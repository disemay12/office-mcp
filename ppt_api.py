"""PowerPoint automation exposed as MCP tools."""

from __future__ import annotations

import os
import re
from typing import Any

import pywintypes

from com import (
    OfficeError,
    cm_to_points,
    default_output,
    file_info,
    hex_to_bgr,
    is_index,
    normalize_matrix,
    prepare_output,
    resolve_path,
    run,
)

LAYOUTS = {
    "title": 1,
    "title_content": 2,
    "two_content": 3,
    "comparison": 4,
    "title_only": 11,
    "blank": 12,
}

BODY_PLACEHOLDER = 2


def _pres(app: Any, identifier: str | None) -> Any:
    pres_list = app.Presentations
    if pres_list.Count == 0:
        raise OfficeError("No PowerPoint presentation is open. Use ppt_create or ppt_open first.")
    if identifier is None or str(identifier).strip() == "":
        try:
            return app.ActivePresentation
        except pywintypes.com_error:
            return pres_list(1)

    ident = str(identifier).strip()
    if is_index(ident):
        index = int(ident)
        if not 1 <= index <= pres_list.Count:
            raise OfficeError(f"Presentation index {index} is out of range (1..{pres_list.Count}).")
        return pres_list(index)

    wanted_path = os.path.normcase(os.path.abspath(ident)) if os.path.isabs(ident) else None
    lowered = ident.lower()
    for pres in pres_list:
        try:
            if wanted_path and os.path.normcase(os.path.abspath(pres.FullName)) == wanted_path:
                return pres
            if pres.Name.lower() == lowered:
                return pres
        except pywintypes.com_error:
            continue
    names = []
    for pres in pres_list:
        try:
            names.append(pres.Name)
        except pywintypes.com_error:
            pass
    hint = ""
    if os.path.isfile(ident):
        hint = " The file exists but is not open — call ppt_open first."
    raise OfficeError(
        f"Presentation '{identifier}' is not open.{hint} Open presentations: {', '.join(names) or 'none'}."
    )


def _slide(pres: Any, number: int) -> Any:
    index = int(number)
    if not 1 <= index <= pres.Slides.Count:
        raise OfficeError(f"Slide {index} is out of range (1..{pres.Slides.Count}).")
    return pres.Slides(index)


def _body_shape(slide: Any) -> Any:
    for index in range(1, slide.Shapes.Count + 1):
        shape = slide.Shapes(index)
        try:
            if shape.HasTextFrame and int(shape.PlaceholderFormat.Type) == BODY_PLACEHOLDER:
                return shape
        except pywintypes.com_error:
            continue
    return None


def _notes_shape(slide: Any) -> Any:
    try:
        notes = slide.NotesPage
        for index in range(1, notes.Shapes.Count + 1):
            shape = notes.Shapes(index)
            try:
                if shape.HasTextFrame and int(shape.PlaceholderFormat.Type) == BODY_PLACEHOLDER:
                    return shape
            except pywintypes.com_error:
                continue
    except pywintypes.com_error:
        return None
    return None


def _parse_bullets(bullets: list[str]) -> list[tuple[int, str]]:
    parsed: list[tuple[int, str]] = []
    for raw in bullets:
        text = str(raw).rstrip()
        stripped = text.lstrip(" \t")
        indent = text[: len(text) - len(stripped)]
        level = indent.count("\t") + (len(indent.replace("\t", "")) // 2)
        cleaned = re.sub(r"^[-*\u2022]\s+", "", stripped)
        parsed.append((max(0, min(level, 4)), cleaned))
    return parsed


def _write_bullets(shape: Any, bullets: list[str], append: bool = False) -> int:
    parsed = _parse_bullets(bullets)
    text_range = shape.TextFrame.TextRange
    if append and text_range.Text:
        text_range.Text = text_range.Text.rstrip("\r") + "\r" + "\r".join(t for _, t in parsed)
        start_index = max(1, int(text_range.Paragraphs().Count) - len(parsed) + 1)
    else:
        text_range.Text = "\r".join(t for _, t in parsed)
        start_index = 1
    if not parsed:
        text_range.Text = ""
        return 0

    # One bullet-format call and one IndentLevel call per level run instead of
    # two COM calls per bullet.
    try:
        text_range.Paragraphs(start_index, len(parsed)).ParagraphFormat.Bullet.Visible = -1
    except pywintypes.com_error:
        for offset in range(len(parsed)):
            try:
                text_range.Paragraphs(start_index + offset).ParagraphFormat.Bullet.Visible = -1
            except pywintypes.com_error:  # some layouts have no bullet formatting
                pass

    offset = 0
    while offset < len(parsed):
        level = parsed[offset][0]
        run = 1
        while offset + run < len(parsed) and parsed[offset + run][0] == level:
            run += 1
        try:
            text_range.Paragraphs(start_index + offset, run).IndentLevel = level + 1
        except pywintypes.com_error:
            for index in range(offset, offset + run):
                text_range.Paragraphs(start_index + index).IndentLevel = level + 1
        offset += run
    return len(parsed)


def _set_notes(slide: Any, text: str) -> None:
    shape = _notes_shape(slide)
    if shape is None:
        raise OfficeError("This slide has no notes placeholder")
    shape.TextFrame.TextRange.Text = text


def _shape_texts(shapes: Any, depth: int = 0) -> list[str]:
    chunks: list[str] = []
    for index in range(1, shapes.Count + 1):
        shape = shapes(index)
        try:
            if shape.HasTable:
                table = shape.Table
                for row in range(1, table.Rows.Count + 1):
                    cells = []
                    for col in range(1, table.Columns.Count + 1):
                        cell_range = table.Cell(row, col).Shape.TextFrame.TextRange
                        cells.append((cell_range.Text or "").replace("\r", " ").strip())
                    chunks.append(" | ".join(cells))
                continue
            if shape.HasTextFrame and shape.TextFrame.HasText:
                text = (shape.TextFrame.TextRange.Text or "").strip()
                if text:
                    chunks.append(text.replace("\r", "\n"))
            if depth < 1 and shape.Type == 6:  # msoGroup
                chunks.extend(_shape_texts(shape.GroupItems, depth + 1))
        except pywintypes.com_error:
            continue
    return chunks


# -- presentations ------------------------------------------------------


def ppt_list_presentations() -> dict[str, Any]:
    """List open PowerPoint presentations with slide counts and saved state."""

    def op(app: Any) -> dict[str, Any]:
        items = []
        active = None
        try:
            active = app.ActivePresentation.Name
        except pywintypes.com_error:
            pass
        for pres in app.Presentations:
            items.append(
                {
                    "name": pres.Name,
                    "path": pres.FullName if pres.Path else None,
                    "slides": int(pres.Slides.Count),
                    "saved": bool(pres.Saved),
                    "active": pres.Name == active,
                }
            )
        return {"count": len(items), "presentations": items}

    return run("powerpoint", op, create=False)


def ppt_create(path: str | None = None) -> dict[str, Any]:
    """Create a new PowerPoint presentation, optionally saved at `path` (.pptx/.pptm).

    A new presentation starts with zero slides; add them with ppt_add_slide.
    PowerPoint is always visible (it cannot run hidden), so a window will appear.
    """

    def op(app: Any) -> dict[str, Any]:
        target = prepare_output(path) if path else None
        if target and os.path.exists(target):
            raise OfficeError(f"File already exists: {target}. Use ppt_open or choose another path.")
        pres = app.Presentations.Add()
        if target:
            pres.SaveAs(target)
        return {
            "name": pres.Name,
            "path": pres.FullName if pres.Path else None,
            "slides": int(pres.Slides.Count),
            "saved": bool(pres.Saved),
        }

    return run("powerpoint", op)


def ppt_open(path: str) -> dict[str, Any]:
    """Open an existing PowerPoint presentation (or focus it if already open)."""
    source = resolve_path(path, must_exist=True)

    def op(app: Any) -> dict[str, Any]:
        for pres in app.Presentations:
            try:
                if os.path.normcase(pres.FullName) == os.path.normcase(source):
                    pres.Activate()
                    return {
                        "name": pres.Name,
                        "path": pres.FullName,
                        "slides": int(pres.Slides.Count),
                        "message": "Presentation was already open",
                    }
            except pywintypes.com_error:
                continue
        pres = app.Presentations.Open(source, False, False, True)
        pres.Activate()
        return {
            "name": pres.Name,
            "path": pres.FullName,
            "slides": int(pres.Slides.Count),
            "message": "Presentation opened",
        }

    return run("powerpoint", op)


def ppt_add_slide(
    title: str | None = None,
    bullets: list[str] | None = None,
    layout: str = "title_content",
    notes: str | None = None,
    index: int | None = None,
    presentation: str | None = None,
) -> dict[str, Any]:
    """Add a slide and fill its title, bullet body and speaker notes.

    `layout` is one of: title_content, title_only, blank, title, two_content,
    comparison. Bullet levels come from leading tabs or pairs of spaces; a
    leading '- ' is stripped. `index` inserts at a 1-based position.
    """

    key = str(layout).strip().lower()
    if key not in LAYOUTS:
        raise OfficeError(f"Unknown layout '{layout}'. Use: {', '.join(LAYOUTS)}")

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        position = pres.Slides.Count + 1 if index is None else int(index)
        if not 1 <= position <= pres.Slides.Count + 1:
            raise OfficeError(f"Slide index {position} is out of range.")
        slide = pres.Slides.Add(position, LAYOUTS[key])

        if title:
            if slide.Shapes.HasTitle:
                slide.Shapes.Title.TextFrame.TextRange.Text = title
            else:
                box = slide.Shapes.AddTextbox(
                    1, cm_to_points(1.5), cm_to_points(1.0), cm_to_points(23), cm_to_points(2.5)
                )
                box.TextFrame.TextRange.Text = title
                box.TextFrame.TextRange.Font.Size = 30
                box.TextFrame.TextRange.Font.Bold = True

        body = _body_shape(slide)
        if bullets:
            if body is None:
                body = slide.Shapes.AddTextbox(
                    1, cm_to_points(1.5), cm_to_points(4.0), cm_to_points(23), cm_to_points(12)
                )
            _write_bullets(body, bullets)
        elif body is not None and bullets is not None:
            _write_bullets(body, [])

        if notes:
            _set_notes(slide, notes)

        return {
            "presentation": pres.Name,
            "slide": int(slide.SlideIndex),
            "slides": int(pres.Slides.Count),
            "layout": key,
            "bullets": len(bullets or []),
            "saved": bool(pres.Saved),
        }

    return run("powerpoint", op, create=False)


def ppt_set_slide(
    slide_number: int,
    title: str | None = None,
    bullets: list[str] | None = None,
    notes: str | None = None,
    append_bullets: bool = False,
    presentation: str | None = None,
) -> dict[str, Any]:
    """Replace the title, bullets and/or notes of an existing slide.

    Omitted arguments stay untouched. `append_bullets=True` adds to the existing
    body instead of replacing it.
    """

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        slide = _slide(pres, slide_number)
        if title is not None:
            if not slide.Shapes.HasTitle:
                raise OfficeError(f"Slide {slide_number} has no title placeholder")
            slide.Shapes.Title.TextFrame.TextRange.Text = title
        if bullets is not None:
            body = _body_shape(slide)
            if body is None:
                raise OfficeError(f"Slide {slide_number} has no body placeholder; use ppt_add_textbox")
            _write_bullets(body, bullets, append=append_bullets)
        if notes is not None:
            _set_notes(slide, notes)
        return {
            "presentation": pres.Name,
            "slide": int(slide.SlideIndex),
            "saved": bool(pres.Saved),
        }

    return run("powerpoint", op, create=False)


def ppt_add_textbox(
    slide_number: int,
    text: str,
    left_cm: float = 1.5,
    top_cm: float = 1.5,
    width_cm: float = 20,
    height_cm: float = 4,
    font_size: float | None = None,
    bold: bool | None = None,
    color: str | None = None,
    presentation: str | None = None,
) -> dict[str, Any]:
    """Add a free text box to a slide. Colours use '#RRGGBB'."""

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        slide = _slide(pres, slide_number)
        box = slide.Shapes.AddTextbox(
            1,
            cm_to_points(left_cm),
            cm_to_points(top_cm),
            cm_to_points(width_cm),
            cm_to_points(height_cm),
        )
        box.TextFrame.WordWrap = -1
        text_range = box.TextFrame.TextRange
        text_range.Text = str(text).replace("\r\n", "\n").replace("\n", "\r")
        if font_size is not None:
            text_range.Font.Size = float(font_size)
        if bold is not None:
            text_range.Font.Bold = -1 if bold else 0
        if color is not None:
            text_range.Font.Color.RGB = hex_to_bgr(color)
        return {
            "presentation": pres.Name,
            "slide": int(slide.SlideIndex),
            "shape": box.Name,
            "saved": bool(pres.Saved),
        }

    return run("powerpoint", op, create=False)


def ppt_add_image(
    slide_number: int,
    image_path: str,
    left_cm: float = 1.5,
    top_cm: float = 1.5,
    width_cm: float | None = None,
    height_cm: float | None = None,
    presentation: str | None = None,
) -> dict[str, Any]:
    """Place an image on a slide. Omit width/height to keep the original size.

    If only one dimension is given, the aspect ratio is preserved.
    """

    source = resolve_path(image_path, must_exist=True)

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        slide = _slide(pres, slide_number)
        left = cm_to_points(left_cm)
        top = cm_to_points(top_cm)
        width = cm_to_points(width_cm) if width_cm else -1
        height = cm_to_points(height_cm) if height_cm else -1
        shape = slide.Shapes.AddPicture(source, 0, -1, left, top, width, height)
        if (width_cm is None) != (height_cm is None):
            shape.LockAspectRatio = -1
            if width_cm is not None:
                shape.Width = cm_to_points(width_cm)
            else:
                shape.Height = cm_to_points(height_cm)
        return {
            "presentation": pres.Name,
            "slide": int(slide.SlideIndex),
            "shape": shape.Name,
            "saved": bool(pres.Saved),
        }

    return run("powerpoint", op, create=False)


def ppt_add_table(
    slide_number: int,
    rows: list[list[Any]],
    left_cm: float = 1.5,
    top_cm: float = 3.5,
    width_cm: float = 22,
    height_cm: float = 9,
    header: bool = True,
    presentation: str | None = None,
) -> dict[str, Any]:
    """Add a table to a slide from a 2-D list; the first row can be a bold header."""

    matrix = normalize_matrix(rows)
    nrows, ncols = len(matrix), len(matrix[0])
    if nrows * ncols == 0:
        raise OfficeError("Table must have at least one row and one column")

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        slide = _slide(pres, slide_number)
        shape = slide.Shapes.AddTable(
            nrows,
            ncols,
            cm_to_points(left_cm),
            cm_to_points(top_cm),
            cm_to_points(width_cm),
            cm_to_points(height_cm),
        )
        table = shape.Table
        for row_index, row in enumerate(matrix, start=1):
            for col_index, value in enumerate(row, start=1):
                cell_text = table.Cell(row_index, col_index).Shape.TextFrame.TextRange
                cell_text.Text = "" if value is None else str(value)
                cell_text.Font.Size = 14
                if header and row_index == 1:
                    cell_text.Font.Bold = -1
        return {
            "presentation": pres.Name,
            "slide": int(slide.SlideIndex),
            "shape": shape.Name,
            "rows": nrows,
            "columns": ncols,
            "saved": bool(pres.Saved),
        }

    return run("powerpoint", op, create=False)


def ppt_read(presentation: str | None = None, max_chars: int = 20000) -> dict[str, Any]:
    """Read all text from a presentation: per-slide title, body text, tables and notes.

    Text collection stops at `max_chars` to protect the chat context.
    """

    limit = max(1000, min(int(max_chars), 200000))

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        slides = []
        used = 0
        for index in range(1, pres.Slides.Count + 1):
            slide = pres.Slides(index)
            title = ""
            try:
                if slide.Shapes.HasTitle:
                    title = (slide.Shapes.Title.TextFrame.TextRange.Text or "").strip()
            except pywintypes.com_error:
                pass
            chunks = [chunk for chunk in _shape_texts(slide.Shapes) if chunk and chunk != title]
            body = "\n".join(chunks)
            notes = ""
            notes_shape = _notes_shape(slide)
            if notes_shape is not None:
                notes = (notes_shape.TextFrame.TextRange.Text or "").strip()
            remaining = limit - used
            combined = body + ("\n" + notes if notes else "")
            if remaining <= 0:
                break
            if len(combined) > remaining:
                combined = combined[:remaining]
                body = combined
                notes = ""
            used += len(combined)
            slides.append(
                {
                    "slide": index,
                    "title": title,
                    "text": body,
                    "notes": notes,
                }
            )
        return {
            "presentation": pres.Name,
            "path": pres.FullName if pres.Path else None,
            "slide_count": int(pres.Slides.Count),
            "slides": slides,
            "truncated": used >= limit,
            "chars": used,
        }

    return run("powerpoint", op, create=False)


def ppt_delete_slide(slide_number: int, presentation: str | None = None) -> dict[str, Any]:
    """Delete a slide by its 1-based number."""

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        slide = _slide(pres, slide_number)
        deleted = int(slide.SlideIndex)
        slide.Delete()
        return {
            "presentation": pres.Name,
            "deleted": deleted,
            "slides": int(pres.Slides.Count),
            "saved": bool(pres.Saved),
        }

    return run("powerpoint", op, create=False)


def ppt_save(presentation: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Save a presentation in place or to `path`; unsaved files get an automatic .pptx path."""

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        target = prepare_output(path) if path else None
        if target:
            pres.SaveAs(target)
        elif not pres.Path:
            target = default_output("powerpoint", "pptx")
            pres.SaveAs(target)
        else:
            pres.Save()
        return {
            "presentation": pres.Name,
            "path": pres.FullName,
            "saved": bool(pres.Saved),
            "file": file_info(pres.FullName),
        }

    return run("powerpoint", op, create=False)


def ppt_export_pdf(presentation: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Export a presentation to PDF (saving the .pptx first if it is unsaved)."""

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        if not pres.Path:
            pres.SaveAs(default_output("powerpoint", "pptx"))
        target = prepare_output(path) if path else default_output("powerpoint", "pdf")
        pres.SaveAs(target, 32)  # ppSaveAsPDF; the presentation keeps its .pptx path
        return {
            "presentation": pres.Name,
            "pdf": target,
            "file": file_info(target),
        }

    return run("powerpoint", op, create=False)


def ppt_close(presentation: str, save: bool = False) -> dict[str, Any]:
    """Close a presentation; `save=True` saves it first (unsaved files get an automatic path)."""

    def op(app: Any) -> dict[str, Any]:
        pres = _pres(app, presentation)
        name = pres.Name
        if save and not pres.Path:
            pres.SaveAs(default_output("powerpoint", "pptx"))
        if save:
            pres.Save()
        pres.Close()
        return {"closed": name, "presentations_left": int(app.Presentations.Count)}

    return run("powerpoint", op, create=False)
