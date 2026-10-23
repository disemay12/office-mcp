"""Word automation exposed as MCP tools."""

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

# Built-in Word style IDs are locale-independent. Localized style names are not.
STYLE_IDS = {
    "normal": -1,
    "body text": -1,
    "title": -63,
    "subtitle": -75,
    "heading 1": -2,
    "heading 2": -3,
    "heading 3": -4,
    "heading 4": -5,
    "heading 5": -6,
    "heading 6": -7,
    "heading 7": -8,
    "heading 8": -9,
    "heading 9": -10,
}

SAVE_FORMATS = {
    ".docx": 16,
    ".docm": 13,
    ".doc": 0,
    ".pdf": 17,
    ".rtf": 6,
    ".txt": 2,
    ".odt": 23,
    ".htm": 8,
    ".html": 8,
}


def _style_ref(style: str | int | None) -> Any:
    if style is None:
        return None
    if isinstance(style, int):
        return style
    key = re.sub(r"\s+", " ", str(style).strip().lower())
    if key in STYLE_IDS:
        return STYLE_IDS[key]
    match = re.fullmatch(r"(?:heading|h)\s*([1-9])", key)
    if match:
        return -1 - int(match.group(1))
    return str(style)


def _doc(app: Any, identifier: str | None) -> Any:
    docs = app.Documents
    if docs.Count == 0:
        raise OfficeError("No Word document is open. Use word_create or word_open first.")
    if identifier is None or str(identifier).strip() == "":
        try:
            return app.ActiveDocument
        except pywintypes.com_error:
            return docs(1)

    ident = str(identifier).strip()
    if is_index(ident):
        index = int(ident)
        if not 1 <= index <= docs.Count:
            raise OfficeError(f"Document index {index} is out of range (1..{docs.Count}).")
        return docs(index)

    wanted_path = os.path.normcase(os.path.abspath(ident)) if os.path.isabs(ident) else None
    lowered = ident.lower()
    for doc in docs:
        try:
            if wanted_path and os.path.normcase(os.path.abspath(doc.FullName)) == wanted_path:
                return doc
            if doc.Name.lower() == lowered:
                return doc
        except pywintypes.com_error:
            continue
    names = []
    for doc in docs:
        try:
            names.append(doc.Name)
        except pywintypes.com_error:
            pass
    hint = ""
    if os.path.isfile(ident):
        hint = " The file exists but is not open — call word_open first."
    raise OfficeError(
        f"Word document '{identifier}' is not open.{hint} Open documents: {', '.join(names) or 'none'}."
    )


def _plain_text(doc: Any) -> str:
    raw = doc.Content.Text or ""
    return (
        raw.replace("\r\x07", "\t")
        .replace("\x07", "\t")
        .replace("\r", "\n")
        .replace("\x0b", "\n")
        .replace("\x0c", "\n")
    )


def _is_blank(doc: Any) -> bool:
    content = doc.Content
    end = int(content.End)
    if end <= 1:
        return True
    if end > 64:
        return False  # avoid reading the whole body just to detect whitespace
    return (content.Text or "").strip("\r\x07\x0c \t") == ""


def _append(doc: Any, text: str, style: str | int | None = None) -> int:
    """Append text as one paragraph per line; return the number of paragraphs."""
    style_ref = _style_ref(style)
    if style_ref is None:
        style_ref = -1  # Normal: appended paragraphs would otherwise inherit the previous style
    lines = str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")
    blank = _is_blank(doc)
    payload = "\r".join(lines)
    if not blank:
        payload = "\r" + payload

    if not payload and blank:
        doc.Paragraphs.Last.Style = style_ref
        return len(lines)

    # One insert + one style assignment instead of two COM calls per line.
    inserted = doc.Content
    inserted.Collapse(0)
    inserted.InsertAfter(payload)
    if not blank:
        inserted.Start = inserted.Start + 1  # skip the leading paragraph mark
    try:
        inserted.Style = style_ref
    except pywintypes.com_error as exc:
        raise OfficeError(
            f"Unknown Word style '{style}'. Use 'Normal', 'Title', 'Heading 1'..'Heading 9' "
            "or a built-in style ID."
        ) from exc
    return len(lines)


def _document_payload(doc: Any) -> dict[str, Any]:
    return {
        "name": doc.Name,
        "path": doc.FullName if doc.Path else None,
        "saved": bool(doc.Saved),
        "paragraphs": int(doc.Paragraphs.Count),
        "words": int(doc.ComputeStatistics(0)),
        "pages": int(doc.ComputeStatistics(2)),
        "tables": int(doc.Tables.Count),
        "inline_images": int(doc.InlineShapes.Count),
        "active": False,
    }


# -- documents ----------------------------------------------------------


def word_list_documents() -> dict[str, Any]:
    """List every open Word document with its path, paragraph/page counts and saved state."""

    def op(app: Any) -> dict[str, Any]:
        items = []
        active = None
        try:
            active = app.ActiveDocument.Name
        except pywintypes.com_error:
            pass
        for doc in app.Documents:
            item = _document_payload(doc)
            item["active"] = item["name"] == active
            items.append(item)
        return {"count": len(items), "documents": items}

    return run("word", op, create=False)


def word_create(path: str | None = None, visible: bool = True) -> dict[str, Any]:
    """Create a new Word document. Optionally save it immediately at `path` (.docx/.doc/.rtf).

    Without `path` the document stays unsaved; later word_save picks a file in
    Documents\\Office MCP. Set visible=False to work without a window (an unsaved
    invisible document is lost when the MCP server exits).
    """

    def op(app: Any) -> dict[str, Any]:
        target = prepare_output(path) if path else None
        if target and os.path.exists(target):
            raise OfficeError(f"File already exists: {target}. Use word_open or choose another path.")
        doc = app.Documents.Add()
        if target:
            doc.SaveAs2(target, SAVE_FORMATS.get(os.path.splitext(target)[1].lower(), 16))
        payload = _document_payload(doc)
        payload["message"] = "Document created"
        return payload

    return run("word", op, visible=visible)


def word_open(path: str, visible: bool = True, read_only: bool = False) -> dict[str, Any]:
    """Open an existing Word document (or focus it if already open) and return its details."""
    source = resolve_path(path, must_exist=True)

    def op(app: Any) -> dict[str, Any]:
        for doc in app.Documents:
            try:
                if os.path.normcase(doc.FullName) == os.path.normcase(source):
                    doc.Activate()
                    payload = _document_payload(doc)
                    payload["message"] = "Document was already open"
                    return payload
            except pywintypes.com_error:
                continue
        doc = app.Documents.Open(source, False, bool(read_only), False)
        doc.Activate()
        payload = _document_payload(doc)
        payload["message"] = "Document opened"
        return payload

    return run("word", op, visible=visible)


def word_read(
    document: str | None = None,
    max_chars: int = 20000,
) -> dict[str, Any]:
    """Read a Word document as plain text. Tables render as tab-separated cell values.

    `document` accepts a file name, full path or 1-based index; omit it for the
    active document. Text is capped at `max_chars` (default 20000).
    """

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        limit = max(1000, min(int(max_chars), 400000))
        text = _plain_text(doc)
        truncated = len(text) > limit
        return {
            "name": doc.Name,
            "path": doc.FullName if doc.Path else None,
            "text": text[:limit],
            "truncated": truncated,
            "chars": len(text),
            "paragraphs": int(doc.Paragraphs.Count),
            "tables": int(doc.Tables.Count),
        }

    return run("word", op, create=False)


def _native_outline(doc: Any, cap: int) -> list[dict[str, Any]] | None:
    """Collect headings with nine native Find cursors (one per outline level).

    Searching by outline level avoids one COM round trip per paragraph, which
    dominates long documents. Returns None when the Word build rejects the
    outline-level search so the caller can fall back to the paragraph scan.
    """

    # The final paragraph mark confuses a formatting-only Find, and a single
    # heading paragraph has no text before it, so search the body only.
    body_end = max(0, int(doc.Content.End) - 1)

    def make_cursor(level: int) -> dict[str, Any]:
        rng = doc.Range(0, body_end)
        finder = rng.Find
        finder.ClearFormatting()
        finder.Text = ""
        finder.Format = True
        finder.ParagraphFormat.OutlineLevel = level
        finder.Forward = True
        finder.Wrap = 0
        return {"level": level, "range": rng, "finder": finder, "seen": -1, "next": None}

    def advance(cursor: dict[str, Any]) -> tuple[int, int, str] | None:
        rng = cursor["range"]
        finder = cursor["finder"]
        while finder.Execute():
            start = int(rng.Start)
            if start <= cursor["seen"]:
                return None
            cursor["seen"] = start
            text = re.sub(r"[\r\x07]+$", "", rng.Text or "").strip()
            if text:
                return start, cursor["level"], text
            rng.Collapse(0)
        return None

    try:
        cursors = []
        for level in range(1, 10):
            cursor = make_cursor(level)
            cursor["next"] = advance(cursor)
            cursors.append(cursor)

        headings: list[dict[str, Any]] = []
        last_start = -1
        while len(headings) < cap:
            available = [cursor for cursor in cursors if cursor["next"] is not None]
            if not available:
                break
            chosen = min(available, key=lambda cursor: cursor["next"][0])
            start, level, text = chosen["next"]
            if start != last_start:
                headings.append({"level": level, "text": text})
                last_start = start
            chosen["range"].Collapse(0)
            chosen["next"] = advance(chosen)
        return headings
    except pywintypes.com_error:
        return None


def word_outline(document: str | None = None, limit: int = 300) -> dict[str, Any]:
    """Return the heading outline of a Word document (level + text) plus document statistics.

    Useful before editing a long document: it shows numbered sections without
    pulling the whole text into the chat.
    """

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        cap = max(1, min(int(limit), 2000))
        headings = _native_outline(doc, cap)
        if headings is None:
            # Fallback for Word builds that cannot search by outline level.
            headings = []
            for paragraph in doc.Paragraphs:
                if len(headings) >= cap:
                    break
                try:
                    level = int(paragraph.OutlineLevel)
                except pywintypes.com_error:
                    continue
                if 1 <= level <= 9:
                    text = re.sub(r"[\r\x07]+$", "", paragraph.Range.Text or "")
                    if text.strip():
                        headings.append({"level": level, "text": text.strip()})
        return {
            "name": doc.Name,
            "path": doc.FullName if doc.Path else None,
            "headings": headings,
            "truncated": len(headings) >= cap,
            "paragraphs": int(doc.Paragraphs.Count),
            "tables": int(doc.Tables.Count),
            "words": int(doc.ComputeStatistics(0)),
            "pages": int(doc.ComputeStatistics(2)),
        }

    return run("word", op, create=False)


def word_append(
    text: str,
    document: str | None = None,
    style: str | None = None,
) -> dict[str, Any]:
    """Append paragraphs to the end of a Word document.

    Each newline in `text` becomes a separate paragraph. `style` accepts
    'Normal', 'Title', 'Subtitle', 'Heading 1'..'Heading 9' or a built-in style ID;
    omitted style means Normal, so text never inherits the previous paragraph's style.
    """

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        count = _append(doc, text, style)
        return {
            "name": doc.Name,
            "paragraphs_added": count,
            "paragraphs": int(doc.Paragraphs.Count),
            "saved": bool(doc.Saved),
        }

    return run("word", op, create=False)


def word_prepend(
    text: str,
    document: str | None = None,
    style: str | None = None,
) -> dict[str, Any]:
    """Insert paragraphs at the very beginning of a Word document (e.g. a title page)."""

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        lines = str(text).replace("\r\n", "\n").replace("\r", "\n").split("\n")
        payload = "\r".join(lines) + "\r"
        inserted = doc.Range(0, 0)
        inserted.InsertAfter(payload)
        style_ref = _style_ref(style)
        if style_ref is None:
            style_ref = -1
        try:
            inserted.Style = style_ref
        except pywintypes.com_error as exc:
            raise OfficeError(
                f"Unknown Word style '{style}'. Use 'Normal', 'Title', 'Heading 1'..'Heading 9' "
                "or a built-in style ID."
            ) from exc
        return {"name": doc.Name, "paragraphs_added": len(lines), "saved": bool(doc.Saved)}

    return run("word", op, create=False)


def word_set_text(
    text: str,
    document: str | None = None,
    style: str | None = None,
) -> dict[str, Any]:
    """Replace the entire body of a Word document with `text` (headers/footers survive)."""

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        doc.Content.Delete()
        _append(doc, text, style)
        return {
            "name": doc.Name,
            "paragraphs": int(doc.Paragraphs.Count),
            "words": int(doc.ComputeStatistics(0)),
            "saved": bool(doc.Saved),
        }

    return run("word", op, create=False)


def word_add_heading(
    text: str,
    level: int = 1,
    document: str | None = None,
) -> dict[str, Any]:
    """Append a heading paragraph (level 1..9) to a Word document."""

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        lvl = max(1, min(int(level), 9))
        _append(doc, text, f"Heading {lvl}")
        return {"name": doc.Name, "heading": text, "level": lvl, "saved": bool(doc.Saved)}

    return run("word", op, create=False)


def word_replace_text(
    find: str,
    replace: str,
    document: str | None = None,
    match_case: bool = True,
    all_stories: bool = False,
) -> dict[str, Any]:
    """Find and replace text in a Word document; returns the number of replacements.

    `all_stories=True` also covers headers, footers, footnotes and text boxes.
    """

    if not find:
        raise OfficeError("find must not be empty")

    def replace_in(target: Any, wrap: int) -> bool:
        finder = target.Find
        finder.ClearFormatting()
        finder.Replacement.ClearFormatting()
        # Positional arguments: pywin32's dynamic dispatch ignores named args here.
        return bool(
            finder.Execute(
                find,
                bool(match_case),
                False,
                False,
                False,
                False,
                True,
                wrap,
                False,
                replace,
                2,  # wdReplaceAll
            )
        )

    def count_in(target: Any) -> int:
        rng = target.Duplicate
        found = 0
        while found < 5000:
            finder = rng.Find
            finder.ClearFormatting()
            if not finder.Execute(
                find, bool(match_case), False, False, False, False, True, 0, False, "", 0
            ):
                break
            found += 1
            rng.Collapse(0)  # continue after the found text
        return found

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        replacements = 0
        chunks = 0
        if all_stories:
            stories = []
            for story in doc.StoryRanges:
                current = story
                while current is not None and len(stories) < 40:
                    stories.append(current)
                    try:
                        current = current.NextStoryRange
                    except pywintypes.com_error:
                        current = None
            for story in stories:
                replacements += count_in(story)
                if replace_in(story, 0):
                    chunks += 1
        else:
            replacements = count_in(doc.Content)
            if replace_in(doc.Content, 1):
                chunks += 1
        return {
            "name": doc.Name,
            "replacements": replacements,
            "regions_changed": chunks,
            "saved": bool(doc.Saved),
        }

    return run("word", op, create=False)


def word_format_text(
    find: str,
    document: str | None = None,
    bold: bool | None = None,
    italic: bool | None = None,
    underline: bool | None = None,
    font_size: float | None = None,
    color: str | None = None,
    highlight: bool | None = None,
    match_case: bool = True,
) -> dict[str, Any]:
    """Apply character formatting to every occurrence of `find` in a Word document.

    `color` is a hex string like '#C00000'. Only the provided properties change.
    """

    if not find:
        raise OfficeError("find must not be empty")

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        target = doc.Content
        changed = 0
        while changed < 1000:
            finder = target.Find
            finder.ClearFormatting()
            if not finder.Execute(
                find, bool(match_case), False, False, False, False, True, 0, False, "", 0
            ):
                break
            font = target.Font
            if bold is not None:
                font.Bold = -1 if bold else 0
            if italic is not None:
                font.Italic = -1 if italic else 0
            if underline is not None:
                font.Underline = 1 if underline else 0
            if font_size is not None:
                font.Size = float(font_size)
            if color is not None:
                font.Color = hex_to_bgr(color)
            if highlight is not None:
                target.HighlightColorIndex = 7 if highlight else 0  # wdYellow / wdNoHighlight
            changed += 1
            target.Collapse(0)
        return {"name": doc.Name, "occurrences_formatted": changed, "saved": bool(doc.Saved)}

    return run("word", op, create=False)


def word_add_table(
    rows: list[list[Any]],
    document: str | None = None,
    header: bool = True,
    autofit: bool = True,
) -> dict[str, Any]:
    """Append a bordered table to a Word document.

    `rows` is a 2-D list; each inner list is a row. With `header=True` the first
    row is bold and shaded. Tables are appended at the end of the document body.
    """

    matrix = normalize_matrix(rows)
    nrows, ncols = len(matrix), len(matrix[0])
    if nrows * ncols == 0:
        raise OfficeError("Table must have at least one row and one column")

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        if not _is_blank(doc):
            doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
        start = doc.Content.End - 1
        table = doc.Tables.Add(doc.Range(start, start), nrows, ncols)
        table.Borders.InsideLineStyle = 1
        table.Borders.OutsideLineStyle = 1
        if autofit:
            try:
                table.AutoFitBehavior(2)  # wdAutoFitWindow
            except pywintypes.com_error:
                pass
        for row_index, row in enumerate(matrix, start=1):
            for col_index, value in enumerate(row, start=1):
                table.Cell(row_index, col_index).Range.Text = (
                    "" if value is None else str(value)
                )
        # One range-wide assignment instead of one COM call per cell.
        table.Range.ParagraphFormat.SpaceAfter = 0
        if header and nrows > 1:
            head = table.Rows(1)
            head.Range.Font.Bold = True
            head.Shading.BackgroundPatternColor = hex_to_bgr("#E7E6E6")
            head.HeadingFormat = True
        return {
            "name": doc.Name,
            "rows": nrows,
            "columns": ncols,
            "tables": int(doc.Tables.Count),
            "saved": bool(doc.Saved),
        }

    return run("word", op, create=False)


def word_add_image(
    image_path: str,
    document: str | None = None,
    width_cm: float | None = None,
    caption: str | None = None,
) -> dict[str, Any]:
    """Insert an inline image at the end of a Word document, optionally with a caption."""

    source = resolve_path(image_path, must_exist=True)

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        if not _is_blank(doc):
            doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertAfter("\r")
        start = doc.Content.End - 1
        shape = doc.InlineShapes.AddPicture(source, False, True, doc.Range(start, start))
        if width_cm is not None:
            shape.Width = cm_to_points(width_cm)
        if caption:
            doc.Content.InsertAfter("\r" + caption)
        return {
            "name": doc.Name,
            "inline_images": int(doc.InlineShapes.Count),
            "width_cm": round(cm_to_points(width_cm) / 28.3464567, 2) if width_cm else None,
            "saved": bool(doc.Saved),
        }

    return run("word", op, create=False)


def word_add_page_break(document: str | None = None) -> dict[str, Any]:
    """Insert a page break at the end of a Word document."""

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        doc.Range(doc.Content.End - 1, doc.Content.End - 1).InsertBreak(7)  # wdPageBreak
        return {"name": doc.Name, "pages": int(doc.ComputeStatistics(2)), "saved": bool(doc.Saved)}

    return run("word", op, create=False)


def word_set_header_footer(
    text: str,
    where: str = "header",
    document: str | None = None,
    section: int = 1,
    page_numbers: bool = False,
) -> dict[str, Any]:
    """Set the header or footer text of a Word section; optionally add right-aligned page numbers."""

    side = str(where).strip().lower()
    if side not in {"header", "footer"}:
        raise OfficeError("where must be 'header' or 'footer'")

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        sections = doc.Sections
        index = int(section)
        if not 1 <= index <= sections.Count:
            raise OfficeError(f"Section {index} is out of range (1..{sections.Count}).")
        target = sections(index).Headers(1) if side == "header" else sections(index).Footers(1)
        target.Range.Text = text or ""
        if page_numbers:
            try:
                target.PageNumbers.Add(2)  # wdAlignPageNumberRight
            except pywintypes.com_error as exc:
                raise OfficeError(f"Cannot add page numbers: {exc}") from exc
        return {
            "name": doc.Name,
            "section": index,
            "where": side,
            "page_numbers": bool(page_numbers),
            "saved": bool(doc.Saved),
        }

    return run("word", op, create=False)


def word_add_toc(
    document: str | None = None,
    levels: int = 3,
    at_start: bool = True,
) -> dict[str, Any]:
    """Insert a table of contents built from heading styles 1..`levels`."""

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        depth = max(1, min(int(levels), 9))
        if at_start:
            start = doc.Range(0, 0)
            start.InsertAfter("\r")
            start = doc.Range(0, 0)
        else:
            start = doc.Range(doc.Content.End - 1, doc.Content.End - 1)
        toc = doc.TablesOfContents.Add(start, True, 1, depth, False, "", True, True)
        toc.Update()
        doc.Repaginate()
        return {
            "name": doc.Name,
            "levels": depth,
            "pages": int(doc.ComputeStatistics(2)),
            "saved": bool(doc.Saved),
        }

    return run("word", op, create=False)


def word_save(document: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Save a Word document. Without `path` it saves in place; unsaved documents get an automatic path."""

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        target = prepare_output(path) if path else None
        if target:
            fmt = SAVE_FORMATS.get(os.path.splitext(target)[1].lower(), 16)
            doc.SaveAs2(target, fmt)
        elif not doc.Path:
            target = default_output("word", "docx")
            doc.SaveAs2(target, 16)
        else:
            doc.Save()
        return {
            "name": doc.Name,
            "path": doc.FullName,
            "saved": bool(doc.Saved),
            "file": file_info(doc.FullName),
        }

    return run("word", op, create=False)


def word_export_pdf(document: str | None = None, path: str | None = None) -> dict[str, Any]:
    """Export a Word document to PDF (saving the document first if it is unsaved)."""

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        if not doc.Path:
            doc.SaveAs2(default_output("word", "docx"), 16)
        target = prepare_output(path) if path else default_output("word", "pdf")
        doc.ExportAsFixedFormat(target, 17)  # wdExportFormatPDF
        return {
            "name": doc.Name,
            "document": doc.FullName,
            "pdf": target,
            "file": file_info(target),
        }

    return run("word", op, create=False)


def word_close(document: str, save: bool = False) -> dict[str, Any]:
    """Close a Word document; `save=True` saves it first (unsaved files get an automatic path)."""

    def op(app: Any) -> dict[str, Any]:
        doc = _doc(app, document)
        name = doc.Name
        if save and not doc.Path:
            doc.SaveAs2(default_output("word", "docx"), 16)
        doc.Close(-1 if save else 0)
        return {"closed": name, "documents_left": int(app.Documents.Count)}

    return run("word", op, create=False)
