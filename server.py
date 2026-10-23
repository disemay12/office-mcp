"""MCP bridge for Microsoft Word, Excel and PowerPoint on Windows.

Run over stdio:  python -X utf8 server.py
"""

from __future__ import annotations

import os
import sys
from typing import Any

from mcp.server.fastmcp import FastMCP

import excel_api
import ppt_api
import word_api
from com import (
    BASE_DIR,
    DISPATCHER,
    OfficeError,
    default_output,
    file_info,
    resolve_path,
    run,
)

INSTRUCTIONS = """
Control the locally installed Microsoft Word, Excel and PowerPoint through COM.

Rules of engagement:
- Always work on real files the user names. Call word_open / excel_open / ppt_open
  before editing existing documents; use the *_create tools for new ones.
- Every tool returns a small JSON object. Read paths and counts from it instead of
  guessing.
- Prefer few, meaningful operations over per-cell or per-character loops: write a
  whole range with excel_write_range and a whole table with word_add_table.
- Save explicitly (word_save / excel_save / ppt_save) and export PDF when the user
  asks for a deliverable. Export functions save the source file first.
- Documents with no explicit path are stored in Documents\\Office MCP.
- Closing an app with office_quit discards nothing silently: it refuses when
  unsaved work exists unless save=true or force=true.
- PowerPoint cannot run hidden; its window will always appear.
""".strip()

mcp = FastMCP("office", instructions=INSTRUCTIONS, log_level="WARNING")

for _tool in (
    word_api.word_list_documents,
    word_api.word_create,
    word_api.word_open,
    word_api.word_read,
    word_api.word_outline,
    word_api.word_append,
    word_api.word_prepend,
    word_api.word_set_text,
    word_api.word_add_heading,
    word_api.word_replace_text,
    word_api.word_format_text,
    word_api.word_add_table,
    word_api.word_add_image,
    word_api.word_add_page_break,
    word_api.word_set_header_footer,
    word_api.word_add_toc,
    word_api.word_save,
    word_api.word_export_pdf,
    word_api.word_close,
    excel_api.excel_list_workbooks,
    excel_api.excel_create,
    excel_api.excel_open,
    excel_api.excel_list_sheets,
    excel_api.excel_read_range,
    excel_api.excel_write_range,
    excel_api.excel_format_range,
    excel_api.excel_add_sheet,
    excel_api.excel_delete_sheet,
    excel_api.excel_rename_sheet,
    excel_api.excel_add_chart,
    excel_api.excel_add_table,
    excel_api.excel_autofit,
    excel_api.excel_recalculate,
    excel_api.excel_save,
    excel_api.excel_export_pdf,
    excel_api.excel_close,
    ppt_api.ppt_list_presentations,
    ppt_api.ppt_create,
    ppt_api.ppt_open,
    ppt_api.ppt_add_slide,
    ppt_api.ppt_set_slide,
    ppt_api.ppt_add_textbox,
    ppt_api.ppt_add_image,
    ppt_api.ppt_add_table,
    ppt_api.ppt_read,
    ppt_api.ppt_delete_slide,
    ppt_api.ppt_save,
    ppt_api.ppt_export_pdf,
    ppt_api.ppt_close,
):
    mcp.tool()(_tool)


@mcp.tool()
def office_status() -> dict[str, Any]:
    """Report which Office applications are running and what is open in each.

    Call this first when you are not sure whether Word/Excel/PowerPoint already
    has the user's document, or when a tool says a file is not open.
    """

    def word_detail(app: Any) -> dict[str, Any]:
        return {
            "documents": [
                {"name": doc.Name, "path": doc.FullName if doc.Path else None, "saved": bool(doc.Saved)}
                for doc in app.Documents
            ]
        }

    def excel_detail(app: Any) -> dict[str, Any]:
        return {
            "workbooks": [
                {
                    "name": book.Name,
                    "path": book.FullName if book.Path else None,
                    "saved": bool(book.Saved),
                    "sheets": [sheet.Name for sheet in book.Worksheets],
                }
                for book in app.Workbooks
            ]
        }

    def ppt_detail(app: Any) -> dict[str, Any]:
        return {
            "presentations": [
                {
                    "name": pres.Name,
                    "path": pres.FullName if pres.Path else None,
                    "slides": int(pres.Slides.Count),
                    "saved": bool(pres.Saved),
                }
                for pres in app.Presentations
            ]
        }

    # Each detail function runs on the COM thread; wrap it in the same probe.
    def probe(app_key: str, detail_fn: Any) -> dict[str, Any]:
        def job(app: Any) -> dict[str, Any]:
            payload = {"visible": bool(app.Visible), "running": True}
            payload.update(detail_fn(app))
            return payload

        try:
            return run(app_key, job, create=False)
        except OfficeError as exc:
            return {"running": False, "detail": str(exc)}

    return {
        "base_dir": BASE_DIR,
        "cwd": os.getcwd(),
        "python": sys.version.split()[0],
        "applications": {
            "word": probe("word", word_detail),
            "excel": probe("excel", excel_detail),
            "powerpoint": probe("powerpoint", ppt_detail),
        },
    }


@mcp.tool()
def office_show(app: str) -> dict[str, Any]:
    """Bring a running (or newly started) Office window to the front.

    `app` is one of 'word', 'excel', 'powerpoint'.
    """
    key = str(app).strip().lower()
    if key not in {"word", "excel", "powerpoint"}:
        raise OfficeError("app must be word, excel or powerpoint")

    def op(application: Any) -> dict[str, Any]:
        application.Visible = True
        application.Activate()
        return {"app": key, "visible": bool(application.Visible)}

    return run(key, op)


@mcp.tool()
def office_quit(app: str, save: bool = False, force: bool = False) -> dict[str, Any]:
    """Quit an Office application and close everything in it.

    Refuses to discard unsaved work unless `save=True` (save all first) or
    `force=True` (discard). This closes the user's windows, so use it only when
    the user asks to quit the application.
    """
    key = str(app).strip().lower()
    if key not in {"word", "excel", "powerpoint"}:
        raise OfficeError("app must be word, excel or powerpoint")

    def op(application: Any) -> dict[str, Any]:
        if key == "word":
            documents = [doc for doc in application.Documents]
            unsaved = [doc.Name for doc in documents if not doc.Saved]
            if unsaved and not save and not force:
                raise OfficeError(
                    "Unsaved Word documents: " + ", ".join(unsaved) + ". Pass save=true or force=true."
                )
            for doc in documents:
                if save:
                    if not doc.Path:
                        doc.SaveAs2(default_output("word", "docx"), 16)
                    doc.Save()
                doc.Close(0)
            application.Quit(0)
            count = len(documents)
        elif key == "excel":
            books = [book for book in application.Workbooks]
            unsaved = [book.Name for book in books if not book.Saved]
            if unsaved and not save and not force:
                raise OfficeError(
                    "Unsaved Excel workbooks: " + ", ".join(unsaved) + ". Pass save=true or force=true."
                )
            for book in books:
                if save:
                    if not book.Path:
                        book.SaveAs(default_output("excel", "xlsx"))
                    book.Save()
                book.Close(False)
            application.Quit()
            count = len(books)
        else:
            presentations = [pres for pres in application.Presentations]
            unsaved = [pres.Name for pres in presentations if not pres.Saved]
            if unsaved and not save and not force:
                raise OfficeError(
                    "Unsaved PowerPoint presentations: "
                    + ", ".join(unsaved)
                    + ". Pass save=true or force=true."
                )
            for pres in presentations:
                if save:
                    if not pres.Path:
                        pres.SaveAs(default_output("powerpoint", "pptx"))
                    pres.Save()
                pres.Close()
            application.Quit()
            count = len(presentations)

        DISPATCHER.forget(key)
        return {"app": key, "closed_items": count, "quit": True}

    return run(key, op, create=False)


@mcp.tool()
def office_open_file(path: str) -> dict[str, Any]:
    """Open a saved file in its default Windows application (e.g. show a generated PDF)."""
    source = resolve_path(path, must_exist=True)
    os.startfile(source)  # noqa: S606 - intentional shell association
    return {"opened": source, "file": file_info(source)}


def main() -> None:
    mcp.run(transport="stdio")


if __name__ == "__main__":
    main()
