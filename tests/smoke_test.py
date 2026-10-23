"""End-to-end smoke test for the Office MCP bridge.

Creates real documents in C:\\Temp\\office-mcp-test, drives Word, Excel and
PowerPoint through every major tool, verifies the artifacts and quits the apps.
Run:  python tests\\smoke_test.py
"""

from __future__ import annotations

import json
import os
import shutil
import sys
import traceback
from pathlib import Path

TEST_DIR = Path(r"C:\Temp\office-mcp-test")
ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))
os.environ["OFFICE_MCP_BASE_DIR"] = str(TEST_DIR)

import excel_api  # noqa: E402
import ppt_api  # noqa: E402
import server  # noqa: E402
import word_api  # noqa: E402

FAILURES: list[str] = []


def check(condition: bool, message: str) -> None:
    if not condition:
        raise AssertionError(message)


def step(name: str, fn) -> object:
    try:
        result = fn()
        text = json.dumps(result, ensure_ascii=False, default=str)
        print(f"PASS  {name}: {text[:260]}")
        return result
    except Exception:
        FAILURES.append(name)
        print(f"FAIL  {name}")
        traceback.print_exc()
        return None


def make_png(path: Path) -> Path:
    try:
        from PIL import Image

        Image.new("RGB", (320, 160), (30, 90, 160)).save(path)
    except Exception:
        import struct
        import zlib

        width, height = 320, 160
        raw = b"".join(b"\x00" + bytes((30, 90, 160)) * width for _ in range(height))

        def chunk(kind: bytes, data: bytes) -> bytes:
            body = kind + data
            return struct.pack(">I", len(data)) + body + struct.pack(">I", zlib.crc32(body))

        png = (
            b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(raw, 9))
            + chunk(b"IEND", b"")
        )
        path.write_bytes(png)
    return path


def main() -> int:
    if TEST_DIR.exists():
        shutil.rmtree(TEST_DIR)
    TEST_DIR.mkdir(parents=True, exist_ok=True)
    image = make_png(TEST_DIR / "logo.png")

    print("== Word ==")
    word_file = str(TEST_DIR / "report.docx")
    word_doc = step("word_create", lambda: word_api.word_create(word_file, visible=False))
    step("word_append", lambda: word_api.word_append("alpha beta alpha", style="Normal"))
    step("word_add_heading", lambda: word_api.word_add_heading("Overview", 1))
    step("word_append_multi", lambda: word_api.word_append("Line one\nLine two"))
    step("word_prepend", lambda: word_api.word_prepend("Quarterly Report", style="Title"))
    replaced = step("word_replace_text", lambda: word_api.word_replace_text("alpha", "omega"))
    step("word_format_text", lambda: word_api.word_format_text("omega", bold=True, color="#C00000"))
    read = step("word_read", lambda: word_api.word_read(max_chars=5000))
    outline = step("word_outline", word_api.word_outline)
    step("word_add_table", lambda: word_api.word_add_table([["Metric", "Value"], ["Revenue", "1.2M"], ["Costs", "0.4M"]]))
    step("word_add_image", lambda: word_api.word_add_image(str(image), width_cm=4))
    step("word_add_page_break", word_api.word_add_page_break)
    step("word_set_header_footer", lambda: word_api.word_set_header_footer("Confidential", "header"))
    step("word_set_footer_numbers", lambda: word_api.word_set_header_footer("", "footer", page_numbers=True))
    step("word_add_toc", word_api.word_add_toc)
    saved = step("word_save", word_api.word_save)
    pdf = step("word_export_pdf", word_api.word_export_pdf)

    if replaced is not None:
        check(replaced["replacements"] == 2, f"expected 2 replacements, got {replaced}")
    if read is not None:
        check("omega" in read["text"], "replacement missing from text")
        check("Quarterly Report" in read["text"], "prepended title missing")
    if outline is not None:
        levels = [item["level"] for item in outline["headings"]]
        check(1 in levels, f"heading 1 missing from outline: {outline['headings']}")
    if saved is not None:
        check(os.path.getsize(saved["path"]) > 1000, "docx is too small")
    if pdf is not None:
        check(pdf["file"]["size_bytes"] > 1000, "pdf is too small")

    print("== Excel ==")
    excel_file = str(TEST_DIR / "sales.xlsx")
    workbook = step("excel_create", lambda: excel_api.excel_create(excel_file, visible=False))
    sheet = workbook["sheets"][0] if workbook else "Sheet1"
    rows = [
        ["Region", "Qty", "Price", "Total"],
        ["North", 10, 2.5, "=B2*C2"],
        ["South", 5, 3.0, "=B3*C3"],
    ]
    written = step(
        "excel_write_range",
        lambda: excel_api.excel_write_range(rows, "A1", sheet=sheet),
    )
    values = step("excel_read_range", lambda: excel_api.excel_read_range("A1:D3", sheet=sheet))
    formatted = step(
        "excel_read_range_formatted",
        lambda: excel_api.excel_read_range("A1:A3", sheet=sheet, formatted=True),
    )
    step(
        "excel_format_range",
        lambda: excel_api.excel_format_range(
            "A1:D1",
            sheet=sheet,
            bold=True,
            bg_color="#DDEBF7",
            borders="all",
            align="center",
        ),
    )
    step(
        "excel_format_numbers",
        lambda: excel_api.excel_format_range("D2:D3", sheet=sheet, number_format="#,##0.00"),
    )
    step("excel_add_table", lambda: excel_api.excel_add_table("A1:D3", "Sales", sheet=sheet))
    step(
        "excel_add_chart",
        lambda: excel_api.excel_add_chart("A1:B3", "column", "Quantity", sheet=sheet, anchor="F2"),
    )
    step("excel_add_sheet", lambda: excel_api.excel_add_sheet("Report"))
    step("excel_rename_sheet", lambda: excel_api.excel_rename_sheet("Report", "Summary"))
    step("excel_delete_sheet", lambda: excel_api.excel_delete_sheet("Summary"))
    step("excel_autofit", lambda: excel_api.excel_autofit(sheet=sheet))
    step("excel_recalculate", excel_api.excel_recalculate)
    step("excel_save", excel_api.excel_save)
    excel_pdf = step("excel_export_pdf", excel_api.excel_export_pdf)

    if written is not None:
        check(written["address"] == "A1:D3", f"unexpected write address {written['address']}")
    if values is not None:
        grid = values["values"]
        check(grid[0][0] == "Region", f"header mismatch: {grid[0]}")
        check(grid[1][1] == 10, f"quantity mismatch: {grid[1]}")
        check(abs(float(grid[1][3]) - 25.0) < 1e-9, f"formula result mismatch: {grid[1][3]}")
    if formatted is not None:
        check(formatted["values"][0][0] == "Region", "formatted read mismatch")
    if excel_pdf is not None:
        check(excel_pdf["file"]["size_bytes"] > 1000, "excel pdf is too small")

    print("== PowerPoint ==")
    ppt_file = str(TEST_DIR / "deck.pptx")
    step("ppt_create", lambda: ppt_api.ppt_create(ppt_file))
    step(
        "ppt_add_slide",
        lambda: ppt_api.ppt_add_slide(
            "Quarterly Report",
            ["Revenue up", "\tCosts flat", "  - Margin improved"],
            notes="Speaker notes for slide one",
        ),
    )
    step(
        "ppt_add_slide_title_only",
        lambda: ppt_api.ppt_add_slide("Details", layout="title_only"),
    )
    step("ppt_add_textbox", lambda: ppt_api.ppt_add_textbox(2, "Free text", font_size=18, color="#C00000"))
    step(
        "ppt_add_table",
        lambda: ppt_api.ppt_add_table(2, [["Metric", "Value"], ["Revenue", "1.2M"], ["Costs", "0.4M"]], top_cm=6),
    )
    step(
        "ppt_add_image",
        lambda: ppt_api.ppt_add_image(2, str(image), left_cm=18, top_cm=1, width_cm=2),
    )
    ppt_read = step("ppt_read", lambda: ppt_api.ppt_read(max_chars=5000))
    step("ppt_set_slide", lambda: ppt_api.ppt_set_slide(1, notes="Updated notes"))
    step("ppt_save", ppt_api.ppt_save)
    ppt_pdf = step("ppt_export_pdf", ppt_api.ppt_export_pdf)

    if ppt_read is not None:
        check(ppt_read["slide_count"] == 2, f"expected 2 slides, got {ppt_read['slide_count']}")
        first = ppt_read["slides"][0]
        check(first["title"] == "Quarterly Report", f"title mismatch: {first['title']}")
        check("Speaker notes" in first["notes"], f"notes mismatch: {first['notes']!r}")
        second_text = ppt_read["slides"][1]["text"]
        check("Free text" in second_text, f"textbox text missing: {second_text!r}")
        check("Revenue" in second_text, f"table text missing: {second_text!r}")
    if ppt_pdf is not None:
        check(ppt_pdf["file"]["size_bytes"] > 1000, "ppt pdf is too small")

    print("== Status ==")
    status = step("office_status", server.office_status)
    if status is not None:
        for key in ("word", "excel", "powerpoint"):
            check(status["applications"][key]["running"], f"{key} not running after tests")

    print("== Cleanup ==")
    for key in ("word", "excel", "powerpoint"):
        step(f"office_quit_{key}", lambda key=key: server.office_quit(key, force=True))

    print()
    if FAILURES:
        print(f"FAILED ({len(FAILURES)}): " + ", ".join(FAILURES))
        return 1
    print("ALL CHECKS PASSED")
    return 0


if __name__ == "__main__":
    code = main()
    try:
        from com import DISPATCHER

        DISPATCHER.shutdown()
    except Exception:
        pass
    sys.exit(code)
