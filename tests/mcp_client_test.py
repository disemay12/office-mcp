"""MCP protocol test: starts server.py over stdio and drives it as a real client.

Run:  python tests\\mcp_client_test.py
"""

from __future__ import annotations

import asyncio
import json
import os
import shutil
import sys
from pathlib import Path

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client

TEST_DIR = Path(r"C:\Temp\office-mcp-protocol-test")
SERVER = Path(__file__).resolve().parent.parent / "server.py"

EXPECTED_TOOLS = {
    "word_create",
    "word_append",
    "word_read",
    "word_add_table",
    "word_export_pdf",
    "excel_create",
    "excel_write_range",
    "excel_read_range",
    "excel_add_chart",
    "ppt_create",
    "ppt_add_slide",
    "ppt_read",
    "office_status",
    "office_quit",
}


def payload(result) -> dict:
    if result.structuredContent:
        return result.structuredContent
    for item in result.content:
        text = getattr(item, "text", None)
        if text:
            return json.loads(text)
    raise AssertionError("tool returned no JSON payload")


async def main() -> int:
    if TEST_DIR.exists():
        shutil.rmtree(TEST_DIR)
    TEST_DIR.mkdir(parents=True, exist_ok=True)

    env = dict(os.environ)
    env["OFFICE_MCP_BASE_DIR"] = str(TEST_DIR)

    params = StdioServerParameters(
        command=sys.executable,
        args=["-X", "utf8", str(SERVER)],
        env=env,
    )

    async with stdio_client(params) as (read, write):
        async with ClientSession(read, write) as session:
            init = await session.initialize()
            print("server:", init.serverInfo.name, init.serverInfo.version)
            print("instructions:", (init.instructions or "")[:100].replace("\n", " "))

            tools = await session.list_tools()
            names = {tool.name for tool in tools.tools}
            print(f"tools: {len(names)}")
            missing = EXPECTED_TOOLS - names
            assert not missing, f"missing tools: {sorted(missing)}"

            append_schema = next(t for t in tools.tools if t.name == "word_append")
            assert "text" in append_schema.inputSchema["properties"]
            assert append_schema.description

            status = payload(await session.call_tool("office_status", {}))
            assert "applications" in status
            print("status ok; word running:", status["applications"]["word"]["running"])

            created = payload(
                await session.call_tool(
                    "word_create", {"path": str(TEST_DIR / "mcp.docx"), "visible": False}
                )
            )
            assert created["name"] == "mcp.docx"

            await session.call_tool(
                "word_append", {"text": "Hello from the MCP client", "style": "Heading 1"}
            )
            read_back = payload(await session.call_tool("word_read", {}))
            assert "Hello from the MCP client" in read_back["text"], read_back
            saved = payload(await session.call_tool("word_save", {}))
            assert Path(saved["path"]).exists()
            pdf = payload(await session.call_tool("word_export_pdf", {}))
            assert Path(pdf["pdf"]).stat().st_size > 1000

            workbook = payload(
                await session.call_tool(
                    "excel_create", {"path": str(TEST_DIR / "mcp.xlsx"), "visible": False}
                )
            )
            sheet = workbook["sheets"][0]
            await session.call_tool(
                "excel_write_range",
                {"values": [["A", "B"], [1, "=A2*2"]], "range": "A1", "sheet": sheet},
            )
            grid = payload(
                await session.call_tool(
                    "excel_read_range", {"range": "A1:B2", "sheet": sheet}
                )
            )
            assert grid["values"][1][1] == 2, grid
            await session.call_tool("excel_save", {})

            deck = payload(
                await session.call_tool(
                    "ppt_create", {"path": str(TEST_DIR / "mcp.pptx")}
                )
            )
            await session.call_tool(
                "ppt_add_slide",
                {"title": "Protocol", "bullets": ["one", "two"], "notes": "n"},
            )
            slides = payload(await session.call_tool("ppt_read", {}))
            assert slides["slide_count"] == 1
            await session.call_tool("ppt_save", {})

            error = await session.call_tool("word_open", {"path": str(TEST_DIR / "nope.docx")})
            assert error.isError, "expected an error for a missing file"
            error_text = " ".join(
                getattr(item, "text", "") for item in error.content
            ).strip()
            assert "not found" in error_text.lower(), error_text
            print("error handling ok:", error_text[:100])

            for key in ("word", "excel", "powerpoint"):
                print(
                    "quit",
                    key,
                    payload(await session.call_tool("office_quit", {"app": key, "force": True})),
                )

    print("MCP PROTOCOL TEST PASSED")
    return 0


if __name__ == "__main__":
    raise SystemExit(asyncio.run(main()))
