# Office MCP Bridge

**Drive your installed Microsoft Word, Excel and PowerPoint from any MCP client — on real files, in real desktop windows.**

[![Python](https://img.shields.io/badge/python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D4?logo=windows&logoColor=white)](#requirements)
[![MCP](https://img.shields.io/badge/protocol-MCP-6E4AFF)](https://modelcontextprotocol.io/)
[![Tools](https://img.shields.io/badge/tools-53-brightgreen)](#tools)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![CI](https://github.com/disemay12/office-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/disemay12/office-mcp/actions/workflows/ci.yml)

[Русская версия](README.ru.md)

---

## Why

Most "office" MCP servers parse OpenXML files or re-implement a document model. This one is a thin, disciplined bridge to the **already installed desktop Office** over COM:

- works on the actual `.docx` / `.xlsx` / `.pptx` the user has open, with Office's own rendering, styles, charts and PDF export;
- no cloud, no API keys, no document upload — everything stays on the machine;
- one STA COM thread, serialized calls and a named mutex, so several agent sessions can share Office without corrupting each other;
- attaches to a running Office instance instead of starting a second one.

## Features

| | |
|---|---|
| **53 tools** | 19 Word, 17 Excel, 13 PowerPoint, 4 lifecycle |
| **Real files** | Opens, edits, saves, exports to PDF; never rewrites the file behind Office's back |
| **Live windows** | Edits appear in the visible app; `visible=false` hides Word/Excel |
| **Concurrency-safe** | Serialized COM dispatcher + cross-process mutex |
| **Zero config** | Sane defaults: outputs go to `Documents\Office MCP`, overridable by env var |
| **Honest errors** | Refuses to quit with unsaved work, reports busy sessions and stuck dialogs |

## Requirements

- **Windows 10/11** with desktop **Microsoft Word, Excel and PowerPoint** installed (Microsoft 365 or Office 2016+).
- **Python 3.10+** (tested on 3.14).
- Packages: `mcp>=1.10`, `pywin32>=306`.

> Word, Excel and PowerPoint are independent — install only what you need; tools for missing apps return a clear error.

## Installation

### 1. Clone

```powershell
git clone https://github.com/disemay12/office-mcp.git
cd office-mcp
```

### 2. Create an isolated environment

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Note the absolute path to the interpreter — every client config below needs it:

```powershell
(Resolve-Path .\.venv\Scripts\python.exe).Path
# C:\...\office-mcp\.venv\Scripts\python.exe
```

### 3. Register the server

The entry point is `office-mcp/server.py`, started over **stdio**:

```
<python.exe> -X utf8 <repo>\office-mcp\server.py
```

`-X utf8` matters: tool arguments and results are UTF-8, including non-ASCII paths.

<details open>
<summary><b>Claude Desktop</b> — <code>%APPDATA%\Claude\claude_desktop_config.json</code></summary>

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\path\\to\\office-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-X", "utf8", "C:\\path\\to\\office-mcp\\office-mcp\\server.py"]
    }
  }
}
```
</details>

<details>
<summary><b>Claude Code</b></summary>

```powershell
claude mcp add office -- "C:\path\to\office-mcp\.venv\Scripts\python.exe" -X utf8 "C:\path\to\office-mcp\\server.py"
```
</details>

<details>
<summary><b>OpenCode</b> — <code>opencode.jsonc</code></summary>

```jsonc
{
  "mcp": {
    "servers": {
      "office": {
        "type": "local",
        "command": [
          "C:/path/to/office-mcp/.venv/Scripts/python.exe",
          "-X", "utf8",
          "C:/path/to/office-mcp/server.py"
        ],
        "timeout": { "startup": 60000, "catalog": 120000, "execution": 300000 }
      }
    }
  }
}
```

Raise the timeouts: the first call in a session may have to cold-start Word.
</details>

<details>
<summary><b>Cursor / VS Code / any client with <code>mcp.json</code></b></summary>

Same `command` + `args` shape as Claude Desktop; see [`examples/mcp.json`](examples/mcp.json).
</details>

### 4. Verify

Restart the client, then ask:

> Call `office_status`.

A working install answers with the Office apps currently running, the base directory and the Python version. If the tools do not appear, see [Troubleshooting](#troubleshooting).

## Tools

### Common

| Tool | Purpose |
|---|---|
| `office_status()` | Running apps, open files, `base_dir`, `cwd`, Python version |
| `office_show(app)` | Focus the window (`word`, `excel`, `powerpoint`) |
| `office_quit(app, save=false, force=false)` | Close the app; refuses unsaved work without `save`/`force` |
| `office_open_file(path)` | Open any file with its Windows association (PDF, images) |

### Word — 19

| Tool | Purpose |
|---|---|
| `word_list_documents()` | Open documents, paths, saved state |
| `word_create(path=null, visible=true)` | New document |
| `word_open(path, visible=true, read_only=false)` | Open existing |
| `word_read(document=null, max_chars=20000)` | Document text |
| `word_outline(document=null, limit=300)` | Headings with levels |
| `word_append(text, document=null, style=null)` | Append paragraph |
| `word_prepend(text, document=null, style=null)` | Insert paragraph at the top |
| `word_set_text(text, document=null, style=null)` | Replace all content |
| `word_add_heading(text, level=1, document=null)` | `Heading 1..9` |
| `word_replace_text(find, replace, document=null, match_case=true, all_stories=false)` | Find & replace |
| `word_format_text(find, document=null, bold, italic, underline, font_size, color, highlight, match_case=true)` | Format by text; color `#RRGGBB` |
| `word_add_table(rows, document=null, header=true, autofit=true)` | Table from a 2-D array |
| `word_add_image(image_path, document=null, width_cm=null, caption=null)` | Insert image |
| `word_add_page_break(document=null)` | Page break |
| `word_set_header_footer(text, where="header", document=null, section=1, page_numbers=false)` | Header/footer and page numbers |
| `word_add_toc(document=null, levels=3, at_start=true)` | Table of contents |
| `word_save(document=null, path=null)` | Save, or Save As with `path` |
| `word_export_pdf(document=null, path=null)` | Save, then export PDF |
| `word_close(document, save=false)` | Close document |

### Excel — 17

| Tool | Purpose |
|---|---|
| `excel_list_workbooks()` | Open workbooks, sheets, used ranges |
| `excel_create(path=null, visible=true)` | New workbook |
| `excel_open(path, visible=true, read_only=false)` | Open workbook |
| `excel_list_sheets(workbook=null)` | Sheet names |
| `excel_read_range(range=null, sheet=null, workbook=null, formatted=false, include_formulas=false, max_rows=200, max_cols=50)` | Read a range; used range when omitted |
| `excel_write_range(values, range="A1", sheet=null, workbook=null, clear_first=false, autofit=false)` | Write a matrix; `=` starts a formula |
| `excel_format_range(range, sheet=null, workbook=null, bold, italic, font_size, number_format, bg_color, font_color, align, wrap, borders, column_width, row_height)` | Format a range; `#RRGGBB` colors |
| `excel_add_sheet(name, workbook=null, activate=true)` | Add sheet |
| `excel_delete_sheet(name, workbook=null)` | Delete sheet |
| `excel_rename_sheet(old_name, new_name, workbook=null)` | Rename sheet |
| `excel_add_chart(range, chart_type="column", title=null, sheet=null, workbook=null, anchor="A1", width_cm=16, height_cm=8)` | Chart: column, bar, line, line_markers, pie, doughnut, area, scatter, radar |
| `excel_add_table(range, name=null, sheet=null, workbook=null, style="TableStyleMedium2")` | Excel table |
| `excel_autofit(sheet=null, workbook=null)` | Autofit columns |
| `excel_recalculate(workbook=null)` | Recalculate formulas |
| `excel_save(workbook=null, path=null)` | Save, or Save As |
| `excel_export_pdf(workbook=null, path=null)` | Save, then export PDF |
| `excel_close(workbook, save=false)` | Close workbook |

### PowerPoint — 13

| Tool | Purpose |
|---|---|
| `ppt_list_presentations()` | Open presentations, slide counts |
| `ppt_create(path=null)` | New presentation |
| `ppt_open(path)` | Open existing |
| `ppt_add_slide(title=null, bullets=null, layout="title_content", notes=null, index=null, presentation=null)` | Add slide; layouts `title`, `title_content`, `two_content`, `comparison`, `title_only`, `blank` |
| `ppt_set_slide(slide_number, title=null, bullets=null, notes=null, append_bullets=false, presentation=null)` | Edit a slide |
| `ppt_add_textbox(slide_number, text, left_cm=1.5, top_cm=1.5, width_cm=20, height_cm=4, font_size=null, bold=null, color=null, presentation=null)` | Text box |
| `ppt_add_image(slide_number, image_path, left_cm=1.5, top_cm=1.5, width_cm=null, height_cm=null, presentation=null)` | Image |
| `ppt_add_table(slide_number, rows, left_cm=1.5, top_cm=3.5, width_cm=22, height_cm=9, header=true, presentation=null)` | Table |
| `ppt_read(presentation=null, max_chars=20000)` | Slide text and speaker notes |
| `ppt_delete_slide(slide_number, presentation=null)` | Delete slide |
| `ppt_save(presentation=null, path=null)` | Save, or Save As |
| `ppt_export_pdf(presentation=null, path=null)` | Save, then export PDF |
| `ppt_close(presentation, save=false)` | Close presentation |

In every tool the `document` / `workbook` / `presentation` argument may be omitted when exactly one file is open — the active one is used. A file name, a full path, or a string index is also accepted.

## Paths

- Relative paths resolve against `%USERPROFILE%\Documents\Office MCP`.
- Files created without a path (`word_create()`, `excel_create()`, `ppt_create()`) are saved there as `word-YYYYMMDD-HHMMSS.docx` and so on.
- Absolute paths and `~` are supported.
- Override the base directory with the `OFFICE_MCP_BASE_DIR` environment variable (set it **before** the client starts).
- `office_open_file` hands the file to the default Windows application (useful for PDFs).

## Behavior

- **Visibility.** Word and Excel show by default; `visible=false` hides the window while the app keeps running. PowerPoint cannot run hidden — its window is always visible.
- **Dialogs.** `DisplayAlerts` is off, so Office never asks for confirmation. A modal dialog opened by someone else (file locked, recovery, sign-in) makes the call fail after a 180 s timeout.
- **Busy.** Calls are serialized. A second session waits up to 60 s, then gets `Another Office MCP session is busy`.
- **Reuse.** The bridge attaches to a running instance via `GetActiveObject`, otherwise it starts one. `office_quit` terminates the app.
- **Macros.** Documents open in normal mode; macros are not forcibly disabled. Only use trusted files.
- **Saving.** PDF export saves the source file first. `office_quit` without `save`/`force` will not close an app holding unsaved work.

## Example prompts

- "Create a Word report with a TOC, a table and headers, save it to `C:\Temp\report.docx` and export it as PDF."
- "Open `sales.xlsx`, add a Report sheet with a regional summary and a chart, then save."
- "Build a 5-slide deck from this text, add speaker notes, export to PDF."
- "Read the open Word document and rewrite the *Summary* section shorter."

## Testing

```powershell
python -X utf8 tests\smoke_test.py       # 49 live operations across Word, Excel, PowerPoint
python -X utf8 tests\mcp_client_test.py  # stdio handshake, 53-tool catalog, real calls
```

Artifacts are written to `%TEMP%`-adjacent paths `C:\Temp\office-mcp-test` and `C:\Temp\office-mcp-protocol-test`. Both tests launch real Office applications and close them afterwards — save your own work first.

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `Office did not answer within 180s` | A modal Office dialog. Close it; if needed kill the process and retry. |
| `word is not running. Start it or create a document first.` | Expected for `office_status` / `office_show`. Create or open a file. |
| `File exists but is not open — call ..._open first` | Open the file with the matching `*_open` tool; a bare path is not enough. |
| `Another Office MCP session is busy` | Another chat holds the mutex for more than 60 s. Wait, or kill the stuck process. |
| `Cannot apply number format` | The Excel locale rejects the format. Pass the local variant, e.g. `# ##0,00`. |
| Tools do not appear | Check the Python and `server.py` paths, restart the client, read the MCP server's stderr. |
| Slow first call | Cold-starting Office; the example configs already raise the timeouts. |
| Leftover Office processes | `office_quit(app, force=true)`; after a client crash close them manually. |

## Security

The bridge runs as the current Windows user and holds full COM access to Office: it can read, modify, save and delete document content, and launch the applications. There is no sandbox. Do not attach it to untrusted agents, and keep it pointed at the directories you actually need. No network port is opened — transport is stdio only.

## License

[MIT](LICENSE)
