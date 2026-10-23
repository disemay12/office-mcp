# Office MCP Bridge

**Управление установленными Microsoft Word, Excel и PowerPoint из любого MCP-клиента — на реальных файлах, в реальных окнах.**

[![Python](https://img.shields.io/badge/python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows-0078D4?logo=windows&logoColor=white)](#требования)
[![MCP](https://img.shields.io/badge/protocol-MCP-6E4AFF)](https://modelcontextprotocol.io/)
[![Tools](https://img.shields.io/badge/tools-53-brightgreen)](#инструменты)
[![License](https://img.shields.io/badge/license-MIT-blue)](LICENSE)
[![CI](https://github.com/disemay12/office-mcp/actions/workflows/ci.yml/badge.svg)](https://github.com/disemay12/office-mcp/actions/workflows/ci.yml)

[English version](README.md)

---

## Зачем

Большинство «офисных» MCP-серверов парсят OpenXML или заново реализуют модель документа. Этот — тонкий и аккуратный мост к **уже установленному настольному Office** через COM:

- работает с настоящими `.docx` / `.xlsx` / `.pptx`, которые открыты у пользователя: родной рендеринг, стили, диаграммы и экспорт в PDF;
- без облака, ключей и выгрузки документов — всё остаётся на машине;
- один STA-поток COM, сериализованные вызовы и именованный мьютекс, поэтому несколько агентских сессий не портят работу друг друга;
- подключается к уже запущенному Office, а не поднимает второй экземпляр.

## Возможности

| | |
|---|---|
| **53 инструмента** | 19 Word, 17 Excel, 13 PowerPoint, 4 служебных |
| **Реальные файлы** | Открытие, правка, сохранение, экспорт в PDF; файл не переписывается в обход Office |
| **Живые окна** | Правки видны в приложении; `visible=false` скрывает Word/Excel |
| **Безопасная конкурентность** | Сериализованный COM-диспетчер + межпроцессный мьютекс |
| **Без настройки** | Результаты по умолчанию идут в `Documents\Office MCP`, база переопределяется переменной окружения |
| **Честные ошибки** | Не закрывает приложение с несохранённой работой, сообщает о занятой сессии и зависших диалогах |

## Требования

- **Windows 10/11** с настольным **Microsoft Word, Excel и PowerPoint** (Microsoft 365 или Office 2016+).
- **Python 3.10+** (проверено на 3.14).
- Пакеты: `mcp>=1.10,<2` (в SDK 2.x `FastMCP` переименован, мост рассчитан на 1.x), `pywin32>=306`.

> Приложения независимы: ставьте только нужные. Инструменты отсутствующего приложения вернут понятную ошибку.

## Установка

### 1. Клонирование

```powershell
git clone https://github.com/disemay12/office-mcp.git
cd office-mcp
```

### 2. Изолированное окружение

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Запомните абсолютный путь к интерпретатору — он нужен во всех конфигах ниже:

```powershell
(Resolve-Path .\.venv\Scripts\python.exe).Path
# C:\...\office-mcp\.venv\Scripts\python.exe
```

### 3. Регистрация сервера

Точка входа — `server.py`, транспорт **stdio**:

```
<python.exe> -X utf8 <repo>\server.py
```

`-X utf8` обязателен: аргументы и результаты инструментов — UTF-8, включая не-ASCII пути.

<details open>
<summary><b>Claude Desktop</b> — <code>%APPDATA%\Claude\claude_desktop_config.json</code></summary>

```json
{
  "mcpServers": {
    "office": {
      "command": "C:\\path\\to\\office-mcp\\.venv\\Scripts\\python.exe",
      "args": ["-X", "utf8", "C:\\path\\to\\office-mcp\\server.py"]
    }
  }
}
```
</details>

<details>
<summary><b>Claude Code</b></summary>

```powershell
claude mcp add office -- "C:\path\to\office-mcp\.venv\Scripts\python.exe" -X utf8 "C:\path\to\office-mcp\server.py"
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

Таймауты повышены не случайно: первый вызов в сессии может холодно стартовать Word.
</details>

<details>
<summary><b>Cursor / VS Code / любой клиент с <code>mcp.json</code></b></summary>

Та же связка `command` + `args`, что у Claude Desktop: [`examples/mcp.json`](examples/mcp.json).
</details>

### 4. Проверка

Перезапустите клиент и попросите:

> Вызови `office_status`.

Работающая установка вернёт запущенные приложения Office, базовый каталог и версию Python. Если инструменты не появились — см. [Диагностика](#диагностика).

## Инструменты

### Общие

| Инструмент | Назначение |
|---|---|
| `office_status()` | Запущенные приложения, открытые файлы, `base_dir`, `cwd`, версия Python |
| `office_show(app)` | Показать/активировать окно (`word`, `excel`, `powerpoint`) |
| `office_quit(app, save=false, force=false)` | Закрыть приложение; без `save`/`force` откажется при несохранённых файлах |
| `office_open_file(path)` | Открыть файл ассоциацией Windows (PDF, изображения) |

### Word — 19

| Инструмент | Назначение |
|---|---|
| `word_list_documents()` | Открытые документы, пути, состояние сохранения |
| `word_create(path=null, visible=true)` | Новый документ |
| `word_open(path, visible=true, read_only=false)` | Открыть существующий |
| `word_read(document=null, max_chars=20000)` | Текст документа |
| `word_outline(document=null, limit=300)` | Заголовки с уровнями |
| `word_append(text, document=null, style=null)` | Добавить абзац в конец |
| `word_prepend(text, document=null, style=null)` | Вставить абзац в начало |
| `word_set_text(text, document=null, style=null)` | Заменить всё содержимое |
| `word_add_heading(text, level=1, document=null)` | Заголовок (`Heading 1..9`) |
| `word_replace_text(find, replace, document=null, match_case=true, all_stories=false)` | Найти и заменить |
| `word_format_text(find, document=null, bold, italic, underline, font_size, color, highlight, match_case=true)` | Форматирование по тексту; цвет `#RRGGBB` |
| `word_add_table(rows, document=null, header=true, autofit=true)` | Таблица из двумерного массива |
| `word_add_image(image_path, document=null, width_cm=null, caption=null)` | Вставить изображение |
| `word_add_page_break(document=null)` | Разрыв страницы |
| `word_set_header_footer(text, where="header", document=null, section=1, page_numbers=false)` | Колонтитул и/или номера страниц |
| `word_add_toc(document=null, levels=3, at_start=true)` | Оглавление |
| `word_save(document=null, path=null)` | Сохранить, при `path` — Save As |
| `word_export_pdf(document=null, path=null)` | Сохранить и экспортировать PDF |
| `word_close(document, save=false)` | Закрыть документ |

### Excel — 17

| Инструмент | Назначение |
|---|---|
| `excel_list_workbooks()` | Открытые книги, листы, used range |
| `excel_create(path=null, visible=true)` | Новая книга |
| `excel_open(path, visible=true, read_only=false)` | Открыть книгу |
| `excel_list_sheets(workbook=null)` | Список листов |
| `excel_read_range(range=null, sheet=null, workbook=null, formatted=false, include_formulas=false, max_rows=200, max_cols=50)` | Чтение диапазона; без `range` — used range |
| `excel_write_range(values, range="A1", sheet=null, workbook=null, clear_first=false, autofit=false)` | Запись матрицы; `=` начинает формулу |
| `excel_format_range(range, sheet=null, workbook=null, bold, italic, font_size, number_format, bg_color, font_color, align, wrap, borders, column_width, row_height)` | Формат диапазона; цвета `#RRGGBB` |
| `excel_add_sheet(name, workbook=null, activate=true)` | Добавить лист |
| `excel_delete_sheet(name, workbook=null)` | Удалить лист |
| `excel_rename_sheet(old_name, new_name, workbook=null)` | Переименовать лист |
| `excel_add_chart(range, chart_type="column", title=null, sheet=null, workbook=null, anchor="A1", width_cm=16, height_cm=8)` | Диаграмма: column, bar, line, line_markers, pie, doughnut, area, scatter, radar |
| `excel_add_table(range, name=null, sheet=null, workbook=null, style="TableStyleMedium2")` | Умная таблица |
| `excel_autofit(sheet=null, workbook=null)` | Автоширина столбцов |
| `excel_recalculate(workbook=null)` | Пересчёт формул |
| `excel_save(workbook=null, path=null)` | Сохранить / Save As |
| `excel_export_pdf(workbook=null, path=null)` | Сохранить и экспортировать PDF |
| `excel_close(workbook, save=false)` | Закрыть книгу |

### PowerPoint — 13

| Инструмент | Назначение |
|---|---|
| `ppt_list_presentations()` | Открытые презентации, число слайдов |
| `ppt_create(path=null)` | Новая презентация |
| `ppt_open(path)` | Открыть существующую |
| `ppt_add_slide(title=null, bullets=null, layout="title_content", notes=null, index=null, presentation=null)` | Добавить слайд; layout: `title`, `title_content`, `two_content`, `comparison`, `title_only`, `blank` |
| `ppt_set_slide(slide_number, title=null, bullets=null, notes=null, append_bullets=false, presentation=null)` | Изменить слайд |
| `ppt_add_textbox(slide_number, text, left_cm=1.5, top_cm=1.5, width_cm=20, height_cm=4, font_size=null, bold=null, color=null, presentation=null)` | Надпись |
| `ppt_add_image(slide_number, image_path, left_cm=1.5, top_cm=1.5, width_cm=null, height_cm=null, presentation=null)` | Изображение |
| `ppt_add_table(slide_number, rows, left_cm=1.5, top_cm=3.5, width_cm=22, height_cm=9, header=true, presentation=null)` | Таблица |
| `ppt_read(presentation=null, max_chars=20000)` | Текст слайдов и заметки |
| `ppt_delete_slide(slide_number, presentation=null)` | Удалить слайд |
| `ppt_save(presentation=null, path=null)` | Сохранить / Save As |
| `ppt_export_pdf(presentation=null, path=null)` | Сохранить и экспортировать PDF |
| `ppt_close(presentation, save=false)` | Закрыть презентацию |

Во всех инструментах `document` / `workbook` / `presentation` можно опустить, если открыт ровно один файл: берётся активный. Также принимается имя файла, полный путь или индекс (строкой).

## Пути

- Относительный путь разрешается от `%USERPROFILE%\Documents\Office MCP`.
- Файл, созданный без пути (`word_create()`, `excel_create()`, `ppt_create()`), сохраняется туда как `word-ГГГГММДД-ЧЧММСС.docx` и т.п.
- Абсолютные пути и `~` поддерживаются.
- База переопределяется переменной окружения `OFFICE_MCP_BASE_DIR` (задаётся **до** запуска клиента).
- `office_open_file` открывает файл приложением Windows по умолчанию (например, PDF).

## Поведение

- **Видимость.** Word и Excel по умолчанию показываются; `visible=false` скрывает окно, приложение продолжает работать. PowerPoint скрыть нельзя — окно всегда видно.
- **Диалоги.** `DisplayAlerts` отключён, Office не спрашивает подтверждений. Если открыт чужой модальный диалог (файл занят, восстановление, вход в аккаунт), вызов завершится ошибкой по таймауту 180 с.
- **Занятость.** Вызовы сериализованы. Параллельная сессия ждёт до 60 с, затем получает `Another Office MCP session is busy`.
- **Переиспользование.** Мост ищет запущенный экземпляр через `GetActiveObject`, иначе запускает новый. `office_quit` завершает приложение.
- **Макросы.** Документы открываются в обычном режиме; макросы принудительно не отключаются. Используйте только доверенные файлы.
- **Сохранение.** Экспорт в PDF сначала сохраняет исходный файл. `office_quit` без `save`/`force` не закрывает приложение при несохранённых данных.

## Типовые запросы

- «Создай в Word отчёт с оглавлением, таблицей и колонтитулами, сохрани в `C:\Temp\report.docx` и сделай PDF».
- «Открой `sales.xlsx`, добавь лист Report со сводкой по регионам и диаграммой, сохрани».
- «Сделай презентацию на 5 слайдов по этому тексту, с заметками, и экспортируй в PDF».
- «Прочитай открытый документ Word и перепиши раздел "Итоги" короче».

## Тесты

```powershell
python -X utf8 tests\smoke_test.py       # 49 живых операций в Word, Excel, PowerPoint
python -X utf8 tests\mcp_client_test.py  # stdio-хендшейк, каталог из 53 инструментов, реальные вызовы
```

Артефакты пишутся в `C:\Temp\office-mcp-test` и `C:\Temp\office-mcp-protocol-test`. Оба теста запускают настоящие приложения Office и закрывают их — сохраните свою работу заранее.

## Диагностика

| Симптом | Причина / решение |
|---|---|
| `Office did not answer within 180s` | Модальное окно Office. Закройте его; при необходимости завершите процесс и повторите. |
| `word is not running. Start it or create a document first.` | Ожидаемо для `office_status` / `office_show`. Создайте или откройте файл. |
| `File exists but is not open — call ..._open first` | Файл нужно открыть инструментом `*_open`, а не только указать путь. |
| `Another Office MCP session is busy` | Другой чат держит мьютекс дольше 60 с. Дождитесь или завершите зависший процесс. |
| `Cannot apply number format` | Локаль Excel не принимает формат. Передайте локальный вариант, например `# ##0,00`. |
| Инструменты не появились | Проверьте путь к Python и `server.py`, перезапустите клиент, посмотрите stderr MCP-сервера. |
| Долгий первый вызов | Холодный старт Office; таймауты в примерах конфигов уже увеличены. |
| Остались процессы Office | `office_quit(app, force=true)`; при аварийном завершении клиента закройте процессы вручную. |

## Безопасность

Мост выполняется под текущей учётной записью Windows и имеет полный COM-доступ к Office: может читать, изменять, сохранять и удалять содержимое документов, а также запускать приложения. Изоляции нет. Не подключайте сервер к недоверенным агентам и не разрешайте ему работать с файлами вне нужных каталогов. Сетевых портов сервер не открывает — только stdio.

## Лицензия

[MIT](LICENSE)
