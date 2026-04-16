# powerbi-mcp-local

MCP server that connects to Power BI Desktop's local Analysis Services instance.
Lets Claude Code (or any MCP client) read, write, and query your Power BI data model programmatically.

## Tools (36)

### Power BI Core

| Tool | Description |
|---|---|
| `pbi_connect` | Auto-discover and connect to the running PBI Desktop instance |
| `pbi_list_instances` | List discovered local Power BI Desktop instances and ports |
| `pbi_list_tables` | List all tables with columns and data types |
| `pbi_list_measures` | List all DAX measures |
| `pbi_list_relationships` | List all relationships |
| `pbi_model_info` | Full model snapshot (tables + measures + relationships) |

### Mutations

| Tool | Description |
|---|---|
| `pbi_create_measure` | Create or update a DAX measure |
| `pbi_delete_measure` | Delete a measure |
| `pbi_set_format` | Batch-format measures (number format strings) |
| `pbi_create_relationship` | Create a relationship between two tables |
| `pbi_create_table` | Create a calculated table (DAX expression) |
| `pbi_create_column` | Create a calculated column in a table |

### Query & Import

| Tool | Description |
|---|---|
| `pbi_execute_dax` | Run any DAX query and get results as JSON |
| `pbi_refresh` | Trigger a data model refresh |
| `pbi_import_dax_file` | Bulk-import measures from a `.dax` file |
| `pbi_export_model` | Export full model as JSON for version control |

### Power Query Tools

| Tool | Description |
|---|---|
| `pbi_get_power_query` | Read the M expression for a specific table partition |
| `pbi_list_power_queries` | List tables, partitions, source types, and current M expressions |
| `pbi_set_power_query` | Inject or replace a validated M expression on a partition |
| `pbi_create_import_query` | Build an Excel-sheet import query and inject it into a partition |
| `pbi_create_csv_import_query` | Build a CSV import query and inject it into a partition |
| `pbi_create_folder_import_query` | Build a folder import query and inject it into a partition |
| `pbi_bulk_import_excel` | Auto-map workbook sheets to model tables and inject queries in bulk |

### Excel Tools

| Tool | Description |
|---|---|
| `excel_list_sheets` | List workbook sheets with row and column counts |
| `excel_read_sheet` | Read rows from a worksheet or range |
| `excel_read_cell` | Read a single cell with cached value, type, format, and formula |
| `excel_search` | Search workbook values across one or all sheets |
| `excel_write_cell` | Write a single cell value with optional number format |
| `excel_write_range` | Write a 2D array into a worksheet |
| `excel_create_sheet` | Create a worksheet in an existing workbook |
| `excel_delete_sheet` | Delete a worksheet from an existing workbook |
| `excel_format_range` | Apply formatting to a worksheet range |
| `excel_auto_width` | Auto-fit worksheet column widths |
| `excel_create_workbook` | Create a new `.xlsx` workbook with optional sheets |
| `excel_workbook_info` | Return workbook metadata, named ranges, and dimensions |
| `excel_to_pbi_check` | Compare Excel sheets and headers against the current PBI model |

## How it works

```
Claude Code ──(MCP/stdio)──> server.py
                                ├──(TOM/.NET)──> PBI Desktop (local SSAS)
                                ├──(openpyxl)──> Excel files (.xlsx)
                                └──(Power Query M)──> workbook / CSV / folder sources
```

Power BI Desktop runs a local Analysis Services engine on a random port.
The server finds the port automatically via:
1. `%LOCALAPPDATA%` workspace scan (MSI install)
2. User-scoped path scan (Microsoft Store install)
3. `msmdsrv.exe` process fallback

Connection is managed by `PowerBIConnectionManager` — thread-safe, auto-reconnects on port change, hard-resets on write failure. Excel operations use `openpyxl` with structured JSON errors, large-file streaming for workbooks over 10 MB, and graceful handling when Excel has a workbook locked for writing. Power Query operations write validated M expressions onto TOM partitions and handle partition source types explicitly (`m`, `query`, `calculated`, `none`).

## Requirements

- **Windows** (Power BI Desktop only runs on Windows)
- **Power BI Desktop** installed and open with a `.pbix` file
- **Python 3.11+**
- **Claude Code** or any MCP-compatible client

## Quick Start

```powershell
# 1. Clone
git clone https://github.com/StealthyLabsHQ/powerbi-mcp-local.git
cd powerbi-mcp-local

# 2. Install deps
pip install -r requirements.txt

# 3. Open Power BI Desktop with any .pbix file

# 4. Test connection
python test_connection.py

# 5. Run with Claude Code
claude
```

Then tell Claude Code:

> "Connect to Power BI and inject the relationships and measures from my .dax file."

## Full Automation Workflow

1. Create or inspect the source file with the Excel tools, or point Power BI to an existing workbook, CSV, or folder with the Power Query tools.
2. Inject the import query into the right table partition with `pbi_create_import_query`, `pbi_create_csv_import_query`, `pbi_create_folder_import_query`, or `pbi_bulk_import_excel`.
3. Create relationships with `pbi_create_relationship`.
4. Add measures individually with `pbi_create_measure` or in bulk with `pbi_import_dax_file`.
5. Trigger `pbi_refresh` and validate the model with `pbi_model_info` or `pbi_execute_dax`.

Example end-to-end sequence:

```text
excel_create_workbook -> excel_write_range -> pbi_create_import_query
-> pbi_create_relationship -> pbi_import_dax_file -> pbi_refresh
-> pbi_execute_dax
```

## Compatible Tools

Works with any MCP-compatible AI tool or IDE:

| Tool | Transport | Config |
|---|---|---|
| **Claude Code** (CLI) | stdio | `.claude/settings.json` |
| **Claude Desktop** | stdio | `%APPDATA%\Claude\claude_desktop_config.json` |
| **Codex CLI** (OpenAI) | stdio | `~/.codex/config.json` |
| **Gemini CLI** (Google) | stdio | `~/.gemini/settings.json` |
| **Cursor** | stdio or sse | `.cursor/mcp.json` |
| **VS Code** (Continue/Cline) | stdio or sse | `.vscode/mcp.json` |
| **JetBrains** (IntelliJ, PyCharm) | stdio or sse | Settings > AI Assistant > MCP |
| **Windsurf / Cline** | stdio | `.mcp/config.json` |

See **[docs/SETUP.md](docs/SETUP.md)** for detailed config per platform.

### Quick Config (Claude Code)

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

### SSE Mode (for IDEs)

```powershell
python server.py --transport sse --port 8765
```

Then point your IDE to `http://localhost:8765/sse`.

## Project Structure

```
powerbi-mcp-local/
├── server.py              Main MCP server (FastMCP, 36 tools)
├── pbi_connection.py      Connection manager (port discovery, TOM, ADOMD)
├── pyproject.toml         Packaging metadata for PyPI / pip installs
├── tools/
│   ├── __init__.py        Tool exports
│   ├── excel.py           Excel read/write/format/pipeline tools
│   ├── model.py           Tables, columns, export, model info
│   ├── measures.py        Measures CRUD, formatting, .dax bulk import
│   ├── power_query.py     Power Query (M) import and partition tools
│   ├── relationships.py   Relationships CRUD
│   └── query.py           DAX execution, data refresh
├── test_connection.py     Standalone connectivity test
├── test_excel.py          Standalone Excel tool test suite
├── test_power_query.py    Standalone Power Query tool test suite
├── requirements.txt       Pinned dependencies
├── CLAUDE.md              Build instructions for Claude Code
├── EXCEL_SPEC.md          Excel extension spec
└── README.md
```

## What it automates vs. what stays manual

| Automated via MCP | Manual in PBI Desktop |
|---|---|
| DAX measures (create, update, delete, bulk import) | Data import (Excel, CSV) |
| Excel workbook reads, writes, formatting, validation | Pivot table refresh in the Excel UI |
| Power Query source setup for workbook, CSV, and folder imports | Visual Power Query editing and step-by-step preview |
| Relationships between tables | Theme import (JSON) |
| Calculated tables and columns | Visual creation (charts, cards) |
| DAX query execution and refresh | Page layout and formatting |
| Model export to JSON | Report publishing |

The MCP handles the **data model layer**. The **visual layer** (charts, layouts, pages) must be built manually in Power BI Desktop — there is no API for it.

## Tech Stack

| Package | Version | Role |
|---|---|---|
| [`mcp[cli]`](https://pypi.org/project/mcp/) | 1.27.0 | Anthropic MCP SDK |
| [`openpyxl`](https://pypi.org/project/openpyxl/) | 3.1.5 | Cross-platform Excel workbook backend |
| [`pbi-pyadomd`](https://pypi.org/project/pbi-pyadomd/) | 1.4.3 | ADOMD.NET wrapper (queries) |
| [`pythonnet`](https://pypi.org/project/pythonnet/) | 3.0.5 | .NET bridge (TOM model writes) |
| [`psutil`](https://pypi.org/project/psutil/) | 7.2.2 | Process port discovery fallback |

## License

MIT
