<div align="center">

# powerbi-mcp-local

**Local-first MCP server for Power BI Desktop automation**

Automate your semantic model, DAX, Power Query, Excel pipeline, and report layout from any MCP-capable AI client.

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue?logo=python&logoColor=white)](https://python.org)
[![Protocol MCP](https://img.shields.io/badge/protocol-MCP-blueviolet)](https://modelcontextprotocol.io)
[![License MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tools 56](https://img.shields.io/badge/tools-56-orange)](#tool-catalog-56-tools)

</div>

## What this gives you

- Connect AI tools directly to a running Power BI Desktop local engine.
- Ship model changes faster: tables, columns, measures, relationships.
- Run DAX and refresh from the same interface.
- Generate and patch Power Query (M) programmatically.
- Edit report pages and visuals through JSON + `pbi-tools`.

No Power BI Pro license is required for this local workflow.

## Who this is for

- Analytics engineers maintaining large Power BI models.
- BI developers who want repeatable model/report changes.
- Teams building AI-assisted BI workflows in editors and IDEs.
- Anyone who wants Power BI + Excel + MCP in one automation layer.

## Use cases

### 1) AI pair-programming for BI
Use Codex/Claude/Cursor prompts to list tables, add measures, fix DAX, and validate results in one loop.

### 2) Excel to Power BI pipelines
Create or update workbooks, generate M imports, refresh, then verify measures and row-level outputs.

### 3) Bulk report layout generation
Extract a `.pbix`, create pages and visuals (`card`, `bar`, `line`, `table`, `gauge`, `slicer`, `map`), apply theme, and compile back.

### 4) Safe automation in local environments
Enforce allowed directories, query guardrails, and readonly mode when needed.

## Architecture

```text
Any MCP Client  --(stdio or sse)-->  src/server.py
                                      |
                                      +-- TOM/.NET -> Power BI Desktop local SSAS
                                      +-- ADOMD    -> DAX query execution
                                      +-- openpyxl -> Excel read/write/format
                                      +-- pbi-tools-> report extract/compile + visuals
                                      +-- security -> path, query, and payload safeguards
```

## Requirements

| Requirement | Why it is needed | Install |
| --- | --- | --- |
| Windows | Power BI Desktop runs on Windows | - |
| Power BI Desktop | Local SSAS engine | `winget install Microsoft.PowerBIDesktop` |
| Python 3.11+ | Runtime | `winget install Python.Python.3.11` |
| .NET 6+ Runtime | Needed by `pythonnet` and `pbi-pyadomd` | `winget install Microsoft.DotNet.Runtime.6` |
| pbi-tools | Needed for report extract/compile + visual tools | `winget install pbi-tools` or `dotnet tool install -g pbi-tools` |

Notes:
- ADOMD.NET ships with Power BI Desktop.
- If `pbi-tools` is not on `PATH`, set `PBI_TOOLS_PATH`.

## 5-minute quick start

### 1) Install

```powershell
git clone https://github.com/StealthyLabsHQ/powerbi-mcp-local.git
cd powerbi-mcp-local
pip install -r requirements.txt
```

### 2) Open Power BI Desktop with a `.pbix`

Keep it running so the local Analysis Services instance is available.

### 3) Verify connectivity

```powershell
python tests/test_connection.py
```

Expected output includes:

```text
Connected to PBI Desktop on port XXXXX
Database: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
Tables: N
```

### 4) Start MCP server

```powershell
python src/server.py
```

Optional modes:

```powershell
# SSE transport
python src/server.py --transport sse --port 8765

# Readonly mode
python src/server.py --readonly
```

## MCP client setup

### Standard `stdio` config

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\src\\server.py"]
    }
  }
}
```

### SSE config

Run:

```powershell
python src/server.py --transport sse --port 8765
```

Then configure your client endpoint as:

```text
http://localhost:8765/sse
```

Guides:
- [docs/SETUP.md](docs/SETUP.md)
- [docs/WINDOWS_SETUP.md](docs/WINDOWS_SETUP.md)

## First prompts to try

- `Connect to Power BI and list all tables with columns.`
- `Create a measure called Total Sales in table Sales.`
- `Run this DAX query and show top 20 rows.`
- `Extract report, add a new page, place 3 visuals, then compile.`

## Tool catalog (56 tools)

### Core model discovery (6)

- `pbi_connect`
- `pbi_list_instances`
- `pbi_list_tables`
- `pbi_list_measures`
- `pbi_list_relationships`
- `pbi_model_info`

### Model mutations (6)

- `pbi_create_measure`
- `pbi_delete_measure`
- `pbi_set_format`
- `pbi_create_relationship`
- `pbi_create_table`
- `pbi_create_column`

### Query and import (4)

- `pbi_execute_dax`
- `pbi_refresh`
- `pbi_import_dax_file`
- `pbi_export_model`

### Power Query (M) tools (7)

- `pbi_get_power_query`
- `pbi_list_power_queries`
- `pbi_set_power_query`
- `pbi_create_import_query`
- `pbi_create_csv_import_query`
- `pbi_create_folder_import_query`
- `pbi_bulk_import_excel`

### Excel tools (13)

- `excel_list_sheets`
- `excel_read_sheet`
- `excel_read_cell`
- `excel_search`
- `excel_write_cell`
- `excel_write_range`
- `excel_create_sheet`
- `excel_delete_sheet`
- `excel_format_range`
- `excel_auto_width`
- `excel_create_workbook`
- `excel_workbook_info`
- `excel_to_pbi_check`

### Report and visual tools (20)

- `pbi_extract_report`
- `pbi_compile_report`
- `pbi_list_pages`
- `pbi_get_page`
- `pbi_create_page`
- `pbi_delete_page`
- `pbi_set_page_size`
- `pbi_add_card`
- `pbi_add_bar_chart`
- `pbi_add_line_chart`
- `pbi_add_donut_chart`
- `pbi_add_gauge`
- `pbi_add_table_visual`
- `pbi_add_waterfall`
- `pbi_add_slicer`
- `pbi_add_text_box`
- `pbi_remove_visual`
- `pbi_move_visual`
- `pbi_apply_theme`
- `pbi_build_dashboard`

## What is automated vs. what remains manual

| Automated via MCP | Requires manual action |
| --- | --- |
| Data source setup (Power Query M) | Custom visuals marketplace management |
| DAX measures — create, update, bulk import | Live visual preview during layout edits |
| Relationships, calculated tables, calculated columns | Report publishing to Power BI Service (requires Pro license) |
| Excel read, write, format, validate | |
| Visual formatting — colors, axes, fonts, borders (via extract + JSON + compile) | |
| Power Query expressions (M) | |
| Report extract / compile / page and visual CRUD | |
| Standard visuals — card, bar, line, donut, gauge, table, waterfall, slicer, text, map | |
| Drillthrough and bookmarks (via Layout JSON in extracted report) | |
| Theme application | |
| Model export to JSON | |

## End-to-end automation flow

```text
1) Build or update Excel input      -> excel_create_workbook / excel_write_range
2) Generate or patch M queries      -> pbi_create_import_query / pbi_bulk_import_excel
3) Model structure updates          -> pbi_create_relationship / pbi_create_column
4) Measures and formatting          -> pbi_create_measure / pbi_set_format
5) Validate                         -> pbi_refresh + pbi_execute_dax
6) Extract report                   -> pbi_extract_report
7) Create pages and visuals         -> pbi_create_page / pbi_build_dashboard / pbi_add_*
8) Apply theme                      -> pbi_apply_theme
9) Compile pbix                     -> pbi_compile_report
```

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| `No module named 'clr'` | Install .NET 6+ runtime, then restart terminal |
| `No running PBI Desktop instance found` | Open a `.pbix` in Power BI Desktop first |
| `pbi-tools not found` | Add to `PATH` or set `PBI_TOOLS_PATH` |
| `PermissionError` on `.xlsx` | Close Excel; workbook files are locked while open |
| Path blocked by policy | Configure `PBI_MCP_ALLOWED_DIRS` |

## FAQ

### Does this work without Power BI Pro?
Yes. This project targets local Power BI Desktop automation.

### Does it support Linux/macOS?
Not for full functionality. Power BI Desktop local engine is Windows-only.

### Do I need `pbi-tools` for every tool?
No. `pbi-tools` is required only for report extract/compile and visual-layout tooling.

### Can I run in readonly mode?
Yes. Start with `python src/server.py --readonly`.

### Is this safe to use on real files?
It includes security middleware (path restrictions, query guards, SSRF protections, audit logging). Review [SECURITY.md](SECURITY.md) before production use.

## Security

Security middleware includes:

- local path restrictions and traversal protection
- DAX/DMV injection and unsafe-query guards
- Power Query SSRF protections
- export redaction controls
- zip safety checks
- tool-call auditing

Details: [SECURITY.md](SECURITY.md)

## Development

Run tests:

```powershell
python -m unittest discover -s tests -p "test_*.py" -v
```

Notes:
- Some tests are local-only and do not require a live Power BI instance.
- Some integration scenarios require Power BI Desktop to be open.

## Repository layout

```text
powerbi-mcp-local/
|-- src/
|   |-- server.py
|   |-- pbi_connection.py
|   |-- security.py
|   `-- tools/
|       |-- model.py
|       |-- measures.py
|       |-- relationships.py
|       |-- query.py
|       |-- power_query.py
|       |-- excel.py
|       `-- visuals.py
|-- tests/
|-- docs/
|-- specs/
|-- SECURITY.md
|-- README.md
|-- pyproject.toml
`-- requirements.txt
```

## License

MIT
