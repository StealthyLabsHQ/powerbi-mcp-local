<div align="center">

# powerbi-mcp-local

<p>
  <img src="docs/assets/powerbi-logo.svg" alt="Power BI logo" width="84" />
</p>

**Local-first MCP server for Power BI Desktop automation**

Automate semantic model changes, DAX, Power Query, Excel, and report layout from MCP-capable AI clients.

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue?logo=python&logoColor=white)](https://python.org)
[![Protocol MCP](https://img.shields.io/badge/protocol-MCP-blueviolet)](https://modelcontextprotocol.io)
[![License MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tools 86](https://img.shields.io/badge/tools-86-orange)](#tool-catalog-en-86-tools)

</div>

## Quick Links

| Start | Setup | Demo | Tools | Security |
| --- | --- | --- | --- | --- |
| [5-minute quick start](#quick-start-en) | [MCP client setup](#mcp-client-setup-en) | [60-second demo](#demo-60s) | [Tool catalog](#tool-catalog-en-86-tools) | [Security](#security-en) |

<a id="demo-60s"></a>
## Demo (60s)

![powerbi-mcp-local demo](docs/assets/demo.gif)

```powershell
# 1) Start the MCP server
python src/server.py

# 2) In your MCP client, run prompts like:
"Connect to Power BI and list all tables with columns."
"Create a measure called Total Sales in table Sales."
"Run this DAX query and show top 20 rows."
```

Expected flow:
- MCP client calls `pbi_connect`
- server auto-discovers local Power BI Desktop SSAS port
- model/query/report tools become available

---

## English

### What this gives you

- Connect AI tools directly to a running local Power BI Desktop engine.
- Automate tables, columns, measures, and relationships.
- Execute DAX and refresh without leaving your MCP client.
- Generate and patch Power Query (M) programmatically.
- Edit report pages and visuals via JSON + `pbi-tools`.

No Power BI Pro license is required for this local workflow.

### Who this is for

- Analytics engineers maintaining large Power BI models.
- BI developers who want repeatable model/report changes.
- Teams building AI-assisted BI workflows in editors and IDEs.

### Architecture

```text
Any MCP Client  --(stdio or sse)-->  src/server.py
                                      |
                                      +-- TOM/.NET -> Power BI Desktop local SSAS
                                      +-- ADOMD    -> DAX query execution
                                      +-- openpyxl -> Excel read/write/format
                                      +-- pbi-tools-> report extract/compile + visuals
                                      +-- security -> path, query, and payload safeguards
```

### Requirements

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

<a id="quick-start-en"></a>
### 5-minute quick start (EN)

1. Install dependencies.

```powershell
git clone https://github.com/StealthyLabsHQ/powerbi-mcp-local.git
cd powerbi-mcp-local
pip install -r requirements.txt
```

2. Open Power BI Desktop with a `.pbix` file.

3. Verify connectivity.

```powershell
python tests/test_connection.py
```

4. Start server.

```powershell
python src/server.py
```

Optional:

```powershell
python src/server.py --transport sse --port 8765
python src/server.py --readonly
python src/server.py --profile readonly   # prune to read-only tools
python src/server.py --profile write      # read + write, no destructive
# SSE auth: set an env var before launch
$env:PBI_MCP_AUTH_TOKEN = "your-secret-token"
python src/server.py --transport sse --port 8765
# clients must send: Authorization: Bearer your-secret-token
```

<a id="mcp-client-setup-en"></a>
### MCP client setup (EN)

Standard `stdio` config:

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

SSE mode:

```powershell
python src/server.py --transport sse --port 8765
```

Endpoint:

```text
http://localhost:8765/sse
```

Guides:
- [docs/SETUP.md](docs/SETUP.md)
- [docs/WINDOWS_SETUP.md](docs/WINDOWS_SETUP.md)

### First prompts to try

- `Connect to Power BI and list all tables with columns.`
- `Create a measure called Total Sales in table Sales.`
- `Run this DAX query and show top 20 rows.`
- `Extract report, add a new page, place 3 visuals, then compile.`

<a id="tool-catalog-en-86-tools"></a>
### Tool catalog (EN, 86 tools)

Core model discovery (7):
- `pbi_connect`
- `pbi_list_instances`
- `pbi_list_tables`
- `pbi_list_measures`
- `pbi_list_relationships`
- `pbi_model_info`
- `pbi_refresh_metadata`

Model mutations (14):
- `pbi_create_measure`
- `pbi_delete_measure`
- `pbi_rename_measure`
- `pbi_set_format`
- `pbi_create_relationship`
- `pbi_update_relationship`
- `pbi_delete_relationship`
- `pbi_create_table`
- `pbi_delete_table`
- `pbi_rename_table`
- `pbi_create_column`
- `pbi_delete_column`
- `pbi_rename_column`
- `pbi_execute_dax_as_role`

Query and import (6):
- `pbi_execute_dax`
- `pbi_trace_query`
- `pbi_validate_dax`
- `pbi_measure_dependencies`
- `pbi_refresh`
- `pbi_import_dax_file`
- `pbi_export_model`

Power Query (M) tools (8):
- `pbi_get_power_query`
- `pbi_list_power_queries`
- `pbi_set_power_query`
- `pbi_create_import_query`
- `pbi_create_csv_import_query`
- `pbi_create_folder_import_query`
- `pbi_bulk_import_excel`
- `pbi_import_excel_workbook`

Workflow tools (3):
- `pbi_model_audit_workflow`
- `pbi_excel_import_workflow`
- `pbi_measure_workflow`

Row-level security (6):
- `pbi_list_roles`
- `pbi_create_role`
- `pbi_delete_role`
- `pbi_set_role_filter`
- `pbi_add_role_member`
- `pbi_remove_role_member`

Calculation groups (3):
- `pbi_list_calc_groups`
- `pbi_create_calc_group`
- `pbi_delete_calc_group`

Unified visual dispatcher (1):
- `pbi_add_visual(visual_type, config)` — dispatches to the 9 legacy `pbi_add_*` tools (kept as shims)

Excel tools (13):
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

Report and visual tools (24):
- `pbi_extract_report`
- `pbi_compile_report`
- `pbi_patch_layout`
- `pbi_list_pages`
- `pbi_validate_report_fields`
- `pbi_repair_report_fields`
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
- `pbi_apply_design`
- `pbi_apply_theme`
- `pbi_build_dashboard`

### End-to-end automation flow

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

### Troubleshooting (EN)

| Symptom | Fix |
| --- | --- |
| `No module named 'clr'` | Install .NET 6+ runtime, then restart terminal |
| `No running PBI Desktop instance found` | Open a `.pbix` in Power BI Desktop first |
| `pbi-tools not found` | Add to `PATH` or set `PBI_TOOLS_PATH` |
| `PermissionError` on `.xlsx` | Close Excel; workbook files are locked while open |
| Path blocked by policy | Configure `PBI_MCP_ALLOWED_DIRS` |

### FAQ (EN)

- Does this work without Power BI Pro? Yes, local Power BI Desktop workflow.
- Linux/macOS support? Not for full functionality. Power BI Desktop local engine is Windows-only.
- Is `pbi-tools` needed for all tools? No, only report extract/compile and visual-layout tooling.
- Can I run readonly? Yes, use `python src/server.py --readonly`.

<a id="security-en"></a>
### Security (EN)

Security middleware includes:
- local path restrictions and traversal protection
- DAX/DMV injection and unsafe-query guards
- Power Query SSRF protections
- export redaction controls
- zip safety checks
- tool-call auditing

Details: [SECURITY.md](SECURITY.md)

### Development

```powershell
pip install -e ".[dev]"
pytest -v
```

---

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
