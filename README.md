<div align="center">

# powerbi-mcp-local

**MCP server for Power BI Desktop + Excel**

Connect any AI coding tool to Power BI Desktop's local engine.
Read, write, and automate your data model — no Pro license required.

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue?logo=python&logoColor=white)](https://python.org)
[![MCP](https://img.shields.io/badge/protocol-MCP-blueviolet)](https://modelcontextprotocol.io)
[![License: MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tools](https://img.shields.io/badge/tools-56-orange)](#tools-56)

</div>

---

## How it works

```
Any AI Tool ──(MCP)──> server.py
                         ├──(TOM/.NET)──────> Power BI Desktop (local SSAS)
                         ├──(Power Query M)──> workbook / CSV / folder sources
                         ├──(openpyxl)───────> Excel files (.xlsx)
                         └──(pbi-tools)──────> Report pages, visuals, and themes
```

Power BI Desktop runs a local Analysis Services engine on a random port.
This server finds the port automatically (MSI, Microsoft Store, or process scan)
and exposes the full data model, file pipeline, and report layer through 56 MCP tools.

---

## Quick Start

```powershell
git clone https://github.com/StealthyLabsHQ/powerbi-mcp-local.git
cd powerbi-mcp-local
pip install -r requirements.txt

# Open Power BI Desktop with any .pbix, then:
python test_connection.py
```

> **Tell your AI tool:** *"Connect to Power BI and list all tables"*

---

## Compatible Tools

Works with **any** MCP-compatible AI tool or IDE:

| Platform | Transport | Config file |
|:---|:---:|:---|
| **Claude Code** | stdio | `.claude/settings.json` |
| **Claude Desktop** | stdio | `%APPDATA%\Claude\claude_desktop_config.json` |
| **Codex CLI** (OpenAI) | stdio | `~/.codex/config.json` |
| **Gemini CLI** (Google) | stdio | `~/.gemini/settings.json` |
| **Cursor** | stdio / sse | `.cursor/mcp.json` |
| **VS Code** (Continue, Cline) | stdio / sse | `.vscode/mcp.json` |
| **JetBrains** (IntelliJ, PyCharm) | stdio / sse | Settings > AI Assistant > MCP |
| **Windsurf / Cline** | stdio | `.mcp/config.json` |

<details>
<summary><strong>Quick config (stdio)</strong></summary>

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

</details>

<details>
<summary><strong>SSE mode (for IDEs)</strong></summary>

```powershell
python server.py --transport sse --port 8765
```

Point your IDE to `http://localhost:8765/sse`

</details>

Full setup guides:
- **[docs/SETUP.md](docs/SETUP.md)** — config for all 8 platforms
- **[docs/WINDOWS_SETUP.md](docs/WINDOWS_SETUP.md)** — step-by-step Windows install

---

## Tools (56)

### Power BI Core

| Tool | Description |
|:---|:---|
| `pbi_connect` | Auto-discover and connect to the running PBI Desktop instance |
| `pbi_list_instances` | List discovered local Power BI Desktop instances and ports |
| `pbi_list_tables` | List all tables with columns and data types |
| `pbi_list_measures` | List all DAX measures |
| `pbi_list_relationships` | List all relationships |
| `pbi_model_info` | Full model snapshot (tables + measures + relationships) |

### Mutations

| Tool | Description |
|:---|:---|
| `pbi_create_measure` | Create or update a DAX measure |
| `pbi_delete_measure` | Delete a measure |
| `pbi_set_format` | Batch-format measures (number format strings) |
| `pbi_create_relationship` | Create a relationship between two tables |
| `pbi_create_table` | Create a calculated table (DAX expression) |
| `pbi_create_column` | Create a calculated column in a table |

### Query & Import

| Tool | Description |
|:---|:---|
| `pbi_execute_dax` | Run any DAX query and get results as JSON |
| `pbi_refresh` | Trigger a data model refresh |
| `pbi_import_dax_file` | Bulk-import measures from a `.dax` file |
| `pbi_export_model` | Export full model as JSON for version control |

### Power Query (M)

| Tool | Description |
|:---|:---|
| `pbi_get_power_query` | Read the M expression for a table partition |
| `pbi_list_power_queries` | List all tables with their M expressions and source types |
| `pbi_set_power_query` | Inject or replace a validated M expression |
| `pbi_create_import_query` | Auto-generate Excel sheet import query |
| `pbi_create_csv_import_query` | Auto-generate CSV file import query |
| `pbi_create_folder_import_query` | Auto-generate folder import query |
| `pbi_bulk_import_excel` | Map all workbook sheets to tables and inject queries |

### Excel

| Tool | Description |
|:---|:---|
| `excel_list_sheets` | List workbook sheets with row and column counts |
| `excel_read_sheet` | Read rows from a worksheet or range |
| `excel_read_cell` | Read a single cell with value, type, format, and formula |
| `excel_search` | Search workbook values across one or all sheets |
| `excel_write_cell` | Write a single cell value with optional number format |
| `excel_write_range` | Write a 2D array into a worksheet |
| `excel_create_sheet` | Create a worksheet in an existing workbook |
| `excel_delete_sheet` | Delete a worksheet from an existing workbook |
| `excel_format_range` | Apply formatting to a worksheet range |
| `excel_auto_width` | Auto-fit worksheet column widths |
| `excel_create_workbook` | Create a new `.xlsx` workbook with optional sheets |
| `excel_workbook_info` | Return workbook metadata, named ranges, and dimensions |
| `excel_to_pbi_check` | Compare Excel sheets against the current PBI model |

### Visual Tools

| Tool | Description |
|:---|:---|
| `pbi_extract_report` | Extract a `.pbix` into a pbi-tools report folder |
| `pbi_compile_report` | Compile an extracted report folder back into a `.pbix` |
| `pbi_list_pages` | List pages in an extracted report |
| `pbi_get_page` | Inspect a page and all of its visuals |
| `pbi_create_page` | Create a new report page |
| `pbi_delete_page` | Delete a report page |
| `pbi_set_page_size` | Resize a report page |
| `pbi_add_card` | Add a KPI card visual |
| `pbi_add_bar_chart` | Add a clustered bar chart visual |
| `pbi_add_line_chart` | Add a line chart visual |
| `pbi_add_donut_chart` | Add a donut chart visual |
| `pbi_add_gauge` | Add a gauge visual |
| `pbi_add_table_visual` | Add a table visual |
| `pbi_add_waterfall` | Add a waterfall chart visual |
| `pbi_add_slicer` | Add a slicer visual |
| `pbi_add_text_box` | Add a text box visual |
| `pbi_remove_visual` | Remove a visual from a page |
| `pbi_move_visual` | Move or resize a visual |
| `pbi_apply_theme` | Apply a theme JSON to an extracted report |
| `pbi_build_dashboard` | Build a full page from a layout spec |

---

## Full Automation Workflow

The complete data pipeline is now automatable end to end:

```
 1. Excel source          excel_create_workbook / excel_write_range
 2. Power Query import    pbi_bulk_import_excel / pbi_create_import_query
 3. Relationships         pbi_create_relationship
 4. DAX measures          pbi_import_dax_file / pbi_create_measure
 5. Refresh & validate    pbi_refresh -> pbi_execute_dax
 6. Report extract        pbi_extract_report
 7. Pages & visuals       pbi_create_page / pbi_build_dashboard / pbi_add_*
 8. Theme                 pbi_apply_theme
 9. Compile               pbi_compile_report
```

---

## What's automated vs. manual

| Automated via MCP | Still manual in PBI Desktop |
|:---|:---|
| Data source setup (Power Query M) | Visual Power Query editor |
| DAX measures (create, update, bulk import) | Advanced visual formatting edge cases |
| Relationships between tables | Drillthrough, bookmarks, and complex interactions |
| Calculated tables and columns | Custom visuals marketplace management |
| Excel read, write, format, validate | Live visual preview while editing layout |
| Report extract / compile / page CRUD | Report publishing |
| Standard visuals (card, charts, table, slicer, text) | |
| Theme file application | |
| Model export to JSON | |

---

## Project Structure

```
powerbi-mcp-local/
├── server.py               56 MCP tools (stdio + sse transport)
├── pbi_connection.py       Connection manager (port discovery, TOM, ADOMD)
├── security.py             Security policy, validation, and redaction helpers
├── SECURITY.md             Threat model and hardening guide
├── pyproject.toml          PyPI packaging metadata
├── tools/
│   ├── model.py            Tables, columns, export, model info
│   ├── measures.py         Measures CRUD, .dax bulk import
│   ├── relationships.py    Relationships CRUD
│   ├── query.py            DAX execution, data refresh
│   ├── power_query.py      Power Query (M) partition tools
│   ├── excel.py            Excel read/write/format/pipeline
│   └── visuals.py          Report pages, visuals, themes, and pbi-tools bridge
├── test_connection.py      PBI connectivity test
├── test_security.py        Security controls test suite (8 tests)
├── test_excel.py           Excel tools test suite
├── test_power_query.py     Power Query tools test suite (7 tests)
├── test_visuals.py         Visual layout tool test suite (5 tests)
├── docs/
│   ├── SETUP.md            Multi-platform setup guide
│   └── WINDOWS_SETUP.md    Step-by-step Windows install
├── CLAUDE.md               Build instructions for Claude Code
├── EXCEL_SPEC.md           Excel extension spec
└── VISUAL_SPEC.md          Visual layer extension spec
```

---

## Tech Stack

| Package | Version | Role |
|:---|:---:|:---|
| [`mcp[cli]`](https://pypi.org/project/mcp/) | 1.27.0 | Anthropic MCP SDK |
| [`openpyxl`](https://pypi.org/project/openpyxl/) | 3.1.5 | Excel workbook backend |
| [`pbi-pyadomd`](https://pypi.org/project/pbi-pyadomd/) | 1.4.3 | ADOMD.NET wrapper (DAX queries) |
| [`pythonnet`](https://pypi.org/project/pythonnet/) | 3.0.5 | .NET bridge (TOM model writes) |
| [`psutil`](https://pypi.org/project/psutil/) | 7.2.2 | Process port discovery |
| `pbi-tools` | external CLI | Report extract / compile for visual automation |

## Requirements

- **Windows** — Power BI Desktop only runs on Windows
- **Power BI Desktop** (free) — installed and open with a `.pbix` file
- **Python 3.11+**
- **pbi-tools** (optional) — only needed for visual layer tools

---

## Security

The server includes a security middleware with path traversal protection,
DAX/DMV injection blocking, M expression SSRF prevention, audit logging,
export redaction, ZIP bomb protection, and a configurable security policy.

See **[SECURITY.md](SECURITY.md)** for the full threat model, controls, and
OWASP/CWE mapping.

---

## License

MIT
