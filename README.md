<div align="center">

![powerbi-mcp-local banner](docs/assets/powerbi-mcp-local-banner.png)

# powerbi-mcp-local

**Local-first MCP server for Power BI Desktop automation**

Automate semantic model changes, DAX, Power Query, Excel, and report layout from MCP-capable AI clients.

[![Python 3.11+](https://img.shields.io/badge/python-3.11%2B-blue?logo=python&logoColor=white)](https://python.org)
[![Protocol MCP](https://img.shields.io/badge/protocol-MCP-blueviolet)](https://modelcontextprotocol.io)
[![License MIT](https://img.shields.io/badge/license-MIT-green)](LICENSE)
[![Tools 131](https://img.shields.io/badge/tools-131-orange)](#tool-catalog)
[![CI](https://github.com/StealthyLabsHQ/powerbi-mcp-local/actions/workflows/ci.yml/badge.svg)](https://github.com/StealthyLabsHQ/powerbi-mcp-local/actions/workflows/ci.yml)

</div>

## Quick Links

| Start | Setup | Tools | Security |
| --- | --- | --- | --- |
| [Quick start](#quick-start) | [MCP client setup](#mcp-client-setup) | [Tool catalog](#tool-catalog) | [Security](#security) |

## What It Does

- Connects AI tools to a running local Power BI Desktop engine.
- Automates tables, columns, measures, relationships, DAX, refreshes, and Power Query.
- Reads/writes Excel files for local BI workflows.
- Extracts, patches, validates, and compiles report layouts through `pbi-tools`.

No Power BI Pro license is required for the local Desktop workflow.

## Architecture

```text
MCP Client --(stdio or sse)--> src/server.py
                              |
                              +-- src/mcp_core.py     FastMCP instance, _run, lifecycle
                              +-- src/wrappers/       14 thin wrappers/<domain>.py
                              +-- src/tools/          business logic (the *_tool fns)
                              +-- TOM/.NET ─────────> Power BI Desktop local SSAS
                              +-- ADOMD ────────────> DAX query execution
                              +-- openpyxl ─────────> Excel read/write/format
                              +-- pbi-tools ────────> report extract/compile + visuals
                              +-- src/security.py     path, DAX, payload safeguards
```

Full module layering and the visuals/ submodule fan-out: see [ARCHITECTURE.md](ARCHITECTURE.md).

## Requirements

| Requirement | Install |
| --- | --- |
| Windows | Power BI Desktop local engine is Windows-only |
| Power BI Desktop | `winget install Microsoft.PowerBIDesktop` |
| Python 3.11+ | `winget install Python.Python.3.11` |
| .NET 6+ Runtime | `winget install Microsoft.DotNet.Runtime.6` |
| pbi-tools | `winget install pbi-tools` or `dotnet tool install -g pbi-tools` |

ADOMD.NET ships with Power BI Desktop. If `pbi-tools` is not on `PATH`, set `PBI_TOOLS_PATH`.

<a id="quick-start"></a>
## Quick Start

```powershell
git clone https://github.com/StealthyLabsHQ/powerbi-mcp-local.git
cd powerbi-mcp-local
pip install -r requirements.txt
```

Open Power BI Desktop with a `.pbix` file, then verify connectivity:

```powershell
python tests/test_connection.py
```

Start the MCP server:

```powershell
python src/server.py
```

Useful launch modes:

```powershell
python src/server.py --transport sse --port 8765
python src/server.py --readonly
python src/server.py --profile readonly   # ~56 read tools
python src/server.py --profile write      # readonly + writes (no destructive)
python src/server.py --profile grading    # 25-tool surface for evaluation workflows
python src/server.py --profile all        # default — every tool
```

For SSE auth:

```powershell
$env:PBI_MCP_AUTH_TOKEN = "your-secret-token"
python src/server.py --transport sse --port 8765
```

Clients must send:

```text
Authorization: Bearer your-secret-token
```

<a id="mcp-client-setup"></a>
## MCP Client Setup

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

SSE endpoint:

```text
http://localhost:8765/sse
```

Setup guides:
- [docs/SETUP.md](docs/SETUP.md)
- [docs/WINDOWS_SETUP.md](docs/WINDOWS_SETUP.md)

## Example Prompts

- `Connect to Power BI and list all tables with columns.`
- `Create a measure called Total Sales in table Sales.`
- `Run this DAX query and show top 20 rows.`
- `Extract report, add a new page, place 3 visuals, then compile.`

<a id="tool-catalog"></a>
## Tool Catalog

131 MCP tools are grouped into these areas:

| Area | Coverage |
| --- | --- |
| Model discovery | instances, tables, measures, relationships, metadata, validation |
| Model mutations | measures, columns, tables, relationships, formats, role-based DAX |
| Query and import | DAX execution, traces, validation, refresh, model export |
| Power Query (M) | read, write, create, import, bulk Excel/folder sources |
| PBIP/TMDL | list, read, write, and patch TMDL project files |
| Workflows | model audit, Excel import, measure workflow automation |
| Quality gates | DAX linting, visual checks, persistence, scenarios, report validation |
| RLS and calc groups | roles, filters, members, calculation groups |
| Excel | workbook, sheet, cell/range, formatting, search, Power BI import checks |
| Reports and visuals | extract, compile, pages, cards, charts, slicers, themes, dashboards |

Unified visual creation is available through:

```text
pbi_add_visual(visual_type, config)
```

## Automation Flow

```text
Excel input -> Power Query -> model updates -> measures -> validation -> report layout -> compile PBIX
```

Common tool chain:

```text
excel_write_range
pbi_create_import_query
pbi_create_relationship
pbi_create_measure
pbi_refresh
pbi_execute_dax
pbi_extract_report
pbi_build_dashboard
pbi_compile_report
```

## Troubleshooting

| Symptom | Fix |
| --- | --- |
| `No module named 'clr'` | Install .NET 6+ runtime, then restart terminal |
| `No running PBI Desktop instance found` | Open a `.pbix` in Power BI Desktop first |
| `pbi-tools not found` | Add it to `PATH` or set `PBI_TOOLS_PATH` |
| `PermissionError` on `.xlsx` | Close Excel; workbook files are locked while open |
| Path blocked by policy | Configure `PBI_MCP_ALLOWED_DIRS` |

<a id="security"></a>
## Security

Built-in safeguards include:

- local path restrictions and traversal protection
- DAX/DMV unsafe-query guards
- Power Query SSRF protections
- export redaction controls
- zip safety checks
- tool-call auditing

Details: [SECURITY.md](SECURITY.md)

## Development

```powershell
pip install -e ".[dev]"
pytest -q
ruff check src tests
ruff format --check src tests
```

CI runs `pytest --cov=src --cov-fail-under=54` on Windows + an offline subset on Ubuntu, plus `ruff` lint + format check, on every PR. Strict registry audit (`PBI_MCP_STRICT_REGISTRY=1`) ensures every public `pbi_*_tool` has a matching `@mcp.tool()` wrapper.

## Repository Layout

```text
powerbi-mcp-local/
├── src/
│   ├── server.py            CLI + transport launcher + @mcp.resource/@mcp.prompt (~320 L)
│   ├── mcp_core.py          FastMCP instance + CONNECTION_MANAGER + lifecycle (~250 L)
│   ├── pbi_connection.py    TOM + ADOMD bring-up, write helpers, op history
│   ├── security.py          path / DAX / payload guards + tool category sets
│   ├── wrappers/            14 domain modules — `register_tool(pbi_*_tool)` calls
│   └── tools/               business logic (*_tool functions)
│       └── visuals/         17 focused submodules (layout, bindings, containers, charts, …)
├── tests/                   167 offline unit tests (live-only scripts gitignored)
├── .github/workflows/ci.yml pytest + coverage + ruff on Windows + Ubuntu, py3.11/3.12
├── docs/, specs/
├── ARCHITECTURE.md          module layering, visuals/ tree, profiles, registry audit
├── CHANGELOG.md             active changelog (last 3 releases)
├── CHANGELOG-archive.md     historical changelog
├── SECURITY.md
├── pyproject.toml           ruff, pytest, coverage config + dev deps
└── requirements.txt
```

## License

MIT
