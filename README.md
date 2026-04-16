# powerbi-mcp-local

MCP server that connects to Power BI Desktop's local Analysis Services instance.
Lets Claude Code (or any MCP client) read, write, and query your Power BI data model programmatically.

## Tools (15)

### Core

| Tool | Description |
|---|---|
| `pbi_connect` | Auto-discover and connect to the running PBI Desktop instance |
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

## How it works

```
Claude Code ──(MCP/stdio)──> server.py ──(TOM/.NET)──> PBI Desktop (local SSAS)
```

Power BI Desktop runs a local Analysis Services engine on a random port.
The server finds the port automatically via:
1. `%LOCALAPPDATA%` workspace scan (MSI install)
2. User-scoped path scan (Microsoft Store install)
3. `msmdsrv.exe` process fallback

Connection is managed by `PowerBIConnectionManager` — thread-safe, auto-reconnects on port change, hard-resets on write failure.

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

## Configuration

Add to your project's `.claude/settings.json`:

```json
{
  "mcpServers": {
    "powerbi-desktop": {
      "command": "python",
      "args": ["server.py"]
    }
  }
}
```

Or globally in `%APPDATA%\Claude\claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "powerbi-desktop": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\server.py"],
      "env": {
        "PYTHONPATH": "C:\\Program Files\\Microsoft Power BI Desktop\\bin"
      }
    }
  }
}
```

## Project Structure

```
powerbi-mcp-local/
├── server.py              Main MCP server (FastMCP, 15 tools)
├── pbi_connection.py      Connection manager (port discovery, TOM, ADOMD)
├── tools/
│   ├── __init__.py        Tool exports
│   ├── model.py           Tables, columns, export, model info
│   ├── measures.py        Measures CRUD, formatting, .dax bulk import
│   ├── relationships.py   Relationships CRUD
│   └── query.py           DAX execution, data refresh
├── test_connection.py     Standalone connectivity test
├── requirements.txt       Pinned dependencies
├── CLAUDE.md              Build instructions for Claude Code
└── README.md
```

## What it automates vs. what stays manual

| Automated via MCP | Manual in PBI Desktop |
|---|---|
| DAX measures (create, update, delete, bulk import) | Data import (Excel, CSV) |
| Relationships between tables | Theme import (JSON) |
| Calculated tables and columns | Visual creation (charts, cards) |
| DAX query execution and refresh | Page layout and formatting |
| Model export to JSON | Report publishing |

The MCP handles the **data model layer**. The **visual layer** (charts, layouts, pages) must be built manually in Power BI Desktop — there is no API for it.

## Tech Stack

| Package | Version | Role |
|---|---|---|
| [`mcp[cli]`](https://pypi.org/project/mcp/) | 1.27.0 | Anthropic MCP SDK |
| [`pbi-pyadomd`](https://pypi.org/project/pbi-pyadomd/) | 1.4.3 | ADOMD.NET wrapper (queries) |
| [`pythonnet`](https://pypi.org/project/pythonnet/) | 3.0.5 | .NET bridge (TOM model writes) |
| [`psutil`](https://pypi.org/project/psutil/) | 7.2.2 | Process port discovery fallback |

## License

MIT
