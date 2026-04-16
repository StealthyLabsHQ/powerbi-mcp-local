# powerbi-mcp-local

MCP server that connects to Power BI Desktop's local Analysis Services instance.
Lets Claude Code (or any MCP client) read, write, and query your Power BI data model programmatically.

## What it does

| Tool | Description |
|---|---|
| `pbi_connect` | Auto-discover and connect to the running PBI Desktop instance |
| `pbi_list_tables` | List all tables with columns and data types |
| `pbi_list_measures` | List all DAX measures |
| `pbi_list_relationships` | List all relationships |
| `pbi_create_measure` | Create or update a DAX measure |
| `pbi_create_relationship` | Create a relationship between two tables |
| `pbi_delete_measure` | Delete a measure |
| `pbi_execute_dax` | Run any DAX query and get results |
| `pbi_model_info` | Full model snapshot (tables + measures + relationships) |

## How it works

```
Claude Code ──(MCP/stdio)──> server.py ──(ADOMD.NET)──> PBI Desktop (SSAS)
```

Power BI Desktop runs a local Analysis Services engine on a random port.
This server finds the port automatically and connects via ADOMD.NET.

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
python -c "from server import find_pbi_port; print(f'PBI on port {find_pbi_port()}')"

# 5. Run with Claude Code
claude
```

Then tell Claude Code:

> "Read the CLAUDE.md and build the MCP server. Then connect to Power BI and inject the measures."

## Configuration

Add to your `.claude/settings.json`:

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

## What it automates vs. what stays manual

| Automated via MCP | Manual in PBI Desktop |
|---|---|
| DAX measures (create, update, delete) | Data import (Excel, CSV) |
| Relationships between tables | Theme import (JSON) |
| DAX query execution | Visual creation (charts, cards) |
| Model inspection | Page layout and formatting |

The MCP handles the **data model layer**. The **visual layer** (charts, layouts, pages) must be built manually in Power BI Desktop — there is no API for it.

## Tech Stack

- [`mcp[cli]`](https://github.com/anthropics/anthropic-sdk-python) — Anthropic MCP SDK
- [`pyadomd`](https://github.com/S-C-A-N/pyadomd) — Python ADOMD.NET wrapper
- [`pythonnet`](https://github.com/pythonnet/pythonnet) — .NET/Python bridge (fallback)
- [`psutil`](https://github.com/giampaolo/psutil) — Process utilities

## License

MIT
