# Security

## Threat Model

This MCP server runs on the same Windows machine as Power BI Desktop.
It connects to the local SSAS instance (localhost only) and can read/write
the data model, execute DAX queries, and access Excel files on disk.

**Primary risk**: an AI agent (LLM) being manipulated via prompt injection
to perform unintended actions through the MCP tools.

## Security Controls

### 1. Path Traversal Protection (CWE-22)

Excel tools validate all file paths against allowed base directories.

**Default**: only files under the current working directory are accessible.

**Configure**:
```powershell
# Allow specific directories (semicolon-separated)
set PBI_MCP_ALLOWED_DIRS=C:\Projects\data;D:\Reports
```

Or in code:
```python
from tools.excel import configure_allowed_dirs
configure_allowed_dirs(["C:\\Projects\\data", "D:\\Reports"])
```

### 2. DAX/DMV Injection Protection (CWE-94)

System catalog queries (`$SYSTEM.*`, `DISCOVER_*`, `DBSCHEMA_*`) are
blocked by default to prevent server metadata exfiltration.

**Override** (for advanced users):
```powershell
set PBI_MCP_ALLOW_DMV=1
```

### 3. M Expression Injection / SSRF Protection (CWE-918)

Power Query expressions are validated before injection. The following
M functions are blocked by default because they can make network calls
or connect to external databases:

- `Web.Contents`, `Web.Page`, `Web.BrowserContents`
- `OData.Feed`
- `Sql.Database`, `Oracle.Database`, `PostgreSQL.Database`, `MySQL.Database`
- `Odbc.*`, `OleDb.*`
- `SharePoint.*`, `ActiveDirectory.*`, `AzureStorage.*`

Only local file-based sources (`File.Contents`, `Csv.Document`,
`Excel.Workbook`, `Folder.Files`) are allowed.

**Override** (for trusted environments):
```powershell
set PBI_MCP_ALLOW_EXTERNAL_M=1
```

### 4. Audit Logging (CWE-778)

Every tool call is logged with:
- Tool name
- Parameters (sensitive values truncated at 200 chars)
- Result status (ok / fail)
- Errors with details

Logs go to stderr. Redirect to a file for persistent audit:
```powershell
python server.py 2>> mcp_audit.log
```

### 5. SSE Transport Security (CWE-306)

SSE mode binds to `127.0.0.1` by default (localhost only).

If you bind to `0.0.0.0`, a warning is logged. Only do this on
trusted networks.

```powershell
# Safe (default)
python server.py --transport sse

# Exposed — only on trusted networks
python server.py --transport sse --host 0.0.0.0
```

### 6. Tool Classification

| Category | Tools | Risk |
|---|---|---|
| **Read** | list_tables, list_measures, model_info, read_sheet, ... | Low |
| **Write** | create_measure, write_cell, set_power_query, ... | Medium |
| **Destructive** | delete_measure, delete_sheet, bulk_import_excel | High |

MCP clients with human-in-the-loop support (Claude Code, Cursor) will
prompt for confirmation on write/destructive actions based on their
own permission policies.

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `PBI_MCP_ALLOWED_DIRS` | (cwd) | Semicolon-separated allowed directories for Excel tools |
| `PBI_MCP_ALLOW_DMV` | `0` | Set to `1` to allow DMV/system queries |
| `PBI_MCP_ALLOW_EXTERNAL_M` | `0` | Set to `1` to allow network M functions |
| `PBI_MCP_LOG_LEVEL` | `INFO` | Logging level (DEBUG, INFO, WARNING, ERROR) |

## Reporting Vulnerabilities

If you find a security issue, please open a private issue on GitHub
or email the maintainer directly. Do not open public issues for
security vulnerabilities.
