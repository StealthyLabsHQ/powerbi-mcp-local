# Security

## Threat Model

This MCP server runs on the same host and under the same user context as Power BI Desktop.
Its highest-risk failure mode is not remote code execution through a public API; it is an
LLM-driven local agent being tricked into issuing high-impact model, query, or file operations.

Primary trust assumptions:

- The host is trusted and locally administered.
- Power BI Desktop is already running under the same Windows user.
- MCP clients may issue arbitrary tool calls unless the server enforces policy.

This security pass was cross-checked against the local hardening references in:

- `references/mcp-security.md`
- `references/llm-agent-security.md`
- `references/supply-chain-security.md`
- `references/owasp-top10.md`

## Security Controls

| Control | Default | Purpose | CWE / OWASP |
|---|---|---|---|
| Path allowlist for local files | On | Restrict Excel, `.dax`, and JSON export paths to approved base directories | CWE-22, A01 |
| Symlink blocking | On | Block symlink-based traversal and reduce TOCTOU abuse | CWE-59, CWE-367, A01 |
| Excel extension allowlist | On | Only `.xlsx` and `.xlsm` are accepted by Excel tools unless policy overrides it | CWE-73, A05 |
| Excel ZIP inspection | On | Reject oversized or suspicious workbook archives before parsing | CWE-409, CWE-400, A06 |
| DAX DMV/system query blocking | On | Prevent metadata exfiltration through `$SYSTEM.*` and discovery DMVs | CWE-200, A01 |
| M network-source blocking | On | Reject external/network M functions by default | CWE-918, CWE-200, A10 |
| Expression and name validation | On | Bound input sizes and reject unsafe object names or query-only DAX syntax | CWE-20, CWE-400, A03 |
| Audit logging | On | Emit `TOOL_CALL`, `TOOL_OK`, `TOOL_FAIL` with sanitized parameters | CWE-778, A09 |
| Localhost SSE default | On | SSE binds to `127.0.0.1` unless explicitly overridden | CWE-306, A05 |
| Security policy engine | On | Central policy can allow/deny tool categories, disable tools, cap rows and input sizes | CWE-693, A04 |
| Readonly mode | Off | Blocks all write/destructive tools when enabled via CLI or env/policy | CWE-285, A01 |
| Runaway-call detection | Warn-only by default | Warn on high call volume; optional hard rate limit via policy | CWE-770, CWE-400 |
| Export redaction | On | Redact obvious credentials/tokens from JSON exports and logs | CWE-200, A02 |
| Explicit DLL loading | On | Load Analysis Services assemblies from explicit absolute directories only | CWE-427, A08 |

## Tool Classification

### Read Tools

- `pbi_connect`
- `pbi_list_instances`
- `pbi_list_tables`
- `pbi_list_measures`
- `pbi_list_relationships`
- `pbi_model_info`
- `pbi_execute_dax`
- `pbi_get_power_query`
- `pbi_list_power_queries`
- `excel_list_sheets`
- `excel_read_sheet`
- `excel_read_cell`
- `excel_search`
- `excel_workbook_info`
- `excel_to_pbi_check`

### Write Tools

- `pbi_create_measure`
- `pbi_create_relationship`
- `pbi_set_format`
- `pbi_refresh`
- `pbi_import_dax_file`
- `pbi_create_table`
- `pbi_create_column`
- `pbi_export_model`
- `pbi_set_power_query`
- `pbi_create_import_query`
- `pbi_create_csv_import_query`
- `pbi_create_folder_import_query`
- `excel_write_cell`
- `excel_write_range`
- `excel_create_sheet`
- `excel_format_range`
- `excel_auto_width`
- `excel_create_workbook`

### Destructive Tools

- `pbi_delete_measure`
- `pbi_bulk_import_excel`
- `excel_delete_sheet`

## Current Hardening Details

### 1. Local File Access

Excel tools, `.dax` imports, and model JSON exports are all resolved through the same
policy-aware path validator.

Protections:

- Relative paths are resolved against the current working directory.
- Paths must remain within the configured allowlist.
- Symlink paths are rejected.
- Excel tools accept only policy-approved extensions.
- Existing workbook archives are inspected before `openpyxl` loads them.

### 2. DAX Query Guard

`pbi_execute_dax` blocks common DMV/system metadata queries by default:

- `$SYSTEM.*`
- `DISCOVER_*`
- `DBSCHEMA_*`
- `MDSCHEMA_*`

The guard is intentionally narrow: it blocks high-risk server metadata reads without
preventing legitimate semantic-model DAX queries.

### 3. M Expression Guard

Power Query expressions are validated before injection.

Blocked by default:

- `Web.*`
- `OData.Feed`
- `Sql.*`, `Oracle.*`, `PostgreSQL.*`, `MySQL.*`
- `Odbc.*`, `OleDb.*`
- `SharePoint.*`, `ActiveDirectory.*`, `AzureStorage.*`
- `Expression.Evaluate`
- `Value.NativeQuery`

The validator also performs a basic syntax sanity check:

- balanced `()`, `[]`, `{}`
- string literal balancing
- `let ... in ...` structure presence when applicable

### 4. DAX Measure / Table / Column Validation

The server now applies shared validation before model mutation:

- object names cannot be empty
- control characters are rejected
- measure names cannot contain brackets or quotes
- expression length is capped by policy
- query-only DAX syntax such as `EVALUATE` or `DEFINE` is rejected where a model expression is expected

This prevents the `.dax` bulk importer from accepting crafted measure headers that would
blur the boundary between model expressions and query text.

### 5. Export Redaction

`pbi_export_model` recursively redacts obvious secret-bearing values from both the
returned JSON payload and any written `.json` export.

Patterns redacted:

- `password=...`
- `pwd=...`
- `accountkey=...`
- `sharedaccesssignature=...`
- `clientsecret=...`
- `token=...`
- `user id=...`
- URI credentials in `scheme://user:password@host`

### 6. Security Middleware

`security.py` provides the central enforcement layer used by the server:

- tool classification
- policy loading from env or `security_policy.json`
- per-call validation
- input length limits
- path validation
- rate-limit / high-volume detection
- readonly-mode enforcement
- sanitized logging helpers

### 7. Readonly Mode

Readonly mode blocks every write and destructive tool before execution.

Enable it with:

```powershell
python server.py --readonly
```

or:

```powershell
set PBI_MCP_READONLY=1
python server.py
```

Readonly mode is intentionally off by default because it changes server behavior
for legitimate automation workflows.

### 8. SSE Transport

SSE binds to `127.0.0.1` by default. If you bind to `0.0.0.0`, the server logs a
security warning because the transport becomes reachable off-host.

## Security Policy

The server loads policy from one of these sources:

1. `PBI_MCP_SECURITY_POLICY` set to a JSON string
2. `PBI_MCP_SECURITY_POLICY` set to a path to a JSON file
3. `security_policy.json` in the working directory

Supported policy keys:

| Key | Type | Default | Meaning |
|---|---|---:|---|
| `allow_categories` | list[str] | `["read","write","destructive"]` | Categories allowed |
| `deny_categories` | list[str] | `[]` | Categories explicitly denied |
| `enabled_tools` | list[str] | `null` | Optional allowlist of exact tool names |
| `disabled_tools` | list[str] | `[]` | Exact tools to block |
| `readonly` | bool | `false` | Blocks write/destructive tools |
| `max_string_length` | int | `8192` | Generic string limit |
| `max_name_length` | int | `256` | Object-name limit |
| `max_expression_length` | int | `200000` | DAX/M expression limit |
| `max_query_length` | int | `100000` | DAX query limit |
| `max_path_length` | int | `4096` | Local path limit |
| `max_dax_rows` | int | `5000` | Max `max_rows` accepted by `pbi_execute_dax` |
| `allowed_excel_extensions` | list[str] | `[".xlsx",".xlsm"]` | Allowed Excel extensions |
| `allowed_base_dirs` | list[str] | cwd | Local filesystem roots allowed |
| `max_excel_zip_uncompressed_bytes` | int | `104857600` | Excel ZIP decompressed-size cap |
| `max_excel_zip_members` | int | `10000` | Excel ZIP member cap |
| `max_excel_zip_compression_ratio` | float | `250.0` | Excel ZIP compression-ratio cap |
| `max_excel_cells_scanned` | int | `200000` | Max cells scanned by read/search operations |
| `warn_after_calls_per_minute` | int | `120` | Warning threshold for tool-call rate |
| `rate_limit_calls_per_minute` | int/null | `null` | Hard limit for tool-call rate |

### Production Example

`security_policy.json`

```json
{
  "readonly": false,
  "deny_categories": [],
  "disabled_tools": [
    "pbi_delete_measure",
    "excel_delete_sheet"
  ],
  "allowed_base_dirs": [
    "C:\\\\Data\\\\PowerBI",
    "C:\\\\Repos\\\\powerbi-mcp-local\\\\samples"
  ],
  "allowed_excel_extensions": [".xlsx", ".xlsm"],
  "max_expression_length": 50000,
  "max_query_length": 20000,
  "max_dax_rows": 1000,
  "max_excel_zip_uncompressed_bytes": 52428800,
  "max_excel_zip_members": 4000,
  "max_excel_zip_compression_ratio": 100.0,
  "max_excel_cells_scanned": 100000,
  "warn_after_calls_per_minute": 60,
  "rate_limit_calls_per_minute": 120
}
```

Recommended launch:

```powershell
set PBI_MCP_SECURITY_POLICY=C:\path\to\security_policy.json
python server.py --transport stdio
```

## Environment Variables

| Variable | Default | Description |
|---|---|---|
| `PBI_MCP_ALLOWED_DIRS` | cwd | Semicolon-separated allowed filesystem roots |
| `PBI_MCP_ALLOW_DMV` | `0` | Set to `1` to allow DMV/system queries |
| `PBI_MCP_ALLOW_EXTERNAL_M` | `0` | Set to `1` to allow blocked external M functions |
| `PBI_MCP_LOG_LEVEL` | `INFO` | Logging level |
| `PBI_MCP_SECURITY_POLICY` | unset | JSON string or path to `security_policy.json` |
| `PBI_MCP_READONLY` | `0` | Set to `1` to block write/destructive tools |
| `PBI_DESKTOP_BIN` | unset | Absolute override path to the Power BI Desktop `bin` directory |
| `PBI_DLL_DIR` | unset | Absolute override path to Analysis Services DLLs |
| `PBI_WORKSPACE_ROOTS` | unset | Extra workspace roots to scan for local SSAS instances |

## Known Limitations

- This is not a sandbox. Any process running as the same local user can still target the same Power BI Desktop session.
- The server cannot fully neutralize malicious logic embedded in a legitimate DAX or M expression; it can only validate structure and block known high-risk patterns.
- `pythonnet` / CLR interop still trusts Microsoft Analysis Services assemblies present on disk. The server now loads them by explicit path, but cannot verify their publisher signature.
- `openpyxl` reads Office Open XML packages in Python space. ZIP bomb controls reduce risk, but parser-level vulnerabilities in the dependency itself remain a supply-chain concern.
- Symlink blocking reduces traversal and TOCTOU risk, but no pure-Python path validator can provide a perfect race-free guarantee across every filesystem and OS behavior.
- SSE has no built-in authentication. Localhost binding is the primary protection.
- Secrets already loaded into a live Power BI Desktop process are outside the scope of this server. The server can redact obvious secrets in exports, but cannot prove a model contains none.

## Reporting Vulnerabilities

If you find a security issue, report it privately to the maintainer.
Do not open a public GitHub issue with exploit details before coordinated disclosure.
