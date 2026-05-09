# Security Audit Archive

This folder preserves the 14 findings flagged by the Codex security scanner
during the v0.10.x → v0.12.x window. All issues have been remediated; the
files are kept here for public traceability and so future audits can
cross-reference what was found, what was shipped, and where to look in the
code.

| # | Title | Severity | Status | Shipped in |
| --- | --- | --- | --- | --- |
| [1](1.md) | Unauthenticated SSE exposes powerful MCP tools | Critical | Fixed | v0.12.0 |
| [2](2.md) | New dependency tool bypasses DMV guard | High | Fixed | v0.12.0 |
| [3](3.md) | Quoted M identifiers bypass external connector blocklist | High | Fixed | v0.12.0 |
| [4](4.md) | Power Query M allowlist can be bypassed | High | Fixed | v0.12.0 |
| [5](5.md) | Server ignores working-directory security policy | High | Fixed | v0.12.0 |
| [6](6.md) | Empty tool allowlist is treated as allow-all | High | Fixed | v0.12.0 |
| [7](7.md) | Correction report tool bypasses readonly policy | High | Fixed | v0.12.0 |
| [8](8.md) | Zip Slip in native PBIX extraction fallback | High | Fixed | v0.12.0 |
| [9](9.md) | New write tools bypass readonly policy | High | Fixed | v0.12.0 |
| [10](10.md) | New visual writers bypass readonly policy | High | Fixed | v0.12.0 |
| [11](11.md) | New mutating tools bypass readonly policy | High | Fixed | v0.12.0 |
| [12](12.md) | Windows OCR leaks full desktop screenshot text | High | Fixed | v0.12.0 |
| [13](13.md) | Readonly PBIX reopen tool can capture screen + close PBI | High | Fixed | v0.12.0 |
| [14](14.md) | Readonly profile exposes apply-capable write workflows | High | Fixed | v0.12.0 |

## Follow-up hardening

The same pass surfaced residual issues that were not in the original
report; v0.12.3 + v0.12.4 closed them:

- Constant-time Bearer comparison (`hmac.compare_digest`)
- 32-character minimum on `PBI_MCP_AUTH_TOKEN`
- Power Query M blocklist extended to modern cloud connectors (Snowflake,
  BigQuery, Redshift, Azure SQL/Synapse, ADLS Gen2, S3, SaaS,
  AnalysisServices, `Excel.CurrentWorkbook`)
- `pbi-tools` subprocess timeout
- `pbi_persist_now` switched from `SendInput` to `PostMessage`
- `_save_and_close_powerbi_gracefully` switched from `WScript SendKeys +
  AppActivate` to direct PostMessage
- Default `rate_limit_calls_per_minute = 600`
- `max_response_bytes = 16 MiB` cap with structured `response_too_large`
  error
- PBIX zip-bomb caps applied during native extraction (decompressed size,
  member count, compression ratio)
- Native PBIX extraction logs zip-slip skips and reports
  `skipped_traversal_count`
- PowerShell helpers force UTF-8 I/O (PS 5.1 default mangles non-ASCII)
- Antigravity adapter keeps a stderr WARNING handler so misconfiguration
  surfaces in the client diagnostics view
- `security_policy.json` hot-reload on mtime change
- Tool registry audit no longer crashes stdio in non-strict mode

See `CHANGELOG.md` v0.12.0 / v0.12.3 / v0.12.4 entries for the
file-by-file diff context.
