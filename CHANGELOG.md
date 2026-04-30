# Changelog

## [Unreleased]

### Added

- `pbi_import_excel_workbook` - explicit one-call Excel workbook import tool for Power BI tables.
- `pbi_model_audit_workflow`, `pbi_excel_import_workflow`, `pbi_measure_workflow` - guided workflow tools with dry-run defaults for higher-productivity LLM agents.
- `pbi_validate_report_fields` and `pbi_repair_report_fields` - detect and repair broken report visual field bindings that cause Power BI "Fix this" placeholders.

All notable changes to this project are documented here.

## [0.7.0] — 2026-04-22 — Schema cache + batch measures + model audit + MCP Resources & Prompts

### Added — new tools (+2)

- `pbi_create_measures` — batch create/update multiple DAX measures in a single `SaveChanges()` call; accepts a list of `{name, expression, format_string?, description?, display_folder?, is_hidden?}` items.
- `pbi_validate_model` — model audit: reports empty expressions, visible measures without format strings, orphan tables (no relationships + no measures), and duplicate measure names across tables.

### Added — MCP Resources (3)

- `powerbi://model/schema` — live full model snapshot (tables + measures + relationships)
- `powerbi://model/measures` — live measures list
- `powerbi://model/relationships` — live relationships list

These are fetched on-demand by the MCP client without burning a tool call.

### Added — MCP Prompts (8)

Ready-to-use workflow prompts surfaced natively to any MCP client:
`model_audit`, `time_intelligence_kit`, `star_schema_builder`, `rls_setup`, `dead_measure_scan`, `bulk_measure_format_fix`, `excel_to_pbi_pipeline`, `model_snapshot_export`.

### Performance — schema read cache

`pbi_list_tables`, `pbi_list_measures`, `pbi_list_relationships` now use a write-generation cache inside `PowerBIConnectionManager`. Results are served from memory until the next write (`execute_write`), reconnect, or `pbi_refresh_metadata`. Typical speedup for repeated reads: 10–50× (avoids TOM iteration).

---

## [0.6.0] — 2026-04-21 — Full CRUD + RLS + Calc Groups + Infra (80 tools)

Covers three work streams ("Lot 1/2/3") and end-to-end live validation against a real Power BI Desktop model (78/80 tools exercised; 2 hors-portée due to external tooling — `pbi_extract_report`/`pbi_compile_report` rely on an `extract` action that pbi-tools.core 1.2.0 no longer ships).

### Added — new tools (+22)

- CRUD completion:
  `pbi_delete_relationship`, `pbi_update_relationship`,
  `pbi_delete_table`, `pbi_delete_column`,
  `pbi_rename_table`, `pbi_rename_column`, `pbi_rename_measure`
- DAX introspection:
  `pbi_validate_dax` (parse-check via zero/one-row probe),
  `pbi_measure_dependencies` (DISCOVER_CALC_DEPENDENCY)
- Cache management:
  `pbi_refresh_metadata` (cheap TOM schema reload)
- Row-level security CRUD (6):
  `pbi_list_roles`, `pbi_create_role`, `pbi_delete_role`,
  `pbi_set_role_filter`, `pbi_add_role_member`, `pbi_remove_role_member`
- Calculation groups CRUD (3):
  `pbi_list_calc_groups`, `pbi_create_calc_group`, `pbi_delete_calc_group`
- Unified visual dispatcher:
  `pbi_add_visual(visual_type, config)` — consolidates the 9 per-type add tools (kept as shims for back-compat)

### Added — infrastructure

- `--profile readonly|write|all` startup flag: prunes registered MCP tools at boot (smaller surface for SSE / restricted clients)
- SSE bearer authentication via `PBI_MCP_AUTH_TOKEN` env var (warns when SSE is exposed without a token)
- `timeout_seconds` parameter on `pbi_execute_dax` and `pbi_trace_query`, threaded through the pyadomd and pythonnet backends
- GitHub Actions CI workflow running offline security tests on Windows + Linux (Python 3.11/3.12)
- LICENSE file (MIT) at repo root; `pyproject.toml` now points to it via `{ file = "LICENSE" }`
- End-to-end test scripts under `tests/`: `smoke_e2e.py`, `demo_write_cycle.py`, `demo_design_cycle.py`, `demo_full_cycle.py`, `demo_risky_cycle.py`

### Changed

- `src/__init__.py` bootstraps `sys.path` so flat imports (`from pbi_connection import ...`) work under both script mode and installed-package mode. Fixes the previously broken `powerbi-mcp-local` console entry point.
- Power BI Desktop DLL discovery now also probes the Windows registry (HKLM/HKCU install keys + App Paths), `%PROGRAMFILES%\WindowsApps` (Microsoft Store installs), `%LOCALAPPDATA%\Programs`, and `shutil.which` on the PATH.
- `pbi_bulk_import_excel` reclassified from DESTRUCTIVE to WRITE — it creates or replaces query partitions but does not delete model objects.
- Dependency specifiers relaxed from `==` to `~=` so patch-level security updates are picked up.
- Test runner standardized on pytest (`[tool.pytest.ini_options]` + `[project.optional-dependencies].dev = [pytest]`); README updated accordingly.

### Fixed — real bugs caught by live testing

- `_map_cardinality`: `oneToMany` mapped to `(One, Many)` but Tabular requires the "from" side to always be Many (FK) and "to" to be One (PK). Both `oneToMany` and `manyToOne` now canonicalize to `(Many, One)`.
- `pbi_create_relationship`: one-to-one relationships with `direction=oneDirection` are rejected by SSAS — auto-upgrades to `bothDirections` in that case.
- `_get_target_partition` / `pbi_bulk_import_excel`: `table.Partitions[0]` raised because the .NET `NamedMetadataObjectCollection` indexer expects a string (partition name), not an int. Replaced with `next(iter(table.Partitions))`.
- `pbi_create_calc_group`: the single data column must always have `SourceColumn = "Name"` (only the displayed `Name` may vary), and the model requires `DiscourageImplicitMeasures = True` before any calculation group can be saved — both are now enforced automatically.
- `_build_csv_m`: the multi-line `Csv.Document(…)` call was being joined with the top-level step separator (`",\n"`), inserting stray commas inside the function invocation and producing `M Engine error: Token Literal expected`. Fixed by emitting the call as one pre-joined block.

---

## [0.5.1] — 2026-04-16 — Visual Layer follow-ups

### Added
- `pbi_patch_layout`: direct PBIX Layout patch tool that swaps `Report/Layout`, removes `SecurityBindings`, preserves ZIP entry metadata, and supports `force=True` auto-kill on Windows.
- `pbi_apply_design`: one-shot design preset tool that writes the base theme, updates page backgrounds, and applies card container styling.

### Fixed
- Visual query generation now resolves measure home tables from extract metadata (`Model/tables/*/measures/*.dax`) so `prototypeQuery.From[]` uses the real table entity instead of `"$Measures"` when available.
- Measure fallback now logs a warning and uses `"$Measures"` only when extract metadata is missing or disconnected.
- `.dax` import parsing now strips `//` and `/* ... */` comments safely while preserving quoted strings and measure block boundaries.
- `pbi_compile_report` now accepts `force: bool = False` and can auto-kill `PBIDesktop.exe` on Windows before write operations.
- RLS role-scoped execution now rejects `role` / `username` values containing connection-string separators before building the ADOMD connection string.
- Power Query M validation now blocks `#shared` and rejects function calls outside a strict local-file allowlist unless `PBI_MCP_ALLOW_EXTERNAL_M=1` is set.
- MCP responses that expose model or Power Query expressions now redact secret-like values before returning them to clients.

---

## [0.5.0] — 2026-04-16 — Visual Layer (56 tools)

### Added
- 20 visual tools via pbi-tools extract/compile pipeline:
  `pbi_extract_report`, `pbi_compile_report`, `pbi_create_page`, `pbi_delete_page`,
  `pbi_get_page`, `pbi_list_pages`, `pbi_set_page_size`, `pbi_add_card`,
  `pbi_add_bar_chart`, `pbi_add_line_chart`, `pbi_add_donut_chart`, `pbi_add_gauge`,
  `pbi_add_slicer`, `pbi_add_table_visual`, `pbi_add_waterfall`, `pbi_add_text_box`,
  `pbi_move_visual`, `pbi_remove_visual`, `pbi_apply_theme`, `pbi_build_dashboard`

### Docs
- Visual layer specification
- Windows setup guide
- README redesigned with badges, collapsible config, and pipeline diagram

---

## [0.4.0] — 2026-04-16 — Security Hardening

### Added
- `security.py` middleware: path traversal, DAX injection, and SSRF protection
- 15 security tests

### Fixed
- 7 vulnerabilities: path traversal, DAX injection, SSRF, logging exposure

---

## [0.3.0] — 2026-04-16 — Multi-Platform + Power Query v2 (36 tools)

### Added
- SSE transport support
- Power Query v2: CSV import, folder import
- `pyproject.toml` packaging

---

## [0.2.0] — 2026-04-16 — Excel + Power Query (33 tools)

### Added
- 13 Excel tools: `excel_create_workbook`, `excel_read_sheet`, `excel_write_range`, and more
- `pbi_list_instances`
- 4 Power Query tools: `pbi_get_power_query`, `pbi_set_power_query`, `pbi_list_power_queries`, `pbi_create_import_query`

---

## [0.1.0] — 2026-04-16 — Initial Release (15 tools)

### Added
- Core MCP server implementation
- Model inspection, DAX measure management, relationships, refresh, DAX query execution
- `pbi_connect`, `pbi_model_info`, `pbi_list_tables`, `pbi_list_measures`, `pbi_create_measure`,
  `pbi_delete_measure`, `pbi_execute_dax`, `pbi_refresh`, `pbi_list_relationships`,
  `pbi_create_relationship`, `pbi_export_model`, `pbi_set_format`, `pbi_create_column`,
  `pbi_create_table`, `pbi_import_dax_file`
