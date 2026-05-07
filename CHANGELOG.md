# Changelog

## [0.8.2] — 2026-05-07 — Live-model field validation, error chains, symlink hardening

### Added — fail-fast field validation

- New helper `_validate_field_references_live(manager, references)` checks every measure / column referenced by a visual against the live TOM model index. Raises `PowerBIValidationError` with a structured `details.missing` list naming the bad references and their kind (`measure` vs `column`). When no manager is supplied (offline scripting / tests), validation is silently skipped — preserves the existing offline workflow.
- `pbi_add_card_tool`, `pbi_add_gauge_tool`, `pbi_add_labelled_card_tool` now accept an optional `manager` kwarg. The MCP wrappers forward `CONNECTION_MANAGER` automatically, so a typo in a measure name now fails before the layout is ever written instead of producing a "Fix this" placeholder visible only after refreshing PBI Desktop.

### Fixed — error chain preservation

- `error_payload()` now uses `flatten_exception_message()` for non-`PowerBIError` exceptions, so the response message contains the full chain instead of only the topmost frame. Ex.: `RuntimeError("outer") | ValueError("low-level boom")`.
- New `details.cause_chain` field surfaces a structured `[{type, message}, …]` walk of the exception chain (Python `__cause__` / `__context__` plus .NET `InnerException`). Programmatic clients can now identify the underlying error type instead of regex-matching the message.

### Hardened — symlink rejection

- `_reject_symlink_path()` now walks every ancestor directory and rejects the call if *any* of them is a symlink. The previous implementation only rejected the leaf path; this defends against an unlikely TOCTOU edge case on slow filesystems where `path.resolve()` and the subsequent open could observe different targets if a parent symlink was swapped between the two calls. Permission errors on a parent are deferred to the downstream open (we don't leak a misleading symlink-rejection error).

### Tests

- 94 passing, 2 platform-conditional skips (`pythonnet` on non-Windows, symlink creation requires admin on Windows). New cases:
  - `test_field_validation_blocks_missing_measure` — patched `_live_model_field_index` returns a tiny model; valid measure passes, typo raises with the right `details.missing`.
  - `test_field_validation_skipped_without_manager` — no manager → no live check, offline path still works.
  - `test_field_validation_gauge_checks_target_and_color_measure` — gauge validates Y, target_measure, and `fill_color_measure` together.
  - `test_python_cause_chain_is_preserved` — Python `raise … from …` chain flows through the payload.
  - `test_dotnet_inner_exception_traversed` — pythonnet-style `.InnerException` traversed.
  - `test_rejects_symlinked_parent_directory` — symlink in an ancestor path is rejected (skipped on Windows without dev-mode).

## [0.8.1] — 2026-05-07 — Reliability fixes + visual format productisation

### Added — new tools

- `pbi_set_visual_format_property` — generic format setter for an existing visual. Merges properties into `singleVisual.objects[<object_name>][0].properties`, encoding Python values to PBI's canonical literals (single-quoted text, `L` int suffix, `D` decimal suffix, `'#RRGGBB'` solid color, `true`/`false` bool). Optional `property_types` lets callers force a specific encoding (`text`, `bool`, `int`, `decimal`, `color`, or `raw` for pre-shaped expr dicts). Productises the title / axis-label / data-label patches we used to do via direct layout edits.
- `pbi_disable_card_autoscale` — bulk-set `labelDisplayUnits=1` (None) + `labelPrecision` on every card (or a filtered subset) to kill the "119K K €" double-suffix bug when a measure already pre-divides by 1000 with a `K €` format.

### Fixed — connection manager

- `_is_current_state_usable_locked()` now compares the cached `instance.pid` with the PID of the process currently bound to the cached port; if a different msmdsrv process has reused the port (PBI restart on the same port number), the connection is treated as stale and reopened. Helper `_pid_for_port()` added.

### Fixed — relocate tool

- `pbi_relocate_data_source` now validates each rewritten M expression even when `dry_run=True`. The dry-run output now includes a `validation: ok | invalid` field per entry (with `validation_error` on failure), so callers see syntax / security errors during preview instead of only at commit time.

### Added — startup integrity

- New helper `_audit_tool_registry()` in `server.py` runs at server start, walks `tools.__all__`, and confirms every public `pbi_*_tool` is wrapped by an `@mcp.tool()` registration. Logs WARN on orphaned implementations and INFO on wrappers without a matching `_tool` (workflows, etc.). Set `PBI_MCP_STRICT_REGISTRY=1` to make the server fail-fast on drift — useful in CI. Current state: 102/102 tools, 0 orphans.

### Tests

- 88 passing (`test_visuals + test_visual_field_validation + test_workflows + test_quality + test_security + test_query`).
- New `test_visuals` cases: visual format property writes canonical PBI literals (single-quoted text, `L` int, hex auto-detected as solid color); explicit type hint forces decimal encoding; missing visual id surfaces a structured failure; `pbi_disable_card_autoscale` patches only cards (skips charts), supports `visual_ids` filter.

## [0.8.0] — 2026-05-06 — Portable data sources + visual format toolkit

### Added — new tools

- `pbi_relocate_data_source` — bulk-rewrite a hardcoded file/folder path inside every M partition. Accepts `dry_run` for preview, `case_sensitive` for strict matching. Solves the most common collaborator-handoff issue (DataSource.NotFound on file move).
- `pbi_parameterize_data_source` — create a Power Query M parameter (e.g. `SourcePath`) with `meta [IsParameterQuery=true, Type=type text, IsParameterQueryRequired=true]` so the path is editable via PBI Desktop's *Manage parameters* UI, then rewrites every matching M partition to call `File.Contents(<param>)`. After running once, future moves only need a single parameter edit.
- `pbi_add_labelled_card` — composes a textbox label above a card value, matching docx-style "label-on-top" layouts that the native card visual cannot reproduce.
- `pbi_set_column_data_type` — set an existing column's TOM `DataType` (and optionally `FormatString`) when Power Query type hints are overridden by PBI's downstream inference. Accepts standard names: Int64, Decimal, Double, String, DateTime, Boolean, Currency.

### Enhanced

- `pbi_add_gauge` now accepts:
  - `min_value` / `max_value` — gauge axis range bounds (e.g. 0.10–0.20 for a discount % gauge instead of 0–1)
  - `target_value` — target marker as a constant (alternative to `target_measure`)
  - `fill_color` / `target_color` — `'#RRGGBB'` arc + target colors
  - `fill_color_measure` — bind the arc fill to a DAX measure that returns `'#RRGGBB'`. Conditional formatting reacts to slicer / page filter context (overrides `fill_color`).
- `pbi_add_slicer` now accepts `slicer_type="tile"` → emits native list slicer with `general.orientation = 1L` (horizontal tile band).
- `pbi_build_dashboard` `gauge` spec forwards all new range/color keys; new `labelled_card` / `labeled_card` types.

### Fixed

- M expression validator: `each (...)`, `if (...)`, `let`, `then`, etc. were rejected as "function calls outside the local-file allowlist". Added `M_RESERVED_KEYWORDS` exclusion set so legitimate M idioms pass validation. Without this, partition rewrites referencing `each` (very common) failed.
- `_base_visual_config` now deep-merges `extra_single_visual["objects"]` into `singleVisual.objects`, so callers can add formatting alongside the title without clobbering it.

### Tests

- 83 passing across `test_visuals + test_visual_field_validation + test_workflows + test_quality + test_security + test_query`.
- New cases: gauge axis literals + dataPoint colors, gauge color validation, gauge fill_color_measure binding, gauge measure-binding overrides static fill, tile slicer orientation, slicer type validation, labelled card composition, label-height validation.

## [Unreleased]

### Added

- `pbi_import_excel_workbook` - explicit one-call Excel workbook import tool for Power BI tables.
- `pbi_model_audit_workflow`, `pbi_excel_import_workflow`, `pbi_measure_workflow` - guided workflow tools with dry-run defaults for higher-productivity LLM agents.
- `pbi_validate_report_fields` and `pbi_repair_report_fields` - detect and repair broken report visual field bindings that cause Power BI "Fix this" placeholders.
- `pbi_list_tmdl_files`, `pbi_read_tmdl_file`, `pbi_write_tmdl_file` - offline PBIP/TMDL semantic model file helpers for projects that should be edited without launching Power BI Desktop.
- `pbi_patch_tmdl_measure` - create or replace a measure block in a table TMDL file without rewriting the whole file manually.
- `pbi_create_persistent_report` - optional persistent PBIX builder for DataModel tables, DAX measures, relationships, pages, and native visuals.
- `pbi_audit_model`, `pbi_lint_dax`, `pbi_lint_report_layout`, `pbi_validate_visual_bindings`, `pbi_score_dashboard`, `pbi_run_scenario`, `pbi_compare_report_versions` - QA gates for MCP-generated Power BI reports.
- `pbi_detect_name_collisions`, `pbi_detect_dirty_dates`, `pbi_validate_relationship_plan` - enterprise-grade preflight checks for naming collisions, dirty text dates, and unsafe relationship plans.
- `pbi_detect_empty_visuals`, `pbi_export_validation_report` - report QA probes for visuals that render no rows and reusable JSON validation artifacts.
- `pbi_detect_empty_visuals` now accepts an optional DAX `filter_expression`, and validation report exports include an `overall_valid` summary with total issue and warning counts.
- `pbi_validate_filter_expression` - validates a DAX boolean filter before filtered visual probes run.
- `pbi_generate_measure_tests` - runs smoke tests for measures and reports execution errors, blanks, zeros, unsafe division operators, and format mismatches.
- `pbi_validate_pbix_persistence` - checks patched PBIX ZIP/Layout persistence, visual/page counts, and stale `SecurityBindings`.
- `pbi_validate_pbix_reopen` - opens a PBIX in Power BI Desktop, optionally captures/analyzes a screenshot, and scans UI Automation text for repair-error signals.
- `pbi_validate_pbix_reopen` can optionally run Windows OCR over the screenshot to catch canvas text that UI Automation misses.
- `pbi_patch_layout` now enables `fail_on_persistence_risk=True` by default so PBIX layout patches are blocked when visual bindings rely on live-model metadata missing from the extract.

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
- `pbi_patch_layout`: direct PBIX Layout patch tool that swaps `Report/Layout`, removes `SecurityBindings`, preserves ZIP entry metadata, and supports `force=True` graceful save/close before Windows auto-kill fallback.
- `pbi_apply_design`: one-shot design preset tool that writes the base theme, updates page backgrounds, and applies card container styling.

### Fixed
- Visual query generation now resolves measure home tables from extract metadata (`Model/tables/*/measures/*.dax`) so `prototypeQuery.From[]` uses the real table entity instead of `"$Measures"` when available.
- Measure fallback now logs a warning and uses `"$Measures"` only when extract metadata is missing or disconnected.
- `.dax` import parsing now strips `//` and `/* ... */` comments safely while preserving quoted strings and measure block boundaries.
- `pbi_compile_report` now accepts `force: bool = False` and can gracefully save/close Power BI Desktop before write operations, with auto-kill fallback.
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
