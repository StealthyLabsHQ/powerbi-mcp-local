# Changelog

## [0.10.8] — 2026-05-08 — Refactor phase 3: bindings, home_tables, containers extracted

Continuation of the visuals/ fan-out. Three more focused submodules pulled out of `src/tools/visuals/__init__.py`. No behavior change, no tool surface change. 167/167 tests, registry 131/131 clean.

### Changed

3 new submodules extracted from `src/tools/visuals/__init__.py`:

1. **`_home_tables.py`** (173 L) — measure → home table resolution: `_scan_measure_home_tables`, `_resolve_measure_home_map`, `_augment_measure_home_map_with_live`, `_inspect_value_measures`, `_persistence_risks`. Lazy-imports `pbi_model_info_tool` through the package re-export so existing test patches against `tools.visuals.pbi_model_info_tool` keep working.

2. **`_bindings.py`** (499 L) — visual binding builders + validators: `_build_select_entry`, `_build_prototype_query`, `_select_name_map`, `_from_entity_by_alias`, `_next_alias`, `_sync_container_query`, `_validate_projection_roles`, `_validate_field_references_live`, `_live_model_field_index`, `_visual_binding_issues`, `_scan_visual_bindings`, `_assert_container_bindings`. Same patchable-lookup pattern: `_validate_field_references_live` and `_validate_projection_roles` resolve `_live_model_field_index` through the package so tests patching `tools.visuals._live_model_field_index` keep working.

3. **`_containers.py`** (236 L) — visual container builders: `_unique_visual_id`, `_validate_dimensions`, `_page_next_z`, `_base_visual_config`, `_make_visual_container`, `_visual_payload`, `_find_visual`, `_append_visual`, `_create_chart_container`.

### Internals

- `tools/visuals/__init__.py`: 3248 → 2440 lines (-808 across phases 2 + 3).
- 8 submodules total (`_base`, `_paths`, `_refs`, `_layout`, `_formatting`, `_home_tables`, `_bindings`, `_containers`).
- Registry audit clean: 131/131, 0 orphans.

### Tests

- 167 passing, 2 platform skips. No test changes required.

### Deferred to v0.10.9+

Still in the monolith: I/O block (`_extract_pbix_zip_natively`, `pbi_extract_report_tool`, `pbi_compile_report_tool`, PowerShell helpers, `_run_pbi_tools`, `_find_pbi_tools`), per-domain tool wrappers (`charts.py`, `cards.py`, `structure.py`, `pages.py`, `ops.py`, `design.py`, `dispatcher.py`), and the `server.py` wrapper split into `src/wrappers/`.

## [0.10.7] — 2026-05-08 — Refactor phase 2: visuals/ submodule fan-out (5 modules extracted)

Continuation of the file-split refactor started in v0.10.6. No behavior change, no tool surface change. Breaks the visuals package into focused submodules so future edits target the relevant concern rather than scrolling a 3500-line monolith.

### Changed

Five submodules extracted from `src/tools/visuals/__init__.py`:

1. **`_base.py`** (92 L) — constants (`DEFAULT_PAGE_WIDTH/HEIGHT`, `LAYOUT_RELATIVE_PATH`, `THEMES_RELATIVE_DIR`, `MODEL_TABLES_RELATIVE_DIR`, `HEX_COLOR_RE`, `DEFAULT_VISUAL_SIZES`, `VISUAL_FIELD_ROLES`, `VISUAL_ROLE_KINDS`) and error classes (`VisualToolError`, `PBIToolsNotInstalledError`, `ReportLayoutError`, `PageNotFoundError`, `VisualNotFoundError`). Bottom of the import graph — no internal deps.

2. **`_paths.py`** (25 L) — path resolution helpers (`_resolve_pbix_path`, `_resolve_extract_folder`, `_resolve_theme_path`, `_layout_path`).

3. **`_refs.py`** (61 L) — field reference normalisation (`_BRACKET_REF_RE`, `_normalize_reference`, `_split_column_ref`, `_query_ref`). Accepts `Table.Column`, `Table[Column]`, `'Table'[Column]`, bare measure.

4. **`_layout.py`** (168 L) — layout I/O: `_load_layout`, atomic `_save_layout` with `.bak` fallback (v0.10.5), `dry_run_layout_writes()` context manager (v0.10.5), embedded JSON helpers, page utilities (`_find_page`, `_normalize_page_name`, `_next_page_name`, `_page_summary`).

5. **`_formatting.py`** (135 L) — Power BI visual literal + format-property encoders: `_literal_value`, `_decimal_literal`, `_int_literal`, `_text_literal`, `_solid_color`, `_gauge_axis_objects`, `_datapoint_fill_objects`, `_title_objects`, `_encode_visual_format_value`, `_VISUAL_FORMAT_TYPES`.

The package `__init__.py` re-exports every name to preserve back-compat — every existing `from tools.visuals import …` keeps working unchanged.

### Internals

- `tools/visuals/__init__.py`: 3607 → 3248 lines (-359, after also dropping unused `re`, `threading`, `contextmanager`, `Iterator` imports).
- 5 new submodules totalling 481 lines.
- Registry audit clean: 131/131, 0 orphans.

### Tests

- 167 passing, 2 platform skips. No test changes required.

### Deferred to v0.10.8+

Heavier extractions still in the monolith: bindings (`_build_select_entry`, `_build_prototype_query`, `_validate_field_references_live`, `_validate_projection_roles`, `_scan_visual_bindings`), containers (`_create_chart_container`, `_make_visual_container`, `_base_visual_config`, `_append_visual`), home tables (`_scan_measure_home_tables`, `_resolve_measure_home_map`, `_inspect_value_measures`), I/O (`pbi_extract_report`, `pbi_compile_report`, PowerShell helpers), and per-domain tool wrappers (`charts.py`, `cards.py`, `structure.py`, `pages.py`, `ops.py`, `design.py`, `dispatcher.py`).

## [0.10.6] — 2026-05-08 — Refactor phase 1: package skeleton + mcp_core extraction

Foundational refactor (no behavior change, no tool surface change). Splits the two largest modules into a package + a runtime core, preparing the codebase for incremental fan-out of the visual / wrapper sub-modules in later releases.

### Changed

1. **`src/tools/visuals.py` → `src/tools/visuals/__init__.py`** (package conversion). Every existing import (`from tools.visuals import …`) keeps working unchanged — the file is now an `__init__.py` that exports the same surface. Subsequent releases will extract layout / refs / bindings / containers / formatting / charts / cards / structure / dispatcher into sibling submodules without breaking callers. The internal `Path(__file__).resolve().parents[2]` was updated to `parents[3]` to keep the bundled `tools-bin/pbi-tools.core.exe` resolvable.

2. **`src/mcp_core.py` extracted from `src/server.py`**. The FastMCP instance, the `CONNECTION_MANAGER`, the `_run` helper, the PID lock + parent-watcher lifecycle (v0.10.3), the registry audit (v0.10.3), and the profile filter (v0.10.5) now live in this dedicated runtime module. `server.py` keeps only `main()`, CLI argparsing, and the ~131 `@mcp.tool()` wrappers. Removed unused imports (`atexit`, `signal`, `sys`, `tempfile`, `threading`, `time`, `FastMCP`, `PowerBIConnectionManager`, `error_payload`) — they're now in `mcp_core`.

### Internals

- `server.py`: 3500 → 3215 lines (-285).
- `mcp_core.py`: new, 249 lines.
- `tools/visuals/`: package directory.
- Registry audit clean: 131 wrapped, 131 implementations, 0 orphans.

### Tests

- 167 passing, 2 platform skips. No test changes required.

### Deferred to v0.10.7+

Per-domain wrapper split of `server.py` (`src/wrappers/measures.py`, `wrappers/visuals.py`, …) and per-concern submodule extraction of `tools/visuals/` (`_layout.py`, `_refs.py`, `_bindings.py`, `_containers.py`, …). Each can be done incrementally without breaking external imports.

## [0.10.5] — 2026-05-08 — Atomicity, dry-run, persistence warnings, grading profile

Foundational reliability improvements addressing the four highest-ROI structural gaps identified in the v0.10.4 retrospective. No tool removals; new optional `dry_run` parameter on the generic visual dispatcher; existing signatures unchanged.

### Fixed

1. **Atomic layout writes with `.bak` recovery.** `_save_layout` now serialises to `Layout.tmp.<pid>`, copies the previous-good content to `Layout.bak`, then `os.replace`s the temp file onto `Layout` — atomic on Windows + POSIX. A crash mid-write leaves the original Layout intact and the previous version available at `Layout.bak`. Previously, an interrupted write could corrupt the Layout JSON and break the entire `.pbix`.

2. **Persistence warning on every TOM mutation.** `PowerBIConnectionManager.execute_write()` injects a `persistence: {scope: "memory_only", hint: "..."}` field into every write payload. The hint explicitly states that the change committed to the AS engine in memory and that the `.pbix` on disk is unchanged until Power BI Desktop saves the file. All measure / table / column / relationship / role / calc-group / Power Query / TMDL writes now propagate this warning to the response. Eliminates the silent-loss failure mode where a user closes Power BI Desktop without realising changes weren't persisted.

### Added

3. **Dry-run scaffolding for visual writes.** New `dry_run_layout_writes()` context manager + thread-local flag. While active, `_save_layout` records a per-write log entry (`{folder, section_count, visual_count}`) instead of writing to disk. The generic `pbi_add_visual` MCP tool gained a `dry_run: bool = False` parameter — when True, all validation, binding resolution, and home-table lookups still run, but the layout is left untouched. The response carries `dry_run=True`, a `write_log`, and a `[dry-run]`-prefixed message. Use it to preview a proposed change before committing.

4. **`grading` profile.** New `--profile=grading` exposes a tightly-scoped 25-tool surface (vs. ~130 in `all`): connect / list / model_info / page reads, DAX validation, all six v0.10.4 analysis tools, lint, audit. Drops every visual writer, calc group, role, Power Query mutator, Excel write, and persistence tool. Drastically reduces LLM tool-selection noise during evaluation workflows. Existing `readonly` / `write` / `all` profiles unchanged.

### Internals

- `READ_TOOLS` extended with the v0.10.4 analysis tools + `pbi_describe_page`, `pbi_system_health`, `pbi_operation_history` (these were always read-only but missing from the readonly profile).
- New `GRADING_TOOLS` set in `src/security.py`.
- New imports in `src/tools/visuals.py`: `threading`, `Iterator`, `contextmanager`. New module-level `_LAYOUT_WRITE_TL` thread-local.
- `pbi_add_visual_tool` signature gained `dry_run: bool = False` (kw-only). Existing callers unaffected.

### Tests

- 167 passing (was 162) — 5 new offline tests covering atomic write + `.bak` creation, no-temp-leak after exception, dry-run interception, dry-run context reset, and `execute_write` persistence injection. 2 platform skips unchanged.

### Still pending (deferred to v0.10.6)

- `pbi_persist_now()` — would-be replacement for manual Ctrl+S. Requires UI automation (`pywinauto` or `SendKeys`) which is fragile on Windows; will be opt-in via env var or a separate package extra.
- `dry_run` propagation to per-type visual tools (`pbi_add_line_chart`, `pbi_add_bar_chart`, …). Currently only the generic `pbi_add_visual` dispatcher exposes it. The plumbing is centralised so this is a fan-out, not a redesign.
- `pbi_update_visual_bindings` — patch projections without remove + recreate.

## [0.10.4] — 2026-05-08 — Analysis & scoring tools, line-chart pre-flight, column qualification docs

Six new analysis tools for grading and rubric scoring of Power BI deliverables, plus a pre-flight diagnostic for the line-chart constant-measure failure mode and clarified column-qualification docs across the visual surface. No tool removals; existing signatures unchanged.

### Added

- **`pbi_validate_star_schema`** — classifies tables into fact / dim / bridge / isolated based on relationship topology, flags snowflake (dim-to-dim) chains and fact-to-fact joins as issues, surfaces multiple-fact and isolated-table warnings. Optional `fact_table_hints` lets graders tag fact tables that have no incoming relationships yet.
- **`pbi_detect_circular_dependencies`** — builds a measure dependency graph from DAX `[Name]` tokens, runs DFS to find cycles, reports self-references separately. Returns the actual cycle paths so the user can break the loop.
- **`pbi_validate_power_query_steps`** — checks an M expression for required step patterns. Each entry is a substring or `re:` regex; useful for grading transformations like `Text.PadStart(_, 5, "0")` or null-customer filtering.
- **`pbi_detect_missing_visuals`** — scans a page's layout for required visuals. Each requirement is `{visual_type, count?, contains_field?, label?}`. Catches missing carte géographique, missing tranches-de-factures visual, etc.
- **`pbi_score_rubric`** — weighted aggregator that runs the above validators (`star_schema`, `no_circular_deps`, `power_query_steps`, `missing_visuals`, `measure_exists`) plus per-criterion weights and returns a normalised score in [0, 1].
- **`pbi_export_correction_report`** — generates a Markdown grading report (model overview, star-schema verdict, cycles, audit issues, optional rubric scoring with checkmarks).

### Fixed

- **Line chart now warns on constant measures and unresolved home tables.** New `_inspect_value_measures` helper checks each Y-measure of `pbi_add_line_chart_tool`. When a measure has no DAX column/measure reference (looks like a scalar constant such as `Ratio = 0.92`) or when its home table cannot be resolved (binding falls back to the synthetic `$Measures` entity), the response carries a `warnings` list with a `hint` explaining the fix. Reproduces the cause of the "constant measure breaks line chart" symptom even when full PBI repro isn't available — the linked root cause (model persistence) is tracked for v0.10.5.

### Changed

- **Column qualification documented in tool descriptions.** The MCP wrappers for `pbi_add_line_chart`, `pbi_add_bar_chart`, `pbi_add_donut_chart`, `pbi_add_table_visual`, `pbi_add_scatter_chart`, `pbi_add_waterfall`, and `pbi_add_slicer` now state explicitly that columns require table qualification (`Table[Column]`, `'Table'[Column]`, or `Table.Column`). Bare names like `Year` are rejected — the LLM no longer needs to discover this from an error message.

### Internals

- 6 new tools registered via `tools/__init__.py` and `src/server.py`. Registry audit clean.
- New helper `_inspect_value_measures(value_measures, measure_home_map, manager)` in `src/tools/visuals.py`.
- New file `tests/test_v0_10_4_analysis.py` with 11 offline unit tests (mock manager + patched snapshot, no live PBI required).

### Tests

- 162 passing (was 151 in v0.10.3) — 11 new offline tests for star-schema topology, cycle detection, power-query step matching, missing-visual detection, rubric aggregation, and correction-report writing. Two platform skips unchanged.

### Known limitations (not yet fixed)

- Measures created via `pbi_create_measure` live only in the AS in-memory model, not in the on-disk `.pbix`. Restarting Power BI Desktop drops them. A `pbi_patch_model` writing TMDL to `DataModel/` is tracked for v0.10.5.
- No `pbi_save_model` (forced Ctrl+S) — requires UI automation, opt-in only.
- Visual binding updates still require remove + re-add (no `pbi_update_visual_bindings` yet).

## [0.10.3] — 2026-05-07 — Stability: single-instance lock, parent watcher, instance discovery cache

Server-side reliability and cold-start performance. Eliminates the multi-instance / zombie-process failure mode and removes redundant work from the connection bootstrap. No tool surface changes.

### Fixed

1. **Single-instance enforcement at startup.** `main()` now writes a PID lock to `%TEMP%/powerbi-mcp.pid` (or `/tmp` on POSIX). When a second server starts and the recorded PID is still alive, the older process is killed via `psutil.Process.kill()` before the new one claims the lock. `atexit` plus `SIGINT`/`SIGTERM` handlers remove the lock on clean shutdown. Resolves the conflict where two `python.exe` instances fought for the stdio transport when an LLM CLI (Claude Code, Codex, …) restarted without reaping its child.

2. **Parent-process watcher.** A daemon thread polls the parent PID every 2 seconds via `psutil`. If the parent disappears or becomes a zombie, the server releases the PID lock and exits with `os._exit(0)`. Handles the abnormal-termination case where the parent dies without sending SIGTERM and without closing the stdio pipe — previously the server would linger as a zombie until reboot.

### Performance

3. **Instance discovery cache (TTL 5 s).** `PowerBIConnectionManager._discover_instances()` now memoises results for 5 seconds. Tool calls that resolve through `connect()` / `list_instances()` / `_select_instance()` no longer trigger a full filesystem + process rescan on every invocation. Cache is invalidated on `_disconnect_locked()` so a Power BI Desktop restart is picked up within one TTL window.

4. **Bounded workspace glob.** `_discover_workspace_instances()` no longer calls `Path.rglob`. New `_bounded_glob(root, name, max_depth=5)` walks the workspace tree breadth-limited via `iterdir`, avoiding the recursive scan over Packages/ hierarchies that contained 100 + GUID directories.

5. **Lazy `psutil.process_iter` scan.** `_discover_instances()` skips `_discover_process_instances()` when every workspace-located instance already carries a `port_file` (which proves the process is alive). PID enrichment is no longer the critical path for connecting; `psutil.process_iter` only runs as a fallback when workspace discovery is incomplete.

6. **Tool-registry audit gated by env var.** `_audit_tool_registry()` previously ran on every server start, walking `mcp._tool_manager` introspectively. It is now opt-in via `PBI_MCP_AUDIT=1` (or the existing `PBI_MCP_STRICT_REGISTRY=1` for CI). Production servers skip the introspection cost.

### Internals

- New module-level helpers in `src/server.py`: `_PID_LOCK_PATH`, `_release_pid_lock`, `_pid_alive`, `_acquire_single_instance_lock`, `_start_parent_watcher`. Plumbed into the stdio branch of `main()` only — SSE transport unchanged (port binding already provides single-instance semantics).
- New field `PowerBIConnectionManager._instance_cache: tuple[float, list[DiscoveredInstance]] | None` and helper `_bounded_glob` (static).
- Imports added: `atexit`, `signal`, `sys`, `tempfile`, `threading`, `time` in `server.py`. No new third-party dependencies — `psutil` was already required on Windows.

### Validated

- Local harness exercises the MCP `initialize` handshake (returns `protocolVersion`), the single-instance kill-and-claim path (second spawn terminates first within ~2 s), and the EOF cleanup path (closing parent stdin removes the PID file within 5 s). All three pass.

## [0.10.2] — 2026-05-07 — Field validation, manager propagation, and DAX template hardening

Six follow-up bugs surfaced by a real-world v0.10.1 test pass on a 7-page / 62-visual report. All fixes are localised, signature-compatible, and covered by regression tests. Registry: 125/125, 0 orphans · tests: 112/112 (2 platform skips).

### Fixed

1. **`pbi_add_visual` now forwards the connection manager to every dispatcher.** The generic dispatcher injects the active manager into the dispatch ``cfg`` (under the reserved ``__manager__`` key), and each ``_dispatch_*`` function passes it to the underlying ``pbi_add_*_tool``. Map / scatter / combo / kpi / matrix visuals created via the generic dispatcher now go through live field validation and home-table resolution — eliminating the ``measure_home_table_needs_repair`` follow-up signal previously seen on map writes.

2. **Missing-reference errors now carry a `hint` plus a `did_you_mean` list.** `_validate_field_references_live` runs `difflib.get_close_matches` over the live model's measure / column lists and surfaces up to 5 nearest names. Each missing entry includes a role-aware `hint` (e.g. "axis/category/rows expect a column — qualify with the table; try one of: Date.Year, Date.Year-Month."). The error payload also reports `available_measure_count` / `available_column_count` for context.

3. **Every `pbi_add_*_tool` now passes `expected_kinds` to the validator.** Each visual builds a `{reference: "column" | "measure"}` map matching its role schema (Category/Series/Rows → column, Y/Values/Indicator/Goal → measure, Slicer → column, Map.Location → column, KPI.TrendLine → column). When a user passes a bare `Year` to a Category role, the validator now reports `kind="column"` with a hint listing qualified candidates — instead of silently inferring `kind="measure"` from the format.

4. **Bare references can no longer satisfy a column-expecting role.** Even if a column with the same short name exists somewhere in the model, a `Year`-style bare reference cannot fill an `axis/category/rows` slot because the layout writer needs the table prefix. The validator now flags every bare reference passed to a column role and lists every `Table.Column` candidate that resolves to that short name.

5. **`pbi_validate_dax_semantic_tool` no longer overwrites its `kind` parameter.** A renamed local variable (`ref_kind`) inside the unknown-references loop replaces the previous `kind` shadowing that corrupted the runtime probe call. The tool now correctly returns `{semantic: {unknown_references: […]}}` instead of failing with `"kind must be 'scalar' or 'table'"`.

6. **DAX templates always quote table names** so reserved-word collisions (`Date`, `Time`, `Year`, …) work. Time-intelligence (`YTD`/`MTD`/`QTD`/`SPY`/`MA3`), variance, contribution, top-N, and rolling-average templates now route table names through `_dax_column_ref(table, column)` which emits `'Table'[Column]` (with embedded single quotes doubled per the DAX grammar). Generated DAX is valid against models with date tables literally named `Date`.

### Internals

- New helpers `_dax_table_ref` / `_dax_column_ref` in `src/tools/measures.py` for quoted DAX column references.
- New reserved key `__manager__` in `pbi_add_visual_tool` cfg propagates the live manager to every per-type dispatcher.

### Tests

- 112 passing (2 platform skips). Six new regression cases in `tests/test_security.py` covering: kind shadowing, quoted DAX templates, embedded-quote escaping, hint+did_you_mean enrichment, role-aware column-vs-measure error reporting, and manager propagation through `pbi_add_visual_tool`.

## [0.10.1] — 2026-05-07 — bugfixes (extract fallback, ref parsing, map dispatch, home tables, FORMAT detection, role hints, lint knobs)

Seven concrete bugs surfaced while building the Power BI report. All seven fixes are localised, signature-compatible, and covered by regression tests in `tests/test_security.py`. Net new tool: `pbi_add_map`. Registry: 125/125, 0 orphans · tests: 107/107 (2 platform skips).

### Fixed

1. **`pbi_extract_report` falls back to native ZIP extraction when the bundled `pbi-tools.core` does not support the `extract` action** (it ships `compile` only). The CLI's "Unknown action: 'extract'" / "No action was specified" stdout is detected, the wrapper logs the fallback, and a fresh ZIP-based extraction unpacks `Report/Layout` and `Report/StaticResources/Themes/*` so layout-touching workflows keep functioning. The response carries `extraction_method` so callers know which path ran.

2. **Visual tools accept `Date[Année]`, `'Date'[Année]`, and `Date.Année` interchangeably.** New `_normalize_reference` collapses every PBI reference syntax into the canonical `Table.Column` form before downstream parsing. The dotted form remains unchanged for backward compatibility, and bare measure names stay bare. `_query_ref` and `_split_column_ref` route through the normaliser; user-facing error messages now list the three accepted formats.

3. **`pbi_add_visual` supports `visual_type="map"`** via a new `_dispatch_map` registered in `_VISUAL_TYPE_DISPATCH`. New top-level tool `pbi_add_map` (with manager-aware live field validation) is also exposed for callers that prefer the explicit form. Achieves API parity with the `pbi_build_dashboard` `map` spec.

4. **Visual writes already carry the right measure home table**, so `pbi_validate_report_fields` no longer reports `measure_home_table_needs_repair` after a successful `pbi_add_*`. New helper `_resolve_measure_home_map(extract_folder, manager=…)` combines on-disk PBIP metadata with the live model so the Entity reference is correct from the first write. The 14 visual tools that take `manager` use this path; on-disk metadata still wins on conflict for stable offline behaviour.

5. **`pbi_detect_empty_visuals` no longer flags `FORMAT()` text-returning measures as numeric zero.** The "all-zero" warning fires only when every non-blank value is numeric (`int`/`float`, with `bool` excluded) AND every numeric value equals zero. Mixed/text-only payloads stay silent.

6. **Field-validation errors report the role-expected kind, not just the format-inferred one.** `_validate_field_references_live` now accepts an optional `expected_kinds` map (`reference -> "column" | "measure"`); each missing entry surfaces both `kind` (the role expectation) and `inferred_kind` (the format guess), plus a `hint` explaining how to qualify the reference (e.g. "axis/category/rows expect a column — qualify with the table"). Diagnostic messages match the visual's role contract.

7. **`pbi_lint_report_layout` accepts `ignore_warnings`, `only_pages`, and `max_visuals_per_page`.** `ignore_warnings` drops listed warning types (`too_many_visuals`, `visual_too_small`, `missing_title`, `excessive_whitespace`, `layout_overloaded`) on intentionally dense pages. `only_pages` scopes the lint to specific pages (e.g. only the new pages an LLM produced). `max_visuals_per_page` overrides the default `too_many_visuals` threshold. Issues are never silenced — only warnings.

### Added

- `pbi_add_map` — top-level wrapper for the existing dispatcher path (Bug 3 productisation). Accepts the same flexible reference syntax (`Table[Column]`, `'Table'[Column]`, `Table.Column`).

### Tests

- Seven regression tests added to `tests/test_security.py` (tracked) — one per bug. The visual-side tests live in `tests/test_visuals.py` (local-only by repo convention). Suite stays 100 % green: 107 passing, 2 platform skips.

## [0.10.0] — 2026-05-07 — Stability foundation, LLM-resilient visual surface, advanced visuals, DAX power tools

124 tools registered, 0 orphans · 110/110 tests passing · 17 new tools, 6 hardened, 4 new DAX scaffolding helpers.

### Pillar A — Stability foundation

- `pbi_system_health` — single-call diagnostic. Returns `connected`, `port`, `port_open`, `pid_match`, `tom_available`, `adomd_available`, `model_loaded`, `model_name`, `table_count`, `measure_count`, `cache.{write_generation, entries}`, `last_operation_ts`, `dependencies` (mcp / pythonnet / pyadomd / pbi_pyadomd). Read-only and safe to call without an active connection — useful as a one-stop health check before any LLM agent attempts a write.
- `pbi_operation_history` — exposes a 50-entry ring buffer of recent ops recorded inside `PowerBIConnectionManager` (`{ts, op, kind: read|write, duration_ms, ok, error_type?, error_code?, error_message?}`). `flatten_exception_message` is used so the cause chain travels with each failure entry. Use after a failure to see what already landed.
- Idempotency standardisation: `pbi_create_relationship` and `pbi_add_role_member` now accept `overwrite=False`. With `overwrite=True`, existing endpoints/members are updated in place and the response carries `action="updated"` instead of raising `PowerBIDuplicateError`. `pbi_create_table_tool`, `pbi_create_column_tool`, `pbi_create_role_tool`, `pbi_create_calc_group_tool` already had this — surface aligned.
- `pbi_create_measures` now accepts `dry_run=False`. With `dry_run=True`, every measure is name + expression validated and a per-item `planned_action` is reported (`would_create` / `would_update` / `would_fail`). No model mutation. Mirrors the pattern shipped earlier on `pbi_relocate_data_source`.

### Pillar B — LLM-resilient visual surface

- Field-existence validation extended from 3 to **9** add_* tools: `pbi_add_card`, `pbi_add_gauge`, `pbi_add_labelled_card` (already shipped) plus `pbi_add_bar_chart`, `pbi_add_line_chart`, `pbi_add_donut_chart`, `pbi_add_table_visual`, `pbi_add_waterfall`, `pbi_add_slicer`. Server wrappers forward `CONNECTION_MANAGER` automatically — typos fail fast before the layout is written.
- New pre-flight projection role validator `_validate_projection_roles(visual_type, projections, *, manager=None)` wired into `_create_chart_container`. Two checks: (1) every role is in `VISUAL_FIELD_ROLES[visual_type]`; (2) when a connection manager is passed, each reference's resolved kind matches the new `VISUAL_ROLE_KINDS` table (e.g. measure-in-Category mistakes are rejected at tool-call time). Catches the most common LLM mistake when constructing visuals.
- `pbi_describe_page` — read-only structured snapshot of a page. Per-visual `id`, `type`, `position`, `bindings` (role → list of refs), `formatting` (title, axis titles, label_display_units), and a `binding_health` rollup (`ok` | `missing_field` | `wrong_role` | `issues`). Lets an LLM introspect what's on the page without parsing layout JSON.
- `pbi_auto_grid_layout` — pure utility. Positions a list of visual specs on an N-column grid with configurable padding, supports `col_span`/`row_span`, returns each spec annotated with `x`/`y`/`width`/`height`. No live model touch — saves LLMs from doing arithmetic and prevents overlap.
- `pbi_convert_visual_type` — migrate an existing visual to a different type while preserving compatible bindings. Compatibility groups: `card↔kpi`, `clusteredBarChart↔clusteredColumnChart↔lineChart↔lineClusteredColumnComboChart`, `donutChart↔treemap`. Incompatible source/target combinations are rejected with a structured `details.reason="incompatible"`.

### Pillar C — Advanced visual types

- `pbi_add_scatter_chart` — `scatterChart`. Roles: Category (column), X (measure), Y (measure), Size (measure, optional), Series (column, optional). For correlation analysis between two measures grouped by a dimension.
- `pbi_add_combo_chart` — `lineClusteredColumnComboChart`. Roles: Category, Y (bar measures, list), Y2 (line measures, list). Use for actual-vs-target dashboards.
- `pbi_add_kpi` — native `kpi` visual. Roles: Indicator (measure), TrendLine (column), Goal (measure, optional). `direction="high_is_good"|"low_is_good"` controls the status colour interpretation.
- `pbi_add_matrix` — `pivotTable`. Roles: Rows, Columns, Values. `column_layout="stepped"|"tabular"`, `subtotals=True|False`. Matches the docx-style multi-dim table common in business reports.

### Pillar D — DAX power tools

- `pbi_create_time_intelligence_pack` — batch creates a family of measures (default: YTD, MTD, QTD, SPY, YOY, YOY %, MA3) from one base measure. Dependency-aware: `YOY%` auto-pulls `YOY` and `SPY`. Supports `dry_run=True` for a no-mutation preview, `format_inherit=True` to inherit the base measure's format, and per-pattern format overrides (e.g. YOY % defaults to `0.00%`).
- Per-pattern wrappers: `pbi_create_ytd_measure`, `pbi_create_mtd_measure`, `pbi_create_spy_measure`, `pbi_create_yoy_measure` — same engine, single pattern when only one is needed.
- `pbi_create_variance_measure`, `pbi_create_contribution_measure`, `pbi_create_topn_measure`, `pbi_create_rolling_average_measure` — canonical DAX templates for the four most-asked analytics patterns. Each generates a measure with the right description, display folder, and (when supplied) format string.
- `pbi_apply_format_preset` + `pbi_list_format_presets` — preset library covering `currency_eur`/`currency_eur_k`/`currency_eur_m`, `currency_usd` family, `percent`/`percent_0dp`/`percent_1dp`/`percent_2dp`/`percent_4dp`, `thousands`, `millions`, `decimal_2`, `integer`, `integer_no_sep`, `date_iso`, `date_short_fr`, `date_short_us`, `date_long_fr`, `datetime_iso`. New module `src/tools/formats.py`.
- `pbi_validate_dax_semantic` — three-layer DAX validation. (1) References: parses `Table[Column]` and bare `[Measure]` tokens, checks each against the live model index, surfaces unknown references in `semantic.unknown_references`. (2) Format compatibility heuristic (best-effort, never blocks) — flags percent format on a money expression or currency on a ratio. (3) Runtime probe: delegates to `pbi_validate_dax`. Returns `{valid, syntax, semantic, runtime_error?}`.
- `pbi_generate_dax_context_prompt` — compact markdown snapshot of the model (tables, columns, measures, relationships) ready to paste into an LLM system prompt. `include_dax`/`include_relationships`/`max_chars` knobs. Truncation respects line boundaries with a clear notice.

### Tool registry

`_audit_tool_registry()` confirms 124/124 wrappers (was 102 in v0.8.2, +22 net new tools). 0 orphans on every pillar.

### Tests

110 passing across `test_visuals + test_security + test_workflows + test_quality + test_query + test_visual_field_validation` (2 platform-conditional skips). New cases:
- Stability: ring buffer rolls over at 50, returns newest-first; system health works disconnected; time-intel template renders the canonical strings; dependency expansion adds SPY/YOY automatically when YOY% is requested.
- Visuals: auto-grid places specs on 3-column grid with correct spacing, honours col_span; convert card → kpi preserves bindings; donut → kpi rejected as `incompatible`; describe_page returns structured visuals with positions, bindings, binding_health; projection role validator rejects unknown role and (with manager) measure-in-Category kind mismatch.
- DAX: format preset catalogue and lookup; semantic reference parser extracts `Table[Column]` and bare `[Measure]` tokens correctly.

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
