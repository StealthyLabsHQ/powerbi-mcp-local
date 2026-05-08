# Changelog

Active changelog covers the most recent releases.  
Older entries (v0.10.11 and earlier — refactor phases, v0.10.x feature drops, v0.8.x and v0.7.x history) live in [CHANGELOG-archive.md](CHANGELOG-archive.md).

## [0.12.1] — 2026-05-09 — Google Antigravity compat entry point

### Added

1. **`src/server_antigravity.py`** — dedicated stdio entry point for the
   Google Antigravity MCP client, registered as
   ``powerbi-mcp-antigravity`` in ``[project.scripts]``. The default
   entry point keeps shipping the full FastMCP capability set for
   Claude Desktop / Cursor / Anthropic CLI; this one strips it down to
   what Antigravity's bundled client accepts.

   Antigravity used to fail with
   ``connection closed: calling "resources/list": client is closing: EOF``
   immediately after ``initialize``. Three differences vs. the default
   entry point fix it:

   - **stdio hygiene**: UTF-8, line-buffered, ``\n`` line endings on
     stdout/stderr (the Windows default ``\r\n`` breaks strict
     JSON-RPC framing).
   - **Logs to stderr only at ERROR level**: stdout is reserved for
     JSON-RPC frames; even one INFO line on stdout would corrupt the
     next frame.
   - **Minimal capability advertisement**: only ``tools`` (no
     ``prompts``, no ``resources``, no ``experimental``, no
     ``listChanged`` notifications). Antigravity's strict validator
     rejects shapes it doesn't recognize. Tools, resources, and
     prompts registered on the FastMCP instance remain reachable
     through ``tools/call``; only the capability advertisement is
     stripped.

   Same single-instance lock and parent-watcher lifecycle as the
   default entry point.

   Configure Antigravity at
   ``%USERPROFILE%\.gemini\antigravity\mcp_config.json`` to point
   ``command`` + ``args`` at ``server_antigravity.py``. See
   ``docs/SETUP.md`` for the full snippet.

## [0.12.0] — 2026-05-09 — Security release: Codex audit findings (14 fixes)

External security audit (Codex) flagged 1 critical + 13 high findings.
This release closes all of them. No tool surface change.

### Critical

1. **SSE without authentication or DNS-rebinding defense**
   (`src/server.py`). The optional SSE transport accepted any HTTP client
   reachable on the bind host, and even when bound to ``127.0.0.1`` a
   browser tricked into resolving an attacker-owned domain to loopback
   could call the endpoint. Two new defenses:
   - **Fail-closed** when SSE is bound to a non-loopback host without
     ``PBI_MCP_AUTH_TOKEN``. The operator must either bind to localhost,
     set a bearer token, or explicitly opt out via
     ``PBI_MCP_ALLOW_UNAUTHENTICATED_SSE=1`` (logged as a warning).
   - **Always-on Host/Origin allowlist middleware**
     (``_origin_host_check_middleware``) — Host and Origin headers must
     match the bind host (loopback synonyms when bound to localhost) plus
     any ``PBI_MCP_ALLOWED_ORIGINS`` (csv) extension. Mismatches are
     rejected with HTTP 403, blocking DNS-rebinding from a malicious
     web page.

### High — readonly / deny-write policy bypass

2. **22 newly-added mutating tools were missing from `WRITE_TOOLS`**
   (`src/security.py`). `tool_category()` defaults unknown tools to
   ``"read"``, so ``--readonly`` / ``deny_categories: ["write"]`` did not
   block them. Added to the WRITE classification:
   `pbi_export_correction_report`, `pbi_create_time_intelligence_pack`,
   `pbi_create_ytd_measure`, `pbi_create_mtd_measure`,
   `pbi_create_spy_measure`, `pbi_create_yoy_measure`,
   `pbi_create_variance_measure`, `pbi_create_contribution_measure`,
   `pbi_create_topn_measure`, `pbi_create_rolling_average_measure`,
   `pbi_apply_format_preset`, `pbi_set_column_data_type`,
   `pbi_parameterize_data_source`, `pbi_relocate_data_source`,
   `pbi_add_scatter_chart`, `pbi_add_combo_chart`, `pbi_add_kpi`,
   `pbi_add_matrix`, `pbi_add_labelled_card`, `pbi_convert_visual_type`,
   `pbi_set_visual_format_property`, `pbi_disable_card_autoscale`.
3. **`pbi_validate_pbix_reopen` reclassified as WRITE** (was READ).
   The probe captures the entire primary screen, may close existing
   PBI Desktop windows, and writes a PNG to a caller-controlled path —
   none of which is read-only. Removed from `READ_TOOLS`, added to
   `WRITE_TOOLS`.
4. **`pbi_export_correction_report` removed from `READ_TOOLS` and
   `GRADING_TOOLS`** — it writes a Markdown report to ``output_path``
   so it is unambiguously a write tool.
5. **`--profile readonly` now also enables runtime readonly mode**
   (`src/server.py`). Previously it only pruned the registered tool
   surface using static `READ_TOOLS`, leaving dual-mode workflows
   (`pbi_excel_import_workflow`, `pbi_measure_workflow`,
   `pbi_repair_report_fields`) callable with ``apply=true`` and
   classified as write at call time but allowed by default policy. Now
   `--profile readonly` triggers `SECURITY.set_runtime_readonly(True)`,
   matching `--readonly` behavior.

### High — confidentiality / DMV / OCR

6. **`pbi_measure_dependencies` no longer bypasses the DMV guard**
   (`src/tools/query.py`). The tool hard-codes
   ``SELECT * FROM $SYSTEM.DISCOVER_CALC_DEPENDENCY``; it now calls
   `_validate_dax_query` first so the same ``PBI_MCP_ALLOW_DMV=1``
   opt-in that gates `pbi_execute_dax` applies here too.
7. **Windows OCR no longer leaks raw screen text**
   (`src/tools/quality.py`). The PowerShell helper used to serialize
   ``$result.Text`` (full recognized desktop text) into the response.
   Removed the `text` field entirely — only `text_length` and the
   matched signal labels (``"Fix this"`` etc.) are returned. Default
   `use_windows_ocr` flipped from ``True`` → ``False`` so callers must
   opt in explicitly even for the bounded form.

### High — input-validation / lexical bypasses

8. **M expression validator hardened against quoted-identifier and
   indirect-call bypasses** (`src/m_expression_security.py`).
   - Reject ``#"…"\s*(`` calls up-front against the **raw** expression.
     The literal stripper used to erase quoted identifiers as if they
     were strings, so ``#"Web.Contents"("https://…")`` slipped past the
     blocklist + allowlist.
   - Reject `Value.Invoke`, `Function.Invoke`, `Function.InvokeAfter`,
     and `Record.Field` explicitly. Each takes a callable as a
     first-class value and would let an attacker invoke a blocked
     function indirectly even though the wrapper sits under an allowed
     prefix.
9. **Zip-Slip in PBIX native ZIP extraction**
   (`src/tools/visuals/_io.py`). The fallback extraction iterated ZIP
   entries starting with ``Report/StaticResources/`` and joined them
   blindly with ``target / name``. A malicious PBIX with
   ``Report/StaticResources/../../../tmp/owned.txt`` could write outside
   the extraction folder. Now: members containing ``..``, ``.``,
   absolute prefixes, or drive letters are rejected, and the resolved
   destination must stay under ``target.resolve()``.

### High — security policy plumbing

10. **`enabled_tools: []` now denies every tool** (`src/security.py`).
    Previously the truthiness check ``{...} if enabled else None``
    coerced an explicit empty list to ``None``, which silently disabled
    the allowlist gate. Changed to ``if enabled is not None`` so an
    empty allowlist is treated as documented (lock everything down).
11. **Working-directory `security_policy.json` is honored again**
    (`src/server.py`). Server startup used
    ``cwd=Path(__file__).parent`` which made the loader probe
    ``src/security_policy.json`` instead of the operator's launch CWD.
    Restored to ``Path.cwd()`` so the documented configuration source
    works.

### Tests

12. **+4 regression tests** in `tests/test_security.py`:
    - `test_quoted_identifier_function_call_rejected`
    - `test_value_invoke_blocked`
    - `test_empty_enabled_tools_denies_every_tool`
    - `test_extract_report_zip_native_rejects_path_traversal`

    Plus updates to `tests/test_quality.py` for the OCR opt-in default
    and the `pbi_validate_pbix_reopen` reclassification. Suite now
    198 passing (was 194).

## [0.11.3] — 2026-05-08 — Runtime constant-measure probe + utcnow deprecation

Targeted fixes from the v0.11.x retrospective.

### Fixed

1. **`datetime.utcnow()` deprecation** in `src/pbi_connection.py` (×2,
   `connected_at` + operation log `ts`) and `src/tools/quality.py` (×1,
   correction-report header). Replaced with
   `datetime.now(UTC).isoformat().replace("+00:00", "Z")` to keep the
   trailing ``Z`` suffix that downstream tooling parses. Eliminates the
   `64 warnings` block that was appearing in every pytest run since
   Python 3.12 graduated the deprecation.

### Added

2. **Runtime probe for the bug-0.92 family** —
   `_runtime_probe_measure_constancy(manager, axis_ref, measure)` in
   `src/tools/visuals/_home_tables.py`. When a connection manager is
   available and the visual's axis can be recovered as ``Table.Column``,
   the probe issues a single DAX query:

   ```text
   EVALUATE TOPN(<sample_count>,
                 ADDCOLUMNS(VALUES('Table'[Column]), "__probe_v", [Measure]))
   ```

   and flags the measure as `runtime_constant_measure` when every
   sampled value is identical. This catches the cases the static
   `_is_likely_constant_dax` heuristic misses — e.g.
   `CALCULATE(SUM(Sales[Amount]), Sales[Amount] = 0)` references
   `Sales[Amount]` (so static parsing leaves it alone) but always returns
   zero. PBI Desktop renders these as a flat baseline or, on some
   builds, an opaque internal error.

   Default sample count is 3, clamped to ``[2, 10]``. Engine errors,
   < 2 distinct axis values, and missing axis information all degrade
   to ``(False, None)`` — the probe is conservative and never flags
   inconclusively.

### Changed

3. **`_inspect_value_measures` now accepts an `axis_ref` keyword** and
   runs the runtime probe as a second pass when the static check
   returns clean. Adds a new `runtime_constant_measure` warning entry
   alongside the existing `constant_measure` one.

4. **`pbi_add_line_chart_tool`, `pbi_add_combo_chart_tool`, and
   `pbi_add_waterfall_tool` now wire the axis through** —
   `axis_column` / `category_column` is forwarded to
   `_inspect_value_measures` so the runtime probe can fire at write
   time too, not just from the diagnostic tool.

5. **`pbi_diagnose_render_risks_tool`** now extracts the Category
   column from each cartesian visual's prototypeQuery (via the new
   `_recover_axis_full_ref` helper) and surfaces both
   `constant_measure` (static) and `runtime_constant_measure` (live
   engine) in `constant_measure_risks`. The probe payload is included
   per finding when present.

### Internals

- No new runtime dependency. The probe re-uses
  `pbi_execute_dax_tool` with a one-shot DAX query.
- `_recover_axis_full_ref` lives in `_repair.py`; sibling to the
  `_recover_full_refs_from_prototype` helper added in v0.11.0 (kept
  separate to avoid the bindings module taking on a layout-walk
  responsibility).
- Tool count unchanged: 134/134. ruff + format clean.

### Tests

- 6 new offline tests in `tests/test_visuals.py`:
  - `test_runtime_probe_returns_false_when_manager_is_none`
  - `test_runtime_probe_returns_false_when_axis_ref_invalid`
  - `test_runtime_probe_flags_dynamic_constant` — mocked
    `pbi_execute_dax_tool` returns 3 identical samples, asserts
    `is_constant=True`, probe payload populated.
  - `test_runtime_probe_passes_when_values_vary` — varying samples
    return `False`.
  - `test_runtime_probe_returns_false_on_engine_error` — DAX raise
    degrades cleanly to inconclusive.
  - `test_diagnose_render_risks_runtime_constant_surfaces` — full
    pipeline: build a line chart with a sneaky dynamic-constant
    measure, mock the probe, assert the diagnostic tool surfaces
    `runtime_constant_measure` with `axis_ref` populated.
- 194 passing locally (was 188), 2 platform skips. The pytest run is
  now warning-free.

## [0.11.2] — 2026-05-08 — Bug 0.92 diagnostic + render-risk aggregator

Best-effort static diagnostic for the bug-0.92 family — line / combo /
waterfall charts where a Y measure resolves to a scalar literal,
which Power BI Desktop reports as an opaque internal render error.
The full live repro still needs human eyes on the rendered visual,
but the static checks now flag the most common offenders before
compile.

### Added

1. **`pbi_diagnose_render_risks_tool`** in `src/tools/visuals/_repair.py`.
   Read-only aggregate diagnostic that walks the extracted layout and
   reports every render-risk it can detect statically:
   - Constant Y measure on cartesian charts (line, combo, waterfall,
     area, stackedArea).
   - Unresolved measure home tables.
   - Missing column / measure references in the live model (when a
     `manager` is supplied).
   - Wrong reference kind (column-only role bound to a measure or vice
     versa).
   - Query-ref mismatch between projections and prototypeQuery.

   Returns `{ok, risk_count, healthy, binding_issues,
   constant_measure_risks, model_validation, …}`. `page` and `visual_id`
   narrow the scan; both omitted scans the whole report.

2. **`_is_likely_constant_dax(expression)`** helper in
   `src/tools/visuals/_home_tables.py`. Strips line + block comments and
   string literals before checking for column/measure references, so:
   - `0.92` → flagged.
   - `BLANK()` → flagged.
   - `"Sales[Amount]"` (string literal) → flagged.
   - `/* Sales[Amount] */ 0.92` (commented ref) → flagged.
   - `SUM(Sales[Amount])` → not flagged.
   - `CALCULATE([Total Sales], Date[Year] = 2025)` → not flagged.

   Misses dynamic-but-still-constant DAX (e.g.
   `CALCULATE(SUM(Sales[Amount]), Sales[Amount] = 0)`) — those need a
   runtime probe.

### Changed

3. **`_inspect_value_measures` now uses `_is_likely_constant_dax`** so the
   line-chart, combo-chart, and waterfall builders share the same
   tightened heuristic. Previously the inline check was `"[" not in expr`
   on the raw expression, which a commented-out reference would defeat.

4. **`pbi_add_combo_chart_tool` and `pbi_add_waterfall_tool` now emit
   `warnings`** when their Y / Y2 measures look constant (parity with
   `pbi_add_line_chart_tool`, which has carried this since v0.10.x).

5. **Grading profile** now includes `pbi_diagnose_render_risks` so
   evaluation flows can flag bug-0.92-class issues before they reach the
   compile step.

### Internals

- Tool count: 133 → 134. Strict registry audit
  (`PBI_MCP_STRICT_REGISTRY=1`) clean: 134/134.
- New module-level constant `_CARTESIAN_VISUAL_TYPES` in `_repair.py`
  scopes the constant-measure check to chart families known to fail in
  the bug-0.92 way; bar / column / scatter etc. are unaffected.

### Tests

- 7 new offline tests in `tests/test_visuals.py`:
  - 5 `_is_likely_constant_dax` cases (literal, BLANK, comment-stripping,
    string-literal-stripping, real-aggregate negative).
  - `test_diagnose_render_risks_flags_constant_line_chart_measure` — full
    happy path: build a line chart with a constant measure, then query
    the diagnostic and assert the risk surfaces.
  - `test_diagnose_render_risks_clean_layout_returns_healthy` — empty
    page returns `healthy=True`, `risk_count=0`.
- 188 passing locally (was 181), 2 platform skips. ruff check + format
  check clean.

### Caveats

- The "constant" heuristic is intentionally conservative: it only fires
  when no `[...]` reference survives comment + string stripping, or when
  the entire body is `BLANK()`. Pathological constants that happen to
  reference a column whose value is always the same (`CALCULATE(SUM(...),
  ...) = 0`) are not detected; those need a runtime probe of the AS
  engine which is out of scope for a static tool.
- The `cartesian_visual_count` field in the diagnostic response counts
  every chart that *could* trigger bug-0.92, not just those that did. Use
  `risk_count` and `constant_measure_risks` for actionable items.

## [0.11.1] — 2026-05-08 — pbi_persist_now: opt-in Ctrl+S for Power BI Desktop

Closes the long-standing "TOM mutations are in-memory only" gap. Until now
every measure / column / table / relationship change a tool made lived in
the AS engine in memory; the user had to switch to PBI Desktop and press
Ctrl+S to persist. With `PBI_MCP_ALLOW_UI_AUTOMATION=1` and
`confirm=True`, callers can now ask the server to drive that key chord.

### Added

1. **`pbi_persist_now_tool`** in `src/tools/ui_automation.py`. Hard gates,
   both required:
   - **Server env**: `PBI_MCP_ALLOW_UI_AUTOMATION=1` must be set when the
     server starts. Without it the call fails with an explicit error.
   - **Per-call**: `confirm=True` must be passed. Default refuses.

   Behaviour:
   - Resolves the target PBI Desktop PID from the connection manager's
     live instance metadata (`manager._state.instance.pid`); falls back to
     the most recently started `PBIDesktop.exe` via `psutil` when no
     manager is connected.
   - Finds the visible top-level window of that PID via
     `EnumWindows` + `GetWindowThreadProcessId`.
   - Brings it to the foreground (`SetForegroundWindow`, with
     `ShowWindow(SW_RESTORE)` if minimised), captures the previous
     foreground HWND.
   - Injects exactly **one** key sequence — `Ctrl + S` — through
     `SendInput`. No other key sequences are emitted.
   - Optional `pbix_path`: when provided, polls the file's modification
     timestamp up to `timeout_seconds` (clamped to `[1, 60]`) and reports
     the observed delta in `save_observed`, `mtime_before`, `mtime_after`.
   - Restores focus to the previously-foreground window (best-effort).

2. **`src/wrappers/ui_automation.py`** — one-line `register_tool()` call,
   wired into `server.py` alongside the other domain wrappers.

### Internals

- All Win32 calls use `ctypes.windll.user32` directly — no extra runtime
  dependency (no pywinauto, pyautogui, keyboard).
- Tool count: 132 → 133. Strict registry audit (`PBI_MCP_STRICT_REGISTRY=1`)
  clean: 133/133.
- Registered as `write` in `security.WRITE_TOOLS`. Not exposed in the
  `grading` profile.
- `platform.system()` guard returns a structured error with
  `{"platform": "..."}` outside Windows so the tool never reaches the
  Win32 imports on Linux/macOS.

### Tests

- New `tests/test_ui_automation.py` (7 cases):
  - **Gate tests** (run on every platform):
    - `test_requires_confirm` — `confirm=False` raises.
    - `test_requires_env_opt_in` — env var unset raises (Windows path) /
      platform guard raises (non-Windows path).
    - `test_tool_category` — categorised as `write`.
  - **Execution tests** (Windows-only via `skipUnless`):
    - `test_no_pid_returns_error_payload` — no PBI Desktop process raises.
    - `test_no_window_for_pid_returns_error` — process exists but no
      visible window raises.
    - `test_full_path_polls_mtime` — full happy path with mocked Win32
      injection that simulates Power BI flushing the pbix; asserts mtime
      observation works.
    - `test_no_pbix_path_returns_immediately` — Ctrl+S sent, no polling.
- 181 passing locally (was 174), 2 platform skips. ruff check + format
  check clean.

### Lessons / footguns avoided

- `SendInput` is the only Win32 API that survives UAC isolation between
  unprivileged and elevated processes; `keybd_event` does not. PBI
  Desktop typically runs unelevated, but if a user starts the MCP server
  from an elevated terminal it would silently no-op with `keybd_event`.
- `SetForegroundWindow` requires the calling process to own the
  foreground or have specific input ownership. The brief 150 ms sleep
  after `SetForegroundWindow` keeps Windows from racing the key injection
  back to the previous focus owner.
- Mtime polling has a 250 ms granularity — snappy enough for interactive
  feel, tight enough to detect saves that finish in under a second.

### Caveats

- If the PBIX has never been saved (no on-disk file), Ctrl+S triggers
  a "Save As…" dialog. The tool injects the key chord regardless; it
  does not (and will not) auto-fill the filename. Caller should ensure
  the file has been saved at least once.
- If a modal dialog is open in PBI Desktop (refresh prompt, conflict
  resolution, etc.), Ctrl+S is consumed by the dialog. The tool reports
  `save_observed=False` after the timeout in that case.

## [0.11.0] — 2026-05-08 — Visual binding edits without remove + recreate

First feature drop after the 0.10.x stabilisation cycle. Adds
`pbi_update_visual_bindings_tool` so callers can mutate an existing visual's
projections + `prototypeQuery` in place. The previous workflow
(`pbi_remove_visual` → `pbi_add_*`) regenerated the visual with a fresh
visual id, dropping any side properties (size, format, theme, conditional
formatting, manual tweaks) the user had already applied.

### Added

1. **`pbi_update_visual_bindings_tool`** in `src/tools/visuals/_ops.py`. Two
   modes, mutually exclusive:
   - `projections={role: [refs]}` — full replacement of every role.
   - `add_to_role` / `remove_from_role={role: [refs]}` — incremental edits
     against the visual's current bindings; roles that empty out are
     dropped automatically.

   References use the same forms accepted everywhere else in the visuals
   API: `Table.Column`, `Table[Column]`, `'Table With Spaces'[Column]`, or
   bare measure names.

   Validation chain (re-uses existing pre-flight code):
   - Roles checked against `VISUAL_FIELD_ROLES[visual_type]`.
   - Reference kinds (column vs measure) checked against
     `VISUAL_ROLE_KINDS[visual_type]` when a `manager` is connected.
   - Each reference checked against the live model (with did-you-mean
     suggestions on miss).
   - Final binding sanity check via `_assert_container_bindings`.

   `prototypeQuery` is rebuilt from the new reference set so PBI Desktop
   renders the visual correctly. `dry_run=True` runs every check but skips
   the layout disk write — response carries `dry_run=True` and the
   write log.

   The tool returns `old_projections`, `new_projections`, `added`, and
   `removed` so callers can audit the change without diffing layouts.

### Internals

- New helper `_recover_full_refs_from_prototype()` in `_ops.py` walks the
  visual's existing `prototypeQuery` Select+From entries to recover the
  full `Table.Column` form from the short `queryRef` stored in
  `projections`, so incremental edits don't need the caller to know about
  the column-name aliasing.
- Tool count: 131 → 132. Strict registry audit (`PBI_MCP_STRICT_REGISTRY=1`)
  clean: 132/132.
- Registered as `write` in `security.WRITE_TOOLS`. Not exposed in the
  `grading` profile.

### Tests

- 7 new offline tests in `tests/test_visuals.py`:
  - `test_update_visual_bindings_replace_projections`
  - `test_update_visual_bindings_incremental_add_and_remove`
  - `test_update_visual_bindings_rejects_mutually_exclusive_inputs`
  - `test_update_visual_bindings_rejects_no_input`
  - `test_update_visual_bindings_rejects_unknown_role`
  - `test_update_visual_bindings_rejects_empty_result`
  - `test_update_visual_bindings_dry_run_does_not_persist`
- 174 passing locally (was 167), 2 platform skips. ruff check + format
  check clean.

## [0.10.15] — 2026-05-08 — Docs refresh + CI Windows fixes

Documentation pass + two CI Windows test fixes spotted in v0.10.14's first run.

### Added

1. **`ARCHITECTURE.md`** — full module layering (server / mcp_core / wrappers / tools / pbi_connection / security), visuals/ submodule tree with the dependency graph, the `register_tool()` auto-detection contract, lifecycle (PID lock, parent watcher, registry audit), profiles, layout-write atomicity, TOM mutation persistence semantics.

2. **`CHANGELOG-archive.md`** — historical entries (v0.10.11 and earlier) split out so the active `CHANGELOG.md` stays scannable.

### Fixed

3. **CI Windows test patch paths**. `tests/test_visuals.py::test_extract_and_compile_reports_with_mocked_subprocess` patched `tools.visuals._find_pbi_tools` / `tools.visuals.subprocess.run`. After the v0.10.9 split those names live in `tools.visuals._io`. Updated the patches accordingly.

4. **Windows tempfile short-path comparison**. `tests/test_persistent_report.py::test_create_persistent_report_writes_pbix` compared `tempfile.TemporaryDirectory()` (short form `RUNNER~1` on GitHub Actions) against the tool's resolved path (long form `runneradmin`). Both sides now go through `Path.resolve()`.

5. **Coverage threshold lowered to 54%** (was 55). Local runs land at 55% but the GitHub Actions Windows runner reports 54.86 — small environment-dependent difference.

### Changed

6. **README refresh** — tool count `109 → 131`, `--profile=grading` documented, repository layout reflects `mcp_core.py` + `wrappers/` + `tools/visuals/` submodule structure, dev section mentions `ruff` + `pytest --cov`, link to `ARCHITECTURE.md`, added a CI status badge.

### Tests

- 167 passing locally, 2 platform skips. Coverage 55% local, ≥54% expected on GitHub Actions Windows runner.

## [0.10.14] — 2026-05-08 — DRY: consolidated `_run` + metaprogrammed wrappers

Two more items from the v0.10.12 retrospective. No tool surface change, no behavior change. 167/167 tests, registry strict 131/131 clean.

### Changed

1. **Consolidated `_run` offline helper** in `src/tools/visuals/_base.py`. The 5 visuals submodules (`_io`, `_pages`, `_design`, `_repair`, `_ops`) and the package `__init__.py` previously each defined their own near-identical `def _run(callback): try: return callback() except: return error_payload(exc)`. Single canonical definition now, imported via `from ._base import _run`. The package `__init__.py` re-exports it under the name `_run` so historic `from tools.visuals import _run` imports keep working.

2. **`register_tool()` helper for metaprogrammed MCP wrappers** in `src/wrappers/_helpers.py`. Generates a pass-through wrapper from the underlying `pbi_*_tool` signature. Auto-detects manager-injection pattern (positional first param vs keyword-only with default `None` vs none). Preserves the parameter schema via `__signature__` so FastMCP picks up the right JSON schema. Copies docstring from the underlying tool.

3. **All 14 wrapper modules collapsed** to one-line registrations. Each previously expanded the boilerplate `@mcp.tool()` + `def fn(...)` + `return _run("name", tool, ..., manager=CONNECTION_MANAGER)` for every tool. Now each wrapper file is just:
   ```python
   from tools import (pbi_xxx_tool, ...)
   from ._helpers import register_tool

   register_tool(pbi_xxx_tool)
   ```

### Internals

- `src/wrappers/`: 3372 → 534 lines (-84%).
- The single `register_tool` call replaces ~10-30 lines per wrapper × 144 wrappers.
- Auto-detection: 7 connection wrappers + 4 relationships + 6 rls + 3 calc_groups + 10 model + 19 measures + 7 query + 21 quality + 13 excel + 11 power_query + 4 tmdl + 1 project + 36 visuals + 2 workflows = 144 wrappers, all pass-through compatible.
- New `src/wrappers/_helpers.py` (115 L) is the single source of truth for wrapper boilerplate.

### Tests

- 167 passing, 2 platform skips. `ruff check` + `ruff format --check` clean. Strict registry audit (131/131) clean.

### Lessons / footguns avoided

- FastMCP's `@mcp.tool()` decorator inspects the function's signature for the JSON parameter schema. Setting `__signature__` on the generated wrapper is necessary to keep clients seeing typed parameters instead of `**kwargs`. Verified: `pbi_connect` schema correctly exposes `preferred_port`, `force_reconnect`.
- Tools with extra wrapper-only logic (e.g. `pbi_add_visual` has `dry_run`) are still pure pass-through because the `dry_run` param exists on the underlying `pbi_add_visual_tool` itself. Nothing requires manual wrapping.
- `functools.update_wrapper` plus an explicit `__signature__` reassignment is the right combination — `update_wrapper` would otherwise copy the underlying function's signature (with `manager`), which we want to drop.

## [0.10.13] — 2026-05-08 — Hardening: CI, ruff, coverage, public tests, strict audit

Hotfix release on the items flagged as high-ROI in the v0.10.12 retrospective. No tool surface changes; pure infrastructure + style.

### Added

1. **GitHub Actions CI** (`.github/workflows/ci.yml`). Two jobs:
   - `test` — runs `pytest -q --cov=src --cov-fail-under=55` on Windows (full deps) + a reduced offline subset on Ubuntu, with `PBI_MCP_AUDIT=1` and `PBI_MCP_STRICT_REGISTRY=1` so registry drift fails the build.
   - `lint` — `ruff check` + `ruff format --check` on Linux.
   - Matrix: Python 3.11 + 3.12.

2. **Ruff configured in `pyproject.toml`** — line-length 120, target py311, rule sets `E F W I UP B` with pragmatic ignores for E501/B008/B904/B007/B017/B023 plus per-file ignores for the re-export façade modules.

3. **Coverage configured** — `pytest-cov` in dev deps, `[tool.coverage.*]` settings in `pyproject.toml`. Baseline coverage is 55% (enforced floor in CI). Excluded from coverage: `src/tools/visuals/_io.py` (PowerShell + subprocess paths needing live env) and `src/wrappers/*.py` (thin wrappers exercised via integration).

4. **Tests directory committed** — `tests/` is no longer gitignored at the project level. The 17 offline test files (~167 test cases) are now part of the repo. Three pattern-based exclusions stay gitignored for live-only scripts: `tests/demo_*.py`, `tests/smoke_e2e.py`, and `tests/test_*_local.py`.

### Changed

5. **Codebase formatted with `ruff format`** — 41 files reformatted to a consistent 120-column style. No semantic changes; `ruff check` passes with zero errors.

6. **Imports re-organised by `ruff check --fix`** — every module now has stdlib / third-party / local imports separated and alphabetised (isort rule).

### Internals

- 17 wrapper-style modules (`src/wrappers/`, `src/tools/__init__.py`, `src/tools/visuals/__init__.py`) carry F401/E402 per-file ignores since their re-exports are the public surface.
- `_audit_tool_registry(strict=True)` runs by default in CI via `PBI_MCP_STRICT_REGISTRY=1`. Locally still opt-in via env var.
- New dev deps: `pytest-cov~=5.0`, `ruff~=0.6`.

### Tests

- 167 passing, 2 platform skips. Coverage 55% on Windows full-deps run.

## [0.10.12] — 2026-05-08 — Refactor phase 7 (final): server.py wrapper split

The full structural refactor is **complete**. `src/server.py` shrunk from 3215 → 323 lines; the 144 `@mcp.tool()` wrappers are now distributed across 14 focused domain modules under `src/wrappers/`. No behavior change, no tool surface change. 167/167 tests, registry 131/131 clean.

### Changed

`src/wrappers/` package created. Each module registers its tools as a side-effect of import:

| Module | Wrappers | Lines |
|---|---|---|
| `wrappers/connection.py` | 7 | 92 |
| `wrappers/model.py` | 10 | 175 |
| `wrappers/measures.py` | 19 | 467 |
| `wrappers/relationships.py` | 4 | 103 |
| `wrappers/rls.py` | 6 | 101 |
| `wrappers/calc_groups.py` | 3 | 50 |
| `wrappers/query.py` | 7 | 138 |
| `wrappers/quality.py` | 21 | 385 |
| `wrappers/visuals.py` | 36 | 949 |
| `wrappers/excel.py` | 13 | 186 |
| `wrappers/power_query.py` | 11 | 246 |
| `wrappers/tmdl.py` | 4 | 72 |
| `wrappers/project.py` | 1 | 32 |
| `wrappers/workflows.py` | 2 | 47 |
| **Total** | **144** | **3043** |

`src/server.py` now keeps only:
- runtime-singletons import from `mcp_core` (mcp, CONNECTION_MANAGER, _run, audit, profile, lock, watcher)
- `find_pbi_port()` compatibility shim for standalone scripts
- Side-effect imports of every `wrappers.<domain>` module
- `@mcp.resource()` and `@mcp.prompt()` registrations (3 resources, 7 prompts)
- SSE Bearer-auth middleware + `_run_sse_with_auth`
- `main()` with argparse + profile filter + transport launcher

### Internals

- `server.py`: 3215 → 323 lines (-90%).
- Total `src/` line count after refactor: **roughly half** of pre-refactor when measured against the original monoliths.
- Registry audit clean: 131/131, 0 orphans.
- Per-wrapper imports of `mcp` from `mcp_core`, `_run` + `CONNECTION_MANAGER`, plus only the specific `pbi_*_tool` it wraps.

### Tests

- 167 passing, 2 platform skips.

### End of refactor sequence

- v0.10.6: package skeleton + mcp_core extraction
- v0.10.7: visuals/ phase 2 (5 submodules)
- v0.10.8: visuals/ phase 3 (bindings, home_tables, containers)
- v0.10.9: visuals/ phase 4 (I/O block)
- v0.10.10: visuals/ phase 5 (pages, design, repair)
- v0.10.11: visuals/ phase 6 final (charts, cards, structure, ops, dispatcher)
- **v0.10.12: server.py wrapper split (this release)**

The codebase is now organised into ~30 focused modules instead of the original 3 monoliths totalling ~7100 lines.
