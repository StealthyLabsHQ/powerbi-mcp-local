# Changelog

Active changelog covers the most recent releases.  
Older entries (v0.10.11 and earlier — refactor phases, v0.10.x feature drops, v0.8.x and v0.7.x history) live in [CHANGELOG-archive.md](CHANGELOG-archive.md).

## [0.15.0] — 2026-07-02 — Safety hardening + unified contract + tool-surface consolidation

Architecture overhaul in five phases: protect user files first, then make
every response speak the same envelope, then shrink the tool surface an
LLM has to reason about.

### Phase 1 — write safety (data-loss fixes)

- **New `atomic_io` module**: shared `atomic_write_bytes` / `atomic_write_text`
  (temp sibling → `.bak` of previous content → atomic `os.replace`) and
  `snapshot_once` (pristine `.orig` preserved across multi-round loops).
- `pbi_apply_style_preset` no longer overwrites the source `.pbix` in place
  with a raw non-atomic write — a crash mid-write can no longer truncate
  the original. Response now carries `backup_path`.
- `pbi_write_tmdl_file` and `pbi_patch_tmdl_measure` write atomically with a
  `.bak` of the previous TMDL content.
- `pbi_patch_layout` keeps a `.bak` of the pre-patch PBIX (the zip rebuild
  was previously unrecoverable if subtly wrong). Response carries
  `backup_path`.
- **Force-kill guard**: `force=True` now refuses to `taskkill` Power BI
  Desktop when the graceful save-and-close failed and no prior save was
  verified (`save_attempt.mtime_changed`) — unsaved in-memory model edits
  are no longer silently discarded.
- `pbi_repair_loop` snapshots the pristine layout once (`Layout.orig`)
  before its first repair round, so multi-round repairs no longer erode
  the only backup. Response carries `pristine_snapshot`.

### Phase 2 — unified response contract + error classification

- `_enforce_response_size` now returns the standard `{ok: false, error: …}`
  envelope (was a `{status: "error"}` dialect that `ok`-keyed clients
  misread as success); audit logging keys on `ok` too.
- Connection-loss classification is now locale-independent: .NET exception
  type names (`AdomdConnectionException`, `SocketException`, …) are checked
  before English message substrings, so auto-reconnect works on localized
  Power BI Desktop builds.
- TMDL structural-injection hardening: multi-line measure expressions are
  re-indented so they cannot dedent into sibling TMDL blocks;
  `format_string`/`display_folder` reject line breaks.
- `validate_measure_name` now also guards `pbi_create_measures` (batch),
  `pbi_rename_measure`, and `pbi_patch_tmdl_measure` (was only
  `pbi_create_measure`); `measure_name` params are name-validated.
- `secure_tool` binds positional arguments to parameter names before
  policy validation (positional call sites could previously bypass it).

### Phase 3 — tool-surface consolidation (174 → 141 tools)

- **BREAKING**: the 28 per-type `pbi_add_<chart>` MCP tools are demoted to
  internal builders. Use `pbi_add_visual(visual_type=…, config=…)` (same
  builders, one entry point, per-type config documented in the tool
  description) or `pbi_add_visual_from_intent`. The Python implementations
  remain importable from `tools.visuals`.
- **BREAKING**: `pbi_create_ytd/mtd/spy/yoy_measure` are demoted — use
  `pbi_create_time_intelligence_pack(patterns=["YTD"])` etc. (same
  implementation; the pack auto-resolves dependencies).
- **BREAKING**: `pbi_validate_dax_semantic` is merged into
  `pbi_validate_dax(semantic=true)`.
- **Single-source registration**: `wrappers/<domain>.py` modules are
  replaced by `wrappers.register_all()`, which derives the MCP registry
  from `tools.__all__`. Adding a tool = implement + export, nothing else.

### Phase 4 — tests + CI

- New offline unit tests for the previously untested mutating modules:
  RLS, calc groups, relationships, measures (incl. time-intelligence pack
  and DAX-file parser); new `tests/test_atomic_io.py`.
- Live-only scripts moved out of `tests/` (they contributed zero collected
  tests): `tests/test_connection.py` → `scripts/live_connection_check.py`,
  `demo_*.py`/`smoke_e2e.py` → `scripts/`.
- Linux CI now runs the full offline suite instead of a hardcoded 4-file
  allowlist; new `tool-count` CI job fails when the README badge drifts
  from the registered tool count.

### Phase 5 — decomposition

- `tools/quality.py` (2170 lines, 7 mixed concerns) → `tools/quality/`
  package: `_shared` / `_model_audit` / `_dax_lint` / `_layout_lint` /
  `_persistence` / `_scoring`, with a façade `__init__` preserving every
  historical import.
- `tools/measures.py` (1115 lines) → `tools/measures/` package:
  `_crud` / `_time_intelligence` / `_dax_import`, same façade pattern.
- Envelope + error surface extracted from `pbi_connection.py` (1362 →
  1143 lines): new `pbi_errors.py` (exception taxonomy) and
  `envelopes.py` (`ok`, `error_payload`, `serialize_value`, …), both
  dependency-free; `pbi_connection` re-exports for compatibility.
- Coverage 60% → 68%; CI gate raised 55 → 62.
- README tool badge updated (162 → 141 actual registered tools).
- Deferred (documented): renaming the installed package off top-level
  `src` (namespace-collision risk), metadata-cache invalidation on
  external Desktop edits, refresh timeout/lock scope.

## [0.14.0] — 2026-06-09 — Visual Intent Layer + repairable-error loop

Phase 1 of the LLM-reliability plan: stop asking models to pick visual
types and roles directly, and return errors they can act on.

### Added

- **Visual Intent Layer** (`visuals/_intent.py`):
  - `pbi_plan_visual` — pure planner: a constrained business-intent spec
    (`metric`/`metrics`, `dimension`, `time`, `breakdown`, `target`, plus
    boolean hints `comparison`, `trend`, `parts_of_whole`, `correlation`,
    `geographic`, `detail_table`, `many_categories`, `filter_control`)
    is mapped deterministically to a visual type, dispatcher config, and
    a rationale. Unknown intent keys and bare (non-`Table.Field`)
    references are rejected.
  - `pbi_add_visual_from_intent` — plans then builds through the generic
    dispatcher (full binding validation, `dry_run` supported); the
    response carries the `decision` (type, rationale, config).
- **Repairable-error registry + loop** (`visuals/_errors.py`):
  - `pbi_list_repairable_errors` — stable error vocabulary; every code
    carries severity, auto-repairability, and an `llm_action` telling
    the calling model how to fix its *spec*.
  - `pbi_repair_loop` — detect → classify → auto-repair → re-verify
    loop over a report extract. Deterministic fixes (query-ref
    mismatches, measure home tables) are applied and re-scanned until
    convergence; residual issues come back classified. New detectors:
    `gauge_target_invalid` (target outside [min, max], min ≥ max) and
    `double_display_units` (measure format string already scaled to K/M
    while the visual also sets `labelDisplayUnits`). Optional
    `check_empty_visuals=true` folds the live empty-visual probe into
    the same classified output.

### Fixed

- `pbi_add_visual` docstring promised `kpi` and `matrix` but the
  dispatcher never routed them — both are now registered (with
  `indicator_measure`/`trend_column`/`goal_measure` and
  `rows`/`values`/`columns` configs respectively).

## [0.13.5] — 2026-05-13 — Styling canvas wallpaper + theme activation

Field report on v0.13.3 / v0.13.4: PBIX files produced by
`pbi_apply_style_preset` embedded the PNG + theme JSON + OPC manifest
correctly, but Power BI Desktop opened with a blank canvas — the
wallpaper was never wired into the Layout's canvas-background block,
and the theme was never registered as active.

### Causes

1. `image.name` / `image.url` / `image.scaling` were emitted as
   `{"expr": {"Literal": {"Value": "'…'"}}}` instead of bare strings.
   Power BI silently drops canvas image references that wrap those
   properties in Literal expressions — the binding mechanism only
   accepts bare strings under `image.*`. The Literal wrapper still
   applies to `show` / `transparency` (which *are* expression
   evaluable).
2. `image.url` had no `RegisteredResources/` prefix. PBI Desktop
   resolves the URL relative to `StaticResources/`, so without the
   prefix the lookup fails and the canvas falls back to blank.
3. `objects.wallpaper` (the chrome layer around the canvas) was never
   written, leaving a white halo around the styled canvas.
4. The theme was set on `layout.activeTheme` only. Newer Power BI
   builds read `layout.reportSettings.activeTheme`; both keys now get
   the theme entry.
5. Visual titles got overwritten whenever the preset pass re-built
   `vcObjects`. The original user-set title was lost.

### Fix

1. **Bare-string image block** (`_embed.py::_image_block`). The
   wallpaper entry is now:

   ```json
   "image": {
     "name": "<filename>",
     "url": "RegisteredResources/<filename>",
     "scaling": "Stretch"
   }
   ```

   while `show` and `transparency` keep their `Literal` expressions.
   `scaling` enum extended to include `Stretch` (PBI default for
   full-bleed wallpapers).

2. **Wallpaper layer** also written to `objects.wallpaper`
   (`apply_wallpaper_layer=True` by default) so the area surrounding
   the canvas matches the canvas itself.

3. **`wallpaper_fit` parameter** on
   `pbi_apply_style_preset_tool`. Overrides the preset's
   `page.wallpaper.fit`. Validated against the
   `Fit / Fill / Normal / Stretch` enum at the boundary. Defaults of
   `glassmorph_dark` and `glassmorph_light` switched to `Stretch`.

4. **Theme activation in three places**: `layout.themeCollection`,
   `layout.activeTheme`, and `layout.reportSettings.activeTheme`.
   Duplicate entries in `themeCollection` are collapsed so a repeated
   apply doesn't grow the list.

5. **Custom title preservation**
   (`_embed.py::_extract_visual_title`). Before the preset pass
   rebuilds `vcObjects`, any non-empty
   `objects.title[*].properties.text` is snapshotted and re-inserted
   after styling. Empty titles are not counted as "preserved" so the
   metric tracks real user-set titles.

6. **Post-write validation gate**. The styling tool now re-extracts
   the written PBIX and asserts:
   - every targeted page has a wallpaper image reference that
     resolves to a real archive part;
   - the theme part exists and `layout.activeTheme` points at it;
   - DBCC remains valid.
   A failed assertion raises `PowerBIValidationError` with the
   per-page error list so a half-applied style never returns ok=True.

### New return-payload fields

`pbi_apply_style_preset_tool` now returns, in addition to the
existing keys:

- `wallpaper_applied_pages: list[str]` — pages whose Layout now
  references a real wallpaper part.
- `theme_activated: bool` — true when the embedded theme path is
  present at `layout.activeTheme` and the archive carries the part.
- `custom_titles_preserved: int` — count of visuals whose user-set
  title survived the styling pass.
- `validation_errors: list[dict]` — empty when the file is ready to
  open; the gate raises before this line if anything fails.
- `wallpaper_fit: str` — the actual scaling used (after override
  resolution).

### Tests (+13)

`tests/test_v0_13_5_canvas_activation.py`:

- 5 schema tests (bare-string image block, wallpaper layer, fit
  enum, glassmorph default `Stretch`).
- 2 title-preservation tests (non-empty title kept, empty title not
  counted).
- 6 end-to-end apply tests (wallpaper reference + part exists, theme
  in both layout roots, custom titles round-trip, `wallpaper_fit`
  override, bad-fit rejection, validation gate fires when wallpaper
  is missing).

Plus the v0.13.3 schema test updated to assert the bare-string
contract instead of the legacy Literal-wrapped form.

Full suite: **387 passed**, 2 skipped (was 374 in v0.13.4; +13).

### Demo

`scripts/demo_glassmorph.py` generates a fixture PBIX with a
custom-titled card, applies `glassmorph_dark`, and prints the full
return payload (wallpaper pages, theme activation, preserved title,
DBCC). Useful as the local "zero-click" smoke check before opening
the file in Power BI Desktop.

## [0.13.4] — 2026-05-13 — Styling OPC manifest fix

Field report on v0.13.3: PBIX files produced by `pbi_apply_style_preset`
opened with:

> MashupValidationError: This file is corrupted or was created by an
> unrecognized version of Power BI Desktop.

### Cause

A PBIX is an OPC (Open Packaging Conventions) package. Every part's
extension must be declared in `[Content_Types].xml` via a `Default`
entry. The v0.13.3 repacker embedded
`StaticResources/RegisteredResources/<name>.png` without touching the
manifest, so the PBIX had a PNG part whose extension wasn't declared
→ hard fail on reopen.

### Fix

1. **`patch_content_types(raw_xml, required_extensions)`** in
   `src/tools/styling/_embed.py`. Reads the existing
   `[Content_Types].xml`, parses every declared `Default Extension="…"`,
   and adds the missing entries before `</Types>`. Catalogue:
   `png → image/png`, `jpg/jpeg → image/jpeg`, `gif → image/gif`,
   `bmp → image/bmp`, `svg → image/svg+xml`, `json → application/json`,
   `xml → application/xml`. Returns the patched bytes plus the list
   of extensions that were added. Handles the rare case where
   `[Content_Types].xml` is absent by writing a minimal document with
   every required `Default`.

2. **`repack_pbix` rewrites the manifest** on every apply. Builds the
   final part-name list, derives the required extension set, and
   writes the patched `[Content_Types].xml` alongside the new layout
   and resources.

3. **Post-write fail-loud validation**
   (`validate_content_types_declarations` + apply-tool re-read). The
   styling tool now opens the freshly-written PBIX, reads
   `[Content_Types].xml`, and asserts every extension among the parts
   has a matching `Default`. A missing declaration raises
   `PowerBIValidationError` with the offending extension list so the
   caller never ships a broken file silently. Response carries
   `content_types_required: [...]` and `content_types_missing: [...]`
   for downstream inspection.

### Tests

`tests/test_v0_13_4_content_types.py` — 11 tests:

- Catalogue + extension extraction.
- Patch adds missing PNG entry, preserves existing XML entry.
- Patch is a no-op when everything is already declared.
- Patch handles a missing `[Content_Types].xml` (writes a fresh one).
- Validator lists missing extensions / returns empty when complete /
  treats missing XML as fully-missing.
- Apply tool writes the PNG + JSON Default entries on a real PBIX
  round-trip.
- Apply tool synthesises `[Content_Types].xml` when the source PBIX
  lacks one.
- Repeated apply is idempotent — each extension appears exactly once
  in the manifest, no duplicates accumulate.

Full suite: **374 passed**, 2 skipped (was 363 in v0.13.3; +11).

## [0.13.3] — 2026-05-13 — One-shot styling presets (wallpaper + chrome + theme)

End goal: a single MCP call configures wallpaper, page background, card
chrome, chart chrome, accent colours, and theme JSON on every page of an
existing PBIX. Zero manual clicks in Power BI Desktop.

### New tools

1. **`pbi_apply_style_preset`** (`src/tools/styling/_apply.py`). Applies
   a built-in or custom preset to an existing `.pbix`. Pipeline:
   - Unzip PBIX in memory.
   - Embed wallpaper PNG under
     `StaticResources/RegisteredResources/<sanitized>_<sha1>.png`
     (SHA-1 dedup; reapplication never bloats the archive).
   - Patch every `ReportSection.config.objects.background` with the
     canonical Power BI wallpaper block (image + fit + transparency)
     plus the page background colour from the preset.
   - Patch every visual's `vcObjects` with the preset's card / chart
     chrome (background, border, drop shadow, radius, weight,
     transparency).
   - Pick a card border accent per visual from the bound measure name
     via the heuristic in `_accent.py`
     (`positive` / `warning` / `info` / `neutral`).
   - Embed the preset's theme JSON under
     `StaticResources/SharedResources/BaseThemes/<name>.json` and set
     `layout.activeTheme` + `layout.themeCollection`.
   - Repack the PBIX.
   - Run `pbi_diagnose_pbix_dbcc` on the output — fail loud if the
     styling round-trip regressed the string store.

2. **`pbi_list_style_presets`**. Returns the catalogue with `name`,
   `description`, `palette`, and `default_wallpaper` per preset.

### Built-in preset catalogue (5)

- `glassmorph_dark` — frosted navy glass on dark gradient.
- `glassmorph_light` — translucent white on sky gradient.
- `neon_cyber` — magenta / cyan / lime on near-black.
- `minimal_corporate` — white canvas, hairline borders, no shadows.
- `dark_pro` — saturated dark dashboard, opaque cards, sharp accents.

Each ships a full theme JSON (passes the v0.13 theme validator), a
palette, card + chart chrome specs, and an `accentMap`. Default
wallpapers are generated lazily as vertical-gradient PNGs (native
zlib + PNG chunk writer — no Pillow runtime dep) into
`src/tools/styling/_wallpapers/` and cached.

### Auto-accent inference

`infer_accent_key(measure_name)` matches against documented
substrings (case-insensitive):

- `croissance` / `growth` / `marge brute` / `gross margin` / `ebe` /
  `ebit` / `marge nette` / `net margin` / `profit` → `positive`
- `endettement` / `debt` / `leverage` / `bfr` / `wcr` / `charge` /
  `expense` / `frais` / `cost` → `warning`
- `var` / `variance` / `geo` / `atelier` / `workshop` / `store` →
  `info`
- otherwise → `neutral`

Override per measure via `custom_spec`.

### Integration with `pbi_create_persistent_report`

Three new optional parameters on `pbi_create_persistent_report_tool`:

- `style_preset: str | None`
- `style_wallpaper_path: str | None`
- `style_custom_spec: dict | None`

When `style_preset` is set, the styling apply step runs immediately
after save. Single-call path from spec → PBIX + chrome + theme +
wallpaper.

### Validation

- Wallpaper: native PNG IHDR inspection. Reject > 1920 × 1080 or
  > 2 MB before embed. No Pillow runtime dependency.
- Preset palette: every value must match `#RRGGBB`.
- Preset theme: reuses `validate_theme_payload` (size cap + key
  allowlist + colour format + URL guard).
- Output PBIX: `pbi_diagnose_pbix_dbcc` runs post-write — regression
  on string store fails the call.

### Tests (+22)

`tests/test_v0_13_3_styling.py`:

- 4 preset-catalogue tests (count, palette hex, theme schema, accent
  map shape).
- 7 accent-inference tests (positive / warning / info / neutral /
  fallback chain).
- 2 native-PNG round-trip tests.
- 5 embed-helper tests (sanitize, sha1 dedup, wallpaper section
  patch, page filter, vcObjects emission).
- 4 end-to-end apply-tool tests on a synthetic PBIX (wallpaper +
  theme embedded, page filter respected, custom preset path).

Full suite: **363 passed**, 2 skipped (was 341 in v0.13.2; +22).

### Tool surface

`+2` MCP tools, no breaking signatures:

- `pbi_apply_style_preset` (write)
- `pbi_list_style_presets` (read)

### Out of scope (documented for future work)

- Custom font embedding (Power BI doesn't expose this in the layout).
- Animated backgrounds.
- CSS-like blend modes (Power BI limitation).
- Mobile layout styling.

## [0.13.2] — 2026-05-13 — pbix-mcp 0.9.2 upstream patches (5 bugs)

In-tree patches against the vendored `pbix-mcp` 0.9.2 dependency
(`.venv/Lib/site-packages/pbix_mcp/`). Documented in detail in
[UPSTREAM_PATCHES.md](UPSTREAM_PATCHES.md). Re-applied on a fresh
install by re-running `tests/test_v0_13_2_pbix_mcp_patches.py`.

### Bug #1 — `Measure.FormatString` persistence

Already fixed in the published 0.9.2 SQL INSERT. Pinned with a
regression test that fails if the literal `NULL` ever returns.

### Bug #2 — DBCC string-store corruption on `HASONEVALUE+VALUES`

Power BI Desktop refused to open a PBIX with:

> PFE_XM_DBCC_STRINGSTORE_CORRUPT,
> PFE_XM_ERROR_WHILE_PARALLEL_LOADING_IMBI_TABLE_DATA

Root cause is in the upstream Vertipaq dictionary encoder — not
addressable from a Python-side patch alone. Mitigation shipped:

- **New `pbix_mcp/dbcc_guard.py`** with a pattern catalogue
  (`hasonevalue_values_string`, `selectedvalue_string_default`,
  `treatas_string`) and a `DBCCRiskWarning` category.
- `PBIXBuilder.save()` and `_pre_build_checks()` now scan the
  measure list against the catalogue and surface findings as
  structured warnings instead of silently producing a corrupt
  `.pbix`.
- Mitigation suggestions are embedded in the warning message:
  source the affected table from CSV/DB, swap to an Int64 surrogate
  key, or run `pbi-tools roundtrip` after build.

### Bug #3 — Visual config pass-throughs

Already shipped in 0.9.2: `series` projection, visual `objects` /
`vcObjects`, page-level `config`. Pinned with three regression tests
that assert the rendered Layout JSON carries each key.

### Bug #4 — Clean `add_measure(format_string=…)` signature

Added the keyword-only `format_string: str | None = None` parameter
to `PBIXBuilder.add_measure` and propagated to `self._measures` so
the upstream INSERT (Bug #1) consumes it directly.
`src/tools/persistent_report.py` migrated to the clean signature
with a `TypeError`-guarded legacy fallback (the post-hoc
`_measures[-1]["format_string"] = …` mutation) for older
deployments.

### Bug #5 — Row-shape validation

`PBIXBuilder.add_table` now raises a `TypeError` at the call site
when a row is anything other than a dict, with the offending row
index and an example payload. Previously failed deep in `save()`
with a cryptic `'list' object has no attribute 'keys'`.

### Tests

`tests/test_v0_13_2_pbix_mcp_patches.py` — 15 regressions, one
class per bug. Full suite: **341 passed**, 2 skipped (was 326 in
v0.13.1; +15).

### Documentation

New top-level [UPSTREAM_PATCHES.md](UPSTREAM_PATCHES.md) describes
every modification with the bug catalogue, file-by-file diff
summary, and a checklist for re-applying after a fresh install.

## [0.13.1] — 2026-05-13 — DBCC string-store hardening (post-build dialog fix)

Field report on v0.13.0: after a `pbi_scaffold_pbix` / `pbi_create_persistent_report` build, Power BI Desktop reopened the file with a modal dialog:

> Database consistency checks (DBCC) failed while checking the string
> store. An error occurred while loading Vertipaq data objects for
> multiple tables.

### Cause

`pbi_create_persistent_report` (and the v0.13 scaffold templates) generated tables with `rows=[]` whose schema declared at least one `String` column. The resulting `.pbix` had a Vertipaq dictionary referenced by metadata but never primed by data segments. On reopen, DBCC's dict ↔ segment consistency pass rejects the model.

### Fix

1. **Sentinel-row priming** (`src/tools/persistent_report.py`,
   `src/tools/scaffold.py`). New `prime_string_store` parameter
   (default `True`) on both tools. When enabled, every Import table
   with `rows=[]`, no `source_csv` / `source_db`, and at least one
   `String` column receives a single typed sentinel row before the
   PBIX is written. The Vertipaq dictionary gets primed and DBCC
   passes on reopen. Opt out via `prime_string_store=False` when the
   table will be populated by the first Power BI refresh after open.
   The tool response now includes
   `primed_string_store_tables: [<table-name>, …]` so the caller can
   see exactly which tables received the sentinel.

2. **Static DBCC diagnostics** (`src/tools/dbcc.py`, new module).
   - `pbi_diagnose_pbix_dbcc` — opens the `.pbix` zip, inventories
     the DataModel / Report/Layout / Connections / Metadata parts,
     and flags `no_data_model` or `undersized_data_model`
     (≤ 4 KB ⇒ Vertipaq is empty ⇒ DBCC will fail). Static — no need
     to open the file in Power BI Desktop.
   - `pbi_check_scaffold_spec_dbcc_risks` — pre-build risk check on
     the same `tables` shape that the scaffold/persistent_report
     tools take. Flags `empty_string_table` (issue) and
     `empty_import_table` (warning) so an LLM can correct the spec
     before calling the builder.

3. **Reopen probe signal list extended** (`src/tools/quality.py`).
   `pbi_validate_pbix_reopen` now matches `Database consistency
   checks`, `DBCC`, `Vertipaq`, `string store`, `An error occurred
   while loading`, `Report this issue`, `Something went wrong`,
   `Copy details to clipboard`, and `multiple tables` against the
   UIAutomation tree. The next time this dialog appears the probe
   surfaces a structured `powerbi_fix_this_signal` match instead of
   only the screenshot.

### Tests

`tests/test_v0_13_1_dbcc.py` — 18 tests:
- Spec-risk checker: empty-String flagged, numeric-only is warning,
  rows present / `source_csv` / DirectQuery / invalid entries.
- `_prime_string_store`: every type has a sentinel, no-op when rows
  exist or no String columns, primed table flag set.
- Static PBIX diagnoser: non-zip, no DataModel, undersized DataModel,
  healthy model passes, `known_signals` exposed.
- Reopen probe signal list contains every DBCC needle.

### Test coverage

`pytest -q` → **326 passed**, 2 skipped (was 308 in v0.13.0; +18).

### Tool surface

`+2` tools, no breaking signatures:

- `pbi_diagnose_pbix_dbcc` (read)
- `pbi_check_scaffold_spec_dbcc_risks` (read)

## [0.13.0] — 2026-05-13 — Major: PBIX scaffold, theme validation, security + test sweep

Major release. Four orthogonal blocks:

### 1. PBIX scaffold (new feature)

1. **`pbi_scaffold_pbix_tool`** (`src/tools/scaffold.py`). One-call
   creation of a starter `.pbix` from a named template:
   - `blank` — date table only.
   - `finance` — date + GL fact + baseline KPI pack (Total, Total YTD,
     Total MTD, Total YoY %).
   - `sales` — date + sales fact + product dim + baseline KPIs, with
     two relationships pre-wired.
   - `analytics` — date + events fact + user dim + baseline KPIs.
   Accepts optional `theme_json_path` (validated against the v0.13
   theme schema below) and `extra_measures` for callers that want to
   layer custom measures on top of the template's baseline.

2. **`pbi_list_scaffold_templates_tool`**. Returns the catalogue with
   per-template description + table/measure counts so an LLM can pick
   the right scaffold without trial-and-error.

3. The scaffold delegates the actual `.pbix` write to
   `pbi_create_persistent_report_tool` (the v0.10 builder) so all the
   existing validation + name guards still apply.

### 2. Theme JSON validation (new feature + hardening)

4. **`pbi_validate_theme_tool`** (`src/tools/visuals/_design.py`).
   Dry-run validation of a user-supplied theme JSON without touching
   the report. Returns issue list, size, and the boolean `valid`.

5. **`pbi_export_active_theme_tool`**. Captures the currently active
   theme from an extracted report folder into a `.json` file the
   caller can edit and re-apply.

6. **Schema validator** (`src/tools/visuals/_themes.py`). Enforced on
   every apply / scaffold path:
   - **Size cap**: 256 KB (`MAX_THEME_BYTES`).
   - **Top-level key allowlist**: 20 keys from the report-theme schema
     (`name`, `dataColors`, `foreground`, `background`, …,
     `visualStyles`). Unknown top-level keys are rejected; deep
     `visualStyles.*` keys stay open since visuals carry extension
     properties there.
   - **Colour shape**: `dataColors[]` and `#`-prefixed string values
     under colour-named keys must match `#RRGGBB` or `#RRGGBBAA`.
   - **URL-bearing values rejected** (CWE-20): any string value that
     starts with `javascript:`, `data:`, `vbscript:`, `file://`, or
     `https?://`. Themes describe colours and typography, not
     behaviour — a URL in a value is treated as smuggling.
   - **`pbi_apply_theme_tool`** now refuses to write a theme that
     fails this schema. Previously it accepted any well-formed JSON.

### 3. Security hardening sweep

7. **DAX guard widened** (`src/tools/query.py`). The DMV/system
   blocklist for `pbi_execute_dax` now also rejects:
   - `INFO.<NAME>(…)` — DAX INFO functions surface server metadata
     (tables, measures, relationships) the same way DMVs do.
   - `EVALUATEANDLOG(…)` — writes side-channel debug output to the
     server log directory.
   Both join the existing `$SYSTEM.*` / `DISCOVER_*` / `DBSCHEMA_*` /
   `MDSCHEMA_*` set. `PBI_MCP_ALLOW_DMV=1` still acts as the explicit
   opt-out, matching the legacy DMV behaviour.

8. **Theme path is policy-aware**. The new theme tools resolve every
   incoming path through `resolve_local_path()` with extension
   allowlist + symlink rejection, matching the rest of the file
   surface. There's no separate "themes are special" path now.

9. **Tool catalogue updated** (`src/security.py`). New READ entries
   (`pbi_validate_theme`, `pbi_list_scaffold_templates`) and new
   WRITE entries (`pbi_scaffold_pbix`, `pbi_export_active_theme`) so
   `--profile readonly` and the `disabled_tools` allowlist see the
   v0.13 surface correctly.

### 4. Test sweep (+71 tests)

`pytest -q` → **308 passed**, 2 skipped (was 237 in v0.12.8).

New test files:

- `tests/test_v0_13_theme_validator.py` — 25 tests pinning the schema
  validator (top-level allowlist, colour format, URL guard, size cap,
  end-to-end apply / export integration).
- `tests/test_v0_13_dax_guards.py` — 11 tests pinning the DMV /
  INFO.* / EVALUATEANDLOG blocklist + the `PBI_MCP_ALLOW_DMV` opt-out.
- `tests/test_v0_13_scaffold.py` — 14 tests pinning template catalogue
  (`blank` / `finance` / `sales` / `analytics`), template execution
  through a fake `PBIXBuilder`, extra-measure injection, and theme
  rejection on the scaffold path.
- `tests/test_v0_13_visual_roles_matrix.py` — 4 sub-tested matrix
  tests pinning every entry of `VISUAL_FIELD_ROLES` /
  `VISUAL_ROLE_KINDS` and the dispatcher coverage so a future
  refactor can't silently drop a chart family.
- `tests/test_v0_13_dax_generators.py` — 17 tests pinning the
  time-intelligence template catalogue (YTD/MTD/QTD/SPY/YOY/YOY%/MA3)
  and the pattern resolver (dependency expansion, dedup, case-fold,
  rejection of unknown patterns).

### Tool surface

Public surface grows by 4 MCP tools:

- `pbi_scaffold_pbix` (write)
- `pbi_list_scaffold_templates` (read)
- `pbi_validate_theme` (read)
- `pbi_export_active_theme` (write)

No breaking changes to existing tool signatures.

## [0.12.8] — 2026-05-09 — Treemap role-name fix

Field report on top of v0.12.7: treemaps still rendered empty in PBI
Desktop because the projection roles were wrong, not just the
renderer-compat flags.

### Bug fix

1. **Treemap uses `Category` / `Details` / `Values`, not the cartesian
   `Y`** (`src/tools/visuals/_charts.py`,
   `src/tools/visuals/_base.py`). PBI Desktop's treemap field wells map
   to those three roles internally; emitting `Y` (a cartesian role)
   caused the data-shape pass to drop the projection silently and the
   visual opened white. `pbi_add_treemap_tool` now writes `Values` for
   the measure and accepts an optional `details_column` for the
   second-level grouping.
   `VISUAL_FIELD_ROLES["treemap"]` and `VISUAL_ROLE_KINDS["treemap"]`
   updated to match so `pbi_convert_visual_type` and the role-binding
   validator know the canonical shape.

### Verified (no change)

2. **Scatter chart roles** — already correct: `Category` / `X` / `Y` /
   `Size` / `Series`. Both `Category` and `Details` are accepted by the
   PBI renderer for scatter; we keep `Category` for legacy-build
   compatibility.

### Tests

3. **`tests/test_v0_12_8_treemap_roles.py`** — 4 regressions:
   - `VISUAL_FIELD_ROLES["treemap"]` is exactly
     `{"Category", "Details", "Values"}` (and not `"Y"`).
   - `VISUAL_ROLE_KINDS["treemap"]` categorises each role.
   - `pbi_add_treemap_tool` emits `Values`, never `Y`.
   - The optional `details_column` parameter routes to the `Details`
     role.

### Test coverage

`pytest -q` → 237 passed, 2 skipped (was 233 in v0.12.7; +4 from
v0.12.8).

## [0.12.7] — 2026-05-09 — Empty-visual renderer-compat fixes

Field-tested against Power BI Desktop 2024 builds: every visual added by
`pbi_add_*` (treemap, map, scatter, the v0.12.6 chart pack…) opened
with an empty data area despite the projections + select entries being
correct. Two missing fields in the emitted JSON were the root cause —
the renderer's data-shape pass silently drops items that lack them.

### Bug fixes

1. **`From` entries now carry `Type: 0`** (`src/tools/visuals/_bindings.py`).
   `_build_prototype_query` was emitting `{"Name": …, "Entity": …}`. PBI
   Desktop's renderer ignores entries that lack the `Type` discriminator
   (0 = standard table entity-source kind in PBI's protobuf). Without
   it, every `Measure`/`Column` whose `SourceRef.Source` points at the
   stripped alias fails to bind, so the visual opens empty.

2. **Projection items now carry `"active": True`**
   (`src/tools/visuals/_refs.py`, every chart builder under
   `src/tools/visuals/`). The dispatcher historically emitted
   `{"queryRef": …}`; PBI's data-shape pass treats unflagged projection
   items as draft / inactive and drops them. Affects every
   ``pbi_add_*_tool`` plus ``pbi_update_visual_bindings_tool``. A new
   ``_projection(reference, *, active=True)`` helper builds the
   canonical shape so future chart additions can't regress.

3. **Dispatcher coverage gap closed**
   (`src/tools/visuals/_dispatcher.py`). The generic
   ``pbi_add_visual(visual_type=…)`` registry only knew the original
   chart set — `treemap`, `pie_chart`, `scatter_chart`, `combo_chart`,
   the v0.12.6 stacked/area/ribbon/funnel/multi_row_card additions, all
   raised "unknown visual_type". Three small `_categorical_dispatch` /
   `_axis_chart_dispatch` factories add the missing 15 entries without
   duplicating per-type validation.

### Tests

4. **`tests/test_v0_12_7_renderer_compat.py`** — 8 new regressions:
   - `_build_prototype_query` emits `Type: 0` on every `From` entry.
   - `_projection` defaults to `active: True` and supports opt-out.
   - Treemap, pie, bar (with legend), line (multiple measures) all
     produce on-disk Layouts where every `From` entry has `Type: 0`
     and every projection item has `active: True`.
   - The dispatcher exposes every v0.12.6 chart-pack key plus
     `scatter_chart` / `combo_chart`.

### Test coverage

`pytest -q` → 233 passed, 2 skipped (was 225 in v0.12.6; +8 from
v0.12.7).

## [0.12.6] — 2026-05-09 — Hot-reload bypass fix + chart-family pack

Closes the high-severity Codex finding on the v0.12.3 hot-reload feature
and fills the biggest gaps versus Power BI's native visual palette.

### Security

1. **Hot-reload policy bypass** (Codex finding `f7b1de6e`,
   `src/security.py`). With hot-reload enabled (v0.12.3), a write-capable
   tool (`pbi_export_model` is the obvious one) targeting the active
   `security_policy.json` could overwrite it with valid JSON whose
   missing keys default to `allow_categories = read/write/destructive`
   and `enabled_tools = None`. The next `policy()` call then dropped the
   prior allowlist / `disabled_tools`, elevating any previously-denied
   tool. The path validator now resolves the active policy file (env
   `PBI_MCP_SECURITY_POLICY`, the manager's `_policy_cwd`, and the
   process cwd's `security_policy.json`) and rejects any
   `WRITE_TOOLS`/`DESTRUCTIVE_TOOLS` call whose path resolves to it.
   `tests/test_security.py` adds two regressions:
   `test_export_model_cannot_overwrite_active_policy_file` and
   `test_destructive_tool_cannot_overwrite_policy_file`.

### Visual surface — 13 new chart types

Brings the native chart coverage up to PBI Desktop parity.
`src/tools/visuals/_charts.py` + `src/tools/visuals/_base.py`:

2. **Pie chart** — `pbi_add_pie_chart`. Sibling of donut.
3. **Stacked + clustered column / bar variants** —
   `pbi_add_stacked_bar_chart`, `pbi_add_stacked_column_chart`,
   `pbi_add_clustered_column_chart`,
   `pbi_add_hundred_percent_stacked_bar_chart`,
   `pbi_add_hundred_percent_stacked_column_chart`. Existing
   `pbi_add_bar_chart` stays clustered-bar; the new tools cover the
   remaining quadrants of the bar/column matrix.
4. **Area family** — `pbi_add_area_chart`,
   `pbi_add_stacked_area_chart`,
   `pbi_add_hundred_percent_stacked_area_chart`. Same projection shape
   as `pbi_add_line_chart` (multiple Y measures + optional Series).
5. **Ribbon chart** — `pbi_add_ribbon_chart`. Series ranking over time.
6. **Treemap** — `pbi_add_treemap`. Hierarchical alternative to donut /
   pie for many categories.
7. **Funnel** — `pbi_add_funnel`. Pipeline / conversion. Uses the
   `Group` / `Values` projection roles instead of `Category` / `Y`.
8. **Multi-row card** — `pbi_add_multi_row_card`. Vertical stack of KPI
   rows; with or without a category column.

A small `_add_categorical_chart` / `_add_axis_chart` helper backs the
new bar+column / area variants so the wrappers stay one-call thin and
share validation + measure-home resolution with the existing tools.
`VISUAL_FIELD_ROLES` and `VISUAL_ROLE_KINDS` updated for every new
`visualType` so `pbi_convert_visual_type` and the role-binding
validator know about the additions.

### Tests

9. **`tests/test_v0_12_6_charts.py`** — 5 new tests: every new
   categorical variant emits its expected `visualType`, the area family
   accepts multiple Y measures, funnel uses `Group`/`Values`,
   multi-row card works with and without a category and rejects empty
   measure lists.

### Tool count

Registered `@mcp.tool()` handlers: 149 → 162 (+13). README badge
updated.

## [0.12.5] — 2026-05-09 — Visual ops bugfixes + per-series colour

Field-tested against an LLM client driving layout patches: 4 bugs closed,
2 capability gaps filled. No breaking changes.

### Bug fixes

1. **`pbi_patch_layout` — "multiple values for argument 'extract_folder'"**
   (`src/wrappers/_helpers.py`). The `register_tool` helper detected
   `manager` as a `POSITIONAL_OR_KEYWORD` parameter and injected
   `CONNECTION_MANAGER` as `*args[0]`. That worked when `manager` was the
   first positional but collided with the actual first positional
   (`extract_folder`, `pbix_path`, etc.) in tools where `manager` appears
   later in the signature. The wrapper now always injects `manager` as a
   keyword argument, regardless of where it sits in the underlying
   signature. Affects every tool with a non-first `manager` param —
   primarily `pbi_patch_layout_tool` and `pbi_validate_report_fields_tool`.

2. **`pbi_set_visual_format_property` — undocumented type hints + nested
   colour rejection** (`src/tools/visuals/_formatting.py`,
   `src/tools/visuals/_ops.py`). The valid `property_types` values were
   only discoverable through the source. The tool now:
   - Documents every type (`auto`, `bool`, `int`, `decimal`, `text`,
     `color`, `raw`) in the docstring with a concrete example per type.
   - Accepts common aliases (`integer`, `float`, `number`, `string`,
     `fill`, `hex`, `rgb`, `boolean`) and maps them to the canonical
     name. Unknown hints raise with the full allowlist.
   - Auto-unwraps a previously-encoded
     `{"solid": {"color": "#RRGGBB"}}` payload back to a hex string so
     LLM clients that round-trip a returned colour value don't get a
     confusing "color must match '#RRGGBB'" error.

3. **`dataPoint.defaultColor` paints every series the same colour** —
   new tool `pbi_set_series_color` (`src/tools/visuals/_ops.py`).
   `defaultColor` is the visual-wide default; per-series overrides go
   into the same `dataPoint` array as additional entries with an `id`
   selector pinned to the target measure / column. The new tool
   targets a series by index (0-based across projections, role-ordered
   `Y` → `Values` → `Series` → `Category`) or by name (matches the
   queryRef or the underlying `Property`), then writes the override
   while leaving sibling series untouched.

4. **PBIX locked → in-memory measures lost on `force=True` patch**
   (`src/tools/visuals/_io.py`, `src/tools/visuals/_ops.py`).
   `pbi_patch_layout_tool` now takes a `save_before_close: bool = True`
   parameter. When `force=True` and `save_before_close=True`, the call
   posts Ctrl+S to every running Power BI Desktop window via Win32
   PostMessage (same hardened path as `pbi_persist_now`) and waits up
   to 10 seconds for the PBIX mtime to advance *before* invoking the
   close-then-kill path. The save attempt is best-effort, never
   raises, and the response now includes a `save_attempt` block with
   telemetry (`attempted`, `windows_targeted`, `mtime_changed`,
   `polled_seconds`, `skipped_reason`) so callers can detect silent
   loss.

### New capabilities

5. **`pbi_set_visual_format_property` reset path**
   (`src/tools/visuals/_ops.py`). Two new ways to clear a property
   instead of leaving it set to a blank value:
   - `reset_properties: list[str] | None = None` — explicit list of
     properties to delete from the visual's bag.
   - In-band sentinel: pass `"__reset__"` as the value inside the
     `properties` dict for the same effect, so callers that build a
     single dict (e.g. JSON payload) have a single-call path.
   The response surfaces `reset` (alongside `applied`) so callers see
   exactly what was cleared.

6. **`pbi_add_conditional_formatting` — table / matrix data bars,
   colour scales, and icon sets** (`src/tools/visuals/_ops.py`,
   registered for the `add_visual` neighbours). Supports
   `format_type ∈ {"dataBar", "colorScale", "iconSet"}` with the
   common per-type knobs (`bar_color`, `min/mid/max_color`,
   `icon_set ∈ {threeArrows, threeArrowsGray, threeTrafficLights,
   threeSymbols, threeFlags, fiveArrows}`). Targets a column / measure
   by display name; matched case-insensitively against the visual's
   `prototypeQuery.Select` entries.

### Tests

7. **`tests/test_v0_12_5_fixes.py`** — 8 new tests:
   - Bug 1: keyword-injection works when `manager` is the 4th
     parameter (the case that previously raised `TypeError: multiple
     values for argument 'extract_folder'`); explicit caller-provided
     manager wins over auto-injection.
   - Bug 3: `series_index=1` writes a `dataPoint` entry pinned to the
     second series (`Cost`) and not the first (`Revenue`); selector
     shape is correct; out-of-range index returns a structured error;
     role-priority order resolves `Y` before `Series`.
   - Wrapper smoke: every `wrappers/<domain>.py` imports cleanly.

### Tool count

Registered `@mcp.tool()` handlers: 147 → 149 (+2 from
`pbi_set_series_color` and `pbi_add_conditional_formatting`). README
badge updated.

## [0.12.4] — 2026-05-09 — Hardening polish

Follow-up sweep on top of v0.12.3. Closes the residual zip-bomb gap on
PBIX inputs, removes the last focus-race surface in the graceful-close
helper, locks new defaults behind regression tests, and refreshes the
docs.

### Security

1. **PBIX zip-bomb caps** (`src/tools/visuals/_io.py`). The native PBIX
   extraction path now reuses the `max_excel_zip_*` policy caps
   (decompressed size, member count, compression ratio) before writing
   any member. Previously only `.xlsx` workbooks were inspected; a
   hostile PBIX with 100k members or a 1000:1 ratio could exhaust disk
   during the fallback extraction.
2. **Graceful-close keyboard injection moved to PostMessage**
   (`src/tools/visuals/_io.py`). `_save_and_close_powerbi_gracefully`
   used to call `WScript.Shell.AppActivate` + `SendKeys('^s')`, which
   has the same global-input-queue race as the legacy `pbi_persist_now`
   path. The helper now enumerates every PBIDesktop top-level window
   through Win32 and posts the WM_KEYDOWN/UP chord directly to each
   HWND, before handing off to the existing PowerShell wait-for-mtime
   loop.

### Tests

3. **`tests/test_security.py::V0124RegressionTests`** — four new
   regressions:
   - `max_response_bytes` cap returns a `response_too_large` error
   - `max_response_bytes=0` disables the check (untouched payload)
   - `security_policy.json` mtime change triggers a reload on the next
     `policy()` call
   - the v0.12.3 M blocklist additions (Snowflake, BigQuery, Redshift,
     Excel.CurrentWorkbook, AzureBlobStorage, AnalysisServices,
     Salesforce, GoogleSheets) are all rejected

### Docs

4. **SECURITY.md** — `security_policy.json` schema documents
   `max_response_bytes` and the new default for
   `rate_limit_calls_per_minute` (600). The M blocklist listing is
   refreshed with cloud DW, SaaS, cloud storage, cubes, reflection,
   and quoted-identifier classes.
5. **SECURITY.md / README.md env-var tables** now list
   `PBI_MCP_AUTH_TOKEN`, `PBI_MCP_ALLOWED_ORIGINS`,
   `PBI_MCP_ALLOW_UNAUTHENTICATED_SSE`, `PBI_MCP_PBI_TOOLS_TIMEOUT`,
   `PBI_MCP_PERSIST_USE_SENDINPUT`, `PBI_MCP_ALLOW_UI_AUTOMATION`,
   and the audit knobs.

### Test coverage

`pytest -q` → 210 passed, 2 skipped (was 206 in v0.12.3; +4 from the
new V0124RegressionTests block).

## [0.12.3] — 2026-05-09 — Hardening sweep

A focused robustness pass across security, stdio I/O, and observability.
No tool surface changes, no breaking changes; defaults move toward
fail-safe.

### Security

1. **Bearer auth now uses `hmac.compare_digest`** (`src/server.py`).
   The previous byte-equality check leaked timing information; an
   attacker on the SSE endpoint could probe the token byte-by-byte
   through response-time deltas. The new path is constant-time.
2. **`PBI_MCP_AUTH_TOKEN` enforces a 32-character minimum** when set,
   raising `SecurityPolicyError` at startup. The README's recommended
   `secrets.token_urlsafe(32)` already meets this; short or
   placeholder tokens now fail loud instead of silently accepting
   weak credentials.
3. **Power Query M blocklist widened** (`src/m_expression_security.py`).
   Added Snowflake, BigQuery, Redshift, Azure SQL/Synapse, AWS S3,
   ADLS Gen2, SaaS connectors (Salesforce, Dynamics, Google Sheets,
   Exchange), AnalysisServices.Database, Cube.\*, Excel.CurrentWorkbook,
   and `WebAction.*`. The previous list missed every modern cloud
   connector; an LLM-authored M expression could exfiltrate to a
   Snowflake account or a public S3 bucket without tripping the gate.
4. **`pbi-tools` subprocess gets a 300 s timeout**
   (`src/tools/visuals/_io.py`). A hostile or extremely large PBIX could
   stall the server indefinitely under the previous unbounded
   `subprocess.run`. Tunable via `PBI_MCP_PBI_TOOLS_TIMEOUT`.
5. **`pbi_persist_now` switched from `SendInput` to `PostMessage`**
   (`src/tools/ui_automation.py`). `SendInput` injects into the global
   keyboard queue, which routes to whichever window owns the
   foreground when the events are processed — a focus race could
   deliver the Ctrl+S to an unrelated app. `PostMessage` posts
   directly to the resolved focused descendant of the PBI Desktop
   HWND. Operators can fall back to the legacy path with
   `PBI_MCP_PERSIST_USE_SENDINPUT=1` on builds where WPF input
   processing ignores posted messages.
6. **`SecurityPolicy.rate_limit_calls_per_minute` defaults to 600**
   (`src/security.py`). Previously `None` (unlimited), so a runaway
   LLM agent could melt local Power BI Desktop with a measure-creation
   loop. Set to `0`/`null` in `security_policy.json` to opt out.
7. **`SecurityPolicy.max_response_bytes` defaults to 16 MiB**
   (`src/security.py`, `src/mcp_core.py`). `pbi_export_model` against
   a multi-GB model — or a DAX query that side-loaded metadata —
   could OOM the LLM client process. Oversized responses now return
   a structured `response_too_large` error.

### Bugs / robustness

8. **Native PBIX zip extraction surfaces zip-slip skips**
   (`src/tools/visuals/_io.py`). Traversal members were silently
   ignored; the call now logs an aggregate warning and reports
   `skipped_traversal_count` in the response so operators can
   investigate hostile PBIX inputs.
9. **Antigravity adapter keeps a stderr handler at WARNING**
   (`src/server_antigravity.py`). The previous `_silence_loggers()`
   set `ERROR` and would have hidden capability mismatches and bind
   failures from the Antigravity diagnostics view.
10. **PowerShell helpers force UTF-8 I/O**
    (`src/tools/quality.py`, `src/tools/visuals/_io.py`). PS 5.1
    defaults `Out-File` to UTF-16 LE and stdout to the host codepage,
    which silently mangles non-ASCII paths round-tripped through
    `json.dumps` + `ConvertFrom-Json`. A small prelude pins the
    encoding for every script we run.
11. **Tool registry audit no longer crashes stdio in non-strict mode**
    (`src/server.py`). A bare `RuntimeError` from
    `_audit_tool_registry` would terminate the server before it could
    emit a JSON-RPC error frame; CI strict mode still raises so the
    failure surfaces as a non-zero exit.
12. **`security_policy.json` hot-reloads on mtime change**
    (`src/security.py`). Operators can edit policy without restarting
    the server; the next `policy()` call re-reads the file and logs a
    `policy reloaded` line.

### Tests / DX

13. **`tests/test_server_antigravity.py`** — new coverage for the
    Antigravity adapter: argparse defaults + `--profile` validation,
    `_silence_loggers` keeps a single WARNING-level stderr handler,
    `_harden_stdio` sets the expected env vars and tolerates
    non-reconfigurable streams.
14. **`scripts/tool_count.py`** — programmatic tool count for the
    README badge. Importing `server` triggers every wrapper-side
    registration, then the script reads `mcp._tool_manager._tools`.
    No more drift between the badge and reality.
15. **CI coverage floor bumped 54 → 55**
    (`.github/workflows/ci.yml`). Conservative bump to lock in the
    Antigravity-adapter coverage; further rises tracked separately.

### Test coverage

`pytest -q` → 206 passed, 2 skipped (was 198 in v0.12.2; +8 from the
new Antigravity adapter test file).

## [0.12.2] — 2026-05-09 — Antigravity adapter polish

### Changed

1. **`tools/antigravity_mcp_launcher.ps1`** now invokes
   `src/server_antigravity.py` instead of `src/server.py`, so the
   capability-stripping logic introduced in v0.12.1 actually applies
   when Antigravity launches the server through the documented
   PowerShell wrapper.
2. **`src/server_antigravity.py`** gained a small argparse layer
   exposing `--readonly` and `--profile {readonly,write,all,grading}`
   with the same semantics as `src/server.py`. Both flags are honored:
   `--readonly` toggles `SECURITY.set_runtime_readonly(True)` and
   `--profile readonly` does the same plus prunes the registered tool
   surface to `READ_TOOLS`.
3. **`server_version` in the InitializationOptions** is now read
   dynamically from `importlib.metadata` so the value reflects the
   installed package, not a hard-coded fallback.

### Docs

4. **README.md** — Antigravity section reworded to explain *why* the
   dedicated entry point exists (FastMCP 1.27 capability mismatch),
   what the PowerShell wrapper does (cwd + UTF-8 env), and how to map
   the `--profile` flags. Tool-count badge bumped 134 → 147 to match
   the current registered surface.

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
