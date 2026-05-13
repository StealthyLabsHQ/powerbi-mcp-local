# Upstream patches — pbix-mcp 0.9.2

Tracks every modification this repo applies to its `pbix-mcp` dependency
beyond the published 0.9.2 release. The patches live in
`.venv/Lib/site-packages/pbix_mcp/` and are re-applied on every fresh
install via the regression tests in
`tests/test_v0_13_2_pbix_mcp_patches.py` (they assert the patches are
in place — a re-install without the patches turns the suite red).

Goal: a single, reviewable record of "what we changed in the vendored
library and why" so anyone can re-apply them after `pip install
--force-reinstall pbix-mcp` or generate an upstream PR.

| # | Bug | Status in 0.9.2 | Patch |
|---|---|---|---|
| 1 | `Measure.FormatString` always `NULL` | already fixed | regression test pins it |
| 2 | DBCC string-store corruption on `HASONEVALUE+VALUES` over embedded tables | **not fixed upstream** — root cause in the Vertipaq encoder, not addressable from Python | new `pbix_mcp/dbcc_guard.py` scans measures, emits `DBCCRiskWarning` on `save()` and surfaces findings in `_pre_build_checks()` |
| 3 | Visual config schema too narrow (`series`, `objects`, `vcObjects`, page `config`) | already fixed | regression tests pin each pass-through |
| 4 | `add_measure()` doesn't accept `format_string=` | new patch | added keyword-only `format_string=None` parameter to `PBIXBuilder.add_measure`; `src/tools/persistent_report.py` migrated to the clean signature with a legacy fallback |
| 5 | Cryptic `'list' object has no attribute 'keys'` deep in `save()` | new patch | added early row-shape validation in `PBIXBuilder.add_table` that raises a `TypeError` with the offending row index + an example payload |

## File-by-file diff summary

### `pbix_mcp/builder.py`

`add_table` (Bug #5): early row-shape validation rejecting non-dict rows
with a usable `TypeError`.

`add_measure` (Bug #4): added `format_string: str | None = None` and
propagated to `self._measures` so the SQL INSERT (Bug #1) can persist it
directly.

`_pre_build_checks` (Bug #2): tail invocation of
`scan_measures_for_dbcc_risks(...)` surfaces matches as
`WARNING: DBCC string-store risk in measure '<name>' ...` so callers see
the issue in `pre_build_issues` without `save()`.

`save` (Bug #2): wraps `build()` in a `scan_measures_for_dbcc_risks` +
`emit_runtime_warnings` call so a `DBCCRiskWarning` is always issued
before disk-write when the pattern matches. Scanner failures are
swallowed — guarding never blocks a save.

### `pbix_mcp/dbcc_guard.py` (new file)

`scan_measures_for_dbcc_risks(measures, tables)` — pattern catalogue:

- `hasonevalue_values_string` — `HASONEVALUE(T[col])` paired with
  `VALUES(T[col])` on a string column.
- `selectedvalue_string_default` —
  `SELECTEDVALUE(T[col], "<default>")` over a string column.
- `treatas_string` — `TREATAS(VALUES(T[StringCol]), …)`.

Risk is only reported when the referenced table is in our embedded-rows
set (no `source_csv` / `source_db`, mode≠`directquery`). Tables sourced
externally use a different encoder path that AS accepts.

`emit_runtime_warnings(findings)` issues one
`warnings.warn(..., DBCCRiskWarning)` per finding.

## Re-applying after a fresh install

```powershell
.\.venv\Scripts\python.exe -m pytest tests/test_v0_13_2_pbix_mcp_patches.py -v
```

A failure on any test means a patch is missing — open
`.venv/Lib/site-packages/pbix_mcp/builder.py` (and `dbcc_guard.py`) and
re-apply the corresponding hunk. The test names map 1:1 to the bug
catalogue above.

## Mitigations for Bug #2 (DBCC string-store)

When the scanner flags a measure:

1. **Source the affected table from CSV / DB.** Embedded `rows=[…]`
   goes through our serializer; `source_csv` / `source_db` does not.
2. **Replace string lookups with Int64 surrogates.** Add a numeric ID
   column to the dimension and reference it via `LOOKUPVALUE`. The
   numeric dictionary encoder is independent of the string store.
3. **Run `pbi-tools roundtrip`** on the produced `.pbix`. AS rebuilds
   the segments / dictionaries and DBCC then passes.

The runtime warning includes the same suggestions inline so a caller
that pipes warnings into logs can act without reading this file.

## Upstream PR (TODO)

Open an upstream PR against `github.com/d0nk3yhm/pbix-mcp` carrying:

- the `add_measure` keyword + the matching SQL INSERT placeholder (Bug
  #1 + Bug #4 are coupled),
- the row-shape validation (Bug #5),
- the DBCC guard module + `save()` / `_pre_build_checks` hook (Bug #2 —
  defensive surface, not a fix).
