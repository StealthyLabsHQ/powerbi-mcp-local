# Changelog

All notable changes to this project are documented here.

## [Unreleased] — 2026-04-16

### Fixed
- `pbi_build_dashboard`: all `queryRef` values now use `_query_ref()` — table-prefixed refs (e.g. `FaitsCA.Annee`) correctly emit the short column name
- `pbi_build_dashboard`: gauge projection role corrected (`"Value"` → `"Y"`)
- Individual visual tools (`pbi_add_bar_chart`, `pbi_add_line_chart`, `pbi_add_donut_chart`, `pbi_add_table_visual`, `pbi_add_waterfall`, `pbi_add_slicer`, `pbi_add_gauge`): same `_query_ref()` fix applied
- `_build_select_entry`: `Select[Name]` now emits the short column name instead of the full `Table.Column` reference — required by Power BI's prototypeQuery parser
- `pbi_add_gauge_tool`: projection role `"Value"` → `"Y"` (correct PBI gauge data role)
- Security policy path resolution, pbi-tools discovery, and Layout ZIP extraction
- Windows: support split DLL locations and alternate assembly names for ADOMD.NET

### Added
- `_query_ref()` helper — strips table prefix from any `Table.Column` reference, returns column name only

### Refactored
- Project reorganized into `src/` / `tests/` / `specs/` structure

---

## [0.5.0] — 2026-04-16 — Visual Layer (56 tools)

### Added
- 20 visual tools via pbi-tools extract/compile pipeline:
  `pbi_extract_report`, `pbi_compile_report`, `pbi_create_page`, `pbi_delete_page`,
  `pbi_get_page`, `pbi_list_pages`, `pbi_set_page_size`, `pbi_add_card`,
  `pbi_add_bar_chart`, `pbi_add_line_chart`, `pbi_add_donut_chart`, `pbi_add_gauge`,
  `pbi_add_slicer`, `pbi_add_table_visual`, `pbi_add_waterfall`, `pbi_add_text_box`,
  `pbi_move_visual`, `pbi_remove_visual`, `pbi_apply_theme`, `pbi_build_dashboard`

### Fixed
- English-only specs and tests (removed French content)

### Docs
- Visual layer spec added
- Windows setup guide (zero to operational)
- README redesigned with badges, collapsible config, pipeline diagram

---

## [0.4.0] — 2026-04-16 — Security Hardening

### Added
- `security.py` middleware — path traversal, DAX injection, SSRF protection
- 15 security tests

### Fixed
- 7 vulnerabilities: path traversal, DAX injection, SSRF, logging exposure

---

## [0.3.0] — 2026-04-16 — Multi-Platform + Power Query v2 (36 tools)

### Added
- SSE transport support (multi-platform)
- Power Query v2: CSV import, folder import
- `pyproject.toml` packaging

---

## [0.2.0] — 2026-04-16 — Excel + Power Query (33 tools)

### Added
- 13 Excel tools: `excel_create_workbook`, `excel_read_sheet`, `excel_write_range`, etc.
- `pbi_list_instances`
- 4 Power Query (M) tools: `pbi_get_power_query`, `pbi_set_power_query`, `pbi_list_power_queries`, `pbi_create_import_query`

---

## [0.1.0] — 2026-04-16 — Initial Release (15 tools)

### Added
- Core MCP server implementation
- 15 tools: model inspection, DAX measure management, relationships, refresh, DAX query execution
- `pbi_connect`, `pbi_model_info`, `pbi_list_tables`, `pbi_list_measures`, `pbi_create_measure`,
  `pbi_delete_measure`, `pbi_execute_dax`, `pbi_refresh`, `pbi_list_relationships`,
  `pbi_create_relationship`, `pbi_export_model`, `pbi_set_format`, `pbi_create_column`,
  `pbi_create_table`, `pbi_import_dax_file`
