# Changelog

All notable changes to this project are documented here.

## [Unreleased] — 2026-04-16

### Added
- `pbi_patch_layout`: direct PBIX Layout patch tool that swaps `Report/Layout`, removes `SecurityBindings`, preserves ZIP entry metadata, and supports `force=True` auto-kill on Windows.

### Fixed
- Visual query generation now resolves measure home tables from extract metadata (`Model/tables/*/measures/*.dax`) so `prototypeQuery.From[]` uses the real table entity instead of `"$Measures"` when available.
- Measure fallback now logs a warning and uses `"$Measures"` only when extract metadata is missing or disconnected.
- `.dax` import parsing now strips `//` and `/* ... */` comments safely while preserving quoted strings and measure block boundaries.
- `pbi_compile_report` now accepts `force: bool = False` and can auto-kill `PBIDesktop.exe` on Windows before write operations.

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
