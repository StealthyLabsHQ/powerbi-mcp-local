# Plugins — MCP tools exposed by powerbi-mcp-local

80 tools grouped by domain. Each tool is wired through FastMCP in `src/server.py`.

Startup flags:
- `--transport stdio|sse` (default: stdio)
- `--readonly` — block write/destructive categories at runtime
- `--profile readonly|write|all` — prune tool registry at boot (smaller surface)
- `PBI_MCP_AUTH_TOKEN` env → Bearer auth on SSE transport

---

## Connection & discovery (2)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_connect` | read | Connects to running PBI Desktop instance |
| `pbi_list_instances` | read | Lists all discovered PBI Desktop ports |

## Model inspection (5)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_list_tables` | read | Tables + columns + optional row counts |
| `pbi_list_measures` | read | All DAX measures |
| `pbi_list_relationships` | read | All relationships with cardinality + direction |
| `pbi_model_info` | read | Full snapshot (tables + measures + relationships) |
| `pbi_refresh_metadata` | read | Reload cached TOM schema without full reconnect |

## DAX query & diagnostics (4)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_execute_dax` | read | Execute DAX (`max_rows`, `timeout_seconds`) |
| `pbi_execute_dax_as_role` | read | Run under a specific RLS role |
| `pbi_trace_query` | read | Execute + performance diagnostics |
| `pbi_validate_dax` | read | Parse-check expression (`kind=scalar|table`) |
| `pbi_measure_dependencies` | read | DISCOVER_CALC_DEPENDENCY rows |

## Measure CRUD (5)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_create_measure` | write | Create or update a measure |
| `pbi_rename_measure` | write | Rename (callers must update downstream DAX) |
| `pbi_delete_measure` | destructive | Drop a measure |
| `pbi_set_format` | write | Batch format string on measures or columns |
| `pbi_import_dax_file` | write | Bulk import from `.dax` file |

## Table & column CRUD (7)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_create_table` | write | Create or replace a calculated table |
| `pbi_rename_table` | write | Rename a table |
| `pbi_delete_table` | destructive | Drop a table |
| `pbi_create_column` | write | Create or update a calculated column |
| `pbi_rename_column` | write | Rename a column |
| `pbi_delete_column` | destructive | Drop a column |
| `pbi_export_model` | write | Export full model JSON (optional file output) |

## Relationships (4)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_create_relationship` | write | Create by from/to columns |
| `pbi_update_relationship` | write | Mutate cardinality/direction/is_active/name |
| `pbi_delete_relationship` | destructive | Drop by name or endpoint columns |
| `pbi_refresh` | write | Refresh a specific table or the whole model |

## Power Query / M (7)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_list_power_queries` | read | Partitions + source expressions |
| `pbi_get_power_query` | read | Read a table's M expression |
| `pbi_set_power_query` | write | Replace a partition's M expression |
| `pbi_create_import_query` | write | Inject an Excel-import M for a table |
| `pbi_create_csv_import_query` | write | Inject a CSV-import M |
| `pbi_create_folder_import_query` | write | Inject a folder-scan M |
| `pbi_bulk_import_excel` | write | Bulk rewrite multiple tables from one workbook |

## Row-level security (6)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_list_roles` | read | Roles + members + filters |
| `pbi_create_role` | write | Create or update a role (permission level) |
| `pbi_delete_role` | destructive | Drop a role |
| `pbi_set_role_filter` | write | Apply or clear a DAX filter on a table for a role |
| `pbi_add_role_member` | write | Add external/windows principal to a role |
| `pbi_remove_role_member` | write | Remove a principal by member name |

## Calculation groups (3)
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_list_calc_groups` | read | Groups + items |
| `pbi_create_calc_group` | write | Create or replace a calc group with items |
| `pbi_delete_calc_group` | destructive | Drop a calc group table |

## Excel workbook (13)
| Tool | Category | Summary |
| --- | --- | --- |
| `excel_create_workbook` | write | New workbook with optional sheets |
| `excel_list_sheets` | read | Sheets with row/col counts |
| `excel_workbook_info` | read | Full workbook metadata |
| `excel_create_sheet` | write | Add a sheet |
| `excel_delete_sheet` | destructive | Remove a sheet |
| `excel_read_cell` | read | Single cell value |
| `excel_read_sheet` | read | Rows with optional range + limit |
| `excel_search` | read | Search values across one or all sheets |
| `excel_write_cell` | write | Write single cell |
| `excel_write_range` | write | Write 2D array at start cell |
| `excel_format_range` | write | Font/fill/border/number format |
| `excel_auto_width` | write | Auto-fit column widths |
| `excel_to_pbi_check` | read | Compare workbook with active PBI model |

## Report & visual layer (22)

**Core (3)**
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_extract_report` | write | Unpack a `.pbix` via pbi-tools (requires legacy pbi-tools extract action, not available in pbi-tools.core 1.2.0) |
| `pbi_compile_report` | write | Rebuild a `.pbix` from an extracted folder |
| `pbi_patch_layout` | write | Inject modified `Report/Layout` into an existing `.pbix` via native zip (no pbi-tools needed) |

**Pages (5)**
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_list_pages` | read | Enumerate pages in an extracted report |
| `pbi_get_page` | read | Page details + visuals |
| `pbi_create_page` | write | Add a page |
| `pbi_set_page_size` | write | Resize a page |
| `pbi_delete_page` | destructive | Remove a page |

**Visuals (13)**
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_add_visual` | write | Unified dispatcher (`visual_type` + `config`) — covers all types below |
| `pbi_add_card` | write | Legacy: single-value card |
| `pbi_add_bar_chart` | write | Legacy: clustered bar |
| `pbi_add_line_chart` | write | Legacy: multi-measure line |
| `pbi_add_donut_chart` | write | Legacy: donut |
| `pbi_add_table_visual` | write | Legacy: table visual |
| `pbi_add_waterfall` | write | Legacy: waterfall |
| `pbi_add_slicer` | write | Legacy: slicer (`slicer_type` dropdown/list) |
| `pbi_add_gauge` | write | Legacy: gauge (optional target measure) |
| `pbi_add_text_box` | write | Legacy: formatted text box |
| `pbi_move_visual` | write | Reposition and/or resize |
| `pbi_remove_visual` | destructive | Delete by visual id |

**Design (3)**
| Tool | Category | Summary |
| --- | --- | --- |
| `pbi_apply_theme` | write | Load a theme JSON into the report |
| `pbi_apply_design` | write | Apply a preset (theme + page background + card styling) |
| `pbi_build_dashboard` | write | Build a dashboard page from a bulk layout spec |

---

## Category summary

| Category | Count |
| --- | --- |
| read | 24 |
| write | 44 |
| destructive | 12 |
| **Total** | **80** |
