# Suggested Power BI commands

Ready-to-use prompts for common Power BI automation workflows. Paste any
prompt into an MCP client session that has the `powerbi-desktop` server
configured. The MCP takes care of the TOM/DAX plumbing.

Legend:
- **Read-only** — safe on any model
- **Write** — mutates the model in memory; saved only if the user hits Ctrl+S in PBI Desktop
- **Destructive** — permanently removes objects

---

## 1. Onboarding & audit

### 1.1 Model tour (read-only)
```
Connect to Power BI Desktop. Give me a compact audit of the active model:
- Table count + row counts for each fact table
- Measure count per table and anything obviously misnamed
- Relationship graph (from -> to + cardinality)
- Any table with multiple partitions or compressed measure dependencies
End with 3 concrete improvement suggestions.
```
Tools: `pbi_connect`, `pbi_list_tables(include_row_counts=True)`,
`pbi_list_measures`, `pbi_list_relationships`, `pbi_measure_dependencies`.

### 1.2 Dead measure scan (read-only)
```
Find measures that are not referenced by any other measure or visual.
For each, show the definition and suggest whether to keep, rename, or delete.
```
Tools: `pbi_list_measures`, `pbi_measure_dependencies`, `pbi_list_pages` + `pbi_get_page` (scan visuals).

### 1.3 Export full documentation
```
Export the full model definition to model.json in ./docs/,
then summarize the top 10 most complex measures.
```
Tools: `pbi_export_model`, `pbi_list_measures`.

---

## 2. Measure authoring

### 2.1 Generate base aggregates (write)
```
For every numeric column in table FaitsCA, create a pair of measures:
- "<col> Total" = SUM
- "<col> Avg"   = AVERAGE
Format: "#,##0.00". Place them in display folder "Aggregates".
```
Tools: `pbi_list_tables`, `pbi_create_measure` (per column).

### 2.2 Time intelligence kit (write)
```
Assuming Dim_Temps is marked as date table, generate for measure "CA":
- CA MTD  = TOTALMTD([CA], Dim_Temps[Date])
- CA QTD  = TOTALQTD(...)
- CA YTD  = TOTALYTD(...)
- CA YoY  = CALCULATE([CA], SAMEPERIODLASTYEAR(Dim_Temps[Date]))
- CA YoY% = DIVIDE([CA] - [CA YoY], [CA YoY])
Validate each with pbi_validate_dax before saving.
```
Tools: `pbi_validate_dax`, `pbi_create_measure`, `pbi_set_format`.

### 2.3 Safe rename with impact check (write)
```
Rename measure Sales -> "Total Sales".
Before renaming: run pbi_measure_dependencies on "Sales" and list every
measure/visual that references it. If any consumer exists, pause and
report; otherwise rename and continue.
```
Tools: `pbi_measure_dependencies`, `pbi_rename_measure`.

### 2.4 Bulk import from .dax file (write)
```
Import the measures defined in ./dax/kpis.dax into table "Measures",
overwriting any with the same name. Stop on first parse error.
```
Tools: `pbi_import_dax_file`.

---

## 3. Star schema & relationships

### 3.1 Detect candidate relationships (read-only)
```
Scan all tables. Propose missing relationships based on matching
column names (e.g. Dim_Geo[Region] <-> FaitsCA[Region]) and unique-value
heuristics. Do not create anything — list recommendations with cardinality.
```
Tools: `pbi_list_tables`, `pbi_execute_dax` (COUNTROWS / DISTINCTCOUNT per candidate).

### 3.2 Create star schema (write)
```
Given dimension tables Dim_Temps, Dim_Geo, Dim_Famille and fact FaitsCA,
create manyToOne relationships from fact to each dimension on matching
column names. All single-direction. Skip if a relationship already exists.
```
Tools: `pbi_list_relationships`, `pbi_create_relationship`.

### 3.3 Relationship hygiene (destructive)
```
List inactive relationships. For each, decide keep or delete based on
whether any measure uses USERELATIONSHIP on it. Report first; apply
deletions only after explicit confirmation.
```
Tools: `pbi_list_relationships`, `pbi_list_measures`, `pbi_delete_relationship`.

---

## 4. Calculation groups

### 4.1 Time-intelligence calc group (write)
```
Create calculation group "Time Intelligence" (column "Period") with items:
- Current   = SELECTEDMEASURE()
- MTD       = CALCULATE(SELECTEDMEASURE(), DATESMTD(Dim_Temps[Date]))
- QTD       = CALCULATE(SELECTEDMEASURE(), DATESQTD(Dim_Temps[Date]))
- YTD       = CALCULATE(SELECTEDMEASURE(), DATESYTD(Dim_Temps[Date]))
- YoY       = CALCULATE(SELECTEDMEASURE(), SAMEPERIODLASTYEAR(Dim_Temps[Date]))
- YoY Delta = SELECTEDMEASURE() - CALCULATE(SELECTEDMEASURE(), SAMEPERIODLASTYEAR(Dim_Temps[Date]))
Precedence 10.
```
Tools: `pbi_create_calc_group`.

### 4.2 Currency calc group (write)
```
Create calc group "Currency" with items Actual / Budget / Variance,
each applying an appropriate CALCULATE filter against Dim_Scenario.
```
Tools: `pbi_create_calc_group`.

---

## 5. Row-level security (RLS)

### 5.1 Regional RLS template (write)
```
Create a role "Region_North" with Read permission:
- Filter on Dim_Geo[Region]: [Region] = "North"
- Member: north-team@company.com (external, AzureAD)
Then verify with pbi_execute_dax_as_role that COUNTROWS(FaitsCA)
differs from the unfiltered count.
```
Tools: `pbi_create_role`, `pbi_set_role_filter`, `pbi_add_role_member`,
`pbi_execute_dax_as_role`.

### 5.2 Bulk role provisioning (write)
```
Given the list of regions in Dim_Geo, create one Read role per region,
each with a filter restricted to its region. Print the role matrix.
```
Tools: `pbi_execute_dax` (fetch region list), `pbi_create_role`,
`pbi_set_role_filter`.

### 5.3 RLS audit (read-only)
```
For every role, list:
- permission level
- member count (and member names)
- which tables have a filter + the filter expression
Flag any role without members, or any filter that references a missing column.
```
Tools: `pbi_list_roles`, `pbi_list_tables`.

---

## 6. Power Query & data source hygiene

### 6.1 Repoint model to a new Excel file (write)
```
All tables currently load from ./old_data.xlsx. Repoint them to
./new_data.xlsx by sheet name. Use pbi_bulk_import_excel, then refresh
each table and report any schema mismatches.
```
Tools: `pbi_bulk_import_excel`, `pbi_refresh`.

### 6.2 CSV migration (write)
```
Table "Orders" currently imports from an xlsx. Migrate it to a CSV at
./data/orders.csv with headers, UTF-8, comma delimiter. Keep headers promoted.
```
Tools: `pbi_create_csv_import_query`.

### 6.3 Incremental folder loader (write)
```
Point table "RawLogs" at folder ./logs/ filtered to .csv files only, hidden files excluded. Refresh afterwards.
```
Tools: `pbi_create_folder_import_query`, `pbi_refresh`.

### 6.4 Power Query audit (read-only)
```
Dump every table's current M expression. Redact any literal secret-looking
tokens (already handled by the server). Highlight any expression touching
blocked functions (Web.Contents, Sql.Database, ...) — those will be
rejected by the validator on save.
```
Tools: `pbi_list_power_queries`, `pbi_get_power_query`.

---

## 7. Reporting / visuals

### 7.1 Dashboard from spec (write + requires extracted folder)
```
I have a report extracted at ./extract/. Build a new page "Executive" with:
- 4 KPI cards across the top (Total Sales, Gross Margin, Orders, Customers)
- Bar chart by Region below (Total Sales)
- Line chart by Month for Total Sales MTD vs YTD
- Slicer on Year
Apply the "powerbi-navy-pro" design preset. Patch the layout into demo.pbix
without recompiling.
```
Tools: `pbi_list_pages`, `pbi_create_page`, `pbi_add_visual` (×N),
`pbi_apply_design`, `pbi_patch_layout`.

### 7.2 Theme migration (write)
```
Apply ./themes/corporate.json to the report, then set the background
color of every page to #F0F4FB.
```
Tools: `pbi_apply_theme`, `pbi_apply_design(page_background=...)`.

### 7.3 Bulk layout (write)
```
Use pbi_build_dashboard to materialize this layout spec on page "Overview":
[ {type: card, measure: "CA", x: 20,  y: 20},
  {type: card, measure: "Achats", x: 240, y: 20},
  {type: bar_chart, category_column: "Dim_Geo.Region",
   value_measure: "CA", x: 20, y: 160, width: 600, height: 320} ]
```
Tools: `pbi_build_dashboard`.

### 7.4 Clean slate (destructive, on a disposable pbix)
```
On the extracted folder ./scratch/, delete all pages except the first.
Remove every visual from page 1. Apply the minimal design preset.
Leave the model untouched.
```
Tools: `pbi_list_pages`, `pbi_delete_page`, `pbi_get_page`,
`pbi_remove_visual`, `pbi_apply_design`.

---

## 8. Excel pipeline helpers

### 8.1 Create input template (write)
```
Create ./data/input.xlsx with sheets [Dim_Temps, Dim_Geo, FaitsCA].
Write headers matching the PBI model columns. Format header rows bold
with a navy background. Auto-fit column widths.
```
Tools: `excel_create_workbook`, `excel_write_cell`, `excel_write_range`,
`excel_format_range`, `excel_auto_width`.

### 8.2 Excel <-> PBI drift check (read-only)
```
Compare ./data/input.xlsx against the active model. For every expected
table/column that is missing from the workbook, list it. Do the same
for workbook columns that no PBI table knows about.
```
Tools: `excel_to_pbi_check`.

---

## 9. Performance diagnostics

### 9.1 Slow measure profile (read-only)
```
For each measure in a list, run pbi_trace_query with a simple EVALUATE ROW
wrapper and report duration_ms + SE calls + formula engine time.
Sort descending. Suggest simplifications for the top 5.
```
Tools: `pbi_trace_query`.

### 9.2 DAX dry-run before commit (read-only)
```
Here is a DAX expression I plan to save as a measure. Run pbi_validate_dax
(kind=scalar) and report any error before we touch the model.
```
Tools: `pbi_validate_dax`.

---

## 10. Maintenance / shutdown

### 10.1 Refresh cache after GUI edits (read-only)
```
I just changed the model in the PBI Desktop GUI. Reload the MCP's cached
metadata without a full reconnect.
```
Tools: `pbi_refresh_metadata`.

### 10.2 Pre-release lint (read-only)
```
Before I save this pbix, run the full audit:
- measures without format strings
- hidden tables that are referenced by measures
- relationships with cross-filter "bothDirections" (potential perf risk)
- roles without members
Produce a prioritized action list.
```
Tools: `pbi_list_tables`, `pbi_list_measures`, `pbi_list_relationships`,
`pbi_list_roles`, `pbi_measure_dependencies`.
