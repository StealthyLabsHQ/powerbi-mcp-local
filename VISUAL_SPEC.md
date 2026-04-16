# Power BI Visual Layer — Extension Spec

## Goal

Add tools to create and manage Power BI report visuals programmatically
by manipulating the `.pbix` report layout JSON via pbi-tools.

This completes the full automation pipeline:
```
Excel source -> Power Query import -> Relationships -> DAX measures ->
Refresh -> Report pages & visuals -> Theme
```

## Technical Background

A `.pbix` file is a ZIP containing:
- `DataModel` — the SSAS tabular model (handled by TOM, already implemented)
- `Report/Layout` — a JSON blob defining pages, visuals, filters, styles
- `Report/StaticResources/` — images, custom visuals

**pbi-tools** (https://pbi.tools) can:
1. `pbi-tools extract report.pbix -extractFolder ./report` — decompose into files
2. Modify the JSON files
3. `pbi-tools compile ./report -outPath report.pbix` — recompile

The Layout JSON structure:
```json
{
  "id": 0,
  "reportId": "...",
  "sections": [
    {
      "name": "ReportSection1",
      "displayName": "Page 1",
      "displayOption": 0,
      "width": 1280,
      "height": 720,
      "visualContainers": [
        {
          "x": 50,
          "y": 100,
          "z": 0,
          "width": 400,
          "height": 300,
          "config": "{ ... visual config JSON ... }",
          "filters": "[ ... ]",
          "query": "{ ... }"
        }
      ]
    }
  ]
}
```

Each `visualContainer.config` is a stringified JSON containing:
```json
{
  "name": "unique-visual-id",
  "layouts": [{"id": 0, "position": {"x": 50, "y": 100, "width": 400, "height": 300}}],
  "singleVisual": {
    "visualType": "barChart",
    "projections": {
      "Category": [{"queryRef": "Dim_Temps.Annee"}],
      "Y": [{"queryRef": "CA Total"}]
    },
    "prototypeQuery": { ... },
    "objects": { ... }
  }
}
```

## Architecture

```
Claude Code ──(MCP)──> server.py
                         ├──(TOM/.NET)──────> PBI Desktop (SSAS) [existing]
                         ├──(pbi-tools)─────> .pbix extract/compile
                         └──(JSON manip)────> Report/Layout JSON
```

## Dependencies

- **pbi-tools** — CLI tool, must be on PATH or configured via env var
  - Install: download from https://pbi.tools or `winget install pbi-tools`
  - Or: `dotnet tool install -g pbi-tools`
- No new Python packages needed (just subprocess + json)

## New Tools

### Report Management

#### `pbi_extract_report(pbix_path: str, extract_folder?: str)`
Extract a .pbix into a folder of JSON files using pbi-tools.
- Default extract_folder: `{pbix_name}_extracted/` next to the .pbix
- Returns: `{extract_folder, pages: [...], visual_count}`

#### `pbi_compile_report(extract_folder: str, output_path: str)`
Compile an extracted report folder back into a .pbix.
- Returns: `{output_path, size_bytes}`

#### `pbi_list_pages(extract_folder: str)`
List all report pages with their dimensions and visual counts.
- Returns: `[{name, display_name, width, height, visual_count}]`

#### `pbi_get_page(extract_folder: str, page: str)`
Get full details of a page including all visuals.
- Returns: `{name, display_name, visuals: [{id, type, x, y, width, height, data}]}`

### Page Operations

#### `pbi_create_page(extract_folder: str, display_name: str, width?: int, height?: int)`
Create a new report page.
- Default: 1280x720 (standard 16:9)
- Returns: `{page_name, display_name}`

#### `pbi_delete_page(extract_folder: str, page: str)`
Delete a report page.

#### `pbi_set_page_size(extract_folder: str, page: str, width: int, height: int)`
Change page dimensions.

### Visual Operations

#### `pbi_add_card(extract_folder: str, page: str, measure: str, x: int, y: int, width?: int, height?: int, title?: str)`
Add a KPI card visual displaying a single measure.
- Default size: 200x120
- Returns: `{visual_id}`

#### `pbi_add_bar_chart(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width?: int, height?: int, title?: str, legend_column?: str)`
Add a bar/column chart.
- category_column: "TableName.ColumnName" format
- value_measure: measure name
- Default size: 400x300

#### `pbi_add_line_chart(extract_folder: str, page: str, axis_column: str, value_measures: list[str], x: int, y: int, width?: int, height?: int, title?: str)`
Add a line chart with one or more measures.

#### `pbi_add_donut_chart(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width?: int, height?: int, title?: str)`
Add a donut/pie chart.

#### `pbi_add_table_visual(extract_folder: str, page: str, columns: list[str], x: int, y: int, width?: int, height?: int, title?: str)`
Add a table/matrix visual.
- columns: list of "TableName.ColumnName" or measure names

#### `pbi_add_waterfall(extract_folder: str, page: str, category_column: str, value_measure: str, x: int, y: int, width?: int, height?: int, title?: str)`
Add a waterfall chart (useful for SIG cascade).

#### `pbi_add_slicer(extract_folder: str, page: str, column: str, x: int, y: int, width?: int, height?: int, slicer_type?: str)`
Add a slicer (dropdown, list, or range).
- slicer_type: "dropdown" (default), "list", "range"

#### `pbi_add_text_box(extract_folder: str, page: str, text: str, x: int, y: int, width?: int, height?: int, font_size?: int, bold?: bool, color?: str)`
Add a text label/title on the page.

#### `pbi_remove_visual(extract_folder: str, page: str, visual_id: str)`
Remove a visual from a page.

#### `pbi_move_visual(extract_folder: str, page: str, visual_id: str, x: int, y: int, width?: int, height?: int)`
Reposition or resize a visual.

### Theme

#### `pbi_apply_theme(extract_folder: str, theme_json_path: str)`
Apply a Power BI theme JSON to the extracted report.
- Copies the theme into StaticResources and references it in the layout.

### Bulk Operations

#### `pbi_build_dashboard(extract_folder: str, page: str, layout: list[dict])`
Build an entire dashboard page from a layout spec. Each dict in the list:
```json
{
  "type": "card|bar_chart|line_chart|donut|table|waterfall|slicer|text",
  "x": 50, "y": 100, "width": 400, "height": 300,
  "measure": "CA Total",
  "category": "Dim_Temps.Annee",
  "title": "CA par année"
}
```
This is the power tool — build a full TB in one call.

## Implementation Notes

### pbi-tools Detection
```python
import subprocess
def _find_pbi_tools() -> str:
    """Find pbi-tools executable on PATH or via PBI_TOOLS_PATH env var."""
    custom = os.environ.get("PBI_TOOLS_PATH")
    if custom and Path(custom).exists():
        return custom
    result = subprocess.run(["pbi-tools", "--version"], capture_output=True, text=True)
    if result.returncode == 0:
        return "pbi-tools"
    raise FileNotFoundError("pbi-tools not found. Install from https://pbi.tools")
```

### Layout JSON Parsing
The Layout JSON is deeply nested and uses stringified JSON inside strings.
Key parsing pattern:
```python
import json

def _parse_layout(extract_folder: str) -> dict:
    layout_path = Path(extract_folder) / "Report" / "Layout"
    with open(layout_path, encoding="utf-16-le") as f:
        return json.load(f)

def _get_visual_config(container: dict) -> dict:
    """Parse the stringified config JSON inside a visual container."""
    return json.loads(container.get("config", "{}"))
```

### Visual ID Generation
Each visual needs a unique ID (GUID-like). Use:
```python
import uuid
visual_id = str(uuid.uuid4()).replace("-", "")[:20]
```

### Query Reference Format
Columns: `"TableName.ColumnName"`
Measures: just the measure name (resolved from the model)

### File Encoding
The Layout file uses UTF-16-LE encoding. Always read/write with
`encoding="utf-16-le"`.

## Security Considerations

- `pbi_extract_report` and `pbi_compile_report` execute subprocess calls.
  Validate all paths through the existing `resolve_local_path()` security
  middleware before passing to pbi-tools.
- The extract folder must be inside ALLOWED_BASE_DIRS.
- Never pass unsanitized strings to subprocess — use list form, not shell=True.
- Theme JSON files must be validated before applying.

## Limitations

1. **pbi-tools required** — must be installed separately (not a pip package)
2. **PBI Desktop must be closed** during extract/compile (file lock)
3. **No live preview** — you extract, modify, compile, then reopen in PBI Desktop
4. **Complex visuals** — advanced formatting (conditional formatting, custom
   tooltips, drillthrough) requires deep knowledge of the Layout JSON schema
   which is undocumented by Microsoft
5. **Version sensitivity** — Layout JSON structure can change between PBI Desktop
   versions. The tools should handle unknown fields gracefully (preserve, don't delete)
