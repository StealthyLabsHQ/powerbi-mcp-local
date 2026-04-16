# Excel Tools — Extension Spec

## Goal

Add Excel read/write tools to the existing Power BI MCP server.
This creates a unified data pipeline: Excel (source) -> Power BI (model).

Claude Code can then read source data, modify it, and refresh the PBI model
in a single session without leaving the MCP.

## Architecture

```
Claude Code ──(MCP/stdio)──> server.py
                                ├──(TOM/.NET)──> PBI Desktop (SSAS)
                                └──(openpyxl)──> Excel files (.xlsx)
                                └──(COM/win32)──> Excel Application (optional)
```

Two backends:
1. **openpyxl** — file-based, cross-platform, no Excel needed. Read/write .xlsx directly.
2. **win32com** (optional) — controls the running Excel application on Windows.
   Can trigger recalculations, refresh pivot tables, interact with open workbooks.

openpyxl is the primary backend. win32com is optional and only used when
the user explicitly wants to interact with a running Excel instance.

## New Dependencies

```
openpyxl>=3.1.0
pywin32>=306        # optional, Windows-only, for COM automation
```

## Project Structure (additions)

```
powerbi-mcp-local/
├── ...existing files...
├── tools/
│   ├── ...existing tools...
│   └── excel.py           (all Excel tools)
├── EXCEL_SPEC.md          (this file)
```

## Excel Tools to Implement

### Read Operations

#### `excel_list_sheets(file_path: str)`
List all sheets in a workbook with row/column counts.
- Returns: `[{name, rows, columns, has_data}]`

#### `excel_read_sheet(file_path: str, sheet: str, range?: str, limit?: int)`
Read data from a sheet. Optional range (e.g. "A1:D20") or row limit.
- Returns: `{headers: [...], rows: [[...], ...], total_rows}`
- Default limit: 500 rows (prevent huge payloads)

#### `excel_read_cell(file_path: str, sheet: str, cell: str)`
Read a single cell value with its format and formula (if any).
- Returns: `{value, type, format, formula}`

#### `excel_search(file_path: str, query: str, sheet?: str)`
Search for a value across sheets.
- Returns: `[{sheet, cell, value}]`

### Write Operations

#### `excel_write_cell(file_path: str, sheet: str, cell: str, value: any, format?: str)`
Write a value to a cell with optional number format.
- Returns: confirmation

#### `excel_write_range(file_path: str, sheet: str, start_cell: str, data: list[list])`
Write a 2D array starting at a cell (e.g. "A1").
- Returns: `{rows_written, columns_written}`

#### `excel_create_sheet(file_path: str, name: str, position?: int)`
Create a new sheet in an existing workbook.
- Returns: confirmation

#### `excel_delete_sheet(file_path: str, name: str)`
Delete a sheet from a workbook.
- Returns: confirmation

### Formatting

#### `excel_format_range(file_path: str, sheet: str, range: str, format: dict)`
Apply formatting to a range. Format dict supports:
- `bold`, `italic`, `font_size`, `font_color` (hex)
- `fill_color` (hex), `number_format`
- `border` ("thin", "medium", "thick")
- `alignment` ("left", "center", "right")
- Returns: confirmation

#### `excel_auto_width(file_path: str, sheet: str)`
Auto-fit column widths based on content.
- Returns: confirmation

### Workbook Operations

#### `excel_create_workbook(file_path: str, sheets?: list[str])`
Create a new .xlsx file with optional named sheets.
- Returns: confirmation

#### `excel_workbook_info(file_path: str)`
Full workbook summary: sheets, named ranges, dimensions.
- Returns: `{sheets: [...], named_ranges: [...], properties: {...}}`

### Pipeline Tools (Excel -> PBI)

#### `excel_to_pbi_check(file_path: str)`
Validate an Excel file against the current PBI model:
which sheets match which tables, column name mismatches, data type issues.
- Returns: `{matches: [...], mismatches: [...], suggestions: [...]}`

## Implementation Notes

### File Locking
Excel locks .xlsx files when open. openpyxl can still read (read_only mode)
but cannot write if Excel has the file open. Handle this gracefully:
- Try write -> if PermissionError, return "File locked by Excel, close it first"

### Large Files
Use openpyxl's read_only mode for files > 10MB. Stream rows instead of
loading entire workbook into memory.

### Path Handling
Accept both absolute and relative paths. Resolve relative to the current
working directory. Normalize backslashes on Windows.

### Data Type Mapping
- `int`/`float` -> number
- `str` -> string
- `datetime` -> ISO 8601 string
- `None` -> null
- Formula cells -> return the cached value, expose formula separately

## Error Handling

All tools should return structured errors:
```json
{"error": "FileNotFoundError", "message": "File not found: data.xlsx", "path": "C:\\..."}
```

Never raise exceptions to the MCP client. Always return a JSON error.
