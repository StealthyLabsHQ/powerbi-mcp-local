# MCP Power BI Desktop — Local Server

## Goal

Build a Python MCP (Model Context Protocol) server that connects to
Power BI Desktop locally via its embedded Analysis Services instance.

This MCP enables Claude Code (or any MCP client) to:
- Read the data model (tables, columns, relationships, measures)
- Create DAX measures
- Create relationships between tables
- Execute DAX queries
- Automate Power BI report model construction

## Technical Background

When Power BI Desktop is open with a `.pbix` file, it spawns a local
Analysis Services (SSAS) instance on a random port. We can find this port
and connect via the XMLA/ADOMD.NET protocol.

### Finding the port

```
%LOCALAPPDATA%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\
```

Inside this folder there is a subfolder per instance. The file
`msmdsrv.port.txt` contains the port number.

Alternative: find the `msmdsrv.exe` process and its listening port.

## Architecture

```
claude-code (MCP client)
    │
    ▼
mcp-powerbi-server (Python, stdio)
    │
    ▼
Power BI Desktop (local SSAS, dynamic port)
```

## Tech Stack

- **Python 3.11+**
- **`mcp[cli]`** — Anthropic's official MCP SDK (pip install mcp[cli])
- **`pyadomd`** — ADOMD.NET connection to SSAS (pip install pyadomd)
  - Requires .NET Framework (already present on Windows)
  - Alternative: `clr` via `pythonnet` + Microsoft.AnalysisServices.Tabular
- **`psutil`** — To auto-discover the PBI port

## Dependencies

```powershell
pip install "mcp[cli]" pyadomd psutil
```

If `pyadomd` fails to install, use `pythonnet` + TOM instead:

```powershell
pip install pythonnet psutil "mcp[cli]"
```

With pythonnet you also need the Microsoft DLLs:
- `Microsoft.AnalysisServices.Tabular.dll`
- `Microsoft.AnalysisServices.AdomdClient.dll`

These are in the Power BI Desktop install folder:
```
C:\Program Files\Microsoft Power BI Desktop\bin\
```

## Project Structure

```
powerbi-mcp-local/
├── src/
│   ├── server.py           (main MCP server)
│   ├── pbi_connection.py   (SSAS / PBI Desktop connection logic)
│   ├── security.py         (security middleware)
│   └── tools/
│       ├── model.py        (tables, columns, export)
│       ├── measures.py     (DAX measures CRUD)
│       ├── relationships.py (relationships CRUD)
│       ├── query.py        (DAX query execution)
│       ├── power_query.py  (Power Query M tools)
│       ├── excel.py        (Excel read/write)
│       └── visuals.py      (report visuals via pbi-tools)
├── tests/
│   ├── test_connection.py
│   ├── test_security.py
│   ├── test_excel.py
│   ├── test_power_query.py
│   └── test_visuals.py
├── docs/
├── specs/
├── CLAUDE.md           (this file)
├── requirements.txt
└── README.md
```

## MCP Tools to Implement

### 1. `pbi_connect`
Find and connect to the running PBI Desktop instance.
- Scan `%LOCALAPPDATA%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\`
- Read `msmdsrv.port.txt`
- Open ADOMD connection
- Returns: database name, port, status

### 2. `pbi_list_tables`
List all tables in the model with their columns.
- Returns: `[{name, columns: [{name, dataType}], rowCount}]`

### 3. `pbi_list_measures`
List all existing DAX measures.
- Returns: `[{name, table, expression, formatString}]`

### 4. `pbi_list_relationships`
List all relationships in the model.
- Returns: `[{from_table, from_column, to_table, to_column, cardinality, direction}]`

### 5. `pbi_execute_dax(query: str)`
Execute a DAX query and return results.
- Input: DAX query (e.g. `EVALUATE SUMMARIZE(Sales, Dates[Year], "Total", [Total Sales])`)
- Returns: result table (JSON)

### 6. `pbi_create_measure(table: str, name: str, expression: str, format_string?: str)`
Create a new DAX measure in a table.
- Input: target table name, measure name, DAX expression, optional format string
- Returns: confirmation or DAX syntax error

### 7. `pbi_create_relationship(from_table: str, from_column: str, to_table: str, to_column: str, cardinality?: str)`
Create a relationship between two tables.
- cardinality: "oneToMany" (default), "manyToOne", "oneToOne"
- Returns: confirmation

### 8. `pbi_delete_measure(table: str, name: str)`
Delete an existing measure.

### 9. `pbi_model_info`
Return a full model summary (tables, measures, relationships) in one call.

## Implementation — server.py (skeleton)

```python
"""MCP Server — Power BI Desktop local connection."""

import os
import glob
import psutil
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("powerbi-desktop")

# ── Global connection state ──────────────────

_connection = None
_server = None

def find_pbi_port():
    """Find the SSAS port of the running PBI Desktop instance."""
    base = os.path.join(
        os.environ["LOCALAPPDATA"],
        "Microsoft", "Power BI Desktop", "AnalysisServicesWorkspaces"
    )
    if not os.path.exists(base):
        raise FileNotFoundError(f"PBI workspace not found: {base}")

    port_files = glob.glob(os.path.join(base, "*/Data/msmdsrv.port.txt"))
    if not port_files:
        raise FileNotFoundError("No running PBI Desktop instance found")

    # Pick the most recent if multiple instances
    port_file = max(port_files, key=os.path.getmtime)
    with open(port_file) as f:
        return int(f.read().strip())


def get_connection():
    """Return the active ADOMD connection, create if needed."""
    global _connection
    if _connection is not None:
        return _connection

    port = find_pbi_port()
    conn_str = f"Provider=MSOLAP;Data Source=localhost:{port};"

    # Method 1: pyadomd
    try:
        from pyadomd import Pyadomd
        _connection = Pyadomd(conn_str)
        _connection.open()
        return _connection
    except ImportError:
        pass

    # Method 2: pythonnet + ADOMD.NET
    import clr
    clr.AddReference("Microsoft.AnalysisServices.AdomdClient")
    from Microsoft.AnalysisServices.AdomdClient import AdomdConnection
    _connection = AdomdConnection(conn_str)
    _connection.Open()
    return _connection


def get_tom_server():
    """Return a TOM Server object for model manipulation."""
    global _server
    if _server is not None:
        return _server

    port = find_pbi_port()
    import clr
    # Add PBI DLL path if needed
    pbi_bin = r"C:\Program Files\Microsoft Power BI Desktop\bin"
    if os.path.exists(pbi_bin):
        import sys
        sys.path.append(pbi_bin)

    clr.AddReference("Microsoft.AnalysisServices.Tabular")
    from Microsoft.AnalysisServices.Tabular import Server
    _server = Server()
    _server.Connect(f"localhost:{port}")
    return _server


# ── MCP Tools ────────────────────────────────

@mcp.tool()
def pbi_connect() -> str:
    """Find and connect to Power BI Desktop."""
    port = find_pbi_port()
    server = get_tom_server()
    db = server.Databases[0]
    return f"Connected to PBI Desktop on port {port}. Database: {db.Name}. " \
           f"Tables: {db.Model.Tables.Count}."


@mcp.tool()
def pbi_list_tables() -> list:
    """List all tables in the model with columns."""
    server = get_tom_server()
    db = server.Databases[0]
    result = []
    for table in db.Model.Tables:
        cols = [{"name": c.Name, "type": str(c.DataType)}
                for c in table.Columns]
        result.append({"name": table.Name, "columns": cols})
    return result


@mcp.tool()
def pbi_list_measures() -> list:
    """List all DAX measures."""
    server = get_tom_server()
    db = server.Databases[0]
    result = []
    for table in db.Model.Tables:
        for measure in table.Measures:
            result.append({
                "name": measure.Name,
                "table": table.Name,
                "expression": measure.Expression,
                "format": measure.FormatString or ""
            })
    return result


@mcp.tool()
def pbi_create_measure(table: str, name: str, expression: str,
                       format_string: str = "") -> str:
    """Create a DAX measure in the specified table."""
    server = get_tom_server()
    db = server.Databases[0]
    model = db.Model

    target_table = model.Tables.Find(table)
    if target_table is None:
        return f"Error: table '{table}' not found"

    existing = target_table.Measures.Find(name)
    if existing:
        existing.Expression = expression
        if format_string:
            existing.FormatString = format_string
    else:
        from Microsoft.AnalysisServices.Tabular import Measure
        m = Measure()
        m.Name = name
        m.Expression = expression
        if format_string:
            m.FormatString = format_string
        target_table.Measures.Add(m)

    model.SaveChanges()
    return f"Measure '{name}' created/updated in '{table}'"


@mcp.tool()
def pbi_create_relationship(from_table: str, from_column: str,
                            to_table: str, to_column: str) -> str:
    """Create a relationship between two tables."""
    server = get_tom_server()
    db = server.Databases[0]
    model = db.Model

    ft = model.Tables.Find(from_table)
    tt = model.Tables.Find(to_table)
    if not ft or not tt:
        return f"Error: table not found ({from_table} or {to_table})"

    fc = ft.Columns.Find(from_column)
    tc = tt.Columns.Find(to_column)
    if not fc or not tc:
        return f"Error: column not found"

    from Microsoft.AnalysisServices.Tabular import (
        SingleColumnRelationship, RelationshipEndCardinality,
        CrossFilteringBehavior
    )
    rel = SingleColumnRelationship()
    rel.Name = f"{from_table}_{from_column}_{to_table}_{to_column}"
    rel.FromColumn = fc
    rel.ToColumn = tc
    rel.FromCardinality = RelationshipEndCardinality.Many
    rel.ToCardinality = RelationshipEndCardinality.One
    rel.CrossFilteringBehavior = CrossFilteringBehavior.OneDirection

    model.Relationships.Add(rel)
    model.SaveChanges()
    return f"Relationship created: {from_table}[{from_column}] -> {to_table}[{to_column}]"


@mcp.tool()
def pbi_execute_dax(query: str) -> list:
    """Execute a DAX query and return results."""
    conn = get_connection()
    # pyadomd path
    if hasattr(conn, 'cursor'):
        cursor = conn.cursor()
        cursor.execute(query)
        columns = [col.name for col in cursor.description]
        rows = [dict(zip(columns, row)) for row in cursor.fetchall()]
        return rows
    # pythonnet ADOMD path
    else:
        from Microsoft.AnalysisServices.AdomdClient import AdomdCommand
        cmd = AdomdCommand(query, conn)
        reader = cmd.ExecuteReader()
        cols = [reader.GetName(i) for i in range(reader.FieldCount)]
        rows = []
        while reader.Read():
            rows.append({cols[i]: reader.GetValue(i) for i in range(len(cols))})
        reader.Close()
        return rows


@mcp.tool()
def pbi_model_info() -> dict:
    """Full model summary: tables, measures, relationships."""
    tables = pbi_list_tables()
    measures = pbi_list_measures()
    server = get_tom_server()
    db = server.Databases[0]
    rels = []
    for rel in db.Model.Relationships:
        rels.append({
            "from": f"{rel.FromTable.Name}[{rel.FromColumn.Name}]",
            "to": f"{rel.ToTable.Name}[{rel.ToColumn.Name}]"
        })
    return {"tables": tables, "measures": measures, "relationships": rels}


if __name__ == "__main__":
    mcp.run(transport="stdio")
```

## Claude Code Configuration

### Option 1 — Claude Code CLI (project `.claude/settings.json`)

```json
{
  "mcpServers": {
    "powerbi-desktop": {
      "command": "python",
      "args": ["src/server.py"]
    }
  }
}
```

### Option 2 — Claude Desktop App (`%APPDATA%\Claude\claude_desktop_config.json`)

```json
{
  "mcpServers": {
    "powerbi-desktop": {
      "command": "python",
      "args": ["C:\\path\\to\\powerbi-mcp-local\\src\\server.py"],
      "env": {
        "PYTHONPATH": "C:\\Program Files\\Microsoft Power BI Desktop\\bin"
      }
    }
  }
}
```

## Typical Workflow

Once the MCP is built and PBI Desktop is open with a `.pbix`:

```
1. pbi_connect()                              -> verify connection
2. pbi_list_tables()                          -> see imported tables
3. pbi_create_relationship(...)               -> create relationships
4. pbi_create_measure("Sales", "Total", "SUM(Sales[Amount])")
5. pbi_execute_dax("EVALUATE ROW(...)")       -> validate
6. pbi_model_info()                           -> final summary
```

## Known Limitations

1. **Visuals**: the SSAS API only manages the data model (tables, measures,
   relationships). The visual layer (charts, layouts, pages) is not
   accessible via API. Visuals must be created manually in PBI Desktop.

2. **Data import**: the initial Excel import must be done manually in
   PBI Desktop. The API cannot create new data sources.

3. **Theme**: theme JSON import must be done manually
   (View -> Themes -> Browse).

4. **Dynamic port**: the SSAS port changes every time PBI Desktop is opened.
   The MCP server auto-discovers it.

5. **Windows only**: Power BI Desktop only runs on Windows. The MCP server
   must run on the same Windows machine.

## Testing

Before connecting to Claude Code, test standalone:

```powershell
# 1. Make sure PBI Desktop is open with a .pbix file
# 2. Test port detection
python -c "from server import find_pbi_port; print(find_pbi_port())"

# 3. Test MCP server in dev mode
mcp dev src/server.py
```
