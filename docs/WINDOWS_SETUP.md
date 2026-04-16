# Windows Setup — Step by Step

Everything you need to make the MCP server operational on your Windows machine.

---

## 1. Prerequisites

### Python 3.11+

```powershell
# Check if already installed
python --version

# If not, install via winget
winget install Python.Python.3.11
```

Restart your terminal after installing.

### Power BI Desktop (free)

```powershell
# Option A — Microsoft Store (recommended, auto-updates)
# Open Microsoft Store > search "Power BI Desktop" > Install

# Option B — MSI installer
winget install Microsoft.PowerBIDesktop
```

### Claude Code

```powershell
npm install -g @anthropic-ai/claude-code
```

### pbi-tools (required for visual tools only)

```powershell
# Option A — winget
winget install pbi-tools

# Option B — .NET tool
dotnet tool install -g pbi-tools

# Option C — manual download
# https://pbi.tools/downloads/ > extract to a folder on your PATH
```

Verify: `pbi-tools --version`

---

## 2. Clone and Install

```powershell
cd C:\Projects
git clone https://github.com/StealthyLabsHQ/powerbi-mcp-local.git
cd powerbi-mcp-local

pip install -r requirements.txt
```

---

## 3. Test Connection

Open Power BI Desktop with any `.pbix` file (or create a blank one), then:

```powershell
python test_connection.py
```

Expected output:
```
Connected to PBI Desktop on port XXXXX
Database: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
Tables: N
```

If it fails:
- Make sure PBI Desktop is **open** with a file
- Check if it's an MSI or Store install (see Troubleshooting below)

---

## 4. Configure Claude Code

Create the file `.claude/settings.json` in the project folder:

```powershell
mkdir .claude
```

Write this content into `.claude\settings.json`:

```json
{
  "mcpServers": {
    "powerbi": {
      "command": "python",
      "args": ["C:\\Projects\\powerbi-mcp-local\\server.py"]
    }
  }
}
```

> Replace `C:\Projects\powerbi-mcp-local` with your actual path.

---

## 5. Verify MCP Connection

```powershell
cd C:\Projects\powerbi-mcp-local
claude
```

Then type:

```
Call pbi_connect to verify the Power BI connection
```

You should see: port number, database name, table count.

Then try:

```
Call pbi_list_tables to see all tables in the model
```

---

## 6. ALTI & BIKE Workflow

### Step 1 — Import Excel data

Open PBI Desktop > blank file > save as `alti_bike.pbix`.

Import the data manually (one-time):
- Home > Get Data > Excel
- Select `alti_bike_data.xlsx` from the alti-bike-sae repo
- Check all 14 sheets (skip `_KPI`) > Load

Or use the MCP (after import):
```
Use pbi_bulk_import_excel to point all tables to alti_bike_data.xlsx, then pbi_refresh
```

### Step 2 — Create relationships

Tell Claude Code:
```
Create these relationships:
- FaitsCA[Annee] -> Dim_Temps[Annee]
- FaitsCA[Famille] -> Dim_Famille[Famille]
- FaitsVar[Annee] -> Dim_Temps[Annee]
- FaitsDept[Annee] -> Dim_Temps[Annee]
- CompteResultat[Annee] -> Dim_Temps[Annee]
- Bilan[Annee] -> Dim_Temps[Annee]
```

### Step 3 — Import DAX measures

Tell Claude Code:
```
Import all measures from C:\path\to\alti-bike-sae\powerbi\mesures_DAX.dax into table FaitsCA
```

### Step 4 — Validate

```
Execute this DAX query: EVALUATE ROW("CA 2025", [CA Total])
```

Should return 1,765,000.

### Step 5 — Apply theme

```
Apply the theme from C:\path\to\alti-bike-sae\powerbi\theme\alti_bike_theme.json
```

### Step 6 — Build the 4 dashboards

Option A — manual (follow `guide_visuels.md`):
Create each visual by hand using the guide.

Option B — via MCP (requires pbi-tools):
```
Close PBI Desktop, then:
1. Extract the .pbix: pbi_extract_report("alti_bike.pbix")
2. Create 4 pages: TB1 Commercial, TB2 Performance, TB3 Financement, TB4 RSE
3. Use pbi_build_dashboard for each page with the visual layout
4. Apply theme
5. Compile back: pbi_compile_report
6. Reopen in PBI Desktop to verify
```

### Step 7 — Save and push

Save the `.pbix`, take screenshots for the rapport, push to GitHub.

---

## Troubleshooting

### "No running PBI Desktop instance found"

PBI Desktop must be **open** with a `.pbix` file. A blank window is not enough — you need to have data loaded or at least a saved empty file.

### Wrong DLL path (TOM connection fails)

The server auto-detects the DLL path. If it fails:

```powershell
# Find where PBI Desktop is installed
where /R "C:\Program Files" msmdsrv.exe
where /R "%LOCALAPPDATA%\Microsoft\Power BI Desktop" msmdsrv.exe
```

Then set the env var:
```powershell
set PYTHONPATH=C:\Program Files\Microsoft Power BI Desktop\bin
```

### Microsoft Store vs MSI install

| | MSI | Store |
|---|---|---|
| Path | `C:\Program Files\Microsoft Power BI Desktop\` | `%LOCALAPPDATA%\Microsoft\Power BI Desktop\` |
| Workspace | `%LOCALAPPDATA%\Microsoft\Power BI Desktop\AnalysisServicesWorkspaces\` | User-scoped subfolder |
| Detection | Auto | Auto (fallback to process scan) |

The server handles both. If neither works, it falls back to scanning `msmdsrv.exe` processes.

### pbi-tools not found (visual tools only)

```powershell
pbi-tools --version
```

If not found:
```powershell
# Add to PATH manually
set PATH=%PATH%;C:\path\to\pbi-tools

# Or set env var
set PBI_TOOLS_PATH=C:\path\to\pbi-tools.exe
```

### File locked by Excel

If you get "PermissionError" on Excel tools, close Excel first. PBI Desktop does not lock `.xlsx` files, but Excel does.

### Security: path blocked

By default, Excel tools only access files in the current directory. To allow other directories:

```powershell
set PBI_MCP_ALLOWED_DIRS=C:\Projects\alti-bike-sae\data;C:\Projects\powerbi-mcp-local
```

---

## Environment Variables Reference

| Variable | Default | Description |
|---|---|---|
| `PBI_MCP_ALLOWED_DIRS` | (cwd) | Allowed directories for file access |
| `PBI_MCP_ALLOW_DMV` | `0` | Allow DMV/system queries |
| `PBI_MCP_ALLOW_EXTERNAL_M` | `0` | Allow network M functions |
| `PBI_MCP_READONLY` | `0` | Block all write operations |
| `PBI_MCP_LOG_LEVEL` | `INFO` | Logging level |
| `PBI_TOOLS_PATH` | (auto) | Path to pbi-tools executable |
| `PYTHONPATH` | (auto) | Path to PBI Desktop bin folder |
