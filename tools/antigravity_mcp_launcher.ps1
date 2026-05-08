param(
    [Parameter(ValueFromRemainingArguments = $true)]
    [string[]]$ServerArgs
)

$ErrorActionPreference = "Stop"

$RepoRoot = Split-Path -Parent $PSScriptRoot
$Python = Join-Path $RepoRoot ".venv\Scripts\python.exe"
$Server = Join-Path $RepoRoot "src\server.py"

if (-not (Test-Path -LiteralPath $Python)) {
    Write-Error "Python venv not found: $Python"
    exit 1
}

if (-not (Test-Path -LiteralPath $Server)) {
    Write-Error "MCP server entrypoint not found: $Server"
    exit 1
}

$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"

Set-Location -LiteralPath $RepoRoot
& $Python $Server @ServerArgs
exit $LASTEXITCODE
