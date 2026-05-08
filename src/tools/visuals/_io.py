"""External-tool I/O: pbi-tools CLI dispatch, PBIX zip extraction,
and PowerBI Desktop graceful-close + force-kill helpers.
"""

from __future__ import annotations

import json
import logging
import os
import shutil
import subprocess
import time
import zipfile
from pathlib import Path
from typing import Any

from pbi_connection import ok

from ._base import (
    LAYOUT_RELATIVE_PATH,
    PBIToolsNotInstalledError,
    ReportLayoutError,
    VisualToolError,
)
from ._layout import _load_layout, _page_summary
from ._paths import _resolve_extract_folder, _resolve_pbix_path

logger = logging.getLogger("tools.visuals._io")


def _run(callback):  # pragma: no cover — thin error-payload wrapper, mirrored by package
    from pbi_connection import error_payload

    try:
        return callback()
    except Exception as exc:
        return error_payload(exc)


def _find_pbi_tools() -> str:
    custom = os.environ.get("PBI_TOOLS_PATH", "").strip()
    if custom:
        candidate = Path(custom).expanduser()
        if candidate.exists():
            return str(candidate)
        raise PBIToolsNotInstalledError(
            "PBI_TOOLS_PATH points to a missing executable.",
            details={"path": str(candidate)},
        )
    discovered = shutil.which("pbi-tools") or shutil.which("pbi-tools.exe") or shutil.which("pbi-tools.core.exe")
    if discovered:
        return discovered
    # __file__ is src/tools/visuals/_io.py — repo root is parents[3].
    bundled = Path(__file__).resolve().parents[3] / "tools-bin" / "pbi-tools.core.exe"
    if bundled.exists():
        return str(bundled)
    fallback_paths = [
        Path.home() / "AppData" / "Local" / "pbi-tools" / "full" / "pbi-tools.exe",
        Path.home() / "AppData" / "Local" / "pbi-tools" / "pbi-tools.core.exe",
    ]
    for fallback in fallback_paths:
        if fallback.exists():
            return str(fallback)
    raise PBIToolsNotInstalledError(
        "pbi-tools was not found on PATH. Install it with winget or dotnet tool install -g pbi-tools."
    )


def _run_pbi_tools(arguments: list[str]) -> dict[str, Any]:
    executable = _find_pbi_tools()
    try:
        completed = subprocess.run(
            [executable, *arguments],
            capture_output=True,
            text=True,
            check=False,
            shell=False,
        )
    except FileNotFoundError as exc:
        raise PBIToolsNotInstalledError("pbi-tools executable could not be launched.") from exc
    if completed.returncode != 0:
        raise VisualToolError(
            "pbi-tools command failed.",
            details={
                "command": [executable, *arguments],
                "returncode": completed.returncode,
                "stdout": completed.stdout[-2000:],
                "stderr": completed.stderr[-2000:],
            },
        )
    return {
        "stdout": completed.stdout,
        "stderr": completed.stderr,
        "returncode": completed.returncode,
    }


def _extract_pbix_zip_natively(pbix: Path, target: Path) -> dict[str, Any]:
    """Fallback PBIX extraction using the standard ZIP. Used when the bundled
    pbi-tools.core does not support 'extract' (it only ships 'compile').

    Copies the Report payload (Layout, StaticResources/Themes) so downstream
    layout-touching tools work. The data model stays inside the PBIX —
    consumers needing model definitions should rely on the live TOM
    connection via pbi_connect.
    """
    target.mkdir(parents=True, exist_ok=True)
    extracted: list[str] = []
    layout_path = target / LAYOUT_RELATIVE_PATH
    layout_path.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(pbix, "r") as zf:
        names = set(zf.namelist())
        if "Report/Layout" in names:
            layout_path.write_bytes(zf.read("Report/Layout"))
            extracted.append("Report/Layout")
        for name in names:
            if name.startswith("Report/StaticResources/") and not name.endswith("/"):
                dest = target / name
                dest.parent.mkdir(parents=True, exist_ok=True)
                dest.write_bytes(zf.read(name))
                extracted.append(name)
    return {"method": "zip_native", "extracted_entries": extracted}


def pbi_extract_report_tool(pbix_path: str, extract_folder: str | None = None) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        pbix = _resolve_pbix_path(pbix_path, must_exist=True)
        target = _resolve_extract_folder(
            str(extract_folder or pbix.with_name(f"{pbix.stem}_extracted")), must_exist=False
        )
        target.mkdir(parents=True, exist_ok=True)
        method = "pbi_tools_extract"
        try:
            _run_pbi_tools(["extract", str(pbix), "-extractFolder", str(target), "-modelSerialization", "Legacy"])
        except (VisualToolError, PBIToolsNotInstalledError) as exc:
            details = getattr(exc, "details", {}) or {}
            stdout = str(details.get("stdout", "")) + str(details.get("stderr", ""))
            cli_lacks_extract = (
                "Unknown action" in stdout
                or "No action was specified" in stdout
                or isinstance(exc, PBIToolsNotInstalledError)
            )
            if not cli_lacks_extract:
                raise
            logger.info("pbi-tools CLI cannot extract (likely the .core build); falling back to ZIP-native extraction.")
            fallback = _extract_pbix_zip_natively(pbix, target)
            method = fallback["method"]
        layout_path = target / LAYOUT_RELATIVE_PATH
        if not layout_path.exists():
            _extract_pbix_zip_natively(pbix, target)
            method = method + "+zip_native_fallback"
        _, layout = _load_layout(target)
        pages = [_page_summary(section) for section in layout.get("sections", [])]
        return ok(
            "Report extracted successfully.",
            pbix_path=str(pbix),
            extract_folder=str(target),
            extraction_method=method,
            pages=pages,
            visual_count=sum(page["visual_count"] for page in pages),
        )

    return _run(_impl)


def _run_powershell(script: str, *, timeout: float = 20.0) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-Command", script],
        capture_output=True,
        text=True,
        check=False,
        shell=False,
        timeout=timeout,
    )


def _save_and_close_powerbi_gracefully(pbix_path: Path | None = None) -> bool:
    target_path = str(pbix_path) if pbix_path is not None else ""
    script = (
        "$TargetPath = "
        + json.dumps(target_path)
        + r"""
$ErrorActionPreference = 'SilentlyContinue'
$ws = New-Object -ComObject WScript.Shell
$names = @('PBIDesktop', 'pbidesktoprs')
$initialWrite = $null
if ($TargetPath -and (Test-Path -LiteralPath $TargetPath)) {
    $initialWrite = (Get-Item -LiteralPath $TargetPath).LastWriteTimeUtc
}
$procs = Get-Process -Name $names | Where-Object { $_.MainWindowHandle -ne 0 }
foreach ($proc in $procs) {
    [void]$ws.AppActivate($proc.Id)
    Start-Sleep -Milliseconds 500
    $ws.SendKeys('^s')
}
if ($initialWrite -ne $null) {
    $deadline = (Get-Date).AddSeconds(30)
    do {
        Start-Sleep -Seconds 1
        $currentWrite = (Get-Item -LiteralPath $TargetPath).LastWriteTimeUtc
    } while ($currentWrite -le $initialWrite -and (Get-Date) -lt $deadline)
} else {
    Start-Sleep -Seconds 8
}
foreach ($proc in @($procs)) {
    $proc.Refresh()
    if (-not $proc.HasExited) {
        [void]$proc.CloseMainWindow()
    }
}
$deadline = (Get-Date).AddSeconds(12)
do {
    Start-Sleep -Milliseconds 500
    $open = @(Get-Process -Name $names | Where-Object { $_.MainWindowHandle -ne 0 }).Count
} while ($open -gt 0 -and (Get-Date) -lt $deadline)
if ($open -gt 0) { exit 1 }
exit 0
"""
    )
    try:
        return _run_powershell(script, timeout=45.0).returncode == 0
    except Exception:
        return False


def _force_kill_powerbi() -> None:
    for image in ("PBIDesktop.exe", "pbidesktoprs.exe"):
        try:
            subprocess.run(
                ["taskkill", "/F", "/IM", image],
                capture_output=True,
                text=True,
                check=False,
                shell=False,
            )
        except Exception:
            pass


def _maybe_force_close_powerbi(force: bool, pbix_path: Path | None = None) -> None:
    if not force:
        return
    if os.name != "nt":
        logger.debug("force=True ignored on non-Windows platform for PBIDesktop termination.")
        return
    # Resolve through the package re-export so test patches against
    # ``tools.visuals._save_and_close_powerbi_gracefully`` /
    # ``tools.visuals._force_kill_powerbi`` keep working.
    from . import _force_kill_powerbi as _kill
    from . import _save_and_close_powerbi_gracefully as _graceful

    if not _graceful(pbix_path):
        _kill()
    time.sleep(1.5)


def _page_names_from_layout_bytes(layout_bytes: bytes) -> list[str]:
    try:
        layout = json.loads(layout_bytes.decode("utf-16-le"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ReportLayoutError("Report/Layout content is invalid UTF-16-LE JSON.") from exc
    if not isinstance(layout, dict):
        raise ReportLayoutError("Report/Layout root must be a JSON object.")
    names: list[str] = []
    for section in layout.get("sections", []):
        if not isinstance(section, dict):
            continue
        names.append(str(section.get("displayName") or section.get("name") or ""))
    return names


def pbi_compile_report_tool(extract_folder: str, output_path: str, force: bool = False) -> dict[str, Any]:
    def _impl() -> dict[str, Any]:
        folder = _resolve_extract_folder(extract_folder, must_exist=True)
        output = _resolve_pbix_path(output_path, must_exist=False)
        output.parent.mkdir(parents=True, exist_ok=True)
        _maybe_force_close_powerbi(force, output if output.exists() else None)
        _run_pbi_tools(["compile", str(folder), "-outPath", str(output), "-overwrite"])
        return ok(
            "Report compiled successfully.",
            extract_folder=str(folder),
            output_path=str(output),
            size_bytes=output.stat().st_size if output.exists() else None,
        )

    return _run(_impl)
