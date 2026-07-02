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
    _run,
)
from ._layout import _load_layout, _page_summary
from ._paths import _resolve_extract_folder, _resolve_pbix_path

logger = logging.getLogger("tools.visuals._io")


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


_PBI_TOOLS_TIMEOUT_SECONDS = int(os.environ.get("PBI_MCP_PBI_TOOLS_TIMEOUT", "300"))


def _run_pbi_tools(arguments: list[str]) -> dict[str, Any]:
    executable = _find_pbi_tools()
    try:
        completed = subprocess.run(
            [executable, *arguments],
            capture_output=True,
            text=True,
            check=False,
            shell=False,
            timeout=_PBI_TOOLS_TIMEOUT_SECONDS,
        )
    except subprocess.TimeoutExpired as exc:
        raise VisualToolError(
            f"pbi-tools command timed out after {_PBI_TOOLS_TIMEOUT_SECONDS}s.",
            details={"command": [executable, *arguments], "timeout_seconds": _PBI_TOOLS_TIMEOUT_SECONDS},
        ) from exc
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


def _inspect_pbix_archive(pbix: Path) -> None:
    """Reject PBIX archives that look like zip bombs.

    The active SecurityPolicy ships caps for Excel workbooks
    (``max_excel_zip_*``); PBIX is also a ZIP container, so we reuse the
    same caps here. A hostile PBIX with 100k members or a 1000:1
    compression ratio could exhaust disk / memory during native
    extraction; this preflight raises ``SecurityPolicyError`` before any
    member is written.
    """
    from security import SECURITY, SecurityPolicyError

    policy = SECURITY.policy()
    if not zipfile.is_zipfile(pbix):
        raise SecurityPolicyError("PBIX is not a valid ZIP archive.", details={"path": str(pbix)})
    member_limit = policy.max_excel_zip_members
    uncompressed_limit = policy.max_excel_zip_uncompressed_bytes
    ratio_limit = policy.max_excel_zip_compression_ratio
    with zipfile.ZipFile(pbix) as archive:
        infos = archive.infolist()
        if len(infos) > member_limit:
            raise SecurityPolicyError(
                "PBIX exceeds the maximum number of ZIP members.",
                details={"members": len(infos), "limit": member_limit, "path": str(pbix)},
            )
        total_uncompressed = 0
        for info in infos:
            total_uncompressed += int(info.file_size)
            if total_uncompressed > uncompressed_limit:
                raise SecurityPolicyError(
                    "PBIX exceeds the maximum decompressed size.",
                    details={
                        "bytes": total_uncompressed,
                        "limit": uncompressed_limit,
                        "path": str(pbix),
                    },
                )
            compressed = max(int(info.compress_size), 1)
            if info.file_size and (info.file_size / compressed) > ratio_limit:
                raise SecurityPolicyError(
                    "PBIX looks like a ZIP bomb.",
                    details={
                        "member": info.filename,
                        "ratio": round(info.file_size / compressed, 2),
                        "limit": ratio_limit,
                        "path": str(pbix),
                    },
                )


def _extract_pbix_zip_natively(pbix: Path, target: Path) -> dict[str, Any]:
    """Fallback PBIX extraction using the standard ZIP. Used when the bundled
    pbi-tools.core does not support 'extract' (it only ships 'compile').

    Copies the Report payload (Layout, StaticResources/Themes) so downstream
    layout-touching tools work. The data model stays inside the PBIX —
    consumers needing model definitions should rely on the live TOM
    connection via pbi_connect.
    """
    _inspect_pbix_archive(pbix)
    target.mkdir(parents=True, exist_ok=True)
    target_resolved = target.resolve()
    extracted: list[str] = []
    layout_path = target / LAYOUT_RELATIVE_PATH
    layout_path.parent.mkdir(parents=True, exist_ok=True)

    def _is_safe_member(name: str) -> bool:
        # Reject absolute paths, drive letters, and any traversal component
        # before joining with target. Power BI member names always use
        # forward slashes, so split on both separators defensively.
        if not name or name.endswith("/") or name.endswith("\\"):
            return False
        if name.startswith("/") or name.startswith("\\") or (len(name) >= 2 and name[1] == ":"):
            return False
        parts = [p for p in name.replace("\\", "/").split("/") if p]
        if any(p == ".." or p == "." for p in parts):
            return False
        return True

    def _safe_dest(name: str) -> Path | None:
        if not _is_safe_member(name):
            return None
        dest = (target / name).resolve()
        try:
            dest.relative_to(target_resolved)
        except ValueError:
            return None
        return dest

    skipped_traversal = 0
    with zipfile.ZipFile(pbix, "r") as zf:
        names = set(zf.namelist())
        if "Report/Layout" in names:
            layout_path.write_bytes(zf.read("Report/Layout"))
            extracted.append("Report/Layout")
        for name in names:
            if not name.startswith("Report/StaticResources/"):
                continue
            dest = _safe_dest(name)
            if dest is None:
                # Path-traversal attempt (zip-slip) — skip without surfacing
                # the malicious member name in extracted[]. Count for one
                # aggregate warning so the operator can investigate.
                skipped_traversal += 1
                continue
            dest.parent.mkdir(parents=True, exist_ok=True)
            dest.write_bytes(zf.read(name))
            extracted.append(name)
    if skipped_traversal:
        logger.warning(
            "PBIX %s contained %d zip-slip member(s); they were skipped during native extraction.",
            pbix.name,
            skipped_traversal,
        )
    return {
        "method": "zip_native",
        "extracted_entries": extracted,
        "skipped_traversal_count": skipped_traversal,
    }


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


# Force UTF-8 I/O on Windows PowerShell 5.1 (default is UTF-16 LE for Out-File
# and the locale codepage for stdout). Without this, paths or output text
# containing non-ASCII characters can be silently mangled when round-tripped
# through json.dumps + ConvertFrom-Json on the Python side.
_PS_UTF8_PRELUDE = (
    "$OutputEncoding = [System.Text.UTF8Encoding]::new($false);"
    "[Console]::OutputEncoding = [System.Text.UTF8Encoding]::new($false);"
    "[Console]::InputEncoding = [System.Text.UTF8Encoding]::new($false);"
    "$PSDefaultParameterValues['Out-File:Encoding']='utf8';"
    "$PSDefaultParameterValues['*:Encoding']='utf8';\n"
)


def _run_powershell(script: str, *, timeout: float = 20.0) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        [
            "powershell",
            "-NoProfile",
            "-NonInteractive",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            _PS_UTF8_PRELUDE + script,
        ],
        capture_output=True,
        text=True,
        check=False,
        shell=False,
        timeout=timeout,
    )


def _post_ctrl_s_to_pbi_processes() -> int:
    """Send Ctrl+S to every running PBI Desktop window via PostMessage.

    Returns the number of windows the chord was posted to. Uses
    PostMessage instead of WScript.Shell SendKeys + AppActivate (the
    legacy approach) so the chord stays bound to PBI Desktop's HWND
    even if focus moves mid-call. Same rationale as
    ``ui_automation._send_ctrl_s``: no global-input-queue routing.
    """
    if os.name != "nt":
        return 0
    try:
        import ctypes
        from ctypes import wintypes

        import psutil
    except ImportError:
        return 0

    pbi_pids = {
        int(proc.info["pid"])
        for proc in psutil.process_iter(attrs=["pid", "name"])
        if (proc.info.get("name") or "").lower() in {"pbidesktop.exe", "pbidesktoprs.exe"}
    }
    if not pbi_pids:
        return 0

    user32 = ctypes.windll.user32
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    user32.IsWindowVisible.restype = wintypes.BOOL
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.PostMessageW.restype = wintypes.BOOL

    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)
    targets: list[int] = []

    def _cb(hwnd: int, _lparam: int) -> bool:
        proc_pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_pid))
        if int(proc_pid.value) not in pbi_pids:
            return True
        if not user32.IsWindowVisible(hwnd) or user32.GetWindowTextLengthW(hwnd) <= 0:
            return True
        targets.append(int(hwnd))
        return True

    user32.EnumWindows(EnumWindowsProc(_cb), 0)

    WM_KEYDOWN = 0x0100
    WM_KEYUP = 0x0101
    VK_CONTROL = 0x11
    VK_S = 0x53
    posted = 0
    for hwnd in targets:
        ok_ = all(
            user32.PostMessageW(hwnd, msg, vk, 0)
            for msg, vk in (
                (WM_KEYDOWN, VK_CONTROL),
                (WM_KEYDOWN, VK_S),
                (WM_KEYUP, VK_S),
                (WM_KEYUP, VK_CONTROL),
            )
        )
        if ok_:
            posted += 1
    return posted


def _save_and_close_powerbi_gracefully(pbix_path: Path | None = None) -> bool:
    # Python-side keyboard injection via PostMessage avoids the focus race
    # in the previous WScript.Shell SendKeys path. The PowerShell helper
    # below now only handles the wait-for-mtime + CloseMainWindow phase.
    posted = _post_ctrl_s_to_pbi_processes()
    if posted == 0:
        return True  # nothing to save; caller treats as success
    target_path = str(pbix_path) if pbix_path is not None else ""
    script = (
        "$TargetPath = "
        + json.dumps(target_path)
        + r"""
$ErrorActionPreference = 'SilentlyContinue'
$names = @('PBIDesktop', 'pbidesktoprs')
$initialWrite = $null
if ($TargetPath -and (Test-Path -LiteralPath $TargetPath)) {
    $initialWrite = (Get-Item -LiteralPath $TargetPath).LastWriteTimeUtc
}
$procs = Get-Process -Name $names | Where-Object { $_.MainWindowHandle -ne 0 }
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


def attempt_pbi_save_before_close(pbix_path: Path | None, timeout_seconds: float = 10.0) -> dict[str, Any]:
    """Best-effort Ctrl+S to every running PBI Desktop window, then wait
    for ``pbix_path`` mtime to change.

    Used by :func:`pbi_patch_layout_tool` when ``save_before_close=True``
    so in-memory TOM mutations (measures, columns, role filters) get
    flushed to the PBIX before ``_maybe_force_close_powerbi`` kills the
    Desktop process. Always returns a status payload — never raises —
    because the patch-layout flow must continue regardless.
    """
    info: dict[str, Any] = {
        "attempted": False,
        "windows_targeted": 0,
        "mtime_changed": False,
        "polled_seconds": 0.0,
        "skipped_reason": None,
    }
    if os.name != "nt":
        info["skipped_reason"] = "non_windows_platform"
        return info
    posted = _post_ctrl_s_to_pbi_processes()
    info["windows_targeted"] = posted
    info["attempted"] = posted > 0
    if posted == 0:
        info["skipped_reason"] = "no_pbi_desktop_window"
        return info
    if pbix_path is None or not pbix_path.exists():
        info["skipped_reason"] = "pbix_path_missing"
        return info
    deadline = time.monotonic() + max(1.0, float(timeout_seconds))
    initial_mtime = pbix_path.stat().st_mtime
    while time.monotonic() < deadline:
        time.sleep(0.25)
        try:
            current_mtime = pbix_path.stat().st_mtime
        except OSError:
            continue
        if current_mtime > initial_mtime:
            info["mtime_changed"] = True
            break
    info["polled_seconds"] = round(timeout_seconds - max(0.0, deadline - time.monotonic()), 3)
    return info


def _maybe_force_close_powerbi(force: bool, pbix_path: Path | None = None, *, save_verified: bool = False) -> None:
    """Close (and if safe, kill) Power BI Desktop when ``force=True``.

    ``taskkill /F`` discards every TOM mutation that has not been flushed
    to the PBIX (writes commit to the in-memory AS engine only), so the
    kill is refused unless either the graceful save-and-close succeeded
    or the caller verified a save beforehand (``save_verified=True``,
    e.g. via ``attempt_pbi_save_before_close`` observing an mtime change).
    """
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
        if not save_verified:
            raise ReportLayoutError(
                "Refusing to force-kill Power BI Desktop: the graceful save-and-close "
                "failed and no prior save was verified, so unsaved model edits would be "
                "lost. Save manually in Power BI Desktop (Ctrl+S) and retry, or rerun "
                "with save_before_close=True and confirm save_attempt.mtime_changed.",
                details={"pbix_path": str(pbix_path) if pbix_path else None},
            )
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
