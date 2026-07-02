"""PBIX persistence validation and Power BI Desktop reopen probe machinery."""

from __future__ import annotations

import json
import os
import subprocess
import zipfile
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, ok
from security import resolve_local_path

from ._shared import _PS_UTF8_PRELUDE, _load_layout


def _layout_summary(layout: dict[str, Any]) -> dict[str, Any]:
    pages: list[dict[str, Any]] = []
    visual_count = 0
    for section in layout.get("sections", []) or []:
        if not isinstance(section, dict):
            continue
        visuals = [item for item in section.get("visualContainers", []) or [] if isinstance(item, dict)]
        visual_count += len(visuals)
        pages.append(
            {
                "name": section.get("name"),
                "display_name": section.get("displayName") or section.get("name"),
                "visual_count": len(visuals),
            }
        )
    return {"page_count": len(pages), "visual_count": visual_count, "pages": pages}


def pbi_validate_pbix_persistence_tool(
    *,
    pbix_path: str,
    extract_folder: str | None = None,
    require_security_bindings_removed: bool = True,
) -> dict[str, Any]:
    """Validate that a patched PBIX still contains a readable, persistent report layout."""
    pbix = resolve_local_path(pbix_path, must_exist=True, allowed_extensions={".pbix"})
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    pbix_summary: dict[str, Any] | None = None
    extract_summary: dict[str, Any] | None = None

    if not zipfile.is_zipfile(pbix):
        issues.append({"type": "pbix_not_zip", "pbix_path": str(pbix)})
    else:
        with zipfile.ZipFile(pbix, "r") as archive:
            names = set(archive.namelist())
            if "Report/Layout" not in names:
                issues.append({"type": "missing_report_layout", "pbix_path": str(pbix)})
            else:
                try:
                    layout = json.loads(archive.read("Report/Layout").decode("utf-16-le"))
                    pbix_summary = _layout_summary(layout)
                    if pbix_summary["page_count"] == 0:
                        warnings.append({"type": "no_report_pages", "pbix_path": str(pbix)})
                    if pbix_summary["visual_count"] == 0:
                        warnings.append({"type": "no_report_visuals", "pbix_path": str(pbix)})
                except (UnicodeDecodeError, json.JSONDecodeError) as exc:
                    issues.append({"type": "invalid_report_layout", "pbix_path": str(pbix), "error": str(exc)})
            if require_security_bindings_removed and "SecurityBindings" in names:
                issues.append({"type": "security_bindings_present", "pbix_path": str(pbix)})

    if extract_folder:
        _, extract_layout = _load_layout(extract_folder)
        extract_summary = _layout_summary(extract_layout)
        if pbix_summary and extract_summary:
            if pbix_summary["page_count"] != extract_summary["page_count"]:
                issues.append(
                    {
                        "type": "page_count_mismatch",
                        "pbix_count": pbix_summary["page_count"],
                        "extract_count": extract_summary["page_count"],
                    }
                )
            if pbix_summary["visual_count"] != extract_summary["visual_count"]:
                issues.append(
                    {
                        "type": "visual_count_mismatch",
                        "pbix_count": pbix_summary["visual_count"],
                        "extract_count": extract_summary["visual_count"],
                    }
                )

    return ok(
        f"PBIX persistence validation found {len(issues)} issue(s), {len(warnings)} warning(s).",
        pbix_path=str(pbix),
        extract_folder=extract_folder,
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        pbix_summary=pbix_summary,
        extract_summary=extract_summary,
    )


def _run_reopen_probe(
    *,
    pbix_path: Path,
    timeout_seconds: int,
    screenshot_path: Path | None,
    close_after: bool,
) -> dict[str, Any]:
    script = (
        "$PbixPath = "
        + json.dumps(str(pbix_path))
        + "\n$TimeoutSeconds = "
        + str(timeout_seconds)
        + "\n$ScreenshotPath = "
        + json.dumps(str(screenshot_path) if screenshot_path else "")
        + "\n$CloseAfter = $"
        + ("true" if close_after else "false")
        + r"""
$ErrorActionPreference = 'SilentlyContinue'
$before = @(Get-Process -Name PBIDesktop,pbidesktoprs | Select-Object -ExpandProperty Id)
Start-Process -FilePath $PbixPath | Out-Null
$deadline = (Get-Date).AddSeconds($TimeoutSeconds)
$proc = $null
do {
    Start-Sleep -Seconds 1
    $candidates = @(Get-Process -Name PBIDesktop,pbidesktoprs | Where-Object { $_.MainWindowHandle -ne 0 })
    $proc = $candidates | Where-Object { $before -notcontains $_.Id } | Select-Object -First 1
    if ($proc -eq $null) { $proc = $candidates | Select-Object -First 1 }
} while ($proc -eq $null -and (Get-Date) -lt $deadline)

$texts = @()
$screenshotOk = $false
if ($proc -ne $null) {
    Add-Type -AssemblyName UIAutomationClient | Out-Null
    $root = [System.Windows.Automation.AutomationElement]::RootElement
    $condition = New-Object System.Windows.Automation.PropertyCondition([System.Windows.Automation.AutomationElement]::ProcessIdProperty, $proc.Id)
    $window = $root.FindFirst([System.Windows.Automation.TreeScope]::Children, $condition)
    if ($window -ne $null) {
        $nodes = $window.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
        foreach ($node in $nodes) {
            $name = $node.Current.Name
            if ($name) { $texts += $name }
        }
    }
    if ($ScreenshotPath) {
        Add-Type -AssemblyName System.Windows.Forms | Out-Null
        Add-Type -AssemblyName System.Drawing | Out-Null
        $bounds = [System.Windows.Forms.Screen]::PrimaryScreen.Bounds
        $bitmap = New-Object System.Drawing.Bitmap $bounds.Width, $bounds.Height
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($bounds.Location, [System.Drawing.Point]::Empty, $bounds.Size)
        $bitmap.Save($ScreenshotPath, [System.Drawing.Imaging.ImageFormat]::Png)
        $graphics.Dispose()
        $bitmap.Dispose()
        $screenshotOk = Test-Path -LiteralPath $ScreenshotPath
    }
    if ($CloseAfter) {
        [void]$proc.CloseMainWindow()
    }
}

$uniqueTexts = @($texts | Select-Object -Unique | Select-Object -First 200)
$signals = @(
    'Fix this',
    "Something's wrong with one or more fields",
    'See details',
    'Something went wrong',
    'Database consistency checks',
    'DBCC',
    'Vertipaq',
    'string store',
    'An error occurred while loading',
    'Report this issue',
    'Copy details to clipboard',
    'multiple tables'
)
$matches = @()
foreach ($signal in $signals) {
    if (($uniqueTexts -join "`n") -like "*$signal*") { $matches += $signal }
}
[PSCustomObject]@{
    opened = ($proc -ne $null)
    process_id = if ($proc -ne $null) { $proc.Id } else { $null }
    process_name = if ($proc -ne $null) { $proc.ProcessName } else { $null }
    window_title = if ($proc -ne $null) { $proc.MainWindowTitle } else { $null }
    ui_text_count = $uniqueTexts.Count
    ui_text_matches = $matches
    screenshot_path = if ($screenshotOk) { $ScreenshotPath } else { $null }
} | ConvertTo-Json -Depth 4
"""
    )
    result = subprocess.run(
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
        timeout=timeout_seconds + 20,
    )
    if result.returncode != 0:
        raise PowerBIValidationError(
            "Power BI Desktop reopen probe failed.",
            details={"returncode": result.returncode, "stderr": result.stderr.strip(), "stdout": result.stdout.strip()},
        )
    return json.loads(result.stdout)


def _analyze_reopen_screenshot(screenshot_path: Path) -> dict[str, Any]:
    script = (
        "$ScreenshotPath = "
        + json.dumps(str(screenshot_path))
        + r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Drawing | Out-Null
$bitmap = [System.Drawing.Bitmap]::FromFile($ScreenshotPath)
$width = $bitmap.Width
$height = $bitmap.Height
$maxSamples = 80000
$step = [Math]::Max(1, [int][Math]::Ceiling([Math]::Sqrt(($width * $height) / $maxSamples)))
$samples = 0
$dark = 0
$teal = 0
for ($y = 0; $y -lt $height; $y += $step) {
    for ($x = 0; $x -lt $width; $x += $step) {
        $pixel = $bitmap.GetPixel($x, $y)
        $samples += 1
        if ($pixel.R -lt 35 -and $pixel.G -lt 35 -and $pixel.B -lt 35) { $dark += 1 }
        if ($pixel.R -lt 45 -and $pixel.G -ge 85 -and $pixel.G -le 170 -and $pixel.B -ge 80 -and $pixel.B -le 170) { $teal += 1 }
    }
}
$bitmap.Dispose()
$darkRatio = if ($samples -gt 0) { $dark / $samples } else { 0 }
$tealRatio = if ($samples -gt 0) { $teal / $samples } else { 0 }
[PSCustomObject]@{
    width = $width
    height = $height
    sample_count = $samples
    dark_pixel_ratio = [Math]::Round($darkRatio, 4)
    teal_pixel_ratio = [Math]::Round($tealRatio, 4)
    fix_this_like = ($darkRatio -ge 0.32 -and $tealRatio -ge 0.0005)
} | ConvertTo-Json -Depth 3
"""
    )
    result = subprocess.run(
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
        timeout=30,
    )
    if result.returncode != 0:
        error = result.stderr.strip() or result.stdout.strip()
        return {"available": False, "error": error.splitlines()[0] if error else "Windows OCR failed."}
    payload = json.loads(result.stdout)
    payload["available"] = True
    return payload


def _ocr_reopen_screenshot(screenshot_path: Path) -> dict[str, Any]:
    script = (
        "$ScreenshotPath = "
        + json.dumps(str(screenshot_path))
        + r"""
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Runtime.WindowsRuntime | Out-Null
$null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
$null = [Windows.Graphics.Imaging.BitmapDecoder, Windows.Graphics.Imaging, ContentType = WindowsRuntime]
$null = [Windows.Media.Ocr.OcrEngine, Windows.Media.Ocr, ContentType = WindowsRuntime]
$asTaskGeneric = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.IsGenericMethod
} | Select-Object -First 1)
function Await-WinRt($operation, $type) {
    $task = $asTaskGeneric.MakeGenericMethod($type).Invoke($null, @($operation))
    try {
        $task.Wait()
    } catch {
        $message = $_.Exception.Message
        if ($_.Exception.InnerException -ne $null) {
            $message = "$message | $($_.Exception.InnerException.Message)"
        }
        throw $message
    }
    return $task.Result
}
$file = Await-WinRt ([Windows.Storage.StorageFile]::GetFileFromPathAsync($ScreenshotPath)) ([Windows.Storage.StorageFile])
$stream = Await-WinRt ($file.OpenReadAsync()) ([Windows.Storage.Streams.IRandomAccessStreamWithContentType])
$decoder = Await-WinRt ([Windows.Graphics.Imaging.BitmapDecoder]::CreateAsync($stream)) ([Windows.Graphics.Imaging.BitmapDecoder])
$bitmap = Await-WinRt ($decoder.GetSoftwareBitmapAsync()) ([Windows.Graphics.Imaging.SoftwareBitmap])
$engine = [Windows.Media.Ocr.OcrEngine]::TryCreateFromUserProfileLanguages()
if ($engine -eq $null) { throw 'Windows OCR engine unavailable' }
$result = Await-WinRt ($engine.RecognizeAsync($bitmap)) ([Windows.Media.Ocr.OcrResult])
$text = $result.Text
$signals = @('Fix this', "Something's wrong with one or more fields", 'See details')
$matches = @()
foreach ($signal in $signals) {
    if ($text -like "*$signal*") { $matches += $signal }
}
[PSCustomObject]@{
    text_length = if ($text) { $text.Length } else { 0 }
    matches = $matches
} | ConvertTo-Json -Depth 3
"""
    )
    result = subprocess.run(
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
        timeout=45,
    )
    if result.returncode != 0:
        error = result.stderr.strip() or result.stdout.strip()
        return {"available": False, "error": error.splitlines()[0] if error else "Windows OCR failed."}
    payload = json.loads(result.stdout)
    payload["available"] = True
    return payload


def pbi_validate_pbix_reopen_tool(
    *,
    pbix_path: str,
    timeout_seconds: int = 60,
    screenshot_path: str | None = None,
    close_after: bool = False,
    analyze_screenshot: bool = True,
    use_windows_ocr: bool = False,
) -> dict[str, Any]:
    """Open a PBIX in Power BI Desktop and scan for visible repair-error signals.

    ``use_windows_ocr`` defaults to ``False`` because the underlying screenshot
    captures the entire primary desktop, not just the Power BI window. The
    OCR helper now returns only the matched signal labels and a length —
    never raw recognized text — so screen contents from other applications
    cannot leak through the response.
    """
    from . import _analyze_reopen_screenshot, _ocr_reopen_screenshot, _run_reopen_probe

    if timeout_seconds < 10 or timeout_seconds > 300:
        raise PowerBIValidationError(
            "timeout_seconds must be between 10 and 300.", details={"timeout_seconds": timeout_seconds}
        )
    if os.name != "nt":
        raise PowerBIValidationError("PBIX reopen validation is only supported on Windows.")
    pbix = resolve_local_path(pbix_path, must_exist=True, allowed_extensions={".pbix"})
    screenshot = (
        resolve_local_path(screenshot_path, must_exist=False, allowed_extensions={".png"}) if screenshot_path else None
    )
    if screenshot:
        screenshot.parent.mkdir(parents=True, exist_ok=True)
    persistence = pbi_validate_pbix_persistence_tool(pbix_path=str(pbix), require_security_bindings_removed=False)
    probe = _run_reopen_probe(
        pbix_path=pbix, timeout_seconds=timeout_seconds, screenshot_path=screenshot, close_after=close_after
    )
    issues: list[dict[str, Any]] = []
    warnings: list[dict[str, Any]] = []
    screenshot_analysis: dict[str, Any] | None = None
    ocr: dict[str, Any] | None = None
    if not persistence.get("valid"):
        issues.append({"type": "pbix_persistence_invalid", "issues": persistence.get("issues", [])})
    if not probe.get("opened"):
        issues.append({"type": "powerbi_window_not_opened", "timeout_seconds": timeout_seconds})
    if probe.get("ui_text_matches"):
        issues.append({"type": "powerbi_fix_this_signal", "matches": probe["ui_text_matches"]})
    if screenshot and not probe.get("screenshot_path"):
        warnings.append({"type": "screenshot_not_created", "screenshot_path": str(screenshot)})
    if analyze_screenshot and probe.get("screenshot_path"):
        screenshot_analysis = _analyze_reopen_screenshot(Path(str(probe["screenshot_path"])))
        if not screenshot_analysis.get("available", False):
            warnings.append({"type": "screenshot_analysis_failed", "error": screenshot_analysis.get("error")})
        elif screenshot_analysis.get("fix_this_like"):
            issues.append({"type": "screenshot_fix_this_like_regions", "analysis": screenshot_analysis})
    if use_windows_ocr and probe.get("screenshot_path"):
        ocr = _ocr_reopen_screenshot(Path(str(probe["screenshot_path"])))
        if not ocr.get("available", False):
            warnings.append({"type": "windows_ocr_unavailable", "error": ocr.get("error")})
        elif ocr.get("matches"):
            issues.append({"type": "windows_ocr_fix_this_signal", "matches": ocr["matches"]})
    return ok(
        f"PBIX reopen validation found {len(issues)} issue(s), {len(warnings)} warning(s).",
        pbix_path=str(pbix),
        valid=not issues,
        issue_count=len(issues),
        warning_count=len(warnings),
        issues=issues,
        warnings=warnings,
        persistence=persistence,
        reopen=probe,
        screenshot_analysis=screenshot_analysis,
        windows_ocr=ocr,
    )
