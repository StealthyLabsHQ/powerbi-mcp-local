"""UI-automation tools (Windows-only, opt-in).

These tools drive Power BI Desktop's UI through Win32 keyboard injection
because the Tabular SSAS engine has no public "save the .pbix to disk" API
— Ctrl+S in the Desktop app is the only way. Every TOM mutation done by
this server lives in memory until the user persists. ``pbi_persist_now``
gives the LLM client a way to ask Desktop to flush.

Safety model:
- Hard gated by env var ``PBI_MCP_ALLOW_UI_AUTOMATION=1`` (must be set when
  the server starts).
- Hard gated by call-time ``confirm=True``.
- Only sends the ``Ctrl+S`` key combination — no other key sequences.
- Targets only the foreground window of the PID published by the live
  AS instance (``manager.state.instance.pid``); falls back to the most
  recent ``PBIDesktop.exe`` process if no manager is connected.
- Best-effort observation: when ``pbix_path`` is supplied, the call polls
  the file's modification timestamp up to ``timeout_seconds`` and returns
  the observed delta.
"""

from __future__ import annotations

import os
import platform
import time
from pathlib import Path
from typing import Any

from pbi_connection import PowerBIValidationError, ok
from security import resolve_local_path


def _ensure_windows() -> None:
    if platform.system() != "Windows":
        raise PowerBIValidationError(
            "pbi_persist_now requires Windows (Power BI Desktop is Windows-only).",
            details={"platform": platform.system()},
        )


def _ensure_opt_in() -> None:
    if os.environ.get("PBI_MCP_ALLOW_UI_AUTOMATION", "0") != "1":
        raise PowerBIValidationError(
            "UI automation is disabled. Set PBI_MCP_ALLOW_UI_AUTOMATION=1 in the server's "
            "environment before starting the server to enable pbi_persist_now.",
            details={"env_var": "PBI_MCP_ALLOW_UI_AUTOMATION"},
        )


def _resolve_pid_from_manager(manager: Any | None) -> int | None:
    if manager is None:
        return None
    state = getattr(manager, "_state", None)
    if state is None:
        return None
    instance = getattr(state, "instance", None)
    pid = getattr(instance, "pid", None) if instance is not None else None
    return int(pid) if pid else None


def _fallback_pbidesktop_pid() -> int | None:
    """Return the PID of the most recently started PBIDesktop.exe, or None."""
    try:
        import psutil
    except ImportError:
        return None
    candidates: list[tuple[float, int]] = []
    for proc in psutil.process_iter(attrs=["pid", "name", "create_time"]):
        try:
            name = (proc.info.get("name") or "").lower()
        except Exception:
            continue
        if name == "pbidesktop.exe":
            candidates.append((float(proc.info.get("create_time") or 0), int(proc.info["pid"])))
    if not candidates:
        return None
    candidates.sort()
    return candidates[-1][1]


def _find_main_window_hwnd(pid: int) -> int | None:
    """Return the visible top-level window HWND owned by ``pid``.

    Picks the first window with non-empty title, no owner, and IsWindowVisible.
    """
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32

    EnumWindowsProc = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)

    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    user32.IsWindowVisible.argtypes = [wintypes.HWND]
    user32.IsWindowVisible.restype = wintypes.BOOL
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.GetWindow.argtypes = [wintypes.HWND, ctypes.c_uint]
    user32.GetWindow.restype = wintypes.HWND

    GW_OWNER = 4
    found: list[int] = []

    def _callback(hwnd: int, _lparam: int) -> bool:
        proc_pid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(proc_pid))
        if int(proc_pid.value) != int(pid):
            return True
        if not user32.IsWindowVisible(hwnd):
            return True
        if user32.GetWindowTextLengthW(hwnd) <= 0:
            return True
        if user32.GetWindow(hwnd, GW_OWNER):
            return True
        found.append(int(hwnd))
        return True

    user32.EnumWindows(EnumWindowsProc(_callback), 0)
    return found[0] if found else None


def _read_window_title(hwnd: int) -> str:
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = [wintypes.HWND, ctypes.c_wchar_p, ctypes.c_int]
    user32.GetWindowTextW.restype = ctypes.c_int

    length = user32.GetWindowTextLengthW(hwnd)
    if length <= 0:
        return ""
    buf = ctypes.create_unicode_buffer(length + 1)
    user32.GetWindowTextW(hwnd, buf, length + 1)
    return buf.value


def _bring_to_foreground(hwnd: int) -> int | None:
    """Restore + bring the given window to foreground. Returns the previous
    foreground HWND on success.
    """
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.IsIconic.argtypes = [wintypes.HWND]
    user32.IsIconic.restype = wintypes.BOOL
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.ShowWindow.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.restype = wintypes.BOOL

    SW_RESTORE = 9
    previous = user32.GetForegroundWindow()
    if user32.IsIconic(hwnd):
        user32.ShowWindow(hwnd, SW_RESTORE)
    success = bool(user32.SetForegroundWindow(hwnd))
    return int(previous) if success else None


def _send_ctrl_s_via_sendinput(hwnd: int) -> None:
    """Legacy fallback: inject Ctrl+S into the global input queue.

    Some hosts (notably WPF) only respond to translated keyboard input.
    Operators can opt back into this path with
    ``PBI_MCP_PERSIST_USE_SENDINPUT=1`` if the safer PostMessage variant
    fails to trigger a save on their build of Power BI Desktop.
    """
    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32

    INPUT_KEYBOARD = 1
    KEYEVENTF_KEYUP = 2
    VK_CONTROL = 0x11
    VK_S = 0x53

    class KEYBDINPUT(ctypes.Structure):
        _fields_ = [
            ("wVk", wintypes.WORD),
            ("wScan", wintypes.WORD),
            ("dwFlags", wintypes.DWORD),
            ("time", wintypes.DWORD),
            ("dwExtraInfo", ctypes.c_void_p),
        ]

    class _UN(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    class INPUT(ctypes.Structure):
        _fields_ = [("type", wintypes.DWORD), ("u", _UN)]

    inputs = (INPUT * 4)()
    for index in range(4):
        inputs[index].type = INPUT_KEYBOARD
        inputs[index].u.ki.time = 0
        inputs[index].u.ki.wScan = 0
        inputs[index].u.ki.dwExtraInfo = None
    inputs[0].u.ki.wVk = VK_CONTROL
    inputs[0].u.ki.dwFlags = 0
    inputs[1].u.ki.wVk = VK_S
    inputs[1].u.ki.dwFlags = 0
    inputs[2].u.ki.wVk = VK_S
    inputs[2].u.ki.dwFlags = KEYEVENTF_KEYUP
    inputs[3].u.ki.wVk = VK_CONTROL
    inputs[3].u.ki.dwFlags = KEYEVENTF_KEYUP

    user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
    user32.SendInput.restype = wintypes.UINT
    sent = user32.SendInput(4, inputs, ctypes.sizeof(INPUT))
    if sent != 4:
        raise PowerBIValidationError(
            f"SendInput injected {sent}/4 events — keyboard layer rejected the call.",
            details={"hwnd": hwnd, "sent": sent},
        )


def _send_ctrl_s(hwnd: int) -> None:
    """Inject Ctrl+S to the specific Power BI Desktop ``hwnd``.

    Default path uses ``PostMessage`` against the focused descendant of
    ``hwnd`` instead of ``SendInput``. ``SendInput`` injects key events
    into the global input queue, which routes to whichever window happens
    to own the foreground when the events are processed — meaning a focus
    race between ``SetForegroundWindow`` and event processing could
    deliver Ctrl+S to an unrelated app. ``PostMessage`` posts directly to
    a window handle so the chord stays bound to PBI Desktop even if focus
    moves.

    Set ``PBI_MCP_PERSIST_USE_SENDINPUT=1`` to fall back to the legacy
    SendInput path on builds where WPF's input pipeline ignores posted
    keyboard messages.
    """
    if os.environ.get("PBI_MCP_PERSIST_USE_SENDINPUT", "0") == "1":
        _send_ctrl_s_via_sendinput(hwnd)
        return

    import ctypes
    from ctypes import wintypes

    user32 = ctypes.windll.user32

    WM_KEYDOWN = 0x0100
    WM_KEYUP = 0x0101
    VK_CONTROL = 0x11
    VK_S = 0x53

    user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
    user32.PostMessageW.restype = wintypes.BOOL
    user32.GetGUIThreadInfo.restype = wintypes.BOOL
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD

    # Resolve the focused descendant of hwnd through GetGUIThreadInfo so the
    # chord lands on the actual edit-control / canvas owning input focus.
    class GUITHREADINFO(ctypes.Structure):
        _fields_ = [
            ("cbSize", wintypes.DWORD),
            ("flags", wintypes.DWORD),
            ("hwndActive", wintypes.HWND),
            ("hwndFocus", wintypes.HWND),
            ("hwndCapture", wintypes.HWND),
            ("hwndMenuOwner", wintypes.HWND),
            ("hwndMoveSize", wintypes.HWND),
            ("hwndCaret", wintypes.HWND),
            ("rcCaret", wintypes.RECT),
        ]

    info = GUITHREADINFO()
    info.cbSize = ctypes.sizeof(GUITHREADINFO)
    user32.GetGUIThreadInfo.argtypes = [wintypes.DWORD, ctypes.POINTER(GUITHREADINFO)]
    target_thread = user32.GetWindowThreadProcessId(hwnd, None)
    target_hwnd = hwnd
    if user32.GetGUIThreadInfo(target_thread, ctypes.byref(info)) and info.hwndFocus:
        target_hwnd = info.hwndFocus

    posts = [
        (WM_KEYDOWN, VK_CONTROL),
        (WM_KEYDOWN, VK_S),
        (WM_KEYUP, VK_S),
        (WM_KEYUP, VK_CONTROL),
    ]
    for msg, vk in posts:
        if not user32.PostMessageW(target_hwnd, msg, vk, 0):
            raise PowerBIValidationError(
                "PostMessage failed — Power BI Desktop rejected the key chord.",
                details={"hwnd": int(target_hwnd), "msg": msg, "vk": vk},
            )


def pbi_persist_now_tool(
    pbix_path: str | None = None,
    confirm: bool = False,
    timeout_seconds: int = 10,
    *,
    manager: Any | None = None,
) -> dict[str, Any]:
    """Trigger Power BI Desktop to save the open PBIX (Ctrl+S in the UI).

    The Tabular engine has no programmatic save — every TOM mutation lives
    in memory until the user presses Ctrl+S. This tool drives that key
    chord through Win32 keyboard injection on the connected Desktop
    instance's main window.

    **Hard gates** (both required):

    - The server process must have ``PBI_MCP_ALLOW_UI_AUTOMATION=1`` set in
      its environment before launch.
    - The caller must pass ``confirm=True`` per call.

    Either gate failing returns a structured error without acting. The
    automation only sends ``Ctrl+S``; no other key sequences are emitted.

    Parameters
    ----------
    pbix_path:
        Optional. When provided, the call polls the file's modification
        timestamp for up to ``timeout_seconds`` after sending Ctrl+S, and
        reports the observed delta in the response. Without it the call
        returns immediately after key injection.
    confirm:
        Must be ``True``. Refusing-by-default avoids accidental focus
        steals from other tools.
    timeout_seconds:
        Bound on how long to wait for ``pbix_path`` mtime to change.
        Clamped to ``[1, 60]``.
    """
    if not confirm:
        raise PowerBIValidationError(
            "pbi_persist_now requires confirm=True (UI automation is destructive to focus).",
            details={"confirm": confirm},
        )
    _ensure_windows()
    _ensure_opt_in()

    timeout = max(1, min(int(timeout_seconds), 60))

    pid = _resolve_pid_from_manager(manager) or _fallback_pbidesktop_pid()
    if pid is None:
        raise PowerBIValidationError(
            "No Power BI Desktop process found. Open a PBIX in Desktop first.",
            details={"hint": "psutil could not locate PBIDesktop.exe"},
        )

    hwnd = _find_main_window_hwnd(pid)
    if hwnd is None:
        raise PowerBIValidationError(
            f"PBI Desktop PID {pid} has no visible top-level window. The app may be starting up.",
            details={"pid": pid},
        )

    pbix: Path | None = None
    mtime_before: float | None = None
    if pbix_path:
        pbix = resolve_local_path(pbix_path, must_exist=False, allowed_extensions={".pbix"})
        if pbix.exists():
            mtime_before = pbix.stat().st_mtime

    title = _read_window_title(hwnd)
    previous_hwnd = _bring_to_foreground(hwnd)
    # Brief settle so Ctrl+S routes to PBI Desktop instead of the previous focus owner.
    time.sleep(0.15)
    _send_ctrl_s(hwnd)

    saved = False
    waited_seconds = 0.0
    if pbix is not None and pbix_path:
        deadline = time.monotonic() + timeout
        while time.monotonic() < deadline:
            time.sleep(0.25)
            waited_seconds = round(timeout - max(0.0, deadline - time.monotonic()), 3)
            if pbix.exists():
                current = pbix.stat().st_mtime
                if mtime_before is None and current > 0:
                    saved = True
                    break
                if mtime_before is not None and current > mtime_before:
                    saved = True
                    break

    if previous_hwnd:
        try:
            _bring_to_foreground(previous_hwnd)
        except Exception:  # focus restore is best-effort
            pass

    return ok(
        "Ctrl+S sent to Power BI Desktop." + (" Save observed." if saved else ""),
        pid=pid,
        window_title=title,
        pbix_path=str(pbix) if pbix is not None else None,
        save_observed=bool(saved),
        mtime_before=mtime_before,
        mtime_after=pbix.stat().st_mtime if pbix is not None and pbix.exists() else None,
        polled_for_seconds=waited_seconds,
        timeout_seconds=timeout,
    )


__all__ = ["pbi_persist_now_tool"]
