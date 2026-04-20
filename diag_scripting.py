#!/usr/bin/env python3
"""Standalone diagnostic: does DaVinci Resolve's scripting API respond?

Runs outside the GUI so we can rule out the App's code. If this script
prints ``scriptapp -> None`` even though Resolve is open with a project
loaded and scripting is set to 'Local', then the root cause is on the
Resolve side (server not actually listening) or a Windows-level issue
(firewall, privilege mismatch) — not our pipeline.

Usage:
    py -3 diag_scripting.py
"""

from __future__ import annotations

import ctypes
import os
import platform
import subprocess
import sys
import time


def running_exe() -> str | None:
    try:
        out = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                "(Get-Process -Name Resolve -ErrorAction SilentlyContinue | "
                "Select-Object -First 1).Path",
            ],
            stderr=subprocess.DEVNULL,
            timeout=5,
        )
    except Exception:
        return None
    p = out.decode("utf-8", "replace").strip()
    return p if p and os.path.isfile(p) else None


def edition(exe: str) -> str:
    if not exe:
        return "?"
    try:
        out = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                f"(Get-Item '{exe}').VersionInfo.ProductName",
            ],
            stderr=subprocess.DEVNULL,
            timeout=5,
        )
    except Exception:
        return "?"
    return out.decode("utf-8", "replace").strip() or "?"


def is_admin() -> bool:
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def main() -> int:
    # 1. Show environment
    print("=" * 60)
    print(f"Python:        {sys.version.split()[0]} ({platform.architecture()[0]})")
    print(f"This process:  {'admin' if is_admin() else 'user'}")
    exe = running_exe()
    print(f"Running exe:   {exe or '(no Resolve.exe running)'}")
    if exe:
        print(f"Edition:       {edition(exe)}")
    print("=" * 60)

    if not exe:
        print("FAIL: Resolve.exe is not running. Start Resolve and retry.")
        return 2

    # 2. Bootstrap env vars from the RUNNING Resolve.exe's directory so we
    # guarantee we bind to the right edition's fusionscript.dll.
    install_dir = os.path.dirname(exe)
    dll = os.path.join(install_dir, "fusionscript.dll")
    edition_name = os.path.basename(install_dir)
    api_dir = os.path.join(
        r"C:\ProgramData\Blackmagic Design",
        edition_name,
        r"Support\Developer\Scripting",
    )
    modules_dir = os.path.join(api_dir, "Modules")

    print(f"Using DLL:     {dll}  exists={os.path.isfile(dll)}")
    print(f"Using API dir: {api_dir}  exists={os.path.isdir(api_dir)}")
    print(f"Using Modules: {modules_dir}  exists={os.path.isdir(modules_dir)}")

    if not (os.path.isfile(dll) and os.path.isdir(modules_dir)):
        print("FAIL: required Resolve files missing at derived paths.")
        return 3

    for key in ("RESOLVE_SCRIPT_API", "RESOLVE_SCRIPT_LIB", "PYTHONPATH"):
        os.environ.pop(key, None)
    os.environ["RESOLVE_SCRIPT_API"] = api_dir
    os.environ["RESOLVE_SCRIPT_LIB"] = dll
    if modules_dir not in sys.path:
        sys.path.insert(0, modules_dir)

    # 3. Import + single scriptapp attempt with tight timeout.
    print("=" * 60)
    print("Importing DaVinciResolveScript…")
    time.sleep(2)  # fusionscript.dll file-lock window
    try:
        import DaVinciResolveScript as bmd  # type: ignore
    except Exception as exc:
        print(f"IMPORT FAILED: {type(exc).__name__}: {exc}")
        return 4
    print("Import OK.")

    print("Calling scriptapp('Resolve') 5x, 1s apart…")
    last = None
    for i in range(5):
        r = bmd.scriptapp("Resolve")
        print(f"  attempt {i + 1}: {r!r}")
        if r is not None:
            last = r
            break
        time.sleep(1.0)

    print("=" * 60)
    if last is None:
        print("scriptapp() returned None on every attempt.")
        print("")
        print("This confirms the Resolve scripting SERVER is not answering,")
        print("independent of the Audioconverter app. Now verify INSIDE")
        print("Resolve itself:")
        print("")
        print("  1. Resolve → Workspace → Console → 'Py3' tab")
        print("  2. Paste: ")
        print("       resolve = bmd.scriptapp('Resolve'); print(resolve)")
        print("  3. If that ALSO prints None → External Scripting preference")
        print("     did not take effect. Fix: Preferences → System → General")
        print("     → External scripting using = 'Local', Save, then fully")
        print("     QUIT Resolve (Tray-icon too!) and restart it.")
        print("  4. If that prints an object there but None here → privilege")
        print("     mismatch. Close Resolve, right-click Resolve.exe →")
        print("     Properties → Compatibility → UNCHECK 'Run as admin',")
        print("     then start both fresh.")
        return 1

    print(f"scriptapp() returned: {last!r}")
    pm = last.GetProjectManager()
    print(f"GetProjectManager(): {pm!r}")
    if pm:
        proj = pm.GetCurrentProject()
        print(f"GetCurrentProject(): {proj!r}")
        if proj:
            print(f"Project name: {proj.GetName()!r}")
    print("SCRIPTING WORKS. The app's timeout must have a different cause.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
