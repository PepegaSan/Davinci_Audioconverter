#!/usr/bin/env python3
"""DaVinci Auto Audioconverter.

Four-phase pipeline:
    0. Audio preprocess: extract the source video's audio track with FFmpeg
                        and run it through DeepFilterNet to strip background
                        noise. The cleaned WAV is what Resolve will actually
                        get on the timeline — the operator never hears the
                        raw (noisy) audio.
    1. Preparation    : drop a video, import it into the Media Pool, put the
                        video on V1 and the cleaned WAV (from phase 0) on A1
                        of a fresh timeline.
    2. Manual bridge  : the UI pauses while the operator uses the Resolve
                        "AI Audio Converter" to generate a WAV. Operator clicks
                        OK when the WAV is in the Media Pool.
    3. Finalise+render: newest WAV is detected, the timeline is rebuilt with
                        video on V1 (video-only) and WAV on A1 (audio-only),
                        render queue is purged, preset is loaded and the
                        render job is launched + monitored.

All DaVinci API rules from the product spec are enforced centrally in
``_bootstrap_resolve_api`` and ``ResolveController``. Audio preprocessing
lives in ``audio_preprocess.py`` so the DaVinci integration stays
completely unaware of FFmpeg / DeepFilterNet.
"""

from __future__ import annotations

import os
import platform
import subprocess
import sys
import time
import threading
import traceback
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Any, Callable

import customtkinter as ctk

# Audio preprocessing (FFmpeg extract + DeepFilterNet denoise). Kept in a
# separate module so the DaVinci controller below doesn't import it — the
# preprocessing phase runs entirely before the Resolve API is touched and
# only its returned WAV path is handed to the controller.
from audio_preprocess import (
    AudioPreprocessError,
    EQ_DEFAULT_FREQ_HZ,
    EQ_DEFAULT_GAIN_DB,
    EQ_DEFAULT_WIDTH_Q,
    PREVIEW_DEFAULT_DURATION_S,
    build_equalizer_filter,
    cleanup_temp_files,
    preprocess_video_audio,
    preview_video_audio,
)
# Pulled in for preflight diagnostics (ffmpeg / DeepFilterNet resolution)
# — kept under an alias so the module's private helpers don't leak into
# the rest of main.py's namespace.
import audio_preprocess as _ap

from settings import AppSettings, settings_path

# Drag & drop integration. tkinterdnd2 ships a Tk subclass; we mix it into
# ``ctk.CTk`` so the window keeps the CustomTkinter chrome while supporting
# native file drops from Explorer / Finder.
try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
except Exception:  # pragma: no cover - dependency missing is handled in UI
    DND_FILES = None
    TkinterDnD = None  # type: ignore

# Self-contained theme module (lives in the repo root).
_HERE = Path(__file__).resolve().parent
if str(_HERE) not in sys.path:
    sys.path.insert(0, str(_HERE))
from theme import (  # noqa: E402
    PALETTE_DARK,
    PALETTE_LIGHT,
    BTN_H,
    FONT_APP_TITLE,
    FONT_SECTION,
    FONT_UI,
    FONT_UI_SM,
    FONT_HINT,
    button_kwargs,
)


# ---------------------------------------------------------------------------
# Resolve scripting bootstrap
# ---------------------------------------------------------------------------

# Candidate locations for the scripting ``Modules`` directory (contains
# ``DaVinciResolveScript.py``) - Blackmagic ships Free / Studio / Beta side by
# side, each under its own ProgramData folder. First match wins.
_RESOLVE_MODULE_DIRS: tuple[str, ...] = (
    r"C:\ProgramData\Blackmagic Design\DaVinci Resolve Studio\Support\Developer\Scripting\Modules",
    r"C:\ProgramData\Blackmagic Design\DaVinci Resolve\Support\Developer\Scripting\Modules",
    r"C:\ProgramData\Blackmagic Design\DaVinci Resolve Studio 21 Beta\Support\Developer\Scripting\Modules",
    r"C:\ProgramData\Blackmagic Design\DaVinci Resolve 21 Beta\Support\Developer\Scripting\Modules",
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve Studio\Support\Developer\Scripting\Modules",
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve\Support\Developer\Scripting\Modules",
)

# Candidate locations for ``fusionscript.dll`` - Blackmagic's stock
# ``DaVinciResolveScript.py`` only falls back to the Free path, so Studio/Beta
# installs break without an explicit ``RESOLVE_SCRIPT_LIB``.
_RESOLVE_LIB_CANDIDATES: tuple[str, ...] = (
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve Studio\fusionscript.dll",
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve\fusionscript.dll",
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve Studio 21 Beta\fusionscript.dll",
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve 21 Beta\fusionscript.dll",
)

# Candidate ``Resolve.exe`` paths.
_RESOLVE_EXE_CANDIDATES: tuple[str, ...] = (
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve Studio\Resolve.exe",
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve\Resolve.exe",
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve Studio 21 Beta\Resolve.exe",
    r"C:\Program Files\Blackmagic Design\DaVinci Resolve 21 Beta\Resolve.exe",
)

_DAVINCI_MODULE: Any = None  # cached imported module


def _first_existing(paths: tuple[str, ...]) -> str | None:
    for p in paths:
        if os.path.isfile(p) or os.path.isdir(p):
            return p
    return None


def _bootstrap_resolve_api() -> Any:
    """Import ``DaVinciResolveScript`` following the API-rules spec.

    Enforced here:
        - Rule 1: purge every hardcoded env var that could carry stale values
                  from a previous session before we set fresh ones.
        - Rule 2: ``time.sleep(2)`` before importing the scripting module to
                  let Resolve release the fusionscript.dll file lock that can
                  linger right after Resolve starts.

    Unlike the naive version, we search multiple install roots so Studio /
    Beta installs work out of the box - Blackmagic's stock loader only knows
    the Free path.
    """
    global _DAVINCI_MODULE
    if _DAVINCI_MODULE is not None:
        return _DAVINCI_MODULE

    # Rule 1 - strip stale env vars first.
    for key in ("RESOLVE_SCRIPT_API", "RESOLVE_SCRIPT_LIB", "PYTHONPATH"):
        os.environ.pop(key, None)

    lib_path: str | None = None
    modules_dir: str | None = None

    # Prefer the DLL that sits next to the currently-running Resolve.exe so
    # we always talk to the matching edition (Studio 21 Beta vs stable 20.x
    # etc.). Falls through to the static candidate lists when Resolve isn't
    # running or no match is found.
    running_dir = _running_resolve_dir()
    if running_dir:
        dll_candidate = os.path.join(running_dir, "fusionscript.dll")
        if os.path.isfile(dll_candidate):
            lib_path = dll_candidate
        # Resolve's installer mirrors the install folder name under
        # ProgramData\Blackmagic Design\<edition>\Support\Developer\Scripting\Modules.
        edition_name = os.path.basename(running_dir)
        mirrored = os.path.join(
            r"C:\ProgramData\Blackmagic Design",
            edition_name,
            r"Support\Developer\Scripting\Modules",
        )
        if os.path.isdir(mirrored):
            modules_dir = mirrored

    if modules_dir is None:
        modules_dir = _first_existing(_RESOLVE_MODULE_DIRS)
    if lib_path is None:
        lib_path = _first_existing(_RESOLVE_LIB_CANDIDATES)

    if modules_dir is None or lib_path is None:
        raise ResolveError(
            "Could not locate the DaVinci Resolve scripting files. Expected "
            "Modules dir under one of: "
            + ", ".join(_RESOLVE_MODULE_DIRS)
            + " and fusionscript.dll under one of: "
            + ", ".join(_RESOLVE_LIB_CANDIDATES)
        )

    # Derive RESOLVE_SCRIPT_API from the Modules dir (parent of ``Modules``).
    api_dir = os.path.dirname(modules_dir)
    os.environ["RESOLVE_SCRIPT_API"] = api_dir
    os.environ["RESOLVE_SCRIPT_LIB"] = lib_path

    # Rule 2 - guard against the fusionscript.dll file-lock race.
    time.sleep(2)

    if modules_dir not in sys.path:
        sys.path.insert(0, modules_dir)

    import DaVinciResolveScript as dvr_script  # type: ignore  # noqa: E402
    _DAVINCI_MODULE = dvr_script
    return dvr_script


def _to_forward(path: str | os.PathLike) -> str:
    """Rule 3: every path handed to the Resolve API must use forward slashes.

    Backslashes silently break media import on Windows, so we normalise early.
    """
    return str(path).replace("\\", "/")


# ---------------------------------------------------------------------------
# Resolve controller
# ---------------------------------------------------------------------------


def _is_resolve_process_running() -> bool:
    """Return True if ``Resolve.exe`` currently appears in the Windows task list.

    Used only to decide whether to skip ``Popen`` — a second launch on an
    already-running Resolve wobbles the scripting socket. NOT a gate for the
    first ``scriptapp()`` call: that one is cheap on a running Resolve.

    We read raw bytes because ``tasklist`` on a non-English Windows emits
    OEM codepage output (cp850/cp437/…) that Python's implicit cp1252
    decoder chokes on, returning ``None`` and crashing the ``in`` check.
    """
    if not sys.platform.startswith("win"):
        return False
    try:
        raw = subprocess.check_output(
            ["tasklist", "/FI", "IMAGENAME eq Resolve.exe", "/NH"],
            stderr=subprocess.DEVNULL,
            timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError):
        return False
    if not raw:
        return False
    return "Resolve.exe" in raw.decode("utf-8", errors="replace")


def _running_resolve_exe() -> str | None:
    """Return the full path to the currently-running ``Resolve.exe``, or None."""
    if not sys.platform.startswith("win"):
        return None
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
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError):
        return None
    path = out.decode("utf-8", errors="replace").strip()
    if path and os.path.isfile(path):
        return path
    return None


def _running_resolve_dir() -> str | None:
    """Return the directory of the currently-running ``Resolve.exe``, or None."""
    exe = _running_resolve_exe()
    return os.path.dirname(exe) if exe else None


def _resolve_product_name(exe_path: str) -> str | None:
    """Return the PE version-resource ``ProductName`` field of Resolve.exe.

    NOTE: in Resolve 21 this field reports ``"DaVinci Resolve"`` for BOTH
    the Free and Studio editions — it is informational only and MUST NOT
    be used to gate scripting. If you need to know whether the user has
    Studio you have to check somewhere else (license file, registry, or
    just try scriptapp and let it fail cleanly).
    """
    if not sys.platform.startswith("win") or not exe_path:
        return None
    try:
        ps_literal = exe_path.replace("'", "''")
        out = subprocess.check_output(
            [
                "powershell",
                "-NoProfile",
                "-Command",
                f"(Get-Item '{ps_literal}').VersionInfo.ProductName",
            ],
            stderr=subprocess.DEVNULL,
            timeout=5,
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError):
        return None
    product = out.decode("utf-8", errors="replace").strip()
    return product or None


def _is_python_elevated() -> bool:
    """Return ``True`` if this Python process is running with admin rights."""
    if not sys.platform.startswith("win"):
        return False
    try:
        import ctypes

        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:  # pragma: no cover - defensive: never crash on the check
        return False


# ---------------------------------------------------------------------------
# Filesystem helper used by the "Open output folder" shortcut
# ---------------------------------------------------------------------------


def _open_in_file_manager(path: str) -> None:
    """Open a filesystem path in the native file manager (Explorer / Finder
    / xdg-open). Best-effort: missing commands or a non-existent path are
    logged by the caller rather than raised, because this is always an
    optional convenience after a successful render."""
    if not path or not os.path.isdir(path):
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError:
        pass


def _open_audio_file(path: str) -> None:
    """Open an audio file in the OS's default audio player.

    Uses the same platform-specific fallbacks as :func:`_open_in_file_manager`
    — we don't bundle a player; we just hand the WAV to whatever the user
    has associated with ``.wav`` on their box (Windows Media Player,
    Groove, QuickTime, …). Perfectly sufficient for quick A/B tuning of
    the EQ advanced panel.
    """
    if not path or not os.path.isfile(path):
        return
    try:
        if sys.platform.startswith("win"):
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError:
        pass


# ---------------------------------------------------------------------------
# Preflight health-check
# ---------------------------------------------------------------------------


def run_preflight_diagnostics() -> list[tuple[str, str, str]]:
    """Return a list of ``(label, status, detail)`` tuples describing the
    current environment's readiness for the four-phase pipeline.

    ``status`` is always one of:
        ``"OK"``     — check passed (green)
        ``"WARN"``   — not strictly required, but may degrade features
        ``"FAIL"``   — a hard blocker for the pipeline

    The check is entirely read-only — it does NOT start Resolve, does
    NOT download models, and does NOT call ``scriptapp``. All it does
    is look at paths on disk, environment variables, and the task list,
    so it is safe to invoke as often as the user wants.

    Design: we collect cheap signals and leave interpretation to the UI
    (which formats each tuple with a coloured prefix). That keeps the
    diagnostic layer testable without any GUI plumbing.
    """
    results: list[tuple[str, str, str]] = []

    py_ver = sys.version.split()[0]
    py_bits = platform.architecture()[0]
    py_elev = _is_python_elevated()
    results.append((
        "Python",
        "OK",
        f"{py_ver} ({py_bits})" + (" [admin]" if py_elev else " [user]"),
    ))

    # FFmpeg — hard requirement for Phase 0 (extraction + EQ). Not
    # needed when Phase 0 is OFF, so mark missing as WARN.
    try:
        ff = _ap._resolve_ffmpeg()
        results.append(("FFmpeg", "OK", ff))
    except AudioPreprocessError as err:
        results.append(("FFmpeg", "WARN", str(err).splitlines()[0]))

    # DeepFilterNet — WARN because users can still run with Phase 0 off.
    cli = _ap._resolve_deepfilter_cli()
    if cli:
        results.append(("DeepFilterNet CLI", "OK", cli))
    else:
        try:
            import df.enhance  # type: ignore  # noqa: F401
            results.append((
                "DeepFilterNet",
                "OK",
                "Python package available (no standalone CLI)",
            ))
        except ImportError:
            results.append((
                "DeepFilterNet",
                "WARN",
                "no CLI or Python package — Phase 0 will fail if enabled. "
                "Drop deep-filter*.exe next to main.py or disable the "
                "'Clean source audio' switch.",
            ))

    # Resolve scripting files. Missing = FAIL, because nothing in Phase
    # 1-3 can run without them.
    modules_dir = _first_existing(_RESOLVE_MODULE_DIRS)
    if modules_dir:
        results.append(("Resolve scripting modules", "OK", modules_dir))
    else:
        results.append((
            "Resolve scripting modules",
            "FAIL",
            "DaVinciResolveScript.py not found under any of the standard "
            "paths — reinstall Resolve Studio.",
        ))

    lib = _first_existing(_RESOLVE_LIB_CANDIDATES)
    if lib:
        results.append(("Resolve fusionscript.dll", "OK", lib))
    else:
        results.append((
            "Resolve fusionscript.dll",
            "FAIL",
            "fusionscript.dll not found next to any installed Resolve.exe.",
        ))

    exe = _first_existing(_RESOLVE_EXE_CANDIDATES)
    if exe:
        results.append(("Resolve.exe", "OK", exe))
    else:
        results.append((
            "Resolve.exe",
            "FAIL",
            "Resolve.exe not at the expected install path.",
        ))

    running = _is_resolve_process_running()
    running_exe = _running_resolve_exe()
    if running and running_exe:
        # Flag a privilege-level mismatch early — it's the #1 cause of
        # "connection hangs for 90s then fails" reports.
        resolve_admin_hint = ""
        try:
            # Heuristic: if Resolve.exe lives under Program Files *and*
            # we were launched without elevation, Resolve might have
            # been started as admin from a shortcut. We can't detect
            # that directly without querying the process token, so we
            # only flag the case where the current python IS elevated
            # but resolve's scripting socket might not be — which is
            # rarer but still worth surfacing.
            if py_elev:
                resolve_admin_hint = (
                    "  NOTE: Python is running as admin. If Resolve was "
                    "started as a normal user, scripting will fail until "
                    "both run at the SAME privilege level."
                )
        except Exception:
            pass
        results.append((
            "Resolve process",
            "OK",
            f"running: {running_exe}{resolve_admin_hint}",
        ))
    elif running:
        results.append((
            "Resolve process",
            "OK",
            "Resolve.exe is in the task list (path lookup failed — ok).",
        ))
    else:
        results.append((
            "Resolve process",
            "WARN",
            "not running — the pipeline will auto-launch it on Start.",
        ))

    # Environment variables — set by _bootstrap_resolve_api on first
    # use; until then they're empty, which is expected, not an error.
    lib_env = os.environ.get("RESOLVE_SCRIPT_LIB", "")
    api_env = os.environ.get("RESOLVE_SCRIPT_API", "")
    if lib_env and api_env:
        results.append(("RESOLVE_SCRIPT_* env", "OK", f"lib={lib_env}"))
    else:
        results.append((
            "RESOLVE_SCRIPT_* env",
            "WARN",
            "not yet set (bootstrap happens on first connect — harmless).",
        ))

    return results


class ResolveError(RuntimeError):
    """Raised when the Resolve API returns an unexpected / missing object."""


class ResolveController:
    """Encapsulates every Resolve API call used by the pipeline.

    Each step is wrapped with defensive checks so failure modes surface with a
    readable message rather than a ``None`` dereference deep in the API.
    """

    DEFAULT_RENDER_PRESET = "YouTube - 1080p"
    FALLBACK_RENDER_PRESET = "H.264 Master"
    RENDER_TIMEOUT_S = 60 * 60  # Rule 8: hard stop after 1h to avoid infinite loops.
    MEDIA_IMPORT_SETTLE_S = 1.5  # Rule 6: allow the Media Pool to settle after import.

    # Each scriptapp() call blocks up to ~5s during Resolve's cold boot while
    # waiting for the scripting server socket, so a healthy cold start resolves
    # in 2-3 attempts (~10-15s wall clock). 90s covers slow boxes comfortably
    # without hanging forever on a broken setup (Rule 8).
    RESOLVE_STARTUP_TIMEOUT_S = 90
    RESOLVE_POLL_INTERVAL_S = 2.0
    RESOLVE_DIAG_AFTER_S = 18.0

    def __init__(self) -> None:
        self._resolve: Any = None
        self._project: Any = None
        self._media_pool: Any = None
        self._root_folder: Any = None

    # ------------------------------------------------------------------ setup
    def connect(
        self,
        status_callback: Callable[[str], None] | None = None,
        auto_launch: bool = True,
    ) -> None:
        """Connect to a Resolve instance, optionally launching it first.

        Ordering matters:
            1. Try ``scriptapp("Resolve")`` right away — on an already-running
               Resolve it returns instantly; no point running preliminary
               checks that would only add latency.
            2. If it came back ``None`` **and** Resolve.exe is not in the task
               list, launch Resolve. We explicitly skip ``Popen`` when the
               process is already running: a second launch causes the
               already-running instance's scripting socket to wobble, which is
               the most common reason automation mysteriously stops working
               mid-session.
            3. Poll ``scriptapp`` every ``RESOLVE_POLL_INTERVAL_S`` seconds
               until it succeeds or ``RESOLVE_STARTUP_TIMEOUT_S`` elapses.
               Each call holds the GIL only briefly, keeping the Tk main loop
               responsive.

        Also creates a scratch project if Resolve is up but lands the user in
        the Project Manager — otherwise ``GetCurrentProject`` stays ``None``
        forever and every subsequent API call errors out.
        """
        def _log(msg: str) -> None:
            if status_callback is not None:
                status_callback(msg)

        dvr_script = _bootstrap_resolve_api()

        # Surface exactly which files the API bound to — critical when a
        # connection never succeeds because a mismatched DLL silently
        # loaded nothing.
        lib_env = os.environ.get("RESOLVE_SCRIPT_LIB", "?")
        api_env = os.environ.get("RESOLVE_SCRIPT_API", "?")
        running_exe = _running_resolve_exe()
        product_name = _resolve_product_name(running_exe) if running_exe else None
        py_elevated = _is_python_elevated()

        _log(f"Python: {sys.version.split()[0]} ({platform.architecture()[0]})"
             + (" [admin]" if py_elevated else " [user]"))
        _log(f"Scripting lib: {lib_env}")
        _log(f"Scripting API: {api_env}")
        if running_exe:
            _log(f"Running exe:   {running_exe}"
                 + (f"  ({product_name})" if product_name else ""))

        # NOTE: Resolve 21's DaVinciResolveScript.py ends with
        # ``sys.modules[__name__] = script_module`` — if the DLL fails to
        # load the *whole import* raises ImportError. So there is no
        # "silent scriptModule=None" state to guard against here; any
        # silent failure would've raised in _bootstrap_resolve_api(). A
        # connection failure reaching this point means the DLL loaded,
        # but Resolve's scripting SERVER is not answering.
        # We deliberately do NOT gate on Studio vs Free detection any
        # more: Resolve 21 reports the same ``ProductName`` for both
        # editions, so any such check is a false-positive risk. Let
        # scriptapp() fail cleanly and surface that in the error.

        resolve = dvr_script.scriptapp("Resolve")
        if resolve is None:
            if auto_launch and not _is_resolve_process_running():
                _log("DaVinci Resolve not running — launching…")
                if not self._launch_resolve():
                    raise ResolveError(
                        "Could not find Resolve.exe under any of the default "
                        "install paths: " + ", ".join(_RESOLVE_EXE_CANDIDATES)
                    )
            elif _is_resolve_process_running():
                # Process is up but scripting server isn't responding yet -
                # just wait, don't hit it with another Popen.
                _log("Resolve is starting — waiting for scripting server…")
            elif not auto_launch:
                raise ResolveError(
                    "DaVinci Resolve is not running. Start Resolve Studio and "
                    "open a project first."
                )

            resolve = self._poll_for_scriptapp(dvr_script, log=_log)
            if resolve is None:
                running_exe_now = _running_resolve_exe() or "(Resolve.exe not detected)"
                py_admin = "admin" if py_elevated else "user"
                raise ResolveError(
                    "Could not connect to Resolve within "
                    f"{self.RESOLVE_STARTUP_TIMEOUT_S}s.\n\n"
                    f"Python:        {sys.version.split()[0]} "
                    f"({platform.architecture()[0]}) [{py_admin}]\n"
                    f"Running exe:   {running_exe_now}\n"
                    f"Scripting lib: {lib_env}\n"
                    f"Scripting API: {api_env}\n\n"
                    "Remaining causes when paths look right: "
                    "(1) Resolve was started as administrator but Python is "
                    f"currently running as '{py_admin}' — Windows isolates "
                    "the scripting socket per privilege level, so you must "
                    "run both at the SAME level; "
                    "(2) a modal dialog or onboarding screen is blocking "
                    "Resolve (click through any 'Welcome' / 'Quick Setup' "
                    "wizard); "
                    "(3) Resolve is still on the Project Manager screen — "
                    "double-click a project so scriptapp() can bind; "
                    "(4) another Python/VSCode session still holds the "
                    "scripting socket — close everything else and retry."
                )

        self._resolve = resolve

        project_manager = self._resolve.GetProjectManager()
        if project_manager is None:
            raise ResolveError("Could not access the Project Manager.")

        self._project = project_manager.GetCurrentProject()
        if self._project is None:
            # Cold-launched Resolve often lands in the Project Manager. Create
            # a scratch project so the pipeline can proceed unattended.
            _log("No project open — creating a scratch project…")
            fallback_name = f"AutoAudio_{int(time.time())}"
            self._project = project_manager.CreateProject(fallback_name)
            if self._project is None:
                raise ResolveError(
                    "No project is open and a fallback project could not be "
                    "created. Open a project in Resolve and retry."
                )

        self._media_pool = self._project.GetMediaPool()
        if self._media_pool is None:
            raise ResolveError("Could not access the Media Pool.")
        self._root_folder = self._media_pool.GetRootFolder()

    def _poll_for_scriptapp(
        self,
        dvr_script: Any,
        *,
        log: Callable[[str], None],
    ) -> Any:
        """Poll ``scriptapp("Resolve")`` until it hands back a live object.

        Returns the connected object or ``None`` on timeout. Heartbeat status
        messages every ~8s so the user sees progress instead of a frozen UI.

        Also surfaces the most common root cause (External Scripting not set
        to 'Local') after the second failed attempt, rather than burying it
        in a timeout message 90s later.
        """
        start = time.monotonic()
        deadline = start + self.RESOLVE_STARTUP_TIMEOUT_S
        last_heartbeat = start
        attempt = 0
        preference_hint_logged = False
        while time.monotonic() < deadline:
            attempt += 1
            time.sleep(self.RESOLVE_POLL_INTERVAL_S)
            resolve = dvr_script.scriptapp("Resolve")
            if resolve is not None:
                elapsed = time.monotonic() - start
                log(f"Resolve is up after ~{elapsed:.0f}s (attempt {attempt}).")
                return resolve

            # Early actionable hint: Resolve is clearly running (process
            # check passes) and we've already waited ~4s — the scripting
            # socket never binds if 'External scripting using' is not set
            # to 'Local', so the poll loop will *never* succeed.
            if (
                not preference_hint_logged
                and attempt >= 2
                and _is_resolve_process_running()
            ):
                log(
                    "Resolve is running but not answering scripting. "
                    "Fix: Preferences → System → General → 'External "
                    "scripting using' = Local, THEN restart Resolve."
                )
                preference_hint_logged = True

            now = time.monotonic()
            if now - last_heartbeat >= 8.0:
                remaining = max(0.0, deadline - now)
                log(f"Waiting for Resolve scripting server… ({remaining:.0f}s left)")
                last_heartbeat = now
        return None

    @staticmethod
    def _launch_resolve() -> bool:
        """Start ``Resolve.exe`` detached from this process.

        Returns ``True`` if one of the candidate paths existed and was spawned.
        """
        for exe in _RESOLVE_EXE_CANDIDATES:
            if os.path.isfile(exe):
                creation = 0
                if hasattr(subprocess, "DETACHED_PROCESS"):
                    creation = subprocess.DETACHED_PROCESS  # type: ignore[attr-defined]
                try:
                    subprocess.Popen(
                        [exe],
                        close_fds=True,
                        creationflags=creation,
                        cwd=os.path.dirname(exe),
                    )
                    return True
                except OSError:
                    continue
        return False

    # --------------------------------------------------------------- media IO
    def import_video(self, video_path: str) -> Any:
        """Import the dropped video into the Media Pool and return the clip."""
        assert self._media_pool is not None
        forward = _to_forward(video_path)
        imported = self._media_pool.ImportMedia([forward])
        if not imported:
            raise ResolveError(f"Resolve refused to import the file: {forward}")
        # Rule 6: Media Pool indexing is asynchronous; give it a moment.
        time.sleep(self.MEDIA_IMPORT_SETTLE_S)
        return imported[0]

    @staticmethod
    def probe_clip(clip: Any) -> tuple[str, str]:
        """Return (fps_str, resolution_label) with safe fallbacks.

        Rule 6: never trust metadata blindly. Some containers report no FPS or
        no resolution to the API, so we substitute "25" / "1920x1080" if the
        fields are missing or unparseable.

        FPS comes back as a **string** deliberately. Resolve's
        ``SetSetting('timelineFrameRate', …)`` is pedantic about format:
        "25" is accepted, "25.0" can be rejected silently; "29.97" works,
        "29.97002997" may not. Forwarding the exact string Resolve itself
        handed us via ``GetClipProperty('FPS')`` is what reliably sticks.
        """
        fps = "25"
        resolution = "1920x1080"
        try:
            raw_fps = clip.GetClipProperty("FPS") if clip else None
            if raw_fps:
                candidate = str(raw_fps).strip()
                # Trim trailing ".0" so "25.0" becomes "25" — Resolve
                # accepts both sometimes, but the bare int form is the
                # format Resolve exposes in its own preset dropdown and is
                # the safest bet. Keep fractional values ("23.976",
                # "29.97") verbatim.
                try:
                    as_float = float(candidate.split()[0])
                    if as_float > 0:
                        candidate = (
                            str(int(as_float))
                            if as_float.is_integer()
                            else candidate.split()[0]
                        )
                    else:
                        candidate = "25"
                except (TypeError, ValueError):
                    pass
                fps = candidate
        except Exception:
            fps = "25"
        try:
            w = clip.GetClipProperty("Resolution") if clip else None
            if w:
                resolution = str(w).strip()
        except Exception:
            resolution = "1920x1080"
        return fps, resolution

    # --------------------------------------------------------------- timeline
    AUTO_TIMELINE_PREFIX = "AutoAudio_"

    def cleanup_auto_timelines(self) -> int:
        """Delete previous auto-generated timelines so the project frame
        rate is no longer locked.

        Resolve refuses ``SetSetting('timelineFrameRate', …)`` silently
        whenever *any* timeline exists in the project at a different
        rate. When a user runs this tool more than once the leftover
        ``AutoAudio_<ts>`` timelines from previous runs cause the new
        FPS to be silently ignored.

        We only delete our own — identified by the ``AutoAudio_``
        prefix — so user timelines stay untouched. Returns the number
        of timelines actually removed.
        """
        assert self._project is not None and self._media_pool is not None
        try:
            count = int(self._project.GetTimelineCount() or 0)
        except Exception:
            return 0
        if count <= 0:
            return 0

        victims: list[Any] = []
        for i in range(1, count + 1):
            try:
                tl = self._project.GetTimelineByIndex(i)
            except Exception:
                continue
            if not tl:
                continue
            try:
                name = str(tl.GetName() or "")
            except Exception:
                continue
            if name.startswith(self.AUTO_TIMELINE_PREFIX):
                victims.append(tl)
        if not victims:
            return 0
        try:
            # Resolve ≥17: DeleteTimelines takes a list of timeline objects
            # and returns True/False. Wrap in try so a broken API shape
            # degrades gracefully instead of crashing the pipeline.
            self._media_pool.DeleteTimelines(victims)
        except Exception:
            return 0
        return len(victims)

    def apply_project_timeline_settings(
        self, fps: Any, resolution: str
    ) -> tuple[int, int, str]:
        """Force the project's timeline settings to match the source clip
        so any timeline created next inherits matching width/height/FPS.

        Must be called BEFORE ``create_fresh_timeline`` — ``CreateEmptyTimeline``
        snapshots the project settings at creation time, so setting them
        afterwards has no effect on the active timeline.

        Accepts ``fps`` as a string ("25", "29.97", "23.976") OR float;
        strings are forwarded verbatim so Resolve sees exactly what it
        itself reported via ``GetClipProperty('FPS')``. This matters —
        ``SetSetting('timelineFrameRate', ...)`` is silent-fail on odd
        forms like "25.0" or "29.97002997".

        Returns the ``(width, height, fps_str)`` actually applied (via
        ``GetSetting`` read-back so the caller can verify Resolve
        accepted the values). Falls back to 1920×1080 @ "25" on
        unparseable metadata — Rule 6.
        """
        assert self._project is not None

        # Parse "1920x1080" / "1920 x 1080" with whitespace tolerance.
        width, height = 1920, 1080
        try:
            w_str, h_str = (
                str(resolution).lower().replace(" ", "").split("x", 1)
            )
            width = int(w_str)
            height = int(h_str)
        except (ValueError, AttributeError):
            pass  # keep 1920x1080 fallback

        # Normalise FPS: keep the original string form if it looks
        # numeric, otherwise fall back. Avoid the ``float → str`` round
        # trip because that introduces ``.0`` for integer rates and can
        # shorten fractional ones, both of which Resolve has been
        # observed to reject silently.
        fps_str = str(fps).strip() if fps is not None else ""
        try:
            fps_float = float(fps_str)
            if fps_float <= 0:
                raise ValueError
        except (TypeError, ValueError):
            fps_str = "25"
            fps_float = 25.0
        # If the caller handed us a float like ``25.0``, collapse to "25"
        # so it matches Resolve's own preset values.
        if fps_float.is_integer() and ("." in fps_str or fps_str.lower().endswith(("e", "e+"))):
            fps_str = str(int(fps_float))

        # Resolve expects string values for these settings.
        self._project.SetSetting("timelineResolutionWidth", str(width))
        self._project.SetSetting("timelineResolutionHeight", str(height))
        self._project.SetSetting("timelineFrameRate", fps_str)

        # Tiny settle — Resolve sometimes lags one poll behind when you
        # SetSetting + GetSetting + CreateEmptyTimeline back to back.
        time.sleep(0.3)

        # Verify what actually took. If Resolve rejected our FPS string
        # (e.g. mid-project change that requires an empty timeline) this
        # is where we find out instead of silently rendering at the wrong
        # frame rate.
        try:
            applied_fps = self._project.GetSetting("timelineFrameRate") or fps_str
            applied_fps = str(applied_fps).strip() or fps_str
        except Exception:
            applied_fps = fps_str

        return width, height, applied_fps

    def create_fresh_timeline(self, name: str) -> Any:
        """Create a new empty timeline and mark it as current (Rule 5)."""
        assert self._media_pool is not None and self._project is not None
        timeline = self._media_pool.CreateEmptyTimeline(name)
        if timeline is None:
            raise ResolveError(f"Failed to create timeline '{name}'.")
        if not self._project.SetCurrentTimeline(timeline):
            raise ResolveError(f"Failed to activate timeline '{name}'.")
        return timeline

    def clear_current_timeline(self) -> None:
        """Remove every clip from the current timeline (Rule 5)."""
        assert self._project is not None
        timeline = self._project.GetCurrentTimeline()
        if timeline is None:
            return
        items: list[Any] = []
        for kind in ("video", "audio"):
            count = timeline.GetTrackCount(kind) or 0
            for idx in range(1, count + 1):
                track_items = timeline.GetItemListInTrack(kind, idx) or []
                items.extend(track_items)
        if items:
            timeline.DeleteClips(items, False)

    def append_full_clip(self, clip: Any) -> None:
        """Append ``clip`` with BOTH video and audio tracks, starting at 0.

        Used in Phase 1 so the operator has the source audio in the
        timeline to feed into Resolve's AI Voice / Audio Converter. We
        don't pass ``mediaType`` here — omitting it tells the API to
        include every native track the clip carries.
        """
        assert self._media_pool is not None
        ok = self._media_pool.AppendToTimeline([{"mediaPoolItem": clip}])
        if not ok:
            raise ResolveError("AppendToTimeline() failed for the full clip.")

    def append_video_only(self, clip: Any) -> None:
        """Append the picture part of ``clip`` to V1, starting at frame 0."""
        assert self._media_pool is not None
        ok = self._media_pool.AppendToTimeline([{"mediaPoolItem": clip, "mediaType": 1}])
        if not ok:
            raise ResolveError("AppendToTimeline() failed for the video track.")

    def append_audio_only(self, clip: Any) -> None:
        """Append the sound part of ``clip`` to A1, starting at frame 0."""
        assert self._media_pool is not None
        ok = self._media_pool.AppendToTimeline([{"mediaPoolItem": clip, "mediaType": 2}])
        if not ok:
            raise ResolveError("AppendToTimeline() failed for the audio track.")

    # ------------------------------------------------------------- media pool
    def snapshot_wav_clips(self) -> set[str]:
        """Return the set of current WAV file paths in the Media Pool."""
        return {c.GetClipProperty("File Path") for c in self._iter_wav_clips()}

    def newest_wav_since(self, known_paths: set[str]) -> Any:
        """Return the Media Pool item that was added after ``known_paths``.

        Falls back to the WAV with the most recent mtime if the delta is empty
        (for example because the operator reused an existing WAV name).
        """
        candidates = list(self._iter_wav_clips())
        if not candidates:
            raise ResolveError("No WAV file found in the Media Pool.")

        new_ones = [
            c for c in candidates
            if c.GetClipProperty("File Path") not in known_paths
        ]
        if new_ones:
            return new_ones[-1]

        # Fallback: newest mtime on disk.
        def _mtime(clip: Any) -> float:
            try:
                return Path(clip.GetClipProperty("File Path")).stat().st_mtime
            except OSError:
                return 0.0

        candidates.sort(key=_mtime, reverse=True)
        return candidates[0]

    def _iter_wav_clips(self):
        assert self._root_folder is not None
        stack = [self._root_folder]
        while stack:
            folder = stack.pop()
            for clip in folder.GetClipList() or []:
                path = clip.GetClipProperty("File Path") or ""
                if path.lower().endswith(".wav"):
                    yield clip
            stack.extend(folder.GetSubFolderList() or [])

    @staticmethod
    def _normalise_path(path: str | None) -> str:
        """Return a canonical form of ``path`` for equality comparisons.

        Media Pool clip paths come back from Resolve with mixed separators
        on Windows, so we normalise to lowercase + forward slashes before
        matching. Handles ``None`` / empty strings without raising so the
        callers can feed clip properties directly.
        """
        if not path:
            return ""
        try:
            norm = os.path.normcase(os.path.normpath(str(path)))
        except (TypeError, ValueError):
            norm = str(path).lower()
        return norm.replace("\\", "/")

    def remove_mediapool_clips(
        self,
        paths_to_remove: set[str],
        *,
        log: Callable[[str], None] | None = None,
    ) -> int:
        """Delete Media Pool items whose file path matches any entry in
        ``paths_to_remove``. Returns the number of clips actually removed.

        Matches are case-insensitive and separator-agnostic on Windows
        (see :meth:`_normalise_path`). Missing / invalid entries in the
        set are silently ignored so callers can pass *any* path
        (including the AI WAV path we only just learned about).

        Does NOT touch the file on disk — this is purely a Media Pool
        hygiene operation. The Phase-0 WAV's actual bytes are covered
        by :func:`audio_preprocess.cleanup_temp_files`; the AI WAV
        belongs to the user's Resolve session and stays put.
        """
        if not paths_to_remove:
            return 0
        assert self._media_pool is not None and self._root_folder is not None

        wanted: set[str] = {
            self._normalise_path(p) for p in paths_to_remove if p
        }
        if not wanted:
            return 0

        victims: list[Any] = []
        victim_names: list[str] = []
        stack = [self._root_folder]
        while stack:
            folder = stack.pop()
            for clip in folder.GetClipList() or []:
                try:
                    clip_path = clip.GetClipProperty("File Path") or ""
                except Exception:
                    continue
                if self._normalise_path(clip_path) in wanted:
                    victims.append(clip)
                    victim_names.append(os.path.basename(clip_path))
            stack.extend(folder.GetSubFolderList() or [])

        if not victims:
            if log:
                log("Media Pool cleanup — no matching clips to remove.")
            return 0

        try:
            ok = self._media_pool.DeleteClips(victims)
        except Exception as err:  # noqa: BLE001 - API surface is loose
            if log:
                log(f"Media Pool cleanup — DeleteClips failed: {err}")
            return 0
        if not ok:
            if log:
                log("Media Pool cleanup — DeleteClips returned False.")
            return 0
        if log:
            log(
                f"Media Pool cleanup — removed {len(victims)} clip(s): "
                + ", ".join(victim_names)
            )
        return len(victims)

    # ---------------------------------------------------------------- render
    def list_render_presets(self) -> list[str]:
        """Return every render-preset name currently available in the
        connected project (user-defined + factory). Sorted, de-duplicated.

        Requires ``connect()`` to have succeeded. Empty list on any API
        failure — callers fall back to the default/fallback constants.
        """
        assert self._project is not None, "connect() must be called first"
        try:
            names = self._project.GetRenderPresetList() or []
        except Exception:
            return []
        # Resolve sometimes hands back duplicates between factory and user
        # preset folders — dedupe case-sensitively but preserve display form.
        seen: set[str] = set()
        unique: list[str] = []
        for name in names:
            if name and name not in seen:
                seen.add(name)
                unique.append(name)
        unique.sort(key=str.lower)
        return unique

    def stop_render(self) -> None:
        """Ask Resolve to abort any render currently in progress.

        Safe to call from a UI thread — ``StopRendering`` is a simple
        RPC and the render polling loop in :meth:`render` notices via
        either the ``cancel_event`` it was handed or via
        ``IsRenderingInProgress()`` flipping to ``False``. A missing
        project (never connected) is swallowed quietly so the UI's
        Cancel button is idempotent.
        """
        if self._project is None:
            return
        try:
            self._project.StopRendering()
        except Exception:
            pass

    def render(
        self,
        output_dir: str,
        output_name: str,
        *,
        preset_name: str | None = None,
        cancel_event: threading.Event | None = None,
    ) -> bool:
        """Configure and execute a render job, monitored with a timeout.

        Rule 7: purge the queue. Rule 8: preset fallback + bounded wait
        loop. ``preset_name`` overrides the default; if it fails to load
        we still walk the default → fallback chain so the pipeline never
        silently picks an unexpected preset.

        ``cancel_event`` is an optional ``threading.Event`` the UI can
        flip to request an early stop. When set, we call
        ``StopRendering``, wait for the job to actually wind down, and
        return ``False`` so the caller can branch cleanly. Returns
        ``True`` on a completed render. Timeout still raises.
        """
        assert self._project is not None
        self._project.DeleteAllRenderJobs()

        tried: list[str] = []
        loaded = False
        for candidate in (
            preset_name,
            self.DEFAULT_RENDER_PRESET,
            self.FALLBACK_RENDER_PRESET,
        ):
            if not candidate or candidate in tried:
                continue
            tried.append(candidate)
            if self._project.LoadRenderPreset(candidate):
                loaded = True
                break
        if not loaded:
            raise ResolveError(
                "Could not load any render preset. Tried: "
                + ", ".join(tried)
                + ". Check the names in Resolve's Deliver page."
            )

        self._project.SetRenderSettings({
            "TargetDir": _to_forward(output_dir),
            "CustomName": output_name,
        })

        job_id = self._project.AddRenderJob()
        if not job_id:
            raise ResolveError("Resolve refused to queue the render job.")
        self._project.StartRendering(job_id)

        # Rule 8: bounded polling loop. Additionally checks cancel_event
        # every iteration so a UI-driven cancel stops the render within
        # at most one poll interval (~1s).
        started = time.time()
        cancelled = False
        while self._project.IsRenderingInProgress():
            if cancel_event is not None and cancel_event.is_set():
                # Ask Resolve to abort; keep polling until the flag
                # actually flips so we don't return while a stale job
                # is still flushing to disk.
                self.stop_render()
                cancelled = True
            if time.time() - started > self.RENDER_TIMEOUT_S:
                self.stop_render()
                raise ResolveError(
                    f"Render exceeded the {self.RENDER_TIMEOUT_S}s timeout and was aborted."
                )
            time.sleep(1.0)
        return not cancelled


# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------

# Cleanup-mode dropdown: human-readable label ⇄ internal key. Ordered
# from least to most aggressive so the dropdown reads top-down. Kept at
# module scope so both the widget builder and the pipeline reader share
# one source of truth for the label strings.
CLEANUP_MODE_LABELS: tuple[tuple[str, str], ...] = (
    ("Off — keep everything",                       "off"),
    ("Temp files only — reclaim %TEMP% disk space", "temp"),
    ("Temp + Resolve — also remove our Media Pool clips", "full"),
)
CLEANUP_LABEL_TO_MODE: dict[str, str] = {
    label: key for label, key in CLEANUP_MODE_LABELS
}
CLEANUP_MODE_TO_LABEL: dict[str, str] = {
    key: label for label, key in CLEANUP_MODE_LABELS
}


class _DnDCTk(ctk.CTk):
    """``ctk.CTk`` + tkinterdnd2 drop-target support."""

    def __init__(self, *args, **kwargs) -> None:
        super().__init__(*args, **kwargs)
        if TkinterDnD is not None:
            try:
                self.TkdndVersion = TkinterDnD._require(self)
            except Exception:  # pragma: no cover - gracefully degrade to file dialog
                self.TkdndVersion = None


class App(_DnDCTk):
    """Main application window."""

    def __init__(self) -> None:
        # Load persisted settings BEFORE any widget is built so every
        # Tk variable can be seeded with the user's last values. Load
        # failures fall back to dataclass defaults — the user never
        # sees a broken UI because of a corrupt JSON.
        self._settings = AppSettings.load()

        # Silent-save guard. Flipped True during bulk seeding so the
        # dozen ``.set()`` calls in ``__init__`` don't fire a dozen
        # disk writes — the first real user interaction saves instead.
        self._settings_silent = True

        ctk.set_default_color_theme("blue")
        initial_mode = (self._settings.appearance or "dark").strip().lower()
        ctk.set_appearance_mode("light" if initial_mode == "light" else "dark")
        palette = PALETTE_LIGHT if initial_mode == "light" else PALETTE_DARK
        super().__init__(fg_color=palette["bg"])
        self._pal: dict[str, str] = dict(palette)

        self.title("DaVinci Auto Audioconverter")
        # Honour the saved geometry if it parses as WxH; otherwise fall
        # back to the larger default (the previous 820x580 was too
        # cramped for the preset + cleanup rows).
        geom = (self._settings.window_geometry or "").strip() or "1200x900"
        self.geometry(geom)
        self.minsize(900, 640)
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        self._video_path: str | None = None
        self._phase2_event = threading.Event()
        # Flipped by the topbar's Cancel-render button; the pipeline
        # worker checks it at the start of Phase 3 and the
        # ResolveController.render() polling loop picks it up mid-render
        # to abort. Cleared at the start of every new pipeline run.
        self._render_cancel_event = threading.Event()
        self._controller = ResolveController()
        self._appearance = ctk.StringVar(value=initial_mode)

        # Remember the output dir of the most recent render so the
        # post-render dialog's "Open output folder" shortcut doesn't
        # have to re-derive it from the source path on the Tk main
        # thread (which can race with the user picking a new video).
        self._last_output_dir: str | None = None

        # Post-render cleanup mode (see CLEANUP_MODE_LABELS). Seeded
        # from persisted settings; mapped to the label via the
        # dropdown command when a user picks a different option.
        self._cleanup_mode = ctk.StringVar(
            value=self._settings.cleanup_mode or "off"
        )
        self._cleanup_expanded = bool(self._settings.cleanup_expanded)

        # Preview state — True while a preview is being processed so
        # the button can't be double-clicked, which would race on the
        # same output WAV path.
        self._preview_running = False

        # Render preset picker. StringVar is the single source of truth;
        # the combobox edits it and the pipeline reads from it right
        # before calling render(). Empty-string → "use the controller's
        # DEFAULT_RENDER_PRESET". We seed from the saved value so a
        # user-typed custom preset survives a restart even if Resolve
        # isn't running yet.
        self._render_preset = ctk.StringVar(
            value=(
                self._settings.render_preset
                or ResolveController.DEFAULT_RENDER_PRESET
            )
        )
        self._preset_loading = False

        # Phase-0 toggle — persisted. True means "denoise the audio
        # before Phase 2 so the AI voice converter gets a clean signal".
        self._audio_clean_enabled = ctk.BooleanVar(
            value=self._settings.audio_clean_enabled
        )

        # Optional parametric-EQ pass after denoising (gives voices more
        # chest depth). Kept as StringVars rather than DoubleVars so the
        # user can have an empty/mid-edit field without the Tk binding
        # rejecting it; parsed to float at pipeline start.
        self._eq_enabled = ctk.BooleanVar(value=self._settings.eq_enabled)
        self._eq_freq_str = ctk.StringVar(value=f"{self._settings.eq_freq:g}")
        self._eq_width_str = ctk.StringVar(value=f"{self._settings.eq_width:g}")
        self._eq_gain_str = ctk.StringVar(value=f"{self._settings.eq_gain:g}")
        self._eq_expanded = bool(self._settings.eq_expanded)

        # Log buffer powers the collapsible log panel at the bottom. Lines
        # are kept in-memory so the textbox can be re-rendered after a
        # theme switch (CTkTextbox loses its colours otherwise).
        self._log_lines: list[str] = []
        self._log_expanded = False

        self._build_topbar()
        self._build_body()
        self._build_logpanel()
        self._build_bottombar()

        # Trace-based autosave on every relevant StringVar / BooleanVar.
        # Using ``trace_add`` catches changes no matter how they happen
        # (dropdown command, bound Entry, keyboard shortcut, or direct
        # ``.set()`` from code), so the UI never has to remember to
        # call save().
        for var in (
            self._render_preset,
            self._audio_clean_enabled,
            self._eq_enabled,
            self._eq_freq_str,
            self._eq_width_str,
            self._eq_gain_str,
            self._cleanup_mode,
            self._appearance,
        ):
            var.trace_add("write", lambda *_a: self._save_settings())

        # Persist window geometry + expand-states on close, regardless
        # of whether the user clicked X or we got WM_DELETE_WINDOW.
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Settings-seeding is done; from here on, every widget change
        # may trigger a disk write.
        self._settings_silent = False

        # Apply the log-expanded restore AFTER build so the toggle
        # method can resize grid rows correctly. Same for cleanup.
        if self._settings.log_expanded and not self._log_expanded:
            self._toggle_log()
        if self._settings.cleanup_expanded and not self._cleanup_expanded:
            # Flip to collapsed first so _toggle_cleanup_options() opens
            # consistently (it's a no-arg flip of the state flag).
            self._cleanup_expanded = False
            self._toggle_cleanup_options()

        self._set_status(
            f"Idle — drop a video to begin. (settings: {settings_path()})"
        )

    # ------------------------------------------------------------------ build
    def _build_topbar(self) -> None:
        p = self._pal
        self._top = ctk.CTkFrame(self, fg_color=p["panel"], corner_radius=0, height=56)
        self._top.grid(row=0, column=0, sticky="ew")
        self._top.grid_columnconfigure(1, weight=1)

        self._title_label = ctk.CTkLabel(
            self._top,
            text="DaVinci Auto Audioconverter",
            font=FONT_APP_TITLE,
            text_color=p["text"],
            fg_color="transparent",
        )
        self._title_label.grid(row=0, column=0, padx=(18, 10), pady=12, sticky="w")

        # column 1 is the flexible spacer that used to hold the status label;
        # status is now surfaced in the collapsible log panel at the bottom.

        # "Check setup" button — quick read-only health check of FFmpeg,
        # DeepFilterNet, Resolve install paths, Python bitness, etc.
        # Labelled in plain English (not "Preflight", which aviation
        # jargon newcomers won't recognise). Ghost variant so it doesn't
        # visually compete with the primary "Start pipeline" action.
        self._preflight_btn = ctk.CTkButton(
            self._top,
            text="Check setup",
            width=120,
            command=self._on_preflight_clicked,
            **self._button_kw("ghost"),
        )
        self._preflight_btn.grid(row=0, column=2, padx=(6, 6), pady=10, sticky="e")

        self._start_btn = ctk.CTkButton(
            self._top,
            text="Start pipeline",
            width=150,
            command=self._on_start_clicked,
            **self._button_kw("primary"),
        )
        self._start_btn.grid(row=0, column=3, padx=(6, 14), pady=10, sticky="e")

        # Cancel-render button — created now, hidden (grid_remove'd)
        # until a render actually starts so the topbar stays clean
        # during idle / Phase 0-2. Uses a ghost variant with an explicit
        # red-ish tint in _apply_palette so users can't confuse it with
        # the primary action. Sits in the same grid cell as Start; we
        # swap them at phase transitions.
        self._cancel_render_btn = ctk.CTkButton(
            self._top,
            text="Cancel render",
            width=150,
            command=self._on_cancel_render_clicked,
            **self._button_kw("ghost"),
        )
        self._cancel_render_btn.grid(
            row=0, column=3, padx=(6, 14), pady=10, sticky="e"
        )
        self._cancel_render_btn.grid_remove()

    def _build_body(self) -> None:
        p = self._pal
        self._body = ctk.CTkFrame(self, fg_color=p["bg"])
        self._body.grid(row=1, column=0, sticky="nsew", padx=12, pady=(8, 8))
        self._body.grid_columnconfigure(0, weight=1)
        self._body.grid_rowconfigure(0, weight=1)

        self._content = ctk.CTkFrame(
            self._body,
            fg_color=p["panel"],
            corner_radius=10,
            border_width=1,
            border_color=p["border"],
        )
        self._content.grid(row=0, column=0, sticky="nsew")
        self._content.grid_columnconfigure(0, weight=1)
        self._content.grid_rowconfigure(1, weight=1)

        self._section_label = ctk.CTkLabel(
            self._content,
            text="1. Drop a source video",
            font=FONT_SECTION,
            text_color=p["text"],
            fg_color="transparent",
            anchor="w",
        )
        self._section_label.grid(row=0, column=0, sticky="ew", padx=18, pady=(16, 6))

        self._drop_zone = ctk.CTkFrame(
            self._content,
            fg_color=p["panel_elev"],
            corner_radius=10,
            border_width=2,
            border_color=p["border"],
        )
        self._drop_zone.grid(row=1, column=0, sticky="nsew", padx=18, pady=6)
        self._drop_zone.grid_columnconfigure(0, weight=1)
        self._drop_zone.grid_rowconfigure(0, weight=1)

        self._drop_label = ctk.CTkLabel(
            self._drop_zone,
            text="Drop a .mp4 / .mov here\nor click to browse",
            font=FONT_UI,
            text_color=p["muted"],
            fg_color="transparent",
        )
        self._drop_label.grid(row=0, column=0, padx=16, pady=16)
        self._drop_label.bind("<Button-1>", lambda _e: self._browse_file())
        self._drop_zone.bind("<Button-1>", lambda _e: self._browse_file())

        if TkinterDnD is not None and DND_FILES is not None:
            try:
                self._drop_zone.drop_target_register(DND_FILES)
                self._drop_zone.dnd_bind("<<Drop>>", self._on_file_dropped)
                self._drop_label.drop_target_register(DND_FILES)
                self._drop_label.dnd_bind("<<Drop>>", self._on_file_dropped)
            except Exception:
                pass  # fall back to click-to-browse only

        # Audio-cleaning toggle row. Lets the operator skip Phase 0
        # entirely when the source audio is already clean (saves the
        # ~20-60s FFmpeg + DeepFilterNet round trip and avoids the
        # 100 MB model download on the first run). When OFF, the
        # original audio track of the source video is used on A1,
        # matching the pre-preprocessing behaviour of earlier builds.
        self._clean_row = ctk.CTkFrame(self._content, fg_color="transparent")
        self._clean_row.grid(row=2, column=0, sticky="ew", padx=18, pady=(10, 0))
        self._clean_row.grid_columnconfigure(1, weight=1)

        self._clean_switch = ctk.CTkSwitch(
            self._clean_row,
            text="Clean source audio (FFmpeg + DeepFilterNet, Phase 0)",
            variable=self._audio_clean_enabled,
            onvalue=True,
            offvalue=False,
            command=self._on_clean_toggled,
            font=FONT_UI,
            text_color=p["text"],
            progress_color=p["cyan"],
            button_color=p["panel"],
            button_hover_color=p["panel_elev"],
            fg_color=p["panel_elev"],
        )
        self._clean_switch.grid(row=0, column=0, sticky="w")

        self._clean_hint = ctk.CTkLabel(
            self._clean_row,
            text=(
                "Removes background noise before Phase 2 for a cleaner AI "
                "voice. Turn OFF to skip FFmpeg + DeepFilterNet entirely and "
                "use the original audio."
            ),
            font=FONT_HINT,
            text_color=p["muted"],
            fg_color="transparent",
            justify="left",
            anchor="w",
            wraplength=520,
        )
        self._clean_hint.grid(
            row=1, column=0, columnspan=2, sticky="ew", pady=(4, 0)
        )

        # Optional EQ row: sub-toggle + expandable/foldable parameter
        # block. The EQ pass only runs when BOTH this switch and the
        # audio-cleaning switch above are on — an EQ on top of the raw
        # noisy audio would just amplify rumble we didn't want.
        self._eq_row = ctk.CTkFrame(self._content, fg_color="transparent")
        self._eq_row.grid(row=3, column=0, sticky="ew", padx=18, pady=(8, 0))
        self._eq_row.grid_columnconfigure(0, weight=1)

        self._eq_header = ctk.CTkFrame(self._eq_row, fg_color="transparent")
        self._eq_header.grid(row=0, column=0, sticky="ew")
        self._eq_header.grid_columnconfigure(0, weight=1)

        self._eq_switch = ctk.CTkSwitch(
            self._eq_header,
            text="Apply bass-boost EQ after denoise",
            variable=self._eq_enabled,
            onvalue=True,
            offvalue=False,
            command=self._on_eq_toggled,
            font=FONT_UI,
            text_color=p["text"],
            progress_color=p["cyan"],
            button_color=p["panel"],
            button_hover_color=p["panel_elev"],
            fg_color=p["panel_elev"],
        )
        self._eq_switch.grid(row=0, column=0, sticky="w")

        self._eq_toggle_btn = ctk.CTkButton(
            self._eq_header,
            text="Advanced ▾",
            width=110,
            command=self._toggle_eq_options,
            **self._button_kw("ghost"),
        )
        self._eq_toggle_btn.grid(row=0, column=1, sticky="e")

        self._eq_hint = ctk.CTkLabel(
            self._eq_row,
            text=(
                "Adds chest depth to the denoised voice. Speech defaults: "
                f"f={EQ_DEFAULT_FREQ_HZ:g} Hz · Q={EQ_DEFAULT_WIDTH_Q:g} · "
                f"+{EQ_DEFAULT_GAIN_DB:g} dB (optimal range 100–150 Hz, "
                "+2 to +4 dB)."
            ),
            font=FONT_HINT,
            text_color=p["muted"],
            fg_color="transparent",
            justify="left",
            anchor="w",
            wraplength=520,
        )
        self._eq_hint.grid(row=1, column=0, sticky="ew", pady=(4, 0))

        # Expandable parameter sub-frame. Gridded in/out by
        # _toggle_eq_options; starts collapsed so the main form stays
        # compact for users who don't care about the defaults.
        self._eq_options = ctk.CTkFrame(
            self._eq_row,
            fg_color=p["panel_elev"],
            corner_radius=8,
            border_width=1,
            border_color=p["border"],
        )
        self._eq_options.grid_columnconfigure(1, weight=1)

        def _mk_numeric_row(
            row: int,
            label_text: str,
            var: ctk.StringVar,
            unit: str,
            tip: str,
        ) -> ctk.CTkEntry:
            lbl = ctk.CTkLabel(
                self._eq_options,
                text=label_text,
                font=FONT_UI,
                text_color=p["text"],
                fg_color="transparent",
                anchor="w",
                width=120,
            )
            lbl.grid(row=row, column=0, sticky="w", padx=(12, 8), pady=(8 if row == 0 else 4, 4))

            entry = ctk.CTkEntry(
                self._eq_options,
                textvariable=var,
                width=110,
                font=FONT_UI,
                fg_color=p["panel"],
                border_color=p["border"],
                text_color=p["text"],
            )
            entry.grid(row=row, column=1, sticky="w", padx=0, pady=(8 if row == 0 else 4, 4))
            entry.bind("<KeyRelease>", lambda _e: self._refresh_eq_preview())
            entry.bind("<FocusOut>", lambda _e: self._refresh_eq_preview())

            unit_lbl = ctk.CTkLabel(
                self._eq_options,
                text=unit,
                font=FONT_UI_SM,
                text_color=p["muted"],
                fg_color="transparent",
                anchor="w",
            )
            unit_lbl.grid(row=row, column=2, sticky="w", padx=(6, 12))

            hint_lbl = ctk.CTkLabel(
                self._eq_options,
                text=tip,
                font=FONT_HINT,
                text_color=p["muted"],
                fg_color="transparent",
                justify="left",
                anchor="w",
            )
            hint_lbl.grid(row=row, column=3, sticky="w", padx=(12, 12))
            return entry

        self._eq_freq_entry = _mk_numeric_row(
            0, "Frequency", self._eq_freq_str, "Hz",
            "chest/body of the voice — try 100–150 Hz",
        )
        self._eq_width_entry = _mk_numeric_row(
            1, "Width (Q)", self._eq_width_str, "",
            "bell width — narrower = more surgical",
        )
        self._eq_gain_entry = _mk_numeric_row(
            2, "Gain", self._eq_gain_str, "dB",
            "boost amount — optimal +2 to +4 dB",
        )

        self._eq_preview = ctk.CTkLabel(
            self._eq_options,
            text="",  # filled by _refresh_eq_preview()
            font=FONT_UI_SM,
            text_color=p["muted"],
            fg_color="transparent",
            anchor="w",
            justify="left",
            wraplength=560,
        )
        self._eq_preview.grid(
            row=3, column=0, columnspan=4, sticky="ew",
            padx=12, pady=(6, 4),
        )
        self._refresh_eq_preview()

        # Preview button row — runs the Phase-0 chain on a ~10s slice
        # of the dropped video so the user can A/B the denoise + EQ
        # against the source before committing to a full Resolve run.
        # Sits inside the EQ advanced panel because that's the main
        # knob users will iterate on. Disabled until a video is loaded.
        self._preview_btn = ctk.CTkButton(
            self._eq_options,
            text=f"Preview ({int(PREVIEW_DEFAULT_DURATION_S)}s slice)",
            width=170,
            command=self._on_preview_clicked,
            state="disabled",
            **self._button_kw("primary"),
        )
        self._preview_btn.grid(
            row=4, column=0, sticky="w", padx=(12, 8), pady=(4, 10),
        )
        self._preview_hint = ctk.CTkLabel(
            self._eq_options,
            text=(
                "Processes the first 10 s of the dropped video with the "
                "current toggles and opens the result in your default "
                "audio player. Handy for tuning EQ."
            ),
            font=FONT_HINT,
            text_color=p["muted"],
            fg_color="transparent",
            justify="left",
            anchor="w",
            wraplength=430,
        )
        self._preview_hint.grid(
            row=4, column=1, columnspan=3, sticky="ew",
            padx=(0, 12), pady=(4, 10),
        )

        # Preset row: label + editable combobox + "Load from Resolve"
        # button. The combobox is editable so users can type a custom
        # preset name that isn't in the default list; it validates
        # against Resolve at render time via the fallback chain in
        # ResolveController.render().
        self._preset_row = ctk.CTkFrame(self._content, fg_color="transparent")
        self._preset_row.grid(row=4, column=0, sticky="ew", padx=18, pady=(10, 4))
        self._preset_row.grid_columnconfigure(1, weight=1)

        self._preset_label = ctk.CTkLabel(
            self._preset_row,
            text="Render preset",
            font=FONT_UI,
            text_color=p["text"],
            fg_color="transparent",
            anchor="w",
        )
        self._preset_label.grid(row=0, column=0, padx=(0, 10), sticky="w")

        self._preset_combo = ctk.CTkComboBox(
            self._preset_row,
            values=[
                ResolveController.DEFAULT_RENDER_PRESET,
                ResolveController.FALLBACK_RENDER_PRESET,
            ],
            variable=self._render_preset,
            state="normal",  # explicit: free-form typing is allowed
            command=self._on_preset_committed,
            font=FONT_UI,
            dropdown_font=FONT_UI,
            fg_color=p["panel_elev"],
            border_color=p["border"],
            button_color=p["cyan_dim"],
            button_hover_color=p["cyan"],
            text_color=p["text"],
            dropdown_fg_color=p["panel_elev"],
            dropdown_text_color=p["text"],
            dropdown_hover_color=p["border"],
        )
        self._preset_combo.grid(row=0, column=1, sticky="ew")
        # CTkComboBox's internal Entry only pushes typed text into the
        # bound ``variable`` when Enter / Tab fires or focus leaves.
        # That bites the common case where a user types the preset
        # name and clicks "Start pipeline" without pressing Enter
        # first — the StringVar is still on the old value.
        # We bind Enter + FocusOut to explicitly commit, and the
        # pipeline also reads ``self._preset_combo.get()`` directly at
        # start time as a second safety net.
        try:
            internal_entry = self._preset_combo._entry  # type: ignore[attr-defined]
        except Exception:  # pragma: no cover - internal attr guarded
            internal_entry = None
        if internal_entry is not None:
            internal_entry.bind("<Return>",   lambda _e: self._commit_preset_entry())
            internal_entry.bind("<KP_Enter>", lambda _e: self._commit_preset_entry())
            internal_entry.bind("<FocusOut>", lambda _e: self._commit_preset_entry())

        self._preset_load_btn = ctk.CTkButton(
            self._preset_row,
            text="Load from Resolve",
            width=150,
            command=self._load_presets_from_resolve,
            **self._button_kw("ghost"),
        )
        self._preset_load_btn.grid(row=0, column=2, padx=(8, 0))

        # Affordance: the combobox *is* editable by default, but without
        # an explicit hint users assume it's a strict dropdown and try
        # loading from Resolve even when Resolve isn't running yet. The
        # hint line below the combobox spells out that you can type the
        # preset name directly, which is the only way to set it *before*
        # Resolve is up.
        self._preset_hint = ctk.CTkLabel(
            self._preset_row,
            text=(
                "Type the exact preset name (case-sensitive) — or click "
                "'Load from Resolve' once Resolve is running to pick from "
                "the live list."
            ),
            font=FONT_HINT,
            text_color=p["muted"],
            fg_color="transparent",
            justify="left",
            anchor="w",
            wraplength=520,
        )
        self._preset_hint.grid(
            row=1, column=0, columnspan=3, sticky="ew", pady=(4, 0)
        )

        # Post-render cleanup mode — three-way dropdown replacing the
        # earlier single on/off switch so users can pick exactly how
        # aggressive the post-render hygiene should be. See
        # CLEANUP_MODE_LABELS at the top of the UI section for the
        # label ⇄ key mapping; the internal key lives on
        # self._cleanup_mode as a StringVar.
        # --- Post-render cleanup (foldable) ---
        # Mirrors the EQ pattern: a compact always-visible header shows
        # the current mode at a glance; the dropdown + longer hint are
        # tucked into a subframe that the user only opens when they
        # want to change the setting. Keeps the main form short for
        # the 95% of runs where the default is fine.
        self._cleanup_row = ctk.CTkFrame(self._content, fg_color="transparent")
        self._cleanup_row.grid(row=5, column=0, sticky="ew", padx=18, pady=(10, 0))
        self._cleanup_row.grid_columnconfigure(0, weight=1)

        self._cleanup_header = ctk.CTkFrame(
            self._cleanup_row, fg_color="transparent"
        )
        self._cleanup_header.grid(row=0, column=0, sticky="ew")
        self._cleanup_header.grid_columnconfigure(1, weight=1)

        self._cleanup_label = ctk.CTkLabel(
            self._cleanup_header,
            text="Post-render cleanup",
            font=FONT_UI,
            text_color=p["text"],
            fg_color="transparent",
            anchor="w",
        )
        self._cleanup_label.grid(row=0, column=0, padx=(0, 10), sticky="w")

        # Small inline summary of the currently-selected mode so the
        # user sees the state without expanding the section.
        self._cleanup_summary = ctk.CTkLabel(
            self._cleanup_header,
            text=self._format_cleanup_summary(self._cleanup_mode.get()),
            font=FONT_HINT,
            text_color=p["muted"],
            fg_color="transparent",
            anchor="w",
        )
        self._cleanup_summary.grid(row=0, column=1, sticky="w")

        self._cleanup_toggle_btn = ctk.CTkButton(
            self._cleanup_header,
            text="Options ▾",
            width=110,
            command=self._toggle_cleanup_options,
            **self._button_kw("ghost"),
        )
        self._cleanup_toggle_btn.grid(row=0, column=2, sticky="e")

        # Collapsible subframe — grid_forget by default; shown by
        # _toggle_cleanup_options when the user clicks the header button.
        self._cleanup_options = ctk.CTkFrame(
            self._cleanup_row,
            fg_color=p["panel_elev"],
            corner_radius=10,
            border_width=1,
            border_color=p["border"],
        )
        self._cleanup_options.grid_columnconfigure(0, weight=1)

        self._cleanup_menu = ctk.CTkOptionMenu(
            self._cleanup_options,
            values=[label for label, _ in CLEANUP_MODE_LABELS],
            command=self._on_cleanup_mode_changed,
            width=380,
            font=FONT_UI,
            dropdown_font=FONT_UI,
            fg_color=p["panel_elev"],
            button_color=p["cyan_dim"],
            button_hover_color=p["cyan"],
            text_color=p["text"],
            dropdown_fg_color=p["panel_elev"],
            dropdown_text_color=p["text"],
            dropdown_hover_color=p["border"],
        )
        # Seed with the label that matches our StringVar's current key
        # (defaults to "off" on first launch).
        self._cleanup_menu.set(
            CLEANUP_MODE_TO_LABEL.get(self._cleanup_mode.get(), "")
        )
        self._cleanup_menu.grid(row=0, column=0, sticky="w", padx=12, pady=(10, 6))

        # Concise 3-bullet explainer — each option gets exactly one line
        # so the section stays scannable. The on-disk AI-WAV is never
        # deleted; mentioning that once here is enough.
        self._cleanup_hint = ctk.CTkLabel(
            self._cleanup_options,
            text=(
                "Off: nothing is deleted.\n"
                "Temp files: delete extracted/denoised/EQ'd WAVs from %TEMP% "
                "(reclaims several GB for long videos).\n"
                "Temp + Resolve: above, plus removes the clips we added to "
                "the Media Pool. The AI-WAV file on disk and the rendered "
                "output are always kept."
            ),
            font=FONT_HINT,
            text_color=p["muted"],
            fg_color="transparent",
            justify="left",
            anchor="w",
            wraplength=560,
        )
        self._cleanup_hint.grid(
            row=1, column=0, sticky="ew", padx=12, pady=(0, 10)
        )

        hint = (
            "Next steps: 1) import & fresh timeline · 2) you generate the "
            "AI-WAV in Resolve · 3) render is queued and monitored "
            "automatically."
        )
        self._hint_label = ctk.CTkLabel(
            self._content,
            text=hint,
            font=FONT_HINT,
            text_color=p["muted"],
            fg_color="transparent",
            justify="left",
            anchor="w",
        )
        self._hint_label.grid(row=6, column=0, sticky="ew", padx=18, pady=(8, 16))

    def _build_logpanel(self) -> None:
        """Build the collapsible log panel. Hidden by default; toggled via
        the Log button in the bottom bar. Lives in grid row 2 so it slots
        between the body (row 1) and the bottom bar (row 3)."""
        p = self._pal
        self._log_frame = ctk.CTkFrame(
            self,
            fg_color=p["panel"],
            corner_radius=10,
            border_width=1,
            border_color=p["border"],
        )
        # Not gridded here — _toggle_log() manages visibility.
        self._log_frame.grid_columnconfigure(0, weight=1)
        self._log_frame.grid_rowconfigure(0, weight=1)

        self._log_textbox = ctk.CTkTextbox(
            self._log_frame,
            fg_color=p["panel_elev"],
            text_color=p["text"],
            border_color=p["border"],
            border_width=1,
            corner_radius=6,
            font=FONT_UI_SM,
            wrap="word",
        )
        self._log_textbox.grid(row=0, column=0, sticky="nsew", padx=10, pady=(10, 4))
        self._log_textbox.configure(state="disabled")

        self._log_actions = ctk.CTkFrame(self._log_frame, fg_color="transparent")
        self._log_actions.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 10))
        self._log_actions.grid_columnconfigure(0, weight=1)

        self._log_copy_btn = ctk.CTkButton(
            self._log_actions,
            text="Copy",
            width=70,
            command=self._copy_log,
            **self._button_kw("ghost"),
        )
        self._log_copy_btn.grid(row=0, column=1, padx=(0, 6))

        self._log_clear_btn = ctk.CTkButton(
            self._log_actions,
            text="Clear",
            width=70,
            command=self._clear_log,
            **self._button_kw("ghost"),
        )
        self._log_clear_btn.grid(row=0, column=2)

    def _build_bottombar(self) -> None:
        p = self._pal
        self._bottom_bar = ctk.CTkFrame(self, fg_color=p["panel_elev"], corner_radius=0, height=44)
        self._bottom_bar.grid(row=3, column=0, sticky="ew")
        self._bottom_bar.grid_columnconfigure(1, weight=1)

        self._segmented = ctk.CTkSegmentedButton(
            self._bottom_bar,
            values=["dark", "light"],
            variable=self._appearance,
            command=lambda _v: self._on_appearance(),
            font=FONT_UI_SM,
            fg_color=p["panel"],
            selected_color=p["cyan_dim"],
            selected_hover_color=p["cyan"],
            unselected_color=p["panel_elev"],
            unselected_hover_color=p["border"],
            text_color=p["text"],
        )
        self._segmented.grid(row=0, column=0, padx=12, pady=6, sticky="w")

        self._file_label = ctk.CTkLabel(
            self._bottom_bar,
            text="No file selected",
            font=FONT_UI_SM,
            text_color=p["muted"],
            fg_color="transparent",
            anchor="e",
        )
        self._file_label.grid(row=0, column=1, sticky="ew", padx=12)

        # ``▾`` when collapsed, ``▴`` when expanded. The caret flips in
        # _toggle_log so users know the direction of the action.
        self._log_toggle_btn = ctk.CTkButton(
            self._bottom_bar,
            text="Log ▾",
            width=90,
            command=self._toggle_log,
            **self._button_kw("ghost"),
        )
        self._log_toggle_btn.grid(row=0, column=2, padx=(6, 12), pady=6, sticky="e")

    # -------------------------------------------------------------- palette
    def _button_kw(self, variant: str = "ghost", *, height: int = BTN_H,
                   font: tuple | None = None, width: int | None = None) -> dict:
        # Delegate to the shared theme module so every variant stays consistent.
        return button_kwargs(self._pal, variant, height=height, font=font, width=width)

    def _on_appearance(self) -> None:
        mode = (self._appearance.get() or "dark").strip().lower()
        if mode == "light":
            self._pal = dict(PALETTE_LIGHT)
            ctk.set_appearance_mode("light")
        else:
            self._pal = dict(PALETTE_DARK)
            ctk.set_appearance_mode("dark")
        self._apply_palette()

    def _apply_palette(self) -> None:
        """Re-apply every themed colour on every widget from ``self._pal``."""
        p = self._pal

        # Root + containers
        self.configure(fg_color=p["bg"])
        self._top.configure(fg_color=p["panel"])
        self._body.configure(fg_color=p["bg"])
        self._content.configure(fg_color=p["panel"], border_color=p["border"])
        self._drop_zone.configure(fg_color=p["panel_elev"], border_color=p["border"])
        self._bottom_bar.configure(fg_color=p["panel_elev"])
        self._log_frame.configure(fg_color=p["panel"], border_color=p["border"])
        self._log_textbox.configure(
            fg_color=p["panel_elev"],
            text_color=p["text"],
            border_color=p["border"],
        )

        # Static labels
        self._title_label.configure(text_color=p["text"])
        self._section_label.configure(text_color=p["text"])
        self._hint_label.configure(text_color=p["muted"])

        # Labels that adopt the "text" colour only once a file is loaded,
        # and fall back to "muted" placeholder otherwise.
        if self._video_path:
            self._drop_label.configure(text_color=p["text"])
            self._file_label.configure(text_color=p["text"])
        else:
            self._drop_label.configure(text_color=p["muted"])
            self._file_label.configure(text_color=p["muted"])

        # Buttons — rebuild kwargs because every variant references the palette.
        ghost_kw = self._button_kw("ghost")
        primary_kw = self._button_kw("primary")
        for kw in (ghost_kw, primary_kw):
            for key in ("height",):
                kw.pop(key, None)
        # Start uses the regular "primary" variant so its font weight
        # matches every other button in the topbar; the cyan fill is
        # enough to signal the primary action without an emphasis bold.
        self._start_btn.configure(**primary_kw)
        self._log_toggle_btn.configure(**ghost_kw)
        self._log_copy_btn.configure(**ghost_kw)
        self._log_clear_btn.configure(**ghost_kw)
        self._preset_load_btn.configure(**ghost_kw)
        self._preflight_btn.configure(**ghost_kw)
        # Cancel-render retains the ghost base but overrides the border
        # + hover tint towards a warning colour so it's visually
        # distinct from the other ghost actions on the topbar.
        cancel_kw = dict(ghost_kw)
        cancel_kw.update(
            border_color=p.get("stop", p["border"]),
            hover_color=p.get("stop", p["border"]),
            text_color=p["text"],
        )
        self._cancel_render_btn.configure(**cancel_kw)
        self._preview_btn.configure(**primary_kw)

        # Preset picker row.
        self._preset_label.configure(text_color=p["text"])
        self._preset_combo.configure(
            fg_color=p["panel_elev"],
            border_color=p["border"],
            button_color=p["cyan_dim"],
            button_hover_color=p["cyan"],
            text_color=p["text"],
            dropdown_fg_color=p["panel_elev"],
            dropdown_text_color=p["text"],
            dropdown_hover_color=p["border"],
        )
        self._preset_hint.configure(text_color=p["muted"])

        # Audio-cleaning toggle row.
        self._clean_switch.configure(
            text_color=p["text"],
            progress_color=p["cyan"],
            button_color=p["panel"],
            button_hover_color=p["panel_elev"],
            fg_color=p["panel_elev"],
        )
        self._clean_hint.configure(text_color=p["muted"])

        # Post-render cleanup row. The OptionMenu mirrors the preset
        # picker's colour scheme for visual consistency — arrow button
        # on cyan, field on elevated panel, dropdown list on elevated
        # panel. Same theme keys as the Render-preset combobox so the
        # two rows read as a pair of equivalent settings.
        self._cleanup_label.configure(text_color=p["text"])
        self._cleanup_menu.configure(
            fg_color=p["panel_elev"],
            button_color=p["cyan_dim"],
            button_hover_color=p["cyan"],
            text_color=p["text"],
            dropdown_fg_color=p["panel_elev"],
            dropdown_text_color=p["text"],
            dropdown_hover_color=p["border"],
        )
        self._cleanup_hint.configure(text_color=p["muted"])
        self._cleanup_summary.configure(text_color=p["muted"])
        self._cleanup_toggle_btn.configure(**ghost_kw)
        self._cleanup_options.configure(
            fg_color=p["panel_elev"], border_color=p["border"]
        )

        # Preview button hint text colour.
        self._preview_hint.configure(text_color=p["muted"])

        # EQ row: switch, expand button, hint, and (when open) the
        # parameter sub-frame with its entries and preview.
        self._eq_switch.configure(
            text_color=p["text"],
            progress_color=p["cyan"],
            button_color=p["panel"],
            button_hover_color=p["panel_elev"],
            fg_color=p["panel_elev"],
        )
        self._eq_toggle_btn.configure(**ghost_kw)
        self._eq_hint.configure(text_color=p["muted"])
        self._eq_options.configure(
            fg_color=p["panel_elev"], border_color=p["border"]
        )
        for entry in (
            self._eq_freq_entry,
            self._eq_width_entry,
            self._eq_gain_entry,
        ):
            entry.configure(
                fg_color=p["panel"],
                border_color=p["border"],
                text_color=p["text"],
            )
        # The dynamic numeric/unit/hint labels inside the options frame
        # pick up the palette on their next render — re-walk the children
        # so a theme switch refreshes them immediately.
        for child in self._eq_options.winfo_children():
            if isinstance(child, ctk.CTkLabel):
                text = str(child.cget("text") or "")
                # Static-label heuristic: short labels + unit strings use
                # the primary text colour, longer hint strings use muted.
                if len(text) < 16:
                    child.configure(text_color=p["text"])
                else:
                    child.configure(text_color=p["muted"])
        self._eq_preview.configure(text_color=p["muted"])

        # Segmented appearance toggle.
        self._segmented.configure(
            fg_color=p["panel"],
            selected_color=p["cyan_dim"],
            selected_hover_color=p["cyan"],
            unselected_color=p["panel_elev"],
            unselected_hover_color=p["border"],
            text_color=p["text"],
        )

    # ----------------------------------------------------------------- events
    def _browse_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Select source video",
            filetypes=[("Video files", "*.mp4 *.mov *.mkv *.avi *.mxf"), ("All files", "*.*")],
        )
        if path:
            self._set_video_path(path)

    def _on_file_dropped(self, event: Any) -> None:
        # tkinterdnd2 encodes multiple paths as "{a} {b}". Take the first one.
        raw = (event.data or "").strip()
        if raw.startswith("{"):
            raw = raw.split("}")[0].lstrip("{")
        else:
            raw = raw.split(" ")[0]
        if raw:
            self._set_video_path(raw)

    def _set_video_path(self, path: str) -> None:
        self._video_path = path
        p = self._pal
        self._file_label.configure(text=Path(path).name, text_color=p["text"])
        self._drop_label.configure(
            text=f"Ready: {Path(path).name}",
            text_color=p["text"],
        )
        # Preview is only meaningful once a video is loaded; enable
        # the button now so the user can iterate on EQ settings before
        # even starting the full pipeline.
        if hasattr(self, "_preview_btn"):
            self._refresh_preview_btn_state()
        self._set_status(f"Loaded: {Path(path).name}")

    def _refresh_preview_btn_state(self) -> None:
        """Enable / disable the EQ preview button based on the current
        toggles. Preview needs the Phase-0 chain, so it stays disabled
        while audio cleaning is off (there's nothing to listen to that
        would differ from just opening the source video)."""
        if self._preview_running:
            return  # leave "Processing…" label untouched
        ready = (
            self._video_path is not None
            and bool(self._audio_clean_enabled.get())
        )
        self._preview_btn.configure(
            state="normal" if ready else "disabled",
            text=(
                f"Preview ({int(PREVIEW_DEFAULT_DURATION_S)}s slice)"
                if ready or self._video_path is None
                else "Preview (enable Clean)"
            ),
        )

    # ------------------------------------------------------------------
    # Persistence helpers
    # ------------------------------------------------------------------

    def _snapshot_settings(self) -> AppSettings:
        """Build a fresh :class:`AppSettings` from the current UI state.

        Every Tk variable is read and coerced back into its dataclass
        type. EQ numeric fields might be mid-edit text (``"145.0"`` ·
        ``""`` · ``"1x"``); we fall back to the previous value on any
        parse error rather than clobbering persisted settings with
        transient user typos.
        """
        def _f(var: "ctk.StringVar", fallback: float) -> float:
            try:
                return float(var.get())
            except (TypeError, ValueError):
                return fallback

        return AppSettings(
            audio_clean_enabled=bool(self._audio_clean_enabled.get()),
            eq_enabled=bool(self._eq_enabled.get()),
            eq_freq=_f(self._eq_freq_str, self._settings.eq_freq),
            eq_width=_f(self._eq_width_str, self._settings.eq_width),
            eq_gain=_f(self._eq_gain_str, self._settings.eq_gain),
            render_preset=str(self._render_preset.get() or ""),
            cleanup_mode=str(self._cleanup_mode.get() or "off"),
            cleanup_expanded=bool(self._cleanup_expanded),
            appearance=str(self._appearance.get() or "dark"),
            window_geometry=self._settings.window_geometry,
            log_expanded=bool(self._log_expanded),
            eq_expanded=bool(self._eq_expanded),
        )

    def _save_settings(self) -> None:
        """Persist the current UI state. Cheap enough to call on every
        tiny change — ``AppSettings.save`` writes a <1KB JSON atomically.

        Guards:
        * ``_settings_silent``: skip writes while ``__init__`` is still
          seeding variables (avoids a dozen redundant saves on launch).
        * Any save error is swallowed and reported to the log — a
          broken settings file must NEVER break the pipeline.
        """
        if getattr(self, "_settings_silent", False):
            return
        self._settings = self._snapshot_settings()
        self._settings.save(log=lambda m: self._set_status(m))

    def _on_close(self) -> None:
        """WM_DELETE_WINDOW handler — captures final window size/state
        into the settings file so the next launch opens where the user
        left off, then destroys the window."""
        try:
            # Strip any "+x+y" offset so we re-open at a predictable
            # size without pinning to a monitor that might be gone.
            geom = self.geometry()
            if "+" in geom:
                geom = geom.split("+", 1)[0]
            self._settings.window_geometry = geom
            self._settings.log_expanded = bool(self._log_expanded)
            self._settings.eq_expanded = bool(self._eq_expanded)
            self._settings.cleanup_expanded = bool(self._cleanup_expanded)
            self._settings.save()
        except Exception:
            # Closing the app must always succeed even if the disk is
            # full or read-only — never trap the user with a dead
            # window because of a settings write.
            pass
        try:
            self.destroy()
        except Exception:
            pass

    def _set_status(self, text: str) -> None:
        """Append a line to the collapsible log panel.

        Named ``_set_status`` for backwards compat with every call site in
        the pipeline — status messages *are* the log now, there is no
        separate single-line status widget anymore. Thread-safe: marshals
        onto the Tk main thread via ``after(0, …)``.
        """
        stamp = time.strftime("%H:%M:%S")
        line = f"[{stamp}] {text}"
        self._log_lines.append(line)
        self.after(0, self._append_log_line, line)

    def _append_log_line(self, line: str) -> None:
        self._log_textbox.configure(state="normal")
        self._log_textbox.insert("end", line + "\n")
        self._log_textbox.see("end")
        self._log_textbox.configure(state="disabled")

    def _toggle_log(self) -> None:
        """Show or hide the log panel and adjust row weights so the body
        doesn't fight the log for vertical space."""
        self._log_expanded = not self._log_expanded
        if self._log_expanded:
            self._log_frame.grid(row=2, column=0, sticky="nsew", padx=12, pady=(0, 8))
            self._log_toggle_btn.configure(text="Log ▴")
            # Body keeps priority but log gets a guaranteed slice.
            self.grid_rowconfigure(1, weight=2)
            self.grid_rowconfigure(2, weight=1, minsize=180)
        else:
            self._log_frame.grid_forget()
            self._log_toggle_btn.configure(text="Log ▾")
            self.grid_rowconfigure(1, weight=1)
            self.grid_rowconfigure(2, weight=0, minsize=0)

    def _copy_log(self) -> None:
        self.clipboard_clear()
        self.clipboard_append("\n".join(self._log_lines))

    def _clear_log(self) -> None:
        self._log_lines.clear()
        self._log_textbox.configure(state="normal")
        self._log_textbox.delete("1.0", "end")
        self._log_textbox.configure(state="disabled")

    def _expand_log_if_collapsed(self) -> None:
        if not self._log_expanded:
            self._toggle_log()

    def _load_presets_from_resolve(self) -> None:
        """Background-query Resolve for the real render preset list and
        repopulate the combobox. Keeps the user's current selection if
        the name still exists; otherwise defaults to the first entry.

        Runs in a worker thread because ``connect()`` may block for up
        to 90s on a cold Resolve boot and we don't want to freeze the
        UI. COM is initialised on the worker for the same reason as
        ``_run_pipeline``: fusionscript.dll is a COM module.
        """
        if self._preset_loading:
            return
        self._preset_loading = True
        self._preset_load_btn.configure(state="disabled", text="Loading…")

        def _worker() -> None:
            com_initialised = False
            if sys.platform.startswith("win"):
                try:
                    import ctypes

                    if ctypes.windll.ole32.CoInitializeEx(None, 0x0) >= 0:
                        com_initialised = True
                except Exception:
                    pass
            try:
                self._set_status("Loading render presets from Resolve…")
                self._controller.connect(
                    status_callback=lambda msg: self._set_status(f"Presets — {msg}")
                )
                presets = self._controller.list_render_presets()
                if not presets:
                    self._set_status("Resolve returned an empty preset list.")
                    self.after(0, self._expand_log_if_collapsed)
                    return
                self._set_status(f"Loaded {len(presets)} render presets.")
                self.after(0, lambda: self._apply_preset_list(presets))
            except Exception as exc:  # noqa: BLE001
                trace = traceback.format_exc()
                self._set_status(f"Preset load failed: {exc}")
                for tb_line in trace.rstrip().splitlines():
                    self._set_status(tb_line)
                self.after(0, self._expand_log_if_collapsed)
            finally:
                if com_initialised:
                    try:
                        import ctypes

                        ctypes.windll.ole32.CoUninitialize()
                    except Exception:
                        pass
                self.after(0, self._reset_preset_load_btn)

        threading.Thread(target=_worker, daemon=True).start()

    def _apply_preset_list(self, presets: list[str]) -> None:
        """Push a freshly-loaded preset list into the combobox on the
        main thread; keep the current selection if it's still valid."""
        current = (self._render_preset.get() or "").strip()
        self._preset_combo.configure(values=presets)
        if current and current in presets:
            self._render_preset.set(current)
        elif ResolveController.DEFAULT_RENDER_PRESET in presets:
            self._render_preset.set(ResolveController.DEFAULT_RENDER_PRESET)
        else:
            self._render_preset.set(presets[0])

    def _reset_preset_load_btn(self) -> None:
        self._preset_loading = False
        self._preset_load_btn.configure(state="normal", text="Load from Resolve")

    def _commit_preset_entry(self) -> None:
        """Mirror the combobox's current *displayed* text into the bound
        StringVar. Called on Enter / Tab / FocusOut and as a safety net
        at pipeline start — see the comment where the entry bindings
        are installed."""
        try:
            current = self._preset_combo.get()
        except Exception:
            return
        current = (current or "").strip()
        if current and current != (self._render_preset.get() or "").strip():
            self._render_preset.set(current)

    def _on_preset_committed(self, value: str) -> None:
        """Fired by CTkComboBox when the user picks an item from the
        dropdown. We keep the StringVar in sync and log the choice so
        the operator sees their selection landed."""
        value = (value or "").strip()
        if value:
            self._render_preset.set(value)
            self._set_status(f"Render preset: '{value}'")

    def _on_clean_toggled(self) -> None:
        """Log the new Phase-0 state on toggle so the operator sees at
        a glance which audio source will end up on the timeline."""
        if self._audio_clean_enabled.get():
            self._set_status(
                "Audio cleaning ON — Phase 0 will denoise the source audio."
            )
        else:
            self._set_status(
                "Audio cleaning OFF — original video audio goes straight to A1."
            )
        # Preview piggybacks on the Phase-0 chain, so its availability
        # tracks the clean switch exactly.
        if hasattr(self, "_preview_btn"):
            self._refresh_preview_btn_state()

    def _on_cleanup_mode_changed(self, label: str) -> None:
        """Translate the dropdown's human-readable label back to an
        internal mode key and surface the new intent in the log.

        CustomTkinter's OptionMenu passes the selected label string to
        the command; the mode key is what the pipeline actually reads,
        so we mirror the label into the StringVar via the
        ``CLEANUP_LABEL_TO_MODE`` map. An unknown label defaults to
        ``off`` so a misspelled seed value can't silently turn the
        cleanup into "full".
        """
        mode = CLEANUP_LABEL_TO_MODE.get(label, "off")
        self._cleanup_mode.set(mode)
        if hasattr(self, "_cleanup_summary"):
            self._cleanup_summary.configure(
                text=self._format_cleanup_summary(mode)
            )
        if mode == "off":
            self._set_status(
                "Cleanup mode: off — nothing is deleted after the next render."
            )
        elif mode == "temp":
            self._set_status(
                "Cleanup mode: temp files — %TEMP% WAVs will be removed "
                "after the next successful render."
            )
        else:  # "full"
            self._set_status(
                "Cleanup mode: temp + Resolve — also removes our Media "
                "Pool clips (source-audio + AI-WAV) after render. The AI "
                "WAV file on disk is kept."
            )

    def _on_eq_toggled(self) -> None:
        """Log the new EQ state. The EQ pass runs only after a successful
        denoise step, so we remind the user when audio cleaning is off."""
        if self._eq_enabled.get():
            if self._audio_clean_enabled.get():
                self._set_status(
                    "EQ ON — will apply bass-boost after denoise: "
                    + self._current_eq_filter_or_placeholder()
                )
            else:
                self._set_status(
                    "EQ ON but audio cleaning is OFF — EQ will be skipped "
                    "(requires the denoised WAV as input)."
                )
        else:
            self._set_status("EQ OFF — denoised WAV goes to A1 unprocessed.")

    def _toggle_eq_options(self) -> None:
        """Expand / collapse the EQ parameter sub-frame. The button caret
        mirrors the state so users can see the fold direction at a glance."""
        self._eq_expanded = not self._eq_expanded
        if self._eq_expanded:
            self._eq_options.grid(
                row=2, column=0, sticky="ew", padx=0, pady=(8, 0)
            )
            self._eq_toggle_btn.configure(text="Advanced ▴")
            self._refresh_eq_preview()
        else:
            self._eq_options.grid_forget()
            self._eq_toggle_btn.configure(text="Advanced ▾")
        self._save_settings()

    def _format_cleanup_summary(self, mode: str) -> str:
        """One-word tag shown next to the header when the section is
        collapsed, so the user never has to expand the panel just to
        check which mode is currently active."""
        if mode == "temp":
            return "current: temp files"
        if mode == "full":
            return "current: temp + Resolve"
        return "current: off"

    def _toggle_cleanup_options(self) -> None:
        """Expand / collapse the post-render-cleanup options subframe.
        Collapsed by default; the always-visible summary on the header
        keeps the user informed about the active mode."""
        self._cleanup_expanded = not self._cleanup_expanded
        if self._cleanup_expanded:
            self._cleanup_options.grid(
                row=1, column=0, sticky="ew", padx=0, pady=(8, 0)
            )
            self._cleanup_toggle_btn.configure(text="Options ▴")
        else:
            self._cleanup_options.grid_forget()
            self._cleanup_toggle_btn.configure(text="Options ▾")
        self._save_settings()

    def _parse_eq_params(self) -> tuple[float, float, float]:
        """Return ``(freq, width, gain)`` as floats, falling back to the
        module-level defaults for any field that cannot be parsed.

        We deliberately fall back rather than raise on bad input because
        the UI's *"ON"* toggle represents user intent ("I want EQ"); a
        typo shouldn't sabotage the pipeline — it should apply sensible
        defaults and note the substitution in the log.
        """
        def _float_or_default(raw: str, default: float, label: str) -> float:
            raw = (raw or "").strip().replace(",", ".")
            try:
                val = float(raw)
            except ValueError:
                self._set_status(
                    f"EQ — '{label}' value '{raw}' is not numeric, "
                    f"falling back to default {default:g}."
                )
                return default
            if label in ("frequency", "width") and val <= 0:
                self._set_status(
                    f"EQ — '{label}' must be > 0, falling back to "
                    f"default {default:g}."
                )
                return default
            return val

        return (
            _float_or_default(self._eq_freq_str.get(), EQ_DEFAULT_FREQ_HZ, "frequency"),
            _float_or_default(self._eq_width_str.get(), EQ_DEFAULT_WIDTH_Q, "width"),
            _float_or_default(self._eq_gain_str.get(), EQ_DEFAULT_GAIN_DB, "gain"),
        )

    def _current_eq_filter_or_placeholder(self) -> str:
        """Return the live FFmpeg filter string for the current fields,
        or a placeholder if any field is unparseable right now."""
        try:
            freq = float((self._eq_freq_str.get() or "").replace(",", "."))
            width = float((self._eq_width_str.get() or "").replace(",", "."))
            gain = float((self._eq_gain_str.get() or "").replace(",", "."))
        except ValueError:
            return "equalizer=<invalid values>"
        return build_equalizer_filter(freq, width, gain)

    def _refresh_eq_preview(self) -> None:
        """Update the live filter-string label as the user types."""
        if not hasattr(self, "_eq_preview"):
            return
        self._eq_preview.configure(
            text="FFmpeg filter:  " + self._current_eq_filter_or_placeholder()
        )

    # -------------------------------------------------------- preview + cancel
    def _on_preview_clicked(self) -> None:
        """Kick off a short Phase-0 preview on a background thread.

        The preview uses the CURRENT toggle / EQ values so the user can
        fiddle with the knobs, hit Preview, listen, and iterate without
        ever touching Resolve. We serialise with ``_preview_running``
        so a double-click can't race two concurrent FFmpeg runs on the
        same output WAV path.
        """
        if self._preview_running:
            return
        if not self._video_path:
            messagebox.showwarning(
                "No video selected",
                "Drop a video file on the window first, then hit Preview.",
            )
            return
        if not self._audio_clean_enabled.get():
            messagebox.showinfo(
                "Audio cleaning is OFF",
                "Preview runs the Phase 0 chain (denoise + optional EQ). "
                "Enable 'Clean source audio' to use Preview.",
            )
            return

        self._preview_running = True
        self._preview_btn.configure(state="disabled", text="Processing…")
        eq_enabled = bool(self._eq_enabled.get())
        eq_freq, eq_width, eq_gain = self._parse_eq_params()
        video_path = self._video_path

        def _worker() -> None:
            try:
                self._set_status(
                    "Preview — extracting + denoising a 10s slice "
                    "(first run downloads the model, ~30 s)…"
                )
                out_wav = preview_video_audio(
                    video_path,
                    log=lambda msg: self._set_status(f"Preview — {msg}"),
                    apply_eq=eq_enabled,
                    eq_freq=eq_freq,
                    eq_width=eq_width,
                    eq_gain=eq_gain,
                )
                self._set_status(
                    "Preview — ready: " + os.path.basename(out_wav)
                )
                self.after(0, lambda: _open_audio_file(out_wav))
            except AudioPreprocessError as err:
                self._set_status(f"Preview failed: {err}")
                self.after(0, self._expand_log_if_collapsed)
                self.after(0, lambda e=err: messagebox.showerror(
                    "Preview failed",
                    f"{e}",
                ))
            except Exception as err:  # noqa: BLE001 - surface any failure
                trace = traceback.format_exc()
                self._set_status(f"Preview error: {err}")
                for tb_line in trace.rstrip().splitlines():
                    self._set_status(tb_line)
                self.after(0, self._expand_log_if_collapsed)
            finally:
                def _reset() -> None:
                    self._preview_running = False
                    self._refresh_preview_btn_state()
                self.after(0, _reset)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_preflight_clicked(self) -> None:
        """Run the read-only environment health check and dump the
        result into the log, auto-expanding the panel so the user sees
        the outcome immediately. Runs on a background thread because
        ``tasklist`` / PowerShell calls inside
        :func:`run_preflight_diagnostics` can each take ~500ms."""
        self._preflight_btn.configure(state="disabled", text="Checking…")
        self._set_status("Setup check — running environment diagnostics…")

        def _worker() -> None:
            try:
                results = run_preflight_diagnostics()
            except Exception as err:  # noqa: BLE001 - last-resort safety
                self._set_status(f"Setup check crashed: {err}")
                results = []
            # Format each result consistently: 7-char status badge +
            # the component label (left-padded) + free-form detail.
            fails = 0
            warns = 0
            for label, status, detail in results:
                badge = {
                    "OK":   "[  OK  ]",
                    "WARN": "[ WARN ]",
                    "FAIL": "[ FAIL ]",
                }.get(status, f"[ {status:^4} ]")
                if status == "FAIL":
                    fails += 1
                elif status == "WARN":
                    warns += 1
                self._set_status(f"Setup check  {badge} {label:<28} {detail}")
            summary = f"Setup check done — {fails} fail(s), {warns} warning(s)."
            self._set_status(summary)
            self.after(0, self._expand_log_if_collapsed)

            def _reset_btn() -> None:
                self._preflight_btn.configure(
                    state="normal", text="Check setup"
                )
            self.after(0, _reset_btn)

        threading.Thread(target=_worker, daemon=True).start()

    def _on_cancel_render_clicked(self) -> None:
        """Signal the pipeline worker to abort the current render.

        This is the *only* cancellable step — Phase 1 (API calls) and
        Phase 2 (user dialog) are effectively synchronous from the
        UI's perspective, and Phase 0 runs FFmpeg / DeepFilterNet as a
        subprocess we can also interrupt, but the common use case users
        ask for is "my render is taking too long, get me out". So we
        keep the button scoped to Phase 3 and flip it in/out as phases
        transition.
        """
        if not self._render_cancel_event.is_set():
            self._set_status("Cancel requested — stopping current render…")
            self._render_cancel_event.set()
            # Best-effort synchronous stop so Resolve starts tearing
            # down immediately rather than waiting up to 1s for the
            # polling loop's next iteration to notice the event.
            try:
                self._controller.stop_render()
            except Exception:
                pass
            self._cancel_render_btn.configure(
                text="Cancelling…", state="disabled"
            )

    def _show_cancel_render_btn(self) -> None:
        """Swap the idle Start button out and the Cancel-render button
        in. Called as Phase 3 begins its render polling loop."""
        self._start_btn.grid_remove()
        self._cancel_render_btn.configure(text="Cancel render", state="normal")
        self._cancel_render_btn.grid()

    def _hide_cancel_render_btn(self) -> None:
        """Inverse of :meth:`_show_cancel_render_btn`. Called from the
        pipeline's ``finally`` block so the topbar resets regardless of
        success / cancel / failure."""
        self._cancel_render_btn.grid_remove()
        self._start_btn.grid()

    # ------------------------------------------------------- pipeline driver
    def _on_start_clicked(self) -> None:
        if not self._video_path:
            messagebox.showwarning(
                "No video selected",
                "Drop a video file on the window (or click the drop zone) before starting.",
            )
            return
        # Safety net: CTkComboBox only pushes typed text into the
        # bound variable on Enter / FocusOut. If the user typed a
        # preset name and clicked Start without confirming, the
        # StringVar would still hold the previous value. Force a
        # commit here so Phase 3 sees what's actually on screen.
        self._commit_preset_entry()
        self._start_btn.configure(state="disabled")
        self._phase2_event.clear()
        # Fresh cancel event per run — otherwise a previous cancel
        # would still be latched and the new render would abort
        # immediately.
        self._render_cancel_event.clear()
        # Heavy work runs on a worker thread so the UI stays responsive while
        # still being able to pop up a blocking prompt from the main thread.
        threading.Thread(target=self._run_pipeline, daemon=True).start()

    def _run_pipeline(self) -> None:
        video_path = self._video_path
        if video_path is None:  # should not happen, _on_start_clicked guards it
            return

        # Resolve's scripting C-module uses COM internally. On Windows every
        # thread that touches a COM object must first call CoInitializeEx,
        # otherwise ``scriptapp("Resolve")`` silently returns ``None`` from
        # a worker thread even though it works fine from the main thread of
        # a standalone script. We initialise as MTA because we don't run a
        # Windows message pump on this worker. Always paired with
        # CoUninitialize in finally to keep the thread clean on shutdown.
        com_initialised = False
        if sys.platform.startswith("win"):
            try:
                import ctypes

                # COINIT_MULTITHREADED = 0x0; returns S_OK(0) or S_FALSE(1)
                # when already initialised — both are fine. Negative HRESULT
                # means a real failure.
                hr = ctypes.windll.ole32.CoInitializeEx(None, 0x0)
                if hr >= 0:
                    com_initialised = True
            except Exception:
                pass  # best-effort; don't block pipeline on a diag helper

        # Snapshot the Phase-0 toggles at pipeline start so they cannot
        # flip mid-run if the user fidgets with the switches after
        # clicking Start. ``clean_audio_path`` is None when preprocessing
        # is skipped; Phase 1 then uses the source video's own audio.
        audio_clean_enabled = bool(self._audio_clean_enabled.get())
        eq_enabled = bool(self._eq_enabled.get())
        eq_freq, eq_width, eq_gain = self._parse_eq_params()
        clean_audio_path: str | None = None

        try:
            # --- Phase 0 (audio preprocessing) ---------------------------
            # Runs BEFORE any Resolve API call so FFmpeg / DeepFilterNet
            # failures surface before we even try to talk to Resolve.
            # The resulting WAV is the operator's "source" audio for the
            # whole pipeline — the raw (noisy) track from the video
            # never reaches the timeline.
            if audio_clean_enabled:
                stages = "FFmpeg + DeepFilterNet"
                if eq_enabled:
                    stages += " + FFmpeg EQ"
                self._set_status(f"Phase 0/3 — preprocessing audio ({stages})…")
                try:
                    clean_audio_path = preprocess_video_audio(
                        video_path,
                        log=lambda msg: self._set_status(f"Phase 0/3 — {msg}"),
                        apply_eq=eq_enabled,
                        eq_freq=eq_freq,
                        eq_width=eq_width,
                        eq_gain=eq_gain,
                    )
                except AudioPreprocessError as err:
                    # Surface the hint inline; the generic handler below
                    # would still catch this but a dedicated message makes
                    # the source of the failure obvious in the log.
                    self._set_status(
                        "Phase 0/3 — audio preprocessing failed. Fix the tool "
                        "issue below and restart."
                    )
                    raise RuntimeError(
                        f"Audio preprocessing failed: {err}"
                    ) from err
                self._set_status(
                    "Phase 0/3 — clean audio ready: "
                    + os.path.basename(clean_audio_path)
                )
            else:
                self._set_status(
                    "Phase 0/3 — SKIPPED (audio cleaning disabled). Using "
                    "the original video audio on A1."
                )
                if eq_enabled:
                    self._set_status(
                        "Phase 0/3 — EQ also skipped (needs the denoised WAV)."
                    )

            # --- Phase 1 --------------------------------------------------
            self._set_status("Phase 1/3 — connecting to Resolve…")
            self._controller.connect(
                status_callback=lambda msg: self._set_status(f"Phase 1/3 — {msg}"),
            )

            self._set_status("Phase 1/3 — importing video…")
            clip = self._controller.import_video(video_path)
            fps, resolution = ResolveController.probe_clip(clip)

            # Resolve locks the project frame rate as soon as any
            # timeline exists in the project. Leftover auto-generated
            # timelines from previous runs of this tool would silently
            # prevent the new FPS from sticking. We purge only our own
            # (``AutoAudio_*`` prefix) so any timelines the user made
            # themselves are left alone.
            removed = self._controller.cleanup_auto_timelines()
            if removed:
                self._set_status(
                    f"Phase 1/3 — cleared {removed} leftover auto-"
                    "timeline(s) so the project frame rate is unlocked."
                )

            # Apply the source clip's resolution + FPS to the *project*
            # settings BEFORE creating the timeline. CreateEmptyTimeline
            # snapshots project settings at creation time; without this
            # step a new timeline would inherit Resolve's last project
            # defaults (often 1920x1080 @ 24fps) regardless of the clip.
            width, height, applied_fps = (
                self._controller.apply_project_timeline_settings(fps, resolution)
            )
            # Log what Resolve actually echoed back (read-back via
            # GetSetting). If ``applied_fps`` doesn't match what we sent,
            # that's a signal Resolve rejected the change — usually
            # because an existing timeline with incompatible settings
            # locked the project's frame rate.
            if applied_fps != fps:
                self._set_status(
                    f"Phase 1/3 — project set to {width}x{height} "
                    f"@ {applied_fps} fps (requested {fps}, Resolve kept "
                    f"{applied_fps} — check if a prior timeline locks the rate)"
                )
            else:
                self._set_status(
                    f"Phase 1/3 — project set to {width}x{height} "
                    f"@ {applied_fps} fps"
                )

            self._set_status(
                f"Phase 1/3 — creating timeline ({resolution} @ {fps} fps)…"
            )
            timeline_name = f"AutoAudio_{int(time.time())}"
            self._controller.create_fresh_timeline(timeline_name)

            # Rule 5: clear before the first append as well, in case the newly
            # created timeline inherited any default placeholders from the
            # project settings.
            #
            # Two paths depending on the Phase-0 toggle:
            #   * cleaning ON  → V1 gets the source video (video-only)
            #                    and A1 gets the denoised WAV from Phase 0.
            #   * cleaning OFF → the source video is appended with its
            #                    native tracks so the operator's AI voice
            #                    converter works from the raw audio.
            #
            # Phase 3 later rebuilds the timeline with the newly
            # generated voice WAV regardless of which path was taken
            # here, so the final render is identical.
            self._controller.clear_current_timeline()
            if clean_audio_path is not None:
                self._set_status(
                    "Phase 1/3 — importing cleaned WAV into Media Pool…"
                )
                clean_audio_clip = self._controller.import_video(clean_audio_path)
                self._controller.append_video_only(clip)
                self._controller.append_audio_only(clean_audio_clip)
            else:
                self._set_status(
                    "Phase 1/3 — placing source video with original audio on V1/A1…"
                )
                self._controller.append_full_clip(clip)

            # --- Phase 2 (manual bridge) ---------------------------------
            self._set_status("Phase 2/3 — waiting for AI voice WAV…")
            known_wavs = self._controller.snapshot_wav_clips()
            self.after(0, lambda: self._prompt_manual_step(known_wavs))
            # Block the worker thread until the user clicks OK on the prompt.
            self._phase2_event.wait()

            # --- Phase 3 --------------------------------------------------
            self._set_status("Phase 3/3 — locating new WAV…")
            wav_clip = self._controller.newest_wav_since(known_wavs)

            self._set_status("Phase 3/3 — rebuilding timeline…")
            self._controller.clear_current_timeline()
            self._controller.append_video_only(clip)
            self._controller.append_audio_only(wav_clip)

            # Render output lands next to the source video by default.
            source = Path(video_path)
            output_dir = source.parent
            output_name = f"{source.stem}_autoaudio"
            # Read the combobox twice for belt-and-braces: the StringVar
            # is usually correct thanks to _commit_preset_entry(), but
            # the widget.get() call always reflects the exact text on
            # screen even if an internal CTk race skipped the binding.
            try:
                typed = (self._preset_combo.get() or "").strip()
            except Exception:
                typed = ""
            chosen_preset = typed or (self._render_preset.get() or "").strip() or None
            self._set_status(
                f"Phase 3/3 — rendering (preset: {chosen_preset or 'default'})…"
            )
            self._last_output_dir = str(output_dir)
            # Flip the topbar into "cancellable render" mode so the
            # user can abort if the render takes too long. Swapped
            # back in the finally block below.
            self.after(0, self._show_cancel_render_btn)
            completed = self._controller.render(
                str(output_dir), output_name,
                preset_name=chosen_preset,
                cancel_event=self._render_cancel_event,
            )

            if not completed:
                self._set_status(
                    "Render cancelled by user. Partial output may exist in "
                    + str(output_dir)
                )
                self.after(0, self._expand_log_if_collapsed)
                self.after(0, lambda: messagebox.showwarning(
                    "Render cancelled",
                    "The render was aborted at your request.\n\n"
                    f"Partial output (if any) is in:\n{output_dir}",
                ))
                return

            self._set_status("Done — render finished.")

            # Post-render cleanup. Three modes mirrored from the UI
            # dropdown (see CLEANUP_MODE_LABELS).
            #
            #   "off"  — no-op, user wants everything kept.
            #   "temp" — wipe our %TEMP%/davinci_auto_audio_* dir.
            #   "full" — above, plus remove the Media Pool clips we
            #            imported during this run (the cleaned-audio
            #            WAV and the AI WAV). Files on disk for the AI
            #            WAV stay put because it's the user's manual
            #            output — the final render already embeds it.
            #
            # Only runs on a clean success so a failed run leaves the
            # artefacts in place for post-mortem debugging.
            cleanup_mode = (self._cleanup_mode.get() or "off").strip().lower()
            if cleanup_mode in ("temp", "full"):
                try:
                    freed = cleanup_temp_files(
                        log=lambda msg: self._set_status(f"Cleanup — {msg}")
                    )
                    mib = freed / (1024 * 1024)
                    if freed > 0:
                        self._set_status(
                            f"Cleanup — freed {mib:.1f} MB of temp WAVs."
                        )
                except Exception as err:  # noqa: BLE001 - non-fatal
                    self._set_status(f"Cleanup — temp step skipped ({err}).")
            if cleanup_mode == "full":
                try:
                    # Build the set of paths we know we added during
                    # this pipeline run. The cleaned-audio WAV path is
                    # whatever Phase 0 produced (may be None if Phase
                    # 0 was skipped), and the AI WAV's path is read
                    # from the clip we just finished rendering from.
                    mp_paths: set[str] = set()
                    if clean_audio_path:
                        mp_paths.add(clean_audio_path)
                    try:
                        ai_path = wav_clip.GetClipProperty("File Path") or ""
                    except Exception:
                        ai_path = ""
                    if ai_path:
                        mp_paths.add(ai_path)
                    if mp_paths:
                        self._controller.remove_mediapool_clips(
                            mp_paths,
                            log=lambda msg: self._set_status(f"Cleanup — {msg}"),
                        )
                    else:
                        self._set_status(
                            "Cleanup — no Media Pool clips to remove "
                            "(nothing was imported this run)."
                        )
                except Exception as err:  # noqa: BLE001 - non-fatal
                    self._set_status(f"Cleanup — Resolve step skipped ({err}).")

            # Post-render dialog: offer to open the output folder in
            # Explorer instead of forcing the user to dig through the
            # path in the log. ``askyesno`` returns True on Yes.
            def _finish_dialog() -> None:
                open_it = messagebox.askyesno(
                    "DaVinci Auto Audioconverter",
                    f"Render completed.\n\nOutput folder:\n{output_dir}\n\n"
                    "Open the output folder now?",
                )
                if open_it:
                    _open_in_file_manager(str(output_dir))
            self.after(0, _finish_dialog)

        except Exception as exc:  # noqa: BLE001 - surface *any* failure to the UI
            trace = traceback.format_exc()
            self._set_status(f"Error: {exc}")
            for tb_line in trace.rstrip().splitlines():
                self._set_status(tb_line)
            # Auto-expand the log so the user sees what actually went wrong
            # without having to click "Log" first.
            self.after(0, self._expand_log_if_collapsed)
            self.after(0, lambda e=exc, t=trace: messagebox.showerror(
                "Pipeline failed",
                f"{e}\n\n{t}",
            ))
        finally:
            if com_initialised:
                try:
                    import ctypes

                    ctypes.windll.ole32.CoUninitialize()
                except Exception:
                    pass
            # Restore the topbar: Start visible + enabled, Cancel hidden.
            self.after(0, self._hide_cancel_render_btn)
            self.after(0, lambda: self._start_btn.configure(state="normal"))

    def _prompt_manual_step(self, known_wavs: set[str]) -> None:
        """Blocking confirmation dialog for Phase 2.

        Runs on the Tk main thread. Releasing ``_phase2_event`` wakes the
        worker thread so Phase 3 can continue.
        """
        del known_wavs  # passed for symmetry; actual diffing happens in phase 3
        try:
            messagebox.showinfo(
                "Manual step required",
                "Use the Audio Converter in Resolve now to generate the "
                "speech WAV.\n\nClick OK once the finished WAV is in the "
                "Media Pool.",
            )
        finally:
            self._phase2_event.set()


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
