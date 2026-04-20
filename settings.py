"""Persistent user settings for the DaVinci Auto Audioconverter.

A thin dataclass wrapper around a JSON file in the platform-appropriate
per-user config directory. The UI seeds every widget from
:class:`AppSettings` at launch and calls :meth:`AppSettings.save` on
each individual change so a crash (or a hard close) can only ever
lose the change that's currently being edited, never the whole session.

Design:

* One flat JSON file. No migrations, no nested objects. Easy to inspect
  and hand-edit if someone needs to.
* A ``schema_version`` integer lets us bump the shape later — on load
  we clamp unknown fields to the dataclass defaults and keep only the
  recognised ones, so forward- and backward-compat are both trivial.
* Corrupt or missing files are **never** fatal: we log a one-line note
  (if a logger is supplied) and return defaults. Losing persisted
  settings should never prevent the tool from running.

The module is deliberately dependency-free (stdlib only) so it can be
imported from anywhere in the codebase without creating cycles with
``main.py`` / ``audio_preprocess.py``.
"""

from __future__ import annotations

import json
import os
import sys
import tempfile
from dataclasses import asdict, dataclass, field, fields
from pathlib import Path
from typing import Any, Callable, Optional

LogFn = Callable[[str], None]

_SCHEMA_VERSION = 1

# Filename used inside the per-user config dir. Kept short so the full
# path stays comfortably under Windows' 260-char MAX_PATH on the most
# deeply-nested user folders.
_SETTINGS_FILENAME = "settings.json"


def _settings_dir() -> Path:
    """Return the platform-appropriate directory for our config file.

    * Windows → ``%APPDATA%\\DavinciAutoAudioConverter``
    * macOS   → ``~/Library/Application Support/DavinciAutoAudioConverter``
    * Linux   → ``$XDG_CONFIG_HOME/davinci-auto-audioconverter``
                (falls back to ``~/.config/davinci-auto-audioconverter``)

    The directory is **not** created here; ``save()`` creates it on
    first write so a read-only user can still launch the app and run
    with defaults.
    """
    if sys.platform.startswith("win"):
        base = os.environ.get("APPDATA")
        if not base:
            # Ancient or stripped-down Windows without APPDATA — fall
            # back to the user profile so we don't crash.
            base = str(Path.home())
        return Path(base) / "DavinciAutoAudioConverter"
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "DavinciAutoAudioConverter"
    xdg = os.environ.get("XDG_CONFIG_HOME")
    root = Path(xdg) if xdg else Path.home() / ".config"
    return root / "davinci-auto-audioconverter"


def settings_path() -> Path:
    """Full path to the JSON file (file itself may not exist yet)."""
    return _settings_dir() / _SETTINGS_FILENAME


@dataclass
class AppSettings:
    """Every user-tweakable knob that should survive an app restart.

    Defaults here are the authoritative fallback — both for a fresh
    install and for any key that went missing from an older settings
    file. Keep them in sync with the equivalent constants in the UI
    builder so launching with a deleted settings file produces the
    same initial state as before the persistence feature existed.
    """

    # --- Pipeline toggles ---------------------------------------------
    audio_clean_enabled: bool = True
    eq_enabled: bool = True
    eq_freq: float = 145.0
    eq_width: float = 2.3
    eq_gain: float = 3.5

    # --- Render -------------------------------------------------------
    # Empty string means "use the controller's DEFAULT_RENDER_PRESET".
    # Stored as a free-form string so a custom name the user typed is
    # preserved even if Resolve isn't running at startup.
    render_preset: str = ""

    # --- Post-render cleanup ("off" / "temp" / "full") ----------------
    cleanup_mode: str = "off"
    cleanup_expanded: bool = False

    # --- UI shell -----------------------------------------------------
    appearance: str = "dark"
    # Geometry uses the Tk "WxH" form (no +x+y offset so we don't pin
    # the window off-screen after a monitor layout change). Updated
    # from the current window size at app shutdown.
    window_geometry: str = "1200x900"
    log_expanded: bool = False
    eq_expanded: bool = False

    # --- Meta ---------------------------------------------------------
    schema_version: int = field(default=_SCHEMA_VERSION)

    # ------------------------------------------------------------------
    # Persistence helpers
    # ------------------------------------------------------------------

    @classmethod
    def load(cls, *, log: Optional[LogFn] = None) -> "AppSettings":
        """Read the settings file and return an :class:`AppSettings`.

        Missing file, unreadable file, invalid JSON, or a dict that
        doesn't match the dataclass shape all silently fall back to
        defaults. The caller gets a usable object either way.
        """
        path = settings_path()
        if not path.is_file():
            return cls()

        try:
            raw = json.loads(path.read_text(encoding="utf-8"))
        except (OSError, ValueError) as err:
            if log:
                log(f"Settings — could not read {path} ({err}); using defaults.")
            return cls()

        if not isinstance(raw, dict):
            if log:
                log(f"Settings — {path} is not a JSON object; using defaults.")
            return cls()

        known = {f.name for f in fields(cls)}
        filtered = {k: v for k, v in raw.items() if k in known}

        try:
            inst = cls(**filtered)
        except TypeError as err:
            if log:
                log(f"Settings — {path} has incompatible fields ({err}); using defaults.")
            return cls()

        if log:
            log(f"Settings — loaded {len(filtered)} field(s) from {path}")
        return inst

    def save(self, *, log: Optional[LogFn] = None) -> bool:
        """Persist current values to disk. Returns True on success.

        Writes are atomic: we dump to a sibling temp file in the same
        directory, then ``os.replace`` onto the real path. This avoids
        a half-written JSON if the app is killed mid-write, which
        would otherwise lose *all* settings on the next load.
        """
        directory = _settings_dir()
        try:
            directory.mkdir(parents=True, exist_ok=True)
        except OSError as err:
            if log:
                log(f"Settings — could not create {directory} ({err}); not saving.")
            return False

        payload = asdict(self)
        payload["schema_version"] = _SCHEMA_VERSION
        try:
            # NamedTemporaryFile keeps its own fd open on Windows and
            # that prevents the replace(); use mkstemp + manual close
            # instead for cross-platform atomic write semantics.
            fd, tmp_name = tempfile.mkstemp(
                prefix=".settings_", suffix=".tmp", dir=str(directory)
            )
            try:
                with os.fdopen(fd, "w", encoding="utf-8") as fh:
                    json.dump(payload, fh, indent=2, sort_keys=True)
                os.replace(tmp_name, str(settings_path()))
            except Exception:
                # Best-effort cleanup so we don't leave .settings_*.tmp
                # files lying around on failure.
                try:
                    os.unlink(tmp_name)
                except OSError:
                    pass
                raise
        except OSError as err:
            if log:
                log(f"Settings — save failed ({err}); continuing with in-memory state.")
            return False
        return True

    def to_dict(self) -> dict[str, Any]:
        """Return a plain dict copy — useful for diffing or logging."""
        return asdict(self)
