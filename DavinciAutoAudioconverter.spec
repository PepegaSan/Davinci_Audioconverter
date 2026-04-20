# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec file for DaVinci Auto Audioconverter.

Build target: a **--onedir** bundle containing:

  * DavinciAutoAudioconverter.exe     — the Tk GUI app
  * deep-filter.exe                   — DeepFilterNet CLI, MIT-licensed,
                                        redistributed under the terms in
                                        third_party/DeepFilterNet/LICENSE.txt
  * third_party/DeepFilterNet/        — upstream licence + README for
                                        the bundled binary
  * _internal/                        — PyInstaller runtime + Python deps

This spec deliberately uses --onedir (not --onefile) because:

  * the app starts in milliseconds instead of seconds (no
    extract-to-temp step on every launch);
  * the user can see at a glance which DeepFilterNet version ships
    with the bundle;
  * upgrading the bundled denoiser is a drag-and-drop swap on
    deep-filter.exe, no rebuild required;
  * it keeps the MIT attribution notice physically next to the binary
    it covers, which is exactly what the licence asks for.

How to invoke:

    pyinstaller DavinciAutoAudioconverter.spec

Run ``build.bat`` from the repo root to install PyInstaller, locate
the newest ``deep-filter-*.exe`` on disk, run this spec, and copy the
licence docs into ``dist/DavinciAutoAudioconverter/``.
"""

from __future__ import annotations

import shutil
import sys
import tempfile
from pathlib import Path

from PyInstaller.utils.hooks import collect_data_files


# ---------------------------------------------------------------------------
# Locate the DeepFilterNet binary to bundle. We accept either the canonical
# ``deep-filter.exe`` name (if the user renamed it) or the original versioned
# filename from the GitHub release, e.g.
# ``deep-filter-0.5.6-x86_64-pc-windows-msvc.exe``. When several are present,
# the newest version (by lexicographic sort, reversed) wins. This matches
# the runtime discovery logic in ``audio_preprocess._resolve_deepfilter_cli``
# so dev-mode behaviour and the bundled behaviour stay identical.
# ---------------------------------------------------------------------------

REPO_ROOT = Path(SPECPATH).resolve()

def _pick_deepfilter_binary() -> Path:
    canonical = REPO_ROOT / "deep-filter.exe"
    if canonical.is_file():
        return canonical
    candidates = sorted(REPO_ROOT.glob("deep-filter-*.exe"), reverse=True)
    if candidates:
        return candidates[0]
    raise SystemExit(
        "[spec] No deep-filter*.exe found next to this spec file.\n"
        "       Download the standalone binary from\n"
        "       https://github.com/Rikorose/DeepFilterNet/releases\n"
        "       and drop it into the repo root before running the build."
    )


_deepfilter_src_raw = _pick_deepfilter_binary()

# PyInstaller's ``binaries = [(source, dest_folder)]`` tuple only controls
# the destination **folder**, never the destination **filename** — the
# bundled file always keeps the source's basename. To guarantee the end
# user's install directory contains a predictable ``deep-filter.exe``
# (rather than ``deep-filter-0.5.6-x86_64-pc-windows-msvc.exe``), we
# stage the binary under the canonical name in a build-time temp dir
# and register THAT path with PyInstaller. Canonical filename makes
# upgrading trivial — drop the new build onto ``deep-filter.exe`` and
# you're done, no runtime logic has to know about the version suffix.
_STAGE_DIR = Path(tempfile.mkdtemp(prefix="davinci_audioconv_stage_"))
_deepfilter_src = _STAGE_DIR / "deep-filter.exe"
shutil.copy2(_deepfilter_src_raw, _deepfilter_src)
print(
    f"[spec] Bundling DeepFilterNet: {_deepfilter_src_raw.name}\n"
    f"       -> staged as {_deepfilter_src.name} in {_STAGE_DIR}"
)

# ---------------------------------------------------------------------------
# Extra data that must travel with the binary (licence + README). Tuple
# format: (source_path_on_build_machine, destination_subdir_in_bundle).
# Keeping the licence next to the binary it covers is the single condition
# imposed by the upstream MIT licence when we redistribute.
# ---------------------------------------------------------------------------

datas = [
    ("third_party/DeepFilterNet/LICENSE.txt", "third_party/DeepFilterNet"),
    ("third_party/DeepFilterNet/README.txt",  "third_party/DeepFilterNet"),
]

# CustomTkinter ships JSON theme files + PNGs under its package dir;
# PyInstaller can't pick those up by import analysis alone.
datas += collect_data_files("customtkinter")

# tkinterdnd2 needs its compiled ``tkdnd`` Tcl extension folder bundled as
# data, otherwise drag-and-drop silently no-ops in the frozen build.
datas += collect_data_files("tkinterdnd2")

# ---------------------------------------------------------------------------
# Binaries tuple format: (source_path, destination_subdir_in_bundle).
# Empty string means "put it right next to the main exe" — exactly what we
# want for deep-filter so runtime discovery can find it as an exe sibling.
# ---------------------------------------------------------------------------

binaries = [
    (str(_deepfilter_src), "."),
]

# ---------------------------------------------------------------------------
# Hidden imports: modules PyInstaller's static analysis can miss.
# ---------------------------------------------------------------------------

hiddenimports = [
    # CustomTkinter's colour detector is pulled in at runtime only.
    "darkdetect",
    # tkinterdnd2 loads Tcl extensions from its package dir lazily.
    "tkinterdnd2",
]

# Modules we explicitly don't need inside the bundle — prunes several MB
# from the output and shortens the start-up scan.
excludes = [
    "matplotlib",
    "numpy.tests",
    "PIL.ImageQt",
    "PyQt5",
    "PyQt6",
    "pytest",
    "tests",
    "unittest",
]


# ---------------------------------------------------------------------------
# Standard PyInstaller pipeline
# ---------------------------------------------------------------------------

block_cipher = None

a = Analysis(
    ["main.py"],
    pathex=[str(REPO_ROOT)],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="DavinciAutoAudioconverter",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,                # UPX often flags false-positives with AV engines
    console=False,            # GUI app — hide the console on launch
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name="DavinciAutoAudioconverter",
)
