# DaVinci Auto Audioconverter

A small CustomTkinter utility that drives **DaVinci Resolve Studio 21+**
through the Resolve scripting API. It splits the workflow of turning a
source video into a rendered clip with an AI-generated voice track into
four phases and handles the fully-automatic parts around a single manual
pause.

> **Resolve Studio 21 or newer is required** — earlier versions (and the
> free edition) lack the AI Audio Converter panel the pipeline pauses on.
> See [Requirements](#requirements) for the full story.

## Pipeline

0. **Audio preprocessing (automatic, before Resolve is touched)** — the
   source video's audio track is extracted with FFmpeg and denoised with
   DeepFilterNet, optionally followed by a single-band parametric EQ
   pass for chest-depth. Resolve only ever sees the resulting WAV, so
   the operator's AI-voice generation in phase 2 works from noise-free
   input. Toggles:
   - *"Clean source audio"* turns the extract + denoise step on/off.
     When OFF, Phase 0 is skipped entirely and the original audio
     track of the source video is placed on A1 — useful when the input
     is already clean or you just want a faster dry-run.
   - *"Apply bass-boost EQ after denoise"* runs FFmpeg's `equalizer`
     filter on the denoised WAV. Click the **Advanced ▾** button to
     expand Frequency / Width(Q) / Gain entries; the live FFmpeg filter
     string is shown right below them. Optimal ranges for speech:
     **100–150 Hz** on the frequency and **+2 to +4 dB** on the gain.
     The defaults (`f=145 Hz, width=2.3, g=+3.5 dB`) produce the
     filter string `equalizer=f=145:width_type=q:width=2.3:g=3.5`.
     EQ is automatically skipped when audio cleaning is off, since it
     needs the denoised WAV as input.
1. **Preparation (automatic)** — the script connects to Resolve, imports the
   video into the Media Pool, purges leftover auto-timelines so the project
   frame rate is unlocked, applies the clip's FPS + resolution to the
   project, creates a fresh timeline and places the source video on V1 and
   the cleaned WAV from phase 0 on A1.
2. **Manual bridge (pause)** — a blocking dialog appears. Use the **Audio
   Converter** panel inside Resolve to generate the voice WAV. When the WAV
   shows up in the Media Pool, click **OK**.
3. **Finalise + render (automatic)** — the newest WAV is detected, the
   timeline is rebuilt with the original picture on V1 and the generated WAV
   on A1 (both starting at 0), the render queue is purged, the selected
   preset is loaded with a fallback chain and the render job is launched and
   monitored with a hard timeout.

## Requirements

- Windows 10/11 with **DaVinci Resolve Studio 21 or newer**. This is a
  hard requirement, not a suggestion:
  - The free DaVinci Resolve edition does not expose the scripting API at
    all, so Phases 1 and 3 can't run.
  - Resolve **20** *does* speak the scripting API, but it lacks the
    **AI Audio Converter** panel that Phase 2 depends on — the pipeline
    will pause correctly, but you'll have no way to generate the voice
    WAV inside Resolve, so there's nothing to render in Phase 3.
  - **Resolve 21+ Studio** is the only combination that covers all three
    phases end-to-end.
- **Python 3.11** recommended — see [Supported Python versions](#supported-python-versions)
  below. 3.10, 3.12 and 3.13 also work.
- DaVinci Resolve Studio running with a project open before you start the
  script. If no project is open, the tool auto-creates a scratch project.

### Supported Python versions

The three pip requirements (`customtkinter`, `tkinterdnd2`, and the stdlib
Tk that ships with Python) support **Python 3.10 – 3.13** on Windows. We
pick **3.11** as the default because:

- Rock-solid Tk 8.6 (Tk 9 in Python 3.13 still has minor drag-and-drop
  quirks on some Windows builds).
- Pre-built wheels exist for every optional downstream dep (including the
  *optional* PyPI `deepfilternet`, which on 3.12+ typically falls back to
  a source build and demands the Rust toolchain).
- Installed side-by-side with any other Python — the Windows launcher
  `py -3.11` picks it unambiguously.

Grab the latest 3.11.x patch release here:
<https://www.python.org/downloads/windows/>. During installation enable
**"Add Python to PATH"** *and* **"py launcher"**.

`install.bat` auto-detects Python in this order: `py -3.11` → `py -3.12` →
`py -3.10` → `py -3.13` → any `py -3` → `python` on PATH. You can force a
specific version with `install.bat --py 3.11`.

## Setup (do these three steps once)

### 1. Install the Python dependencies

By default the installer uses your current Python (`pip --user`) and does
**not** create a virtual environment:

```powershell
install.bat
```

To isolate the install into `.\.venv`, answer `y` at the prompt or pass
`--venv` to skip it. Combine with `--py` to pin the venv's Python version:

```powershell
install.bat --venv --py 3.11
```

> **Heads-up about `venv`.** Python's `venv` module is **not** a version
> manager — it can only reuse a Python interpreter that is already
> installed on the machine, it does **not** download Python from the
> internet. So:
>
> - If you want the recommended **Python 3.11** sweet spot but don't have
>   it yet, grab the official Windows installer from
>   <https://www.python.org/downloads/windows/> **before** running
>   `install.bat --venv --py 3.11`. Tick both *"Add Python to PATH"* and
>   *"py launcher"* during setup.
> - If you are on a newer Python (**3.12 / 3.13**) and don't feel like
>   installing 3.11, no problem — the app works fine there too. Just make
>   sure to use the **standalone `deep-filter.exe`** binary (see
>   [Setup step 3](#3-install-the-deepfilternet-binary) below); the
>   PyPI `deepfilternet` package is the only thing that actually cares
>   about 3.11, and we don't need it.
>
> `install.bat` will print this same hint if you ask for a Python
> version that isn't installed.

`run.bat` prefers `.venv\Scripts\python.exe` when present, and otherwise
applies the same 3.11→3.12→3.10→3.13 fallback chain as `install.bat`.

### 2. Install FFmpeg

FFmpeg extracts the video's audio stream as a 48 kHz mono WAV (DeepFilterNet's
native rate). It is not a Python package — install it once:

- **Windows**: `winget install Gyan.FFmpeg` (or Chocolatey / Scoop).
- **macOS**: `brew install ffmpeg`.
- **Linux**: `sudo apt install ffmpeg` (or your distro's equivalent).
- **Portable**: drop `ffmpeg.exe` directly next to `main.py`.

Verify with `ffmpeg -version` in a terminal.

### 3. Install the DeepFilterNet binary

DeepFilterNet removes background noise from the extracted audio. The
**recommended** route is the pre-compiled standalone Rust binary — it
has **ZERO** Python / Rust / Cargo dependencies, so the notorious
`deepfilterlib` *"Cargo is not installed"* build failure (common on
Python 3.12+ when no pre-built wheel is available) cannot bite you.

1. Open the releases page:
   **<https://github.com/Rikorose/DeepFilterNet/releases>**
2. Download the asset for your platform, e.g.:
   - Windows: `deep-filter-<version>-x86_64-pc-windows-msvc.exe`
     (or the `.tar.gz` archive)
   - Linux: `deep-filter-<version>-x86_64-unknown-linux-musl.tar.gz`
   - macOS: `deep-filter-<version>-x86_64-apple-darwin.tar.gz`
3. **Drop the binary next to `main.py`** — i.e. directly into this project
   folder. You can **leave the full versioned filename intact** (e.g.
   `deep-filter-0.5.6-x86_64-pc-windows-msvc.exe`); the tool auto-detects
   `deep-filter.exe`, `deepFilter.exe` **and** any
   `deep-filter-*.exe` / `deepFilter-*.exe` in the project root and picks
   the newest version when several are present.
4. **Don't double-click it.** `deep-filter.exe` is a CLI, not a GUI. A
   double-click launches it without arguments, it prints its help text
   and closes immediately — this is normal. Our pipeline calls it as a
   subprocess; no manual launch is ever needed.

On the first Phase 0 run, DeepFilterNet downloads ~100 MB of model
weights (cached under your home dir thereafter). The log panel shows a
*"first run downloads the model"* hint so you know what's happening.

> **Fallbacks.** The app will also accept any `deep-filter` / `deepFilter`
> binary on your `PATH`, the `deepFilter` CLI that the PyPI package
> `deepfilternet` installs into the env's `Scripts/` dir, or the
> `df.enhance` Python API — but the PyPI package only works if a
> pre-built wheel exists for your Python version **or** Rust + Cargo are
> installed. The standalone binary avoids all that.

## Run

Use the launcher — it picks `.\.venv\Scripts\python.exe` when it exists and
falls back to the system Python otherwise:

```powershell
run.bat
```

1. Drop a `.mp4` / `.mov` / `.mkv` file onto the window.
2. (Optional) Click **Check setup** in the top bar to run a read-only
   diagnostics pass: FFmpeg, DeepFilterNet, Resolve install paths,
   Python bitness, admin status. Results print into the log panel with
   `[ OK ] / [ WARN ] / [ FAIL ]` badges so you can spot a missing tool
   before the pipeline even starts.
3. Decide whether to run audio cleaning — the *"Clean source audio"*
   switch controls Phase 0. Leave it ON (default) for the full
   FFmpeg + DeepFilterNet denoise; flip it OFF to skip Phase 0 entirely
   and keep the original video audio on A1.
4. (Optional) With audio cleaning on, click **Preview** under the EQ
   row to process a 10-second slice through the Phase 0 chain and
   open it in the system's default audio player — useful for
   auditioning EQ settings without running the full pipeline.
5. Choose a render preset: **type the name directly into the combobox**
   (it's editable — just click in it and start typing the exact preset
   name as shown in Resolve's Deliver page, including case). Or click
   **Load from Resolve** once Resolve is running to pick from the live
   list. Typed values are committed on Enter / Tab / click-out.
6. (Optional) Expand **Post-render cleanup · Options ▾** to pick what
   the tool should tidy up after a successful render:
   - *Off* (default) — nothing is deleted.
   - *Temp files only* — deletes the extracted/denoised/EQ'd WAVs in
     `%TEMP%` to reclaim several GB per long video.
   - *Temp + Resolve* — above, plus removes the clips this tool added
     to Resolve's Media Pool. The AI-WAV file on disk and the rendered
     output are always kept.
7. Click **Start pipeline**.
8. While Phase 0 runs (FFmpeg extract + DeepFilterNet denoise), watch
   the log panel at the bottom — click the **Log ▾** header to expand.
   If Phase 0 was toggled off, the pipeline jumps straight to Phase 1.
9. When the dialog asks you to, generate the voice WAV inside Resolve
   and click **OK**.
10. The **Start pipeline** button swaps to **Cancel render** during
    Phase 3 — click it to abort a long render job; the tool calls
    Resolve's `StopRendering()` and aborts cleanly.
11. After a successful render, the tool offers to open the output
    folder in Explorer. The output file is saved next to the source
    video.

### Persistent settings

Your toggles (audio cleaning, EQ enabled + values, render preset,
cleanup mode, theme, window size and fold states) are written to
`%APPDATA%\DavinciAutoAudioConverter\settings.json` after every
change, so the next launch starts exactly where you left off. The
settings file is plain JSON — delete it to reset the app to defaults.

> The settings file lives **outside** the project folder (in the
> per-user config dir), so nothing sensitive is ever at risk of
> being committed to git.

## DaVinci API hardening

The controller in `main.py` enforces every rule from the product spec:

| # | Rule | Where |
|---|------|-------|
| 1 | Purge stale hardcoded env vars before importing the API | `_bootstrap_resolve_api` |
| 2 | `time.sleep(2)` before loading `DaVinciResolveScript` | `_bootstrap_resolve_api` |
| 3 | Forward slashes for every path handed to the API | `_to_forward` |
| 4 | Resolve running + project open guard | `ResolveController.connect` |
| 5 | `SetCurrentTimeline()` on creation + clear before `AppendToTimeline()` | `create_fresh_timeline`, `clear_current_timeline` |
| 6 | Settle delay after import + FPS/resolution fallbacks + project-rate unlock + read-back verification | `import_video`, `probe_clip`, `cleanup_auto_timelines`, `apply_project_timeline_settings` |
| 7 | `DeleteAllRenderJobs()` before queueing | `render` |
| 8 | Render preset fallback + timeout-guarded status loop | `render` |

The worker thread that talks to Resolve also initialises COM via
`ctypes.windll.ole32.CoInitializeEx(None, 0)` and releases it in the
matching `finally` — without this, `scriptapp('Resolve')` silently returns
`None` on Windows GUI threads. See `Davinci API start/` for a standalone
reference implementation of the same pattern.

## Styling

All GUI styling lives in `theme.py` at the repo root (palette, typography and
`button_kwargs` helper for CustomTkinter button variants). Toggle dark/light
with the segmented button in the bottom bar.

The `design_kit/` folder is only a local reference copy and is ignored by
git — it is not part of the shipped project.

## Building a standalone bundle (no Python needed on the target PC)

For shipping to a machine that does not have Python installed, use the
included PyInstaller recipe. It builds a **one-folder** (`--onedir`)
bundle containing the app, the CustomTkinter assets, the DeepFilterNet
CLI and its MIT licence file — zip it up and hand the folder to a user.

```powershell
build.bat             :: build using the currently selected Python
build.bat --clean     :: wipe build\ and dist\ first, then build
```

The builder will:

1. Locate the most recent `deep-filter-*.exe` in the repo root (you
   must have downloaded it per **Setup step 3** above — the build
   refuses to run without it).
2. Install PyInstaller into the active Python if it isn't there yet.
3. Run `pyinstaller DavinciAutoAudioconverter.spec`.
4. Verify that both `DavinciAutoAudioconverter.exe` and
   `deep-filter.exe`, plus the DeepFilterNet MIT licence, actually
   landed in `dist\DavinciAutoAudioconverter\`.

Resulting bundle layout:

```
dist\DavinciAutoAudioconverter\
├── DavinciAutoAudioconverter.exe           ← launch this
├── deep-filter.exe                         ← DeepFilterNet CLI, MIT
├── third_party\
│   └── DeepFilterNet\
│       ├── LICENSE.txt                     ← MIT notice (must travel
│       └── README.txt                         with the binary)
└── _internal\                              ← Python runtime + deps
```

### Licensing when redistributing the bundle

`deep-filter.exe` is DeepFilterNet, released by its author under the
MIT licence (alternatively Apache-2.0; we pick MIT for its brevity).
The MIT licence is a permissive one — you may redistribute the
binary inside your own bundle, commercial use included, **as long
as you keep the copyright notice with it**. That is why the builder
copies `third_party/DeepFilterNet/LICENSE.txt` next to the binary
and the post-build check refuses to succeed if the file is missing.

When you ship `dist\DavinciAutoAudioconverter\` further:

- **Keep the `third_party\DeepFilterNet\` folder** next to
  `deep-filter.exe`. It contains the MIT notice and a brief README
  pointing at the upstream project. Removing it would violate the
  licence.
- **FFmpeg is *not* bundled.** FFmpeg's GPL/LGPL terms make
  in-bundle redistribution a significant compliance exercise (source
  availability, build instructions, possibly LGPL vs. GPL build
  variants). Far simpler to let the end user install FFmpeg
  themselves with `winget install Gyan.FFmpeg`, or to drop
  `ffmpeg.exe` into the bundle folder alongside (in which case you
  become responsible for the FFmpeg licence obligations yourself).
- **Model weights** (~100 MB, DeepFilterNet's neural net) are *not*
  in the bundle. The binary downloads them on its first run and
  caches them under `%USERPROFILE%\.cache\DeepFilterNet\`. Subsequent
  runs are fully offline. For a truly offline bundle you would need
  to pre-populate that cache dir on the target machine.

Upgrading the bundled DeepFilterNet binary later is a straight
drag-and-drop onto `deep-filter.exe` — no rebuild required.

## Repository layout

| Path | Purpose |
|---|---|
| `main.py` | CustomTkinter UI + ResolveController + pipeline orchestration. |
| `audio_preprocess.py` | Phase 0 helper: FFmpeg extract → DeepFilterNet denoise (+ optional EQ). Independent of Resolve. |
| `settings.py` | Persistent user settings (`%APPDATA%\DavinciAutoAudioConverter\settings.json`). |
| `theme.py` | Palette, fonts, button style helper. |
| `diag_scripting.py` | Standalone diagnostic script — checks whether Resolve's scripting API responds outside of the GUI. |
| `install.bat` / `run.bat` | Windows launchers (Python version detection + optional venv). |
| `build.bat` / `DavinciAutoAudioconverter.spec` | PyInstaller recipe for a redistributable `--onedir` bundle. |
| `requirements.txt` | Python deps only (FFmpeg + DeepFilterNet are external). |
| `Davinci API start/` | Standalone reference kit for the Resolve connect pattern — copy into new projects. |
| `third_party/DeepFilterNet/` | Upstream MIT licence + README shipped next to the bundled `deep-filter.exe`. |
| `design_kit/` | Local CustomTkinter design reference, gitignored. |

## Troubleshooting

First line of defence: click **Check setup** in the top bar and read
the `[ OK / WARN / FAIL ]` lines in the log. It reports on FFmpeg,
DeepFilterNet, Resolve install paths, Python bitness and privilege
level — the vast majority of setup issues announce themselves there.

For anything that survives the check:

| Symptom | Fix |
|---|---|
| Phase 2 dialog shows up but there's no **AI Audio Converter** panel in Resolve | You're on Resolve 20 or the free edition. The Audio Converter only exists in **Resolve Studio 21+**. Upgrade Resolve. |
| `install.bat` can't find a Python | Install Python 3.11 from <https://www.python.org/downloads/windows/> with both "Add to PATH" and "py launcher" enabled. Then re-run `install.bat` (optionally with `--py 3.11`). |
| `Cargo … is not installed` during `pip install deepfilternet` | Use the standalone `deep-filter.exe` from GitHub Releases (see **Setup step 3**). Zero compiler setup. |
| Phase 0 never finds DeepFilterNet | Drop `deep-filter.exe` (or the versioned `deep-filter-*.exe`) next to `main.py`, or add its folder to PATH. |
| `FFmpeg not found` | `winget install Gyan.FFmpeg`, or drop `ffmpeg.exe` next to `main.py`. |
| Resolve connects but project FPS doesn't match the clip | A prior timeline is locking the rate. Our tool's leftover `AutoAudio_*` timelines are purged automatically; user-made timelines aren't touched. Delete them in Resolve if needed, or the log will show `Resolve kept <old rate>`. |
| `scriptapp('Resolve')` keeps returning `None` | Open `Preferences → System → General → External scripting using` = `Local`, then **restart Resolve**. The preference socket binds at startup. |
| Tool runs but Resolve ignores it | Privilege mismatch — Resolve started as admin while the tool runs as user (or vice versa). Run both at the same elevation. |
| Settings don't persist between launches | Check that the `%APPDATA%\DavinciAutoAudioConverter\` folder is writeable — on some locked-down corporate profiles `%APPDATA%` is read-only or roamed late. Errors are logged but never fatal; the app just falls back to defaults for that session. |
| Render preset shown in the tool doesn't exist in Resolve | The fallback chain (`your preset → "YouTube - 1080p" → "H.264 Master"`) kicks in automatically and the log reports which one was actually loaded. Fix the typo or pick from **Load from Resolve** next time. |

For a deeper connectivity check (Resolve API from the command line,
no GUI involved), run `py -3 diag_scripting.py`.

## Contributing / license

This repository is published as-is. See individual third-party
projects for their own licences:

- **FFmpeg** — LGPL/GPL (depending on the build you install).
- **DeepFilterNet** — check <https://github.com/Rikorose/DeepFilterNet>
  for the current licence of the pre-compiled binary you download.
- **DaVinci Resolve Scripting API** — Blackmagic Design; distributed
  as part of DaVinci Resolve Studio.

None of those binaries are committed to this repo; you install them
yourself via the steps in **Setup**.
