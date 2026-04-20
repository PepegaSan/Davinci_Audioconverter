# DaVinci API start — reference kit

Drop-in reference for reliably connecting Python scripts to DaVinci Resolve
Studio on Windows. Copy `davinci_api.py` into a new project and call
`connect_resolve()` (optionally wrapped in `scripting_thread()`) — every
edge case that usually costs hours of debugging is handled centrally.

## Quick start

### CLI / single-threaded script

```python
from davinci_api import connect_resolve, to_forward

resolve, project, media_pool, root_folder = connect_resolve(
    status_callback=print,
    auto_launch=True,
)

media_pool.ImportMedia([to_forward(r"C:\clips\take01.mp4")])
```

### GUI worker thread (CustomTkinter / PySide / any Tk thread)

```python
import threading
from davinci_api import connect_resolve, scripting_thread

def worker():
    with scripting_thread():              # COM init/uninit for this thread
        resolve, project, *_ = connect_resolve()
        print(project.GetName())

threading.Thread(target=worker, daemon=True).start()
```

Without `scripting_thread()`, `scriptapp("Resolve")` returns `None` forever
on any thread that is not the main thread. This is the single nastiest
"works in CLI, breaks in GUI" trap in the Resolve scripting API.

### Full pipeline — render preset dropdown + matching timeline

```python
import time
from davinci_api import (
    connect_resolve, scripting_thread, to_forward,
    cleanup_timelines, list_render_presets,
    apply_project_timeline_settings, render_with_preset,
)

with scripting_thread():
    resolve, project, media_pool, root = connect_resolve()

    clips = media_pool.ImportMedia([to_forward(video_path)])
    clip = clips[0]

    # Resolve silently refuses SetSetting('timelineFrameRate', ...) as long
    # as ANY timeline exists in the project at a different rate. Purge only
    # timelines whose names we own (prefix filter) so user-made timelines
    # stay untouched.
    cleanup_timelines(project, media_pool, name_prefix="AutoRun_")

    # Pass the RAW string Resolve reported via GetClipProperty. The helper
    # normalises "25.0" → "25" internally but keeps fractional rates
    # ("29.97", "23.976") verbatim — that's the only form Resolve reliably
    # accepts.
    width, height, applied_fps = apply_project_timeline_settings(
        project,
        clip.GetClipProperty("FPS") or "25",
        clip.GetClipProperty("Resolution") or "1920x1080",
    )
    print(f"Timeline will be {width}x{height} @ {applied_fps} fps")

    timeline = media_pool.CreateEmptyTimeline(f"AutoRun_{int(time.time())}")
    project.SetCurrentTimeline(timeline)
    media_pool.AppendToTimeline([{"mediaPoolItem": clip}])

    presets = list_render_presets(project)            # feed UI dropdown
    render_with_preset(
        project,
        output_dir=r"C:\output",
        output_name="take01_autoedit",
        preset_name=presets[0] if presets else None,  # falls back to defaults
        status_callback=print,
    )
```

### Render-preset picker UX (GUI pattern)

The naive "populate dropdown from Resolve" flow breaks when Resolve isn't
running yet — and users don't want to launch Resolve just to pick a
preset. Use this pattern instead:

```python
import customtkinter as ctk

preset_var = ctk.StringVar(value="YouTube - 1080p")

combo = ctk.CTkComboBox(
    parent,
    values=["YouTube - 1080p", "H.264 Master"],  # sensible defaults
    variable=preset_var,
    state="normal",                              # MUST be editable, NOT "readonly"
)

# Show users that typing is allowed — without this hint they assume it's
# a strict dropdown and get stuck when Resolve isn't running.
ctk.CTkLabel(
    parent,
    text=(
        "Type the preset name exactly as it appears in Resolve "
        "(case-sensitive) — or click 'Load from Resolve' once Resolve "
        "is running."
    ),
    wraplength=520,
).grid(...)

def load_from_resolve():
    # Spawn a worker thread with scripting_thread() wrapping, call
    # list_render_presets(project), then update combo.configure(values=...).
    ...

ctk.CTkButton(parent, text="Load from Resolve", command=load_from_resolve).grid(...)
```

Why this matters:

- `state="normal"` keeps the combobox editable — the user can type
  preset names before Resolve is up.
- `render_with_preset()` validates the typed name through its fallback
  chain (`preset_name → "YouTube - 1080p" → "H.264 Master"`), so a typo
  lands on a safe default instead of crashing the render.
- The "Load from Resolve" button is the *optional* path for users who
  already have Resolve open and prefer picking from the live list.

## Smoke test

Run the file directly — it connects, lists the current project's render
presets, and exits with status 0 on success:

```powershell
python davinci_api.py
```

## What it solves (and why the naive version breaks)

| Problem | Symptom | Fix in `davinci_api.py` |
|---|---|---|
| Hardcoded Free-only install path | `ImportError("Could not locate module dependencies")` on Studio / Beta | `_RESOLVE_MODULE_DIRS` + `_RESOLVE_LIB_CANDIDATES`, first-match wins |
| Wrong edition's DLL when multiple are installed | Silent `None` on a perfectly configured box | `running_resolve_exe()` drives the DLL/modules choice before the static list |
| Stale env vars from a previous edition | Silent scripting failure after editing Resolve installs | Purge `RESOLVE_SCRIPT_API` / `RESOLVE_SCRIPT_LIB` / `PYTHONPATH` before setting fresh values |
| fusionscript.dll file-locked right after Resolve start | Import succeeds, but later calls return `None` | `time.sleep(2)` before `import DaVinciResolveScript` |
| Backslashes in paths | `ImportMedia` / `TargetDir` silently ignored | `to_forward()` helper |
| **Worker thread calls `scriptapp()` without COM init** | **90s timeout in GUI even though standalone Python works** | **`scripting_thread()` context manager wraps `CoInitializeEx`/`CoUninitialize`** |
| **Resolve started as admin, Python as user (or vice versa)** | **90s timeout, no clear reason** | **Admin state logged and listed as cause #3 in the timeout error** |
| Second `Popen` on running Resolve | Scripting socket wobbles mid-session | `is_resolve_process_running()` guard |
| No project open | `GetCurrentProject()` returns `None` forever on the Project Manager screen | Auto-create a scratch project (opt-out via `create_scratch_project_name=None`) |
| Cold start exceeds naive 12-second wait | `None` return from `scriptapp`, caller bails | Poll for up to 90s with 2s intervals, heartbeat status every 8s |
| Timeout hits, user has no idea why | Generic `"could not connect"` error | Early actionable hint after ~4s + full diagnostic block after 18s, listing 5 causes in priority order |
| Timeline created at project defaults, not clip's resolution / FPS | Wrong-size render, letterboxing, re-conform pain | `apply_project_timeline_settings(project, fps, res)` **before** `CreateEmptyTimeline` |
| **FPS "25" round-tripped to "25.0" via `float`** | **`SetSetting('timelineFrameRate', ...)` silently rejected; new timeline lands at the old rate** | **`_normalise_fps()` keeps raw strings verbatim; collapses `25.0 → "25"` only where safe** |
| **Leftover auto-timeline from a prior run locks the project frame rate** | **Second run of the tool silently renders at the first run's FPS** | **`cleanup_timelines(project, media_pool, name_prefix="AutoRun_")` deletes only name-prefixed ones before `SetSetting`** |
| **`SetSetting` silently ignored, caller never knows** | **Wrong-framerate render, no warning** | **`apply_project_timeline_settings` read-backs via `GetSetting` and returns what Resolve actually kept** |
| Hardcoded render preset breaks on non-YouTube workflows | `LoadRenderPreset` fails, no fallback | `render_with_preset(preset_name=..., fallback_presets=(...))` with 3-step fallback chain |
| No way to populate a preset picker UI | Users type names that don't exist | `list_render_presets(project)` returns the real list |
| **Preset picker forces users to launch Resolve before they can even pick** | **Blocked UX: can't configure the tool offline** | **Combobox with `state="normal"` + hint label: type preset name directly; "Load from Resolve" is optional** |

## Connect strategy (4-step)

```
1. scriptapp("Resolve")            ← cheap on a running Resolve, returns in <1s
2. if None and not running:        ← tasklist check, not a second Popen
       launch Resolve.exe
3. poll scriptapp every 2s up to 90s:
       healthy cold start: 2-3 attempts, ~10-15s wall clock
       after ~4s + Resolve-is-running → "External scripting = Local" hint
       after ~18s of failure → full 5-cause diagnostic block
4. if no project open:
       CreateProject("<scratch>_<ts>")
```

## Common failure modes (timeout diagnostic order)

1. **External scripting not enabled** — `Preferences → System → General →
   External scripting using` must be `Local`. Saving alone is not enough;
   the scripting socket binds at Resolve startup, so Resolve must be
   restarted after you toggle the setting.
2. **Worker thread never called `CoInitializeEx`** — wrap your Resolve
   interaction in `with scripting_thread():` on any non-main thread.
   `scriptapp('Resolve')` returns `None` silently without COM init.
3. **Privilege mismatch** — Resolve was started as admin while Python
   runs as user (or vice versa). Windows isolates the scripting socket
   per privilege level. Fix: run both at the same elevation, or remove
   Resolve.exe's `Properties → Compatibility → Run as administrator`
   flag.
4. **Modal dialog open inside Resolve** — unsaved-changes prompt, render
   progress, auto-save confirmation. These block the scripting server from
   responding. Dismiss before connecting.
5. **No project open** — on the Project Manager screen, `scriptapp('Resolve')`
   returns `None` until a project is loaded. The scratch-project fallback
   avoids this when enabled.

### What we deliberately do NOT gate on

Resolve Free cannot use external scripting, but on **Resolve 21** both
Free and Studio report `ProductName = "DaVinci Resolve"` in the EXE
version resource. Hard-gating on that string would false-flag Studio
users. We log the `ProductName` for info and let `scriptapp` fail
cleanly instead.

## The 8 DaVinci API rules, mapped

| # | Rule | Where it lives |
|---|---|---|
| 1 | Purge stale hardcoded env vars before importing | `bootstrap_resolve_api` |
| 2 | `time.sleep(2)` before the import | `bootstrap_resolve_api` |
| 3 | Forward slashes for every path handed to the API | `to_forward` (used in `render_with_preset`) |
| 4 | Resolve running + project open guard | `connect_resolve` |
| 5 | `SetCurrentTimeline()` + clear before `AppendToTimeline()` | *Caller's responsibility — `connect_resolve` returns the pool so you can call them* |
| 6 | Settle delay after import + FPS/resolution fallback + **FPS format normalisation** + **project frame rate unlock** + **read-back verification** | `apply_project_timeline_settings` (fallback parsing + defaults + `_normalise_fps` + `GetSetting` echo), `cleanup_timelines` |
| 7 | `DeleteAllRenderJobs()` before queueing | `render_with_preset` |
| 8 | Bounded startup polling + render timeout | `_poll_for_scriptapp` (`RESOLVE_STARTUP_TIMEOUT_S`), `render_with_preset(timeout_s=...)` |

The kit now covers rules 1, 2, 3, 4, 6, 7, and 8 directly. Only rule 5
(timeline clear + `SetCurrentTimeline`) stays in caller code because it
is pipeline-specific.

### Why FPS needs its own machinery (rule 6 extended)

Two silent-fail traps Resolve doesn't document:

1. **Format pedantry.** `SetSetting('timelineFrameRate', ...)` accepts
   `"25"` and `"29.97"` but rejects `"25.0"` and `"29.97002997"`. A
   naive `float(fps) → str(...)` round-trip turns the former into the
   latter. `_normalise_fps` keeps the raw string, collapses `25.0 → "25"`
   only, and leaves fractional rates verbatim.
2. **Project-level lock.** The moment *any* timeline exists in the
   project, Resolve refuses frame-rate changes until all timelines at
   the old rate are removed. `cleanup_timelines(..., name_prefix=...)`
   purges only timelines whose names you own so user timelines stay
   untouched.

`apply_project_timeline_settings` read-backs via `GetSetting` after every
write so you can log the actual applied value and warn the user if
Resolve silently kept the old one.

## Copy-paste block for future AI chats

See `AI_PROMPT.md` — paste it together with `davinci_api.py` into a new
assistant thread to bootstrap a new Resolve-automation project with the
same patterns.

## Files

| File | Purpose |
|---|---|
| `davinci_api.py` | The drop-in module. Self-contained. Runs as a smoke test via `python davinci_api.py`. |
| `README.md` | This file. |
| `AI_PROMPT.md` | Copy-paste prompt so future AI chats know which pattern to use without re-explanation. |
