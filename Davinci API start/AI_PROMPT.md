# Copy-paste prompt — DaVinci Resolve API connect pattern

Paste everything below the line into a new AI chat when you want the
assistant to wire a Python script against DaVinci Resolve Studio with the
same hardened connect logic captured in this folder. Attach
`davinci_api.py` from this folder as a reference file so the assistant
has the canonical implementation to work against.

---

**Role:** You are adding DaVinci Resolve Studio automation to a Python
project on Windows. Use the connect pattern from the attached
`davinci_api.py` — do **not** invent a new one.

**Rules that MUST be preserved (these are the ones that fail silently and
waste hours every time someone reinvents them):**

1. **Purge then set env vars.** Before importing `DaVinciResolveScript`,
   pop `RESOLVE_SCRIPT_API`, `RESOLVE_SCRIPT_LIB`, and `PYTHONPATH` from
   `os.environ`, then set `RESOLVE_SCRIPT_API` + `RESOLVE_SCRIPT_LIB`
   from the first existing candidate — preferring the edition of the
   currently running `Resolve.exe` (query it via PowerShell
   `(Get-Process Resolve).Path`), not hardcoded to the Free path.
2. **`time.sleep(2)`** between setting the env vars and
   `import DaVinciResolveScript` — fusionscript.dll is file-locked right
   after Resolve starts.
3. **Convert every path to forward slashes** before handing it to the API
   (`ImportMedia`, render `TargetDir`, …). Backslashes fail silently on
   Windows.
4. **Project-open guard** — `GetCurrentProject()` returns `None` on the
   Project Manager screen. Either create a scratch project automatically
   or raise with a clear error.
5. **Clear + `SetCurrentTimeline()`** before every `AppendToTimeline()`.
6. **Match timeline to the source clip — there are THREE sub-rules here
   that all have to be right or the FPS silently doesn't stick:**

   a. Call `apply_project_timeline_settings(project, fps, resolution)`
      BEFORE `CreateEmptyTimeline`. `CreateEmptyTimeline` snapshots
      project settings at creation time; changing them afterwards does
      nothing.

   b. **Pass the FPS as the RAW STRING** you got from
      `clip.GetClipProperty('FPS')`. Do NOT round-trip through
      `float()` and back. `SetSetting('timelineFrameRate', ...)`
      accepts `"25"` and `"29.97"` but silently rejects `"25.0"` and
      `"29.97002997"`. `_normalise_fps()` in `davinci_api.py`
      collapses `25.0 → "25"` only; fractional rates stay verbatim.

   c. **Purge timelines at the old rate FIRST.** Resolve refuses
      frame-rate changes as long as any timeline exists in the project
      at a different rate. Use
      `cleanup_timelines(project, media_pool, name_prefix="YourPrefix_")`
      (never without a prefix unless you're in a scratch project — you'd
      wipe user work). Without this step, the tool's second run silently
      renders at the first run's rate because the leftover timeline
      locks the project.

   d. **Verify via read-back.** `apply_project_timeline_settings` calls
      `GetSetting('timelineFrameRate')` after the write and returns
      what Resolve actually kept. If the returned value differs from
      what you sent, log a warning — the pipeline is about to render
      at the wrong rate otherwise.

   Include a parse fallback (default 1920×1080 @ `"25"`) when
   `GetClipProperty` returns empty.
7. **`DeleteAllRenderJobs()`** before queueing a new render job. Load
   the render preset with a fallback chain
   (`preset_name → "YouTube - 1080p" → "H.264 Master"`) and raise with
   the list of attempted names when none load.
8. **Bounded polling** on `scriptapp("Resolve")` with a timeout
   (90s default), and a bounded render-monitoring loop with its own
   timeout that calls `StopRendering()` on expiry.
9. **COM init on every worker thread.** On Windows, any thread that
   touches a Resolve scripting object must call
   `ctypes.windll.ole32.CoInitializeEx(None, 0)` first and
   `CoUninitialize()` at the end. Without it, `scriptapp('Resolve')`
   silently returns `None` from worker threads (GUI background tasks,
   `threading.Thread`, Tk's `after` callbacks that spin up threads,
   etc.). Use the `scripting_thread()` context manager from
   `davinci_api.py` — wrap every Resolve-touching worker in it.
10. **Do NOT gate on `ProductName` to detect Studio vs Free.** On
    Resolve 21 the PE `ProductName` is `"DaVinci Resolve"` for both
    editions, so this check false-flags Studio. Log it for diagnostics,
    but let `scriptapp()` fail cleanly if it's actually Free.

**Connect strategy (copy this verbatim — it is already in
`davinci_api.py::connect_resolve`):**

```
1. Call scriptapp("Resolve"). On a running Resolve this returns in <1s.
2. If None AND Resolve.exe is NOT in tasklist: Popen Resolve.exe.
3. If None AND Resolve.exe IS in tasklist: DO NOT Popen again
   (second launch wobbles the scripting socket). Just keep polling.
4. Poll scriptapp every 2s, up to 90s total, with heartbeat status
   messages every 8s. After ~4s of failure WITH Resolve running, emit
   an early "External scripting = Local" hint — this is the most common
   cause and worth surfacing before the 18s diagnostic dump.
5. If no project is open after connecting, create a scratch project
   named "<ScriptName>_<unix_ts>" so the pipeline proceeds unattended.
```

**Install-path discovery:** search every edition's install root, not just
the Free one (see `_RESOLVE_MODULE_DIRS`, `_RESOLVE_LIB_CANDIDATES`,
`_RESOLVE_EXE_CANDIDATES` in `davinci_api.py`). Blackmagic's stock loader
only knows the Free path, so Studio / Studio 21 Beta / 21 Beta break on
import without explicit overrides. When `Resolve.exe` is already
running, prefer its own folder's `fusionscript.dll` and the matching
`C:\ProgramData\...\Support\Developer\Scripting\Modules` dir over the
static candidate list — mismatched DLL/Resolve versions break scripting
silently.

**Failure modes to spell out in error messages:** when the connect loop
times out, the error must include the bound paths + running exe +
Python bitness/admin state, then list these five causes in priority
order so the user can act:

1. `Preferences → System → General → External scripting using` is not
   `Local`. Saving alone is not enough; Resolve must be restarted.
2. Worker thread did not initialise COM — wrap the call in
   `with scripting_thread(): ...`.
3. Privilege mismatch — Resolve started as admin while Python runs as
   user (or vice versa). Both must run at the same privilege level.
4. Modal dialog open inside Resolve (unsaved-changes, render-progress,
   auto-save, "Quick Setup" wizard).
5. Resolve is on the Project Manager screen — open a project.

**Process check rules:** `is_resolve_process_running()` is used ONLY to
decide whether to skip `Popen` (second launch destabilises scripting).
Do not use it as a gate on `scriptapp()` — `scriptapp` is cheap on a
running Resolve and should be called first.

**`tasklist` robustness:** read raw bytes with
`subprocess.check_output(..., stderr=DEVNULL, timeout=5,
creationflags=CREATE_NO_WINDOW)` and decode `utf-8` with
`errors="replace"`. Non-English Windows emits OEM codepage output
(cp850/cp437) that Python's implicit cp1252 decoder rejects, returning
`None` and crashing the `in` check.

**Render-preset UX pattern:** users must be able to pick a preset
WITHOUT Resolve being up yet — so a "query Resolve for preset list"
flow alone doesn't cut it. Build it as:

- An **editable** combobox (`CTkComboBox(state="normal", ...)` or
  equivalent). Pre-seed `values=[DEFAULT, FALLBACK]` so the two safe
  names show even before Resolve is running.
- A visible hint label below the combobox that spells out *"Type the
  preset name exactly as it appears in Resolve (case-sensitive), or
  click 'Load from Resolve' once Resolve is running to pick from the
  live list."* Without this hint users assume the combobox is a strict
  dropdown and get stuck when Resolve isn't up.
- A **"Load from Resolve" button** that runs `list_render_presets(project)`
  on a `scripting_thread()`-wrapped worker and updates
  `combo.configure(values=...)` — but preserve any text the user
  already typed (don't clobber).
- Pass the user's choice (typed OR selected) through
  `render_with_preset(preset_name=...)`. The helper's three-step
  fallback chain (`preset_name → "YouTube - 1080p" → "H.264 Master"`)
  catches typos by loading a safe default instead of crashing the
  render.

`list_render_presets(project)` itself already dedupes + sorts the real
names Resolve reports, so just feed the result straight into the
combobox `values=`.

**Task for this codebase:** wire `connect_resolve()` from `davinci_api.py`
into the target script, surface the `status_callback` strings in whatever
UI (CLI stderr, Tk status bar, log file) the project already has. On
every background thread that touches Resolve, wrap the block in
`with scripting_thread(): ...`. Keep `davinci_api.py` untouched — copy
it into the project root instead of editing so future updates stay
trivial.

**Done when:**
- The script launches Resolve automatically if not running.
- Worker threads wrap Resolve calls in `scripting_thread()`.
- A cold start surfaces heartbeat status every ~8 seconds.
- A running-but-unresponsive Resolve shows the preference hint within
  ~4 seconds.
- A stuck Resolve (wrong preference / modal dialog / privilege mismatch
  / no COM / no project) fails with the actionable 5-item diagnostic
  instead of a generic timeout.
- Every path handed to the API goes through `to_forward`.
- Leftover auto-generated timelines are purged with `cleanup_timelines(..., name_prefix=...)` BEFORE changing the frame rate.
- FPS is forwarded as the raw string from `GetClipProperty` (no
  `float()` round-trip), via `apply_project_timeline_settings`, BEFORE
  `CreateEmptyTimeline`.
- `apply_project_timeline_settings` read-back (`GetSetting`) is logged
  so a silent Resolve rejection is visible in the UI.
- The preset combobox is editable (`state="normal"`) with a visible
  typing hint. "Load from Resolve" is an optional convenience, not the
  only way to set the preset.
- Render presets run through `render_with_preset()` with a 3-step
  fallback chain so typos can't crash the render.

---

_End of prompt block._
