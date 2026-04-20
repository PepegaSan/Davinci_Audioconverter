"""Audio preprocessing for the DaVinci Auto Audioconverter pipeline.

Runs *before* Resolve is touched:

    1. Extract the source video's audio into a WAV file via FFmpeg
       (48 kHz mono — DeepFilterNet's native sample rate, so no
       resampling happens during the denoise step).

    2. Remove background noise with DeepFilterNet. We prefer the
       installed ``deepFilter`` CLI (bundled with the PyPI package,
       always an independent subprocess so a crash cannot take down
       the GUI thread), and fall back to the ``df.enhance`` Python
       API for environments where the CLI isn't on PATH.

    3. (Optional) Apply a single-band parametric EQ via FFmpeg's
       ``equalizer`` filter to add chest depth to the cleaned voice.
       Parameters are user-tunable; sensible defaults for speech
       are f=145 Hz, width(Q)=2.3, gain=+3.5 dB. Skipped entirely
       when ``apply_eq=False`` is passed to
       :func:`preprocess_video_audio`.

    4. Return the final WAV's path so Resolve can import it and
       drop it on the timeline as the A1 replacement for the
       original (noisy) source audio.

The module is deliberately self-contained: it does not import
anything from ``main.py`` or the ``ResolveController`` — the Resolve
integration calls :func:`preprocess_video_audio` once and treats the
returned path as an opaque file on disk. That isolation is what lets
us drop the audio phase in front of the existing pipeline without
touching the DaVinci API code.
"""

from __future__ import annotations

import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Callable, Optional

LogFn = Callable[[str], None]

# Sample rate DeepFilterNet was trained at. Feed it anything else and
# the CLI/model will resample internally, which is wasted work and a
# source of subtle artefacts — far cleaner to resample once at the
# FFmpeg stage.
_DEEPFILTERNET_SR = 48000

# Where we keep intermediate WAVs. One temp dir per process lifetime;
# Windows cleans ``%TEMP%`` between sessions, so no explicit teardown
# is needed. Centralising the directory keeps multiple runs from
# leaking files sibling-to-sibling in the source folder and avoids
# polluting the user's workspace.
_TMP_ROOT: Optional[Path] = None


class AudioPreprocessError(RuntimeError):
    """Raised when extraction or denoising can't complete. Always carries
    a human-readable hint pointing at the likely fix (missing FFmpeg,
    missing DeepFilterNet, permission issue) so it surfaces cleanly in
    the app's log panel."""


# ---------------------------------------------------------------------------
# Tool discovery
# ---------------------------------------------------------------------------


def _no_console_flags() -> int:
    """Hide the ``cmd.exe`` popup that Popen otherwise flashes on Windows."""
    return getattr(subprocess, "CREATE_NO_WINDOW", 0)


def _find_executable(name: str, *extra_candidates: str) -> Optional[str]:
    """Find ``name`` on PATH, then fall back to explicit candidate paths.

    Returns the first hit or ``None``. We look at extra candidates last
    so a user's bundled ``ffmpeg.exe`` in the project folder (first
    candidate by convention) wins over a stale system one only if the
    caller actually puts it first — otherwise PATH is authoritative.
    """
    hit = shutil.which(name)
    if hit:
        return hit
    for cand in extra_candidates:
        if cand and os.path.isfile(cand):
            return cand
    return None


def _resolve_ffmpeg() -> str:
    """Return an absolute path to an ffmpeg binary or raise."""
    project_root = Path(__file__).resolve().parent
    bundled = project_root / ("ffmpeg.exe" if sys.platform.startswith("win") else "ffmpeg")
    exe = _find_executable("ffmpeg", str(bundled))
    if exe is None:
        raise AudioPreprocessError(
            "FFmpeg not found. Install it (``winget install Gyan.FFmpeg`` "
            "on Windows, ``brew install ffmpeg`` on macOS, or your distro's "
            "package) and make sure ``ffmpeg`` is on PATH — or drop "
            "``ffmpeg.exe`` next to ``main.py``."
        )
    return exe


def _bundle_search_roots() -> list[Path]:
    """Return the directories to search for a bundled ``deep-filter``
    binary, in priority order.

    When the app is running from a PyInstaller bundle, two extra paths
    come into play:

    * ``sys._MEIPASS`` — the temp directory PyInstaller extracts a
      **onefile** bundle into. ``--add-binary`` payloads land here.
    * The directory containing ``sys.executable`` — for a **onedir**
      bundle this is where siblings of the main exe live, including
      ``deep-filter.exe`` if it was dropped next to it by the build.

    In dev (not frozen), we fall back to the directory this module
    lives in — i.e. the project root next to ``main.py``.
    """
    roots: list[Path] = []
    if getattr(sys, "frozen", False):
        # Onedir build: the exe and its siblings. Our ``build.bat``
        # moves ``deep-filter.exe`` up to this directory after PyInstaller
        # finishes, so this is the primary hit for a packaged release.
        exe_dir = Path(sys.executable).resolve().parent
        roots.append(exe_dir)
        # PyInstaller >=6 defaults to placing bundled binaries/datas
        # inside a ``_internal\`` subfolder of ``exe_dir``. If a user
        # ran PyInstaller directly (bypassing our post-build move) the
        # binary will still be there — cover that case so the app
        # doesn't spuriously complain about a missing denoiser.
        internal = exe_dir / "_internal"
        if internal.is_dir():
            roots.append(internal)
        # Onefile build: extraction temp dir.
        meipass = getattr(sys, "_MEIPASS", None)
        if meipass:
            roots.append(Path(meipass))
    roots.append(Path(__file__).resolve().parent)
    # Deduplicate while preserving order.
    seen: set[Path] = set()
    out: list[Path] = []
    for r in roots:
        if r in seen:
            continue
        seen.add(r)
        out.append(r)
    return out


def _resolve_deepfilter_cli() -> Optional[str]:
    """Return the path to a DeepFilterNet CLI binary if available, else None.

    Two flavours of CLI exist in the wild — we try both:

    * ``deep-filter[.exe]`` — the **standalone Rust binary** published
      on the DeepFilterNet GitHub releases page. This is the preferred
      route because it has ZERO Python / Rust / Cargo dependencies:
      download the archive, drop ``deep-filter.exe`` next to
      ``main.py``, done. Especially important on Python 3.12+ where
      ``pip install deepfilternet`` often falls back to a source build
      and fails with ``Cargo … is not installed``.
    * ``deepFilter[.exe]`` — the **PyPI package's** CLI entry point
      (``pip install deepfilternet`` installs it into the env's
      ``Scripts`` / ``bin`` dir). Works when the wheel installed
      cleanly.

    Both CLIs accept ``-o OUTPUT_DIR INPUT_WAV`` and write
    ``<stem>_DeepFilterNet3.wav`` into the output dir, so we treat
    them interchangeably.

    Search order:
      1. Directory containing ``sys.executable`` (frozen onedir
         bundle - siblings next to the main exe).
      2. ``sys._MEIPASS`` (frozen onefile - extraction temp dir).
      3. Project root next to ``main.py`` (dev runs).
      4. Active Python env's Scripts/bin dir (PyPI package install).
      5. ``shutil.which`` on PATH.
    """
    candidates: list[str] = []

    for search_root in _bundle_search_roots():
        # 1) Canonical names.
        if sys.platform.startswith("win"):
            candidates.extend(
                str(search_root / n)
                for n in ("deep-filter.exe", "deepFilter.exe")
            )
            for pat in ("deep-filter-*.exe", "deepFilter-*.exe"):
                candidates.extend(
                    str(p) for p in sorted(search_root.glob(pat), reverse=True)
                )
        else:
            candidates.extend(
                str(search_root / n)
                for n in ("deep-filter", "deepFilter")
            )
            for pat in ("deep-filter-*", "deepFilter-*"):
                for p in sorted(search_root.glob(pat), reverse=True):
                    if p.is_file() and os.access(p, os.X_OK):
                        candidates.append(str(p))

    # 2) PyPI package's CLI: lives in the active Python env's bin dir.
    bin_dir = Path(sys.executable).resolve().parent
    if sys.platform.startswith("win"):
        for sub in (bin_dir, bin_dir / "Scripts"):
            if sub.is_dir():
                candidates.extend(
                    str(sub / n)
                    for n in (
                        "deep-filter.exe",
                        "deep-filter.cmd",
                        "deepFilter.exe",
                        "deepFilter.cmd",
                    )
                )
    else:
        candidates.extend(
            str(bin_dir / n) for n in ("deep-filter", "deepFilter")
        )

    # 3) Direct filesystem check on every candidate first — this catches
    #    the versioned filenames above which ``shutil.which`` would
    #    never find (it only resolves command *names* via PATH).
    for cand in candidates:
        if cand and os.path.isfile(cand):
            return cand

    # 4) Fall through to PATH for the canonical command names.
    for name in ("deep-filter", "deepFilter"):
        hit = shutil.which(name)
        if hit:
            return hit
    return None


# ---------------------------------------------------------------------------
# Temp file plumbing
# ---------------------------------------------------------------------------


def _get_tmp_dir() -> Path:
    global _TMP_ROOT
    if _TMP_ROOT is None or not _TMP_ROOT.is_dir():
        _TMP_ROOT = Path(tempfile.mkdtemp(prefix="davinci_auto_audio_"))
    return _TMP_ROOT


def _safe_stem(video_path: str) -> str:
    """Return a filesystem-safe stem based on the source filename so the
    temp WAVs stay traceable back to their source. Collapses anything
    that might trip ffmpeg's output path quoting (spaces, brackets)."""
    stem = Path(video_path).stem or "source"
    keep = "".join(c if c.isalnum() or c in ("_", "-") else "_" for c in stem)
    # Collapse runs of underscores for readability.
    while "__" in keep:
        keep = keep.replace("__", "_")
    return keep.strip("_") or "source"


# ---------------------------------------------------------------------------
# Step 1 — extract audio via FFmpeg
# ---------------------------------------------------------------------------


def extract_audio(
    video_path: str,
    output_wav_path: str,
    *,
    log: Optional[LogFn] = None,
    start_s: float = 0.0,
    duration_s: Optional[float] = None,
) -> str:
    """Extract the first audio stream of ``video_path`` to a 48 kHz mono WAV.

    Mono + 48 kHz matches DeepFilterNet's native rate so the model
    doesn't have to resample. ``-vn`` drops video, ``-y`` forces
    overwrite (the temp dir is ours, so clobbering is safe), and
    ``pcm_s16le`` keeps the file ffmpeg-standard so Resolve has no
    trouble importing it later.

    ``start_s`` / ``duration_s`` let the preview path extract only a
    small slice (``-ss`` / ``-t``) so the whole chain stays snappy
    enough for interactive EQ tuning without processing a 30-minute
    video end-to-end. Default (``duration_s=None``) keeps the original
    behaviour — full file, unchanged.
    """
    if not os.path.isfile(video_path):
        raise AudioPreprocessError(f"Source video does not exist: {video_path}")

    ffmpeg = _resolve_ffmpeg()
    out = Path(output_wav_path)
    out.parent.mkdir(parents=True, exist_ok=True)

    cmd = [
        ffmpeg,
        "-hide_banner",
        "-loglevel", "error",
        "-y",
    ]
    # ``-ss`` before ``-i`` gives us fast keyframe-accurate seeking;
    # ``-t`` caps the duration. Both are optional — skipped entirely
    # when the caller doesn't ask for a slice.
    if start_s and start_s > 0:
        cmd += ["-ss", f"{start_s:.3f}"]
    cmd += ["-i", str(video_path)]
    if duration_s and duration_s > 0:
        cmd += ["-t", f"{duration_s:.3f}"]
    cmd += [
        "-vn",
        "-ac", "1",
        "-ar", str(_DEEPFILTERNET_SR),
        "-c:a", "pcm_s16le",
        str(out),
    ]
    if log:
        slice_note = ""
        if duration_s and duration_s > 0:
            slice_note = f" (slice: {start_s:.1f}s…{start_s + duration_s:.1f}s)"
        log(f"FFmpeg: extracting audio → {out.name}{slice_note}")

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            creationflags=_no_console_flags(),
            timeout=600,
        )
    except (OSError, subprocess.SubprocessError) as err:
        raise AudioPreprocessError(f"FFmpeg failed to start: {err}") from err

    if result.returncode != 0:
        stderr_tail = (result.stderr or "").strip().splitlines()[-5:]
        raise AudioPreprocessError(
            "FFmpeg failed to extract the audio track.\n"
            f"Exit code: {result.returncode}\n"
            "Last lines:\n  " + "\n  ".join(stderr_tail or ["(no stderr)"])
        )
    if not out.is_file() or out.stat().st_size == 0:
        raise AudioPreprocessError(
            "FFmpeg reported success but produced no audio — does the "
            "video actually have an audio stream? Check in a media player "
            "or ``ffprobe``."
        )
    return str(out)


# ---------------------------------------------------------------------------
# Step 2 — denoise via DeepFilterNet
# ---------------------------------------------------------------------------


def _denoise_with_cli(
    cli: str,
    input_wav: str,
    output_wav: str,
    *,
    log: Optional[LogFn] = None,
) -> bool:
    """Run the ``deepFilter`` CLI. Returns True on success.

    The CLI writes ``<stem>_DeepFilterNet3.wav`` into ``--output-dir``.
    We run it in an isolated temp directory so the mangled output name
    can't collide with anything, then move / rename into place.
    """
    work = Path(tempfile.mkdtemp(prefix="df_cli_", dir=str(_get_tmp_dir())))
    cmd = [
        cli,
        "--output-dir", str(work),
        str(input_wav),
    ]
    if log:
        log("DeepFilterNet CLI: denoising (first run downloads the model, "
            "which can take ~30 s)…")
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            creationflags=_no_console_flags(),
            timeout=1800,
        )
    except (OSError, subprocess.SubprocessError) as err:
        if log:
            log(f"DeepFilterNet CLI launch failed ({err!s}); trying Python API…")
        return False
    if result.returncode != 0:
        tail = (result.stderr or result.stdout or "").strip().splitlines()[-6:]
        if log:
            log(
                "DeepFilterNet CLI returned "
                f"{result.returncode}; trying Python API.\n  "
                + "\n  ".join(tail or ["(no output)"])
            )
        return False

    produced = list(work.glob("*.wav"))
    if not produced:
        if log:
            log("DeepFilterNet CLI reported success but produced no file; "
                "trying Python API…")
        return False
    # There's only ever one, but `sorted` keeps things deterministic.
    src = sorted(produced)[-1]
    out = Path(output_wav)
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(src), str(out))
    return True


_STANDALONE_HINT = (
    "No DeepFilterNet backend is available. Pick the EASIEST route:\n"
    "\n"
    "  (A) Standalone binary (recommended — no Python/Rust needed)\n"
    "      1. Open https://github.com/Rikorose/DeepFilterNet/releases\n"
    "      2. Download the asset for your platform, e.g.\n"
    "         ``deep-filter-<version>-x86_64-pc-windows-msvc.tar.gz``\n"
    "      3. Extract and drop ``deep-filter.exe`` next to ``main.py``.\n"
    "      4. Re-run the pipeline.\n"
    "\n"
    "  (B) Python package (needs a pre-built wheel OR Rust toolchain)\n"
    "      ``pip install deepfilternet``  — if that fails with "
    "``Cargo is not installed``, either install Rust from "
    "https://rustup.rs/ and retry, or use option (A) which has no\n"
    "      build step at all."
)


def _denoise_with_python(
    input_wav: str,
    output_wav: str,
    *,
    log: Optional[LogFn] = None,
) -> None:
    """Fallback: run DeepFilterNet via its Python package (``df``)."""
    try:
        # Imports are deferred so the app can still start when the
        # optional dependency isn't installed — the user only finds out
        # when they actually try to preprocess.
        from df.enhance import enhance, init_df, load_audio, save_audio  # type: ignore
    except ImportError as err:
        raise AudioPreprocessError(_STANDALONE_HINT) from err

    if log:
        log("DeepFilterNet (Python): loading model — first run downloads "
            "weights (~30 s)…")
    try:
        model, df_state, _ = init_df()
        sr = df_state.sr()
        audio, _ = load_audio(input_wav, sr=sr)
        if log:
            log("DeepFilterNet (Python): enhancing…")
        enhanced = enhance(model, df_state, audio)
        save_audio(output_wav, enhanced, sr)
    except Exception as err:  # noqa: BLE001 - fold into our own error type
        raise AudioPreprocessError(
            f"DeepFilterNet Python backend failed: {err}"
        ) from err


def denoise_audio(
    input_wav: str,
    output_wav: str,
    *,
    log: Optional[LogFn] = None,
) -> str:
    """Remove background noise from ``input_wav``, writing ``output_wav``.

    Strategy — CLI first, Python API second:

    * The CLI is a separate subprocess, so if the model crashes (OOM,
      CUDA init failure, model checkpoint corruption) it cannot take
      down our GUI thread. Much more robust in a desktop app.
    * The Python API is kept as a fallback so users without
      ``deepFilter`` on PATH can still run the pipeline — at the cost
      of pulling the whole model into this process's memory.
    """
    if not os.path.isfile(input_wav):
        raise AudioPreprocessError(f"Input WAV does not exist: {input_wav}")

    cli = _resolve_deepfilter_cli()
    if cli and _denoise_with_cli(cli, input_wav, output_wav, log=log):
        if log:
            log(f"DeepFilterNet CLI: wrote clean WAV → {Path(output_wav).name}")
        return output_wav

    _denoise_with_python(input_wav, output_wav, log=log)
    if log:
        log(f"DeepFilterNet (Python): wrote clean WAV → {Path(output_wav).name}")
    return output_wav


# ---------------------------------------------------------------------------
# Step 3 — optional parametric EQ via FFmpeg
# ---------------------------------------------------------------------------


# Reasonable defaults for a voice "chest-depth" lift. Kept as module-level
# constants so the UI can import them instead of hard-coding the same
# numbers in two places.
EQ_DEFAULT_FREQ_HZ: float = 145.0
EQ_DEFAULT_WIDTH_Q: float = 2.3
EQ_DEFAULT_GAIN_DB: float = 3.5


def build_equalizer_filter(
    freq: float,
    width: float,
    gain: float,
) -> str:
    """Return the ``-af`` string FFmpeg expects for a one-band parametric EQ.

    Exposed so the UI can render the *exact* filter it will run, which
    makes tuning far less guess-y than fiddling with sliders blind.
    ``%g`` trims trailing zeros so "145.0" becomes "145" — matches how
    FFmpeg itself prints these values in log output.
    """
    return (
        f"equalizer=f={freq:g}:width_type=q:width={width:g}:g={gain:g}"
    )


def apply_equalizer(
    input_wav: str,
    output_wav: str,
    *,
    freq: float = EQ_DEFAULT_FREQ_HZ,
    width: float = EQ_DEFAULT_WIDTH_Q,
    gain: float = EQ_DEFAULT_GAIN_DB,
    log: Optional[LogFn] = None,
) -> str:
    """Apply a single-band parametric EQ to ``input_wav`` via FFmpeg.

    Used after denoising to give voices more chest depth without
    re-introducing the broad-band rumble that DeepFilterNet just
    stripped out. Sensible ranges for speech:

    * ``freq``   100–150 Hz — chest / body of the voice
    * ``width``  1.5 – 3.0  — Q-factor of the bell (narrow = surgical,
                              wide = warm-but-muddy)
    * ``gain``   +2 – +4 dB — audible lift without pumping any
                              downstream compressor / de-esser

    Raises :class:`AudioPreprocessError` on any parameter or ffmpeg
    failure. Output format matches the denoise step (48 kHz mono
    pcm_s16le) so the chain stays lossless.
    """
    if not os.path.isfile(input_wav):
        raise AudioPreprocessError(f"EQ input WAV does not exist: {input_wav}")
    if freq <= 0:
        raise AudioPreprocessError(
            f"EQ frequency must be > 0 Hz (got {freq})."
        )
    if width <= 0:
        raise AudioPreprocessError(
            f"EQ width (Q) must be > 0 (got {width})."
        )
    # Gain can legitimately be negative (cut), so no lower bound check.

    ffmpeg = _resolve_ffmpeg()
    out = Path(output_wav)
    out.parent.mkdir(parents=True, exist_ok=True)

    filter_str = build_equalizer_filter(freq, width, gain)
    cmd = [
        ffmpeg,
        "-hide_banner",
        "-loglevel", "error",
        "-y",
        "-i", str(input_wav),
        "-af", filter_str,
        "-ar", str(_DEEPFILTERNET_SR),
        "-ac", "1",
        "-c:a", "pcm_s16le",
        str(out),
    ]
    if log:
        log(f"FFmpeg EQ: {filter_str}")

    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            creationflags=_no_console_flags(),
            timeout=600,
        )
    except (OSError, subprocess.SubprocessError) as err:
        raise AudioPreprocessError(
            f"FFmpeg EQ step failed to start: {err}"
        ) from err

    if result.returncode != 0:
        stderr_tail = (result.stderr or "").strip().splitlines()[-5:]
        raise AudioPreprocessError(
            "FFmpeg EQ filter failed.\n"
            f"Filter: {filter_str}\n"
            f"Exit code: {result.returncode}\n"
            "Last lines:\n  " + "\n  ".join(stderr_tail or ["(no stderr)"])
        )
    if not out.is_file() or out.stat().st_size == 0:
        raise AudioPreprocessError(
            f"FFmpeg EQ reported success but produced no output: {out}"
        )
    return str(out)


# ---------------------------------------------------------------------------
# High-level entry point
# ---------------------------------------------------------------------------


def preprocess_video_audio(
    video_path: str,
    *,
    log: Optional[LogFn] = None,
    apply_eq: bool = False,
    eq_freq: float = EQ_DEFAULT_FREQ_HZ,
    eq_width: float = EQ_DEFAULT_WIDTH_Q,
    eq_gain: float = EQ_DEFAULT_GAIN_DB,
    start_s: float = 0.0,
    duration_s: Optional[float] = None,
    output_suffix: str = "",
) -> str:
    """Run extract → denoise → (optional) EQ and return the final WAV path.

    The returned WAV lives inside the module-managed temp directory so
    it survives for the entire process lifetime (Resolve imports it
    later and needs it on disk while the project is open). Windows'
    ``%TEMP%`` cleaner removes it between sessions automatically.

    When ``apply_eq`` is False (default), the pipeline stops after the
    denoise step — byte-for-byte identical to the pre-EQ behaviour.
    When True, a third FFmpeg pass applies the parametric EQ and the
    *EQ'd* WAV is what the caller gets. The intermediate denoised file
    stays on disk in the temp dir so the user can A/B compare if they
    ever need to.

    ``start_s`` / ``duration_s`` / ``output_suffix`` are there for the
    preview path: a short slice of the source gets its own namespaced
    WAVs (``<stem>_preview_source.wav`` etc) so preview runs don't
    clobber the "real" full-length WAVs that Resolve may still have
    open. Default leaves the full-pipeline behaviour unchanged.

    Raises
    ------
    AudioPreprocessError
        With a human-readable hint on any missing tool / failed step —
        the caller can surface it directly in the UI log.
    """
    tmp = _get_tmp_dir()
    stem = _safe_stem(video_path)
    suffix = f"_{output_suffix}" if output_suffix else ""
    raw_wav = str(tmp / f"{stem}{suffix}_source.wav")
    clean_wav = str(tmp / f"{stem}{suffix}_clean.wav")

    extract_audio(
        video_path,
        raw_wav,
        log=log,
        start_s=start_s,
        duration_s=duration_s,
    )
    denoise_audio(raw_wav, clean_wav, log=log)

    if not apply_eq:
        return clean_wav

    eq_wav = str(tmp / f"{stem}{suffix}_clean_eq.wav")
    apply_equalizer(
        clean_wav,
        eq_wav,
        freq=eq_freq,
        width=eq_width,
        gain=eq_gain,
        log=log,
    )
    return eq_wav


# ---------------------------------------------------------------------------
# Preview + cleanup helpers (UI-driven)
# ---------------------------------------------------------------------------


# Default slice for the "preview" button — long enough to judge
# denoise quality + EQ colour, short enough that interactive tuning
# doesn't feel like a full pipeline run.
PREVIEW_DEFAULT_DURATION_S: float = 10.0


def preview_video_audio(
    video_path: str,
    *,
    log: Optional[LogFn] = None,
    apply_eq: bool = False,
    eq_freq: float = EQ_DEFAULT_FREQ_HZ,
    eq_width: float = EQ_DEFAULT_WIDTH_Q,
    eq_gain: float = EQ_DEFAULT_GAIN_DB,
    duration_s: float = PREVIEW_DEFAULT_DURATION_S,
    start_s: float = 0.0,
) -> str:
    """Process a short slice of the source so the user can audition it.

    Thin wrapper around :func:`preprocess_video_audio` that namespaces
    output paths with ``preview`` and trims the audio to
    ``duration_s`` seconds. The returned WAV is ready to be handed to
    ``os.startfile`` / ``xdg-open`` so the OS's default audio player
    opens it for A/B listening against the source.
    """
    return preprocess_video_audio(
        video_path,
        log=log,
        apply_eq=apply_eq,
        eq_freq=eq_freq,
        eq_width=eq_width,
        eq_gain=eq_gain,
        start_s=start_s,
        duration_s=duration_s,
        output_suffix="preview",
    )


def _dir_size_bytes(path: Path) -> int:
    """Sum the size of every file under ``path``. Broken symlinks and
    permission-denied entries are silently skipped so this can't raise
    from a best-effort cleanup caller."""
    total = 0
    try:
        for child in path.rglob("*"):
            try:
                if child.is_file():
                    total += child.stat().st_size
            except OSError:
                continue
    except OSError:
        pass
    return total


def cleanup_temp_files(log: Optional[LogFn] = None) -> int:
    """Delete every intermediate WAV the module produced during this
    process's lifetime. Returns the number of bytes reclaimed so the
    UI can surface a nice "freed 42 MB" message.

    Safe to call any number of times — missing directory is a no-op.
    On the next preprocess call :func:`_get_tmp_dir` will transparently
    re-create a fresh temp directory.
    """
    global _TMP_ROOT
    root = _TMP_ROOT
    if root is None or not root.is_dir():
        if log:
            log("Temp cleanup: nothing to remove.")
        return 0
    size = _dir_size_bytes(root)
    try:
        shutil.rmtree(root, ignore_errors=True)
    finally:
        _TMP_ROOT = None
    if log:
        mib = size / (1024 * 1024)
        log(f"Temp cleanup: removed {root} (~{mib:.1f} MB)")
    return size
