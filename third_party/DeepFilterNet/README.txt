DeepFilterNet - standalone Rust binary redistributed with this project
======================================================================

This bundle ships the `deep-filter` executable produced by the
DeepFilterNet project. DeepFilterNet is a deep-learning-based speech
enhancement (denoising) library and CLI tool.

  Upstream project : https://github.com/Rikorose/DeepFilterNet
  License          : Apache-2.0 OR MIT (dual-licensed, at your option)
                     We redistribute under the MIT terms; see
                     LICENSE.txt next to this file.
  Copyright        : (c) 2021 Hendrik Schröter and contributors

What the binary does
--------------------
`deep-filter` reads a WAV file, runs it through the DeepFilterNet3
model to remove background noise, and writes a cleaned WAV back out.
It is invoked as a subprocess by `audio_preprocess.py` during
Phase 0 of the pipeline. No network access is performed except on
the **first run**, where the model weights (~100 MB) are downloaded
and cached into the current user's home directory under
`~/.cache/DeepFilterNet/`. Subsequent runs are fully offline.

Why we bundle it
----------------
Shipping the binary alongside the application removes two common
installation pitfalls:

  * the PyPI `deepfilternet` package tries to build a Rust component
    from source when no pre-built wheel exists for the user's Python
    version, which fails without Rust/Cargo installed;
  * asking every end user to manually download a GitHub Release
    asset and drop it next to the executable is a friction point
    that trips non-technical operators.

Upgrading / replacing the bundled binary
----------------------------------------
The file `deep-filter.exe` (or the original versioned name
`deep-filter-<ver>-x86_64-pc-windows-msvc.exe`) next to the app
executable is the binary this application calls. To upgrade:

  1. Download the newer release asset from
       https://github.com/Rikorose/DeepFilterNet/releases
  2. Replace the existing file with the new one.
  3. Leave LICENSE.txt and this README.txt intact (upstream licence
     requires preserving the MIT notice).

Attribution / notices
---------------------
When you redistribute this application's bundle further, you must
keep this `third_party/DeepFilterNet/` folder (or the equivalent
LICENSE.txt + README.txt pair) next to the `deep-filter` binary, so
the MIT copyright notice travels with the binary. That is the only
condition the upstream licence places on redistribution.

This application itself is independent of DeepFilterNet; the two are
linked at runtime only through a subprocess call.
