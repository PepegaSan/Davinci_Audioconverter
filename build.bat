@echo off
REM ---------------------------------------------------------------------------
REM  DaVinci Auto Audioconverter - standalone bundle builder (--onedir)
REM
REM  Produces: dist\DavinciAutoAudioconverter\  (self-contained folder)
REM            Zip it up and hand it to an end user; no Python needed on
REM            the target machine.
REM
REM  The bundled "deep-filter.exe" is DeepFilterNet, MIT-licensed and
REM  redistributed under the terms in third_party\DeepFilterNet\LICENSE.txt.
REM  That licence file is copied next to the binary automatically so the
REM  attribution travels with the bundle.
REM
REM  Usage:
REM      build.bat           - build using auto-detected Python
REM      build.bat --clean   - wipe build\ and dist\ first, then build
REM ---------------------------------------------------------------------------

setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo === DaVinci Auto Audioconverter bundle builder ===
echo.

set "EXITCODE=0"

REM --- flag parsing -----------------------------------------------------------
set "DO_CLEAN=0"
:parse
if "%~1"=="" goto parsed
if /i "%~1"=="--clean" ( set "DO_CLEAN=1" & shift & goto parse )
if /i "%~1"=="-clean"  ( set "DO_CLEAN=1" & shift & goto parse )
if /i "%~1"=="/clean"  ( set "DO_CLEAN=1" & shift & goto parse )
shift
goto parse
:parsed

if "!DO_CLEAN!"=="1" (
    echo Wiping build\ and dist\ ...
    if exist "build" rmdir /s /q "build"
    if exist "dist"  rmdir /s /q "dist"
    echo.
)

REM --- locate deep-filter*.exe ------------------------------------------------
REM We don't hardcode a filename because the user will typically download
REM whichever version is current from GitHub. Pick the first match, prefer
REM the canonical name.

set "DFILTER="
if exist "deep-filter.exe" set "DFILTER=deep-filter.exe"
if not defined DFILTER (
    for /f "delims=" %%F in ('dir /b /o-n "deep-filter-*.exe" 2^>nul') do (
        if not defined DFILTER set "DFILTER=%%F"
    )
)

if not defined DFILTER (
    echo [ERROR] No deep-filter.exe found in the repo root.
    echo.
    echo Download the Windows binary from:
    echo     https://github.com/Rikorose/DeepFilterNet/releases
    echo and drop it here ^(e.g. deep-filter-0.5.6-x86_64-pc-windows-msvc.exe^).
    echo The spec file picks it up automatically.
    echo.
    set "EXITCODE=1"
    goto :end
)
echo Bundling DeepFilterNet binary: !DFILTER!
echo.

REM --- locate Python interpreter ---------------------------------------------
REM Same priority as install.bat / run.bat.

set "PYEXE="
set "PYVER_USED="

if exist ".venv\Scripts\python.exe" (
    set "PYEXE=.venv\Scripts\python.exe"
    set "PYVER_USED=.venv"
)

if not defined PYEXE (
    where py >nul 2>nul
    if not errorlevel 1 (
        for %%V in (3.11 3.12 3.10 3.13) do (
            if not defined PYEXE (
                py -%%V -c "import sys" >nul 2>nul
                if not errorlevel 1 (
                    set "PYEXE=py -%%V"
                    set "PYVER_USED=%%V"
                )
            )
        )
        if not defined PYEXE set "PYEXE=py -3"
        if not defined PYVER_USED set "PYVER_USED=default (py -3)"
    )
)

if not defined PYEXE (
    where python >nul 2>nul
    if not errorlevel 1 (
        set "PYEXE=python"
        set "PYVER_USED=python on PATH"
    )
)

if not defined PYEXE (
    echo [ERROR] No Python 3 interpreter found. Run install.bat first.
    set "EXITCODE=1"
    goto :end
)

echo Using interpreter: !PYEXE!    ^(version: !PYVER_USED!^)
!PYEXE! --version
echo.

REM --- make sure PyInstaller is available ------------------------------------
!PYEXE! -m pip show pyinstaller >nul 2>nul
if errorlevel 1 (
    echo Installing PyInstaller ^(one-time^) ...
    !PYEXE! -m pip install --upgrade pyinstaller
    if errorlevel 1 (
        echo [ERROR] PyInstaller install failed.
        set "EXITCODE=1"
        goto :end
    )
    echo.
)

REM --- make sure runtime deps are installed in the selected interpreter ------
REM PyInstaller imports the app to analyse it; missing deps would crash the
REM build very early with a cryptic ModuleNotFoundError.

!PYEXE! -c "import customtkinter, tkinterdnd2" >nul 2>nul
if errorlevel 1 (
    echo Installing runtime dependencies into build interpreter ...
    !PYEXE! -m pip install -r requirements.txt
    if errorlevel 1 (
        echo [ERROR] Dependency install failed.
        set "EXITCODE=1"
        goto :end
    )
    echo.
)

REM --- run the build ----------------------------------------------------------
echo Running PyInstaller (this can take 1-2 minutes) ...
echo.
!PYEXE! -m PyInstaller --noconfirm DavinciAutoAudioconverter.spec
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    set "EXITCODE=1"
    goto :end
)

REM --- post-build: surface user-facing files to the bundle root --------------
REM PyInstaller >=6 collects everything into a ``_internal\`` subfolder by
REM default, which buries the DeepFilterNet binary and its licence notice
REM where nobody will find them. Move both up to the bundle root so:
REM   (a) the user can drop in a newer ``deep-filter.exe`` without digging;
REM   (b) the MIT licence notice is visible next to the binary it covers;
REM   (c) the runtime discovery in audio_preprocess._bundle_search_roots
REM       hits the sibling-to-exe path first and avoids a _MEIPASS walk.

set "OUT_DIR=dist\DavinciAutoAudioconverter"
set "INT_DIR=!OUT_DIR!\_internal"

if exist "!INT_DIR!\deep-filter.exe" (
    echo Moving deep-filter.exe to bundle root ...
    move /y "!INT_DIR!\deep-filter.exe" "!OUT_DIR!\deep-filter.exe" >nul
)
if exist "!INT_DIR!\third_party" (
    echo Moving third_party\ to bundle root ...
    if exist "!OUT_DIR!\third_party" rmdir /s /q "!OUT_DIR!\third_party"
    move /y "!INT_DIR!\third_party" "!OUT_DIR!\third_party" >nul
)

REM --- post-build sanity check -----------------------------------------------
if not exist "!OUT_DIR!\DavinciAutoAudioconverter.exe" (
    echo [ERROR] Build finished but the main exe is missing in !OUT_DIR!.
    set "EXITCODE=1"
    goto :end
)
if not exist "!OUT_DIR!\deep-filter.exe" (
    echo [ERROR] deep-filter.exe is missing from the bundle.
    echo         Expected at: !OUT_DIR!\deep-filter.exe
    echo         PyInstaller placed files into _internal\ — the post-build
    echo         move step above should have surfaced it. Re-run with
    echo         --clean and check the build log for move errors.
    set "EXITCODE=1"
    goto :end
)
if not exist "!OUT_DIR!\third_party\DeepFilterNet\LICENSE.txt" (
    echo [ERROR] DeepFilterNet LICENSE.txt missing from the bundle.
    echo         Re-run with --clean to rebuild from scratch.
    set "EXITCODE=1"
    goto :end
)

echo.
echo === Build complete ===
echo.
echo Bundle directory : !OUT_DIR!
echo Main executable  : !OUT_DIR!\DavinciAutoAudioconverter.exe
echo DeepFilterNet    : !OUT_DIR!\deep-filter.exe       ^(MIT - see third_party\DeepFilterNet\LICENSE.txt^)
echo.
echo Ship the entire !OUT_DIR! folder. The LICENSE.txt next to
echo the binary MUST stay with it when you redistribute.

:end
echo.
pause
endlocal & exit /b %EXITCODE%
