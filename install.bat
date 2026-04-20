@echo off
REM ---------------------------------------------------------------------------
REM  DaVinci Auto Audioconverter - dependency installer
REM
REM  Default: installs requirements.txt into the current user's Python
REM  (pip --user). A local .venv is only created when explicitly requested
REM  via "install.bat --venv" or by answering "y" at the prompt.
REM
REM  Python selection:
REM      Preferred versions (descending): 3.11, 3.12, 3.10, 3.13.
REM      3.11 is the sweet spot for this project (stable Tk + CustomTkinter,
REM      still has pre-built wheels for the optional deepfilternet package).
REM      The Windows Python launcher "py -X.Y" is used to pin the version
REM      when available; otherwise we fall back to "py -3" and finally to
REM      a "python" on PATH.
REM
REM  Flags:
REM      --venv      force virtual-environment install (no prompt)
REM      --system    force system/user install         (no prompt)
REM      --py X.Y    force a specific Python version  (e.g. --py 3.11)
REM ---------------------------------------------------------------------------

setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo === DaVinci Auto Audioconverter installer ===
echo.

set "EXITCODE=0"
set "USE_VENV=0"
set "ASK_USER=1"
set "FORCE_PYVER="

REM --- parse flags ------------------------------------------------------------
:parse_args
if "%~1"=="" goto args_done
if /i "%~1"=="--venv"   ( set "USE_VENV=1" & set "ASK_USER=0" & shift & goto parse_args )
if /i "%~1"=="-venv"    ( set "USE_VENV=1" & set "ASK_USER=0" & shift & goto parse_args )
if /i "%~1"=="/venv"    ( set "USE_VENV=1" & set "ASK_USER=0" & shift & goto parse_args )
if /i "%~1"=="--system" ( set "USE_VENV=0" & set "ASK_USER=0" & shift & goto parse_args )
if /i "%~1"=="--py"     ( set "FORCE_PYVER=%~2" & shift & shift & goto parse_args )
shift
goto parse_args
:args_done

REM --- locate a Python interpreter --------------------------------------------
REM We prefer 3.11, then 3.12, 3.10, 3.13. Override with --py X.Y if needed.
set "PYEXE="
set "PYVER_USED="

where py >nul 2>nul
if errorlevel 1 goto try_python_on_path

if defined FORCE_PYVER (
    py -!FORCE_PYVER! -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "PYEXE=py -!FORCE_PYVER!"
        set "PYVER_USED=!FORCE_PYVER!"
        goto py_found
    )
    echo [WARN] Python !FORCE_PYVER! is not installed on this machine.
    echo.
    echo        To get it, download the Windows installer from
    echo            https://www.python.org/downloads/windows/
    echo        During setup, tick BOTH
    echo            [x] Add Python to PATH
    echo            [x] py launcher
    echo        then re-run:  install.bat --venv --py !FORCE_PYVER!
    echo.
    echo        Alternatively, if you are on a newer Python version
    echo        ^(e.g. 3.12 / 3.13^) and don't want to install 3.11, it
    echo        is perfectly fine to run the app on that newer Python
    echo        as long as you use the standalone ``deep-filter.exe``
    echo        binary from the DeepFilterNet releases page — see the
    echo        README section "Setup step 3" for the link and install
    echo        instructions.
    echo.
    echo        Falling back to automatic version detection...
    echo.
)

for %%V in (3.11 3.12 3.10 3.13) do (
    if not defined PYEXE (
        py -%%V -c "import sys" >nul 2>nul
        if not errorlevel 1 (
            set "PYEXE=py -%%V"
            set "PYVER_USED=%%V"
        )
    )
)

if not defined PYEXE (
    py -3 -c "import sys" >nul 2>nul
    if not errorlevel 1 (
        set "PYEXE=py -3"
        set "PYVER_USED=default (py -3)"
    )
)

:try_python_on_path
if not defined PYEXE (
    where python >nul 2>nul
    if not errorlevel 1 (
        set "PYEXE=python"
        set "PYVER_USED=python on PATH"
    )
)

:py_found
if not defined PYEXE (
    echo [ERROR] No compatible Python 3 interpreter was found.
    echo.
    echo This project is tested with Python 3.10 - 3.13 ^(sweet spot: 3.11^).
    echo Please install one from https://www.python.org/downloads/release/python-3119/
    echo and make sure "Add Python to PATH" + "py launcher" are enabled.
    echo.
    set "EXITCODE=1"
    goto :end
)

echo Using interpreter: !PYEXE!    ^(version: !PYVER_USED!^)
!PYEXE! --version
echo.

REM --- optional prompt --------------------------------------------------------
if "!ASK_USER!"=="1" (
    echo Install into a local virtual environment ^(.venv, Python !PYVER_USED!^) ?
    echo   [y] yes - isolates dependencies in .\.venv   ^(recommended^)
    echo   [n] no  - installs into the current Python via pip --user
    echo.
    echo Note: ``venv`` uses whichever Python is already installed — it
    echo       does NOT download Python itself. The recommended sweet spot
    echo       for this project is 3.11; if you want to lock the venv to
    echo       that version but don't have it yet, cancel here ^(Ctrl+C^),
    echo       grab the installer from
    echo           https://www.python.org/downloads/windows/
    echo       and re-run ``install.bat --venv --py 3.11``. If you would
    echo       rather stay on your current newer Python ^(3.12/3.13^), that
    echo       is fine — just make sure to use the standalone
    echo       ``deep-filter.exe`` binary from the DeepFilterNet releases
    echo       page ^(README, Setup step 3^).
    echo.
    set "ANSWER="
    set /p "ANSWER=Your choice [y/N]: "
    if /i "!ANSWER!"=="y"   set "USE_VENV=1"
    if /i "!ANSWER!"=="yes" set "USE_VENV=1"
)

echo.

if "!USE_VENV!"=="1" goto :install_venv
goto :install_system


REM ---------------------------------------------------------------------------
:install_venv
REM ---------------------------------------------------------------------------
if not exist ".venv\Scripts\python.exe" (
    echo Creating virtual environment in .venv ^(Python !PYVER_USED!^) ...
    !PYEXE! -m venv .venv
    if errorlevel 1 (
        echo [ERROR] Failed to create the virtual environment.
        set "EXITCODE=1"
        goto :end
    )
) else (
    echo Reusing existing virtual environment in .venv
    echo       ^(delete the folder to rebuild with a different Python version^)
)

set "VENV_PY=.venv\Scripts\python.exe"

echo.
echo venv Python version:
"!VENV_PY!" --version
echo.
echo Upgrading pip ...
"!VENV_PY!" -m pip install --upgrade pip
if errorlevel 1 (
    echo [ERROR] pip upgrade failed.
    set "EXITCODE=1"
    goto :end
)

echo.
echo Installing dependencies from requirements.txt ...
"!VENV_PY!" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Dependency installation failed.
    set "EXITCODE=1"
    goto :end
)

echo.
echo === Done (virtual environment). ===
echo Start the app with:  run.bat
echo                  or  .venv\Scripts\python.exe main.py
goto :end


REM ---------------------------------------------------------------------------
:install_system
REM ---------------------------------------------------------------------------
echo Installing dependencies into the current Python ^(pip --user^) ...
echo.

!PYEXE! -m pip install --user --upgrade pip
if errorlevel 1 (
    echo [ERROR] pip upgrade failed.
    set "EXITCODE=1"
    goto :end
)

!PYEXE! -m pip install --user -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Dependency installation failed.
    set "EXITCODE=1"
    goto :end
)

echo.
echo === Done (user install). ===
echo Start the app with:  run.bat
echo                  or  !PYEXE! main.py
goto :end


REM ---------------------------------------------------------------------------
:end
echo.
pause
endlocal & exit /b %EXITCODE%
