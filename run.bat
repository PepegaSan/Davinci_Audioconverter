@echo off
REM Launch the app. Prefers the local .venv (if install.bat --venv was used)
REM and otherwise picks the most compatible Python via the py launcher
REM (3.11 > 3.12 > 3.10 > 3.13 > any py -3 > python on PATH).

setlocal EnableExtensions EnableDelayedExpansion
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" main.py
    set "RC=%errorlevel%"
    endlocal & exit /b %RC%
)

set "PYEXE="
where py >nul 2>nul
if not errorlevel 1 (
    for %%V in (3.11 3.12 3.10 3.13) do (
        if not defined PYEXE (
            py -%%V -c "import sys" >nul 2>nul
            if not errorlevel 1 set "PYEXE=py -%%V"
        )
    )
    if not defined PYEXE set "PYEXE=py -3"
)

if not defined PYEXE (
    where python >nul 2>nul
    if not errorlevel 1 set "PYEXE=python"
)

if not defined PYEXE (
    echo [ERROR] Python 3 was not found on PATH. Run install.bat first.
    pause
    endlocal & exit /b 1
)

!PYEXE! main.py
set "RC=%errorlevel%"
endlocal & exit /b %RC%
