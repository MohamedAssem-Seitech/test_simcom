@echo off
setlocal
cd /d "%~dp0"
if exist ".venv\Scripts\python.exe" (
    ".venv\Scripts\python.exe" "version_history_gui.py"
) else (
    py -3 "version_history_gui.py"
)
if errorlevel 1 (
    echo.
    echo Install Python 3.10+ with Tcl/Tk and the requirements listed in README.md.
    pause
)
