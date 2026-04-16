@echo off
cd /d "%~dp0"
python waypoint_map_installer.py
if errorlevel 1 (
    echo.
    echo === ERROR: App crashed. See message above. ===
    pause
)
