@echo off
REM Startet die CBP 7501 Import-App
cd /d "%~dp0"
python cbp7501_app.py 2>>startup_error.log
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo === FEHLER beim Starten (siehe startup_error.log) ===
    type startup_error.log
    echo.
    pause
)
