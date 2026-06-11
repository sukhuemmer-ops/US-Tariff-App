@echo off
REM Verknuepfung: startet die CBP 7501 Import-App (Drag & Drop -> Datenbank)
cd /d "%~dp0"
start "" python cbp7501_app.py
