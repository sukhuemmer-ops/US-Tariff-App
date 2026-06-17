@echo off
setlocal
cd /d "C:\DEV\Tariff-Database"

REM Remote sicherstellen
git remote get-url origin >nul 2>&1
if errorlevel 1 (
    git remote add origin https://github.com/sukhuemmer-ops/US-Tariff-App.git
)

REM Prüfen ob Änderungen vorhanden
git status --porcelain > "%TEMP%\git_status.txt" 2>&1
for %%A in ("%TEMP%\git_status.txt") do if %%~zA==0 (
    exit /b 0
)

REM Datum/Uhrzeit für Commit-Message
for /f "tokens=*" %%d in ('powershell -command "Get-Date -Format \"yyyy-MM-dd HH:mm\""') do set DATUM=%%d

REM Alle Änderungen stagen (außer .gitignore-Einträge)
git add .

REM Committen
git commit -m "Auto-Sync: %DATUM%" >nul 2>&1

REM Pushen
git push origin main >> "%TEMP%\git_autosync.log" 2>&1

endlocal
