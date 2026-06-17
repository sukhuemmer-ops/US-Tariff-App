@echo off
setlocal
cd /d "C:\DEV\Tariff-Database"

echo.
echo === GitHub Sync - US-Tariff-App ===
echo.

git add .

for /f "tokens=*" %%i in ('git status --porcelain') do set CHANGES=%%i
if not defined CHANGES (
    echo Keine Aenderungen gefunden - alles aktuell.
    echo.
    pause
    exit /b 0
)

echo Geaenderte Dateien:
git status --short
echo.

set /p MSG="Beschreibung der Aenderung (Enter fuer automatisch): "
if "%MSG%"=="" (
    for /f "tokens=*" %%d in ('powershell -command "Get-Date -Format \"yyyy-MM-dd HH:mm\""') do set DATUM=%%d
    set MSG=Update %DATUM%
)

git commit -m "%MSG%"
git push origin dev

echo.
echo === Sync abgeschlossen! ===
echo https://github.com/sukhuemmer-ops/US-Tariff-App/tree/dev
echo.
pause
endlocal
