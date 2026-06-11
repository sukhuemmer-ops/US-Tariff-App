@echo off
REM Kopiert 3163973032.pdf aus dem SharePoint/Netzordner in den lokalen PDF-Ordner
REM Dann neu importieren via App (Datenbank leeren -> Drag & Drop)

set SRC=C:\Users\huemmkma\Catensys Germany GmbH\GRP_CUS_Troy - US_Trump-Tariffs - Trump-Tariffs\2026-03\2026-03-All-Duties\3163973032.pdf
set DST=C:\DEV\Tariff-Database\PDF\3163973032.pdf

if not exist "%SRC%" (
    echo FEHLER: Quelldatei nicht gefunden:
    echo %SRC%
    pause
    exit /b 1
)

copy /Y "%SRC%" "%DST%"
if %ERRORLEVEL%==0 (
    echo OK: 3163973032.pdf wurde nach PDF\ kopiert.
) else (
    echo FEHLER beim Kopieren.
)
pause
