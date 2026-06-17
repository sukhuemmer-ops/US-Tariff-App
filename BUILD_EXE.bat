@echo off
:: ============================================================
::  CBP7501 App – EXE Build Script
::  Aufruf: Doppelklick oder in PowerShell: .\BUILD_EXE.bat
:: ============================================================
cd /d "%~dp0"

echo.
echo ============================================================
echo   CBP7501 EXE Build
echo ============================================================
echo.

:: Prüfe ob venv existiert
if not exist "venv\Scripts\activate.bat" (
    echo [1/5] Erstelle virtuelle Umgebung ...
    python -m venv venv
) else (
    echo [1/5] venv bereits vorhanden – übersprungen.
)

:: Aktiviere venv
call venv\Scripts\activate.bat

:: Pakete installieren
echo.
echo [2/5] Installiere Pakete aus requirements.txt ...
pip install -r requirements.txt --quiet

:: Tesseract prüfen (wird nicht eingebunden, muss installiert sein)
echo.
echo [3/5] Prüfe Tesseract-OCR ...
where tesseract >nul 2>&1
if %errorlevel% neq 0 (
    echo    WARNUNG: Tesseract nicht gefunden!
    echo    Bitte installieren: https://github.com/UB-Mannheim/tesseract/wiki
    echo    Danach Tesseract-Pfad zu PATH hinzufügen.
    echo    OCR-Funktionen werden im EXE nicht verfügbar sein.
) else (
    for /f "tokens=*" %%i in ('where tesseract') do echo    Gefunden: %%i
)

:: Alte Build-Artefakte löschen
echo.
echo [4/5] Bereinige alte Build-Artefakte ...
if exist "dist\CBP7501" rmdir /s /q "dist\CBP7501"
if exist "build\CBP7501" rmdir /s /q "build\CBP7501"

:: Build
echo.
echo [5/5] Starte PyInstaller Build ...
echo.
pyinstaller cbp7501_app.spec --noconfirm

if %errorlevel% neq 0 (
    echo.
    echo ============================================================
    echo   BUILD FEHLGESCHLAGEN – siehe Fehler oben
    echo ============================================================
    pause
    exit /b 1
)

:: Datenbank + Einstellungen automatisch kopieren
echo.
echo [6/6] Kopiere Daten in dist\CBP7501\ ...

if exist "cbp7501.db" (
    copy /Y "cbp7501.db" "dist\CBP7501\cbp7501.db" >nul
    echo    cbp7501.db       – kopiert
) else (
    echo    cbp7501.db       – NICHT GEFUNDEN (wird beim ersten Start neu angelegt)
)

if exist "sap_settings.ini" (
    copy /Y "sap_settings.ini" "dist\CBP7501\sap_settings.ini" >nul
    echo    sap_settings.ini – kopiert
) else (
    echo    sap_settings.ini – nicht vorhanden (SAP-Verbindung muss in App eingerichtet werden)
)

if exist "pdf_archive" (
    echo    pdf_archive\     – wird NICHT kopiert (zu gross; bei Bedarf manuell)
)

echo.
echo ============================================================
echo   BUILD ERFOLGREICH
echo   EXE: dist\CBP7501\CBP7501.exe
echo.
echo   Datenbank und Einstellungen wurden automatisch kopiert.
echo   Tesseract muss auf dem Ziel-PC installiert sein
echo   fuer OCR-Funktionen (Bild-PDFs).
echo ============================================================
echo.
pause
