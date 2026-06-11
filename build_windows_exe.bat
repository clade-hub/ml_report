@echo off
REM Change to the directory where this bat file is located
cd /d "%~dp0"

echo ========================================
echo ML Report - Windows EXE Builder
echo ========================================
echo Arbeitsverzeichnis: %cd%
echo.

REM Check if Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python ist nicht installiert oder nicht im PATH
    echo Bitte Python installieren von https://python.org
    echo WICHTIG: Bei der Installation "Add Python to PATH" ankreuzen!
    pause
    exit /b 1
)

echo Python gefunden:
python --version

echo.
echo Installiere benoetigte Pakete...
python -m pip install --upgrade pip
python -m pip install pyinstaller odfpy Pillow pdf2image

echo.
echo ========================================
echo Erstelle EXE Datei...
echo ========================================
echo.

REM First clean old build files
echo Loesche alte Build-Dateien...
if exist "build" rmdir /s /q "build"
if exist "dist" rmdir /s /q "dist"
if exist "ML_Report.spec" del "ML_Report.spec"

echo.
echo Starte PyInstaller...
python -m PyInstaller --onefile --windowed --name "ML_Report" report_generator.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo FEHLER: PyInstaller ist fehlgeschlagen!
    echo ========================================
    echo Bitte Fehlermeldungen oben lesen.
    pause
    exit /b 1
)

if not exist "dist\ML_Report.exe" (
    echo.
    echo ========================================
    echo FEHLER: EXE wurde nicht erstellt!
    echo ========================================
    echo Bitte Fehlermeldungen oben lesen.
    pause
    exit /b 1
)

echo.
echo ========================================
echo Fertig!
echo ========================================
echo.
echo Deine EXE Datei ist hier: dist\ML_Report.exe
echo.
echo WICHTIG: Poppler muss installiert sein fuer PDF Konvertierung:
echo 1. Download: https://github.com/osborn/poppler-windows/releases
echo 2. Entpacken nach C:\Programme\poppler
echo 3. C:\Programme\poppler\Library\bin zum PATH hinzufuegen
echo.
pause
