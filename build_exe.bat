@echo off
echo 🔨 Erstelle EXE-Datei für Inventur-Programm V2...
echo.

REM Prüfe ob Python verfügbar ist
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ❌ Python ist nicht installiert oder nicht im PATH!
    pause
    exit /b 1
)

REM Installiere Dependencies falls nötig
echo 📦 Installiere Dependencies...
pip install -r requirements.txt

REM Führe Build-Skript aus
echo 🚀 Starte Build-Prozess...
python build_exe.py

echo.
echo ✅ Build abgeschlossen!
echo 📁 Prüfen Sie den deployment/ Ordner
echo.
pause
