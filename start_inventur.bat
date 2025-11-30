@echo off
title FORBO INVENTUR-PROGRAMM
echo ===============================================
echo FORBO INVENTUR-PROGRAMM WIRD GESTARTET...
echo ===============================================
echo.

:: Prüfe ob Python installiert ist
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [FEHLER] Python ist nicht installiert!
    echo.
    echo Bitte führen Sie zuerst install_python.bat aus.
    echo.
    pause
    exit /b 1
)

:: Wechsle zum Programmverzeichnis
cd /d "%~dp0"

:: Erstelle data-Ordner falls nicht vorhanden
if not exist "data" mkdir data
if not exist "config" mkdir config

:: Prüfe ob Arbeitstabelle vorhanden ist
if not exist "data\Arbeitstabelle.xlsx" (
    echo [WARNUNG] Arbeitstabelle.xlsx nicht gefunden!
    echo.
    echo Bitte kopieren Sie die Arbeitstabelle.xlsx in den data\ Ordner.
    echo Pfad: %~dp0data\Arbeitstabelle.xlsx
    echo.
    echo Möchten Sie das Programm trotzdem starten? (j/n)
    set /p choice=
    if /i not "%choice%"=="j" exit /b 0
)

echo [INFO] Starte Inventur-Programm...
echo.

:: Starte Python-Programm
python inventur_app.py

:: Falls Fehler aufgetreten
if %errorlevel% neq 0 (
    echo.
    echo [FEHLER] Das Programm wurde mit einem Fehler beendet.
    echo Fehlercode: %errorlevel%
    echo.
    pause
)
