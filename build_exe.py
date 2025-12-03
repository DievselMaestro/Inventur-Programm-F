#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Build-Skript für Inventur-Programm V2
Erstellt eine EXE-Datei mit PyInstaller
"""

import os
import sys
import shutil
from pathlib import Path

def build_exe():
    """Erstellt die EXE-Datei"""
    print("🔨 Erstelle EXE-Datei für Inventur-Programm V2...")
    
    # PyInstaller-Befehl
    cmd = [
        "python", "-m", "PyInstaller",
        "--onefile",                    # Eine einzige EXE-Datei
        "--windowed",                   # Kein Konsolen-Fenster
        "--name=InventurApp_V2",        # Name der EXE
        "--icon=config/forbo_icon.ico", # Icon (falls vorhanden)
        "--add-data=config;config",     # Config-Ordner mitpacken
        "--distpath=dist",              # Output-Verzeichnis
        "--workpath=build",             # Temporäres Build-Verzeichnis
        "--specpath=.",                 # .spec Datei hier erstellen
        "inventur_app.py"               # Haupt-Python-Datei
    ]
    
    # Führe PyInstaller aus
    result = os.system(" ".join(cmd))
    
    if result == 0:
        print("✅ EXE erfolgreich erstellt!")
        
        # Erstelle Deployment-Ordner
        deployment_dir = Path("deployment")
        if deployment_dir.exists():
            shutil.rmtree(deployment_dir)
        deployment_dir.mkdir()
        
        # Kopiere EXE
        exe_source = Path("dist/InventurApp_V2.exe")
        exe_target = deployment_dir / "InventurApp_V2.exe"
        shutil.copy2(exe_source, exe_target)
        
        # Erstelle data-Ordner Struktur
        data_dir = deployment_dir / "data"
        data_dir.mkdir()
        (data_dir / "backups").mkdir()
        
        # Erstelle config-Ordner
        config_dir = deployment_dir / "config"
        config_dir.mkdir()
        
        # Kopiere Standard-Config
        if Path("config/settings.json").exists():
            shutil.copy2("config/settings.json", config_dir / "settings.json")
        
        # Erstelle README für Deployment
        readme_content = """# Inventur-Programm V2 - Deployment

## Installation:
1. Kopieren Sie den gesamten Ordner an den gewünschten Ort
2. Starten Sie InventurApp_V2.exe

## Wichtige Dateien:
- InventurApp_V2.exe: Hauptprogramm
- data/Arbeitstabelle.xlsx: Ihre Arbeitstabelle (muss von Ihnen eingefügt werden)
- data/: Hier werden alle Inventur-Dateien gespeichert
- config/: Konfigurationsdateien

## Erste Schritte:
1. Kopieren Sie Ihre Arbeitstabelle.xlsx in den data/ Ordner
2. Die Arbeitstabelle muss zwei Tabellenblätter haben: "Rollen" und "Granulate"
3. Starten Sie das Programm

## Datenstruktur:
- Inventur_Rollen.xlsx: Alle erfassten Rollen
- Inventur_Granulat.xlsx: Alle erfassten Granulate
- backups/: Automatische Backups bei Export

Das Programm erstellt automatisch alle benötigten Dateien beim ersten Start.
"""
        
        with open(deployment_dir / "README.txt", "w", encoding="utf-8") as f:
            f.write(readme_content)
        
        print(f"📁 Deployment-Paket erstellt in: {deployment_dir.absolute()}")
        print("\n🎯 Nächste Schritte:")
        print("1. Kopieren Sie Ihre Arbeitstabelle.xlsx in deployment/data/")
        print("2. Testen Sie die EXE: deployment/InventurApp_V2.exe")
        print("3. Verteilen Sie den gesamten deployment/ Ordner")
        
    else:
        print("❌ Fehler beim Erstellen der EXE!")
        print("Stellen Sie sicher, dass PyInstaller installiert ist:")
        print("pip install pyinstaller")

if __name__ == "__main__":
    build_exe()
