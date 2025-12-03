# 🚀 EXE-Erstellung für Inventur-Programm V2

## 📋 Voraussetzungen

1. **Python 3.11+** installiert
2. **Alle Dependencies** installiert:
   ```bash
   pip install -r requirements.txt
   ```

## 🔨 EXE erstellen

### Automatisch (empfohlen):
```bash
python build_exe.py
```

### Manuell mit PyInstaller:
```bash
pyinstaller --onefile --windowed --name=InventurApp_V2 --add-data="config;config" inventur_app.py
```

## 📁 Deployment-Struktur

Nach dem Build wird automatisch ein `deployment/` Ordner erstellt:

```
deployment/
├── InventurApp_V2.exe          # Hauptprogramm
├── data/                       # Datenordner
│   ├── Arbeitstabelle.xlsx     # ← Ihre Datei hier einfügen
│   └── backups/                # Automatische Backups
├── config/                     # Konfiguration
│   └── settings.json           # Programmeinstellungen
└── README.txt                  # Installationsanleitung
```

## ✅ Wichtige Punkte für EXE-Deployment

### 1. **Pfad-Handling**
- ✅ Das Programm erkennt automatisch ob es als .py oder .exe läuft
- ✅ Alle Pfade sind relativ zum EXE-Standort
- ✅ `data/` Ordner wird immer neben der EXE gesucht

### 2. **Datenordner-Struktur**
```
InventurApp_V2.exe              # Hier starten
data/
├── Arbeitstabelle.xlsx         # Input (von Ihnen)
├── Inventur_Rollen.xlsx        # Output (automatisch)
├── Inventur_Granulat.xlsx      # Output (automatisch)
├── inventur.log                # Logdatei
└── backups/                    # Export-Backups
    ├── Inventur_Rollen_Backup_20241202_143022.xlsx
    └── Inventur_Granulat_Backup_20241202_143022.xlsx
```

### 3. **Erste Installation**
1. **Deployment-Ordner kopieren** an gewünschten Ort
2. **Arbeitstabelle.xlsx** in `data/` Ordner kopieren
3. **EXE starten** → Programm läuft sofort

### 4. **Arbeitstabelle-Format**
Die `Arbeitstabelle.xlsx` muss **zwei Tabellenblätter** haben:
- **"Rollen"** - mit Spalten: Charge, Material, Materialkurztext, Länge m, Breite mm, Frei verwendbar
- **"Granulate"** - mit Spalten: Charge, Material/Materialnummer, Materialkurztext, Frei verwendbar

## 🔧 Troubleshooting

### Problem: "Arbeitstabelle nicht gefunden"
**Lösung:** Kopieren Sie `Arbeitstabelle.xlsx` in den `data/` Ordner neben der EXE

### Problem: "Fehlende Spalten"
**Lösung:** Überprüfen Sie die Spaltennamen in beiden Tabellenblättern

### Problem: EXE startet nicht
**Lösung:** 
1. Starten Sie über Kommandozeile für Fehlermeldungen
2. Prüfen Sie ob alle Dateien im deployment/ Ordner vorhanden sind

## 📊 Vorteile der EXE-Version

- ✅ **Keine Python-Installation** nötig auf Zielrechner
- ✅ **Portable** - einfach kopieren und starten
- ✅ **Automatische Pfade** - funktioniert überall
- ✅ **Professionell** - sieht aus wie normale Software
- ✅ **Einfache Verteilung** - ein Ordner für alles

## 🎯 Deployment-Workflow

1. **Entwicklung:** Python-Skript testen
2. **Build:** `python build_exe.py` ausführen
3. **Test:** EXE im deployment/ Ordner testen
4. **Verteilung:** Gesamten deployment/ Ordner kopieren
5. **Installation:** Arbeitstabelle.xlsx einfügen → fertig!

Die EXE wird **immer** im Ordner suchen, wo sie gestartet wird. Das macht sie sehr portabel und einfach zu verwenden.
