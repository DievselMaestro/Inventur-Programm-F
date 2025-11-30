# INVENTUR-PROGRAMM FÜR FORBO MOVEMENT SYSTEMS

## 📋 Übersicht

Desktop-Anwendung für die Lagerpflege/Inventur bei Forbo Movement Systems. Das Programm ermöglicht es Mitarbeitern, mit einem Barcodescanner Waren im Lager zu erfassen und die Ergebnisse in Excel-Dateien zu speichern.

### ✨ **Neue Features (Version 1.1):**
- **Duplikat-Schutz:** Verhindert versehentliches Doppelscannen
- **Vollbild-Modus:** Startet automatisch im maximierten Modus
- **Führende Nullen:** Charge-Nummern bleiben vollständig erhalten
- **Professionelle Excel-Formatierung:** Saubere Text-Spalten ohne Apostrophe
- **Optimierte Benutzeroberfläche:** Entfernung des Suchfelds für besseren Workflow

## 💻 Systemanforderungen

- **Betriebssystem:** Windows 11 (oder Windows 10)
- **Python:** 3.11 oder höher (wird automatisch installiert)
- **Speicherplatz:** Mindestens 500 MB freier Speicherplatz
- **Hardware:** Barcodescanner (Keyboard Wedge)

## 🚀 Installation

### Schritt 1: Dateien herunterladen
Kopieren Sie alle Programmdateien in einen Ordner auf Ihrem Computer.

### Schritt 2: Python installieren
1. Doppelklicken Sie auf `install_python.bat`
2. Das Script lädt automatisch Python herunter und installiert alle benötigten Module
3. Warten Sie bis "Installation abgeschlossen" angezeigt wird

### Schritt 3: Arbeitstabelle bereitstellen
1. Kopieren Sie Ihre `Arbeitstabelle.xlsx` in den `data/` Ordner
2. Die Datei muss folgende Spalten enthalten:
   - Werk, Lagerort, Material, Materialkurztext, Charge, Länge m, Breite mm, Frei verwendbar, Rollenstatus, Fach

### Schritt 4: Programm starten
1. Doppelklicken Sie auf `start_inventur.bat`
2. Das Programm öffnet sich automatisch im **Vollbild-Modus**

## 📖 Bedienungsanleitung

### Grundlegende Bedienung

1. **Barcode scannen:**
   - Das Eingabefeld ist immer fokussiert
   - Scannen Sie den Barcode oder geben Sie die Charge-Nummer ein
   - Drücken Sie ENTER (oder Scanner sendet automatisch ENTER)

2. **Gefundene Ware:**
   - Daten werden automatisch angezeigt
   - Geben Sie die **Fach-Nummer** ein (Pflichtfeld)
   - Optional: Bemerkung hinzufügen
   - Drücken Sie ENTER zum Speichern

3. **Nicht gefundene Ware:**
   - Dialog öffnet sich automatisch mit professionellem Layout
   - Geben Sie alle Daten manuell ein (Labels über Eingabefeldern)
   - Klicken Sie "Speichern" oder drücken Sie ENTER

4. **🔒 Duplikat-Schutz:**
   - Bereits gescannte Artikel werden automatisch erkannt
   - Warnung bei versehentlichem Doppelscan
   - Verhindert Doppeleinträge in der Inventur

### Tastenkürzel

- **ENTER:** Scan abschließen / Speichern
- **ESC:** Aktuellen Scan abbrechen
- **Ctrl+Z:** Letzte Aktion rückgängig machen
- **Ctrl+S:** Manuell speichern
- **F11:** Vollbild-Modus ein/aus

### Listen-Funktionen

- **Automatische Sortierung:** Neueste Einträge stehen oben
- **Löschen:** Rechtsklick auf Eintrag → "Löschen"
- **Status-Anzeige:** ✅ Gefunden / ⚠️ Nicht gefunden
- **Artikel-Zähler:** Zeigt Gesamtanzahl gescannter Artikel

## 📊 Datenstruktur

### Eingabe: Arbeitstabelle.xlsx
Die Excel-Datei mit der Lager-Datenbank. Enthält alle verfügbaren Artikel mit Charge-Nummern.

### Ausgabe: Inventur.xlsx
Wird automatisch erstellt und enthält zwei Blätter:
- **"Inventur":** Alle gefundenen Waren
- **"Nicht_gefunden":** Manuell erfasste, unbekannte Waren

## 🔧 Erweiterte Funktionen

### 💾 Auto-Save
- Das Programm speichert automatisch nach jedem Scan
- Zusätzlich wird bei jeder Eingabe in Fach/Bemerkung gespeichert
- Bei Programmabsturz gehen keine Daten verloren

### 📤 Export-Funktion
- Klicken Sie "💾 Inventur exportieren"
- Erstellt eine Backup-Kopie mit Zeitstempel
- Originaldatei bleibt unverändert

### 🖥️ Vollbild-Modus
- **Startet automatisch maximiert** für optimale Arbeitsplatznutzung
- Drücken Sie F11 zum Umschalten zwischen Vollbild und Fenster-Modus
- Ideal für Touch-Screens oder große Monitore

### 🔢 Charge-Nummern mit führenden Nullen
- **Vollständige Erhaltung** aller Charge-Nummern (z.B. 0618639923)
- **Professionelle Excel-Formatierung** als Text-Spalten
- **Keine Datenverluste** durch automatische Zahlenkonvertierung

### 🚫 Duplikat-Schutz
- **Automatische Erkennung** bereits gescannter Artikel
- **Sofortige Warnung** bei Doppelscan-Versuchen
- **Verhindert Fehler** in der Inventur-Erfassung

## 🔄 Jährlicher Neustart

Zu Beginn einer neuen Inventur:

1. **Alte Inventur archivieren:**
   - Benennen Sie `data/Inventur.xlsx` um (z.B. `Inventur_2024.xlsx`)
   - Oder löschen Sie die Datei

2. **Neue Arbeitstabelle einsetzen:**
   - Ersetzen Sie `data/Arbeitstabelle.xlsx` mit der neuen Datei

3. **Programm starten:**
   - Das Programm erstellt automatisch eine neue `Inventur.xlsx`

## 🛠️ Fehlerbehebung

### Programm startet nicht
- Prüfen Sie, ob Python installiert ist: `python --version` in der Eingabeaufforderung
- Führen Sie `install_python.bat` erneut aus

### Arbeitstabelle nicht gefunden
- Stellen Sie sicher, dass `Arbeitstabelle.xlsx` im `data/` Ordner liegt
- Prüfen Sie die Spalten-Namen in der Excel-Datei

### Scanner funktioniert nicht
- Testen Sie den Scanner in einem Texteditor
- Falls kein automatisches ENTER: Drücken Sie manuell ENTER nach dem Scan
- Oder verwenden Sie den "Scannen"-Button

### Excel-Fehler
- Schließen Sie alle Excel-Dateien vor dem Programmstart
- Prüfen Sie Schreibrechte im `data/` Ordner

### Führende Nullen verschwinden
- **Problem gelöst:** Charge-Spalten werden automatisch als Text formatiert
- Alle Charge-Nummern bleiben vollständig erhalten (z.B. 0618639923)

### "nan" in Bemerkung-Spalte
- **Problem gelöst:** Leere Bemerkungen werden korrekt als leer angezeigt
- Keine störenden "nan" Texte mehr

### Performance-Probleme
- Bei sehr großen Arbeitstabellen (>5000 Artikel) kann die Suche langsamer werden
- Schließen Sie andere Programme für bessere Performance

### Duplikat-Warnung erscheint fälschlicherweise
- Prüfen Sie ob die Charge-Nummer bereits in der Liste steht
- Bei Bedarf können Sie den Eintrag über Rechtsklick → "Löschen" entfernen

## 📁 Dateistruktur

```
inventur_programm/
├── inventur_app.py          # Hauptprogramm
├── install_python.bat       # Python-Installation
├── start_inventur.bat       # Programm-Start
├── requirements.txt         # Python-Module
├── README.md               # Diese Dokumentation
├── data/                   # Daten-Verzeichnis
│   ├── Arbeitstabelle.xlsx # Lager-Datenbank (manuell kopieren)
│   └── Inventur.xlsx       # Inventur-Ergebnisse (automatisch)
└── config/                 # Konfiguration
    ├── settings.json       # Programmeinstellungen
    └── inventur.log        # Log-Datei
```

## 🔍 Log-Dateien

Das Programm protokolliert alle Aktivitäten in `config/inventur.log`:
- Programmstart/-ende
- Gescannte Artikel
- Fehler und Warnungen

## ⚙️ Konfiguration

Erweiterte Einstellungen in `config/settings.json`:
```json
{
  "auto_save": true,
  "schriftgroesse": 12,
  "farbe_gefunden": "#E8F5E8",
  "farbe_nicht_gefunden": "#FFF2CC",
  "vollbild": true
}
```

### Konfigurationsoptionen:
- **auto_save:** Automatisches Speichern nach jedem Scan
- **schriftgroesse:** Schriftgröße der Benutzeroberfläche
- **farbe_gefunden:** Hintergrundfarbe für gefundene Artikel
- **farbe_nicht_gefunden:** Hintergrundfarbe für nicht gefundene Artikel
- **vollbild:** Startet im maximierten Modus (empfohlen: true)

## 📞 Support

Bei Problemen oder Fragen:
1. Prüfen Sie die Log-Datei `config/inventur.log`
2. Starten Sie das Programm neu
3. Kontaktieren Sie den IT-Support mit der Log-Datei

## 📝 Versionshistorie

### **Version 1.1** - November 2024 ✨
- **🚫 Duplikat-Schutz:** Verhindert versehentliches Doppelscannen
- **🖥️ Vollbild-Modus:** Startet automatisch maximiert
- **🔢 Führende Nullen:** Vollständige Erhaltung aller Charge-Nummern
- **📊 Excel-Formatierung:** Professionelle Text-Spalten ohne Apostrophe
- **🎨 UI-Optimierung:** Entfernung des Suchfelds, verbessertes Dialog-Layout
- **🧹 Datenbereinigung:** Keine "nan" Werte mehr in Bemerkungen
- **📋 Benutzerfreundlichkeit:** Labels über Eingabefeldern im Dialog

### **Version 1.0** - November 2024
- Erste vollständige Version
- Alle Kernfunktionen implementiert
- Getestet für Windows 11

## 🎯 Roadmap

### Geplante Features:
- **Bearbeitungsfunktion:** Nachträgliche Änderung von Einträgen
- **Erweiterte Statistiken:** Inventur-Fortschritt und Auswertungen
- **Backup-Automatisierung:** Automatische tägliche Backups

---

**Entwickelt für Forbo Movement Systems**  
*Professionelle Lagerverwaltung mit Barcodescanner-Integration*

### 🏆 **Produktionsreif für den Einsatz!**
*Alle kritischen Features implementiert und getestet*
