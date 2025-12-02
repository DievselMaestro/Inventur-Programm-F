# INVENTUR-PROGRAMM V2 FÜR FORBO MOVEMENT SYSTEMS

## 📋 Übersicht

Desktop-Anwendung für die Lagerpflege/Inventur bei Forbo Movement Systems. Das Programm unterstützt jetzt **ROLLEN UND GRANULATE** mit separaten Excel-Dateien und erweiterten Eingabefeldern.

### 🎆 **Neue Features (Version 2.0):**
- **🔵 Rollen-Unterstützung:** Separate Erfassung mit Fach + Breite kontrolliert
- **🟨 Granulat-Unterstützung:** Gewichts-Erfassung mit Zählmenge
- **📁 Zwei Excel-Dateien:** Inventur_Rollen.xlsx und Inventur_Granulat.xlsx
- **🎨 Visuelle Unterscheidung:** Blaue und gelbe Farbkodierung
- **📊 Erweiterte Statistiken:** Getrennte Zähler für beide Warentypen
- **⚙️ Intelligente Validierung:** Breite (1-4 Ziffern) und Gewicht (Dezimalzahlen)

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
2. **WICHTIG:** Die Datei muss **zwei Tabellenblätter** enthalten:

#### Tabellenblatt "Rollen":
   - Werk, Lagerort, Material, Materialkurztext, Charge, Länge m, Breite mm, Frei verwendbar, Rollenstatus, Fach

#### Tabellenblatt "Granulate":
   - Werk, LOrt, Materialnummer, Materialkurztext, Charge, Frei verwendbar, BME

### Schritt 4: Programm starten
1. Doppelklicken Sie auf `start_inventur.bat`
2. Das Programm öffnet sich automatisch im **Vollbild-Modus**

## 📖 Bedienungsanleitung

### Grundlegende Bedienung

1. **Barcode scannen:**
   - Das Eingabefeld ist immer fokussiert
   - Scannen Sie den Barcode oder geben Sie die Charge-Nummer ein
   - Drücken Sie ENTER (oder Scanner sendet automatisch ENTER)

2. **Gefundene Rolle (🔵 BLAU):**
   - Daten werden automatisch angezeigt
   - Geben Sie **Fach (Lagerort)** ein (Pflichtfeld)
   - Geben Sie **Breite kontrolliert (mm)** ein (Pflichtfeld, 1-4 Ziffern)
   - Optional: Bemerkung hinzufügen
   - Drücken Sie ENTER zum Speichern

3. **Gefundenes Granulat (🟨 GELB):**
   - Daten werden automatisch angezeigt
   - Geben Sie **Zählmenge (KG)** ein (Pflichtfeld, Dezimalzahl möglich)
   - Optional: Bemerkung hinzufügen
   - Drücken Sie ENTER zum Speichern

4. **Nicht gefundene Ware:**
   - Dialog öffnet sich mit **Typ-Auswahl** (🔵 Rolle oder 🟨 Granulat)
   - Geben Sie alle Daten manuell ein (Labels über Eingabefeldern)
   - Eingabefelder passen sich automatisch an den gewählten Typ an
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
- **Typ-Anzeige:** 🔵 Rolle / 🟨 Granu mit visueller Unterscheidung
- **Löschen:** Rechtsklick auf Eintrag → "Löschen"
- **Status-Anzeige:** ✅ Gefunden / ⚠️ Nicht gefunden
- **Erweiterte Zähler:** Zeigt Rollen und Granulate separat

## 📊 Datenstruktur V2

### Eingabe: Arbeitstabelle.xlsx (Erweitert)
Die Excel-Datei mit **zwei Tabellenblättern:**
- **"Rollen":** Alle Rollen mit Länge, Breite, Fläche, Fach
- **"Granulate":** Alle Granulate mit Gewichts-Informationen

### Ausgabe: Zwei separate Inventur-Dateien

#### 1. Inventur_Rollen.xlsx
- **"Inventur":** Gefundene Rollen mit Original- und kontrollierter Breite
- **"Nicht_gefunden":** Manuell erfasste, unbekannte Rollen

#### 2. Inventur_Granulat.xlsx
- **"Inventur":** Gefundene Granulate mit Soll- und Ist-Gewicht
- **"Nicht_gefunden":** Manuell erfasste, unbekannte Granulate

## 🔧 Erweiterte Funktionen

### 💾 Auto-Save
- Das Programm speichert automatisch nach jedem Scan
- Zusätzlich wird bei jeder Eingabe in Fach/Bemerkung gespeichert
- Bei Programmabsturz gehen keine Daten verloren

### 📊 Export-Funktion (V2)
- Klicken Sie "💾 Inventur exportieren"
- Erstellt **zwei Backup-Dateien** mit Zeitstempel:
  - `Inventur_Rollen_Backup_YYYYMMDD_HHMMSS.xlsx`
  - `Inventur_Granulat_Backup_YYYYMMDD_HHMMSS.xlsx`
- Originaldateien bleiben unverändert

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

## 🔄 Jährlicher Neustart (V2)

Zu Beginn einer neuen Inventur:

1. **Alte Inventur archivieren:**
   - Benennen Sie `data/Inventur_Rollen.xlsx` um (z.B. `Inventur_Rollen_2024.xlsx`)
   - Benennen Sie `data/Inventur_Granulat.xlsx` um (z.B. `Inventur_Granulat_2024.xlsx`)
   - Oder löschen Sie beide Dateien

2. **Neue Arbeitstabelle einsetzen:**
   - Ersetzen Sie `data/Arbeitstabelle.xlsx` mit der neuen Datei
   - **Wichtig:** Muss zwei Tabellenblätter haben ("Rollen" und "Granulate")

3. **Programm starten:**
   - Das Programm erstellt automatisch neue Excel-Dateien für beide Typen

## 🛠️ Fehlerbehebung

### Programm startet nicht
- Prüfen Sie, ob Python installiert ist: `python --version` in der Eingabeaufforderung
- Führen Sie `install_python.bat` erneut aus

### Arbeitstabelle nicht gefunden
- Stellen Sie sicher, dass `Arbeitstabelle.xlsx` im `data/` Ordner liegt
- **Wichtig:** Die Datei muss zwei Tabellenblätter haben: "Rollen" und "Granulate"
- Prüfen Sie die Spalten-Namen in beiden Tabellenblättern

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
inventur_programm_v2/
├── inventur_app.py          # Hauptprogramm V2
├── install_python.bat       # Python-Installation
├── start_inventur.bat       # Programm-Start
├── requirements.txt         # Python-Module
├── README.md               # Diese Dokumentation
├── data/                   # Daten-Verzeichnis
│   ├── Arbeitstabelle.xlsx # Lager-Datenbank (2 Blätter: Rollen + Granulate)
│   ├── Inventur_Rollen.xlsx    # Rollen-Inventur (automatisch)
│   ├── Inventur_Granulat.xlsx  # Granulat-Inventur (automatisch)
│   └── backups/            # Backup-Verzeichnis
└── config/                 # Konfiguration
    ├── settings.json       # Programmeinstellungen (erweitert)
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
  "farbe_rolle_bg": "#E3F2FD",
  "farbe_rolle_text": "#1976D2",
  "farbe_granulat_bg": "#FFF9C4",
  "farbe_granulat_text": "#F57F17",
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

### **Version 2.0** - Dezember 2024 🎆
- **🔵 Rollen-Unterstützung:** Separate Erfassung mit Fach + Breite kontrolliert
- **🟨 Granulat-Unterstützung:** Gewichts-Erfassung mit Zählmenge
- **📁 Zwei Excel-Dateien:** Inventur_Rollen.xlsx und Inventur_Granulat.xlsx
- **🎨 Visuelle Unterscheidung:** Blaue und gelbe Farbkodierung
- **📊 Erweiterte Statistiken:** Getrennte Zähler für beide Warentypen
- **⚙️ Intelligente Validierung:** Breite (1-4 Ziffern) und Gewicht (Dezimalzahlen)
- **📋 Dynamische UI:** Eingabefelder passen sich automatisch an den Warentyp an
- **📄 Erweiterte Spaltenstruktur:** Original- und kontrollierte Werte für Rollen

### **Version 1.1** - November 2024 ✨
- **🚫 Duplikat-Schutz:** Verhindert versehentliches Doppelscannen
- **🖥️ Vollbild-Modus:** Startet automatisch maximiert
- **🔢 Führende Nullen:** Vollständige Erhaltung aller Charge-Nummern
- **📊 Excel-Formatierung:** Professionelle Text-Spalten ohne Apostrophe
- **🎨 UI-Optimierung:** Entfernung des Suchfelds, verbessertes Dialog-Layout

### **Version 1.0** - November 2024
- Erste vollständige Version
- Alle Kernfunktionen implementiert
- Getestet für Windows 11

## 🎯 Roadmap

### Geplante Features:
- **Bearbeitungsfunktion:** Nachträgliche Änderung von Einträgen
- **Statistik-Dashboard:** Detaillierte Auswertungen nach Warentyp
- **Backup-Automatisierung:** Automatische tägliche Backups
- **Weitere Warentypen:** Erweiterung für zusätzliche Produktkategorien

---

**Entwickelt für Forbo Movement Systems**  
*Professionelle Lagerverwaltung mit Barcodescanner-Integration*

### 🏆 **VERSION 2.0 - PRODUKTIONSREIF!**
*Vollständige Unterstützung für Rollen und Granulate*

**✨ Neue Funktionen erfolgreich implementiert:**
- 🔵 Rollen mit Breiten-Kontrolle
- 🟨 Granulate mit Gewichts-Erfassung
- 📁 Separate Excel-Dateien
- 🎨 Visuelle Typ-Unterscheidung
