#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INVENTUR-PROGRAMM FÜR FORBO MOVEMENT SYSTEMS
===========================================

Desktop-Anwendung für Lagerpflege/Inventur mit Barcodescanner-Integration
Entwickelt für Windows 11, Python 3.11+

Autor: Cascade AI
Datum: November 2024
Version: 1.0
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from openpyxl import Workbook, load_workbook
from datetime import datetime
import os
import sys
from pathlib import Path
import json
import logging

class InventurApp:
    def __init__(self):
        """Initialisiert die Inventur-Anwendung"""
        self.root = tk.Tk()
        self.setup_logging()
        self.setup_paths()
        self.load_config()
        self.init_data()
        self.setup_ui()
        self.load_existing_inventur()
        self.bind_shortcuts()
        
    def setup_logging(self):
        """Konfiguriert das Logging-System"""
        if not os.path.exists('config'):
            os.makedirs('config')
        
        logging.basicConfig(
            filename='config/inventur.log',
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%d.%m.%Y %H:%M:%S'
        )
        self.logger = logging.getLogger(__name__)
        self.logger.info("Inventur-Programm gestartet")
    
    def setup_paths(self):
        """Definiert die Dateipfade"""
        self.data_dir = Path('data')
        self.config_dir = Path('config')
        self.arbeitstabelle_path = self.data_dir / 'Arbeitstabelle.xlsx'
        self.inventur_path = self.data_dir / 'Inventur.xlsx'
        
        # Erstelle Verzeichnisse falls nicht vorhanden
        self.data_dir.mkdir(exist_ok=True)
        self.config_dir.mkdir(exist_ok=True)
    
    def load_config(self):
        """Lädt die Konfigurationsdatei"""
        config_path = self.config_dir / 'settings.json'
        default_config = {
            "auto_save": True,
            "schriftgroesse": 12,
            "farbe_gefunden": "#E8F5E8",
            "farbe_nicht_gefunden": "#FFF2CC",
            "farbe_fehler": "#FFE6E6",
            "vollbild": True
        }
        
        try:
            if config_path.exists():
                with open(config_path, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                self.config = default_config
                self.save_config()
        except Exception as e:
            self.config = default_config
            self.logger.error(f"Fehler beim Laden der Konfiguration: {e}")
    
    def save_config(self):
        """Speichert die Konfiguration"""
        try:
            config_path = self.config_dir / 'settings.json'
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            self.logger.error(f"Fehler beim Speichern der Konfiguration: {e}")
    
    def init_data(self):
        """Initialisiert die Datenstrukturen"""
        self.df_arbeitstabelle = None
        self.inventur_data = []
        self.nicht_gefunden_data = []
        self.current_scan = None
        self.undo_stack = []
        
        # Lade Arbeitstabelle
        self.load_arbeitstabelle()
    
    def load_arbeitstabelle(self):
        """Lädt die Arbeitstabelle in ein DataFrame"""
        try:
            if self.arbeitstabelle_path.exists():
                # Lade Excel mit Charge als Text (behält führende Nullen)
                self.df_arbeitstabelle = pd.read_excel(self.arbeitstabelle_path, dtype={'Charge': str})
                
                # Stelle sicher, dass Charge-Spalte als String formatiert ist
                if 'Charge' in self.df_arbeitstabelle.columns:
                    self.df_arbeitstabelle['Charge'] = self.df_arbeitstabelle['Charge'].astype(str)
                
                self.logger.info(f"Arbeitstabelle geladen: {len(self.df_arbeitstabelle)} Artikel")
                
                # Prüfe erforderliche Spalten
                required_columns = ['Charge', 'Material', 'Materialkurztext', 'Länge m', 'Breite mm', 'Frei verwendbar']
                missing_columns = [col for col in required_columns if col not in self.df_arbeitstabelle.columns]
                
                if missing_columns:
                    messagebox.showerror("Fehler", f"Fehlende Spalten in Arbeitstabelle: {', '.join(missing_columns)}")
                    sys.exit(1)
                    
            else:
                messagebox.showwarning("Warnung", 
                    f"Arbeitstabelle nicht gefunden!\n\n"
                    f"Bitte kopieren Sie die Arbeitstabelle.xlsx in:\n"
                    f"{self.arbeitstabelle_path.absolute()}")
                
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der Arbeitstabelle:\n{e}")
            self.logger.error(f"Fehler beim Laden der Arbeitstabelle: {e}")
    
    def setup_ui(self):
        """Erstellt die Benutzeroberfläche"""
        # Hauptfenster konfigurieren
        self.root.title("INVENTUR Forbo - Lagerverwaltung")
        self.root.geometry("1200x800")
        self.root.configure(bg='#f0f0f0')
        
        # Icon setzen (falls vorhanden)
        try:
            self.root.iconbitmap('config/forbo_icon.ico')
        except:
            pass
        
        # Vollbild-Modus - standardmäßig aktiviert für Inventur-Anwendung
        self.root.state('zoomed')  # Maximiert das Fenster
        # Optional: Echter Vollbild-Modus (ohne Taskleiste)
        # self.root.attributes('-fullscreen', True)
        
        # Hauptframe
        self.main_frame = ttk.Frame(self.root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Grid-Konfiguration
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(0, weight=1)
        
        self.create_header()
        self.create_scan_section()
        self.create_current_scan_section()
        self.create_list_section()
        self.create_button_section()
        
        # Status-Bar
        self.create_status_bar()
        
        # Fokus auf Scan-Eingabefeld
        self.scan_entry.focus_set()
    
    def create_header(self):
        """Erstellt den Header-Bereich"""
        header_frame = ttk.Frame(self.main_frame)
        header_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 20))
        header_frame.columnconfigure(1, weight=1)
        
        # Logo-Platzhalter (links)
        logo_label = ttk.Label(header_frame, text="📦", font=("Arial", 24))
        logo_label.grid(row=0, column=0, padx=(0, 20))
        
        # Titel (mitte)
        title_label = ttk.Label(header_frame, 
                               text="INVENTUR Forbo - Lagerverwaltung",
                               font=("Arial", 18, "bold"),
                               foreground="#1f4e79")
        title_label.grid(row=0, column=1)
        
        # Info (rechts)
        if self.df_arbeitstabelle is not None:
            info_text = f"Artikel in DB: {len(self.df_arbeitstabelle)}"
        else:
            info_text = "Keine Arbeitstabelle"
        
        info_label = ttk.Label(header_frame, text=info_text, font=("Arial", 10))
        info_label.grid(row=0, column=2, padx=(20, 0))
    
    def create_scan_section(self):
        """Erstellt den Barcode-Scan-Bereich"""
        # Scan-Frame
        scan_frame = ttk.LabelFrame(self.main_frame, text="🔍 BARCODE SCANNEN", padding="10")
        scan_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        scan_frame.columnconfigure(0, weight=1)
        
        # Scan-Eingabefeld
        self.scan_var = tk.StringVar()
        self.scan_entry = ttk.Entry(scan_frame, 
                                   textvariable=self.scan_var,
                                   font=("Arial", 14),
                                   width=50)
        self.scan_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        # Scan-Button (falls Scanner kein Enter sendet)
        scan_button = ttk.Button(scan_frame, text="Scannen", command=self.process_scan)
        scan_button.grid(row=0, column=1)
        
        # Enter-Binding
        self.scan_entry.bind('<Return>', lambda e: self.process_scan())
        
        # Info-Label
        info_label = ttk.Label(scan_frame, 
                              text="Eingabefeld ist immer fokussiert. Barcode scannen oder eingeben + ENTER drücken.",
                              font=("Arial", 9),
                              foreground="gray")
        info_label.grid(row=1, column=0, columnspan=2, pady=(5, 0))
    
    def create_current_scan_section(self):
        """Erstellt den Bereich für den aktuellen Scan"""
        # Current-Scan-Frame
        self.current_frame = ttk.LabelFrame(self.main_frame, text="✅ AKTUELLER SCAN", padding="10")
        self.current_frame.grid(row=2, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        self.current_frame.columnconfigure(1, weight=1)
        
        # Datenfelder
        self.create_data_labels()
        
        # Eingabefelder
        self.create_input_fields()
        
        # Verstecke initial
        self.current_frame.grid_remove()
    
    def create_data_labels(self):
        """Erstellt die Datenanzeigefelder"""
        row = 0
        
        # Charge
        ttk.Label(self.current_frame, text="Charge:", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=2)
        self.charge_label = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.charge_label.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        row += 1
        
        # Material
        ttk.Label(self.current_frame, text="Material:", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=2)
        self.material_label = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.material_label.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        row += 1
        
        # Materialkurztext
        ttk.Label(self.current_frame, text="Materialkurztext:", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=2)
        self.kurztext_label = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.kurztext_label.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        row += 1
        
        # Länge
        ttk.Label(self.current_frame, text="Länge:", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=2)
        self.laenge_label = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.laenge_label.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        row += 1
        
        # Breite
        ttk.Label(self.current_frame, text="Breite:", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=2)
        self.breite_label = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.breite_label.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        row += 1
        
        # Fläche
        ttk.Label(self.current_frame, text="Fläche:", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=2)
        self.flaeche_label = ttk.Label(self.current_frame, text="", font=("Arial", 11))
        self.flaeche_label.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=2)
        row += 1
        
        # Trennlinie
        separator = ttk.Separator(self.current_frame, orient='horizontal')
        separator.grid(row=row, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        row += 1
        
        self.input_start_row = row
    
    def create_input_fields(self):
        """Erstellt die Eingabefelder für Fach und Bemerkung"""
        row = self.input_start_row
        
        # Fach-Eingabe
        ttk.Label(self.current_frame, text="🏷️ Fach (Lagerort):", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=5)
        self.fach_var = tk.StringVar()
        self.fach_entry = ttk.Entry(self.current_frame, 
                                   textvariable=self.fach_var,
                                   font=("Arial", 12),
                                   width=20)
        self.fach_entry.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        self.fach_entry.bind('<Return>', self.save_current_scan)
        self.fach_entry.bind('<KeyRelease>', self.on_field_change)
        row += 1
        
        # Bemerkung-Eingabe
        ttk.Label(self.current_frame, text="📝 Bemerkung (optional):", font=("Arial", 11, "bold")).grid(row=row, column=0, sticky=tk.W, pady=5)
        self.bemerkung_var = tk.StringVar()
        self.bemerkung_entry = ttk.Entry(self.current_frame, 
                                        textvariable=self.bemerkung_var,
                                        font=("Arial", 12),
                                        width=50)
        self.bemerkung_entry.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=5)
        self.bemerkung_entry.bind('<Return>', self.save_current_scan)
        self.bemerkung_entry.bind('<KeyRelease>', self.on_field_change)
        row += 1
        
        # Speichern-Button
        save_button = ttk.Button(self.current_frame, text="💾 Speichern", command=self.save_current_scan)
        save_button.grid(row=row, column=1, sticky=tk.W, padx=(10, 0), pady=10)
    
    def create_list_section(self):
        """Erstellt die Liste der gescannten Artikel"""
        # List-Frame
        list_frame = ttk.LabelFrame(self.main_frame, text="📋 GESCANNTE ARTIKEL", padding="10")
        list_frame.grid(row=3, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(1, weight=1)
        self.main_frame.rowconfigure(3, weight=1)
        
        # Artikel-Anzahl (ohne Suchbereich)
        count_frame = ttk.Frame(list_frame)
        count_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        
        self.count_label = ttk.Label(count_frame, text="0 Artikel", font=("Arial", 12, "bold"))
        self.count_label.pack(side=tk.LEFT)
        
        # Treeview für Artikelliste
        columns = ('Zeit', 'Charge', 'Material', 'Kurztext', 'Fach', 'Bemerkung', 'Status')
        self.tree = ttk.Treeview(list_frame, columns=columns, show='headings', height=15)
        
        # Spalten konfigurieren
        self.tree.heading('Zeit', text='Zeit')
        self.tree.heading('Charge', text='Charge')
        self.tree.heading('Material', text='Material')
        self.tree.heading('Kurztext', text='Materialkurztext')
        self.tree.heading('Fach', text='Fach')
        self.tree.heading('Bemerkung', text='Bemerkung')
        self.tree.heading('Status', text='Status')
        
        # Spaltenbreiten
        self.tree.column('Zeit', width=120, minwidth=100)
        self.tree.column('Charge', width=100, minwidth=80)
        self.tree.column('Material', width=100, minwidth=80)
        self.tree.column('Kurztext', width=200, minwidth=150)
        self.tree.column('Fach', width=80, minwidth=60)
        self.tree.column('Bemerkung', width=150, minwidth=100)
        self.tree.column('Status', width=100, minwidth=80)
        
        self.tree.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        v_scrollbar.grid(row=1, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=v_scrollbar.set)
        
        h_scrollbar = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        h_scrollbar.grid(row=2, column=0, sticky=(tk.W, tk.E))
        self.tree.configure(xscrollcommand=h_scrollbar.set)
        
        # Kontextmenü
        self.tree.bind('<Button-3>', self.show_context_menu)
        self.tree.bind('<Double-1>', self.edit_entry)
    
    def create_button_section(self):
        """Erstellt die Button-Leiste"""
        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=4, column=0, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Export-Button
        export_button = ttk.Button(button_frame, text="💾 Inventur exportieren", command=self.export_inventur)
        export_button.grid(row=0, column=0, padx=(0, 10))
        
        # Vollbild-Toggle
        fullscreen_button = ttk.Button(button_frame, text="🖥️ Vollbild", command=self.toggle_fullscreen)
        fullscreen_button.grid(row=0, column=1, padx=(0, 10))
        
        # Beenden-Button
        exit_button = ttk.Button(button_frame, text="❌ Programm beenden", command=self.quit_app)
        exit_button.grid(row=0, column=2)
    
    def create_status_bar(self):
        """Erstellt die Status-Leiste"""
        self.status_var = tk.StringVar()
        self.status_var.set("Bereit zum Scannen...")
        
        status_bar = ttk.Label(self.root, textvariable=self.status_var, 
                              relief=tk.SUNKEN, anchor=tk.W, font=("Arial", 9))
        status_bar.grid(row=1, column=0, sticky=(tk.W, tk.E))
    
    def bind_shortcuts(self):
        """Bindet Tastenkürzel"""
        self.root.bind('<Control-z>', lambda e: self.undo_last_action())
        self.root.bind('<Control-Z>', lambda e: self.undo_last_action())
        self.root.bind('<Control-s>', lambda e: self.manual_save())
        self.root.bind('<Control-S>', lambda e: self.manual_save())
        self.root.bind('<Escape>', lambda e: self.reset_scan())
        self.root.bind('<F11>', lambda e: self.toggle_fullscreen())
        
        # Fokus immer zurück zum Scan-Feld
        self.root.bind('<FocusIn>', self.ensure_scan_focus)
    
    def process_scan(self):
        """Verarbeitet einen gescannten Barcode"""
        charge = self.scan_var.get().strip()
        
        if not charge:
            self.status_var.set("Bitte Barcode eingeben oder scannen")
            return
        
        # Prüfe ob Charge bereits gescannt wurde
        if self.is_already_scanned(charge):
            messagebox.showwarning("Bereits gescannt", 
                f"Die Ware mit Charge {charge} wurde bereits eingescannt!\n\n"
                f"Bitte prüfen Sie die Liste der gescannten Artikel.")
            self.reset_scan()
            return
        
        self.logger.info(f"Scan verarbeitet: {charge}")
        
        # Suche in Arbeitstabelle
        if self.df_arbeitstabelle is not None:
            try:
                # Suche nach Charge als String (behält führende Nullen)
                result = self.df_arbeitstabelle[self.df_arbeitstabelle['Charge'] == charge]
                
                # Falls nicht gefunden, versuche auch ohne führende Nullen
                if result.empty:
                    try:
                        charge_int = str(int(charge))  # Entfernt führende Nullen
                        result = self.df_arbeitstabelle[self.df_arbeitstabelle['Charge'] == charge_int]
                    except ValueError:
                        pass
                
                if not result.empty:
                    # Ware gefunden
                    self.show_found_item(result.iloc[0], charge)
                else:
                    # Ware nicht gefunden
                    self.show_not_found_dialog(charge)
                    
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler bei der Suche: {e}")
                self.logger.error(f"Fehler bei Suche: {e}")
                self.reset_scan()
        else:
            messagebox.showerror("Fehler", "Keine Arbeitstabelle geladen")
    
    def is_already_scanned(self, charge):
        """Prüft ob eine Charge bereits gescannt wurde"""
        # Prüfe in beiden Listen
        for item in self.inventur_data:
            if str(item.get('charge', '')) == str(charge):
                return True
        
        for item in self.nicht_gefunden_data:
            if str(item.get('charge', '')) == str(charge):
                return True
        
        return False
    
    def show_found_item(self, item, charge):
        """Zeigt gefundene Ware an"""
        self.current_scan = {
            'charge': charge,
            'material': str(item.get('Material', '')),
            'kurztext': str(item.get('Materialkurztext', '')),
            'laenge': float(item.get('Länge m', 0)),
            'breite': int(item.get('Breite mm', 0)),
            'flaeche': float(item.get('Frei verwendbar', 0)),
            'status': 'gefunden'
        }
        
        # Aktualisiere Labels
        self.charge_label.config(text=charge)
        self.material_label.config(text=self.current_scan['material'])
        self.kurztext_label.config(text=self.current_scan['kurztext'])
        self.laenge_label.config(text=f"{self.current_scan['laenge']:.2f} m")
        self.breite_label.config(text=f"{self.current_scan['breite']} mm")
        self.flaeche_label.config(text=f"{self.current_scan['flaeche']:.2f} m²")
        
        # Zeige Current-Scan-Frame
        self.current_frame.grid()
        
        # Fokus auf Fach-Eingabe
        self.fach_entry.focus_set()
        
        # Status aktualisieren
        self.status_var.set(f"Ware gefunden: {self.current_scan['kurztext']}")
        
        # Scan-Feld leeren
        self.scan_var.set("")
    
    def show_not_found_dialog(self, charge):
        """Zeigt Dialog für nicht gefundene Ware"""
        dialog = NotFoundDialog(self.root, charge)
        result = dialog.result
        
        if result:
            # Manuelle Eingabe wurde gespeichert
            self.current_scan = result
            self.current_scan['status'] = 'nicht_gefunden'
            
            # Speichere direkt (ohne Fach-Eingabe)
            self.save_scan_to_data()
            self.reset_scan()
        else:
            # Abgebrochen
            self.reset_scan()
    
    def save_current_scan(self, event=None):
        """Speichert den aktuellen Scan"""
        if not self.current_scan:
            return
        
        fach = self.fach_var.get().strip()
        if not fach:
            messagebox.showwarning("Warnung", "Bitte Fach-Nummer eingeben")
            self.fach_entry.focus_set()
            return
        
        # Füge Fach und Bemerkung hinzu
        self.current_scan['fach'] = fach
        self.current_scan['bemerkung'] = self.bemerkung_var.get().strip()
        
        # Speichere in Datenstrukturen
        self.save_scan_to_data()
        
        # Reset für nächsten Scan
        self.reset_scan()
    
    def save_scan_to_data(self):
        """Speichert Scan in Datenstrukturen und Excel"""
        if not self.current_scan:
            return
        
        # Zeitstempel hinzufügen
        self.current_scan['zeitstempel'] = datetime.now().strftime('%d.%m.%Y %H:%M:%S')
        
        # Zu entsprechender Liste hinzufügen
        if self.current_scan['status'] == 'gefunden':
            self.inventur_data.append(self.current_scan.copy())
        else:
            self.nicht_gefunden_data.append(self.current_scan.copy())
        
        # Zur Undo-Liste hinzufügen
        self.undo_stack.append(('add', self.current_scan.copy()))
        if len(self.undo_stack) > 50:  # Begrenze Undo-Stack
            self.undo_stack.pop(0)
        
        # In Excel speichern (Auto-Save)
        if self.config.get('auto_save', True):
            self.save_to_excel()
        
        # Liste aktualisieren
        self.update_list()
        
        # Status aktualisieren
        total_scans = len(self.inventur_data) + len(self.nicht_gefunden_data)
        self.status_var.set(f"Artikel gespeichert. Gesamt: {total_scans}")
        
        self.logger.info(f"Artikel gespeichert: {self.current_scan['charge']}")
    
    def reset_scan(self):
        """Setzt den aktuellen Scan zurück"""
        self.current_scan = None
        self.scan_var.set("")
        self.fach_var.set("")
        self.bemerkung_var.set("")
        
        # Verstecke Current-Scan-Frame
        self.current_frame.grid_remove()
        
        # Fokus zurück zu Scan-Eingabe
        self.scan_entry.focus_set()
        
        self.status_var.set("Bereit zum Scannen...")
    
    def on_field_change(self, event=None):
        """Wird aufgerufen wenn Fach oder Bemerkung geändert wird"""
        if self.config.get('auto_save', True) and self.current_scan:
            # Auto-Save bei Feldeingabe (mit kleiner Verzögerung)
            self.root.after(1000, self.delayed_save)
    
    def delayed_save(self):
        """Verzögertes Speichern für Auto-Save"""
        if self.current_scan and self.fach_var.get().strip():
            # Nur speichern wenn Fach ausgefüllt ist
            pass  # Wird durch save_current_scan() erledigt
    
    def update_list(self):
        """Aktualisiert die Artikelliste"""
        # Lösche alle Einträge
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Kombiniere beide Listen
        all_items = []
        for item in self.inventur_data:
            all_items.append((item, 'Gefunden'))
        for item in self.nicht_gefunden_data:
            all_items.append((item, 'Nicht gefunden'))
        
        # Sortiere nach Zeitstempel (neueste zuerst)
        all_items.sort(key=lambda x: x[0]['zeitstempel'], reverse=True)
        
        # Füge zur TreeView hinzu
        for item_data, status in all_items:
            # Kürze Kurztext falls zu lang
            kurztext = item_data['kurztext']
            if len(kurztext) > 30:
                kurztext = kurztext[:27] + "..."
            
            # Bereinige Bemerkung von "nan" Werten
            bemerkung = item_data.get('bemerkung', '')
            if bemerkung == 'nan' or str(bemerkung).lower() == 'nan':
                bemerkung = ''
            
            values = (
                item_data['zeitstempel'].split()[1],  # Nur Zeit anzeigen
                item_data['charge'],
                item_data['material'],
                kurztext,
                item_data.get('fach', ''),
                bemerkung,
                status
            )
            
            item_id = self.tree.insert('', 'end', values=values)
            
            # Farbcodierung
            if status == 'Nicht gefunden':
                self.tree.set(item_id, 'Status', '⚠️ Nicht gefunden')
            else:
                self.tree.set(item_id, 'Status', '✅ Gefunden')
        
        # Anzahl aktualisieren
        total = len(all_items)
        self.count_label.config(text=f"{total} Artikel")
    
    def save_to_excel(self):
        """Speichert Daten in Excel-Datei"""
        try:
            # Erstelle oder lade Workbook
            if self.inventur_path.exists():
                wb = load_workbook(self.inventur_path)
            else:
                wb = Workbook()
                # Entferne Standard-Sheet
                if 'Sheet' in wb.sheetnames:
                    wb.remove(wb['Sheet'])
            
            # Erstelle/aktualisiere Inventur-Sheet
            if 'Inventur' in wb.sheetnames:
                wb.remove(wb['Inventur'])
            
            ws_inventur = wb.create_sheet('Inventur')
            
            # Header für Inventur
            headers = ['Datum/Uhrzeit', 'Charge', 'Material', 'Materialkurztext', 
                      'Länge m', 'Breite mm', 'Fläche m²', 'Fach', 'Bemerkung']
            ws_inventur.append(headers)
            
            # Daten für Inventur
            for item in self.inventur_data:
                # Bereinige Bemerkung beim Speichern
                bemerkung = item.get('bemerkung', '')
                if bemerkung == 'nan' or str(bemerkung).lower() == 'nan':
                    bemerkung = ''
                
                row = [
                    item['zeitstempel'],
                    item['charge'],  # Ohne Apostroph - wird über Formatierung als Text gesetzt
                    item['material'],
                    item['kurztext'],
                    item['laenge'],
                    item['breite'],
                    item['flaeche'],
                    item.get('fach', ''),
                    bemerkung
                ]
                ws_inventur.append(row)
            
            # Formatiere Charge-Spalte als Text
            from openpyxl.styles import NamedStyle
            from openpyxl.utils import get_column_letter
            
            # Charge-Spalte (Spalte B) als Text formatieren
            charge_col = get_column_letter(2)  # Spalte B
            for row in range(2, len(self.inventur_data) + 2):  # Ab Zeile 2 (nach Header)
                cell = ws_inventur[f'{charge_col}{row}']
                cell.number_format = '@'  # Text-Format
            
            # Erstelle/aktualisiere Nicht_gefunden-Sheet
            if 'Nicht_gefunden' in wb.sheetnames:
                wb.remove(wb['Nicht_gefunden'])
            
            ws_nicht_gefunden = wb.create_sheet('Nicht_gefunden')
            ws_nicht_gefunden.append(headers)
            
            # Daten für Nicht_gefunden
            for item in self.nicht_gefunden_data:
                # Bereinige Bemerkung beim Speichern
                bemerkung = item.get('bemerkung', '')
                if bemerkung == 'nan' or str(bemerkung).lower() == 'nan':
                    bemerkung = ''
                
                row = [
                    item['zeitstempel'],
                    item['charge'],  # Ohne Apostroph - wird über Formatierung als Text gesetzt
                    item['material'],
                    item['kurztext'],
                    item['laenge'],
                    item['breite'],
                    item['flaeche'],
                    item.get('fach', ''),
                    bemerkung
                ]
                ws_nicht_gefunden.append(row)
            
            # Formatiere Charge-Spalte als Text auch im Nicht_gefunden-Sheet
            for row in range(2, len(self.nicht_gefunden_data) + 2):  # Ab Zeile 2 (nach Header)
                cell = ws_nicht_gefunden[f'{charge_col}{row}']
                cell.number_format = '@'  # Text-Format
            
            # Speichern
            wb.save(self.inventur_path)
            self.logger.info("Excel-Datei gespeichert")
            
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Speichern der Excel-Datei:\n{e}")
            self.logger.error(f"Fehler beim Speichern: {e}")
    
    def load_existing_inventur(self):
        """Lädt bestehende Inventur-Daten"""
        if not self.inventur_path.exists():
            return
        
        try:
            # Lade Inventur-Sheet mit Charge als String
            df_inventur = pd.read_excel(self.inventur_path, sheet_name='Inventur', dtype={'Charge': str})
            for _, row in df_inventur.iterrows():
                # Bereinige Bemerkung von nan-Werten
                bemerkung = row.get('Bemerkung', '')
                if pd.isna(bemerkung) or str(bemerkung).lower() == 'nan':
                    bemerkung = ''
                
                item = {
                    'zeitstempel': str(row.get('Datum/Uhrzeit', '')),
                    'charge': str(row.get('Charge', '')),
                    'material': str(row.get('Material', '')),
                    'kurztext': str(row.get('Materialkurztext', '')),
                    'laenge': float(row.get('Länge m', 0)),
                    'breite': int(row.get('Breite mm', 0)),
                    'flaeche': float(row.get('Fläche m²', 0)),
                    'fach': str(row.get('Fach', '')),
                    'bemerkung': str(bemerkung),
                    'status': 'gefunden'
                }
                self.inventur_data.append(item)
            
            # Lade Nicht_gefunden-Sheet mit Charge als String
            try:
                df_nicht_gefunden = pd.read_excel(self.inventur_path, sheet_name='Nicht_gefunden', dtype={'Charge': str})
                for _, row in df_nicht_gefunden.iterrows():
                    # Bereinige Bemerkung von nan-Werten
                    bemerkung = row.get('Bemerkung', '')
                    if pd.isna(bemerkung) or str(bemerkung).lower() == 'nan':
                        bemerkung = ''
                    
                    item = {
                        'zeitstempel': str(row.get('Datum/Uhrzeit', '')),
                        'charge': str(row.get('Charge', '')),
                        'material': str(row.get('Material', '')),
                        'kurztext': str(row.get('Materialkurztext', '')),
                        'laenge': float(row.get('Länge m', 0)),
                        'breite': int(row.get('Breite mm', 0)),
                        'flaeche': float(row.get('Fläche m²', 0)),
                        'fach': str(row.get('Fach', '')),
                        'bemerkung': str(bemerkung),
                        'status': 'nicht_gefunden'
                    }
                    self.nicht_gefunden_data.append(item)
            except:
                pass  # Sheet existiert noch nicht
            
            # Liste aktualisieren
            self.update_list()
            
            total = len(self.inventur_data) + len(self.nicht_gefunden_data)
            self.status_var.set(f"Bestehende Inventur geladen: {total} Artikel")
            self.logger.info(f"Bestehende Inventur geladen: {total} Artikel")
            
        except Exception as e:
            self.logger.error(f"Fehler beim Laden der bestehenden Inventur: {e}")
    
    def export_inventur(self):
        """Exportiert Inventur als Backup"""
        try:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            backup_filename = f"Inventur_Backup_{timestamp}.xlsx"
            
            backup_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx")],
                initialfile=backup_filename,
                title="Inventur-Backup speichern"
            )
            
            if backup_path:
                # Kopiere aktuelle Inventur-Datei
                if self.inventur_path.exists():
                    import shutil
                    shutil.copy2(self.inventur_path, backup_path)
                else:
                    # Erstelle neue Datei
                    self.save_to_excel()
                    shutil.copy2(self.inventur_path, backup_path)
                
                messagebox.showinfo("Export erfolgreich", 
                    f"Inventur-Backup gespeichert:\n{backup_path}")
                self.logger.info(f"Backup erstellt: {backup_path}")
                
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Export:\n{e}")
            self.logger.error(f"Fehler beim Export: {e}")
    
    def show_context_menu(self, event):
        """Zeigt Kontextmenü für Listeneinträge"""
        item = self.tree.selection()[0] if self.tree.selection() else None
        if not item:
            return
        
        context_menu = tk.Menu(self.root, tearoff=0)
        context_menu.add_command(label="🗑️ Löschen", command=lambda: self.delete_entry(item))
        context_menu.add_command(label="✏️ Bearbeiten", command=lambda: self.edit_entry())
        
        try:
            context_menu.tk_popup(event.x_root, event.y_root)
        finally:
            context_menu.grab_release()
    
    def delete_entry(self, item_id):
        """Löscht einen Eintrag"""
        if not messagebox.askyesno("Löschen bestätigen", 
                                  "Möchten Sie diesen Eintrag wirklich löschen?"):
            return
        
        # Finde Eintrag in Daten
        values = self.tree.item(item_id, 'values')
        if len(values) >= 2:
            charge = values[1]  # Charge ist in Spalte 1
            
            # Entferne aus entsprechender Liste
            self.inventur_data = [item for item in self.inventur_data if item['charge'] != charge]
            self.nicht_gefunden_data = [item for item in self.nicht_gefunden_data if item['charge'] != charge]
            
            # Speichere und aktualisiere
            self.save_to_excel()
            self.update_list()
            
            self.status_var.set("Eintrag gelöscht")
            self.logger.info(f"Eintrag gelöscht: {charge}")
    
    def edit_entry(self):
        """Bearbeitet einen Eintrag"""
        selection = self.tree.selection()
        if not selection:
            return
        
        messagebox.showinfo("Info", "Bearbeitungsfunktion wird in einer späteren Version implementiert.")
    
    def undo_last_action(self):
        """Macht die letzte Aktion rückgängig"""
        if not self.undo_stack:
            self.status_var.set("Nichts zum Rückgängigmachen")
            return
        
        action, data = self.undo_stack.pop()
        
        if action == 'add':
            # Entferne letzten Eintrag
            charge = data['charge']
            if data['status'] == 'gefunden':
                self.inventur_data = [item for item in self.inventur_data if item['charge'] != charge]
            else:
                self.nicht_gefunden_data = [item for item in self.nicht_gefunden_data if item['charge'] != charge]
            
            self.save_to_excel()
            self.update_list()
            self.status_var.set(f"Eintrag rückgängig gemacht: {charge}")
            self.logger.info(f"Undo: {charge}")
    
    def manual_save(self):
        """Manuelles Speichern"""
        self.save_to_excel()
        self.status_var.set("Manuell gespeichert")
    
    def toggle_fullscreen(self):
        """Schaltet Vollbild-Modus um"""
        current_state = self.root.attributes('-fullscreen')
        self.root.attributes('-fullscreen', not current_state)
        
        if not current_state:
            self.status_var.set("Vollbild-Modus aktiviert (F11 zum Beenden)")
        else:
            self.status_var.set("Vollbild-Modus deaktiviert")
    
    def ensure_scan_focus(self, event=None):
        """Stellt sicher, dass Scan-Feld fokussiert bleibt"""
        if hasattr(self, 'scan_entry') and not self.current_scan:
            self.root.after(100, lambda: self.scan_entry.focus_set())
    
    def quit_app(self):
        """Beendet die Anwendung"""
        if messagebox.askyesno("Beenden", "Möchten Sie das Programm wirklich beenden?"):
            self.logger.info("Programm beendet")
            self.root.quit()
    
    def run(self):
        """Startet die Anwendung"""
        self.root.mainloop()


class NotFoundDialog:
    """Dialog für nicht gefundene Waren"""
    
    def __init__(self, parent, charge):
        self.result = None
        
        # Dialog-Fenster
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("⚠️ Ware nicht gefunden")
        self.dialog.geometry("450x650")
        self.dialog.resizable(True, True)
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Zentriere Dialog
        self.dialog.geometry("+%d+%d" % (parent.winfo_rootx() + 50, parent.winfo_rooty() + 50))
        
        self.create_widgets(charge)
        
        # Warte auf Schließen
        self.dialog.wait_window()
    
    def create_widgets(self, charge):
        """Erstellt Dialog-Widgets"""
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Titel
        title_label = ttk.Label(main_frame, 
                               text=f"Ware mit Charge {charge} nicht gefunden!",
                               font=("Arial", 12, "bold"),
                               foreground="red")
        title_label.pack(pady=(0, 20))
        
        # Info
        info_label = ttk.Label(main_frame, 
                              text="Bitte geben Sie die Daten manuell ein:",
                              font=("Arial", 10))
        info_label.pack(pady=(0, 15))
        
        # Eingabefelder
        self.charge_var = tk.StringVar(value=charge)
        self.material_var = tk.StringVar()
        self.kurztext_var = tk.StringVar()
        self.laenge_var = tk.StringVar()
        self.breite_var = tk.StringVar()
        self.flaeche_var = tk.StringVar()
        self.fach_var = tk.StringVar()
        self.bemerkung_var = tk.StringVar()
        
        fields = [
            ("Charge:", self.charge_var, False),
            ("Material-Nummer:", self.material_var, True),
            ("Materialkurztext:", self.kurztext_var, False),
            ("Länge (m):", self.laenge_var, True),
            ("Breite (mm):", self.breite_var, True),
            ("Fläche (m²):", self.flaeche_var, True),
            ("Fach:", self.fach_var, True),
            ("Bemerkung:", self.bemerkung_var, False)
        ]
        
        for i, (label_text, var, required) in enumerate(fields):
            # Frame für jedes Feld
            field_frame = ttk.Frame(main_frame)
            field_frame.pack(fill=tk.X, pady=3)
            
            # Label über dem Eingabefeld
            display_text = label_text + " *" if required else label_text
            label = ttk.Label(field_frame, text=display_text, font=("Arial", 9))
            label.pack(anchor=tk.W)
            
            # Eingabefeld unter dem Label
            entry = ttk.Entry(field_frame, textvariable=var, width=50, font=("Arial", 10))
            entry.pack(fill=tk.X, pady=(1, 0))
            
            if i == 1:  # Fokus auf Material-Nummer
                entry.focus_set()
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(15, 10), fill=tk.X)
        
        # Zentriere Buttons
        inner_button_frame = ttk.Frame(button_frame)
        inner_button_frame.pack()
        
        save_button = ttk.Button(inner_button_frame, text="💾 Speichern", 
                                command=self.save_data, width=12)
        save_button.pack(side=tk.LEFT, padx=(0, 15))
        
        cancel_button = ttk.Button(inner_button_frame, text="❌ Abbrechen", 
                                  command=self.cancel, width=12)
        cancel_button.pack(side=tk.LEFT)
        
        # Enter-Binding
        self.dialog.bind('<Return>', lambda e: self.save_data())
        self.dialog.bind('<Escape>', lambda e: self.cancel())
    
    def save_data(self):
        """Speichert die eingegebenen Daten"""
        # Validierung
        required_fields = [
            (self.material_var.get().strip(), "Material-Nummer"),
            (self.laenge_var.get().strip(), "Länge"),
            (self.breite_var.get().strip(), "Breite"),
            (self.flaeche_var.get().strip(), "Fläche"),
            (self.fach_var.get().strip(), "Fach")
        ]
        
        for value, field_name in required_fields:
            if not value:
                messagebox.showerror("Fehler", f"Bitte {field_name} eingeben")
                return
        
        try:
            # Konvertiere numerische Werte
            laenge = float(self.laenge_var.get().replace(',', '.'))
            breite = int(self.breite_var.get())
            flaeche = float(self.flaeche_var.get().replace(',', '.'))
            
            self.result = {
                'charge': self.charge_var.get().strip(),
                'material': self.material_var.get().strip(),
                'kurztext': self.kurztext_var.get().strip(),
                'laenge': laenge,
                'breite': breite,
                'flaeche': flaeche,
                'fach': self.fach_var.get().strip(),
                'bemerkung': self.bemerkung_var.get().strip()
            }
            
            self.dialog.destroy()
            
        except ValueError:
            messagebox.showerror("Fehler", "Ungültige Zahlenwerte eingegeben")
    
    def cancel(self):
        """Bricht den Dialog ab"""
        self.result = None
        self.dialog.destroy()


def main():
    """Hauptfunktion"""
    try:
        app = InventurApp()
        app.run()
    except Exception as e:
        messagebox.showerror("Kritischer Fehler", f"Unerwarteter Fehler:\n{e}")
        logging.error(f"Kritischer Fehler: {e}")


if __name__ == "__main__":
    main()
