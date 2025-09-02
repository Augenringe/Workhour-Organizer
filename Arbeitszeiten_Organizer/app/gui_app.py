import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import os
import threading
import subprocess
import sys
from pathlib import Path

class ArbeitszeitenGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Arbeitszeiten-Verarbeitung - Agilos QCS GmbH")
        self.root.geometry("1300x750")
        self.root.minsize(1200, 700)
        self.root.configure(bg='#f0f0f0')
        
        # Stil für ttk
        style = ttk.Style()
        style.theme_use('clam')
        
        self.setup_ui()
        self.check_folders()
        
    def setup_ui(self):
        """Erstellt die Benutzeroberfläche"""
        # Hauptframe
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Titel
        title_label = ttk.Label(main_frame, text="Arbeitszeiten-Verarbeitung", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        subtitle_label = ttk.Label(main_frame, text="Agilos QCS GmbH - Professionelle Excel-Generierung", 
                                  font=('Arial', 10))
        subtitle_label.grid(row=1, column=0, columnspan=3, pady=(0, 30))
        
        # Ordner-Status
        self.setup_folder_status(main_frame)
        
        # CSV-Dateien Bereich
        self.setup_csv_section(main_frame)
        
        # Verarbeitungsbereich
        self.setup_processing_section(main_frame)
        
        # Log-Bereich
        self.setup_log_section(main_frame)
        
        # Buttons
        self.setup_buttons(main_frame)
        
        # Grid-Konfiguration
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
    def setup_folder_status(self, parent):
        """Ordner-Status Anzeige"""
        folder_frame = ttk.LabelFrame(parent, text="Ordner-Status", padding="10")
        folder_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # CSV Eingabe
        ttk.Label(folder_frame, text="CSV Eingabe:").grid(row=0, column=0, sticky=tk.W)
        self.csv_status = ttk.Label(folder_frame, text="❌ Nicht gefunden", foreground="red")
        self.csv_status.grid(row=0, column=1, sticky=tk.W, padx=(10, 0))
        
        # Excel Ausgabe
        ttk.Label(folder_frame, text="Excel Ausgabe:").grid(row=1, column=0, sticky=tk.W)
        self.excel_status = ttk.Label(folder_frame, text="❌ Nicht gefunden", foreground="red")
        self.excel_status.grid(row=1, column=1, sticky=tk.W, padx=(10, 0))
        
        # CSV Archiv
        ttk.Label(folder_frame, text="CSV Archiv:").grid(row=2, column=0, sticky=tk.W)
        self.archive_status = ttk.Label(folder_frame, text="❌ Nicht gefunden", foreground="red")
        self.archive_status.grid(row=2, column=1, sticky=tk.W, padx=(10, 0))
        
    def setup_csv_section(self, parent):
        """CSV-Dateien Bereich"""
        csv_frame = ttk.LabelFrame(parent, text="CSV-Dateien", padding="10")
        csv_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # CSV-Dateien Liste
        ttk.Label(csv_frame, text="Verfügbare CSV-Dateien:").grid(row=0, column=0, sticky=tk.W)
        
        # Listbox mit Scrollbar
        list_frame = ttk.Frame(csv_frame)
        list_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(5, 0))
        
        self.csv_listbox = tk.Listbox(list_frame, height=4, width=80, selectmode=tk.MULTIPLE)
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.csv_listbox.yview)
        self.csv_listbox.configure(yscrollcommand=scrollbar.set)
        
        self.csv_listbox.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        list_frame.columnconfigure(0, weight=1)
        
        # Refresh Button
        ttk.Button(csv_frame, text="🔄 Aktualisieren", 
                  command=self.refresh_csv_files).grid(row=2, column=0, pady=(10, 0))
        
    def setup_processing_section(self, parent):
        """Verarbeitungsbereich"""
        process_frame = ttk.LabelFrame(parent, text="Verarbeitung", padding="10")
        process_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 20))
        
        # Fortschrittsbalken
        ttk.Label(process_frame, text="Fortschritt:").grid(row=0, column=0, sticky=tk.W)
        self.progress = ttk.Progressbar(process_frame, mode='indeterminate')
        self.progress.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(10, 0))
        
        # Status
        self.status_label = ttk.Label(process_frame, text="Bereit")
        self.status_label.grid(row=1, column=0, columnspan=2, pady=(5, 0))
        
        process_frame.columnconfigure(1, weight=1)
        
    def setup_log_section(self, parent):
        """Log-Bereich"""
        log_frame = ttk.LabelFrame(parent, text="Verarbeitungslog", padding="10")
        log_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 20))
        
        self.log_text = scrolledtext.ScrolledText(log_frame, height=6, width=120)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        
    def setup_buttons(self, parent):
        """Buttons"""
        button_frame = ttk.Frame(parent)
        button_frame.grid(row=6, column=0, columnspan=3, pady=(0, 20))
        
        # Verarbeitungsbuttons
        self.process_selected_btn = ttk.Button(button_frame, text="📊 Ausgewählte verarbeiten", 
                                              command=self.process_selected, state='disabled')
        self.process_selected_btn.grid(row=0, column=0, padx=(0, 10))
        
        self.process_all_btn = ttk.Button(button_frame, text="📈 Alle verarbeiten", 
                                         command=self.process_all, state='disabled')
        self.process_all_btn.grid(row=0, column=1, padx=(0, 10))
        
        # Utility Buttons
        ttk.Button(button_frame, text="📁 Ordner öffnen", 
                  command=self.open_folders).grid(row=0, column=2, padx=(0, 10))
        
        ttk.Button(button_frame, text="❓ Hilfe", 
                  command=self.show_help).grid(row=0, column=3)
        
    def check_folders(self):
        """Überprüft und erstellt Ordner"""
        base_path = Path.cwd().parent
        
        folders = {
            'csv': base_path / "CSV_Eingabe",
            'excel': base_path / "Excel_Ausgabe", 
            'archive': base_path / "CSV_Archiv"
        }
        
        for folder_type, folder_path in folders.items():
            if not folder_path.exists():
                try:
                    folder_path.mkdir()
                    self.log(f"✅ Ordner erstellt: {folder_path.name}")
                except Exception as e:
                    self.log(f"❌ Fehler beim Erstellen von {folder_path.name}: {e}")
        
        # Status aktualisieren
        self.update_folder_status()
        
        self.refresh_csv_files()
        
    def update_folder_status(self):
        """Aktualisiert die Ordner-Status Anzeige"""
        # Bestimme den Basis-Ordner relativ zum Skript
        script_dir = Path(__file__).parent
        base_path = script_dir.parent
        
        # CSV Eingabe Status
        csv_folder = base_path / "CSV_Eingabe"
        if csv_folder.exists():
            csv_files = list(csv_folder.glob("*.csv"))
            if csv_files:
                self.csv_status.config(text="✅ Bereit", foreground="green")
            else:
                self.csv_status.config(text="❌ Keine Dateien", foreground="red")
        else:
            self.csv_status.config(text="❌ Fehler", foreground="red")
        
        # Excel Ausgabe Status
        excel_folder = base_path / "Excel_Ausgabe"
        self.excel_status.config(text="✅ Bereit" if excel_folder.exists() else "❌ Fehler",
                                foreground="green" if excel_folder.exists() else "red")
        
        # CSV Archiv Status
        archive_folder = base_path / "CSV_Archiv"
        self.archive_status.config(text="✅ Bereit" if archive_folder.exists() else "❌ Fehler",
                                  foreground="green" if archive_folder.exists() else "red")
        
    def refresh_csv_files(self):
        """Aktualisiert die CSV-Dateien Liste"""
        self.csv_listbox.delete(0, tk.END)
        
        # Bestimme den CSV-Eingabe-Ordner relativ zum Skript
        script_dir = Path(__file__).parent
        base_path = script_dir.parent
        csv_folder = base_path / "CSV_Eingabe"
        if csv_folder.exists():
            csv_files = list(csv_folder.glob("*.csv"))
            if csv_files:
                for csv_file in csv_files:
                    self.csv_listbox.insert(tk.END, csv_file.name)
                self.log(f"📁 {len(csv_files)} CSV-Dateien gefunden")
                
                # Buttons aktivieren
                self.process_selected_btn.config(state='normal')
                self.process_all_btn.config(state='normal')
            else:
                # Hilfreiche Meldung in der Listbox anzeigen
                self.csv_listbox.insert(tk.END, "📁 Keine CSV-Dateien gefunden!")
                self.csv_listbox.insert(tk.END, "")
                self.csv_listbox.insert(tk.END, "Bitte legen Sie CSV-Dateien in den")
                self.csv_listbox.insert(tk.END, "CSV_Eingabe-Ordner ab.")
                self.csv_listbox.insert(tk.END, "")
                self.csv_listbox.insert(tk.END, "Unterstützte Formate:")
                self.csv_listbox.insert(tk.END, "• Juli_2025.csv")
                self.csv_listbox.insert(tk.END, "• 08_2025.csv")
                self.csv_listbox.insert(tk.END, "• August_2025.csv")
                
                self.log("📁 Keine CSV-Dateien im CSV_Eingabe-Ordner gefunden")
                self.process_selected_btn.config(state='disabled')
                self.process_all_btn.config(state='disabled')
        else:
            self.log("❌ CSV_Eingabe-Ordner nicht gefunden")
        
        # Status aktualisieren
        self.update_folder_status()
            
    def process_selected(self):
        """Verarbeitet ausgewählte CSV-Dateien"""
        selected_indices = self.csv_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("Keine Auswahl", "Bitte wählen Sie mindestens eine CSV-Datei aus.")
            return
            
        # Bestimme den CSV-Eingabe-Ordner
        script_dir = Path(__file__).parent
        base_path = script_dir.parent
        csv_folder = base_path / "CSV_Eingabe"
        selected_files = []
        
        # Nur echte CSV-Dateien verarbeiten (Hilfsmeldungen ignorieren)
        for i in selected_indices:
            filename = self.csv_listbox.get(i)
            if filename.endswith('.csv') and not filename.startswith('📁') and not filename.startswith('•') and filename.strip():
                selected_files.append(filename)
        
        if not selected_files:
            messagebox.showwarning("Keine CSV-Dateien", "Bitte wählen Sie gültige CSV-Dateien aus.")
            return
        
        self.log(f"🚀 Starte Verarbeitung von {len(selected_files)} Dateien...")
        self.start_processing(selected_files)
        
    def process_all(self):
        """Verarbeitet alle CSV-Dateien"""
        # Bestimme den CSV-Eingabe-Ordner
        script_dir = Path(__file__).parent
        base_path = script_dir.parent
        csv_folder = base_path / "CSV_Eingabe"
        if csv_folder.exists():
            csv_files = list(csv_folder.glob("*.csv"))
            if csv_files:
                selected_files = [f.name for f in csv_files]
                self.log(f"🚀 Starte Verarbeitung aller {len(selected_files)} Dateien...")
                self.start_processing(selected_files)
            else:
                messagebox.showinfo("Keine Dateien", "Keine CSV-Dateien gefunden.")
        else:
            messagebox.showerror("Fehler", "CSV_Eingabe-Ordner nicht gefunden.")
            
    def start_processing(self, files):
        """Startet die Verarbeitung in einem separaten Thread"""
        self.progress.start()
        self.status_label.config(text="Verarbeitung läuft...")
        self.process_selected_btn.config(state='disabled')
        self.process_all_btn.config(state='disabled')
        
        # Thread für Verarbeitung
        thread = threading.Thread(target=self.process_files_thread, args=(files,))
        thread.daemon = True
        thread.start()
        
    def process_files_thread(self, files):
        """Verarbeitet Dateien in separatem Thread"""
        try:
            # Bestimme den CSV-Eingabe-Ordner
        script_dir = Path(__file__).parent
        base_path = script_dir.parent
        csv_folder = base_path / "CSV_Eingabe"
            
            for i, filename in enumerate(files):
                file_path = csv_folder / filename
                self.log(f"📊 Verarbeite: {filename}")
                
                # Python-Script aufrufen
                result = subprocess.run([
                    sys.executable, 
                    "simple_excel_processor.py", 
                    str(file_path)
                ], capture_output=True, text=True, cwd=Path.cwd())
                
                if result.returncode == 0:
                    self.log(f"✅ {filename} erfolgreich verarbeitet")
                else:
                    self.log(f"❌ Fehler bei {filename}: {result.stderr}")
                    
            self.log("🎉 Alle Dateien verarbeitet!")
            
        except Exception as e:
            self.log(f"❌ Unerwarteter Fehler: {e}")
        finally:
            # UI zurücksetzen
            self.root.after(0, self.processing_finished)
            
    def processing_finished(self):
        """Wird aufgerufen wenn Verarbeitung abgeschlossen ist"""
        self.progress.stop()
        self.status_label.config(text="Verarbeitung abgeschlossen")
        self.process_selected_btn.config(state='normal')
        self.process_all_btn.config(state='normal')
        self.refresh_csv_files()
        
    def open_folders(self):
        """Öffnet die Ordner im Explorer"""
        script_dir = Path(__file__).parent
        base_path = script_dir.parent
        
        folders = {
            "CSV Eingabe": base_path / "CSV_Eingabe",
            "Excel Ausgabe": base_path / "Excel_Ausgabe",
            "CSV Archiv": base_path / "CSV_Archiv"
        }
        
        for name, folder_path in folders.items():
            if folder_path.exists():
                try:
                    if sys.platform == "win32":
                        os.startfile(folder_path)
                    elif sys.platform == "darwin":
                        subprocess.run(["open", folder_path])
                    else:
                        subprocess.run(["xdg-open", folder_path])
                    self.log(f"📁 {name}-Ordner geöffnet")
                except Exception as e:
                    self.log(f"❌ Fehler beim Öffnen von {name}: {e}")
            else:
                self.log(f"❌ {name}-Ordner nicht gefunden")
                
    def show_help(self):
        """Zeigt Hilfe-Dialog"""
        help_text = """
Arbeitszeiten-Verarbeitung - Hilfe

📁 Ordnerstruktur:
• CSV_Eingabe: Legen Sie hier Ihre CSV-Dateien ab
• Excel_Ausgabe: Hier finden Sie die erstellten Excel-Dateien
• CSV_Archiv: Verarbeitete CSV-Dateien werden hier archiviert

🚀 Verarbeitung:
• Wählen Sie CSV-Dateien aus und klicken Sie "Ausgewählte verarbeiten"
• Oder klicken Sie "Alle verarbeiten" für alle CSV-Dateien
• Das Programm erkennt automatisch den Monat aus dem Dateinamen

📊 Excel-Ausgabe:
• Rohdaten: Alle ursprünglichen Daten
• Mitarbeiter-Übersicht: Summen pro Mitarbeiter
• Tages-Übersicht: Statistiken pro Tag
• Monats-Übersicht: Gesamtkennzahlen
• Individuelle Mitarbeiter-Blätter: Ein Blatt pro Mitarbeiter

💡 Tipps:
• Unterstützte Dateinamen: Juli_2025, 08_2025, August_2025, etc.
• CSV-Dateien werden nach Verarbeitung automatisch archiviert
• Nutzen Sie "Ordner öffnen" für schnellen Zugriff
        """
        
        messagebox.showinfo("Hilfe", help_text)
        
    def log(self, message):
        """Fügt Nachricht zum Log hinzu"""
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        
        self.log_text.insert(tk.END, log_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

def main():
    root = tk.Tk()
    app = ArbeitszeitenGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
