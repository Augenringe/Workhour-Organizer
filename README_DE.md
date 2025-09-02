# Arbeitszeiten-Verarbeitungssystem
## 🚀 Modulare Version - Komplett selbstständig!

## 📋 Was ist das?

Diese **professionelle Anwendung** verwandelt Ihre CSV-Dateien mit Arbeitszeiten automatisch in schöne Excel-Dateien mit Übersichten und Berechnungen. **Diese Version ist vollständig modular und funktioniert überall - einfach den ganzen Ordner `Working_Hours_Organizer` kopieren!**

## 🎯 Was bekommen Sie?

Eine professionelle Excel-Datei mit **mehreren Arbeitsblättern** für jeden Monat:

1. **📊 Rohdaten** - Alle Ihre ursprünglichen Daten, schön formatiert
2. **👥 Mitarbeiter-Übersicht** - Summen pro Mitarbeiter
3. **📅 Tages-Übersicht** - Statistiken pro Tag
4. **📈 Monats-Übersicht** - Gesamtkennzahlen für den Monat
5. **👤 Individuelle Mitarbeiter-Blätter** - Ein eigenes Arbeitsblatt für jeden Mitarbeiter

## 📁 Modulare Ordnerstruktur

Das Programm ist vollständig modular und erstellt automatisch eine saubere Ordnerstruktur:

```
Working_Hours_Organizer/       ← Hauptordner (kann überall hin kopiert werden!)
├── CSV_Input/                 ← **HIER CSV-DATEIEN ABLEGEN**
│   ├── Company_July_2025.csv
│   ├── Working_Hours_August_2025.csv
│   └── Time_Tracking_September_2025.csv
├── Excel_Output/              ← **HIER FINDEN SIE IHRE EXCEL-DATEIEN**
│   ├── Working_Hours_Analysis_July_2025.xlsx
│   ├── Working_Hours_Analysis_August_2025.xlsx
│   └── Working_Hours_Analysis_September_2025.xlsx
├── CSV_Archive/               ← **VERARBEITETE CSV-DATEIEN LANDEN HIER**
│   ├── Company_July_2025.csv
│   └── Working_Hours_August_2025.csv
├── app/                       ← Anwendungsordner
│   ├── simple_excel_processor.py
│   ├── gui_app.py
│   └── requirements_excel.txt
├── start_excel.bat            ← **STARTDATEI - HIER KLICKEN!**
└── README.md                  ← Diese Anleitung
```

## 🚀 So einfach geht's!

### **Schritt 1: CSV-Dateien ablegen**
- Legen Sie Ihre CSV-Dateien in den Ordner `CSV_Input` ab
- Das Programm erkennt automatisch den Monat aus dem Dateinamen

### **Schritt 2: Anwendung starten**
- **Doppelklick auf `start_excel.bat`**
- Wählen Sie zwischen GUI-Version (empfohlen) oder Kommandozeilen-Version

### **Schritt 3: Fertig!**
- Excel-Dateien erscheinen im Ordner `Excel_Output`
- Verarbeitete CSV-Dateien werden automatisch ins `CSV_Archive` verschoben

## 🖥️ GUI-Version (Empfohlen)

Die **GUI-Version** bietet eine benutzerfreundliche Oberfläche:

1. **Start:** Doppelklick auf `start_excel.bat` → Option "1" wählen
2. **Übersicht:** Alle Ordner und CSV-Dateien werden angezeigt
3. **Verarbeitung:** Einfach auf "Verarbeiten" klicken
4. **Status:** Fortschritt wird live angezeigt
5. **Ordner öffnen:** Direkte Links zu allen Ordnern

### **Vorteile der GUI:**
- ✅ **Benutzerfreundlich** - Keine Kommandozeile nötig
- ✅ **Übersichtlich** - Alle Dateien auf einen Blick
- ✅ **Sicher** - Automatische Ordner-Erstellung
- ✅ **Schnell** - Ein Klick für alles

## 📋 Kommandozeilen-Version

Für erfahrene Benutzer:

1. **Start:** Doppelklick auf `start_excel.bat` → Option "2" wählen
2. **Datei wählen:** Nummer der gewünschten CSV-Datei eingeben
3. **Verarbeitung:** Automatische Excel-Erstellung
4. **Fertig:** Excel-Datei im `Excel_Output`-Ordner

## 🔧 Automatische Monatserkennung

Das Programm erkennt automatisch den Monat aus dem Dateinamen:

- ✅ `Company_July_2025.csv` → **Juli 2025**
- ✅ `Working_Hours_08_2025.csv` → **August 2025**
- ✅ `Time_Tracking_September_2025.csv` → **September 2025**
- ✅ `Working_Hours_12_2024.csv` → **Dezember 2024**

## 📊 Excel-Ausgabe

Jede Excel-Datei enthält:

### **📋 Rohdaten-Blatt**
- Alle ursprünglichen Daten, schön formatiert
- Übersichtliche Tabelle mit allen Details

### **👥 Mitarbeiter-Übersicht**
- Summe der Arbeitsstunden pro Mitarbeiter
- Gesamtbetrag pro Mitarbeiter
- Durchschnittlicher Stundenlohn

### **📅 Tages-Übersicht**
- Anzahl Mitarbeiter pro Tag
- Gesamtstunden pro Tag
- Gesamtbetrag pro Tag
- Durchschnittlicher Stundenlohn pro Tag

### **📈 Monats-Übersicht**
- Gesamtstunden des Monats
- Gesamtbetrag des Monats
- Durchschnittliche Stunden pro Tag
- Durchschnittlicher Stundenlohn

### **👤 Individuelle Mitarbeiter-Blätter**
- Ein separates Arbeitsblatt für jeden Mitarbeiter
- Detaillierte Aufstellung aller Arbeitstage
- Summen und Durchschnitte pro Mitarbeiter

## 🗂️ Automatisches Archivieren

- **Verarbeitete CSV-Dateien** werden automatisch in den `CSV_Archive`-Ordner verschoben
- **Keine Duplikate** - Bei gleichen Dateinamen wird ein Zeitstempel hinzugefügt
- **Saubere Trennung** - Eingabe und Archiv sind getrennt

## 🛠️ Was ist in diesem Ordner?

- **`start_excel.bat`** - Hauptstartdatei (Doppelklick zum Starten)
- **`app/simple_excel_processor.py`** - Kernverarbeitung (Python-Skript)
- **`app/gui_app.py`** - Benutzeroberfläche (GUI-Version)
- **`app/requirements_excel.txt`** - Benötigte Python-Bibliotheken
- **`README.md`** - Diese Anleitung

## 🚀 Einfacher Start

1. **Doppelklick auf `start_excel.bat`**
2. **Wählen Sie "1" für GUI-Version** (empfohlen)
3. **Legen Sie CSV-Dateien in `CSV_Input` ab**
4. **Klicken Sie auf "Verarbeiten"**
5. **Fertig!** Excel-Dateien finden Sie in `Excel_Output`

## 🔧 Systemanforderungen

- **Windows** (getestet auf Windows 10/11)
- **Python 3.7+** (wird automatisch installiert)
- **Internetverbindung** (nur für erste Installation)

## ❓ Häufige Fragen

### **"Keine CSV-Dateien gefunden!"**
- Legen Sie CSV-Dateien in den `CSV_Input`-Ordner ab
- Stellen Sie sicher, dass die Dateien die Endung `.csv` haben

### **"Fehler beim Installieren der Abhängigkeiten!"**
- Stellen Sie sicher, dass Sie eine Internetverbindung haben
- Führen Sie `start_excel.bat` als Administrator aus

### **"Excel-Datei wird nicht erstellt!"**
- Überprüfen Sie, ob die CSV-Datei das richtige Format hat
- Schauen Sie in den `CSV_Input`-Ordner für Fehlermeldungen

### **"Kann ich mehrere Monate gleichzeitig verarbeiten?"**
- Ja! Legen Sie einfach alle CSV-Dateien in den `CSV_Input`-Ordner
- Das Programm erstellt für jeden Monat eine separate Excel-Datei

## 📞 Support

Bei Problemen oder Fragen:
1. Überprüfen Sie diese Anleitung
2. Schauen Sie in die Log-Ausgabe der Anwendung
3. Stellen Sie sicher, dass alle Ordner korrekt erstellt wurden

---

**🎉 Viel Erfolg mit Ihrer Arbeitszeiten-Verarbeitung!**
