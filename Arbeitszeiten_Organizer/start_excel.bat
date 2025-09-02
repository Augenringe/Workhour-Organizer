@echo off
echo Arbeitszeiten-Verarbeitung - Agilos QCS GmbH
echo ============================================
echo.

echo Installiere Abhängigkeiten...
cd app
pip install -r requirements_excel.txt
cd ..
echo.

echo Erstelle Ordnerstruktur...
if not exist "CSV_Eingabe" mkdir "CSV_Eingabe"
if not exist "Excel_Ausgabe" mkdir "Excel_Ausgabe"
if not exist "CSV_Archiv" mkdir "CSV_Archiv"
echo.

echo Wählen Sie die Verarbeitungsart:
echo.
echo 1. GUI-Version (Empfohlen) - Benutzerfreundliche Oberfläche
echo 2. Kommandozeilen-Version - Für erfahrene Benutzer
echo.
set /p choice="Ihre Wahl (1 oder 2): "

if "%choice%"=="1" (
    echo.
    echo Starte GUI-Anwendung...
    python app\gui_app.py
    goto :end
)

if "%choice%"=="2" (
    echo.
    echo Verfügbare CSV-Dateien im CSV_Eingabe-Ordner:
    echo.
    set /a counter=0
    for %%f in (CSV_Eingabe\*.csv) do (
        set /a counter+=1
        echo !counter! - %%~nxf
    )
    echo.

    if %counter%==0 (
        echo Keine CSV-Dateien im CSV_Eingabe-Ordner gefunden!
        echo.
        echo Bitte legen Sie Ihre CSV-Dateien in den Ordner "CSV_Eingabe" ab.
        echo Das Programm erkennt automatisch den Monat aus dem Dateinamen.
        echo.
        pause
        exit /b
    )

    set /p file_choice="Welche CSV-Datei verarbeiten? (Nummer eingeben oder 'alle' für alle Dateien): "

    if "%file_choice%"=="alle" (
        echo.
        echo Verarbeite alle CSV-Dateien...
        for %%f in (CSV_Eingabe\*.csv) do (
            echo.
            echo Verarbeite: %%~nxf
            python app\simple_excel_processor.py "%%f"
        )
    ) else (
        echo.
        echo Verarbeite ausgewählte CSV-Datei...
        set /a file_num=1
        for %%f in (CSV_Eingabe\*.csv) do (
            if !file_num!==%file_choice% (
                python app\simple_excel_processor.py "%%f"
                goto :done
            )
            set /a file_num+=1
        )
        echo Fehler: Ungültige Auswahl!
    )

    :done
    echo.
    echo Fertig! Excel-Dateien wurden im Ordner "Excel_Ausgabe" erstellt.
    echo Verarbeitete CSV-Dateien wurden ins "CSV_Archiv" verschoben.
    echo.
    pause
    goto :end
)

echo.
echo Ungültige Auswahl! Bitte wählen Sie 1 oder 2.
pause

:end
