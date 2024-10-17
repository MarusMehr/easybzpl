import streamlit as st
import pandas as pd
from ics import Calendar, Event
import datetime as dt
from io import BytesIO

# Funktion zur Erstellung eines minimalistischen Kalenders
def create_minimal_ics(df):
    cal = Calendar()
    for i, row in df.iterrows():
        event = Event()
        event.name = row['Gewerk']  # Setzt Gewerk als SUMMARY
        
        # Startdatum und Enddatum verarbeiten
        try:
            if pd.notna(row['Beginn\n']):  # Überprüfen, ob Beginn nicht NaT ist
                event.begin = dt.datetime.strptime(str(row['Beginn\n']), '%Y-%m-%d %H:%M:%S')
            else:
                continue  # Überspringen, falls Beginn leer ist
        except ValueError:
            event.begin = dt.datetime.strptime(str(row['Beginn\n']), '%Y-%m-%d')  # Fallback, falls Zeitangabe fehlt
            
        try:
            if pd.notna(row['Ende\n\n']):  # Überprüfen, ob Ende nicht NaT ist
                event.end = dt.datetime.strptime(str(row['Ende\n\n']), '%Y-%m-%d %H:%M:%S')
            else:
                event.end = event.begin  # Falls kein Enddatum vorhanden, nur Startdatum verwenden
        except ValueError:
            event.end = dt.datetime.strptime(str(row['Ende\n\n']), '%Y-%m-%d')  # Fallback, falls Zeitangabe fehlt
        
        event.location = row['Unternehmen']  # Unternehmen als LOCATION
        event.description = row['Bemerkung'] if pd.notna(row['Bemerkung']) else "Keine Bemerkung"
        event.transparent = True  # Markiert das Ereignis als transparent (frei)
        cal.events.add(event)
    return cal

# Funktion zum Generieren der ICS-Datei
def generate_ics_file(cal):
    ics_file = BytesIO()
    ics_file.write(str(cal).encode('utf-8'))
    ics_file.seek(0)
    return ics_file

# Streamlit-Benutzeroberfläche
st.title("Excel zu ICS-Konverter für Bauzeitenplan")

st.write("Lade eine Excel-Datei hoch, die die Spalten `Gewerk`, `Bemerkung`, `Unternehmen`, `Beginn`, `Ende` enthält.")

# Datei-Upload-Bereich
uploaded_file = st.file_uploader("Wähle eine Excel-Datei aus", type="xlsx")

if uploaded_file is not None:
    xls = pd.ExcelFile(uploaded_file)
    sheet_name = st.selectbox("Wähle das Blatt aus", xls.sheet_names)
    
    # Die Datei wird so geladen, dass Zeile 12 (Index 11) als Header genutzt wird
    df = pd.read_excel(xls, sheet_name=sheet_name, header=11)
    
    # Überprüfen, ob die erforderlichen Spalten vorhanden sind
    required_columns = ["Gewerk", "Bemerkung", "Unternehmen", "Beginn\n", "Ende\n\n"]
    if all(col in df.columns for col in required_columns):
        # Kalender erstellen und Datei generieren
        cal = create_minimal_ics(df)
        ics_file = generate_ics_file(cal)
        
        st.download_button(
            label="ICS-Datei herunterladen",
            data=ics_file,
            file_name="minimal_calendar.ics",
            mime="text/calendar"
        )
    else:
        st.error("Die Excel-Datei enthält nicht alle erforderlichen Spalten.")

