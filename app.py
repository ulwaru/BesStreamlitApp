from pathlib import Path
import streamlit as st
st.title("📂 Mein Online-Textarchiv")

# Ein Eingabefeld für Text
name = st.text_input("Wie heißen Sie?", "Max Mustermann")

# Ein Schieberegler für Zahlen
alter = st.slider("Wählen Sie Ihr Alter:", 18, 100, 75)

# Ein Button mit Aktion
if st.button("Begrüßung anzeigen"):
    st.success(f"Hallo {name}! Sie sind {alter} Jahre alt.")

# Pfad zu Ihrem Ordner festlegen
ORDNER_PFAD = Path("F:\Archiv Trampolin 1900-1999")  # Oder relativer Pfad zu Ihren Dateien

if ORDNER_PFAD.exists():
    # Alle .txt-Dateien im Ordner suchen
    txt_dateien = list(ORDNER_PFAD.glob("*.txt"))

    if txt_dateien:
        # Dateiauswahl für mobile Ansicht
        ausgewaehlte_datei = st.selectbox(
            "Wählen Sie eine Textdatei aus:",
            options=txt_dateien,
            format_func=lambda x: x.name,
        )

        if ausgewaehlte_datei:
            st.markdown(f"### Inhalt von: `{ausgewaehlte_datei.name}`")
            # Datei einlesen und anzeigen
            inhalt = ausgewaehlte_datei.read_text(encoding="utf-8")
            st.text_area("Inhalt:", inhalt, height=300)
    else:
        st.info("Keine .txt-Dateien im Ordner gefunden.")
else:
    st.warning("Der angegebene Ordner existiert nicht.")