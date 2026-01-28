import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io

st.set_page_config(page_title="Convertitore XML", layout="centered")

st.title("📂 Convertitore XML Fatture → Excel")
st.write("Carica uno o più file XML per estrarre i dati in un unico foglio Excel.")

# Caricamento file
uploaded_files = st.file_uploader("Trascina qui i tuoi file XML", type="xml", accept_multiple_files=True)

if uploaded_files:
    risultati = []
    
    for file in uploaded_files:
        try:
            # Leggiamo il contenuto del file
            string_data = file.read()
            root = ET.fromstring(string_data)
            
            # Esempio di estrazione (questi campi dipendono dalla struttura del tuo XML)
            # Proviamo a cercare dei tag comuni nelle fatture
            nome_file = file.name
            dati_estratte = {"File": nome_file}
            
            # Esempio generico: cerca tutti i tag e prendi i primi valori
            for child in root.iter():
                if child.text and len(child.text.strip()) > 0:
                    tag_name = child.tag.split('}')[-1] # Rimuove il namespace se presente
                    if tag_name not in dati_estratte:
                        dati_estratte[tag_name] = child.text
            
            risultati.append(dati_estratte)
        except Exception as e:
            st.error(f"Errore nella lettura di {file.name}: {e}")

    if risultati:
        df = pd.DataFrame(risultati)
        st.success(f"Elaborati {len(risultati)} file con successo!")
        st.dataframe(df.head()) # Mostra un'anteprima dei primi 5

        # Generazione Excel in memoria
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False)
        
        st.download_button(
            label="📥 Scarica Excel Completo",
            data=output.getvalue(),
            file_name="dati_fatture.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
