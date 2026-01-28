import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
from datetime import datetime

# --- COSTANTI COLORI (equivalenti VBA) ---
COLOR_FATTURA = "green"
COLOR_PARCELLA = "purple"
COLOR_NOTA_CREDITO = "red"
COLOR_NEGATIVO = "red"

def determina_tipo_documento(tipo_doc):
    mapping = {
        "TD01": "FATTURA", "TD02": "FATTURA", "TD03": "FATTURA",
        "TD04": "NOTA CREDITO", "TD05": "NOTA DEBITO", "TD06": "PARCELLA"
    }
    return mapping.get(tipo_doc, "ALTRO")

def processa_xml(file):
    tree = ET.parse(file)
    root = tree.getroot()
    # Rimuoviamo i namespace per semplicità di ricerca
    for elem in root.iter():
        if '}' in elem.tag:
            elem.tag = elem.tag.split('}', 1)[1]

    dati = {}
    
    # 1. Cedente Prestatore
    cp = root.find(".//CedentePrestatore/DatiAnagrafici")
    if cp is not None:
        dati['P.IVA'] = cp.findtext(".//IdCodice", "")
        dati['COD.FISC'] = cp.findtext("CodiceFiscale", "")
        denominazione = cp.findtext(".//Denominazione", "")
        if not denominazione:
            denominazione = f"{cp.findtext('.//Nome', '')} {cp.findtext('.//Cognome', '')}".strip()
        dati['DENOMINAZIONE'] = denominazione

    # 2. Dati Generali
    dg = root.find(".//DatiGeneraliDocumento")
    if dg is not None:
        dati['NUMERO'] = dg.findtext("Numero", "")
        data_str = dg.findtext("Data", "")
        dati['DATA'] = datetime.strptime(data_str, '%Y-%m-%d').date() if data_str else None
        dati['TIPO'] = determina_tipo_documento(dg.findtext("TipoDocumento", ""))
        dati['BOLLO'] = float(dg.findtext("ImportoBollo", "0").replace(",", "."))
        dati['TOTALE'] = float(dg.findtext("ImportoTotaleDocumento", "0").replace(",", "."))
        dati['CAUSALE'] = dg.findtext("Causale", "")

    # 3. Riepilogo (Somma Imponibile e IVA come da macro)
    tot_imponibile = 0.0
    tot_iva = 0.0
    for riepilogo in root.findall(".//DatiRiepilogo"):
        imp = riepilogo.findtext("ImponibileImporto", "0").replace(",", ".")
        iva = riepilogo.findtext("Imposta", "0").replace(",", ".")
        tot_imponibile += float(imp)
        tot_iva += float(iva)
    
    dati['IMPONIBILE'] = tot_imponibile
    dati['IVA'] = tot_iva

    # 4. Ritenute
    dr = root.find(".//DatiRitenuta")
    if dr is not None:
        dati['RITENUTE'] = float(dr.findtext("ImportoRitenuta", "0").replace(",", "."))
        dati['TIPO RIT.'] = dr.findtext("TipoRitenuta", "")
    else:
        dati['RITENUTE'] = 0.0
        dati['TIPO RIT.'] = ""

    # 5. Descrizione (Unione righe)
    descrizioni = [linea.text for linea in root.findall(".//DettaglioLinee/Descrizione") if linea.text]
    full_desc = "; ".join(descrizioni)
    dati['DESCRIZIONE'] = (full_desc[:252] + '...') if len(full_desc) > 255 else full_desc
    dati['NOME FILE'] = file.name

    return dati

# --- INTERFACCIA STREAMLIT ---
st.title("🚀 XML Invoice Processor")

uploaded_files = st.file_uploader("Carica file XML", type="xml", accept_multiple_files=True)

if uploaded_files:
    lista_finale = []
    visti = set() # Per controllo duplicati (P.IVA + Numero + Anno)

    for f in uploaded_files:
        try:
            d = processa_xml(f)
            chiave = (d['P.IVA'], d['NUMERO'], d['DATA'].year if d['DATA'] else "")
            
            if chiave not in visti:
                lista_finale.append(d)
                visti.add(chiave)
            else:
                st.warning(f"Duplicato ignorato: {d['NOME FILE']}")
        except Exception as e:
            st.error(f"Errore nel file {f.name}: {e}")

    if lista_finale:
        df = pd.DataFrame(lista_finale)
        
        # Riordino colonne come da macro
        cols = ["P.IVA", "DENOMINAZIONE", "NUMERO", "DATA", "TIPO", "IMPONIBILE", "IVA", "TOTALE", "RITENUTE", "BOLLO", "TIPO RIT.", "CAUSALE", "COD.FISC", "DESCRIZIONE", "NOME FILE"]
        df = df[cols]

        st.dataframe(df)

        # Export Excel con formattazione
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, index=False, sheet_name='Fatture')
            # Qui potremmo aggiungere colori alle celle, ma intanto esportiamo i dati puri
            
        st.download_button("📥 Scarica Excel", output.getvalue(), "fatture_elaborate.xlsx")
