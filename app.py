import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
import io
from datetime import datetime
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

# --- FUNZIONI DI LOGICA (Traduzione Macro VBA) ---

def determina_tipo_documento(tipo_doc):
    mapping = {
        "TD01": "FATTURA", "TD02": "FATTURA", "TD03": "FATTURA",
        "TD04": "NOTA CREDITO", "TD05": "NOTA DEBITO", "TD06": "PARCELLA"
    }
    return mapping.get(tipo_doc, "ALTRO")

def processa_xml(file):
    tree = ET.parse(file)
    root = tree.getroot()
    for elem in root.iter():
        if '}' in elem.tag:
            elem.tag = elem.tag.split('}', 1)[1]

    dati = {}
    cp = root.find(".//CedentePrestatore/DatiAnagrafici")
    if cp is not None:
        dati['P.IVA'] = cp.findtext(".//IdCodice", "")
        dati['COD.FISC'] = cp.findtext("CodiceFiscale", "")
        denominazione = cp.findtext(".//Denominazione", "")
        if not denominazione:
            denominazione = f"{cp.findtext('.//Nome', '')} {cp.findtext('.//Cognome', '')}".strip()
        dati['DENOMINAZIONE'] = denominazione

    dg = root.find(".//DatiGeneraliDocumento")
    if dg is not None:
        dati['NUMERO'] = dg.findtext("Numero", "")
        data_str = dg.findtext("Data", "")
        dati['DATA'] = datetime.strptime(data_str, '%Y-%m-%d').date() if data_str else None
        dati['TIPO'] = determina_tipo_documento(dg.findtext("TipoDocumento", ""))
        dati['BOLLO'] = float(dg.findtext("ImportoBollo", "0").replace(",", "."))
        dati['TOTALE'] = float(dg.findtext("ImportoTotaleDocumento", "0").replace(",", "."))
        dati['CAUSALE'] = dg.findtext("Causale", "")

    tot_imponibile = 0.0
    tot_iva = 0.0
    for riepilogo in root.findall(".//DatiRiepilogo"):
        imp = riepilogo.findtext("ImponibileImporto", "0").replace(",", ".")
        iva = riepilogo.findtext("Imposta", "0").replace(",", ".")
        tot_imponibile += float(imp)
        tot_iva += float(iva)
    
    dati['IMPONIBILE'] = tot_imponibile
    dati['IVA'] = tot_iva

    dr = root.find(".//DatiRitenuta")
    dati['RITENUTE'] = float(dr.findtext("ImportoRitenuta", "0").replace(",", ".")) if dr is not None else 0.0
    dati['TIPO RIT.'] = dr.findtext("TipoRitenuta", "") if dr is not None else ""

    descrizioni = [linea.text for linea in root.findall(".//DettaglioLinee/Descrizione") if linea.text]
    full_desc = "; ".join(descrizioni)
    dati['DESCRIZIONE'] = (full_desc[:252] + '...') if len(full_desc) > 255 else full_desc
    dati['NOME FILE'] = file.name
    return dati

def esporta_excel_formattato(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Fatture')
        workbook = writer.book
        worksheet = writer.sheets['Fatture']

        # Stili
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        font_bold = Font(bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        
        color_map = {
            "FATTURA": Font(color="008000", bold=True),      # Verde
            "PARCELLA": Font(color="800080", bold=True),     # Viola
            "NOTA CREDITO": Font(color="800000", bold=True)  # Rosso scuro
        }

        # Blocca riquadri (Prima riga)
        worksheet.freeze_panes = "A2"

        # Intestazioni
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = font_bold
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")

        # Dati
        for row in worksheet.iter_rows(min_row=2, max_row=len(df)+1):
            for cell in row:
                cell.border = thin_border
                if cell.column in range(6, 11): # Colonne Valuta
                    cell.number_format = '#,##0.00 "€"'
                if cell.column == 5: # Colore TIPO
                    valore_tipo = str(cell.value).upper()
                    if valore_tipo in color_map:
                        cell.font = color_map[valore_tipo]
            worksheet.cell(row=row[0].row, column=4).number_format = 'DD/MM/YYYY'

        # Auto-fit
        for column in worksheet.columns:
            max_length = max(len(str(cell.value or "")) for cell in column)
            worksheet.column_dimensions[column[0].column_letter].width = max_length + 2

    return output.getvalue()

# --- INTERFACCIA STREAMLIT ---

st.set_page_config(page_title="Convertitore XML", layout="wide")
st.title("📂 Elaborazione Fatture XML")

uploaded_files = st.file_uploader("Carica i tuoi file XML", type="xml", accept_multiple_files=True)

if uploaded_files:
    lista_finale = []
    visti = set()

    for f in uploaded_files:
        try:
            d = processa_xml(f)
            # Controllo duplicati: P.IVA + Numero + Anno
            anno = d['DATA'].year if d['DATA'] else ""
            chiave = (d['P.IVA'], d['NUMERO'], anno)
            
            if chiave not in visti:
                lista_finale.append(d)
                visti.add(chiave)
            else:
