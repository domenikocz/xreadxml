from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

def esporta_excel_formattato(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Fatture')
        workbook = writer.book
        worksheet = writer.sheets['Fatture']

        # 1. Stili (equivalenti alle tue costanti VBA)
        header_fill = PatternFill(start_color="C0C0C0", end_color="C0C0C0", fill_type="solid")
        font_bold = Font(bold=True)
        thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                             top=Side(style='thin'), bottom=Side(style='thin'))
        
        # Colori per il TIPO
        color_map = {
            "FATTURA": Font(color="008000", bold=True),      # Verde
            "PARCELLA": Font(color="800080", bold=True),     # Viola
            "NOTA CREDITO": Font(color="800000", bold=True)  # Amaranto
        }

        # 2. Formattazione Intestazioni
        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = font_bold
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center")

        # 3. Formattazione Righe e Dati
        for row in worksheet.iter_rows(min_row=2, max_row=len(df)+1):
            for cell in row:
                cell.border = thin_border
                
                # Applica formattazione valuta (€) alle colonne numeriche (6 a 10)
                if cell.column in range(6, 11): 
                    cell.number_format = '#,##0.00 "€"'
                
                # Colore condizionale nella colonna TIPO (Colonna 5)
                if cell.column == 5:
                    valore_tipo = str(cell.value).upper()
                    if valore_tipo in color_map:
                        cell.font = color_map[valore_tipo]
            
            # Formattazione Data (Colonna 4)
            worksheet.cell(row=row[0].row, column=4).number_format = 'DD/MM/YYYY'

        # Auto-fit delle colonne
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            worksheet.column_dimensions[column_letter].width = max_length + 2

    return output.getvalue()

# --- Modifica questa parte nel blocco 'if lista_finale:' ---
if lista_finale:
    df = pd.DataFrame(lista_finale)
    cols = ["P.IVA", "DENOMINAZIONE", "NUMERO", "DATA", "TIPO", "IMPONIBILE", "IVA", "TOTALE", "RITENUTE", "BOLLO", "TIPO RIT.", "CAUSALE", "COD.FISC", "DESCRIZIONE", "NOME FILE"]
    df = df[cols]
    
    st.dataframe(df)
    
    # Chiamata alla nuova funzione di formattazione
    excel_data = esporta_excel_formattato(df)
    
    st.download_button(
        label="📥 Scarica Excel Formattato",
        data=excel_data,
        file_name="fatture_elaborate.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
