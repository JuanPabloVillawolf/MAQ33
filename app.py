import streamlit as st
import pandas as pd
import pdfplumber
import re
import os
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
import warnings
warnings.filterwarnings('ignore')

st.set_page_config(page_title="FLEXCON Extractor", layout="wide")

st.title("📊 Extractor Universal FLEXCON - PDF a Excel")

# ===============================
# FUNCIONES ORIGINALES (ADAPTADAS)
# ===============================

def analyze_pdf_structure(pdf_path):
    info = {
        'total_pages': 0,
        'transfer_type': None
    }

    with pdfplumber.open(pdf_path) as pdf:
        info['total_pages'] = len(pdf.pages)
        text = pdf.pages[0].extract_text()

        if text:
            if 'Non-Inventory Transfer' in text:
                info['transfer_type'] = 'non-inventory'
            elif 'INV TRANSFER' in text or 'INVENTORY TRANSFER' in text.upper():
                info['transfer_type'] = 'inventory'

    return info


def extract_inventory_transfer_data(pdf_path):
    all_data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text()
            if not text:
                continue

            lines = text.split('\n')

            for i, line in enumerate(lines):
                match = re.match(r'^(\d{1,3})\s+(\d{3})\s+([\w-]+)', line)

                if match:
                    line_num = match.group(1)
                    bp = match.group(2)
                    item = match.group(3)

                    rest_of_line = line[match.end():].strip()

                    parts = rest_of_line.split()

                    origin = ''
                    origin_idx = -1
                    for idx, part in enumerate(parts):
                        if part in ['USA', 'JPN', 'NLD', 'MEX', 'CHN']:
                            origin = part
                            origin_idx = idx
                            break

                    if origin_idx == -1:
                        continue

                    description = ' '.join(parts[:origin_idx])
                    after_origin = parts[origin_idx + 1:]
                    after_origin = [p.replace(',', '') for p in after_origin]

                    quantity = after_origin[0] if len(after_origin) >= 1 else ''
                    uom = after_origin[1] if len(after_origin) >= 2 else ''
                    qty_m = after_origin[2] if len(after_origin) >= 3 else ''

                    value = ''
                    for j in range(len(after_origin) - 1, -1, -1):
                        if re.match(r'^\d+\.\d{2}$', after_origin[j]):
                            value = after_origin[j]
                            break

                    record = {
                        'LINE': line_num,
                        'B/P': bp,
                        'ITEM': item,
                        'DESCRIPTION/LOT': description.strip(),
                        'ORIGIN': origin,
                        'QUANTITY': quantity,
                        'UOM': uom,
                        'QTY(M)': qty_m,
                        'VALUE': value
                    }

                    all_data.append(record)

    return all_data


# ===============================
# INTERFAZ WEB
# ===============================

uploaded_file = st.file_uploader("📤 Sube tu PDF FLEXCON", type="pdf")

if uploaded_file is not None:

    temp_pdf = "temp.pdf"
    with open(temp_pdf, "wb") as f:
        f.write(uploaded_file.read())

    st.info("Analizando PDF...")

    pdf_info = analyze_pdf_structure(temp_pdf)
    data = extract_inventory_transfer_data(temp_pdf)

    if len(data) == 0:
        st.error("❌ No se pudieron extraer datos.")
    else:
        df = pd.DataFrame(data)
        df['LINE'] = pd.to_numeric(df['LINE'], errors='coerce')
        df['VALUE'] = pd.to_numeric(df['VALUE'], errors='coerce')
        df = df.sort_values('LINE').reset_index(drop=True)

        st.success(f"✅ {len(df)} registros extraídos")
        st.dataframe(df)

        output_file = "resultado.xlsx"
        df.to_excel(output_file, index=False)

        with open(output_file, "rb") as f:
            st.download_button(
                label="📥 Descargar Excel",
                data=f,
                file_name="FLEXCON_EXTRACTED.xlsx"
            )
