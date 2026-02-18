# -*- coding: utf-8 -*-
"""
FLEXCON_Extractor_Universal_v2.py
==================================
App Streamlit — Extractor de documentos FLEXCON PDF → Excel

Hoja 1 : FLEXCON INV TRANSFER
Hoja 2 : SEIP Non-Inventory Transfer

Uso:
    streamlit run flexcon_extractor_universal_v2.py
"""

import re
import io
import os
import warnings

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

warnings.filterwarnings("ignore")

# ══════════════════════════════════════════════════════════════════════════════
#  CONSTANTES
# ══════════════════════════════════════════════════════════════════════════════

ORIGINS     = {"USA", "JPN", "NLD", "MEX", "CHN", "KOR", "TWN", "DEU"}
UOM_TYPES   = {"PC", "FT", "EA", "GAL", "PT", "IN", "BOX", "KG", "LB", "M", "SET"}
LOT_PATTERN = re.compile(r"^[A-Z]\d{7,}", re.IGNORECASE)


# ══════════════════════════════════════════════════════════════════════════════
#  UTILIDADES
# ══════════════════════════════════════════════════════════════════════════════

def clean_num(s: str) -> str:
    return s.replace(",", "").replace("$", "").strip()

def safe_int(val):
    try:
        if pd.isna(val) or str(val).strip() in ("", "N/A", "N/D"):
            return None
        return int(clean_num(str(val)))
    except Exception:
        return None

def safe_float(val):
    try:
        if pd.isna(val) or str(val).strip() in ("", "N/A", "N/D"):
            return None
        return float(clean_num(str(val)))
    except Exception:
        return None

def extract_page_text(page) -> str:
    try:
        return page.extract_text(x_tolerance=3, y_tolerance=3) or ""
    except Exception:
        return page.extract_text() or ""

def is_non_inventory_page(text: str) -> bool:
    markers = [
        "non-inventory transfer", "requestor name",
        "ship to address", "m27009",
    ]
    t = text.lower()
    return any(m in t for m in markers)

def is_inventory_page(text: str) -> bool:
    markers = ["inv transfer", "inventory transfer", "flexcon inv", "p5546"]
    t = text.upper()
    return any(m in t for m in markers)

def extract_document_meta(text: str) -> dict:
    meta = {"doc_number": "", "transfer_date": "", "total_value": None, "total_boxes": None}
    m = re.search(r"\b(P\d{7,})\b", text)
    if m:
        meta["doc_number"] = m.group(1)
    m = re.search(r"Transfer Date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})", text, re.IGNORECASE)
    if m:
        meta["transfer_date"] = m.group(1)
    else:
        m = re.search(r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b", text)
        if m:
            meta["transfer_date"] = m.group(1)
    m = re.search(r"Total\s+Value[:\s]+\$?([\d,]+\.\d{2})", text, re.IGNORECASE)
    if m:
        meta["total_value"] = float(clean_num(m.group(1)))
    m = re.search(r"Total\s+Boxes[:\s]+(\d+)", text, re.IGNORECASE)
    if m:
        meta["total_boxes"] = int(m.group(1))
    return meta


# ══════════════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN — INV TRANSFER
# ══════════════════════════════════════════════════════════════════════════════

def parse_inventory_line(line_parts: list, line_num: str, bp: str, item: str):
    origin_idx = -1
    for i, tok in enumerate(line_parts):
        if tok.upper() in ORIGINS:
            origin_idx = i
            break
    if origin_idx == -1:
        return None

    description = " ".join(line_parts[:origin_idx])
    origin      = line_parts[origin_idx].upper()
    after       = [clean_num(t) for t in line_parts[origin_idx + 1:]]

    int_vals   = [v for v in after if re.match(r"^\d+$", v)]
    float_vals = [v for v in after if re.match(r"^\d+\.\d{2}$", v)]
    uom_vals   = [v for v in after if v.upper() in UOM_TYPES]

    quantity = int_vals[0]             if int_vals            else ""
    uom      = uom_vals[0].upper()     if uom_vals            else ""
    qty_m    = int_vals[1]             if len(int_vals) >= 2  else ""
    value    = float_vals[-1]          if float_vals          else ""

    return {
        "LINE":            line_num,
        "B/P":             bp,
        "ITEM":            item,
        "DESCRIPTION/LOT": description.strip(),
        "ORIGIN":          origin,
        "QUANTITY":        quantity,
        "UOM":             uom,
        "QTY(M)":          qty_m,
        "BOXES":           "N/D",
        "WEIGHT(KG)":      "N/D",
        "VALUE":           value,
    }


def extract_inventory_data(pdf_bytes: bytes):
    all_records = []
    doc_meta    = {}
    DATA_LINE_RE = re.compile(r"^(\d{1,3})\s+(\d{3})\s+([\w\-\.]+)(.*)")

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = extract_page_text(page)
            if not text or is_non_inventory_page(text):
                continue

            if not doc_meta:
                doc_meta = extract_document_meta(text)

            lines = text.split("\n")
            i = 0
            while i < len(lines):
                raw = lines[i].strip()
                m = DATA_LINE_RE.match(raw)
                if not m:
                    i += 1
                    continue

                line_num = m.group(1)
                bp       = m.group(2)
                item     = m.group(3)
                rest     = m.group(4).strip()
                tokens   = rest.split()
                j = i + 1

                while j < min(i + 5, len(lines)):
                    candidate = lines[j].strip()
                    if LOT_PATTERN.match(candidate) and not any(
                        candidate.upper().startswith(orig) for orig in ORIGINS
                    ):
                        tokens = tokens + [candidate]
                        j += 1
                        continue
                    if any(
                        f" {o} " in f" {candidate} " or candidate.startswith(o + " ")
                        for o in ORIGINS
                    ):
                        tokens += candidate.split()
                        j += 1
                        break
                    if re.match(r"^\d", candidate):
                        if any(t.upper() in ORIGINS for t in tokens):
                            break
                        tokens += candidate.split()
                        j += 1
                        continue
                    break

                record = parse_inventory_line(tokens, line_num, bp, item)
                if record:
                    all_records.append(record)
                i = j

    return all_records, doc_meta


# ══════════════════════════════════════════════════════════════════════════════
#  EXTRACCIÓN — NON-INVENTORY TRANSFER
# ══════════════════════════════════════════════════════════════════════════════

def extract_non_inventory_data(pdf_bytes: bytes):
    all_records = []
    doc_meta    = {}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            text = extract_page_text(page)
            if not text or not is_non_inventory_page(text):
                continue

            if not doc_meta:
                for pattern, key in [
                    (r"Requestor Name[:\s]+(.+)",                       "requestor"),
                    (r"From Location[:\s]+(.+)",                        "from_location"),
                    (r"Shipment Date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})",  "shipment_date"),
                    (r"Ship To Address[:\s]+(.+)",                       "ship_to"),
                ]:
                    m = re.search(pattern, text, re.IGNORECASE)
                    if m:
                        doc_meta[key] = m.group(1).strip()
                m = re.search(r"TOTAL\s+VALUE[:\s]+\$?([\d,]+\.\d{2})", text, re.IGNORECASE)
                if m:
                    doc_meta["total_value"] = float(clean_num(m.group(1)))

            lines = text.split("\n")
            for line in lines:
                m = re.match(r"^(\d{1,2})\s+(\S+)\s+(.+)", line.strip())
                if not m:
                    continue
                line_num = int(m.group(1))
                if line_num < 1 or line_num > 20:
                    continue

                item_num  = m.group(2)
                remainder = m.group(3).strip()

                origin_match = None
                for orig in ORIGINS:
                    pat = re.search(rf"\b{orig}\b(.*)", remainder, re.IGNORECASE)
                    if pat:
                        origin_match = (orig.upper(), pat.start(), pat.group(1).strip())
                        break
                if not origin_match:
                    continue

                origin        = origin_match[0]
                before_origin = remainder[:origin_match[1]].strip()
                after_origin  = origin_match[2]

                uom          = ""
                description  = before_origin
                words_before = before_origin.split()
                if words_before and words_before[-1].upper() in UOM_TYPES:
                    uom         = words_before[-1].upper()
                    description = " ".join(words_before[:-1])

                after_tokens = [clean_num(t) for t in after_origin.split()]
                quantity = ""
                lot      = "N/A"
                value    = ""
                for tok in after_tokens:
                    if tok.upper() in ("N/A", "NA"):
                        lot = "N/A"
                    elif re.match(r"^\d+\.\d{2}$", tok):
                        value = tok
                    elif re.match(r"^\d+$", tok) and not quantity:
                        quantity = tok

                all_records.append({
                    "LINE":          str(line_num),
                    "ITEM NUMBER":   item_num,
                    "DESCRIPTION":   description.strip(),
                    "U/M":           uom,
                    "ORIGIN":        origin,
                    "QTY REQUESTED": quantity,
                    "LOT NUMBER":    lot,
                    "VALUE":         value,
                    "BOXES":         "N/D",
                    "WEIGHT(KG)":    "N/D",
                })

    return all_records, doc_meta


# ══════════════════════════════════════════════════════════════════════════════
#  CONSTRUCCIÓN DEL EXCEL  (devuelve bytes en memoria)
# ══════════════════════════════════════════════════════════════════════════════

def _style_header_row(ws, row_num: int, hex_color: str):
    thick = Side(style="medium", color="404040")
    fill  = PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")
    for cell in ws[row_num]:
        cell.fill      = fill
        cell.font      = Font(bold=True, color="FFFFFF", name="Arial", size=10)
        cell.border    = Border(left=thick, right=thick, top=thick, bottom=thick)
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row_num].height = 28


def _style_data_rows(ws, data_start: int, data_end: int, alt_hex: str):
    thin     = Side(style="thin", color="C0C0C0")
    thin_b   = Border(left=thin, right=thin, top=thin, bottom=thin)
    alt_fill = PatternFill(start_color=alt_hex, end_color=alt_hex, fill_type="solid")
    for row_idx in range(data_start, data_end + 1):
        for cell in ws[row_idx]:
            cell.border    = thin_b
            cell.font      = Font(name="Arial", size=9)
            cell.alignment = Alignment(vertical="center")
            if row_idx % 2 == 0:
                cell.fill = alt_fill


def _add_total_row(ws, data_start: int, data_end: int,
                   headers_map: dict, money_col: str,
                   sum_cols: list, total_hex: str):
    thick  = Side(style="medium", color="404040")
    t_fill = PatternFill(start_color=total_hex, end_color=total_hex, fill_type="solid")
    total_r = data_end + 1
    ws.append(["TOTAL"] + [""] * (len(headers_map) - 1))
    for col_name in sum_cols:
        if col_name in headers_map:
            cl = headers_map[col_name]
            ws[f"{cl}{total_r}"] = f'=IFERROR(SUM({cl}{data_start}:{cl}{data_end}),"")'
            ws[f"{cl}{total_r}"].number_format = "#,##0"
    if money_col in headers_map:
        cl = headers_map[money_col]
        ws[f"{cl}{total_r}"] = f'=IFERROR(SUM({cl}{data_start}:{cl}{data_end}),"")'
        ws[f"{cl}{total_r}"].number_format = "$#,##0.00"
    for cell in ws[total_r]:
        cell.fill   = t_fill
        cell.font   = Font(bold=True, name="Arial", size=9)
        cell.border = Border(left=thick, right=thick, top=thick, bottom=thick)


def build_excel(inv_records, inv_meta, noninv_records, noninv_meta) -> bytes:
    wb = Workbook()

    # ── Hoja 1: INV TRANSFER ────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = "INV Transfer"

    if inv_records:
        df1 = pd.DataFrame(inv_records)
        df1["LINE"]     = df1["LINE"].apply(safe_int)
        df1["QUANTITY"] = df1["QUANTITY"].apply(safe_int)
        df1["QTY(M)"]   = df1["QTY(M)"].apply(safe_int)
        df1["VALUE"]    = df1["VALUE"].apply(safe_float)
        df1.sort_values("LINE", inplace=True, ignore_index=True)

        meta_str = (
            f"Doc: {inv_meta.get('doc_number','')}   "
            f"Fecha: {inv_meta.get('transfer_date','')}   "
            f"Total Declarado: ${inv_meta.get('total_value', 0) or 0:,.2f}"
        )
        ws1.append([meta_str])
        ws1.merge_cells(f"A1:{get_column_letter(len(df1.columns))}1")
        ws1["A1"].font      = Font(bold=True, italic=True, name="Arial", size=9, color="404040")
        ws1["A1"].alignment = Alignment(horizontal="left")

        ws1.append(list(df1.columns))
        _style_header_row(ws1, 2, "1F4E78")

        for _, row in df1.iterrows():
            ws1.append(list(row))

        data_start, data_end = 3, 2 + len(df1)
        _style_data_rows(ws1, data_start, data_end, "F0F4F8")

        headers_map1 = {cell.value: cell.column_letter for cell in ws1[2]}

        for r in range(data_start, data_end + 1):
            for c in ["QUANTITY", "QTY(M)"]:
                if c in headers_map1:
                    ws1[f"{headers_map1[c]}{r}"].number_format = "#,##0"
            if "VALUE" in headers_map1:
                ws1[f"{headers_map1['VALUE']}{r}"].number_format = "$#,##0.00"

        _add_total_row(ws1, data_start, data_end, headers_map1,
                       "VALUE", ["QUANTITY", "QTY(M)"], "D6E4F0")

        col_widths1 = {
            "A": 6,  "B": 7,  "C": 16, "D": 58,
            "E": 9,  "F": 12, "G": 7,  "H": 10,
            "I": 8,  "J": 11, "K": 13,
        }
        for cl, w in col_widths1.items():
            ws1.column_dimensions[cl].width = w
        ws1.freeze_panes = "A3"
    else:
        ws1["A1"] = "⚠️ No se encontraron registros de tipo INV TRANSFER"

    # ── Hoja 2: NON-INVENTORY TRANSFER ──────────────────────────────────────
    ws2 = wb.create_sheet("Non-Inv Transfer")

    if noninv_records:
        df2 = pd.DataFrame(noninv_records)
        df2["LINE"]          = df2["LINE"].apply(safe_int)
        df2["QTY REQUESTED"] = df2["QTY REQUESTED"].apply(safe_int)
        df2["VALUE"]         = df2["VALUE"].apply(safe_float)
        df2.sort_values("LINE", inplace=True, ignore_index=True)

        meta_str2 = (
            f"Requestor: {noninv_meta.get('requestor','')}   "
            f"From: {noninv_meta.get('from_location','')}   "
            f"Ship To: {noninv_meta.get('ship_to','')}   "
            f"Fecha: {noninv_meta.get('shipment_date','')}   "
            f"Total Declarado: ${noninv_meta.get('total_value', 0) or 0:,.2f}"
        )
        ws2.append([meta_str2])
        ws2.merge_cells(f"A1:{get_column_letter(len(df2.columns))}1")
        ws2["A1"].font      = Font(bold=True, italic=True, name="Arial", size=9, color="404040")
        ws2["A1"].alignment = Alignment(horizontal="left")

        ws2.append(list(df2.columns))
        _style_header_row(ws2, 2, "1D6A3A")

        for _, row in df2.iterrows():
            ws2.append(list(row))

        data_start2, data_end2 = 3, 2 + len(df2)
        _style_data_rows(ws2, data_start2, data_end2, "EDF7EF")

        headers_map2 = {cell.value: cell.column_letter for cell in ws2[2]}

        for r in range(data_start2, data_end2 + 1):
            if "QTY REQUESTED" in headers_map2:
                ws2[f"{headers_map2['QTY REQUESTED']}{r}"].number_format = "#,##0"
            if "VALUE" in headers_map2:
                ws2[f"{headers_map2['VALUE']}{r}"].number_format = "$#,##0.00"

        _add_total_row(ws2, data_start2, data_end2, headers_map2,
                       "VALUE", ["QTY REQUESTED"], "D5F0DC")

        col_widths2 = {
            "A": 6,  "B": 14, "C": 50, "D": 7,
            "E": 9,  "F": 13, "G": 11, "H": 12,
            "I": 8,  "J": 11,
        }
        for cl, w in col_widths2.items():
            ws2.column_dimensions[cl].width = w
        ws2.freeze_panes = "A3"
    else:
        ws2["A1"] = "⚠️ No se encontraron registros de tipo Non-Inventory Transfer"

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()


# ══════════════════════════════════════════════════════════════════════════════
#  INTERFAZ STREAMLIT
# ══════════════════════════════════════════════════════════════════════════════

st.set_page_config(
    page_title="FLEXCON Extractor",
    page_icon="📦",
    layout="centered",
)

st.title("📦 FLEXCON PDF → Excel Extractor")
st.markdown(
    "Sube un documento **FLEXCON INV TRANSFER** (puede contener también páginas "
    "**Non-Inventory Transfer**). Se generará un Excel con dos hojas."
)

uploaded_file = st.file_uploader("Selecciona el archivo PDF", type=["pdf"])

if uploaded_file is not None:
    # Leer todo el PDF como bytes una sola vez
    pdf_bytes = uploaded_file.read()

    with st.spinner("Extrayendo datos del PDF..."):
        inv_records,    inv_meta    = extract_inventory_data(pdf_bytes)
        noninv_records, noninv_meta = extract_non_inventory_data(pdf_bytes)

    # ── Métricas ─────────────────────────────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🗂️ INV Transfer")
        st.metric("Registros extraídos", len(inv_records))
        if inv_records:
            df_inv = pd.DataFrame(inv_records)
            df_inv["VALUE"] = df_inv["VALUE"].apply(safe_float)
            total_calc = df_inv["VALUE"].sum()
            st.metric("Valor calculado", f"${total_calc:,.2f}")
            total_decl = inv_meta.get("total_value") or 0
            if total_decl:
                diff   = total_calc - total_decl
                status = "✅ Coincide" if abs(diff) < 1 else f"⚠️ Dif: ${diff:,.2f}"
                st.metric("Total declarado", f"${total_decl:,.2f}", delta=status)

    with col2:
        st.subheader("📋 Non-Inv Transfer")
        st.metric("Registros extraídos", len(noninv_records))
        if noninv_records:
            df_ni = pd.DataFrame(noninv_records)
            df_ni["VALUE"] = df_ni["VALUE"].apply(safe_float)
            total_calc_ni = df_ni["VALUE"].sum()
            st.metric("Valor calculado", f"${total_calc_ni:,.2f}")
            total_decl_ni = noninv_meta.get("total_value") or 0
            if total_decl_ni:
                diff_ni   = total_calc_ni - total_decl_ni
                status_ni = "✅ Coincide" if abs(diff_ni) < 1 else f"⚠️ Dif: ${diff_ni:,.2f}"
                st.metric("Total declarado", f"${total_decl_ni:,.2f}", delta=status_ni)

    # ── Preview ───────────────────────────────────────────────────────────────
    if inv_records:
        with st.expander("👁️ Preview — INV Transfer", expanded=False):
            st.dataframe(pd.DataFrame(inv_records), use_container_width=True)

    if noninv_records:
        with st.expander("👁️ Preview — Non-Inv Transfer", expanded=False):
            st.dataframe(pd.DataFrame(noninv_records), use_container_width=True)

    # ── Generar y descargar Excel ─────────────────────────────────────────────
    if inv_records or noninv_records:
        with st.spinner("Generando Excel..."):
            excel_bytes = build_excel(inv_records, inv_meta, noninv_records, noninv_meta)

        base_name   = os.path.splitext(uploaded_file.name)[0]
        output_name = f"{base_name}_EXTRACTED.xlsx"

        st.success("✅ Excel generado correctamente")
        st.download_button(
            label="⬇️ Descargar Excel",
            data=excel_bytes,
            file_name=output_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        st.error(
            "❌ No se pudieron extraer datos. "
            "Verifica que el archivo sea un documento FLEXCON válido."
        )

else:
    st.info("👆 Sube un archivo PDF para comenzar.")
