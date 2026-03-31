import streamlit as st
import pandas as pd
import pdfplumber
import re
import json
import io
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side
)
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────
# CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="Validador CCP vs ACEM",
    page_icon="🔍",
    layout="wide",
)

HISTORY_FILE = "historico.json"
TOLERANCE_PESO  = 0.5    # kg
TOLERANCE_VALOR = 1.0    # USD

# Palabras clave para detección automática de columnas
KEYWORDS = {
    "cantidad": ["cantidad", "piezas", "pieces", "qty", "units", "unidades"],
    "peso":     ["peso", "weight", "kg", "kilogram"],
    "valor":    ["valor", "value", "monto", "amount", "importe", "usd"],
    "factura":  ["factura", "invoice", "inv", "no factura", "num factura", "bill"],
}

# ─────────────────────────────────────────────
# UTILIDADES
# ─────────────────────────────────────────────

def detect_columns(df: pd.DataFrame) -> dict[str, str | None]:
    """Detecta automáticamente las columnas relevantes del Excel."""
    mapping = {k: None for k in KEYWORDS}
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for field, keywords in KEYWORDS.items():
        for kw in keywords:
            matches = [orig for low, orig in cols_lower.items() if kw in low]
            if matches:
                mapping[field] = matches[0]
                break
    return mapping


def parse_number(text: str) -> float | None:
    """Extrae el primer número flotante de un string."""
    if not text:
        return None
    clean = re.sub(r"[^\d.,]", "", str(text))
    clean = clean.replace(",", "")
    try:
        return float(clean)
    except ValueError:
        return None


def extract_pdf_data(pdf_file) -> dict:
    """
    Extrae Value, Weight y Pieces del manifiesto ACEM.
    Intenta detectar patrones de forma flexible.
    """
    result = {"value": None, "weight": None, "pieces": None,
              "ref": None, "consignee": None, "facturas": [], "raw_text": ""}
    try:
        with pdfplumber.open(pdf_file) as pdf:
            full_text = "\n".join(
                page.extract_text() or "" for page in pdf.pages
            )
        result["raw_text"] = full_text

        # Referencia / Trip Number
        m = re.search(r"Ref\s*[:\s]+(\S+)", full_text, re.IGNORECASE)
        if m:
            result["ref"] = m.group(1)

        # Consignatario
        m = re.search(r"([A-Z][A-Z ,\.]+(?:INC|LLC|SA|CORP|CO)[A-Z ,\.]*)\.",
                      full_text, re.IGNORECASE)
        if m:
            result["consignee"] = m.group(0).strip()

        # Weight  ─ busca número seguido de Kg
        m = re.search(r"([\d,\.]+)\s*[Kk][Gg]", full_text)
        if m:
            result["weight"] = parse_number(m.group(1))

        # Pieces ─ número que aparece DESPUÉS del peso en la misma línea
        m = re.search(
            r"([\d,\.]+)\s*[Kk][Gg]\s+([\d,\.]+)", full_text
        )
        if m:
            result["pieces"] = parse_number(m.group(2))

        # Value ─ línea con "Value"
        m = re.search(r"Value\s+([\d,\.]+)", full_text, re.IGNORECASE)
        if m:
            result["value"] = parse_number(m.group(1))

        # Facturas ─ después de INV#  (ej: "INV# RPT-0401-ENS / RIN-0090-ENS")
        m = re.search(r"INV#\s+(.+?)(?:\n|$)", full_text, re.IGNORECASE)
        if m:
            raw_inv = m.group(1).strip()
            result["facturas"] = [
                f.strip() for f in re.split(r"\s*/\s*", raw_inv) if f.strip()
            ]

    except Exception as e:
        st.error(f"Error leyendo PDF: {e}")
    return result


def semaforo(diferencia: float, tolerancia: float) -> tuple[str, str]:
    """Devuelve (emoji, etiqueta) según la diferencia vs tolerancia."""
    abs_diff = abs(diferencia)
    if abs_diff == 0:
        return "🟢", "Exacto"
    elif abs_diff <= tolerancia:
        return "🟡", "Redondeo"
    else:
        return "🔴", "Discrepancia"


def pct(diff: float, base: float) -> str:
    if base and base != 0:
        return f"{(diff / base * 100):+.4f}%"
    return "N/A"


def compare_facturas(excel_set: set[str], pdf_set: set[str]) -> dict:
    """Compara los conjuntos de facturas entre Excel y PDF."""
    en_ambos   = sorted(excel_set & pdf_set)
    solo_excel = sorted(excel_set - pdf_set)
    solo_pdf   = sorted(pdf_set  - excel_set)
    return {
        "en_ambos":    en_ambos,
        "solo_excel":  solo_excel,
        "solo_pdf":    solo_pdf,
        "total_excel": len(excel_set),
        "total_pdf":   len(pdf_set),
        "ok":          len(solo_excel) == 0 and len(solo_pdf) == 0,
    }


# ─────────────────────────────────────────────
# HISTÓRICO  (JSON local)
# ─────────────────────────────────────────────

def load_history() -> list:
    try:
        with open(HISTORY_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []


def save_history(record: dict):
    history = load_history()
    history.insert(0, record)
    history = history[:50]          # máximo 50 registros
    with open(HISTORY_FILE, "w", encoding="utf-8") as f:
        json.dump(history, f, ensure_ascii=False, indent=2)


# ─────────────────────────────────────────────
# EXPORTAR A EXCEL
# ─────────────────────────────────────────────

def build_excel_report(
    df_lines: pd.DataFrame,
    comparison: list[dict],
    meta: dict,
    fact_result: dict | None = None,
) -> bytes:
    wb = Workbook()

    # ── Paleta ──
    GREEN  = "C6EFCE"; GREEN_F  = "276221"
    YELLOW = "FFEB9C"; YELLOW_F = "9C6500"
    RED    = "FFC7CE"; RED_F    = "9C0006"
    BLUE_H = "1F4E79"; WHITE    = "FFFFFF"
    GRAY_L = "F2F2F2"; GRAY_M   = "D9D9D9"

    thin = Side(style="thin", color="BFBFBF")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    def hdr_style(ws, row, cols, bg=BLUE_H, fg=WHITE):
        for c in cols:
            cell = ws.cell(row=row, column=c)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font = Font(bold=True, color=fg, size=10)
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)
            cell.border = border

    def color_row(ws, row, ncols, bg, fg):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row, column=c)
            cell.fill = PatternFill("solid", fgColor=bg)
            cell.font  = Font(color=fg, size=10)
            cell.border = border
            cell.alignment = Alignment(horizontal="center", vertical="center")

    # ──────────────────────────────────────
    # Hoja 1 — Resumen comparativo
    # ──────────────────────────────────────
    ws1 = wb.active
    ws1.title = "Comparación"
    ws1.row_dimensions[1].height = 30

    # Encabezado principal
    ws1.merge_cells("A1:J1")
    ws1["A1"] = "REPORTE DE VALIDACIÓN DOCUMENTAL — CCP vs ACEM"
    ws1["A1"].font = Font(bold=True, color=WHITE, size=13)
    ws1["A1"].fill = PatternFill("solid", fgColor=BLUE_H)
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")

    # Metadata
    ws1["A2"] = "Referencia:"; ws1["B2"] = meta.get("ref", "—")
    ws1["A3"] = "Fecha:";      ws1["B3"] = meta.get("fecha", "—")
    ws1["A4"] = "Consignatario:"; ws1["B4"] = meta.get("consignee", "—")
    ws1["A5"] = "Archivo Excel:";  ws1["B5"] = meta.get("excel_name", "—")
    ws1["A6"] = "Archivo PDF:";    ws1["B6"] = meta.get("pdf_name", "—")
    for r in range(2, 7):
        ws1.cell(r, 1).font = Font(bold=True, size=10)
        ws1.cell(r, 2).font = Font(size=10)

    # Tabla comparación
    headers = ["Campo", "Etiqueta Excel", "Etiqueta PDF",
               "Valor Excel", "Valor PDF",
               "Diferencia absoluta", "Diferencia %",
               "Tolerancia", "Estado", "Observaciones"]
    row_h = 8
    ws1.row_dimensions[row_h].height = 22
    for i, h in enumerate(headers, 1):
        ws1.cell(row_h, i).value = h
    hdr_style(ws1, row_h, range(1, len(headers) + 1))

    status_colors = {"Exacto": (GREEN, GREEN_F),
                     "Redondeo": (YELLOW, YELLOW_F),
                     "Discrepancia": (RED, RED_F)}

    for i, row in enumerate(comparison, row_h + 1):
        bg, fg = status_colors.get(row["estado"], (GRAY_L, "000000"))
        cells = [
            row["campo"], row["etiqueta_excel"], row["etiqueta_pdf"],
            row["val_excel"], row["val_pdf"],
            row["diff_abs"], row["diff_pct"],
            row["tolerancia"], f"{row['emoji']} {row['estado']}",
            row["obs"],
        ]
        for j, v in enumerate(cells, 1):
            c = ws1.cell(i, j)
            c.value = v
            c.fill  = PatternFill("solid", fgColor=bg)
            c.font  = Font(color=fg, size=10)
            c.alignment = Alignment(horizontal="center", vertical="center",
                                    wrap_text=True)
            c.border = border

    # Anchos de columna hoja 1
    widths = [20, 18, 18, 15, 15, 18, 13, 12, 16, 40]
    for i, w in enumerate(widths, 1):
        ws1.column_dimensions[get_column_letter(i)].width = w

    # ──────────────────────────────────────
    # Hoja 2 — Desglose de líneas
    # ──────────────────────────────────────
    ws2 = wb.create_sheet("Desglose Excel")
    ws2.merge_cells("A1:H1")
    ws2["A1"] = "DESGLOSE DE LÍNEAS — ARCHIVO EXCEL"
    ws2["A1"].font = Font(bold=True, color=WHITE, size=12)
    ws2["A1"].fill = PatternFill("solid", fgColor=BLUE_H)
    ws2["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws2.row_dimensions[1].height = 25

    if not df_lines.empty:
        cols = list(df_lines.columns)
        for j, col in enumerate(cols, 1):
            ws2.cell(2, j).value = col
        hdr_style(ws2, 2, range(1, len(cols) + 1))

        for i, (_, r) in enumerate(df_lines.iterrows(), 3):
            bg = GRAY_L if i % 2 == 1 else WHITE
            for j, v in enumerate(r.values, 1):
                c = ws2.cell(i, j)
                c.value = v
                c.fill  = PatternFill("solid", fgColor=bg)
                c.font  = Font(size=10)
                c.border = border
                c.alignment = Alignment(horizontal="center", vertical="center")

        # Fila de totales
        tot_row = len(df_lines) + 3
        ws2.cell(tot_row, 1).value = "TOTAL"
        ws2.cell(tot_row, 1).font = Font(bold=True, size=10, color=WHITE)
        ws2.cell(tot_row, 1).fill = PatternFill("solid", fgColor=BLUE_H)
        ws2.cell(tot_row, 1).border = border
        for j, col in enumerate(cols, 1):
            if pd.api.types.is_numeric_dtype(df_lines[col]):
                c = ws2.cell(tot_row, j)
                c.value = f"=SUM({get_column_letter(j)}3:{get_column_letter(j)}{tot_row-1})"
                c.font  = Font(bold=True, size=10, color=WHITE)
                c.fill  = PatternFill("solid", fgColor=BLUE_H)
                c.border = border
                c.alignment = Alignment(horizontal="center")
            elif j > 1:
                c = ws2.cell(tot_row, j)
                c.fill  = PatternFill("solid", fgColor=BLUE_H)
                c.border = border

        for j in range(1, len(cols) + 1):
            ws2.column_dimensions[get_column_letter(j)].width = 20

    # ──────────────────────────────────────
    # Hoja 3 — Diagnóstico
    # ──────────────────────────────────────
    ws3 = wb.create_sheet("Diagnóstico")
    ws3.merge_cells("A1:D1")
    ws3["A1"] = "DIAGNÓSTICO Y OBSERVACIONES"
    ws3["A1"].font = Font(bold=True, color=WHITE, size=12)
    ws3["A1"].fill = PatternFill("solid", fgColor=BLUE_H)
    ws3["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws3.row_dimensions[1].height = 25

    exact   = sum(1 for r in comparison if r["estado"] == "Exacto")
    redond  = sum(1 for r in comparison if r["estado"] == "Redondeo")
    discr   = sum(1 for r in comparison if r["estado"] == "Discrepancia")
    total   = len(comparison)

    diag_rows = [
        ("Total campos comparados", total),
        ("Coincidencias exactas 🟢", exact),
        ("Diferencias por redondeo 🟡", redond),
        ("Discrepancias críticas 🔴", discr),
        ("Resultado global", "✅ SIN ERRORES CRÍTICOS" if discr == 0 else "⚠️ REVISAR DISCREPANCIAS"),
    ]
    for i, (label, val) in enumerate(diag_rows, 3):
        ws3.cell(i, 1).value = label
        ws3.cell(i, 2).value = val
        ws3.cell(i, 1).font = Font(bold=True, size=11)
        ws3.cell(i, 2).font = Font(size=11)
        bg = GRAY_L if i % 2 == 0 else WHITE
        for j in [1, 2]:
            ws3.cell(i, j).fill = PatternFill("solid", fgColor=bg)
            ws3.cell(i, j).border = border
            ws3.cell(i, j).alignment = Alignment(vertical="center", horizontal="center")
        ws3.row_dimensions[i].height = 20

    ws3.column_dimensions["A"].width = 35
    ws3.column_dimensions["B"].width = 35

    # ── Sección de facturas en Hoja 3 ──
    if fact_result:
        start = len(diag_rows) + 5   # deja una fila en blanco
        ws3.cell(start, 1).value = "VALIDACIÓN DE FACTURAS (INV#)"
        ws3.cell(start, 1).font  = Font(bold=True, color=WHITE, size=11)
        ws3.cell(start, 1).fill  = PatternFill("solid", fgColor=BLUE_H)
        ws3.cell(start, 1).border = border
        ws3.cell(start, 2).fill  = PatternFill("solid", fgColor=BLUE_H)
        ws3.cell(start, 2).border = border
        ws3.row_dimensions[start].height = 20

        fact_hdr = ["Factura", "En Excel", "En PDF (INV#)", "Estado"]
        hdr_style(ws3, start + 1, range(1, len(fact_hdr) + 1))
        for j, h in enumerate(fact_hdr, 1):
            ws3.cell(start + 1, j).value = h

        all_inv = sorted(
            set(fact_result.get("en_ambos", []))
            | set(fact_result.get("solo_excel", []))
            | set(fact_result.get("solo_pdf", []))
        )
        for idx, inv in enumerate(all_inv, start + 2):
            en_xl  = inv in fact_result.get("en_ambos", []) + fact_result.get("solo_excel", [])
            en_pdf = inv in fact_result.get("en_ambos", []) + fact_result.get("solo_pdf", [])
            if en_xl and en_pdf:
                estado_f, bg_f, fg_f = "✅ Coincide", GREEN, GREEN_F
            elif en_xl:
                estado_f, bg_f, fg_f = "🔴 Solo en Excel", RED, RED_F
            else:
                estado_f, bg_f, fg_f = "🔴 Solo en PDF", RED, RED_F

            vals = [inv, "✔" if en_xl else "✘", "✔" if en_pdf else "✘", estado_f]
            for j, v in enumerate(vals, 1):
                c = ws3.cell(idx, j)
                c.value = v
                c.fill  = PatternFill("solid", fgColor=bg_f)
                c.font  = Font(color=fg_f, size=10)
                c.border = border
                c.alignment = Alignment(horizontal="center", vertical="center")

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ─────────────────────────────────────────────
# UI PRINCIPAL
# ─────────────────────────────────────────────

def main():
    st.title("🔍 Validador Documental — CCP-EFO vs Manifiesto ACEM")
    st.caption("Comparación automática de valores entre el archivo Excel del SEER y el manifiesto electrónico ACEM.")

    # ── Sidebar ──
    with st.sidebar:
        st.header("📁 Cargar archivos")
        excel_file = st.file_uploader(
            "Excel (CCP-EFO)",
            type=["xlsx", "xls"],
            help="Exportación del SEER Tráfico con columnas de cantidad, peso y valor."
        )
        pdf_file = st.file_uploader(
            "PDF (Manifiesto ACEM)",
            type=["pdf"],
            help="Manifiesto electrónico generado por el sistema ACE."
        )
        st.divider()
        st.subheader("⚙️ Tolerancias")
        tol_peso  = st.number_input("Tolerancia peso (Kg)",  value=TOLERANCE_PESO,  step=0.1, format="%.2f")
        tol_valor = st.number_input("Tolerancia valor (USD)", value=TOLERANCE_VALOR, step=0.1, format="%.2f")

    if not excel_file or not pdf_file:
        st.info("⬅️ Sube el archivo Excel y el PDF en el panel izquierdo para iniciar la validación.")
        _show_history()
        return

    # ── Leer Excel ──
    try:
        df = pd.read_excel(excel_file)
    except Exception as e:
        st.error(f"No se pudo leer el Excel: {e}")
        return

    col_map = detect_columns(df)

    # Verificar que se detectaron las columnas mínimas
    missing = [k for k, v in col_map.items() if v is None]
    if missing:
        st.warning(
            f"No se detectaron automáticamente: **{', '.join(missing)}**. "
            "Selecciónalas manualmente:"
        )
        cols_available = ["(ninguna)"] + list(df.columns)
        for field in missing:
            sel = st.selectbox(f"Columna para '{field}'", cols_available)
            if sel != "(ninguna)":
                col_map[field] = sel

    # ── Leer PDF ──
    pdf_data = extract_pdf_data(pdf_file)

    # ── Calcular totales Excel ──
    def safe_sum(col):
        if col and col in df.columns:
            return pd.to_numeric(df[col], errors="coerce").sum()
        return None

    excel_cantidad = safe_sum(col_map["cantidad"])
    excel_peso     = safe_sum(col_map["peso"])
    excel_valor    = safe_sum(col_map["valor"])

    pdf_pieces = pdf_data["pieces"]
    pdf_weight = pdf_data["weight"]
    pdf_value  = pdf_data["value"]

    # ── Tabla comparación ──
    comparison = []
    for campo, etiq_xl, etiq_pdf, val_xl, val_pdf, tol, unidad in [
        ("Cantidad / Piezas",
         col_map["cantidad"] or "—", "Pieces",
         excel_cantidad, pdf_pieces, 0, "pcs"),
        ("Peso bruto",
         col_map["peso"]     or "—", "Weight (Kg)",
         excel_peso,    pdf_weight,  tol_peso, "Kg"),
        ("Valor mercancía",
         col_map["valor"]    or "—", "Value",
         excel_valor,   pdf_value,   tol_valor, "USD"),
    ]:
        if val_xl is None or val_pdf is None:
            emoji, estado = "⚪", "Sin dato"
            diff_abs = diff_pct_str = "N/A"
            obs = "Valor no encontrado en uno de los documentos."
        else:
            diff = val_xl - val_pdf
            emoji, estado = semaforo(diff, tol)
            diff_abs      = round(diff, 4)
            diff_pct_str  = pct(diff, val_pdf)
            obs = (
                "Coincidencia exacta." if estado == "Exacto"
                else f"Diferencia de {diff_abs} {unidad} atribuible a redondeo."
                if estado == "Redondeo"
                else f"Discrepancia de {diff_abs} {unidad}. Verificar con agente."
            )
        comparison.append({
            "campo": campo, "etiqueta_excel": etiq_xl, "etiqueta_pdf": etiq_pdf,
            "val_excel": round(val_xl, 4) if isinstance(val_xl, float) else val_xl,
            "val_pdf":   round(val_pdf, 4) if isinstance(val_pdf, float) else val_pdf,
            "diff_abs": diff_abs, "diff_pct": diff_pct_str,
            "tolerancia": f"±{tol} {unidad}",
            "emoji": emoji, "estado": estado, "obs": obs,
        })

    # ── MÉTRICAS ──
    st.subheader("📊 Resumen de validación")

    # Comparación de facturas
    col_factura = col_map.get("factura")
    if col_factura and col_factura in df.columns:
        excel_facturas = set(df[col_factura].dropna().astype(str).str.strip().unique())
    else:
        excel_facturas = set()

    pdf_facturas = set(pdf_data.get("facturas", []))
    fact_result  = compare_facturas(excel_facturas, pdf_facturas)

    c1, c2, c3, c4, c5 = st.columns(5)
    total_c = len(comparison)
    exact_c = sum(1 for r in comparison if r["estado"] == "Exacto")
    redond_c = sum(1 for r in comparison if r["estado"] == "Redondeo")
    discr_c  = sum(1 for r in comparison if r["estado"] == "Discrepancia")
    fact_ok  = fact_result["ok"]

    c1.metric("Campos comparados", total_c)
    c2.metric("🟢 Exactos",        exact_c)
    c3.metric("🟡 Redondeo",       redond_c)
    c4.metric("🔴 Discrepancias",  discr_c,
              delta="REVISAR" if discr_c > 0 else None,
              delta_color="inverse")
    c5.metric("🧾 Facturas",
              f"{len(fact_result['en_ambos'])}/{max(fact_result['total_excel'], fact_result['total_pdf'])}",
              delta=None if fact_ok else "Sin coincidencia",
              delta_color="inverse")

    # ── SEMÁFORO PRINCIPAL ──
    if discr_c == 0 and redond_c == 0:
        st.success("✅ Todos los valores coinciden exactamente entre el Excel y el PDF.")
    elif discr_c == 0:
        st.warning("🟡 Los documentos son consistentes. Las diferencias se deben únicamente a redondeo.")
    else:
        st.error("🔴 Se detectaron discrepancias críticas. Revisa el reporte detallado.")

    # ── SEMÁFORO FACTURAS ──
    if not excel_facturas and not pdf_facturas:
        st.info("⚪ No se encontró la columna de facturas en el Excel ni facturas en el PDF.")
    elif fact_result["ok"]:
        st.success(
            f"✅ Facturas coinciden — {len(fact_result['en_ambos'])} factura(s) presentes en ambos documentos: "
            f"**{', '.join(fact_result['en_ambos'])}**"
        )
    else:
        msgs = []
        if fact_result["solo_excel"]:
            msgs.append(f"Solo en Excel: **{', '.join(fact_result['solo_excel'])}**")
        if fact_result["solo_pdf"]:
            msgs.append(f"Solo en PDF: **{', '.join(fact_result['solo_pdf'])}**")
        st.error("🔴 Discrepancia en facturas — " + " | ".join(msgs))

    st.divider()

    # ── DETALLE FACTURAS ──
    st.subheader("🧾 Validación de facturas (INV#)")

    # Construir tabla fila a fila
    all_facturas = sorted(excel_facturas | pdf_facturas)
    if all_facturas:
        rows_fact = []
        for inv in all_facturas:
            en_xl  = inv in excel_facturas
            en_pdf = inv in pdf_facturas
            if en_xl and en_pdf:
                estado_f, emoji_f = "Coincide", "🟢"
            elif en_xl:
                estado_f, emoji_f = "Solo en Excel", "🔴"
            else:
                estado_f, emoji_f = "Solo en PDF", "🔴"
            rows_fact.append({
                "Factura":         inv,
                "En Excel":        "✔" if en_xl  else "✘",
                "En PDF (INV#)":   "✔" if en_pdf else "✘",
                "Estado":          f"{emoji_f} {estado_f}",
            })
        df_fact = pd.DataFrame(rows_fact)

        def highlight_fact(row):
            if "Coincide"    in row["Estado"]: return ["background-color:#c6efce"] * len(row)
            if "Solo en"     in row["Estado"]: return ["background-color:#ffc7ce"] * len(row)
            return [""] * len(row)

        st.dataframe(
            df_fact.style.apply(highlight_fact, axis=1),
            use_container_width=True, hide_index=True,
        )

        # Fuente de extracción
        if not col_factura:
            st.caption("⚠️ Columna de factura no detectada en el Excel — selecciónala manualmente en '🔧 Columnas detectadas'.")
        else:
            st.caption(f"Columna Excel usada: **{col_factura}** · Facturas extraídas del PDF tras INV#: {pdf_facturas or '(ninguna)'}")
    else:
        st.info("No hay facturas para comparar. Verifica que el PDF contenga una línea INV# y que el Excel tenga columna de factura.")

    st.divider()

    # ── TABLA COMPARATIVA ──
    st.subheader("📋 Tabla comparativa")
    df_cmp = pd.DataFrame([{
        "Campo":            r["campo"],
        "Etiqueta Excel":   r["etiqueta_excel"],
        "Etiqueta PDF":     r["etiqueta_pdf"],
        "Valor Excel":      r["val_excel"],
        "Valor PDF":        r["val_pdf"],
        "Diferencia abs.":  r["diff_abs"],
        "Diferencia %":     r["diff_pct"],
        "Tolerancia":       r["tolerancia"],
        "Estado":           f"{r['emoji']} {r['estado']}",
        "Observaciones":    r["obs"],
    } for r in comparison])

    def highlight_estado(row):
        e = row["Estado"]
        if "Exacto"      in e: return ["background-color:#c6efce"] * len(row)
        if "Redondeo"    in e: return ["background-color:#ffeb9c"] * len(row)
        if "Discrepancia"in e: return ["background-color:#ffc7ce"] * len(row)
        return [""] * len(row)

    st.dataframe(
        df_cmp.style.apply(highlight_estado, axis=1),
        use_container_width=True, hide_index=True,
    )

    # ── DESGLOSE EXCEL ──
    with st.expander("📄 Ver desglose de líneas del Excel"):
        st.dataframe(df, use_container_width=True, hide_index=True)

    # ── TEXTO RAW DEL PDF ──
    with st.expander("🔎 Ver texto extraído del PDF"):
        st.text(pdf_data["raw_text"] or "(sin texto extraído)")

    st.divider()

    # ── DETECTAR COLUMNAS ──
    with st.expander("🔧 Columnas detectadas automáticamente"):
        det_df = pd.DataFrame([
            {"Campo lógico": k, "Columna Excel detectada": v or "⚠️ No detectada"}
            for k, v in col_map.items()
        ])
        st.dataframe(det_df, use_container_width=True, hide_index=True)

    # ── META PARA EXPORT / HISTORY ──
    meta = {
        "ref":         pdf_data.get("ref") or excel_file.name,
        "fecha":       datetime.now().strftime("%Y-%m-%d %H:%M"),
        "consignee":   pdf_data.get("consignee", "—"),
        "excel_name":  excel_file.name,
        "pdf_name":    pdf_file.name,
    }

    # ── EXPORTAR EXCEL ──
    excel_bytes = build_excel_report(df, comparison, meta, fact_result)
    st.download_button(
        label="⬇️ Descargar reporte en Excel",
        data=excel_bytes,
        file_name=f"validacion_{meta['ref']}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # ── GUARDAR EN HISTÍRICO ──
    record = {
        "fecha":        meta["fecha"],
        "ref":          meta["ref"],
        "excel":        meta["excel_name"],
        "pdf":          meta["pdf_name"],
        "exactos":      exact_c,
        "redondeo":     redond_c,
        "discrepancias": discr_c,
        "resultado":    "SIN ERRORES" if discr_c == 0 else "CON DISCREPANCIAS",
        "detalle":      [
            {k: str(v) for k, v in r.items()}
            for r in comparison
        ]
    }
    if "ultimo_guardado" not in st.session_state or st.session_state.ultimo_guardado != meta["ref"]:
        save_history(record)
        st.session_state.ultimo_guardado = meta["ref"]

    st.divider()
    _show_history()


def _show_history():
    history = load_history()
    if not history:
        return

    st.subheader("📂 Historial de comparaciones")
    df_hist = pd.DataFrame([{
        "Fecha":          h["fecha"],
        "Referencia":     h["ref"],
        "Excel":          h["excel"],
        "PDF":            h["pdf"],
        "🟢 Exactos":     h.get("exactos", 0),
        "🟡 Redondeo":    h.get("redondeo", 0),
        "🔴 Discrepancias": h.get("discrepancias", 0),
        "Resultado":      h.get("resultado", "—"),
    } for h in history])

    def color_resultado(val):
        if "SIN ERRORES" in str(val):  return "color: green; font-weight: bold"
        if "DISCREPANCIAS" in str(val): return "color: red;  font-weight: bold"
        return ""

    st.dataframe(
        df_hist.style.applymap(color_resultado, subset=["Resultado"]),
        use_container_width=True, hide_index=True,
    )

    if st.button("🗑️ Limpiar historial"):
        with open(HISTORY_FILE, "w") as f:
            json.dump([], f)
        st.rerun()


if __name__ == "__main__":
    main()
