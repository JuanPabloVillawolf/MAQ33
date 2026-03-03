"""
app.py — Supply List Enricher
Versión : 1.2.0
Descripción:
    Enriquece un Supply List con datos del Catálogo de Partes mediante
    un LEFT JOIN en "Part No." / "NumParte", aplica transformaciones de
    COO y Tracking Number, excluye columnas internas, reordena el resultado
    y exporta a Excel. Acepta archivos .xlsx, .xls y .csv.

Uso:
    streamlit run app.py
"""

import io
import pandas as pd
import streamlit as st


# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTES
# ─────────────────────────────────────────────────────────────────────────────

# Columnas obligatorias del Supply List
SUPPLY_REQUIRED_COLS = [
    "Part No.",
    "Po Number",
    "Qty",
    "Unit of Measure",
    "County of Origin (Made In)",
    "Unit Price (USD)",
    "Weight",
    "Weight Unit",
    "Item Type",
    "Tracking Number",
]

# Columnas que se requieren en el Catálogo de Partes
CATALOG_REQUIRED_COLS = [
    "NumParte",
    "Tim_Clave",
    "Par_DescripcionEsp",
    "FraccionMX",
    "UM",          # ← extraída del catálogo
]

# Columnas internas que NO deben aparecer en el archivo final
COLUMNS_TO_EXCLUDE = [
    "Created By",
    "Description",
    "Brand",
    "Model",
    "Series",
    "Other",
    "Vendor Link",
    "Supplier",
    "Packing Slip Number",
    "Shipping Company",
    "Southbound Demand Level",
    "Date of Shipment",
    "Estimated Arrival Date",
    "Tracking Received By",
    "Actual Arrival Date",
    "Status Invoice",
    "Invoice Number",
    "No. Entry Warehouse",
    "Size of Entry Warehouse",
]

# Orden final de columnas en el Excel de salida.
# "Part No." se renombra a "NumParte" para consistencia con el catálogo.
# "Tracking" es copia de "Tracking Number" (con prefijo ya aplicado).
FINAL_COLUMN_ORDER = [
    "NumParte",
    "Po Number",
    "Qty",
    "Unit of Measure",
    "Country of Origin",
    "Unit Price (USD)",
    "Weight",
    "Weight Unit",
    "Item Type",
    "Tracking Number",
    "Tim_Clave",
    "Par_DescripcionEsp",
    "FraccionMX",
    "Tracking",
    "COO",
    "UM",
]

# Mapeo de códigos ISO → nombre completo para la columna COO
COO_MAP = {
    "CHN": "CHINA",
    "CRI": "COSTA RICA",
    "USA": "USA",
    "KOR": "KOREA",
    "TWN": "TAIWAN",
    "HKG": "HONG KONG",
    "VNM": "VIETNAM",
    "CAN": "CANADA",
}


# ─────────────────────────────────────────────────────────────────────────────
# LECTURA DE ARCHIVOS (EXCEL O CSV)
# ─────────────────────────────────────────────────────────────────────────────

def read_file(file: io.BytesIO, label: str) -> pd.DataFrame:
    """
    Lee un archivo y retorna un DataFrame.
    Soporta .xlsx, .xls y .csv.

    Parámetros
    ----------
    file  : Objeto de archivo cargado por Streamlit.
    label : Nombre descriptivo para mensajes de error.

    Lanza
    -----
    ValueError si el formato no es soportado.
    """
    name = file.name.lower()

    if name.endswith((".xlsx", ".xls")):
        df = pd.read_excel(file, dtype=str)
    elif name.endswith(".csv"):
        # Intentar detectar separador automáticamente (coma o punto y coma)
        raw = file.read()
        file.seek(0)  # rebobinar para que pandas pueda leerlo
        sample = raw[:2048].decode("utf-8", errors="ignore")
        sep = ";" if sample.count(";") > sample.count(",") else ","
        df = pd.read_csv(file, dtype=str, sep=sep, encoding="utf-8-sig")
    else:
        raise ValueError(
            f"❌ Formato no soportado en '{label}'. "
            f"Usa archivos .xlsx, .xls o .csv."
        )

    # Limpiar espacios en los encabezados
    df.columns = df.columns.str.strip()
    return df


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES DE VALIDACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def validate_columns(df: pd.DataFrame, required: list[str], file_label: str) -> None:
    """
    Verifica que un DataFrame contenga todas las columnas requeridas.
    Lanza ValueError con el detalle de las faltantes.
    """
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(
            f"❌ El archivo '{file_label}' no tiene las columnas requeridas:\n"
            f"   Faltantes → {missing}\n"
            f"   Presentes → {list(df.columns)}"
        )


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES DE TRANSFORMACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def enrich_with_catalog(
    supply_df: pd.DataFrame,
    catalog_df: pd.DataFrame,
) -> pd.DataFrame:
    """
    LEFT JOIN entre Supply List ("Part No.") y Catálogo ("NumParte").
    Agrega: Tim_Clave | Par_DescripcionEsp | FraccionMX | UM
    """
    catalog_slim = catalog_df[CATALOG_REQUIRED_COLS].copy()

    enriched_df = supply_df.merge(
        catalog_slim,
        how="left",
        left_on="Part No.",
        right_on="NumParte",
    )

    # Eliminar columna llave duplicada del catálogo
    enriched_df.drop(columns=["NumParte"], inplace=True, errors="ignore")

    return enriched_df


def apply_coo_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea la columna 'COO' a partir de 'County of Origin (Made In)'
    usando el mapa COO_MAP. Conserva el valor original si no hay match.
    """
    df = df.copy()
    df["COO"] = (
        df["County of Origin (Made In)"]
        .astype(str)
        .str.strip()
        .str.upper()
        .map(lambda x: COO_MAP.get(x, x))
    )
    return df

def extract_tracking_value(df: pd.DataFrame) -> pd.DataFrame:
    """
    Limpia la columna 'Tracking Number' dejando únicamente el valor
    sin el prefijo 'TRACKING' o 'TRACKING:'.
    """
    df = df.copy()

    def _clean_value(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        
        val_str = str(val).strip()

        # Elimina TRACKING o TRACKING:
        if val_str.upper().startswith("TRACKING"):
            val_str = val_str.replace("TRACKING:", "")
            val_str = val_str.replace("TRACKING", "")
        
        return val_str.strip()

    df["Tracking Number"] = df["Tracking Number"].apply(_clean_value)
    return df

def apply_tracking_prefix(df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega el prefijo 'TRACKING:' a la columna 'Tracking Number'.
    Omite NaN/vacíos y evita duplicar el prefijo.
    """
    df = df.copy()

    def _add_prefix(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        
        val_str = str(val).strip()

        if val_str.upper().startswith("TRACKING"):
            return val_str
        
        return f"TRACKING: {val_str}"

    df["Tracking Number"] = df["Tracking Number"].apply(_add_prefix)
    return df




def drop_excluded_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Elimina las columnas internas listadas en COLUMNS_TO_EXCLUDE.
    Ignora silenciosamente las que no existan.
    """
    return df.drop(columns=COLUMNS_TO_EXCLUDE, errors="ignore")


def rename_and_reorder_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    1. Renombra 'Part No.' → 'NumParte' y
       'County of Origin (Made In)' → 'Country of Origin'.
    2. Reordena las columnas según FINAL_COLUMN_ORDER.
       Las columnas no listadas se descartan del resultado final.
    """
    df = df.copy()

    # Renombrar columnas para consistencia con el orden final
    df.rename(columns={
        "Part No.":                  "NumParte",
        "County of Origin (Made In)": "Country of Origin",
    }, inplace=True)

    # Filtrar solo las columnas del orden final que existan en el DataFrame
    cols_present = [c for c in FINAL_COLUMN_ORDER if c in df.columns]
    return df[cols_present]


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN PRINCIPAL DE PROCESAMIENTO
# ─────────────────────────────────────────────────────────────────────────────

def process_files(
    supply_file: io.BytesIO,
    catalog_file: io.BytesIO,
) -> pd.DataFrame:
    """
    Orquesta lectura, validación y todas las transformaciones.

    Retorna DataFrame listo para exportar.
    """
    # ── 1. Lectura ───────────────────────────────────────────────────────────
    supply_df  = read_file(supply_file,  "Supply List")
    catalog_df = read_file(catalog_file, "Catálogo de Partes")

    # ── 2. Validación ────────────────────────────────────────────────────────
    validate_columns(supply_df,  SUPPLY_REQUIRED_COLS,  "Supply List")
    validate_columns(catalog_df, CATALOG_REQUIRED_COLS, "Catálogo de Partes")

    # ── 3. LEFT JOIN con catálogo ────────────────────────────────────────────
    result_df = enrich_with_catalog(supply_df, catalog_df)

    # ── 4. Columna COO ───────────────────────────────────────────────────────
    result_df = apply_coo_column(result_df)

    # ── 5. Prefijo TRACKING en Tracking Number ───────────────────────────────
    result_df = apply_tracking_prefix(result_df)

    # ── 6. Columna Tracking (copia de Tracking Number) ───────────────────────
    result_df = add_tracking_copy(result_df)

    # ── 7. Eliminar columnas internas ────────────────────────────────────────
    result_df = drop_excluded_columns(result_df)

    # ── 8. Renombrar y reordenar columnas finales ────────────────────────────
    result_df = rename_and_reorder_columns(result_df)

    return result_df


# ─────────────────────────────────────────────────────────────────────────────
# EXPORTACIÓN A EXCEL EN MEMORIA
# ─────────────────────────────────────────────────────────────────────────────

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    Convierte un DataFrame a bytes .xlsx en memoria.
    Listo para st.download_button.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Supply List Enriquecido")
    return buffer.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Supply List Enricher",
        page_icon="📦",
        layout="centered",
    )

    st.title("📦 Supply List Enricher")
    st.markdown(
        "Carga tu **Supply List** y el **Catálogo de Partes** para generar "
        "automáticamente una tabla enriquecida lista para exportar. "
        "Acepta archivos **.xlsx**, **.xls** y **.csv**."
    )
    st.divider()

    # ── Carga de archivos ────────────────────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 Supply List")
        supply_file = st.file_uploader(
            "Selecciona el archivo Supply List",
            type=["xlsx", "xls", "csv"],
            key="supply",
        )

    with col2:
        st.subheader("📋 Catálogo de Partes")
        catalog_file = st.file_uploader(
            "Selecciona el Catálogo de Partes",
            type=["xlsx", "xls", "csv"],
            key="catalog",
        )

    st.divider()

    # ── Botón de procesamiento ───────────────────────────────────────────────
    if st.button("⚙️ Procesar archivos", type="primary", use_container_width=True):

        if not supply_file:
            st.warning("⚠️ Por favor carga el archivo **Supply List**.")
            st.stop()
        if not catalog_file:
            st.warning("⚠️ Por favor carga el archivo **Catálogo de Partes**.")
            st.stop()

        with st.spinner("Procesando archivos..."):
            try:
                result_df = process_files(supply_file, catalog_file)
                st.session_state["result_df"] = result_df
                st.success(
                    f"✅ Procesamiento completado. "
                    f"**{len(result_df):,} registros** · "
                    f"**{len(result_df.columns)} columnas**."
                )
            except ValueError as ve:
                st.error(str(ve))
                st.stop()
            except Exception as ex:
                st.error(f"❌ Error inesperado: {ex}")
                st.stop()

    # ── Vista previa y descarga ──────────────────────────────────────────────
    if "result_df" in st.session_state:
        result_df = st.session_state["result_df"]

        st.subheader("🔍 Vista previa del resultado")
        st.dataframe(result_df.head(50), use_container_width=True)
        st.caption(
            f"Mostrando las primeras 50 filas de {len(result_df):,} registros totales."
        )

        excel_bytes = to_excel_bytes(result_df)
        st.download_button(
            label="⬇️ Descargar Excel enriquecido",
            data=excel_bytes,
            file_name="supply_list_enriquecido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


# ─────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    main()
