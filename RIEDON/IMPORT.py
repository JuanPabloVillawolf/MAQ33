"""
app.py — Supply List Enricher
Versión : 1.0.0
Descripción:
    Enriquece un Supply List con datos del Catálogo de Partes mediante
    un LEFT JOIN en "Part No." / "NumParte", aplica transformaciones de
    COO y Tracking Number, y expone una interfaz Streamlit con carga de
    archivos y descarga del resultado en formato Excel.

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
    "Country of Origin",
    "Unit Price (USD)",
    "Weight",
    "Weight Unit",
    "Item Type",
    "Tracking Number",
]

# Columnas que se requieren en el Catálogo de Partes
CATALOG_REQUIRED_COLS = ["NumParte", "Tim_Clave", "Par_DescripcionEsp", "FraccionMx"]

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
# FUNCIONES DE VALIDACIÓN
# ─────────────────────────────────────────────────────────────────────────────

def validate_columns(df: pd.DataFrame, required: list[str], file_label: str) -> None:
    """
    Verifica que un DataFrame contenga todas las columnas requeridas.

    Parámetros
    ----------
    df          : DataFrame a validar.
    required    : Lista de nombres de columnas obligatorias.
    file_label  : Nombre descriptivo del archivo (para el mensaje de error).

    Lanza
    -----
    ValueError si faltan columnas.
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
    Realiza un LEFT JOIN entre el Supply List y el Catálogo de Partes.

    Llave de unión:
        Supply List  → "Part No."
        Catálogo     → "NumParte"

    Columnas agregadas al resultado:
        Tim_Clave | Par_DescripcionEsp | FraccionMx

    Registros sin coincidencia conservan NaN en esas columnas.
    """
    # Solo se toman las columnas necesarias del catálogo para evitar
    # traer columnas extra al resultado final
    catalog_slim = catalog_df[CATALOG_REQUIRED_COLS].copy()

    enriched_df = supply_df.merge(
        catalog_slim,
        how="left",
        left_on="Part No.",
        right_on="NumParte",
    )

    # Eliminar la columna llave duplicada proveniente del catálogo
    enriched_df.drop(columns=["NumParte"], inplace=True, errors="ignore")

    return enriched_df


def apply_coo_column(df: pd.DataFrame) -> pd.DataFrame:
    """
    Crea la columna 'COO' mapeando los códigos de 'Country of Origin'
    a su nombre completo según COO_MAP.
    Si el código no existe en el mapa, conserva el valor original.
    """
    df = df.copy()
    df["COO"] = (
        df["Country of Origin"]
        .astype(str)
        .str.strip()
        .str.upper()
        .map(lambda x: COO_MAP.get(x, x))
    )
    return df


def apply_tracking_prefix(df: pd.DataFrame) -> pd.DataFrame:
    """
    Agrega el prefijo 'TRACKING ' al inicio de cada valor en
    la columna 'Tracking Number'.

    - Omite celdas vacías o NaN.
    - Evita duplicar el prefijo si ya estuviera presente.
    """
    df = df.copy()

    def _add_prefix(val):
        if pd.isna(val) or str(val).strip() == "":
            return val
        val_str = str(val).strip()
        if val_str.upper().startswith("TRACKING "):
            return val_str                    # ya tiene el prefijo, no duplicar
        return f"TRACKING {val_str}"

    df["Tracking Number"] = df["Tracking Number"].apply(_add_prefix)
    return df


# ─────────────────────────────────────────────────────────────────────────────
# FUNCIÓN PRINCIPAL DE PROCESAMIENTO
# ─────────────────────────────────────────────────────────────────────────────

def process_files(
    supply_file: io.BytesIO,
    catalog_file: io.BytesIO,
) -> pd.DataFrame:
    """
    Orquesta la lectura, validación y transformación completa.

    Parámetros
    ----------
    supply_file  : Objeto de archivo del Supply List  (.xlsx / .xls).
    catalog_file : Objeto de archivo del Catálogo     (.xlsx / .xls).

    Retorna
    -------
    DataFrame enriquecido y transformado, listo para exportar.
    """
    # ── 1. Lectura ──────────────────────────────────────────────────────────
    supply_df  = pd.read_excel(supply_file,  dtype=str)   # dtype=str evita
    catalog_df = pd.read_excel(catalog_file, dtype=str)   # conversiones no deseadas

    # Limpiar espacios en los encabezados por si vienen con padding
    supply_df.columns  = supply_df.columns.str.strip()
    catalog_df.columns = catalog_df.columns.str.strip()

    # ── 2. Validación de columnas ────────────────────────────────────────────
    validate_columns(supply_df,  SUPPLY_REQUIRED_COLS,  "Supply List")
    validate_columns(catalog_df, CATALOG_REQUIRED_COLS, "Catálogo de Partes")

    # ── 3. Enriquecimiento con el catálogo (LEFT JOIN) ───────────────────────
    result_df = enrich_with_catalog(supply_df, catalog_df)

    # ── 4. Nueva columna COO ─────────────────────────────────────────────────
    result_df = apply_coo_column(result_df)

    # ── 5. Prefijo en Tracking Number ────────────────────────────────────────
    result_df = apply_tracking_prefix(result_df)

    return result_df


# ─────────────────────────────────────────────────────────────────────────────
# EXPORTACIÓN A EXCEL EN MEMORIA
# ─────────────────────────────────────────────────────────────────────────────

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """
    Convierte un DataFrame a bytes de un archivo Excel (.xlsx) en memoria,
    sin necesidad de escribir en disco. Listo para st.download_button.
    """
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Supply List Enriquecido")
    return buffer.getvalue()


# ─────────────────────────────────────────────────────────────────────────────
# INTERFAZ STREAMLIT
# ─────────────────────────────────────────────────────────────────────────────

def main():
    # ── Configuración de página ──────────────────────────────────────────────
    st.set_page_config(
        page_title="Supply List Enricher",
        page_icon="📦",
        layout="centered",
    )

    st.title("📦 Supply List Enricher")
    st.markdown(
        "Carga tu **Supply List** y el **Catálogo de Partes** para generar "
        "automáticamente una tabla enriquecida lista para exportar."
    )
    st.divider()

    # ── Carga de archivos ────────────────────────────────────────────────────
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📄 Supply List")
        supply_file = st.file_uploader(
            "Selecciona el archivo Supply List",
            type=["xlsx", "xls"],
            key="supply",
        )

    with col2:
        st.subheader("📋 Catálogo de Partes")
        catalog_file = st.file_uploader(
            "Selecciona el Catálogo de Partes",
            type=["xlsx", "xls"],
            key="catalog",
        )

    st.divider()

    # ── Botón de procesamiento ───────────────────────────────────────────────
    if st.button("⚙️ Procesar archivos", type="primary", use_container_width=True):

        # Validar que ambos archivos estén cargados antes de procesar
        if not supply_file:
            st.warning("⚠️ Por favor carga el archivo **Supply List**.")
            st.stop()
        if not catalog_file:
            st.warning("⚠️ Por favor carga el archivo **Catálogo de Partes**.")
            st.stop()

        with st.spinner("Procesando archivos..."):
            try:
                result_df = process_files(supply_file, catalog_file)

                # Guardar resultado en session_state para que persista entre
                # reruns de Streamlit (evita reprocesar al hacer clic en descargar)
                st.session_state["result_df"] = result_df
                st.success(
                    f"✅ Procesamiento completado. "
                    f"**{len(result_df):,} registros** generados con "
                    f"**{len(result_df.columns)} columnas**."
                )

            except ValueError as ve:
                st.error(str(ve))
                st.stop()
            except Exception as ex:
                st.error(f"❌ Error inesperado: {ex}")
                st.stop()

    # ── Vista previa y descarga (solo si ya se procesó) ──────────────────────
    if "result_df" in st.session_state:
        result_df = st.session_state["result_df"]

        st.subheader("🔍 Vista previa del resultado")
        st.dataframe(result_df.head(50), use_container_width=True)

        st.caption(
            f"Mostrando las primeras 50 filas de {len(result_df):,} registros totales."
        )

        # ── Botón de descarga ────────────────────────────────────────────────
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
