# -*- coding: utf-8 -*-
"""
FLEXCON_Extractor_Universal_v2.py
==================================
App Streamlit — Extractor de documentos FLEXCON PDF → Excel

Hoja 1 : FLEXCON INV TRANSFER
Hoja 2 : SEIP Non-Inventory Transfer

Uso:
    streamlit run flexcon_extractor_universal_v2.py

# Instalación de librerías necesarias
!pip install pandas openpyxl pdfplumber tabula-py -q
print("✅ Librerías instaladas correctamente")
""" 

# Importar librerías
import pandas as pd
import pdfplumber
from google.colab import files
import re
from collections import defaultdict
import warnings
warnings.filterwarnings('ignore')

print("✅ Librerías importadas")

# Subir el archivo PDF
print("📤 Por favor, sube tu archivo PDF de FLEXCON INV TRANSFER")
print("   (Puede ser de cualquier fecha o formato)\n")
uploaded = files.upload()
pdf_filename = list(uploaded.keys())[0]
print(f"\n✅ Archivo cargado: {pdf_filename}")

# PASO 1: Análisis del PDF y detección de estructura
print("\n" + "="*80)
print("PASO 1: ANALIZANDO ESTRUCTURA DEL PDF")
print("="*80)

def analyze_pdf_structure(pdf_path):
    """Analiza el PDF para detectar su estructura y tipo de contenido"""
    info = {
        'total_pages': 0,
        'has_tables': False,
        'sample_lines': [],
        'transfer_type': None  # 'inventory' o 'non-inventory'
    }

    with pdfplumber.open(pdf_path) as pdf:
        info['total_pages'] = len(pdf.pages)

        # Analizar primera página
        first_page = pdf.pages[0]
        text = first_page.extract_text()

        # Detectar tipo de transferencia
        if 'Non-Inventory Transfer' in text:
            info['transfer_type'] = 'non-inventory'
        elif 'INV TRANSFER' in text or 'INVENTORY TRANSFER' in text.upper():
            info['transfer_type'] = 'inventory'

        # Verificar si hay tablas
        tables = first_page.extract_tables()
        info['has_tables'] = len(tables) > 0

        # Obtener líneas de muestra
        lines = text.split('\n')
        info['sample_lines'] = lines[:30]

    return info

pdf_info = analyze_pdf_structure(pdf_filename)

print(f"\n📄 Total de páginas: {pdf_info['total_pages']}")
print(f"📋 Tipo de transferencia: {pdf_info['transfer_type']}")
print(f"📊 Contiene tablas: {'Sí' if pdf_info['has_tables'] else 'No'}")
print("\n✅ Análisis completado")

# PASO 2: Extracción inteligente de datos
print("\n" + "="*80)
print("PASO 2: EXTRAYENDO DATOS DEL PDF")
print("="*80)

def extract_inventory_transfer_data(pdf_path):
    """Extrae datos de documentos FLEXCON INV TRANSFER (formato estándar)"""
    all_data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"   Procesando página {page_num}/{len(pdf.pages)}...")

            # Extraer texto completo
            text = page.extract_text()
            if not text:
                continue

            lines = text.split('\n')

            # Buscar líneas de datos (empiezan con número de línea)
            for i, line in enumerate(lines):
                # Patrón: empieza con 1-3 dígitos, seguido de espacio y otro número
                match = re.match(r'^(\d{1,3})\s+(\d{3})\s+([\w-]+)', line)

                if match:
                    line_num = match.group(1)
                    bp = match.group(2)
                    item = match.group(3)

                    # Extraer el resto de la línea
                    rest_of_line = line[match.end():].strip()

                    # Buscar LOT en la siguiente línea (si existe)
                    lot_number = ''
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        # Detectar códigos LOT (empiezan con P, W, o letras similares)
                        if re.match(r'^[A-Z]\d{8,}', next_line):
                            lot_number = next_line

                    # Parsear el resto de la línea
                    parts = rest_of_line.split()

                    # Buscar ORIGIN (USA, JPN, NLD, MEX, CHN)
                    origin = ''
                    origin_idx = -1
                    for idx, part in enumerate(parts):
                        if part in ['USA', 'JPN', 'NLD', 'MEX', 'CHN']:
                            origin = part
                            origin_idx = idx
                            break

                    if origin_idx == -1:
                        continue  # No se encontró origen, saltar esta línea

                    # DESCRIPTION: todo antes del ORIGIN
                    description = ' '.join(parts[:origin_idx])
                    if lot_number:
                        description += ' ' + lot_number

                    # Campos después del ORIGIN
                    after_origin = parts[origin_idx + 1:]

                    # Limpiar comas de los números
                    after_origin = [p.replace(',', '') for p in after_origin]

                    # Extraer campos numéricos
                    quantity = ''
                    uom = ''
                    qty_m = ''
                    boxes = ''
                    weight = ''
                    value = ''

                    # QUANTITY (primer campo después de ORIGIN)
                    if len(after_origin) >= 1:
                        quantity = after_origin[0]

                    # UOM (segundo campo)
                    if len(after_origin) >= 2:
                        uom = after_origin[1]

                    # QTY(M) (tercer campo)
                    if len(after_origin) >= 3:
                        qty_m = after_origin[2]

                    # VALUE: buscar el último número con punto decimal
                    for j in range(len(after_origin) - 1, -1, -1):
                        if re.match(r'^\d+\.\d{2}$', after_origin[j]):
                            value = after_origin[j]
                            value_idx = j

                            # BOXES y WEIGHT están entre QTY(M) y VALUE
                            middle_fields = after_origin[3:value_idx]

                            # Filtrar solo números
                            numeric_middle = []
                            for field in middle_fields:
                                # Verificar si es número (entero o decimal)
                                if re.match(r'^\d+(\.\d+)?$', field):
                                    numeric_middle.append(field)

                            # Asignar BOXES y WEIGHT
                            if len(numeric_middle) >= 1:
                                boxes = numeric_middle[0]
                            if len(numeric_middle) >= 2:
                                weight = numeric_middle[1]

                            break

                    # Crear registro
                    record = {
                        'LINE': line_num,
                        'B/P': bp,
                        'ITEM': item,
                        'DESCRIPTION/LOT': description.strip(),
                        'ORIGIN': origin,
                        'QUANTITY': quantity,
                        'UOM': uom,
                        'QTY(M)': qty_m,
                        'BOXES': boxes,
                        'WEIGHT(KG)': weight,
                        'VALUE': value
                    }

                    all_data.append(record)

    return all_data

def extract_non_inventory_transfer_data(pdf_path):
    """Extrae datos de documentos Non-Inventory Transfer"""
    all_data = []

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            print(f"   Procesando página {page_num}/{len(pdf.pages)}...")

            text = page.extract_text()
            if not text:
                continue

            lines = text.split('\n')

            # Buscar líneas de datos (empiezan con número de línea)
            for line in lines:
                # Patrón: empieza con 1-2 dígitos seguido de datos
                match = re.match(r'^(\d{1,2})\s+(\S+)', line)

                if match and int(match.group(1)) <= 20:  # Líneas típicamente 1-20
                    line_num = match.group(1)
                    item = match.group(2)

                    # Extraer el resto
                    rest = line[match.end():].strip()
                    parts = rest.split()

                    # Buscar ORIGIN
                    origin = ''
                    origin_idx = -1
                    for idx, part in enumerate(parts):
                        if part in ['USA', 'JPN', 'NLD', 'MEX', 'CHN']:
                            origin = part
                            origin_idx = idx
                            break

                    if origin_idx == -1:
                        continue

                    # Descripción antes del origen
                    description = ' '.join(parts[:origin_idx])

                    # Después del origen
                    after_origin = parts[origin_idx + 1:]
                    after_origin = [p.replace(',', '').replace('$', '') for p in after_origin]

                    quantity = ''
                    uom = ''
                    value = ''
                    boxes = ''

                    # Cantidad
                    if len(after_origin) >= 1:
                        quantity = after_origin[0]

                    # UOM (puede ser Box, GAL, PC, etc.)
                    if len(after_origin) >= 2 and not re.match(r'^\d+\.\d{2}$', after_origin[1]):
                        uom = 'N/A'  # Non-inventory suele tener N/A

                    # VALUE: buscar el último número con formato $X.XX
                    for j in range(len(after_origin) - 1, -1, -1):
                        if re.match(r'^\d+\.\d{2}$', after_origin[j]):
                            value = after_origin[j]
                            break

                    record = {
                        'LINE': line_num,
                        'B/P': '',
                        'ITEM': item,
                        'DESCRIPTION/LOT': description.strip(),
                        'ORIGIN': origin,
                        'QUANTITY': quantity,
                        'UOM': uom,
                        'QTY(M)': '',
                        'BOXES': boxes,
                        'WEIGHT(KG)': '',
                        'VALUE': value
                    }

                    all_data.append(record)

    return all_data

# Extraer según el tipo de documento
if pdf_info['transfer_type'] == 'inventory':
    print("   Tipo: INVENTORY TRANSFER")
    extracted_data = extract_inventory_transfer_data(pdf_filename)
elif pdf_info['transfer_type'] == 'non-inventory':
    print("   Tipo: NON-INVENTORY TRANSFER")
    extracted_data = extract_non_inventory_transfer_data(pdf_filename)
else:
    print("   Tipo: DESCONOCIDO - Intentando extracción genérica...")
    extracted_data = extract_inventory_transfer_data(pdf_filename)

print(f"\n✅ Datos extraídos: {len(extracted_data)} registros")

# PASO 3: Validación y limpieza de datos
print("\n" + "="*80)
print("PASO 3: VALIDANDO Y LIMPIANDO DATOS")
print("="*80)

if len(extracted_data) == 0:
    print("\n⚠️  No se pudieron extraer datos automáticamente.")
    print("    Por favor verifica que el PDF sea un documento FLEXCON válido.")
else:
    # Crear DataFrame
    df = pd.DataFrame(extracted_data)

    # Funciones de conversión segura
    def safe_int(val):
        try:
            if pd.isna(val) or val == '' or val == 'N/A':
                return None
            return int(str(val).replace(',', '').replace('$', '').strip())
        except:
            return None

    def safe_float(val):
        try:
            if pd.isna(val) or val == '' or val == 'N/A':
                return None
            return float(str(val).replace(',', '').replace('$', '').strip())
        except:
            return None

    # Convertir tipos de datos
    df['LINE'] = df['LINE'].apply(safe_int)
    df['B/P'] = df['B/P'].apply(lambda x: safe_int(x) if x else None)
    df['QUANTITY'] = df['QUANTITY'].apply(safe_int)
    df['QTY(M)'] = df['QTY(M)'].apply(safe_int)
    df['BOXES'] = df['BOXES'].apply(safe_int)
    df['WEIGHT(KG)'] = df['WEIGHT(KG)'].apply(safe_float)
    df['VALUE'] = df['VALUE'].apply(safe_float)

    # Ordenar por LINE
    df = df.sort_values('LINE').reset_index(drop=True)

    print(f"\n✅ Datos validados: {len(df)} registros")
    print(f"   - Registros con VALUE: {df['VALUE'].notna().sum()}")
    print(f"   - Total VALUE: ${df['VALUE'].sum():,.2f}" if df['VALUE'].sum() > 0 else "")

# PASO 4: Mostrar preview de datos
print("\n" + "="*80)
print("PASO 4: PREVIEW DE DATOS EXTRAÍDOS")
print("="*80)

if len(extracted_data) > 0:
    print("\nPRIMEROS 10 REGISTROS:")
    print("-" * 80)
    display(df.head(10))

    if len(df) > 10:
        print("\nÚLTIMOS 5 REGISTROS:")
        print("-" * 80)
        display(df.tail(5))

# PASO 5: Exportar a Excel con formato profesional
print("\n" + "="*80)
print("PASO 5: EXPORTANDO A EXCEL")
print("="*80)

if len(extracted_data) > 0:
    # Nombre del archivo de salida
    import os
    base_name = os.path.splitext(pdf_filename)[0]
    output_filename = f'{base_name}_EXTRACTED.xlsx'

    # Crear Excel con formato
    with pd.ExcelWriter(output_filename, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Transfer Data', index=False)

        # Obtener el workbook y worksheet
        workbook = writer.book
        worksheet = writer.sheets['Transfer Data']

        from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

        # Formato de encabezados
        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF", size=11)

        for cell in worksheet[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # Ajustar ancho de columnas
        column_widths = {
            'A': 8,   # LINE
            'B': 8,   # B/P
            'C': 16,  # ITEM
            'D': 60,  # DESCRIPTION/LOT
            'E': 10,  # ORIGIN
            'F': 13,  # QUANTITY
            'G': 8,   # UOM
            'H': 11,  # QTY(M)
            'I': 10,  # BOXES
            'J': 13,  # WEIGHT(KG)
            'K': 14   # VALUE
        }

        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width

        # Formato numérico
        for row in range(2, len(df) + 2):
            # VALUE con formato de moneda
            if worksheet[f'K{row}'].value is not None:
                worksheet[f'K{row}'].number_format = '$#,##0.00'
            # QUANTITY con separador de miles
            if worksheet[f'F{row}'].value is not None:
                worksheet[f'F{row}'].number_format = '#,##0'
            # QTY(M) con separador de miles
            if worksheet[f'H{row}'].value is not None:
                worksheet[f'H{row}'].number_format = '#,##0'

        # Agregar bordes
        thin_border = Border(
            left=Side(style='thin', color='000000'),
            right=Side(style='thin', color='000000'),
            top=Side(style='thin', color='000000'),
            bottom=Side(style='thin', color='000000')
        )

        for row in worksheet.iter_rows(min_row=1, max_row=len(df)+1, min_col=1, max_col=11):
            for cell in row:
                cell.border = thin_border
                if cell.row > 1:
                    cell.alignment = Alignment(vertical='center', wrap_text=False)

# PASO 6: Descargar archivo
print("\n" + "="*80)
print("PASO 6: DESCARGANDO ARCHIVO")
print("="*80)

if len(extracted_data) > 0:
    files.download(output_filename)
    print(f"\n✅ Descarga iniciada: {output_filename}")

# ESTADÍSTICAS FINALES
print("\n" + "="*80)
print("📊 ESTADÍSTICAS FINALES")
print("="*80)

if len(extracted_data) > 0:
    print(f"\n📄 Archivo procesado: {pdf_filename}")
    print(f"📋 Tipo de documento: {pdf_info['transfer_type'].upper() if pdf_info['transfer_type'] else 'INVENTORY'}")
    print(f"📃 Páginas procesadas: {pdf_info['total_pages']}")
    print(f"\n{'RESUMEN DE DATOS':^80}")
    print("-" * 80)
    print(f"Total de Items (Líneas): {len(df)}")
    print(f"Registros con VALUE: {df['VALUE'].notna().sum()} ({df['VALUE'].notna().sum()/len(df)*100:.1f}%)")

    if df['BOXES'].sum() > 0:
        print(f"Total de Cajas: {df['BOXES'].sum():.0f}")
    if df['WEIGHT(KG)'].sum() > 0:
        print(f"Peso Total (KG): {df['WEIGHT(KG)'].sum():.2f}")

    print(f"\n💰 VALOR TOTAL: ${df['VALUE'].sum():,.2f}")

    # Distribución por ORIGIN
    if df['ORIGIN'].notna().sum() > 0:
        print(f"\n{'DISTRIBUCIÓN POR ORIGEN':^80}")
        print("-" * 80)
        origin_stats = df.groupby('ORIGIN').agg({
            'LINE': 'count',
            'VALUE': 'sum'
        }).round(2)
        origin_stats.columns = ['Items', 'Total Value ($)']
        print(origin_stats.to_string())

    # Distribución por UOM
    if df['UOM'].notna().sum() > 0:
        print(f"\n{'DISTRIBUCIÓN POR UNIDAD DE MEDIDA':^80}")
        print("-" * 80)
        uom_stats = df.groupby('UOM').agg({
            'LINE': 'count',
            'QUANTITY': 'sum'
        }).round(0)
        uom_stats.columns = ['Items', 'Cantidad Total']
        print(uom_stats.to_string())

    # Top items por valor
    if df['VALUE'].sum() > 0:
        print(f"\n{'TOP 5 ITEMS POR VALOR':^80}")
        print("-" * 80)
        top_items = df.nlargest(5, 'VALUE')[['LINE', 'ITEM', 'DESCRIPTION/LOT', 'VALUE']]
        for idx, row in top_items.iterrows():
            print(f"{int(row['LINE']):>3}. {row['ITEM']:<18} ${row['VALUE']:>12,.2f}")
            desc = row['DESCRIPTION/LOT'][:75]
            print(f"     {desc}")

    print("\n" + "="*80)
    print("✅ PROCESO COMPLETADO EXITOSAMENTE")
    print("="*80)
else:
    print("\n❌ No se pudieron extraer datos del PDF.")
    print("   Verifica que el archivo sea un documento FLEXCON válido.")
