# -*- coding: utf-8 -*-
"""FLEXCON_Extractor_v4.py — Streamlit App
Uso: streamlit run flexcon_extractor_v4.py
"""
import re, io, os, warnings
import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
warnings.filterwarnings("ignore")

# ── Zonas de columna ─────────────────────────────────────────────────────────
INV_COLS = {
    "LINE":(0,43),"B_P":(43,72),"ITEM":(72,143),"DESCRIPTION":(143,369),
    "ORIGIN":(369,427),"QUANTITY":(427,491),"UOM":(491,533),
    "QTY_M":(533,583),"VALUE":(712,9999),
}
NI_COLS = {
    "LINE":(0,100),"ITEM_NUMBER":(100,230),"DESCRIPTION":(230,483),
    "UOM":(483,515),"ORIGIN":(515,545),"QTY":(545,605),
    "LOT":(605,655),"VALUE":(655,710),
}
ORIGINS    = {"USA","JPN","NLD","MEX","CHN","KOR","TWN","DEU"}
LOT_RE     = re.compile(r'^[A-Za-z][Oo0-9]{7,}$')
FORM_WORDS = {"shipped","received","shipping","notes","date","total",
              "value","boxes","by","transfer","truck","flexcon","seip"}
COLUMNS    = ['Transfer Date','Doc ID','Doc Seq','Page','Line','B/P','Item',
              'Description','Lot','Origin','Quantity','UOM','QTY(M)',
              'Boxes','Weight(KG)','Value','Doc Type']
COL_WIDTHS = {'Transfer Date':13,'Doc ID':11,'Doc Seq':8,'Page':6,'Line':6,
              'B/P':6,'Item':13,'Description':52,'Lot':16,'Origin':8,
              'Quantity':11,'UOM':7,'QTY(M)':9,'Boxes':7,'Weight(KG)':11,
              'Value':12,'Doc Type':22}

# ── Utilidades ───────────────────────────────────────────────────────────────
def col_of(x, zones):
    for k,(lo,hi) in zones.items():
        if lo<=x<hi: return k
    return None

def is_garbage(t):
    t=t.strip()
    if not t: return True
    return sum(1 for c in t if c.isalnum() or c in '$.,') / max(len(t),1) < 0.45

def fix_value(s):
    """2.279.90 → 2279.90  |  3,582.70 → 3582.70"""
    s = s.strip().replace('$','').replace(' ','')
    if not s: return None
    parts = s.split('.')
    if len(parts) >= 3:
        s = ''.join(parts[:-1]) + '.' + parts[-1]
    s = s.replace(',','')
    try:
        v = float(s)
        return v if v > 0 else None
    except: return None

def parse_qty(s):
    s = s.strip().replace(',','')
    if re.match(r'^\d+\.\d{3,}$', s): return s.replace('.','')
    if re.match(r'^\d+$', s): return s
    return ''

def ro(items, bucket=3):
    """Reading order: sort by (rounded_top, x0)."""
    return [t for _,_,t in sorted(items, key=lambda w:(round(w[0]/bucket),w[1]))]

def ptext(page):
    try: return page.extract_text(x_tolerance=3,y_tolerance=3) or ""
    except: return page.extract_text() or ""

def pwords(page):
    try: return page.extract_words(x_tolerance=3,y_tolerance=3)
    except: return page.extract_words()

def clean_words(page, min_top=70):
    return [{'x0':w['x0'],'top':w['top'],'text':w['text'].strip()}
            for w in pwords(page)
            if w['text'].strip() and not is_garbage(w['text']) and w['top']>min_top]

# ── Detección de metadatos de página ─────────────────────────────────────────
def detect_meta(page):
    txt   = ptext(page)
    words = pwords(page)
    hw    = ' '.join(w['text'] for w in words if w['top']<55).upper()
    is_ni = any(m in txt.lower() for m in
                ['non-inventory transfer','non-lnventory','requestor name'])
    is_inv= any(m in hw for m in ['INV TRANSFER','FLEXCON INV'])
    doc_type = 'NON_INV' if is_ni else 'INV'

    date = ''
    for pat in [r'Transfer Date[:\s]+(\d{1,2}/\d{1,2}/\d{2,4})',
                r'Shipment Dal[e;]+[:\s]+(\d{2}/\d{2}/\d{2,4})',
                r'Shipment Date[:\s]+(\d{2}/\d{2}/\d{2,4})']:
        m = re.search(pat, txt, re.I)
        if m: date = m.group(1).strip(); break
    if not date:
        m = re.search(r'\b(\d{1,2}/\d{1,2}/\d{2,4})\b', hw)
        if m: date = m.group(1)

    pg = 1
    m = re.search(r'Page[;:\s]+(\d+)', txt, re.I)
    if m: pg = int(m.group(1))

    ts = ''
    m = re.search(r'\b(\d{2}:\d{2}:\d{2})\b', txt)
    if m: ts = m.group(1)

    doc_id = ''
    m = re.search(r'\b(P\d{7,})\b', txt)
    if m: doc_id = m.group(1)

    tv = None
    for pat in [r'Total\s+Value[:\s]+\$?\s*([\d,.]+)',
                r'TOTAL\s+VALUE[;:\s]+\$?([\d,.]+)']:
        m = re.search(pat, txt, re.I)
        if m:
            tv = fix_value(m.group(1))
            if tv: break

    return dict(doc_type=doc_type,date=date,page_num=pg,
                timestamp=ts,doc_id=doc_id,total_val=tv)

# ── Extracción INV (una página) ───────────────────────────────────────────────
def extract_inv_page(page, meta):
    cw_list = clean_words(page, min_top=70)

    anchors = []
    for cw in cw_list:
        if col_of(cw['x0'], INV_COLS)=='LINE':
            if re.match(r'^\d{1,3}$', cw['text']) and 1<=int(cw['text'])<=99:
                anchors.append((cw['top'], int(cw['text'])))
    anchors.sort(key=lambda x:x[0])

    records = []
    for anchor_top, line_num in anchors:

        def coll(keys, dy):
            r={k:[] for k in keys}
            for cw in cw_list:
                if abs(cw['top']-anchor_top)<=dy:
                    c=col_of(cw['x0'],INV_COLS)
                    if c in keys: r[c].append((cw['top'],cw['x0'],cw['text']))
            return r

        def coll_below(keys, dmin, dmax):
            r={k:[] for k in keys}
            for cw in cw_list:
                dy=cw['top']-anchor_top
                if dmin<=dy<=dmax:
                    c=col_of(cw['x0'],INV_COLS)
                    if c in keys: r[c].append(cw['text'])
            return r

        same   = coll(['B_P','ITEM','DESCRIPTION','ORIGIN','QUANTITY'], 6)
        flt    = coll(['UOM','QTY_M','VALUE'], 18)
        below  = coll_below(['DESCRIPTION'], 6, 20)

        bp   = ' '.join(ro(same['B_P']))
        item = ' '.join(ro(same['ITEM']))
        desc = ' '.join(ro(same['DESCRIPTION']))
        if bp and not re.match(r'^\d{3}$', bp.strip()): bp=''

        origin = next((t for _,_,t in sorted(same['ORIGIN'],
                       key=lambda w:(round(w[0]/3),w[1])) if t.upper() in ORIGINS),'')

        qlist  = [parse_qty(t) for _,_,t in sorted(same['QUANTITY'],
                  key=lambda w:(round(w[0]/3),w[1])) if parse_qty(t)]
        qty    = int(qlist[0]) if qlist else None

        ulist  = [t.upper() for _,_,t in sorted(flt['UOM'],
                  key=lambda w:(round(w[0]/3),w[1])) if not is_garbage(t) and len(t)<=5]
        uom    = ulist[0].lstrip('-').strip() if ulist else ''

        qmlist = [parse_qty(t) for _,_,t in sorted(flt['QTY_M'],
                  key=lambda w:(round(w[0]/3),w[1])) if parse_qty(t)]
        qtym   = int(qmlist[0]) if qmlist else None

        vlist  = []
        for _,_,t in sorted(flt['VALUE'], key=lambda w:(round(w[0]/3),w[1])):
            v=fix_value(t)
            if v: vlist.append(v)
        value = vlist[-1] if vlist else None

        lot = next((t for t in below['DESCRIPTION']
                    if LOT_RE.match(t.replace('O','0').replace('o','0'))),'')
        if lot: desc=(desc+' '+lot).strip()

        records.append({
            'Transfer Date': meta['date'],
            'Doc ID':        meta['doc_id'],
            'Page':          meta['page_num'],
            'Line':          line_num,
            'B/P':           bp,
            'Item':          item,
            'Description':   desc,
            'Lot':           lot,
            'Origin':        origin.upper() if origin else '',
            'Quantity':      qty,
            'UOM':           uom,
            'QTY(M)':        qtym,
            'Boxes':         '',
            'Weight(KG)':    '',
            'Value':         value,
            'Doc Type':      'INV Transfer',
        })
    return records

# ── Extracción Non-Inv (una página) ──────────────────────────────────────────
def extract_ni_page(page, meta):
    cw_list = clean_words(page, min_top=95)

    anchors = []
    for cw in cw_list:
        if col_of(cw['x0'],NI_COLS)=='LINE':
            if re.match(r'^\d{1,2}$',cw['text']) and 1<=int(cw['text'])<=20:
                anchors.append((cw['top'],int(cw['text'])))
    anchors.sort(key=lambda x:x[0])

    records = []
    for i,(anchor_top,line_num) in enumerate(anchors):
        y_hi = anchors[i+1][0]-2 if i+1<len(anchors) else anchor_top+25
        lo,hi = anchor_top-3, y_hi

        combined={}
        for cw in cw_list:
            if lo<=cw['top']<=hi:
                c=col_of(cw['x0'],NI_COLS)
                if c and c!='LINE':
                    combined.setdefault(c,[]).append((cw['top'],cw['x0'],cw['text']))

        item = ' '.join(ro(combined.get('ITEM_NUMBER',[])))
        desc = ' '.join(ro(combined.get('DESCRIPTION',[])))
        uom  = ' '.join(ro(combined.get('UOM',[]))).upper()
        origin = next((t for _,_,t in combined.get('ORIGIN',[]) if t.upper() in ORIGINS),'')
        qlist = [parse_qty(t) for _,_,t in combined.get('QTY',[]) if parse_qty(t)]
        qty   = int(qlist[0]) if qlist else None
        lot   = next((t for _,_,t in combined.get('LOT',[]) if t.upper() in ('N/A','NA')),'N/A')
        vlist = []
        for _,_,t in combined.get('VALUE',[]):
            v=fix_value(t.replace('$',''))
            if v: vlist.append(v)
        value = vlist[0] if vlist else None

        if not item.strip() and not desc.strip(): continue
        if any(fw in item.lower() for fw in FORM_WORDS): continue

        records.append({
            'Transfer Date': meta['date'],
            'Doc ID':        '',
            'Page':          meta['page_num'],
            'Line':          line_num,
            'B/P':           '',
            'Item':          item,
            'Description':   desc,
            'Lot':           lot,
            'Origin':        origin.upper() if origin else '',
            'Quantity':      qty,
            'UOM':           uom,
            'QTY(M)':        None,
            'Boxes':         '',
            'Weight(KG)':    '',
            'Value':         value,
            'Doc Type':      'Non-Inventory Transfer',
        })
    return records

# ── Pipeline principal ────────────────────────────────────────────────────────
def extract_all(pdf_bytes):
    """Procesa todo el PDF en orden físico. Nunca reordena."""
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        page_metas = [(p, detect_meta(p)) for p in pdf.pages]

    # Asignar doc_seq por cambio de timestamp
    cur_ts, cur_seq = None, 0
    for _, meta in page_metas:
        if meta['timestamp'] != cur_ts:
            cur_seq += 1; cur_ts = meta['timestamp']
        meta['doc_seq'] = cur_seq

    all_records   = []
    doc_summaries = {}

    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        for page_obj, meta in zip(pdf.pages, [m for _,m in page_metas]):
            if meta['doc_type']=='NON_INV':
                recs = extract_ni_page(page_obj, meta)
            else:
                recs = extract_inv_page(page_obj, meta)
            for r in recs:
                r['Doc Seq'] = meta['doc_seq']
            all_records.extend(recs)

            seq = meta['doc_seq']
            if seq not in doc_summaries:
                doc_summaries[seq] = dict(doc_seq=seq, doc_id=meta['doc_id'],
                    date=meta['date'], doc_type=meta['doc_type'],
                    timestamp=meta['timestamp'], total_val=None)
            if meta['total_val']:
                doc_summaries[seq]['total_val'] = meta['total_val']

    return all_records, list(doc_summaries.values())

# ── Construcción Excel ────────────────────────────────────────────────────────
def _fill(h): return PatternFill(start_color=h,end_color=h,fill_type='solid')
def _side(s='thin',c='C0C0C0'): return Side(style=s,color=c)

def write_sheet(ws, df, header_hex, alt_hex, total_hex, title):
    thick = _side('medium','404040')
    # Título
    ws.append([title])
    ws.merge_cells(f'A1:{get_column_letter(len(COLUMNS))}1')
    ws['A1'].font=Font(bold=True,italic=True,name='Arial',size=10,color='404040')
    ws['A1'].alignment=Alignment(horizontal='left',vertical='center')
    ws.row_dimensions[1].height=18
    # Encabezados
    ws.append(COLUMNS)
    for cell in ws[2]:
        cell.fill=_fill(header_hex)
        cell.font=Font(bold=True,color='FFFFFF',name='Arial',size=10)
        cell.border=Border(left=thick,right=thick,top=thick,bottom=thick)
        cell.alignment=Alignment(horizontal='center',vertical='center',wrap_text=True)
    ws.row_dimensions[2].height=28
    # Datos
    alt  = _fill(alt_hex)
    bord = Border(left=_side(),right=_side(),top=_side(),bottom=_side())
    for ri,(_, row) in enumerate(df.iterrows(), 3):
        ws.append([row.get(c,'') if pd.notna(row.get(c,'')) else '' for c in COLUMNS])
        for cell in ws[ri]:
            cell.border=bord; cell.font=Font(name='Arial',size=9)
            cell.alignment=Alignment(vertical='center')
            if ri%2==0: cell.fill=alt
    # Formato numérico
    hdr={c.value:c.column_letter for c in ws[2]}
    ds,de=3,2+len(df)
    for ri in range(ds,de+1):
        for col in ['Quantity','QTY(M)']:
            if col in hdr: ws[f"{hdr[col]}{ri}"].number_format='#,##0'
        if 'Value' in hdr: ws[f"{hdr['Value']}{ri}"].number_format='$#,##0.00'
    # Totales
    tr=de+1
    ws.append(['TOTALES']+['']*( len(COLUMNS)-1))
    for col in ['Quantity','QTY(M)']:
        if col in hdr:
            cl=hdr[col]; ws[f'{cl}{tr}']=f'=IFERROR(SUM({cl}{ds}:{cl}{de}),"")'; ws[f'{cl}{tr}'].number_format='#,##0'
    if 'Value' in hdr:
        cl=hdr['Value']; ws[f'{cl}{tr}']=f'=IFERROR(SUM({cl}{ds}:{cl}{de}),"")'; ws[f'{cl}{tr}'].number_format='$#,##0.00'
    for cell in ws[tr]:
        cell.fill=_fill(total_hex); cell.font=Font(bold=True,name='Arial',size=10)
        cell.border=Border(left=thick,right=thick,top=thick,bottom=thick)
    # Anchos
    for i,col in enumerate(COLUMNS,1):
        ws.column_dimensions[get_column_letter(i)].width=COL_WIDTHS.get(col,12)
    ws.freeze_panes='A3'
    ws.auto_filter.ref=f'A2:{get_column_letter(len(COLUMNS))}{de}'

def build_excel(all_records, doc_summaries):
    wb=Workbook()
    df_all=pd.DataFrame(all_records)
    if 'Doc Seq' not in df_all.columns: df_all['Doc Seq']=1
    # Asegurar columnas COLUMNS presentes
    for c in COLUMNS:
        if c not in df_all.columns: df_all[c]=''

    # Hoja 1 — Tabla unificada en orden físico
    ws1=wb.active; ws1.title='Todos los Registros'
    doc_info='  |  '.join(
        f"Doc{d['doc_seq']}: {d['date']} {d['doc_type']}"
        +(f" TV=${d['total_val']:,.2f}" if d['total_val'] else '')
        for d in doc_summaries)
    write_sheet(ws1,df_all,'1F4E78','EEF3F8','C9D8EC',
                f'FLEXCON — Todos los registros en orden físico  |  {doc_info}')

    # Hojas por documento lógico
    colors=[('2E5FA3','EEF3F8','C9D8EC'),('1A6B3C','EDF7EF','C6E8CC'),
            ('7B3F9E','F5EEF8','DCC6EC'),('B05000','FEF3E8','F5D5B0')]
    for ds_meta in doc_summaries:
        seq=ds_meta['doc_seq']
        df_d=df_all[df_all['Doc Seq']==seq].copy()
        if df_d.empty: continue
        dt=ds_meta['date'].replace('/','_')
        dtype='INV' if ds_meta['doc_type']=='INV' else 'NonInv'
        name=f"Doc{seq} {dt} {dtype}"[:31]
        ws=wb.create_sheet(name)
        hc,ac,tc=colors[(seq-1)%len(colors)]
        tv=f'  Total declarado: ${ds_meta["total_val"]:,.2f}' if ds_meta['total_val'] else ''
        write_sheet(ws,df_d,hc,ac,tc,
                    f'{ds_meta["doc_id"]}  Fecha: {ds_meta["date"]}  Tipo: {ds_meta["doc_type"]}{tv}')

    buf=io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()

# ── Streamlit UI ──────────────────────────────────────────────────────────────
st.set_page_config(page_title='FLEXCON Extractor v4',page_icon='📦',layout='wide')
st.title('📦 FLEXCON PDF → Excel  (v4 — Multi-Doc / Orden Físico)')
st.markdown('Detecta automáticamente cuántos documentos contiene el PDF y extrae '
            '**todos los registros en orden físico exacto**, sin reordenar.')

uploaded_file=st.file_uploader('Selecciona el archivo PDF',type=['pdf'])

if uploaded_file:
    pdf_bytes=uploaded_file.read()
    with st.spinner('Procesando páginas en orden…'):
        all_records,doc_summaries=extract_all(pdf_bytes)

    if not all_records:
        st.error('❌ No se extrajeron registros.')
    else:
        df=pd.DataFrame(all_records)
        st.subheader(f'📋 {len(doc_summaries)} documento(s) detectado(s)')
        cols_ui=st.columns(min(len(doc_summaries),4))
        for i,dm in enumerate(doc_summaries):
            df_d=df[df['Doc Seq']==dm['doc_seq']]
            calc=df_d['Value'].apply(lambda v: v if isinstance(v,(int,float)) else 0).sum()
            icon='🗂️' if dm['doc_type']=='INV' else '📋'
            with cols_ui[i%len(cols_ui)]:
                st.metric(f"{icon} Doc {dm['doc_seq']} — {dm['date']}",
                          f"{len(df_d)} registros", delta=f"${calc:,.2f}")
                if dm['total_val']:
                    diff=calc-dm['total_val']
                    st.caption(f"Declarado: ${dm['total_val']:,.2f} — "
                               f"{'✅ OK' if abs(diff)<1 else f'⚠️ Dif ${diff:+,.2f}'}")
                st.caption(dm['doc_type'])

        with st.expander('👁️ Preview — Todos los registros en orden físico',expanded=True):
            show=['Transfer Date','Doc Seq','Page','Line','B/P','Item',
                  'Description','Lot','Origin','Quantity','UOM','QTY(M)','Value','Doc Type']
            st.dataframe(df[[c for c in show if c in df.columns]],
                         use_container_width=True,height=400)

        with st.spinner('Generando Excel…'):
            xlsx=build_excel(all_records,doc_summaries)
        st.success(f'✅ {len(all_records)} registros  |  {len(doc_summaries)} documento(s)')
        st.download_button('⬇️ Descargar Excel',data=xlsx,
            file_name=os.path.splitext(uploaded_file.name)[0]+'_v4.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
else:
    st.info('👆 Sube un archivo PDF para comenzar.')
