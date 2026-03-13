import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from collections import defaultdict
import io, math, zipfile, os

st.set_page_config(page_title="Template CM | INTERLOG", page_icon="📋", layout="wide")

# ── CSS estilo Portal Grupo 8 ─────────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Roboto', sans-serif;
        background-color: #0e1117;
        color: #fafafa;
    }
    .block-container { padding-top: 2rem; }

    .titulo-app {
        font-size: 2rem; font-weight: 700;
        color: #00b4d8; margin-bottom: 0;
    }
    .subtitulo-app {
        font-size: 0.95rem; color: #8899aa;
        margin-bottom: 1.5rem;
    }
    .seccion-titulo {
        font-size: 1rem; font-weight: 600;
        color: #00b4d8; text-transform: uppercase;
        letter-spacing: 1px; margin-bottom: 1rem;
        border-bottom: 1px solid #1e3a4a; padding-bottom: 6px;
    }
    .card {
        background: #161b22; border: 1px solid #1e3a4a;
        border-radius: 10px; padding: 20px; margin-bottom: 16px;
    }
    .badge-ok {
        background: #0d2e1a; color: #3dd68c;
        padding: 5px 12px; border-radius: 20px;
        font-size: 12px; font-weight: 500;
        border: 1px solid #3dd68c;
        display: inline-block; margin: 3px 2px;
    }
    .badge-err {
        background: #2e0d0d; color: #ff6b6b;
        padding: 5px 12px; border-radius: 20px;
        font-size: 12px; font-weight: 500;
        border: 1px solid #ff6b6b;
        display: inline-block; margin: 3px 2px;
    }
    .badge-fixed {
        background: #1a2a1a; color: #90ee90;
        padding: 5px 12px; border-radius: 20px;
        font-size: 12px; font-weight: 500;
        border: 1px solid #90ee90;
        display: inline-block; margin: 3px 2px;
    }
    .resumen-card {
        background: #161b22; border: 1px solid #1e3a4a;
        border-radius: 8px; padding: 14px; text-align: center;
        margin: 4px;
    }
    .resumen-inv { font-size: 0.85rem; color: #00b4d8; font-weight: 600; }
    .resumen-items { font-size: 1.4rem; font-weight: 700; color: #fafafa; }
    .resumen-fob { font-size: 0.8rem; color: #8899aa; }

    .stButton > button {
        background: linear-gradient(135deg, #00b4d8, #0077b6) !important;
        color: white !important; font-weight: 600 !important;
        border: none !important; border-radius: 8px !important;
        padding: 0.6rem 1.5rem !important;
        font-size: 1rem !important; letter-spacing: 0.5px;
        transition: opacity 0.2s;
    }
    .stButton > button:hover { opacity: 0.85; }
    .stButton > button:disabled { opacity: 0.4 !important; }

    div[data-testid="stFileUploader"] {
        background: #161b22 !important;
        border: 1px dashed #1e3a4a !important;
        border-radius: 8px !important;
        padding: 8px !important;
    }
    .stSelectbox > div > div {
        background: #161b22 !important;
        border: 1px solid #1e3a4a !important;
        border-radius: 8px !important;
    }
    .stTextInput > div > div > input {
        background: #161b22 !important;
        border: 1px solid #1e3a4a !important;
        border-radius: 8px !important;
        color: #fafafa !important;
    }
    .alerta-descarte {
        background: #2a1f00; border: 1px solid #f0c040;
        border-radius: 8px; padding: 12px 16px;
        color: #f0c040; font-size: 0.88rem; margin: 6px 0;
    }
</style>
""", unsafe_allow_html=True)

# ── CONSTANTES ────────────────────────────────────────────────────────────────
ARCHIVOS_FIJOS = {
    'razon': 'RAZON_SOCIAL_FSM.xlsx',
    'proyectos': 'PROYECTOS.xlsx',
    'ncm': 'NCM_LEY_MINERA.xlsx',
}

PAIS_MAP = {
    'USA': 'ESTADOS UNIDOS', 'UNITED STATES': 'ESTADOS UNIDOS',
    'MEXICO': 'MEXICO', 'INDIA': 'INDIA', 'CANADA': 'CANADA', 'CHINA': 'CHINA',
    'DENMARK': 'DINAMARCA', 'GERMANY': 'ALEMANIA', 'JAPAN': 'JAPON',
    'FRANCE': 'FRANCIA', 'ITALY': 'ITALIA', 'BRAZIL': 'BRASIL',
    'UK': 'REINO UNIDO', 'UNITED KINGDOM': 'REINO UNIDO', 'SWEDEN': 'SUECIA',
    'SPAIN': 'ESPAÑA', 'NETHERLANDS': 'PAISES BAJOS',
    'SOUTH KOREA': 'COREA DEL SUR', 'KOREA': 'COREA DEL SUR',
    'AUSTRALIA': 'AUSTRALIA', 'FINLAND': 'FINLANDIA',
}
UNIDAD_MAP = {
    '07 - UNIDAD': 'UNIDAD', '01 - KILOGRAMO': 'KILOGRAMO',
    '06 - METRO': 'METRO', '10 - LITRO': 'LITRO',
}
COLUMNAS = [
    'ID','InscRUMP','ActiServ','NroInsc','RazonSocial','CUIT','ImpDirecta','CondMerca',
    'AnioFabricacion','RazonSocialLeasing','CuitLeasing','SimiSira','ValorFOBTotal',
    'ProyectoMinero','Radicacion','DetalleTransitorioDeposito','ClasificacionDeArticulo',
    'TipoDeFactura','NumeroDeFactura','OrdenDeCompra','Descripcion','Cantidad',
    'UnidadMedida','PosicionArancelaria','Marca','Modelo','NroDeSerie','ValorUnitario',
    'ValorTotalItem','SerieDelMotor','MarcaDelMotor','ModeloMotor','TipoPlantaDestino',
    'FinalidadDeUso','CodigoParte','TipoDeMaquina','MarcaMaquina','ModeloMaquina',
    'Expedientes Escalonados','Proveedor','PaisOrigenDeLaMercaderia','Observaciones',
    'ITEM_DESPACHO'
]
SIN_COLOR = {
    'ID','InscRUMP','ActiServ','NroInsc','RazonSocial','CUIT','ImpDirecta','CondMerca',
    'SimiSira','ProyectoMinero','Radicacion','ClasificacionDeArticulo','TipoDeFactura',
    'Observaciones','ITEM_DESPACHO'
}

# ── HELPERS ───────────────────────────────────────────────────────────────────
def safe_float(v):
    try:
        f = float(v); return 0.0 if math.isnan(f) else f
    except: return 0.0

def traducir_pais(p):
    return PAIS_MAP.get(str(p).strip().upper(), str(p).strip().upper())

def parsear_equipo(eq):
    if not eq or eq.strip() in ('', 'nan'): return '', 'CAT', ''
    eq = eq.strip()
    if ' - ' in eq:
        p = eq.split(' - ', 1); return p[0].strip(), 'CAT', p[1].strip()
    tokens = eq.split()
    for i, t in enumerate(tokens):
        if t.upper() in ['CAT', 'CATERPILLAR']:
            return ' '.join(tokens[:i]).strip(), 'CAT', ' '.join(tokens[i+1:]).strip()
    return (' '.join(tokens[:-1]).strip(), 'CAT', tokens[-1].strip()) if len(tokens) > 1 else (eq, 'CAT', '')

def calcular_fobs(items):
    total_ext = sum(safe_float(i.get('EXTENDED_PRICE')) for i in items)
    total_gastos = sum(
        safe_float(i.get('SPECIAL PACKING')) + safe_float(i.get('FREIGHT_CHARGE')) +
        safe_float(i.get('BO_FREIGHT_CHARGE')) + safe_float(i.get('EMERGENCY_FILL_CHARGE_VAL'))
        for i in items
    )
    return [
        round(safe_float(i.get('EXTENDED_PRICE')) +
              (safe_float(i.get('EXTENDED_PRICE')) * total_gastos / total_ext if total_ext else 0), 2)
        for i in items
    ]

def match_fob(pn, qty, fob_calc, subitem_df, usados):
    pn_str = str(pn).strip()
    try: qty_val = float(qty)
    except: return None, None, 'sin_match'
    mask = (subitem_df['MODELO'].astype(str).str.strip() == pn_str) & \
           (subitem_df['CANTIDAD'].astype(float) == qty_val)
    disponibles = subitem_df[mask & ~subitem_df.index.isin(usados)].copy()
    if disponibles.empty: return None, None, 'sin_match'
    match = disponibles[(disponibles['MONTO FOB'].astype(float) - fob_calc).abs() <= 0.02]
    if not match.empty:
        idx = match.index[0]; usados.add(idx)
        return match.loc[idx], str(match.loc[idx]['ITEM']), 'unico'
    idx = disponibles.index[0]; usados.add(idx)
    return disponibles.loc[idx], str(disponibles.loc[idx]['ITEM']), 'descarte'

def generar_excel_bytes(filas):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'general'
    FILL_AM = PatternFill('solid', fgColor='FFFF00')
    FH = Font(name='Arial', size=11, bold=True)
    FD = Font(name='Calibri', size=11)
    for col_idx, col_name in enumerate(COLUMNAS, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name)
        cell.font = FH; cell.alignment = Alignment(horizontal='left', vertical='center')
        if col_name not in SIN_COLOR: cell.fill = FILL_AM
    for row_idx, fila in enumerate(filas, 2):
        for col_idx, col_name in enumerate(COLUMNAS, 1):
            val = fila.get(col_name)
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = FD; cell.alignment = Alignment(vertical='center')
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value is not None else 0) for c in col)
        ws.column_dimensions[col[0].column_letter].width = min(max(max_len + 2, 10), 45)
    ws.freeze_panes = 'A2'
    buf = io.BytesIO(); wb.save(buf); buf.seek(0); return buf.getvalue()

def procesar(f_madre, f_despacho, f_equipos, f_desc, cond_merca):
    # Archivos dinámicos
    df_madre = pd.read_excel(f_madre, sheet_name='Hoja2', dtype=str)
    df_subitem = pd.read_excel(f_despacho, sheet_name='Subitem', dtype=str)

    xl_eq = pd.ExcelFile(f_equipos)
    sheet_eq = 'data' if 'data' in xl_eq.sheet_names else xl_eq.sheet_names[0]
    df_eq = pd.read_excel(f_equipos, sheet_name=sheet_eq, dtype=str)

    xl_desc = pd.ExcelFile(f_desc)
    sheet_desc = 'Hoja2' if 'Hoja2' in xl_desc.sheet_names else xl_desc.sheet_names[0]
    df_desc_df = pd.read_excel(f_desc, sheet_name=sheet_desc, dtype=str)

    # Archivos fijos del repo
    df_razon = pd.read_excel(ARCHIVOS_FIJOS['razon'], dtype=str)
    df_ncm_raw = pd.read_excel(ARCHIVOS_FIJOS['ncm'], dtype=str)

    xl_proy = pd.ExcelFile(ARCHIVOS_FIJOS['proyectos'])
    sheet_proy = 'Hoja2' if 'Hoja2' in xl_proy.sheet_names else xl_proy.sheet_names[0]
    df_proy = pd.read_excel(ARCHIVOS_FIJOS['proyectos'], sheet_name=sheet_proy, dtype=str)

    for df in [df_madre, df_subitem, df_razon, df_proy, df_eq, df_desc_df, df_ncm_raw]:
        df.columns = df.columns.str.strip()
    for col in df_proy.columns:
        df_proy[col] = df_proy[col].astype(str).str.replace('\xa0', ' ', regex=False).str.strip()

    ncm_validas = set(df_ncm_raw['NCM'].str.strip().tolist())
    rs = df_razon.iloc[0]
    proyectos = {str(r['CUST_CD']).strip(): r.to_dict() for _, r in df_proy.iterrows()}

    eq_col_pn = 'Part number' if 'Part number' in df_eq.columns else df_eq.columns[0]
    eq_col_eq = 'Equipos' if 'Equipos' in df_eq.columns else df_eq.columns[-1]
    equipos = {str(r[eq_col_pn]).strip(): (str(r[eq_col_eq]).strip() if pd.notna(r[eq_col_eq]) else '') for _, r in df_eq.iterrows()}

    desc_col_pn = 'PART_NUMBER' if 'PART_NUMBER' in df_desc_df.columns else df_desc_df.columns[0]
    desc_col_d = 'DESCRIPCION' if 'DESCRIPCION' in df_desc_df.columns else df_desc_df.columns[-1]
    descs = {str(r[desc_col_pn]).strip(): (str(r[desc_col_d]).strip() if pd.notna(r[desc_col_d]) else '') for _, r in df_desc_df.iterrows()}

    try: cuit_val = int(float(str(rs['CUIT'])))
    except: cuit_val = str(rs['CUIT'])

    facturas = defaultdict(list)
    for _, row in df_madre.iterrows():
        inv = str(row.get('INVOICE_NUMBER', '')).strip()
        if inv and inv != 'nan':
            facturas[inv].append(row.to_dict())

    resultados = {}
    alertas = []
    usados_global = set()

    for inv, items in facturas.items():
        fobs_calc = calcular_fobs(items)
        filas_candidatas = []

        for item, fob_calc in zip(items, fobs_calc):
            pn = str(item.get('PART_NUMBER', '')).strip()
            qty_str = str(item.get('QTY', '')).strip()
            cust_cd = str(item.get('CUST_CD', '')).strip()

            sub_row, item_despacho, estado = match_fob(pn, qty_str, fob_calc, df_subitem, usados_global)

            fob_final = ncm = valor_unit = valor_total = None
            unidad = 'UNIDAD'; ncm_10 = ''

            if sub_row is not None:
                fob_final = round(float(sub_row['MONTO FOB']), 2)
                ncm = str(sub_row.get('NCM', '')).strip()
                ncm_10 = ncm[:10]
                try:
                    qty_f = float(qty_str)
                    valor_unit = round(fob_final / qty_f, 2) if qty_f else fob_final
                except: valor_unit = fob_final
                valor_total = fob_final
                unidad = UNIDAD_MAP.get(str(sub_row.get('UNIDAD DECLARADA', '')).strip(), 'UNIDAD')
                if estado == 'descarte':
                    alertas.append({
                        'factura': inv, 'pn': pn, 'fob_calc': fob_calc,
                        'fob_despacho': fob_final, 'diff': round(fob_final - fob_calc, 2),
                        'item_despacho': item_despacho
                    })

            if ncm_10 not in ncm_validas:
                continue

            tipo_maq, marca_maq, modelo_maq = parsear_equipo(equipos.get(pn, ''))
            descripcion = descs.get(pn, str(item.get('PART_NAME', '')).strip())
            proy = proyectos.get(cust_cd, {})

            try: qty_int = int(float(qty_str)) if qty_str and qty_str != 'nan' else None
            except: qty_int = qty_str

            filas_candidatas.append({
                'pn': pn, 'qty_int': qty_int, 'fob_final': fob_final, 'ncm': ncm,
                'valor_unit': valor_unit, 'valor_total': valor_total, 'unidad': unidad,
                'tipo_maq': tipo_maq, 'marca_maq': marca_maq, 'modelo_maq': modelo_maq,
                'descripcion': descripcion, 'proy': proy,
                'origen': traducir_pais(str(item.get('PART_ORIGIN', '')).strip()),
                'item_despacho': item_despacho,
            })

        if not filas_candidatas:
            continue

        valor_fob_total = round(sum(f['fob_final'] for f in filas_candidatas if f['fob_final']), 2)
        filas = []
        for f in filas_candidatas:
            proy = f['proy']
            filas.append({
                'ID': 0,
                'InscRUMP': str(rs['InscRUMP']), 'ActiServ': str(rs['ActiServ']),
                'NroInsc': str(rs['NroInsc']), 'RazonSocial': str(rs['RazonSocial']),
                'CUIT': cuit_val, 'ImpDirecta': str(rs['ImpDirecta']),
                'CondMerca': cond_merca,
                'AnioFabricacion': None, 'RazonSocialLeasing': None, 'CuitLeasing': None, 'SimiSira': None,
                'ValorFOBTotal': valor_fob_total,
                'ProyectoMinero': str(proy.get('ProyectoMinero', '')).upper() if proy else '',
                'Radicacion': str(proy.get('Radicacion', '')).upper() if proy else '',
                'DetalleTransitorioDeposito': str(proy.get('DetalleTransitorioDeposito', '')) if proy else '',
                'ClasificacionDeArticulo': str(rs['ClasificacionDeArticulo']),
                'TipoDeFactura': 'DEFINITIVA', 'NumeroDeFactura': inv, 'OrdenDeCompra': None,
                'Descripcion': f['descripcion'], 'Cantidad': f['qty_int'], 'UnidadMedida': f['unidad'],
                'PosicionArancelaria': f['ncm'], 'Marca': 'CATERPILLAR', 'Modelo': 'NO POSEE',
                'NroDeSerie': 'NO POSEE', 'ValorUnitario': f['valor_unit'], 'ValorTotalItem': f['valor_total'],
                'SerieDelMotor': None, 'MarcaDelMotor': None, 'ModeloMotor': None,
                'TipoPlantaDestino': None, 'FinalidadDeUso': None,
                'CodigoParte': f['pn'], 'TipoDeMaquina': f['tipo_maq'],
                'MarcaMaquina': f['marca_maq'], 'ModeloMaquina': f['modelo_maq'],
                'Expedientes Escalonados': 'no', 'Proveedor': 'CATERPILLAR',
                'PaisOrigenDeLaMercaderia': f['origen'], 'Observaciones': None,
                'ITEM_DESPACHO': f['item_despacho'],
            })
        resultados[inv] = filas

    return resultados, alertas


# ── UI ────────────────────────────────────────────────────────────────────────
st.markdown('<p class="titulo-app">📋 Template CM</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitulo-app">INTERLOG Comercio Exterior — Generación automática de planillas de clasificación de mercadería</p>', unsafe_allow_html=True)
st.markdown("---")

col_main, col_config = st.columns([3, 1])

with col_main:
    st.markdown('<p class="seccion-titulo">📂 Archivos de la operación</p>', unsafe_allow_html=True)

    nro_ref = st.text_input("Número de referencia de la operación", placeholder="ej: 982755")

    c1, c2 = st.columns(2)
    with c1:
        label_facaero = f"Excel FACAERO {nro_ref}" if nro_ref else "Excel FACAERO (ingresá el número de referencia)"
        f_madre = st.file_uploader(label_facaero, type=["xlsx"])

        label_eq = f"Excel Equipos {nro_ref}" if nro_ref else "Excel Equipos"
        f_equipos = st.file_uploader(label_eq, type=["xlsx"])

    with c2:
        label_di = f"Excel DI {nro_ref}" if nro_ref else "Excel DI (ingresá el número de referencia)"
        f_despacho = st.file_uploader(label_di, type=["xlsx"])

        label_desc = f"Excel Descripciones {nro_ref}" if nro_ref else "Excel Descripciones"
        f_descripciones = st.file_uploader(label_desc, type=["xlsx"])

with col_config:
    st.markdown('<p class="seccion-titulo">⚙️ Configuración</p>', unsafe_allow_html=True)

    cond_merca = st.selectbox("Condición de Mercadería", ["NUEVA", "USADA"])

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<p class="seccion-titulo">📋 Estado</p>', unsafe_allow_html=True)

    archivos_din = [f_madre, f_despacho, f_equipos, f_descripciones]
    nombres_din = [
        f"FACAERO {nro_ref or ''}",
        f"DI {nro_ref or ''}",
        "Equipos",
        "Descripciones"
    ]
    for f, n in zip(archivos_din, nombres_din):
        if f:
            st.markdown(f'<span class="badge-ok">✓ {n}</span>', unsafe_allow_html=True)
        else:
            st.markdown(f'<span class="badge-err">✗ {n}</span>', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown("**Archivos fijos del sistema:**")
    for nombre in ["Razón Social", "Proyectos", "NCM Ley Minera"]:
        st.markdown(f'<span class="badge-fixed">⚙ {nombre}</span>', unsafe_allow_html=True)

st.markdown("---")

todos_cargados = all(archivos_din) and bool(nro_ref)

if not nro_ref:
    st.info("📝 Ingresá el número de referencia de la operación para continuar.")

if st.button("🚀 GENERAR TEMPLATES", disabled=not todos_cargados, use_container_width=True):
    with st.spinner("Procesando operación..."):
        try:
            resultados, alertas = procesar(f_madre, f_despacho, f_equipos, f_descripciones, cond_merca)
            st.session_state['resultados'] = resultados
            st.session_state['alertas'] = alertas
            st.session_state['procesado'] = True
            st.session_state['nro_ref'] = nro_ref
        except Exception as e:
            st.error(f"Error al procesar: {e}")
            import traceback; st.code(traceback.format_exc())

if st.session_state.get('procesado'):
    resultados = st.session_state['resultados']
    alertas = st.session_state['alertas']
    nro = st.session_state.get('nro_ref', '')

    if alertas:
        st.markdown("### ⚠️ Alertas — Verificar asignaciones por descarte")
        for a in alertas:
            st.markdown(f"""
            <div class="alerta-descarte">
                ⚠️ <b>{a['factura']}</b> | PN <code>{a['pn']}</code> → ITEM <code>{a['item_despacho']}</code> asignado por descarte |
                FOB calculado: {a['fob_calc']} | FOB despacho: {a['fob_despacho']} | <b>Diferencia: {a['diff']}</b>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("### 📊 Resumen de la operación")
    total_items = sum(len(v) for v in resultados.values())

    cols_res = st.columns(min(len(resultados), 4))
    for i, (inv, filas) in enumerate(resultados.items()):
        with cols_res[i % 4]:
            st.markdown(f"""
            <div class="resumen-card">
                <div class="resumen-inv">{inv}</div>
                <div class="resumen-items">{len(filas)}</div>
                <div class="resumen-fob">ítems | FOB: USD {filas[0]['ValorFOBTotal'] if filas else '-'}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown(f"**Total: {len(resultados)} facturas — {total_items} ítems procesados**")
    st.markdown("---")

    st.markdown("### 📥 Descargar templates")
    excel_bytes = {}
    cols_dl = st.columns(min(len(resultados), 4))
    for i, (inv, filas) in enumerate(resultados.items()):
        b = generar_excel_bytes(filas)
        excel_bytes[inv] = b
        with cols_dl[i % 4]:
            st.download_button(
                label=f"📄 {inv}",
                data=b,
                file_name=f"TEMPLATE_{nro}_{inv}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"dl_{inv}"
            )

    st.markdown("<br>", unsafe_allow_html=True)
    zip_buf = io.BytesIO()
    with zipfile.ZipFile(zip_buf, 'w') as zf:
        for inv, b in excel_bytes.items():
            zf.writestr(f"TEMPLATE_{nro}_{inv}.xlsx", b)
    zip_buf.seek(0)

    st.download_button(
        label=f"📦 DESCARGAR TODOS — Operación {nro}",
        data=zip_buf.getvalue(),
        file_name=f"TEMPLATES_{nro}.zip",
        mime="application/zip",
        use_container_width=True
    )
