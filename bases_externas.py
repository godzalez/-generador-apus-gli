"""
bases_externas.py — Motor de búsqueda en bases de precios externas.
Soporta: Gobernación de Boyacá, INVIAS (y cualquier base con estructura similar).
"""
import unicodedata, re
from openpyxl import load_workbook


# ══════════════════════════════════════════════════════════════════
# NORMALIZACIÓN
# ══════════════════════════════════════════════════════════════════

def _norm(texto):
    s = str(texto).upper().strip()
    s = ''.join(c for c in unicodedata.normalize('NFD', s)
                if unicodedata.category(c) != 'Mn')
    return re.sub(r'\s+', ' ', s)

_STOPWORDS = {
    'DE','LA','EL','EN','CON','Y','A','E','O','UN','UNA','LOS','LAS',
    'DEL','AL','POR','SU','ES','SE','QUE','PARA','HASTA','DESDE','SIN',
    'NO','SI','NI','MAS','MAS','TIPO','SEGUN','INCLUYE','INCLUYE',
}


# ══════════════════════════════════════════════════════════════════
# DETECTORES DE FORMATO
# ══════════════════════════════════════════════════════════════════

def _detectar_formato(ws):
    """
    Detecta el formato de la base de precios.
    Retorna 'boyaca' | 'invias' | 'generico' | None
    """
    primera = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    cols = [str(v).upper().strip() if v else '' for v in primera]
    cols_str = ' '.join(cols)

    # Formato Gobernación de Boyacá
    if ('CODIGO ITEM' in cols_str and 'UNIDAD' in cols_str and 'TOTAL' in cols_str
            and 'CAPITULO' in cols_str):
        return 'boyaca'

    # Formato INVIAS (a implementar cuando se cargue)
    if 'INVIAS' in cols_str or ('CODIGO' in cols_str and 'PRECIO' in cols_str):
        return 'invias'

    # Formato genérico: código, descripción, unidad, precio
    if any(x in cols_str for x in ('CODIGO', 'ITEM', 'DESCRIPCION')) and \
       any(x in cols_str for x in ('PRECIO', 'VALOR', 'TOTAL', 'UNITARIO')):
        return 'generico'

    return None


# ══════════════════════════════════════════════════════════════════
# CARGADORES POR FORMATO
# ══════════════════════════════════════════════════════════════════

def _cargar_boyaca(ws):
    """
    Estructura Gobernación de Boyacá:
    col0=cod_cap, col1=capitulo, col2=cod_sub, col3=subcapitulo,
    col4=codigo_item, col5=descripcion, col6=unidad,
    col7=vr_mano_obra, col8=total
    """
    bd = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row[4] or not row[8]: continue
        total = float(row[8]) if isinstance(row[8], (int, float)) else 0
        if total <= 0: continue
        mo = float(row[7]) if isinstance(row[7], (int, float)) and row[7] else 0
        codigo = str(row[4]).strip()
        desc   = str(row[5]).strip() if row[5] else ''
        bd[codigo] = {
            'codigo':      codigo,
            'fuente':      'Gobernación de Boyacá',
            'capitulo':    str(row[1]).strip() if row[1] else '',
            'subcapitulo': str(row[3]).strip() if row[3] else '',
            'descripcion': desc,
            'unidad':      str(row[6]).strip() if row[6] else '',
            'vr_mo':       mo,
            'vr_mat_equip': total - mo,
            'total':       total,
            'desc_norm':   _norm(desc),
        }
    return bd


def _cargar_generico(ws):
    """
    Intenta cargar cualquier base con encabezados detectados automáticamente.
    """
    primera = list(ws.iter_rows(min_row=1, max_row=1, values_only=True))[0]
    cols = [str(v).upper().strip() if v else '' for v in primera]

    # Detectar columnas
    col_cod  = next((j for j, c in enumerate(cols) if 'CODIGO' in c or c == 'ITEM'), None)
    col_desc = next((j for j, c in enumerate(cols) if 'DESCRIPCION' in c or 'DESCRIPCI' in c), None)
    col_und  = next((j for j, c in enumerate(cols) if c in ('UNIDAD','UND','UN','UND.')), None)
    col_tot  = next((j for j, c in enumerate(cols)
                     if any(x in c for x in ('TOTAL','PRECIO UNIT','VALOR UNIT'))), None)

    if col_desc is None or col_tot is None:
        return {}

    bd = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        total = row[col_tot] if col_tot < len(row) else None
        if not isinstance(total, (int, float)) or total <= 0: continue
        desc = str(row[col_desc]).strip() if col_desc < len(row) and row[col_desc] else ''
        if not desc: continue
        codigo = str(row[col_cod]).strip() if col_cod is not None and col_cod < len(row) and row[col_cod] else f'EXT-{len(bd)+1}'
        und = str(row[col_und]).strip() if col_und is not None and col_und < len(row) and row[col_und] else ''
        bd[codigo] = {
            'codigo': codigo, 'fuente': 'Base externa',
            'capitulo': '', 'subcapitulo': '', 'descripcion': desc,
            'unidad': und, 'vr_mo': 0, 'vr_mat_equip': float(total),
            'total': float(total), 'desc_norm': _norm(desc),
        }
    return bd


# ══════════════════════════════════════════════════════════════════
# FUNCIÓN PRINCIPAL DE CARGA
# ══════════════════════════════════════════════════════════════════

def cargar_base_externa(archivo, nombre_fuente=None):
    """
    Carga un archivo de base de precios externa.
    Retorna (bd_dict, formato_detectado, error).
    bd_dict: {codigo → ítem}
    """
    try:
        wb = load_workbook(archivo, data_only=True)
    except Exception as e:
        return {}, None, f"No se pudo abrir el archivo: {e}"

    # Intentar cada hoja hasta encontrar una con datos
    for nombre in wb.sheetnames:
        ws = wb[nombre]
        if ws.max_row < 5:
            continue
        fmt = _detectar_formato(ws)
        if fmt == 'boyaca':
            bd = _cargar_boyaca(ws)
            fuente = nombre_fuente or 'Gobernación de Boyacá'
            for v in bd.values():
                v['fuente'] = fuente
            return bd, 'boyaca', None
        elif fmt in ('invias', 'generico'):
            bd = _cargar_generico(ws)
            fuente = nombre_fuente or 'Base externa'
            for v in bd.values():
                v['fuente'] = fuente
            return bd, fmt, None

    return {}, None, "No se reconoció el formato del archivo de base de precios."


# ══════════════════════════════════════════════════════════════════
# BÚSQUEDA POR SIMILITUD
# ══════════════════════════════════════════════════════════════════

def buscar_en_base(descripcion, unidad, bd, top_n=5, umbral_pct=30):
    """
    Busca los ítems más similares en una base de precios externa.
    Retorna lista de (score, pct_coincidencia, item) ordenada por score desc.
    """
    palabras = {p for p in _norm(descripcion).split()
                if len(p) >= 4 and p not in _STOPWORDS}
    if not palabras:
        return []

    und_norm = _norm(unidad) if unidad else ''

    resultados = []
    for item in bd.values():
        palabras_item = set(item['desc_norm'].split())
        coincidencias = len(palabras & palabras_item)
        if coincidencias == 0:
            continue
        pct = coincidencias / len(palabras) * 100
        if pct < umbral_pct:
            continue
        bonus = 0.5 if und_norm and und_norm == _norm(item['unidad']) else 0
        score = coincidencias + bonus
        resultados.append((score, pct, item))

    resultados.sort(key=lambda x: -x[0])
    return resultados[:top_n]


def construir_apu_desde_base(item_bd, valor_ofrecido):
    """
    Construye un APU simplificado a partir de un ítem de base de precios.
    Escala los componentes al valor ofrecido.
    Retorna dict compatible con el generator.
    """
    total_ref = item_bd['total']
    factor = valor_ofrecido / total_ref if total_ref > 0 else 1.0

    mo_ref       = item_bd['vr_mo']
    mat_eq_ref   = item_bd['vr_mat_equip']

    mo_escalado     = round(mo_ref * factor, 2)
    mat_eq_escalado = round(mat_eq_ref * factor, 2)

    # Ajuste para que cierre exacto
    diferencia = valor_ofrecido - mo_escalado - mat_eq_escalado
    if abs(diferencia) > 0:
        if mo_escalado > 0:
            mo_escalado = round(mo_escalado + diferencia, 2)
        else:
            mat_eq_escalado = round(mat_eq_escalado + diferencia, 2)

    componentes_mdo = []
    if mo_escalado > 0:
        componentes_mdo = [{
            'description': 'Mano de obra (cuadrilla)',
            'unit': 'HR',
            'rend': 1.0,
            'unit_price': mo_escalado,
            'parcial': mo_escalado,
        }]

    componentes_mat = []
    if mat_eq_escalado > 0:
        componentes_mat = [{
            'description': 'Materiales, insumos y equipos',
            'unit': 'GL',
            'rend': 1.0,
            'unit_price': mat_eq_escalado,
            'parcial': mat_eq_escalado,
        }]

    return {
        'materiales':   componentes_mat,
        'herramientas': [],
        'transporte':   [],
        'mano_de_obra': componentes_mdo,
        'fuente_bd':    item_bd['fuente'],
        'codigo_bd':    item_bd['codigo'],
        'total_referencia': total_ref,
    }

"""
Generador Automático de APUs — Gerencia Legal Integral Colombia S.A.S.
v9: fix hoja proponente + generación IA con manejo robusto de errores.
"""
import streamlit as st
import pandas as pd
import json
import requests
from generator import leer_oferta_economica, leer_base_datos_apu, generate_apu_excel
from bases_externas import cargar_base_externa
from salarios import años_disponibles, jornal_minimo, smmlv_vigente, info_año

st.set_page_config(page_title="Generador de APUs – GLI", page_icon="🏗️", layout="wide")

st.markdown("""
    <div style='background-color:#1B3A6B;padding:18px 24px;border-radius:8px;margin-bottom:18px'>
        <h2 style='color:white;margin:0'>🏗️ Generador Automático de APUs</h2>
        <p style='color:#BDD7EE;margin:4px 0 0 0'>
        Gerencia Legal Integral Colombia S.A.S. &nbsp;·&nbsp;
        Desglosa cada precio unitario con sus componentes exactos.</p>
    </div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuración")

    # ── Año del proceso / SMMLV ───────────────────────────────────────────────
    st.subheader("📅 Año del proceso")
    año_opciones = años_disponibles()
    import datetime
    año_default_idx = año_opciones.index(datetime.date.today().year) \
        if datetime.date.today().year in año_opciones else 0
    año_manual = st.selectbox(
        "Año vigente del proceso",
        options=[None] + año_opciones,
        index=0,
        format_func=lambda x: "🔍 Auto-detectar" if x is None else str(x),
    )
    if año_manual:
        datos = info_año(año_manual)
        st.info(
            f"**SMMLV {año_manual}:** ${datos['smmlv']:,}\n\n"
            f"**Jornal mín. legal:** ${datos['jornal_minimo']:,.0f}/día\n\n"
            f"_(Base: SMMLV / {datos['dias_mes']} días × {datos['factor']}x prestacional)_"
        )
    else:
        st.caption("El sistema detectará el año del proceso automáticamente desde la lista de insumos o el archivo Excel.")

    st.divider()

    # ── AIU ───────────────────────────────────────────────────────────────────
    include_aiu = st.toggle("Incluir AIU en los APUs", value=False)
    aiu_pct = 0.0
    if include_aiu:
        val = st.number_input("Porcentaje AIU (%)", 0.0, 100.0, 25.0, 0.1, "%.2f")
        aiu_pct = val / 100.0
        st.info(f"AIU: **{val:.2f}%**")
    else:
        st.success("Sin AIU — precio = costo directo")

    st.divider()
    st.markdown("⚠️ Use descripciones **genéricas** (sin marcas comerciales).")
    st.caption("v10.0 · GLI Colombia · 2026")

# ── Paso 1: Archivo del proceso ───────────────────────────────────────────────
st.subheader("1. Cargue el Excel de la oferta económica del proceso")
st.caption(
    "Suba el Excel del proceso — puede ser cualquier formato: PRESUPUESTO, PROPUESTA, "
    "FORMULARIO, etc. Si el archivo tiene una hoja del proponente con sus precios, "
    "el software la usará automáticamente."
)
uploaded_proceso = st.file_uploader(
    "Excel del proceso (cualquier formato — PRESUPUESTO, PROPUESTA, FORMULARIO, etc.)",
    type=["xlsx","xlsm"], key="proceso"
)

# ── Paso 2: Base de datos de referencia ───────────────────────────────────────
st.subheader("2. Cargue su base de datos de referencia APU (opcional pero recomendado)")
st.caption(
    "Si tiene un archivo de un proceso anterior con hoja APU, cárguelo aquí. "
    "El software cruzará los códigos y tomará los componentes, ajustando los precios a su oferta actual."
)
uploaded_ref = st.file_uploader(
    "Archivo de referencia con hoja APU",
    type=["xlsx","xlsm"], key="referencia"
)

if not uploaded_proceso:
    st.info("👆 Cargue el Excel del proceso para comenzar.")
    st.stop()

# ── Leer referencia ───────────────────────────────────────────────────────────
bd_referencia = {}
if uploaded_ref:
    with st.spinner("Leyendo base de datos de referencia..."):
        bd_ref, err_ref = leer_base_datos_apu(uploaded_ref)
        if bd_ref:
            bd_referencia = bd_ref
            st.success(f"✅ Referencia cargada: **{len(bd_referencia)}** ítems con componentes.")
        else:
            st.warning(f"⚠️ Referencia: {err_ref}")

# ── Paso 3: Bases de precios externas ────────────────────────────────────────
st.subheader("3. Cargue bases de precios externas (Gobernación, INVIAS, etc.)")
st.caption(
    "Opcional. Si carga una base de precios (Gobernación de Boyacá, INVIAS u otra), "
    "el software la usará automáticamente para completar los ítems que no tengan APU de referencia. "
    "Puede cargar múltiples archivos."
)
uploaded_bases = st.file_uploader(
    "Bases de precios externas (Gobernación de Boyacá, INVIAS, etc.)",
    type=["xlsx","xlsm"], key="bases_ext", accept_multiple_files=True
)

bds_externas = []
if uploaded_bases:
    for ub in uploaded_bases:
        bd_ext, fmt_ext, err_ext = cargar_base_externa(ub, nombre_fuente=ub.name.replace('.xlsx','').replace('.xlsm',''))
        if bd_ext:
            bds_externas.append(bd_ext)
            st.success(f"✅ **{ub.name}**: {len(bd_ext)} ítems cargados (formato: {fmt_ext})")
        else:
            st.warning(f"⚠️ {ub.name}: {err_ext}")

# ── Leer proceso ──────────────────────────────────────────────────────────────
with st.spinner("Leyendo el archivo del proceso..."):
    resultado, error = leer_oferta_economica(
        uploaded_proceso,
        bd_referencia=bd_referencia,
        año_proceso=año_manual,
    )

if error:
    st.error(f"❌ {error}")
    st.stop()

con_apu = resultado['items_con_apu']
sin_apu = resultado['items_sin_apu']
n_total = resultado['total_proceso']
hoja    = resultado['hoja_usada']
año_proc       = resultado['año_proceso']
año_fuente     = resultado['año_fuente']
jornal_min     = resultado['jornal_min_legal']
alertas        = resultado['alertas']

# ── Año detectado y SMMLV ─────────────────────────────────────────────────────
st.info(f"📅 **Año del proceso detectado: {año_proc}** — {año_fuente}")

# ── Panel de validaciones críticas ───────────────────────────────────────────
alertas_supera   = alertas['supera_entidad']
alertas_negativa = alertas['zona_negativa']
alertas_rend     = alertas['rendimiento_alto']
alertas_jornal   = alertas['jornal_bajo']

if alertas_supera:
    with st.expander(
        f"🔴 **ALERTA CRÍTICA: {len(alertas_supera)} ítem(s) superan el APU de la entidad** — expandir para ver",
        expanded=True
    ):
        st.error(
            "Los siguientes ítems tienen un precio ofertado **MAYOR** al APU oficial de la entidad. "
            "Esto no es aceptable en una propuesta licitatoria. Revise su oferta económica antes de continuar."
        )
        data_sup = []
        for it in alertas_supera:
            data_sup.append({
                "Ítem": it['code'],
                "Descripción": it['description'][:60],
                "Precio ofertado": f"${it['valor_ofrecido']:,.0f}",
                "APU entidad": f"${it['apu_entidad_total']:,.0f}",
                "Exceso": f"${it['valor_ofrecido'] - it['apu_entidad_total']:,.0f}",
            })
        st.dataframe(data_sup, use_container_width=True)

if alertas_negativa:
    with st.expander(
        f"🔴 **{len(alertas_negativa)} ítem(s) en zona negativa** — costos fijos superan el precio ofertado",
        expanded=True
    ):
        st.warning(
            "En estos ítems, la suma de materiales + equipos + transporte (valores fijos del APU oficial) "
            "ya supera su precio ofertado. No hay margen para mano de obra. "
            "**Requieren justificación documental** (cotizaciones, fichas técnicas de eficiencia)."
        )
        data_neg = []
        for it in alertas_negativa:
            data_neg.append({
                "Ítem": it['code'],
                "Descripción": it['description'][:60],
                "Precio ofertado": f"${it['valor_ofrecido']:,.0f}",
                "Labor residual": f"${it.get('labor_residual', 0):,.0f}",
            })
        st.dataframe(data_neg, use_container_width=True)

if alertas_rend:
    with st.expander(
        f"⚠️ **{len(alertas_rend)} ítem(s) con rendimiento MO muy alto** (precio bajo para ese ítem)"
    ):
        st.warning(
            "El rendimiento calculado de mano de obra es mayor a 5.0x el rendimiento de referencia. "
            "Requieren acreditar rendimientos extraordinarios con registros de obra."
        )
        for it in alertas_rend:
            st.markdown(f"- **`{it['code']}`** — {it['description'][:55]} | Factor MO: **{it.get('factor_mo', '?')}x**")

if alertas_jornal:
    with st.expander(f"⚠️ **{len(alertas_jornal)} ítem(s)** con jornales ajustados al mínimo legal"):
        st.info(
            f"El APU de referencia traía jornales inferiores al mínimo legal {año_proc} "
            f"(${jornal_min:,.0f}/día). Se ajustaron automáticamente al mínimo."
        )

if not (alertas_supera or alertas_negativa or alertas_rend or alertas_jornal):
    if con_apu:
        st.success(f"✅ Todas las validaciones pasaron. {len(con_apu)} APU(s) listos con algoritmo R2.")

# ── Resumen ───────────────────────────────────────────────────────────────────
st.subheader("4. Resumen")
st.caption(f"Hoja leída: **{hoja}**")
c1, c2, c3 = st.columns(3)
c1.metric("Total ítems", n_total)
c2.metric("Con componentes APU", len(con_apu))
c3.metric("Sin componentes (requieren IA)", len(sin_apu))

if con_apu:
    with st.expander(f"✅ {len(con_apu)} ítems con componentes completos", expanded=False):
        for item in con_apu:
            n = sum(len(item[s]) for s in ('materiales','herramientas','transporte','mano_de_obra'))
            st.markdown(f"**`{item['code']}`** — {item['description'][:65]} | **${item['valor_ofrecido']:,.0f}** | {n} componentes")

# ── IA para ítems sin componentes ─────────────────────────────────────────────
items_ia = []

if sin_apu:
    st.subheader("5. Generación de componentes con IA")
    st.markdown(
        f"Los siguientes **{len(sin_apu)} ítems** no tienen componentes en ninguna fuente disponible. "
        "Haga clic en **Generar con IA** para que el asistente proponga la desagregación automáticamente."
    )

    if st.button("🤖 Generar componentes con IA para todos los ítems pendientes",
                 type="primary", use_container_width=True):

        progress = st.progress(0, text="Iniciando...")
        ok_count = 0
        err_count = 0
        items_generados = []
        jornal_min = resultado.get('jornal_min_legal', jornal_minimo(resultado.get('año_proceso', 2026)))
        año_proc   = resultado.get('año_proceso', 2026)

        for idx, item in enumerate(sin_apu):
            pct = idx / len(sin_apu)
            progress.progress(pct, text=f"Procesando [{item['code']}] ({idx+1}/{len(sin_apu)})...")

            precio_total = int(round(item['valor_ofrecido'], 0))

            prompt = f"""Eres un ingeniero civil colombiano experto en presupuestos de obra pública.
Desagrega el siguiente ítem en sus componentes de COSTO DIRECTO para una propuesta licitatoria colombiana.

ÍTEM: {item['description']}
UNIDAD DE MEDIDA: {item['unit']}
PRECIO UNITARIO OFERTADO: ${precio_total:,} COP
AÑO DEL PROCESO: {año_proc}
JORNAL MÍNIMO LEGAL {año_proc}: ${jornal_min:,.0f} COP/día (SMMLV/{jornal_min:.0f} × factor 1.5 prestacional)

REGLAS OBLIGATORIAS:
1. Descripciones GENÉRICAS sin marcas comerciales.
2. La suma exacta de (rend × unit_price) de TODOS los componentes debe ser exactamente {precio_total}.
3. Clasifica cada componente en: materiales, herramientas, mano_de_obra o transporte.
4. Unidades colombianas: M3, M2, ML, KG, GL, UND, HR, DIA, TON.
5. CRÍTICO — Mano de obra: unit_price NUNCA puede ser inferior a ${jornal_min:,.0f} COP/día.
   Este es el mínimo legal {año_proc}. Si el precio ofertado no permite MO a ese jornal,
   reduzca materiales o equipos pero NUNCA el jornal.
6. Rendimientos coherentes con la unidad del ítem ({item['unit']}).
7. Mínimo 2 componentes, máximo 8.
8. Sin valores negativos en ningún componente.

Responde ÚNICAMENTE con JSON válido sin texto adicional, sin comillas markdown:
{{"materiales":[{{"description":"nombre genérico","unit":"und","rend":0.5,"unit_price":10000}}],"herramientas":[{{"description":"nombre","unit":"und","rend":0.1,"unit_price":5000}}],"mano_de_obra":[{{"description":"cuadrilla operario+ayudante","unit":"DIA","rend":0.4,"unit_price":{jornal_min:.0f}}}],"transporte":[]}}"""

            try:
                # Obtener API key: primero de Streamlit Secrets (nube), luego de variable local
                import os
                api_key = None
                try:
                    api_key = st.secrets.get("ANTHROPIC_API_KEY", None)
                except Exception:
                    pass
                if not api_key:
                    api_key = os.environ.get("ANTHROPIC_API_KEY", "")

                headers_api = {"Content-Type": "application/json"}
                if api_key:
                    headers_api["x-api-key"] = api_key

                response = requests.post(
                    "https://api.anthropic.com/v1/messages",
                    headers=headers_api,
                    json={
                        "model": "claude-sonnet-4-20250514",
                        "max_tokens": 1000,
                        "messages": [{"role": "user", "content": prompt}]
                    },
                    timeout=45
                )

                if response.status_code == 401:
                    st.error(
                        "❌ **Error de autenticación con la API de IA.** "
                        "La generación automática de componentes requiere conexión a internet "
                        "y acceso a la API de Anthropic. "
                        "Verifique su conexión o ingrese los componentes manualmente en la sección siguiente."
                    )
                    break

                if response.status_code != 200:
                    err_count += 1
                    items_generados.append(item)
                    continue

                data = response.json()
                if 'error' in data:
                    err_count += 1
                    items_generados.append(item)
                    continue

                texto = data['content'][0]['text'].strip()

                # Limpiar backticks si los hay
                if '```' in texto:
                    partes = texto.split('```')
                    for p in partes:
                        p = p.strip()
                        if p.startswith('json'):
                            p = p[4:]
                        if p.startswith('{'):
                            texto = p
                            break

                componentes = json.loads(texto)

                # ── Validar jornales mínimos en componentes IA ────────────────
                for comp in componentes.get('mano_de_obra', []):
                    if comp.get('unit_price', 0) < jornal_min:
                        comp['unit_price'] = jornal_min  # corregir al mínimo legal

                # Ajustar para que sume exactamente el precio
                total_comp = sum(
                    c['rend'] * c['unit_price']
                    for sec in ('materiales','herramientas','mano_de_obra','transporte')
                    for c in componentes.get(sec, [])
                )
                diferencia = precio_total - total_comp

                if abs(diferencia) > 1:
                    # Ajustar el último componente de mano_de_obra o materiales
                    for sec in ('mano_de_obra','materiales','herramientas'):
                        comps = componentes.get(sec, [])
                        if comps and comps[-1]['rend'] > 0:
                            comps[-1]['unit_price'] = round(
                                comps[-1]['unit_price'] + diferencia / comps[-1]['rend'], 2
                            )
                            break

                item_generado = {
                    **item,
                    'materiales':   componentes.get('materiales', []),
                    'herramientas': componentes.get('herramientas', []),
                    'mano_de_obra': componentes.get('mano_de_obra', []),
                    'transporte':   componentes.get('transporte', []),
                    'tiene_apu':    True,
                }
                items_generados.append(item_generado)
                ok_count += 1

            except json.JSONDecodeError:
                # JSON inválido — guardar sin componentes
                items_generados.append(item)
                err_count += 1
            except requests.exceptions.Timeout:
                items_generados.append(item)
                err_count += 1
            except Exception:
                items_generados.append(item)
                err_count += 1

        progress.progress(1.0, text="✅ Completado")
        st.session_state.items_ia = items_generados

        if ok_count > 0:
            st.success(f"✅ IA generó componentes para **{ok_count}** ítems.")
        if err_count > 0:
            st.warning(
                f"⚠️ **{err_count} ítems** no pudieron generarse automáticamente. "
                "Puede ingresarlos manualmente en la sección siguiente o generar el Excel "
                "con las celdas vacías para completarlas después."
            )

    # Resultado IA generado
    if 'items_ia' in st.session_state and st.session_state.items_ia:
        con_comp = [i for i in st.session_state.items_ia if sum(len(i.get(s,[])) for s in ('materiales','herramientas','mano_de_obra')) > 0]
        sin_comp = [i for i in st.session_state.items_ia if sum(len(i.get(s,[])) for s in ('materiales','herramientas','mano_de_obra')) == 0]

        with st.expander(f"🤖 Componentes generados por IA: {len(con_comp)} ítems completos, {len(sin_comp)} vacíos", expanded=False):
            for item in con_comp[:10]:
                n = sum(len(item.get(s,[])) for s in ('materiales','herramientas','mano_de_obra'))
                precio = item['valor_ofrecido']
                suma = sum(c['rend']*c['unit_price'] for s in ('materiales','herramientas','mano_de_obra','transporte') for c in item.get(s,[]))
                cierra = "✅" if abs(precio-suma) < 2 else f"⚠️ diff=${abs(precio-suma):,.0f}"
                st.markdown(f"**`{item['code']}`** — {item['description'][:55]} | ${precio:,.0f} | {n} comp | {cierra}")

# ── Generación final ──────────────────────────────────────────────────────────
seccion_num = "6." if sin_apu else "5."
st.subheader(f"{seccion_num} Generar y descargar APUs")

items_manuales_final = st.session_state.get('items_ia', [])
items_sin_procesar = [
    i for i in sin_apu
    if i['code'] not in {x['code'] for x in items_manuales_final}
]
todos_manuales = items_manuales_final + items_sin_procesar
n_gen = len(con_apu) + len(todos_manuales)
aiu_txt = f"con AIU ({aiu_pct*100:.1f}%)" if include_aiu else "sin AIU"

if n_gen == 0:
    st.warning("No hay ítems para generar.")
else:
    if items_sin_procesar:
        st.info(
            f"ℹ️ {len(items_sin_procesar)} ítem(s) se generarán con estructura vacía "
            "(sin componentes). Puede completarlos manualmente en el Excel descargado."
        )

    if st.button(f"🚀 Generar {n_gen} APU(s) — {aiu_txt}",
                 type="primary", use_container_width=True):
        with st.spinner(f"Generando {n_gen} APUs..."):
            try:
                excel = generate_apu_excel(
                    resultado,
                    items_manuales=todos_manuales if todos_manuales else None,
                    include_aiu=include_aiu,
                    aiu_pct=aiu_pct,
                    bd_externas=bds_externas if bds_externas else None,
                )
                st.session_state.excel = excel
                st.session_state.nombre = uploaded_proceso.name.replace('.xlsx','').replace('.xlsm','')
                st.success(
                    f"✅ **{n_gen} APUs generados.** "
                    f"El PRECIO UNITARIO de cada APU cuadra exactamente con su oferta."
                )
            except Exception as e:
                st.error(f"❌ Error al generar: {e}")
                raise e

    if "excel" in st.session_state:
        st.download_button(
            label="⬇️ Descargar APUs en Excel",
            data=st.session_state.excel,
            file_name=f"APUs_{st.session_state.nombre}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary",
        )
