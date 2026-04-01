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
    st.caption("v9.0 · GLI Colombia · 2026")

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
    resultado, error = leer_oferta_economica(uploaded_proceso, bd_referencia=bd_referencia)

if error:
    st.error(f"❌ {error}")
    st.stop()

con_apu = resultado['items_con_apu']
sin_apu = resultado['items_sin_apu']
n_total = resultado['total_proceso']
hoja    = resultado['hoja_usada']

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

        for idx, item in enumerate(sin_apu):
            pct = idx / len(sin_apu)
            progress.progress(pct, text=f"Procesando [{item['code']}] ({idx+1}/{len(sin_apu)})...")

            precio_total = int(round(item['valor_ofrecido'], 0))

            prompt = f"""Eres un ingeniero civil colombiano experto en presupuestos de obra pública.
Desagrega el siguiente ítem en sus componentes de COSTO DIRECTO.

ÍTEM: {item['description']}
UNIDAD DE MEDIDA DEL ÍTEM: {item['unit']}
PRECIO UNITARIO: ${precio_total:,} COP

REGLAS OBLIGATORIAS:
1. Descripciones GENÉRICAS sin marcas comerciales.
2. La suma exacta de (rend × unit_price) de TODOS los componentes debe ser {precio_total}.
3. Clasifica cada componente en: materiales, herramientas, mano_de_obra o transporte.
4. UNIDADES colombianas para cada componente (campo "unit"): M3, M2, ML, KG, GL, UND, HR, DIA, M3-KM, TON, VJE. NUNCA dejes "unit" vacío.
5. unit_price debe ser un número ENTERO (sin decimales).
6. Rendimientos coherentes con la unidad del ítem ({item['unit']}).
7. Mínimo 2 componentes, máximo 8.
8. La MANO DE OBRA debe representar al menos el 15% del total (no puede quedar en cero ni negativa).

Responde ÚNICAMENTE con JSON válido sin texto adicional, sin comillas markdown:
{{"materiales":[{{"description":"nombre genérico","unit":"UND","rend":0.5,"unit_price":10000}}],"herramientas":[{{"description":"nombre","unit":"HR","rend":0.1,"unit_price":5000}}],"mano_de_obra":[{{"description":"cuadrilla nombre","unit":"DIA","rend":0.4,"unit_price":45000}}],"transporte":[]}}"""

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
        mo_baja  = [i for i in con_comp if i.get('alerta_mo', False)]

        if mo_baja:
            st.warning(
                f"⚠️ **{len(mo_baja)} ítem(s) tienen mano de obra inferior al 5% del costo directo.** "
                "Revíselos en la hoja RESUMEN (celdas naranjas) y ajuste la tabla 2 antes de entregar."
            )

        with st.expander(f"🤖 Componentes generados por IA: {len(con_comp)} ítems completos, {len(sin_comp)} vacíos", expanded=False):
            for item in con_comp[:10]:
                n = sum(len(item.get(s,[])) for s in ('materiales','herramientas','mano_de_obra'))
                precio = item['valor_ofrecido']
                suma = sum(c['rend']*c['unit_price'] for s in ('materiales','herramientas','mano_de_obra','transporte') for c in item.get(s,[]))
                cierra = "✅" if abs(precio-suma) < 2 else f"⚠️ diff=${abs(precio-suma):,.0f}"
                mo_tag = f" | ⚠️ MO {item['pct_mo']}%" if item.get('alerta_mo') else ""
                st.markdown(f"**`{item['code']}`** — {item['description'][:55]} | ${precio:,.0f} | {n} comp | {cierra}{mo_tag}")

# ══════════════════════════════════════════════════════════════════
# PANTALLA DE REVISIÓN ANTES DE EXPORTAR
# ══════════════════════════════════════════════════════════════════
seccion_num = "6." if sin_apu else "5."
st.subheader(f"{seccion_num} Revisión y aprobación de APUs")
st.caption(
    "Revise cada ítem antes de exportar. Puede aprobar o rechazar individualmente. "
    "Solo los ítems **aprobados** se incluirán en el Excel final."
)

# ── Construir lista unificada de todos los ítems ─────────────────
items_manuales_final = st.session_state.get('items_ia', [])
items_sin_procesar   = [
    i for i in sin_apu
    if i['code'] not in {x['code'] for x in items_manuales_final}
]
todos_manuales = items_manuales_final + items_sin_procesar

# Lista maestra de todos los ítems (con_apu + manuales/IA + sin componentes)
lista_revision = list(con_apu) + todos_manuales

# Marcar fuente de cada ítem
for item in lista_revision:
    if 'fuente_revision' not in item:
        fuente = item.get('fuente_bd', '')
        n_comp = sum(len(item.get(s,[])) for s in ('materiales','herramientas','transporte','mano_de_obra'))
        if n_comp == 0:
            item['fuente_revision'] = 'Sin componentes'
        elif fuente:
            item['fuente_revision'] = f'Base externa ({fuente[:30]})'
        elif item in con_apu:
            item['fuente_revision'] = 'Referencia propia'
        else:
            item['fuente_revision'] = 'IA generado'

if not lista_revision:
    st.warning("No hay ítems para revisar.")
    st.stop()

# ── Inicializar estado de aprobación ─────────────────────────────
if 'aprobacion' not in st.session_state:
    st.session_state.aprobacion = {}

# Pre-aprobar automáticamente los que tienen referencia propia y cuadran
for item in lista_revision:
    code = item['code']
    if code not in st.session_state.aprobacion:
        n_comp = sum(len(item.get(s,[])) for s in
                     ('materiales','herramientas','transporte','mano_de_obra'))
        precio = item.get('valor_ofrecido', 0)
        suma   = sum(c['rend']*c['unit_price']
                     for s in ('materiales','herramientas','transporte','mano_de_obra')
                     for c in item.get(s, []))
        cierra = abs(precio - suma) < 2 if precio > 0 else False
        alerta = item.get('alerta_mo', False)
        # Auto-aprobar: referencia propia + cierra + sin alerta MO
        auto_ok = (item.get('fuente_revision') == 'Referencia propia'
                   and cierra and not alerta and n_comp > 0)
        st.session_state.aprobacion[code] = auto_ok

# ── Controles globales ────────────────────────────────────────────
col_ap, col_rec, col_info = st.columns([1, 1, 3])
with col_ap:
    if st.button("✅ Aprobar todos", use_container_width=True):
        for item in lista_revision:
            st.session_state.aprobacion[item['code']] = True
        st.rerun()
with col_rec:
    if st.button("❌ Rechazar todos", use_container_width=True):
        for item in lista_revision:
            st.session_state.aprobacion[item['code']] = False
        st.rerun()
with col_info:
    n_aprobados  = sum(1 for v in st.session_state.aprobacion.values() if v)
    n_rechazados = len(lista_revision) - n_aprobados
    st.info(f"**{n_aprobados}** aprobados · **{n_rechazados}** pendientes de {len(lista_revision)} total")

st.divider()

# ── Tabla de revisión ítem por ítem ──────────────────────────────
SECCIONES = ('materiales','herramientas','transporte','mano_de_obra')

# Agrupar por estado para mostrar pendientes primero
pendientes = [i for i in lista_revision if not st.session_state.aprobacion.get(i['code'], False)]
aprobados  = [i for i in lista_revision if st.session_state.aprobacion.get(i['code'], False)]

for grupo_label, grupo_items, expandido in [
    ("⚠️ Pendientes de revisión", pendientes, True),
    ("✅ Aprobados",              aprobados,  False),
]:
    if not grupo_items:
        continue
    with st.expander(f"{grupo_label} — {len(grupo_items)} ítems", expanded=expandido):
        for item in grupo_items:
            code    = item['code']
            precio  = item.get('valor_ofrecido', 0)
            n_comp  = sum(len(item.get(s,[])) for s in SECCIONES)
            suma    = sum(c['rend']*c['unit_price']
                          for s in SECCIONES for c in item.get(s, []))
            diff    = precio - suma
            cierra  = abs(diff) < 2
            alerta_mo = item.get('alerta_mo', False)
            pct_mo  = item.get('pct_mo', None)
            fuente  = item.get('fuente_revision', '—')
            aprobado = st.session_state.aprobacion.get(code, False)

            # Color de borde según estado
            if not aprobado:
                borde_color = '#D32F2F' if (not cierra or alerta_mo or n_comp == 0) else '#F57C00'
            else:
                borde_color = '#388E3C'

            st.markdown(f"""
            <div style='border-left:4px solid {borde_color};
                        padding:6px 12px; margin-bottom:4px;
                        background:#FAFAFA; border-radius:4px;'>
                <strong>{code}</strong> &nbsp;·&nbsp;
                <span style='font-size:0.9em'>{item['description'][:80]}</span><br>
                <span style='font-size:0.82em; color:#555'>
                    {item.get('unit','—')} &nbsp;|&nbsp;
                    <b>${precio:,.0f}</b> &nbsp;|&nbsp;
                    {n_comp} componentes &nbsp;|&nbsp;
                    {'✅ Cierra' if cierra else f'⚠️ Diferencia ${abs(diff):,.0f}'} &nbsp;|&nbsp;
                    {'⚠️ MO ' + str(pct_mo) + '%' if alerta_mo else ''} &nbsp;|&nbsp;
                    <em>{fuente}</em>
                </span>
            </div>
            """, unsafe_allow_html=True)

            c1, c2, c3 = st.columns([1, 1, 5])
            with c1:
                if st.button("✅ Aprobar", key=f"ap_{code}", use_container_width=True):
                    st.session_state.aprobacion[code] = True
                    st.rerun()
            with c2:
                if st.button("❌ Rechazar", key=f"rec_{code}", use_container_width=True):
                    st.session_state.aprobacion[code] = False
                    st.rerun()
            with c3:
                # Mostrar componentes desplegables
                with st.expander(f"Ver componentes ({n_comp})", expanded=False):
                    if n_comp == 0:
                        st.warning("Sin componentes — se exportará con estructura vacía.")
                    else:
                        for sec in SECCIONES:
                            comps = item.get(sec, [])
                            if not comps: continue
                            st.markdown(f"**{sec.upper().replace('_',' ')}**")
                            for c in comps:
                                und  = c.get('unit','') or '⚠️SIN UND'
                                rend = c.get('rend', 0)
                                up   = c.get('unit_price', 0)
                                parc = rend * up
                                und_color = 'red' if not c.get('unit','') else 'inherit'
                                st.markdown(
                                    f"&nbsp;&nbsp;• {c['description']} &nbsp;"
                                    f"<span style='color:{und_color}'>[{und}]</span> &nbsp;"
                                    f"rend={rend:.4f} × ${up:,.0f} = **${parc:,.0f}**",
                                    unsafe_allow_html=True
                                )
                            st.markdown("---")

                    # Resumen de cierre
                    if n_comp > 0:
                        col_s1, col_s2, col_s3 = st.columns(3)
                        col_s1.metric("Precio ofrecido", f"${precio:,.0f}")
                        col_s2.metric("Suma componentes", f"${suma:,.0f}")
                        col_s3.metric(
                            "Diferencia",
                            f"${abs(diff):,.0f}",
                            delta=f"{'OK' if cierra else 'Revisar'}",
                            delta_color="normal" if cierra else "inverse"
                        )

            st.markdown("")  # separador visual

# ══════════════════════════════════════════════════════════════════
# EXPORTAR SOLO LOS APROBADOS
# ══════════════════════════════════════════════════════════════════
st.divider()
seccion_exp = "7." if sin_apu else "6."
st.subheader(f"{seccion_exp} Exportar APUs aprobados")

items_aprobados = [i for i in lista_revision
                   if st.session_state.aprobacion.get(i['code'], False)]
items_rechazados = [i for i in lista_revision
                    if not st.session_state.aprobacion.get(i['code'], False)]

aiu_txt = f"con AIU ({aiu_pct*100:.1f}%)" if include_aiu else "sin AIU"

if not items_aprobados:
    st.warning("⚠️ No hay ítems aprobados. Apruebe al menos uno para exportar.")
else:
    if items_rechazados:
        st.info(
            f"ℹ️ Se exportarán **{len(items_aprobados)}** ítems aprobados. "
            f"**{len(items_rechazados)}** rechazados quedarán fuera del Excel."
        )
    else:
        st.success(f"✅ Todos los {len(items_aprobados)} ítems aprobados listos para exportar.")

    if st.button(
        f"🚀 Generar Excel — {len(items_aprobados)} APUs aprobados ({aiu_txt})",
        type="primary", use_container_width=True
    ):
        with st.spinner(f"Generando {len(items_aprobados)} APUs..."):
            try:
                # Separar aprobados en con_apu y manuales para respetar la firma de generate_apu_excel
                codes_con_apu   = {i['code'] for i in con_apu}
                aprobados_ref   = [i for i in items_aprobados if i['code'] in codes_con_apu]
                aprobados_otros = [i for i in items_aprobados if i['code'] not in codes_con_apu]

                resultado_filtrado = {
                    **resultado,
                    'items_con_apu': aprobados_ref,
                    'items_sin_apu': [],
                }

                excel = generate_apu_excel(
                    resultado_filtrado,
                    items_manuales=aprobados_otros if aprobados_otros else None,
                    include_aiu=include_aiu,
                    aiu_pct=aiu_pct,
                    bd_externas=bds_externas if bds_externas else None,
                )
                st.session_state.excel  = excel
                st.session_state.nombre = uploaded_proceso.name.replace('.xlsx','').replace('.xlsm','')
                st.success(
                    f"✅ **{len(items_aprobados)} APUs generados.** "
                    "Descargue el archivo a continuación."
                )
            except Exception as e:
                st.error(f"❌ Error al generar: {e}")
                raise e

    if "excel" in st.session_state:
        st.download_button(
            label="⬇️ Descargar APUs en Excel",
            data=st.session_state.excel,
            file_name=f"APUs_{st.session_state.get('nombre','export')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, type="primary",
        )

