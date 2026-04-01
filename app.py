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
