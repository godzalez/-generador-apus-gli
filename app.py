"""
Generador Automático de APUs — Gerencia Legal Integral Colombia S.A.S.
v11: Textos claros + AIU leído automáticamente del archivo de la entidad.
"""
import streamlit as st
import sys, traceback

# ── Import con diagnóstico visible ───────────────────────────────────────────
try:
    import pandas as pd
    from generator import (
        leer_apu_entidad,
        leer_propuesta_economica,
        cruzar_y_ajustar,
        generate_apu_excel,
    )
except Exception as _e:
    st.set_page_config(page_title="Error – APUs GLI", page_icon="❌")
    st.error(f"**Error al cargar la aplicación:** `{type(_e).__name__}: {_e}`")
    st.code(traceback.format_exc(), language="text")
    st.info(f"Python {sys.version} | sys.path: {sys.path[:3]}")
    st.stop()

# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Generador de APUs – GLI", page_icon="🏗️", layout="wide")

st.markdown("""
    <div style='background-color:#1B3A6B;padding:18px 24px;border-radius:8px;margin-bottom:18px'>
        <h2 style='color:white;margin:0'>🏗️ Generador Automático de APUs</h2>
        <p style='color:#BDD7EE;margin:4px 0 0 0'>
        Gerencia Legal Integral Colombia S.A.S. &nbsp;·&nbsp;
        Ajusta los APUs de la entidad al precio de su propuesta económica.</p>
    </div>
""", unsafe_allow_html=True)

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Configuración")
    st.info(
        "**AIU automático**  \n"
        "El factor AIU se lee directamente del archivo de la entidad "
        "(columna J de la hoja APU). No necesita ingresarlo manualmente."
    )
    st.divider()
    modo_avanzado = st.toggle("🔧 Modo avanzado — bases externas", value=False)
    if modo_avanzado:
        st.caption(
            "Activa la carga de bases de precios externas (Gobernación, INVIAS) "
            "para completar ítems que no estén en el archivo de la entidad."
        )
    st.divider()
    st.markdown("⚠️ Use descripciones **genéricas** (sin marcas comerciales).")
    st.caption("v13.0 · GLI Colombia · 2026")


# ══════════════════════════════════════════════════════════════════════════════
# PASO 1 — APUs de la Entidad
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
col_icon, col_title = st.columns([0.06, 0.94])
col_icon.markdown("## 1️⃣")
col_title.markdown("## APUs de la Entidad")

st.markdown("""
> **¿Qué archivo cargar aquí?**  
> El archivo que **la entidad contratante le entregó** con los análisis de precios unitarios del proceso.  
> Puede ser **Excel (.xlsx)** o **PDF** — la aplicación detecta el formato automáticamente.  
> Debe contener los componentes de cada ítem: materiales, equipos, mano de obra y transporte.  
> La aplicación tomará estos componentes como base y **solo ajustará el rendimiento de la mano de obra**.
""")

uploaded_entidad = st.file_uploader(
    "📂 Seleccione el archivo de APUs de la Entidad",
    type=["xlsx", "xlsm", "pdf"],
    key="entidad",
    help="Excel (.xlsx) o PDF con los APUs de la entidad. El sistema detecta el formato automáticamente."
)
if uploaded_entidad:
    st.success(f"✅ Archivo cargado: **{uploaded_entidad.name}**")


# ══════════════════════════════════════════════════════════════════════════════
# PASO 2 — Mi Propuesta Económica
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
col_icon, col_title = st.columns([0.06, 0.94])
col_icon.markdown("## 2️⃣")
col_title.markdown("## Mi Propuesta Económica")

st.markdown("""
> **¿Qué archivo cargar aquí?**  
> El formulario o presupuesto con **los precios que usted va a ofrecer** en el proceso.
> La aplicación ajustará el rendimiento de mano de obra para que cada APU cierre
> **exactamente** con ese valor (a dos decimales).
""")

uploaded_propuesta = st.file_uploader(
    "📂 Seleccione el archivo con su Propuesta Económica",
    type=["xlsx", "xlsm"],
    key="propuesta",
    help="Su formulario de oferta económica con los precios unitarios que va a proponer."
)
if uploaded_propuesta:
    st.success(f"✅ Archivo cargado: **{uploaded_propuesta.name}**")


# ══════════════════════════════════════════════════════════════════════════════
# PASO 3 (Modo avanzado) — Bases externas
# ══════════════════════════════════════════════════════════════════════════════
bds_externas = []
if modo_avanzado:
    st.markdown("---")
    col_icon, col_title = st.columns([0.06, 0.94])
    col_icon.markdown("## 3️⃣")
    col_title.markdown("## Bases de Precios Externas *(opcional)*")

    st.markdown("""
    > **¿Para qué sirve esto?**  
    > Si algunos ítems de su propuesta **no están en el archivo de la entidad**, puede cargar
    > bases de precios de Gobernación de Boyacá, INVIAS u otras para completarlos.
    > Si no carga nada aquí, esos ítems quedarán en rojo para completar manualmente.
    """)

    uploaded_bases = st.file_uploader(
        "📂 Seleccione bases de precios externas (puede cargar varios archivos)",
        type=["xlsx", "xlsm"],
        key="bases_ext",
        accept_multiple_files=True,
    )
    if uploaded_bases:
        from bases_externas import cargar_base_externa
        for ub in uploaded_bases:
            bd_ext, fmt_ext, err_ext = cargar_base_externa(
                ub, nombre_fuente=ub.name.replace('.xlsx','').replace('.xlsm','')
            )
            if bd_ext:
                bds_externas.append(bd_ext)
                st.success(f"✅ **{ub.name}**: {len(bd_ext)} ítems ({fmt_ext})")
            else:
                st.warning(f"⚠️ {ub.name}: {err_ext}")


# ══════════════════════════════════════════════════════════════════════════════
# Validación
# ══════════════════════════════════════════════════════════════════════════════
if not uploaded_entidad or not uploaded_propuesta:
    st.markdown("---")
    faltantes = []
    if not uploaded_entidad:   faltantes.append("**APUs de la Entidad** (Paso 1)")
    if not uploaded_propuesta: faltantes.append("**Propuesta Económica** (Paso 2)")
    st.info(f"👆 Para continuar, cargue: {' y '.join(faltantes)}.")
    st.stop()


# ══════════════════════════════════════════════════════════════════════════════
# Procesamiento
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
n_paso = 4 if modo_avanzado else 3

with st.spinner("Leyendo APUs de la entidad..."):
    bd_entidad, err_entidad = leer_apu_entidad(uploaded_entidad)

if err_entidad:
    st.error(f"❌ No se pudo leer el archivo de APUs de la entidad: {err_entidad}")
    st.stop()

aiu_values = [v.get('aiu_factor', 1.0) for v in bd_entidad.values() if v.get('aiu_factor', 1.0) > 1.0]
if aiu_values:
    aiu_prom   = sum(aiu_values) / len(aiu_values)
    aiu_unico  = len(set(round(x, 4) for x in aiu_values)) == 1
    aiu_txt    = f"{(aiu_prom-1)*100:.2f}%" if aiu_unico else f"variable ({(min(aiu_values)-1)*100:.2f}% – {(max(aiu_values)-1)*100:.2f}%)"
    st.success(f"✅ APUs de la entidad: **{len(bd_entidad)}** ítems  |  Factor AIU detectado: **{aiu_txt}**")
else:
    st.success(f"✅ APUs de la entidad: **{len(bd_entidad)}** ítems (sin AIU)")

with st.spinner("Leyendo propuesta económica..."):
    propuesta, err_prop = leer_propuesta_economica(uploaded_propuesta)

if err_prop:
    st.error(f"❌ No se pudo leer la propuesta económica: {err_prop}")
    st.stop()

st.success(f"✅ Propuesta económica: **{len(propuesta)}** ítems leídos")

with st.spinner("Cruzando ítems y ajustando rendimientos de mano de obra..."):
    resultado = cruzar_y_ajustar(propuesta, bd_entidad)

con_apu = resultado['items_ajustados']
sin_apu = resultado['items_sin_apu']
n_total = resultado['total']
cruces  = resultado['detalle_cruce']


# ══════════════════════════════════════════════════════════════════════════════
# Resultado del cruce
# ══════════════════════════════════════════════════════════════════════════════
col_icon, col_title = st.columns([0.06, 0.94])
col_icon.markdown(f"## {n_paso}️⃣")
col_title.markdown("## Resultado del Cruce")
n_paso += 1

c1, c2, c3 = st.columns(3)
c1.metric("Total ítems en propuesta", n_total)
c2.metric("✅ APUs ajustados",         len(con_apu))
c3.metric("🔴 Sin APU (completar)",   len(sin_apu))

if con_apu:
    with st.expander(f"✅ Ver los {len(con_apu)} ítems ajustados", expanded=False):
        filas = []
        for item in con_apu:
            aiu_f  = item.get('aiu_factor', 1.0)
            cd     = item.get('cd_final', 0)
            precio = item['valor_ofrecido']
            diff   = abs(precio - round(cd * aiu_f, 2))
            filas.append({
                'Ítem':          item['code'],
                'Descripción':   item['description'][:55],
                'Unidad':        item['unit'],
                'Precio ($)':    f"${precio:,.2f}",
                'Costo Directo': f"${cd:,.2f}",
                'AIU':           f"{(aiu_f-1)*100:.2f}%" if aiu_f > 1 else '—',
                'Cierre':        "✅" if diff < 0.05 else f"⚠️ dif=${diff:,.2f}",
                'Ajuste vía':    cruces.get(item['code'], {}).get('metodo_ajuste', ''),
            })
        st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

if sin_apu:
    with st.expander(
        f"🔴 {len(sin_apu)} ítems que quedarán en ROJO — completar manualmente",
        expanded=True
    ):
        st.warning(
            "Estos ítems no se encontraron en el archivo de APUs de la entidad.  \n"
            "Se incluirán en el Excel con la hoja **vacía y en rojo**."
        )
        filas_r = []
        for item in sin_apu:
            info = cruces.get(item['code'], {})
            filas_r.append({
                'Ítem':        item['code'],
                'Descripción': item['description'][:65],
                'Precio ($)':  f"${item['valor_ofrecido']:,.2f}",
                'Motivo':      info.get('razon', info.get('metodo', '—')),
            })
        st.dataframe(pd.DataFrame(filas_r), use_container_width=True, hide_index=True)


# ══════════════════════════════════════════════════════════════════════════════
# Generación del Excel
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
col_icon, col_title = st.columns([0.06, 0.94])
col_icon.markdown(f"## {n_paso}️⃣")
col_title.markdown("## Generar y Descargar APUs en Excel")

n_gen = len(con_apu) + len(sin_apu)

if n_gen == 0:
    st.warning("No hay ítems para generar.")
else:
    st.markdown(f"""
    El Excel resultante tendrá:
    - **{len(con_apu)} hojas en verde** — APUs ajustados, rendimiento MO calibrado al precio ofrecido
    - **{len(sin_apu)} hojas en rojo** — ítems sin APU para completar manualmente
    - **1 hoja RESUMEN** — tabla general de todos los ítems
    """)

    if st.button(f"🚀 Generar {n_gen} APU(s) en Excel", type="primary", use_container_width=True):
        with st.spinner(f"Generando {n_gen} APUs..."):
            try:
                excel_bytes = generate_apu_excel(
                    items_ajustados=con_apu,
                    items_sin_apu=sin_apu,
                    bd_externas=bds_externas if bds_externas else None,
                )
                st.session_state['excel_bytes']    = excel_bytes
                st.session_state['nombre_archivo'] = (
                    uploaded_propuesta.name.replace('.xlsx','').replace('.xlsm','')
                )
                st.success(
                    f"✅ **{len(con_apu)} APUs generados** — cierra exactamente con su oferta.  \n"
                    f"🔴 **{len(sin_apu)} ítems en rojo** — completar manualmente."
                )
            except Exception as e:
                st.error(f"❌ Error al generar: {e}")
                st.code(traceback.format_exc(), language="text")

    if 'excel_bytes' in st.session_state:
        nombre_dl = f"APUs_{st.session_state['nombre_archivo']}.xlsx"
        st.download_button(
            label="⬇️ Descargar APUs en Excel",
            data=st.session_state['excel_bytes'],
            file_name=nombre_dl,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            type="primary",
        )
