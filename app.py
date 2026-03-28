"""
Generador Automático de APUs — Gerencia Legal Integral Colombia S.A.S.
v10: Flujo rediseñado — APUs Entidad + Propuesta Económica.
     Motor de ajuste por rendimiento de mano de obra (Reglas Generales).
     Bases externas desactivadas por defecto (modo avanzado).
"""
import streamlit as st
import pandas as pd
import io
from generator import (
    leer_apu_entidad,
    leer_propuesta_economica,
    cruzar_y_ajustar,
    generate_apu_excel,
)

st.set_page_config(page_title="Generador de APUs – GLI", page_icon="🏗️", layout="wide")

st.markdown("""
    <div style='background-color:#1B3A6B;padding:18px 24px;border-radius:8px;margin-bottom:18px'>
        <h2 style='color:white;margin:0'>🏗️ Generador Automático de APUs</h2>
        <p style='color:#BDD7EE;margin:4px 0 0 0'>
        Gerencia Legal Integral Colombia S.A.S. &nbsp;·&nbsp;
        Ajusta los APUs de la entidad a los precios de su propuesta económica.</p>
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

    # Modo avanzado — bases externas (oculto por defecto)
    modo_avanzado = st.toggle("🔧 Modo avanzado (bases externas)", value=False)

    st.divider()
    st.markdown("⚠️ Use descripciones **genéricas** (sin marcas comerciales).")
    st.caption("v10.0 · GLI Colombia · 2026")

# ── PASO 1: APUs de la Entidad ────────────────────────────────────────────────
st.subheader("1. APUs de la Entidad")
st.caption(
    "Cargue el archivo Excel que contiene los APUs del proceso de contratación "
    "(archivo entregado por la entidad). Debe contener **todos los ítems** con sus "
    "componentes: materiales, equipos, herramientas, transporte y mano de obra."
)
uploaded_entidad = st.file_uploader(
    "Excel con APUs de la Entidad",
    type=["xlsx", "xlsm"], key="entidad"
)

# ── PASO 2: Mi Propuesta Económica ────────────────────────────────────────────
st.subheader("2. Mi Propuesta Económica")
st.caption(
    "Cargue el archivo Excel con los precios que usted va a ofrecer. "
    "El sistema cruzará ítem por ítem y ajustará el rendimiento de la mano de obra "
    "para que cada APU cierre exactamente con su precio ofrecido (a 2 decimales)."
)
uploaded_propuesta = st.file_uploader(
    "Excel con su Propuesta Económica (PRESUPUESTO, PROPUESTA, FORMULARIO, etc.)",
    type=["xlsx", "xlsm"], key="propuesta"
)

if not uploaded_entidad or not uploaded_propuesta:
    pasos = []
    if not uploaded_entidad:   pasos.append("APUs de la Entidad")
    if not uploaded_propuesta: pasos.append("Propuesta Económica")
    st.info(f"👆 Cargue: **{' y '.join(pasos)}** para continuar.")
    st.stop()

# ── PASO 3 (Modo avanzado): Bases externas ────────────────────────────────────
bds_externas = []
if modo_avanzado:
    st.subheader("3. Bases de precios externas (avanzado)")
    st.caption(
        "Opcional. Solo para ítems que no tengan APU en el archivo de la entidad. "
        "Puede cargar bases de Gobernación de Boyacá, INVIAS u otras."
    )
    uploaded_bases = st.file_uploader(
        "Bases de precios externas",
        type=["xlsx", "xlsm"], key="bases_ext", accept_multiple_files=True
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

# ── Procesar archivos ─────────────────────────────────────────────────────────
n_paso = 4 if modo_avanzado else 3

with st.spinner("Leyendo APUs de la entidad..."):
    bd_entidad, err_entidad = leer_apu_entidad(uploaded_entidad)

if err_entidad:
    st.error(f"❌ Error en APUs de la entidad: {err_entidad}")
    st.stop()

st.success(f"✅ APUs de la entidad: **{len(bd_entidad)}** ítems con componentes cargados.")

with st.spinner("Leyendo propuesta económica..."):
    propuesta, err_prop = leer_propuesta_economica(uploaded_propuesta)

if err_prop:
    st.error(f"❌ Error en propuesta económica: {err_prop}")
    st.stop()

st.success(f"✅ Propuesta económica: **{len(propuesta)}** ítems leídos.")

# ── Cruce y ajuste ────────────────────────────────────────────────────────────
with st.spinner("Cruzando ítems y ajustando rendimientos de mano de obra..."):
    resultado = cruzar_y_ajustar(propuesta, bd_entidad)

con_apu = resultado['items_ajustados']
sin_apu = resultado['items_sin_apu']
n_total = resultado['total']
cruces  = resultado['detalle_cruce']

# ── Resumen del cruce ─────────────────────────────────────────────────────────
st.subheader(f"{n_paso}. Resultado del cruce")
n_paso += 1

c1, c2, c3 = st.columns(3)
c1.metric("Total ítems en propuesta",  n_total)
c2.metric("✅ Con APU ajustado",        len(con_apu))
c3.metric("🔴 Sin APU (marcar rojo)",  len(sin_apu))

# Detalle ítems ajustados
if con_apu:
    with st.expander(f"✅ {len(con_apu)} ítems ajustados correctamente", expanded=False):
        filas = []
        for item in con_apu:
            precio = item['valor_ofrecido']
            suma   = sum(
                round(c['rend'] * c['unit_price'], 2)
                for s in ('materiales', 'herramientas', 'transporte', 'mano_de_obra')
                for c in item.get(s, [])
            )
            diff   = abs(precio - suma)
            cierre = "✅" if diff < 0.02 else f"⚠️ dif=${diff:,.2f}"
            metodo = cruces.get(item['code'], {}).get('metodo', '')
            filas.append({
                'Ítem':        item['code'],
                'Descripción': item['description'][:60],
                'Unidad':      item['unit'],
                'Precio ($)':  f"${precio:,.2f}",
                'Cierre':      cierre,
                'Cruce por':   metodo,
            })
        st.dataframe(pd.DataFrame(filas), use_container_width=True, hide_index=True)

# Detalle ítems sin APU
if sin_apu:
    with st.expander(
        f"🔴 {len(sin_apu)} ítems SIN APU en la entidad — se marcarán en rojo",
        expanded=True
    ):
        st.warning(
            "Estos ítems no tienen APU correspondiente en el archivo de la entidad. "
            "Se incluirán en el Excel con la hoja **vacía y en rojo** para que los complete manualmente."
        )
        filas_r = []
        for item in sin_apu:
            razon = cruces.get(item['code'], {}).get('razon', 'No encontrado en APUs de la entidad')
            filas_r.append({
                'Ítem':        item['code'],
                'Descripción': item['description'][:65],
                'Precio ($)':  f"${item['valor_ofrecido']:,.2f}",
                'Motivo':      razon,
            })
        st.dataframe(pd.DataFrame(filas_r), use_container_width=True, hide_index=True)

# ── Generación del Excel ──────────────────────────────────────────────────────
st.subheader(f"{n_paso}. Generar APUs en Excel")

n_gen   = len(con_apu) + len(sin_apu)
aiu_txt = f"con AIU ({aiu_pct*100:.1f}%)" if include_aiu else "sin AIU"

if n_gen == 0:
    st.warning("No hay ítems para generar.")
else:
    if st.button(
        f"🚀 Generar {n_gen} APU(s) — {aiu_txt}",
        type="primary", use_container_width=True
    ):
        with st.spinner(f"Generando {n_gen} APUs en Excel..."):
            try:
                excel_bytes = generate_apu_excel(
                    items_ajustados=con_apu,
                    items_sin_apu=sin_apu,
                    include_aiu=include_aiu,
                    aiu_pct=aiu_pct,
                    bd_externas=bds_externas if bds_externas else None,
                )
                st.session_state['excel_bytes']    = excel_bytes
                st.session_state['nombre_archivo'] = (
                    uploaded_propuesta.name.replace('.xlsx','').replace('.xlsm','')
                )
                n_ok  = len(con_apu)
                n_roj = len(sin_apu)
                st.success(
                    f"✅ **{n_ok} APUs ajustados** — precio unitario cierra exactamente con su oferta.  \n"
                    f"🔴 **{n_roj} ítems en rojo** — sin componentes, para completar manualmente."
                )
            except Exception as e:
                st.error(f"❌ Error al generar: {e}")
                raise e

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
