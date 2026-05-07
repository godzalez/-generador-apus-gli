"""
Generador de APUs — Gerencia Legal Integral Colombia S.A.S.
v12: cascada de fuentes (entidad → Policía → INVIAS → genérica) + reporte cobertura.
"""
import streamlit as st
import io
from generator import leer_oferta_economica, leer_base_datos_apu, generate_apu_excel
from bases_externas import (
    cargar_catalogo_policia, cargar_base_externa,
    buscar_en_cascada, generar_reporte_cobertura,
)

st.set_page_config(page_title="Generador de APUs – GLI", page_icon="🏗️", layout="wide")

st.markdown("""
    <div style='background:#1B3A6B;padding:18px 24px;border-radius:8px;margin-bottom:18px'>
        <h2 style='color:white;margin:0'>🏗️ Generador Automático de APUs</h2>
        <p style='color:#BDD7EE;margin:4px 0 0 0'>
        Gerencia Legal Integral Colombia S.A.S. &nbsp;·&nbsp;
        Entidad, Policía Nacional, INVIAS o base externa.</p>
    </div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════════
# SIDEBAR — Configuración y fuentes externas
# ══════════════════════════════════════════════════════════════════
with st.sidebar:
    st.header("⚙️ Configuración")
    include_aiu = st.toggle("Incluir línea AIU en cada APU", value=False)
    aiu_pct = 0.0
    if include_aiu:
        val = st.number_input("Porcentaje AIU (%)", 0.0, 100.0, 25.0, 0.1, "%.2f")
        aiu_pct = val / 100.0
        st.info(f"AIU aplicado: **{val:.2f}%**")
    else:
        st.success("Sin AIU — precio = costo directo")

    st.divider()

    # ── FUENTE 1: Catálogo Policía Nacional ──────────────────────
    st.subheader("📚 Fuente 1 — Policía Nacional 2026")
    uploaded_catalogo = st.file_uploader(
        "Catálogo oficial (1LF-FR-0206)",
        type=["xlsx", "xlsm"], key="catalogo_policia",
        help="Archivo: 10. PRESUPUESTO APU ESPECIFICACIONES TECNICAS PARTICULARES.xlsx"
    )
    zona_proceso = st.selectbox(
        "Zona de accesibilidad",
        ["A1","A2","A3","A4","A5","A6"], index=0,
        help="A1=grandes ciudades · A4=municipios pequeños · A6=zona remota"
    )
    catalogo_policia = {}
    if uploaded_catalogo:
        data_cat = uploaded_catalogo.read()
        act, ins, eq, mo, err_cat = cargar_catalogo_policia(
            io.BytesIO(data_cat), zona=zona_proceso
        )
        if err_cat:
            st.error(f"❌ {err_cat}")
        else:
            catalogo_policia = {"actividades": act, "insumos": ins,
                                 "equipos": eq, "mano_obra": mo, "zona": zona_proceso}
            st.success(
                f"✅ Zona **{zona_proceso}** — "
                f"{len(act)} actividades · {len(ins)} insumos"
            )
    else:
        st.caption("Opcional — para procesos de la Policía Nacional.")

    st.divider()

    # ── FUENTE 2: APU INVIAS por provincia ───────────────────────
    st.subheader("📚 Fuente 2 — INVIAS regionalizado")
    st.caption(
        "Descargue el Excel de su provincia en "
        "[hermes.invias.gov.co](https://hermes.invias.gov.co/APUs/IndiceAPU/) "
        "y cárguelo aquí."
    )
    uploaded_invias = st.file_uploader(
        "APU INVIAS provincia (Excel)",
        type=["xlsx", "xlsm"], key="apu_invias",
    )
    bd_invias = {}
    if uploaded_invias:
        data_inv = uploaded_invias.read()
        bd_invias, fmt_inv, err_inv = cargar_base_externa(
            io.BytesIO(data_inv), nombre_fuente="INVIAS 2025-2"
        )
        if err_inv:
            st.error(f"❌ {err_inv}")
        elif not bd_invias:
            st.warning("⚠️ No se reconoció el formato INVIAS.")
        else:
            st.success(f"✅ {len(bd_invias)} actividades INVIAS cargadas.")

    st.divider()

    # ── FUENTE 3: Base genérica adicional ─────────────────────────
    st.subheader("📚 Fuente 3 — Base adicional")
    st.caption("Gobernación, IDU, municipio u otra entidad.")
    uploaded_extra = st.file_uploader(
        "Base de precios adicional (Excel)",
        type=["xlsx", "xlsm"], key="base_extra",
    )
    bd_extra = {}
    nombre_extra = st.text_input("Nombre de esta fuente", value="Base externa",
                                  key="nombre_extra")
    if uploaded_extra:
        data_ext = uploaded_extra.read()
        bd_extra, fmt_ext, err_ext = cargar_base_externa(
            io.BytesIO(data_ext), nombre_fuente=nombre_extra
        )
        if err_ext:
            st.error(f"❌ {err_ext}")
        elif not bd_extra:
            st.warning("⚠️ No se reconoció el formato.")
        else:
            st.success(f"✅ {len(bd_extra)} actividades '{nombre_extra}' cargadas.")

    st.divider()
    st.caption("v12.0 · GLI Colombia · 2026")

# ══════════════════════════════════════════════════════════════════
# PASO 1 — Oferta
# ══════════════════════════════════════════════════════════════════
st.subheader("1. Oferta económica del proponente")
st.caption("El sistema detecta automáticamente la hoja y columnas.")
uploaded_oferta = st.file_uploader(
    "Excel de la oferta (PROPUESTA, FORMULARIO, OFERTA…)",
    type=["xlsx","xlsm"], key="oferta"
)

# ══════════════════════════════════════════════════════════════════
# PASO 2 — APU entidad (opcional en Camino B)
# ══════════════════════════════════════════════════════════════════
st.subheader("2. Archivo APU de la entidad contratante")
st.caption(
    "Si la entidad publicó sus APUs, cárguelos aquí (Camino A). "
    "Si no los publicó, deje en blanco y el sistema usará las fuentes del sidebar (Camino B)."
)
uploaded_apu = st.file_uploader(
    "Excel APU de la entidad (opcional)",
    type=["xlsx","xlsm"], key="apu_entidad"
)

if not uploaded_oferta:
    st.info("👆 Cargue primero la oferta económica para continuar.")
    st.stop()

# ══════════════════════════════════════════════════════════════════
# LECTURA
# ══════════════════════════════════════════════════════════════════
bd_entidad = {}
if uploaded_apu:
    with st.spinner("Leyendo APU de la entidad…"):
        bd_entidad, err_bd = leer_base_datos_apu(uploaded_apu)
    if err_bd and not bd_entidad:
        st.warning(f"⚠️ No se pudo leer el APU de la entidad: {err_bd}")

with st.spinner("Leyendo oferta económica…"):
    resultado, err_oferta = leer_oferta_economica(
        uploaded_oferta, bd_referencia=bd_entidad
    )

if err_oferta:
    st.error(f"❌ {err_oferta}")
    st.stop()

con_apu = resultado["items_con_apu"]
sin_apu = resultado["items_sin_apu"]
n_total = resultado["total_proceso"]
hoja    = resultado["hoja_usada"]

# ══════════════════════════════════════════════════════════════════
# CAMINO B — Cascada de fuentes para ítems sin APU
# ══════════════════════════════════════════════════════════════════
fuentes_cascada = []
if catalogo_policia.get("actividades"):
    fuentes_cascada.append({
        "tipo": "policia", "nombre": f"Policía Nacional 2026 · zona {zona_proceso}",
        "actividades": catalogo_policia["actividades"],
    })
if bd_invias:
    fuentes_cascada.append({"tipo": "invias", "nombre": "INVIAS 2025-2", "bd": bd_invias})
if bd_extra:
    fuentes_cascada.append({"tipo": "generica", "nombre": nombre_extra, "bd": bd_extra})

items_cascada_ok  = []
items_cascada_no  = []

if sin_apu and fuentes_cascada:
    with st.spinner(f"Buscando {len(sin_apu)} ítems sin APU en fuentes externas…"):
        for item in sin_apu:
            apu_ext = buscar_en_cascada(
                item["description"], item["unit"],
                item["valor_ofrecido"], fuentes_cascada
            )
            if apu_ext:
                item.update(apu_ext)
                item["tiene_apu"] = True
                item["_via_cascada"] = True
                items_cascada_ok.append(item)
            else:
                item["_via_cascada"] = False
                items_cascada_no.append(item)
    # Mover los encontrados a con_apu
    con_apu = list(con_apu) + items_cascada_ok
    sin_apu = items_cascada_no

# ══════════════════════════════════════════════════════════════════
# PASO 3 — DIAGNÓSTICO + REPORTE DE COBERTURA
# ══════════════════════════════════════════════════════════════════
st.subheader("3. Diagnóstico del cruce")
st.caption(f"Hoja leída: **{hoja}** · APUs entidad: **{len(bd_entidad)}**")

c1, c2, c3, c4 = st.columns(4)
c1.metric("Total ítems oferta", n_total)
c2.metric("APU entidad (Camino A)", resultado["total_proceso"] - len(resultado["items_sin_apu"]))
c3.metric("APU fuentes externas (Camino B)", len(items_cascada_ok))
c4.metric("Sin APU", len(sin_apu),
          delta="⚠️ Pendientes" if sin_apu else "✅ Ninguno",
          delta_color="inverse" if sin_apu else "normal")

# Reporte de cobertura
reporte = generar_reporte_cobertura(con_apu, sin_apu)
pct = reporte["pct_cobertura"]
color_pct = "#388E3C" if pct == 100 else ("#F57C00" if pct >= 70 else "#D32F2F")
st.markdown(
    f"<div style='background:{color_pct};color:white;padding:10px 18px;"
    f"border-radius:6px;font-size:1.1em;margin:8px 0'>"
    f"<b>Cobertura total: {pct:.1f}%</b> — "
    f"{reporte['con_apu']} de {reporte['total']} ítems con APU</div>",
    unsafe_allow_html=True
)

if reporte["por_fuente"]:
    with st.expander("📊 Ver cobertura por fuente", expanded=False):
        for fuente, cant in reporte["por_fuente"].items():
            st.markdown(f"• **{fuente}**: {cant} ítems")

if sin_apu:
    st.warning(f"⚠️ **{len(sin_apu)} ítem(s) sin APU** en ninguna fuente disponible.")
    with st.expander(f"Ver {len(sin_apu)} ítems sin APU — completar manualmente", expanded=True):
        for it in sin_apu:
            st.markdown(
                f"**`{it['code']}`** — {it['description'][:80]} "
                f"| {it.get('unit','—')} | **${it['valor_ofrecido']:,.0f}**"
            )
else:
    st.success(f"✅ **Cobertura completa.** Todos los {len(con_apu)} ítems tienen APU.")

# ══════════════════════════════════════════════════════════════════
# PASO 4 — REVISIÓN
# ══════════════════════════════════════════════════════════════════
st.subheader("4. Revisión y aprobación de APUs")
st.caption("Solo los **aprobados** se incluirán en el Excel final.")

SECCIONES = ("materiales","herramientas","transporte","mano_de_obra")
lista_revision = list(con_apu) + list(sin_apu)

for item in lista_revision:
    n = sum(len(item.get(s,[])) for s in SECCIONES)
    via = item.get("_via_cascada", False)
    fuente_bd = item.get("fuente_bd","")
    if n > 0 and not via:
        item["_fuente"] = "APU entidad"
    elif n > 0 and via:
        item["_fuente"] = fuente_bd or "Fuente externa"
    else:
        item["_fuente"] = "Sin APU"

if "aprobacion" not in st.session_state:
    st.session_state.aprobacion = {}

for item in lista_revision:
    code = item["code"]
    if code not in st.session_state.aprobacion:
        n     = sum(len(item.get(s,[])) for s in SECCIONES)
        precio = item.get("valor_ofrecido",0)
        suma  = sum(c["rend"]*c["unit_price"] for s in SECCIONES for c in item.get(s,[]))
        cierra = abs(precio-suma) < 2
        alerta = item.get("alerta_mo",False)
        st.session_state.aprobacion[code] = (n > 0 and cierra and not alerta)

ca, cr, ci = st.columns([1,1,3])
with ca:
    if st.button("✅ Aprobar todos", use_container_width=True):
        for i in lista_revision: st.session_state.aprobacion[i["code"]] = True
        st.rerun()
with cr:
    if st.button("❌ Rechazar todos", use_container_width=True):
        for i in lista_revision: st.session_state.aprobacion[i["code"]] = False
        st.rerun()
with ci:
    n_ap = sum(1 for v in st.session_state.aprobacion.values() if v)
    st.info(f"**{n_ap}** aprobados · **{len(lista_revision)-n_ap}** pendientes de **{len(lista_revision)}** total")

st.divider()

pendientes = [i for i in lista_revision if not st.session_state.aprobacion.get(i["code"],False)]
aprobados  = [i for i in lista_revision if     st.session_state.aprobacion.get(i["code"],False)]

for label, grupo, abierto in [
    ("⚠️ Pendientes de revisión", pendientes, True),
    ("✅ Aprobados",               aprobados,  False),
]:
    if not grupo: continue
    with st.expander(f"{label} — {len(grupo)} ítems", expanded=abierto):
        for item in grupo:
            code   = item["code"]
            precio = item.get("valor_ofrecido",0)
            n_comp = sum(len(item.get(s,[])) for s in SECCIONES)
            suma   = sum(c["rend"]*c["unit_price"] for s in SECCIONES for c in item.get(s,[]))
            diff   = precio - suma
            cierra = abs(diff) < 2
            alerta_mo = item.get("alerta_mo",False)
            pct_mo    = item.get("pct_mo",None)
            aprobado  = st.session_state.aprobacion.get(code,False)
            via_cascada = item.get("_via_cascada", False)

            color = "#388E3C" if aprobado else ("#D32F2F" if (n_comp==0 or not cierra or alerta_mo) else "#F57C00")
            tags  = []
            if n_comp == 0:   tags.append("🔴 Sin componentes")
            if not cierra:    tags.append(f"⚠️ Diferencia ${abs(diff):,.0f}")
            if alerta_mo:     tags.append(f"⚠️ MO {pct_mo}%")
            if via_cascada:   tags.append(f"🔵 {item['_fuente']}")
            estado = " &nbsp;·&nbsp; ".join(tags) if tags else "✅ OK"

            st.markdown(f"""
            <div style='border-left:4px solid {color};padding:6px 14px;
                        margin-bottom:6px;background:#FAFAFA;border-radius:4px;'>
                <strong>{code}</strong> &nbsp;·&nbsp;
                <span style='font-size:.9em'>{item['description'][:85]}</span><br>
                <span style='font-size:.82em;color:#555'>
                    <b>{item.get('unit','—')}</b> &nbsp;|&nbsp;
                    Ofrecido: <b>${precio:,.0f}</b> &nbsp;|&nbsp;
                    Componentes: <b>${suma:,.0f}</b> &nbsp;|&nbsp;
                    {estado}
                </span>
            </div>""", unsafe_allow_html=True)

            c1, c2, c3 = st.columns([1,1,5])
            with c1:
                if st.button("✅ Aprobar", key=f"ap_{code}", use_container_width=True):
                    st.session_state.aprobacion[code] = True
                    st.rerun()
            with c2:
                if st.button("❌ Rechazar", key=f"rec_{code}", use_container_width=True):
                    st.session_state.aprobacion[code] = False
                    st.rerun()
            with c3:
                with st.expander(f"Ver {n_comp} componentes", expanded=False):
                    if n_comp == 0:
                        st.warning("Sin APU disponible — se exportará vacío para completar manualmente.")
                    else:
                        for sec in SECCIONES:
                            comps = item.get(sec,[])
                            if not comps: continue
                            st.markdown(f"**{sec.replace('_',' ').upper()}**")
                            for c in comps:
                                und  = c.get("unit","") or "⚠️ SIN UND"
                                rend = c.get("rend",0)
                                up   = c.get("unit_price",0)
                                cu   = "red" if not c.get("unit","") else "inherit"
                                est  = " *(ajustado)*" if c.get("_rend_ajustado") else ""
                                st.markdown(
                                    f"&nbsp;&nbsp;• {c['description']} "
                                    f"<span style='color:{cu}'>[{und}]</span> "
                                    f"&nbsp;{rend:.4f} × ${up:,.0f} = **${rend*up:,.0f}**{est}",
                                    unsafe_allow_html=True)
                        m1,m2,m3 = st.columns(3)
                        m1.metric("Precio ofrecido",  f"${precio:,.0f}")
                        m2.metric("Suma componentes", f"${suma:,.0f}")
                        m3.metric("Diferencia", f"${abs(diff):,.0f}",
                                  delta="OK" if cierra else "Revisar",
                                  delta_color="normal" if cierra else "inverse")
            st.markdown("")

# ══════════════════════════════════════════════════════════════════
# PASO 5 — EXPORTAR
# ══════════════════════════════════════════════════════════════════
st.divider()
st.subheader("5. Exportar APUs aprobados")

items_aprobados  = [i for i in lista_revision if  st.session_state.aprobacion.get(i["code"],False)]
items_rechazados = [i for i in lista_revision if not st.session_state.aprobacion.get(i["code"],False)]
aiu_txt = f"con AIU ({aiu_pct*100:.1f}%)" if include_aiu else "sin AIU"

if not items_aprobados:
    st.warning("⚠️ No hay ítems aprobados. Apruebe al menos uno para exportar.")
    st.stop()

if items_rechazados:
    st.info(f"Se exportarán **{len(items_aprobados)}** ítems aprobados. **{len(items_rechazados)}** rechazados quedarán fuera.")
else:
    st.success(f"✅ Todos los **{len(items_aprobados)}** ítems aprobados y listos.")

if st.button(f"🚀 Generar Excel — {len(items_aprobados)} APUs ({aiu_txt})", type="primary", use_container_width=True):
    with st.spinner(f"Generando {len(items_aprobados)} APUs…"):
        try:
            codes_ref        = {i["code"] for i in con_apu}
            aprobados_ref    = [i for i in items_aprobados if     i["code"] in codes_ref]
            aprobados_vacios = [i for i in items_aprobados if not i["code"] in codes_ref]
            resultado_export = {**resultado, "items_con_apu": aprobados_ref, "items_sin_apu": []}
            excel = generate_apu_excel(
                resultado_export,
                items_manuales=aprobados_vacios if aprobados_vacios else None,
                include_aiu=include_aiu, aiu_pct=aiu_pct, bd_externas=None,
            )
            st.session_state.excel  = excel
            st.session_state.nombre = uploaded_oferta.name.replace(".xlsx","").replace(".xlsm","")
            st.success(f"✅ **{len(items_aprobados)} APUs generados.** Descargue el archivo.")
        except Exception as e:
            st.error(f"❌ Error al generar el Excel: {e}")
            raise e

if "excel" in st.session_state:
    st.download_button(
        label="⬇️ Descargar APUs en Excel",
        data=st.session_state.excel,
        file_name=f"APUs_{st.session_state.get('nombre','export')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True, type="primary",
    )
