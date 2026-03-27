"""
Generador de APUs - Gerencia Legal Integral Colombia S.A.S.
v7: detección robusta — acepta cualquier variante de encabezado y número de ítem.
"""
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io, re


# ══════════════════════════════════════════════════════════════════
# UTILIDADES DE DETECCIÓN — importadas desde detector.py
# ══════════════════════════════════════════════════════════════════
from numero_letras import numero_a_letras
from bases_externas import buscar_en_base, construir_apu_desde_base
from salarios import jornal_minimo, determinar_año_proceso
from detector import (
    detectar_columnas as _detectar_columnas_ext,
    es_fila_item_valida as _es_fila_item_valida_ext,
    extraer_codigo as _extraer_codigo_ext,
    es_codigo_item,
    encontrar_hoja_presupuesto as _encontrar_hoja_presupuesto,
)

# Alias para compatibilidad con el resto del código
def _detectar_columnas(ws):     return _detectar_columnas_ext(ws)
def _es_fila_item(r, m):        return _es_fila_item_valida_ext(r, m)
def _extraer_codigo(r, m):      return _extraer_codigo_ext(r, m)

# ══════════════════════════════════════════════════════════════════
# LECTURA DE APUs EN HOJAS INDIVIDUALES (formato "APU X.XX")
# ══════════════════════════════════════════════════════════════════

def _leer_apu_hoja_individual(ws):
    """
    Lee un APU desde una hoja individual tipo SENA/Gobernación.
    Estructura: INSUMOS/MATERIALES, EQUIPO/HERRAMIENTA, MANO DE OBRA
    con columnas: ÍTEM | DESCRIPCIÓN | UNIDAD | V. UNITARIO | CANTIDAD | V. PARCIAL
    Retorna dict {total_referencia, materiales, herramientas, transporte, mano_de_obra}.
    """
    apu = {'total_referencia': 0,
           'materiales': [], 'herramientas': [], 'transporte': [], 'mano_de_obra': []}

    seccion_actual = None
    COL_DESC  = None  # se detecta dinámicamente
    COL_VUNIT = None
    COL_CANT  = None

    for row in ws.iter_rows(min_row=1, values_only=True):
        vals = [v for v in row if v is not None]
        if not vals: continue

        primera = str(vals[0]).upper().strip()

        # Detectar inicio de sección
        if any(x in primera for x in ('INSUMO','MATERIAL')):
            seccion_actual = 'materiales'
            COL_DESC = COL_VUNIT = COL_CANT = None
            continue
        if any(x in primera for x in ('EQUIPO','HERRAMIENTA','TRANSPORTE')):
            seccion_actual = 'herramientas'
            COL_DESC = COL_VUNIT = COL_CANT = None
            continue
        if any(x in primera for x in ('MANO DE OB','MANO OB')):
            seccion_actual = 'mano_de_obra'
            COL_DESC = COL_VUNIT = COL_CANT = None
            continue

        # Detectar fila de encabezados de la sección
        row_list = list(row)
        row_upper = [str(v).upper().strip() if v else '' for v in row_list]
        if 'V. UNITARIO' in row_upper or 'V.UNITARIO' in row_upper or 'UNITARIO' in ' '.join(row_upper):
            for j, t in enumerate(row_upper):
                if 'DESCRIPCI' in t: COL_DESC = j
                if 'V. UNIT' in t or 'UNITARIO' in t: COL_VUNIT = j
                if 'CANTIDAD' in t or 'CANT' in t: COL_CANT = j
            continue

        # Detectar total costo directo
        if 'TOTAL COSTO' in primera or 'TOTAL DIRECTO' in primera:
            for v in row:
                if isinstance(v, (int, float)) and v > 0:
                    apu['total_referencia'] = float(v)
                    break
            continue

        # Fila de componente: debe tener número, descripción y valor
        if seccion_actual is None: continue
        if COL_VUNIT is None: continue

        # Detectar si es fila de datos: tiene descripción + V.UNITARIO + CANTIDAD válidos
        # No exigimos número de ítem — algunos archivos lo omiten o usan texto
        if COL_VUNIT is None or COL_CANT is None: continue
        vunit_check = row_list[COL_VUNIT] if COL_VUNIT < len(row_list) else None
        cant_check  = row_list[COL_CANT]  if COL_CANT  < len(row_list) else None
        if not isinstance(vunit_check, (int,float)) or vunit_check <= 0: continue
        if not isinstance(cant_check,  (int,float)) or cant_check  <= 0: continue

        desc   = str(row_list[COL_DESC]).strip()  if COL_DESC  is not None and COL_DESC  < len(row_list) and row_list[COL_DESC]  else ''
        vunit  = row_list[COL_VUNIT] if COL_VUNIT is not None and COL_VUNIT < len(row_list) else None
        cant   = row_list[COL_CANT]  if COL_CANT  is not None and COL_CANT  < len(row_list) else None
        und    = ''
        # Buscar unidad (suele ser la columna entre descripción y v.unitario)
        if COL_DESC is not None and COL_VUNIT is not None:
            for j in range(COL_DESC+1, COL_VUNIT):
                if j < len(row_list) and row_list[j] and isinstance(row_list[j], str):
                    und = str(row_list[j]).strip()
                    break

        if not desc or not isinstance(vunit, (int, float)) or not isinstance(cant, (int, float)):
            continue
        if vunit <= 0 or cant <= 0: continue

        comp = {'description': desc, 'unit': und,
                'rend': float(cant), 'unit_price': float(vunit)}
        apu[seccion_actual].append(comp)

    return apu

def _leer_hojas_apu_individuales(wb):
    """
    Lee todas las hojas que empiezan con 'APU' en un archivo de referencia.
    Retorna dict {codigo → apu_dict}.
    """
    bd = {}
    for nombre in wb.sheetnames:
        nombre_limpio = nombre.strip()
        if not nombre_limpio.upper().startswith('APU'):
            continue
        # Extraer código del nombre: "APU 2.16" → "2.16"
        partes = nombre_limpio.split()
        if len(partes) >= 2:
            codigo_hoja = partes[-1].strip()
        else:
            continue

        ws = wb[nombre]
        apu = _leer_apu_hoja_individual(ws)

        # También intentar extraer código desde la celda F6 col A
        try:
            codigo_celda = list(ws.iter_rows(min_row=6, max_row=6, values_only=True))[0][0]
            if codigo_celda and isinstance(codigo_celda, float):
                codigo_hoja = f"{codigo_celda:.2f}".rstrip('0').rstrip('.')
        except Exception:
            pass

        if apu['total_referencia'] > 0 or any(apu[s] for s in ('materiales','herramientas','mano_de_obra')):
            bd[codigo_hoja] = apu

    return bd


# ══════════════════════════════════════════════════════════════════
# LECTOR DE BASE DE DATOS APU (archivo de referencia)
# ══════════════════════════════════════════════════════════════════

def _leer_hoja_apu_columnar(ws_apu):
    """Lee hoja APU en formato columnar estándar (CODINS='-')."""
    bd = {}
    current = None
    primera = [c.value for c in ws_apu[1]]
    fila_ini = 2 if any(v and 'CODIGO' in str(v).upper() for v in primera if v) else 1
    for row in ws_apu.iter_rows(min_row=fila_ini, values_only=True):
        if not row[0]: continue
        code   = str(row[0]).strip()
        codins = str(row[1]).strip() if row[1] is not None else ''
        insumo = str(row[2]).strip() if row[2] else ''
        tipo   = str(row[3]).strip() if row[3] else ''
        unit   = str(row[4]).strip() if row[4] else ''
        rend   = row[5]; uprice = row[6]; total = row[8]
        if not tipo: continue
        if codins == '-':
            current = code
            bd[code] = {'total_referencia': float(total) if total else 0,
                        'materiales':[],'herramientas':[],'transporte':[],'mano_de_obra':[]}
        elif current and code == current and rend is not None and uprice is not None:
            comp = {'description': insumo, 'unit': unit,
                    'rend': float(rend), 'unit_price': float(uprice)}
            t = tipo.lower()
            if any(x in t for x in ('insumo','analisis basico','actividad','ensayo')):
                bd[current]['materiales'].append(comp)
            elif any(x in t for x in ('herramienta','equipo')):
                bd[current]['herramientas'].append(comp)
            elif any(x in t for x in ('cuadrilla','personal')):
                bd[current]['mano_de_obra'].append(comp)
            elif 'transporte' in t:
                bd[current]['transporte'].append(comp)
            else:
                bd[current]['materiales'].append(comp)
    return bd

def _construir_mapa_codigos(wb):
    """
    Lee la hoja PRESUPUESTO (o similar) del archivo de referencia
    y construye un mapeo {cod_gobernacion → item_num_interno}.
    Sirve para cruzar la oferta (usa cod gober) con los APUs (usan num interno).
    """
    mapa = {}  # cod_gober → item_num  (ej: "2.01.34" → "2.01")
    ws = None
    for nombre in wb.sheetnames:
        if 'PRESUPUESTO' in nombre.upper() or 'RESUMEN' in nombre.upper():
            ws = wb[nombre]; break
    if not ws: return mapa

    _, col_mapa = _detectar_columnas(ws)
    # Buscar las dos primeras columnas numéricas
    # En CIMM: col A = ítem interno, col B = código gobernación
    for row in ws.iter_rows(min_row=1, max_row=200, values_only=True):
        item_num = row[0]
        cod_gob  = row[1]
        if (item_num and isinstance(item_num, float) and item_num > 0
                and cod_gob and isinstance(cod_gob, str) and '.' in cod_gob):
            num_str = f"{item_num:.2f}".rstrip('0').rstrip('.')
            mapa[cod_gob.strip()] = num_str
    return mapa


def _leer_apu_formato_presupuesto_directo(ws):
    """
    Lee APUs del formato de presupuesto interno (Pascual Bravo / similar).
    Todos los APUs en una sola hoja con patrón:
      Señal: 'ANALISIS DE PRECIOS UNITARIOS - APU'
      Encabezado: col0='ITEM' col5='UNIDAD' col6='COSTO DIRECTO'
      Ítem:       col0=código, col1=desc, col5=unidad, col6=costo_directo
      Secciones:  col0=10000/20000/30000/40000/50000
      Componente: col0=cod_num, col1=desc, col2=und, col3=cant, col4=rend_o_desperd, col5=vr_unit, col6=vr_total
      Fin APU:    col1='TOTAL UNITARIO'
    """
    SECCIONES = {
        10000: 'materiales', 20000: 'mano_de_obra',
        30000: 'herramientas', 40000: 'materiales', 50000: 'transporte',
    }
    COD_SEC = set(SECCIONES.keys())

    apus = {}
    current_apu = None
    estado = 'buscando'
    seccion = None
    es_mo = False

    for row in ws.iter_rows(values_only=True):
        c = list(row) + [None] * 11
        c0,c1,c2,c3,c4,c5,c6 = c[0],c[1],c[2],c[3],c[4],c[5],c[6]

        if c0 and str(c0).strip() == 'ANALISIS DE PRECIOS UNITARIOS - APU':
            estado = 'enc_item'; continue

        if estado == 'enc_item' and c0 and str(c0).strip() == 'ITEM':
            estado = 'item'; continue

        if estado == 'item':
            estado = 'leyendo'; seccion = None
            if c0 and c1 and c6 is not None and isinstance(c6,(int,float)) and c6 > 0:
                code = str(c0).strip()
                current_apu = {
                    'code': code, 'description': str(c1).strip(),
                    'unit': str(c5).strip() if c5 else '',
                    'total_referencia': float(c6),
                    'materiales':[],'herramientas':[],'transporte':[],'mano_de_obra':[],
                }
                apus[code] = current_apu
            else:
                current_apu = None
            continue

        if estado != 'leyendo' or current_apu is None: continue

        if c1 and 'TOTAL UNITARIO' in str(c1).upper():
            current_apu = None; seccion = None; estado = 'buscando'; continue

        if isinstance(c0, int) and c0 in COD_SEC:
            seccion = SECCIONES[c0]; es_mo = (c0 == 20000); continue

        if c0 and str(c0).strip() in ('COD','#N/A'): continue
        if seccion is None: continue

        if (isinstance(c0, int) and c0 not in COD_SEC and
                c1 and isinstance(c1, str) and len(c1.strip()) > 2 and
                c5 is not None and isinstance(c5,(int,float)) and c5 > 0 and
                c6 is not None and isinstance(c6,(int,float)) and c6 > 0):
            # Derivar rendimiento desde VR.TOTAL / VR.UNITARIO
            # Esto captura desperdicio (materiales), CANT2 (equipos) y jornales/rendimiento (MO)
            vr_total = float(c6)
            vr_unit  = float(c5)
            rend = vr_total / vr_unit if vr_unit > 0 else 0
            if rend > 0:
                current_apu[seccion].append({
                    'description': str(c1).strip(), 'unit': str(c2).strip() if c2 else '',
                    'rend': round(rend, 6), 'unit_price': vr_unit,
                    'parcial': vr_total,
                })
    return apus

def leer_base_datos_apu(archivo_referencia):
    """
    Lee un archivo de referencia e intenta extraer APUs por cualquier método.
    Soporta: hoja A.P.U columnar, hojas individuales "APU X.XX", y PRESUPUESTO.
    """
    try:
        if hasattr(archivo_referencia, 'read'):
            data = archivo_referencia.read()
            archivo_referencia = io.BytesIO(data)
        wb = load_workbook(archivo_referencia, data_only=True)
    except Exception as e:
        return {}, f"No se pudo abrir el archivo de referencia: {e}"

    bd = {}

    # Método 1: hojas individuales "APU X.XX"
    bd_ind = _leer_hojas_apu_individuales(wb)
    if bd_ind:
        bd.update(bd_ind)

    # Método 2: hoja columnar A.P.U / APU
    for nombre in wb.sheetnames:
        if 'A.P.U' in nombre.upper() or nombre.upper() == 'APU':
            bd_col = _leer_hoja_apu_columnar(wb[nombre])
            bd.update(bd_col)
            break

    # Método 3: formato presupuesto directo (hoja con múltiples APUs, señal textual)
    # Buscar hojas con nombre APU, APU Obra Civil, APU Extraccion, APUS ELEC, etc.
    for nombre in wb.sheetnames:
        n_up = nombre.upper().strip()
        if n_up.startswith('APU') or 'APU' in n_up:
            bd_pd = _leer_apu_formato_presupuesto_directo(wb[nombre])
            if bd_pd:
                bd.update(bd_pd)

    # Método 4: construir puente entre código gobernación y número interno
    # (necesario cuando el archivo de referencia tiene APUs por número interno
    #  pero la oferta usa códigos de gobernación)
    mapa_codigos = _construir_mapa_codigos(wb)
    for cod_gob, num_int in mapa_codigos.items():
        if num_int in bd and cod_gob not in bd:
            bd[cod_gob] = bd[num_int]  # registrar bajo ambas claves

    return bd, None if bd else "No se encontraron APUs en el archivo de referencia."


# ══════════════════════════════════════════════════════════════════
# LECTURA PRINCIPAL DEL PROCESO
# ══════════════════════════════════════════════════════════════════

def leer_oferta_economica(uploaded_file, bd_referencia=None, año_proceso: int | None = None):
    try:
        data = uploaded_file.read()
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        return None, f"No se pudo abrir el archivo: {e}"

    # ── Determinar año del proceso (para validación de jornales mínimos) ──────
    año_proc, año_fuente = determinar_año_proceso(wb=wb, año_manual=año_proceso)
    jornal_min_legal = jornal_minimo(año_proc)

    ws_pres = _encontrar_hoja_presupuesto(wb)
    fila_enc, mapa = _detectar_columnas(ws_pres)

    if fila_enc is None or 'col_valor' not in mapa:
        return None, (
            f"No se encontró columna de valor unitario en la hoja '{ws_pres.title}'.\n"
            f"Hojas disponibles: {', '.join(wb.sheetnames)}\n\n"
            f"Se buscaron variantes de: VALOR UNITARIO, Vr. Unit., V. Unitario, P.U., etc."
        )

    col_valor = mapa['col_valor']
    col_desc  = mapa.get('col_desc')
    col_und   = mapa.get('col_und')

    items_proceso = {}
    for row in ws_pres.iter_rows(min_row=fila_enc + 2, values_only=True):
        if not _es_fila_item(row, mapa): continue
        codigo = _extraer_codigo(row, mapa)
        desc   = str(row[col_desc]).strip() if (col_desc is not None and col_desc < len(row) and row[col_desc]) else codigo
        und    = str(row[col_und]).strip()  if (col_und  is not None and col_und  < len(row) and row[col_und])  else ''
        valor  = float(row[col_valor])
        if codigo not in items_proceso:
            items_proceso[codigo] = {
                'code': codigo, 'description': desc, 'unit': und,
                'valor_ofrecido': valor,
                'materiales': [], 'herramientas': [],
                'transporte': [], 'mano_de_obra': [],
                'tiene_apu': False,
            }

    if not items_proceso:
        # Segundo intento: buscar otra hoja en el archivo
        hojas_restantes = [h for h in wb.sheetnames if h != ws_pres.title]
        for hoja_alt in hojas_restantes:
            ws_alt = wb[hoja_alt]
            fila_alt, mapa_alt = _detectar_columnas(ws_alt)
            if fila_alt is None or 'col_valor' not in mapa_alt:
                continue
            items_alt = {}
            col_v2 = mapa_alt['col_valor']
            col_d2 = mapa_alt.get('col_desc')
            col_u2 = mapa_alt.get('col_und')
            for row in ws_alt.iter_rows(min_row=fila_alt + 2, values_only=True):
                if not _es_fila_item(row, mapa_alt): continue
                cod2 = _extraer_codigo(row, mapa_alt)
                desc2 = str(row[col_d2]).strip() if (col_d2 is not None and col_d2 < len(row) and row[col_d2]) else cod2
                und2  = str(row[col_u2]).strip()  if (col_u2 is not None and col_u2 < len(row) and row[col_u2])  else ''
                val2  = float(row[col_v2])
                if cod2 not in items_alt:
                    items_alt[cod2] = {'code':cod2,'description':desc2,'unit':und2,
                                       'valor_ofrecido':val2,'materiales':[],'herramientas':[],
                                       'transporte':[],'mano_de_obra':[],'tiene_apu':False}
            if items_alt:
                items_proceso = items_alt
                ws_pres = ws_alt
                mapa = mapa_alt
                break

    if not items_proceso:
        return None, (
            f"No se encontraron ítems con valor unitario en '{ws_pres.title}'.\n"
            f"Encabezados detectados en fila {fila_enc+1}: {mapa}"
        )

    # ── APU del mismo archivo (columnar o individual) ─────────────────────────
    bd_interna = {}
    bd_ind = _leer_hojas_apu_individuales(wb)
    if bd_ind:
        bd_interna.update(bd_ind)
    for nombre in wb.sheetnames:
        if 'A.P.U' in nombre.upper() or nombre.upper() == 'APU':
            bd_interna.update(_leer_hoja_apu_columnar(wb[nombre]))
            break

    # ── Combinar fuentes: interna + referencia externa ────────────────────────
    bd_ext = bd_referencia or {}
    bd_combinada = {**bd_ext, **bd_interna}  # interna tiene prioridad

    # ── Cruzar y aplicar algoritmo R2 ────────────────────────────────────────
    for code, item in items_proceso.items():
        apu = bd_combinada.get(code)
        if not apu:
            continue
        ref = apu['total_referencia']
        if ref <= 0:
            continue

        precio_ofertado = item['valor_ofrecido']

        # ── Validación R: precio ofertado NO debe superar APU entidad ─────────
        if precio_ofertado > ref + 1:   # tolerancia $1 por redondeo
            item['alerta_supera_entidad'] = True
            item['apu_entidad_total']     = ref

        # ── Componentes FIJOS: VU y REND inalterados ─────────────────────────
        # Materiales, herramientas/equipos y transporte: valores exactos del APU oficial
        for sec in ('materiales', 'herramientas', 'transporte'):
            item[sec] = [
                {**c, 'parcial': round(c['rend'] * c['unit_price'], 2)}
                for c in apu.get(sec, [])
            ]

        subtotal_mat   = sum(c['rend'] * c['unit_price'] for c in item['materiales'])
        subtotal_herr  = sum(c['rend'] * c['unit_price'] for c in item['herramientas'])
        subtotal_trans = sum(c['rend'] * c['unit_price'] for c in item['transporte'])
        fijos_total    = subtotal_mat + subtotal_herr + subtotal_trans

        # ── Labor residual: lo que queda para mano de obra ───────────────────
        labor_residual = precio_ofertado - fijos_total

        if labor_residual <= 0:
            # R3: zona negativa — costos fijos ya superan el precio ofertado
            item['alerta_zona_negativa'] = True
            item['labor_residual']       = labor_residual
            # Copiar MO de referencia sin escalar (estructura visible para el usuario)
            item['mano_de_obra'] = [
                {**c, 'parcial': round(c['rend'] * c['unit_price'], 2)}
                for c in apu.get('mano_de_obra', [])
            ]
        else:
            # R2: calcular nuevo rendimiento para MO — jornales FIJOS, rend varía
            mo_ref = apu.get('mano_de_obra', [])
            suma_mo_original = sum(c['rend'] * c['unit_price'] for c in mo_ref)

            if mo_ref and suma_mo_original > 0:
                # factor = labor_residual / coste_MO_original → escala solo el rend
                factor_mo = labor_residual / suma_mo_original

                # Señales de alerta por rendimiento (según prompt maestro)
                if factor_mo > 5.0:
                    item['alerta_rendimiento_alto'] = True
                    item['factor_mo'] = round(factor_mo, 4)
                elif factor_mo < 0.3:
                    item['alerta_rendimiento_bajo'] = True
                    item['factor_mo'] = round(factor_mo, 4)

                # Validar que ningún jornal quede por debajo del mínimo legal
                mo_nueva = []
                for c in mo_ref:
                    jornal_actual = c['unit_price']
                    if jornal_actual < jornal_min_legal:
                        # Ajustar jornal al mínimo y recalcular rend
                        jornal_actual = jornal_min_legal
                        item['alerta_jornal_bajo'] = True
                    rend_nuevo = round(c['rend'] * factor_mo, 6)
                    mo_nueva.append({
                        **c,
                        'unit_price': jornal_actual,
                        'rend':       rend_nuevo,
                        'parcial':    round(rend_nuevo * jornal_actual, 2),
                    })
                item['mano_de_obra'] = mo_nueva
            else:
                # Sin MO en referencia: crear componente genérico con jornal mínimo
                item['mano_de_obra'] = [{
                    'description': 'Cuadrilla de obra (operario + ayudante)',
                    'unit':        'DIA',
                    'rend':        round(labor_residual / jornal_min_legal, 6),
                    'unit_price':  jornal_min_legal,
                    'parcial':     round(labor_residual, 2),
                }]
            item['labor_residual'] = labor_residual

        item['tiene_apu'] = True

    con_apu = [i for i in items_proceso.values() if i['tiene_apu']]
    sin_apu = [i for i in items_proceso.values() if not i['tiene_apu']]

    # ── Estadísticas de validación ────────────────────────────────────────────
    supera_entidad  = [i for i in con_apu if i.get('alerta_supera_entidad')]
    zona_negativa   = [i for i in con_apu if i.get('alerta_zona_negativa')]
    rend_alto       = [i for i in con_apu if i.get('alerta_rendimiento_alto')]
    jornal_bajo     = [i for i in con_apu if i.get('alerta_jornal_bajo')]

    return {
        'items_con_apu':       con_apu,
        'items_sin_apu':       sin_apu,
        'total_proceso':       len(items_proceso),
        'tiene_hoja_apu':      bool(bd_interna),
        'hoja_usada':          ws_pres.title,
        'año_proceso':         año_proc,
        'año_fuente':          año_fuente,
        'jornal_min_legal':    jornal_min_legal,
        'alertas': {
            'supera_entidad':  supera_entidad,
            'zona_negativa':   zona_negativa,
            'rendimiento_alto': rend_alto,
            'jornal_bajo':     jornal_bajo,
        },
    }, None


# ══════════════════════════════════════════════════════════════════
# GENERACIÓN DEL EXCEL
# ══════════════════════════════════════════════════════════════════

def _safe_name(code):
    name = str(code)[:31]
    for c in '/\\?*:[]': name = name.replace(c, '-')
    return name

def _build_apu_sheet(ws, item, include_aiu=False, aiu_pct=0.0):
    bold9  = Font(bold=True, name='Arial', size=9)
    norm9  = Font(bold=False, name='Arial', size=9)
    thin   = Side(style='thin')
    brd    = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_w = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    right  = Alignment(horizontal='right',  vertical='center')
    gray   = PatternFill('solid', fgColor='D9D9D9')
    blue   = PatternFill('solid', fgColor='BDD7EE')
    CUR    = '$ #,##0'

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 4
    ws.column_dimensions['C'].width = 4
    ws.column_dimensions['D'].width = 4
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 15

    r = 1
    ws.row_dimensions[r].height = 35
    ws.merge_cells(f'A{r}:C{r}')
    for col, val, al in [
        ('A','DESCRIPCIÓN',center), ('D',item['description'],left_w),
        ('E','UNIDAD',center),      ('F',item['unit'],center),
        ('G','ITEM',center),        ('H',item['code'],center),
    ]:
        ws[f'{col}{r}'] = val
        ws[f'{col}{r}'].font = bold9
        ws[f'{col}{r}'].alignment = al
        ws[f'{col}{r}'].border = brd
    r += 1

    def write_section(titulo, col_header, componentes):
        nonlocal r
        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:H{r}')
        ws[f'A{r}'] = titulo
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = center
        ws[f'A{r}'].fill = gray;  ws[f'A{r}'].border = brd
        r += 1
        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:D{r}')
        for col, val in [('A','DESCRIPCION'),('E','UNIDAD'),
                         ('F',col_header),('G','VR UNIT'),('H','VR TOTAL')]:
            ws[f'{col}{r}'] = val
            ws[f'{col}{r}'].font = bold9
            ws[f'{col}{r}'].alignment = center
            ws[f'{col}{r}'].border = brd
        r += 1
        start = r
        n = max(len(componentes), 2)
        for i in range(n):
            ws.row_dimensions[r].height = 13
            comp = componentes[i] if i < len(componentes) else None
            ws.merge_cells(f'A{r}:D{r}')
            if comp:
                ws[f'A{r}'] = comp['description']
                ws[f'E{r}'] = comp['unit']
                ws[f'F{r}'] = comp['rend']
                ws[f'G{r}'] = comp['unit_price']
                ws[f'H{r}'] = f'=ROUND(F{r}*G{r},0)'
            for col, al, fmt in [
                ('A',left_w,None),('E',center,None),
                ('F',right,'#,##0.00'),('G',right,CUR),('H',right,CUR)]:
                ws[f'{col}{r}'].font = norm9
                ws[f'{col}{r}'].border = brd
                ws[f'{col}{r}'].alignment = al
                if fmt: ws[f'{col}{r}'].number_format = fmt
            r += 1
        end = r - 1
        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:G{r}')
        ws[f'A{r}'] = f'SUBTOTAL {titulo}'
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = right
        ws[f'A{r}'].fill = gray;  ws[f'A{r}'].border = brd
        ws[f'H{r}'] = f'=SUM(H{start}:H{end})'
        ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
        ws[f'H{r}'].fill = gray;  ws[f'H{r}'].border = brd
        ws[f'H{r}'].number_format = CUR
        sub = r; r += 1
        return sub

    row_mat = write_section('MATERIALES',               'CANTID.', item['materiales'])
    row_her = write_section('HERRAMIENTAS Y/O EQUIPOS', 'RENDIM.', item['herramientas'])
    row_tra = write_section('TRANSPORTE',               'CANTID.', item['transporte'])
    row_mdo = write_section('MANO DE OBRA',             'RENDIM.', item['mano_de_obra'])

    aiu_row = None
    if include_aiu and aiu_pct > 0:
        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:D{r}')
        ws[f'A{r}'] = 'A.I.U.'
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = left_w; ws[f'A{r}'].border = brd
        ws[f'E{r}'] = '%'
        ws[f'E{r}'].font = bold9; ws[f'E{r}'].alignment = center; ws[f'E{r}'].border = brd
        ws[f'F{r}'] = aiu_pct
        ws[f'F{r}'].font = bold9; ws[f'F{r}'].alignment = right
        ws[f'F{r}'].border = brd; ws[f'F{r}'].number_format = '0.00%'
        ws[f'G{r}'].border = brd
        ws[f'H{r}'] = f'=ROUND((H{row_mat}+H{row_her}+H{row_tra}+H{row_mdo})*F{r},0)'
        ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
        ws[f'H{r}'].border = brd; ws[f'H{r}'].number_format = CUR
        aiu_row = r; r += 1

    # ── PRECIO UNITARIO ─────────────────────────────────────────────────────
    ws.row_dimensions[r].height = 18
    ws.merge_cells(f'A{r}:E{r}')
    ws[f'A{r}'] = 'PRECIO UNITARIO'
    ws[f'A{r}'].font = Font(bold=True, name='Arial', size=10)
    ws[f'A{r}'].alignment = center; ws[f'A{r}'].fill = blue; ws[f'A{r}'].border = brd
    ws.merge_cells(f'F{r}:H{r}')
    valor_precio = int(round(item['valor_ofrecido'], 0))
    ws[f'F{r}'] = valor_precio
    ws[f'F{r}'].font = Font(bold=True, name='Arial', size=10)
    ws[f'F{r}'].alignment = right; ws[f'F{r}'].fill = blue
    ws[f'F{r}'].border = brd; ws[f'F{r}'].number_format = CUR
    r += 1

    # ── TOTAL EN LETRAS ───────────────────────────────────────────────────
    gold = PatternFill('solid', fgColor='FFF2CC')
    ws.row_dimensions[r].height = 22
    ws.merge_cells(f'A{r}:H{r}')
    letras = numero_a_letras(valor_precio)
    ws[f'A{r}'] = letras
    ws[f'A{r}'].font = Font(bold=True, name='Arial', size=9, italic=True)
    ws[f'A{r}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws[f'A{r}'].fill = gold
    ws[f'A{r}'].border = brd


def _build_resumen_sheet(wb, items_generados):
    """
    Crea la hoja RESUMEN (primera hoja del workbook) con tabla de todos los APUs:
    Ítem | Descripción | Unidad | Precio ofrecido | Precio APU | Cierre | N° comp | Fuente
    """
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    ws = wb.create_sheet(title='RESUMEN', index=0)

    # Estilos
    bold   = Font(bold=True, name='Arial', size=9)
    norm   = Font(bold=False, name='Arial', size=9)
    thin   = Side(style='thin')
    brd    = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left   = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    right  = Alignment(horizontal='right',  vertical='center')
    azul   = PatternFill('solid', fgColor='1B3A6B')
    gris   = PatternFill('solid', fgColor='D9D9D9')
    verde  = PatternFill('solid', fgColor='E2EFDA')
    rojo   = PatternFill('solid', fgColor='FCE4D6')
    ambar  = PatternFill('solid', fgColor='FFF2CC')
    blanco = Font(bold=True, name='Arial', size=9, color='FFFFFF')
    CUR    = '$ #,##0'

    # Anchos de columna
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 10
    ws.column_dimensions['H'].width = 35

    # Título
    ws.row_dimensions[1].height = 28
    ws.merge_cells('A1:H1')
    ws['A1'] = 'RESUMEN DE APUs GENERADOS — Gerencia Legal Integral Colombia S.A.S.'
    ws['A1'].font = blanco
    ws['A1'].fill = azul
    ws['A1'].alignment = center

    # Encabezados
    ws.row_dimensions[2].height = 30
    encabezados = [
        ('A2', 'ÍTEM'),
        ('B2', 'DESCRIPCIÓN'),
        ('C2', 'UND.'),
        ('D2', 'PRECIO OFRECIDO'),
        ('E2', 'PRECIO APU'),
        ('F2', 'CIERRE'),
        ('G2', 'COMP.'),
        ('H2', 'FUENTE'),
    ]
    for celda, texto in encabezados:
        ws[celda] = texto
        ws[celda].font = bold
        ws[celda].fill = gris
        ws[celda].alignment = center
        ws[celda].border = brd

    # Datos
    fila = 3
    total_ofrecido = 0
    total_apu = 0

    for item in items_generados:
        ws.row_dimensions[fila].height = 14

        n_comp = sum(len(item.get(s, [])) for s in
                     ('materiales', 'herramientas', 'transporte', 'mano_de_obra'))

        # Calcular precio APU desde componentes
        precio_ofrecido = item.get('valor_ofrecido', 0)
        precio_apu_comp = sum(
            c['rend'] * c['unit_price']
            for s in ('materiales', 'herramientas', 'transporte', 'mano_de_obra')
            for c in item.get(s, [])
        )
        # El PRECIO UNITARIO del APU es siempre el hardcodeado = valor_ofrecido
        precio_apu = precio_ofrecido
        diferencia = abs(precio_ofrecido - precio_apu_comp)
        cierra = diferencia < 2

        # Fuente
        fuente = item.get('fuente_bd', '')
        if not fuente:
            fuente = 'Misma entidad (APU referencia)'
        if n_comp == 0:
            fuente = 'Sin componentes'

        fila_fill = verde if cierra and n_comp > 0 else (rojo if n_comp == 0 else ambar)

        valores = [
            ('A', item.get('code', ''), center),
            ('B', item.get('description', ''), left),
            ('C', item.get('unit', ''), center),
            ('D', precio_ofrecido, right),
            ('E', precio_apu, right),
            ('F', '✅' if cierra and n_comp > 0 else ('⚠️' if n_comp > 0 else '—'), center),
            ('G', n_comp, center),
            ('H', fuente, left),
        ]

        for col, val, aln in valores:
            c = ws[f'{col}{fila}']
            c.value = val
            c.font = norm
            c.alignment = aln
            c.border = brd
            c.fill = fila_fill
            if col in ('D', 'E'):
                c.number_format = CUR

        total_ofrecido += precio_ofrecido
        total_apu += precio_apu
        fila += 1

    # Fila de totales
    ws.row_dimensions[fila].height = 16
    ws.merge_cells(f'A{fila}:C{fila}')
    ws[f'A{fila}'] = f'TOTAL — {fila - 3} ítems APU'
    ws[f'A{fila}'].font = bold
    ws[f'A{fila}'].fill = gris
    ws[f'A{fila}'].alignment = right
    ws[f'A{fila}'].border = brd

    for col, val in [('D', total_ofrecido), ('E', total_apu)]:
        ws[f'{col}{fila}'] = val
        ws[f'{col}{fila}'].font = bold
        ws[f'{col}{fila}'].fill = gris
        ws[f'{col}{fila}'].alignment = right
        ws[f'{col}{fila}'].border = brd
        ws[f'{col}{fila}'].number_format = CUR

    for col in ('F', 'G', 'H'):
        ws[f'{col}{fila}'].border = brd
        ws[f'{col}{fila}'].fill = gris

    # Leyenda de fuentes
    fila += 2
    leyenda = [
        ('', verde,  'Verde — APU completo con componentes desde archivo de referencia (misma entidad)'),
        ('', ambar,  'Amarillo — APU con componentes desde base de precios externa (Gobernación, INVIAS, etc.)'),
        ('', rojo,   'Rojo — Sin componentes (ítems pendientes de completar)'),
    ]
    for _, fill, texto in leyenda:
        ws.row_dimensions[fila].height = 14
        ws.merge_cells(f'A{fila}:H{fila}')
        ws[f'A{fila}'] = texto
        ws[f'A{fila}'].font = Font(italic=True, name='Arial', size=8)
        ws[f'A{fila}'].fill = fill
        ws[f'A{fila}'].alignment = left
        fila += 1

def generate_apu_excel(resultado, items_manuales=None, include_aiu=False,
                       aiu_pct=0.0, bd_externas=None):
    """
    Genera el Excel final.
    bd_externas: lista de dicts cargados desde bases_externas.cargar_base_externa()
    """
    wb = Workbook(); wb.remove(wb.active)

    # Construir lista completa de ítems
    todos = list(resultado['items_con_apu']) + (items_manuales or [])

    # Para ítems sin APU, intentar base externa antes de dejar vacío
    sin_apu = resultado['items_sin_apu']
    items_sin_procesar = [i for i in sin_apu
                          if i['code'] not in {x['code'] for x in todos}]

    if bd_externas and items_sin_procesar:
        for item in items_sin_procesar:
            mejor = None
            mejor_pct = 0
            for bd in bd_externas:
                resultados = buscar_en_base(
                    item['description'], item['unit'], bd, top_n=1, umbral_pct=25
                )
                if resultados and resultados[0][1] > mejor_pct:
                    mejor_pct = resultados[0][1]
                    mejor = resultados[0][2]
            if mejor:
                apu_ext = construir_apu_desde_base(mejor, item['valor_ofrecido'])
                item['materiales']   = apu_ext['materiales']
                item['herramientas'] = apu_ext['herramientas']
                item['transporte']   = apu_ext['transporte']
                item['mano_de_obra'] = apu_ext['mano_de_obra']
                item['fuente_bd']    = apu_ext['fuente_bd']
                item['codigo_bd']    = apu_ext['codigo_bd']
            todos.append(item)
    else:
        todos += items_sin_procesar

    if not todos:
        todos = sin_apu

    # Marcar fuente de los ítems con APU de referencia
    for item in resultado['items_con_apu']:
        if 'fuente_bd' not in item:
            item['fuente_bd'] = ''  # vacío = misma entidad

    # Generar hojas de APU
    for item in todos:
        name = _safe_name(item['code'])
        if name in [s.title for s in wb.worksheets]: name = name[:27] + '_b'
        ws = wb.create_sheet(title=name)
        _build_apu_sheet(ws, item, include_aiu, aiu_pct)

    # Hoja RESUMEN (se inserta al inicio)
    _build_resumen_sheet(wb, todos)

    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf.getvalue()
