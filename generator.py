"""
Generador de APUs — Gerencia Legal Integral Colombia S.A.S.
v10: Flujo rediseñado.
  - leer_apu_entidad()        → lee APUs del archivo de la entidad (componentes completos)
  - leer_propuesta_economica() → lee precios del proponente
  - cruzar_y_ajustar()        → cruza por código (luego descripción) y ajusta MO
  - generate_apu_excel()       → genera el Excel final con hojas de APU + RESUMEN

REGLAS DE AJUSTE (Reglas Generales):
  1. Mantener materiales, equipos, herramientas y transporte SIN CAMBIO.
  2. Ajustar ÚNICAMENTE el rendimiento de la mano de obra.
  3. Si con MO se generan valores negativos o es imposible el cierre exacto,
     ajustar rendimiento de equipos/herramientas como último recurso.
  4. Precio unitario final = exactamente el precio ofrecido (2 decimales).
  5. Ningún componente puede quedar con valor negativo.
"""
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io, re, unicodedata

from numero_letras import numero_a_letras
from detector import (
    detectar_columnas          as _detectar_columnas,
    es_fila_item_valida        as _es_fila_item,
    extraer_codigo             as _extraer_codigo,
    encontrar_hoja_presupuesto as _encontrar_hoja_presupuesto,
)


# ══════════════════════════════════════════════════════════════════════════════
# UTILIDADES
# ══════════════════════════════════════════════════════════════════════════════

def _normalizar(texto):
    """Minúsculas, sin tildes, sin caracteres especiales — para comparación."""
    t = str(texto).lower().strip()
    t = unicodedata.normalize('NFD', t)
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')
    t = re.sub(r'[^a-z0-9\s]', ' ', t)
    t = re.sub(r'\s+', ' ', t).strip()
    return t


def _similitud_tokens(a, b):
    """Porcentaje de tokens comunes / máximo tokens. Simple pero robusto."""
    ta = set(_normalizar(a).split())
    tb = set(_normalizar(b).split())
    if not ta or not tb:
        return 0.0
    comunes = len(ta & tb)
    return comunes / max(len(ta), len(tb))


def _safe_name(code):
    name = str(code)[:31]
    for c in '/\\?*:[]': name = name.replace(c, '-')
    return name


# ══════════════════════════════════════════════════════════════════════════════
# LECTOR DE APUs DE LA ENTIDAD
# ══════════════════════════════════════════════════════════════════════════════

def _leer_apu_hoja_individual(ws):
    """
    Lee un APU desde hoja individual (formato SENA/Gobernación).
    Retorna dict con keys: total_referencia, materiales, herramientas,
                           transporte, mano_de_obra, description, unit.
    """
    apu = {
        'description': '', 'unit': '', 'total_referencia': 0,
        'materiales': [], 'herramientas': [], 'transporte': [], 'mano_de_obra': []
    }
    seccion = None
    COL_DESC = COL_VUNIT = COL_CANT = None

    for row in ws.iter_rows(min_row=1, values_only=True):
        vals = [v for v in row if v is not None]
        if not vals:
            continue
        primera = str(vals[0]).upper().strip()

        # Detectar sección
        if any(x in primera for x in ('INSUMO', 'MATERIAL')):
            seccion = 'materiales'; COL_DESC = COL_VUNIT = COL_CANT = None; continue
        if any(x in primera for x in ('EQUIPO', 'HERRAMIENTA', 'TRANSPORTE')):
            seccion = 'herramientas'; COL_DESC = COL_VUNIT = COL_CANT = None; continue
        if any(x in primera for x in ('MANO DE OB', 'MANO OB')):
            seccion = 'mano_de_obra'; COL_DESC = COL_VUNIT = COL_CANT = None; continue

        # Detectar encabezados de columna
        row_list = list(row)
        row_upper = [str(v).upper().strip() if v else '' for v in row_list]
        if 'UNITARIO' in ' '.join(row_upper):
            for j, t in enumerate(row_upper):
                if 'DESCRIPCI' in t: COL_DESC = j
                if 'V. UNIT' in t or ('UNITARIO' in t and 'V' in t): COL_VUNIT = j
                if 'CANTIDAD' in t or 'CANT' in t: COL_CANT = j
            continue

        # Total costo directo
        if 'TOTAL COSTO' in primera or 'TOTAL DIRECTO' in primera:
            for v in row:
                if isinstance(v, (int, float)) and v > 0:
                    apu['total_referencia'] = float(v); break
            continue

        if seccion is None or COL_VUNIT is None or COL_CANT is None:
            continue

        vunit_v = row_list[COL_VUNIT] if COL_VUNIT < len(row_list) else None
        cant_v  = row_list[COL_CANT]  if COL_CANT  < len(row_list) else None
        if not isinstance(vunit_v, (int, float)) or vunit_v <= 0: continue
        if not isinstance(cant_v,  (int, float)) or cant_v  <= 0: continue

        desc = str(row_list[COL_DESC]).strip() if (COL_DESC is not None and COL_DESC < len(row_list) and row_list[COL_DESC]) else ''
        und  = ''
        if COL_DESC is not None and COL_VUNIT is not None:
            for j in range(COL_DESC + 1, COL_VUNIT):
                if j < len(row_list) and row_list[j] and isinstance(row_list[j], str):
                    und = str(row_list[j]).strip(); break

        if not desc: continue
        apu[seccion].append({
            'description': desc, 'unit': und,
            'rend': float(cant_v), 'unit_price': float(vunit_v),
        })

    return apu


def _leer_hojas_apu_individuales(wb):
    """Lee todas las hojas 'APU X.XX' de un workbook. Retorna {codigo: apu_dict}."""
    bd = {}
    for nombre in wb.sheetnames:
        n = nombre.strip()
        if not n.upper().startswith('APU'):
            continue
        partes = n.split()
        if len(partes) < 2:
            continue
        codigo_hoja = partes[-1].strip()
        ws  = wb[nombre]
        apu = _leer_apu_hoja_individual(ws)
        # Intentar extraer descripción desde celda de encabezado
        try:
            for row in ws.iter_rows(min_row=1, max_row=8, values_only=True):
                for v in row:
                    if v and isinstance(v, str) and len(v) > 8 and not any(
                        x in v.upper() for x in ('GERENCIA','GOBERN','ENTIDAD','S.A.S','FECHA','HOJA','ITEM')
                    ):
                        if not apu['description']:
                            apu['description'] = v.strip()
                        break
                if apu['description']:
                    break
        except Exception:
            pass
        if apu['total_referencia'] > 0 or any(apu[s] for s in ('materiales','herramientas','mano_de_obra')):
            bd[codigo_hoja] = apu
    return bd


def _leer_apu_columnar(ws_apu):
    """Lee hoja APU columnar estándar (CODINS='-')."""
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
            bd[code] = {
                'description': insumo, 'unit': unit,
                'total_referencia': float(total) if total else 0,
                'materiales':[], 'herramientas':[], 'transporte':[], 'mano_de_obra':[]
            }
        elif current and code == current and rend is not None and uprice is not None:
            comp = {'description': insumo, 'unit': unit,
                    'rend': float(rend), 'unit_price': float(uprice)}
            t = tipo.lower()
            if any(x in t for x in ('insumo','analisis','actividad','ensayo')):
                bd[current]['materiales'].append(comp)
            elif any(x in t for x in ('herramienta','equipo')):
                bd[current]['herramientas'].append(comp)
            elif any(x in t for x in ('cuadrilla','personal','mano')):
                bd[current]['mano_de_obra'].append(comp)
            elif 'transporte' in t:
                bd[current]['transporte'].append(comp)
            else:
                bd[current]['materiales'].append(comp)
    return bd


def _leer_apu_presupuesto_directo(ws):
    """
    Lee APUs embebidos en hoja de presupuesto (señal 'ANALISIS DE PRECIOS UNITARIOS - APU').
    """
    SECCIONES = {
        10000: 'materiales', 20000: 'mano_de_obra',
        30000: 'herramientas', 40000: 'materiales', 50000: 'transporte',
    }
    COD_SEC = set(SECCIONES.keys())
    apus = {}
    current = None
    estado  = 'buscando'
    seccion = None

    for row in ws.iter_rows(values_only=True):
        c = list(row) + [None] * 11
        c0, c1, c2, c5, c6 = c[0], c[1], c[2], c[5], c[6]

        if c0 and str(c0).strip() == 'ANALISIS DE PRECIOS UNITARIOS - APU':
            estado = 'enc_item'; continue
        if estado == 'enc_item' and c0 and str(c0).strip() == 'ITEM':
            estado = 'item'; continue
        if estado == 'item':
            estado = 'leyendo'; seccion = None
            if c0 and c1 and c6 is not None and isinstance(c6, (int, float)) and c6 > 0:
                code = str(c0).strip()
                current = {
                    'code': code, 'description': str(c1).strip(),
                    'unit': str(c5).strip() if c5 else '',
                    'total_referencia': float(c6),
                    'materiales':[], 'herramientas':[], 'transporte':[], 'mano_de_obra':[],
                }
                apus[code] = current
            else:
                current = None
            continue

        if estado != 'leyendo' or current is None: continue
        if c1 and 'TOTAL UNITARIO' in str(c1).upper():
            current = None; seccion = None; estado = 'buscando'; continue
        if isinstance(c0, int) and c0 in COD_SEC:
            seccion = SECCIONES[c0]; continue
        if c0 and str(c0).strip() in ('COD', '#N/A'): continue
        if seccion is None: continue

        if (isinstance(c0, int) and c0 not in COD_SEC and
                c1 and isinstance(c1, str) and len(c1.strip()) > 2 and
                c5 is not None and isinstance(c5, (int, float)) and c5 > 0 and
                c6 is not None and isinstance(c6, (int, float)) and c6 > 0):
            vr_unit = float(c5)
            rend    = float(c6) / vr_unit if vr_unit > 0 else 0
            if rend > 0:
                current[seccion].append({
                    'description': str(c1).strip(),
                    'unit':        str(c2).strip() if c2 else '',
                    'rend':        round(rend, 6),
                    'unit_price':  vr_unit,
                })
    return apus


def leer_apu_entidad(archivo):
    """
    Punto de entrada principal.
    Lee el archivo Excel de la entidad e intenta extraer APUs por todos los métodos:
      1. Hojas individuales 'APU X.XX'
      2. Hoja columnar 'A.P.U' / 'APU'
      3. Formato presupuesto directo con señal textual
    Retorna: (dict{codigo → apu_dict}, error_str | None)
    """
    try:
        if hasattr(archivo, 'read'):
            data = archivo.read()
            archivo = io.BytesIO(data)
        wb = load_workbook(archivo, data_only=True)
    except Exception as e:
        return {}, f"No se pudo abrir el archivo de la entidad: {e}"

    bd = {}

    # Método 1: hojas individuales
    bd.update(_leer_hojas_apu_individuales(wb))

    # Método 2: hoja columnar
    for nombre in wb.sheetnames:
        if 'A.P.U' in nombre.upper() or nombre.upper() == 'APU':
            bd.update(_leer_apu_columnar(wb[nombre]))
            break

    # Método 3: formato presupuesto directo
    for nombre in wb.sheetnames:
        n_up = nombre.upper().strip()
        if n_up.startswith('APU') or 'APU' in n_up:
            apus_pd = _leer_apu_presupuesto_directo(wb[nombre])
            if apus_pd:
                bd.update(apus_pd)

    if not bd:
        return {}, (
            "No se encontraron APUs en el archivo de la entidad. "
            "Verifique que el archivo contenga hojas con nombre 'APU X.XX', "
            "una hoja 'A.P.U' columnar, o APUs en formato de presupuesto directo."
        )
    return bd, None


# ══════════════════════════════════════════════════════════════════════════════
# LECTOR DE PROPUESTA ECONÓMICA
# ══════════════════════════════════════════════════════════════════════════════

def leer_propuesta_economica(archivo):
    """
    Lee la propuesta económica del proponente.
    Retorna: (list[dict], error_str | None)
    Cada dict: {code, description, unit, valor_ofrecido}
    """
    try:
        data = archivo.read()
        wb   = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        return [], f"No se pudo abrir la propuesta económica: {e}"

    ws = _encontrar_hoja_presupuesto(wb)
    fila_enc, mapa = _detectar_columnas(ws)

    if fila_enc is None or 'col_valor' not in mapa:
        # Intentar otras hojas
        for nombre in wb.sheetnames:
            if nombre == ws.title:
                continue
            ws_alt = wb[nombre]
            fila_alt, mapa_alt = _detectar_columnas(ws_alt)
            if fila_alt is not None and 'col_valor' in mapa_alt:
                ws = ws_alt; fila_enc = fila_alt; mapa = mapa_alt; break

    if fila_enc is None or 'col_valor' not in mapa:
        return [], (
            f"No se encontró columna de valor unitario en '{ws.title}'. "
            f"Hojas disponibles: {', '.join(wb.sheetnames)}"
        )

    col_valor = mapa['col_valor']
    col_desc  = mapa.get('col_desc')
    col_und   = mapa.get('col_und')

    items = []
    vistos = set()
    for row in ws.iter_rows(min_row=fila_enc + 2, values_only=True):
        if not _es_fila_item(row, mapa):
            continue
        codigo = _extraer_codigo(row, mapa)
        if not codigo or codigo in vistos:
            continue
        desc  = str(row[col_desc]).strip() if (col_desc is not None and col_desc < len(row) and row[col_desc]) else codigo
        und   = str(row[col_und]).strip()  if (col_und  is not None and col_und  < len(row) and row[col_und])  else ''
        valor = float(row[col_valor])
        if valor <= 0:
            continue
        items.append({'code': codigo, 'description': desc, 'unit': und, 'valor_ofrecido': valor})
        vistos.add(codigo)

    if not items:
        return [], (
            f"No se encontraron ítems con valor unitario en '{ws.title}'. "
            "Verifique el formato del archivo."
        )
    return items, None


# ══════════════════════════════════════════════════════════════════════════════
# MOTOR DE CRUCE Y AJUSTE
# ══════════════════════════════════════════════════════════════════════════════

def _rend_exacto(parcial_rd2, unit_price):
    """
    Calcula el rendimiento (10 decimales) tal que ROUND(rend * unit_price, 2) == parcial_rd2.
    Necesario para que la fórmula Excel =ROUND(F*G,2) cierre exactamente.
    """
    if unit_price <= 0:
        return 0.0
    base = parcial_rd2 / unit_price
    for delta in [0, 1e-10, -1e-10, 2e-10, -2e-10, 5e-10, -5e-10, 1e-9, -1e-9]:
        r = round(base + delta, 10)
        if round(r * unit_price, 2) == round(parcial_rd2, 2):
            return r
    return round(base, 10)


def _ajustar_rendimiento_mo(item, precio_objetivo):
    """
    Ajusta el rendimiento de la mano de obra para que el APU cierre exactamente
    con el precio_objetivo (a 2 decimales).

    REGLAS:
      1. Calcular suma fija = materiales + herramientas + transporte.
      2. Necesario_mo = precio_objetivo - suma_fija.
      3. Si necesario_mo < 0: no es posible con MO → intentar con equipos.
      4. Distribuir necesario_mo entre los componentes de MO proporcional a su peso original.
      5. Si MO está vacía, intentar ajustar herramientas.
      6. Si tampoco es viable: registrar advertencia.

    Retorna: item ajustado (in-place) + dict de resultado.
    """
    import copy

    def subtotal(comps):
        return sum(round(c['rend'] * c['unit_price'], 10) for c in comps)

    suma_fija = (
        subtotal(item['materiales']) +
        subtotal(item['transporte'])
    )
    suma_her  = subtotal(item['herramientas'])
    suma_mo   = subtotal(item['mano_de_obra'])

    necesario_mo = round(precio_objetivo - suma_fija - suma_her, 10)

    # ── Caso 1: ajuste en MO posible ─────────────────────────────────────────
    if item['mano_de_obra'] and necesario_mo >= 0:
        componentes  = copy.deepcopy(item['mano_de_obra'])
        mo_original  = suma_mo if suma_mo > 0 else 1.0
        acumulado_rd = 0.0
        mo_ajustada  = []

        for i, comp in enumerate(componentes):
            up = comp['unit_price']
            if i == len(componentes) - 1:
                # Residuo exacto: total fijado por los anteriores con ROUND(,2)
                residuo_rd2 = round(necesario_mo - acumulado_rd, 2)
                comp['rend'] = _rend_exacto(residuo_rd2, up)
            else:
                peso       = (comp['rend'] * up) / mo_original if mo_original > 0 else 0
                parcial_rd = round(necesario_mo * peso, 2)
                comp['rend'] = _rend_exacto(parcial_rd, up)
                acumulado_rd += round(comp['rend'] * up, 2)
            if comp['rend'] < 0:
                comp['rend'] = 0
            mo_ajustada.append(comp)

        item['mano_de_obra'] = mo_ajustada
        total_real = round(suma_fija + suma_her + sum(round(c['rend']*c['unit_price'],2) for c in mo_ajustada), 2)
        return {'ok': True, 'metodo_ajuste': 'mano_de_obra', 'total_final': total_real}

    # ── Caso 2: MO vacía o necesario_mo < 0 → intentar con herramientas ─────
    necesario_her = round(precio_objetivo - suma_fija - suma_mo, 10)
    if item['herramientas'] and necesario_her >= 0:
        her_original = suma_her if suma_her > 0 else 1.0
        acumulado_rd = 0.0
        her_ajustada = []
        componentes  = copy.deepcopy(item['herramientas'])

        for i, comp in enumerate(componentes):
            up = comp['unit_price']
            if i == len(componentes) - 1:
                residuo_rd2  = round(necesario_her - acumulado_rd, 2)
                comp['rend'] = _rend_exacto(residuo_rd2, up)
            else:
                peso       = (comp['rend'] * up) / her_original if her_original > 0 else 0
                parcial_rd = round(necesario_her * peso, 2)
                comp['rend'] = _rend_exacto(parcial_rd, up)
                acumulado_rd += round(comp['rend'] * up, 2)
            if comp['rend'] < 0:
                comp['rend'] = 0
            her_ajustada.append(comp)

        item['herramientas'] = her_ajustada
        total_real = round(suma_fija + sum(round(c['rend']*c['unit_price'],2) for c in her_ajustada) + suma_mo, 2)
        return {'ok': True, 'metodo_ajuste': 'herramientas_equipo', 'total_final': total_real}

    # ── Caso 3: no viable ────────────────────────────────────────────────────
    return {
        'ok': False,
        'metodo_ajuste': 'ninguno',
        'total_final': round(suma_fija + suma_her + suma_mo, 2),
        'razon': f"No fue posible ajustar: precio={precio_objetivo:,.2f}, "
                 f"costo_fijo={suma_fija:,.2f}, MO disponible={suma_mo:,.2f}",
    }


def cruzar_y_ajustar(propuesta, bd_entidad):
    """
    Cruza la propuesta económica con los APUs de la entidad.
    Estrategia:
      1. Por código exacto.
      2. Si no hay match de código, buscar por similitud de descripción (umbral 60%).

    Para cada ítem cruzado, aplica _ajustar_rendimiento_mo().

    Retorna dict con:
      items_ajustados: list[item_dict]  — ítems con APU ajustado
      items_sin_apu:   list[item_dict]  — ítems sin APU (marcar rojo)
      total:           int
      detalle_cruce:   dict{codigo → {metodo, razon, metodo_ajuste}}
    """
    import copy

    # Índice de descripción normalizada para búsqueda fuzzy
    indice_desc = {
        _normalizar(apu.get('description', '')) or _normalizar(code): (code, apu)
        for code, apu in bd_entidad.items()
    }

    items_ajustados = []
    items_sin_apu   = []
    detalle_cruce   = {}

    for item_prop in propuesta:
        code    = item_prop['code']
        precio  = round(item_prop['valor_ofrecido'], 2)

        # ── Paso A: cruce por código exacto ──────────────────────────────────
        apu = bd_entidad.get(code)
        metodo_cruce = 'código exacto'

        # ── Paso B: cruce por descripción (si no hay match de código) ─────────
        if not apu:
            desc_prop = _normalizar(item_prop['description'])
            mejor_sim = 0.0
            mejor_key = None
            for desc_norm, (cod_ent, apu_ent) in indice_desc.items():
                sim = _similitud_tokens(desc_prop, desc_norm)
                if sim > mejor_sim:
                    mejor_sim = sim
                    mejor_key = (cod_ent, apu_ent)
            if mejor_sim >= 0.60 and mejor_key:
                apu = mejor_key[1]
                metodo_cruce = f"descripción similar ({mejor_sim:.0%}) → {mejor_key[0]}"
            else:
                metodo_cruce = 'sin match'

        # ── Sin APU ──────────────────────────────────────────────────────────
        if not apu:
            item_vacio = {
                **item_prop,
                'materiales': [], 'herramientas': [],
                'transporte': [], 'mano_de_obra': [],
                'tiene_apu': False,
            }
            items_sin_apu.append(item_vacio)
            detalle_cruce[code] = {
                'metodo': 'sin match',
                'razon':  'No se encontró APU por código ni por descripción en el archivo de la entidad.',
                'metodo_ajuste': None,
            }
            continue

        # ── Con APU: copiar componentes y ajustar ────────────────────────────
        item_ajustado = {
            **item_prop,
            'materiales':   copy.deepcopy(apu.get('materiales',   [])),
            'herramientas': copy.deepcopy(apu.get('herramientas', [])),
            'transporte':   copy.deepcopy(apu.get('transporte',   [])),
            'mano_de_obra': copy.deepcopy(apu.get('mano_de_obra', [])),
            'tiene_apu': True,
            'fuente_bd': 'APUs entidad',
        }

        resultado_ajuste = _ajustar_rendimiento_mo(item_ajustado, precio)
        item_ajustado['tiene_apu'] = resultado_ajuste['ok']

        if resultado_ajuste['ok']:
            items_ajustados.append(item_ajustado)
        else:
            # No fue posible ajustar — incluir con componentes originales
            # (sin modificar) marcado como sin ajuste
            item_ajustado['tiene_apu'] = False
            items_sin_apu.append(item_ajustado)

        detalle_cruce[code] = {
            'metodo':         metodo_cruce,
            'metodo_ajuste':  resultado_ajuste.get('metodo_ajuste'),
            'razon':          resultado_ajuste.get('razon', ''),
        }

    return {
        'items_ajustados': items_ajustados,
        'items_sin_apu':   items_sin_apu,
        'total':           len(propuesta),
        'detalle_cruce':   detalle_cruce,
    }


# ══════════════════════════════════════════════════════════════════════════════
# GENERACIÓN DEL EXCEL
# ══════════════════════════════════════════════════════════════════════════════

def _build_apu_sheet(ws, item, include_aiu=False, aiu_pct=0.0, sin_componentes=False):
    """Construye la hoja de APU para un ítem."""
    bold9  = Font(bold=True, name='Arial', size=9)
    norm9  = Font(bold=False, name='Arial', size=9)
    thin   = Side(style='thin')
    brd    = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_w = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    right  = Alignment(horizontal='right',  vertical='center')
    gray   = PatternFill('solid', fgColor='D9D9D9')
    blue   = PatternFill('solid', fgColor='BDD7EE')
    rojo   = PatternFill('solid', fgColor='FCE4D6')
    CUR    = '$ #,##0.00'

    ws.column_dimensions['A'].width = 38
    ws.column_dimensions['B'].width = 4
    ws.column_dimensions['C'].width = 4
    ws.column_dimensions['D'].width = 4
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 15

    r = 1

    # ── Cabecera ──────────────────────────────────────────────────────────────
    ws.row_dimensions[r].height = 35
    ws.merge_cells(f'A{r}:C{r}')
    for col, val, al in [
        ('A', 'DESCRIPCIÓN', center),
        ('D', item['description'], left_w),
        ('E', 'UNIDAD', center),
        ('F', item['unit'], center),
        ('G', 'ITEM', center),
        ('H', item['code'], center),
    ]:
        ws[f'{col}{r}'] = val
        ws[f'{col}{r}'].font = bold9
        ws[f'{col}{r}'].alignment = al
        ws[f'{col}{r}'].border = brd
        if sin_componentes:
            ws[f'{col}{r}'].fill = rojo
    r += 1

    # ── Secciones ─────────────────────────────────────────────────────────────
    def write_section(titulo, col_header, componentes):
        nonlocal r
        fill_sec = rojo if sin_componentes else gray

        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:H{r}')
        ws[f'A{r}'] = titulo
        ws[f'A{r}'].font = bold9
        ws[f'A{r}'].alignment = center
        ws[f'A{r}'].fill = fill_sec
        ws[f'A{r}'].border = brd
        r += 1

        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:D{r}')
        for col, val in [
            ('A', 'DESCRIPCION'), ('E', 'UNIDAD'),
            ('F', col_header),    ('G', 'VR UNIT'), ('H', 'VR TOTAL')
        ]:
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
                ws[f'F{r}'] = round(comp['rend'], 6)
                ws[f'G{r}'] = round(comp['unit_price'], 2)
                ws[f'H{r}'] = f'=ROUND(F{r}*G{r},2)'
            for col, al, fmt in [
                ('A', left_w, None), ('E', center, None),
                ('F', right, '#,##0.000000'), ('G', right, CUR), ('H', right, CUR)
            ]:
                ws[f'{col}{r}'].font = norm9
                ws[f'{col}{r}'].border = brd
                ws[f'{col}{r}'].alignment = al
                if fmt: ws[f'{col}{r}'].number_format = fmt
            r += 1

        end = r - 1
        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:G{r}')
        ws[f'A{r}'] = f'SUBTOTAL {titulo}'
        ws[f'A{r}'].font = bold9
        ws[f'A{r}'].alignment = right
        ws[f'A{r}'].fill = fill_sec
        ws[f'A{r}'].border = brd
        ws[f'H{r}'] = f'=ROUND(SUM(H{start}:H{end}),2)'
        ws[f'H{r}'].font = bold9
        ws[f'H{r}'].alignment = right
        ws[f'H{r}'].fill = fill_sec
        ws[f'H{r}'].border = brd
        ws[f'H{r}'].number_format = CUR
        sub = r; r += 1
        return sub

    row_mat = write_section('MATERIALES',               'CANTID.',  item['materiales'])
    row_her = write_section('HERRAMIENTAS Y/O EQUIPOS', 'RENDIM.',  item['herramientas'])
    row_tra = write_section('TRANSPORTE',               'CANTID.',  item['transporte'])
    row_mdo = write_section('MANO DE OBRA',             'RENDIM.',  item['mano_de_obra'])

    # ── AIU (opcional) ────────────────────────────────────────────────────────
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
        ws[f'H{r}'] = f'=ROUND((H{row_mat}+H{row_her}+H{row_tra}+H{row_mdo})*F{r},2)'
        ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
        ws[f'H{r}'].border = brd; ws[f'H{r}'].number_format = CUR
        aiu_row = r; r += 1

    # ── PRECIO UNITARIO ───────────────────────────────────────────────────────
    fill_pu = rojo if sin_componentes else PatternFill('solid', fgColor='BDD7EE')
    ws.row_dimensions[r].height = 18
    ws.merge_cells(f'A{r}:E{r}')
    ws[f'A{r}'] = 'PRECIO UNITARIO'
    ws[f'A{r}'].font = Font(bold=True, name='Arial', size=10)
    ws[f'A{r}'].alignment = center
    ws[f'A{r}'].fill = fill_pu
    ws[f'A{r}'].border = brd
    ws.merge_cells(f'F{r}:H{r}')
    valor_precio = round(item['valor_ofrecido'], 2)
    ws[f'F{r}'] = valor_precio
    ws[f'F{r}'].font = Font(bold=True, name='Arial', size=10)
    ws[f'F{r}'].alignment = right
    ws[f'F{r}'].fill = fill_pu
    ws[f'F{r}'].border = brd
    ws[f'F{r}'].number_format = CUR
    r += 1

    # ── TOTAL EN LETRAS ───────────────────────────────────────────────────────
    gold = PatternFill('solid', fgColor='FFF2CC')
    ws.row_dimensions[r].height = 22
    ws.merge_cells(f'A{r}:H{r}')
    ws[f'A{r}'] = numero_a_letras(int(round(valor_precio, 0)))
    ws[f'A{r}'].font = Font(bold=True, name='Arial', size=9, italic=True)
    ws[f'A{r}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws[f'A{r}'].fill = gold
    ws[f'A{r}'].border = brd


def _build_resumen_sheet(wb, items_ajustados, items_sin_apu):
    """Crea hoja RESUMEN al inicio del workbook."""
    ws = wb.create_sheet(title='RESUMEN', index=0)

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
    blanco = Font(bold=True, name='Arial', size=9, color='FFFFFF')
    CUR    = '$ #,##0.00'

    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 12
    ws.column_dimensions['H'].width = 30

    ws.row_dimensions[1].height = 28
    ws.merge_cells('A1:H1')
    ws['A1'] = 'RESUMEN DE APUs GENERADOS — Gerencia Legal Integral Colombia S.A.S.'
    ws['A1'].font = blanco
    ws['A1'].fill = azul
    ws['A1'].alignment = center

    ws.row_dimensions[2].height = 30
    encabezados = [
        ('A2', 'ÍTEM'), ('B2', 'DESCRIPCIÓN'), ('C2', 'UND.'),
        ('D2', 'PRECIO OFRECIDO'), ('E2', 'PRECIO APU'),
        ('F2', 'CIERRE'), ('G2', 'COMP.'), ('H2', 'ESTADO'),
    ]
    for celda, texto in encabezados:
        ws[celda] = texto
        ws[celda].font = bold
        ws[celda].fill = gris
        ws[celda].alignment = center
        ws[celda].border = brd

    fila = 3
    total_ofrecido = 0
    total_apu      = 0
    todos          = [(i, False) for i in items_ajustados] + [(i, True) for i in items_sin_apu]

    for item, es_sin_apu in todos:
        ws.row_dimensions[fila].height = 14
        precio_ofrecido = item.get('valor_ofrecido', 0)
        n_comp = sum(len(item.get(s, [])) for s in ('materiales','herramientas','transporte','mano_de_obra'))
        precio_apu = sum(
            round(c['rend'] * c['unit_price'], 2)
            for s in ('materiales','herramientas','transporte','mano_de_obra')
            for c in item.get(s, [])
        ) if not es_sin_apu else 0
        diff   = abs(precio_ofrecido - precio_apu)
        cierra = not es_sin_apu and diff < 0.02

        fila_fill = verde if cierra else rojo
        estado    = '✅ Ajustado' if cierra else '🔴 Sin componentes'

        for col, val, al, fmt in [
            ('A', item.get('code',''),        center, None),
            ('B', item.get('description',''), left,   None),
            ('C', item.get('unit',''),         center, None),
            ('D', precio_ofrecido,             right,  CUR),
            ('E', precio_apu if not es_sin_apu else '',  right, CUR),
            ('F', '✅' if cierra else '🔴',   center, None),
            ('G', n_comp,                      center, None),
            ('H', estado,                      left,   None),
        ]:
            c = ws[f'{col}{fila}']
            c.value = val
            c.font  = norm
            c.alignment = al
            c.border = brd
            c.fill   = fila_fill
            if fmt and isinstance(val, (int, float)):
                c.number_format = fmt

        total_ofrecido += precio_ofrecido
        total_apu      += (precio_apu if not es_sin_apu else 0)
        fila += 1

    # Fila de totales
    ws.row_dimensions[fila].height = 16
    ws.merge_cells(f'A{fila}:C{fila}')
    ws[f'A{fila}'] = f'TOTAL — {fila - 3} ítems'
    ws[f'A{fila}'].font = bold; ws[f'A{fila}'].fill = gris
    ws[f'A{fila}'].alignment = right; ws[f'A{fila}'].border = brd

    for col, val in [('D', total_ofrecido), ('E', total_apu)]:
        ws[f'{col}{fila}'] = val
        ws[f'{col}{fila}'].font = bold; ws[f'{col}{fila}'].fill = gris
        ws[f'{col}{fila}'].alignment = right; ws[f'{col}{fila}'].border = brd
        ws[f'{col}{fila}'].number_format = CUR
    for col in ('F','G','H'):
        ws[f'{col}{fila}'].border = brd; ws[f'{col}{fila}'].fill = gris

    # Leyenda
    fila += 2
    for fill, texto in [
        (verde, 'Verde — APU ajustado: rendimiento MO calibrado al precio ofrecido'),
        (rojo,  'Rojo  — Sin componentes: ítem pendiente de completar manualmente'),
    ]:
        ws.row_dimensions[fila].height = 14
        ws.merge_cells(f'A{fila}:H{fila}')
        ws[f'A{fila}'] = texto
        ws[f'A{fila}'].font = Font(italic=True, name='Arial', size=8)
        ws[f'A{fila}'].fill = fill
        ws[f'A{fila}'].alignment = left
        fila += 1


def generate_apu_excel(items_ajustados, items_sin_apu,
                       include_aiu=False, aiu_pct=0.0, bd_externas=None):
    """
    Genera el archivo Excel con:
      - Hoja RESUMEN (primera)
      - Una hoja por cada ítem (ajustados en verde, sin APU en rojo)
    """
    wb = Workbook()
    wb.remove(wb.active)

    # Intentar completar ítems sin APU con bases externas si están disponibles
    if bd_externas and items_sin_apu:
        from bases_externas import buscar_en_base, construir_apu_desde_base
        import copy
        completados = []
        aun_sin_apu = []
        for item in items_sin_apu:
            mejor = None; mejor_pct = 0
            for bd in bd_externas:
                res = buscar_en_base(item['description'], item['unit'], bd, top_n=1, umbral_pct=25)
                if res and res[0][1] > mejor_pct:
                    mejor_pct = res[0][1]; mejor = res[0][2]
            if mejor:
                apu_ext = construir_apu_desde_base(mejor, item['valor_ofrecido'])
                item_c  = {**item, **apu_ext, 'tiene_apu': True}
                completados.append(item_c)
            else:
                aun_sin_apu.append(item)
        items_ajustados = items_ajustados + completados
        items_sin_apu   = aun_sin_apu

    # Generar hojas APU — primero los ajustados, luego los rojos
    nombres_usados = set()
    for item in items_ajustados:
        name = _safe_name(item['code'])
        if name in nombres_usados: name = name[:27] + '_b'
        nombres_usados.add(name)
        ws = wb.create_sheet(title=name)
        _build_apu_sheet(ws, item, include_aiu, aiu_pct, sin_componentes=False)

    for item in items_sin_apu:
        name = _safe_name(item['code'])
        if name in nombres_usados: name = name[:27] + '_b'
        nombres_usados.add(name)
        ws = wb.create_sheet(title=name)
        _build_apu_sheet(ws, item, include_aiu, aiu_pct, sin_componentes=True)

    # Hoja RESUMEN al inicio
    _build_resumen_sheet(wb, items_ajustados, items_sin_apu)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf.getvalue()
