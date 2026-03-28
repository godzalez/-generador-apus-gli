"""
Generador de APUs — Gerencia Legal Integral Colombia S.A.S.
v11: Corrección AIU por ítem + precios unitarios siempre intactos.

FLUJO:
  1. leer_apu_entidad()         → lee APUs del archivo de la entidad
                                   (componentes + factor AIU por ítem desde col J)
  2. leer_propuesta_economica() → lee precios ofrecidos por el proponente
  3. cruzar_y_ajustar()         → cruza por código (luego descripción)
                                   CD_objetivo = precio_ofrecido ÷ aiu_factor
                                   Ajusta SOLO el rendimiento de MO
  4. generate_apu_excel()        → Excel: hojas APU + RESUMEN

REGLAS DE AJUSTE (Reglas Generales):
  - Materiales, herramientas, transporte → precios unitarios INTACTOS, sin cambios.
  - Único ajuste: rendimiento de mano de obra.
  - Último recurso si MO genera negativos: rendimiento de equipos/herramientas.
  - Cierre exacto a 2 decimales, sin valores negativos.
"""
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io, re, unicodedata, copy

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
    t = str(texto).lower().strip()
    t = unicodedata.normalize('NFD', t)
    t = ''.join(c for c in t if unicodedata.category(c) != 'Mn')
    t = re.sub(r'[^a-z0-9\s]', ' ', t)
    return re.sub(r'\s+', ' ', t).strip()


def _similitud_tokens(a, b):
    ta = set(_normalizar(a).split())
    tb = set(_normalizar(b).split())
    if not ta or not tb:
        return 0.0
    return len(ta & tb) / max(len(ta), len(tb))


def _safe_name(code):
    name = str(code)[:31]
    for c in '/\\?*:[]':
        name = name.replace(c, '-')
    return name


def _rend_exacto(parcial_rd2, unit_price):
    """
    Devuelve rendimiento (10 dec) tal que ROUND(rend × unit_price, 2) == parcial_rd2.
    Garantiza que la fórmula Excel =ROUND(F*G,2) cierre exactamente.
    """
    if unit_price <= 0:
        return 0.0
    base = parcial_rd2 / unit_price
    for delta in [0, 1e-10, -1e-10, 2e-10, -2e-10, 5e-10, -5e-10, 1e-9, -1e-9]:
        r = round(base + delta, 10)
        if round(r * unit_price, 2) == round(parcial_rd2, 2):
            return r
    return round(base, 10)


# ══════════════════════════════════════════════════════════════════════════════
# LECTORES DE APUs DE LA ENTIDAD
# ══════════════════════════════════════════════════════════════════════════════

def _leer_apu_hoja_individual(ws):
    """Lee APU desde hoja individual tipo SENA/Gobernación. AIU factor = 1.0."""
    apu = {
        'description': '', 'unit': '', 'total_referencia': 0,
        'aiu_factor': 1.0,
        'materiales': [], 'herramientas': [], 'transporte': [], 'mano_de_obra': []
    }
    seccion = None
    COL_DESC = COL_VUNIT = COL_CANT = None

    for row in ws.iter_rows(min_row=1, values_only=True):
        vals = [v for v in row if v is not None]
        if not vals:
            continue
        primera = str(vals[0]).upper().strip()

        if any(x in primera for x in ('INSUMO', 'MATERIAL')):
            seccion = 'materiales'; COL_DESC = COL_VUNIT = COL_CANT = None; continue
        if any(x in primera for x in ('EQUIPO', 'HERRAMIENTA', 'TRANSPORTE')):
            seccion = 'herramientas'; COL_DESC = COL_VUNIT = COL_CANT = None; continue
        if any(x in primera for x in ('MANO DE OB', 'MANO OB')):
            seccion = 'mano_de_obra'; COL_DESC = COL_VUNIT = COL_CANT = None; continue

        row_list  = list(row)
        row_upper = [str(v).upper().strip() if v else '' for v in row_list]
        if 'UNITARIO' in ' '.join(row_upper):
            for j, t in enumerate(row_upper):
                if 'DESCRIPCI' in t:               COL_DESC  = j
                if 'V. UNIT' in t or 'UNITARIO' in t: COL_VUNIT = j
                if 'CANTIDAD' in t or 'CANT' in t: COL_CANT  = j
            continue

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
        if not desc:
            continue
        apu[seccion].append({'description': desc, 'unit': und,
                              'rend': float(cant_v), 'unit_price': float(vunit_v)})
    return apu


def _leer_hojas_apu_individuales(wb):
    """Lee todas las hojas 'APU X.XX' del workbook."""
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
    """
    Lee hoja APU columnar estándar (CODINS='-').
    LEE el factor AIU desde columna J (índice 9) — formato COMM/CIMM/entidades.
    Si el valor de col J > 1, se interpreta como factor multiplicador del costo directo.
    Ejemplo: 1.3502 → el precio = costo_directo × 1.3502 → AIU = 35.02 %.
    """
    bd = {}
    current = None
    primera  = [c.value for c in ws_apu[1]]
    fila_ini = 2 if any(v and 'CODIGO' in str(v).upper() for v in primera if v) else 1

    for row in ws_apu.iter_rows(min_row=fila_ini, values_only=True):
        if not row[0]:
            continue
        code   = str(row[0]).strip()
        codins = str(row[1]).strip() if row[1] is not None else ''
        insumo = str(row[2]).strip() if row[2] else ''
        tipo   = str(row[3]).strip() if row[3] else ''
        unit   = str(row[4]).strip() if row[4] else ''
        rend   = row[5]
        uprice = row[6]
        total  = row[8]

        # Factor AIU en columna J (índice 9)
        aiu_raw = row[9] if len(row) > 9 else None
        try:
            aiu_factor = float(aiu_raw) if (aiu_raw is not None and
                                             isinstance(aiu_raw, (int, float)) and
                                             float(aiu_raw) > 1.0) else 1.0
        except Exception:
            aiu_factor = 1.0

        if not tipo:
            continue

        if codins == '-':
            current = code
            bd[code] = {
                'description': insumo, 'unit': unit,
                'total_referencia': float(total) if total else 0,
                'aiu_factor': aiu_factor,
                'materiales':[], 'herramientas':[], 'transporte':[], 'mano_de_obra':[]
            }
        elif current and code == current and rend is not None and uprice is not None:
            try:
                comp = {'description': insumo, 'unit': unit,
                        'rend': float(rend), 'unit_price': float(uprice)}
            except Exception:
                continue
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
    AIU factor = 1.0 por defecto (este formato no lo incluye).
    """
    SECCIONES = {
        10000: 'materiales', 20000: 'mano_de_obra',
        30000: 'herramientas', 40000: 'materiales', 50000: 'transporte',
    }
    COD_SEC = set(SECCIONES.keys())
    apus    = {}
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
                    'total_referencia': float(c6), 'aiu_factor': 1.0,
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
    Lee el archivo Excel de APUs de la entidad.
    Combina todos los métodos de lectura disponibles.
    Retorna: (dict{codigo → apu_dict}, error_str | None)
    """
    try:
        if hasattr(archivo, 'read'):
            data = archivo.read()
            archivo = io.BytesIO(data)
        wb = load_workbook(archivo, data_only=True)
    except Exception as e:
        return {}, f"No se pudo abrir el archivo de APUs de la entidad: {e}"

    bd = {}
    bd.update(_leer_hojas_apu_individuales(wb))
    for nombre in wb.sheetnames:
        if 'A.P.U' in nombre.upper() or nombre.upper() == 'APU':
            bd.update(_leer_apu_columnar(wb[nombre]))
            break
    for nombre in wb.sheetnames:
        n_up = nombre.upper().strip()
        if n_up.startswith('APU') or 'APU' in n_up:
            apus_pd = _leer_apu_presupuesto_directo(wb[nombre])
            if apus_pd:
                bd.update(apus_pd)

    if not bd:
        return {}, (
            "No se encontraron APUs en el archivo. "
            "Verifique que contenga hojas 'APU X.XX', hoja 'A.P.U' columnar "
            "o APUs en formato de presupuesto directo."
        )
    return bd, None


# ══════════════════════════════════════════════════════════════════════════════
# LECTOR DE PROPUESTA ECONÓMICA
# ══════════════════════════════════════════════════════════════════════════════

def leer_propuesta_economica(archivo):
    """
    Lee la propuesta económica del proponente.
    Retorna: (list[dict], error_str | None)
    """
    try:
        data = archivo.read()
        wb   = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        return [], f"No se pudo abrir la propuesta económica: {e}"

    ws       = _encontrar_hoja_presupuesto(wb)
    fila_enc, mapa = _detectar_columnas(ws)

    if fila_enc is None or 'col_valor' not in mapa:
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

    items  = []
    vistos = set()
    for row in ws.iter_rows(min_row=fila_enc + 2, values_only=True):
        if not _es_fila_item(row, mapa):
            continue
        codigo = _extraer_codigo(row, mapa)
        if not codigo or codigo in vistos:
            continue
        desc  = str(row[col_desc]).strip() if (col_desc is not None and col_desc < len(row) and row[col_desc]) else codigo
        und   = str(row[col_und]).strip()  if (col_und  is not None and col_und  < len(row) and row[col_und])  else ''
        try:
            valor = float(row[col_valor])
        except Exception:
            continue
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
# MOTOR DE AJUSTE
# ══════════════════════════════════════════════════════════════════════════════

def _ajustar_rendimiento_mo(item, cd_objetivo):
    """
    Ajusta SOLO el rendimiento de MO (o herramientas como último recurso)
    para que Σ(componentes) = cd_objetivo (costo directo = precio ÷ AIU_factor).

    Los precios unitarios de todos los insumos permanecen INTACTOS.
    """
    def subtotal(comps):
        return sum(round(c['rend'] * c['unit_price'], 10) for c in comps)

    suma_fija = subtotal(item['materiales']) + subtotal(item['transporte'])
    suma_her  = subtotal(item['herramientas'])
    suma_mo   = subtotal(item['mano_de_obra'])

    necesario_mo = round(cd_objetivo - suma_fija - suma_her, 10)

    # ── Prioridad 1: ajustar Mano de Obra ────────────────────────────────────
    if item['mano_de_obra'] and necesario_mo >= 0:
        mo_original  = suma_mo if suma_mo > 0 else 1.0
        componentes  = copy.deepcopy(item['mano_de_obra'])
        acumulado_rd = 0.0
        mo_ajustada  = []

        for i, comp in enumerate(componentes):
            up = comp['unit_price']
            if i == len(componentes) - 1:
                residuo_rd2  = round(necesario_mo - acumulado_rd, 2)
                comp['rend'] = _rend_exacto(residuo_rd2, up)
            else:
                peso       = (comp['rend'] * up) / mo_original if mo_original > 0 else 0
                parcial_rd = round(necesario_mo * peso, 2)
                comp['rend'] = _rend_exacto(parcial_rd, up)
                acumulado_rd += round(comp['rend'] * up, 2)
            comp['rend'] = max(comp['rend'], 0)
            mo_ajustada.append(comp)

        item['mano_de_obra'] = mo_ajustada
        total_real = round(suma_fija + suma_her +
                           sum(round(c['rend'] * c['unit_price'], 2) for c in mo_ajustada), 2)
        return {'ok': True, 'metodo_ajuste': 'Mano de obra', 'total_final': total_real}

    # ── Prioridad 2 (último recurso): ajustar Herramientas/Equipos ───────────
    necesario_her = round(cd_objetivo - suma_fija - suma_mo, 10)
    if item['herramientas'] and necesario_her >= 0:
        her_original = suma_her if suma_her > 0 else 1.0
        componentes  = copy.deepcopy(item['herramientas'])
        acumulado_rd = 0.0
        her_ajustada = []

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
            comp['rend'] = max(comp['rend'], 0)
            her_ajustada.append(comp)

        item['herramientas'] = her_ajustada
        total_real = round(suma_fija +
                           sum(round(c['rend'] * c['unit_price'], 2) for c in her_ajustada) +
                           suma_mo, 2)
        return {'ok': True, 'metodo_ajuste': 'Herramientas/Equipos (último recurso)', 'total_final': total_real}

    return {
        'ok': False,
        'metodo_ajuste': 'ninguno',
        'total_final': round(suma_fija + suma_her + suma_mo, 2),
        'razon': (
            f"El precio ofrecido (CD objetivo={cd_objetivo:,.2f}) es menor que el "
            f"costo fijo de materiales+transporte ({suma_fija:,.2f}). "
            "No es posible ajustar sin generar valores negativos."
        ),
    }


def cruzar_y_ajustar(propuesta, bd_entidad):
    """
    Cruza ítem por ítem la propuesta económica con los APUs de la entidad.

    Cruce: 1) código exacto → 2) similitud de descripción (umbral 60%).
    Ajuste: CD_objetivo = precio_ofrecido ÷ aiu_factor (leído del archivo entidad).
    """
    indice_desc = {
        _normalizar(apu.get('description', '')) or _normalizar(code): (code, apu)
        for code, apu in bd_entidad.items()
    }

    items_ajustados = []
    items_sin_apu   = []
    detalle_cruce   = {}

    for item_prop in propuesta:
        code   = item_prop['code']
        precio = round(item_prop['valor_ofrecido'], 2)

        # Cruce por código exacto
        apu          = bd_entidad.get(code)
        metodo_cruce = 'Código exacto'

        # Cruce por descripción si no hay match
        if not apu:
            desc_prop = _normalizar(item_prop['description'])
            mejor_sim = 0.0; mejor_key = None
            for desc_norm, (cod_ent, apu_ent) in indice_desc.items():
                sim = _similitud_tokens(desc_prop, desc_norm)
                if sim > mejor_sim:
                    mejor_sim = sim; mejor_key = (cod_ent, apu_ent)
            if mejor_sim >= 0.60 and mejor_key:
                apu          = mejor_key[1]
                metodo_cruce = f"Descripción similar ({mejor_sim:.0%}) → {mejor_key[0]}"
            else:
                metodo_cruce = 'Sin match'

        # Sin APU → marcar rojo
        if not apu:
            items_sin_apu.append({
                **item_prop,
                'materiales': [], 'herramientas': [],
                'transporte': [], 'mano_de_obra': [],
                'aiu_factor': 1.0, 'cd_final': 0, 'tiene_apu': False,
            })
            detalle_cruce[code] = {
                'metodo': 'Sin match',
                'razon': 'No se encontró APU por código ni descripción en el archivo de la entidad.',
                'metodo_ajuste': None,
            }
            continue

        # Con APU: copiar componentes y ajustar rendimiento MO
        aiu_factor  = max(float(apu.get('aiu_factor', 1.0)), 1.0)
        cd_objetivo = round(precio / aiu_factor, 10)

        item_ajustado = {
            **item_prop,
            'materiales':   copy.deepcopy(apu.get('materiales',   [])),
            'herramientas': copy.deepcopy(apu.get('herramientas', [])),
            'transporte':   copy.deepcopy(apu.get('transporte',   [])),
            'mano_de_obra': copy.deepcopy(apu.get('mano_de_obra', [])),
            'aiu_factor':   aiu_factor,
            'fuente_bd':    'APUs entidad',
            'tiene_apu':    True,
        }

        res = _ajustar_rendimiento_mo(item_ajustado, cd_objetivo)
        item_ajustado['tiene_apu'] = res['ok']
        item_ajustado['cd_final']  = res['total_final']

        if res['ok']:
            items_ajustados.append(item_ajustado)
        else:
            item_ajustado['tiene_apu'] = False
            items_sin_apu.append(item_ajustado)

        detalle_cruce[code] = {
            'metodo':        metodo_cruce,
            'aiu_factor':    aiu_factor,
            'cd_objetivo':   round(cd_objetivo, 2),
            'metodo_ajuste': res.get('metodo_ajuste'),
            'razon':         res.get('razon', ''),
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

def _build_apu_sheet(ws, item, sin_componentes=False):
    """
    Construye la hoja APU de un ítem.
    Estructura: MATERIALES → HERRAMIENTAS → TRANSPORTE → MANO DE OBRA
                COSTO DIRECTO → A.I.U. (si factor > 1) → PRECIO UNITARIO → LETRAS
    """
    bold9  = Font(bold=True,  name='Arial', size=9)
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

    fill_base = rojo if sin_componentes else gray
    fill_pu   = rojo if sin_componentes else blue

    for col, w in zip('ABCDEFGH', [38, 4, 4, 4, 14, 10, 15, 15]):
        ws.column_dimensions[col].width = w

    r = 1

    # Cabecera
    ws.row_dimensions[r].height = 35
    ws.merge_cells(f'A{r}:C{r}')
    for col, val, al in [
        ('A', 'DESCRIPCIÓN',       center),
        ('D', item['description'], left_w),
        ('E', 'UNIDAD',            center),
        ('F', item['unit'],        center),
        ('G', 'ITEM',              center),
        ('H', item['code'],        center),
    ]:
        ws[f'{col}{r}'] = val
        ws[f'{col}{r}'].font = bold9; ws[f'{col}{r}'].alignment = al
        ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = fill_base
    r += 1

    def write_section(titulo, col_header, componentes):
        nonlocal r
        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:H{r}')
        ws[f'A{r}'] = titulo
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = center
        ws[f'A{r}'].fill = fill_base; ws[f'A{r}'].border = brd
        r += 1

        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:D{r}')
        for col, val in [('A','DESCRIPCION'), ('E','UNIDAD'),
                         ('F', col_header),   ('G','VR UNIT'), ('H','VR TOTAL')]:
            ws[f'{col}{r}'] = val
            ws[f'{col}{r}'].font = bold9; ws[f'{col}{r}'].alignment = center
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
                ws[f'{col}{r}'].font = norm9; ws[f'{col}{r}'].border = brd
                ws[f'{col}{r}'].alignment = al
                if fmt: ws[f'{col}{r}'].number_format = fmt
            r += 1

        end = r - 1
        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:G{r}')
        ws[f'A{r}'] = f'SUBTOTAL {titulo}'
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = right
        ws[f'A{r}'].fill = fill_base; ws[f'A{r}'].border = brd
        ws[f'H{r}'] = f'=ROUND(SUM(H{start}:H{end}),2)'
        ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
        ws[f'H{r}'].fill = fill_base; ws[f'H{r}'].border = brd
        ws[f'H{r}'].number_format = CUR
        sub = r; r += 1
        return sub

    row_mat = write_section('MATERIALES',               'CANTID.', item['materiales'])
    row_her = write_section('HERRAMIENTAS Y/O EQUIPOS', 'RENDIM.', item['herramientas'])
    row_tra = write_section('TRANSPORTE',               'CANTID.', item['transporte'])
    row_mdo = write_section('MANO DE OBRA',             'RENDIM.', item['mano_de_obra'])

    # Costo Directo
    ws.row_dimensions[r].height = 15
    ws.merge_cells(f'A{r}:G{r}')
    ws[f'A{r}'] = 'COSTO DIRECTO'
    ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = right
    ws[f'A{r}'].fill = gray;  ws[f'A{r}'].border = brd
    ws[f'H{r}'] = f'=ROUND(H{row_mat}+H{row_her}+H{row_tra}+H{row_mdo},2)'
    ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
    ws[f'H{r}'].fill = gray;  ws[f'H{r}'].border = brd
    ws[f'H{r}'].number_format = CUR
    row_cd = r; r += 1

    # A.I.U. (si factor > 1 — leído del archivo de la entidad)
    aiu_factor = item.get('aiu_factor', 1.0)
    if aiu_factor > 1.0:
        aiu_pct = round(aiu_factor - 1.0, 6)
        ws.row_dimensions[r].height = 14
        ws.merge_cells(f'A{r}:D{r}')
        ws[f'A{r}'] = f'A.I.U.  ({aiu_pct*100:.2f}%)'
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = left_w; ws[f'A{r}'].border = brd
        ws[f'E{r}'] = '%'
        ws[f'E{r}'].font = bold9; ws[f'E{r}'].alignment = center; ws[f'E{r}'].border = brd
        ws[f'F{r}'] = aiu_pct
        ws[f'F{r}'].font = bold9; ws[f'F{r}'].alignment = right
        ws[f'F{r}'].border = brd; ws[f'F{r}'].number_format = '0.00%'
        ws[f'G{r}'].border = brd
        ws[f'H{r}'] = f'=ROUND(H{row_cd}*F{r},2)'
        ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
        ws[f'H{r}'].border = brd; ws[f'H{r}'].number_format = CUR
        r += 1

    # Precio Unitario
    ws.row_dimensions[r].height = 18
    ws.merge_cells(f'A{r}:E{r}')
    ws[f'A{r}'] = 'PRECIO UNITARIO'
    ws[f'A{r}'].font = Font(bold=True, name='Arial', size=10)
    ws[f'A{r}'].alignment = center; ws[f'A{r}'].fill = fill_pu; ws[f'A{r}'].border = brd
    ws.merge_cells(f'F{r}:H{r}')
    valor_precio = round(item['valor_ofrecido'], 2)
    ws[f'F{r}'] = valor_precio
    ws[f'F{r}'].font = Font(bold=True, name='Arial', size=10)
    ws[f'F{r}'].alignment = right; ws[f'F{r}'].fill = fill_pu
    ws[f'F{r}'].border = brd; ws[f'F{r}'].number_format = CUR
    r += 1

    # Valor en letras
    gold = PatternFill('solid', fgColor='FFF2CC')
    ws.row_dimensions[r].height = 22
    ws.merge_cells(f'A{r}:H{r}')
    ws[f'A{r}'] = numero_a_letras(int(round(valor_precio, 0)))
    ws[f'A{r}'].font = Font(bold=True, name='Arial', size=9, italic=True)
    ws[f'A{r}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws[f'A{r}'].fill = gold; ws[f'A{r}'].border = brd


def _build_resumen_sheet(wb, items_ajustados, items_sin_apu):
    """Crea hoja RESUMEN (primera del workbook)."""
    ws = wb.create_sheet(title='RESUMEN', index=0)

    bold   = Font(bold=True,  name='Arial', size=9)
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

    for col, w in zip('ABCDEFGH', [12, 50, 8, 18, 18, 10, 10, 22]):
        ws.column_dimensions[col].width = w

    ws.row_dimensions[1].height = 28
    ws.merge_cells('A1:H1')
    ws['A1'] = 'RESUMEN DE APUs — Gerencia Legal Integral Colombia S.A.S.'
    ws['A1'].font = blanco; ws['A1'].fill = azul; ws['A1'].alignment = center

    ws.row_dimensions[2].height = 30
    for celda, texto in [
        ('A2','ÍTEM'), ('B2','DESCRIPCIÓN'), ('C2','UND.'),
        ('D2','PRECIO OFRECIDO'), ('E2','COSTO DIRECTO'),
        ('F2','AIU %'), ('G2','CIERRE'), ('H2','ESTADO'),
    ]:
        ws[celda] = texto
        ws[celda].font = bold; ws[celda].fill = gris
        ws[celda].alignment = center; ws[celda].border = brd

    fila = 3
    total_precio = 0
    todos = [(i, False) for i in items_ajustados] + [(i, True) for i in items_sin_apu]

    for item, es_sin in todos:
        ws.row_dimensions[fila].height = 14
        precio  = item.get('valor_ofrecido', 0)
        aiu_f   = item.get('aiu_factor', 1.0)
        cd_real = item.get('cd_final', 0) if not es_sin else 0
        # Verificar cierre: cd × aiu_factor ≈ precio
        precio_calc = round(cd_real * aiu_f, 2) if not es_sin else 0
        cierra  = not es_sin and abs(precio - precio_calc) < 0.05
        fill    = verde if cierra else rojo
        estado  = '✅ Ajustado' if cierra else '🔴 Completar manualmente'

        for col, val, al, fmt in [
            ('A', item.get('code',''),        center, None),
            ('B', item.get('description',''), left,   None),
            ('C', item.get('unit',''),        center, None),
            ('D', precio,                     right,  CUR),
            ('E', round(cd_real, 2) if not es_sin else '', right, CUR),
            ('F', f'{(aiu_f-1)*100:.2f}%' if aiu_f > 1 else '—', center, None),
            ('G', '✅' if cierra else '🔴',  center, None),
            ('H', estado,                    left,   None),
        ]:
            c = ws[f'{col}{fila}']
            c.value = val; c.font = norm; c.alignment = al
            c.border = brd; c.fill = fill
            if fmt and isinstance(val, (int, float)):
                c.number_format = fmt

        total_precio += precio
        fila += 1

    # Totales
    ws.row_dimensions[fila].height = 16
    ws.merge_cells(f'A{fila}:C{fila}')
    ws[f'A{fila}'] = f'TOTAL — {fila - 3} ítems'
    ws[f'A{fila}'].font = bold; ws[f'A{fila}'].fill = gris
    ws[f'A{fila}'].alignment = right; ws[f'A{fila}'].border = brd
    ws[f'D{fila}'] = total_precio
    ws[f'D{fila}'].font = bold; ws[f'D{fila}'].fill = gris
    ws[f'D{fila}'].alignment = right; ws[f'D{fila}'].border = brd
    ws[f'D{fila}'].number_format = CUR
    for col in ('E','F','G','H'):
        ws[f'{col}{fila}'].border = brd; ws[f'{col}{fila}'].fill = gris

    fila += 2
    for fill, txt in [
        (verde, 'Verde — APU ajustado: rendimiento MO calibrado, precio unitario cierra exactamente'),
        (rojo,  'Rojo  — Sin APU o no viable: ítem para completar manualmente en el Excel'),
    ]:
        ws.row_dimensions[fila].height = 14
        ws.merge_cells(f'A{fila}:H{fila}')
        ws[f'A{fila}'] = txt
        ws[f'A{fila}'].font = Font(italic=True, name='Arial', size=8)
        ws[f'A{fila}'].fill = fill; ws[f'A{fila}'].alignment = left
        fila += 1


def generate_apu_excel(items_ajustados, items_sin_apu,
                       include_aiu=False, aiu_pct=0.0, bd_externas=None):
    """Genera el Excel final con hojas de APU + hoja RESUMEN."""
    wb = Workbook()
    wb.remove(wb.active)

    if bd_externas and items_sin_apu:
        from bases_externas import buscar_en_base, construir_apu_desde_base
        completados = []; aun_sin = []
        for item in items_sin_apu:
            mejor = None; mejor_pct = 0
            for bd in bd_externas:
                res = buscar_en_base(item['description'], item['unit'], bd, top_n=1, umbral_pct=25)
                if res and res[0][1] > mejor_pct:
                    mejor_pct = res[0][1]; mejor = res[0][2]
            if mejor:
                apu_ext = construir_apu_desde_base(mejor, item['valor_ofrecido'])
                completados.append({**item, **apu_ext, 'tiene_apu': True})
            else:
                aun_sin.append(item)
        items_ajustados = items_ajustados + completados
        items_sin_apu   = aun_sin

    nombres = set()
    for item in items_ajustados:
        name = _safe_name(item['code'])
        if name in nombres: name = name[:27] + '_b'
        nombres.add(name)
        _build_apu_sheet(wb.create_sheet(title=name), item, sin_componentes=False)

    for item in items_sin_apu:
        name = _safe_name(item['code'])
        if name in nombres: name = name[:27] + '_b'
        nombres.add(name)
        _build_apu_sheet(wb.create_sheet(title=name), item, sin_componentes=True)

    _build_resumen_sheet(wb, items_ajustados, items_sin_apu)

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf.getvalue()
