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


# ══════════════════════════════════════════════════════════════════
# CATÁLOGO POLICÍA NACIONAL 2026 — Cargador dedicado
# Fuente: 10. PRESUPUESTO APU ESPECIFICACIONES TECNICAS PARTICULARES.xlsx
# Código oficial: 1LF-FR-0206 · Vigencia 2026
# ══════════════════════════════════════════════════════════════════

import pandas as pd

def _zona_col(zona: str) -> str:
    """Mapea zona A1..A6 al nombre de columna en las hojas del catálogo."""
    return {
        "A1": "A1", "A2": "A2", "A3": "A3",
        "A4": "A4", "A5": "A5", "A6": "A6",
    }.get(zona.upper(), "A1")


def cargar_catalogo_policia(ruta_archivo, zona: str = "A1"):
    """
    Carga el catálogo Policía Nacional 2026 y retorna cuatro diccionarios:
      actividades : {codigo → {descripcion, unidad, precio}}
      insumos     : {descripcion_norm → {descripcion, unidad, precio}}
      equipos     : {descripcion_norm → {descripcion, unidad, precio_hora}}
      mano_obra   : {descripcion_norm → {descripcion, unidad, precio_hora}}

    zona: 'A1'..'A6' — selecciona la columna de precio correspondiente.
    """
    try:
        xls = pd.ExcelFile(ruta_archivo)
    except Exception as e:
        return {}, {}, {}, {}, f"No se pudo abrir el catálogo Policía: {e}"

    zona = zona.upper()
    col_precio_act  = f"VR. UNIT."    # LISTADO ACTIVIDADES usa nombres A1..A6 en fila 0
    col_precio_insumo = "Vigencia 2026"
    col_precio_equipo = f"COSTO  {zona}"  # EQUIPOS: "COSTO  A1", "COSTO  A2", etc.

    # ── ACTIVIDADES PRE-TARIFADAS ─────────────────────────────────────────────
    actividades = {}
    try:
        df_act = pd.read_excel(ruta_archivo, sheet_name="LISTADO ACTIVIDADES", header=None)
        # Fila 0: encabezados. Columnas: 0=CAP, 1=Código, 2=ACTIVIDAD, 3=UN, 4=A1..9=A6
        # Detectar columna de zona (A1=col4, A2=col5 ... A6=col9)
        zona_idx = {"A1": 4, "A2": 5, "A3": 6, "A4": 7, "A5": 8, "A6": 9}
        col_z = zona_idx.get(zona, 4)
        for _, row in df_act.iterrows():
            cod = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            desc = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ""
            und = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ""
            precio = row.iloc[col_z] if col_z < len(row) else None
            if not cod or not desc or cod in ("NaN", "nan", "CAP.", "ACTIVIDAD"):
                continue
            if not isinstance(precio, (int, float)) or precio <= 0:
                continue
            actividades[cod] = {
                "codigo":      cod,
                "descripcion": desc,
                "unidad":      und,
                "precio":      float(precio),
                "zona":        zona,
                "fuente":      "Catálogo Policía 1LF-FR-0206 · 2026",
            }
    except Exception as e:
        pass   # hoja no encontrada — continuar con las demás

    # ── INSUMOS (precio único nacional) ──────────────────────────────────────
    insumos = {}
    try:
        df_ins = pd.read_excel(ruta_archivo, sheet_name="INSUMOS", header=0)
        # Columnas: Descripción, UND, Vigencia 2026
        for _, row in df_ins.iterrows():
            desc = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            und  = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            precio = row.iloc[2] if len(row) > 2 else None
            if not desc or desc in ("Descripción",):
                continue
            if not isinstance(precio, (int, float)) or precio <= 0:
                continue
            key = _norm(desc)
            insumos[key] = {
                "descripcion": desc,
                "unidad":      und,
                "precio":      float(precio),
                "fuente":      "Catálogo Policía 1LF-FR-0206 · 2026 · Insumos",
            }
    except Exception:
        pass

    # ── EQUIPOS (precio por zona) ─────────────────────────────────────────────
    equipos = {}
    try:
        df_eq = pd.read_excel(ruta_archivo, sheet_name="EQUIPOS", header=None)
        # Fila 0: PRECIOS ALQUILER... Fila 1: DESCRIPCIÓN, UND, COSTO A1..COSTO A6
        # Detectar columna de zona desde fila 1
        header_row = list(df_eq.iloc[1])
        col_zona_eq = None
        for j, h in enumerate(header_row):
            if h and isinstance(h, str) and zona in str(h).upper():
                col_zona_eq = j
                break
        if col_zona_eq is None:
            col_zona_eq = {"A1":2,"A2":3,"A3":4,"A4":5,"A5":6,"A6":7}.get(zona, 2)
        for _, row in df_eq.iloc[2:].iterrows():
            desc = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            und  = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            precio = row.iloc[col_zona_eq] if col_zona_eq < len(row) else None
            if not desc or not isinstance(precio, (int, float)) or precio <= 0:
                continue
            key = _norm(desc)
            equipos[key] = {
                "descripcion": desc,
                "unidad":      und,
                "precio_hora": float(precio),
                "zona":        zona,
                "fuente":      "Catálogo Policía 1LF-FR-0206 · 2026 · Equipos",
            }
    except Exception:
        pass

    # ── MANO DE OBRA (precio por zona) ───────────────────────────────────────
    mano_obra = {}
    try:
        df_mo = pd.read_excel(ruta_archivo, sheet_name="MANO DE OBRA", header=None)
        # Estructura variable — buscar filas con valor numérico > 0 en la columna de zona
        # Fila 11: encabezado con DESCRIPCIÓN, UNIDAD, VALOR UNITARIO, etc.
        # Las tarifas individuales y cuadrillas están en col2 (valor sin FP)
        # Para zona A1 usar col2; para otras zonas buscar col correspondiente
        # Simplificación robusta: leer todo y filtrar filas numéricas con descripción
        for _, row in df_mo.iterrows():
            desc = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            und  = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""
            # La tarifa sin factor prestacional está en col 2
            # Usar col 2 para A1; para otras zonas multiplicar por el factor de zona
            from salarios import MULTIPLICADORES_ZONA
            tarifa_base = row.iloc[2] if len(row) > 2 else None
            if not desc or len(desc) < 5:
                continue
            if not isinstance(tarifa_base, (int, float)) or tarifa_base <= 0:
                continue
            # Filtrar encabezados y filas no-tarifa
            if any(x in desc.upper() for x in ("DESCRIPCI", "SALARIO", "CLASIFIC",
                                                 "PERSONAL", "CUADRILLA", "ESTRUCTURA")):
                # "CUADRILLA" puede ser encabezado de sección — ignorar si no tiene número
                if not isinstance(tarifa_base, float) or tarifa_base > 50_000_000:
                    continue
            mult = MULTIPLICADORES_ZONA.get(zona, 1.0)
            precio_zona = round(float(tarifa_base) * mult, 6)
            key = _norm(desc)
            mano_obra[key] = {
                "descripcion":  desc,
                "unidad":       und if und else "mes",
                "precio_hora":  precio_zona,
                "precio_base":  float(tarifa_base),
                "zona":         zona,
                "fuente":       "Catálogo Policía 1LF-FR-0206 · 2026 · Mano de Obra",
            }
    except Exception:
        pass

    return actividades, insumos, equipos, mano_obra, None


def buscar_actividad_policia(descripcion: str, unidad: str,
                              actividades: dict, top_n: int = 5,
                              umbral_pct: int = 25) -> list:
    """
    Busca actividades del catálogo Policía por similitud de descripción.
    Retorna lista de (score, pct, item) ordenada por relevancia.
    """
    palabras = {p for p in _norm(descripcion).split()
                if len(p) >= 4 and p not in _STOPWORDS}
    if not palabras:
        return []
    und_norm = _norm(unidad) if unidad else ""
    resultados = []
    for item in actividades.values():
        palabras_item = set(_norm(item["descripcion"]).split())
        coincidencias = len(palabras & palabras_item)
        if coincidencias == 0:
            continue
        pct = coincidencias / len(palabras) * 100
        if pct < umbral_pct:
            continue
        bonus = 0.5 if und_norm and und_norm == _norm(item["unidad"]) else 0
        resultados.append((coincidencias + bonus, pct, item))
    resultados.sort(key=lambda x: -x[0])
    return resultados[:top_n]


def precio_insumo_policia(descripcion: str, insumos: dict) -> dict | None:
    """
    Busca el precio de un insumo en el catálogo Policía por nombre exacto o similar.
    Retorna el ítem más cercano o None si no hay coincidencia con score > 50%.
    """
    palabras = {p for p in _norm(descripcion).split()
                if len(p) >= 4 and p not in _STOPWORDS}
    if not palabras:
        return None
    mejor = None
    mejor_score = 0
    for key, item in insumos.items():
        palabras_item = set(key.split())
        score = len(palabras & palabras_item) / len(palabras)
        if score > mejor_score:
            mejor_score = score
            mejor = item
    return mejor if mejor_score >= 0.5 else None
