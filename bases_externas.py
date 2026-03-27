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
