"""
detector.py — Motor de detección universal de presupuestos
Gerencia Legal Integral Colombia S.A.S. — v1.0
"""
import re
import unicodedata

# ══════════════════════════════════════════════════════════════════
# NORMALIZACIÓN DE TEXTO
# ══════════════════════════════════════════════════════════════════

def _norm(texto):
    """Normaliza a mayúsculas sin tildes para comparaciones robustas."""
    if not texto: return ''
    s = str(texto).upper().strip()
    return ''.join(
        c for c in unicodedata.normalize('NFD', s)
        if unicodedata.category(c) != 'Mn'
    )

# ══════════════════════════════════════════════════════════════════
# ALIASES DE "VALOR UNITARIO"
# ══════════════════════════════════════════════════════════════════

# Regla principal: si la celda contiene "UNIT" y NO contiene "TOTAL"
# → candidata a valor unitario.
# Lista adicional de formas cortas que no contienen "UNIT":
_ALIASES_CORTOS = {
    'P.U.', 'P.U', 'VR.U.', 'VR.U', 'VU', 'V/U',
    'VALOR/UNIT', 'VLR.U', 'V.UN.', 'PRECIO UN.',
}

def es_encabezado_valor_unitario(celda):
    """Retorna True si la celda es un encabezado de Valor Unitario."""
    n = _norm(celda)
    if not n: return False
    # Regla principal
    if 'UNIT' in n and 'TOTAL' not in n and 'PARCIAL' not in n:
        return True
    # Aliases cortos
    return n in {_norm(a) for a in _ALIASES_CORTOS}

# ══════════════════════════════════════════════════════════════════
# IDENTIFICACIÓN DE FILAS NO-ÍTEM
# ══════════════════════════════════════════════════════════════════

# Palabras que inician filas de subtotal/capítulo
_PALABRAS_NO_ITEM = {
    'SUBTOTAL', 'TOTAL', 'GRAN TOTAL', 'TOTAL GENERAL',
    'COSTO DIRECTO', 'VALOR OBRA', 'NOTA:', 'NOTA ',
    'PRESUPUESTO DE OBRA', 'FACTOR', 'GENERAL', 'PARTICULAR',
    'AIU', 'A.I.U', 'IMPREVISTOS', 'UTILIDAD',
    'ADMINISTRACION', 'ADMINISTRACIÓN',
}

def es_fila_no_item(desc_texto):
    """True si la descripción corresponde a capítulo, subtotal o encabezado."""
    n = _norm(desc_texto)
    if not n: return True
    # Descripción muy corta → capítulo o encabezado
    if len(n) < 9: return True
    # Comienza con palabra reservada
    for palabra in _PALABRAS_NO_ITEM:
        if n.startswith(_norm(palabra)):
            return True
    return False

# ══════════════════════════════════════════════════════════════════
# DETECCIÓN DE COLUMNAS
# ══════════════════════════════════════════════════════════════════

def detectar_columnas(ws, max_fila_busqueda=40):
    """
    Escanea las primeras max_fila_busqueda filas del worksheet
    buscando la fila de encabezados.
    Retorna (fila_enc_0based, mapa) donde mapa tiene:
        col_valor, col_desc, col_cod, col_und, col_cant (índices 0-based)
    """
    for i, row in enumerate(ws.iter_rows(max_row=max_fila_busqueda, values_only=True)):
        # Fila candidata a encabezado: debe tener al menos 3 celdas no vacías
        celdas_no_vacias = [v for v in row if v is not None and str(v).strip()]
        if len(celdas_no_vacias) < 3:
            continue

        # Encontrar columna Valor Unitario en esta fila
        # La celda debe ser texto corto (encabezado), no un párrafo de nota
        col_valor = None
        for j, celda in enumerate(row):
            if celda and isinstance(celda, str) and len(str(celda)) < 60:
                if es_encabezado_valor_unitario(celda):
                    col_valor = j
                    break
        if col_valor is None:
            continue

        # Mapear todas las columnas de esta fila
        mapa = {'col_valor': col_valor}
        row_n = [_norm(v) if v else '' for v in row]

        for j, n in enumerate(row_n):
            if j == col_valor: continue
            if not n: continue

            if ('VALOR TOTAL' in n or 'VR. PARC' in n or 'VR PARC' in n
                    or 'VR.PARC' in n or 'VALOR PARC' in n or 'VLR TOTAL' in n):
                mapa.setdefault('col_total', j)
            elif ('DESCRIPCION' in n or 'DESCRIPCI' in n):
                mapa.setdefault('col_desc', j)
            elif ('ITEM DE PAGO' in n or 'ITEM GEN' in n or
                  'ITEM PAGO' in n or 'CODIGO' in n or n == 'ITEM'):
                mapa.setdefault('col_cod', j)
            elif n in ('UND.','UND','UNIDAD','UN.','UN'):
                mapa.setdefault('col_und', j)
            elif n in ('CANTIDAD','CANT.','CANT','CANT. TOTAL'):
                mapa.setdefault('col_cant', j)

        # Si no encontró descripción por nombre, buscar en filas siguientes
        # la columna con textos más largos
        if 'col_desc' not in mapa:
            for row2 in ws.iter_rows(min_row=i+2, max_row=i+10, values_only=True):
                for j, val in enumerate(row2):
                    if val and isinstance(val, str) and len(val.strip()) > 15:
                        mapa['col_desc'] = j
                        break
                if 'col_desc' in mapa: break

        # Si no encontró código, buscar en columnas justo antes de descripción
        if 'col_cod' not in mapa and 'col_desc' in mapa:
            col_desc = mapa['col_desc']
            for row2 in ws.iter_rows(min_row=i+2, max_row=i+12, values_only=True):
                for j in range(max(0, col_desc - 5), col_desc):
                    v = row2[j]
                    if v is not None and es_codigo_item(v):
                        mapa['col_cod'] = j
                        break
                if 'col_cod' in mapa: break

        return i, mapa  # fila 0-based

    return None, {}

# ══════════════════════════════════════════════════════════════════
# IDENTIFICACIÓN DE CÓDIGO DE ÍTEM
# ══════════════════════════════════════════════════════════════════

# Patrones de código válidos (ordenados de más específico a más general)
_RE_CODIGOS = [
    re.compile(r'^\d{3}\.\d{3}$'),              # CCE: 001.044
    re.compile(r'^\d{3}\.\d{3}\.\d+$'),          # CCE extendido: 001.044.01
    re.compile(r'^[A-Z]{2,4}-\d+[A-Z\-]*$'),     # APU-001, APU-002-A
    re.compile(r'^\d+\.\d+\.\d+-[A-Z]+$'),        # 3.1.3-EPC
    re.compile(r'^\d+\.\d+\.\d+$'),               # 3.13.08 / 3.1.1
    re.compile(r'^\d+\-[A-Z]+$'),                 # 16.2-C
    re.compile(r'^\d+\.\d{2,3}$'),                # 2.16 / 3.01
    re.compile(r'^\d+\.\d+$'),                    # 1.1 / 2.3
]

_TEXTOS_NO_CODIGO = {
    'SUBTOTAL','TOTAL','PRELIMINARES','EXCAVACIONES','DEMOLICION',
    'SUMINISTRO','INSTALACION','RELLENO','CONCRETO','MAMPOSTERIA',
    'PINTURA','RETIRO','MOVIMIENTO','CERRAJERIA','ENCHAPE',
    'CUBIERTA','CARPINTERIA','PROYECTO','CENTRO','NOTA',
    'RESUMEN','PRESUPUESTO','FORMULARIO','GENERAL','PARTICULAR',
    'CAPITULO','DESCRIPCION','UNIDAD','ITEM','CODIGO',
}

def es_codigo_item(valor):
    """True si el valor parece ser un código de ítem de obra."""
    if valor is None: return False
    if isinstance(valor, bool): return False
    s = str(valor).strip()
    if not s or len(s) > 30 or len(s) < 2: return False
    n = _norm(s)
    # Rechazar textos que claramente no son códigos
    if any(palabra in n for palabra in _TEXTOS_NO_CODIGO):
        return False
    # Verificar contra patrones conocidos
    return any(pat.match(s) for pat in _RE_CODIGOS)

def extraer_codigo(row, mapa):
    """
    Extrae el código del ítem con sistema de prioridades:
    1. Columna detectada como col_cod
    2. Columna PARTICULAR (sub-encabezado ESPECIFICACIONES)
    3. Búsqueda de patrón válido cerca de la descripción
    4. Número de ítem (float/int como string)
    """
    col_cod  = mapa.get('col_cod')
    col_desc = mapa.get('col_desc')

    # Prioridad 1: columna de código detectada
    if col_cod is not None and col_cod < len(row):
        val = row[col_cod]
        if val is not None and str(val).strip() and es_codigo_item(val):
            return _formatear_codigo(val)

    # Prioridad 2: buscar patrón válido en columnas cercanas a descripción
    if col_desc is not None:
        for j in range(max(0, col_desc - 5), col_desc):
            if j < len(row) and es_codigo_item(row[j]):
                return _formatear_codigo(row[j])
        # También revisar columnas después (algunos formatos ponen código después de desc)
        for j in range(col_desc + 1, min(col_desc + 4, len(row))):
            if j < len(row) and es_codigo_item(row[j]):
                return _formatear_codigo(row[j])

    # Prioridad 3: número de ítem como código
    for j, val in enumerate(row):
        if isinstance(val, float) and 0 < val < 10000:
            s = f"{val:.4f}".rstrip('0').rstrip('.')
            if re.match(r'^\d+\.\d+$', s):
                return s
        elif isinstance(val, int) and 0 < val < 10000:
            return str(val)

    return 'SIN_CODIGO'

def _formatear_codigo(val):
    """Formatea un código limpiamente."""
    if isinstance(val, float):
        s = f"{val:.4f}".rstrip('0').rstrip('.')
        return s
    return str(val).strip()

# ══════════════════════════════════════════════════════════════════
# VALIDACIÓN DE FILA ÍTEM
# ══════════════════════════════════════════════════════════════════

def es_fila_item_valida(row, mapa):
    """
    Valida que una fila sea un ítem real de presupuesto.
    Requiere las 3 condiciones simultáneamente:
    1. Valor numérico > 0 en la columna VALOR UNITARIO exacta
    2. Descripción de más de 8 caracteres
    3. La descripción NO inicia con palabras de capítulo/subtotal
    """
    col_valor = mapa.get('col_valor')
    col_desc  = mapa.get('col_desc')

    if col_valor is None or col_valor >= len(row):
        return False

    # Condición 1: valor unitario numérico > 0
    valor = row[col_valor]
    if not isinstance(valor, (int, float)) or isinstance(valor, bool):
        return False
    if valor <= 0:
        return False

    # Condición 2 y 3: descripción válida
    if col_desc is not None and col_desc < len(row):
        desc = row[col_desc]
        if not desc:
            return False
        desc_str = str(desc).strip()
        if len(desc_str) < 9:
            return False
        if es_fila_no_item(desc_str):
            return False

    return True


# ══════════════════════════════════════════════════════════════════
# BÚSQUEDA DE HOJA DE PRESUPUESTO
# ══════════════════════════════════════════════════════════════════

_NOMBRES_HOJA = [
    'PRESUPUESTO', 'PROPUESTA', 'FORMULARIO', 'OFERTA',
    'ICSSA', 'LOTE', 'PRECIOS',
]

# Palabras que indican hoja del PROPONENTE (prioridad alta — son sus precios)
_HOJA_PROPONENTE = ['PROPONENTE', 'PROPUESTA ECON', 'OFERTA ECON']
# Palabras que indican hoja de la ENTIDAD (evitar si hay alternativa del proponente)
_HOJA_ENTIDAD = ['OFICIAL (ENTIDAD)', 'PRESUPUESTO OFICIAL', 'OFICIAL (ENTI']

def encontrar_hoja_presupuesto(wb):
    """
    Estrategia en cascada:
    1. Hoja del PROPONENTE (sus precios — máxima prioridad)
    2. Nombre conocido genérico
    3. Primera hoja con encabezado VALOR UNITARIO
    4. Primera hoja del archivo
    """
    nombres_upper = [n.upper() for n in wb.sheetnames]

    # Prioridad 1: hoja del proponente — solo si tiene valores (no es plantilla vacía)
    for i, n_up in enumerate(nombres_upper):
        if any(p in n_up for p in _HOJA_PROPONENTE):
            ws_candidata = wb[wb.sheetnames[i]]
            # Verificar que tiene al menos 1 valor unitario > 0
            tiene_precios = False
            for row in ws_candidata.iter_rows(max_row=50, values_only=True):
                for v in row:
                    if isinstance(v, (int, float)) and not isinstance(v, bool) and v > 1000:
                        tiene_precios = True
                        break
                if tiene_precios:
                    break
            if tiene_precios:
                return ws_candidata

    # Prioridad 2: hoja genérica conocida
    # Si el archivo tiene PROPUESTA ECONOMICA (Entidad) Y PRESUPUESTO OFICIAL,
    # preferir el PRESUPUESTO porque la propuesta puede ser la hoja del proponente vacía
    for nombre in wb.sheetnames:
        n_up = nombre.upper()
        for clave in _NOMBRES_HOJA:
            if clave.upper() in n_up:
                # No saltar ninguna — el detector de columnas filtrará si no tiene datos
                return wb[nombre]

    for nombre in wb.sheetnames:
        ws = wb[nombre]
        for row in ws.iter_rows(max_row=30, values_only=True):
            celdas = [v for v in row if v and isinstance(v, str) and len(v) < 60]
            if any(es_encabezado_valor_unitario(c) for c in celdas):
                return ws

    return wb[wb.sheetnames[0]]
