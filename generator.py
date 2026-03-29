"""
Generador de APUs — Gerencia Legal Integral Colombia S.A.S.
v13: Soporte de APUs en PDF (Gobernación/entidades sin Excel).

CAMBIOS v12:
  - Columnas más anchas (descripción visible completa)
  - Bordes en todas las celdas (todos los lados)
  - TRANSPORTE detectado por descripción además de tipo
  - PRECIO UNITARIO con fórmula (no hardcodeado)
  - Motor de ajuste de 3 niveles:
      1. Rendimiento Mano de Obra
      2. Rendimiento Herramientas/Equipos
      3. Valor unitario de Materiales (último recurso)
  - Ítems sin MO: ajusta materiales directamente
"""
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import io, re, unicodedata, copy, subprocess, tempfile, os

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
    """Rendimiento (10 dec) tal que ROUND(rend × unit_price, 2) == parcial_rd2."""
    if unit_price <= 0:
        return 0.0
    base = parcial_rd2 / unit_price
    for delta in [0, 1e-10, -1e-10, 2e-10, -2e-10, 5e-10, -5e-10, 1e-9, -1e-9]:
        r = round(base + delta, 10)
        if round(r * unit_price, 2) == round(parcial_rd2, 2):
            return r
    return round(base, 10)


def _es_transporte(tipo_str, desc_str):
    """
    True si el componente debe ir a la sección TRANSPORTE.
    
    REGLA CORREGIDA: se basa principalmente en el TIPO del componente
    (columna TIPO en el archivo APU), NO en la descripción.
    La descripción puede decir 'SUMINISTRO, TRANSPORTE E INSTALACION' para
    ítems de mobiliario/equipos — eso no los convierte en componentes de transporte.
    
    Solo se acepta 'transporte' por descripción si el TIPO es genérico (vacío, Insumos)
    Y la descripción contiene palabras muy específicas como 'FLETE' o 'ACARREO'.
    """
    t = tipo_str.lower().strip()
    d = desc_str.lower().strip()
    
    # El TIPO dice explícitamente transporte → sí
    if 'transporte' in t:
        return True
    # La descripción tiene palabras de flete/acarreo (servicios de transporte reales)
    # pero NO si el tipo es 'Insumos' o 'Actividad' (esos son materiales/ítems completos)
    if any(x in d for x in ('flete', 'acarreo')):
        if t not in ('insumos', 'actividad', 'analisis basico'):
            return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
# LECTORES DE APUs DE LA ENTIDAD
# ══════════════════════════════════════════════════════════════════════════════

def _leer_apu_hoja_individual(ws):
    """Lee APU desde hoja individual tipo SENA/Gobernación."""
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
        # Detectar TRANSPORTE antes de EQUIPO para no confundir
        if 'TRANSPORTE' in primera:
            seccion = 'transporte'; COL_DESC = COL_VUNIT = COL_CANT = None; continue
        if any(x in primera for x in ('EQUIPO', 'HERRAMIENTA')):
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

        # Reclasificar a TRANSPORTE si la descripción lo indica
        sec_real = 'transporte' if _es_transporte('', desc) else seccion
        apu[sec_real].append({'description': desc, 'unit': und,
                               'rend': float(cant_v), 'unit_price': float(vunit_v)})
    return apu


def _leer_hojas_apu_individuales(wb):
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
    Lee hoja APU columnar (CODINS='-').
    Clasifica TRANSPORTE por tipo O por descripción del insumo.
    Lee factor AIU desde columna J (índice 9).
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

        # Factor AIU: buscar en columnas J, K, L (índices 9-11)
        # Acepta formato factor (>1, ej: 1.3502) o porcentaje (0<x<1, ej: 0.3502)
        aiu_factor = 1.0
        for aiu_idx in [9, 10, 11]:
            aiu_raw = row[aiu_idx] if len(row) > aiu_idx else None
            if aiu_raw is None:
                continue
            try:
                aiu_v = float(aiu_raw)
                if aiu_v > 1.0:           # Factor multiplicador (ej: 1.3502)
                    aiu_factor = aiu_v; break
                elif 0.01 < aiu_v < 1.0:  # Porcentaje (ej: 0.3502 = 35.02%)
                    aiu_factor = 1.0 + aiu_v; break
            except Exception:
                pass

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
            # TRANSPORTE: chequeado primero por tipo Y por descripción
            if _es_transporte(tipo, insumo):
                bd[current]['transporte'].append(comp)
            elif any(x in t for x in ('herramienta','equipo')):
                bd[current]['herramientas'].append(comp)
            elif any(x in t for x in ('cuadrilla','personal','mano')):
                bd[current]['mano_de_obra'].append(comp)
            elif any(x in t for x in ('insumo','analisis','actividad','ensayo')):
                bd[current]['materiales'].append(comp)
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
            desc    = str(c1).strip()
            if rend > 0:
                # Reclasificar a transporte si descripción lo indica
                sec_real = 'transporte' if _es_transporte('', desc) else seccion
                current[sec_real].append({
                    'description': desc,
                    'unit':        str(c2).strip() if c2 else '',
                    'rend':        round(rend, 6),
                    'unit_price':  vr_unit,
                })
    return apus


# ══════════════════════════════════════════════════════════════════════════════
# LECTOR DE APUs DESDE PDF (Gobernación, entidades sin Excel)
# ══════════════════════════════════════════════════════════════════════════════

def _tiene_pdftotext():
    """Verifica si pdftotext (poppler-utils) está disponible en el sistema."""
    try:
        r = subprocess.run(["pdftotext", "-v"], capture_output=True, timeout=5)
        return True
    except (FileNotFoundError, OSError, Exception):
        return False


def _extraer_texto_pagina_pdf(ruta_pdf, numero_pagina):
    """
    Extrae el texto de UNA página del PDF preservando el layout espacial.
    Usa pdftotext -layout (poppler) si está disponible; si no, pdfplumber.
    """
    if _tiene_pdftotext():
        r = subprocess.run(
            ["pdftotext", "-layout",
             "-f", str(numero_pagina), "-l", str(numero_pagina),
             ruta_pdf, "-"],
            capture_output=True, text=True, timeout=30
        )
        return r.stdout
    else:
        try:
            import pdfplumber
            with pdfplumber.open(ruta_pdf) as pdf:
                if numero_pagina - 1 < len(pdf.pages):
                    return pdf.pages[numero_pagina - 1].extract_text(layout=True) or ""
        except Exception:
            pass
        return ""


def _contar_paginas_pdf(ruta_pdf):
    """Cuenta las páginas de un PDF."""
    try:
        r = subprocess.run(["pdfinfo", ruta_pdf], capture_output=True, text=True, timeout=15)
        for linea in r.stdout.split("\n"):
            if linea.startswith("Pages:"):
                return int(linea.split(":")[1].strip())
    except Exception:
        pass
    try:
        import pdfplumber
        with pdfplumber.open(ruta_pdf) as pdf:
            return len(pdf.pages)
    except Exception:
        return 0


def _num_pdf(txt):
    """Convierte texto con formato numérico colombiano a float."""
    s = str(txt).replace("$", "").replace(",", "").strip()
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def _parsear_pagina_apu_pdf(texto):
    """
    Parsea el texto de UNA página de APU (formato pdftotext -layout).
    Detecta: código, descripción, unidad, costo directo y componentes
    de las 4 secciones estándar colombianas:
      I.   EQUIPO           → herramientas
      II.  MATERIALES       → materiales
      III. TRANSPORTES      → transporte
      IV.  MANO DE OBRA     → mano_de_obra

    El dict generado es 100% compatible con el resto del flujo (cruce y ajuste).
    """
    PAT_ITEM  = re.compile(r"ITEM:\s*(\d+(?:\.\d+)+)\.?\s*(.*)", re.IGNORECASE)
    PAT_UNIT  = re.compile(r"UNIDAD:\s*(\S+)",                        re.IGNORECASE)
    PAT_TOTAL = re.compile(r"Total\s+Costo\s+Directo\s+\$\s*([\d,. ]+)", re.IGNORECASE)
    UNIDADES  = {
        "M3", "M2", "ML", "GL", "UND", "UN", "KG", "LB", "D", "TON",
        "HR", "JN", "MES", "HA", "KM", "M3-KM", "%", "H", "GBL", "VJE",
        "DIA", "VIAJE", "BULTO", "PAR", "JGO", "SET", "M", "CM", "MM",
    }

    apu = {
        "code": None, "description": "", "unit": "",
        "total_referencia": 0.0, "aiu_factor": 1.0,
        "materiales": [], "herramientas": [], "transporte": [], "mano_de_obra": [],
    }
    seccion    = None
    desc_lines = []

    for linea in texto.split("\n"):
        l = linea.strip()

        # Inicio de ítem
        m = PAT_ITEM.search(linea)
        if m:
            apu["code"] = m.group(1).rstrip(".")
            resto = m.group(2).strip()
            desc_lines = [resto] if resto else []
            seccion    = None
            continue

        if apu["code"] is None:
            continue

        # Continuación de descripción
        if not apu["description"] and not seccion and l:
            if (not l.startswith("PROVINCIA") and not l.startswith("MUNICIPIO")
                    and not re.match(r"^[IVX]+\.", l)
                    and not l.startswith("ESPECIF") and not l.startswith("UNIDAD")):
                if re.search(r"[A-ZÁÉÍÓÚÑ]{3,}", l) and not re.search(r"\d{3,}", l):
                    desc_lines.append(l)

        # Unidad del ítem
        m_u = PAT_UNIT.search(linea)
        if m_u:
            apu["unit"] = m_u.group(1).strip()
            if desc_lines and not apu["description"]:
                apu["description"] = " ".join(desc_lines).strip()
            continue

        # Secciones
        if re.match(r"\s*I\.\s+EQUIPO",              linea, re.I): seccion = "herramientas"; continue
        if re.match(r"\s*II\.\s+MATERIALES",          linea, re.I): seccion = "materiales";   continue
        if re.match(r"\s*III\.\s+TRANSPORTE",         linea, re.I): seccion = "transporte";   continue
        if re.match(r"\s*IV\.\s+MANO\s+DE\s+OBRA", linea, re.I): seccion = "mano_de_obra"; continue

        # Ignorar encabezados y subtotales
        if re.search(r"Descripci[oó]n|Trabajador|Material\s+Vol|Sub-Total", linea, re.I):
            continue

        # Total Costo Directo
        m_t = PAT_TOTAL.search(linea)
        if m_t:
            val = _num_pdf(m_t.group(1).replace(" ", ""))
            if val:
                apu["total_referencia"] = val
            continue

        if seccion is None:
            continue

        # Parsear línea de componente
        tokens = linea.split()
        if len(tokens) < 3:
            continue

        nums     = []
        desc_tok = []
        for tok in tokens:
            n = _num_pdf(tok)
            if n is not None and n > 0:
                nums.append(n)
            else:
                desc_tok.append(tok)

        if len(nums) < 2:
            continue

        valor_unit = nums[-1]
        desc_parte = " ".join(desc_tok).strip()
        unit_comp  = ""

        m_uc = re.match(
            r"^(.*?)\s+(" + "|".join(re.escape(u) for u in UNIDADES) + r")\s*$",
            desc_parte, re.IGNORECASE
        )
        if m_uc:
            desc_final = m_uc.group(1).strip()
            unit_comp  = m_uc.group(2).upper()
        else:
            desc_final = desc_parte

        if not desc_final or len(desc_final) < 2:
            continue

        if seccion in ("herramientas", "mano_de_obra"):
            rendimiento = nums[-2] if len(nums) >= 2 else 1.0
            precio_base = round(valor_unit * rendimiento, 4) if rendimiento > 0 else valor_unit
            rend_real   = round(1.0 / rendimiento, 6)        if rendimiento > 0 else 0
            apu[seccion].append({
                "description": desc_final, "unit": unit_comp,
                "rend": rend_real, "unit_price": round(precio_base, 2),
            })
        else:
            cantidad = nums[-2] if len(nums) >= 2 else 1.0
            precio_u = nums[-3] if len(nums) >= 3 else valor_unit
            apu[seccion].append({
                "description": desc_final, "unit": unit_comp,
                "rend": round(float(cantidad), 6),
                "unit_price": round(float(precio_u), 2),
            })

    return apu if apu["code"] else None


def leer_apu_entidad_pdf(archivo_pdf):
    """
    Lee APUs desde un archivo PDF (Gobernación, entidades sin Excel).
    Procesa página por página — una página = un APU.
    Si el mismo código aparece en múltiples municipios, conserva el último.

    Retorna: (dict{codigo → apu_dict}, error_str | None)
    100% compatible con cruzar_y_ajustar() y generate_apu_excel().
    """
    ruta_temporal = None
    try:
        if hasattr(archivo_pdf, "read"):
            data = archivo_pdf.read()
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
                tmp.write(data)
                ruta_temporal = tmp.name
        else:
            ruta_temporal = str(archivo_pdf)

        n_pages = _contar_paginas_pdf(ruta_temporal)
        if n_pages == 0:
            return {}, "No se pudo determinar el número de páginas del PDF."

        bd = {}
        for pag in range(1, n_pages + 1):
            texto = _extraer_texto_pagina_pdf(ruta_temporal, pag)
            apu   = _parsear_pagina_apu_pdf(texto)
            if apu and apu["code"]:
                bd[apu["code"]] = apu

        if not bd:
            return {}, (
                "No se encontraron APUs en el PDF.\n"
                "Verifique que tenga el formato estándar colombiano:\n"
                "  ITEM: X.XX → I. EQUIPO → II. MATERIALES → IV. MANO DE OBRA\n"
                "  Total Costo Directo: $..."
            )
        return bd, None

    except Exception as e:
        return {}, f"Error al procesar el PDF: {e}"
    finally:
        if ruta_temporal and hasattr(archivo_pdf, "read"):
            try:
                os.unlink(ruta_temporal)
            except Exception:
                pass


def _leer_aiu_global(wb):
    """
    Intenta extraer un factor AIU global del workbook buscando:
      - Hoja AU / AIU / A.I.U que tenga un número > 1 en sus celdas
      - Celda con el texto 'SOLVER' o 'FACTOR' cerca de un número > 1
    Retorna float (factor) o 1.0 si no encuentra nada.
    """
    for nombre in wb.sheetnames:
        n = nombre.upper().strip()
        if any(x in n for x in ('AU', 'AIU', 'A.I.U', 'UTILIDAD', 'INDIRECTO', 'GASTOS')):
            ws = wb[nombre]
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, values_only=True):
                for val in row:
                    if isinstance(val, (int, float)) and 1.01 < val < 3.0:
                        return float(val)
    return 1.0


def leer_apu_entidad(archivo):
    """
    Lee el archivo de APUs de la entidad.
    Acepta XLSX/XLSM (Excel) y PDF.

    Para Excel, intenta todos los formatos en orden de prioridad:
      1. Hojas individuales 'APU X.XX' (formato SENA/Gobernación)
      2. Hoja columnar 'APU' con CODINS='-' (formato CIMM/COMM) — lee AIU por ítem
      3. APUs embebidos con señal textual 'ANALISIS DE PRECIOS UNITARIOS - APU'
    Para formatos 1 y 3, intenta leer el AIU global desde hoja AU/AIU si existe.

    Para PDF: parsea página por página con _parsear_pagina_apu_pdf().
    El dict resultante es idéntico al de Excel — compatible con todo el flujo.

    Retorna: (dict{codigo → apu_dict}, error_str | None)
    """
    # ── Detectar si es PDF ────────────────────────────────────────────────────
    nombre = getattr(archivo, "name", "") or ""
    if nombre.lower().endswith(".pdf"):
        return leer_apu_entidad_pdf(archivo)

    # Si es BytesIO o archivo ya leído, verificar los primeros bytes (magic number PDF)
    if hasattr(archivo, "read"):
        data = archivo.read()
        if data[:4] == b"%PDF":
            import io as _io2
            return leer_apu_entidad_pdf(_io2.BytesIO(data))
        archivo = io.BytesIO(data)

    # ── Procesar Excel ────────────────────────────────────────────────────────
    try:
        wb = load_workbook(archivo, data_only=True)
    except Exception as e:
        return {}, f"No se pudo abrir el archivo de APUs de la entidad: {e}"

    bd = {}

    # Método 1: hojas individuales 'APU X.XX'
    bd.update(_leer_hojas_apu_individuales(wb))

    # Método 2: hoja columnar (lee AIU por ítem desde col J/K/L)
    for nombre in wb.sheetnames:
        if 'A.P.U' in nombre.upper() or nombre.upper() == 'APU':
            bd.update(_leer_apu_columnar(wb[nombre]))
            break

    # Método 3: APUs embebidos con señal textual
    # PROTECCIÓN: solo procesar hojas que realmente contienen la señal,
    # para no activar este método en hojas de resumen o listas.
    for nombre in wb.sheetnames:
        n_up = nombre.upper().strip()
        if n_up.startswith('APU') or 'APU' in n_up:
            apus_pd = _leer_apu_presupuesto_directo(wb[nombre])
            if apus_pd:
                bd.update(apus_pd)

    # Si se leyeron APUs con AIU=1.0 (formatos 1 y 3), intentar asignar
    # el factor AIU global desde la hoja AU si existe en el workbook.
    sin_aiu = [c for c, a in bd.items() if a.get('aiu_factor', 1.0) <= 1.0]
    if sin_aiu:
        aiu_global = _leer_aiu_global(wb)
        if aiu_global > 1.0:
            for code in sin_aiu:
                bd[code]['aiu_factor'] = aiu_global

    if not bd:
        return {}, (
            "No se encontraron APUs en el archivo de la entidad.  \n"
            "Formatos aceptados:\n"
            "  · Hojas con nombre 'APU X.XX' (una por ítem)\n"
            "  · Hoja 'APU' con formato columnar (CODINS='-')\n"
            "  · APUs embebidos con señal 'ANALISIS DE PRECIOS UNITARIOS - APU'"
        )
    return bd, None


# ══════════════════════════════════════════════════════════════════════════════
# LECTOR DE PROPUESTA ECONÓMICA
# ══════════════════════════════════════════════════════════════════════════════

# Firma única que identifica el output de esta app (aparece en A1)
_FIRMA_APP = 'RESUMEN DE APUs'

def _leer_resumen_propio(wb):
    """
    Detecta si el workbook es un output previo de esta app y lee los ítems.
    Para ser output propio debe cumplir TODAS estas condiciones:
      1. Tener hoja llamada 'RESUMEN'
      2. La celda A1 contiene la firma de la app ('RESUMEN DE APUs')
      3. La fila 2 tiene el encabezado 'PRECIO OFRECIDO' en alguna columna
    Este triple filtro evita falsos positivos con archivos de entidades que
    tengan una hoja llamada 'RESUMEN' por coincidencia.
    Retorna list[dict] o [] si no aplica.
    """
    ws_res = None
    for nombre in wb.sheetnames:
        if nombre.strip().upper() == 'RESUMEN':
            ws_res = wb[nombre]; break
    if ws_res is None:
        return []

    # Verificar firma en A1 (debe contener el texto de la app)
    a1 = str(ws_res['A1'].value or '').upper()
    if _FIRMA_APP.upper() not in a1:
        return []  # No es output de esta app

    # Verificar encabezado en fila 2
    fila2 = [str(v).upper().strip() if v else '' for v in
             next(ws_res.iter_rows(min_row=2, max_row=2, values_only=True))]
    tiene_precio_ofrecido = any('PRECIO' in c and 'OFRECIDO' in c for c in fila2)
    if not tiene_precio_ofrecido:
        return []

    # Columnas: A=código, B=desc, C=unidad, D=PRECIO OFRECIDO
    items  = []
    vistos = set()
    for row in ws_res.iter_rows(min_row=3, values_only=True):
        code = str(row[0]).strip() if row[0] else ''
        if not code or code in vistos:
            continue
        # Ignorar fila de totales (código muy largo o empieza con TOTAL)
        if 'TOTAL' in code.upper() or len(code) > 20:
            continue
        desc  = str(row[1]).strip() if row[1] else code
        und   = str(row[2]).strip() if row[2] else ''
        precio_raw = row[3]
        try:
            precio = float(precio_raw)
        except Exception:
            continue
        if precio <= 0:
            continue
        items.append({'code': code, 'description': desc, 'unit': und,
                      'valor_ofrecido': precio})
        vistos.add(code)
    return items


def leer_propuesta_economica(archivo):
    """
    Lee la propuesta económica del proponente (precios ofrecidos).

    Soporta dos formatos:
      A) Archivo con hoja RESUMEN generada por esta app (output previo).
         Columnas: ÍTEM | DESCRIPCIÓN | UND. | PRECIO OFRECIDO
      B) Cualquier Excel con columna de VALOR UNITARIO / PRECIO UNITARIO
         (formato PRESUPUESTO, FORMULARIO, etc.)

    Retorna: (list[dict], error_str | None)
    """
    try:
        data = archivo.read()
        wb   = load_workbook(io.BytesIO(data), data_only=True)
    except Exception as e:
        return [], f"No se pudo abrir la propuesta económica: {e}"

    # ── Método A: hoja RESUMEN propia de la app ───────────────────────────────
    items_resumen = _leer_resumen_propio(wb)
    if items_resumen:
        return items_resumen, None

    # ── Método B: hoja estándar de presupuesto ────────────────────────────────
    # Priorizar hojas con nombres típicos de presupuesto/formulario/oferta.
    # Si no, usar el detector genérico. Excluir hojas que parezcan APUs individuales
    # (nombre corto numérico o empieza con 'APU').
    NOMBRES_PRESUPUESTO = [
        'PRESUPUESTO', 'FORMULARIO', 'OFERTA', 'PROPUESTA', 'PROPONENTE',
        'ACTIVIDADES', 'ITEMS', 'ÍTEMS', 'PRECIOS', 'CANTIDADES', 'OFERTA ECONOMICA',
        'OFERTA ECONÓMICA', 'PRESUPUESTO OFICIAL', 'BOQ', 'BILL OF QUANTITIES',
    ]
    def _es_hoja_presupuesto_probable(nombre):
        n = nombre.upper().strip()
        if any(k in n for k in NOMBRES_PRESUPUESTO):
            return True
        return False
    def _es_hoja_apu_individual(nombre):
        """Verdadero si parece una hoja de APU individual (no queremos leerla como presupuesto)."""
        n = nombre.strip()
        # Nombre muy corto o puramente numérico → hoja APU individual
        if n.replace('.','').replace('-','').isdigit(): return True
        # Nombre tipo '003.014', '08B.844', etc.
        if re.match(r'^[0-9A-Za-z]{2,4}\.[0-9]{3,}$', n): return True
        return False

    ws       = _encontrar_hoja_presupuesto(wb)
    fila_enc, mapa = _detectar_columnas(ws)

    # Si la hoja detectada es probablemente un APU individual, buscar mejor candidato
    if _es_hoja_apu_individual(ws.title) or fila_enc is None or 'col_valor' not in mapa:
        # Primero las hojas con nombres de presupuesto
        candidatas = (
            [n for n in wb.sheetnames if _es_hoja_presupuesto_probable(n)] +
            [n for n in wb.sheetnames if not _es_hoja_presupuesto_probable(n)
                                      and not _es_hoja_apu_individual(n)]
        )
        for nombre in candidatas:
            if nombre == ws.title: continue
            ws_alt = wb[nombre]
            fila_alt, mapa_alt = _detectar_columnas(ws_alt)
            if fila_alt is not None and 'col_valor' in mapa_alt:
                ws = ws_alt; fila_enc = fila_alt; mapa = mapa_alt; break

    if fila_enc is None or 'col_valor' not in mapa:
        return [], (
            f"No se encontró columna de valor unitario en el archivo de propuesta.  \n"
            f"Hojas disponibles: {', '.join(wb.sheetnames[:10])}...  \n"
            "Verifique que el archivo tenga una hoja con columnas de PRESUPUESTO "
            "o que sea un output previo de esta aplicación."
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
        items.append({'code': codigo, 'description': desc, 'unit': und,
                      'valor_ofrecido': valor})
        vistos.add(codigo)

    if not items:
        return [], (
            f"No se encontraron ítems con valor unitario en '{ws.title}'.  \n"
            "Verifique el formato del archivo."
        )
    return items, None


# ══════════════════════════════════════════════════════════════════════════════
# MOTOR DE AJUSTE — 3 NIVELES + caso sin MO
# ══════════════════════════════════════════════════════════════════════════════

def _ajustar_rendimiento_mo(item, cd_objetivo):
    """
    Ajusta componentes para que Σ = cd_objetivo.

    Orden de prioridad (Reglas Generales):
      1. Rendimiento Mano de Obra (cuadrilla)
      2. Rendimiento Herramientas/Equipos
      3. Valor unitario de Materiales (proporcional, último recurso)
      Sin MO: ajusta materiales directamente.
    """
    def subtotal(comps):
        return sum(round(c['rend'] * c['unit_price'], 10) for c in comps)

    def _ajustar_rendimientos(componentes, necesario, original_sum):
        """Ajusta los rendimientos de una lista de componentes para sumar 'necesario'."""
        if not componentes or necesario < 0:
            return None
        orig = original_sum if original_sum > 0 else 1.0
        ajustados    = []
        acumulado_rd = 0.0
        for i, comp in enumerate(copy.deepcopy(componentes)):
            up = comp['unit_price']
            if i == len(componentes) - 1:
                residuo      = round(necesario - acumulado_rd, 2)
                comp['rend'] = _rend_exacto(residuo, up)
            else:
                peso         = (comp['rend'] * up) / orig
                parcial      = round(necesario * peso, 2)
                comp['rend'] = _rend_exacto(parcial, up)
                acumulado_rd += round(comp['rend'] * up, 2)
            comp['rend'] = max(comp['rend'], 0)
            ajustados.append(comp)
        return ajustados

    def _ajustar_precio_materiales(componentes, necesario):
        """
        Ajusta el precio unitario de los materiales proporcional al peso original.
        Solo se llama como último recurso.
        """
        if not componentes or necesario < 0:
            return None
        total_orig = subtotal(componentes)
        if total_orig <= 0:
            return None
        ajustados = []
        acumulado  = 0.0
        for i, comp in enumerate(copy.deepcopy(componentes)):
            if i == len(componentes) - 1:
                parcial_restante = round(necesario - acumulado, 2)
                comp['unit_price'] = round(parcial_restante / comp['rend'], 2) if comp['rend'] > 0 else comp['unit_price']
            else:
                peso               = (comp['rend'] * comp['unit_price']) / total_orig
                parcial            = round(necesario * peso, 2)
                comp['unit_price'] = round(parcial / comp['rend'], 2) if comp['rend'] > 0 else comp['unit_price']
                acumulado         += round(comp['rend'] * comp['unit_price'], 2)
            comp['unit_price'] = max(comp['unit_price'], 0)
            ajustados.append(comp)
        return ajustados

    suma_mat = subtotal(item['materiales'])
    suma_her = subtotal(item['herramientas'])
    suma_tra = subtotal(item['transporte'])
    suma_mo  = subtotal(item['mano_de_obra'])
    tiene_mo = bool(item['mano_de_obra'])

    # ── CASO: Ítem SIN mano de obra → ajustar materiales directamente ─────────
    if not tiene_mo:
        necesario_mat = round(cd_objetivo - suma_her - suma_tra, 10)
        if necesario_mat >= 0 and item['materiales']:
            ajustados = _ajustar_precio_materiales(item['materiales'], necesario_mat)
            if ajustados:
                item['materiales'] = ajustados
                total_real = round(subtotal(ajustados) + suma_her + suma_tra, 2)
                return {'ok': True, 'metodo_ajuste': 'Precio materiales (sin MO)', 'total_final': total_real}
        return {
            'ok': False, 'metodo_ajuste': 'ninguno',
            'total_final': round(suma_mat + suma_her + suma_tra, 2),
            'razon': f"Sin MO y no es posible ajustar materiales (CD objetivo={cd_objetivo:,.2f})",
        }

    # ── NIVEL 1: Rendimiento Mano de Obra ─────────────────────────────────────
    necesario_mo = round(cd_objetivo - suma_mat - suma_her - suma_tra, 10)
    if necesario_mo >= 0:
        ajustados = _ajustar_rendimientos(item['mano_de_obra'], necesario_mo, suma_mo)
        if ajustados:
            item['mano_de_obra'] = ajustados
            total_real = round(suma_mat + suma_her + suma_tra +
                               sum(round(c['rend']*c['unit_price'],2) for c in ajustados), 2)
            return {'ok': True, 'metodo_ajuste': 'Rendimiento mano de obra', 'total_final': total_real}

    # ── NIVEL 2: Rendimiento Herramientas/Equipos ─────────────────────────────
    necesario_her = round(cd_objetivo - suma_mat - suma_mo - suma_tra, 10)
    if necesario_her >= 0 and item['herramientas']:
        ajustados = _ajustar_rendimientos(item['herramientas'], necesario_her, suma_her)
        if ajustados:
            item['herramientas'] = ajustados
            total_real = round(suma_mat + sum(round(c['rend']*c['unit_price'],2) for c in ajustados) +
                               suma_tra + suma_mo, 2)
            return {'ok': True, 'metodo_ajuste': 'Rendimiento herramientas (nivel 2)', 'total_final': total_real}

    # ── NIVEL 3: Precio unitario de Materiales (último recurso) ───────────────
    necesario_mat = round(cd_objetivo - suma_her - suma_mo - suma_tra, 10)
    if necesario_mat >= 0 and item['materiales']:
        ajustados = _ajustar_precio_materiales(item['materiales'], necesario_mat)
        if ajustados:
            item['materiales'] = ajustados
            total_real = round(sum(round(c['rend']*c['unit_price'],2) for c in ajustados) +
                               suma_her + suma_tra + suma_mo, 2)
            return {'ok': True, 'metodo_ajuste': 'Precio materiales (último recurso)', 'total_final': total_real}

    return {
        'ok': False, 'metodo_ajuste': 'ninguno',
        'total_final': round(suma_mat + suma_her + suma_tra + suma_mo, 2),
        'razon': (
            f"No fue posible ajustar ningún componente para alcanzar "
            f"CD={cd_objetivo:,.2f} sin generar valores negativos."
        ),
    }


def cruzar_y_ajustar(propuesta, bd_entidad):
    """
    Cruza ítem por ítem la propuesta con los APUs de la entidad y ajusta.
    Cruce: 1) código exacto → 2) similitud descripción ≥ 60%.
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

        apu          = bd_entidad.get(code)
        metodo_cruce = 'Código exacto'

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
# GENERACIÓN DEL EXCEL — formato corregido
# ══════════════════════════════════════════════════════════════════════════════

def _brd_todos():
    """Border con los 4 lados en todos los bordes."""
    thin = Side(style='thin')
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def _aplicar_fila(ws, fila, cols_vals, brd, estilos_extra=None):
    """
    Aplica valor, borde y estilos a las celdas indicadas.
    cols_vals: list[(col_letra, valor, font, alignment, fill, number_format)]
    Garantiza que TODAS las columnas A-H tengan borde.
    """
    # Primero poner borde a todas las columnas aunque estén vacías
    for col in 'ABCDEFGH':
        ws[f'{col}{fila}'].border = brd
    # Luego aplicar los valores y estilos
    for item in cols_vals:
        col = item[0]; val = item[1]
        font = item[2] if len(item) > 2 else None
        aln  = item[3] if len(item) > 3 else None
        fill = item[4] if len(item) > 4 else None
        fmt  = item[5] if len(item) > 5 else None
        c = ws[f'{col}{fila}']
        c.value  = val
        c.border = brd
        if font: c.font      = font
        if aln:  c.alignment = aln
        if fill: c.fill      = fill
        if fmt:  c.number_format = fmt


def _build_apu_sheet(ws, item, sin_componentes=False):
    """
    Construye la hoja APU de un ítem.
    PRECIO UNITARIO es una fórmula (=COSTO DIRECTO + AIU).
    Todas las celdas tienen bordes en los 4 lados.
    Columnas con ancho suficiente para descripciones completas.
    """
    bold9  = Font(bold=True,  name='Arial', size=9)
    norm9  = Font(bold=False, name='Arial', size=9)
    brd    = _brd_todos()
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left_w = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    right  = Alignment(horizontal='right',  vertical='center')
    gray   = PatternFill('solid', fgColor='D9D9D9')
    blue   = PatternFill('solid', fgColor='BDD7EE')
    rojo   = PatternFill('solid', fgColor='FCE4D6')
    gold   = PatternFill('solid', fgColor='FFF2CC')
    CUR    = '$ #,##0.00'

    fill_base = rojo if sin_componentes else gray
    fill_pu   = rojo if sin_componentes else blue

    # ── Anchos de columna (descripción visible completa) ─────────────────────
    ws.column_dimensions['A'].width = 45   # descripción principal
    ws.column_dimensions['B'].width = 5
    ws.column_dimensions['C'].width = 5
    ws.column_dimensions['D'].width = 5
    ws.column_dimensions['E'].width = 12   # unidad
    ws.column_dimensions['F'].width = 14   # rendimiento / %
    ws.column_dimensions['G'].width = 16   # VR unitario
    ws.column_dimensions['H'].width = 16   # VR total

    r = 1

    # ── Cabecera: descripción / unidad / código ───────────────────────────────
    ws.row_dimensions[r].height = 40
    ws.merge_cells(f'A{r}:C{r}')
    ws.merge_cells(f'D{r}:E{r}')
    ws.merge_cells(f'F{r}:H{r}')  # unidad y código en la misma banda
    for col, val, al, fill in [
        ('A', 'DESCRIPCIÓN',       center, fill_base),
        ('D', item['description'], left_w, fill_base),
        ('F', '',                  center, fill_base),  # espacio
    ]:
        ws[f'{col}{r}'] = val
        ws[f'{col}{r}'].font = bold9; ws[f'{col}{r}'].alignment = al
        ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = fill
    # Segunda fila de cabecera: unidad y código
    r += 1
    ws.row_dimensions[r].height = 20
    ws.merge_cells(f'A{r}:C{r}')
    ws.merge_cells(f'D{r}:E{r}')
    for col, val, al, fill in [
        ('A', 'UNIDAD',    center, fill_base),
        ('D', item['unit'], center, fill_base),
        ('F', 'ITEM',      center, fill_base),
        ('G', item['code'], center, fill_base),
        ('H', '',          center, fill_base),
    ]:
        ws[f'{col}{r}'] = val
        ws[f'{col}{r}'].font = bold9; ws[f'{col}{r}'].alignment = al
        ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = fill
    for col in 'BH':
        ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = fill_base
    r += 1

    # ── Función para escribir una sección ─────────────────────────────────────
    def write_section(titulo, col_header, componentes):
        nonlocal r

        # Título de sección
        ws.row_dimensions[r].height = 15
        ws.merge_cells(f'A{r}:H{r}')
        ws[f'A{r}'] = titulo
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = center
        ws[f'A{r}'].fill = fill_base; ws[f'A{r}'].border = brd
        for col in 'BCDEFGH':
            ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = fill_base
        r += 1

        # Encabezados de columna
        ws.row_dimensions[r].height = 15
        ws.merge_cells(f'A{r}:E{r}')
        for col, val in [('A','DESCRIPCION'), ('F', col_header), ('G','VR UNIT'), ('H','VR TOTAL')]:
            ws[f'{col}{r}'] = val
            ws[f'{col}{r}'].font = bold9; ws[f'{col}{r}'].alignment = center
            ws[f'{col}{r}'].border = brd
        for col in 'BCDE':
            ws[f'{col}{r}'].border = brd
        r += 1

        start = r
        n = max(len(componentes), 2)
        for i in range(n):
            ws.row_dimensions[r].height = 15
            comp = componentes[i] if i < len(componentes) else None
            ws.merge_cells(f'A{r}:E{r}')
            if comp:
                ws[f'A{r}'] = comp['description']
                ws[f'F{r}'] = round(comp['rend'], 6)
                ws[f'G{r}'] = round(comp['unit_price'], 2)
                ws[f'H{r}'] = f'=ROUND(F{r}*G{r},2)'
            # Bordes en todas las celdas
            for col in 'ABCDE':
                ws[f'{col}{r}'].border = brd
            ws[f'A{r}'].font = norm9; ws[f'A{r}'].alignment = left_w
            ws[f'F{r}'].font = norm9; ws[f'F{r}'].alignment = right
            ws[f'F{r}'].border = brd; ws[f'F{r}'].number_format = '#,##0.000000'
            ws[f'G{r}'].font = norm9; ws[f'G{r}'].alignment = right
            ws[f'G{r}'].border = brd; ws[f'G{r}'].number_format = CUR
            ws[f'H{r}'].font = norm9; ws[f'H{r}'].alignment = right
            ws[f'H{r}'].border = brd; ws[f'H{r}'].number_format = CUR
            r += 1

        end = r - 1
        # Subtotal de sección
        ws.row_dimensions[r].height = 15
        ws.merge_cells(f'A{r}:G{r}')
        ws[f'A{r}'] = f'SUBTOTAL {titulo}'
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = right
        ws[f'A{r}'].fill = fill_base; ws[f'A{r}'].border = brd
        for col in 'BCDEFG':
            ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = fill_base
        ws[f'H{r}'] = f'=ROUND(SUM(H{start}:H{end}),2)'
        ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
        ws[f'H{r}'].fill = fill_base; ws[f'H{r}'].border = brd
        ws[f'H{r}'].number_format = CUR
        sub = r; r += 1
        return sub

    row_mat = write_section('MATERIALES',               'RENDIM./CANT.', item['materiales'])
    row_her = write_section('HERRAMIENTAS Y/O EQUIPOS', 'RENDIM.',        item['herramientas'])
    row_tra = write_section('TRANSPORTE',               'CANT.',          item['transporte'])
    row_mdo = write_section('MANO DE OBRA',             'RENDIM.',        item['mano_de_obra'])

    # ── COSTO DIRECTO ─────────────────────────────────────────────────────────
    ws.row_dimensions[r].height = 16
    ws.merge_cells(f'A{r}:G{r}')
    ws[f'A{r}'] = 'COSTO DIRECTO'
    ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = right
    ws[f'A{r}'].fill = gray; ws[f'A{r}'].border = brd
    for col in 'BCDEFG':
        ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = gray
    ws[f'H{r}'] = f'=ROUND(H{row_mat}+H{row_her}+H{row_tra}+H{row_mdo},2)'
    ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
    ws[f'H{r}'].fill = gray; ws[f'H{r}'].border = brd
    ws[f'H{r}'].number_format = CUR
    row_cd = r; r += 1

    # ── A.I.U. (si factor > 1) ────────────────────────────────────────────────
    aiu_factor = item.get('aiu_factor', 1.0)
    row_aiu    = None
    if aiu_factor > 1.0:
        aiu_pct = round(aiu_factor - 1.0, 6)
        ws.row_dimensions[r].height = 15
        ws.merge_cells(f'A{r}:D{r}')
        ws[f'A{r}'] = f'A.I.U.  ({aiu_pct*100:.2f}%)'
        ws[f'A{r}'].font = bold9; ws[f'A{r}'].alignment = left_w; ws[f'A{r}'].border = brd
        for col in 'BCD':
            ws[f'{col}{r}'].border = brd
        ws[f'E{r}'] = '%'
        ws[f'E{r}'].font = bold9; ws[f'E{r}'].alignment = center; ws[f'E{r}'].border = brd
        ws[f'F{r}'] = aiu_pct
        ws[f'F{r}'].font = bold9; ws[f'F{r}'].alignment = right
        ws[f'F{r}'].border = brd; ws[f'F{r}'].number_format = '0.0000%'
        ws[f'G{r}'].border = brd
        ws[f'H{r}'] = f'=ROUND(H{row_cd}*F{r},2)'
        ws[f'H{r}'].font = bold9; ws[f'H{r}'].alignment = right
        ws[f'H{r}'].border = brd; ws[f'H{r}'].number_format = CUR
        row_aiu = r; r += 1

    # ── PRECIO UNITARIO — con fórmula ─────────────────────────────────────────
    ws.row_dimensions[r].height = 20
    ws.merge_cells(f'A{r}:E{r}')
    ws[f'A{r}'] = 'PRECIO UNITARIO'
    ws[f'A{r}'].font = Font(bold=True, name='Arial', size=10)
    ws[f'A{r}'].alignment = center; ws[f'A{r}'].fill = fill_pu; ws[f'A{r}'].border = brd
    for col in 'BCDE':
        ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = fill_pu
    ws.merge_cells(f'F{r}:H{r}')

    # Fórmula: CD + AIU o solo CD
    if row_aiu:
        formula_pu = f'=ROUND(H{row_cd}+H{row_aiu},2)'
    else:
        formula_pu = f'=ROUND(H{row_cd},2)'

    ws[f'F{r}'] = formula_pu
    ws[f'F{r}'].font = Font(bold=True, name='Arial', size=10)
    ws[f'F{r}'].alignment = right; ws[f'F{r}'].fill = fill_pu
    ws[f'F{r}'].border = brd; ws[f'F{r}'].number_format = CUR
    for col in 'GH':
        ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = fill_pu
    row_pu = r; r += 1

    # ── VALOR EN LETRAS ───────────────────────────────────────────────────────
    ws.row_dimensions[r].height = 25
    ws.merge_cells(f'A{r}:H{r}')
    valor_precio = round(item['valor_ofrecido'], 2)
    ws[f'A{r}'] = numero_a_letras(int(round(valor_precio, 0)))
    ws[f'A{r}'].font = Font(bold=True, name='Arial', size=9, italic=True)
    ws[f'A{r}'].alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws[f'A{r}'].fill = gold; ws[f'A{r}'].border = brd
    for col in 'BCDEFGH':
        ws[f'{col}{r}'].border = brd; ws[f'{col}{r}'].fill = gold


def _build_resumen_sheet(wb, items_ajustados, items_sin_apu):
    """Crea hoja RESUMEN (primera del workbook) con bordes completos."""
    ws = wb.create_sheet(title='RESUMEN', index=0)

    bold   = Font(bold=True,  name='Arial', size=9)
    norm   = Font(bold=False, name='Arial', size=9)
    brd    = _brd_todos()
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    left   = Alignment(horizontal='left',   vertical='center', wrap_text=True)
    right  = Alignment(horizontal='right',  vertical='center')
    azul   = PatternFill('solid', fgColor='1B3A6B')
    gris   = PatternFill('solid', fgColor='D9D9D9')
    verde  = PatternFill('solid', fgColor='E2EFDA')
    rojo   = PatternFill('solid', fgColor='FCE4D6')
    blanco = Font(bold=True, name='Arial', size=9, color='FFFFFF')
    CUR    = '$ #,##0.00'

    for col, w in zip('ABCDEFGH', [14, 55, 8, 18, 18, 10, 14, 26]):
        ws.column_dimensions[col].width = w

    ws.row_dimensions[1].height = 28
    ws.merge_cells('A1:H1')
    ws['A1'] = 'RESUMEN DE APUs — Gerencia Legal Integral Colombia S.A.S.'
    ws['A1'].font = blanco; ws['A1'].fill = azul; ws['A1'].alignment = center
    for col in 'BCDEFGH':
        ws[f'{col}1'].border = brd

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
        precio_calc = round(cd_real * aiu_f, 2) if not es_sin else 0
        cierra  = not es_sin and abs(precio - precio_calc) < 0.05
        fill    = verde if cierra else rojo
        estado  = '✅ Ajustado' if cierra else '🔴 Completar manualmente'
        met     = item.get('metodo_ajuste', '') if not es_sin else ''

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
    for col in 'BC':
        ws[f'{col}{fila}'].border = brd; ws[f'{col}{fila}'].fill = gris
    ws[f'D{fila}'] = total_precio
    ws[f'D{fila}'].font = bold; ws[f'D{fila}'].fill = gris
    ws[f'D{fila}'].alignment = right; ws[f'D{fila}'].border = brd
    ws[f'D{fila}'].number_format = CUR
    for col in 'EFGH':
        ws[f'{col}{fila}'].border = brd; ws[f'{col}{fila}'].fill = gris

    # Leyenda
    fila += 2
    for fill, txt in [
        (verde, 'Verde — APU ajustado: rendimiento MO/Her o precio materiales calibrado'),
        (rojo,  'Rojo  — Sin APU o no viable: ítem para completar manualmente'),
    ]:
        ws.row_dimensions[fila].height = 14
        ws.merge_cells(f'A{fila}:H{fila}')
        ws[f'A{fila}'] = txt
        ws[f'A{fila}'].font = Font(italic=True, name='Arial', size=8)
        ws[f'A{fila}'].fill = fill; ws[f'A{fila}'].alignment = left
        ws[f'A{fila}'].border = brd
        for col in 'BCDEFGH':
            ws[f'{col}{fila}'].border = brd; ws[f'{col}{fila}'].fill = fill
        fila += 1


def generate_apu_excel(items_ajustados, items_sin_apu,
                       include_aiu=False, aiu_pct=0.0, bd_externas=None):
    """Genera el Excel final con hojas APU + RESUMEN."""
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
