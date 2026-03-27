"""
salarios.py — Salarios mínimos legales Colombia (SMMLV) y jornales de obra pública.
Gerencia Legal Integral Colombia S.A.S. (GLI)
v1.0 — 2026
"""
from __future__ import annotations
import datetime

# ─── Tabla SMMLV oficial + proyecciones ───────────────────────────────────────
# Fuentes: Decretos anuales Ministerio del Trabajo
SMMLV: dict[int, int] = {
    2024: 1_300_000,   # Decreto 2193/2023  (+12.07%)
    2025: 1_423_500,   # Decreto 2551/2024  (+9.47%)
    2026: 1_634_027,   # Decreto 1562/2025  (+14.80%)
    2027: 1_764_749,   # Proyección          (+8.0%)
    2028: 1_905_929,   # Proyección          (+8.0%)
}

# ─── Parámetros estándar obras públicas Colombia ──────────────────────────────
DIAS_LABORALES_MES     = 26     # Estándar contractual Colombia
FACTOR_PRESTACIONAL    = 1.5    # Factor 1.5x (cubre SS, caja, SENA, ARL grado I-II)
HERR_MENOR_PCT_DEFAULT = 0.05   # 5% herramienta menor sobre subtotal MO


# ─── Funciones de cálculo ─────────────────────────────────────────────────────

def jornal_minimo(año: int) -> float:
    """
    Jornal mínimo legal para obra pública.
    = SMMLV(año) / 26 días × 1.5 (prestacional)
    """
    smmlv = SMMLV.get(año, SMMLV[max(SMMLV)])
    return round(smmlv / DIAS_LABORALES_MES * FACTOR_PRESTACIONAL, 0)


def smmlv_vigente(año: int) -> int:
    """Devuelve el SMMLV mensual para el año dado."""
    return SMMLV.get(año, SMMLV[max(SMMLV)])


def años_disponibles() -> list[int]:
    return sorted(SMMLV.keys())


def info_año(año: int) -> dict:
    """Devuelve resumen completo de un año: SMMLV, jornal, factor."""
    smmlv = smmlv_vigente(año)
    j = jornal_minimo(año)
    return {
        'año':              año,
        'smmlv':            smmlv,
        'jornal_minimo':    j,
        'factor':           FACTOR_PRESTACIONAL,
        'dias_mes':         DIAS_LABORALES_MES,
    }


# ─── Detección automática del año del proceso ─────────────────────────────────

def detectar_año_de_jornal(jornal_ref: float) -> int | None:
    """
    Dado un jornal de la lista de insumos de la entidad (código 20xxx),
    infiere el año del proceso comparando con la tabla SMMLV.
    Tolerancia: ±30% sobre el jornal mínimo legal de cada año.
    Retorna None si no puede determinar el año con certeza.
    """
    if not jornal_ref or jornal_ref <= 0:
        return None
    for año in sorted(SMMLV):
        j_min = jornal_minimo(año)
        if j_min * 0.70 <= jornal_ref <= j_min * 1.60:
            return año
    return None


def detectar_año_de_archivo(wb) -> int | None:
    """
    Intenta leer el año de las propiedades del workbook openpyxl.
    Útil cuando la entidad crea el archivo en el año del proceso.
    """
    try:
        props = wb.properties
        for attr in ('created', 'modified'):
            val = getattr(props, attr, None)
            if val and hasattr(val, 'year') and 2024 <= val.year <= 2030:
                return val.year
    except Exception:
        pass
    return None


def determinar_año_proceso(
    wb=None,
    jornales_entidad: list[float] | None = None,
    año_manual: int | None = None,
) -> tuple[int, str]:
    """
    Determina el año vigente del proceso con descripción de la fuente.

    Prioridad:
      1. Año ingresado manualmente por el usuario
      2. Inferido de los jornales de la lista de insumos de la entidad
      3. Leído de las propiedades del archivo Excel
      4. Año actual del sistema operativo

    Returns:
        (año: int, fuente: str)
    """
    # 1. Manual
    if año_manual and 2024 <= año_manual <= 2030:
        smmlv = smmlv_vigente(año_manual)
        j = jornal_minimo(año_manual)
        return año_manual, (
            f"Manual — {año_manual} | SMMLV: ${smmlv:,} | "
            f"Jornal mín.: ${j:,.0f}/día"
        )

    # 2. Jornales de la entidad
    if jornales_entidad:
        for j_ref in jornales_entidad:
            año = detectar_año_de_jornal(j_ref)
            if año:
                j = jornal_minimo(año)
                return año, (
                    f"Inferido de lista de insumos (jornal ${j_ref:,.0f} → {año}) | "
                    f"Jornal mín. legal: ${j:,.0f}/día"
                )

    # 3. Propiedades del archivo
    if wb:
        año = detectar_año_de_archivo(wb)
        if año:
            smmlv = smmlv_vigente(año)
            j = jornal_minimo(año)
            return año, (
                f"Propiedades del archivo Excel ({año}) | "
                f"SMMLV: ${smmlv:,} | Jornal mín.: ${j:,.0f}/día"
            )

    # 4. Año del sistema
    año_actual = datetime.date.today().year
    smmlv = smmlv_vigente(año_actual)
    j = jornal_minimo(año_actual)
    return año_actual, (
        f"Año actual del sistema ({año_actual}) | "
        f"SMMLV: ${smmlv:,} | Jornal mín.: ${j:,.0f}/día"
    )


def validar_jornal(unit_price: float, año: int) -> tuple[bool, float]:
    """
    Valida que un jornal de mano de obra sea ≥ mínimo legal.
    Returns: (cumple: bool, jornal_minimo_legal: float)
    """
    j_min = jornal_minimo(año)
    return unit_price >= j_min, j_min
