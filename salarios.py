"""
salarios.py — Salarios mínimos legales Colombia (SMMLV) y jornales de obra pública.
Gerencia Legal Integral Colombia S.A.S. (GLI)
v2.0 — 2026

FUENTE OFICIAL 2026: Catálogo Policía Nacional, código 1LF-FR-0206, vigencia 2026.
Salario mínimo base confirmado: $1.750.905 (fuente: hoja MANO DE OBRA del catálogo).
Factor prestacional total: 1.763466 (fuente: hoja FP del catálogo).
"""
from __future__ import annotations

# ─── SMMLV por decreto general ────────────────────────────────────────────────
SMMLV: dict[int, int] = {
    2024: 1_300_000,
    2025: 1_423_500,
    2026: 1_423_500,   # decreto general — NO usar para Policía Nacional
    2027: 1_764_749,
    2028: 1_905_929,
}

# ─── Catálogo Policía Nacional 2026 (fuente: 1LF-FR-0206) ────────────────────
SMMLV_POLICIA_2026        = 1_750_905
FACTOR_PRESTACIONAL_2026  = 1.763466

FP_COMPONENTES = {
    "cesantia_anual":         0.0833,
    "intereses_cesantias":    0.0100,
    "vacaciones":             0.0417,
    "prima":                  0.0833,
    "dotacion":               0.0333,
    "pensiones":              0.1200,
    "salud":                  0.0850,
    "riesgos_laborales":      0.0696,
    "fondo_ind_construccion": 0.0250,
    "icbf":                   0.0300,
    "caja_compensacion":      0.0400,
}

# Tarifas hora zona A1 — fuente: hoja MANO DE OBRA del catálogo
CUADRILLAS_A1 = {
    "AYUDANTE":                                    19_298.256047,
    "MAESTRO DE OBRA":                             25_731.008063,
    "OPERARIO MAQUINARIA PESADA":                  19_298.256047,
    "TECNICO":                                     25_731.008063,
    "TOPOGRAFO":                                   34_361.510625,
    "CADENERO":                                    19_298.256047,
    "CUADRILLA I (oficial + ayudante)":            45_029.264109,
    "CUADRILLA II (técnico + ayudante)":           45_029.264109,
    "CUADRILLA III (oficial + 2 ayudantes)":       64_327.520156,
    "CUADRILLA IV (técnico + 2 ayudantes)":        64_327.520156,
    "CUADRILLA V (oficial + 3 ayudantes)":         83_625.776203,
    "CUADRILLA VI (técnico + 3 ayudantes)":        83_625.776203,
    "CUADRILLA VII (oficial + 4 ayudantes)":      102_924.032250,
    "CUADRILLA VIII (técnico + 4 ayudantes)":     102_924.032250,
    "CUADRILLA IX (2 oficiales + 2 ayudantes)":    90_760.272219,
}

MULTIPLICADORES_ZONA = {
    "A1": 1.00, "A2": 1.02, "A3": 1.05,
    "A4": 1.15, "A5": 1.60, "A6": 2.50,
}

DIAS_LABORALES_MES     = 30
HERR_MENOR_PCT_DEFAULT = 0.05


def smmlv_vigente(año: int, fuente: str = "general") -> int:
    if fuente == "policia" and año == 2026:
        return SMMLV_POLICIA_2026
    return SMMLV.get(año, SMMLV[max(SMMLV)])


def factor_prestacional(año: int = 2026) -> float:
    return FACTOR_PRESTACIONAL_2026


def cuadrilla_hora(nombre: str, zona: str = "A1") -> float:
    nombre_up = nombre.upper()
    for k, v in CUADRILLAS_A1.items():
        if nombre_up in k.upper() or k.upper() in nombre_up:
            mult = MULTIPLICADORES_ZONA.get(zona.upper(), 1.0)
            return round(v * mult, 6)
    raise ValueError(f"Cuadrilla no encontrada: {nombre!r}")


def años_disponibles() -> list[int]:
    return sorted(SMMLV.keys())


def info_año(año: int, fuente: str = "policia") -> dict:
    smmlv = smmlv_vigente(año, fuente)
    return {
        "año":                año,
        "smmlv":              smmlv,
        "factor_prestacional":factor_prestacional(año),
        "cuadrilla_III_A1":   CUADRILLAS_A1.get("CUADRILLA III (oficial + 2 ayudantes)"),
        "dias_mes":           DIAS_LABORALES_MES,
        "fuente":             "Catálogo Policía 1LF-FR-0206" if fuente == "policia" else "Decreto SMMLV general",
    }


def detectar_año_de_jornal(jornal_ref: float) -> int | None:
    base = CUADRILLAS_A1["CUADRILLA III (oficial + 2 ayudantes)"]
    for mult in MULTIPLICADORES_ZONA.values():
        if abs(jornal_ref - round(base * mult, 0)) < 500:
            return 2026
    return None
