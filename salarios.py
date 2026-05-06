
SMMLV_POLICIA_2026        = 1_750_905
FACTOR_PRESTACIONAL_2026  = 1.763466

SMMLV = {2024:1_300_000, 2025:1_423_500, 2026:1_423_500}

FP_COMPONENTES = {
    "cesantia_anual":0.0833,"intereses_cesantias":0.0100,
    "vacaciones":0.0417,"prima":0.0833,"dotacion":0.0333,
    "pensiones":0.1200,"salud":0.0850,"riesgos_laborales":0.0696,
    "fondo_ind_construccion":0.0250,"icbf":0.0300,"caja_compensacion":0.0400,
}

CUADRILLAS_A1 = {
    "AYUDANTE":19298.256047,
    "MAESTRO DE OBRA":25731.008063,
    "OPERARIO MAQUINARIA PESADA":19298.256047,
    "TECNICO":25731.008063,
    "TOPOGRAFO":34361.510625,
    "CADENERO":19298.256047,
    "CUADRILLA I (oficial + ayudante)":45029.264109,
    "CUADRILLA II (tecnico + ayudante)":45029.264109,
    "CUADRILLA III (oficial + 2 ayudantes)":64327.520156,
    "CUADRILLA IV (tecnico + 2 ayudantes)":64327.520156,
    "CUADRILLA V (oficial + 3 ayudantes)":83625.776203,
    "CUADRILLA VI (tecnico + 3 ayudantes)":83625.776203,
    "CUADRILLA VII (oficial + 4 ayudantes)":102924.032250,
    "CUADRILLA VIII (tecnico + 4 ayudantes)":102924.032250,
    "CUADRILLA IX (2 oficiales + 2 ayudantes)":90760.272219,
}

MULTIPLICADORES_ZONA = {
    "A1":1.00,"A2":1.02,"A3":1.05,
    "A4":1.15,"A5":1.60,"A6":2.50,
}

DIAS_LABORALES_MES     = 30
HERR_MENOR_PCT_DEFAULT = 0.05

def smmlv_vigente(ano, fuente="general"):
    if fuente == "policia" and ano == 2026:
        return SMMLV_POLICIA_2026
    return SMMLV.get(ano, 1_750_905)

def factor_prestacional(ano=2026):
    return FACTOR_PRESTACIONAL_2026

def cuadrilla_hora(nombre, zona="A1"):
    nombre_up = nombre.upper()
    for k, v in CUADRILLAS_A1.items():
        if nombre_up in k.upper() or k.upper() in nombre_up:
            mult = MULTIPLICADORES_ZONA.get(zona.upper(), 1.0)
            return round(v * mult, 6)
    raise ValueError(f"Cuadrilla no encontrada: {nombre!r}")

def anos_disponibles():
    return sorted(SMMLV.keys())

def info_ano(ano=2026, fuente="policia"):
    return {
        "ano": ano,
        "smmlv": smmlv_vigente(ano, fuente),
        "factor_prestacional": factor_prestacional(ano),
        "cuadrilla_III_A1": CUADRILLAS_A1.get("CUADRILLA III (oficial + 2 ayudantes)"),
        "dias_mes": DIAS_LABORALES_MES,
        "fuente": "Catalogo Policia 1LF-FR-0206" if fuente=="policia" else "Decreto SMMLV general",
    }

def detectar_ano_de_jornal(jornal_ref):
    base = CUADRILLAS_A1["CUADRILLA III (oficial + 2 ayudantes)"]
    for mult in MULTIPLICADORES_ZONA.values():
        if abs(jornal_ref - round(base * mult, 0)) < 500:
            return 2026
    return None
