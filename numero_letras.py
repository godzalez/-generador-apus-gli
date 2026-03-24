"""
Conversión de números a letras en español — pesos colombianos.
Uso: numero_a_letras(3552333) → "TRES MILLONES QUINIENTOS CINCUENTA Y DOS MIL TRESCIENTOS TREINTA Y TRES PESOS M/CTE"
"""

_UNIDADES = [
    '', 'UN', 'DOS', 'TRES', 'CUATRO', 'CINCO',
    'SEIS', 'SIETE', 'OCHO', 'NUEVE', 'DIEZ',
    'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE',
    'DIECISÉIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE', 'VEINTE',
    'VEINTIÚN', 'VEINTIDÓS', 'VEINTITRÉS', 'VEINTICUATRO', 'VEINTICINCO',
    'VEINTISÉIS', 'VEINTISIETE', 'VEINTIOCHO', 'VEINTINUEVE',
]

_DECENAS = [
    '', '', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA',
    'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA',
]

_CENTENAS = [
    '', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS',
    'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS',
]


def _menos_de_mil(n: int) -> str:
    """Convierte un número entre 0 y 999 a letras."""
    if n == 0:
        return ''
    if n == 100:
        return 'CIEN'
    if n < 30:
        return _UNIDADES[n]

    centenas = n // 100
    resto    = n % 100
    decenas  = resto // 10
    unidades = resto % 10

    partes = []

    if centenas:
        partes.append(_CENTENAS[centenas])

    if resto == 0:
        pass
    elif resto < 30:
        partes.append(_UNIDADES[resto])
    else:
        if unidades == 0:
            partes.append(_DECENAS[decenas])
        else:
            partes.append(f"{_DECENAS[decenas]} Y {_UNIDADES[unidades]}")

    return ' '.join(p for p in partes if p)


def numero_a_letras(valor: float) -> str:
    """
    Convierte un valor numérico a su representación en letras en pesos colombianos.
    Ejemplos:
        3552333   → "TRES MILLONES QUINIENTOS CINCUENTA Y DOS MIL TRESCIENTOS TREINTA Y TRES PESOS M/CTE"
        1000000   → "UN MILLÓN PESOS M/CTE"
        47526     → "CUARENTA Y SIETE MIL QUINIENTOS VEINTISÉIS PESOS M/CTE"
        100       → "CIEN PESOS M/CTE"
        0         → "CERO PESOS M/CTE"
    """
    n = int(round(valor, 0))

    if n == 0:
        return "CERO PESOS M/CTE"
    if n < 0:
        return f"MENOS {numero_a_letras(-n)}"

    partes = []

    # Billones (1_000_000_000_000)
    billones = n // 1_000_000_000_000
    n %= 1_000_000_000_000
    if billones == 1:
        partes.append("UN BILLÓN")
    elif billones > 1:
        partes.append(f"{_menos_de_mil(billones)} BILLONES")

    # Millones
    millones = n // 1_000_000
    n %= 1_000_000
    if millones == 1:
        partes.append("UN MILLÓN")
    elif millones > 1:
        partes.append(f"{_menos_de_mil(millones)} MILLONES")

    # Miles
    miles = n // 1_000
    n %= 1_000
    if miles == 1:
        partes.append("MIL")
    elif miles > 1:
        texto_miles = _menos_de_mil(miles)
        # "VEINTIÚN MIL" → "VEINTIÚN MIL"  (correcto)
        partes.append(f"{texto_miles} MIL")

    # Centenas/decenas/unidades
    if n > 0:
        partes.append(_menos_de_mil(n))

    resultado = ' '.join(p for p in partes if p)
    return f"{resultado} PESOS M/CTE"


# ── Tests ────────────────────────────────────────────────────────────────────
if __name__ == '__main__':
    casos = [
        (0,          "CERO PESOS M/CTE"),
        (1,          "UN PESOS M/CTE"),
        (100,        "CIEN PESOS M/CTE"),
        (1000,       "MIL PESOS M/CTE"),
        (47526,      "CUARENTA Y SIETE MIL QUINIENTOS VEINTISÉIS PESOS M/CTE"),
        (3552333,    "TRES MILLONES QUINIENTOS CINCUENTA Y DOS MIL TRESCIENTOS TREINTA Y TRES PESOS M/CTE"),
        (1000000,    "UN MILLÓN PESOS M/CTE"),
        (3848681,    "TRES MILLONES OCHOCIENTOS CUARENTA Y OCHO MIL SEISCIENTOS OCHENTA Y UN PESOS M/CTE"),
        (121004,     "CIENTO VEINTIÚN MIL CUATRO PESOS M/CTE"),
    ]
    for num, esperado in casos:
        resultado = numero_a_letras(num)
        ok = "✅" if resultado == esperado else "❌"
        print(f"{ok} {num:>15,} → {resultado}")
        if resultado != esperado:
            print(f"   Esperado: {esperado}")
