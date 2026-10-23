# -*- coding: utf-8 -*-
"""
Carga del conjunto de datos desde los CSV de `datos/`.

La fuente de verdad son los archivos CSV, no este módulo: para incorporar un ejercicio nuevo
se añaden filas a los CSV y no se toca una línea de código. Este módulo solo los lee y arma
las estructuras que consumen `kpis.py`, `modelo.py` y `m3s.py`.

Estructura resultante, idéntica a la que usaban esos módulos:

    PUERTOS[codigo] = {
        ...campos de identificación...,
        "eerr":    {partida: [valor por año, en orden de YEARS]},
        "esf":     {partida: [valor por año]},
        "trafico": {serie:   [valor por año]},
        "margen_ebitda_of": {año: margen},      # solo donde el regulador lo publica
        "dep_amort": {año: valor},              # solo donde el informe la desagrega
    }

Todas las cifras monetarias están en miles de dólares, tal como las publica el Ositrán.
"""
import csv
from pathlib import Path

DATOS = Path(__file__).resolve().parent.parent / "datos"

# Campos de identificación que no son texto
_NUM = {"inicio": int, "plazo": int, "tm_contenedorizada_2025": int,
        "compromiso_inv": float, "inv_acum": float, "inv_2025": float, "avance_inv": float}


def _celda(v):
    """Convierte una celda a número cuando procede; la cadena vacía es un dato ausente."""
    if v == "" or v is None:
        return None
    try:
        return int(v)
    except ValueError:
        try:
            return float(v)
        except ValueError:
            return v


def _leer(nombre):
    with open(DATOS / nombre, encoding="utf-8") as f:
        return list(csv.DictReader(f))


# --------------------------------------------------------------- identificación
PUERTOS, ORDEN = {}, []
for fila in _leer("terminales.csv"):
    k = fila.pop("codigo")
    ORDEN.append(k)
    PUERTOS[k] = {c: (_NUM[c](v) if c in _NUM else v) for c, v in fila.items()}

# --------------------------------------------------------------- series anuales
YEARS = sorted({int(f["anio"]) for f in _leer("estado_resultados.csv")})


def _serie(archivo, grupo):
    filas = _leer(archivo)
    campos = [c for c in filas[0] if c not in ("codigo", "anio")]
    for k in ORDEN:
        PUERTOS[k][grupo] = {c: [None] * len(YEARS) for c in campos}
    for f in filas:
        i = YEARS.index(int(f["anio"]))
        for c in campos:
            PUERTOS[f["codigo"]][grupo][c][i] = _celda(f[c])


_serie("estado_resultados.csv", "eerr")
_serie("estado_situacion.csv", "esf")
_serie("trafico.csv", "trafico")

# ------------------------------------------- margen EBITDA publicado por el regulador
for k in ORDEN:
    PUERTOS[k]["margen_ebitda_of"] = {}
for f in _leer("margen_ebitda_oficial.csv"):
    PUERTOS[f["codigo"]]["margen_ebitda_of"][int(f["anio"])] = float(f["margen_ebitda"])

# --- depreciación y amortización desagregada, donde el informe la publica directamente
for f in _leer("depreciacion_amortizacion.csv"):
    PUERTOS[f["codigo"]].setdefault("dep_amort", {})[int(f["anio"])] = _celda(f["dep_amort"])

# ------------------------------------------- ratios publicados, para contraste
# Una celda vacía significa que el regulador no publicó ese ratio para ese ejercicio,
# distinto de un cero: se conserva como ausencia para que la validación la salte.
RATIOS_OFICIALES = {}
for f in _leer("ratios_oficiales.csv"):
    RATIOS_OFICIALES.setdefault(f["codigo"], {}).setdefault(int(f["anio"]), {})[f["ratio"]] = _celda(f["valor"])


if __name__ == "__main__":
    print(f"{len(ORDEN)} terminales · ejercicios {YEARS[0]}–{YEARS[-1]}")
    for k in ORDEN:
        p = PUERTOS[k]
        print(f"  {k:<6} {p['nombre'][:34]:<34} ingresos {p['eerr']['ingresos_totales'][-1]:>8,} "
              f"· carga {p['trafico']['carga_tm'][-1]:>12,} TM")
