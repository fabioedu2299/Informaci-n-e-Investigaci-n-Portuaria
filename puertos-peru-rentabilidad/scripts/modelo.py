# -*- coding: utf-8 -*-
"""
Modelo financiero por terminal: del EBIT al flujo de caja libre.

Construcción (método indirecto, sin estado de flujos publicado por el regulador):
  NOPAT      = EBIT x (1 - tasa efectiva)
  + D&A
  - Delta capital de trabajo operativo
  - Capex                       (derivado del balance: variación de activo fijo + D&A)
  = FCFF  flujo de caja libre para la empresa
  - intereses netos despues de impuestos
  + variación de deuda financiera
  = FCFE  flujo de caja libre para el accionista

Capital de trabajo operativo = (activo corriente - efectivo - fondos restringidos)
                             - (pasivo corriente - deuda financiera corriente)

Limitación declarada: para Paita, parte de la inversión se reconoce como activo financiero
de la concesión (modelo bifurcado CINIIF 12) y no como intangible, de modo que el capex
derivado del balance subestima su inversión real. Se marca en el resultado.
"""
from dataset import PUERTOS, ORDEN

YEARS = [2021, 2022, 2023, 2024, 2025]
TASA_LEGAL = 0.295          # impuesto a la renta corporativo, Perú
WACC_ESC = [0.08, 0.10, 0.12]   # escenarios de costo de capital en USD
CAPEX_FIN = {"TPP"}         # concesiones con activo financiero relevante


def _i(y):
    return YEARS.index(y)


def serie(k, campo):
    p = PUERTOS[k]
    return p["esf"].get(campo) or p["eerr"].get(campo)


def ebitda(k, y):
    p = PUERTOS[k]
    i = _i(y)
    uo = p["eerr"]["utilidad_operativa"][i]
    if "dep_amort" in p:
        return uo + p["dep_amort"][y], p["dep_amort"][y]
    me = p.get("margen_ebitda_of", {}).get(y)
    if me is None:
        return None, None
    eb = me * p["eerr"]["ingresos_totales"][i]
    return eb, eb - uo


def capital_trabajo(k, y):
    """Capital de trabajo estrictamente operativo.

    Se excluyen del activo la caja y los fondos restringidos en fideicomiso, y del pasivo
    la deuda financiera corriente y las cuentas por pagar a entidades relacionadas: estas
    últimas son financiamiento intragrupo y no crédito comercial, y su oscilación llegó a
    USD 30 millones en un solo ejercicio en una de las concesiones.
    """
    p, i = PUERTOS[k], _i(y)
    b = p["esf"]
    return ((b["activo_corriente"][i] - b["efectivo"][i] - b["fondos_restringidos"][i])
            - (b["pasivo_corriente"][i] - b["deuda_fin_cp"][i] - b["cxp_relacionadas"][i]))


def activo_fijo(k, y):
    p, i = PUERTOS[k], _i(y)
    return p["esf"]["intangibles"][i] + p["esf"]["ppe"][i]


def tasa_efectiva(k, y):
    p, i = PUERTOS[k], _i(y)
    uai, imp = p["eerr"]["uai"][i], p["eerr"]["impuesto_renta"][i]
    if uai is None or uai <= 0:
        return TASA_LEGAL, "legal"
    t = imp / uai
    if t < 0 or t > 0.5:          # créditos fiscales o diferidos atípicos
        return TASA_LEGAL, "legal"
    return t, "efectiva"


def modelo(k, y):
    """Devuelve el puente completo del EBIT al FCFE para un terminal y un año."""
    if y == YEARS[0]:
        return None                       # hace falta el balance del año anterior
    p, i = PUERTOS[k], _i(y)
    eb, da = ebitda(k, y)
    if eb is None:
        return None
    ebit = p["eerr"]["utilidad_operativa"][i]
    t, origen_t = tasa_efectiva(k, y)
    nopat = ebit * (1 - t)

    dwc = capital_trabajo(k, y) - capital_trabajo(k, y - 1)
    capex = activo_fijo(k, y) - activo_fijo(k, y - 1) + da

    fcff = nopat + da - dwc - capex

    # Contraste con la inversión ejecutada que reporta el regulador (solo 2025, incluye IGV)
    capex_reg = p["inv_2025"] * 1000 if y == 2025 else None
    fcff_reg = (nopat + da - dwc - capex_reg) if capex_reg is not None else None

    gf = p["eerr"]["gastos_financieros"][i]
    gfi = p["eerr"]["ingresos_financieros"][i]
    int_neto_dt = (gf - gfi) * (1 - t)
    # La deuda incluye las cuentas por pagar a relacionadas de largo plazo: en Matarani y
    # Salaverry devengan interés y son financiamiento intragrupo, no crédito comercial.
    # Mismo criterio que el modelo de tres estados y las fichas por terminal.
    def _d(j):
        b = p["esf"]
        return b["deuda_fin_cp"][j] + b["deuda_fin_lp"][j] + b["deuda_relacionadas_lp"][j]

    deuda, deuda_1 = _d(i), _d(i - 1)
    dd = deuda - deuda_1
    fcfe = fcff - int_neto_dt + dd

    ing = p["eerr"]["ingresos_totales"][i]
    cap_inv = deuda + p["esf"]["patrimonio"][i]
    cap_inv_1 = deuda_1 + p["esf"]["patrimonio"][i - 1]
    cap_med = (cap_inv + cap_inv_1) / 2
    roic = nopat / cap_med if cap_med else None

    return {
        "anio": y,
        "ebit": ebit, "dep_amort": da, "ebitda": eb,
        "tasa_impositiva": t, "origen_tasa": origen_t,
        "nopat": nopat,
        "capital_trabajo": capital_trabajo(k, y),
        "var_capital_trabajo": dwc,
        "capex": capex,
        "capex_regulador": capex_reg,
        "fcff": fcff,
        "fcff_capex_regulador": fcff_reg,
        "intereses_netos_dt": int_neto_dt,
        "var_deuda": dd,
        "fcfe": fcfe,
        "deuda_financiera": deuda,
        "capital_invertido": cap_inv,
        "capital_invertido_medio": cap_med,
        # --- métricas derivadas
        "roic": roic,
        "margen_fcff": fcff / ing if ing else None,
        "conversion_caja": fcff / eb if eb else None,
        "capex_sobre_da": capex / da if da else None,
        "capex_sobre_ingresos": capex / ing if ing else None,
        "fcf_yield_capital": fcff / cap_med if cap_med else None,
        "spread": {f"{int(w*100)}": (roic - w) if roic is not None else None for w in WACC_ESC},
        "eva": {f"{int(w*100)}": ((roic - w) * cap_med) if roic is not None else None for w in WACC_ESC},
        "capex_subestimado": k in CAPEX_FIN,
    }


# ---------------- unidad de trabajo por naturaleza de la operación ----------------
SEGMENTO = {
    "DPW":   "Contenedores",
    "APMT":  "Multipropósito con contenedores",
    "TPP":   "Multipropósito con contenedores",
    "TISUR": "Graneles y carga fraccionada",
    "STI":   "Graneles y carga fraccionada",
    "TPGSM": "Graneles y carga fraccionada",
    "TCSA":  "Granel mineral especializado",
}
# unidad de trabajo pertinente para comparar dentro de cada segmento
UNIDAD = {"Contenedores": "teu",
          "Multipropósito con contenedores": "ambas",
          "Graneles y carga fraccionada": "tm",
          "Granel mineral especializado": "tm"}


def unitarios(k, y):
    """Métricas unitarias, cada una solo donde la unidad de trabajo tiene sentido."""
    p, i = PUERTOS[k], _i(y)
    ings = p["eerr"]["ing_servicios"][i]
    cs = p["eerr"]["costo_servicios"][i]
    eb, _ = ebitda(k, y)
    tm = p["trafico"]["carga_tm"][i]
    teu = p["trafico"]["teu"][i]
    nav = p["trafico"]["naves"][i]
    seg = SEGMENTO[k]
    u = UNIDAD[seg]
    d = {"segmento": seg,
         "pct_contenedorizada": p["tm_contenedorizada_2025"] / p["trafico"]["carga_tm"][-1]}
    # por tonelada: solo donde el granel domina o el terminal es multipropósito
    if u in ("tm", "ambas"):
        d.update(ingreso_por_tm=ings * 1000 / tm, costo_por_tm=cs * 1000 / tm,
                 ebitda_por_tm=(eb * 1000 / tm) if eb else None)
    # por TEU: solo donde el contenedor es la operación
    if u in ("teu", "ambas") and teu > 0:
        d.update(ingreso_por_teu=ings * 1000 / teu, costo_por_teu=cs * 1000 / teu,
                 ebitda_por_teu=(eb * 1000 / teu) if eb else None)
    # transversales: comparables entre los siete sin depender del tipo de carga
    d.update(ingreso_por_recalada=ings * 1000 / nav,
             ebitda_por_recalada=(eb * 1000 / nav) if eb else None,
             ebitda_sobre_activo_fijo=(eb / activo_fijo(k, y)) if eb else None,
             ebitda_sobre_activo=(eb / p["esf"]["activo_total"][i]) if eb else None,
             nopat_sobre_activo=None,
             rotacion_activo=p["eerr"]["ingresos_totales"][i] / p["esf"]["activo_total"][i])
    return d


def construir():
    out = {}
    for k in ORDEN:
        out[k] = {"segmento": SEGMENTO[k],
                  "modelo": [m for y in YEARS if (m := modelo(k, y))],
                  "unitarios": {str(y): unitarios(k, y) for y in YEARS}}
    return out


if __name__ == "__main__":
    import json
    d = construir()
    print(f"{'Terminal':<8}{'Año':>6}{'EBIT':>9}{'D&A':>9}{'NOPAT':>9}{'ΔWC':>9}"
          f"{'Capex':>9}{'FCFF':>10}{'FCFE':>10}{'Conv.':>8}{'ROIC':>8}")
    print("-" * 95)
    for k in ORDEN:
        for m in d[k]["modelo"]:
            print(f"{k:<8}{m['anio']:>6}{m['ebit']:>9,.0f}{m['dep_amort']:>9,.0f}{m['nopat']:>9,.0f}"
                  f"{m['var_capital_trabajo']:>9,.0f}{m['capex']:>9,.0f}{m['fcff']:>10,.0f}"
                  f"{m['fcfe']:>10,.0f}{m['conversion_caja']*100:>7.0f}%{m['roic']*100:>7.1f}%")
    json.dump(d, open("modelo.json", "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    print("\nmodelo.json escrito")
