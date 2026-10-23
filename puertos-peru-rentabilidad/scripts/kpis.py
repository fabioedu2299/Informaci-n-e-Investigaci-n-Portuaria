# -*- coding: utf-8 -*-
"""Cálculo de KPIs financieros y operativos + validación contra ratios oficiales OSITRAN."""
import json
from dataset import PUERTOS, ORDEN, YEARS, RATIOS_OFICIALES


def div(a, b):
    try:
        return a / b if b not in (0, None) and a is not None else None
    except Exception:
        return None


def calcular():
    out = {}
    for k in ORDEN:
        p = PUERTOS[k]
        e, b = p["eerr"], p["esf"]
        rows = []
        for i, y in enumerate(YEARS):
            ing = e["ingresos_totales"][i]
            ing_serv = e["ing_servicios"][i]
            ub, uo, un = e["utilidad_bruta"][i], e["utilidad_operativa"][i], e["utilidad_neta"][i]
            act, pat = b["activo_total"][i], b["patrimonio"][i]
            pas = b["pasivo_total"][i]
            ac, pc = b["activo_corriente"][i], b["pasivo_corriente"][i]
            inv = b["inventarios"][i]
            # Criterio unificado con el modelo de tres estados y las fichas: las cuentas por
            # pagar a relacionadas de largo plazo devengan interés en Matarani y Salaverry,
            # de modo que son financiamiento intragrupo y no crédito comercial.
            deuda = (b["deuda_fin_cp"][i] + b["deuda_fin_lp"][i]
                     + b["deuda_relacionadas_lp"][i])

            # EBITDA: D&A explícita si existe; si no, desde el margen EBITDA oficial de OSITRAN
            da = None
            if "dep_amort" in p:
                da = p["dep_amort"][y]
                ebitda = uo + da
            elif y in p.get("margen_ebitda_of", {}):
                ebitda = p["margen_ebitda_of"][y] * ing
                da = ebitda - uo
            else:
                ebitda = None

            rows.append({
                "anio": y,
                # --- Estado de resultados
                "ingresos_totales": ing, "ingresos_servicios": ing_serv,
                "ingresos_construccion": e["ing_construccion"][i],
                "costos_totales": e["costos_totales"][i], "costo_servicios": e["costo_servicios"][i],
                "utilidad_bruta": ub,
                "gastos_admin": e["gastos_admin"][i], "gastos_venta": e["gastos_venta"][i],
                "utilidad_operativa": uo,
                "ingresos_financieros": e["ingresos_financieros"][i],
                "gastos_financieros": e["gastos_financieros"][i],
                "uai": e["uai"][i], "impuesto_renta": e["impuesto_renta"][i],
                "utilidad_neta": un,
                "ebitda": round(ebitda, 1) if ebitda is not None else None,
                "dep_amort": round(da, 1) if da is not None else None,
                # --- Balance
                "efectivo": b["efectivo"][i], "ctas_cobrar_com": b["ctas_cobrar_com"][i],
                "inventarios": inv, "activo_corriente": ac,
                "activo_no_corriente": b["activo_no_corriente"][i], "activo_total": act,
                "intangibles": b["intangibles"][i], "ppe": b["ppe"][i],
                "pasivo_corriente": pc, "pasivo_no_corriente": b["pasivo_no_corriente"][i],
                "pasivo_total": pas, "deuda_financiera": deuda, "patrimonio": pat,
                # --- Rentabilidad
                "margen_bruto": div(ub, ing), "margen_operativo": div(uo, ing),
                "margen_neto": div(un, ing), "margen_ebitda": div(ebitda, ing),
                "roa": div(un, act), "roe": div(un, pat),
                "roic_aprox": div(uo * (1 - div(e["impuesto_renta"][i], e["uai"][i])
                                        if e["uai"][i] not in (0, None) else 0.7),
                                  (deuda + pat)),
                # --- Liquidez y solvencia
                "liquidez_corriente": div(ac, pc),
                "prueba_acida": div(ac - inv, pc),
                "capital_trabajo": ac - pc,
                "apalancamiento": div(act, pat),
                "deuda_patrimonio": div(pas, pat),
                "endeudamiento_activo": div(pas, act),
                "deuda_neta_ebitda": div(deuda - b["efectivo"][i], ebitda) if ebitda else None,
                "cobertura_intereses": div(ebitda, e["gastos_financieros"][i]) if ebitda else None,
                # --- Operativo
                "naves": p["trafico"]["naves"][i],
                "carga_tm": p["trafico"]["carga_tm"][i],
                "teu": p["trafico"]["teu"][i],
                "contenedores": p["trafico"]["contenedores"][i],
                # --- Eficiencia (USD por unidad física; ingresos en miles -> x1000)
                "ingreso_por_tm": div(ing_serv * 1000, p["trafico"]["carga_tm"][i]),
                "costo_por_tm": div(e["costo_servicios"][i] * 1000, p["trafico"]["carga_tm"][i]),
                "ebitda_por_tm": div(ebitda * 1000, p["trafico"]["carga_tm"][i]) if ebitda else None,
                "ingreso_por_teu": div(ing_serv * 1000, p["trafico"]["teu"][i]),
                "tm_por_nave": div(p["trafico"]["carga_tm"][i], p["trafico"]["naves"][i]),
                "ingreso_por_nave": div(ing_serv * 1000, p["trafico"]["naves"][i]),
                "rotacion_activos": div(ing, act),
                "intensidad_capital": div(act, ing),
                "dias_cobro": div(b["ctas_cobrar_com"][i] * 365, ing_serv),
            })
        out[k] = {"meta": {kk: vv for kk, vv in p.items()
                           if kk not in ("eerr", "esf", "trafico", "dep_amort", "margen_ebitda_of")},
                  "serie": rows}
    return out


def validar(data):
    """Contrasta los ratios calculados contra los publicados por OSITRAN."""
    print("VALIDACIÓN vs ratios oficiales OSITRAN (desviación en puntos porcentuales)")
    print("-" * 78)
    maxdev = 0.0
    alertas = []
    for k in ORDEN:
        for y in (2024, 2025):
            of = RATIOS_OFICIALES[k][y]
            r = next(x for x in data[k]["serie"] if x["anio"] == y)
            pares = [("MB", of["mb"], r["margen_bruto"]), ("MO", of["mo"], r["margen_operativo"]),
                     ("MN", of["mn"], r["margen_neto"]), ("EBITDA", of["ebitda"], r["margen_ebitda"]),
                     ("ROE", of["roe"], r["roe"]), ("ROA", of["roa"], r["roa"])]
            devs = []
            for nom, o, c in pares:
                if o is None or c is None:
                    continue
                d = abs(o - c) * 100
                maxdev = max(maxdev, d)
                devs.append(f"{nom} {d:.2f}")
                if d > 1.0:
                    alertas.append(f"{k} {y} {nom}: oficial {o:.3f} vs calc {c:.3f}")
            print(f"{k:6s} {y}  " + "  ".join(devs))
    print("-" * 78)
    print(f"Desviación máxima: {maxdev:.2f} pp")
    if alertas:
        print("\nALERTAS (>1 pp):")
        for a in alertas:
            print("  -", a)
    else:
        print("Sin desviaciones mayores a 1 pp: los EEFF cargados reproducen los ratios de OSITRAN.")
    return maxdev


if __name__ == "__main__":
    data = calcular()
    validar(data)
    with open("datos.json", "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    print("\ndatos.json escrito")
