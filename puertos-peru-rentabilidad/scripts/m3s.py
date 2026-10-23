# -*- coding: utf-8 -*-
"""
Modelo de tres estados integrado (three-statement model) para las siete concesiones
portuarias del Perú, siguiendo la arquitectura de la plantilla de CFI:

  Portada  -> índice y controles del modelo
  Inputs   -> drivers por terminal, con escenarios Alto / Base / Bajo
  <TERMINAL> x7 -> Estado de Resultados, Flujo de Efectivo y Situación Financiera
                   enlazados, más las cédulas de apoyo
  Resumen  -> tablero comparativo
  Notas    -> método, fuentes y limitaciones

Histórico 2021-2025: estados financieros auditados reproducidos por el Ositrán.
Proyección 2026-2030: construida con fórmulas vivas de Excel; al cambiar un input
el modelo entero recalcula.

Convención de color: azul = dato de entrada editable; negro = fórmula calculada.
"""
import json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from dataset import PUERTOS, ORDEN

# Ratios y series por terminal (kpis.py) y puente del EBIT al flujo libre (modelo.py).
DAT = json.load(open("datos.json", encoding="utf-8"))
MDL = json.load(open("modelo.json", encoding="utf-8"))
for _k, _v in MDL.items():
    DAT[_k]["modelo"] = {str(m["anio"]): m for m in _v["modelo"]}
    DAT[_k]["unit"] = _v["unitarios"]
HIST = [2021, 2022, 2023, 2024, 2025]
FC = [2026, 2027, 2028, 2029, 2030]
ALL = HIST + FC
CH = {y: get_column_letter(4 + i) for i, y in enumerate(ALL)}      # D..M
C1, CN = CH[2021], CH[2030]
CF1, CFN = CH[2026], CH[2030]

# ---------------------------------------------------------------- estilo
AZ, MG, OC = "1B5FA8", "C9426E", "B5830F"
INK, INK2, INK3 = "141A20", "4A5560", "7D8893"
BG, BG2, BGI = "F4F1EA", "ECEADF", "EAF1FA"
F_IN = Font(name="Calibri", size=10, color="0C4FA0")       # entrada editable
F_FX = Font(name="Calibri", size=10, color=INK)            # fórmula
F_LB = Font(name="Calibri", size=10, color=INK)
F_SB = Font(name="Calibri", size=10, color=INK2)
F_TT = Font(name="Calibri", size=10, color=INK, bold=True)
F_H1 = Font(name="Calibri", size=16, color=INK, bold=True)
F_H2 = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
F_H3 = Font(name="Calibri", size=10, color=AZ, bold=True)
F_NT = Font(name="Calibri", size=8, color=INK3, italic=True)
FILL_H2 = PatternFill("solid", fgColor=AZ)
FILL_SEC = PatternFill("solid", fgColor=BG2)
FILL_TOT = PatternFill("solid", fgColor=BG)
FILL_IN = PatternFill("solid", fgColor=BGI)
THIN = Side(style="thin", color="D8D2C6")
MED = Side(style="medium", color=AZ)
NUM = '#,##0;[Red](#,##0);"—"'
NUM1 = '#,##0.0;[Red](#,##0.0);"—"'
NUM2 = '#,##0.00;[Red](#,##0.00);"—"'
PCT = '0.0%;[Red](0.0%);"—"'
PCT2 = '0.00%'
XTI = '0.00"×"'


class Hoja:
    """Escritor de hojas con cursor de fila y registro de filas con nombre."""

    def __init__(self, ws):
        self.ws, self.r, self.ref = ws, 1, {}

    def salto(self, n=1):
        self.r += n

    def _set(self, col, val, font=None, fmt=None, fill=None, align=None, border=None):
        c = self.ws[f"{col}{self.r}"]
        if isinstance(val, str) and val.startswith("="):
            c.value = val
            c.font = font or F_FX
        else:
            c.value = val
            c.font = font or F_LB
        if fmt:
            c.number_format = fmt
        if fill:
            c.fill = fill
        if align:
            c.alignment = align
        if border:
            c.border = border
        return c

    def titulo(self, txt, sub=""):
        self._set("B", txt, F_H1)
        if sub:
            self._set("C", "")
            self.r += 1
            self._set("B", sub, F_SB)
        self.r += 2

    def banda(self, txt):
        for col in "BCDEFGHIJKLM":
            self._set(col, txt if col == "B" else None, F_H2, fill=FILL_H2)
        self.ws.row_dimensions[self.r].height = 17
        self.r += 2

    def sub(self, txt):
        self._set("B", txt, F_H3)
        self.r += 1

    def anios(self, hist=True, fc=True):
        self._set("B", "Miles de USD salvo indicación", F_NT)
        for y in ALL:
            if (y in HIST and hist) or (y in FC and fc):
                c = self._set(CH[y], y, F_TT, fmt="0", align=Alignment(horizontal="right"))
                c.border = Border(bottom=THIN)
        self.r += 1
        self._set("B", None)
        for y in ALL:
            if (y in HIST and hist) or (y in FC and fc):
                self._set(CH[y], "Auditado" if y in HIST else "Proyectado", F_NT,
                          align=Alignment(horizontal="right"))
        self.r += 1

    def fila(self, nombre, etq, vals=None, fml=None, fmt=NUM, estilo="n", nota=None, reg=True):
        """vals: dict año->valor (histórico). fml: función año->fórmula (proyección)."""
        f = {"n": F_LB, "s": F_SB, "t": F_TT}[estilo]
        c = self._set("B", ("    " + etq) if estilo == "s" else etq, f)
        if estilo == "t":
            c.fill = FILL_TOT
        if nota:
            self._set("C", nota, F_NT)
        # si la fila trae datos históricos duros, la fórmula solo gobierna la proyección
        solo_fc = bool(vals) and any(y in HIST for y in vals)
        for y in ALL:
            cell = None
            if vals and y in vals and vals[y] is not None:
                cell = self._set(CH[y], vals[y], F_FX if isinstance(vals[y], str) else F_LB, fmt=fmt)
            elif fml and (y in FC or not solo_fc):
                v = fml(y)
                if v is not None:
                    cell = self._set(CH[y], v, F_FX, fmt=fmt)
            if cell is not None and estilo == "t":
                cell.fill = FILL_TOT
                cell.font = Font(name="Calibri", size=10, color=INK, bold=True)
        if reg and nombre:
            self.ref[nombre] = self.r
        self.r += 1
        return self.r - 1

    def entrada(self, nombre, etq, vals, fmt=NUM, nota=None, solo_fc=True):
        self._set("B", etq, F_LB)
        if nota:
            self._set("C", nota, F_NT)
        for y in (FC if solo_fc else ALL):
            v = vals.get(y) if isinstance(vals, dict) else vals
            c = self._set(CH[y], v, F_IN, fmt=fmt, fill=FILL_IN)
            c.border = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
        if nombre:
            self.ref[nombre] = self.r
        self.r += 1
        return self.r - 1

    def nota(self, txt):
        self._set("B", txt, F_NT)
        self.r += 1


def S(k, f, y):
    p = PUERTOS[k]
    src = p["esf"] if f in p["esf"] else p["eerr"]
    return src[f][HIST.index(y)]


def MOD(k, y, f):
    m = DAT[k]["modelo"].get(str(y))
    return m[f] if m else None


def DA(k, y):
    """Depreciación y amortización del ejercicio, donde la fuente permite derivarla."""
    r = next((x for x in DAT[k]["serie"] if x["anio"] == y), None)
    return r["dep_amort"] if r else None


def hist(k, f, tr=None):
    """dict año->valor histórico, opcionalmente transformado."""
    return {y: (tr(S(k, f, y)) if tr else S(k, f, y)) for y in HIST}


# ================================================================ SUPUESTOS
UNID = {"DPW": ("TEU", "teu"), "APMT": ("TM", "carga_tm"), "TPP": ("TM", "carga_tm"),
        "TISUR": ("TM", "carga_tm"), "STI": ("TM", "carga_tm"),
        "TPGSM": ("TM", "carga_tm"), "TCSA": ("TM", "carga_tm")}
PAYOUT = {"APMT": .40, "DPW": .70, "TISUR": .70, "TPP": .30,
          "TPGSM": .00, "STI": .20, "TCSA": .00}
PCT_VAR = {"APMT": .45, "DPW": .35, "TISUR": .50, "TPP": .45,
           "TPGSM": .55, "STI": .50, "TCSA": .30}
TASA_DEUDA = {"APMT": .065, "DPW": .062, "TISUR": .072, "TPP": .108,
              "TPGSM": .073, "STI": .089, "TCSA": .076}
INFLA = {2026: .025, 2027: .025, 2028: .022, 2029: .022, 2030: .020}
TASA_IR, TASA_CAJA, TASA_REVOLV = .295, .030, .095


def sup(k):
    """Supuestos base derivados del histórico de cada terminal."""
    u, campo = UNID[k]
    vol = {y: PUERTOS[k]["trafico"][campo][HIST.index(y)] for y in HIST}
    cagr = (vol[2025] / vol[2021]) ** .25 - 1 if vol[2021] else 0
    g = max(min(cagr, .06), -.01)
    ings = S(k, "ing_servicios", 2025)
    da = MOD(k, 2025, "dep_amort")
    cap25 = MOD(k, 2025, "capex")
    base_cap = max(da, cap25)
    deuda = S(k, "deuda_fin_cp", 2025) + S(k, "deuda_fin_lp", 2025) + S(k, "deuda_relacionadas_lp", 2025)
    rest = PUERTOS[k]["inicio"] + PUERTOS[k]["plazo"] - 2025
    return dict(
        unidad=u, vol25=vol[2025], tarifa25=ings * 1000 / vol[2025], cagr=cagr,
        g_alto=g + .02, g_base=g, g_bajo=max(g - .02, -.02),
        t_alto=.030, t_base=.020, t_bajo=.010,
        cap_alto=base_cap * 1.3, cap_base=base_cap, cap_bajo=base_cap * .7,
        costo_ef25=S(k, "costo_servicios", 2025) - da, gav25=S(k, "gastos_admin", 2025) + S(k, "gastos_venta", 2025),
        da25=da, deuda25=deuda, rest=rest,
        pend=PUERTOS[k]["compromiso_inv"] - PUERTOS[k]["inv_acum"],
        dias_ar=S(k, "ctas_cobrar_com", 2025) * 365 / ings,
        dias_inv=S(k, "inventarios", 2025) * 365 / max(S(k, "costo_servicios", 2025) - da, 1),
        dias_ap=S(k, "cxp_comerciales", 2025) * 365 / max(S(k, "costo_servicios", 2025) - da, 1),
        amort=deuda / 10)


SUP = {k: sup(k) for k in ORDEN}


def hoja_inputs(wb):
    ws = wb.create_sheet("Inputs")
    h = Hoja(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 26
    for y in ALL:
        ws.column_dimensions[CH[y]].width = 12
    ws.sheet_view.showGridLines = False

    h.titulo("Inputs del modelo",
             "Las celdas en azul son editables. Todo lo demás se calcula. "
             "El selector de escenario de la fila 8 gobierna las tres hojas de drivers de cada terminal.")

    h.banda("CONTROL GENERAL")
    h.ws[f"B{h.r}"] = "Escenario activo"
    h.ws[f"B{h.r}"].font = F_TT
    c = h.ws[f"D{h.r}"]
    c.value = 2
    c.font = F_IN
    c.fill = FILL_IN
    c.border = Border(left=MED, right=MED, top=MED, bottom=MED)
    c.alignment = Alignment(horizontal="center")
    h.ws[f"E{h.r}"] = '=CHOOSE($D$' + str(h.r) + ',"ALTO","BASE","BAJO")'
    h.ws[f"E{h.r}"].font = Font(name="Calibri", size=10, bold=True, color=MG)
    h.ws[f"G{h.r}"] = "1 = Alto   ·   2 = Base   ·   3 = Bajo"
    h.ws[f"G{h.r}"].font = F_NT
    dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add(c)
    ESC = f"$D${h.r}"
    h.r += 2

    h.sub("Supuestos macro y fiscales, comunes a los siete terminales")
    h.anios(hist=False)
    h.entrada("inflacion", "Inflación de costos en USD", INFLA, PCT2)
    r_ir = h.entrada("tasa_ir", "Tasa de impuesto a la renta", {y: TASA_IR for y in FC}, PCT2,
                     "Tasa corporativa vigente en Perú")
    r_cj = h.entrada("tasa_caja", "Rendimiento de los saldos de caja", {y: TASA_CAJA for y in FC}, PCT2)
    r_rv = h.entrada("tasa_revolv", "Tasa de la línea revolvente", {y: TASA_REVOLV for y in FC}, PCT2,
                     "Se activa solo si la caja se agota")
    h.nota("Los intereses se calculan sobre el saldo inicial de cada periodo para evitar referencias circulares.")
    h.r += 1

    reg = {"ESC": ESC, "inflacion": h.ref["inflacion"], "tasa_ir": r_ir,
           "tasa_caja": r_cj, "tasa_revolv": r_rv}

    for k in ORDEN:
        s = SUP[k]
        h.banda(f"{PUERTOS[k]['nombre'].upper()}  ·  {PUERTOS[k]['terminal']}")
        h.anios(hist=False)
        rk = {}

        h.sub(f"Crecimiento del volumen ({s['unidad']})")
        h.fila(None, "Escenario activo", fml=lambda y, rr=h.r: None, reg=False)
        h.r -= 1
        ra = h.entrada(None, "Alto", {y: s["g_alto"] for y in FC}, PCT2)
        rb = h.entrada(None, "Base", {y: s["g_base"] for y in FC}, PCT2,
                       f"CAGR 2021-2025 observado: {s['cagr']*100:,.1f} %")
        rc = h.entrada(None, "Bajo", {y: s["g_bajo"] for y in FC}, PCT2)
        rk["g"] = h.fila("g", "Crecimiento aplicado",
                         fml=lambda y: f"=CHOOSE(Inputs!{ESC},{CH[y]}{ra},{CH[y]}{rb},{CH[y]}{rc})",
                         fmt=PCT2, estilo="t")
        h.r += 1

        h.sub("Ajuste de la tarifa media")
        ra = h.entrada(None, "Alto", {y: s["t_alto"] for y in FC}, PCT2)
        rb = h.entrada(None, "Base", {y: s["t_base"] for y in FC}, PCT2,
                       "Mecanismo RPI-X: inflación menos factor de productividad")
        rc = h.entrada(None, "Bajo", {y: s["t_bajo"] for y in FC}, PCT2)
        rk["t"] = h.fila("t", "Ajuste aplicado",
                         fml=lambda y: f"=CHOOSE(Inputs!{ESC},{CH[y]}{ra},{CH[y]}{rb},{CH[y]}{rc})",
                         fmt=PCT2, estilo="t")
        h.r += 1

        h.sub("Inversión en activo fijo (capex)")
        ra = h.entrada(None, "Alto", {y: round(s["cap_alto"]) for y in FC}, NUM)
        rb = h.entrada(None, "Base", {y: round(s["cap_base"]) for y in FC}, NUM,
                       f"Compromiso contractual pendiente: USD {s['pend']:,.0f} M")
        rc = h.entrada(None, "Bajo", {y: round(s["cap_bajo"]) for y in FC}, NUM)
        rk["cap"] = h.fila("cap", "Capex aplicado",
                           fml=lambda y: f"=CHOOSE(Inputs!{ESC},{CH[y]}{ra},{CH[y]}{rb},{CH[y]}{rc})",
                           fmt=NUM, estilo="t")
        h.r += 1

        h.sub("Capital de trabajo, estructura de costos y financiamiento")
        rk["ar"] = h.entrada(None, "Días de cuentas por cobrar", {y: round(s["dias_ar"]) for y in FC}, NUM1,
                             f"2025 observado: {s['dias_ar']:,.0f} días")
        rk["inv"] = h.entrada(None, "Días de inventario", {y: round(s["dias_inv"]) for y in FC}, NUM1,
                              f"2025 observado: {s['dias_inv']:,.0f} días")
        rk["ap"] = h.entrada(None, "Días de cuentas por pagar", {y: round(s["dias_ap"]) for y in FC}, NUM1,
                             f"2025 observado: {s['dias_ap']:,.0f} días")
        rk["pv"] = h.entrada(None, "Porción variable del costo de servicios",
                             {y: PCT_VAR[k] for y in FC}, PCT2,
                             "El resto se trata como costo fijo, indexado a inflación")
        rk["td"] = h.entrada(None, "Tasa de interés de la deuda", {y: TASA_DEUDA[k] for y in FC}, PCT2,
                             "Implícita en el gasto financiero de 2025")
        rk["am"] = h.entrada(None, "Amortización de deuda del año", {y: round(s["amort"]) for y in FC}, NUM,
                             "Importe positivo = repago")
        rk["ap_cap"] = h.entrada(None, "Aportes (retiros) de capital", {y: 0 for y in FC}, NUM)
        rk["po"] = h.entrada(None, "Payout de dividendos", {y: PAYOUT[k] for y in FC}, PCT2,
                             "No se reparte si los resultados acumulados son negativos")
        rk["vu_int"] = h.entrada(None, "Años remanentes de la concesión",
                                 {y: max(s["rest"], 1) for y in FC}, NUM1,
                                 "Determina la amortización del intangible")
        rk["vu_ppe"] = h.entrada(None, "Vida útil de maquinaria y equipo (años)", {y: 10 for y in FC}, NUM1)
        h.r += 1
        reg[k] = rk

    return reg


# ================================================================ HOJA POR TERMINAL
def hoja_puerto(wb, k, reg):
    s, P = SUP[k], PUERTOS[k]
    I, rk = reg, reg[k]
    ws = wb.create_sheet(k)
    h = Hoja(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 46
    ws.column_dimensions["C"].width = 24
    for y in ALL:
        ws.column_dimensions[CH[y]].width = 12
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D1"

    def inp(campo, y):
        return f"Inputs!{CH[y]}${rk[campo]}"

    def gl(campo, y):
        return f"Inputs!{CH[y]}${I[campo]}"

    def pv(y):           # columna del año anterior
        return CH[y - 1] if y > 2021 else CH[2021]

    # ---------------- encabezado
    h.titulo(P["nombre"], f"{P['terminal']} · {P['operador']}")
    for etq, val in [("Tipo de carga", P["tipo"]), ("Modalidad", P["modalidad"]),
                     ("Inicio y plazo", f"{P['inicio']} · {P['plazo']} años"),
                     ("Unidad de trabajo del modelo", s["unidad"]),
                     ("Compromiso de inversión", f"USD {P['compromiso_inv']:,.1f} M · ejecutado {P['avance_inv']:,.1f} %"),
                     ("Vencimiento de la concesión",
                      f"{P['inicio'] + P['plazo']}"
                      + ("  ⚠ dentro del horizonte del modelo: la amortización del intangible se "
                         "acelera y los ejercicios posteriores se muestran solo por continuidad de la serie"
                         if P['inicio'] + P['plazo'] <= 2030 else ""))]:
        h._set("B", etq, F_SB)
        h._set("C", val, F_LB)
        h.r += 1
    h.r += 1

    # ================= CÉDULA DE INGRESOS
    h.banda("CÉDULA DE INGRESOS")
    h.anios()
    vol = {y: P["trafico"][UNID[k][1]][HIST.index(y)] for y in HIST}
    r_vol = h.fila("vol", f"Volumen movilizado ({s['unidad']})", vals=vol,
                   fml=lambda y: f"={pv(y)}{h.r}*(1+{inp('g', y)})", fmt=NUM)
    h.fila("gvol", "Crecimiento del volumen",
           fml=lambda y: f"={CH[y]}{r_vol}/{pv(y)}{r_vol}-1" if y != 2021 else None,
           fmt=PCT, estilo="s")
    r_ing_s = h.fila("ings", "Ingresos por servicios portuarios", vals=hist(k, "ing_servicios"),
                     fml=lambda y: f"={CH[y]}{r_vol}*{CH[y]}{h.r+1}/1000")
    r_tar = h.fila("tarifa", f"Tarifa media implícita (USD/{s['unidad']})",
                   vals={y: f"={CH[y]}{r_ing_s}*1000/{CH[y]}{r_vol}" for y in HIST},
                   fml=lambda y: f"={pv(y)}{h.r}*(1+{inp('t', y)})", fmt=NUM2, estilo="s")
    r_ing_c = h.fila("ingc", "Ingresos por servicios de construcción", vals=hist(k, "ing_construccion"),
                     fml=lambda y: 0, estilo="s",
                     nota="CINIIF 12; sin margen. No se proyecta: el margen de 2026 sube por eso.")
    r_ing_o = h.fila("ingo", "Otros ingresos de actividades ordinarias",
                     vals={y: S(k, "ingresos_totales", y) - S(k, "ing_servicios", y)
                           - S(k, "ing_construccion", y) for y in HIST},
                     fml=lambda y: 0, estilo="s")
    r_ing = h.fila("ing", "Ingresos totales",
                   vals={y: f"=SUM({CH[y]}{r_ing_s},{CH[y]}{r_ing_c},{CH[y]}{r_ing_o})" for y in ALL},
                   estilo="t")
    h.r += 1

    # ================= CÉDULA DE COSTOS
    h.banda("CÉDULA DE COSTOS Y GASTOS")
    h.anios()
    r_cv_u = h.fila("cvu", f"Costo variable unitario (USD/{s['unidad']})",
                    vals={2025: s["costo_ef25"] * PCT_VAR[k] * 1000 / s["vol25"]},
                    fml=lambda y: f"={pv(y)}{h.r}*(1+{gl('inflacion', y)})", fmt=NUM2)
    r_cv = h.fila("cv", "Costo variable total",
                  fml=lambda y: f"={CH[y]}{r_vol}*{CH[y]}{r_cv_u}/1000", estilo="s")
    r_cf = h.fila("cf", "Costo fijo de operación",
                  vals={2025: round(s["costo_ef25"] * (1 - PCT_VAR[k]))},
                  fml=lambda y: f"={pv(y)}{h.r}*(1+{gl('inflacion', y)})", estilo="s")
    r_ce = h.fila("ce", "Costo de servicios en efectivo",
                  vals={y: S(k, "costo_servicios", y) - (DA(k, y) or 0) for y in HIST},
                  fml=lambda y: f"=SUM({CH[y]}{r_cv},{CH[y]}{r_cf})", estilo="t")
    r_gav = h.fila("gav", "Gastos de administración y venta",
                   vals={y: S(k, "gastos_admin", y) + S(k, "gastos_venta", y) for y in HIST},
                   fml=lambda y: f"={pv(y)}{h.r}*(1+{gl('inflacion', y)})")
    h.nota("El costo de servicios en efectivo histórico se obtiene restando al costo reportado la "
           "depreciación y amortización del periodo, derivable solo desde 2024 en la mayoría de los terminales.")
    h.r += 1

    # ================= CÉDULA DE ACTIVO FIJO
    h.banda("CÉDULA DE ACTIVO FIJO Y DEPRECIACIÓN")
    h.anios()
    h.sub("Activos intangibles de la concesión")
    r_int_i = h.fila("int_i", "Saldo inicial",
                     fml=lambda y: f"={pv(y)}{h.r+3}" if y != 2021 else None, estilo="s")
    r_int_a = h.fila("int_a", "Adiciones",
                     fml=lambda y: f"={inp('cap', y)}*0.85" if y in FC else None, estilo="s",
                     nota="85 % del capex se capitaliza como intangible")
    r_int_d = h.fila("int_d", "Amortización del ejercicio",
                     fml=lambda y: f"=-{CH[y]}{r_int_i}/{inp('vu_int', y)}" if y in FC else None, estilo="s")
    r_int = h.fila("int", "Saldo final", vals=hist(k, "intangibles"),
                   fml=lambda y: f"=SUM({CH[y]}{r_int_i}:{CH[y]}{r_int_d})", estilo="t")
    h.sub("Instalaciones, maquinaria y equipo")
    r_ppe_i = h.fila("ppe_i", "Saldo inicial",
                     fml=lambda y: f"={pv(y)}{h.r+3}" if y != 2021 else None, estilo="s")
    r_ppe_a = h.fila("ppe_a", "Adiciones",
                     fml=lambda y: f"={inp('cap', y)}*0.15" if y in FC else None, estilo="s")
    r_ppe_d = h.fila("ppe_d", "Depreciación del ejercicio",
                     fml=lambda y: f"=-{CH[y]}{r_ppe_i}/{inp('vu_ppe', y)}" if y in FC else None, estilo="s")
    r_ppe = h.fila("ppe", "Saldo final", vals=hist(k, "ppe"),
                   fml=lambda y: f"=SUM({CH[y]}{r_ppe_i}:{CH[y]}{r_ppe_d})", estilo="t")
    r_da = h.fila("da", "Depreciación y amortización total",
                  vals={y: DA(k, y) for y in HIST},
                  fml=lambda y: f"=-SUM({CH[y]}{r_int_d},{CH[y]}{r_ppe_d})", estilo="t")
    r_capex = h.fila("capex", "Capex del ejercicio",
                     vals={y: MOD(k, y, "capex") for y in HIST},
                     fml=lambda y: f"={inp('cap', y)}")
    h.r += 1

    # ================= CÉDULA DE CAPITAL DE TRABAJO
    h.banda("CÉDULA DE CAPITAL DE TRABAJO")
    h.anios()
    r_ar = h.fila("ar", "Cuentas por cobrar comerciales", vals=hist(k, "ctas_cobrar_com"),
                  fml=lambda y: f"={inp('ar', y)}/365*{CH[y]}{r_ing_s}")
    r_in = h.fila("inv", "Inventarios y suministros", vals=hist(k, "inventarios"),
                  fml=lambda y: f"={inp('inv', y)}/365*{CH[y]}{r_ce}")
    r_cp = h.fila("cxp", "Cuentas por pagar comerciales", vals=hist(k, "cxp_comerciales"),
                  fml=lambda y: f"={inp('ap', y)}/365*{CH[y]}{r_ce}")
    r_wc = h.fila("wc", "Capital de trabajo comercial",
                  vals={y: f"={CH[y]}{r_ar}+{CH[y]}{r_in}-{CH[y]}{r_cp}" for y in ALL}, estilo="t")
    r_dwc = h.fila("dwc", "Caja generada (consumida) por el capital de trabajo",
                   fml=lambda y: f"=-({CH[y]}{r_wc}-{pv(y)}{r_wc})" if y != 2021 else None, estilo="s")
    h.r += 1

    # ================= CÉDULA DE DEUDA
    h.banda("CÉDULA DE DEUDA Y PATRIMONIO")
    h.anios()
    h.sub("Deuda financiera e intragrupo")
    r_dd_i = h.fila("dd_i", "Saldo inicial",
                    fml=lambda y: f"={pv(y)}{h.r+2}" if y != 2021 else None, estilo="s")
    r_dd_v = h.fila("dd_v", "Amortización del ejercicio",
                    fml=lambda y: f"=-MIN({inp('am', y)},{CH[y]}{r_dd_i})" if y in FC else None, estilo="s")
    r_dd = h.fila("dd", "Saldo final",
                  vals={y: S(k, "deuda_fin_cp", y) + S(k, "deuda_fin_lp", y)
                        + S(k, "deuda_relacionadas_lp", y) for y in HIST},
                  fml=lambda y: f"=SUM({CH[y]}{r_dd_i}:{CH[y]}{r_dd_v})", estilo="t")
    r_int_g = h.fila("int_g", "Gasto financiero sobre la deuda",
                     fml=lambda y: f"={CH[y]}{r_dd_i}*{inp('td', y)}" if y in FC else None, estilo="s")
    h.sub("Línea revolvente de capital de trabajo")
    r_rv_i = h.fila("rv_i", "Saldo inicial",
                    fml=lambda y: f"={pv(y)}{h.r+2}" if y != 2021 else 0, estilo="s")
    r_rv_v = h.fila("rv_v", "Giro (repago) del ejercicio", fml=lambda y: None, estilo="s")
    r_rv = h.fila("rv", "Saldo final",
                  vals={y: 0 for y in HIST},
                  fml=lambda y: f"=SUM({CH[y]}{r_rv_i}:{CH[y]}{r_rv_v})", estilo="t")
    r_int_r = h.fila("int_r", "Gasto financiero de la línea revolvente",
                     fml=lambda y: f"={CH[y]}{r_rv_i}*{gl('tasa_revolv', y)}" if y in FC else None, estilo="s")
    h.sub("Patrimonio")
    r_cap = h.fila("cap_res", "Capital y reservas",
                   vals={y: S(k, "patrimonio", y) - S(k, "resultados_acumulados", y) for y in HIST},
                   fml=lambda y: f"={pv(y)}{h.r}+{inp('ap_cap', y)}")
    r_ra_i = h.fila("ra_i", "Resultados acumulados al inicio",
                    fml=lambda y: f"={pv(y)}{h.r+2}" if y != 2021 else None, estilo="s")
    r_div = h.fila("div", "Dividendos declarados", fml=lambda y: None, estilo="s")
    r_ra = h.fila("ra", "Resultados acumulados al cierre", vals=hist(k, "resultados_acumulados"),
                  estilo="t")   # la fórmula se fija en cerrar(), cuando existe la utilidad neta
    r_pat = h.fila("pat", "Total patrimonio", vals=hist(k, "patrimonio"),
                   fml=lambda y: f"=SUM({CH[y]}{r_cap},{CH[y]}{r_ra})", estilo="t")
    h.r += 1

    # ================= CÉDULA DE IMPUESTOS
    h.banda("CÉDULA DE IMPUESTO A LA RENTA")
    h.anios()
    r_pf_i = h.fila("pf_i", "Pérdidas tributarias acumuladas al inicio",
                    fml=lambda y: f"={pv(y)}{h.r+3}" if y != 2021 else None, estilo="s")
    r_pf_n = h.fila("pf_n", "Pérdidas generadas en el ejercicio", fml=lambda y: None, estilo="s")
    r_pf_u = h.fila("pf_u", "Pérdidas aplicadas en el ejercicio", fml=lambda y: None, estilo="s")
    r_pf = h.fila("pf", "Pérdidas tributarias al cierre",
                  vals={2025: max(0, -S(k, "resultados_acumulados", 2025))},
                  fml=lambda y: f"=SUM({CH[y]}{r_pf_i}:{CH[y]}{r_pf_u})", estilo="t")
    r_bi = h.fila("bi", "Base imponible", fml=lambda y: None, estilo="s")
    r_ir = h.fila("ir", "Impuesto a la renta del ejercicio",
                  vals={y: S(k, "impuesto_renta", y) for y in HIST},
                  fml=lambda y: f"={CH[y]}{r_bi}*{gl('tasa_ir', y)}", estilo="t")
    h.nota("Las pérdidas tributarias de arranque se aproximan con los resultados acumulados negativos al cierre de 2025.")
    h.r += 1

    # ================= ESTADO DE RESULTADOS
    h.banda("ESTADO DE RESULTADOS INTEGRALES")
    h.anios()
    h.fila(None, "Ingresos totales", vals={y: f"={CH[y]}{r_ing}" for y in ALL}, estilo="t", reg=False)
    h.fila(None, "Ingresos por servicios portuarios",
           vals={y: f"={CH[y]}{r_ing_s}" for y in ALL}, estilo="s", reg=False)
    h.fila(None, "Ingresos por servicios de construcción",
           vals={y: f"={CH[y]}{r_ing_c}" for y in ALL}, estilo="s", reg=False)
    h.fila(None, "Otros ingresos de actividades ordinarias",
           vals={y: f"={CH[y]}{r_ing_o}" for y in ALL}, estilo="s", reg=False)
    r_cs = h.fila("cs", "Costo de servicios",
                  vals={y: -S(k, "costos_totales", y) for y in HIST},
                  fml=lambda y: f"=-{CH[y]}{r_ce}-{CH[y]}{r_da}",
                  nota="Histórico incluye el costo de construcción, que no se proyecta")
    r_ub = h.fila("ub", "Utilidad bruta",
                  vals={y: f"=SUM({CH[y]}{r_ing},{CH[y]}{r_cs})" for y in ALL}, estilo="t")
    r_gv = h.fila("gv", "Gastos de administración y venta",
                  vals={y: f"=-{CH[y]}{r_gav}" for y in ALL})
    r_oo = h.fila("oo", "Otros ingresos (gastos) operativos",
                  vals={y: S(k, "utilidad_operativa", y) - (S(k, "utilidad_bruta", y)
                        - S(k, "gastos_admin", y) - S(k, "gastos_venta", y)) for y in HIST},
                  fml=lambda y: 0)
    r_ebit = h.fila("ebit", "Utilidad operativa (EBIT)",
                    vals={y: f"=SUM({CH[y]}{r_ub},{CH[y]}{r_gv},{CH[y]}{r_oo})" for y in ALL}, estilo="t")
    h.fila(None, "Depreciación y amortización", vals={y: f"={CH[y]}{r_da}" for y in ALL},
           estilo="s", reg=False)
    r_ebitda = h.fila("ebitda", "EBITDA",
                      vals={y: f"={CH[y]}{r_ebit}+{CH[y]}{r_da}" for y in ALL}, estilo="t")
    r_if = h.fila("if", "Ingresos financieros", vals=hist(k, "ingresos_financieros"),
                  fml=lambda y: None)
    r_gf = h.fila("gf", "Gastos financieros",
                  vals={y: -S(k, "gastos_financieros", y) for y in HIST},
                  fml=lambda y: f"=-{CH[y]}{r_int_g}-{CH[y]}{r_int_r}")
    r_dc = h.fila("dc", "Diferencia en cambio y otros",
                  vals={y: S(k, "uai", y) - (S(k, "utilidad_operativa", y)
                        + S(k, "ingresos_financieros", y) - S(k, "gastos_financieros", y)) for y in HIST},
                  fml=lambda y: 0)
    r_uai = h.fila("uai", "Utilidad antes de impuestos",
                   vals={y: f"=SUM({CH[y]}{r_ebit},{CH[y]}{r_if},{CH[y]}{r_gf},{CH[y]}{r_dc})" for y in ALL},
                   estilo="t")
    r_imp = h.fila("imp", "Impuesto a la renta",
                   vals={y: -S(k, "impuesto_renta", y) for y in HIST},
                   fml=lambda y: f"=-{CH[y]}{r_ir}")
    r_un = h.fila("un", "Utilidad neta",
                  vals={y: f"=SUM({CH[y]}{r_uai},{CH[y]}{r_imp})" for y in ALL}, estilo="t")
    h.r += 1
    h.sub("Márgenes")
    for etq, rr in [("Margen bruto", r_ub), ("Margen EBITDA", r_ebitda),
                    ("Margen operativo", r_ebit), ("Margen neto", r_un)]:
        h.fila(None, etq, vals={y: f"={CH[y]}{rr}/{CH[y]}{r_ing}" for y in ALL},
               fmt=PCT, estilo="s", reg=False)
    h.r += 1
    refs = dict(ing=r_ing, ings=r_ing_s, ebit=r_ebit, ebitda=r_ebitda, un=r_un, uai=r_uai,
                ub=r_ub, cs=r_cs, gav=r_gav, ce=r_ce,
                da=r_da, capex=r_capex, dwc=r_dwc, ar=r_ar, inv=r_in, cxp=r_cp, wc=r_wc,
                dd=r_dd, dd_i=r_dd_i, dd_v=r_dd_v, rv=r_rv, rv_i=r_rv_i, rv_v=r_rv_v,
                pat=r_pat, ra=r_ra, ra_i=r_ra_i, div=r_div, cap=r_cap, int=r_int, ppe=r_ppe,
                imp=r_imp, ir=r_ir, bi=r_bi, pf=r_pf, pf_i=r_pf_i, pf_n=r_pf_n, pf_u=r_pf_u,
                gf=r_gf, iff=r_if, int_g=r_int_g, int_r=r_int_r, vol=r_vol, tar=r_tar)
    return h, ws, refs, inp, gl, pv


def estados(h, ws, k, R, inp, gl, pv, reg):
    """Flujo de efectivo y situación financiera, y cierre de las referencias cruzadas."""
    s = SUP[k]
    I = reg

    # ================= FLUJO DE EFECTIVO
    h.banda("ESTADO DE FLUJOS DE EFECTIVO")
    h.anios()
    h.sub("Actividades de operación")
    f_un = h.fila(None, "Utilidad neta", vals={y: f"={CH[y]}{R['un']}" for y in ALL}, estilo="s", reg=False)
    f_da = h.fila(None, "Depreciación y amortización", vals={y: f"={CH[y]}{R['da']}" for y in ALL},
                  estilo="s", reg=False)
    f_wc = h.fila(None, "Variación del capital de trabajo comercial",
                  vals={y: f"={CH[y]}{R['dwc']}" for y in ALL if y != 2021}, estilo="s", reg=False)
    f_ot = h.fila("ot", "Otras partidas operativas", fml=lambda y: 0 if y in FC else None, estilo="s",
                  nota="En el histórico absorbe lo no desagregado por la fuente")
    f_op = h.fila("op", "Flujo de actividades de operación",
                  vals={y: f"=SUM({CH[y]}{f_un}:{CH[y]}{f_ot})" for y in ALL}, estilo="t")
    h.sub("Actividades de inversión")
    f_cx = h.fila(None, "Inversión en activo fijo", vals={y: f"=-{CH[y]}{R['capex']}" for y in ALL},
                  estilo="s", reg=False)
    f_iv = h.fila("iv", "Flujo de actividades de inversión",
                  vals={y: f"={CH[y]}{f_cx}" for y in ALL}, estilo="t")
    h.sub("Actividades de financiamiento")
    f_dd = h.fila(None, "Variación de la deuda financiera e intragrupo",
                  vals={y: f"={CH[y]}{R['dd']}-{pv(y)}{R['dd']}" for y in ALL if y != 2021},
                  estilo="s", reg=False)
    f_cp = h.fila(None, "Aportes (retiros) de capital",
                  vals={y: f"={CH[y]}{R['cap']}-{pv(y)}{R['cap']}" for y in ALL if y != 2021},
                  estilo="s", reg=False)
    f_dv = h.fila(None, "Dividendos pagados", vals={y: f"={CH[y]}{R['div']}" for y in ALL},
                  estilo="s", reg=False)
    f_fs = h.fila("fs", "Subtotal antes de la línea revolvente",
                  vals={y: f"=SUM({CH[y]}{f_dd}:{CH[y]}{f_dv})" for y in ALL}, estilo="s")
    f_rv = h.fila(None, "Giro (repago) de la línea revolvente",
                  vals={y: f"={CH[y]}{R['rv_v']}" for y in FC}, estilo="s", reg=False)
    f_fi = h.fila("fi", "Flujo de actividades de financiamiento",
                  vals={y: f"=SUM({CH[y]}{f_fs},{CH[y]}{f_rv})" for y in ALL}, estilo="t")
    h.sub("Variación de la caja")
    f_ci = h.fila("ci", "Caja al inicio del ejercicio", estilo="s")
    f_dp = h.fila("disp", "Caja disponible antes de la línea revolvente",
                  vals={y: f"=SUM({CH[y]}{f_ci},{CH[y]}{f_op},{CH[y]}{f_iv},{CH[y]}{f_fs})"
                        for y in FC}, estilo="s",
                  nota="Rompe la circularidad: la revolvente se dimensiona sobre este saldo")
    f_cv = h.fila("cv", "Aumento (disminución) de la caja",
                  vals={y: f"=SUM({CH[y]}{f_op},{CH[y]}{f_iv},{CH[y]}{f_fi})" for y in ALL}, estilo="s")
    f_cf = h.fila("cf", "Caja al cierre del ejercicio",
                  vals={y: S(k, "efectivo", y) for y in HIST},
                  fml=lambda y: f"=SUM({CH[y]}{f_ci},{CH[y]}{f_cv})", estilo="t")
    # la caja inicial es el cierre del año anterior; se fija ahora que f_cf es conocido
    for y in ALL:
        if y == 2021:
            continue
        c = ws[f"{CH[y]}{f_ci}"]
        c.value = f"={pv(y)}{f_cf}"
        c.font = F_FX
        c.number_format = NUM
    h.r += 1

    # ================= SITUACIÓN FINANCIERA
    h.banda("ESTADO DE SITUACIÓN FINANCIERA")
    h.anios()
    h.sub("Activo")
    b_ca = h.fila(None, "Efectivo y equivalentes", vals={y: f"={CH[y]}{f_cf}" for y in ALL},
                  estilo="s", reg=False)
    b_fr = h.fila(None, "Fondos restringidos en fideicomiso", vals=hist(k, "fondos_restringidos"),
                  fml=lambda y: f"={pv(y)}{h.r}", estilo="s", reg=False)
    b_ar = h.fila(None, "Cuentas por cobrar comerciales", vals={y: f"={CH[y]}{R['ar']}" for y in ALL},
                  estilo="s", reg=False)
    b_in = h.fila(None, "Inventarios y suministros", vals={y: f"={CH[y]}{R['inv']}" for y in ALL},
                  estilo="s", reg=False)
    oac = {y: S(k, "activo_corriente", y) - S(k, "efectivo", y) - S(k, "fondos_restringidos", y)
           - S(k, "ctas_cobrar_com", y) - S(k, "inventarios", y) for y in HIST}
    b_oa = h.fila(None, "Otros activos corrientes", vals=oac, fml=lambda y: f"={pv(y)}{h.r}",
                  estilo="s", reg=False)
    b_ac = h.fila("ac", "Total activo corriente",
                  vals={y: f"=SUM({CH[y]}{b_ca}:{CH[y]}{b_oa})" for y in ALL}, estilo="t")
    b_it = h.fila(None, "Activos intangibles de la concesión", vals={y: f"={CH[y]}{R['int']}" for y in ALL},
                  estilo="s", reg=False)
    b_pp = h.fila(None, "Instalaciones, maquinaria y equipo", vals={y: f"={CH[y]}{R['ppe']}" for y in ALL},
                  estilo="s", reg=False)
    oanc = {y: S(k, "activo_no_corriente", y) - S(k, "intangibles", y) - S(k, "ppe", y) for y in HIST}
    b_on = h.fila(None, "Otros activos no corrientes", vals=oanc, fml=lambda y: f"={pv(y)}{h.r}",
                  estilo="s", reg=False)
    b_an = h.fila("an", "Total activo no corriente",
                  vals={y: f"=SUM({CH[y]}{b_it}:{CH[y]}{b_on})" for y in ALL}, estilo="t")
    b_at = h.fila("at", "TOTAL ACTIVO", vals={y: f"={CH[y]}{b_ac}+{CH[y]}{b_an}" for y in ALL}, estilo="t")
    h.sub("Pasivo")
    b_cx = h.fila(None, "Cuentas por pagar comerciales", vals={y: f"={CH[y]}{R['cxp']}" for y in ALL},
                  estilo="s", reg=False)
    b_cr = h.fila(None, "Cuentas por pagar a relacionadas, corriente", vals=hist(k, "cxp_relacionadas"),
                  fml=lambda y: f"={pv(y)}{h.r}", estilo="s", reg=False)
    b_rv = h.fila(None, "Línea revolvente", vals={y: f"={CH[y]}{R['rv']}" for y in ALL},
                  estilo="s", reg=False)
    b_dc = h.fila(None, "Deuda financiera, porción corriente", vals=hist(k, "deuda_fin_cp"),
                  fml=lambda y: (f"=MIN({inp('am', min(y+1, 2030))},{CH[y]}{R['dd']})"), estilo="s", reg=False)
    opc = {y: S(k, "pasivo_corriente", y) - S(k, "cxp_comerciales", y) - S(k, "cxp_relacionadas", y)
           - S(k, "deuda_fin_cp", y) for y in HIST}
    b_op = h.fila(None, "Otros pasivos corrientes", vals=opc, fml=lambda y: f"={pv(y)}{h.r}",
                  estilo="s", reg=False)
    b_pc = h.fila("pc", "Total pasivo corriente",
                  vals={y: f"=SUM({CH[y]}{b_cx}:{CH[y]}{b_op})" for y in ALL}, estilo="t")
    b_dl = h.fila(None, "Deuda financiera e intragrupo, porción no corriente",
                  vals={y: S(k, "deuda_fin_lp", y) + S(k, "deuda_relacionadas_lp", y) for y in HIST},
                  fml=lambda y: f"={CH[y]}{R['dd']}-{CH[y]}{b_dc}", estilo="s", reg=False)
    opnc = {y: S(k, "pasivo_no_corriente", y) - S(k, "deuda_fin_lp", y)
            - S(k, "deuda_relacionadas_lp", y) for y in HIST}
    b_ol = h.fila(None, "Otros pasivos no corrientes", vals=opnc, fml=lambda y: f"={pv(y)}{h.r}",
                  estilo="s", reg=False)
    b_pn = h.fila("pn", "Total pasivo no corriente",
                  vals={y: f"={CH[y]}{b_dl}+{CH[y]}{b_ol}" for y in ALL}, estilo="t")
    b_pt = h.fila("pt", "TOTAL PASIVO", vals={y: f"={CH[y]}{b_pc}+{CH[y]}{b_pn}" for y in ALL}, estilo="t")
    h.sub("Patrimonio")
    b_k1 = h.fila(None, "Capital y reservas", vals={y: f"={CH[y]}{R['cap']}" for y in ALL},
                  estilo="s", reg=False)
    b_k2 = h.fila(None, "Resultados acumulados", vals={y: f"={CH[y]}{R['ra']}" for y in ALL},
                  estilo="s", reg=False)
    b_pa = h.fila("pa", "TOTAL PATRIMONIO", vals={y: f"={CH[y]}{b_k1}+{CH[y]}{b_k2}" for y in ALL},
                  estilo="t")
    b_tp = h.fila("tp", "TOTAL PASIVO Y PATRIMONIO",
                  vals={y: f"={CH[y]}{b_pt}+{CH[y]}{b_pa}" for y in ALL}, estilo="t")
    b_ck = h.fila("ck", "Control: activo menos pasivo y patrimonio",
                  vals={y: f"=ROUND({CH[y]}{b_at}-{CH[y]}{b_tp},0)" for y in ALL}, estilo="t",
                  nota="Debe ser cero en todos los ejercicios")
    for y in ALL:
        ws[f"{CH[y]}{b_ck}"].font = Font(name="Calibri", size=10, bold=True, color="0F7A2E")
    h.r += 1
    R.update(at=b_at, pc=b_pc, pt=b_pt, pa=b_pa, ck=b_ck, ac=b_ac, caja=f_cf, op=f_op, iv=f_iv,
             fi=f_fi, fs=f_fs, ot=f_ot, cv=f_cv, ci=f_ci, disp=f_dp, dc=b_dc)
    return R


def cerrar(ws, k, R, inp, gl, pv):
    """Rellena las fórmulas que dependen de filas definidas más abajo en la hoja."""
    def put(row, y, f):
        c = ws[f"{CH[y]}{row}"]
        c.value = f
        c.font = F_FX
        c.number_format = NUM

    for y in ALL:
        # --- resultados acumulados: arrastre con utilidad neta y dividendos
        if y in FC:
            put(R["ra"], y, f"=SUM({CH[y]}{R['ra_i']},{CH[y]}{R['un']},{CH[y]}{R['div']})")
        # --- dividendos: histórico por diferencia en resultados acumulados; proyección por payout
        if y == 2021:
            pass
        elif y in HIST:
            put(R["div"], y, f"={CH[y]}{R['ra']}-{pv(y)}{R['ra']}-{CH[y]}{R['un']}")
        else:
            put(R["div"], y, f"=-MAX(0,MIN({CH[y]}{R['un']}*{inp('po', y)},"
                             f"{CH[y]}{R['ra_i']}+{CH[y]}{R['un']}))")
        # --- ingresos financieros sobre la caja inicial
        if y in FC:
            put(R["iff"], y, f"={CH[y]}{R['ci']}*{gl('tasa_caja', y)}")
        # --- impuesto a la renta con arrastre de pérdidas
        if y in FC:
            put(R["pf_n"], y, f"=MAX(0,-{CH[y]}{R['uai']})")
            put(R["pf_u"], y, f"=-MIN({CH[y]}{R['pf_i']},MAX(0,{CH[y]}{R['uai']}))")
            put(R["bi"], y, f"=MAX(0,{CH[y]}{R['uai']}+{CH[y]}{R['pf_u']})")
        # --- línea revolvente: se gira cuando la caja no alcanza y se repaga con el excedente
        if y in FC:
            antes = f"{CH[y]}{R['disp']}"
            put(R["rv_v"], y, f"=MAX(0,-{antes})-MIN({CH[y]}{R['rv_i']},MAX(0,{antes}))")
    # pérdidas tributarias de arranque en el primer año histórico de la cédula
    ws[f"{CH[2025]}{R['pf']}"].value = max(0, -S(k, "resultados_acumulados", 2025))
    ws[f"{CH[2025]}{R['pf']}"].font = F_IN
    ws[f"{CH[2025]}{R['pf']}"].fill = FILL_IN


# ================================================================ RESUMEN
def hoja_resumen(wb, refs):
    ws = wb.create_sheet("Resumen")
    h = Hoja(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 20
    for y in ALL:
        ws.column_dimensions[CH[y]].width = 12
    ws.sheet_view.showGridLines = False
    h.titulo("Resumen comparativo",
             "Todas las cifras se enlazan a las hojas de cada terminal. Cambie el escenario en Inputs "
             "y este tablero se recalcula.")
    bloques = [("INGRESOS TOTALES", "ing"), ("EBITDA", "ebitda"), ("MARGEN EBITDA", "_meb"),
               ("UTILIDAD NETA", "un"), ("FLUJO DE OPERACIÓN", "op"),
               ("INVERSIÓN EN ACTIVO FIJO", "capex"), ("FLUJO DE CAJA LIBRE", "_fcf"),
               ("DEUDA FINANCIERA E INTRAGRUPO", "dd"), ("DEUDA NETA / EBITDA", "_dn"),
               ("TOTAL ACTIVO", "at"), ("TOTAL PATRIMONIO", "pa"),
               ("RENTABILIDAD DEL ACTIVO (ROA)", "_roa"),
               ("CONTROL: BALANCE CUADRADO", "ck")]
    for tit, campo in bloques:
        h.banda(tit)
        h.anios()
        pctfmt = campo in ("_meb", "_roa")
        for k in ORDEN:
            R = refs[k]
            if campo == "_meb":
                f = lambda y, k=k, R=R: f"={k}!{CH[y]}{R['ebitda']}/{k}!{CH[y]}{R['ing']}"
            elif campo == "_fcf":
                f = lambda y, k=k, R=R: f"={k}!{CH[y]}{R['op']}+{k}!{CH[y]}{R['iv']}"
            elif campo == "_dn":
                f = lambda y, k=k, R=R: (f"=IFERROR(({k}!{CH[y]}{R['dd']}-{k}!{CH[y]}{R['caja']})"
                                         f"/{k}!{CH[y]}{R['ebitda']},\"\")")
            elif campo == "_roa":
                f = lambda y, k=k, R=R: f"={k}!{CH[y]}{R['un']}/{k}!{CH[y]}{R['at']}"
            else:
                f = lambda y, k=k, R=R, c=campo: f"={k}!{CH[y]}{R[c]}"
            h.fila(None, PUERTOS[k]["nombre"],
                   vals={y: f(y) for y in ALL},
                   fmt=PCT if pctfmt else (XTI if campo == "_dn" else NUM), reg=False)
        if campo in ("ing", "ebitda", "un", "op", "capex", "_fcf", "at", "pa", "dd"):
            h.fila(None, "Total del sistema",
                   vals={y: f"=SUM({CH[y]}{h.r-7}:{CH[y]}{h.r-1})" for y in ALL}, estilo="t", reg=False)
        h.r += 1
    return ws


# ================================================================ PORTADA Y NOTAS
def hoja_portada(wb, refs):
    ws = wb.create_sheet("Portada", 0)
    h = Hoja(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 50
    ws.column_dimensions["C"].width = 16
    for c in "DEFGHIJKLM":
        ws.column_dimensions[c].width = 12
    ws.sheet_view.showGridLines = False
    h.r = 3
    h._set("B", "Modelo de tres estados integrado", F_H1)
    h.r += 1
    h._set("B", "Concesiones portuarias del Perú · histórico auditado 2021-2025 y proyección 2026-2030", F_SB)
    h.r += 1
    h._set("B", "Fuente: Informes de Desempeño 2025 del Ositrán. Cifras en miles de USD.", F_NT)
    h.r += 3

    h.sub("Contenido")
    NOM = {"APMT": "APM Terminals Callao · Muelle Norte",
           "DPW": "DP World Callao · Muelle Sur",
           "TISUR": "Terminal Internacional del Sur · Matarani",
           "TPP": "Terminales Portuarios Euroandinos · Paita",
           "TPGSM": "Terminal Portuario Paracas · General San Martín, Pisco",
           "STI": "Salaverry Terminal Internacional",
           "TCSA": "Transportadora Callao · concentrados de minerales"}
    items = [("", "FICHAS POR TERMINAL — todo lo encontrado sobre cada concesión")]
    items += [(f"Ficha {k}", NOM[k]) for k in ORDEN]
    items += [("", ""), ("", "COMPARATIVOS"),
              ("Resumen", "Tablero comparativo de los siete terminales"),
              ("Indicadores", "Ratios financieros, histórico y proyección"),
              ("Unitarios", "Economía unitaria por segmento de operación"),
              ("", ""), ("", "MODELO DE TRES ESTADOS PROYECTADO"),
              ("Inputs", "Drivers por terminal con escenarios Alto, Base y Bajo")]
    items += [(k, NOM[k]) for k in ORDEN]
    items += [("", ""), ("Notas", "Método, fuentes y limitaciones")]
    for nm, desc in items:
        if not nm:
            if desc:
                h._set("B", desc, F_H3)
            h.r += 1
            continue
        ref = f"'{nm}'!A1" if " " in nm else f"{nm}!A1"
        c = h._set("B", nm, Font(name="Calibri", size=10, color="0C4FA0", underline="single"))
        c.hyperlink = f"#{ref}"
        h._set("C", desc, F_SB)
        h.r += 1
    h.r += 2

    h.sub("Controles del modelo")
    h._set("B", "El balance debe cuadrar en los diez ejercicios de cada terminal y la caja no debe "
                "quedar negativa. Un cero en cada fila indica que el modelo es consistente.", F_NT)
    h.r += 2
    h._set("B", "Terminal", F_TT)
    h._set("C", "Balance descuadrado", F_TT)
    h._set("D", "Caja negativa", F_TT)
    h._set("E", "Patrimonio negativo", F_TT)
    h.r += 1
    f0 = h.r
    for k in ORDEN:
        R = refs[k]
        h._set("B", PUERTOS[k]["nombre"], F_LB)
        rg = lambda c: f"{k}!${C1}${R[c]}:${CN}${R[c]}"
        for col, f in [("C", f"=SUMPRODUCT(ABS({rg('ck')}))"),
                       ("D", f"=SUMPRODUCT(--({rg('caja')}<0))"),
                       ("E", f"=SUMPRODUCT(--({rg('pa')}<0))")]:
            c = h._set(col, f, fmt="0", align=Alignment(horizontal="center"))
            c.font = Font(name="Calibri", size=10, bold=True, color=INK)
        h.r += 1
    f1 = h.r - 1
    h._set("B", "Total de excepciones", F_TT, fill=FILL_TOT)
    for col in "CDE":
        c = h._set(col, f"=SUM({col}{f0}:{col}{f1})", fmt="0", fill=FILL_TOT,
                   align=Alignment(horizontal="center"))
        c.font = Font(name="Calibri", size=11, bold=True, color="0F7A2E")
    h.r += 3
    h._set("B", "Sin referencias circulares: los intereses de deuda, caja y línea revolvente se "
                "calculan sobre saldos iniciales, de modo que el libro no requiere cálculo iterativo.", F_NT)
    h.r += 1
    h._set("B", "Convención de color: azul = dato de entrada editable; negro = fórmula calculada.", F_NT)
    return ws


def hoja_notas(wb):
    ws = wb.create_sheet("Notas")
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 118
    ws.sheet_view.showGridLines = False
    txt = [
        ("H1", "Método, fuentes y limitaciones"), ("", ""),
        ("H3", "Fuentes"),
        ("", "Los estados financieros históricos 2021-2025 proceden de los estados auditados de cada concesionario, "
             "reproducidos por el Ositrán en su Informe de Desempeño 2025. El tráfico procede de la Declaración "
             "Estadística del Ositrán. Las cifras de inversión comprometida y ejecutada provienen de los reportes de "
             "inversiones valorizadas en los contratos de concesión, a diciembre de 2025."),
        ("", "Los ratios calculados reproducen los publicados por el regulador para 2024 y 2025 con una desviación "
             "máxima de 0,40 puntos porcentuales, atribuible a redondeo."),
        ("", ""),
        ("H3", "Arquitectura del modelo"),
        ("", "Cada hoja de terminal contiene las cédulas de apoyo (ingresos, costos, activo fijo, capital de trabajo, "
             "deuda, patrimonio e impuestos) y, a partir de ellas, los tres estados enlazados. El estado de resultados "
             "alimenta el flujo de efectivo; el flujo determina la caja; la caja cierra el balance. La fila de control "
             "del balance debe dar cero en los diez ejercicios."),
        ("", "No hay referencias circulares. Los intereses de la deuda, de la caja y de la línea revolvente se calculan "
             "sobre los saldos iniciales de cada periodo, de modo que el libro no requiere activar el cálculo iterativo "
             "de Excel."),
        ("", "La línea revolvente actúa como variable de ajuste: se gira cuando el flujo del ejercicio dejaría la caja "
             "en negativo y se repaga con el excedente de caja disponible."),
        ("", ""),
        ("H3", "Unidad de trabajo"),
        ("", "El modelo de ingresos de cada terminal se construye sobre la unidad física que corresponde a su operación. "
             "DP World se modela en TEU por tratarse de un terminal de contenedores puro. Los demás se modelan en "
             "toneladas. Los unitarios de un terminal no son comparables con los de otro que atiende una clase distinta "
             "de carga: un contenedor de veinte pies pesa una fracción de lo que pesa el mismo espacio de muelle "
             "ocupado por granel."),
        ("", ""),
        ("H3", "Tratamiento de partidas específicas"),
        ("", "Ingresos por servicios de construcción. Se reconocen bajo la CINIIF 12 contra un costo casi idéntico, no "
             "generan margen y dependen del cronograma de obra. Se muestran separados y no se proyectan."),
        ("", "Depreciación y amortización. Para DP World se usa la desagregación publicada en su informe. Para los demás "
             "terminales se deriva del margen EBITDA que publica el Ositrán, disponible desde 2024, y desde 2021 en el "
             "caso de Matarani. Donde no es derivable sin estimación, la celda queda vacía."),
        ("", "Capex histórico. Se deriva del balance como variación del activo fijo neto más la depreciación y "
             "amortización del periodo. Difiere de la inversión que reporta el regulador, que se valoriza según contrato "
             "e incluye IGV y obra en curso."),
        ("", "Deuda intragrupo. En Matarani y Salaverry una parte relevante del pasivo son cuentas por pagar a entidades "
             "relacionadas de largo plazo que devengan interés. El modelo las trata como deuda, no como capital de "
             "trabajo: en Matarani el gasto financiero de 2025 solo tiene sentido contra ese saldo."),
        ("", "Pérdidas tributarias. El saldo de arranque se aproxima con los resultados acumulados negativos al cierre de "
             "2025, y se aplica contra la utilidad antes de impuestos de los ejercicios proyectados."),
        ("", ""),
        ("H3", "Limitaciones declaradas"),
        ("", "El Ositrán no publica el estado de flujos de efectivo de las concesiones. El flujo histórico se reconstruye "
             "por el método indirecto a partir de la variación del balance, y la línea «Otras partidas operativas» "
             "absorbe lo que la fuente no desagrega, de modo que el flujo cuadra con la variación real de la caja."),
        ("", "En Paita una parte relevante de la inversión de la concesión se reconoce como activo financiero y no como "
             "intangible, por el modelo bifurcado que admite la CINIIF 12 cuando el concedente garantiza un ingreso "
             "mínimo. El capex derivado del balance no captura ese movimiento y, por tanto, lo subestima."),
        ("", "La apertura entre costo variable y costo fijo no la publica el regulador salvo en DP World y Matarani. El "
             "modelo la trata como un input editable, con un valor inicial razonable para cada terminal."),
        ("", "Los escenarios de volumen, tarifa y capex son supuestos del modelo, no proyecciones del concesionario ni "
             "del regulador. El escenario base parte del crecimiento observado entre 2021 y 2025 y de un ajuste "
             "tarifario coherente con el mecanismo RPI-X."),
        ("", "El modelo no incorpora la competencia de terminales ajenos al régimen de concesión del Ositrán, cuyo "
             "efecto sobre tarifa y volumen debe introducirse manualmente en los drivers."),
    ]
    r = 3
    for tag, t in txt:
        c = ws.cell(row=r, column=2, value=t)
        c.font = {"H1": F_H1, "H3": F_H3, "": Font(name="Calibri", size=10, color=INK2)}[tag]
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if tag == "":
            ws.row_dimensions[r].height = max(14, 13 * (len(t) // 118 + 1))
        r += 1 if tag else 1
    return ws


# ================================================================ CONSTRUCCIÓN
def pulir(wb):
    """Colores de pestaña, formato condicional de los controles y configuración de impresión."""
    from openpyxl.formatting.rule import CellIsRule
    verde = PatternFill("solid", fgColor="D8EFDD")
    rojo = PatternFill("solid", fgColor="F8D7D5")
    tabs = {"Portada": AZ, "Inputs": OC, "Resumen": MG, "Indicadores": MG, "Unitarios": MG,
            "Notas": "7D8893", **{f"Ficha {k}": "1B5FA8" for k in ORDEN}}
    for ws in wb.worksheets:
        ws.sheetView = ws.sheet_view
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.sheet_properties.tabColor = tabs.get(ws.title, "4A5560")
        if ws.title in tabs:
            if ws.title.startswith("Ficha "):
                ws.freeze_panes = "D1"
            continue
        for r in range(1, ws.max_row + 1):
            if str(ws[f"B{r}"].value).startswith("Control: activo"):
                rng = f"{C1}{r}:{CN}{r}"
                ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=["0"], fill=verde))
                ws.conditional_formatting.add(rng, CellIsRule(operator="notEqual", formula=["0"], fill=rojo))
                break
    for nm in ("Portada", "Inputs", "Resumen", "Indicadores", "Unitarios"):
        wb[nm].freeze_panes = "A1" if nm == "Portada" else "D1"


def construir(destino="Modelo_3Estados_Puertos.xlsx"):
    wb = Workbook()
    wb.remove(wb.active)
    reg = hoja_inputs(wb)
    refs = {}
    for k in ORDEN:
        h, ws, R, inp, gl, pv = hoja_puerto(wb, k, reg)
        R = estados(h, ws, k, R, inp, gl, pv, reg)
        cerrar(ws, k, R, inp, gl, pv)
        refs[k] = R
    for k in ORDEN:
        hoja_ficha(wb, k)
    hoja_resumen(wb, refs)
    hoja_indicadores(wb, refs)
    hoja_unitarios(wb, refs)
    hoja_notas(wb)
    hoja_portada(wb, refs)
    pulir(wb)
    wb._sheets = [wb[n] for n in ["Portada"] + [f"Ficha {k}" for k in ORDEN]
              + ["Resumen", "Indicadores", "Unitarios", "Inputs"] + ORDEN + ["Notas"]]
    wb.active = 0
    wb.save(destino)
    return destino, refs




# ================================================================ INDICADORES Y UNITARIOS
def hoja_indicadores(wb, refs):
    """Ratios calculados sobre las hojas de cada terminal, histórico y proyección."""
    ws = wb.create_sheet("Indicadores")
    h = Hoja(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 18
    for y in ALL:
        ws.column_dimensions[CH[y]].width = 12
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D1"
    h.titulo("Indicadores financieros",
             "Calculados sobre las hojas de cada terminal. Los ratios de 2024 y 2025 reproducen los "
             "publicados por el Ositrán con una desviación máxima de 0,40 puntos porcentuales.")
    grupos = [
        ("RENTABILIDAD", [
            ("Margen bruto", lambda k, R, y: f"={k}!{CH[y]}{R['ub']}/{k}!{CH[y]}{R['ing']}", PCT),
            ("Margen EBITDA", lambda k, R, y: f"={k}!{CH[y]}{R['ebitda']}/{k}!{CH[y]}{R['ing']}", PCT),
            ("Margen operativo", lambda k, R, y: f"={k}!{CH[y]}{R['ebit']}/{k}!{CH[y]}{R['ing']}", PCT),
            ("Margen neto", lambda k, R, y: f"={k}!{CH[y]}{R['un']}/{k}!{CH[y]}{R['ing']}", PCT),
            ("ROA · utilidad neta sobre activo",
             lambda k, R, y: f"={k}!{CH[y]}{R['un']}/{k}!{CH[y]}{R['at']}", PCT),
            ("ROE · utilidad neta sobre patrimonio",
             lambda k, R, y: f"=IFERROR({k}!{CH[y]}{R['un']}/{k}!{CH[y]}{R['pa']},\"\")", PCT),
            ("EBITDA sobre activo total",
             lambda k, R, y: f"={k}!{CH[y]}{R['ebitda']}/{k}!{CH[y]}{R['at']}", PCT)]),
        ("LIQUIDEZ Y SOLVENCIA", [
            ("Liquidez corriente", lambda k, R, y: f"={k}!{CH[y]}{R['ac']}/{k}!{CH[y]}{R['pc']}", XTI),
            ("Prueba ácida",
             lambda k, R, y: f"=({k}!{CH[y]}{R['ac']}-{k}!{CH[y]}{R['inv']})/{k}!{CH[y]}{R['pc']}", XTI),
            ("Apalancamiento · activo sobre patrimonio",
             lambda k, R, y: f"=IFERROR({k}!{CH[y]}{R['at']}/{k}!{CH[y]}{R['pa']},\"\")", XTI),
            ("Pasivo sobre patrimonio",
             lambda k, R, y: f"=IFERROR({k}!{CH[y]}{R['pt']}/{k}!{CH[y]}{R['pa']},\"\")", XTI),
            ("Endeudamiento del activo",
             lambda k, R, y: f"={k}!{CH[y]}{R['pt']}/{k}!{CH[y]}{R['at']}", PCT),
            ("Deuda bruta sobre EBITDA",
             lambda k, R, y: f"=IFERROR(({k}!{CH[y]}{R['dd']}+{k}!{CH[y]}{R['rv']})"
                             f"/{k}!{CH[y]}{R['ebitda']},\"\")", XTI),
            ("Deuda neta sobre EBITDA",
             lambda k, R, y: f"=IFERROR(({k}!{CH[y]}{R['dd']}+{k}!{CH[y]}{R['rv']}-{k}!{CH[y]}{R['caja']})"
                             f"/{k}!{CH[y]}{R['ebitda']},\"\")", XTI),
            ("Cobertura de intereses",
             lambda k, R, y: f"=IFERROR(-{k}!{CH[y]}{R['ebitda']}/{k}!{CH[y]}{R['gf']},\"\")", XTI)]),
        ("GENERACIÓN DE CAJA", [
            ("Flujo de caja libre",
             lambda k, R, y: f"={k}!{CH[y]}{R['op']}+{k}!{CH[y]}{R['iv']}", NUM),
            ("Conversión de EBITDA en caja",
             lambda k, R, y: f"=IFERROR(({k}!{CH[y]}{R['op']}+{k}!{CH[y]}{R['iv']})"
                             f"/{k}!{CH[y]}{R['ebitda']},\"\")", PCT),
            ("Capex sobre depreciación y amortización",
             lambda k, R, y: f"=IFERROR({k}!{CH[y]}{R['capex']}/{k}!{CH[y]}{R['da']},\"\")", XTI),
            ("Capex sobre ingresos",
             lambda k, R, y: f"={k}!{CH[y]}{R['capex']}/{k}!{CH[y]}{R['ing']}", PCT)]),
        ("EFICIENCIA", [
            ("Rotación del activo",
             lambda k, R, y: f"={k}!{CH[y]}{R['ing']}/{k}!{CH[y]}{R['at']}", XTI),
            ("Días de cobro",
             lambda k, R, y: f"={k}!{CH[y]}{R['ar']}*365/{k}!{CH[y]}{R['ings']}", NUM1),
            ("Días de pago",
             lambda k, R, y: f"=IFERROR({k}!{CH[y]}{R['cxp']}*365/({k}!{CH[y]}{R['ing']}"
                             f"+{k}!{CH[y]}{R['cs']}),\"\")", NUM1)]),
    ]
    for tit, filas in grupos:
        h.banda(tit)
        for etq, f, fmt in filas:
            h.sub(etq)
            h.anios()
            for k in ORDEN:
                R = refs[k]
                h.fila(None, PUERTOS[k]["nombre"],
                       vals={y: f(k, R, y) for y in ALL}, fmt=fmt, reg=False)
            h.r += 1
    return ws


def hoja_unitarios(wb, refs):
    """Economía unitaria, con la unidad de trabajo que corresponde a cada segmento."""
    from modelo import SEGMENTO
    ws = wb.create_sheet("Unitarios")
    h = Hoja(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 18
    for y in ALL:
        ws.column_dimensions[CH[y]].width = 12
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D1"
    h.titulo("Economía unitaria por segmento de operación",
             "La tonelada no es la unidad de trabajo de todos los terminales. Un terminal de contenedores "
             "vende movimientos de caja y uno de granel mineral vende tonelaje por faja. Cada terminal se "
             "mide con su unidad y solo se compara con sus pares de segmento.")
    segs = ["Contenedores", "Multipropósito con contenedores",
            "Graneles y carga fraccionada", "Granel mineral especializado"]
    for seg in segs:
        ks = [k for k in ORDEN if SEGMENTO[k] == seg]
        if not ks:
            continue
        h.banda(seg.upper() + ("  ·  ÚNICO TERMINAL DEL SISTEMA EN EL SEGMENTO" if len(ks) == 1 else ""))
        porteu = any(PUERTOS[k]["trafico"]["teu"][4] > 1000 for k in ks)
        bloques = []
        if porteu:
            bloques += [("Ingreso por TEU (USD)", "ings", "teu"), ("EBITDA por TEU (USD)", "ebitda", "teu")]
        bloques += [("Ingreso por tonelada (USD)", "ings", "tm"), ("Costo por tonelada (USD)", "cs", "tm"),
                    ("EBITDA por tonelada (USD)", "ebitda", "tm")]
        for etq, num, den in bloques:
            h.sub(etq)
            h.anios()
            for k in ORDEN:
                if k not in ks:
                    continue
                R = refs[k]
                u = UNID[k][0]
                aplica = (den == "teu" and PUERTOS[k]["trafico"]["teu"][4] > 1000) or \
                         (den == "tm" and not (k == "DPW"))
                if not aplica:
                    h.fila(None, PUERTOS[k]["nombre"],
                           vals={y: "n.a." for y in ALL}, fmt="@", reg=False)
                    continue
                if den == "teu" and u == "TEU":
                    dref = f"{k}!{{c}}{R['vol']}"
                elif den == "tm" and u == "TM":
                    dref = f"{k}!{{c}}{R['vol']}"
                else:
                    dref = None
                if dref is None:
                    h.fila(None, PUERTOS[k]["nombre"], vals={y: "n.d." for y in ALL}, fmt="@", reg=False)
                    continue
                sg = "-" if num == "cs" else ""
                h.fila(None, PUERTOS[k]["nombre"],
                       vals={y: f"=IFERROR({sg}{k}!{CH[y]}{R[num]}*1000/"
                                f"{dref.format(c=CH[y])},\"\")" for y in ALL},
                       fmt=NUM2, reg=False)
            h.r += 1
        h.sub("Comparables entre segmentos")
        for etq, num in [("Ingreso por recalada (USD)", "ings"), ("EBITDA por recalada (USD)", "ebitda")]:
            h.anios()
            for k in ks:
                R = refs[k]
                nav = {y: PUERTOS[k]["trafico"]["naves"][HIST.index(y)] for y in HIST}
                h.fila(None, f"{PUERTOS[k]['nombre']} — {etq}",
                       vals={y: f"={k}!{CH[y]}{R[num]}*1000/{nav[y]}" for y in HIST},
                       fmt=NUM, reg=False)
            h.r += 1
        h.nota("Las recaladas no se proyectan: el modelo de ingresos se construye sobre el volumen, "
               "no sobre el número de escalas.")
        h.r += 1
    return ws


DIAG = {
 "APMT": ("Escala y solvencia, con el volumen estancado",
  "Es el mayor operador del sistema por ingresos y el único que cerró 2025 sin deuda financiera. Su margen "
  "EBITDA mejoró y el ROA es el más alto del conjunto. El punto débil está en el volumen: la carga creció "
  "apenas 0,4 % anual en cinco años, de modo que el avance de la facturación proviene íntegramente de tarifa "
  "y mezcla. Con más de USD 866 millones de compromiso de inversión pendiente, el retorno de ese capex depende "
  "de un crecimiento de carga que hoy no se registra."),
 "DPW": ("Máxima eficiencia operativa, en fase de cosecha",
  "Lidera el sistema en margen EBITDA y en rendimiento por recalada, con el costo unitario más bajo de su "
  "segmento. Completó el 99 % de su compromiso de inversión y ha pasado a distribuir: durante 2025 repartió "
  "USD 180,6 millones en dividendos, lo que redujo su patrimonio e infló el ROE. El ROA es la medida correcta "
  "de su desempeño. Su deuda, la mayor del sistema en términos absolutos, está holgadamente cubierta."),
 "TISUR": ("Generación de caja sólida sobre un patrimonio en retroceso",
  "Margen EBITDA y ROA en el cuartil superior. Su ROE es el más alto del sistema pero carece de valor "
  "comparativo: el patrimonio cayó de USD 90 a USD 60 millones entre 2021 y 2025 con utilidad neta estancada. "
  "Tres señales de atención: el capital de trabajo más negativo del conjunto, la posición de efectivo más "
  "ajustada, y ninguna inversión ejecutada en 2025 con el 57 % del compromiso contractual pendiente y la "
  "concesión venciendo en 2029."),
 "TPP": ("La mejor conversión de tonelaje en margen",
  "Con el 5,2 % de la carga del sistema captura el 8,2 % de los ingresos y el 8,3 % del EBITDA. Su EBITDA por "
  "tonelada es el más alto del conjunto y su margen neto el mayor del sistema. El ejercicio 2025 fue "
  "excepcional: los ingresos por servicios crecieron 58,8 % y la utilidad operativa casi se duplicó. Mantiene "
  "la liquidez corriente más alta por el peso de los fondos restringidos en fideicomiso. Está cerca de cerrar "
  "su ciclo de inversión."),
 "TPGSM": ("Crecimiento de volumen sin conversión a valor",
  "Es el terminal de mayor crecimiento del sistema y el único con resultado neto negativo en 2025, el cuarto "
  "en cinco ejercicios. Genera caja: su flujo libre fue positivo pese a la pérdida contable, que se explica "
  "por la amortización del intangible y el gasto financiero. El problema no es liquidez sino retorno: su ROIC "
  "de 2,7 % queda muy por debajo de cualquier costo de capital razonable. Tiene el costo por tonelada más alto "
  "de su segmento y el EBITDA por tonelada más bajo del sistema."),
 "STI": ("Operación rentable bajo una estructura de deuda exigente",
  "Margen EBITDA y margen operativo en línea con el sistema, y el segundo mejor rendimiento por recalada del "
  "conjunto. La restricción está en el balance: una deuda que triplica el patrimonio, cobertura de intereses "
  "ajustada y una posición de caja casi nula al cierre de 2025. El gasto financiero absorbe más de una cuarta "
  "parte de su EBITDA."),
 "TCSA": ("Activo maduro con concesión en su tramo final",
  "Concesión de 20 años iniciada en 2014, con el 100 % de la inversión ejecutada y sin capex adicional "
  "comprometido. El ejercicio 2025 fue el mejor de la serie. El volumen está plano y el intangible de "
  "concesión se amortiza hacia cero. El punto crítico de gestión es el ciclo de cobro, 110 días, que "
  "inmoviliza casi un tercio de la facturación anual sobre una base de clientes muy concentrada."),
}


def _rk(k, campo, mayor=True, y=2025):
    """Posición del terminal entre los siete según un indicador."""
    def val(kk):
        if campo == "deuda_neta_ebitda":
            r0 = next((x for x in DAT[kk]["serie"] if x["anio"] == y), None)
            eb = r0["ebitda"] if r0 else None
            if not eb:
                return -1e18 if mayor else 1e18
            d = (PUERTOS[kk]["esf"]["deuda_fin_cp"][HIST.index(y)]
                 + PUERTOS[kk]["esf"]["deuda_fin_lp"][HIST.index(y)]
                 + PUERTOS[kk]["esf"]["deuda_relacionadas_lp"][HIST.index(y)]
                 - PUERTOS[kk]["esf"]["efectivo"][HIST.index(y)])
            return d / eb
        r = next((x for x in DAT[kk]["serie"] if x["anio"] == y), None)
        v = r.get(campo) if r else None
        if v is None:
            v = (DAT[kk]["unit"][str(y)].get(campo) if campo in DAT[kk]["unit"][str(y)] else None)
        if v is None:
            v = (DAT[kk]["modelo"].get(str(y)) or {}).get(campo)
        return v if v is not None else (-1e18 if mayor else 1e18)
    orden = sorted(ORDEN, key=lambda kk: -val(kk) if mayor else val(kk))
    return orden.index(k) + 1



# ============================================ FACTOR DE PRODUCTIVIDAD (RPI - X)
# Revisiones tarifarias publicadas por el Ositrán en los Informes de Desempeño 2025.
# Convención: en la fórmula RPI - X, un factor NEGATIVO eleva la tarifa por encima de la
# inflación, y uno POSITIVO traslada ganancias de eficiencia al usuario.
FACTOR_X = {
 "APMT": {
  "formula": "Tarifa máxima ajustada anualmente por RPI − X, con RPI igual a la variación del CPI "
             "de los Estados Unidos de los últimos doce meses disponibles.",
  "rev": [("Inicio de explotación – jun 2016", "Primer quinquenio", 0.0,
           "Cláusula 8.25 del Contrato de Concesión: factor cero durante los primeros cinco años"),
          ("1 jul 2016 – 30 jun 2021", "Primera revisión", -0.0963,
           "Resolución de Consejo Directivo N° 030-2016-CD-OSITRAN, junio de 2016"),
          ("1 jul 2021 – 30 jun 2026", "Segunda revisión", -0.0604,
           "Resolución de Consejo Directivo N° 0024-2021-CD-OSITRAN, junio de 2021")],
  "ajuste": [("jul 2025 – jun 2026", 0.0235, -0.0604, 0.0839,
              "RPI 2,35 % menos X de −6,04 %")]},
 "DPW": {
  "formula": "Tarifa máxima ajustada anualmente por RPI − X sobre dos canastas: servicios en "
             "función a la nave y servicios en función a la carga.",
  "rev": [("Primeros cinco años con dos amarraderos", "Primer periodo", 0.0,
           "Factor cero por contrato hasta la primera revisión"),
          ("Desde 2015", "Primera revisión", 0.0414,
           "Primera revisión de tarifas máximas; factor calculado en 4,14 %"),
          ("Hasta el 17 ago 2025", "Segunda revisión", -0.0266,
           "Vigente hasta el cierre del quinquenio anterior"),
          ("18 ago 2025 – 17 ago 2030", "Tercera revisión", None,
           "Procedimiento de revisión tarifaria de oficio aún en proceso al cierre de 2025")],
  "ajuste": []},
 "TISUR": {
  "formula": "Tarifa máxima ajustada por RPI corregido por la variación del tipo de cambio, menos X. "
             "Es el único terminal del sistema cuyo reajuste incorpora el tipo de cambio.",
  "rev": [("Vigente al cierre de 2025", "Último factor aprobado", 0.0183,
           "Factor positivo: traslada ganancias de productividad al usuario")],
  "ajuste": [("2025 frente a 2024", 0.0169, 0.0183, 0.0737,
              "RPI 1,69 % menos X de 1,83 % menos variación del tipo de cambio de −7,51 %")]},
 "TPP": {
  "formula": "Tarifa máxima ajustada anualmente por RPI − X, con RPI igual a la inflación de los "
             "Estados Unidos. Primera revisión al quinto año del Nuevo Muelle de Contenedores.",
  "rev": [("2014 – nov 2019", "Primer quinquenio", 0.0,
           "Factor cero durante los primeros cinco años de explotación del Nuevo Muelle"),
          ("nov 2019 – 2 oct 2024", "Primera revisión", -0.0470,
           "Aprobado por el Consejo Directivo del Ositrán en noviembre de 2019"),
          ("3 oct 2024 – 2 oct 2029", "Segunda revisión", -0.0270,
           "Resolución de Presidencia N° 0054-2024-PD-OSITRAN, 20 de agosto de 2024")],
  "ajuste": [("Ajuste 2024, vigente hasta el 2 oct 2025", None, -0.0470, 0.0523,
              "Variación máxima permitida de +5,23 %"),
             ("Ajuste 2025, vigente desde el 3 oct 2025", None, -0.0270, 0.0562,
              "Variación máxima permitida de +5,62 %")]},
 "TPGSM": {
  "formula": "Tarifa máxima ajustada anualmente por RPI − X. Desde 2025 el reajuste se practica el "
             "17 de enero con el RPI de enero a diciembre del año anterior.",
  "rev": [("21 ago 2014 – 16 ene 2025", "Primer periodo", 0.0,
           "Factor cero: el ajuste consideró únicamente la variación del RPI"),
          ("17 ene 2025 – 16 dic 2030", "Primera revisión", -0.0550,
           "Resolución de Presidencia N° 005-2025-PD-OSITRAN, 15 de enero de 2025")],
  "ajuste": [("2025", 0.0034, -0.0550, 0.0584,
              "RPI 0,34 % menos X de −5,50 %; tope por canasta de 5,84 %")]},
 "STI": {
  "formula": "Tarifa máxima ajustada anualmente por RPI − X. El factor permanece en cero hasta el "
             "quinto año contado desde la explotación de la totalidad de las Etapas 1 y 2.",
  "rev": [("Desde el inicio de explotación", "Primer periodo", 0.0,
           "Cláusula 8.24: factor cero. La APN recibió las obras de las Etapas 1 y 2 el 2 de junio de 2023, "
           "de modo que la primera revisión corresponde al quinto año desde esa fecha")],
  "ajuste": [("2025", 0.0301, 0.0, 0.0301, "RPI 3,01 % menos X de 0 %")]},
 "TCSA": {
  "formula": "Tarifa máxima ajustada anualmente por RPI − X sobre las canastas de servicios estándar "
             "en función a la nave y en función a la carga.",
  "rev": [("2014 – oct 2019", "Primer quinquenio", 0.0,
           "Cláusula 8.20 del Contrato de Concesión: factor cero durante los primeros cinco años"),
          ("oct 2019 – 2024", "Primera revisión", 0.0144,
           "Resolución de Consejo Directivo N° 0046-2019-CD-OSITRAN, octubre de 2019"),
          ("22 jul 2024 – 22 may 2029", "Segunda revisión", -0.0188,
           "Resolución de Presidencia N° 0049-2024-PD-OSITRAN, 18 de julio de 2024. Precedida de "
           "tarifas provisionales desde el 23 de mayo de 2024")],
  "ajuste": [("may 2025 – may 2026", None, -0.0188, 0.0419,
              "Variación máxima permitida de +4,19 %")]},
}


def bloque_dupont(h, k, ser):
    """Descomposición DuPont de tres y de cinco factores."""
    h.banda("ANÁLISIS DUPONT, 2021 – 2025")
    h.anios(fc=False)
    h.nota("Descompone el retorno sobre el patrimonio en los factores que lo generan. Permite ver si el "
           "ROE viene del margen, de la rotación del activo o simplemente del apalancamiento.")
    h.sub("Descomposición de tres factores")
    f_mn = {y: ser("margen_neto", y) for y in HIST}
    f_rt = {y: ser("rotacion_activos", y) for y in HIST}
    f_ap = {y: ser("apalancamiento", y) for y in HIST}
    h.fila(None, "Margen neto · utilidad neta sobre ingresos", vals=f_mn, fmt=PCT, estilo="s", reg=False)
    h.fila(None, "Rotación del activo · ingresos sobre activo", vals=f_rt, fmt=XTI, estilo="s", reg=False)
    h.fila(None, "Apalancamiento · activo sobre patrimonio", vals=f_ap, fmt=XTI, estilo="s", reg=False)
    h.fila(None, "ROE resultante", fmt=PCT, estilo="t", reg=False,
           vals={y: (f_mn[y] * f_rt[y] * f_ap[y]) if None not in (f_mn[y], f_rt[y], f_ap[y]) else None
                 for y in HIST})
    h.fila(None, "ROE publicado, para contraste", vals={y: ser("roe", y) for y in HIST},
           fmt=PCT, estilo="s", reg=False)
    h.r += 1
    h.sub("Descomposición de cinco factores")
    g_cf = {y: (ser("utilidad_neta", y) / ser("uai", y)) if ser("uai", y) else None for y in HIST}
    g_ci = {y: (ser("uai", y) / ser("utilidad_operativa", y)) if ser("utilidad_operativa", y) else None
            for y in HIST}
    g_mo = {y: ser("margen_operativo", y) for y in HIST}
    h.fila(None, "Carga fiscal · utilidad neta sobre utilidad antes de impuestos",
           vals=g_cf, fmt=XTI, estilo="s", reg=False)
    h.fila(None, "Carga de intereses · utilidad antes de impuestos sobre EBIT",
           vals=g_ci, fmt=XTI, estilo="s", reg=False)
    h.fila(None, "Margen operativo · EBIT sobre ingresos", vals=g_mo, fmt=PCT, estilo="s", reg=False)
    h.fila(None, "Rotación del activo · ingresos sobre activo", vals=f_rt, fmt=XTI, estilo="s", reg=False)
    h.fila(None, "Apalancamiento · activo sobre patrimonio", vals=f_ap, fmt=XTI, estilo="s", reg=False)
    h.fila(None, "ROE resultante", fmt=PCT, estilo="t", reg=False,
           vals={y: (g_cf[y] * g_ci[y] * g_mo[y] * f_rt[y] * f_ap[y])
                 if None not in (g_cf[y], g_ci[y], g_mo[y], f_rt[y], f_ap[y]) else None for y in HIST})
    h.r += 1
    h.sub("Lectura")
    h.fila(None, "Retorno operativo · margen neto por rotación (ROA)", fmt=PCT, estilo="s", reg=False,
           vals={y: (f_mn[y] * f_rt[y]) if None not in (f_mn[y], f_rt[y]) else None for y in HIST})
    h.fila(None, "Aporte del apalancamiento al ROE", fmt=PCT, estilo="s", reg=False,
           vals={y: (f_mn[y] * f_rt[y] * (f_ap[y] - 1)) if None not in (f_mn[y], f_rt[y], f_ap[y])
                 else None for y in HIST})
    h.nota("El aporte del apalancamiento es la parte del ROE que no procede de la operación sino de la "
           "estructura de financiamiento. Cuando el patrimonio contable se erosiona por reparto de "
           "dividendos, ese componente se dispara sin que medie mejora operativa alguna.")
    h.r += 1


def bloque_factorx(h, k, ws):
    """Factor de productividad X y revisiones tarifarias."""
    F = FACTOR_X[k]
    h.banda("FACTOR DE PRODUCTIVIDAD (X) Y REVISIONES TARIFARIAS")
    h._set("B", F["formula"], F_SB)
    ws.merge_cells(start_row=h.r, start_column=2, end_row=h.r + 1, end_column=11)
    ws[f"B{h.r}"].alignment = Alignment(wrap_text=True, vertical="top")
    h.r += 3
    h._set("B", "En la fórmula RPI − X, un factor NEGATIVO eleva la tarifa por encima de la inflación; "
                "uno POSITIVO traslada ganancias de eficiencia al usuario.", F_NT)
    h.r += 2
    h.sub("Factores aprobados por el regulador")
    for col, etq in [("B", "Periodo de vigencia"), ("D", "Revisión"), ("F", "Factor X"),
                     ("H", "Resolución y observaciones")]:
        c = h._set(col, etq, F_TT)
        c.border = Border(bottom=THIN)
    h.r += 1
    for per, rev, x, obs in F["rev"]:
        h._set("B", per, F_LB)
        h._set("D", rev, F_SB)
        c = h._set("F", x if x is not None else "en revisión",
                   Font(name="Calibri", size=10, bold=True,
                        color=("B3332F" if (x or 0) < 0 else ("0F7A2E" if (x or 0) > 0 else INK2))),
                   fmt=(PCT2 if x is not None else "@"), align=Alignment(horizontal="right"))
        h._set("H", obs, F_NT)
        h.r += 1
    h.r += 1
    if F["ajuste"]:
        h.sub("Ajuste tarifario resultante")
        for col, etq in [("B", "Periodo"), ("D", "RPI"), ("E", "Factor X"),
                         ("F", "Var. máxima"), ("H", "Detalle")]:
            c = h._set(col, etq, F_TT)
            c.border = Border(bottom=THIN)
        h.r += 1
        for per, rpi, x, var, det in F["ajuste"]:
            h._set("B", per, F_LB)
            h._set("D", rpi if rpi is not None else "n.d.", F_LB,
                   fmt=(PCT2 if rpi is not None else "@"), align=Alignment(horizontal="right"))
            h._set("E", x, F_LB, fmt=PCT2, align=Alignment(horizontal="right"))
            h._set("F", var, F_TT, fmt=PCT2, align=Alignment(horizontal="right"))
            h._set("H", det, F_NT)
            h.r += 1
        h.r += 1
    h.sub("Lectura para el modelo")
    h.nota("El input «Ajuste de la tarifa media» de la hoja Inputs debe ser coherente con este cuadro: "
           "es el reajuste anual que el marco regulatorio permite, no una decisión comercial del "
           "concesionario. El escenario base del modelo usa 2,0 % anual, por debajo del tope vigente, "
           "porque la tarifa media efectiva depende además de la mezcla de servicios y de los acuerdos "
           "comerciales con las líneas navieras.")
    h.r += 1

def hoja_ficha(wb, k):
    """Ficha de resumen por terminal.

    Convención del libro: el dato de origen —estados auditados, tráfico, y la depreciación
    derivada del margen EBITDA que publica el Ositrán— va en azul y fijo. Todo lo que se
    puede calcular a partir de él va en negro y como fórmula viva.
    """
    from modelo import SEGMENTO
    P, s = PUERTOS[k], SUP[k]
    SEG = SEGMENTO[k]
    ws = wb.create_sheet(f"Ficha {k}")
    h = Hoja(ws)
    ws.column_dimensions["A"].width = 2
    ws.column_dimensions["B"].width = 44
    ws.column_dimensions["C"].width = 22
    for y in ALL:
        ws.column_dimensions[CH[y]].width = 13
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "D1"
    R = {}

    def ser(c, y):
        r = next((x for x in DAT[k]["serie"] if x["anio"] == y), None)
        return r.get(c) if r else None

    def uni(c, y):
        return DAT[k]["unit"][str(y)].get(c)

    def mo(c, y):
        return (DAT[k]["modelo"].get(str(y)) or {}).get(c)

    def orig(campo, y):
        """Resuelve una línea de origen, incluidas las derivadas por diferencia de la fuente."""
        if campo == "_ingo":
            return (S(k, "ingresos_totales", y) - S(k, "ing_servicios", y)
                    - S(k, "ing_construccion", y))
        if campo == "_cc":
            return S(k, "costos_totales", y) - S(k, "costo_servicios", y)
        if campo == "_oo":
            return (S(k, "utilidad_operativa", y) - (S(k, "utilidad_bruta", y)
                    - S(k, "gastos_admin", y) - S(k, "gastos_venta", y)))
        if campo == "_dc":
            return (S(k, "uai", y) - (S(k, "utilidad_operativa", y)
                    + S(k, "ingresos_financieros", y) - S(k, "gastos_financieros", y)))
        if campo == "_oac":
            return (S(k, "activo_corriente", y) - S(k, "efectivo", y) - S(k, "fondos_restringidos", y)
                    - S(k, "ctas_cobrar_com", y) - S(k, "inventarios", y))
        if campo == "_oanc":
            return S(k, "activo_no_corriente", y) - S(k, "intangibles", y) - S(k, "ppe", y)
        if campo == "_opc":
            return (S(k, "pasivo_corriente", y) - S(k, "cxp_comerciales", y)
                    - S(k, "cxp_relacionadas", y) - S(k, "deuda_fin_cp", y))
        if campo == "_opnc":
            return (S(k, "pasivo_no_corriente", y) - S(k, "deuda_fin_lp", y)
                    - S(k, "deuda_relacionadas_lp", y))
        if campo == "_cap":
            return S(k, "patrimonio", y) - S(k, "resultados_acumulados", y)
        if campo == "dep_amort":
            return ser("dep_amort", y)
        return S(k, campo, y)

    def dato(nombre, etq, campo, est="s", fmt=NUM, nota=None):
        f = {"n": F_LB, "s": F_SB, "t": F_TT}[est]
        h._set("B", ("    " + etq) if est == "s" else etq, f)
        if nota:
            h._set("C", nota, F_NT)
        for y in HIST:
            c = h._set(CH[y], orig(campo, y), F_IN, fmt=fmt)
            c.fill = FILL_IN
        R[nombre] = h.r
        h.r += 1
        return h.r - 1

    def calc(nombre, etq, fml, est="s", fmt=NUM, nota=None, anios=None):
        f = {"n": F_LB, "s": F_SB, "t": F_TT}[est]
        c0 = h._set("B", ("    " + etq) if est == "s" else etq, f)
        if est == "t":
            c0.fill = FILL_TOT
        if nota:
            h._set("C", nota, F_NT)
        for y in (anios or HIST):
            v = fml(y)
            if v is None:
                continue
            c = h._set(CH[y], v, F_FX, fmt=fmt)
            if est == "t":
                c.fill = FILL_TOT
                c.font = Font(name="Calibri", size=10, color=INK, bold=True)
        if nombre:
            R[nombre] = h.r
        h.r += 1
        return h.r - 1

    def texto(etq, est="s"):
        h._set("B", ("    " + etq) if est == "s" else etq, F_SB)
        for y in HIST:
            h._set(CH[y], "n.a.", F_NT, align=Alignment(horizontal="right"))
        h.r += 1

    def rf(n, y):
        return f"{CH[y]}{R[n]}"

    def seb(y, expr):
        return f'=IF({rf("da", y)}="","",{expr})'

    # ===================================================== identificación
    h.titulo(P["nombre"], f"{P['terminal']} · {P['operador']}")
    h._set("B", "Ficha de resumen. Dato de origen en azul; todo lo calculado es fórmula viva. "
                f"El modelo de tres estados proyectado está en la hoja «{k}».", F_NT)
    h.r += 2
    h.banda("IDENTIFICACIÓN DE LA CONCESIÓN")
    venc = P["inicio"] + P["plazo"]
    for etq, val in [("Concesionario", P["operador"]), ("Tipo de carga", P["tipo"]),
                     ("Segmento de operación", SEG),
                     ("Participación contenedorizada del tonelaje",
                      f"{uni('pct_contenedorizada', 2025) * 100:,.1f} %"),
                     ("Unidad de trabajo pertinente", s["unidad"]),
                     ("Modalidad", P["modalidad"]),
                     ("Inicio de la concesión", P["inicio"]), ("Plazo", f"{P['plazo']} años"),
                     ("Vencimiento", f"{venc}" + ("  ⚠ dentro del horizonte del modelo" if venc <= 2030 else "")),
                     ("Compromiso de inversión", f"USD {P['compromiso_inv']:,.1f} millones"),
                     ("Inversión ejecutada acumulada", f"USD {P['inv_acum']:,.1f} millones"),
                     ("Avance del compromiso", f"{P['avance_inv']:,.1f} %"),
                     ("Inversión pendiente", f"USD {P['compromiso_inv'] - P['inv_acum']:,.1f} millones"),
                     ("Inversión ejecutada en 2025", f"USD {P['inv_2025']:,.2f} millones")]:
        h._set("B", etq, F_SB)
        h._set("C", val, F_LB, align=Alignment(horizontal="left"))
        h.r += 1
    h.r += 1

    # ===================================================== posición en el sistema
    h.banda("POSICIÓN EN EL SISTEMA, 2025")
    for col, etq in [("B", "Indicador"), ("C", "Valor 2025"), ("D", "Cuota"),
                     ("E", "Puesto")]:
        c = h._set(col, etq, F_TT)
        c.border = Border(bottom=THIN)
    h.r += 1
    POS = [("Ingresos totales", "ing", NUM, True, True),
           ("EBITDA", "ebitda", NUM, True, True),
           ("Utilidad neta", "un", NUM, True, True),
           ("Carga movilizada (TM)", "carga", NUM, True, True),
           ("Contenedores (TEU)", "teu", NUM, True, True),
           ("Naves atendidas", "naves", NUM, True, True),
           ("Activo total", "at", NUM, True, True),
           ("Margen EBITDA", "m_ebitda", PCT, True, False),
           ("Rentabilidad del activo (ROA)", "roa", PCT, True, False),
           ("EBITDA sobre activo total", "eb_act", PCT, True, False),
           ("Deuda neta / EBITDA", "dn_eb", XTI, False, False)]
    pos0 = h.r
    for etq, _, _, _, _ in POS:
        h._set("B", etq, F_LB)
        h.r += 1
    h.r += 1

    # ===================================================== estado de resultados
    h.banda("ESTADO DE RESULTADOS AUDITADO, 2021 – 2025")
    h.anios(fc=False)
    dato("ings", "Ingresos por servicios portuarios", "ing_servicios")
    dato("ingc", "Ingresos por servicios de construcción", "ing_construccion")
    dato("ingo", "Otros ingresos de actividades ordinarias", "_ingo")
    calc("ing", "Ingresos totales",
         lambda y: f'=SUM({rf("ings", y)},{rf("ingc", y)},{rf("ingo", y)})', est="t")
    dato("cs", "Costo de servicios portuarios", "costo_servicios")
    dato("cc", "Costo de servicios de construcción", "_cc")
    calc("ct", "Costos totales", lambda y: f'={rf("cs", y)}+{rf("cc", y)}', est="n")
    calc("ub", "Utilidad bruta", lambda y: f'={rf("ing", y)}-{rf("ct", y)}', est="t")
    dato("ga", "Gastos de administración", "gastos_admin", est="n")
    dato("gv", "Gastos de venta", "gastos_venta", est="n")
    dato("oo", "Otros ingresos (gastos) operativos", "_oo", est="n")
    calc("ebit", "Utilidad operativa (EBIT)",
         lambda y: f'={rf("ub", y)}-{rf("ga", y)}-{rf("gv", y)}+{rf("oo", y)}', est="t")
    dato("da", "Depreciación y amortización", "dep_amort",
         nota="Derivada del margen EBITDA del Ositrán")
    calc("ebitda", "EBITDA", lambda y: seb(y, f'{rf("ebit", y)}+{rf("da", y)}'), est="t")
    dato("if", "Ingresos financieros", "ingresos_financieros", est="n")
    dato("gf", "Gastos financieros", "gastos_financieros", est="n")
    dato("dc", "Diferencia en cambio y otros", "_dc", est="n")
    calc("uai", "Utilidad antes de impuestos",
         lambda y: f'={rf("ebit", y)}+{rf("if", y)}-{rf("gf", y)}+{rf("dc", y)}', est="t")
    dato("imp", "Impuesto a la renta", "impuesto_renta", est="n")
    calc("un", "Utilidad neta", lambda y: f'={rf("uai", y)}-{rf("imp", y)}', est="t")
    h.r += 1

    # ===================================================== situación financiera
    h.banda("ESTADO DE SITUACIÓN FINANCIERA AUDITADO, 2021 – 2025")
    h.anios(fc=False)
    dato("caja", "Efectivo y equivalentes", "efectivo")
    dato("fr", "Fondos restringidos en fideicomiso", "fondos_restringidos")
    dato("ar", "Cuentas por cobrar comerciales", "ctas_cobrar_com")
    dato("inv", "Inventarios y suministros", "inventarios")
    dato("oac", "Otros activos corrientes", "_oac")
    calc("ac", "Total activo corriente",
         lambda y: f'=SUM({rf("caja", y)},{rf("fr", y)},{rf("ar", y)},{rf("inv", y)},{rf("oac", y)})',
         est="t")
    dato("int", "Activos intangibles de la concesión", "intangibles")
    dato("ppe", "Instalaciones, maquinaria y equipo", "ppe")
    dato("oanc", "Otros activos no corrientes", "_oanc")
    calc("anc", "Total activo no corriente",
         lambda y: f'=SUM({rf("int", y)},{rf("ppe", y)},{rf("oanc", y)})', est="t")
    calc("at", "TOTAL ACTIVO", lambda y: f'={rf("ac", y)}+{rf("anc", y)}', est="t")
    dato("cxp", "Cuentas por pagar comerciales", "cxp_comerciales")
    dato("cxr", "Cuentas por pagar a relacionadas, corriente", "cxp_relacionadas")
    dato("dcp", "Deuda financiera, porción corriente", "deuda_fin_cp")
    dato("opc", "Otros pasivos corrientes", "_opc")
    calc("pc", "Total pasivo corriente",
         lambda y: f'=SUM({rf("cxp", y)},{rf("cxr", y)},{rf("dcp", y)},{rf("opc", y)})', est="t")
    dato("dlp", "Deuda financiera, porción no corriente", "deuda_fin_lp")
    dato("drl", "Deuda con relacionadas, no corriente", "deuda_relacionadas_lp")
    dato("opnc", "Otros pasivos no corrientes", "_opnc")
    calc("pnc", "Total pasivo no corriente",
         lambda y: f'=SUM({rf("dlp", y)},{rf("drl", y)},{rf("opnc", y)})', est="t")
    calc("pt", "TOTAL PASIVO", lambda y: f'={rf("pc", y)}+{rf("pnc", y)}', est="t")
    calc("deuda", "Deuda financiera e intragrupo",
         lambda y: f'=SUM({rf("dcp", y)},{rf("dlp", y)},{rf("drl", y)})')
    dato("cap", "Capital y reservas", "_cap")
    dato("ra", "Resultados acumulados", "resultados_acumulados")
    calc("pat", "TOTAL PATRIMONIO", lambda y: f'={rf("cap", y)}+{rf("ra", y)}', est="t")
    calc("ck", "Control: activo menos pasivo y patrimonio",
         lambda y: f'=ROUND({rf("at", y)}-{rf("pt", y)}-{rf("pat", y)},0)', est="t",
         nota="Debe ser cero en los cinco ejercicios")
    for y in HIST:
        ws[f"{CH[y]}{R['ck']}"].font = Font(name="Calibri", size=10, bold=True, color="0F7A2E")
    calc("ctrab", "Capital de trabajo", lambda y: f'={rf("ac", y)}-{rf("pc", y)}')
    h.r += 1

    # ===================================================== tráfico
    h.banda("TRÁFICO, 2021 – 2025")
    h.anios(fc=False)
    for nm, etq, campo in [("naves", "Naves atendidas", "naves"),
                           ("carga", "Carga movilizada (toneladas)", "carga_tm"),
                           ("cont", "Contenedores (unidades)", "contenedores"),
                           ("teu", "Contenedores (TEU)", "teu")]:
        h._set("B", "    " + etq, F_SB)
        for y in HIST:
            c = h._set(CH[y], PUERTOS[k]["trafico"][campo][HIST.index(y)], F_IN, fmt=NUM)
            c.fill = FILL_IN
        R[nm] = h.r
        h.r += 1
    h.r += 1

    # ===================================================== indicadores
    h.banda("INDICADORES FINANCIEROS Y OPERATIVOS, 2021 – 2025")
    h.anios(fc=False)
    h.sub("Rentabilidad")
    calc("mb", "Margen bruto", lambda y: f'={rf("ub", y)}/{rf("ing", y)}', fmt=PCT)
    calc("mo", "Margen operativo", lambda y: f'={rf("ebit", y)}/{rf("ing", y)}', fmt=PCT)
    calc("m_ebitda", "Margen EBITDA", lambda y: seb(y, f'{rf("ebitda", y)}/{rf("ing", y)}'), fmt=PCT)
    calc("mn", "Margen neto", lambda y: f'={rf("un", y)}/{rf("ing", y)}', fmt=PCT)
    calc("roa", "ROA — utilidad neta / activo", lambda y: f'={rf("un", y)}/{rf("at", y)}', fmt=PCT)
    calc("roe", "ROE — utilidad neta / patrimonio",
         lambda y: f'=IFERROR({rf("un", y)}/{rf("pat", y)},"")', fmt=PCT)
    calc("roic", "ROIC aproximado",
         lambda y: f'=IFERROR({rf("ebit", y)}*(1-{rf("imp", y)}/{rf("uai", y)})'
                   f'/({rf("deuda", y)}+{rf("pat", y)}),"")', fmt=PCT,
         nota="EBIT después de la tasa efectiva, sobre deuda más patrimonio")
    calc("eb_act", "EBITDA sobre activo total",
         lambda y: seb(y, f'{rf("ebitda", y)}/{rf("at", y)}'), fmt=PCT)

    h.sub("Liquidez y solvencia")
    calc("liq", "Liquidez corriente", lambda y: f'={rf("ac", y)}/{rf("pc", y)}', fmt=XTI)
    calc("acid", "Prueba ácida", lambda y: f'=({rf("ac", y)}-{rf("inv", y)})/{rf("pc", y)}', fmt=XTI)
    calc("apal", "Apalancamiento — activo / patrimonio",
         lambda y: f'=IFERROR({rf("at", y)}/{rf("pat", y)},"")', fmt=XTI)
    calc("pas_pat", "Pasivo / patrimonio",
         lambda y: f'=IFERROR({rf("pt", y)}/{rf("pat", y)},"")', fmt=XTI)
    calc("end", "Endeudamiento del activo", lambda y: f'={rf("pt", y)}/{rf("at", y)}', fmt=PCT)
    calc("db_eb", "Deuda bruta / EBITDA",
         lambda y: seb(y, f'{rf("deuda", y)}/{rf("ebitda", y)}'), fmt=XTI)
    calc("dn_eb", "Deuda neta / EBITDA",
         lambda y: seb(y, f'({rf("deuda", y)}-{rf("caja", y)})/{rf("ebitda", y)}'), fmt=XTI)
    calc("cob", "Cobertura de intereses",
         lambda y: seb(y, f'IFERROR({rf("ebitda", y)}/{rf("gf", y)},"")'), fmt=XTI)
    calc(None, "Capital de trabajo", lambda y: f'={rf("ctrab", y)}')
    h.nota("La deuda incluye las cuentas por pagar a relacionadas de largo plazo, que en Matarani y "
           "Salaverry devengan interés y son financiamiento, no crédito comercial. Mismo criterio que la hoja del modelo.")

    h.sub("Eficiencia operativa — volumen")
    for etq, nm in [("Carga movilizada (TM)", "carga"), ("TEU movilizados", "teu"),
                    ("Contenedores (unidades)", "cont"), ("Naves atendidas", "naves")]:
        calc(None, etq, lambda y, n=nm: f'={rf(n, y)}')
    calc("tm_nave", "Toneladas por recalada", lambda y: f'={rf("carga", y)}/{rf("naves", y)}')

    h.sub(f"Eficiencia operativa — unitarios ({SEG})")
    h.nota("Los unitarios solo son comparables con los de los terminales del mismo segmento. "
           "«n.a.» indica que la unidad de trabajo no aplica a esta operación.")
    hay_tm = uni("ingreso_por_tm", 2025) is not None
    hay_teu = uni("ingreso_por_teu", 2025) is not None
    if hay_tm:
        calc("i_tm", "Ingreso por tonelada (USD)",
             lambda y: f'={rf("ings", y)}*1000/{rf("carga", y)}', fmt=NUM2)
        calc("c_tm", "Costo por tonelada (USD)",
             lambda y: f'={rf("cs", y)}*1000/{rf("carga", y)}', fmt=NUM2)
        calc("e_tm", "EBITDA por tonelada (USD)",
             lambda y: seb(y, f'{rf("ebitda", y)}*1000/{rf("carga", y)}'), fmt=NUM2)
    else:
        for etq in ["Ingreso por tonelada (USD)", "Costo por tonelada (USD)", "EBITDA por tonelada (USD)"]:
            texto(etq)
    if hay_teu:
        calc("i_teu", "Ingreso por TEU (USD)",
             lambda y: f'=IFERROR({rf("ings", y)}*1000/{rf("teu", y)},"")', fmt=NUM2)
        calc("c_teu", "Costo por TEU (USD)",
             lambda y: f'=IFERROR({rf("cs", y)}*1000/{rf("teu", y)},"")', fmt=NUM2)
        calc("e_teu", "EBITDA por TEU (USD)",
             lambda y: seb(y, f'IFERROR({rf("ebitda", y)}*1000/{rf("teu", y)},"")'), fmt=NUM2)
    else:
        for etq in ["Ingreso por TEU (USD)", "Costo por TEU (USD)", "EBITDA por TEU (USD)"]:
            texto(etq)
    calc("i_rec", "Ingreso por recalada (USD)", lambda y: f'={rf("ings", y)}*1000/{rf("naves", y)}')
    calc("e_rec", "EBITDA por recalada (USD)",
         lambda y: seb(y, f'{rf("ebitda", y)}*1000/{rf("naves", y)}'))

    h.sub("Eficiencia operativa — capital y ciclo")
    calc("rot", "Rotación del activo", lambda y: f'={rf("ing", y)}/{rf("at", y)}', fmt=XTI)
    calc("d_cob", "Días de cobro", lambda y: f'={rf("ar", y)}*365/{rf("ings", y)}', fmt=NUM1)
    calc("d_pag", "Días de pago", lambda y: f'={rf("cxp", y)}*365/{rf("ct", y)}', fmt=NUM1)
    calc("d_inv", "Días de inventario", lambda y: f'={rf("inv", y)}*365/{rf("ct", y)}', fmt=NUM1)
    h.r += 1

    # ===================================================== modelo de caja
    h.banda("MODELO DE CAJA: DEL RESULTADO OPERATIVO AL FLUJO LIBRE")
    h.anios(fc=False)
    h.nota("Método indirecto: el Ositrán no publica estado de flujos. Cubre los ejercicios en que la "
           "depreciación y amortización es derivable sin estimación y existe balance del año anterior.")
    AC = [y for y in HIST if mo("fcff", y) is not None]

    def pvh(y):
        return CH[y - 1]

    calc(None, "Resultado operativo (EBIT)", lambda y: f'={rf("ebit", y)}', est="t", anios=AC)
    calc("tasa", "Tasa impositiva efectiva aplicada",
         lambda y: f'=IF(OR({rf("uai", y)}<=0,{rf("imp", y)}/{rf("uai", y)}<0,'
                   f'{rf("imp", y)}/{rf("uai", y)}>0.5),0.295,{rf("imp", y)}/{rf("uai", y)})',
         fmt=PCT, anios=AC)
    calc("nopat", "Resultado operativo después de impuestos (NOPAT)",
         lambda y: f'={rf("ebit", y)}*(1-{rf("tasa", y)})', est="t", anios=AC)
    calc(None, "Más depreciación y amortización", lambda y: f'={rf("da", y)}', anios=AC)
    calc("wc_op", "Capital de trabajo operativo",
         lambda y: f'=({rf("ac", y)}-{rf("caja", y)}-{rf("fr", y)})'
                   f'-({rf("pc", y)}-{rf("dcp", y)}-{rf("cxr", y)})')
    calc("dwc", "Variación del capital de trabajo operativo",
         lambda y: f'={rf("wc_op", y)}-{pvh(y)}{R["wc_op"]}', anios=AC)
    calc("capex", "Inversión en activo fijo (capex)",
         lambda y: f'=({rf("int", y)}+{rf("ppe", y)})-({pvh(y)}{R["int"]}+{pvh(y)}{R["ppe"]})'
                   f'+{rf("da", y)}', anios=AC)
    calc("fcff", "Flujo de caja libre para la empresa (FCFF)",
         lambda y: f'={rf("nopat", y)}+{rf("da", y)}-{rf("dwc", y)}-{rf("capex", y)}',
         est="t", anios=AC)
    calc("int_dt", "Intereses netos después de impuestos",
         lambda y: f'=({rf("gf", y)}-{rf("if", y)})*(1-{rf("tasa", y)})', anios=AC)
    calc("ddeuda", "Variación de la deuda financiera",
         lambda y: f'={rf("deuda", y)}-{pvh(y)}{R["deuda"]}', anios=AC)
    calc("fcfe", "Flujo de caja libre para el accionista (FCFE)",
         lambda y: f'={rf("fcff", y)}-{rf("int_dt", y)}+{rf("ddeuda", y)}', est="t", anios=AC)
    calc(None, "Margen de flujo libre sobre ingresos",
         lambda y: f'={rf("fcff", y)}/{rf("ing", y)}', fmt=PCT, anios=AC)
    calc(None, "Conversión de EBITDA en caja",
         lambda y: f'={rf("fcff", y)}/{rf("ebitda", y)}', fmt=PCT, anios=AC)
    calc(None, "Capex sobre depreciación y amortización",
         lambda y: f'={rf("capex", y)}/{rf("da", y)}', fmt=XTI, anios=AC)
    calc(None, "Capex sobre ingresos", lambda y: f'={rf("capex", y)}/{rf("ing", y)}',
         fmt=PCT, anios=AC)
    calc("cap_inv", "Capital invertido medio",
         lambda y: f'=(({rf("deuda", y)}+{rf("pat", y)})'
                   f'+({pvh(y)}{R["deuda"]}+{pvh(y)}{R["pat"]}))/2', anios=AC)
    calc("roic_m", "Retorno sobre el capital invertido (ROIC)",
         lambda y: f'={rf("nopat", y)}/{rf("cap_inv", y)}', est="t", fmt=PCT, anios=AC)
    h.sub("Valor económico añadido según el costo de capital supuesto")
    for w in (8, 10, 12):
        calc(None, f"EVA con costo de capital de {w} %",
             lambda y, w=w: f'=({rf("roic_m", y)}-{w / 100})*{rf("cap_inv", y)}', anios=AC)
    h.sub("Sensibilidad del capex, 2025")
    calc(None, "Capex derivado del balance", lambda y: f'={rf("capex", y)}', anios=[2025])
    h._set("B", "    Inversión reportada al regulador", F_SB)
    c = h._set(CH[2025], P["inv_2025"] * 1000, F_IN, fmt=NUM)
    c.fill = FILL_IN
    R["cap_reg"] = h.r
    h.r += 1
    calc(None, "FCFF con la inversión reportada",
         lambda y: f'={rf("nopat", y)}+{rf("da", y)}-{rf("dwc", y)}-{rf("cap_reg", y)}', anios=[2025])
    h.r += 1

    # ===================================================== DuPont
    h.banda("ANÁLISIS DUPONT, 2021 – 2025")
    h.anios(fc=False)
    h.nota("Descompone el retorno sobre el patrimonio en los factores que lo generan. Permite ver si el "
           "ROE viene del margen, de la rotación del activo o simplemente del apalancamiento.")
    h.sub("Descomposición de tres factores")
    calc(None, "Margen neto — utilidad neta / ingresos", lambda y: f'={rf("mn", y)}', fmt=PCT)
    calc(None, "Rotación del activo — ingresos / activo", lambda y: f'={rf("rot", y)}', fmt=XTI)
    calc(None, "Apalancamiento — activo / patrimonio", lambda y: f'={rf("apal", y)}', fmt=XTI)
    calc(None, "ROE resultante",
         lambda y: f'=IFERROR({rf("mn", y)}*{rf("rot", y)}*{rf("apal", y)},"")', est="t", fmt=PCT)
    calc(None, "ROE publicado, para contraste", lambda y: f'={rf("roe", y)}', fmt=PCT)
    h.r += 1
    h.sub("Descomposición de cinco factores")
    calc("cf", "Carga fiscal — utilidad neta / utilidad antes de impuestos",
         lambda y: f'=IFERROR({rf("un", y)}/{rf("uai", y)},"")', fmt=XTI)
    calc("ci", "Carga de intereses — utilidad antes de impuestos / EBIT",
         lambda y: f'=IFERROR({rf("uai", y)}/{rf("ebit", y)},"")', fmt=XTI)
    calc(None, "Margen operativo — EBIT / ingresos", lambda y: f'={rf("mo", y)}', fmt=PCT)
    calc(None, "Rotación del activo — ingresos / activo", lambda y: f'={rf("rot", y)}', fmt=XTI)
    calc(None, "Apalancamiento — activo / patrimonio", lambda y: f'={rf("apal", y)}', fmt=XTI)
    calc(None, "ROE resultante",
         lambda y: f'=IFERROR({rf("cf", y)}*{rf("ci", y)}*{rf("mo", y)}*{rf("rot", y)}'
                   f'*{rf("apal", y)},"")', est="t", fmt=PCT)
    h.r += 1
    h.sub("Lectura")
    calc(None, "Retorno operativo — margen neto por rotación (ROA)",
         lambda y: f'={rf("mn", y)}*{rf("rot", y)}', fmt=PCT)
    calc(None, "Aporte del apalancamiento al ROE",
         lambda y: f'=IFERROR({rf("mn", y)}*{rf("rot", y)}*({rf("apal", y)}-1),"")', fmt=PCT)
    h.nota("El aporte del apalancamiento es la parte del ROE que no procede de la operación sino de la "
           "estructura de financiamiento. Cuando el patrimonio contable se erosiona por reparto de "
           "dividendos, ese componente se dispara sin que medie mejora operativa alguna.")
    h.r += 1

    # ===================================================== factor X
    bloque_factorx(h, k, ws)

    # ===================================================== diagnóstico
    tit, txt = DIAG[k]
    h.banda("DIAGNÓSTICO")
    h._set("B", tit, Font(name="Calibri", size=11, bold=True, color=MG))
    h.r += 1
    ws.merge_cells(start_row=h.r, start_column=2, end_row=h.r + 5, end_column=10)
    c = h._set("B", txt, Font(name="Calibri", size=10, color=INK2))
    c.alignment = Alignment(wrap_text=True, vertical="top")
    h.r += 7
    h._set("B", "Fuente: Informe de Desempeño 2025 del Ositrán · estados financieros auditados de "
                f"{P['operador']} · Declaración Estadística del Ositrán. Cifras en miles de USD.", F_NT)

    # ===================================================== posición: enlaces y cuotas
    for i, (etq, nm, fmt, mayor, cuota) in enumerate(POS):
        r = pos0 + i
        ws[f"C{r}"] = f'={CH[2025]}{R[nm]}'
        ws[f"C{r}"].font = F_FX
        ws[f"C{r}"].number_format = fmt
        ws[f"C{r}"].alignment = Alignment(horizontal="right")
        if cuota:
            suma = "+".join(f"'Ficha {kk}'!{CH[2025]}{R[nm]}" for kk in ORDEN)
            ws[f"D{r}"] = f'=IFERROR({CH[2025]}{R[nm]}/({suma}),"")'
            ws[f"D{r}"].font = F_FX
            ws[f"D{r}"].number_format = PCT
            ws[f"D{r}"].alignment = Alignment(horizontal="right")
        lista = ",".join(f"'Ficha {kk}'!{CH[2025]}{R[nm]}" for kk in ORDEN)
        orden = "" if mayor else ",1"
        ws[f"E{r}"] = (f'=IFERROR(RANK({CH[2025]}{R[nm]},({lista}){orden})&" de 7","")')
        ws[f"E{r}"].font = F_SB
        ws[f"E{r}"].alignment = Alignment(horizontal="center")
    return ws

if __name__ == "__main__":
    d, refs = construir()
    import os
    print(d, os.path.getsize(d), "bytes")
    for k in ORDEN:
        print(f"  {k:<6} filas clave: ck={refs[k]['ck']} caja={refs[k]['caja']} at={refs[k]['at']}")


# ================================================================ FICHA POR TERMINAL
