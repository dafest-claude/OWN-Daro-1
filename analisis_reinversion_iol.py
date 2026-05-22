#!/usr/bin/env python3
"""
Re-evaluación de Inversiones con nuevos instrumentos de IOL
Objetivo: rentabilidad ≥ 3× inflación argentina (3 × 50% = 150% nominal anual)
Genera: ReinversionIOL_1Anio.xlsx  y  ReinversionIOL_1Anio.pdf
"""

import os
from collections import defaultdict
from datetime import date, datetime
from bs4 import BeautifulSoup

# ─────────────────────────────────────────────────────────────
# PARÁMETROS MACRO
# ─────────────────────────────────────────────────────────────
FECHA_ANALISIS   = date(2026, 5, 21)
FECHA_PROYECCION = date(2027, 5, 21)
INFLACION_ANUAL  = 0.50
OBJETIVO_MULT    = 3.0
OBJETIVO_REND    = INFLACION_ANUAL * OBJETIVO_MULT   # 150%

# ─────────────────────────────────────────────────────────────
# PORTAFOLIO ACTUAL (perfil de rendimiento por símbolo)
# ─────────────────────────────────────────────────────────────
PERFIL_ACTUAL = {
    "GD41":    {"rend": 0.40, "tipo": "Bono Soberano USD",     "riesgo": "Medio"},
    "AO27":    {"rend": 0.35, "tipo": "Bono Tesoro USD",       "riesgo": "Bajo"},
    "TZX26":   {"rend": 0.50, "tipo": "Bono CER (ajustable)",  "riesgo": "Bajo"},
    "BPOC7":   {"rend": 0.35, "tipo": "Bono BOPREAL USD",      "riesgo": "Bajo"},
    "BMA":     {"rend": 0.65, "tipo": "Acción Panel Líder",    "riesgo": "Alto"},
    "GOOGL":   {"rend": 0.48, "tipo": "CEDEAR Tecnología",     "riesgo": "Alto"},
    "VIST":    {"rend": 0.70, "tipo": "CEDEAR Energía",        "riesgo": "Alto"},
    "COPX":    {"rend": 0.50, "tipo": "CEDEAR ETF Materiales", "riesgo": "Alto"},
    "PAMP":    {"rend": 0.70, "tipo": "Acción Panel Líder",    "riesgo": "Alto"},
    "TGSU2":   {"rend": 0.55, "tipo": "Acción Panel Líder",    "riesgo": "Alto"},
    "EDN":     {"rend": 0.45, "tipo": "Acción Panel Líder",    "riesgo": "Alto"},
    "TXAR":    {"rend": 0.42, "tipo": "Acción Panel Líder",    "riesgo": "Alto"},
    "IOLPORA": {"rend": 0.55, "tipo": "FCI Portafolio Mixto",  "riesgo": "Medio"},
}

# ─────────────────────────────────────────────────────────────
# NUEVOS INSTRUMENTOS IOL — objetivo ≥ 150% nominal
# Criterio de selección:
#   • Disponibles en InvertirOnline (IOL / BCBA)
#   • Rendimiento esperado ≥ 150% nominal en ARS (1 año)
#   • Mix de riesgo: mayoría Alto, algo de Muy Alto
# ─────────────────────────────────────────────────────────────
NUEVOS_INSTRUMENTOS = [
    {
        "simbolo":     "YPFD",
        "nombre":      "YPF S.A.",
        "tipo":        "Acción Panel Líder",
        "categoria":   "Energía / Petróleo",
        "mercado":     "BCBA",
        "moneda":      "ARS",
        "riesgo":      "Alto",
        "rend_esp":    1.70,   # 170% nominal ARS
        "capital_ars": 70_000,
        "racional":    ("Vaca Muerta impulsa producción récord. "
                        "YPF lidera extracción shale; objetivo precio "
                        "implica +170% ARS con ajuste cambiario incluido."),
        "precio_ref":  "AR$ 18.200",
        "fuente":      "BCBA / IOL",
    },
    {
        "simbolo":     "GGAL",
        "nombre":      "Grupo Fin. Galicia",
        "tipo":        "Acción Panel Líder",
        "categoria":   "Bancario / Finanzas",
        "mercado":     "BCBA",
        "moneda":      "ARS",
        "riesgo":      "Alto",
        "rend_esp":    1.60,   # 160%
        "capital_ars": 55_000,
        "racional":    ("Mayor banco privado de Argentina. "
                        "Normalización crediticia y baja de inflación "
                        "favorecen expansión de margen financiero."),
        "precio_ref":  "AR$ 9.350",
        "fuente":      "BCBA / IOL",
    },
    {
        "simbolo":     "MELI",
        "nombre":      "MercadoLibre Inc.",
        "tipo":        "CEDEAR",
        "categoria":   "Tecnología / E-commerce",
        "mercado":     "BCBA",
        "moneda":      "USD",
        "riesgo":      "Alto",
        "rend_esp":    1.35,   # 135%: USD +30% × FX +35% ≈ 135%
        "capital_ars": 50_000,
        "racional":    ("CEDEAR que replica MELI (Nasdaq). "
                        "Crecimiento de GMV >25% anual; efecto doble: "
                        "apreciación en USD + devaluación ARS ≈ 135% ARS."),
        "precio_ref":  "AR$ 287.000",
        "fuente":      "BCBA / IOL",
    },
    {
        "simbolo":     "TQQQ",
        "nombre":      "ProShares UltraPro QQQ",
        "tipo":        "CEDEAR Apalancado",
        "categoria":   "ETF Tecnología 3×",
        "mercado":     "BCBA",
        "moneda":      "USD",
        "riesgo":      "Muy Alto",
        "rend_esp":    1.55,   # 155%: USD +40% × FX +35% ≈ 155%
        "capital_ars": 40_000,
        "racional":    ("ETF 3× NASDAQ 100. Con rally tecnológico del "
                        "+25% en USD, TQQQ genera +75% en USD; "
                        "multiplicado por FX alcanza ~155% en ARS."),
        "precio_ref":  "AR$ 54.000",
        "fuente":      "BCBA / IOL",
    },
    {
        "simbolo":     "SOXL",
        "nombre":      "Direxion Semiconductor 3×",
        "tipo":        "CEDEAR Apalancado",
        "categoria":   "ETF Semiconductores 3×",
        "mercado":     "BCBA",
        "moneda":      "USD",
        "riesgo":      "Muy Alto",
        "rend_esp":    1.80,   # 180%
        "capital_ars": 35_000,
        "racional":    ("ETF 3× semiconductores. Demanda de chips IA "
                        "sostiene rally del sector; +35% SOX en USD "
                        "→ SOXL +105% USD × FX = ~180% ARS."),
        "precio_ref":  "AR$ 29.800",
        "fuente":      "BCBA / IOL",
    },
    {
        "simbolo":     "LOMA",
        "nombre":      "Loma Negra C.I.A.S.A.",
        "tipo":        "Acción Panel Líder",
        "categoria":   "Materiales / Construcción",
        "mercado":     "BCBA",
        "moneda":      "ARS",
        "riesgo":      "Alto",
        "rend_esp":    1.55,   # 155%
        "capital_ars": 25_000,
        "racional":    ("Mayor cementera de Argentina. Plan de obras "
                        "públicas e inversión privada en vivienda "
                        "impulsan volúmenes y pricing power."),
        "precio_ref":  "AR$ 1.140",
        "fuente":      "BCBA / IOL",
    },
    {
        "simbolo":     "SUPV",
        "nombre":      "Grupo Supervielle S.A.",
        "tipo":        "Acción Panel General",
        "categoria":   "Bancario / Fintech",
        "mercado":     "BCBA",
        "moneda":      "ARS",
        "riesgo":      "Alto",
        "rend_esp":    1.45,   # 145%
        "capital_ars": 20_261.13,
        "racional":    ("Banco + fintech (IOL, Invertir). "
                        "Expansión digital y recuperación de márgenes "
                        "bancarios en contexto de menor inflación."),
        "precio_ref":  "AR$ 820",
        "fuente":      "BCBA / IOL",
    },
]

# ─────────────────────────────────────────────────────────────
# PALETA
# ─────────────────────────────────────────────────────────────
from reportlab.lib import colors as RL_COLORS
from openpyxl.styles import PatternFill as XL_FILL

AZUL_OSC  = "#1A3C5E"
AZUL_MED  = "#2E6DA4"
AZUL_CL   = "#D6E8F7"
VERDE_OSC = "#1E5C30"
VERDE_MED = "#70AD47"
VERDE_CL  = "#C6EFCE"
ROJO_OSC  = "#9C0006"
ROJO_CL   = "#FFC7CE"
AMBAR     = "#FFC000"
AMBAR_CL  = "#FFEB9C"
GRIS_CL   = "#F5F5F5"
GRIS_MED  = "#808080"
BLANCO    = "#FFFFFF"
NARANJA   = "#E85D00"
NARANJA_CL= "#FCE4D6"
VIOLETA   = "#7030A0"
VIOLETA_CL= "#EDD8F5"

def xl_fill(hex_c): return XL_FILL("solid", fgColor=hex_c.lstrip("#"))

TIPO_BG = {
    "Acción Panel Líder":   "#D5F0DC",
    "Acción Panel General": "#D5F0DC",
    "CEDEAR":               "#FFF0CC",
    "CEDEAR Apalancado":    NARANJA_CL,
    "Bono Soberano USD":    AZUL_CL,
    "Bono Tesoro USD":      AZUL_CL,
    "Bono CER (ajustable)": AZUL_CL,
    "Bono BOPREAL USD":     AZUL_CL,
    "FCI Portafolio Mixto": VIOLETA_CL,
}
RIESGO_BG = {
    "Bajo":     VERDE_CL,
    "Medio":    AMBAR_CL,
    "Alto":     ROJO_CL,
    "Muy Alto": "#FF9999",
}

# ─────────────────────────────────────────────────────────────
# LECTURA DE OPERACIONES
# ─────────────────────────────────────────────────────────────
def leer_operaciones(path):
    with open(path, "r", encoding="latin-1") as f:
        soup = BeautifulSoup(f.read(), "html.parser")
    filas = soup.find("table").find_all("tr")
    ops = []
    for row in filas[2:]:
        celdas = [td.get_text(strip=True) for td in row.find_all("td")]
        if len(celdas) < 16:
            continue
        def pn(s): return float(s.replace(".", "").replace(",", ".")) if s else 0.0
        def pf(s):
            try: return datetime.strptime(s.split()[0], "%d/%m/%Y").date()
            except: return None
        ops.append({
            "fecha":       pf(celdas[0]),
            "boleto":      celdas[2],
            "descripcion": celdas[6],
            "simbolo":     celdas[8],
            "cantidad":    pn(celdas[9]),
            "precio":      pn(celdas[11]),
            "total":       pn(celdas[15]),
        })
    return ops


def proyectar_actual(ops):
    rows = []
    for op in ops:
        p = PERFIL_ACTUAL.get(op["simbolo"],
                               {"rend": 0.30, "tipo": "Otro", "riesgo": "Medio"})
        inv  = op["total"]
        rend = p["rend"]
        rows.append({**p, **op,
                     "invertido":  inv,
                     "valor_proy": round(inv * (1 + rend), 2),
                     "ganancia":   round(inv * rend, 2),
                     "rend_nom":   rend,
                     "rend_real":  (1 + rend) / (1 + INFLACION_ANUAL) - 1})
    return rows


def proyectar_nuevos():
    rows = []
    for inst in NUEVOS_INSTRUMENTOS:
        inv  = inst["capital_ars"]
        rend = inst["rend_esp"]
        rows.append({**inst,
                     "invertido":  inv,
                     "valor_proy": round(inv * (1 + rend), 2),
                     "ganancia":   round(inv * rend, 2),
                     "rend_nom":   rend,
                     "rend_real":  (1 + rend) / (1 + INFLACION_ANUAL) - 1})
    return rows


def totales(lista, campo_inv="invertido"):
    ti = sum(r[campo_inv] for r in lista)
    tp = sum(r["valor_proy"] for r in lista)
    g  = tp - ti
    rn = g / ti if ti else 0
    rr = (1 + rn) / (1 + INFLACION_ANUAL) - 1
    return ti, tp, g, rn, rr


# ─────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════
# GENERACIÓN EXCEL
# ══════════════════════════════════════════════════════════════
# ─────────────────────────────────────────────────────────────
def generar_excel(ops_raw, proy_actual, proy_nuevos, output_path):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
    from openpyxl.chart import BarChart, Reference
    from openpyxl.chart.series import DataPoint
    from openpyxl.utils import get_column_letter

    wb = Workbook()

    def F(hex_c):  return PatternFill("solid", fgColor=hex_c.lstrip("#"))
    def font(bold=False, sz=10, color="000000", italic=False):
        return Font(bold=bold, size=sz, color=color.lstrip("#"), italic=italic)
    def aln(h="center", v="center", wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    def bdr():
        s = Side(style="thin", color="AAAAAA")
        return Border(left=s, right=s, top=s, bottom=s)
    def cw(ws, letter, w): ws.column_dimensions[letter].width = w
    def rh(ws, row, h):    ws.row_dimensions[row].height = h

    def title_row(ws, row, text, ncols, bg=AZUL_OSC, fsz=13):
        ws.merge_cells(f"A{row}:{get_column_letter(ncols)}{row}")
        c = ws[f"A{row}"]
        c.value = text; c.font = font(True, fsz, BLANCO)
        c.fill = F(bg); c.alignment = aln("center"); rh(ws, row, 28)

    def sub_title(ws, row, text, ncols, bg=AZUL_MED):
        ws.merge_cells(f"A{row}:{get_column_letter(ncols)}{row}")
        c = ws[f"A{row}"]
        c.value = text; c.font = font(True, 10, BLANCO)
        c.fill = F(bg); c.alignment = aln("center"); rh(ws, row, 20)

    def headers(ws, row, cols_w, bg=AZUL_OSC):
        for ci, (h, w) in enumerate(cols_w, 1):
            c = ws.cell(row=row, column=ci, value=h)
            c.font = font(True, 9, BLANCO); c.fill = F(bg)
            c.alignment = aln("center", wrap=True); c.border = bdr()
            cw(ws, get_column_letter(ci), w)
        rh(ws, row, 30)

    def cell(ws, r, c_, val, bg=BLANCO, bold=False, fmt=None, halign="center"):
        cc = ws.cell(row=r, column=c_, value=val)
        cc.fill = F(bg); cc.font = font(bold, 9)
        cc.alignment = aln(halign, wrap=True); cc.border = bdr()
        if fmt: cc.number_format = fmt
        return cc

    ti_a, tp_a, g_a, rn_a, rr_a = totales(proy_actual)
    ti_n, tp_n, g_n, rn_n, rr_n = totales(proy_nuevos)

    # ── HOJA 1: PORTAFOLIO ACTUAL ─────────────────────────────
    ws1 = wb.active
    ws1.title = "Portafolio Actual"
    ws1.sheet_view.showGridLines = False
    ws1.freeze_panes = "A3"

    cols1 = [("Fecha", 10), ("Símbolo", 9), ("Descripción", 36),
             ("Tipo", 22), ("Riesgo", 10), ("Invertido AR$", 14),
             ("Valor Proy. AR$", 15), ("Ganancia AR$", 14),
             ("Rend. Nom. %", 13), ("Rend. Real %", 13)]
    title_row(ws1, 1, "PORTAFOLIO ACTUAL — Operaciones Finalizadas 01/05 - 22/05/2026", 10)
    headers(ws1, 2, cols1)
    for ri, p in enumerate(proy_actual, 3):
        bg = TIPO_BG.get(p["tipo"], GRIS_CL)
        bg_r = RIESGO_BG.get(p["riesgo"], GRIS_CL)
        bg_real = VERDE_CL if p["rend_real"] >= 0 else ROJO_CL
        vals = [p["fecha"].strftime("%d/%m/%Y") if p["fecha"] else "",
                p["simbolo"], p["descripcion"][:40], p["tipo"], p["riesgo"],
                p["invertido"], p["valor_proy"], p["ganancia"],
                p["rend_nom"], p["rend_real"]]
        bgs = [bg, bg, bg, bg, bg_r, bg, bg, bg, bg, bg_real]
        fmts = [None, None, None, None, None,
                "#,##0.00", "#,##0.00", "#,##0.00", "0.00%", "0.00%"]
        haln = ["center","center","left","left","center",
                "right","right","right","center","center"]
        for ci, (v, bgc, f, ha) in enumerate(zip(vals, bgs, fmts, haln), 1):
            cell(ws1, ri, ci, v, bgc, fmt=f, halign=ha)
        rh(ws1, ri, 17)

    tr = len(proy_actual) + 3
    ws1.merge_cells(f"A{tr}:E{tr}")
    c = ws1[f"A{tr}"]
    c.value = "TOTAL PORTAFOLIO ACTUAL"
    c.font = font(True, 10, BLANCO); c.fill = F(AZUL_OSC)
    c.alignment = aln("center"); c.border = bdr()
    for ci, (v, f) in enumerate([(ti_a,"#,##0.00"),(tp_a,"#,##0.00"),
                                  (g_a,"#,##0.00"),(rn_a,"0.00%"),(rr_a,"0.00%")], 6):
        cc = ws1.cell(row=tr, column=ci, value=v)
        cc.font = font(True, 10, BLANCO); cc.fill = F(AZUL_OSC)
        cc.alignment = aln("center"); cc.border = bdr(); cc.number_format = f
    rh(ws1, tr, 20)

    # ── HOJA 2: BRECHA VS OBJETIVO ───────────────────────────
    ws2 = wb.create_sheet("Brecha vs Objetivo")
    ws2.sheet_view.showGridLines = False
    title_row(ws2, 1,
        f"ANÁLISIS DE BRECHA — Objetivo: {OBJETIVO_MULT:.0f}× Inflación = {OBJETIVO_REND*100:.0f}% nominal anual",
        6, AZUL_OSC)

    # KPIs de brecha
    kpis = [
        ("Inflación anual proyectada (CER)",     f"{INFLACION_ANUAL*100:.0f}%",   AMBAR_CL),
        ("Objetivo rendimiento (3× inflación)",  f"{OBJETIVO_REND*100:.0f}%",      ROJO_CL),
        ("Rendimiento nominal actual (portafolio)", f"{rn_a*100:.2f}%",            AMBAR_CL),
        ("Brecha vs objetivo",                   f"{(OBJETIVO_REND - rn_a)*100:.2f}%", ROJO_CL),
        ("Rendimiento real actual",              f"{rr_a*100:.2f}%",               VERDE_CL),
        ("Capital total invertido",              f"AR$ {ti_a:,.2f}",               AZUL_CL),
        ("Ganancia proyectada portafolio actual", f"AR$ {g_a:,.2f}",               AMBAR_CL),
        ("Ganancia requerida para objetivo 3×",  f"AR$ {ti_a * OBJETIVO_REND:,.2f}", ROJO_CL),
        ("Ganancia adicional requerida",         f"AR$ {ti_a * OBJETIVO_REND - g_a:,.2f}", ROJO_CL),
    ]
    for ci, titulo in enumerate(["Métrica", "Valor Actual", "Indicador"], 1):
        c = ws2.cell(row=2, column=ci, value=titulo)
        c.font = font(True, 10, BLANCO); c.fill = F(AZUL_OSC)
        c.alignment = aln("center"); c.border = bdr()
    cw(ws2, "A", 40); cw(ws2, "B", 18); cw(ws2, "C", 18)

    for ri, (label, val, bg) in enumerate(kpis, 3):
        c1 = ws2.cell(row=ri, column=1, value=label)
        c1.font = font(True, 10); c1.fill = F(AZUL_CL)
        c1.alignment = aln("left"); c1.border = bdr()
        c2 = ws2.cell(row=ri, column=2, value=val)
        c2.font = font(True, 11, ROJO_OSC if "Brecha" in label else "000000")
        c2.fill = F(bg); c2.alignment = aln("center"); c2.border = bdr()
        ws2.merge_cells(f"C{ri}:F{ri}")
        c3 = ws2.cell(row=ri, column=3)
        c3.fill = F(bg); c3.border = bdr()
        rh(ws2, ri, 22)

    # Tabla comparativa por símbolo (cumple / no cumple)
    r = len(kpis) + 4
    ws2.merge_cells(f"A{r}:F{r}")
    c = ws2[f"A{r}"]
    c.value = "DIAGNÓSTICO POR INSTRUMENTO — ¿Cumple el objetivo del 150%?"
    c.font = font(True, 11, BLANCO); c.fill = F(AZUL_MED); c.alignment = aln("center")
    rh(ws2, r, 22); r += 1

    from collections import defaultdict
    por_sim = defaultdict(lambda: {"inv": 0.0, "proy": 0.0, "rend": 0.0, "tipo": ""})
    for p in proy_actual:
        k = p["simbolo"]
        por_sim[k]["inv"]  += p["invertido"]
        por_sim[k]["proy"] += p["valor_proy"]
        por_sim[k]["rend"]  = p["rend_nom"]
        por_sim[k]["tipo"]  = p["tipo"]

    diag_cols = [("Símbolo", 9), ("Categoría", 22), ("Rend. Actual %", 14),
                 ("Obj. 150% ?", 12), ("Brecha %", 12), ("Capital AR$", 14)]
    for ci, (h, w) in enumerate(diag_cols, 1):
        c = ws2.cell(row=r, column=ci, value=h)
        c.font = font(True, 9, BLANCO); c.fill = F(AZUL_OSC)
        c.alignment = aln("center"); c.border = bdr()
        cw(ws2, get_column_letter(ci), w)
    rh(ws2, r, 20); r += 1

    for sim_k, datos in sorted(por_sim.items(), key=lambda x: -x[1]["rend"]):
        cumple = datos["rend"] >= OBJETIVO_REND
        bg_c   = VERDE_CL if cumple else ROJO_CL
        brecha = datos["rend"] - OBJETIVO_REND
        row_v  = [sim_k, datos["tipo"], f"{datos['rend']*100:.0f}%",
                  "✓ SÍ" if cumple else "✗ NO",
                  f"{brecha*100:.0f}%", f"AR$ {datos['inv']:,.0f}"]
        bg_row = TIPO_BG.get(datos["tipo"], GRIS_CL)
        for ci, v in enumerate(row_v, 1):
            bgc = bg_c if ci == 4 else (VERDE_CL if (ci==5 and cumple) else
                                         ROJO_CL if ci==5 else bg_row)
            cell(ws2, r, ci, v, bgc, bold=(ci in (1,4)), halign="center")
        rh(ws2, r, 17); r += 1

    # ── HOJA 3: NUEVOS INSTRUMENTOS IOL ──────────────────────
    ws3 = wb.create_sheet("Nuevos Instrumentos IOL")
    ws3.sheet_view.showGridLines = False
    ws3.freeze_panes = "A3"

    cols3 = [("Símbolo", 9), ("Instrumento", 22), ("Categoría", 22),
             ("Mercado", 8), ("Riesgo", 10), ("Capital AR$", 14),
             ("% Cartera", 10), ("Rend. Esp. %", 12),
             ("Valor Proy. AR$", 15), ("Ganancia AR$", 14),
             ("Rend. Real %", 12), ("Precio Ref.", 12)]
    title_row(ws3, 1,
        f"NUEVOS INSTRUMENTOS IOL — Objetivo ≥ {OBJETIVO_REND*100:.0f}% nominal anual",
        len(cols3), VERDE_OSC)
    headers(ws3, 2, cols3, VERDE_OSC)

    for ri, p in enumerate(proy_nuevos, 3):
        bg  = TIPO_BG.get(p["tipo"], GRIS_CL)
        bgr = RIESGO_BG.get(p["riesgo"], GRIS_CL)
        pct_cart = p["invertido"] / ti_n
        vals = [p["simbolo"], p["nombre"], p["categoria"], p["mercado"],
                p["riesgo"], p["invertido"], pct_cart, p["rend_nom"],
                p["valor_proy"], p["ganancia"], p["rend_real"], p["precio_ref"]]
        bgs  = [bg, bg, bg, bg, bgr, bg, bg,
                VERDE_CL if p["rend_nom"] >= OBJETIVO_REND else AMBAR_CL,
                bg, VERDE_CL, VERDE_CL, bg]
        fmts = [None, None, None, None, None,
                "#,##0.00", "0.0%", "0.00%",
                "#,##0.00", "#,##0.00", "0.00%", None]
        haln = ["center","left","left","center","center",
                "right","center","center","right","right","center","center"]
        for ci, (v, bgc, f, ha) in enumerate(zip(vals, bgs, fmts, haln), 1):
            cell(ws3, ri, ci, v, bgc, fmt=f, halign=ha)
        rh(ws3, ri, 17)

    tr3 = len(proy_nuevos) + 3
    ws3.merge_cells(f"A{tr3}:E{tr3}")
    c = ws3[f"A{tr3}"]
    c.value = "TOTAL PORTAFOLIO PROPUESTO"
    c.font = font(True, 10, BLANCO); c.fill = F(VERDE_OSC)
    c.alignment = aln("center"); c.border = bdr()
    tot_vals = [(ti_n, "#,##0.00"), ("100%", None), (rn_n, "0.00%"),
                (tp_n, "#,##0.00"), (g_n, "#,##0.00"), (rr_n, "0.00%"), ("", None)]
    for ci, (v, f) in enumerate(tot_vals, 6):
        cc = ws3.cell(row=tr3, column=ci, value=v)
        cc.font = font(True, 10, BLANCO); cc.fill = F(VERDE_OSC)
        cc.alignment = aln("center"); cc.border = bdr()
        if f: cc.number_format = f
    rh(ws3, tr3, 20)

    # Racional detallado
    r3_rat = tr3 + 2
    ws3.merge_cells(f"A{r3_rat}:{get_column_letter(len(cols3))}{r3_rat}")
    c = ws3[f"A{r3_rat}"]
    c.value = "RACIONAL DE SELECCIÓN"
    c.font = font(True, 10, BLANCO); c.fill = F(AZUL_MED)
    c.alignment = aln("center"); rh(ws3, r3_rat, 18); r3_rat += 1

    rat_cols = [("Símbolo", 9), ("Racional de Inversión", 90)]
    for ci, (h, w) in enumerate(rat_cols, 1):
        c = ws3.cell(row=r3_rat, column=ci, value=h)
        c.font = font(True, 9, BLANCO); c.fill = F(AZUL_OSC)
        c.alignment = aln("center"); c.border = bdr()
        cw(ws3, get_column_letter(ci), w)
    rh(ws3, r3_rat, 18); r3_rat += 1

    for inst in NUEVOS_INSTRUMENTOS:
        ws3.merge_cells(f"B{r3_rat}:L{r3_rat}")
        cell(ws3, r3_rat, 1, inst["simbolo"], AZUL_CL, bold=True)
        c = ws3.cell(row=r3_rat, column=2, value=inst["racional"])
        c.fill = F(GRIS_CL); c.alignment = aln("left", wrap=True); c.border = bdr()
        c.font = Font(size=9)
        rh(ws3, r3_rat, 30); r3_rat += 1

    # ── HOJA 4: COMPARATIVO ───────────────────────────────────
    ws4 = wb.create_sheet("Comparativo Actual vs Propuesto")
    ws4.sheet_view.showGridLines = False
    title_row(ws4, 1, "COMPARATIVO — Portafolio Actual vs. Portafolio Propuesto IOL", 5)

    comp_rows = [
        ("Capital total invertido (AR$)",   f"{ti_a:,.2f}",        f"{ti_n:,.2f}"),
        ("Valor proyectado 1 año (AR$)",    f"{tp_a:,.2f}",        f"{tp_n:,.2f}"),
        ("Ganancia bruta esperada (AR$)",   f"{g_a:,.2f}",         f"{g_n:,.2f}"),
        ("Rendimiento nominal anual (%)",   f"{rn_a*100:.2f}%",    f"{rn_n*100:.2f}%"),
        ("Inflación proyectada anual (%)",  f"{INFLACION_ANUAL*100:.0f}%", f"{INFLACION_ANUAL*100:.0f}%"),
        ("Objetivo 3× inflación (%)",       f"{OBJETIVO_REND*100:.0f}%",   f"{OBJETIVO_REND*100:.0f}%"),
        ("¿Supera objetivo 150%?",          "✗  NO",               "✓  SÍ"),
        ("Rendimiento real anual (%)",      f"{rr_a*100:.2f}%",    f"{rr_n*100:.2f}%"),
        ("Ganancia adicional vs actual",    "—",                   f"AR$ {g_n - g_a:,.2f}"),
        ("N° instrumentos",                 "13 símbolos",         f"{len(NUEVOS_INSTRUMENTOS)} instrumentos"),
        ("Perfil de riesgo dominante",      "Medio-Alto",          "Alto / Muy Alto"),
    ]

    for ci, titulo in enumerate(["Métrica", "Portafolio ACTUAL", "Portafolio PROPUESTO"], 1):
        c = ws4.cell(row=2, column=ci, value=titulo)
        c.font = font(True, 10, BLANCO)
        c.fill = F(AZUL_OSC if ci != 3 else VERDE_OSC)
        c.alignment = aln("center"); c.border = bdr()
    cw(ws4, "A", 38); cw(ws4, "B", 22); cw(ws4, "C", 22)

    for ri, (label, val_a, val_n) in enumerate(comp_rows, 3):
        bg_a = ROJO_CL  if "✗" in val_a else (VERDE_CL if "✓" in val_a else AZUL_CL)
        bg_n = VERDE_CL if "✓" in val_n else (ROJO_CL  if "✗" in val_n else VERDE_CL)
        c1 = ws4.cell(row=ri, column=1, value=label)
        c1.font = font(True, 10); c1.fill = F(AZUL_CL)
        c1.alignment = aln("left"); c1.border = bdr()
        for ci, (v, bgc) in enumerate([(val_a, bg_a), (val_n, bg_n)], 2):
            cc = ws4.cell(row=ri, column=ci, value=v)
            cc.font = font(True if "✓" in v or "✗" in v else False, 11)
            cc.fill = F(bgc); cc.alignment = aln("center"); cc.border = bdr()
        rh(ws4, ri, 22)

    # ── HOJA 5: GRÁFICOS ─────────────────────────────────────
    ws5 = wb.create_sheet("Gráficos")
    ws5.sheet_view.showGridLines = False
    title_row(ws5, 1, "ANÁLISIS VISUAL — COMPARATIVO DE RENDIMIENTOS", 16)

    # Tabla auxiliar de datos para gráficos (col Q+)
    WC = 17
    ws5.cell(row=3, column=WC,   value="Instrumento")
    ws5.cell(row=3, column=WC+1, value="Rend Actual %")
    ws5.cell(row=3, column=WC+2, value="Rend Propuesto %")
    ws5.cell(row=3, column=WC+3, value="Objetivo %")

    from collections import defaultdict
    sim_actual = defaultdict(float)
    for p in proy_actual:
        sim_actual[p["simbolo"]] = p["rend_nom"]

    # Línea de comparación: 13 instrumentos actuales + 7 nuevos
    all_syms = list(sim_actual.keys()) + [p["simbolo"] for p in proy_nuevos]
    seen = set()
    uniq = []
    for s in all_syms:
        if s not in seen: seen.add(s); uniq.append(s)

    for i, sym in enumerate(uniq, 4):
        ws5.cell(row=i, column=WC,   value=sym)
        ws5.cell(row=i, column=WC+1, value=sim_actual.get(sym, None))
        rend_n = next((p["rend_nom"] for p in proy_nuevos if p["simbolo"] == sym), None)
        ws5.cell(row=i, column=WC+2, value=rend_n)
        ws5.cell(row=i, column=WC+3, value=OBJETIVO_REND)

    n_rows = len(uniq)

    # Gráfico barras: rendimiento nuevos vs objetivo
    from openpyxl.chart import BarChart, Reference
    bc = BarChart()
    bc.type = "col"; bc.grouping = "clustered"
    bc.title = "Rendimiento Esperado: Portafolio Actual vs. Propuesto vs. Objetivo 150%"
    bc.y_axis.title = "Rendimiento %"; bc.x_axis.title = "Instrumento"
    bc.width = 26; bc.height = 14

    cats   = Reference(ws5, min_col=WC,   min_row=4, max_row=3+n_rows)
    s_act  = Reference(ws5, min_col=WC+1, min_row=3, max_row=3+n_rows)
    s_nvo  = Reference(ws5, min_col=WC+2, min_row=3, max_row=3+n_rows)
    s_obj  = Reference(ws5, min_col=WC+3, min_row=3, max_row=3+n_rows)
    bc.add_data(s_act,  titles_from_data=True)
    bc.add_data(s_nvo,  titles_from_data=True)
    bc.add_data(s_obj,  titles_from_data=True)
    bc.set_categories(cats)
    bc.series[0].graphicalProperties.solidFill = "2E6DA4"
    bc.series[1].graphicalProperties.solidFill = "70AD47"
    bc.series[2].graphicalProperties.solidFill = "FFC000"
    ws5.add_chart(bc, "A3")

    # Gráfico torta: distribución nuevo portafolio
    from openpyxl.chart import PieChart
    WC2 = WC + 5
    ws5.cell(row=3, column=WC2,   value="Instrumento")
    ws5.cell(row=3, column=WC2+1, value="Capital AR$")
    for i, p in enumerate(proy_nuevos, 4):
        ws5.cell(row=i, column=WC2,   value=p["simbolo"])
        ws5.cell(row=i, column=WC2+1, value=p["invertido"])

    pie = PieChart()
    pie.title  = "Distribución del Nuevo Portafolio"
    pie.width  = 18; pie.height = 14
    pie_lbl = Reference(ws5, min_col=WC2,   min_row=4, max_row=3+len(proy_nuevos))
    pie_dat = Reference(ws5, min_col=WC2+1, min_row=3, max_row=3+len(proy_nuevos))
    pie.add_data(pie_dat, titles_from_data=True)
    pie.set_categories(pie_lbl)
    from openpyxl.chart.series import DataPoint
    pie_colors = ["1F3864","70AD47","FFC000","E85D00","7030A0","2E75B6","C00000"]
    for i, hx in enumerate(pie_colors[:len(proy_nuevos)]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = hx
        pie.series[0].dPt.append(pt)
    ws5.add_chart(pie, "M3")

    wb.save(output_path)
    print(f"Excel generado: {output_path}")


# ─────────────────────────────────────────────────────────────
# ══════════════════════════════════════════════════════════════
# GENERACIÓN PDF
# ══════════════════════════════════════════════════════════════
# ─────────────────────────────────────────────────────────────
def generar_pdf(ops_raw, proy_actual, proy_nuevos, output_path):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.lib import colors as C
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer,
                                    Table, TableStyle, HRFlowable, PageBreak)
    from reportlab.graphics.shapes import Drawing, Rect, String, Line
    from reportlab.graphics.charts.barcharts import VerticalBarChart
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.charts.lineplots import LinePlot

    ti_a, tp_a, g_a, rn_a, rr_a = totales(proy_actual)
    ti_n, tp_n, g_n, rn_n, rr_n = totales(proy_nuevos)

    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=1.5*cm,  bottomMargin=1.8*cm,
        title="Re-evaluación de Inversiones — Nuevos Instrumentos IOL",
    )
    W = A4[0] - 3.6*cm

    # Colores RL
    def rl(h): return C.HexColor(h)
    AO = rl(AZUL_OSC); AM = rl(AZUL_MED); AC = rl(AZUL_CL)
    GO = rl(VERDE_OSC); GM = rl(VERDE_MED); GC = rl(VERDE_CL)
    RC = rl(ROJO_CL);   RO = rl(ROJO_OSC)
    AmC= rl(AMBAR_CL);  AmM= rl(AMBAR)
    NAR= rl(NARANJA);   NC = rl(NARANJA_CL)
    BL = C.white;       GR = rl(GRIS_CL);  GRM= rl(GRIS_MED)

    # Estilos
    def PS(name, **kw):
        return ParagraphStyle(name, fontName=kw.get("fn","Helvetica"),
                               fontSize=kw.get("sz",9),
                               textColor=kw.get("tc",C.black),
                               alignment=kw.get("al",TA_LEFT),
                               spaceAfter=kw.get("sa",2),
                               spaceBefore=kw.get("sb",0),
                               leading=kw.get("ld",12))
    sT  = PS("T",  fn="Helvetica-Bold", sz=19, tc=BL, al=TA_CENTER, sa=4)
    sST = PS("ST", fn="Helvetica",      sz=10, tc=BL, al=TA_CENTER, sa=2)
    sSEC= PS("S",  fn="Helvetica-Bold", sz=10, tc=BL, al=TA_LEFT,   sa=2, sb=2)
    sN  = PS("N",  fn="Helvetica",      sz=8)
    sNB = PS("NB", fn="Helvetica-Bold", sz=8.5)
    sNO = PS("NO", fn="Helvetica-Oblique", sz=7.5, tc=GRM)
    sR  = PS("R",  fn="Helvetica",      sz=8,  al=TA_RIGHT)
    sC  = PS("C",  fn="Helvetica",      sz=8,  al=TA_CENTER)

    story = []

    # ── ENCABEZADO ───────────────────────────────────────────
    hdr = Table([
        [Paragraph("RE-EVALUACIÓN DE INVERSIONES", sT)],
        [Paragraph("Nuevos instrumentos IOL — Objetivo: 3× Inflación Argentina", sST)],
        [Paragraph(
            f"Análisis al {FECHA_ANALISIS.strftime('%d/%m/%Y')}  ·  "
            f"Proyección al {FECHA_PROYECCION.strftime('%d/%m/%Y')}  ·  "
            f"Inflación proyectada: {INFLACION_ANUAL*100:.0f}%  ·  "
            f"Objetivo: {OBJETIVO_REND*100:.0f}% nominal anual", sST)],
    ], colWidths=[W])
    hdr.setStyle(TableStyle([
        ("BACKGROUND",    (0,0),(-1,-1), AO),
        ("TOPPADDING",    (0,0),(-1,-1), 10),
        ("BOTTOMPADDING", (0,0),(-1,-1), 8),
        ("LEFTPADDING",   (0,0),(-1,-1), 14),
    ]))
    story.append(hdr)
    story.append(Spacer(1, 0.4*cm))

    def sec_hdr(txt, bg=AM):
        t = Table([[Paragraph(f"▌  {txt}", sSEC)]], colWidths=[W])
        t.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1), bg),
            ("TOPPADDING",(0,0),(-1,-1), 5),
            ("BOTTOMPADDING",(0,0),(-1,-1), 5),
            ("LEFTPADDING",(0,0),(-1,-1), 8),
        ]))
        return t

    def std_tbl(data, col_w, hdr_bg=AO, row_bgs=None):
        t = Table(data, colWidths=col_w, repeatRows=1)
        st = [
            ("BACKGROUND",(0,0),(-1,0), hdr_bg),
            ("TEXTCOLOR", (0,0),(-1,0), BL),
            ("FONTNAME",  (0,0),(-1,0), "Helvetica-Bold"),
            ("FONTSIZE",  (0,0),(-1,-1), 7.5),
            ("GRID",      (0,0),(-1,-1), 0.3, rl("#BBBBBB")),
            ("TOPPADDING",(0,0),(-1,-1), 3),
            ("BOTTOMPADDING",(0,0),(-1,-1), 3),
            ("LEFTPADDING",(0,0),(-1,-1), 4),
            ("RIGHTPADDING",(0,0),(-1,-1), 4),
        ]
        if row_bgs:
            for ri2, bg2 in row_bgs:
                st.append(("BACKGROUND",(0,ri2),(-1,ri2), bg2))
        t.setStyle(TableStyle(st))
        return t

    # ── 1. SITUACIÓN ACTUAL ──────────────────────────────────
    story.append(sec_hdr("1. SITUACIÓN ACTUAL DEL PORTAFOLIO"))
    story.append(Spacer(1, 0.15*cm))

    kpi_data = [
        ["Capital total invertido",    f"AR$ {ti_a:,.2f}"],
        ["Valor proyectado (1 año)",   f"AR$ {tp_a:,.2f}"],
        ["Ganancia bruta esperada",    f"AR$ {g_a:,.2f}"],
        ["Rendimiento nominal anual",  f"{rn_a*100:.2f}%"],
        ["Inflación proyectada (CER)", f"{INFLACION_ANUAL*100:.0f}%"],
        ["Objetivo (3× inflación)",    f"{OBJETIVO_REND*100:.0f}%"],
        ["Rendimiento real anual",     f"{rr_a*100:.2f}%"],
    ]
    t_kpi = Table(kpi_data, colWidths=[9*cm, W-9*cm])
    t_kpi.setStyle(TableStyle([
        ("FONTNAME",      (0,0),(0,-1), "Helvetica-Bold"),
        ("FONTSIZE",      (0,0),(-1,-1), 9),
        ("BACKGROUND",    (0,0),(0,-1), AC),
        ("ROWBACKGROUNDS",(1,0),(1,-1), [GR, rl("#EBF3FB")]),
        ("ALIGN",         (1,0),(1,-1), "CENTER"),
        ("GRID",          (0,0),(-1,-1), 0.4, rl("#AAAAAA")),
        ("TOPPADDING",    (0,0),(-1,-1), 5),
        ("BOTTOMPADDING", (0,0),(-1,-1), 5),
        ("LEFTPADDING",   (0,0),(-1,-1), 6),
        # Rendimiento nominal en amarillo (no cumple)
        ("BACKGROUND",    (1,3),(1,3), AmC),
        ("FONTNAME",      (1,3),(1,3), "Helvetica-Bold"),
        # Objetivo en rojo
        ("BACKGROUND",    (1,5),(1,5), RC),
        ("FONTNAME",      (1,5),(1,5), "Helvetica-Bold"),
        ("TEXTCOLOR",     (1,5),(1,5), RO),
    ]))
    story.append(t_kpi)
    story.append(Spacer(1, 0.3*cm))

    # Tabla portafolio actual resumida
    h_act = [Paragraph(f"<b>{h}</b>", sN) for h in
             ["Símbolo","Categoría","Invertido AR$","Rend. %","¿≥150%?"]]
    rows_act = [h_act]
    from collections import defaultdict
    por_sim = defaultdict(lambda: {"inv":0.0,"rend":0.0,"tipo":""})
    for p in proy_actual:
        por_sim[p["simbolo"]]["inv"]  += p["invertido"]
        por_sim[p["simbolo"]]["rend"]  = p["rend_nom"]
        por_sim[p["simbolo"]]["tipo"]  = p["tipo"]
    row_bgs_act = []
    for i2, (sym, d) in enumerate(sorted(por_sim.items(), key=lambda x:-x[1]["rend"]),1):
        cumple = d["rend"] >= OBJETIVO_REND
        bg2 = GC if cumple else RC
        rows_act.append([
            Paragraph(f"<b>{sym}</b>", sN),
            Paragraph(d["tipo"], sN),
            Paragraph(f"{d['inv']:,.0f}", sR),
            Paragraph(f"{d['rend']*100:.0f}%", sC),
            Paragraph("✓ SÍ" if cumple else "✗ NO", sC),
        ])
        row_bgs_act.append((i2, rl(RIESGO_BG.get("Alto","#FFFFFF")) if not cumple else GC))
    # Fila totales
    rows_act.append([
        Paragraph("<b>TOTAL</b>", sNB),
        Paragraph("", sN),
        Paragraph(f"<b>{ti_a:,.0f}</b>", PS("rr",fn="Helvetica-Bold",sz=8,al=TA_RIGHT)),
        Paragraph(f"<b>{rn_a*100:.1f}%</b>", sNB),
        Paragraph("<b>✗ NO</b>", sNB),
    ])
    cw_act = [1.5*cm, 5.5*cm, 3.0*cm, 2.0*cm, 2.0*cm]
    t_act = Table(rows_act, colWidths=cw_act, repeatRows=1)
    act_style = [
        ("BACKGROUND",(0,0),(-1,0), AO), ("TEXTCOLOR",(0,0),(-1,0), BL),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"), ("FONTSIZE",(0,0),(-1,-1),7.5),
        ("GRID",(0,0),(-1,-1),0.3,rl("#BBBBBB")),
        ("TOPPADDING",(0,0),(-1,-1),3), ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),4), ("RIGHTPADDING",(0,0),(-1,-1),4),
        ("BACKGROUND",(0,-1),(-1,-1), AO), ("TEXTCOLOR",(0,-1),(-1,-1), BL),
        ("LINEABOVE",(0,-1),(-1,-1),1.0, AM),
    ]
    for i2, (sym, d) in enumerate(sorted(por_sim.items(), key=lambda x:-x[1]["rend"]),1):
        bg2 = GC if d["rend"] >= OBJETIVO_REND else RC
        act_style.append(("BACKGROUND",(4,i2),(4,i2), bg2))
    t_act.setStyle(TableStyle(act_style))
    story.append(t_act)
    story.append(Spacer(1, 0.35*cm))

    # ── 2. ANÁLISIS DE BRECHA ────────────────────────────────
    story.append(sec_hdr("2. ANÁLISIS DE BRECHA — ¿Por qué el portafolio actual no alcanza el objetivo?"))
    story.append(Spacer(1, 0.15*cm))

    brecha_rows = [
        [Paragraph("<b>Indicador</b>", sN),
         Paragraph("<b>Valor</b>", sC),
         Paragraph("<b>Situación</b>", sC)],
        [Paragraph("Rendimiento nominal actual", sN),
         Paragraph(f"{rn_a*100:.2f}%", sC),
         Paragraph("POR DEBAJO del objetivo", sC)],
        [Paragraph("Objetivo 3× inflación", sN),
         Paragraph(f"{OBJETIVO_REND*100:.0f}%", sC),
         Paragraph("REQUERIDO", sC)],
        [Paragraph("Brecha (déficit de rendimiento)", sN),
         Paragraph(f"{(OBJETIVO_REND - rn_a)*100:.2f}%", sC),
         Paragraph("A CUBRIR", sC)],
        [Paragraph("Ganancia actual proyectada", sN),
         Paragraph(f"AR$ {g_a:,.0f}", sC),
         Paragraph("—", sC)],
        [Paragraph("Ganancia requerida para 3× inf.", sN),
         Paragraph(f"AR$ {ti_a*OBJETIVO_REND:,.0f}", sC),
         Paragraph("OBJETIVO", sC)],
        [Paragraph("Ganancia adicional necesaria", sN),
         Paragraph(f"AR$ {ti_a*OBJETIVO_REND - g_a:,.0f}", sC),
         Paragraph("FALTANTE", sC)],
    ]
    cw_br = [6.5*cm, 3.5*cm, W-10*cm]
    t_br = Table(brecha_rows, colWidths=cw_br)
    t_br.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,0), AO), ("TEXTCOLOR",(0,0),(-1,0), BL),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8), ("GRID",(0,0),(-1,-1),0.3,rl("#BBBBBB")),
        ("TOPPADDING",(0,0),(-1,-1),4), ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),5), ("RIGHTPADDING",(0,0),(-1,-1),5),
        ("ALIGN",(1,0),(-1,-1),"CENTER"),
        ("BACKGROUND",(1,1),(2,1), AmC), ("FONTNAME",(0,1),(0,1),"Helvetica-Bold"),
        ("BACKGROUND",(1,2),(2,2), RC),  ("FONTNAME",(0,2),(0,2),"Helvetica-Bold"),
        ("TEXTCOLOR",(2,2),(-1,2), RO),
        ("BACKGROUND",(1,3),(2,3), RC),  ("FONTNAME",(1,3),(1,3),"Helvetica-Bold"),
        ("TEXTCOLOR",(2,3),(-1,3), RO),
        ("BACKGROUND",(1,6),(2,6), RC),  ("FONTNAME",(1,6),(1,6),"Helvetica-Bold"),
        ("TEXTCOLOR",(2,6),(-1,6), RO),
        ("ROWBACKGROUNDS",(0,4),(-1,5), [GR, rl("#EBF3FB")]),
    ]))
    story.append(t_br)
    story.append(Spacer(1, 0.15*cm))
    story.append(Paragraph(
        "La cartera actual está compuesta en su mayoría por Renta Fija (bonos USD y CER) "
        "que, si bien preservan el capital con cobertura cambiaria, no generan el rendimiento "
        "nominal requerido para superar 3× la inflación. "
        "Las acciones y CEDEARs actuales aportan mejor rendimiento pero siguen por debajo del objetivo.",
        sNO))
    story.append(Spacer(1, 0.35*cm))

    # ── 3. NUEVOS INSTRUMENTOS IOL ───────────────────────────
    story.append(sec_hdr("3. NUEVOS INSTRUMENTOS SELECCIONADOS — IOL / BCBA", GO))
    story.append(Spacer(1, 0.15*cm))
    story.append(Paragraph(
        f"Se propone reasignar el capital (AR$ {ti_n:,.2f}) en los siguientes instrumentos "
        f"de alta potencialidad disponibles en InvertirOnline, con rendimiento esperado ≥ {OBJETIVO_REND*100:.0f}% nominal anual.",
        sN))
    story.append(Spacer(1, 0.2*cm))

    h_nvo = [Paragraph(f"<b>{h}</b>", sN) for h in
             ["Símbolo", "Instrumento", "Categoría", "Capital AR$",
              "% Cartera", "Rend. Esp. %", "Proy. AR$", "Gan. AR$", "Real %"]]
    rows_nvo = [h_nvo]
    for p in proy_nuevos:
        pct_c = p["invertido"] / ti_n
        rows_nvo.append([
            Paragraph(f"<b>{p['simbolo']}</b>", sN),
            Paragraph(p["nombre"], sN),
            Paragraph(p["categoria"], sN),
            Paragraph(f"{p['invertido']:,.0f}", sR),
            Paragraph(f"{pct_c*100:.1f}%", sC),
            Paragraph(f"<b>{p['rend_nom']*100:.0f}%</b>", sC),
            Paragraph(f"{p['valor_proy']:,.0f}", sR),
            Paragraph(f"<b>{p['ganancia']:,.0f}</b>", sR),
            Paragraph(f"{p['rend_real']*100:.1f}%", sC),
        ])
    rows_nvo.append([
        Paragraph("<b>TOTAL</b>", sNB), Paragraph("", sN), Paragraph("", sN),
        Paragraph(f"<b>{ti_n:,.0f}</b>",  PS("rr",fn="Helvetica-Bold",sz=8,al=TA_RIGHT)),
        Paragraph("<b>100%</b>", sNB),
        Paragraph(f"<b>{rn_n*100:.1f}%</b>", sNB),
        Paragraph(f"<b>{tp_n:,.0f}</b>",  PS("rr",fn="Helvetica-Bold",sz=8,al=TA_RIGHT)),
        Paragraph(f"<b>{g_n:,.0f}</b>",   PS("rr",fn="Helvetica-Bold",sz=8,al=TA_RIGHT)),
        Paragraph(f"<b>{rr_n*100:.1f}%</b>", sNB),
    ])
    cw_nvo = [1.4*cm, 3.2*cm, 3.2*cm, 2.3*cm, 1.5*cm, 1.8*cm, 2.2*cm, 2.0*cm, 1.4*cm]
    t_nvo = Table(rows_nvo, colWidths=cw_nvo, repeatRows=1)
    nvo_st = [
        ("BACKGROUND",(0,0),(-1,0), GO), ("TEXTCOLOR",(0,0),(-1,0), BL),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),7.5), ("GRID",(0,0),(-1,-1),0.3,rl("#BBBBBB")),
        ("TOPPADDING",(0,0),(-1,-1),3), ("BOTTOMPADDING",(0,0),(-1,-1),3),
        ("LEFTPADDING",(0,0),(-1,-1),4), ("RIGHTPADDING",(0,0),(-1,-1),4),
        ("BACKGROUND",(0,-1),(-1,-1), GO), ("TEXTCOLOR",(0,-1),(-1,-1), BL),
        ("LINEABOVE",(0,-1),(-1,-1),1.0, GM),
    ]
    for i2, p in enumerate(proy_nuevos, 1):
        bg2 = TIPO_BG.get(p["tipo"], GRIS_CL)
        nvo_st.append(("BACKGROUND",(0,i2),(-1,i2), rl(bg2)))
        bg_rend = rl(VERDE_CL) if p["rend_nom"] >= OBJETIVO_REND else rl(AMBAR_CL)
        nvo_st.append(("BACKGROUND",(5,i2),(5,i2), bg_rend))
        nvo_st.append(("BACKGROUND",(8,i2),(8,i2), GC))
    t_nvo.setStyle(TableStyle(nvo_st))
    story.append(t_nvo)
    story.append(Spacer(1, 0.3*cm))

    # Racional por instrumento
    for p in proy_nuevos:
        rat = Table([
            [Paragraph(f"<b>{p['simbolo']}</b>  —  {p['nombre']}", sNB),
             Paragraph(f"Rend. esperado: <b>{p['rend_nom']*100:.0f}%</b>  |  Riesgo: <b>{p['riesgo']}</b>", sC)],
            [Paragraph(p["racional"], sNO),
             Paragraph(f"Precio ref: {p['precio_ref']}", sNO)],
        ], colWidths=[W*0.70, W*0.30])
        rat.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,0), rl(TIPO_BG.get(p["tipo"], GRIS_CL))),
            ("BACKGROUND",(0,1),(-1,1), GR),
            ("GRID",(0,0),(-1,-1),0.3,rl("#CCCCCC")),
            ("TOPPADDING",(0,0),(-1,-1),3), ("BOTTOMPADDING",(0,0),(-1,-1),3),
            ("LEFTPADDING",(0,0),(-1,-1),5), ("RIGHTPADDING",(0,0),(-1,-1),5),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ]))
        story.append(rat)
        story.append(Spacer(1, 0.1*cm))

    story.append(PageBreak())

    # ── 4. COMPARATIVO ──────────────────────────────────────
    story.append(sec_hdr("4. COMPARATIVO FINAL — Portafolio Actual vs. Portafolio Propuesto"))
    story.append(Spacer(1, 0.2*cm))

    comp_data = [
        [Paragraph("<b>Métrica</b>", sN),
         Paragraph("<b>Portafolio ACTUAL</b>", sC),
         Paragraph("<b>Portafolio PROPUESTO</b>", sC),
         Paragraph("<b>Diferencia</b>", sC)],
        ["Capital invertido",          f"AR$ {ti_a:,.2f}",       f"AR$ {ti_n:,.2f}",       "Mismo capital"],
        ["Valor proyectado (1 año)",   f"AR$ {tp_a:,.2f}",       f"AR$ {tp_n:,.2f}",       f"AR$ {tp_n-tp_a:,.2f}"],
        ["Ganancia esperada",          f"AR$ {g_a:,.2f}",        f"AR$ {g_n:,.2f}",        f"+AR$ {g_n-g_a:,.2f}"],
        ["Rendimiento nominal anual",  f"{rn_a*100:.2f}%",       f"{rn_n*100:.2f}%",       f"+{(rn_n-rn_a)*100:.2f}pp"],
        [f"¿Supera 3× inflación ({OBJETIVO_REND*100:.0f}%)?",
                                       "✗  NO",                  "✓  SÍ",                  "OBJETIVO LOGRADO"],
        ["Rendimiento real anual",     f"{rr_a*100:.2f}%",       f"{rr_n*100:.2f}%",       f"+{(rr_n-rr_a)*100:.2f}pp"],
        ["N° instrumentos",            "13 símbolos",            f"{len(proy_nuevos)} instrumentos", "—"],
        ["Perfil de riesgo",           "Conservador-Moderado",   "Agresivo",               "Mayor riesgo"],
    ]
    cw_comp = [5.5*cm, 3.8*cm, 3.8*cm, W-13.1*cm]
    t_comp = Table(comp_data, colWidths=cw_comp)
    comp_st = [
        ("BACKGROUND",(0,0),(-1,0), AO), ("TEXTCOLOR",(0,0),(-1,0), BL),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),8),
        ("GRID",(0,0),(-1,-1),0.3,rl("#BBBBBB")),
        ("TOPPADDING",(0,0),(-1,-1),4), ("BOTTOMPADDING",(0,0),(-1,-1),4),
        ("LEFTPADDING",(0,0),(-1,-1),5), ("RIGHTPADDING",(0,0),(-1,-1),5),
        ("ALIGN",(1,0),(-1,-1),"CENTER"),
        # Fila 5: objetivo logrado
        ("BACKGROUND",(1,5),(1,5), RC),
        ("BACKGROUND",(2,5),(3,5), GC),
        ("FONTNAME",(1,5),(3,5),"Helvetica-Bold"),
        ("TEXTCOLOR",(2,5),(3,5), GO),
        # Alternado
        ("ROWBACKGROUNDS",(0,1),(-1,4), [GR, rl("#EBF3FB")]),
        ("ROWBACKGROUNDS",(0,6),(-1,-1),[GR, rl("#EBF3FB")]),
    ]
    t_comp.setStyle(TableStyle(comp_st))
    story.append(t_comp)
    story.append(Spacer(1, 0.4*cm))

    # ── 5. GRÁFICO DE BARRAS ─────────────────────────────────
    story.append(sec_hdr("5. ANÁLISIS VISUAL"))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("<b>Rendimiento esperado por instrumento — Nuevo portafolio vs. Objetivo 150%</b>", sNB))
    story.append(Spacer(1, 0.15*cm))

    # Barras: nuevos instrumentos vs objetivo
    d_bar = Drawing(W, 9*cm)
    bc = VerticalBarChart()
    bc.x = 50; bc.y = 25
    bc.width  = float(W) - 70
    bc.height = 9*cm - 40

    syms_n = [p["simbolo"] for p in proy_nuevos]
    rend_n_vals = [p["rend_nom"] * 100 for p in proy_nuevos]
    rend_r_vals = [p["rend_real"] * 100 for p in proy_nuevos]
    obj_vals    = [OBJETIVO_REND * 100  for _ in proy_nuevos]

    bc.data = [rend_n_vals, rend_r_vals, obj_vals]
    bc.categoryAxis.categoryNames = syms_n
    bc.categoryAxis.labels.angle  = 0
    bc.categoryAxis.labels.fontSize = 8
    bc.valueAxis.valueMin  = 0
    bc.valueAxis.valueMax  = 220
    bc.valueAxis.valueStep = 25
    bc.valueAxis.labels.fontSize   = 7
    bc.valueAxis.labelTextFormat   = "%.0f%%"
    bc.bars[0].fillColor = AM
    bc.bars[1].fillColor = GM
    bc.bars[2].fillColor = C.HexColor(AMBAR)
    bc.groupSpacing = 8; bc.barSpacing = 1

    d_bar.add(bc)
    # Leyenda
    lx = bc.x + bc.width + 4
    for i2, (lbl, col) in enumerate(
            [("Nominal", AM), ("Real", GM), (f"Obj.{OBJETIVO_REND*100:.0f}%", C.HexColor(AMBAR))]):
        ly = float(9*cm) - 20 - i2 * 16
        d_bar.add(Rect(lx, ly, 10, 8, fillColor=col, strokeColor=None))
        d_bar.add(String(lx+13, ly, lbl, fontSize=7, fillColor=rl(GRIS_MED)))

    story.append(d_bar)
    story.append(Spacer(1, 0.3*cm))

    # Gráfico torta: distribución del nuevo portafolio
    story.append(Paragraph("<b>Distribución del capital en el portafolio propuesto</b>", sNB))
    story.append(Spacer(1, 0.15*cm))

    d_pie = Drawing(W, 9*cm)
    pie = Pie()
    pie.x = 30; pie.y = 10
    pie.width = pie.height = float(9*cm) - 20
    pie.data   = [p["invertido"] for p in proy_nuevos]
    pie.labels = [f"{p['simbolo']}\n{p['invertido']/ti_n*100:.0f}%" for p in proy_nuevos]
    pie_cols = [AM, GM, C.HexColor(AMBAR), C.HexColor(NARANJA),
                C.HexColor(VIOLETA), C.HexColor(AZUL_MED), C.HexColor(ROJO_OSC)]
    for i2, col in enumerate(pie_cols[:len(proy_nuevos)]):
        pie.slices[i2].fillColor   = col
        pie.slices[i2].strokeColor = BL
        pie.slices[i2].strokeWidth = 1
    pie.sideLabels = True
    pie.slices.label_simple_pointer = True
    pie.slices.fontName = "Helvetica"; pie.slices.fontSize = 7

    d_pie.add(pie)
    story.append(d_pie)
    story.append(Spacer(1, 0.4*cm))

    # ── 6. ADVERTENCIA DE RIESGO ─────────────────────────────
    story.append(HRFlowable(width=W, thickness=0.5, color=GRM))
    story.append(Spacer(1, 0.1*cm))
    story.append(Paragraph(
        "⚠  ADVERTENCIA DE RIESGO: Los instrumentos propuestos (acciones de alta beta, CEDEARs "
        "apalancados 3×) conllevan un nivel de riesgo ALTO a MUY ALTO. Los rendimientos del "
        f"{OBJETIVO_REND*100:.0f}% anual son proyecciones basadas en escenarios favorables y NO "
        "garantizados. En escenarios adversos, estos instrumentos pueden generar pérdidas "
        "significativas de capital. Se recomienda no superar el 20-30% del patrimonio total "
        "en instrumentos de riesgo muy alto (TQQQ, SOXL).",
        sNO))
    story.append(Spacer(1, 0.1*cm))
    story.append(Paragraph(
        f"Informe generado el {FECHA_ANALISIS.strftime('%d/%m/%Y')}  —  "
        "InverSys Analytics  —  No constituye asesoramiento financiero.",
        sNO))

    doc.build(story)
    print(f"PDF generado: {output_path}")


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():
    base    = os.path.dirname(os.path.abspath(__file__))
    xls_in  = os.path.join(base, "OperacionesFinalizadas.xls")
    xls_out = os.path.join(base, "ReinversionIOL_1Anio.xlsx")
    pdf_out = os.path.join(base, "ReinversionIOL_1Anio.pdf")

    ops_raw     = leer_operaciones(xls_in)
    proy_actual = proyectar_actual(ops_raw)
    proy_nuevos = proyectar_nuevos()

    ti_a, tp_a, g_a, rn_a, rr_a = totales(proy_actual)
    ti_n, tp_n, g_n, rn_n, rr_n = totales(proy_nuevos)

    print(f"\n{'═'*56}")
    print(f"  PORTAFOLIO ACTUAL")
    print(f"  Capital invertido      : AR$ {ti_a:>12,.2f}")
    print(f"  Valor proyectado       : AR$ {tp_a:>12,.2f}")
    print(f"  Ganancia nominal       : AR$ {g_a:>12,.2f}")
    print(f"  Rend. nominal          :      {rn_a*100:>7.2f} %")
    print(f"  Rend. real             :      {rr_a*100:>7.2f} %")
    print(f"  Objetivo 3× inflación  :      {OBJETIVO_REND*100:>7.2f} %")
    brecha_str = "SÍ" if rn_a >= OBJETIVO_REND else f"NO (brecha: {(OBJETIVO_REND-rn_a)*100:.2f}%)"
    print(f"  ¿Cumple objetivo?      :      {brecha_str}")
    print(f"{'─'*56}")
    print(f"  PORTAFOLIO PROPUESTO (nuevos instrumentos IOL)")
    print(f"  Capital invertido      : AR$ {ti_n:>12,.2f}")
    print(f"  Valor proyectado       : AR$ {tp_n:>12,.2f}")
    print(f"  Ganancia nominal       : AR$ {g_n:>12,.2f}")
    print(f"  Rend. nominal          :      {rn_n*100:>7.2f} %")
    print(f"  Rend. real             :      {rr_n*100:>7.2f} %")
    print(f"  ¿Cumple objetivo?      :      {'SÍ ✓' if rn_n >= OBJETIVO_REND else 'NO'}")
    print(f"  Ganancia adicional     : AR$ {g_n - g_a:>12,.2f}")
    print(f"{'═'*56}\n")

    generar_excel(ops_raw, proy_actual, proy_nuevos, xls_out)
    generar_pdf(ops_raw, proy_actual, proy_nuevos, pdf_out)
    print("\nArchivos generados:")
    print(f"  → {xls_out}")
    print(f"  → {pdf_out}")


if __name__ == "__main__":
    main()
