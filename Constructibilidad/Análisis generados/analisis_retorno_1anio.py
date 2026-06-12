#!/usr/bin/env python3
"""
Análisis de Retorno de Inversiones a 1 Año Vista
Fuente: OperacionesFinalizadas.xls  (período 01/05/2026 – 22/05/2026)
Proyección: 21/05/2026 → 21/05/2027
"""

import os
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter
from datetime import date, datetime

# ─────────────────────────────────────────────────────────────
# PARÁMETROS MACROECONÓMICOS
# ─────────────────────────────────────────────────────────────
FECHA_ANALISIS   = date(2026,  5, 21)
FECHA_PROYECCION = date(2027,  5, 21)
INFLACION_ANUAL  = 0.50   # 50 % inflación anual proyectada (CER)

# ─────────────────────────────────────────────────────────────
# RENDIMIENTOS ESPERADOS POR SÍMBOLO  (retorno total ARS, 1 año)
# Metodología:
#   Bonos USD  : carry en USD  + variación precio  + devaluación ARS ~35 %
#   Bonos CER  : inflación CER ~50 % + spread real
#   CEDEARs    : apreciación USD del subyacente × (1 + devaluación ARS)
#   Acciones   : consenso analistas panel Merval 2026-27
#   FCI Mixto  : benchmark histórico del fondo
# ─────────────────────────────────────────────────────────────
PERFIL = {
    "GD41":    {"rend": 0.40, "tipo": "Bono Soberano USD",     "riesgo": "Medio", "moneda_suby": "USD"},
    "AO27":    {"rend": 0.35, "tipo": "Bono Tesoro USD",       "riesgo": "Bajo",  "moneda_suby": "USD"},
    "TZX26":   {"rend": 0.50, "tipo": "Bono CER (ajustable)",  "riesgo": "Bajo",  "moneda_suby": "ARS"},
    "BPOC7":   {"rend": 0.35, "tipo": "Bono BOPREAL USD",      "riesgo": "Bajo",  "moneda_suby": "USD"},
    "BMA":     {"rend": 0.65, "tipo": "Acción Panel Líder",    "riesgo": "Alto",  "moneda_suby": "ARS"},
    "GOOGL":   {"rend": 0.48, "tipo": "CEDEAR Tecnología",     "riesgo": "Alto",  "moneda_suby": "USD"},
    "VIST":    {"rend": 0.70, "tipo": "CEDEAR Energía",        "riesgo": "Alto",  "moneda_suby": "USD"},
    "COPX":    {"rend": 0.50, "tipo": "CEDEAR ETF Materiales", "riesgo": "Alto",  "moneda_suby": "USD"},
    "PAMP":    {"rend": 0.70, "tipo": "Acción Panel Líder",    "riesgo": "Alto",  "moneda_suby": "ARS"},
    "TGSU2":   {"rend": 0.55, "tipo": "Acción Panel Líder",    "riesgo": "Alto",  "moneda_suby": "ARS"},
    "EDN":     {"rend": 0.45, "tipo": "Acción Panel Líder",    "riesgo": "Alto",  "moneda_suby": "ARS"},
    "TXAR":    {"rend": 0.42, "tipo": "Acción Panel Líder",    "riesgo": "Alto",  "moneda_suby": "ARS"},
    "IOLPORA": {"rend": 0.55, "tipo": "FCI Portafolio Mixto",  "riesgo": "Medio", "moneda_suby": "ARS"},
}

# ─────────────────────────────────────────────────────────────
# HELPERS DE ESTILO
# ─────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, size=10, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic)

def align(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def border_thin():
    side = Side(style="thin", color="AAAAAA")
    return Border(left=side, right=side, top=side, bottom=side)

def border_medium():
    side = Side(style="medium", color="555555")
    return Border(left=side, right=side, top=side, bottom=side)

COLOR_HEADER   = "1A3C5E"   # azul oscuro
COLOR_SUBHEAD  = "2E6DA4"   # azul medio
COLOR_RENTA    = "D6E8F7"   # azul claro  (bonos)
COLOR_ACCION   = "D5F0DC"   # verde claro (acciones)
COLOR_CEDEAR   = "FFF0CC"   # amarillo    (CEDEARs)
COLOR_FCI      = "EDD8F5"   # lila        (FCI)
COLOR_RIESGO = {
    "Bajo":  "C6EFCE",  # verde
    "Medio": "FFEB9C",  # amarillo
    "Alto":  "FFC7CE",  # rojo claro
}
COLOR_POSITIVO = "C6EFCE"
COLOR_NEGATIVO = "FFC7CE"

TIPO_COLOR = {
    "Bono Soberano USD":     COLOR_RENTA,
    "Bono Tesoro USD":       COLOR_RENTA,
    "Bono CER (ajustable)":  COLOR_RENTA,
    "Bono BOPREAL USD":      COLOR_RENTA,
    "Acción Panel Líder":    COLOR_ACCION,
    "CEDEAR Tecnología":     COLOR_CEDEAR,
    "CEDEAR Energía":        COLOR_CEDEAR,
    "CEDEAR ETF Materiales": COLOR_CEDEAR,
    "FCI Portafolio Mixto":  COLOR_FCI,
}

def col_width(ws, col_letter, width):
    ws.column_dimensions[col_letter].width = width

def set_row_height(ws, row, height):
    ws.row_dimensions[row].height = height

# ─────────────────────────────────────────────────────────────
# LECTURA DEL ARCHIVO HTML-XLS
# ─────────────────────────────────────────────────────────────
def leer_operaciones(path):
    with open(path, "r", encoding="latin-1") as f:
        soup = BeautifulSoup(f.read(), "html.parser")

    filas = soup.find("table").find_all("tr")
    # fila 0 = título,  fila 1 = cabeceras,  filas 2..N = datos
    ops = []
    for row in filas[2:]:
        celdas = [td.get_text(strip=True) for td in row.find_all("td")]
        if len(celdas) < 16:
            continue

        def parsear_num(s):
            return float(s.replace(".", "").replace(",", ".")) if s else 0.0

        def parsear_fecha(s):
            try:
                return datetime.strptime(s.split()[0], "%d/%m/%Y").date()
            except Exception:
                return None

        ops.append({
            "fecha_tx":    parsear_fecha(celdas[0]),
            "fecha_liq":   parsear_fecha(celdas[1]),
            "boleto":      celdas[2],
            "mercado":     celdas[3],
            "tipo_tx":     celdas[4],
            "cuenta":      celdas[5],
            "descripcion": celdas[6],
            "especie":     celdas[7],
            "simbolo":     celdas[8],
            "cantidad":    parsear_num(celdas[9]),
            "moneda":      celdas[10],
            "precio":      parsear_num(celdas[11]),
            "monto":       parsear_num(celdas[12]),
            "comision":    parsear_num(celdas[13]),
            "iva":         parsear_num(celdas[14]),
            "total":       parsear_num(celdas[15]),
        })
    return ops

# ─────────────────────────────────────────────────────────────
# PROYECCIÓN POR OPERACIÓN
# ─────────────────────────────────────────────────────────────
def proyectar(op):
    simbolo = op["simbolo"]
    perfil  = PERFIL.get(simbolo, {"rend": 0.30, "tipo": "Otro", "riesgo": "Medio", "moneda_suby": "ARS"})

    invertido    = op["total"]             # costo real (incluye comisiones)
    rend_nominal = perfil["rend"]
    valor_proy   = invertido * (1 + rend_nominal)
    ganancia_ars = valor_proy - invertido
    ganancia_pct = rend_nominal
    # Retorno real: deflactar por inflación proyectada
    rend_real    = (1 + ganancia_pct) / (1 + INFLACION_ANUAL) - 1

    return {
        **perfil,
        "simbolo":       simbolo,
        "descripcion":   op["descripcion"],
        "fecha_compra":  op["fecha_tx"],
        "cantidad":      op["cantidad"],
        "precio_compra": op["precio"],
        "invertido_ars": invertido,
        "valor_proy":    round(valor_proy,   2),
        "ganancia_ars":  round(ganancia_ars, 2),
        "rend_nominal":  rend_nominal,
        "rend_real":     rend_real,
    }

# ─────────────────────────────────────────────────────────────
# HOJA 1 – OPERACIONES FINALIZADAS
# ─────────────────────────────────────────────────────────────
COLS_OPS = [
    ("Fecha Transacción", 18),
    ("Boleto",            11),
    ("Descripción",       38),
    ("Símbolo",            9),
    ("Cant.",              8),
    ("Precio Pond.",      14),
    ("Monto AR$",         14),
    ("Comisión AR$",      14),
    ("IVA AR$",           10),
    ("Total AR$",         14),
]

def hoja_operaciones(ws, ops_raw):
    ws.title = "Operaciones"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    # Título principal
    ws.merge_cells("A1:J1")
    c = ws["A1"]
    c.value    = "OPERACIONES FINALIZADAS  —  01/05/2026 al 22/05/2026"
    c.font     = font(bold=True, size=13, color="FFFFFF")
    c.fill     = fill(COLOR_HEADER)
    c.alignment = align("center")
    set_row_height(ws, 1, 28)

    # Encabezados
    for ci, (titulo, ancho) in enumerate(COLS_OPS, start=1):
        c = ws.cell(row=2, column=ci, value=titulo)
        c.font      = font(bold=True, size=10, color="FFFFFF")
        c.fill      = fill(COLOR_SUBHEAD)
        c.alignment = align("center")
        c.border    = border_thin()
        col_width(ws, get_column_letter(ci), ancho)
    set_row_height(ws, 2, 22)

    # Filas de datos
    for ri, op in enumerate(ops_raw, start=3):
        simbolo = op["simbolo"]
        tipo    = PERFIL.get(simbolo, {}).get("tipo", "Otro")
        bg      = TIPO_COLOR.get(tipo, "FFFFFF")

        valores = [
            op["fecha_tx"].strftime("%d/%m/%Y") if op["fecha_tx"] else "",
            op["boleto"],
            op["descripcion"],
            simbolo,
            op["cantidad"],
            op["precio"],
            op["monto"],
            op["comision"],
            op["iva"],
            op["total"],
        ]
        for ci, val in enumerate(valores, start=1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.fill      = fill(bg)
            c.border    = border_thin()
            c.alignment = align("center") if ci != 3 else align("left", wrap=True)
            c.font      = font(size=9)
            if ci in (6, 7, 8, 9, 10):
                c.number_format = '#,##0.00'
        set_row_height(ws, ri, 18)

    # Totales
    tr = len(ops_raw) + 3
    ws.merge_cells(f"A{tr}:F{tr}")
    c = ws[f"A{tr}"]
    c.value     = "TOTAL INVERTIDO"
    c.font      = font(bold=True, size=10, color="FFFFFF")
    c.fill      = fill(COLOR_HEADER)
    c.alignment = align("center")

    total_general = sum(op["total"] for op in ops_raw)
    for ci in range(7, 11):
        c = ws.cell(row=tr, column=ci)
        c.fill   = fill(COLOR_HEADER)
        c.border = border_thin()

    c = ws.cell(row=tr, column=10, value=total_general)
    c.font          = font(bold=True, size=10, color="FFFFFF")
    c.fill          = fill(COLOR_HEADER)
    c.alignment     = align("center")
    c.number_format = '#,##0.00'
    c.border        = border_thin()
    set_row_height(ws, tr, 20)

    # Leyenda de colores
    lr = tr + 2
    ws.merge_cells(f"A{lr}:B{lr}")
    ws[f"A{lr}"].value     = "Leyenda de categorías"
    ws[f"A{lr}"].font      = font(bold=True, size=9)
    ws[f"A{lr}"].alignment = align("left")

    leyenda = [
        ("Renta Fija (Bonos)",  COLOR_RENTA),
        ("Acciones",            COLOR_ACCION),
        ("CEDEARs",             COLOR_CEDEAR),
        ("FCI",                 COLOR_FCI),
    ]
    for i, (texto, color) in enumerate(leyenda):
        r_ = lr + 1 + i
        ws.cell(row=r_, column=1, value="  ").fill = fill(color)
        c2 = ws.cell(row=r_, column=2, value=texto)
        c2.font = font(size=9)
        c2.alignment = align("left")

# ─────────────────────────────────────────────────────────────
# HOJA 2 – PROYECCIÓN A 1 AÑO
# ─────────────────────────────────────────────────────────────
COLS_PROY = [
    ("Símbolo",             9),
    ("Descripción",        36),
    ("Categoría",          22),
    ("Riesgo",             10),
    ("Fecha Compra",       14),
    ("Invertido AR$",      15),
    ("Rend. Esperado %",   16),
    ("Valor Proy. AR$",    15),
    ("Ganancia AR$",       14),
    ("Gan. Nominal %",     14),
    ("Gan. Real %",        13),
]

def hoja_proyeccion(ws, proyecciones):
    ws.title = "Proyección 1 Año"
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "A3"

    # Banda superior
    ws.merge_cells("A1:K1")
    c = ws["A1"]
    c.value     = (f"PROYECCIÓN DE RENDIMIENTO — "
                   f"{FECHA_ANALISIS.strftime('%d/%m/%Y')}  →  "
                   f"{FECHA_PROYECCION.strftime('%d/%m/%Y')}  "
                   f"(Inflación estimada: {INFLACION_ANUAL*100:.0f}% anual)")
    c.font      = font(bold=True, size=13, color="FFFFFF")
    c.fill      = fill(COLOR_HEADER)
    c.alignment = align("center")
    set_row_height(ws, 1, 28)

    # Encabezados
    for ci, (titulo, ancho) in enumerate(COLS_PROY, start=1):
        c = ws.cell(row=2, column=ci, value=titulo)
        c.font      = font(bold=True, size=10, color="FFFFFF")
        c.fill      = fill(COLOR_SUBHEAD)
        c.alignment = align("center", wrap=True)
        c.border    = border_thin()
        col_width(ws, get_column_letter(ci), ancho)
    set_row_height(ws, 2, 30)

    total_inv = total_proy = 0.0

    for ri, p in enumerate(proyecciones, start=3):
        tipo    = p["tipo"]
        riesgo  = p["riesgo"]
        bg      = TIPO_COLOR.get(tipo, "FFFFFF")
        bg_rsg  = COLOR_RIESGO.get(riesgo, "FFFFFF")
        bg_gan  = COLOR_POSITIVO if p["ganancia_ars"] >= 0 else COLOR_NEGATIVO
        bg_real = COLOR_POSITIVO if p["rend_real"] >= 0    else COLOR_NEGATIVO

        total_inv  += p["invertido_ars"]
        total_proy += p["valor_proy"]

        valores = [
            p["simbolo"],
            p["descripcion"],
            tipo,
            riesgo,
            p["fecha_compra"].strftime("%d/%m/%Y") if p["fecha_compra"] else "",
            p["invertido_ars"],
            p["rend_nominal"],
            p["valor_proy"],
            p["ganancia_ars"],
            p["rend_nominal"],
            p["rend_real"],
        ]
        bgs = [bg, bg, bg, bg_rsg, bg, bg, bg, bg, bg_gan, bg_gan, bg_real]
        fmts = [None, None, None, None, None,
                '#,##0.00', '0.00%', '#,##0.00', '#,##0.00', '0.00%', '0.00%']

        for ci, (val, bgc, fmt) in enumerate(zip(valores, bgs, fmts), start=1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.fill      = fill(bgc)
            c.border    = border_thin()
            c.font      = font(size=9)
            c.alignment = align("left", wrap=True) if ci == 2 else align("center")
            if fmt:
                c.number_format = fmt
        set_row_height(ws, ri, 18)

    # Fila de totales
    tr = len(proyecciones) + 3
    ganancia_total  = total_proy - total_inv
    rend_total_nom  = ganancia_total / total_inv if total_inv else 0
    rend_total_real = (1 + rend_total_nom) / (1 + INFLACION_ANUAL) - 1

    ws.merge_cells(f"A{tr}:E{tr}")
    c = ws[f"A{tr}"]
    c.value     = "TOTALES DEL PORTAFOLIO"
    c.font      = font(bold=True, size=11, color="FFFFFF")
    c.fill      = fill(COLOR_HEADER)
    c.alignment = align("center")
    c.border    = border_thin()

    totales_vals = [total_inv, None, total_proy, ganancia_total, rend_total_nom, rend_total_real]
    totales_fmts = ['#,##0.00', None, '#,##0.00', '#,##0.00', '0.00%', '0.00%']
    for i, (val, fmt) in enumerate(zip(totales_vals, totales_fmts), start=6):
        c = ws.cell(row=tr, column=i, value=val)
        c.font          = font(bold=True, size=10, color="FFFFFF")
        c.fill          = fill(COLOR_HEADER)
        c.alignment     = align("center")
        c.border        = border_thin()
        if fmt:
            c.number_format = fmt
    set_row_height(ws, tr, 22)

    return total_inv, total_proy, ganancia_total, rend_total_nom, rend_total_real

# ─────────────────────────────────────────────────────────────
# HOJA 3 – RESUMEN EJECUTIVO
# ─────────────────────────────────────────────────────────────
def hoja_resumen(ws, proyecciones, total_inv, total_proy, gan_total, rend_nom, rend_real):
    ws.title = "Resumen Ejecutivo"
    ws.sheet_view.showGridLines = False

    for col, w in zip("ABCDE", [22, 18, 18, 18, 18]):
        col_width(ws, col, w)

    # ── TÍTULO ──────────────────────────────────────────────────
    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value     = "RESUMEN EJECUTIVO — PROYECCIÓN PORTAFOLIO 1 AÑO"
    c.font      = font(bold=True, size=14, color="FFFFFF")
    c.fill      = fill(COLOR_HEADER)
    c.alignment = align("center")
    set_row_height(ws, 1, 30)

    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value     = (f"Análisis al {FECHA_ANALISIS.strftime('%d/%m/%Y')}  "
                   f"—  Proyección al {FECHA_PROYECCION.strftime('%d/%m/%Y')}")
    c.font      = font(italic=True, size=10, color="FFFFFF")
    c.fill      = fill(COLOR_SUBHEAD)
    c.alignment = align("center")
    set_row_height(ws, 2, 18)

    # ── MÉTRICAS CLAVE ──────────────────────────────────────────
    def kpi_row(row, label, val, fmt=None, bg_val="EBF3FB"):
        ws.merge_cells(f"A{row}:B{row}")
        c1 = ws[f"A{row}"]
        c1.value     = label
        c1.font      = font(bold=True, size=10)
        c1.fill      = fill("D9E8F5")
        c1.alignment = align("left")
        c1.border    = border_thin()
        ws.merge_cells(f"C{row}:E{row}")
        c2 = ws[f"C{row}"]
        c2.value     = val
        c2.font      = font(bold=True, size=11)
        c2.fill      = fill(bg_val)
        c2.alignment = align("center")
        c2.border    = border_thin()
        if fmt:
            c2.number_format = fmt
        set_row_height(ws, row, 22)

    r = 4
    ws.merge_cells(f"A{r}:E{r}")
    c = ws[f"A{r}"]
    c.value     = "■  MÉTRICAS GLOBALES DEL PORTAFOLIO"
    c.font      = font(bold=True, size=11, color="FFFFFF")
    c.fill      = fill(COLOR_SUBHEAD)
    c.alignment = align("left")
    set_row_height(ws, r, 20)

    kpi_row(5,  "Capital total invertido (AR$)",    total_inv,  '#,##0.00')
    kpi_row(6,  "Valor proyectado total (AR$)",      total_proy, '#,##0.00', "C6EFCE")
    kpi_row(7,  "Ganancia bruta esperada (AR$)",     gan_total,  '#,##0.00', "C6EFCE")
    kpi_row(8,  "Rendimiento nominal anual (%)",     rend_nom,   '0.00%',    "C6EFCE")
    kpi_row(9,  "Inflación proyectada anual (%)",    INFLACION_ANUAL, '0.00%', "FFEB9C")
    kpi_row(10, "Rendimiento REAL anual (%)",        rend_real,  '0.00%',
            "C6EFCE" if rend_real >= 0 else "FFC7CE")
    kpi_row(11, "Fecha inicio proyección",           FECHA_ANALISIS.strftime('%d/%m/%Y'))
    kpi_row(12, "Fecha fin proyección",              FECHA_PROYECCION.strftime('%d/%m/%Y'))

    # ── DETALLE POR CATEGORÍA ───────────────────────────────────
    from collections import defaultdict
    por_tipo = defaultdict(lambda: {"inv": 0.0, "proy": 0.0, "n": 0})
    por_riesgo = defaultdict(lambda: {"inv": 0.0, "proy": 0.0, "n": 0})
    for p in proyecciones:
        por_tipo[p["tipo"]]["inv"]  += p["invertido_ars"]
        por_tipo[p["tipo"]]["proy"] += p["valor_proy"]
        por_tipo[p["tipo"]]["n"]    += 1
        por_riesgo[p["riesgo"]]["inv"]  += p["invertido_ars"]
        por_riesgo[p["riesgo"]]["proy"] += p["valor_proy"]
        por_riesgo[p["riesgo"]]["n"]    += 1

    r = 14
    ws.merge_cells(f"A{r}:E{r}")
    c = ws[f"A{r}"]
    c.value     = "■  DESGLOSE POR CATEGORÍA"
    c.font      = font(bold=True, size=11, color="FFFFFF")
    c.fill      = fill(COLOR_SUBHEAD)
    c.alignment = align("left")
    set_row_height(ws, r, 20)

    r += 1
    for ci, titulo in enumerate(["Categoría", "Ops.", "Invertido AR$", "Valor Proy. AR$", "Rend. %"], start=1):
        c = ws.cell(row=r, column=ci, value=titulo)
        c.font      = font(bold=True, size=9, color="FFFFFF")
        c.fill      = fill(COLOR_HEADER)
        c.alignment = align("center")
        c.border    = border_thin()
    set_row_height(ws, r, 18)

    for tipo_k, datos in sorted(por_tipo.items()):
        r += 1
        bg = TIPO_COLOR.get(tipo_k, "FFFFFF")
        rend_t = (datos["proy"] - datos["inv"]) / datos["inv"] if datos["inv"] else 0
        row_vals = [tipo_k, datos["n"], datos["inv"], datos["proy"], rend_t]
        fmts_    = [None, None, '#,##0.00', '#,##0.00', '0.00%']
        for ci, (v, f_) in enumerate(zip(row_vals, fmts_), start=1):
            c = ws.cell(row=r, column=ci, value=v)
            c.fill      = fill(bg)
            c.border    = border_thin()
            c.font      = font(size=9)
            c.alignment = align("center") if ci != 1 else align("left")
            if f_:
                c.number_format = f_
        set_row_height(ws, r, 18)

    # Sub-total categorías
    r += 1
    ws.merge_cells(f"A{r}:B{r}")
    c = ws[f"A{r}"]
    c.value = "TOTAL"
    c.font  = font(bold=True, size=9, color="FFFFFF")
    c.fill  = fill(COLOR_HEADER)
    c.alignment = align("center")
    c.border = border_thin()
    sub_vals = [total_inv, total_proy, rend_nom]
    sub_fmts = ['#,##0.00', '#,##0.00', '0.00%']
    for ci, (v, f_) in enumerate(zip(sub_vals, sub_fmts), start=3):
        c = ws.cell(row=r, column=ci, value=v)
        c.font = font(bold=True, size=9, color="FFFFFF")
        c.fill = fill(COLOR_HEADER)
        c.alignment = align("center")
        c.border = border_thin()
        if f_:
            c.number_format = f_
    set_row_height(ws, r, 18)

    # ── DESGLOSE POR RIESGO ─────────────────────────────────────
    r += 2
    ws.merge_cells(f"A{r}:E{r}")
    c = ws[f"A{r}"]
    c.value     = "■  DESGLOSE POR NIVEL DE RIESGO"
    c.font      = font(bold=True, size=11, color="FFFFFF")
    c.fill      = fill(COLOR_SUBHEAD)
    c.alignment = align("left")
    set_row_height(ws, r, 20)

    r += 1
    for ci, titulo in enumerate(["Riesgo", "Ops.", "Invertido AR$", "Valor Proy. AR$", "Rend. %"], start=1):
        c = ws.cell(row=r, column=ci, value=titulo)
        c.font      = font(bold=True, size=9, color="FFFFFF")
        c.fill      = fill(COLOR_HEADER)
        c.alignment = align("center")
        c.border    = border_thin()
    set_row_height(ws, r, 18)

    for riesgo_k in ["Bajo", "Medio", "Alto"]:
        if riesgo_k not in por_riesgo:
            continue
        datos  = por_riesgo[riesgo_k]
        r += 1
        bg     = COLOR_RIESGO[riesgo_k]
        rend_r = (datos["proy"] - datos["inv"]) / datos["inv"] if datos["inv"] else 0
        row_vals = [riesgo_k, datos["n"], datos["inv"], datos["proy"], rend_r]
        fmts_    = [None, None, '#,##0.00', '#,##0.00', '0.00%']
        for ci, (v, f_) in enumerate(zip(row_vals, fmts_), start=1):
            c = ws.cell(row=r, column=ci, value=v)
            c.fill      = fill(bg)
            c.border    = border_thin()
            c.font      = font(bold=True, size=9)
            c.alignment = align("center") if ci != 1 else align("left")
            if f_:
                c.number_format = f_
        set_row_height(ws, r, 18)

    # ── RANKING: TOP RENDIMIENTO ─────────────────────────────────
    r += 2
    ws.merge_cells(f"A{r}:E{r}")
    c = ws[f"A{r}"]
    c.value     = "■  RANKING DE RENDIMIENTO ESPERADO (mayor a menor)"
    c.font      = font(bold=True, size=11, color="FFFFFF")
    c.fill      = fill(COLOR_SUBHEAD)
    c.alignment = align("left")
    set_row_height(ws, r, 20)

    r += 1
    for ci, titulo in enumerate(["#", "Símbolo", "Descripción corta", "Rend. Nominal %", "Rend. Real %"], start=1):
        c = ws.cell(row=r, column=ci, value=titulo)
        c.font      = font(bold=True, size=9, color="FFFFFF")
        c.fill      = fill(COLOR_HEADER)
        c.alignment = align("center")
        c.border    = border_thin()
    set_row_height(ws, r, 18)

    # Consolidar por símbolo
    from collections import defaultdict
    por_sim = defaultdict(lambda: {"inv": 0.0, "proy": 0.0, "desc": "", "tipo": ""})
    for p in proyecciones:
        por_sim[p["simbolo"]]["inv"]  += p["invertido_ars"]
        por_sim[p["simbolo"]]["proy"] += p["valor_proy"]
        por_sim[p["simbolo"]]["desc"]  = p["descripcion"][:30]
        por_sim[p["simbolo"]]["tipo"]  = p["tipo"]

    ranking = sorted(por_sim.items(), key=lambda x: (x[1]["proy"]-x[1]["inv"])/x[1]["inv"], reverse=True)

    for rank_i, (sim, datos) in enumerate(ranking, start=1):
        r += 1
        rend_n = (datos["proy"] - datos["inv"]) / datos["inv"]
        rend_rl = (1 + rend_n) / (1 + INFLACION_ANUAL) - 1
        bg = TIPO_COLOR.get(datos["tipo"], "FFFFFF")
        bg_r = COLOR_POSITIVO if rend_rl >= 0 else COLOR_NEGATIVO
        row_vals = [rank_i, sim, datos["desc"], rend_n, rend_rl]
        bgs_     = ["D9E8F5", bg, bg, COLOR_POSITIVO, bg_r]
        fmts_    = [None, None, None, '0.00%', '0.00%']
        for ci, (v, bgc, f_) in enumerate(zip(row_vals, bgs_, fmts_), start=1):
            c = ws.cell(row=r, column=ci, value=v)
            c.fill      = fill(bgc)
            c.border    = border_thin()
            c.font      = font(bold=(ci in (1, 4, 5)), size=9)
            c.alignment = align("center") if ci != 3 else align("left")
            if f_:
                c.number_format = f_
        set_row_height(ws, r, 17)

    # Nota al pie
    r += 2
    ws.merge_cells(f"A{r}:E{r}")
    c = ws[f"A{r}"]
    c.value = ("⚠  NOTA: Las proyecciones son estimaciones basadas en análisis fundamental y técnico. "
               "No constituyen asesoramiento financiero ni garantía de retorno.")
    c.font      = font(italic=True, size=8, color="7F7F7F")
    c.alignment = align("left", wrap=True)
    set_row_height(ws, r, 30)

# ─────────────────────────────────────────────────────────────
# HOJA 4 – GRÁFICOS
# ─────────────────────────────────────────────────────────────
def hoja_graficos(ws, wb, proyecciones):
    ws.title = "Gráficos"
    ws.sheet_view.showGridLines = False

    ws.merge_cells("A1:N1")
    c = ws["A1"]
    c.value     = "ANÁLISIS VISUAL — PORTAFOLIO DE INVERSIONES"
    c.font      = font(bold=True, size=14, color="FFFFFF")
    c.fill      = fill(COLOR_HEADER)
    c.alignment = align("center")
    set_row_height(ws, 1, 30)

    # ── Tabla auxiliar para gráficos (col P+) ─────────────────
    from collections import defaultdict
    por_sim = defaultdict(lambda: {"inv": 0.0, "proy": 0.0, "desc": "", "tipo": ""})
    for p in proyecciones:
        k = p["simbolo"]
        por_sim[k]["inv"]  += p["invertido_ars"]
        por_sim[k]["proy"] += p["valor_proy"]
        por_sim[k]["desc"]  = p["simbolo"]
        por_sim[k]["tipo"]  = p["tipo"]

    sim_list = list(por_sim.items())

    # Encabezados auxiliares (col Q, R, S, T)
    WC = 17  # columna Q (data helper)
    ws.cell(row=3, column=WC,   value="Símbolo")
    ws.cell(row=3, column=WC+1, value="Invertido AR$")
    ws.cell(row=3, column=WC+2, value="Valor Proy AR$")
    ws.cell(row=3, column=WC+3, value="Rend. %")

    # Tabla por categoría para pie chart
    por_tipo = defaultdict(float)
    for s, d in sim_list:
        por_tipo[d["tipo"]] += d["inv"]
    tipos_list = list(por_tipo.items())

    WC2 = WC + 5  # col V+
    ws.cell(row=3, column=WC2,   value="Categoría")
    ws.cell(row=3, column=WC2+1, value="Invertido AR$")

    for i, (tipo_k, inv_v) in enumerate(tipos_list, start=4):
        ws.cell(row=i, column=WC2,   value=tipo_k)
        ws.cell(row=i, column=WC2+1, value=inv_v)

    for i, (sim_k, datos) in enumerate(sim_list, start=4):
        rend = (datos["proy"] - datos["inv"]) / datos["inv"]
        ws.cell(row=i, column=WC,   value=sim_k)
        ws.cell(row=i, column=WC+1, value=datos["inv"])
        ws.cell(row=i, column=WC+2, value=datos["proy"])
        ws.cell(row=i, column=WC+3, value=rend)

    n_sim  = len(sim_list)
    n_tipo = len(tipos_list)

    # ── GRÁFICO 1: Barras – Invertido vs. Proyectado ───────────
    bar = BarChart()
    bar.type          = "col"
    bar.grouping      = "clustered"
    bar.title         = "Capital Invertido vs. Valor Proyectado (AR$)"
    bar.y_axis.title  = "AR$"
    bar.x_axis.title  = "Símbolo"
    bar.width  = 22
    bar.height = 14

    cats = Reference(ws, min_col=WC, min_row=4, max_row=3+n_sim)
    inv_ref  = Reference(ws, min_col=WC+1, min_row=3, max_row=3+n_sim)
    proy_ref = Reference(ws, min_col=WC+2, min_row=3, max_row=3+n_sim)

    bar.add_data(inv_ref,  titles_from_data=True)
    bar.add_data(proy_ref, titles_from_data=True)
    bar.set_categories(cats)
    bar.series[0].graphicalProperties.solidFill = "2E6DA4"
    bar.series[1].graphicalProperties.solidFill = "70AD47"

    ws.add_chart(bar, "A3")

    # ── GRÁFICO 2: Barras – Rendimiento % por símbolo ──────────
    rend_bar = BarChart()
    rend_bar.type         = "col"
    rend_bar.grouping     = "clustered"
    rend_bar.title        = "Rendimiento Nominal Esperado por Instrumento (%)"
    rend_bar.y_axis.title = "Rendimiento %"
    rend_bar.x_axis.title = "Símbolo"
    rend_bar.width  = 22
    rend_bar.height = 14

    rend_ref = Reference(ws, min_col=WC+3, min_row=3, max_row=3+n_sim)
    rend_bar.add_data(rend_ref, titles_from_data=True)
    rend_bar.set_categories(cats)
    rend_bar.series[0].graphicalProperties.solidFill = "ED7D31"

    ws.add_chart(rend_bar, "A22")

    # ── GRÁFICO 3: Torta – Composición por categoría ───────────
    pie = PieChart()
    pie.title  = "Composición del Portafolio por Categoría"
    pie.width  = 18
    pie.height = 14

    pie_labels = Reference(ws, min_col=WC2,   min_row=4, max_row=3+n_tipo)
    pie_data   = Reference(ws, min_col=WC2+1, min_row=3, max_row=3+n_tipo)
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_labels)
    pie.dataLabels = None

    # Colores para cada slice
    pie_colors = ["2E6DA4", "70AD47", "FFC000", "9B59B6", "E74C3C"]
    for i, slice_color in enumerate(pie_colors[:n_tipo]):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = slice_color
        pie.series[0].dPt.append(pt)

    ws.add_chart(pie, "M3")

# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():
    base_dir  = os.path.dirname(os.path.abspath(__file__))
    root_dir  = os.path.dirname(base_dir)
    input_xls = os.path.join(root_dir, "datos_entradas", "OperacionesFinalizadas.xls")
    output    = os.path.join(base_dir, "RetornoInversiones_1Anio.xlsx")

    print(f"Leyendo operaciones desde: {input_xls}")
    ops_raw      = leer_operaciones(input_xls)
    proyecciones = [proyectar(op) for op in ops_raw]

    print(f"  → {len(ops_raw)} operaciones procesadas")

    wb = Workbook()
    ws_ops   = wb.active
    ws_proy  = wb.create_sheet()
    ws_res   = wb.create_sheet()
    ws_chart = wb.create_sheet()

    hoja_operaciones(ws_ops, ops_raw)

    total_inv, total_proy, gan_total, rend_nom, rend_real = \
        hoja_proyeccion(ws_proy, proyecciones)

    hoja_resumen(ws_res, proyecciones, total_inv, total_proy, gan_total, rend_nom, rend_real)
    hoja_graficos(ws_chart, wb, proyecciones)

    wb.save(output)
    print(f"\nInforme generado: {output}")
    print(f"\n{'─'*50}")
    print(f"  Capital total invertido : AR$ {total_inv:>14,.2f}")
    print(f"  Valor proyectado 1 año  : AR$ {total_proy:>14,.2f}")
    print(f"  Ganancia bruta esperada : AR$ {gan_total:>14,.2f}")
    print(f"  Rendimiento nominal     :      {rend_nom*100:>7.2f} %")
    print(f"  Inflación proyectada    :      {INFLACION_ANUAL*100:>7.2f} %")
    print(f"  Rendimiento REAL        :      {rend_real*100:>7.2f} %")
    print(f"{'─'*50}\n")

if __name__ == "__main__":
    main()
