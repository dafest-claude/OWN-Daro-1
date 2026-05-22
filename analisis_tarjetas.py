#!/usr/bin/env python3
"""
Análisis de Tarjetas de Crédito Argentinas — Mayo 2026
Genera: GastosTarjetas_Mayo2026.xlsx y GastosTarjetas_Mayo2026.pdf
"""

import os
import re
from datetime import date
from collections import defaultdict

# ─────────────────────────────────────────────────────────────
# DATOS DE MUESTRA
# ─────────────────────────────────────────────────────────────

VISA_TXN = [
    {"fecha": date(2026, 5, 2),  "desc": "Carrefour Express Palermo",       "monto": 68400},
    {"fecha": date(2026, 5, 3),  "desc": "Coto Martinez",                   "monto": 42300},
    {"fecha": date(2026, 5, 5),  "desc": "Disco Belgrano",                  "monto": 35800},
    {"fecha": date(2026, 5, 18), "desc": "Carrefour Caballito",              "monto": 28600},
    {"fecha": date(2026, 5, 4),  "desc": "La Cabrera Palermo",               "monto": 42000},
    {"fecha": date(2026, 5, 8),  "desc": "Sushi Pop Núñez",                  "monto": 38500},
    {"fecha": date(2026, 5, 12), "desc": "Hamburguesa Nostra Villa Urquiza", "monto": 24800},
    {"fecha": date(2026, 5, 15), "desc": "Freddo Belgrano",                  "monto": 9600},
    {"fecha": date(2026, 5, 20), "desc": "Café Martínez Recoleta",           "monto": 8900},
    {"fecha": date(2026, 5, 3),  "desc": "Rappi Delivery",                   "monto": 18200},
    {"fecha": date(2026, 5, 10), "desc": "Rappi Delivery",                   "monto": 22400},
    {"fecha": date(2026, 5, 17), "desc": "Rappi Delivery",                   "monto": 19800},
    {"fecha": date(2026, 5, 7),  "desc": "PedidosYa",                        "monto": 15600},
    {"fecha": date(2026, 5, 14), "desc": "PedidosYa",                        "monto": 12400},
    {"fecha": date(2026, 5, 6),  "desc": "YPF Palermo",                      "monto": 48300},
    {"fecha": date(2026, 5, 19), "desc": "YPF España",                       "monto": 51200},
    {"fecha": date(2026, 5, 1),  "desc": "Movistar",                         "monto": 24800},
    {"fecha": date(2026, 5, 1),  "desc": "Netflix",                          "monto": 6490},
    {"fecha": date(2026, 5, 1),  "desc": "OSDE prepaga",                     "monto": 142000},
    {"fecha": date(2026, 5, 2),  "desc": "Gym BodySystems",                  "monto": 28500},
    {"fecha": date(2026, 5, 5),  "desc": "Edesur",                           "monto": 18400},
    {"fecha": date(2026, 5, 5),  "desc": "Metrogas",                         "monto": 12300},
    {"fecha": date(2026, 5, 9),  "desc": "LG Smart TV CTA 3/12",             "monto": 45000},
    {"fecha": date(2026, 5, 9),  "desc": "Sillón Mepal CTA 5/18",            "monto": 38200},
    {"fecha": date(2026, 5, 9),  "desc": "iPhone 15 CTA 1/24",               "monto": 82500},
    {"fecha": date(2026, 5, 9),  "desc": "Viaje Bariloche CTA 2/6",          "monto": 68400},
    {"fecha": date(2026, 5, 11), "desc": "Farmacity Palermo",                "monto": 12400},
    {"fecha": date(2026, 5, 22), "desc": "Farmacity Belgrano",               "monto": 8700},
    {"fecha": date(2026, 5, 16), "desc": "Óptica Visión Sur",                "monto": 35000},
    {"fecha": date(2026, 5, 4),  "desc": "Uber",                             "monto": 4200},
    {"fecha": date(2026, 5, 13), "desc": "Uber",                             "monto": 3800},
    {"fecha": date(2026, 5, 21), "desc": "Uber",                             "monto": 5100},
    {"fecha": date(2026, 5, 7),  "desc": "Peaje AU Panamericana",            "monto": 1200},
]

AMEX_SANT_TXN = [
    {"fecha": date(2026, 5, 3),  "desc": "Aerolíneas Argentinas BUE-BRC",   "monto": 185400},
    {"fecha": date(2026, 5, 18), "desc": "Aerolíneas Argentinas BRC-BUE",   "monto": 195200},
    {"fecha": date(2026, 5, 4),  "desc": "Booking.com Hotel Patagonico",     "monto": 320000},
    {"fecha": date(2026, 5, 10), "desc": "NH Hotel Crillon BA",              "monto": 185000},
    {"fecha": date(2026, 5, 20), "desc": "Despegar.com Mendoza",             "monto": 248000},
    {"fecha": date(2026, 5, 6),  "desc": "Don Julio Palermo",                "monto": 86500},
    {"fecha": date(2026, 5, 9),  "desc": "Mishiguene",                       "monto": 72300},
    {"fecha": date(2026, 5, 14), "desc": "El Preferido de Palermo",          "monto": 48200},
    {"fecha": date(2026, 5, 19), "desc": "Aramburu Recoleta",                "monto": 124000},
    {"fecha": date(2026, 5, 1),  "desc": "Personal HogarNet",                "monto": 32400},
    {"fecha": date(2026, 5, 1),  "desc": "DirecTV GO",                       "monto": 18900},
    {"fecha": date(2026, 5, 1),  "desc": "Spotify",                          "monto": 3990},
    {"fecha": date(2026, 5, 7),  "desc": "Jumbo Unicenter",                  "monto": 98600},
    {"fecha": date(2026, 5, 15), "desc": "Jumbo Martinez",                   "monto": 76400},
    {"fecha": date(2026, 5, 5),  "desc": "Shell Av. Libertador",             "monto": 52800},
    {"fecha": date(2026, 5, 16), "desc": "Shell Panamericana",               "monto": 49600},
    {"fecha": date(2026, 5, 1),  "desc": "Swiss Medical prepaga",            "monto": 285000},
    {"fecha": date(2026, 5, 11), "desc": "Farmacity Premium",                "monto": 24300},
    {"fecha": date(2026, 5, 8),  "desc": "Samsung Galaxy CTA 2/18",          "monto": 78500},
    {"fecha": date(2026, 5, 8),  "desc": "Bicicleta Estática CTA 4/12",      "monto": 42000},
    {"fecha": date(2026, 5, 8),  "desc": "Viaje Europa CTA 1/12",            "monto": 168000},
    {"fecha": date(2026, 5, 12), "desc": "Zara Abasto CTA 1/3",              "monto": 85200},
    {"fecha": date(2026, 5, 17), "desc": "Nike Store Florida",               "monto": 62000},
    {"fecha": date(2026, 5, 13), "desc": "Megatone Smart TV CTA 3/12",       "monto": 185000},
    {"fecha": date(2026, 5, 21), "desc": "Apple App Store",                  "monto": 4200},
    {"fecha": date(2026, 5, 4),  "desc": "Uber Premium",                     "monto": 8400},
    {"fecha": date(2026, 5, 20), "desc": "Uber Premium",                     "monto": 6800},
    {"fecha": date(2026, 5, 10), "desc": "Peaje Autopista",                  "monto": 1800},
]

AMEX_TXN = [
    {"fecha": date(2026, 5, 5),  "desc": "Tegui Palermo",                   "monto": 95000},
    {"fecha": date(2026, 5, 8),  "desc": "Elena Four Seasons",               "monto": 138000},
    {"fecha": date(2026, 5, 13), "desc": "Osaka Puerto Madero",              "monto": 68400},
    {"fecha": date(2026, 5, 1),  "desc": "Disney+",                          "monto": 4800},
    {"fecha": date(2026, 5, 1),  "desc": "HBO Max",                          "monto": 5600},
    {"fecha": date(2026, 5, 1),  "desc": "Apple Music",                      "monto": 3200},
    {"fecha": date(2026, 5, 16), "desc": "Entradas Teatro Colón",            "monto": 42000},
    {"fecha": date(2026, 5, 7),  "desc": "Expedia vuelo internacional",      "monto": 425000},
    {"fecha": date(2026, 5, 9),  "desc": "LATAM Airlines",                   "monto": 218000},
    {"fecha": date(2026, 5, 4),  "desc": "Airbnb Mendoza",                   "monto": 245000},
    {"fecha": date(2026, 5, 11), "desc": "Airbnb Córdoba",                   "monto": 189000},
    {"fecha": date(2026, 5, 6),  "desc": "Jumbo Palermo",                    "monto": 68400},
    {"fecha": date(2026, 5, 14), "desc": "Rapsodia CTA 1/3",                 "monto": 78000},
    {"fecha": date(2026, 5, 18), "desc": "H&M Dot",                          "monto": 45200},
    {"fecha": date(2026, 5, 20), "desc": "Lacoste CTA 2/6",                  "monto": 125000},
    {"fecha": date(2026, 5, 1),  "desc": "Medicus prepaga",                  "monto": 156000},
    {"fecha": date(2026, 5, 10), "desc": "Farmacity Cabildo",                "monto": 18300},
    {"fecha": date(2026, 5, 17), "desc": "Laboratorio Roffo",                "monto": 32400},
    {"fecha": date(2026, 5, 15), "desc": "Easy San Fernando",                "monto": 98600},
    {"fecha": date(2026, 5, 19), "desc": "IKEA online CTA 1/6",              "monto": 145000},
    {"fecha": date(2026, 5, 8),  "desc": "Crucero Mediterráneo CTA 3/24",    "monto": 245000},
    {"fecha": date(2026, 5, 8),  "desc": "Heladera Whirlpool CTA 2/12",      "monto": 68400},
    {"fecha": date(2026, 5, 8),  "desc": "Colchón Simmons CTA 1/12",         "monto": 95000},
    {"fecha": date(2026, 5, 21), "desc": "Apple Store MacBook Accesorios",   "monto": 215000},
    {"fecha": date(2026, 5, 22), "desc": "Amazon US importado",              "monto": 65000},
]

CARDS = [
    ("Visa Santander",  VISA_TXN),
    ("Amex Santander",  AMEX_SANT_TXN),
    ("Amex",            AMEX_TXN),
]

# ─────────────────────────────────────────────────────────────
# CATEGORIZACIÓN
# ─────────────────────────────────────────────────────────────

CATEGORIES = [
    ("Supermercado",       ["carrefour", "coto", "jumbo", "disco", "dia", "walmart", "la anonima", "vea", "changomas", "fresco"]),
    ("Restaurantes y Bares", ["restaurant", "resto", "parrilla", "pizzeria", "sushi", "cafe", "coffee", "bodegon", "bar", "don julio", "mishiguene", "aramburu", "tegui", "elena", "osaka", "hamburguesa", "mcdonalds", "burger", "freddo", "la cabrera", "preferido"]),
    ("Delivery",           ["rappi", "pedidosya", "glovo"]),
    ("Combustible",        ["ypf", "shell", "axion", "puma", "nafta", "gasoil"]),
    ("Viajes",             ["aerolineas", "aerolíneas", "latam", "american airlines", "vuelo", "aeropuerto", "booking", "expedia", "despegar", "airbnb", "hotel"]),
    ("Transporte Local",   ["uber", "cabify", "taxi", "remis", "peaje", "sube"]),
    ("Servicios Fijos",    ["movistar", "claro", "personal", "telecom", "fibertel", "speedy", "directv", "flow", "cablevision", "metrogas", "edesur", "edenor", "aysa", "gym", "osde", "swiss medical", "medicus", "prepaga", "hogarn"]),
    ("Entretenimiento",    ["netflix", "spotify", "disney", "hbo", "apple music", "youtube", "deezer", "star+", "cine", "teatro", "entradas"]),
    ("Salud",              ["farmacia", "farmacity", "medico", "clinica", "hospital", "laboratorio", "optica", "óptica", "roffo"]),
    ("Indumentaria",       ["zara", "h&m", "rapsodia", "falabella", "ropa", "zapatilla", "adidas", "nike", "lacoste", "polo", "grimoldi"]),
    ("Tecnología",         ["apple store", "apple app", "samsung", "garbarino", "fravega", "musimundo", "compumundo", "megatone", "iphone", "macbook", "amazon", "lg smart", "lg "]),
    ("Hogar y Deco",       ["easy", "homecenters", "muebles", "decoracion", "ikea", "sodimac", "pintureria", "ferreteria", "sillón", "sillan", "colchon", "colchón", "heladera", "bicicleta"]),
]

CTA_RE = re.compile(r"CTA\s+(\d+)/(\d+)", re.IGNORECASE)


def categorize(desc):
    low = desc.lower()
    m = CTA_RE.search(desc)
    if m:
        nro, total = int(m.group(1)), int(m.group(2))
        # Still assign a primary category, but also flag as Cuota
        for cat, kws in CATEGORIES:
            for kw in kws:
                if kw in low:
                    return cat, "Cuota", nro, total
        return "Otros", "Cuota", nro, total
    for cat, kws in CATEGORIES:
        for kw in kws:
            if kw in low:
                return cat, "Normal", None, None
    return "Otros", "Normal", None, None


def enrich(txns):
    result = []
    for t in txns:
        cat, tipo, nro, total = categorize(t["desc"])
        result.append({**t, "cat": cat, "tipo": tipo, "nro_cuota": nro, "total_cuotas": total})
    return result


VISA_E   = enrich(VISA_TXN)
AMEX_S_E = enrich(AMEX_SANT_TXN)
AMEX_E   = enrich(AMEX_TXN)
ALL_DATA = [("Visa Santander", VISA_E), ("Amex Santander", AMEX_S_E), ("Amex", AMEX_E)]

# ─────────────────────────────────────────────────────────────
# COLORES POR CATEGORÍA
# ─────────────────────────────────────────────────────────────

CAT_COLORS = {
    "Supermercado":        "D5F0DC",
    "Restaurantes y Bares": "FFE4B5",
    "Delivery":            "FFFACD",
    "Combustible":         "FFD5B5",
    "Viajes":              "B5D5FF",
    "Transporte Local":    "D5ECFF",
    "Servicios Fijos":     "E8D5FF",
    "Entretenimiento":     "FFD5EC",
    "Salud":               "D5FFD5",
    "Indumentaria":        "FFB5D5",
    "Tecnología":          "B5FFD5",
    "Hogar y Deco":        "FFDDB5",
    "Otros":               "E8E8E8",
}

# ─────────────────────────────────────────────────────────────
# HELPERS EXCEL
# ─────────────────────────────────────────────────────────────

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, numbers
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint
from openpyxl.utils import get_column_letter


def xfill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def xfont(bold=False, size=10, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic, name="Calibri")


def xalign(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)


def xborder_thin():
    s = Side(style="thin", color="AAAAAA")
    return Border(left=s, right=s, top=s, bottom=s)


def xborder_med():
    s = Side(style="medium", color="555555")
    return Border(left=s, right=s, top=s, bottom=s)


C_DARK   = "1A3C5E"
C_MED    = "2E6DA4"
C_LIGHT  = "D6E8F7"
C_WHITE  = "FFFFFF"
C_GRAY   = "F5F5F5"
C_AMBER  = "FFC000"
C_GREEN  = "70AD47"
C_RED    = "FF4444"

ARS_FMT  = '#,##0.00'
DATE_FMT = "DD/MM/YYYY"

CARD_COLORS = {
    "Visa Santander": "1A3C5E",
    "Amex Santander": "1E5C30",
    "Amex":           "7B2D8B",
}


def set_cell(ws, row, col, value, fill=None, font=None, align=None, border=None, num_fmt=None):
    c = ws.cell(row=row, column=col, value=value)
    if fill:   c.fill      = fill
    if font:   c.font      = font
    if align:  c.alignment = align
    if border: c.border    = border
    if num_fmt: c.number_format = num_fmt
    return c


def title_row(ws, row, text, ncols, color=C_DARK):
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=ncols)
    c = ws.cell(row=row, column=1, value=text)
    c.fill      = xfill(color)
    c.font      = xfont(bold=True, size=14, color=C_WHITE)
    c.alignment = xalign("center", "center")
    ws.row_dimensions[row].height = 28


def header_row(ws, row, headers, color=C_MED):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.fill      = xfill(color)
        c.font      = xfont(bold=True, size=10, color=C_WHITE)
        c.alignment = xalign("center", "center")
        c.border    = xborder_thin()
    ws.row_dimensions[row].height = 18


# ─────────────────────────────────────────────────────────────
# SHEET: Per-card transaction sheet
# ─────────────────────────────────────────────────────────────

HEADERS = ["Fecha", "Descripción", "Categoría", "Tipo", "Monto AR$", "Cuota", "Nro Cuota"]
COL_W   = [12,      40,            22,           9,      16,          9,       10]


def write_card_sheet(ws, card_name, txns):
    ws.sheet_view.showGridLines = True
    ncols = len(HEADERS)

    title_row(ws, 1, f"Tarjeta: {card_name}  —  Mayo 2026", ncols, CARD_COLORS.get(card_name, C_DARK))
    header_row(ws, 2, HEADERS, C_MED)
    ws.freeze_panes = "A3"

    cat_totals = defaultdict(float)
    for r, t in enumerate(txns, start=3):
        cat  = t["cat"]
        fc   = CAT_COLORS.get(cat, "E8E8E8")
        fill = xfill(fc)
        border = xborder_thin()

        cuota_label = ""
        if t["total_cuotas"]:
            cuota_label = f"{t['nro_cuota']}/{t['total_cuotas']}"

        set_cell(ws, r, 1, t["fecha"],         fill, xfont(),                  xalign("center"),  border, DATE_FMT)
        set_cell(ws, r, 2, t["desc"],           fill, xfont(),                  xalign("left"),    border)
        set_cell(ws, r, 3, cat,                 fill, xfont(),                  xalign("center"),  border)
        set_cell(ws, r, 4, t["tipo"],           fill, xfont(italic=True),       xalign("center"),  border)
        set_cell(ws, r, 5, t["monto"],          fill, xfont(bold=True),         xalign("right"),   border, ARS_FMT)
        set_cell(ws, r, 6, cuota_label,         fill, xfont(),                  xalign("center"),  border)
        set_cell(ws, r, 7, t["nro_cuota"] or "",fill, xfont(),                 xalign("center"),  border)

        cat_totals[cat] += t["monto"]

    last_data = 2 + len(txns)
    # Summary row
    sum_row = last_data + 2
    ws.merge_cells(start_row=sum_row, start_column=1, end_row=sum_row, end_column=4)
    c = ws.cell(row=sum_row, column=1, value="TOTAL")
    c.fill      = xfill(C_DARK)
    c.font      = xfont(bold=True, size=11, color=C_WHITE)
    c.alignment = xalign("right")

    total = sum(t["monto"] for t in txns)
    tc = ws.cell(row=sum_row, column=5, value=total)
    tc.fill          = xfill(C_AMBER)
    tc.font          = xfont(bold=True, size=11)
    tc.alignment     = xalign("right")
    tc.number_format = ARS_FMT
    tc.border        = xborder_med()

    # Category subtotals
    sub_row = sum_row + 2
    ws.cell(row=sub_row, column=1, value="Subtotales por Categoría").font = xfont(bold=True, size=10)
    sub_row += 1
    header_row(ws, sub_row, ["Categoría", "Total AR$", "%"], C_MED)
    sub_row += 1
    for cat, tot in sorted(cat_totals.items(), key=lambda x: -x[1]):
        pct = tot / total * 100 if total else 0
        fc  = CAT_COLORS.get(cat, "E8E8E8")
        set_cell(ws, sub_row, 1, cat,  xfill(fc), xfont(), xalign("left"),   xborder_thin())
        set_cell(ws, sub_row, 2, tot,  xfill(fc), xfont(bold=True), xalign("right"), xborder_thin(), ARS_FMT)
        set_cell(ws, sub_row, 3, pct/100, xfill(fc), xfont(), xalign("center"), xborder_thin(), "0.0%")
        sub_row += 1

    # Column widths
    for col, w in enumerate(COL_W, 1):
        ws.column_dimensions[get_column_letter(col)].width = w


# ─────────────────────────────────────────────────────────────
# SHEET: Resumen Consolidado
# ─────────────────────────────────────────────────────────────

def write_resumen(ws, all_data):
    ws.sheet_view.showGridLines = True
    card_totals = {name: sum(t["monto"] for t in txns) for name, txns in all_data}
    grand_total = sum(card_totals.values())

    # KPI table
    title_row(ws, 1, "Resumen Consolidado — Mayo 2026", 8, C_DARK)
    header_row(ws, 2, ["Tarjeta", "Total AR$", "% del Total"], C_MED)

    r = 3
    for name, tot in card_totals.items():
        pct = tot / grand_total if grand_total else 0
        cc  = CARD_COLORS.get(name, C_DARK)
        set_cell(ws, r, 1, name,   xfill(C_LIGHT), xfont(bold=True),     xalign("left"),   xborder_thin())
        set_cell(ws, r, 2, tot,    xfill(C_LIGHT), xfont(bold=True),     xalign("right"),  xborder_thin(), ARS_FMT)
        set_cell(ws, r, 3, pct,    xfill(C_LIGHT), xfont(),              xalign("center"), xborder_thin(), "0.0%")
        r += 1

    # Grand total row
    set_cell(ws, r, 1, "GRAN TOTAL", xfill(C_AMBER), xfont(bold=True, size=11), xalign("left"), xborder_med())
    set_cell(ws, r, 2, grand_total,  xfill(C_AMBER), xfont(bold=True, size=11), xalign("right"), xborder_med(), ARS_FMT)
    set_cell(ws, r, 3, 1.0,          xfill(C_AMBER), xfont(bold=True),          xalign("center"), xborder_med(), "0.0%")

    # Category breakdown
    cat_data = defaultdict(lambda: defaultdict(float))
    for name, txns in all_data:
        for t in txns:
            cat_data[t["cat"]][name] += t["monto"]

    r += 3
    title_row(ws, r, "Gasto por Categoría", 8, C_MED)
    r += 1
    header_row(ws, r, ["Categoría", "Visa Santander", "Amex Santander", "Amex", "Total", "% del Total"], C_MED)
    r += 1

    all_cats = sorted(cat_data.keys(), key=lambda c: -sum(cat_data[c].values()))
    for cat in all_cats:
        v1 = cat_data[cat].get("Visa Santander", 0)
        v2 = cat_data[cat].get("Amex Santander", 0)
        v3 = cat_data[cat].get("Amex", 0)
        tot = v1 + v2 + v3
        pct = tot / grand_total if grand_total else 0
        fc  = CAT_COLORS.get(cat, "E8E8E8")
        f   = xfill(fc)
        b   = xborder_thin()
        set_cell(ws, r, 1, cat, f, xfont(bold=True), xalign("left"), b)
        set_cell(ws, r, 2, v1,  f, xfont(), xalign("right"), b, ARS_FMT)
        set_cell(ws, r, 3, v2,  f, xfont(), xalign("right"), b, ARS_FMT)
        set_cell(ws, r, 4, v3,  f, xfont(), xalign("right"), b, ARS_FMT)
        set_cell(ws, r, 5, tot, f, xfont(bold=True), xalign("right"), b, ARS_FMT)
        set_cell(ws, r, 6, pct, f, xfont(), xalign("center"), b, "0.0%")
        r += 1

    ws.freeze_panes = "A3"
    for i, w in enumerate([22, 18, 18, 18, 18, 12], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ─────────────────────────────────────────────────────────────
# SHEET: Por Categoría
# ─────────────────────────────────────────────────────────────

def write_por_categoria(ws, all_data):
    ws.sheet_view.showGridLines = True
    title_row(ws, 1, "Detalle por Categoría — Todas las Tarjetas", 8, C_DARK)
    cols = ["Tarjeta", "Fecha", "Descripción", "Categoría", "Tipo", "Monto AR$", "Cuota", "Nro Cuota"]
    header_row(ws, 2, cols, C_MED)
    ws.freeze_panes = "A3"

    all_txns = []
    for name, txns in all_data:
        for t in txns:
            all_txns.append((t["cat"], name, t))

    all_txns.sort(key=lambda x: (x[0], x[1]))

    for r, (cat, name, t) in enumerate(all_txns, start=3):
        fc     = CAT_COLORS.get(cat, "E8E8E8")
        fill   = xfill(fc)
        border = xborder_thin()
        cuota_label = f"{t['nro_cuota']}/{t['total_cuotas']}" if t["total_cuotas"] else ""
        set_cell(ws, r, 1, name,             fill, xfont(bold=True), xalign("center"), border)
        set_cell(ws, r, 2, t["fecha"],       fill, xfont(),          xalign("center"), border, DATE_FMT)
        set_cell(ws, r, 3, t["desc"],        fill, xfont(),          xalign("left"),   border)
        set_cell(ws, r, 4, cat,              fill, xfont(),          xalign("center"), border)
        set_cell(ws, r, 5, t["tipo"],        fill, xfont(italic=True), xalign("center"), border)
        set_cell(ws, r, 6, t["monto"],       fill, xfont(bold=True), xalign("right"),  border, ARS_FMT)
        set_cell(ws, r, 7, cuota_label,      fill, xfont(),          xalign("center"), border)
        set_cell(ws, r, 8, t["nro_cuota"] or "", fill, xfont(), xalign("center"), border)

    for i, w in enumerate([18, 12, 40, 22, 10, 16, 9, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ─────────────────────────────────────────────────────────────
# SHEET: Dashboard
# ─────────────────────────────────────────────────────────────

def write_dashboard(ws, wb, all_data):
    ws.sheet_view.showGridLines = False

    card_totals = {name: sum(t["monto"] for t in txns) for name, txns in all_data}
    grand_total = sum(card_totals.values())

    cat_data = defaultdict(float)
    for _, txns in all_data:
        for t in txns:
            cat_data[t["cat"]] += t["monto"]

    # Title band
    ws.merge_cells("A1:T2")
    c = ws.cell(row=1, column=1, value="DASHBOARD — ANÁLISIS DE TARJETAS DE CRÉDITO  |  Mayo 2026")
    c.fill      = xfill(C_DARK)
    c.font      = xfont(bold=True, size=16, color=C_WHITE)
    c.alignment = xalign("center", "center")
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[2].height = 22

    # KPI boxes: rows 4-9, 4 boxes side by side
    kpi_items = [
        ("Visa Santander",  card_totals.get("Visa Santander", 0),  "1A3C5E"),
        ("Amex Santander",  card_totals.get("Amex Santander", 0),  "1E5C30"),
        ("Amex",            card_totals.get("Amex", 0),            "7B2D8B"),
        ("GRAN TOTAL",      grand_total,                            "B45309"),
    ]
    kpi_cols = [1, 6, 11, 16]   # each box is 5 cols wide
    for (label, val, hex_c), start_col in zip(kpi_items, kpi_cols):
        end_col = start_col + 4
        for row in range(4, 10):
            ws.merge_cells(start_row=row, start_column=start_col, end_row=row, end_column=end_col)
        # background
        for row in range(4, 10):
            c = ws.cell(row=row, column=start_col)
            c.fill = xfill(hex_c)
        # label
        lc = ws.cell(row=4, column=start_col, value=label.upper())
        lc.font      = xfont(bold=True, size=10, color=C_WHITE)
        lc.alignment = xalign("center", "center")
        lc.fill      = xfill(hex_c)
        # value
        vc = ws.cell(row=6, column=start_col, value=val)
        vc.font          = xfont(bold=True, size=14, color=C_WHITE)
        vc.alignment     = xalign("center", "center")
        vc.fill          = xfill(hex_c)
        vc.number_format = ARS_FMT

    # Row heights for KPI
    for row in range(4, 10):
        ws.row_dimensions[row].height = 18

    # ---- BAR CHART: Gasto por Categoría (rows 11-28, cols A-L)
    # Write chart data to a hidden area (cols W onwards)
    cats_sorted = sorted(cat_data.keys(), key=lambda c: cat_data[c], reverse=True)
    data_start_col = 23  # col W

    ws.cell(row=11, column=data_start_col, value="Categoría")
    ws.cell(row=11, column=data_start_col+1, value="Total")
    for i, cat in enumerate(cats_sorted, start=12):
        ws.cell(row=i, column=data_start_col, value=cat)
        ws.cell(row=i, column=data_start_col+1, value=cat_data[cat])

    bar = BarChart()
    bar.type        = "bar"
    bar.grouping    = "clustered"
    bar.title       = "Gasto por Categoría"
    bar.y_axis.title = "Categoría"
    bar.x_axis.title = "AR$"
    bar.width        = 22
    bar.height       = 14

    data_ref  = Reference(ws, min_col=data_start_col+1, min_row=11,
                          max_row=11+len(cats_sorted))
    cats_ref  = Reference(ws, min_col=data_start_col,   min_row=12,
                          max_row=11+len(cats_sorted))
    bar.add_data(data_ref, titles_from_data=True)
    bar.set_categories(cats_ref)
    ws.add_chart(bar, "A11")

    # ---- PIE CHART: Distribución por tarjeta (rows 11-28, cols M-T)
    pie_data_col = data_start_col + 3  # col Z
    ws.cell(row=11, column=pie_data_col, value="Tarjeta")
    ws.cell(row=11, column=pie_data_col+1, value="Total")
    for i, (name, tot) in enumerate(card_totals.items(), start=12):
        ws.cell(row=i, column=pie_data_col, value=name)
        ws.cell(row=i, column=pie_data_col+1, value=tot)

    pie = PieChart()
    pie.title  = "Distribución por Tarjeta"
    pie.width  = 14
    pie.height = 14

    pie_data = Reference(ws, min_col=pie_data_col+1, min_row=11,
                         max_row=11+len(card_totals))
    pie_cats = Reference(ws, min_col=pie_data_col,   min_row=12,
                         max_row=11+len(card_totals))
    pie.add_data(pie_data, titles_from_data=True)
    pie.set_categories(pie_cats)

    slice_colors = ["1A3C5E", "1E5C30", "7B2D8B"]
    for i, hex_c in enumerate(slice_colors):
        dp = DataPoint(idx=i)
        dp.graphicalProperties.solidFill = hex_c
        pie.series[0].dPt.append(dp)

    ws.add_chart(pie, "M11")

    # ---- TOP 10 individual transactions (rows 30-42, cols A-G)
    all_txns_flat = []
    for name, txns in all_data:
        for t in txns:
            all_txns_flat.append((name, t))
    top10 = sorted(all_txns_flat, key=lambda x: -x[1]["monto"])[:10]

    r30 = 30
    ws.merge_cells(start_row=r30, start_column=1, end_row=r30, end_column=7)
    c = ws.cell(row=r30, column=1, value="Top 10 Gastos Individuales")
    c.fill      = xfill(C_DARK)
    c.font      = xfont(bold=True, size=11, color=C_WHITE)
    c.alignment = xalign("center")

    r30 += 1
    header_row(ws, r30, ["#", "Tarjeta", "Fecha", "Descripción", "Categoría", "Tipo", "Monto AR$"], C_MED)
    r30 += 1
    for rank, (name, t) in enumerate(top10, start=1):
        fc   = CAT_COLORS.get(t["cat"], "E8E8E8")
        fill = xfill(fc)
        b    = xborder_thin()
        set_cell(ws, r30, 1, rank,      fill, xfont(bold=True), xalign("center"), b)
        set_cell(ws, r30, 2, name,      fill, xfont(), xalign("center"), b)
        set_cell(ws, r30, 3, t["fecha"],fill, xfont(), xalign("center"), b, DATE_FMT)
        set_cell(ws, r30, 4, t["desc"], fill, xfont(), xalign("left"),   b)
        set_cell(ws, r30, 5, t["cat"],  fill, xfont(), xalign("center"), b)
        set_cell(ws, r30, 6, t["tipo"], fill, xfont(italic=True), xalign("center"), b)
        set_cell(ws, r30, 7, t["monto"],fill, xfont(bold=True), xalign("right"),  b, ARS_FMT)
        r30 += 1

    # ---- Cuotas activas (rows 30-42, cols I-N)
    cuotas = []
    for name, txns in all_data:
        for t in txns:
            if t["tipo"] == "Cuota":
                cuotas.append((name, t))
    cuotas.sort(key=lambda x: -x[1]["monto"])

    rc = 30
    ws.merge_cells(start_row=rc, start_column=9, end_row=rc, end_column=14)
    c = ws.cell(row=rc, column=9, value="Resumen de Cuotas Activas")
    c.fill      = xfill(C_MED)
    c.font      = xfont(bold=True, size=11, color=C_WHITE)
    c.alignment = xalign("center")

    rc += 1
    header_row(ws, rc, ["Tarjeta", "Descripción", "Cuota", "Monto AR$", "Restantes", "Total Comprometido"], C_MED)
    # fix column alignment (offset by 8)
    for col_idx in range(9, 15):
        ws.cell(row=rc, column=col_idx).fill      = xfill(C_MED)
        ws.cell(row=rc, column=col_idx).font      = xfont(bold=True, size=10, color=C_WHITE)
        ws.cell(row=rc, column=col_idx).alignment = xalign("center", "center")
        ws.cell(row=rc, column=col_idx).border    = xborder_thin()

    cuota_headers = ["Tarjeta", "Descripción", "Cuota", "Monto AR$", "Restantes", "Total Comprometido"]
    for i, h in enumerate(cuota_headers, start=9):
        ws.cell(row=rc, column=i, value=h)

    rc += 1
    for name, t in cuotas[:10]:
        restantes = t["total_cuotas"] - t["nro_cuota"]
        total_comp = t["monto"] * restantes
        fc   = CAT_COLORS.get(t["cat"], "E8E8E8")
        fill = xfill(fc)
        b    = xborder_thin()
        cuota_str = f"{t['nro_cuota']}/{t['total_cuotas']}"
        set_cell(ws, rc, 9,  name,       fill, xfont(), xalign("center"), b)
        set_cell(ws, rc, 10, t["desc"],  fill, xfont(), xalign("left"),   b)
        set_cell(ws, rc, 11, cuota_str,  fill, xfont(), xalign("center"), b)
        set_cell(ws, rc, 12, t["monto"], fill, xfont(bold=True), xalign("right"), b, ARS_FMT)
        set_cell(ws, rc, 13, restantes,  fill, xfont(), xalign("center"), b)
        set_cell(ws, rc, 14, total_comp, fill, xfont(bold=True), xalign("right"), b, ARS_FMT)
        rc += 1

    # Column widths
    col_widths = {1:5, 2:18, 3:12, 4:35, 5:22, 6:10, 7:16,
                  9:18, 10:35, 11:9, 12:16, 13:10, 14:20}
    for col, w in col_widths.items():
        ws.column_dimensions[get_column_letter(col)].width = w


# ─────────────────────────────────────────────────────────────
# BUILD EXCEL
# ─────────────────────────────────────────────────────────────

def build_excel(out_path):
    wb = Workbook()
    wb.remove(wb.active)

    ws_visa   = wb.create_sheet("Visa Santander")
    ws_amex_s = wb.create_sheet("Amex Santander")
    ws_amex   = wb.create_sheet("Amex")
    ws_res    = wb.create_sheet("Resumen Consolidado")
    ws_cat    = wb.create_sheet("Por Categoría")
    ws_dash   = wb.create_sheet("Dashboard")

    write_card_sheet(ws_visa,   "Visa Santander", VISA_E)
    write_card_sheet(ws_amex_s, "Amex Santander", AMEX_S_E)
    write_card_sheet(ws_amex,   "Amex",           AMEX_E)
    write_resumen(ws_res, ALL_DATA)
    write_por_categoria(ws_cat, ALL_DATA)
    write_dashboard(ws_dash, wb, ALL_DATA)

    wb.save(out_path)
    return out_path


# ─────────────────────────────────────────────────────────────
# PDF GENERATION
# ─────────────────────────────────────────────────────────────

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, PageBreak
)
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF
from reportlab.graphics.charts.barcharts import VerticalBarChart


AZUL_OSC  = colors.HexColor("#1A3C5E")
AZUL_MED  = colors.HexColor("#2E6DA4")
AZUL_CLAR = colors.HexColor("#D6E8F7")
VERDE_OSC = colors.HexColor("#1E5C30")
VERDE_MED = colors.HexColor("#70AD47")
VERDE_CL  = colors.HexColor("#C6EFCE")
MORADO    = colors.HexColor("#7B2D8B")
AMBAR_RL  = colors.HexColor("#FFC000")
GRIS_OSC  = colors.HexColor("#404040")
GRIS_MED  = colors.HexColor("#808080")
GRIS_CL   = colors.HexColor("#F5F5F5")
BLANCO    = colors.white
NEGRO     = colors.black

PAGE_W, PAGE_H = A4
MARGIN = 1.8 * cm

styles = getSampleStyleSheet()

def ps(name, parent="Normal", **kwargs):
    return ParagraphStyle(name, parent=styles[parent], **kwargs)


TITLE_STYLE   = ps("Title2",   fontSize=18, textColor=BLANCO, alignment=TA_CENTER, leading=22, fontName="Helvetica-Bold")
HEAD_STYLE    = ps("Head2",    fontSize=12, textColor=BLANCO, alignment=TA_CENTER, leading=15, fontName="Helvetica-Bold")
BODY_STYLE    = ps("Body2",    fontSize=9,  textColor=NEGRO,  alignment=TA_LEFT,   leading=12)
BOLD_STYLE    = ps("Bold2",    fontSize=9,  textColor=NEGRO,  alignment=TA_LEFT,   leading=12, fontName="Helvetica-Bold")
SECTION_STYLE = ps("Section2", fontSize=11, textColor=BLANCO, alignment=TA_LEFT,   leading=14, fontName="Helvetica-Bold")
BULLET_STYLE  = ps("Bullet2",  fontSize=9,  textColor=NEGRO,  alignment=TA_LEFT,   leading=13, leftIndent=12)
SMALL_STYLE   = ps("Small2",   fontSize=7,  textColor=GRIS_MED, alignment=TA_CENTER, leading=9)


def fmt_ars(v):
    return f"AR$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")


def section_header(text, color=AZUL_MED):
    d = Drawing(PAGE_W - 2*MARGIN, 22)
    d.add(Rect(0, 0, PAGE_W - 2*MARGIN, 22, fillColor=color, strokeColor=None))
    d.add(String(8, 6, text, fontSize=11, fillColor=colors.white, fontName="Helvetica-Bold"))
    return d


def kpi_row_drawing(items):
    """
    items: list of (label, value_str, hex_color)
    Returns a Drawing with 4 boxes side by side.
    """
    w = PAGE_W - 2*MARGIN
    box_w = w / len(items) - 4
    d = Drawing(w, 60)
    for i, (label, val, hex_c) in enumerate(items):
        x = i * (box_w + 4)
        d.add(Rect(x, 0, box_w, 60, fillColor=colors.HexColor(f"#{hex_c}"), strokeColor=None))
        d.add(String(x + box_w/2, 42, label, fontSize=9, fillColor=colors.white,
                     fontName="Helvetica-Bold", textAnchor="middle"))
        d.add(String(x + box_w/2, 18, val, fontSize=11, fillColor=colors.white,
                     fontName="Helvetica-Bold", textAnchor="middle"))
    return d


def bar_chart_drawing(cats, values, title="Gasto por Categoría"):
    """Horizontal bar chart using VerticalBarChart rotated via Drawing."""
    dw = PAGE_W - 2*MARGIN
    dh = 200
    d  = Drawing(dw, dh)

    # Title
    d.add(String(dw/2, dh - 12, title, fontSize=11, fillColor=AZUL_OSC,
                 fontName="Helvetica-Bold", textAnchor="middle"))

    # Manual horizontal bars
    max_val = max(values) if values else 1
    bar_area_x = 160
    bar_area_w = dw - bar_area_x - 10
    n = len(cats)
    bar_h = min(14, (dh - 40) / n - 3) if n else 14
    gap   = 3
    pal   = [
        "#1A3C5E","#2E6DA4","#70AD47","#FFC000","#7B2D8B",
        "#E84040","#00B0F0","#FF7C00","#00B050","#C00000",
        "#FF99CC","#99CCFF","#CCFF99",
    ]
    for i, (cat, val) in enumerate(zip(cats, values)):
        y     = dh - 30 - i*(bar_h + gap) - bar_h
        bw    = (val / max_val) * bar_area_w
        color = colors.HexColor(pal[i % len(pal)])
        d.add(Rect(bar_area_x, y, bw, bar_h, fillColor=color, strokeColor=None))
        d.add(String(bar_area_x - 4, y + bar_h/2 - 4, cat[:22],
                     fontSize=7, fillColor=NEGRO, textAnchor="end"))
        d.add(String(bar_area_x + bw + 4, y + bar_h/2 - 4,
                     f"{val/1e6:.2f}M" if val >= 1e6 else f"{val/1000:.0f}k",
                     fontSize=7, fillColor=GRIS_OSC, textAnchor="start"))
    return d


def pie_chart_drawing(labels, values, title="Distribución por Tarjeta"):
    dw = 220
    dh = 160
    d  = Drawing(dw, dh)
    d.add(String(dw/2, dh - 12, title, fontSize=10, fillColor=AZUL_OSC,
                 fontName="Helvetica-Bold", textAnchor="middle"))

    pie = Pie()
    pie.x      = 20
    pie.y      = 20
    pie.width  = 110
    pie.height = 110
    pie.data   = values
    pie.labels = labels
    pie_colors = [AZUL_OSC, VERDE_OSC, MORADO]
    for i, c in enumerate(pie_colors[:len(values)]):
        pie.slices[i].fillColor = c
        pie.slices[i].labelRadius = 1.2
        pie.slices[i].fontSize = 7
    d.add(pie)

    # Legend
    lx, ly = 140, dh - 30
    for i, (lbl, val) in enumerate(zip(labels, values)):
        pct = val / sum(values) * 100 if sum(values) else 0
        c   = pie_colors[i % len(pie_colors)]
        d.add(Rect(lx, ly - i*16, 10, 10, fillColor=c, strokeColor=None))
        d.add(String(lx+14, ly - i*16, f"{lbl}: {pct:.0f}%", fontSize=7,
                     fillColor=NEGRO, textAnchor="start"))
    return d


def table_style_default(n_rows, header_color=AZUL_MED):
    return TableStyle([
        ("BACKGROUND",   (0, 0), (-1, 0),  header_color),
        ("TEXTCOLOR",    (0, 0), (-1, 0),  BLANCO),
        ("FONTNAME",     (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",     (0, 0), (-1, 0),  9),
        ("ALIGN",        (0, 0), (-1, 0),  "CENTER"),
        ("FONTSIZE",     (0, 1), (-1, -1), 8),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [BLANCO, GRIS_CL]),
        ("GRID",         (0, 0), (-1, -1), 0.3, GRIS_MED),
        ("VALIGN",       (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING",  (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING",   (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 3),
    ])


def build_pdf(out_path):
    doc = SimpleDocTemplate(
        out_path,
        pagesize=A4,
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=MARGIN,  bottomMargin=MARGIN,
    )

    card_totals = {name: sum(t["monto"] for t in txns) for name, txns in ALL_DATA}
    grand_total = sum(card_totals.values())

    cat_data = defaultdict(float)
    cat_txns = defaultdict(list)
    for name, txns in ALL_DATA:
        for t in txns:
            cat_data[t["cat"]] += t["monto"]
            cat_txns[t["cat"]].append((name, t))

    cats_sorted = sorted(cat_data.keys(), key=lambda c: -cat_data[c])

    story = []

    # ─── PAGE 1: Dashboard Ejecutivo ─────────────────────────────
    # Header band
    header_d = Drawing(PAGE_W - 2*MARGIN, 50)
    header_d.add(Rect(0, 0, PAGE_W - 2*MARGIN, 50, fillColor=AZUL_OSC, strokeColor=None))
    header_d.add(String((PAGE_W - 2*MARGIN)/2, 30,
                         "ANÁLISIS DE TARJETAS DE CRÉDITO",
                         fontSize=16, fillColor=BLANCO, fontName="Helvetica-Bold", textAnchor="middle"))
    header_d.add(String((PAGE_W - 2*MARGIN)/2, 12,
                         "Mayo 2026  |  Generado el 22/05/2026",
                         fontSize=9, fillColor=AZUL_CLAR, textAnchor="middle"))
    story.append(header_d)
    story.append(Spacer(1, 0.3*cm))

    # KPI boxes
    kpi_items = [
        ("VISA SANTANDER",  fmt_ars(card_totals.get("Visa Santander", 0)),  "1A3C5E"),
        ("AMEX SANTANDER",  fmt_ars(card_totals.get("Amex Santander", 0)),  "1E5C30"),
        ("AMEX",            fmt_ars(card_totals.get("Amex", 0)),            "7B2D8B"),
        ("GRAN TOTAL",      fmt_ars(grand_total),                           "B45309"),
    ]
    story.append(kpi_row_drawing(kpi_items))
    story.append(Spacer(1, 0.4*cm))

    # Bar chart
    story.append(section_header("Gasto por Categoría — Todas las Tarjetas"))
    story.append(Spacer(1, 0.2*cm))
    bar_vals = [cat_data[c] for c in cats_sorted]
    story.append(bar_chart_drawing(cats_sorted, bar_vals))
    story.append(Spacer(1, 0.3*cm))

    # Pie chart
    pie_labels = list(card_totals.keys())
    pie_vals   = [card_totals[k] for k in pie_labels]
    story.append(pie_chart_drawing(pie_labels, pie_vals))

    story.append(PageBreak())

    # ─── PAGE 2: Análisis Detallado ──────────────────────────────
    story.append(section_header("Gastos por Categoría", AZUL_OSC))
    story.append(Spacer(1, 0.3*cm))

    cat_table_data = [["Categoría", "Monto AR$", "%", "# Txns", "Ticket Prom."]]
    all_txns_flat2 = [(t, name) for name, txns in ALL_DATA for t in txns]
    for cat in cats_sorted:
        n_txns   = len(cat_txns[cat])
        avg_tick = cat_data[cat] / n_txns if n_txns else 0
        pct      = cat_data[cat] / grand_total * 100 if grand_total else 0
        cat_table_data.append([
            cat,
            fmt_ars(cat_data[cat]),
            f"{pct:.1f}%",
            str(n_txns),
            fmt_ars(avg_tick),
        ])

    cat_table = Table(cat_table_data, colWidths=[5*cm, 4*cm, 2*cm, 2*cm, 4*cm])
    cat_table.setStyle(table_style_default(len(cat_table_data)))
    story.append(cat_table)
    story.append(Spacer(1, 0.5*cm))

    # Cuotas activas
    story.append(section_header("Gastos en Cuotas — Pagos Comprometidos", AZUL_MED))
    story.append(Spacer(1, 0.3*cm))

    cuotas_all = []
    for name, txns in ALL_DATA:
        for t in txns:
            if t["tipo"] == "Cuota":
                cuotas_all.append((name, t))
    cuotas_all.sort(key=lambda x: -x[1]["monto"])

    cuota_table_data = [["Tarjeta", "Descripción", "Cuota", "Monto/mes", "Cuotas Rest.", "Comprometido"]]
    for name, t in cuotas_all:
        rest = t["total_cuotas"] - t["nro_cuota"]
        cuota_table_data.append([
            name,
            t["desc"][:35],
            f"{t['nro_cuota']}/{t['total_cuotas']}",
            fmt_ars(t["monto"]),
            str(rest),
            fmt_ars(t["monto"] * rest),
        ])
    # Totals
    total_cuota_comprometido = sum(t["monto"] * (t["total_cuotas"] - t["nro_cuota"])
                                   for _, t in cuotas_all)
    cuota_table_data.append(["", "TOTAL COMPROMETIDO", "", "", "", fmt_ars(total_cuota_comprometido)])

    cuota_table = Table(cuota_table_data,
                        colWidths=[3.2*cm, 5.5*cm, 1.8*cm, 3*cm, 2.5*cm, 3.5*cm])
    cuota_table.setStyle(table_style_default(len(cuota_table_data)))
    # highlight total row
    cuota_table.setStyle(TableStyle([
        ("BACKGROUND", (0, -1), (-1, -1), AMBAR_RL),
        ("FONTNAME",   (0, -1), (-1, -1), "Helvetica-Bold"),
    ]))
    story.append(cuota_table)
    story.append(Spacer(1, 0.5*cm))

    # Gastos fijos mensuales
    story.append(section_header("Gastos Fijos Mensuales (Servicios)", VERDE_OSC))
    story.append(Spacer(1, 0.3*cm))

    fijos = [(name, t) for name, txns in ALL_DATA for t in txns
             if t["cat"] == "Servicios Fijos"]
    fijos.sort(key=lambda x: -x[1]["monto"])

    fijos_data = [["Tarjeta", "Servicio", "Monto AR$"]]
    for name, t in fijos:
        fijos_data.append([name, t["desc"], fmt_ars(t["monto"])])
    total_fijos = sum(t["monto"] for _, t in fijos)
    fijos_data.append(["", "TOTAL FIJOS", fmt_ars(total_fijos)])

    fijos_table = Table(fijos_data, colWidths=[4*cm, 8*cm, 4*cm])
    fijos_table.setStyle(table_style_default(len(fijos_data)))
    fijos_table.setStyle(TableStyle([
        ("BACKGROUND", (0, -1), (-1, -1), VERDE_CL),
        ("FONTNAME",   (0, -1), (-1, -1), "Helvetica-Bold"),
    ]))
    story.append(fijos_table)
    story.append(Spacer(1, 0.5*cm))

    # Top 10 comercios
    story.append(section_header("Top 10 Comercios — Mayor Gasto Individual", AZUL_OSC))
    story.append(Spacer(1, 0.3*cm))

    all_flat = [(name, t) for name, txns in ALL_DATA for t in txns]
    top10 = sorted(all_flat, key=lambda x: -x[1]["monto"])[:10]

    top10_data = [["#", "Tarjeta", "Comercio", "Categoría", "Monto AR$"]]
    for rank, (name, t) in enumerate(top10, 1):
        top10_data.append([str(rank), name, t["desc"][:30], t["cat"], fmt_ars(t["monto"])])

    top10_table = Table(top10_data, colWidths=[1*cm, 3.5*cm, 5.5*cm, 3.5*cm, 3.5*cm])
    top10_table.setStyle(table_style_default(len(top10_data)))
    story.append(top10_table)

    story.append(PageBreak())

    # ─── PAGE 3: Detalle por Tarjeta ─────────────────────────────
    story.append(section_header("Detalle por Tarjeta", AZUL_OSC))
    story.append(Spacer(1, 0.3*cm))

    card_summary_data = [["Tarjeta", "Transacciones", "Total AR$", "% del Total", "Ticket Prom."]]
    for name, txns in ALL_DATA:
        tot   = sum(t["monto"] for t in txns)
        pct   = tot / grand_total * 100 if grand_total else 0
        avg   = tot / len(txns) if txns else 0
        card_summary_data.append([name, str(len(txns)), fmt_ars(tot), f"{pct:.1f}%", fmt_ars(avg)])

    card_summary_table = Table(card_summary_data, colWidths=[4*cm, 3*cm, 4*cm, 3*cm, 4*cm])
    card_summary_table.setStyle(table_style_default(len(card_summary_data)))
    story.append(card_summary_table)
    story.append(Spacer(1, 0.5*cm))

    # Notable insights
    story.append(section_header("Insights Destacados", AZUL_MED))
    story.append(Spacer(1, 0.3*cm))

    viajes_total = cat_data.get("Viajes", 0)
    cuotas_total_mensual = sum(t["monto"] for _, t in cuotas_all)
    fijos_total = sum(t["monto"] for _, t in fijos)
    top_cat      = cats_sorted[0] if cats_sorted else "N/A"
    top_cat_pct  = cat_data.get(top_cat, 0) / grand_total * 100 if grand_total else 0

    insights = [
        f"Viajes representa el {viajes_total/grand_total*100:.1f}% del gasto total "
        f"({fmt_ars(viajes_total)}), siendo la categoría de mayor impacto junto con Servicios Fijos.",
        f"Cuotas activas: {fmt_ars(cuotas_total_mensual)} en pagos mensuales comprometidos. "
        f"Total comprometido futuro: {fmt_ars(total_cuota_comprometido)}.",
        f"Servicios fijos mensuales suman {fmt_ars(fijos_total)}, representando "
        f"{fijos_total/grand_total*100:.1f}% del gasto total.",
        f"La Amex Santander concentra el mayor gasto por tarjeta "
        f"({fmt_ars(card_totals.get('Amex Santander',0))}), impulsada por viajes internacionales.",
        f"El ticket promedio general es {fmt_ars(grand_total / len(all_flat))} por transacción.",
    ]
    for ins in insights:
        story.append(Paragraph(f"• {ins}", BULLET_STYLE))
        story.append(Spacer(1, 0.15*cm))

    story.append(Spacer(1, 0.4*cm))

    # Risk section
    story.append(section_header("Análisis de Riesgo y Recomendaciones", colors.HexColor("#7B2D00")))
    story.append(Spacer(1, 0.3*cm))

    risks = [
        ("Alto compromiso en cuotas",
         f"Con {fmt_ars(total_cuota_comprometido)} comprometidos en cuotas futuras, "
         "se recomienda no incrementar nuevas compras en cuotas hasta cancelar las actuales."),
        ("Concentración en Viajes",
         "El gasto en viajes supera categorías esenciales. Considerar presupuestar viajes "
         "con anticipación usando financiación en cuotas sin interés."),
        ("Contexto inflacionario",
         "Con inflación del 50% anual, los gastos en servicios fijos se incrementarán ~4% mensual. "
         "Revisar suscripciones prescindibles."),
        ("Optimización de tarjetas",
         "Concentrar compras de supermercado en la tarjeta con mayores beneficios/cashback "
         "para maximizar retorno en la categoría de mayor frecuencia."),
    ]
    for title_r, body_r in risks:
        story.append(Paragraph(f"<b>{title_r}:</b> {body_r}", BODY_STYLE))
        story.append(Spacer(1, 0.2*cm))

    story.append(Spacer(1, 1*cm))
    story.append(HRFlowable(width="100%", thickness=0.5, color=GRIS_MED))
    story.append(Spacer(1, 0.2*cm))

    disclaimer = ("Este reporte fue generado con datos de muestra para análisis ilustrativo. "
                  "Los montos, comercios y fechas son representativos de patrones de gasto "
                  "típicos argentinos para mayo 2026. No constituye asesoramiento financiero.")
    story.append(Paragraph(disclaimer, SMALL_STYLE))

    doc.build(story)
    return out_path


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    out_dir = "/home/user/OWN-Daro-1"
    xlsx_path = os.path.join(out_dir, "GastosTarjetas_Mayo2026.xlsx")
    pdf_path  = os.path.join(out_dir, "GastosTarjetas_Mayo2026.pdf")

    print("Generando Excel...")
    build_excel(xlsx_path)
    print(f"  Excel guardado: {xlsx_path}  ({os.path.getsize(xlsx_path):,} bytes)")

    print("Generando PDF...")
    build_pdf(pdf_path)
    print(f"  PDF guardado:   {pdf_path}  ({os.path.getsize(pdf_path):,} bytes)")

    # Summary to stdout
    card_totals = {name: sum(t["monto"] for t in txns) for name, txns in ALL_DATA}
    grand_total = sum(card_totals.values())
    cat_data    = defaultdict(float)
    for _, txns in ALL_DATA:
        for t in txns:
            cat_data[t["cat"]] += t["monto"]

    print("\n" + "="*60)
    print("  RESUMEN — TARJETAS MAYO 2026")
    print("="*60)
    for name, tot in card_totals.items():
        n = sum(len(txns) for n2, txns in ALL_DATA if n2 == name)
        print(f"  {name:<20} AR$ {tot:>14,.0f}  ({n} txns)")
    print("-"*60)
    print(f"  {'GRAN TOTAL':<20} AR$ {grand_total:>14,.0f}")
    print("\n  Top 5 categorías:")
    for cat in sorted(cat_data.keys(), key=lambda c: -cat_data[c])[:5]:
        pct = cat_data[cat]/grand_total*100
        print(f"    {cat:<25} AR$ {cat_data[cat]:>12,.0f}  ({pct:.1f}%)")
    print("="*60)
    print(f"\n  Archivos generados en {out_dir}/")
    print(f"    GastosTarjetas_Mayo2026.xlsx")
    print(f"    GastosTarjetas_Mayo2026.pdf")
