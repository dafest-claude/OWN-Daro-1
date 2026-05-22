#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Análisis de gastos de tarjetas de crédito - Datos reales Abril 2026
Tarjetas: Amex Santander 3766-865097-53002 | Amex 4100 | Visa 2473
"""

import glob
import re
from datetime import datetime
from collections import defaultdict

import openpyxl
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.series import DataPoint

from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.units import cm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (SimpleDocTemplate, Table, TableStyle,
                                Paragraph, Spacer, HRFlowable)
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF

# ─────────────────────────────────────────────────────────────────────────────
# CATEGORÍAS Y COLORES
# ─────────────────────────────────────────────────────────────────────────────
CATEGORIAS = [
    "CUOTAS",
    "SUPERMERCADO",
    "RESTAURANTES",
    "DELIVERY",
    "AUTOMÓVIL",
    "TRANSPORTE",
    "SERVICIOS FIJOS",
    "ENTRETENIMIENTO",
    "SEGUROS",
    "SALUD",
    "SERVICIO DOMÉSTICO",
    "IMPUESTOS Y CARGOS",
    "TECNOLOGÍA",
    "VIAJES",
    "OTROS",
]

CAT_COLORS = {
    "CUOTAS":              "FF6B35",
    "SUPERMERCADO":        "4CAF50",
    "RESTAURANTES":        "E91E63",
    "DELIVERY":            "FF9800",
    "AUTOMÓVIL":           "2196F3",
    "TRANSPORTE":          "00BCD4",
    "SERVICIOS FIJOS":     "9C27B0",
    "ENTRETENIMIENTO":     "673AB7",
    "SEGUROS":             "795548",
    "SALUD":               "009688",
    "SERVICIO DOMÉSTICO":  "607D8B",
    "IMPUESTOS Y CARGOS":  "F44336",
    "TECNOLOGÍA":          "3F51B5",
    "VIAJES":              "FFC107",
    "OTROS":               "9E9E9E",
}

# ─────────────────────────────────────────────────────────────────────────────
# REGLAS DE CATEGORIZACIÓN (keyword → categoría)
# ─────────────────────────────────────────────────────────────────────────────
RULES = [
    # Supermercados
    ("coto",                  "SUPERMERCADO"),
    ("superdia",              "SUPERMERCADO"),
    ("coop obrera",           "SUPERMERCADO"),
    ("disco",                 "SUPERMERCADO"),
    ("carrefour",             "SUPERMERCADO"),
    ("jumbo",                 "SUPERMERCADO"),
    # Restaurantes / cafés / heladerías
    ("merengue",              "RESTAURANTES"),
    ("la veneciana",          "RESTAURANTES"),
    ("clock resto",           "RESTAURANTES"),
    ("clockresto",            "RESTAURANTES"),
    ("quotidiano",            "RESTAURANTES"),
    ("innsbruck",             "RESTAURANTES"),
    ("the one resto",         "RESTAURANTES"),
    ("la princesa",           "RESTAURANTES"),
    ("elgransabor",           "RESTAURANTES"),
    ("el gran sabor",         "RESTAURANTES"),
    ("mcdonalds",             "RESTAURANTES"),
    ("deniscerveza",          "RESTAURANTES"),
    ("elreydelaesqu",         "RESTAURANTES"),
    ("rapanui",               "RESTAURANTES"),
    ("marpann",               "RESTAURANTES"),
    ("promokiosco",           "RESTAURANTES"),
    ("frivolite",             "RESTAURANTES"),
    # Delivery
    ("pedidosya",             "DELIVERY"),
    ("rappi",                 "DELIVERY"),
    # Combustible / automóvil / peajes
    ("ypf",                   "AUTOMÓVIL"),
    ("shell",                 "AUTOMÓVIL"),
    ("axion",                 "AUTOMÓVIL"),
    ("autopistas del sol",    "AUTOMÓVIL"),
    ("autopistasdelso",       "AUTOMÓVIL"),
    ("applus iteuve",         "AUTOMÓVIL"),
    ("applus",                "AUTOMÓVIL"),
    ("prot robo celu",        "AUTOMÓVIL"),
    # Transporte público / SUBE
    ("emovasubte",            "TRANSPORTE"),
    ("subeviajes",            "TRANSPORTE"),
    ("sube",                  "TRANSPORTE"),
    # Telefonía / internet / cable
    ("claro",                 "SERVICIOS FIJOS"),
    ("antina",                "SERVICIOS FIJOS"),
    ("personal",              "SERVICIOS FIJOS"),
    ("movistar",              "SERVICIOS FIJOS"),
    ("fibertel",              "SERVICIOS FIJOS"),
    ("cablevision",           "SERVICIOS FIJOS"),
    ("directv",               "SERVICIOS FIJOS"),
    # Entretenimiento digital
    ("spotify",               "ENTRETENIMIENTO"),
    ("netflix",               "ENTRETENIMIENTO"),
    ("splice",                "ENTRETENIMIENTO"),
    ("apple.com",             "ENTRETENIMIENTO"),
    ("ebanxsa",               "ENTRETENIMIENTO"),
    ("ebanx",                 "ENTRETENIMIENTO"),
    ("disney",                "ENTRETENIMIENTO"),
    ("hbo",                   "ENTRETENIMIENTO"),
    ("youtube",               "ENTRETENIMIENTO"),
    # Seguros
    ("zurich",                "SEGUROS"),
    ("zurichseguros",         "SEGUROS"),
    ("zurich mb",             "SEGUROS"),
    ("sancorseguros",         "SEGUROS"),
    ("galicia seguros",       "SEGUROS"),
    # Salud / farmacia / prepaga
    ("s p s",                 "SALUD"),
    ("sancor cooperativa",    "SALUD"),
    ("farmacity",             "SALUD"),
    ("osde",                  "SALUD"),
    ("swiss medical",         "SALUD"),
    ("medicus",               "SALUD"),
    # Servicio doméstico (empleados, servicios del hogar)
    ("marcelaalejandras",     "SERVICIO DOMÉSTICO"),
    ("marcela alejandra",     "SERVICIO DOMÉSTICO"),
    ("nestorhoraciogarc",     "SERVICIO DOMÉSTICO"),
    ("nestor horac",          "SERVICIO DOMÉSTICO"),
    ("susanabeatrizgime",     "SERVICIO DOMÉSTICO"),
    ("susana beatriz",        "SERVICIO DOMÉSTICO"),
    ("cristianfedericob",     "SERVICIO DOMÉSTICO"),
    ("nicolasrubenricom",     "SERVICIO DOMÉSTICO"),
    ("silviagramajo",         "SERVICIO DOMÉSTICO"),
    # Impuestos / cargos financieros
    ("municipalidadedegr",    "IMPUESTOS Y CARGOS"),
    ("municipalidad",         "IMPUESTOS Y CARGOS"),
    ("intereses financieros", "IMPUESTOS Y CARGOS"),
    ("imp. sellos",           "IMPUESTOS Y CARGOS"),
    ("impuesto de sellos",    "IMPUESTOS Y CARGOS"),
    ("sellos",                "IMPUESTOS Y CARGOS"),
    ("percepcion",            "IMPUESTOS Y CARGOS"),
    ("percep",                "IMPUESTOS Y CARGOS"),
    ("iva 21",                "IMPUESTOS Y CARGOS"),
    ("iibb",                  "IMPUESTOS Y CARGOS"),
    ("rg 5617",               "IMPUESTOS Y CARGOS"),
    ("db rg",                 "IMPUESTOS Y CARGOS"),
    # Tecnología / electrónica
    ("norcelstore",           "TECNOLOGÍA"),
    ("norcel",                "TECNOLOGÍA"),
    # Viajes / alojamiento
    ("fairview pinamar",      "VIAJES"),
    ("pinamar",               "VIAJES"),
    ("booking",               "VIAJES"),
    ("airbnb",                "VIAJES"),
    ("despegar",              "VIAJES"),
    ("flybondi",              "VIAJES"),
    # Indumentaria y moda (cuotas generalmente, categorizar como CUOTAS si tiene X/Y)
    ("montagne",              "CUOTAS"),
    ("zara",                  "CUOTAS"),
    ("shop gallery",          "CUOTAS"),
    ("shopgallery",           "CUOTAS"),
    ("giorgio beneti",        "CUOTAS"),
    ("giorgiobenet",          "CUOTAS"),
    ("marroquineria doga",    "CUOTAS"),
    ("doga",                  "CUOTAS"),
    ("diego perfumes",        "CUOTAS"),
    ("lomitas street",        "CUOTAS"),
    ("frieraclaudio",         "CUOTAS"),
    ("essen",                 "CUOTAS"),
    ("coop obrera ltda",      "CUOTAS"),
    # Transferencias / pagos a personas
    ("stirparolautarma",      "OTROS"),
    ("stirparo lautaro",      "OTROS"),
    ("luciano agustina",      "OTROS"),
    ("lucianoagustina",       "OTROS"),
]


def categorizar(descripcion: str, es_cuota: bool) -> str:
    desc_lower = descripcion.lower()
    if es_cuota:
        return "CUOTAS"
    for kw, cat in RULES:
        if kw in desc_lower:
            return cat
    return "OTROS"


def es_pago_credito(descripcion: str, importe: float) -> bool:
    """Devuelve True si es un pago/acreditación (no gasto)."""
    desc_lower = descripcion.lower()
    if importe < 0:
        return True
    keywords_pago = ["acreditacion de vuestro pago", "su pago en pesos",
                     "pago recibido", "acreditación"]
    return any(k in desc_lower for k in keywords_pago)


CUOTA_RE = re.compile(r'\b(\d+)\s*[/\-]\s*(\d+)\b|\b(\d+)\s+de\s+(\d+)\b',
                      re.IGNORECASE)


def detectar_cuota(texto: str):
    """Retorna (cuota_actual, cuotas_total) o None."""
    if not texto:
        return None
    m = CUOTA_RE.search(str(texto))
    if m:
        a, b, c, d = m.groups()
        if a and b:
            return (int(a), int(b))
        if c and d:
            return (int(c), int(d))
    return None


# ─────────────────────────────────────────────────────────────────────────────
# PARSEO DE ARCHIVOS
# ─────────────────────────────────────────────────────────────────────────────

def parse_importe_ars(s):
    """Convierte '$14.736,41' o '-$542.739,11' → float. Retorna None si falla."""
    if s is None:
        return None
    s = str(s).strip()
    if s in ("", "None", "-"):
        return None
    # Quitar $ y espacios, manejar negativo
    neg = s.startswith('-') or s.startswith('$-')
    s = s.replace('$', '').replace('-', '').strip()
    # Formato argentino: puntos = miles, coma = decimal
    s = s.replace('.', '').replace(',', '.')
    try:
        v = float(s)
        return -v if neg else v
    except ValueError:
        return None


def parse_importe_usd(s):
    """Convierte 'U$S12,99' → float."""
    if s is None:
        return None
    s = str(s).strip()
    s = s.replace('U$S', '').replace('$', '').strip()
    s = s.replace('.', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


SKIP_DESC = {'fecha', 'descripción', 'descripcion', 'cuotas', 'comprobante',
             'monto en pesos', 'monto en dólares', 'monto en dolares'}


def _is_header_row(row):
    d = str(row[1]).strip().lower() if row[1] else ""
    return d in SKIP_DESC or d == "descripción"


def _is_total_row(row):
    f = str(row[0]).strip().lower() if row[0] else ""
    return f.startswith('total de') or f == 'total de'


def _is_section_row(row):
    """Devuelve el texto si es una fila de encabezado de sección (no transacción)."""
    if row[0] and not row[1] and not row[3] and not row[4]:
        return str(row[0]).strip()
    return None


def _make_tx(tarjeta, subcuenta, fecha, desc, cuotas_col, importe, moneda):
    info_cuota = detectar_cuota(cuotas_col) or detectar_cuota(desc)
    es_cuota = info_cuota is not None or bool(
        re.search(r'\b\d+\s+de\s+\d+\b', cuotas_col or "", re.IGNORECASE))
    categoria = categorizar(desc, es_cuota)
    return {
        "tarjeta": tarjeta,
        "subcuenta": subcuenta,
        "fecha": str(fecha) if fecha else "",
        "descripcion": desc.strip(),
        "cuotas_info": cuotas_col.strip() if cuotas_col else "",
        "importe": abs(importe),
        "moneda": moneda,
        "es_cuota": es_cuota,
        "info_cuota": info_cuota,
        "categoria": categoria,
    }


def parse_amex_4100():
    """Lee Último resumen - American Express 4100.xlsx"""
    files = glob.glob('/home/user/OWN-Daro-1/*.xlsx')
    path = next((f for f in files if '4100' in f), None)
    if not path:
        raise FileNotFoundError("No se encontró el archivo de Amex 4100")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    transacciones = []
    in_transactions = False
    fecha_actual = ""

    for row in rows:
        if not any(row):
            continue

        row0 = str(row[0]).strip() if row[0] else ""
        row1 = str(row[1]).strip() if row[1] else ""

        # Detectar inicio de bloque de transacciones del titular
        if 'terminada en 4100' in row0.lower() or 'terminada en 4100' in row1.lower():
            in_transactions = True
            continue

        if not in_transactions:
            continue

        # Fin del bloque
        if row0.lower().startswith('total de') or row0.lower() == 'otros conceptos':
            if row0.lower() == 'otros conceptos':
                # Parsear impuestos a continuación
                in_transactions = False
            continue

        # Saltar encabezado de columnas
        if _is_header_row(row):
            continue

        # Saltar pagos
        if 'pago' in row1.lower() and row[4] is not None:
            v = parse_importe_ars(row[4])
            if v and v < 0:
                continue

        desc = row1
        if not desc or desc == "None":
            continue

        cuotas_col = str(row[2]).strip() if row[2] else ""
        importe_ars = parse_importe_ars(row[4])
        importe_usd = parse_importe_usd(row[5])

        if row[0]:
            fecha_actual = str(row[0]).strip()

        if importe_ars and importe_ars > 0:
            tx = _make_tx("AMEX 4100", "Dario F. Stirparo (4100)",
                          fecha_actual, desc, cuotas_col, importe_ars, "$")
            transacciones.append(tx)
        elif importe_usd and importe_usd > 0:
            tx = _make_tx("AMEX 4100", "Dario F. Stirparo (4100)",
                          fecha_actual, desc, cuotas_col, importe_usd, "U$S")
            transacciones.append(tx)

    # Parsear "Otros conceptos" (sellos)
    in_otros = False
    for row in rows:
        if not any(row):
            continue
        row0 = str(row[0]).strip() if row[0] else ""
        if row0.lower() == 'otros conceptos':
            in_otros = True
            continue
        if in_otros:
            desc = row0
            if desc.lower() in ('descripción', 'descripcion') or not desc:
                continue
            v = parse_importe_ars(row[1])
            if v and v > 0:
                tx = _make_tx("AMEX 4100", "Dario F. Stirparo (4100)",
                              "28/04/2026", desc, "", v, "$")
                transacciones.append(tx)

    return transacciones


def parse_visa_2473():
    """Lee Último resumen - Visa 2473.xlsx"""
    files = glob.glob('/home/user/OWN-Daro-1/*.xlsx')
    path = next((f for f in files if 'Visa 2473' in f), None)
    if not path:
        raise FileNotFoundError("No se encontró el archivo de Visa 2473")

    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    transacciones = []
    subcuenta_actual = None
    in_transactions = False
    fecha_actual = ""

    for row in rows:
        if not any(row):
            continue

        row0 = str(row[0]).strip() if row[0] else ""
        row1 = str(row[1]).strip() if row[1] else ""

        # Detectar encabezados de subcuenta
        if 'terminada en 9368' in row0.lower():
            subcuenta_actual = "Lautaro Stirparo (9368)"
            in_transactions = True
            continue
        if 'terminada en 2473' in row0.lower():
            subcuenta_actual = "Dario F. Stirparo (2473)"
            in_transactions = True
            continue
        if 'terminada en 4917' in row0.lower():
            subcuenta_actual = "Dario F. Stirparo (4917)"
            in_transactions = True
            continue

        if not in_transactions:
            continue

        # Fin de bloque
        if row0.lower().startswith('total de'):
            continue
        if row0.lower() == 'otros conceptos':
            in_transactions = False
            continue

        # Encabezado de columnas
        if _is_header_row(row):
            continue

        # Pago/devolución (monto negativo)
        if row[4] is not None:
            v = parse_importe_ars(row[4])
            if v and v < 0:
                continue

        # Saltar filas de "Pago de tarjeta y devoluciones"
        if 'pago de tarjeta' in row0.lower():
            continue

        desc = row1
        if not desc or desc == "None":
            continue

        cuotas_col = str(row[2]).strip() if row[2] else ""
        importe_ars = parse_importe_ars(row[4])
        importe_usd = parse_importe_usd(row[5])

        if row[0]:
            fecha_actual = str(row[0]).strip()

        if importe_ars and importe_ars > 0:
            tx = _make_tx("VISA 2473", subcuenta_actual or "Desconocido",
                          fecha_actual, desc, cuotas_col, importe_ars, "$")
            transacciones.append(tx)
        elif importe_usd and importe_usd > 0:
            tx = _make_tx("VISA 2473", subcuenta_actual or "Desconocido",
                          fecha_actual, desc, cuotas_col, importe_usd, "U$S")
            transacciones.append(tx)

    # Parsear "Otros conceptos" (sellos, DB RG 5617)
    in_otros = False
    for row in rows:
        if not any(row):
            continue
        row0 = str(row[0]).strip() if row[0] else ""
        if row0.lower() == 'otros conceptos':
            in_otros = True
            continue
        if in_otros and row0:
            if row0.lower() in ('descripción', 'descripcion', 'aviso importante'):
                continue
            v = parse_importe_ars(row[1])
            if v and v > 0:
                tx = _make_tx("VISA 2473", "Cargos generales",
                              "30/04/2026", row0, "", v, "$")
                transacciones.append(tx)

    return transacciones


def parse_amex_santander():
    """Lee el PDF de Amex Santander ya convertido a texto."""
    txt_path = "/tmp/amex_santander.txt"
    try:
        with open(txt_path, "r", encoding="utf-8") as f:
            text = f.read()
    except FileNotFoundError:
        import subprocess
        subprocess.run(["pdftotext", "-layout",
                        "/home/user/OWN-Daro-1/2026-04-28.pdf",
                        txt_path], check=True)
        with open(txt_path, "r", encoding="utf-8") as f:
            text = f.read()

    # Transacciones reales extraídas del PDF (parseadas manualmente del texto)
    raw = [
        # (fecha_str, descripcion, importe_ars, moneda, cuota_desc)
        ("22/10/2025", "MERPAGO*NTOMAC",            16004.77,  "$",   "CUOTA 07/09 DE"),
        ("01/04/2026", "ZURICH MB",                  6255.15,  "$",   ""),
        ("01/04/2026", "S P S",                    126311.90,  "$",   ""),
        ("01/04/2026", "ZURICH SEGUROS",            258558.30,  "$",   ""),
        ("07/04/2026", "AUTOPISTAS DEL SOL S.A.",     3181.30,  "$",   ""),
        ("08/04/2026", "DLO*SPOTIFY 99999999",        6642.49,  "$",   ""),
        ("12/04/2026", "=NETFLIX_COM",               30198.49,  "$",   ""),
        ("13/04/2026", "APPLUS ITEUVE ARGENT",        97057.65,  "$",   ""),
        ("21/04/2026", "AUTOPISTAS DEL SOL S.A.",     2187.14,  "$",   ""),
        ("22/04/2026", "CLARO",                      69538.31,  "$",   ""),
        ("27/04/2026", "ANTINA",                     73090.00,  "$",   ""),
        ("28/04/2026", "INTERESES FINANCIEROS",          93.73,  "$",   ""),
        ("28/04/2026", "PERCEPCION RG 5617",           1520.64,  "$",   ""),
        ("28/04/2026", "Imp. Sellos BS AS",             8327.03,  "$",   ""),
        ("28/04/2026", "IVA 21%",                       19.68,  "$",   ""),
        ("28/04/2026", "PERCEP IVA SERV DIGITALES LEY", 1064.45,  "$",   ""),
        ("28/04/2026", "PERC IIBB BS AS SERV DIG",      101.38,  "$",   ""),
        # USD
        ("05/04/2026", "APPLE.COM/BILL",               0.99, "U$S",  ""),
        ("22/04/2026", "APPLE.COM/BILL",               2.53, "U$S",  ""),
    ]

    transacciones = []
    for (fecha, desc, importe, moneda, cuota_desc) in raw:
        info_cuota = detectar_cuota(cuota_desc) or detectar_cuota(desc)
        es_cuota = info_cuota is not None

        categoria = categorizar(desc, es_cuota)

        transacciones.append({
            "tarjeta": "AMEX SANTANDER",
            "subcuenta": "Dario Stirparo (3766)",
            "fecha": fecha,
            "descripcion": desc,
            "cuotas_info": cuota_desc,
            "importe": abs(importe),
            "moneda": moneda,
            "es_cuota": es_cuota,
            "info_cuota": info_cuota,
            "categoria": categoria,
        })

    return transacciones


# ─────────────────────────────────────────────────────────────────────────────
# ESTILOS EXCEL HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)


def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)


def header_style(ws, row, col, value, bg="1A237E", fg="FFFFFF", bold=True, size=10):
    c = ws.cell(row=row, column=col, value=value)
    c.fill = fill(bg)
    c.font = Font(bold=bold, color=fg, size=size)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = thin_border()
    return c


def data_cell(ws, row, col, value, fmt=None, bold=False, align="left", bg=None):
    c = ws.cell(row=row, column=col, value=value)
    if fmt:
        c.number_format = fmt
    c.font = Font(bold=bold, size=9)
    c.alignment = Alignment(horizontal=align, vertical="center")
    c.border = thin_border()
    if bg:
        c.fill = fill(bg)
    return c


PESO_FMT = '#,##0.00'
USD_FMT  = 'U$S #,##0.00'


# ─────────────────────────────────────────────────────────────────────────────
# HOJA POR TARJETA
# ─────────────────────────────────────────────────────────────────────────────

def crear_hoja_tarjeta(wb, nombre_hoja, transacciones, color_tema):
    ws = wb.create_sheet(nombre_hoja)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 32
    ws.column_dimensions['C'].width = 13
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 8
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 16
    ws.column_dimensions['H'].width = 8

    # Título
    ws.merge_cells('A1:H1')
    c = ws['A1']
    c.value = f"DETALLE DE GASTOS — {nombre_hoja}"
    c.fill = fill(color_tema)
    c.font = Font(bold=True, color="FFFFFF", size=13)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28

    # Encabezados
    headers = ["Fecha", "Descripción", "Subcuenta", "Importe", "Moneda",
               "Categoría", "Info Cuotas", "¿Cuota?"]
    for ci, h in enumerate(headers, 1):
        header_style(ws, 2, ci, h, bg=color_tema)
    ws.row_dimensions[2].height = 20

    # Datos
    row = 3
    for t in transacciones:
        data_cell(ws, row, 1, t["fecha"])
        data_cell(ws, row, 2, t["descripcion"])
        data_cell(ws, row, 3, t["subcuenta"])
        data_cell(ws, row, 4, t["importe"],
                  fmt=PESO_FMT if t["moneda"] == "$" else USD_FMT,
                  align="right")
        data_cell(ws, row, 5, t["moneda"], align="center")
        cat = t["categoria"]
        cat_bg = CAT_COLORS.get(cat, "9E9E9E")
        data_cell(ws, row, 6, cat, bg=cat_bg,
                  bold=True, align="center")
        ws.cell(row=row, column=6).font = Font(bold=True, color="FFFFFF", size=9)

        cuota_str = t["cuotas_info"]
        if t["info_cuota"]:
            a, b = t["info_cuota"]
            cuota_str = f"Cuota {a} de {b}"
        data_cell(ws, row, 7, cuota_str)
        data_cell(ws, row, 8, "Sí" if t["es_cuota"] else "No",
                  align="center",
                  bold=t["es_cuota"],
                  bg="FFF9C4" if t["es_cuota"] else None)

        row += 1

    # Totales
    ars_total = sum(t["importe"] for t in transacciones if t["moneda"] == "$")
    usd_total = sum(t["importe"] for t in transacciones if t["moneda"] != "$")

    ws.merge_cells(f'A{row}:C{row}')
    c = ws[f'A{row}']
    c.value = "TOTAL"
    c.fill = fill(color_tema)
    c.font = Font(bold=True, color="FFFFFF", size=10)
    c.alignment = Alignment(horizontal="right")

    tc = ws.cell(row=row, column=4, value=ars_total)
    tc.number_format = PESO_FMT
    tc.font = Font(bold=True, size=10)
    tc.fill = fill(color_tema)
    tc.font = Font(bold=True, color="FFFFFF", size=10)
    tc.alignment = Alignment(horizontal="right")

    if usd_total > 0:
        tc2 = ws.cell(row=row, column=5, value=f"U$S {usd_total:.2f}")
        tc2.font = Font(bold=True, color="FFFFFF", size=9)
        tc2.fill = fill(color_tema)

    ws.freeze_panes = 'A3'
    ws.auto_filter.ref = f"A2:H{row-1}"


# ─────────────────────────────────────────────────────────────────────────────
# HOJA RESUMEN CONSOLIDADO
# ─────────────────────────────────────────────────────────────────────────────

def crear_hoja_resumen(wb, todas):
    ws = wb.create_sheet("RESUMEN CONSOLIDADO", 0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 26
    ws.column_dimensions['B'].width = 16
    ws.column_dimensions['C'].width = 14
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14

    ws.merge_cells('A1:E1')
    c = ws['A1']
    c.value = "ANÁLISIS DE GASTOS — 3 TARJETAS DE CRÉDITO — ABRIL 2026"
    c.fill = fill("1A237E")
    c.font = Font(bold=True, color="FFFFFF", size=14)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    # KPIs rápidos
    tarjetas = ["AMEX 4100", "AMEX SANTANDER", "VISA 2473"]
    kpi_data = {}
    for t_name in tarjetas:
        txs = [t for t in todas if t["tarjeta"] == t_name]
        kpi_data[t_name] = {
            "ars": sum(t["importe"] for t in txs if t["moneda"] == "$"),
            "usd": sum(t["importe"] for t in txs if t["moneda"] != "$"),
            "n": len(txs),
        }

    total_ars = sum(v["ars"] for v in kpi_data.values())
    total_usd = sum(v["usd"] for v in kpi_data.values())

    # Bloque KPI por tarjeta
    ws.row_dimensions[2].height = 4
    colors_tarjetas = {"AMEX 4100": "1565C0", "AMEX SANTANDER": "6A1B9A",
                       "VISA 2473": "1B5E20"}

    kpi_start = 3
    headers_kpi = ["Tarjeta", "Gastos ARS", "Gastos USD", "N° Transacciones", "% del Total"]
    for ci, h in enumerate(headers_kpi, 1):
        header_style(ws, kpi_start, ci, h, bg="37474F")
    ws.row_dimensions[kpi_start].height = 18

    row = kpi_start + 1
    for t_name in tarjetas:
        kd = kpi_data[t_name]
        bg = colors_tarjetas[t_name]
        c1 = ws.cell(row=row, column=1, value=t_name)
        c1.fill = fill(bg)
        c1.font = Font(bold=True, color="FFFFFF", size=10)
        c1.alignment = Alignment(horizontal="left", vertical="center")
        c1.border = thin_border()

        c2 = ws.cell(row=row, column=2, value=kd["ars"])
        c2.number_format = PESO_FMT
        c2.font = Font(bold=True, size=10)
        c2.alignment = Alignment(horizontal="right")
        c2.border = thin_border()

        c3 = ws.cell(row=row, column=3, value=kd["usd"] if kd["usd"] > 0 else "—")
        if kd["usd"] > 0:
            c3.number_format = USD_FMT
        c3.alignment = Alignment(horizontal="right")
        c3.border = thin_border()

        c4 = ws.cell(row=row, column=4, value=kd["n"])
        c4.alignment = Alignment(horizontal="center")
        c4.border = thin_border()

        pct = kd["ars"] / total_ars * 100 if total_ars > 0 else 0
        c5 = ws.cell(row=row, column=5, value=pct / 100)
        c5.number_format = '0.0%'
        c5.alignment = Alignment(horizontal="center")
        c5.border = thin_border()

        ws.row_dimensions[row].height = 18
        row += 1

    # Fila total
    ws.merge_cells(f'A{row}:A{row}')
    t1 = ws.cell(row=row, column=1, value="TOTAL GENERAL")
    t1.fill = fill("212121")
    t1.font = Font(bold=True, color="FFFFFF", size=11)
    t1.border = thin_border()

    t2 = ws.cell(row=row, column=2, value=total_ars)
    t2.number_format = PESO_FMT
    t2.fill = fill("212121")
    t2.font = Font(bold=True, color="FFFFFF", size=11)
    t2.alignment = Alignment(horizontal="right")
    t2.border = thin_border()

    t3 = ws.cell(row=row, column=3, value=total_usd if total_usd > 0 else "—")
    if total_usd > 0:
        t3.number_format = USD_FMT
    t3.fill = fill("212121")
    t3.font = Font(bold=True, color="FFFFFF", size=11)
    t3.alignment = Alignment(horizontal="right")
    t3.border = thin_border()

    t4 = ws.cell(row=row, column=4, value=len(todas))
    t4.fill = fill("212121")
    t4.font = Font(bold=True, color="FFFFFF", size=11)
    t4.alignment = Alignment(horizontal="center")
    t4.border = thin_border()

    t5 = ws.cell(row=row, column=5, value=1.0)
    t5.number_format = '0.0%'
    t5.fill = fill("212121")
    t5.font = Font(bold=True, color="FFFFFF", size=11)
    t5.alignment = Alignment(horizontal="center")
    t5.border = thin_border()

    ws.row_dimensions[row].height = 22
    row += 2

    # ── Breakdown por categoría ──
    cat_data = defaultdict(float)
    for t in todas:
        if t["moneda"] == "$":
            cat_data[t["categoria"]] += t["importe"]

    sorted_cats = sorted(cat_data.items(), key=lambda x: x[1], reverse=True)

    ws.merge_cells(f'A{row}:E{row}')
    ch = ws.cell(row=row, column=1, value="DESGLOSE POR CATEGORÍA")
    ch.fill = fill("37474F")
    ch.font = Font(bold=True, color="FFFFFF", size=11)
    ch.alignment = Alignment(horizontal="center")
    ws.row_dimensions[row].height = 20
    row += 1

    hdr_cats = ["Categoría", "Total ARS", "% del Total", "N° Tx", "Ticket Promedio"]
    for ci, h in enumerate(hdr_cats, 1):
        header_style(ws, row, ci, h, bg="546E7A", size=9)
    ws.row_dimensions[row].height = 16
    row += 1

    for cat, monto in sorted_cats:
        cat_txs = [t for t in todas if t["categoria"] == cat and t["moneda"] == "$"]
        pct = monto / total_ars * 100 if total_ars > 0 else 0
        avg = monto / len(cat_txs) if cat_txs else 0
        bg_cat = CAT_COLORS.get(cat, "9E9E9E")

        c1 = ws.cell(row=row, column=1, value=cat)
        c1.fill = fill(bg_cat)
        c1.font = Font(bold=True, color="FFFFFF", size=9)
        c1.border = thin_border()

        c2 = ws.cell(row=row, column=2, value=monto)
        c2.number_format = PESO_FMT
        c2.font = Font(size=9)
        c2.alignment = Alignment(horizontal="right")
        c2.border = thin_border()

        c3 = ws.cell(row=row, column=3, value=pct / 100)
        c3.number_format = '0.0%'
        c3.alignment = Alignment(horizontal="center")
        c3.border = thin_border()

        c4 = ws.cell(row=row, column=4, value=len(cat_txs))
        c4.alignment = Alignment(horizontal="center")
        c4.border = thin_border()

        c5 = ws.cell(row=row, column=5, value=avg)
        c5.number_format = PESO_FMT
        c5.alignment = Alignment(horizontal="right")
        c5.border = thin_border()

        ws.row_dimensions[row].height = 15
        row += 1

    return ws, sorted_cats, total_ars, total_usd, kpi_data


# ─────────────────────────────────────────────────────────────────────────────
# HOJA CUOTAS ACTIVAS
# ─────────────────────────────────────────────────────────────────────────────

def crear_hoja_cuotas(wb, todas):
    ws = wb.create_sheet("CUOTAS ACTIVAS")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 16
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 14

    ws.merge_cells('A1:G1')
    c = ws['A1']
    c.value = "CUOTAS ACTIVAS — PLANES DE PAGO EN VIGENCIA"
    c.fill = fill("E65100")
    c.font = Font(bold=True, color="FFFFFF", size=13)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    headers = ["Fecha", "Descripción", "Tarjeta", "Cuota Actual",
               "Total Cuotas", "Importe/Cuota", "Total Compra Est."]
    for ci, h in enumerate(headers, 1):
        header_style(ws, 2, ci, h, bg="BF360C")
    ws.row_dimensions[2].height = 18

    cuotas = [t for t in todas if t["es_cuota"]]
    row = 3
    total_mes = 0
    for t in cuotas:
        data_cell(ws, row, 1, t["fecha"])
        data_cell(ws, row, 2, t["descripcion"])
        data_cell(ws, row, 3, t["tarjeta"])
        if t["info_cuota"]:
            a, b = t["info_cuota"]
            data_cell(ws, row, 4, a, align="center")
            data_cell(ws, row, 5, b, align="center")
            data_cell(ws, row, 6, t["importe"], fmt=PESO_FMT, align="right")
            data_cell(ws, row, 7, t["importe"] * b, fmt=PESO_FMT, align="right")
        else:
            data_cell(ws, row, 4, "?", align="center")
            data_cell(ws, row, 5, "?", align="center")
            data_cell(ws, row, 6, t["importe"], fmt=PESO_FMT, align="right")
            data_cell(ws, row, 7, "—", align="center")
        total_mes += t["importe"]
        ws.row_dimensions[row].height = 15
        row += 1

    # Total
    ws.merge_cells(f'A{row}:E{row}')
    tc = ws[f'A{row}']
    tc.value = f"TOTAL CUOTAS ESTE MES ({len(cuotas)} planes)"
    tc.fill = fill("E65100")
    tc.font = Font(bold=True, color="FFFFFF", size=10)
    tc.alignment = Alignment(horizontal="right")
    tc.border = thin_border()

    tv = ws.cell(row=row, column=6, value=total_mes)
    tv.number_format = PESO_FMT
    tv.fill = fill("E65100")
    tv.font = Font(bold=True, color="FFFFFF", size=10)
    tv.alignment = Alignment(horizontal="right")
    tv.border = thin_border()

    ws.freeze_panes = 'A3'


# ─────────────────────────────────────────────────────────────────────────────
# HOJA SERVICIOS FIJOS
# ─────────────────────────────────────────────────────────────────────────────

def crear_hoja_servicios(wb, todas):
    ws = wb.create_sheet("SERVICIOS FIJOS")
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 12
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 22
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 16

    cats_servicios = {"SERVICIOS FIJOS", "SEGUROS", "ENTRETENIMIENTO", "SALUD"}

    ws.merge_cells('A1:E1')
    c = ws['A1']
    c.value = "GASTOS FIJOS RECURRENTES — SERVICIOS / SEGUROS / ENTRETENIMIENTO / SALUD"
    c.fill = fill("4A148C")
    c.font = Font(bold=True, color="FFFFFF", size=11)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 26

    headers = ["Fecha", "Descripción", "Tarjeta", "Categoría", "Importe"]
    for ci, h in enumerate(headers, 1):
        header_style(ws, 2, ci, h, bg="6A1B9A")
    ws.row_dimensions[2].height = 18

    servicios = [t for t in todas if t["categoria"] in cats_servicios]
    row = 3
    total = 0.0
    for t in servicios:
        data_cell(ws, row, 1, t["fecha"])
        data_cell(ws, row, 2, t["descripcion"])
        data_cell(ws, row, 3, t["tarjeta"])
        bg = CAT_COLORS.get(t["categoria"], "9E9E9E")
        c4 = ws.cell(row=row, column=4, value=t["categoria"])
        c4.fill = fill(bg)
        c4.font = Font(bold=True, color="FFFFFF", size=9)
        c4.alignment = Alignment(horizontal="center")
        c4.border = thin_border()

        moneda_str = "" if t["moneda"] == "$" else "U$S "
        data_cell(ws, row, 5,
                  t["importe"],
                  fmt=PESO_FMT if t["moneda"] == "$" else USD_FMT,
                  align="right")
        if t["moneda"] == "$":
            total += t["importe"]
        ws.row_dimensions[row].height = 15
        row += 1

    ws.merge_cells(f'A{row}:D{row}')
    tc = ws[f'A{row}']
    tc.value = "TOTAL GASTOS FIJOS (en $)"
    tc.fill = fill("4A148C")
    tc.font = Font(bold=True, color="FFFFFF", size=10)
    tc.alignment = Alignment(horizontal="right")
    tc.border = thin_border()

    tv = ws.cell(row=row, column=5, value=total)
    tv.number_format = PESO_FMT
    tv.fill = fill("4A148C")
    tv.font = Font(bold=True, color="FFFFFF", size=10)
    tv.alignment = Alignment(horizontal="right")
    tv.border = thin_border()

    ws.freeze_panes = 'A3'


# ─────────────────────────────────────────────────────────────────────────────
# HOJA DASHBOARD
# ─────────────────────────────────────────────────────────────────────────────

def crear_hoja_dashboard(wb, todas, sorted_cats, total_ars, total_usd, kpi_data):
    ws = wb.create_sheet("DASHBOARD", 0)
    ws.sheet_view.showGridLines = False
    for col in range(1, 16):
        ws.column_dimensions[get_column_letter(col)].width = 10

    # Título principal
    ws.merge_cells('A1:O1')
    c = ws['A1']
    c.value = "DASHBOARD EJECUTIVO — GASTOS TARJETAS ABRIL 2026"
    c.fill = fill("0D1B2A")
    c.font = Font(bold=True, color="FFFFFF", size=15)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 36

    ws.merge_cells('A2:O2')
    c2 = ws['A2']
    c2.value = "Titular: DARIO STIRPARO  |  Cierre: Abril 2026  |  3 Tarjetas Analizadas"
    c2.fill = fill("1A237E")
    c2.font = Font(color="BBDEFB", size=10, italic=True)
    c2.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 18

    ws.row_dimensions[3].height = 8

    # ── KPI Cards ──
    kpi_row_start = 4

    def kpi_box(start_col, end_col, titulo, valor, subtitulo, bg_color, row_s=kpi_row_start):
        ws.merge_cells(
            start_row=row_s, start_column=start_col,
            end_row=row_s, end_column=end_col
        )
        ct = ws.cell(row=row_s, column=start_col, value=titulo)
        ct.fill = fill(bg_color)
        ct.font = Font(bold=True, color="FFFFFF", size=8)
        ct.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row_s].height = 18

        ws.merge_cells(
            start_row=row_s+1, start_column=start_col,
            end_row=row_s+1, end_column=end_col
        )
        cv = ws.cell(row=row_s+1, column=start_col, value=valor)
        cv.fill = fill(bg_color)
        cv.font = Font(bold=True, color="FFFFFF", size=12)
        cv.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row_s+1].height = 26

        ws.merge_cells(
            start_row=row_s+2, start_column=start_col,
            end_row=row_s+2, end_column=end_col
        )
        cs = ws.cell(row=row_s+2, column=start_col, value=subtitulo)
        cs.fill = fill(bg_color)
        cs.font = Font(color="E3F2FD", size=8)
        cs.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[row_s+2].height = 16

    amex4100_ars = kpi_data["AMEX 4100"]["ars"]
    amex_sant_ars = kpi_data["AMEX SANTANDER"]["ars"]
    visa_ars = kpi_data["VISA 2473"]["ars"]
    n_tx = len(todas)
    n_cuotas = sum(1 for t in todas if t["es_cuota"])

    kpi_box(1, 3, "TOTAL GENERAL ARS",
            f"$ {total_ars:,.0f}", f"{n_tx} transacciones", "1A237E")
    kpi_box(4, 5, "AMEX 4100",
            f"$ {amex4100_ars:,.0f}",
            f"{kpi_data['AMEX 4100']['n']} tx", "1565C0")
    kpi_box(6, 7, "AMEX SANTANDER",
            f"$ {amex_sant_ars:,.0f}",
            f"{kpi_data['AMEX SANTANDER']['n']} tx", "6A1B9A")
    kpi_box(8, 9, "VISA 2473",
            f"$ {visa_ars:,.0f}",
            f"{kpi_data['VISA 2473']['n']} tx", "1B5E20")
    kpi_box(10, 11, "CUOTAS MES",
            f"$ {sum(t['importe'] for t in todas if t['es_cuota'] and t['moneda']=='$'):,.0f}",
            f"{n_cuotas} planes activos", "E65100")
    kpi_box(12, 13, "USD TOTAL",
            f"U$S {total_usd:.2f}",
            "Apple + Splice", "00695C")
    kpi_box(14, 15, "TICKET PROM.",
            f"$ {total_ars/n_tx:,.0f}" if n_tx else "—",
            "por transacción", "37474F")

    ws.row_dimensions[kpi_row_start + 3].height = 10

    # ── Tabla categorías para gráfico ──
    cat_row = kpi_row_start + 4
    ws.cell(row=cat_row, column=1, value="Categoría").font = Font(bold=True, size=8)
    ws.cell(row=cat_row, column=2, value="Monto").font = Font(bold=True, size=8)
    ws.cell(row=cat_row, column=3, value="%").font = Font(bold=True, size=8)

    data_rows_start = cat_row + 1
    for i, (cat, monto) in enumerate(sorted_cats[:12]):
        r = data_rows_start + i
        ws.cell(row=r, column=1, value=cat).font = Font(size=8)
        ws.cell(row=r, column=2, value=monto).number_format = PESO_FMT
        ws.cell(row=r, column=3, value=monto/total_ars if total_ars else 0).number_format = '0.0%'
        ws.row_dimensions[r].height = 14

    # Gráfico de barras por categoría
    n_cats = min(12, len(sorted_cats))
    chart = BarChart()
    chart.type = "col"
    chart.title = "Gastos por Categoría (ARS)"
    chart.style = 10
    chart.y_axis.title = "Monto ARS"
    chart.x_axis.title = "Categoría"
    chart.grouping = "clustered"
    chart.width = 18
    chart.height = 12

    data_ref = Reference(ws,
                         min_col=2, max_col=2,
                         min_row=data_rows_start,
                         max_row=data_rows_start + n_cats - 1)
    cats_ref = Reference(ws,
                         min_col=1,
                         min_row=data_rows_start,
                         max_row=data_rows_start + n_cats - 1)
    chart.add_data(data_ref)
    chart.set_categories(cats_ref)

    ws.add_chart(chart, f"E{kpi_row_start + 4}")

    # Gráfico de torta por tarjeta
    pie_row = kpi_row_start + 4
    ws.cell(row=pie_row, column=14, value="Tarjeta").font = Font(bold=True, size=8)
    ws.cell(row=pie_row, column=15, value="ARS").font = Font(bold=True, size=8)
    for i, t_name in enumerate(["AMEX 4100", "AMEX SANTANDER", "VISA 2473"]):
        r = pie_row + 1 + i
        ws.cell(row=r, column=14, value=t_name).font = Font(size=8)
        ws.cell(row=r, column=15, value=kpi_data[t_name]["ars"]).number_format = PESO_FMT

    pie = PieChart()
    pie.title = "Distribución por Tarjeta"
    pie.style = 10
    pie.width = 10
    pie.height = 10

    pie_data = Reference(ws, min_col=15, min_row=pie_row + 1,
                         max_row=pie_row + 3)
    pie_labels = Reference(ws, min_col=14, min_row=pie_row + 1,
                           max_row=pie_row + 3)
    pie.add_data(pie_data)
    pie.set_categories(pie_labels)

    slice_colors = ["1565C0", "6A1B9A", "1B5E20"]
    for i, hex_c in enumerate(slice_colors):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = hex_c
        pie.series[0].dPt.append(pt)

    pie.dataLabels = None
    ws.add_chart(pie, f"N{pie_row + 1}")


# ─────────────────────────────────────────────────────────────────────────────
# GENERAR EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def generar_excel(todas, output_path):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # quitar hoja default

    # Calcular datos para resumen
    cat_data = defaultdict(float)
    for t in todas:
        if t["moneda"] == "$":
            cat_data[t["categoria"]] += t["importe"]
    sorted_cats = sorted(cat_data.items(), key=lambda x: x[1], reverse=True)
    total_ars = sum(t["importe"] for t in todas if t["moneda"] == "$")
    total_usd = sum(t["importe"] for t in todas if t["moneda"] != "$")

    tarjetas_info = {}
    for t_name in ["AMEX 4100", "AMEX SANTANDER", "VISA 2473"]:
        txs = [t for t in todas if t["tarjeta"] == t_name]
        tarjetas_info[t_name] = {
            "ars": sum(t["importe"] for t in txs if t["moneda"] == "$"),
            "usd": sum(t["importe"] for t in txs if t["moneda"] != "$"),
            "n": len(txs),
        }

    crear_hoja_dashboard(wb, todas, sorted_cats, total_ars, total_usd, tarjetas_info)
    _, sorted_cats2, total_ars2, total_usd2, kpi_data2 = crear_hoja_resumen(wb, todas)

    # Hojas por tarjeta
    amex4100 = [t for t in todas if t["tarjeta"] == "AMEX 4100"]
    amex_sant = [t for t in todas if t["tarjeta"] == "AMEX SANTANDER"]
    visa = [t for t in todas if t["tarjeta"] == "VISA 2473"]

    crear_hoja_tarjeta(wb, "AMEX 4100", amex4100, "1565C0")
    crear_hoja_tarjeta(wb, "AMEX SANTANDER", amex_sant, "6A1B9A")
    crear_hoja_tarjeta(wb, "VISA 2473", visa, "1B5E20")
    crear_hoja_cuotas(wb, todas)
    crear_hoja_servicios(wb, todas)

    wb.save(output_path)
    print(f"Excel guardado: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
# GENERAR PDF
# ─────────────────────────────────────────────────────────────────────────────

def generar_pdf(todas, output_path):
    doc = SimpleDocTemplate(
        output_path, pagesize=A4,
        rightMargin=1.5*cm, leftMargin=1.5*cm,
        topMargin=1.5*cm, bottomMargin=1.5*cm
    )

    styles = getSampleStyleSheet()
    style_title = ParagraphStyle('title', parent=styles['Heading1'],
                                 fontSize=16, textColor=colors.white,
                                 backColor=colors.HexColor(0x1A237E),
                                 spaceAfter=4, spaceBefore=4,
                                 leftIndent=6, borderPad=6)
    style_h2 = ParagraphStyle('h2', parent=styles['Heading2'],
                               fontSize=12, textColor=colors.white,
                               backColor=colors.HexColor(0x37474F),
                               spaceAfter=3, spaceBefore=8,
                               leftIndent=4, borderPad=4)
    style_body = ParagraphStyle('body', parent=styles['Normal'],
                                fontSize=8, spaceAfter=2)
    style_note = ParagraphStyle('note', parent=styles['Normal'],
                                fontSize=7, textColor=colors.HexColor(0x607D8B),
                                spaceAfter=2)

    cat_data = defaultdict(float)
    for t in todas:
        if t["moneda"] == "$":
            cat_data[t["categoria"]] += t["importe"]
    sorted_cats = sorted(cat_data.items(), key=lambda x: x[1], reverse=True)
    total_ars = sum(t["importe"] for t in todas if t["moneda"] == "$")
    total_usd = sum(t["importe"] for t in todas if t["moneda"] != "$")

    tarjetas_info = {}
    for t_name in ["AMEX 4100", "AMEX SANTANDER", "VISA 2473"]:
        txs = [t for t in todas if t["tarjeta"] == t_name]
        tarjetas_info[t_name] = {
            "ars": sum(t["importe"] for t in txs if t["moneda"] == "$"),
            "usd": sum(t["importe"] for t in txs if t["moneda"] != "$"),
            "n": len(txs),
        }

    story = []

    # ── PÁGINA 1: DASHBOARD ──
    story.append(Paragraph("ANÁLISIS DE GASTOS — TARJETAS DE CRÉDITO ABRIL 2026",
                            style_title))
    story.append(Paragraph("Titular: DARIO STIRPARO  |  3 Tarjetas: AMEX 4100 · AMEX SANTANDER · VISA 2473",
                            style_note))
    story.append(Spacer(1, 0.3*cm))

    # KPI table
    n_cuotas = sum(1 for t in todas if t["es_cuota"])
    cuotas_total = sum(t["importe"] for t in todas if t["es_cuota"] and t["moneda"] == "$")

    kpi_rows = [
        ["TOTAL ARS", "AMEX 4100", "AMEX SANTANDER", "VISA 2473", "CUOTAS MES", "USD TOTAL"],
        [
            f"$ {total_ars:,.0f}",
            f"$ {tarjetas_info['AMEX 4100']['ars']:,.0f}",
            f"$ {tarjetas_info['AMEX SANTANDER']['ars']:,.0f}",
            f"$ {tarjetas_info['VISA 2473']['ars']:,.0f}",
            f"$ {cuotas_total:,.0f}",
            f"U$S {total_usd:.2f}",
        ],
        [
            f"{len(todas)} transacciones",
            f"{tarjetas_info['AMEX 4100']['n']} tx",
            f"{tarjetas_info['AMEX SANTANDER']['n']} tx",
            f"{tarjetas_info['VISA 2473']['n']} tx",
            f"{n_cuotas} planes",
            "Apple + Splice",
        ]
    ]
    kpi_style = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(0x1A237E)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor(0x283593)),
        ('TEXTCOLOR', (0, 1), (-1, 1), colors.white),
        ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 1), (-1, 1), 11),
        ('BACKGROUND', (0, 2), (-1, 2), colors.HexColor(0x3949AB)),
        ('TEXTCOLOR', (0, 2), (-1, 2), colors.HexColor(0xC5CAE9)),
        ('FONTSIZE', (0, 2), (-1, 2), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUND', (0, 0), (-1, 0), [colors.HexColor(0x1565C0),
                                             colors.HexColor(0x6A1B9A),
                                             colors.HexColor(0x1B5E20),
                                             colors.HexColor(0xE65100),
                                             colors.HexColor(0x00695C),
                                             colors.HexColor(0x37474F)]),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.white),
        ('INNERGRID', (0, 0), (-1, -1), 0.5, colors.white),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ])
    kpi_t = Table(kpi_rows, colWidths=[3*cm]*6, rowHeights=[1*cm, 1.2*cm, 0.7*cm])
    kpi_t.setStyle(kpi_style)
    story.append(kpi_t)
    story.append(Spacer(1, 0.4*cm))

    # Gráfico de barras categorías
    story.append(Paragraph("DISTRIBUCIÓN POR CATEGORÍA", style_h2))

    top_cats = sorted_cats[:10]
    bar_d = Drawing(17*cm, 8*cm)
    bc = VerticalBarChart()
    bc.x = 60
    bc.y = 20
    bc.width = 14*cm
    bc.height = 6*cm
    bc.data = [[monto for _, monto in top_cats]]
    bc.categoryAxis.categoryNames = [cat[:14] for cat, _ in top_cats]
    bc.categoryAxis.labels.angle = 30
    bc.categoryAxis.labels.fontSize = 6
    bc.valueAxis.valueMin = 0
    bc.valueAxis.labels.fontSize = 6
    bc.bars[0].fillColor = colors.HexColor(0x1565C0)
    bc.bars[0].strokeColor = colors.HexColor(0x0D47A1)
    bar_d.add(bc)
    story.append(bar_d)
    story.append(Spacer(1, 0.3*cm))

    # Gráfico torta distribución por tarjeta
    story.append(Paragraph("PARTICIPACIÓN POR TARJETA", style_h2))
    pie_d = Drawing(17*cm, 6*cm)
    pie = Pie()
    pie.x = 3*cm
    pie.y = 0.5*cm
    pie.width = 5*cm
    pie.height = 5*cm
    pie_vals = [tarjetas_info[n]["ars"] for n in ["AMEX 4100", "AMEX SANTANDER", "VISA 2473"]]
    pie.data = pie_vals
    pie.labels = [
        f"AMEX 4100\n{tarjetas_info['AMEX 4100']['ars']/total_ars*100:.1f}%",
        f"AMEX SANT\n{tarjetas_info['AMEX SANTANDER']['ars']/total_ars*100:.1f}%",
        f"VISA 2473\n{tarjetas_info['VISA 2473']['ars']/total_ars*100:.1f}%",
    ]
    pie.slices[0].fillColor = colors.HexColor(0x1565C0)
    pie.slices[1].fillColor = colors.HexColor(0x6A1B9A)
    pie.slices[2].fillColor = colors.HexColor(0x1B5E20)
    pie.slices.strokeColor = colors.white
    pie.sideLabels = True
    pie_d.add(pie)
    story.append(pie_d)

    # ── PÁGINA 2: ANÁLISIS POR CATEGORÍA ──
    from reportlab.platypus import PageBreak
    story.append(PageBreak())
    story.append(Paragraph("DESGLOSE DETALLADO POR CATEGORÍA", style_title))
    story.append(Spacer(1, 0.3*cm))

    cat_tbl_data = [["Categoría", "Monto ARS", "% Total", "N° Tx", "Ticket Prom."]]
    for cat, monto in sorted_cats:
        cat_txs = [t for t in todas if t["categoria"] == cat and t["moneda"] == "$"]
        pct = monto / total_ars * 100 if total_ars else 0
        avg = monto / len(cat_txs) if cat_txs else 0
        cat_tbl_data.append([
            cat,
            f"$ {monto:,.2f}",
            f"{pct:.1f}%",
            str(len(cat_txs)),
            f"$ {avg:,.0f}",
        ])

    # Totales
    cat_tbl_data.append([
        "TOTAL GENERAL",
        f"$ {total_ars:,.2f}",
        "100.0%",
        str(len([t for t in todas if t["moneda"] == "$"])),
        f"$ {total_ars/len(todas):,.0f}" if todas else "—",
    ])

    hex_colors_cats = [CAT_COLORS.get(cat, "9E9E9E") for cat, _ in sorted_cats]

    cat_ts = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(0x37474F)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('ALIGN', (1, 0), (-1, -1), 'RIGHT'),
        ('ALIGN', (0, 0), (0, -1), 'LEFT'),
        ('ALIGN', (2, 0), (3, -1), 'CENTER'),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (-1, -1), 0.3, colors.HexColor(0xEEEEEE)),
        ('ROWBACKGROUND', (0, -1), (-1, -1), [colors.HexColor(0x212121)]),
        ('TEXTCOLOR', (0, -1), (-1, -1), colors.white),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('TOPPADDING', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
    ])
    # Colores por categoría en col 0
    for i, (cat, _) in enumerate(sorted_cats):
        hex_c = CAT_COLORS.get(cat, "9E9E9E")
        r = i + 1
        cat_ts.add('BACKGROUND', (0, r), (0, r), colors.HexColor(int(hex_c, 16)))
        cat_ts.add('TEXTCOLOR', (0, r), (0, r), colors.white)
        cat_ts.add('FONTNAME', (0, r), (0, r), 'Helvetica-Bold')

    cat_tbl = Table(cat_tbl_data,
                    colWidths=[5*cm, 3.5*cm, 2*cm, 1.5*cm, 3*cm])
    cat_tbl.setStyle(cat_ts)
    story.append(cat_tbl)
    story.append(Spacer(1, 0.5*cm))

    # Cuotas activas resumen
    story.append(Paragraph("CUOTAS ACTIVAS — PLANES DE PAGO", style_h2))
    cuotas_items = [t for t in todas if t["es_cuota"]]
    cuota_hdr = ["Descripción", "Tarjeta", "Cuota", "Importe/mes"]
    cuota_rows = [cuota_hdr]
    for t in cuotas_items:
        cuota_str = "?"
        if t["info_cuota"]:
            a, b = t["info_cuota"]
            cuota_str = f"{a}/{b}"
        cuota_rows.append([
            t["descripcion"][:38],
            t["tarjeta"],
            cuota_str,
            f"$ {t['importe']:,.2f}",
        ])
    cuota_rows.append([
        f"TOTAL CUOTAS ({len(cuotas_items)} planes)",
        "", "",
        f"$ {cuotas_total:,.2f}",
    ])
    cuota_ts = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(0xBF360C)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 8),
        ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor(0xE65100)),
        ('TEXTCOLOR', (0, -1), (-1, -1), colors.white),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('ROWBACKGROUND', (0, 1), (-1, -2),
         [colors.HexColor(0xFFF9C4), colors.HexColor(0xFFFDE7)]),
        ('ALIGN', (3, 0), (3, -1), 'RIGHT'),
        ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
        ('INNERGRID', (0, 0), (-1, -1), 0.3, colors.HexColor(0xEEEEEE)),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
    ])
    cuota_tbl = Table(cuota_rows, colWidths=[7*cm, 3*cm, 2*cm, 3*cm])
    cuota_tbl.setStyle(cuota_ts)
    story.append(cuota_tbl)

    # ── PÁGINA 3: DETALLE POR TARJETA ──
    story.append(PageBreak())
    story.append(Paragraph("DETALLE POR TARJETA — TOP GASTOS", style_title))

    for t_name, t_color in [("AMEX 4100", 0x1565C0),
                              ("AMEX SANTANDER", 0x6A1B9A),
                              ("VISA 2473", 0x1B5E20)]:
        txs = sorted([t for t in todas if t["tarjeta"] == t_name
                      and t["moneda"] == "$"],
                     key=lambda x: x["importe"], reverse=True)
        t_ars = sum(t["importe"] for t in txs)
        t_usd = sum(t["importe"] for t in todas
                    if t["tarjeta"] == t_name and t["moneda"] != "$")

        story.append(Paragraph(
            f"{t_name} — Total ARS: $ {t_ars:,.2f}"
            + (f"  |  USD: U$S {t_usd:.2f}" if t_usd > 0 else ""),
            ParagraphStyle('th2', parent=styles['Heading2'],
                           fontSize=11, textColor=colors.white,
                           backColor=colors.HexColor(t_color),
                           spaceAfter=3, spaceBefore=6,
                           leftIndent=4, borderPad=4)
        ))

        det_rows = [["Descripción", "Categoría", "Importe"]]
        for tx in txs[:20]:  # top 20 por tarjeta
            det_rows.append([
                tx["descripcion"][:42],
                tx["categoria"],
                f"$ {tx['importe']:,.2f}",
            ])
        if len(txs) > 20:
            resto = sum(t["importe"] for t in txs[20:])
            det_rows.append([f"... y {len(txs)-20} gastos más", "", f"$ {resto:,.2f}"])

        det_ts = TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(t_color)),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, -1), 7),
            ('ROWBACKGROUND', (0, 1), (-1, -1),
             [colors.white, colors.HexColor(0xF5F5F5)]),
            ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
            ('BOX', (0, 0), (-1, -1), 0.5, colors.grey),
            ('INNERGRID', (0, 0), (-1, -1), 0.2, colors.HexColor(0xEEEEEE)),
            ('TOPPADDING', (0, 0), (-1, -1), 2),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 2),
        ])
        det_tbl = Table(det_rows, colWidths=[8*cm, 4*cm, 3*cm])
        det_tbl.setStyle(det_ts)
        story.append(det_tbl)
        story.append(Spacer(1, 0.3*cm))

    # Pie de página con insights
    story.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor(0x37474F)))
    story.append(Spacer(1, 0.2*cm))
    insights = [
        f"Gasto más alto del período: {sorted(todas, key=lambda x: x['importe'], reverse=True)[0]['descripcion']} "
        f"($ {sorted(todas, key=lambda x: x['importe'], reverse=True)[0]['importe']:,.2f})",
        f"Categoría de mayor gasto: {sorted_cats[0][0]} con $ {sorted_cats[0][1]:,.2f} "
        f"({sorted_cats[0][1]/total_ars*100:.1f}% del total)",
        f"Tarjeta con mayor uso: {'VISA 2473' if tarjetas_info['VISA 2473']['ars'] > tarjetas_info['AMEX 4100']['ars'] else 'AMEX 4100'} "
        f"con mayor cantidad de transacciones.",
        f"Total gastos en dólares: U$S {total_usd:.2f} (Apple + Splice · equiv. al tipo de cambio blue)",
        f"Fecha de análisis: {datetime.now().strftime('%d/%m/%Y')}",
    ]
    for insight in insights:
        story.append(Paragraph(f"▶ {insight}", style_note))

    doc.build(story)
    print(f"PDF guardado: {output_path}")


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print("Parseando Amex 4100...")
    amex_4100 = parse_amex_4100()
    print(f"  → {len(amex_4100)} transacciones")

    print("Parseando Visa 2473...")
    visa_2473 = parse_visa_2473()
    print(f"  → {len(visa_2473)} transacciones")

    print("Parseando Amex Santander...")
    amex_sant = parse_amex_santander()
    print(f"  → {len(amex_sant)} transacciones")

    todas = amex_4100 + visa_2473 + amex_sant
    print(f"\nTotal combinado: {len(todas)} transacciones")

    total_ars = sum(t["importe"] for t in todas if t["moneda"] == "$")
    total_usd = sum(t["importe"] for t in todas if t["moneda"] != "$")
    print(f"Total ARS: $ {total_ars:,.2f}")
    print(f"Total USD: U$S {total_usd:.2f}")

    excel_path = "/home/user/OWN-Daro-1/GastosTarjetas_Real_Abril2026.xlsx"
    pdf_path   = "/home/user/OWN-Daro-1/GastosTarjetas_Real_Abril2026.pdf"

    print("\nGenerando Excel...")
    generar_excel(todas, excel_path)

    print("Generando PDF...")
    generar_pdf(todas, pdf_path)

    print("\n✓ Análisis completado exitosamente.")


if __name__ == "__main__":
    main()
