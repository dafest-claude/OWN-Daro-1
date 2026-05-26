#!/usr/bin/env python3
"""
Generador de PDF — Análisis de Retorno de Inversiones a 1 Año Vista
Fuente de datos: OperacionesFinalizadas.xls  (período 01/05/2026 – 22/05/2026)
Proyección:       21/05/2026  →  21/05/2027
"""

import os
from collections import defaultdict
from datetime import datetime, date
from bs4 import BeautifulSoup

from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether, PageBreak
)
from reportlab.platypus.flowables import BalancedColumns
from reportlab.graphics.shapes import Drawing, Rect, String, Line
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.graphics.charts.piecharts import Pie
from reportlab.graphics import renderPDF

# ─────────────────────────────────────────────────────────────
# PARÁMETROS
# ─────────────────────────────────────────────────────────────
FECHA_ANALISIS   = date(2026, 5, 21)
FECHA_PROYECCION = date(2027, 5, 21)
INFLACION_ANUAL  = 0.50

PERFIL = {
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
# PALETA DE COLORES
# ─────────────────────────────────────────────────────────────
AZUL_OSC  = colors.HexColor("#1A3C5E")
AZUL_MED  = colors.HexColor("#2E6DA4")
AZUL_CLAR = colors.HexColor("#D6E8F7")
VERDE_OSC = colors.HexColor("#1E5C30")
VERDE_MED = colors.HexColor("#70AD47")
VERDE_CL  = colors.HexColor("#C6EFCE")
AMBAR     = colors.HexColor("#FFC000")
ROJO_CL   = colors.HexColor("#FFC7CE")
AMBAR_CL  = colors.HexColor("#FFEB9C")
GRIS_OSC  = colors.HexColor("#404040")
GRIS_MED  = colors.HexColor("#808080")
GRIS_CL   = colors.HexColor("#F5F5F5")
BLANCO    = colors.white
NEGRO     = colors.black

C_BONO    = colors.HexColor("#D6E8F7")
C_ACCION  = colors.HexColor("#D5F0DC")
C_CEDEAR  = colors.HexColor("#FFF0CC")
C_FCI     = colors.HexColor("#EDD8F5")

TIPO_COLOR_MAP = {
    "Bono Soberano USD":     C_BONO,
    "Bono Tesoro USD":       C_BONO,
    "Bono CER (ajustable)":  C_BONO,
    "Bono BOPREAL USD":      C_BONO,
    "Acción Panel Líder":    C_ACCION,
    "CEDEAR Tecnología":     C_CEDEAR,
    "CEDEAR Energía":        C_CEDEAR,
    "CEDEAR ETF Materiales": C_CEDEAR,
    "FCI Portafolio Mixto":  C_FCI,
}

RIESGO_COLOR = {
    "Bajo":  VERDE_CL,
    "Medio": AMBAR_CL,
    "Alto":  ROJO_CL,
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

        def pnum(s):
            return float(s.replace(".", "").replace(",", ".")) if s else 0.0

        def pfecha(s):
            try:
                return datetime.strptime(s.split()[0], "%d/%m/%Y").date()
            except Exception:
                return None

        ops.append({
            "fecha_tx":    pfecha(celdas[0]),
            "boleto":      celdas[2],
            "descripcion": celdas[6],
            "simbolo":     celdas[8],
            "cantidad":    pnum(celdas[9]),
            "precio":      pnum(celdas[11]),
            "monto":       pnum(celdas[12]),
            "comision":    pnum(celdas[13]),
            "iva":         pnum(celdas[14]),
            "total":       pnum(celdas[15]),
        })
    return ops


def proyectar(op):
    p = PERFIL.get(op["simbolo"], {"rend": 0.30, "tipo": "Otro", "riesgo": "Medio"})
    inv = op["total"]
    rend = p["rend"]
    vproy = inv * (1 + rend)
    gan = vproy - inv
    rreal = (1 + rend) / (1 + INFLACION_ANUAL) - 1
    return {
        **p,
        "simbolo":      op["simbolo"],
        "descripcion":  op["descripcion"],
        "fecha_compra": op["fecha_tx"],
        "cantidad":     op["cantidad"],
        "precio":       op["precio"],
        "invertido":    inv,
        "valor_proy":   round(vproy, 2),
        "ganancia":     round(gan, 2),
        "rend_nom":     rend,
        "rend_real":    rreal,
    }


def ars(v):
    return f"AR$ {v:,.2f}"


def pct(v):
    signo = "+" if v >= 0 else ""
    return f"{signo}{v*100:.2f}%"


# ─────────────────────────────────────────────────────────────
# ESTILOS
# ─────────────────────────────────────────────────────────────
def build_styles():
    base = getSampleStyleSheet()

    titulo = ParagraphStyle(
        "Titulo",
        parent=base["Normal"],
        fontName="Helvetica-Bold",
        fontSize=20,
        textColor=BLANCO,
        alignment=TA_CENTER,
        spaceAfter=4,
    )
    subtitulo = ParagraphStyle(
        "Subtitulo",
        parent=base["Normal"],
        fontName="Helvetica",
        fontSize=10,
        textColor=BLANCO,
        alignment=TA_CENTER,
        spaceAfter=2,
    )
    seccion = ParagraphStyle(
        "Seccion",
        parent=base["Normal"],
        fontName="Helvetica-Bold",
        fontSize=11,
        textColor=BLANCO,
        alignment=TA_LEFT,
        spaceBefore=4,
        spaceAfter=4,
        leftIndent=6,
    )
    normal = ParagraphStyle(
        "NormalCustom",
        parent=base["Normal"],
        fontName="Helvetica",
        fontSize=8.5,
        textColor=GRIS_OSC,
        spaceAfter=2,
    )
    negrita = ParagraphStyle(
        "Negrita",
        parent=normal,
        fontName="Helvetica-Bold",
        fontSize=9,
    )
    nota = ParagraphStyle(
        "Nota",
        parent=base["Normal"],
        fontName="Helvetica-Oblique",
        fontSize=7.5,
        textColor=GRIS_MED,
        alignment=TA_LEFT,
        spaceAfter=4,
    )
    return titulo, subtitulo, seccion, normal, negrita, nota


# ─────────────────────────────────────────────────────────────
# COMPONENTES REUTILIZABLES
# ─────────────────────────────────────────────────────────────
def encabezado_seccion(texto, style_seccion):
    """Retorna una tabla que actúa como encabezado de sección azul."""
    t = Table([[Paragraph(f"&#9632;  {texto}", style_seccion)]], colWidths=["100%"])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), AZUL_MED),
        ("TOPPADDING",    (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 8),
    ]))
    return t


def tabla_kpi(filas_data, col_widths):
    """Tabla de métricas clave (label | valor)."""
    t = Table(filas_data, colWidths=col_widths)
    style = [
        ("FONTNAME",  (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME",  (1, 0), (1, -1), "Helvetica"),
        ("FONTSIZE",  (0, 0), (-1, -1), 9),
        ("BACKGROUND",(0, 0), (0, -1), AZUL_CLAR),
        ("BACKGROUND",(1, 0), (1, -1), GRIS_CL),
        ("ALIGN",     (1, 0), (1, -1), "CENTER"),
        ("ALIGN",     (0, 0), (0, -1), "LEFT"),
        ("GRID",      (0, 0), (-1, -1), 0.4, colors.HexColor("#AAAAAA")),
        ("ROWBACKGROUNDS", (0, 0), (-1, -1), [AZUL_CLAR, colors.HexColor("#EBF3FB")]),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
    ]
    # Resaltar fila de ganancia (índice 2) en verde
    for i, fila in enumerate(filas_data):
        if fila and "Ganancia" in str(fila[0]):
            style.append(("BACKGROUND", (1, i), (1, i), VERDE_CL))
            style.append(("FONTNAME",   (1, i), (1, i), "Helvetica-Bold"))
        if fila and "Real" in str(fila[0]):
            style.append(("BACKGROUND", (1, i), (1, i), VERDE_CL))
    t.setStyle(TableStyle(style))
    return t


# ─────────────────────────────────────────────────────────────
# GRÁFICO DE BARRAS (reportlab nativo)
# ─────────────────────────────────────────────────────────────
def grafico_barras(sim_data, ancho=16*cm, alto=8*cm):
    """sim_data: lista de (label, rend_nom, rend_real)"""
    d = Drawing(ancho, alto)

    labels    = [s[0] for s in sim_data]
    vals_nom  = [s[1] * 100 for s in sim_data]
    vals_real = [s[2] * 100 for s in sim_data]

    bc = VerticalBarChart()
    bc.x = 55
    bc.y = 20
    bc.width  = ancho - 70
    bc.height = alto - 35
    bc.data   = [vals_nom, vals_real]

    bc.bars[0].fillColor = AZUL_MED
    bc.bars[1].fillColor = VERDE_MED

    bc.categoryAxis.categoryNames = labels
    bc.categoryAxis.labels.angle  = 35
    bc.categoryAxis.labels.dx     = -6
    bc.categoryAxis.labels.dy     = -12
    bc.categoryAxis.labels.fontSize = 7

    bc.valueAxis.valueMin  = min(min(vals_real), 0) - 5
    bc.valueAxis.valueMax  = max(vals_nom) + 10
    bc.valueAxis.valueStep = 10
    bc.valueAxis.labels.fontSize = 7
    bc.valueAxis.labelTextFormat = "%.0f%%"

    bc.groupSpacing = 5
    bc.barSpacing   = 1

    d.add(bc)

    # Leyenda
    lx = bc.x + bc.width + 5
    d.add(Rect(lx, alto - 20, 10, 8, fillColor=AZUL_MED, strokeColor=None))
    d.add(String(lx + 13, alto - 20, "Nominal", fontSize=7, fillColor=GRIS_OSC))
    d.add(Rect(lx, alto - 34, 10, 8, fillColor=VERDE_MED, strokeColor=None))
    d.add(String(lx + 13, alto - 34, "Real", fontSize=7, fillColor=GRIS_OSC))

    return d


# ─────────────────────────────────────────────────────────────
# GRÁFICO DE TORTA
# ─────────────────────────────────────────────────────────────
def grafico_torta(tipo_data, ancho=14*cm, alto=8*cm):
    """tipo_data: lista de (label, monto)"""
    d = Drawing(ancho, alto)

    pie = Pie()
    pie.x      = 20
    pie.y      = 15
    pie.width  = alto - 30
    pie.height = alto - 30
    pie.data   = [t[1] for t in tipo_data]
    pie.labels = [t[0].replace(" ", "\n") for t in tipo_data]

    paleta = [AZUL_MED, VERDE_MED, AMBAR, colors.HexColor("#9B59B6"),
              colors.HexColor("#E74C3C"), colors.HexColor("#1ABC9C")]
    for i, c_ in enumerate(paleta[:len(tipo_data)]):
        pie.slices[i].fillColor = c_
        pie.slices[i].strokeColor = BLANCO
        pie.slices[i].strokeWidth = 1

    pie.sideLabels = True
    pie.slices.label_simple_pointer = True
    pie.slices.fontName  = "Helvetica"
    pie.slices.fontSize  = 6.5
    pie.slices.fillColor = colors.HexColor("#303030")

    d.add(pie)
    return d


# ─────────────────────────────────────────────────────────────
# CONSTRUCCIÓN DEL PDF
# ─────────────────────────────────────────────────────────────
def generar_pdf(ops_raw, proyecciones, output_path):
    doc = SimpleDocTemplate(
        output_path,
        pagesize=A4,
        leftMargin=1.8*cm, rightMargin=1.8*cm,
        topMargin=1.5*cm,  bottomMargin=1.8*cm,
        title="Análisis de Retorno de Inversiones — 1 Año Vista",
        author="InverSys Analytics",
    )

    W = A4[0] - 3.6*cm   # ancho útil

    s_titulo, s_sub, s_sec, s_normal, s_neg, s_nota = build_styles()
    story = []

    # ── PORTADA / ENCABEZADO ─────────────────────────────────
    header_data = [[
        Paragraph("ANÁLISIS DE RETORNO DE INVERSIONES", s_titulo),
    ],[
        Paragraph("Proyección a 1 Año Vista  ·  BCBA  ·  Cuenta 99733", s_sub),
    ],[
        Paragraph(
            f"Fecha de análisis: {FECHA_ANALISIS.strftime('%d/%m/%Y')}  "
            f"→  Proyección al: {FECHA_PROYECCION.strftime('%d/%m/%Y')}",
            s_sub),
    ]]
    tbl_header = Table(header_data, colWidths=[W])
    tbl_header.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), AZUL_OSC),
        ("TOPPADDING",    (0, 0), (-1, -1), 10),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ("LEFTPADDING",   (0, 0), (-1, -1), 12),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 12),
    ]))
    story.append(tbl_header)
    story.append(Spacer(1, 0.4*cm))

    # ── CÁLCULOS AGREGADOS ───────────────────────────────────
    total_inv  = sum(p["invertido"]  for p in proyecciones)
    total_proy = sum(p["valor_proy"] for p in proyecciones)
    gan_total  = total_proy - total_inv
    rend_nom   = gan_total / total_inv
    rend_real  = (1 + rend_nom) / (1 + INFLACION_ANUAL) - 1

    por_tipo   = defaultdict(lambda: {"inv": 0.0, "proy": 0.0, "n": 0})
    por_riesgo = defaultdict(lambda: {"inv": 0.0, "proy": 0.0, "n": 0})
    por_sim    = defaultdict(lambda: {"inv": 0.0, "proy": 0.0, "desc": "", "tipo": "", "rend": 0.0})
    for p in proyecciones:
        por_tipo[p["tipo"]]["inv"]  += p["invertido"]
        por_tipo[p["tipo"]]["proy"] += p["valor_proy"]
        por_tipo[p["tipo"]]["n"]    += 1
        por_riesgo[p["riesgo"]]["inv"]  += p["invertido"]
        por_riesgo[p["riesgo"]]["proy"] += p["valor_proy"]
        por_riesgo[p["riesgo"]]["n"]    += 1
        por_sim[p["simbolo"]]["inv"]  += p["invertido"]
        por_sim[p["simbolo"]]["proy"] += p["valor_proy"]
        por_sim[p["simbolo"]]["desc"]  = p["descripcion"][:40]
        por_sim[p["simbolo"]]["tipo"]  = p["tipo"]
        por_sim[p["simbolo"]]["rend"]  = p["rend_nom"]

    # ── 1. MÉTRICAS GLOBALES ─────────────────────────────────
    story.append(encabezado_seccion("MÉTRICAS GLOBALES DEL PORTAFOLIO", s_sec))
    story.append(Spacer(1, 0.15*cm))

    kpi_rows = [
        ["Capital total invertido",   ars(total_inv)],
        ["Operaciones analizadas",     f"{len(ops_raw)}  (15 boletos BCBA)"],
        ["Valor proyectado a 1 año",  ars(total_proy)],
        ["Ganancia bruta esperada",   ars(gan_total)],
        ["Rendimiento nominal anual", pct(rend_nom)],
        ["Inflación proyectada anual", pct(INFLACION_ANUAL)],
        ["Rendimiento REAL anual",    pct(rend_real)],
        ["Período de análisis",
         f"{FECHA_ANALISIS.strftime('%d/%m/%Y')} → {FECHA_PROYECCION.strftime('%d/%m/%Y')}"],
    ]
    story.append(tabla_kpi(kpi_rows, [9*cm, W - 9*cm]))
    story.append(Spacer(1, 0.4*cm))

    # ── 2. DETALLE DE OPERACIONES ────────────────────────────
    story.append(encabezado_seccion("DETALLE DE OPERACIONES FINALIZADAS (01/05 – 22/05/2026)", s_sec))
    story.append(Spacer(1, 0.15*cm))

    col_w_ops = [1.1*cm, 2.3*cm, 6.5*cm, 2.2*cm, 2.4*cm, 2.5*cm]
    ops_header = [
        Paragraph("<b>Fecha</b>",      s_normal),
        Paragraph("<b>Símbolo</b>",    s_normal),
        Paragraph("<b>Descripción</b>",s_normal),
        Paragraph("<b>Cantidad</b>",   s_normal),
        Paragraph("<b>Precio AR$</b>", s_normal),
        Paragraph("<b>Total AR$</b>",  s_normal),
    ]
    ops_rows = [ops_header]
    for op in ops_raw:
        tipo_op = PERFIL.get(op["simbolo"], {}).get("tipo", "Otro")
        ops_rows.append([
            Paragraph(op["fecha_tx"].strftime("%d/%m") if op["fecha_tx"] else "", s_normal),
            Paragraph(f"<b>{op['simbolo']}</b>", s_normal),
            Paragraph(op["descripcion"][:48], s_normal),
            Paragraph(f"{op['cantidad']:,.0f}", s_normal),
            Paragraph(f"{op['precio']:,.2f}", s_normal),
            Paragraph(f"<b>{op['total']:,.2f}</b>", s_normal),
        ])

    # Fila de totales
    ops_rows.append([
        Paragraph("", s_normal),
        Paragraph("", s_normal),
        Paragraph("<b>TOTAL INVERTIDO</b>", s_neg),
        Paragraph("", s_normal),
        Paragraph("", s_normal),
        Paragraph(f"<b>{total_inv:,.2f}</b>", s_neg),
    ])

    tbl_ops = Table(ops_rows, colWidths=col_w_ops, repeatRows=1)
    ops_style = [
        ("BACKGROUND",    (0, 0), (-1, 0),  AZUL_OSC),
        ("TEXTCOLOR",     (0, 0), (-1, 0),  BLANCO),
        ("FONTNAME",      (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 7.5),
        ("ALIGN",         (3, 0), (-1, -1), "RIGHT"),
        ("ALIGN",         (0, 0), (2, -1),  "LEFT"),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#BBBBBB")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        # Fila de totales
        ("BACKGROUND",    (0, -1), (-1, -1), AZUL_OSC),
        ("TEXTCOLOR",     (0, -1), (-1, -1), BLANCO),
        ("FONTNAME",      (0, -1), (-1, -1), "Helvetica-Bold"),
        ("LINEABOVE",     (0, -1), (-1, -1), 1.0, AZUL_MED),
    ]
    # Colores por categoría en filas de datos
    for i, op in enumerate(ops_raw, start=1):
        tipo_op = PERFIL.get(op["simbolo"], {}).get("tipo", "Otro")
        bg = TIPO_COLOR_MAP.get(tipo_op, GRIS_CL)
        if i % 2 == 0:
            # ligeramente más oscuro en filas pares
            ops_style.append(("BACKGROUND", (0, i), (-1, i), bg))
        else:
            ops_style.append(("BACKGROUND", (0, i), (-1, i), bg))

    tbl_ops.setStyle(TableStyle(ops_style))
    story.append(tbl_ops)
    story.append(Spacer(1, 0.15*cm))

    # Leyenda
    legend_items = [
        ("Renta Fija / Bonos", C_BONO),
        ("Acciones locales",   C_ACCION),
        ("CEDEARs",            C_CEDEAR),
        ("FCI",                C_FCI),
    ]
    legend_data = [[
        Table([[Paragraph(f"  {txt}  ", s_normal)]], colWidths=[3.2*cm],
              style=[("BACKGROUND", (0, 0), (-1, -1), clr),
                     ("GRID", (0, 0), (-1, -1), 0.3, GRIS_MED)])
        for txt, clr in legend_items
    ]]
    tbl_ley = Table(legend_data, colWidths=[W])
    story.append(tbl_ley)
    story.append(Spacer(1, 0.4*cm))

    # ── 3. PROYECCIÓN A 1 AÑO ───────────────────────────────
    story.append(encabezado_seccion("PROYECCIÓN DE RENDIMIENTO A 1 AÑO (21/05/2026 → 21/05/2027)", s_sec))
    story.append(Spacer(1, 0.15*cm))

    # Supuestos
    sup_data = [
        ["Inflación anual proyectada", f"{INFLACION_ANUAL*100:.0f}%",
         "Fuente: estimación CER / REM-BCRA"],
    ]
    tbl_sup = Table(sup_data, colWidths=[5.5*cm, 2*cm, W - 7.5*cm])
    tbl_sup.setStyle(TableStyle([
        ("BACKGROUND",    (0, 0), (-1, -1), colors.HexColor("#FFFDE7")),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("FONTNAME",      (0, 0), (0, -1),  "Helvetica-Bold"),
        ("FONTNAME",      (1, 0), (1, -1),  "Helvetica-Bold"),
        ("TEXTCOLOR",     (1, 0), (1, -1),  GRIS_OSC),
        ("GRID",          (0, 0), (-1, -1), 0.3, AMBAR),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 6),
        ("ALIGN",         (1, 0), (1, -1), "CENTER"),
    ]))
    story.append(tbl_sup)
    story.append(Spacer(1, 0.2*cm))

    col_w_proy = [1.5*cm, 5.5*cm, 2.5*cm, 2.5*cm, 2.0*cm, 1.8*cm, 1.7*cm]
    proy_header = [
        Paragraph("<b>Símbolo</b>",       s_normal),
        Paragraph("<b>Categoría</b>",     s_normal),
        Paragraph("<b>Invertido AR$</b>", s_normal),
        Paragraph("<b>Proy. AR$</b>",     s_normal),
        Paragraph("<b>Ganancia AR$</b>",  s_normal),
        Paragraph("<b>Nom. %</b>",        s_normal),
        Paragraph("<b>Real %</b>",        s_normal),
    ]
    proy_rows = [proy_header]

    for p in proyecciones:
        bg_tipo  = TIPO_COLOR_MAP.get(p["tipo"], GRIS_CL)
        bg_rsg   = RIESGO_COLOR.get(p["riesgo"], GRIS_CL)
        proy_rows.append([
            Paragraph(f"<b>{p['simbolo']}</b>",  s_normal),
            Paragraph(p["tipo"],                  s_normal),
            Paragraph(f"{p['invertido']:,.2f}",   s_normal),
            Paragraph(f"{p['valor_proy']:,.2f}",  s_normal),
            Paragraph(f"<b>{p['ganancia']:,.2f}</b>", s_normal),
            Paragraph(pct(p["rend_nom"]),         s_normal),
            Paragraph(pct(p["rend_real"]),        s_normal),
        ])

    # Fila totales
    proy_rows.append([
        Paragraph("<b>TOTAL</b>",           s_neg),
        Paragraph("",                        s_neg),
        Paragraph(f"<b>{total_inv:,.2f}</b>", s_neg),
        Paragraph(f"<b>{total_proy:,.2f}</b>", s_neg),
        Paragraph(f"<b>{gan_total:,.2f}</b>",  s_neg),
        Paragraph(f"<b>{pct(rend_nom)}</b>",   s_neg),
        Paragraph(f"<b>{pct(rend_real)}</b>",  s_neg),
    ])

    tbl_proy = Table(proy_rows, colWidths=col_w_proy, repeatRows=1)
    proy_style = [
        ("BACKGROUND",    (0, 0), (-1, 0), AZUL_OSC),
        ("TEXTCOLOR",     (0, 0), (-1, 0), BLANCO),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 7.5),
        ("ALIGN",         (2, 0), (-1, -1), "RIGHT"),
        ("ALIGN",         (0, 0), (1, -1),  "LEFT"),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#BBBBBB")),
        ("TOPPADDING",    (0, 0), (-1, -1), 3),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING",   (0, 0), (-1, -1), 4),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 4),
        ("BACKGROUND",    (0, -1), (-1, -1), AZUL_OSC),
        ("TEXTCOLOR",     (0, -1), (-1, -1), BLANCO),
        ("LINEABOVE",     (0, -1), (-1, -1), 1.0, AZUL_MED),
    ]
    for i, p in enumerate(proyecciones, start=1):
        bg = TIPO_COLOR_MAP.get(p["tipo"], GRIS_CL)
        proy_style.append(("BACKGROUND", (0, i), (-1, i), bg))
        # Real% en verde si positivo, rojo si negativo
        real_bg = VERDE_CL if p["rend_real"] >= 0 else ROJO_CL
        proy_style.append(("BACKGROUND", (-1, i), (-1, i), real_bg))

    tbl_proy.setStyle(TableStyle(proy_style))
    story.append(tbl_proy)
    story.append(Spacer(1, 0.4*cm))

    # ── 4. RESUMEN POR CATEGORÍA ─────────────────────────────
    story.append(encabezado_seccion("DESGLOSE POR CATEGORÍA DE INSTRUMENTO", s_sec))
    story.append(Spacer(1, 0.15*cm))

    cat_header = [
        Paragraph("<b>Categoría</b>",        s_normal),
        Paragraph("<b>Ops.</b>",             s_normal),
        Paragraph("<b>Invertido AR$</b>",    s_normal),
        Paragraph("<b>% Cartera</b>",        s_normal),
        Paragraph("<b>Proyectado AR$</b>",   s_normal),
        Paragraph("<b>Rendimiento %</b>",    s_normal),
    ]
    cat_rows = [cat_header]
    for tipo_k, datos in sorted(por_tipo.items(), key=lambda x: -x[1]["inv"]):
        rend_t = (datos["proy"] - datos["inv"]) / datos["inv"]
        pct_c  = datos["inv"] / total_inv
        bg = TIPO_COLOR_MAP.get(tipo_k, GRIS_CL)
        cat_rows.append([
            Paragraph(tipo_k,                     s_normal),
            Paragraph(str(datos["n"]),            s_normal),
            Paragraph(f"{datos['inv']:,.2f}",     s_normal),
            Paragraph(f"{pct_c*100:.1f}%",        s_normal),
            Paragraph(f"{datos['proy']:,.2f}",    s_normal),
            Paragraph(f"<b>{pct(rend_t)}</b>",   s_normal),
        ])
    # Total
    cat_rows.append([
        Paragraph("<b>TOTAL</b>", s_neg),
        Paragraph(f"<b>{len(proyecciones)}</b>", s_neg),
        Paragraph(f"<b>{total_inv:,.2f}</b>",   s_neg),
        Paragraph("<b>100.0%</b>",               s_neg),
        Paragraph(f"<b>{total_proy:,.2f}</b>",  s_neg),
        Paragraph(f"<b>{pct(rend_nom)}</b>",    s_neg),
    ])

    col_w_cat = [5.2*cm, 1.2*cm, 3.0*cm, 2.0*cm, 3.0*cm, 2.6*cm]
    tbl_cat = Table(cat_rows, colWidths=col_w_cat, repeatRows=1)
    cat_style = [
        ("BACKGROUND",    (0, 0), (-1, 0), AZUL_OSC),
        ("TEXTCOLOR",     (0, 0), (-1, 0), BLANCO),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN",         (0, 0), (0, -1),  "LEFT"),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#BBBBBB")),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
        ("BACKGROUND",    (0, -1), (-1, -1), AZUL_OSC),
        ("TEXTCOLOR",     (0, -1), (-1, -1), BLANCO),
        ("LINEABOVE",     (0, -1), (-1, -1), 1.0, AZUL_MED),
    ]
    for i, tipo_k in enumerate(sorted(por_tipo.keys(), key=lambda x: -por_tipo[x]["inv"]), start=1):
        bg = TIPO_COLOR_MAP.get(tipo_k, GRIS_CL)
        cat_style.append(("BACKGROUND", (0, i), (-1, i), bg))
    tbl_cat.setStyle(TableStyle(cat_style))
    story.append(tbl_cat)
    story.append(Spacer(1, 0.35*cm))

    # ── 5. DESGLOSE POR RIESGO ───────────────────────────────
    story.append(encabezado_seccion("DESGLOSE POR NIVEL DE RIESGO", s_sec))
    story.append(Spacer(1, 0.15*cm))

    rsg_header = [
        Paragraph("<b>Riesgo</b>",           s_normal),
        Paragraph("<b>Ops.</b>",             s_normal),
        Paragraph("<b>Invertido AR$</b>",    s_normal),
        Paragraph("<b>% Cartera</b>",        s_normal),
        Paragraph("<b>Proyectado AR$</b>",   s_normal),
        Paragraph("<b>Rendimiento %</b>",    s_normal),
    ]
    rsg_rows = [rsg_header]
    for riesgo_k in ["Bajo", "Medio", "Alto"]:
        if riesgo_k not in por_riesgo:
            continue
        datos  = por_riesgo[riesgo_k]
        rend_r = (datos["proy"] - datos["inv"]) / datos["inv"]
        pct_r  = datos["inv"] / total_inv
        bg     = RIESGO_COLOR.get(riesgo_k, GRIS_CL)
        rsg_rows.append([
            Paragraph(f"<b>{riesgo_k}</b>",         s_normal),
            Paragraph(str(datos["n"]),               s_normal),
            Paragraph(f"{datos['inv']:,.2f}",        s_normal),
            Paragraph(f"{pct_r*100:.1f}%",           s_normal),
            Paragraph(f"{datos['proy']:,.2f}",       s_normal),
            Paragraph(f"<b>{pct(rend_r)}</b>",      s_normal),
        ])

    col_w_rsg = [2.5*cm, 1.2*cm, 3.5*cm, 2.0*cm, 3.5*cm, 4.3*cm]
    tbl_rsg = Table(rsg_rows, colWidths=col_w_rsg, repeatRows=1)
    rsg_style = [
        ("BACKGROUND",    (0, 0), (-1, 0), AZUL_OSC),
        ("TEXTCOLOR",     (0, 0), (-1, 0), BLANCO),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (1, 0), (-1, -1), "RIGHT"),
        ("ALIGN",         (0, 0), (0, -1),  "LEFT"),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#BBBBBB")),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
    ]
    for i, riesgo_k in enumerate(["Bajo", "Medio", "Alto"], start=1):
        if riesgo_k in por_riesgo:
            rsg_style.append(("BACKGROUND", (0, i), (-1, i), RIESGO_COLOR[riesgo_k]))
    tbl_rsg.setStyle(TableStyle(rsg_style))
    story.append(tbl_rsg)
    story.append(Spacer(1, 0.35*cm))

    # ── 6. RANKING ───────────────────────────────────────────
    story.append(encabezado_seccion("RANKING DE RENDIMIENTO ESPERADO (mayor a menor)", s_sec))
    story.append(Spacer(1, 0.15*cm))

    ranking = sorted(por_sim.items(), key=lambda x: -x[1]["rend"])
    rk_header = [
        Paragraph("<b>#</b>",              s_normal),
        Paragraph("<b>Símbolo</b>",        s_normal),
        Paragraph("<b>Categoría</b>",      s_normal),
        Paragraph("<b>Invertido AR$</b>",  s_normal),
        Paragraph("<b>Nom. %</b>",         s_normal),
        Paragraph("<b>Real %</b>",         s_normal),
    ]
    rk_rows = [rk_header]
    for rank_i, (sim_k, datos) in enumerate(ranking, start=1):
        rend_r = (1 + datos["rend"]) / (1 + INFLACION_ANUAL) - 1
        bg = TIPO_COLOR_MAP.get(datos["tipo"], GRIS_CL)
        bg_r = VERDE_CL if rend_r >= 0 else ROJO_CL
        rk_rows.append([
            Paragraph(f"<b>{rank_i}</b>",       s_normal),
            Paragraph(f"<b>{sim_k}</b>",        s_normal),
            Paragraph(datos["tipo"],             s_normal),
            Paragraph(f"{datos['inv']:,.2f}",   s_normal),
            Paragraph(f"<b>{pct(datos['rend'])}</b>", s_normal),
            Paragraph(pct(rend_r),               s_normal),
        ])

    col_w_rk = [1.0*cm, 1.8*cm, 5.5*cm, 3.2*cm, 2.2*cm, 3.3*cm]
    tbl_rk = Table(rk_rows, colWidths=col_w_rk, repeatRows=1)
    rk_style = [
        ("BACKGROUND",    (0, 0), (-1, 0), AZUL_OSC),
        ("TEXTCOLOR",     (0, 0), (-1, 0), BLANCO),
        ("FONTNAME",      (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE",      (0, 0), (-1, -1), 8),
        ("ALIGN",         (3, 0), (-1, -1), "RIGHT"),
        ("ALIGN",         (0, 0), (2, -1),  "LEFT"),
        ("GRID",          (0, 0), (-1, -1), 0.3, colors.HexColor("#BBBBBB")),
        ("TOPPADDING",    (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING",   (0, 0), (-1, -1), 5),
        ("RIGHTPADDING",  (0, 0), (-1, -1), 5),
    ]
    for i, (sim_k, datos) in enumerate(ranking, start=1):
        bg = TIPO_COLOR_MAP.get(datos["tipo"], GRIS_CL)
        rend_r = (1 + datos["rend"]) / (1 + INFLACION_ANUAL) - 1
        rk_style.append(("BACKGROUND", (0, i), (-1, i), bg))
        bg_r = VERDE_CL if rend_r >= 0 else ROJO_CL
        rk_style.append(("BACKGROUND", (-1, i), (-1, i), bg_r))
    tbl_rk.setStyle(TableStyle(rk_style))
    story.append(tbl_rk)

    # ── 7. GRÁFICOS (nueva página) ───────────────────────────
    story.append(PageBreak())
    story.append(encabezado_seccion("ANÁLISIS VISUAL", s_sec))
    story.append(Spacer(1, 0.3*cm))

    # Gráfico de barras — rendimiento por símbolo
    story.append(Paragraph(
        "<b>Rendimiento Nominal vs. Real esperado por instrumento (%)</b>", s_neg))
    story.append(Spacer(1, 0.15*cm))

    sim_chart_data = sorted(
        [(k, d["rend"], (1+d["rend"])/(1+INFLACION_ANUAL)-1)
         for k, d in por_sim.items()],
        key=lambda x: -x[1]
    )
    story.append(grafico_barras(sim_chart_data, ancho=W, alto=9*cm))
    story.append(Spacer(1, 0.4*cm))

    # Gráfico de torta — composición por categoría
    story.append(Paragraph(
        "<b>Composición del portafolio por categoría (capital invertido)</b>", s_neg))
    story.append(Spacer(1, 0.15*cm))

    tipo_chart = sorted(por_tipo.items(), key=lambda x: -x[1]["inv"])
    story.append(grafico_torta(
        [(k, d["inv"]) for k, d in tipo_chart],
        ancho=W, alto=9*cm
    ))
    story.append(Spacer(1, 0.4*cm))

    # ── 8. NOTA AL PIE ──────────────────────────────────────
    story.append(HRFlowable(width=W, thickness=0.5, color=GRIS_MED))
    story.append(Spacer(1, 0.1*cm))
    story.append(Paragraph(
        "⚠  AVISO LEGAL: Las proyecciones presentadas en este informe son estimaciones basadas en "
        "análisis fundamental y técnico de cada instrumento financiero. No constituyen asesoramiento "
        "financiero ni garantía de retorno. Los rendimientos pasados no garantizan resultados futuros. "
        "Los precios y tasas utilizados corresponden al cierre de mercado del 21/05/2026.",
        s_nota,
    ))
    story.append(Paragraph(
        f"Informe generado el {FECHA_ANALISIS.strftime('%d/%m/%Y')}  —  Cuenta BCBA: 99733  —  "
        "InverSys Analytics",
        s_nota,
    ))

    doc.build(story)
    print(f"PDF generado: {output_path}")


# ─────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────
def main():
    base   = os.path.dirname(os.path.abspath(__file__))
    root   = os.path.dirname(base)
    xls_in = os.path.join(root, "datos_entradas", "OperacionesFinalizadas.xls")
    pdf_out = os.path.join(base, "RetornoInversiones_1Anio.pdf")

    ops_raw      = leer_operaciones(xls_in)
    proyecciones = [proyectar(op) for op in ops_raw]
    print(f"Operaciones leídas: {len(ops_raw)}")
    generar_pdf(ops_raw, proyecciones, pdf_out)


if __name__ == "__main__":
    main()
