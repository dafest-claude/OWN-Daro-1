"""
Genera Cronograma_Seguimiento_ABB.pdf usando reportlab.
Provisión ABB — CPF-2 / La Calera II (Vaca Muerta)
"""
import os
from datetime import date, timedelta
from reportlab.lib.pagesizes import A3, landscape
from reportlab.lib.units import cm, mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.lib import colors
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate
from reportlab.lib.colors import HexColor, white, black

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT_PDF   = os.path.join(PLANNING_DIR, "Cronograma_Seguimiento_ABB.pdf")

# ── Colores ───────────────────────────────────────────────────────────────────
R  = HexColor("#CC0000")   # rojo ABB
RD = HexColor("#880000")   # rojo oscuro
BL = HexColor("#2E75B6")   # azul
BD = HexColor("#1F3864")   # azul oscuro
FA = HexColor("#E36C09")   # naranja FAT
DE = HexColor("#7030A0")   # violeta despacho
GR = HexColor("#375623")   # verde integración
YL = HexColor("#FFEB9C")   # amarillo warn
YD = HexColor("#9C5700")   # amarillo oscuro
OK = HexColor("#C6EFCE")   # verde OK
OD = HexColor("#375623")   # verde oscuro texto
RK = HexColor("#FFC7CE")   # rojo riesgo
RKD= HexColor("#9C0006")   # rojo riesgo texto
PP = HexColor("#1F4E79")   # azul pago
S3L= HexColor("#FFE0E0")   # rosa S3
S4L= HexColor("#DEEAF1")   # celeste S4
PML= HexColor("#FFF8DC")   # crema PMS
GRY= HexColor("#F2F2F2")   # gris claro
DGY= HexColor("#595959")   # gris oscuro
WH = white

PAGE = landscape(A3)

MESES = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

def style(name="Normal", sz=9, color=black, bold=False, align=TA_LEFT,
          leading=None, space_before=0, space_after=0):
    return ParagraphStyle(
        name, fontName="Helvetica-Bold" if bold else "Helvetica",
        fontSize=sz, textColor=color,
        alignment=align, leading=leading or (sz + 3),
        spaceBefore=space_before, spaceAfter=space_after
    )

ST = {
    "h1":    style("h1",    sz=18, color=WH,  bold=True,  align=TA_CENTER),
    "h2":    style("h2",    sz=13, color=WH,  bold=True,  align=TA_CENTER),
    "h3":    style("h3",    sz=11, color=BD,  bold=True),
    "body":  style("body",  sz=9,  color=black),
    "small": style("small", sz=8,  color=DGY),
    "th":    style("th",    sz=8,  color=WH,  bold=True,  align=TA_CENTER),
    "td":    style("td",    sz=8,  color=black),
    "tdc":   style("tdc",   sz=8,  color=black, align=TA_CENTER),
    "warn":  style("warn",  sz=8,  color=YD,  bold=True),
    "risk":  style("risk",  sz=8,  color=RKD, bold=True),
    "ok":    style("ok",    sz=8,  color=OD,  bold=True),
    "crit":  style("crit",  sz=8,  color=RD,  bold=True),
    "note":  style("note",  sz=7,  color=DGY, align=TA_LEFT),
}

def P(text, st="body"): return Paragraph(text, ST[st])
def PB(text):           return Paragraph(f"<b>{text}</b>", ST["body"])

def header_band(story, titulo, subtitulo=""):
    story.append(Spacer(1, 4*mm))
    data = [[Paragraph(titulo,    ST["h2"])],
            [Paragraph(subtitulo, style("sub", sz=9, color=HexColor("#BDD7EE"),
                                        align=TA_CENTER))]] if subtitulo else [[Paragraph(titulo, ST["h2"])]]
    t = Table(data, colWidths=[27*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,-1), BD),
        ("TOPPADDING",    (0,0), (-1,-1), 8),
        ("BOTTOMPADDING", (0,0), (-1,-1), 8),
        ("LEFTPADDING",   (0,0), (-1,-1), 12),
        ("RIGHTPADDING",  (0,0), (-1,-1), 12),
    ]))
    story.append(t)
    story.append(Spacer(1, 3*mm))

def th(txt): return Paragraph(txt, ST["th"])
def td(txt, bold=False, color=black, sz=8, align=TA_LEFT):
    s = ParagraphStyle("x", fontName="Helvetica-Bold" if bold else "Helvetica",
                       fontSize=sz, textColor=color, alignment=align, leading=sz+2)
    return Paragraph(txt, s)

# ─────────────────────────────────────────────────────────────────────────────
# DATOS
# ─────────────────────────────────────────────────────────────────────────────
ACTIVIDADES = [
    # (grupo, id, nombre, ini, fin, tipo, critica, oc, holgura_p6)
    # CONTRATO
    ("CONTRATO", "HC-01", "Aceptación OC — Inicio oficial ABB",             date(2026,5,12), date(2026,5,12), "hito", True,  "S3+S4+PMS", None),
    ("CONTRATO", "HC-02", "Envío cronograma preliminar",                     date(2026,5,12), date(2026,5,12), "hito", True,  "S3+S4+PMS", None),
    ("CONTRATO", "HC-03", "KOM Salas Eléctricas #3 y #4",                   date(2026,5,18), date(2026,5,18), "hito", True,  "S3+S4",     None),
    ("CONTRATO", "HC-04", "KOM PMS",                                         date(2026,5,22), date(2026,5,22), "hito", True,  "PMS",       None),
    ("CONTRATO", "HC-05", "◆ Cronograma final ABB con Hold Points",          date(2026,6, 1), date(2026,6, 1), "fp",   True,  "S3+S4+PMS", None),
    ("CONTRATO", "HC-06", "Primer Informe de Avance",                        date(2026,6, 5), date(2026,6, 5), "hito", True,  "S3+S4+PMS", None),
    # PMS
    ("PMS",      "P-01",  "Emisión / APR Ingeniería Básica PMS",             date(2026,5,22), date(2026,6,19), "act",  False, "PMS",       None),
    ("PMS",      "FP-P1", "◆ FP Ing. Básica PMS",                           date(2026,6,19), date(2026,6,19), "fp",   False, "PMS",       None),
    ("PMS",      "P-02",  "Emisión / APR Ingeniería Detalle PMS",            date(2026,6,22), date(2026,7,31), "act",  False, "PMS",       None),
    ("PMS",      "FP-P2", "◆ FP Ing. Detalle PMS",                          date(2026,7,31), date(2026,7,31), "fp",   False, "PMS",       None),
    ("PMS",      "P-03",  "Fabricación Tablero PMS (BsAs)",                  date(2026,8, 3), date(2027,1,29), "act",  False, "PMS",       None),
    ("PMS",      "P-04",  "FAT Maqueta PMS",                                 date(2026,9, 1), date(2026,9,30), "fat",  False, "PMS",       None),
    ("PMS",      "P-05",  "FAT Tablero PMS (integrado)",                     date(2026,10,19),date(2026,10,23),"fat",  False, "PMS",       222),
    ("PMS",      "P-06",  "FAT Final Tableros PMS",                          date(2027,1, 4), date(2027,1,29), "fat",  False, "PMS",       None),
    # Sala #4
    ("Sala #4",  "S4-01", "Emisión / APR IB Sala #4",                        date(2026,5,22), date(2026,6,26), "act",  False, "S4",        None),
    ("Sala #4",  "FP-S4-1","◆ FP Ing. Básica S#4 → OC Celdas MT (Turquía)", date(2026,6,26), date(2026,6,26), "fp",   False, "S4",        None),
    ("Sala #4",  "S4-02", "Emisión / APR ID Sala #4",                        date(2026,6,29), date(2026,8,14), "act",  False, "S4",        None),
    ("Sala #4",  "FP-S4-2","◆ FP Ing. Detalle S#4",                         date(2026,8,14), date(2026,8,14), "fp",   False, "S4",        None),
    ("Sala #4",  "S4-03", "Emisión / APR IC Sala #4",                        date(2026,8,10), date(2026,8,28), "act",  False, "S4",        None),
    ("Sala #4",  "FP-S4-3","◆ FP Ing. Constructiva S#4 ← Cierra Ing S4",    date(2026,8,28), date(2026,8,28), "fp",   False, "S4",        None),
    ("Sala #4",  "S4-04", "Fabricación Tableros S#4 (Brasil) + Ensayos",     date(2026,9, 1), date(2026,11,23),"fat",  False, "S4",        None),
    ("Sala #4",  "S4-05", "Acopio materiales S#4 en Argentina",              date(2026,12, 7),date(2027,1,15), "act",  False, "S4",        None),
    ("Sala #4",  "S4-06", "Montaje Shelter S#4 (Bottino — Mendoza)",         date(2027,1,18), date(2027,3, 5), "act",  False, "S4",        None),
    ("Sala #4",  "S4-07", "FAT Sala Eléctrica #4",                           date(2027,3, 8), date(2027,3,16), "fat",  False, "S4",        None),
    ("Sala #4",  "S4-08", "▶ DESPACHO Sala #4",                             date(2027,4, 2), date(2027,4, 8), "des",  False, "S4",        81),
    ("Sala #4",  "S4-★",  "★ ENTREGA S#4 — OC 4508944973",                  date(2027,4, 8), date(2027,4, 8), "hito", False, "S4",        81),
    # Sala #3 — Cadena Crítica
    ("Sala #3 ★","S3-01", "Emisión / APR IB Sala #3",                        date(2026,5,22), date(2026,6,26), "act",  True,  "S3",        None),
    ("Sala #3 ★","S3-WARN","⚠ OPEN: Confirmar CCM7 (3×22kW→2×45kW+VFD)",   date(2026,5,26), date(2026,6, 5), "warn", True,  "S3",        None),
    ("Sala #3 ★","FP-S3-1","◆ FP Ing. Básica S#3 → OC Ductos Turquía+VFD", date(2026,7, 3), date(2026,7, 3), "fp",   True,  "S3",        None),
    ("Sala #3 ★","S3-02", "Emisión / APR ID Sala #3",                        date(2026,6,29), date(2026,8,14), "act",  True,  "S3",        None),
    ("Sala #3 ★","FP-S3-2","◆ FP Ing. Detalle S#3 → Fabricación 37 CCMs",   date(2026,8,14), date(2026,8,14), "fp",   True,  "S3",        None),
    ("Sala #3 ★","S3-03", "Emisión / APR IC Sala #3",                        date(2026,8,17), date(2026,9,12), "act",  True,  "S3",        None),
    ("Sala #3 ★","FP-S3-3","◆◆ FP Ing. Constructiva S#3 ← ÚLTIMO FP ABB",  date(2026,8,31), date(2026,8,31), "fp",   True,  "S3",        None),
    ("Sala #3 ★","S3-04", "Fabricación 37 CCMs (Brasil) + Ensayos",          date(2026,9,14), date(2026,11,23),"fat",  True,  "S3",        None),
    ("Sala #3 ★","S3-05", "Acopio materiales S#3 en Argentina",              date(2026,12, 7),date(2027,1,29), "act",  True,  "S3",        None),
    ("Sala #3 ★","S3-06", "Montaje Shelter S#3 (Bottino — Mendoza)",         date(2027,2, 1), date(2027,3,26), "act",  True,  "S3",        None),
    ("Sala #3 ★","S3-07", "FAT Sala Eléctrica #3  ← MÁS TARDÍO",           date(2027,4, 2), date(2027,4,19), "fat",  True,  "S3",        35),
    ("Sala #3 ★","S3-08", "▶ DESPACHO Sala #3  ← MÁS TARDÍO",             date(2027,5,10), date(2027,5,14), "des",  True,  "S3",        45),
    ("Sala #3 ★","S3-★",  "★ ENTREGA S#3 — OC 4508944971",                  date(2027,5,14), date(2027,5,14), "hito", True,  "S3",        45),
    # Pagos
    ("PAGOS","PAY-01","§ 10% Envío Ingeniería Básica",          date(2026,6,12),date(2026,6,12),"pago",False,"S3+S4+PMS",None),
    ("PAGOS","PAY-02","§ 10% Aprobación Ingeniería Básica",     date(2026,7, 3),date(2026,7, 3),"pago",False,"S3+S4+PMS",None),
    ("PAGOS","PAY-03","§ 20% FAT completado (S#3 más tardío)",  date(2027,4,19),date(2027,4,19),"pago",True, "S3",        None),
    ("PAGOS","PAY-04","§ 25% Acopio Materiales en Argentina",   date(2027,1,29),date(2027,1,29),"pago",False,"S3+S4+PMS",None),
    ("PAGOS","PAY-05","§ 30% Contra Entrega (despacho S#3)",    date(2027,5,14),date(2027,5,14),"pago",True, "S3",        None),
    ("PAGOS","PAY-06","§ 5% Databook aprobado",                 date(2027,7, 1),date(2027,7, 1),"pago",False,"S3+S4+PMS",None),
    # Integración
    ("EPC","INT-01","Inicio Precomisionado",                     date(2027,5, 3),date(2028,1,13),"int", False,"EPC",        None),
    ("EPC","INT-02","iFAT Integrado (Shelters+PMS+INAUCO+HIMA)",date(2027,6,25),date(2027,7,14),"fat", True, "S3+S4+PMS",3),
    ("EPC","INT-03","Montaje Shelters en campo",                 date(2027,6,28),date(2027,8,21),"int", False,"EPC",        None),
    ("EPC","INT-04","Inicio Comisionado",                        date(2027,7,16),date(2028,1,31),"int", True, "EPC",        None),
    ("EPC","INT-05","★ RFSU — Ready For Start Up",               date(2028,1,31),date(2028,1,31),"hito",True, "PROYECTO",   None),
]

def fmt(d): return d.strftime("%d/%m/%y") if d else "—"

# ─────────────────────────────────────────────────────────────────────────────
# FUNCIONES DE SECCIÓN
# ─────────────────────────────────────────────────────────────────────────────
def portada(story):
    story.append(Spacer(1, 3*cm))
    data = [
        [Paragraph("PROVISIÓN ABB — CPF-2 / LA CALERA II", style("t1",sz=22,color=WH,bold=True,align=TA_CENTER))],
        [Paragraph("VACA MUERTA — PLUSPETROL / AESA",       style("t2",sz=14,color=HexColor("#BDD7EE"),bold=False,align=TA_CENTER))],
        [Paragraph(" ",                                       style("sp",sz=6))],
        [Paragraph("CRONOGRAMA MAESTRO DE SEGUIMIENTO",      style("t3",sz=18,color=WH,bold=True,align=TA_CENTER))],
        [Paragraph("Hitos Críticos · Freezing Points · Camino Crítico · Alertas", style("t4",sz=11,color=HexColor("#BDD7EE"),align=TA_CENTER))],
    ]
    t = Table(data, colWidths=[27*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), R),
        ("TOPPADDING",    (0,0), (-1,-1), 10),
        ("BOTTOMPADDING", (0,0), (-1,-1), 10),
        ("LEFTPADDING",   (0,0), (-1,-1), 20),
        ("RIGHTPADDING",  (0,0), (-1,-1), 20),
        ("ROWBACKGROUNDS",(0,0), (-1,-1), [R]),
    ]))
    story.append(t)
    story.append(Spacer(1, 1*cm))

    # Info proyecto
    info = [
        [th("PARÁMETRO"),           th("SALAS ELÉCTRICAS"),         th("PMS")],
        [td("OC"),                  td("4508944971 (S3) / 4508944973 (S4)"), td("4508945953")],
        [td("Referencia ABB"),      td("R-2631036"),                 td("E-2611027")],
        [td("PM ABB"),              td("Rodrigo Mack"),              td("Pablo Kalis")],
        [td("Monto USD"),           td("S3: $8,127,765.50  |  S4: $3,686,311.55"), td("Por definir")],
        [td("Plazo"),               td("S3: 12 meses  |  S4: 11 meses"),   td("12 meses")],
        [td("Entrega contractual"), td("S3: 14-MAY-2027  |  S4: 08-ABR-2027"), td("14-MAY-2027")],
        [td("Penalidad"),           td("1% semanal — tope 10%"),    td("1% semanal — tope 10%")],
        [td("Garantía extendida"),  td("Hasta 31-ENE-2029"),        td("Hasta 31-ENE-2029")],
        [td("RFSU Proyecto"),       td("31-ENE-2028  (idéntico en P6 y Gantt ABB)"), td("=")],
    ]
    t2 = Table(info, colWidths=[5*cm, 13*cm, 9*cm])
    t2.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  BD),
        ("BACKGROUND",    (0,1), (-1,-1), GRY),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [GRY, WH]),
        ("GRID",          (0,0), (-1,-1), 0.5, HexColor("#CCCCCC")),
        ("TOPPADDING",    (0,0), (-1,-1), 4),
        ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("LEFTPADDING",   (0,0), (-1,-1), 6),
    ]))
    story.append(t2)
    story.append(Spacer(1, 1*cm))

    story.append(Paragraph("Fecha de generación: 26-MAY-2026  |  Basado en: KOM 18-MAY y 22-MAY-2026, Cronograma P6 Rev.0 LB Comercial, Gantt EPC Integrado",
                            style("fn", sz=7, color=DGY, align=TA_CENTER)))


def seccion_camino_critico(story):
    story.append(PageBreak())
    header_band(story, "CAMINO CRÍTICO — SALA #3 (BT)", "La Sala #3 determina el despacho más tardío: 14-MAY-2027  |  OC 4508944971  |  12 meses")

    # Flujograma textual del camino crítico
    flujo = [
        [td("⬇", align=TA_CENTER, bold=True, color=R), td("ETAPA", bold=True, color=WH), td("FECHA PLAN", bold=True, color=WH), td("ID", bold=True, color=WH), td("CONSECUENCIA SI SE DEMORA", bold=True, color=WH)],
        [td("1", align=TA_CENTER, bold=True, color=R),
         td("Aceptación OC — Inicio oficial"),
         td("12-MAY-2026", align=TA_CENTER, bold=True, color=OD),
         td("HC-01"),
         td("Base del proyecto. ✅ Completado")],
        [td("2", align=TA_CENTER, bold=True, color=R),
         td("Emisión / APR Ingeniería Básica S#3"),
         td("22-MAY → 26-JUN-2026", align=TA_CENTER),
         td("S3-01"),
         td("Demora → FP-S3-1 tarde → OC Turquía tarde → todo se corre")],
        [td("⚠", align=TA_CENTER, bold=True, color=HexColor("#9C5700")),
         td("OPEN ISSUE: Confirmar CCM7 (3×22kW → 2×45kW+VFD)"),
         td("VENCE 05-JUN-2026", align=TA_CENTER, bold=True, color=RKD),
         td("S3-WARN"),
         td("Si AESA no confirma antes del 05-JUN: IB no cierra → cadena entera se retrasa")],
        [td("◆", align=TA_CENTER, bold=True, color=R),
         td("FREEZING POINT — Ing. Básica S#3"),
         td("03-JUL-2026", align=TA_CENTER, bold=True, color=R),
         td("FP-S3-1"),
         td("1 día late → OC ductos Turquía (6m lead time) tarde → despacho late → PENALIDAD")],
        [td("3", align=TA_CENTER, bold=True, color=R),
         td("OC Ductos Barras Turquía + VFDs BT Finlandia"),
         td("06-JUL-2026", align=TA_CENTER),
         td("S3-03/04"),
         td("Lead time 6m (ductos) y 4-5m (VFDs). Entrega esperada: Dic-2026 / Ene-2027")],
        [td("4", align=TA_CENTER, bold=True, color=R),
         td("Emisión / APR Ingeniería Detalle S#3"),
         td("29-JUN → 14-AGO-2026", align=TA_CENTER),
         td("S3-02"),
         td("Demora → FP-S3-2 tarde → OC fabricación 37 CCMs (Brasil) tarde")],
        [td("◆", align=TA_CENTER, bold=True, color=R),
         td("FREEZING POINT — Ing. Detalle S#3"),
         td("14-AGO-2026", align=TA_CENTER, bold=True, color=R),
         td("FP-S3-2"),
         td("1 día late → OC 37 CCMs Brasil tarde → ensayos tarde → FAT S#3 tarde")],
        [td("5", align=TA_CENTER, bold=True, color=R),
         td("Emisión / APR Ingeniería Constructiva S#3"),
         td("17-AGO → 12-SEP-2026", align=TA_CENTER),
         td("S3-03"),
         td("Planos constructivos. Errores → retrabajo en Shelter Mendoza")],
        [td("◆◆", align=TA_CENTER, bold=True, color=R),
         td("★★ FREEZING POINT IC S#3 — ÚLTIMO FP DE TODA LA PROVISIÓN ABB"),
         td("31-AGO-2026", align=TA_CENTER, bold=True, color=R),
         td("FP-S3-3"),
         td("HITO MÁS CRÍTICO: 1 semana late → montaje Shelter late → FAT late → PENALIDAD 1%/sem")],
        [td("6", align=TA_CENTER, bold=True, color=R),
         td("Fabricación 37 CCMs BT (Brasil) + Ensayos"),
         td("14-SEP → 23-NOV-2026", align=TA_CENTER),
         td("S3-04"),
         td("37 tableros CCM 4000A 65kA. Ensayos 20 días. Fallo → retrabajo → FAT tarde")],
        [td("7", align=TA_CENTER, bold=True, color=R),
         td("Acopio materiales en Argentina (HITO 25%)"),
         td("07-DIC-26 → 29-ENE-27", align=TA_CENTER),
         td("S3-05"),
         td("Llegada de ductos + VFDs + CCMs a Argentina. Trigger pago 25%")],
        [td("8", align=TA_CENTER, bold=True, color=R),
         td("Montaje e integración en Shelter S#3 (Bottino — Mendoza)"),
         td("01-FEB → 26-MAR-2027", align=TA_CENTER),
         td("S3-06"),
         td("37 CCMs + 6 ductos barras + 3 VFDs BT + HVAC + baterías KMT. 8 semanas")],
        [td("9", align=TA_CENTER, bold=True, color=FA),
         td("FAT Sala Eléctrica #3  ← FAT MÁS TARDÍO de ABB"),
         td("02-ABR → 19-ABR-2027", align=TA_CENTER, bold=True, color=FA),
         td("S3-07"),
         td("Holgura +35d vs P6 (cierra FDE1103 el 24-MAY). Cada semana late reduce margen")],
        [td("▶", align=TA_CENTER, bold=True, color=DE),
         td("DESPACHO Sala #3  ← DESPACHO MÁS TARDÍO"),
         td("10-MAY → 14-MAY-2027", align=TA_CENTER, bold=True, color=DE),
         td("S3-08"),
         td("HOLGURA +45d vs P6 (inicio montaje campo 28-JUN-2027)")],
        [td("★", align=TA_CENTER, bold=True, color=R),
         td("★ ENTREGA S#3 — Cumplimiento OC 4508944971"),
         td("14-MAY-2027", align=TA_CENTER, bold=True, color=R),
         td("S3-★"),
         td("FECHA CONTRACTUAL. Penalidad si excede: 1%/semana. Tope: 10% del contrato")],
    ]

    t = Table(flujo, colWidths=[1.2*cm, 8.5*cm, 4*cm, 2.2*cm, 11.1*cm])
    ts = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  BD),
        ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#CCCCCC")),
        ("TOPPADDING",    (0,0), (-1,-1), 3),
        ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4),
        ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [S3L, WH]),
    ])
    # Resaltar FPs y hitos especiales
    for i, row in enumerate(flujo[1:], 1):
        txt = str(row[1].text if hasattr(row[1],'text') else row[1])
        if "FP-S3-3" in str(flujo[i][3]):
            ts.add("BACKGROUND", (0,i), (-1,i), RK)
        elif "FP-S3" in str(flujo[i][3]):
            ts.add("BACKGROUND", (0,i), (-1,i), HexColor("#FFCCCC"))
        elif "WARN" in str(flujo[i][3]):
            ts.add("BACKGROUND", (0,i), (-1,i), YL)
        elif "S3-★" in str(flujo[i][3]):
            ts.add("BACKGROUND", (0,i), (-1,i), HexColor("#FF9999"))
    t.setStyle(ts)
    story.append(t)


def seccion_freezing(story):
    story.append(PageBreak())
    header_band(story, "FREEZING POINTS — REGISTRO DE SEGUIMIENTO",
                "Congelamiento de diseño: cambio post-FP requiere Nota de Cambio | Sin hito equivalente en P6 comercial ⚠")

    story.append(Paragraph(
        "<b>IMPORTANTE:</b> Ninguno de los 8 Freezing Points tiene representación en el cronograma P6 comercial del proyecto. "
        "Un desvío en ellos NO es visible en el P6 hasta que impacta el despacho. "
        "Se recomienda solicitar a AESA incorporar al menos FP-S3-3 y FP-S4-3 como hitos controlables en el P6.",
        style("w", sz=8, color=HexColor("#9C5700"))))
    story.append(Spacer(1, 3*mm))

    hdrs_fp = [th("ID"), th("Sala"), th("Freezing Point"), th("Fecha Plan"), th("Fecha Real"),
               th("Estado"), th("Días δ"), th("Consecuencia de desvío")]
    fps_data = [
        ("FP-P1",   "PMS",   "Ing. Básica PMS",                   "19-JUN-2026", "Pendiente",
         "Demora procura PMS. PMS tiene holgura amplia (+222d vs P6)"),
        ("FP-P2",   "PMS",   "Ing. Detalle PMS",                  "31-JUL-2026", "Pendiente",
         "Retrasa inicio fabricación tablero PMS"),
        ("FP-S4-1", "S#4",   "Ing. Básica S#4 → OC Celdas Turquía","26-JUN-2026","Pendiente",
         "Lead time 6m. Demora en OC → celdas MT llegan tarde → FAT S#4 tarde"),
        ("FP-S4-2", "S#4",   "Ing. Detalle S#4",                  "14-AGO-2026", "Pendiente",
         "Retrasa fabricación tableros auxiliares BsAs"),
        ("FP-S4-3", "S#4",   "Ing. Constructiva S#4 ← Cierra Ing","28-AGO-2026", "Pendiente",
         "Último FP S#4. Demora → montaje Shelter S#4 tarde → FAT S#4 tarde"),
        ("FP-S3-1", "S#3 ★", "Ing. Básica S#3 → OC Ductos + VFDs","03-JUL-2026", "Pendiente",
         "★ CRÍTICO: Lead time 6m (ductos Turquía). Demora → despacho S#3 tarde → penalidad"),
        ("FP-S3-2", "S#3 ★", "Ing. Detalle S#3 → OC 37 CCMs Brasil","14-AGO-2026","Pendiente",
         "★ CRÍTICO: Demora → fabricación 37 CCMs tarde → ensayos → FAT S#3 tarde"),
        ("FP-S3-3", "S#3 ★", "Ing. Constructiva S#3 ← ÚLTIMO FP ABB","31-AGO-2026","Pendiente",
         "★★ HITO MÁS CRÍTICO: 1 sem late → montaje Shelter late → FAT late → PENALIDAD 1%/sem"),
    ]

    rows = [hdrs_fp]
    for id_fp, sala, nombre, fecha, estado, consecuencia in fps_data:
        critica = "S#3" in sala
        bg = S3L if critica else (S4L if "S#4" in sala else PML)
        rows.append([
            td(id_fp, bold=True, color=R if critica else BD),
            td(sala,  bold=critica, color=R if critica else BD),
            td(nombre, bold=(id_fp == "FP-S3-3")),
            td(fecha, align=TA_CENTER, bold=True, color=R if critica else HexColor("#9C5700")),
            td("—    ", align=TA_CENTER, color=DGY),   # Fecha Real vacía
            td(estado, align=TA_CENTER, color=HexColor("#9C5700")),
            td("—", align=TA_CENTER, color=DGY),        # Días δ vacío
            td(consecuencia, color=RKD if critica else HexColor("#7D4B00")),
        ])

    t = Table(rows, colWidths=[1.8*cm, 1.8*cm, 5.5*cm, 2.8*cm, 2.8*cm, 2.2*cm, 1.6*cm, 9.5*cm])
    ts = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  BD),
        ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#CCCCCC")),
        ("TOPPADDING",    (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4), ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
    ])
    for i, (id_fp, sala, _, __, ___, ____) in enumerate(fps_data, 1):
        critica = "S#3" in sala
        bg = S3L if critica else (S4L if "S#4" in sala else PML)
        if id_fp == "FP-S3-3":
            ts.add("BACKGROUND", (0,i), (-1,i), RK)
        else:
            ts.add("BACKGROUND", (0,i), (-1,i), bg)
    t.setStyle(ts)
    story.append(t)


def seccion_holguras(story):
    story.append(Spacer(1, 5*mm))
    story.append(Paragraph("<b>HOLGURAS ABB vs CRONOGRAMA P6 COMERCIAL</b>",
                            style("hh", sz=11, color=BD, bold=True)))
    story.append(Spacer(1, 2*mm))

    hols = [
        [th("Provisión ABB"), th("Fecha ABB (peor caso)"), th("P6 la necesita"), th("Holgura"), th("Estado")],
        [td("FAT PMS"),         td("23-OCT-2026",align=TA_CENTER), td("02-JUN-2027",align=TA_CENTER), td("+222 días",align=TA_CENTER,bold=True,color=OD), td("✅ AMPLIA",align=TA_CENTER,bold=True,color=OD)],
        [td("FAT Sala #3"),     td("19-ABR-2027",align=TA_CENTER), td("24-MAY-2027 (P6 FDE1103)",align=TA_CENTER), td("+35 días",align=TA_CENTER,bold=True,color=OD), td("✅ OK",align=TA_CENTER,bold=True,color=OD)],
        [td("Despacho Sala #3"),td("14-MAY-2027",align=TA_CENTER), td("28-JUN-2027 (montaje campo)",align=TA_CENTER), td("+45 días",align=TA_CENTER,bold=True,color=OD), td("✅ OK",align=TA_CENTER,bold=True,color=OD)],
        [td("Despacho Sala #4"),td("08-ABR-2027",align=TA_CENTER), td("28-JUN-2027 (montaje campo)",align=TA_CENTER), td("+81 días",align=TA_CENTER,bold=True,color=OD), td("✅ OK",align=TA_CENTER,bold=True,color=OD)],
        [td("iFAT Integración"), td("25-JUN-2027",align=TA_CENTER),td("28-JUN-2027 (TA1124 P6)",align=TA_CENTER), td("+3 días",align=TA_CENTER,bold=True,color=RKD), td("🔴 SIN MARGEN",align=TA_CENTER,bold=True,color=RKD)],
        [td("FDE1104 S#4 en P6"),td("23-JUN-2027 (cierre P6)",align=TA_CENTER),td("28-JUN-2027 (inicio montaje)",align=TA_CENTER), td("+5 días en P6",align=TA_CENTER,bold=True,color=RKD), td("⚠ AJUSTADO",align=TA_CENTER,bold=True,color=HexColor("#9C5700"))],
        [td("RFSU"),            td("31-ENE-2028",align=TA_CENTER), td("31-ENE-2028 (A10290 P6)",align=TA_CENTER), td("= IDÉNTICO",align=TA_CENTER,bold=True,color=OD), td("✅ ALINEADO",align=TA_CENTER,bold=True,color=OD)],
        [td("Freezing Points",bold=True), td("Jun-Ago 2026",align=TA_CENTER), td("SIN HITO EN P6",align=TA_CENTER,bold=True,color=RKD), td("SIN CONTROL",align=TA_CENTER,bold=True,color=RKD), td("⚠ RIESGO",align=TA_CENTER,bold=True,color=HexColor("#9C5700"))],
    ]
    t = Table(hols, colWidths=[4*cm, 4*cm, 6*cm, 3*cm, 4*cm])
    ts = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  BD),
        ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#CCCCCC")),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [GRY, WH]),
        ("TOPPADDING",    (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4), ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ("BACKGROUND",    (0,5), (-1,5),  RK),
        ("BACKGROUND",    (0,6), (-1,6),  YL),
        ("BACKGROUND",    (0,8), (-1,8),  YL),
    ])
    t.setStyle(ts)
    story.append(t)


def seccion_pagos(story):
    story.append(PageBreak())
    header_band(story, "HITOS DE PAGO Y FECHAS CONTRACTUALES",
                "Aplican a las 3 OC — S3: $8,127,765.50  |  S4: $3,686,311.55  |  PMS: a definir")

    pago_rows = [
        [th("#"), th("%"), th("Hito"), th("Fecha Plan"), th("Condición"), th("Póliza / Garantía")],
        [td("1",align=TA_CENTER,bold=True), td("10%",align=TA_CENTER,bold=True,color=BD),
         td("Envío Ingeniería Básica"), td("12-JUN-2026",align=TA_CENTER,bold=True),
         td("Emisión planos principales (IB) de S3, S4 y PMS"),
         td("Póliza Caución Avance Fabricación — debe estar aprobada ANTES de certificar")],
        [td("2",align=TA_CENTER,bold=True), td("10%",align=TA_CENTER,bold=True,color=BD),
         td("Aprobación Ingeniería Básica"), td("03-JUL-2026",align=TA_CENTER,bold=True),
         td("Aprobación de AESA. Según cronograma"),
         td("Póliza Caución Avance Fabricación vigente")],
        [td("3",align=TA_CENTER,bold=True), td("20%",align=TA_CENTER,bold=True,color=BD),
         td("Finalización de FAT"), td("19-ABR-2027",align=TA_CENTER,bold=True),
         td("FAT de integración tableros en fábrica (S#3 más tardío)"),
         td("Póliza Caución Avance Fabricación. PMS: analizar prueba integrada en Shelter")],
        [td("4",align=TA_CENTER,bold=True), td("25%",align=TA_CENTER,bold=True,color=BD),
         td("Acopio Materiales en Argentina"), td("29-ENE-2027",align=TA_CENTER,bold=True),
         td("Arribo de celdas MT + tableros BT + ductos + VFDs MT/BT + PMS + UPS + auxiliares"),
         td("Póliza Caución Acopio de Materiales")],
        [td("5",align=TA_CENTER,bold=True), td("30%",align=TA_CENTER,bold=True,color=BD),
         td("Contra Entrega"), td("14-MAY-2027",align=TA_CENTER,bold=True),
         td("Shelter listo para transportar. Liberación por Calidad del COMPRADOR"),
         td("— (retiro AESA en Mendoza)")],
        [td("6",align=TA_CENTER,bold=True), td("5%",align=TA_CENTER,bold=True,color=BD),
         td("Entrega Databook"), td("~01-JUL-2027",align=TA_CENTER,bold=True),
         td("Databook completo y aprobado por AESA"),
         td("—")],
    ]
    t = Table(pago_rows, colWidths=[1*cm, 1.2*cm, 4.5*cm, 3*cm, 8*cm, 9.3*cm])
    ts = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  BD),
        ("ROWBACKGROUNDS",(0,1), (-1,-1), [HexColor("#EEF4FF"), WH]),
        ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#CCCCCC")),
        ("TOPPADDING",    (0,0), (-1,-1), 4), ("BOTTOMPADDING", (0,0), (-1,-1), 4),
        ("LEFTPADDING",   (0,0), (-1,-1), 4), ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
        ("FONTSIZE",      (0,0), (-1,-1), 8),
    ])
    t.setStyle(ts)
    story.append(t)


def seccion_alertas(story):
    story.append(PageBreak())
    header_band(story, "ALERTAS Y RIESGOS ACTIVOS", "Estado al 26-MAY-2026  |  ABB Ref: R-2631036 / E-2611027")

    alertas = [
        ("🔴 CRÍTICO", "Confirmar CCM7: 3×22kW → 2×45kW+VFD en S#3",
         "AESA (Dario Stirparo)", "05-JUN-2026",
         "IB de S#3 no cierra → FP-S3-1 se corre → OC Turquía tarde → despacho S#3 tarde → PENALIDAD",
         "AESA contactando cliente. ABB: follow-up semanal obligatorio"),

        ("🔴 CRÍTICO", "Cronograma final con Hold Points",
         "ABB (Rodrigo Mack)", "01-JUN-2026",
         "Sin cronograma final: no hay baseline contractual, no hay pólizas, no hay hitos certificables",
         "ABB debe incluir OCs a 3ros confirmadas"),

        ("🟡 IMPORTANTE", "Información ductos de barra: posición transformadores vs Salas",
         "AESA (con cliente Pluspetrol)", "Antes de FP-S4-1 (26-JUN)",
         "Sin datos: ABB no puede finalizar diseño ductos de barras (Turquía, lead 6 meses)",
         "AESA en contacto con Pluspetrol. ABB confirmar fecha de respuesta"),

        ("🟡 IMPORTANTE", "Acceso portal ShareFile para ABB",
         "ABB (interno)", "SEM 26-MAY",
         "Sin acceso: no se pueden emitir documentos vía transmittal oficial (requisito KOM)",
         "ABB gestionando habilitación interna para Pantolini / Palma / Mack"),

        ("🟡 IMPORTANTE", "Pólizas de Caución — requeridas antes del Hito 10%",
         "ABB (Legal/Seguros)", "Antes de PAY-01 (12-JUN)",
         "Sin póliza aprobada por AESA Seguros: no se puede certificar ni cobrar el Hito 10%",
         "Iniciar gestión con aseguradora. Póliza debe estar aprobada antes de la inspección"),

        ("🟡 IMPORTANTE", "Confirmación unifilares tableros CCMs (cambios de cargas)",
         "AESA (Tulio De La Torre)", "Antes de FP-S3-1 (03-JUL)",
         "Cambio de cargas post-FP → Nota de Cambio → impacto costo y plazo",
         "AESA en revisión con cliente. Seguimiento en reuniones técnicas quincenales"),

        ("🟠 ATENCIÓN", "Freezing Points NO en P6 comercial",
         "ABB + AESA (PM)", "SEM 01-JUN",
         "Desvíos en FPs no controlables en P6 hasta que impactan despacho (demasiado tarde)",
         "Solicitar a AESA incorporar FP-IC S#3 (31-AGO), FP-IC S#4 (28-AGO) y FATs en P6"),

        ("🟠 ATENCIÓN", "iFAT: holgura de solo 3 días (ABB 25-JUN vs P6 28-JUN-2027)",
         "ABB + EPC", "JUN-2027",
         "Si cualquier vendor llega tarde al iFAT: comisionado se corre, RFSU en riesgo",
         "Incluir en cronograma final hitos de todos los vendors para el iFAT"),

        ("🟢 INFO", "Prueba integrada PMS en Shelter: factibilidad pendiente",
         "ABB (Pablo Kalis)", "SEM 08-JUN",
         "AESA solicita FAT integral PMS+gabinetes en Shelter. Puede impactar timing despacho PMS",
         "ABB evalúa y define en cronograma final PMS"),
    ]

    sev_bg = {
        "🔴 CRÍTICO":    RK,
        "🟡 IMPORTANTE": YL,
        "🟠 ATENCIÓN":   HexColor("#FFE0B2"),
        "🟢 INFO":       OK,
    }
    sev_fc = {
        "🔴 CRÍTICO":    RKD,
        "🟡 IMPORTANTE": YD,
        "🟠 ATENCIÓN":   HexColor("#7B3F00"),
        "🟢 INFO":       OD,
    }

    rows = [[th("#"), th("Severidad"), th("Alerta"), th("Responsable"),
             th("Fecha límite"), th("Impacto"), th("Mitigación")]]
    for i, (sev, alerta, resp, fecha, impacto, mitig) in enumerate(alertas, 1):
        bg = sev_bg.get(sev, WH)
        fc = sev_fc.get(sev, black)
        rows.append([
            td(str(i), align=TA_CENTER, bold=True),
            td(sev, bold=True, color=fc),
            td(alerta, bold=True, color=fc),
            td(resp),
            td(fecha, align=TA_CENTER, bold=True, color=fc),
            td(impacto, color=fc),
            td(mitig),
        ])

    t = Table(rows, colWidths=[0.6*cm, 2.6*cm, 5.5*cm, 4*cm, 3*cm, 6.5*cm, 5.8*cm])
    ts = TableStyle([
        ("BACKGROUND",    (0,0), (-1,0),  BD),
        ("GRID",          (0,0), (-1,-1), 0.4, HexColor("#CCCCCC")),
        ("TOPPADDING",    (0,0), (-1,-1), 3), ("BOTTOMPADDING", (0,0), (-1,-1), 3),
        ("LEFTPADDING",   (0,0), (-1,-1), 4), ("RIGHTPADDING",  (0,0), (-1,-1), 4),
        ("VALIGN",        (0,0), (-1,-1), "TOP"),
    ])
    for i, (sev, *_) in enumerate(alertas, 1):
        ts.add("BACKGROUND", (0,i), (-1,i), sev_bg.get(sev, WH))
    t.setStyle(ts)
    story.append(t)


def seccion_gantt(story):
    """Gantt textual simplificado en formato tabla por trimestres."""
    story.append(PageBreak())
    header_band(story, "GANTT SIMPLIFICADO — VISIÓN TRIMESTRAL",
                "May 2026 → Feb 2028  |  Colores: Rojo=Cadena Crítica S#3 | Azul=S#4 | Naranja=FAT | Violeta=Despacho")

    QS = [
        ("Q2-2026", "May-Jun 2026"),
        ("Q3-2026", "Jul-Ago-Sep 2026"),
        ("Q4-2026", "Oct-Nov-Dic 2026"),
        ("Q1-2027", "Ene-Feb-Mar 2027"),
        ("Q2-2027", "Abr-May-Jun 2027"),
        ("Q3-2027", "Jul-Ago-Sep 2027"),
        ("Q1-2028", "Oct 27-Feb 2028"),
    ]

    Q_RANGES = [
        (date(2026,4,1),  date(2026,6,30)),
        (date(2026,7,1),  date(2026,9,30)),
        (date(2026,10,1), date(2026,12,31)),
        (date(2027,1,1),  date(2027,3,31)),
        (date(2027,4,1),  date(2027,6,30)),
        (date(2027,7,1),  date(2027,9,30)),
        (date(2027,10,1), date(2028,2,28)),
    ]

    def q_fill(ini, fin, q_ini, q_fin, tipo, critica):
        if fin < q_ini or ini > q_fin:
            return ""
        # overlap
        overlap_start = max(ini, q_ini)
        overlap_end   = min(fin, q_fin)
        total_q = (q_fin - q_ini).days or 1
        total_act = (overlap_end - overlap_start).days + 1
        blocks = max(1, round(total_act / total_q * 8))
        if tipo in ("fp","hito","pago","hito_contrato"):
            return "◆" if ini >= q_ini and ini <= q_fin else "─"
        chars = {"fat":"█","des":"▶","int":"░","act":"█","warn":"⚠"}
        ch = chars.get(tipo, "█")
        return ch * blocks

    act_rows = [[th("Sala"), th("ID"), th("Actividad")] + [th(q) for q,_ in QS]]
    grupo_ant = None
    for (grupo, id_act, nombre, ini, fin, tipo, critica, oc, holgura) in ACTIVIDADES:
        if grupo != grupo_ant:
            lbl = grupo.replace(" ★ CRÍTICA","★").upper()
            sep = [td(lbl, bold=True, color=WH)] + [""] * (len(QS) + 2)
            act_rows.append(sep)
            grupo_ant = grupo

        if tipo in ("pago",):
            continue  # ya tienen su sección propia

        bar_cells = []
        for qi, (q_ini, q_fin) in enumerate(Q_RANGES):
            ch = q_fill(ini, fin, q_ini, q_fin, tipo, critica)
            if ch:
                fc_bar = (R if (critica and tipo not in ("fp","hito","hito_contrato","pago")) else
                          (FA if tipo == "fat" else
                           (DE if tipo == "des" else
                            (GR if tipo in ("int","integracion") else
                             (R if tipo in ("fp","hito","hito_contrato") else BL)))))
                bar_cells.append(td(ch, bold=False, color=fc_bar, align=TA_CENTER))
            else:
                bar_cells.append(td(""))

        short = nombre[:45] + ("…" if len(nombre) > 45 else "")
        act_rows.append([
            td(grupo[:10], sz=7),
            td(id_act, bold=critica or tipo=="fp", color=R if critica else (BL if not critica else BD), sz=7),
            td(short, bold=critica and tipo in ("fp","hito_contrato","des"), sz=7.5,
               color=R if (critica and tipo in ("fp","hito_contrato","des")) else black),
        ] + bar_cells)

    col_ws = [1.5*cm, 1.6*cm, 8*cm] + [2.7*cm] * len(QS)
    t = Table(act_rows, colWidths=col_ws)
    ts = TableStyle([
        ("GRID",          (0,0), (-1,-1), 0.3, HexColor("#DDDDDD")),
        ("TOPPADDING",    (0,0), (-1,-1), 2), ("BOTTOMPADDING", (0,0), (-1,-1), 2),
        ("LEFTPADDING",   (0,0), (-1,-1), 2), ("RIGHTPADDING",  (0,0), (-1,-1), 2),
        ("FONTSIZE",      (0,0), (-1,-1), 7),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
        ("ALIGN",         (3,0), (-1,-1), "CENTER"),
    ])
    # Colorear separadores de grupo
    for i, row in enumerate(act_rows):
        if len(row) > 3 and hasattr(row[0], "text"):
            txt = row[0].text if hasattr(row[0], "text") else ""
        if isinstance(row[0], Paragraph):
            t_text = row[0].text if hasattr(row[0],'text') else ""
    # Zebra y separadores
    gi = 0
    for i, row in enumerate(act_rows):
        cell_val = row[0]
        if isinstance(cell_val, Paragraph):
            raw = ""
            for frag in cell_val.frags if hasattr(cell_val,'frags') else []:
                raw += getattr(frag,'text','')
            if not raw:
                raw = str(cell_val)
        else:
            raw = str(cell_val)
        if any(g in raw.upper() for g in ["CONTRATO","PMS","SALA #3","SALA #4","EPC"]) and i > 0:
            ts.add("BACKGROUND", (0,i), (-1,i), HexColor("#2E3A4E"))
            ts.add("TEXTCOLOR",  (0,i), (-1,i), WH)
        elif i % 2 == 0:
            ts.add("BACKGROUND", (0,i), (-1,i), GRY)
    ts.add("BACKGROUND", (0,0), (-1,0), BD)
    t.setStyle(ts)
    story.append(t)


# ─────────────────────────────────────────────────────────────────────────────
# DOCUMENTO
# ─────────────────────────────────────────────────────────────────────────────
def generar():
    doc = SimpleDocTemplate(
        OUTPUT_PDF,
        pagesize=PAGE,
        leftMargin=1.5*cm, rightMargin=1.5*cm,
        topMargin=1.5*cm,  bottomMargin=1.5*cm,
        title="Cronograma Seguimiento ABB — CPF-2 La Calera II",
        author="ABB Argentina",
        subject="Provisión Salas Eléctricas #3, #4 y PMS",
    )

    story = []
    portada(story)
    seccion_camino_critico(story)
    seccion_freezing(story)
    seccion_holguras(story)
    seccion_pagos(story)
    seccion_alertas(story)
    seccion_gantt(story)

    # Pie final
    story.append(PageBreak())
    story.append(Spacer(1, 3*cm))
    data = [
        [Paragraph("INFORMACIÓN DEL DOCUMENTO", style("fd",sz=14,color=WH,bold=True,align=TA_CENTER))],
        [Paragraph("ABB Argentina — Proyecto R-2631036 / E-2611027\nCPF-2 La Calera II — Vaca Muerta\nGenerado: 26-MAY-2026\nBasado en: KOM 18-MAY y 22-MAY-2026 | Cronograma P6 Rev.0 | Gantt EPC Integrado",
                   style("fi",sz=10,color=WH,align=TA_CENTER,leading=16))],
    ]
    t = Table(data, colWidths=[27*cm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), BD),
        ("TOPPADDING",    (0,0), (-1,-1), 12),
        ("BOTTOMPADDING", (0,0), (-1,-1), 12),
    ]))
    story.append(t)

    doc.build(story)
    print(f"PDF generado: {OUTPUT_PDF}")

if __name__ == "__main__":
    generar()
