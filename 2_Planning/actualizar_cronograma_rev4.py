#!/usr/bin/env python3
"""
Rev 4 del Cronograma General ABB.
Nuevas fuentes PMS:
  - E-2611027 Informe de Avance CPF2 LA CALERA - JUL_2026.pdf  (24-JUL-2026)
  - 5155-00-0007-VE-CO-PG-001_REV B.mpp  (PMS Schedule Rev B)
Genera: CRONOGRAMA_GENERAL_ABB_Rev4_2026-07-27.xlsx
"""

import shutil
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

SRC = "CRONOGRAMA_GENERAL_ABB_Rev3_2026-07-21.xlsx"
DST = "CRONOGRAMA_GENERAL_ABB_Rev4_2026-07-27.xlsx"
REV   = "Rev 4"
FECHA = "27-JUL-2026"

# ── colour helpers ────────────────────────────────────────────────────────────
def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, name="Calibri", italic=italic)

thin  = Side(style="thin",   color="BFBFBF")
thick = Side(style="medium", color="000000")
def border_all():   return Border(left=thin, right=thin, top=thin,  bottom=thin)
def border_thick(): return Border(left=thin, right=thin, top=thick, bottom=thick)
def center(wrap=False): return Alignment(horizontal="center", vertical="center", wrap_text=wrap)
def left(wrap=False):   return Alignment(horizontal="left",   vertical="center", wrap_text=wrap)

# Palette
C_TITLE   = "1F3864"
C_ALERTA  = "C00000"
C_WARN    = "FF9900"
C_OK      = "375623"
C_HEAD    = "D6E4F0"
C_EVEN    = "EBF3FB"
C_YELLOW  = "FFF2CC"
C_ORANGE  = "FCE4D6"
C_RED     = "FFE7E7"
C_GREEN   = "E2EFDA"

# PMS-specific accent colours
C_PMS_HEAD  = "1F4E79"   # dark navy for PMS title
C_PMS_LIGHT = "DEEAF1"   # light blue for PMS section headers

# ── Step 1: copy Rev 3 ───────────────────────────────────────────────────────
shutil.copy(SRC, DST)
wb = load_workbook(DST)

# ── Step 2: update revision strings in every sheet ───────────────────────────
OLD_STRINGS = [
    "Rev 3", "REV 3", "rev3", "Rev3",
    "21-JUL-2026", "2026-07-21", "2026-07-21",
    "CRONOGRAMA_GENERAL_ABB_Rev3",
]
NEW_STRINGS = [
    REV, REV, "rev4", "Rev4",
    FECHA, "2026-07-27", "2026-07-27",
    "CRONOGRAMA_GENERAL_ABB_Rev4",
]

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                v = cell.value
                for old, new in zip(OLD_STRINGS, NEW_STRINGS):
                    v = v.replace(old, new)
                cell.value = v

# ── Step 3: update sheet 2_PMS_Detalle with Rev B + July 2026 status ─────────
ws_pms = wb["2_PMS_Detalle"]

# Find where data starts (look for PP-IB row)
pms_task_col = None
pms_start_row = None
for r in ws_pms.iter_rows():
    for c in r:
        if c.value and "PP-IB" in str(c.value):
            pms_task_col = c.column
            pms_start_row = c.row
            break
    if pms_start_row:
        break

# If found, add/update a status column at the right end of the data
# Add a header row note about Rev B and status
# First, find the maximum column used
max_col = ws_pms.max_column

# Add Rev B source note in a header area — look for title row
for r in ws_pms.iter_rows(max_row=4):
    for c in r:
        if c.value and isinstance(c.value, str) and "PMS" in c.value:
            c.value = (c.value + f"  |  Cronograma Rev B (5155-00-0007-VE-CO-PG-001_REV B)  |  IA Jul-2026"
                      ).replace("  |    |  ", "  |  ")
            break

# PMS Rev B updated schedule table — append after existing content
# Find the last used row
last_row = ws_pms.max_row + 2

# Column widths (re-enforce for PMS detail sheet)
ws_pms.column_dimensions["A"].width = 8
ws_pms.column_dimensions["B"].width = 32
ws_pms.column_dimensions["C"].width = 14
ws_pms.column_dimensions["D"].width = 14
ws_pms.column_dimensions["E"].width = 14
ws_pms.column_dimensions["F"].width = 14
ws_pms.column_dimensions["G"].width = 20

_rp = [last_row]

def pw_title(text, span=7, bg=C_PMS_HEAD):
    r = _rp[0]
    ws_pms.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws_pms.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, color="FFFFFF", size=11)
    c.alignment = center(); c.border = border_all()
    ws_pms.row_dimensions[r].height = 24
    _rp[0] += 1

def pw_section(text, span=7, bg=C_PMS_LIGHT):
    r = _rp[0]
    ws_pms.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws_pms.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, size=9)
    c.alignment = left(); c.border = border_thick()
    ws_pms.row_dimensions[r].height = 16
    _rp[0] += 1

def pw_hdr(cols):
    r = _rp[0]
    for ci, val in enumerate(cols, 1):
        c = ws_pms.cell(row=r, column=ci, value=val)
        c.fill = fill(C_HEAD); c.font = font(bold=True, size=9)
        c.alignment = center(wrap=True); c.border = border_all()
    ws_pms.row_dimensions[r].height = 30
    _rp[0] += 1

def pw_data(values, colors=None):
    r = _rp[0]
    bg = "FFFFFF" if r % 2 == 0 else C_EVEN
    for ci, val in enumerate(values, 1):
        c = ws_pms.cell(row=r, column=ci, value=val)
        bg_use = colors[ci-1] if colors and ci <= len(colors) else bg
        c.fill = fill(bg_use)
        c.font = font(size=9)
        c.alignment = left(wrap=True) if ci in (2, 7) else center()
        c.border = border_all()
    _rp[0] += 1

def pw_blank():
    _rp[0] += 1

pw_blank()
pw_title(f"ACTUALIZACIÓN PMS — CRONOGRAMA REV B + INFORME AVANCE JULIO 2026  |  {REV}  |  {FECHA}")

r = _rp[0]
ws_pms.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
c = ws_pms.cell(row=r, column=1,
    value="Fuentes: E-2611027 Informe Avance JUL-2026 (24-JUL-2026, PM Pablo Kalis) | "
          "5155-00-0007-VE-CO-PG-001_REV B.mpp | OC 4508945953")
c.fill = fill(C_HEAD); c.font = font(size=8); c.alignment = center(); c.border = border_all()
_rp[0] += 1

pw_blank()

# PMS Rev A → Rev B schedule comparison table
pw_section("COMPARATIVO CRONOGRAMA PMS: Rev A vs Rev B  (cambios detectados)")
pw_hdr(["ID", "Actividad PMS", "Inicio Rev A", "Fin Rev A", "Inicio Rev B", "Fin Rev B", "Cambio / Observación"])

# Rev A = prior data from 2_PMS_Detalle sheet; Rev B = updated from MPP binary + IA July
pms_comparison = [
    ("PP-IB", "Emisión IB PMS (docs RevA)",
     "08/06/2026", "17/07/2026",
     "08/06/2026", "17/07/2026",
     "Sin cambio. En progreso (avances reportados JUL-2026)"),
    ("PP-AI", "Aprobación IB por AESA (ciclo completo)",
     "22/06/2026", "14/08/2026",
     "22/06/2026", "14/08/2026",
     "Sin cambio en fecha. Pendiente señales cableadas, Mapa Modbus, Mapas Profinet de AESA"),
    ("PP-ID", "Emisión ID + Procedimientos FAT/SAT",
     "03/08/2026", "09/10/2026",
     "03/08/2026", "09/10/2026",
     "Sin cambio en fecha. ABB confirma inicio agosto"),
    ("PP-ARQTOP", "Arquitectura/Topología Red PMS",
     "—", "—",
     "01/08/2026", "31/08/2026",
     "NUEVA actividad Rev B. Esperada en agosto 2026 según IA Jul-2026"),
    ("PP-DESC", "Descripción Funcional PMS",
     "—", "—",
     "01/08/2026", "31/08/2026",
     "NUEVA actividad Rev B. Esperada en agosto 2026 según IA Jul-2026"),
    ("FP-P1", "Freezing Point: IB PMS aprobada",
     "21/08/2026", "21/08/2026",
     "21/08/2026", "21/08/2026",
     "Sin cambio"),
    ("FP-P2", "Freezing Point: ID PMS aprobado",
     "23/10/2026", "23/10/2026",
     "23/10/2026", "23/10/2026",
     "Sin cambio"),
    ("PP-MQ", "FAT Maqueta Integrada",
     "22/09/2026", "25/09/2026",
     "22/09/2026", "25/09/2026",
     "ALCANCE AMPLIADO Rev B: Incluye PMS/CCM/PCS de CPF1 Y CPF2. Job N° 4JAR001140P0003"),
    ("PP-FAB", "Armado Tableros PMS-001 + PMS-101",
     "11/09/2026", "17/12/2026",
     "11/09/2026", "17/12/2026",
     "Sin cambio en fechas"),
    ("PP-TI", "Pruebas Internas ABB",
     "18/12/2026", "07/01/2027",
     "18/12/2026", "07/01/2027",
     "Sin cambio"),
    ("PP-FAT", "FAT Tableros + Sistema PMS",
     "25/01/2027", "29/01/2027",
     "25/01/2027", "29/01/2027",
     "Sin cambio"),
    ("PP-DES", "Despacho PMS → Shelter",
     "04/02/2027", "17/02/2027",
     "04/02/2027", "17/02/2027",
     "Sin cambio"),
    ("PP-SH3", "FAT Integral Shelter PMS",
     "25/03/2027", "31/03/2027",
     "25/03/2027", "31/03/2027",
     "Sin cambio"),
    ("PP-DB", "DataBook PMS aprobado",
     "31/03/2027", "31/03/2027",
     "31/03/2027", "31/03/2027",
     "Sin cambio"),
]

change_colors = {
    "NUEVA": C_YELLOW,
    "ALCANCE AMPLIADO": C_ORANGE,
    "Sin cambio": None,
    "Sin cambio en": None,
    "Pendiente": C_ORANGE,
    "En progreso": C_EVEN,
}

for row_data in pms_comparison:
    r = _rp[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    obs_text = row_data[6]
    row_bg = bg_base
    for key, col in change_colors.items():
        if key in obs_text and col:
            row_bg = col
            break
    colors_list = [bg_base, bg_base, bg_base, bg_base, bg_base, bg_base, row_bg]
    pw_data(list(row_data), colors=colors_list)

pw_blank()

# Current status of each PMS activity
pw_section("ESTADO ACTUAL DE ACTIVIDADES PMS — Al 24-JUL-2026 (fuente: Informe Avance JUL-2026)")
pw_hdr(["ID", "Actividad", "Estado", "% Avance", "Inicio Real", "Fin Previsto", "Acción Requerida"])

pms_status = [
    ("PP-IB", "Emisión IB PMS",
     "EN CURSO", "~40%", "08/06/2026", "17/07/2026",
     "ABB reporta avances. Definiciones fibra óptica S3/S4 en proceso"),
    ("PP-ARQTOP", "Arquitectura/Topología Red PMS",
     "PENDIENTE", "0%", "01/08/2026", "31/08/2026",
     "Identificada en Rev B. Inicio programado agosto"),
    ("PP-DESC", "Descripción Funcional PMS",
     "PENDIENTE", "0%", "01/08/2026", "31/08/2026",
     "Identificada en Rev B. Inicio programado agosto"),
    ("PP-AI", "Aprobación IB por AESA",
     "BLOQUEADA", "0%", "22/06/2026", "14/08/2026",
     "AESA pendiente: señales cableadas con borneras + Mapa Modbus + Mapas Profinet"),
    ("PP-ID", "Emisión ID + Procedimientos",
     "PENDIENTE", "0%", "03/08/2026", "09/10/2026",
     "No iniciada. Depende de aprobación IB por AESA"),
    ("PP-MQ", "FAT Maqueta Integrada (PMS/CCM/PCS CPF1+CPF2)",
     "PENDIENTE", "0%", "22/09/2026", "25/09/2026",
     "Alcance ampliado incluye CPF1. Coordinación pendiente con CCM/PCS"),
    ("PP-FAB", "Armado Tableros PMS-001 + PMS-101",
     "NO INICIADA", "0%", "11/09/2026", "17/12/2026",
     "Pólizas OC 4508945953 emitidas (fiel cumplimiento + fondo reparo)"),
    ("PP-TI", "Pruebas Internas ABB",
     "NO INICIADA", "0%", "18/12/2026", "07/01/2027",
     "—"),
    ("PP-FAT", "FAT Tableros + Sistema PMS",
     "NO INICIADA", "0%", "25/01/2027", "29/01/2027",
     "Cert. FAT = 20% OC PMS"),
    ("PP-DES", "Despacho PMS → Shelter",
     "NO INICIADA", "0%", "04/02/2027", "17/02/2027",
     "—"),
    ("PP-SH3", "FAT Integral Shelter",
     "NO INICIADA", "0%", "25/03/2027", "31/03/2027",
     "—"),
    ("PP-DB", "DataBook PMS aprobado",
     "NO INICIADA", "0%", "31/03/2027", "31/03/2027",
     "Cert. Entrega DataBook = 5% OC PMS"),
]

status_bg = {
    "EN CURSO":   C_YELLOW,
    "BLOQUEADA":  C_RED,
    "PENDIENTE":  C_ORANGE,
    "NO INICIADA": C_EVEN,
    "COMPLETADO": C_GREEN,
}

for row_data in pms_status:
    r = _rp[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    estado = row_data[2]
    bg_estado = status_bg.get(estado, bg_base)
    colors_list = [bg_base, bg_base, bg_estado, bg_base, bg_base, bg_base, bg_base]
    pw_data(list(row_data), colors=colors_list)

pw_blank()

# Certification status PMS
pw_section("CERTIFICACIONES PMS — OC 4508945953")
pw_hdr(["Hito", "Descripción", "% OC", "Monto Est.", "Condición", "Estado Actual", "Observación"])

cert_pms = [
    ("Cert. 1", "Aprobación Ingeniería (IB+ID)", "10%", "—",
     "IB+ID aprobados por AESA", "NO EMITIDA",
     "Aclaraciones hitos presentadas 8-Jun-2026. En análisis AESA"),
    ("Cert. 2", "FAT Tableros PMS", "20%", "—",
     "FAT completado y aprobado", "NO EMITIDA",
     "Prevista ene-2027"),
    ("Cert. 3", "Entrega DataBook PMS aprobado", "5%", "—",
     "DataBook final aprobado por AESA", "NO EMITIDA",
     "Prevista mar-2027"),
]

for row_data in cert_pms:
    r = _rp[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    colors_list = [bg_base, bg_base, bg_base, bg_base, bg_base, C_RED, bg_base]
    pw_data(list(row_data), colors=colors_list)

# ── Step 4: Add new sheet 9_Estado_PMS_JUL2026 ───────────────────────────────
if "9_Estado_PMS_JUL2026" in wb.sheetnames:
    del wb["9_Estado_PMS_JUL2026"]
ws9 = wb.create_sheet("9_Estado_PMS_JUL2026")

# Column widths
col_widths = [("A",6),("B",34),("C",16),("D",10),("E",14),("F",14),("G",22)]
for col, w in col_widths:
    ws9.column_dimensions[col].width = w

_r9 = [1]

def r9():   return _r9[0]
def nr9():  _r9[0] += 1; return _r9[0] - 1

def w9_title(text, bg=C_PMS_HEAD, fgc="FFFFFF", span=7, height=26):
    r = _r9[0]
    ws9.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws9.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, color=fgc, size=12)
    c.alignment = center(); c.border = border_all()
    ws9.row_dimensions[r].height = height
    _r9[0] += 1

def w9_section(text, bg=C_PMS_LIGHT, span=7):
    r = _r9[0]
    ws9.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws9.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, size=10, color="1F4E79")
    c.alignment = left(); c.border = border_thick()
    ws9.row_dimensions[r].height = 18
    _r9[0] += 1

def w9_hdr(cols, bg=C_HEAD):
    r = _r9[0]
    for ci, val in enumerate(cols, 1):
        c = ws9.cell(row=r, column=ci, value=val)
        c.fill = fill(bg); c.font = font(bold=True, size=9)
        c.alignment = center(wrap=True); c.border = border_all()
    ws9.row_dimensions[r].height = 32
    _r9[0] += 1

def w9_data(values, colors=None, bold=False):
    r = _r9[0]
    bg = "FFFFFF" if r % 2 == 0 else C_EVEN
    for ci, val in enumerate(values, 1):
        c = ws9.cell(row=r, column=ci, value=val)
        bg_use = colors[ci-1] if colors and ci <= len(colors) else bg
        c.fill = fill(bg_use)
        c.font = font(bold=bold, size=9)
        c.alignment = left(wrap=True) if ci in (2, 7) else center()
        c.border = border_all()
    _r9[0] += 1

def w9_kv(key, val, key_bg=C_PMS_LIGHT, val_bg="FFFFFF", key_span=2, val_span=5):
    r = _r9[0]
    ws9.merge_cells(start_row=r, start_column=1, end_row=r, end_column=key_span)
    c = ws9.cell(row=r, column=1, value=key)
    c.fill = fill(key_bg); c.font = font(bold=True, size=9)
    c.alignment = left(); c.border = border_all()
    ws9.merge_cells(start_row=r, start_column=key_span+1, end_row=r, end_column=7)
    c2 = ws9.cell(row=r, column=key_span+1, value=val)
    c2.fill = fill(val_bg); c2.font = font(size=9)
    c2.alignment = left(wrap=True); c2.border = border_all()
    ws9.row_dimensions[r].height = 16
    _r9[0] += 1

def w9_blank():
    _r9[0] += 1

# ── HEADER ───────────────────────────────────────────────────────────────────
w9_title(f"ESTADO PMS — INFORME DE AVANCE JULIO 2026  |  ABB  |  {REV}  |  {FECHA}")
r = _r9[0]
ws9.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
c = ws9.cell(row=r, column=1,
    value="Fuente: E-2611027 Informe de Avance CPF2 LA CALERA - JUL_2026.pdf (24-JUL-2026)  |  "
          "Cronograma Rev B: 5155-00-0007-VE-CO-PG-001_REV B.mpp  |  OC ABB: 4508945953")
c.fill = fill(C_HEAD); c.font = font(size=8, italic=True)
c.alignment = center(); c.border = border_all()
_r9[0] += 1
w9_blank()

# ── A. DATOS DEL INFORME ─────────────────────────────────────────────────────
w9_section("A. DATOS DEL INFORME DE AVANCE PMS — JULIO 2026")
w9_kv("Número Informe", "E-2611027 — Informe de Avance CPF2 LA CALERA, Julio 2026")
w9_kv("Fecha Emisión", "24 de julio de 2026")
w9_kv("PM Responsable", "Pablo Kalis / V. Toledo (ABB)")
w9_kv("OC Principal", "4508945953 (PMS Sala 3 + Sala 4)")
w9_kv("Período Cubierto", "Julio 2026")
w9_kv("Cronograma Referencia", "5155-00-0007-VE-CO-PG-001_REV B  (Rev B — nuevo job: 4JAR001140P0003)")
w9_blank()

# ── B. ALCANCE PMS ───────────────────────────────────────────────────────────
w9_section("B. ALCANCE EQUIPOS PMS (OC 4508945953)")
w9_hdr(["TAG", "Descripción", "Cantidad", "Dim. (H×W×D mm)", "Ubicación", "Cert. Asociada", "Estado Fab."])

scope_pms = [
    ("PMS-001", "Tablero PMS Principal — Sala 3",
     "1 cuerpo", "2200×800×600", "Sala Eléctrica N°3 (BT)", "20% FAT", "NO INICIADA"),
    ("PMS-001 ext.", "Tablero PMS Ext. 1 — Sala 3",
     "1 cuerpo", "2200×800×600", "Sala Eléctrica N°3 (BT)", "20% FAT", "NO INICIADA"),
    ("PMS-001 ext.", "Tablero PMS Ext. 2 — Sala 3",
     "1 cuerpo", "2200×800×600", "Sala Eléctrica N°3 (BT)", "20% FAT", "NO INICIADA"),
    ("PMS-101", "Tablero PMS — Sala 4",
     "1 cuerpo", "2200×800×600", "Sala Eléctrica N°4 (MT)", "20% FAT", "NO INICIADA"),
]

for row_data in scope_pms:
    r = _r9[0]
    bg = "FFFFFF" if r % 2 == 0 else C_EVEN
    colors_list = [bg, bg, bg, bg, bg, bg, C_EVEN]
    w9_data(list(row_data), colors=colors_list)
w9_blank()

# ── C. EVENTOS JULIO 2026 ────────────────────────────────────────────────────
w9_section("C. ACTIVIDADES REALIZADAS — JULIO 2026")
w9_hdr(["N°", "Actividad Realizada", "Responsable", "Estado", "Fecha", "Impacto", "Observación"])

july_events = [
    ("1", "Aclaraciones técnicas PMS (IB en proceso)",
     "ABB", "COMPLETADO", "Jul-2026", "Bajo",
     "Definición de alcance y requerimientos técnicos IB"),
    ("2", "Listado de documentos a emitir acordado con AESA",
     "ABB/AESA", "COMPLETADO", "Jul-2026", "Medio",
     "Lista ID a emitir agosto 2026 definida y acordada"),
    ("3", "Reunión integral Pluspetrol–ABB–AESA",
     "PM Kalis", "COMPLETADO", "Jul-2026", "Alto",
     "Coordinación interfaz PMS con otros proveedores del proyecto"),
    ("4", "Definiciones fibra óptica Sala 3 y Sala 4",
     "ABB/AESA", "EN CURSO", "Jul-2026", "Alto",
     "Afecta topología de red PMS. Resolución esperada agosto"),
    ("5", "Avances en Ingeniería Básica (IB) PMS",
     "ABB", "EN CURSO", "Jul-2026", "Alto",
     "IB en progreso. Entrega formal prevista 17-JUL (con demora leve)"),
    ("6", "Job PMS actualizado a 4JAR001140P0003 (Rev B)",
     "ABB Mgmt", "COMPLETADO", "Jul-2026", "Medio",
     "Formalización cronograma Rev B con nuevo número de job"),
    ("7", "Emisión pólizas OC 4508945953",
     "ABB", "COMPLETADO", "Jul-2026", "Bajo",
     "Póliza fiel cumplimiento y fondo de reparo emitidas para OC PMS"),
]

event_colors = {
    "COMPLETADO": C_GREEN,
    "EN CURSO":   C_YELLOW,
    "PENDIENTE":  C_ORANGE,
    "BLOQUEADO":  C_RED,
}

for row_data in july_events:
    r = _r9[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    estado = row_data[3]
    bg_estado = event_colors.get(estado, bg_base)
    colors_list = [bg_base, bg_base, bg_base, bg_estado, bg_base, bg_base, bg_base]
    w9_data(list(row_data), colors=colors_list)
w9_blank()

# ── D. ACTIVIDADES AGOSTO 2026 ───────────────────────────────────────────────
w9_section("D. ACTIVIDADES PREVISTAS — AGOSTO 2026")
w9_hdr(["N°", "Actividad Prevista", "Responsable", "Tipo", "Fecha Límite", "Prioridad", "Condición"])

aug_events = [
    ("1", "Emisión Arquitectura/Topología de Red PMS (Rev B)",
     "ABB", "Ingeniería", "31/08/2026", "ALTA",
     "Nueva actividad Rev B. Insumo crítico para aprobación AESA"),
    ("2", "Descripción Funcional PMS",
     "ABB", "Ingeniería", "31/08/2026", "ALTA",
     "Nueva actividad Rev B. Insumo para FAT y aprobación"),
    ("3", "Emisión documentos ID (lista acordada con AESA)",
     "ABB", "Ingeniería", "09/10/2026", "ALTA",
     "Proceso iniciado agosto. Lista de ID acordada julio"),
    ("4", "Resolución fibra óptica S3/S4",
     "ABB/AESA", "Ingeniería", "31/08/2026", "ALTA",
     "Pendiente definición final. Afecta topología PMS"),
    ("5", "Aclaraciones hitos certificación PMS (análisis AESA)",
     "AESA", "Contractual", "31/08/2026", "MEDIA",
     "Presentadas 8-Jun-2026. En análisis AESA. Sin respuesta aún"),
]

prio_colors = {"ALTA": C_ORANGE, "MEDIA": C_YELLOW, "BAJA": C_GREEN}

for row_data in aug_events:
    r = _r9[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    prio = row_data[5]
    bg_prio = prio_colors.get(prio, bg_base)
    colors_list = [bg_base, bg_base, bg_base, bg_base, bg_base, bg_prio, bg_base]
    w9_data(list(row_data), colors=colors_list)
w9_blank()

# ── E. PENDIENTES DE AESA ────────────────────────────────────────────────────
w9_section("E. PENDIENTES DE AESA — Información Requerida para Avanzar PMS")
w9_hdr(["N°", "Información Pendiente de AESA", "Necesario Para", "Fecha Requerida", "Impacto", "Estado", "Escalado a"])

aesa_pending = [
    ("1",
     "Señales cableadas con borneras (arquitectura de red PMS)",
     "IB PMS — Aprobación AESA",
     "VENCIDO",
     "CRÍTICO — bloquea aprobación IB",
     "PENDIENTE AESA",
     "PM Pluspetrol"),
    ("2",
     "Mapa Modbus completo (UPS, Generadores, Multimedidores)",
     "IB PMS — Aprobación AESA",
     "VENCIDO",
     "CRÍTICO — bloquea aprobación IB",
     "PENDIENTE AESA",
     "PM Pluspetrol"),
    ("3",
     "Mapas Profinet (UMC y VFD de cada drive)",
     "IB PMS — Aprobación AESA",
     "VENCIDO",
     "CRÍTICO — bloquea aprobación IB",
     "PENDIENTE AESA",
     "PM Pluspetrol"),
    ("4",
     "Respuesta aclaraciones hitos certificación PMS (desde 8-Jun-2026)",
     "Certificación hito 1 PMS",
     "VENCIDO (>45 días)",
     "ALTO — retrasa flujo de caja PMS",
     "EN ANÁLISIS AESA",
     "Contractual"),
    ("5",
     "Definición final fibra óptica Sala 3 / Sala 4",
     "Arquitectura/Topología Red PMS (Rev B)",
     "31/08/2026",
     "ALTO — afecta ID y planos constructivos",
     "EN CURSO",
     "PM Kalis / Reunión integral"),
]

aesa_bg = {
    "PENDIENTE AESA":   C_RED,
    "EN ANÁLISIS AESA": C_ORANGE,
    "EN CURSO":         C_YELLOW,
}

for row_data in aesa_pending:
    r = _r9[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    estado = row_data[5]
    bg_estado = aesa_bg.get(estado, bg_base)
    colors_list = [bg_base, bg_base, bg_base, C_RED if "VENCIDO" in str(row_data[3]) else bg_base,
                   bg_base, bg_estado, bg_base]
    w9_data(list(row_data), colors=colors_list)
w9_blank()

# ── F. CAMBIOS REV A → REV B ─────────────────────────────────────────────────
w9_section("F. CAMBIOS CRONOGRAMA PMS: Rev A → Rev B  (5155-00-0007-VE-CO-PG-001)")
w9_hdr(["Tipo Cambio", "Descripción del Cambio", "Impacto en Fechas", "Impacto en Alcance",
        "Impacto en Costo", "Criticidad", "Acción Requerida"])

rev_changes = [
    ("ACTIVIDADES NUEVAS",
     "Arquitectura/Topología Red PMS + Descripción Funcional PMS",
     "Agregan tareas en agosto 2026, sin impacto en fechas fin",
     "AMPLIADO: documentación adicional ingeniería",
     "Incluido en OC 4508945953",
     "MEDIA",
     "Coordinar recursos ABB agosto"),
    ("AMPLIACIÓN ALCANCE",
     "FAT Maqueta Integrada: ahora incluye PMS/CCM/PCS de CPF1 Y CPF2",
     "Mismas fechas (22-25 Sep 2026) pero mayor complejidad logística",
     "AMPLIADO: coordinación entre 2 proyectos CPF1 y CPF2",
     "Posible impacto. Analizar con OC",
     "ALTA",
     "Confirmar disponibilidad CCM/PCS CPF1 en las mismas fechas"),
    ("JOB NUMBER",
     "Nuevo número job: 4JAR001140P0003 (vs RevA: 4JAR00XXXXP0001)",
     "Ninguno",
     "Formalización contractual",
     "Ninguno",
     "BAJA",
     "Actualizar referencias en documentación"),
    ("PENDIENTES AESA (NUEVO RIESGO)",
     "3 bloqueantes AESA detectados: señales, Modbus, Profinet",
     "⚠ Si AESA no responde antes de 14-AGO, FP-P1 se corre",
     "IB no puede ser aprobada sin datos AESA",
     "Riesgo de claims si demora imputable a AESA",
     "CRÍTICA",
     "Escalar formalmente a Pluspetrol con nota de demora"),
]

crit_colors = {"CRÍTICA": C_RED, "ALTA": C_ORANGE, "MEDIA": C_YELLOW, "BAJA": C_GREEN}

for row_data in rev_changes:
    r = _r9[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    crit = row_data[5]
    bg_crit = crit_colors.get(crit, bg_base)
    colors_list = [bg_base, bg_base, bg_base, bg_base, bg_base, bg_crit, bg_base]
    w9_data(list(row_data), colors=colors_list)
w9_blank()

# ── G. SEMÁFORO RESUMEN PMS ──────────────────────────────────────────────────
w9_section("G. SEMÁFORO RESUMEN PMS — Al 24-JUL-2026")
w9_hdr(["Área", "Ítem", "Estado", "Indicador", "Tendencia", "Último Evento", "Acción Próxima"])

semaforo_pms = [
    ("Ingeniería", "IB PMS en elaboración",
     "AMARILLO", "En curso, avances parciales",
     "→ estable", "Avances reportados JUL-2026",
     "Completar IB y entregar formalmente"),
    ("Aprobación", "AESA: 3 bloqueantes activos",
     "ROJO", "Bloquea aprobación IB completa",
     "↓ empeora", "Sin respuesta AESA al 24-JUL",
     "Escalar a Pluspetrol. Nota formal demora AESA"),
    ("Nuevas actividades", "Arq/Top Red + Desc.Funcional",
     "AMARILLO", "Identificadas Rev B, no iniciadas",
     "→ estable", "Incluidas en Rev B schedule",
     "Inicio agosto 2026"),
    ("FAT Maqueta", "Alcance ampliado CPF1+CPF2",
     "NARANJA", "Mayor complejidad logística",
     "↑ riesgo", "Detectado en Rev B",
     "Confirmar disponibilidad equipos CPF1"),
    ("Fabricación", "Tableros PMS-001 + PMS-101",
     "VERDE", "No crítica para 27-JUL",
     "→ estable", "Pólizas emitidas OC 4508945953",
     "Inicio fabricación: Sep-2026"),
    ("Certificaciones", "Ninguna emitida para OC PMS",
     "AMARILLO", "Aclaraciones en análisis AESA",
     "→ sin avance", "Presentadas 8-Jun-2026",
     "Seguir respuesta AESA — >45 días sin respuesta"),
    ("Contractual", "Pólizas OC 4508945953",
     "VERDE", "Emitidas correctamente",
     "↑ bien", "JUL-2026",
     "OK — sin acción requerida"),
]

sem_colors = {"ROJO": C_RED, "NARANJA": C_ORANGE, "AMARILLO": C_YELLOW,
              "VERDE": C_GREEN, "GRIS": "F2F2F2"}

for row_data in semaforo_pms:
    r = _r9[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    estado = row_data[2]
    bg_estado = sem_colors.get(estado, bg_base)
    colors_list = [bg_base, bg_base, bg_estado, bg_base, bg_base, bg_base, bg_base]
    w9_data(list(row_data), colors=colors_list)
w9_blank()

# Freeze top rows
ws9.freeze_panes = "A3"

# ── Step 5: Update sheet 8_Desvios_Alertas with PMS deviations ───────────────
ws8 = wb["8_Desvios_Alertas"]

# Find last row
last8 = ws8.max_row + 2
_r8 = [last8]

def w8_title(text, bg=C_PMS_HEAD, span=8):
    r = _r8[0]
    ws8.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws8.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, color="FFFFFF", size=11)
    c.alignment = center(); c.border = border_all()
    ws8.row_dimensions[r].height = 24
    _r8[0] += 1

def w8_section(text, bg=C_PMS_LIGHT, span=8):
    r = _r8[0]
    ws8.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws8.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, size=10)
    c.alignment = left(); c.border = border_thick()
    ws8.row_dimensions[r].height = 18
    _r8[0] += 1

def w8_hdr(cols):
    r = _r8[0]
    for ci, val in enumerate(cols, 1):
        c = ws8.cell(row=r, column=ci, value=val)
        c.fill = fill(C_HEAD); c.font = font(bold=True, size=9)
        c.alignment = center(wrap=True); c.border = border_all()
    ws8.row_dimensions[r].height = 30
    _r8[0] += 1

def w8_data(values, colors=None):
    r = _r8[0]
    bg = "FFFFFF" if r % 2 == 0 else C_EVEN
    for ci, val in enumerate(values, 1):
        c = ws8.cell(row=r, column=ci, value=val)
        bg_use = colors[ci-1] if colors and ci <= len(colors) else bg
        c.fill = fill(bg_use)
        c.font = font(size=9)
        c.alignment = left(wrap=True) if ci in (2, 3, 8) else center()
        c.border = border_all()
    _r8[0] += 1

def w8_blank():
    _r8[0] += 1

w8_blank()
w8_title(f"DESVÍOS Y ALERTAS PMS — {REV}  |  {FECHA}  (incorpora Informe Avance JUL-2026 + Cronograma Rev B)")

w8_blank()
w8_section("DESVÍOS ESPECÍFICOS PMS vs CRONOGRAMA BASE (5155-00-0007-VE-CO-PG-001_REV A)")
w8_hdr(["Ref", "Ítem / Actividad", "Desvío Detectado", "Rev A Plan",
        "Rev B / Real JUL-26", "Criticidad", "Tendencia", "Acción Recomendada"])

pms_desvios = [
    ("D-PMS-01",
     "IB PMS: demora en aprobación AESA",
     "3 ítems bloqueantes AESA sin resolver: señales cableadas, Mapa Modbus, Mapas Profinet",
     "Aprobación IB: 14/08/2026",
     "EN RIESGO — bloqueantes activos al 24-JUL",
     "CRÍTICO", "↓",
     "Nota formal a AESA/Pluspetrol. Solicitar respuesta antes 01/AGO"),
    ("D-PMS-02",
     "Aclaraciones hitos certificación: >45 días sin respuesta AESA",
     "Aclaraciones presentadas 08-JUN-2026. Sin respuesta al 24-JUL-2026",
     "Respuesta esperada <15 días",
     ">45 días sin respuesta",
     "ALTO", "↓",
     "Escalar a gerencia Pluspetrol. Potencial claim contractual"),
    ("D-PMS-03",
     "Fibra óptica S3/S4: definición pendiente",
     "Decisión de fibra óptica no tomada. Afecta topología red PMS",
     "Definición: Jun-2026",
     "En proceso JUL-2026 — sin fecha compromiso",
     "ALTO", "→",
     "Urgir resolución antes 31-AGO para no impactar ID"),
    ("D-PMS-04",
     "FAT Maqueta: alcance ampliado a CPF1+CPF2",
     "Rev B incluye integración con equipos CPF1 (PMS/CCM/PCS). Mayor complejidad",
     "Solo CPF2 (RevA)",
     "CPF1 + CPF2 (Rev B) — mismas fechas 22-25 Sep 2026",
     "MEDIO", "↑ riesgo",
     "Confirmar disponibilidad maqueta CPF1 con PM correspondiente"),
    ("D-PMS-05",
     "Nuevas actividades Rev B sin asignación confirmada",
     "Arq/Topología Red y Desc. Funcional: nuevas, inicio agosto 2026",
     "No existían en Rev A",
     "Identificadas Rev B — recursos por confirmar",
     "MEDIO", "→",
     "Confirmar asignación de recursos ABB agosto"),
]

crit_colors2 = {"CRÍTICO": C_RED, "ALTO": C_ORANGE, "MEDIO": C_YELLOW, "BAJO": C_GREEN}

for row_data in pms_desvios:
    r = _r8[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    crit = row_data[5]
    bg_crit = crit_colors2.get(crit, bg_base)
    colors_list = [bg_base, bg_base, bg_base, bg_base, bg_base, bg_crit, bg_base, bg_base]
    w8_data(list(row_data), colors=colors_list)

w8_blank()
w8_section("RESUMEN ALERTAS PMS — Rev 4  (señales tempranas de desvío para el objetivo S#3 14-MAY-2027)")
w8_hdr(["Prioridad", "Alerta", "Impacto en Objetivo", "Fecha Límite Acción",
        "Responsable", "Estado", "Tendencia", "Acción Inmediata"])

alertas_pms = [
    ("🔴 P1",
     "AESA: 3 bloqueantes para aprobación IB PMS",
     "Retraso aprobación IB → impacta ID → posible corrimiento FAT Maqueta Sep/26",
     "01-AGO-2026",
     "PM Kalis / Pluspetrol",
     "ACTIVA",
     "↓ empeorando",
     "Nota formal de demora + solicitud fecha respuesta AESA"),
    ("🔴 P1",
     "Aclaraciones hitos cert. PMS: >45 días sin respuesta",
     "Sin hitos aprobados: no habrá flujo de caja PMS hasta resolución",
     "INMEDIATO",
     "Contractual ABB / Pluspetrol",
     "ACTIVA",
     "↓",
     "Escalar a gerencia. Documentar días de demora AESA"),
    ("🟠 P2",
     "Fibra óptica S3/S4: sin definición final",
     "Afecta Arquitectura Red (Rev B). Si no se define antes 31-AGO → retraso ID",
     "31-AGO-2026",
     "ABB / AESA",
     "EN SEGUIMIENTO",
     "→",
     "Urgir resolución en reunión mensual agosto"),
    ("🟠 P2",
     "FAT Maqueta ampliado CPF1+CPF2 (22-25 Sep 2026)",
     "Mayor complejidad logística. Riesgo de reprogramación si CPF1 no está listo",
     "15-AGO-2026",
     "PM Kalis / PM CPF1",
     "EN ANÁLISIS",
     "↑ riesgo",
     "Reunión coordinación inter-proyecto CPF1/CPF2"),
    ("🟡 P3",
     "Nuevas actividades Rev B: Arq/Top + Desc.Funcional (agosto)",
     "Sin impacto en fechas si recursos disponibles en agosto",
     "31-AGO-2026",
     "ABB Ingeniería",
     "MONITOREO",
     "→",
     "Confirmar asignación recursos agosto"),
]

alert_colors = {"🔴 P1": C_RED, "🟠 P2": C_ORANGE, "🟡 P3": C_YELLOW, "🟢 P4": C_GREEN}
alert_estado = {"ACTIVA": C_RED, "EN SEGUIMIENTO": C_YELLOW,
                "EN ANÁLISIS": C_ORANGE, "MONITOREO": C_EVEN, "RESUELTA": C_GREEN}

for row_data in alertas_pms:
    r = _r8[0]
    bg_base = "FFFFFF" if r % 2 == 0 else C_EVEN
    prio = row_data[0]
    estado = row_data[5]
    bg_prio = alert_colors.get(prio, bg_base)
    bg_est = alert_estado.get(estado, bg_base)
    colors_list = [bg_prio, bg_base, bg_base, bg_base, bg_base, bg_est, bg_base, bg_base]
    w8_data(list(row_data), colors=colors_list)

# ── Step 6: Save ─────────────────────────────────────────────────────────────
wb.save(DST)
print(f"✓ Generado: {DST}")
print(f"  Sheets: {wb.sheetnames}")
print(f"  Hoja 9_Estado_PMS_JUL2026: {ws9.max_row} filas")
print(f"  Hoja 2_PMS_Detalle: actualizada con Rev B")
print(f"  Hoja 8_Desvios_Alertas: ampliada con desvíos PMS")
