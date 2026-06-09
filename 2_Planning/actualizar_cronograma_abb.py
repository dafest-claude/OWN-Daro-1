#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ACTUALIZACIÓN CRONOGRAMA GENERAL ABB — 09-JUN-2026
Fuentes:
  - PG-001 Rev A (5155-00-0009-VE-CO-PG-001_A.mpp/.pdf)  ABB 26-MAY-2026
  - R-2631036 Cronograma de Certificaciones_A.xlsx          ABB 05-JUN-2026
  - R-2631036 Informe de Avance N°1_A (Mayo 2026)          ABB 04-JUN-2026
Cambios vs Preliminar 27-MAY:
  1. Cronograma oficial PG-001 Rev A incorporado
  2. Informe Avance N°1 Mayo 2026: estado de actividades al 05-JUN
  3. Pendientes AESA: 4 items críticos con fechas límite
  4. Cronograma de Certificaciones mensual completo (todos los hitos)
  5. Entrega Sala 4: 08-ABR-2027 | Sala 3: 14-MAY-2027 (confirmadas)
"""

import os, shutil
from datetime import date, timedelta
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))
FECHA_DOC    = "09-JUN-2026"
FECHA_ANT    = "27-MAY-2026"

OUTPUT = os.path.join(PLANNING_DIR,
                      f"CRONOGRAMA_GENERAL_ABB_Actualizado_2026-06-09.xlsx")
BASE   = os.path.join(PLANNING_DIR,
                      "CRONOGRAMA_GENERAL_ABB_Preliminar_2026-05-27.xlsx")

# ── HELPERS ──────────────────────────────────────────────────────────────────
def fill(c): return PatternFill("solid", fgColor=c)
def fnt(bold=False, color="000000", size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")
def aln(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def brd(style="thin"):
    s = Side(style=style)
    return Border(left=s, right=s, top=s, bottom=s)

def cell(ws, row, col, value=None, bold=False, color="000000", bg=None,
         size=10, h="center", v="center", wrap=False, italic=False,
         border=True, num_format=None):
    c = ws.cell(row=row, column=col, value=value)
    c.font      = fnt(bold=bold, color=color, size=size, italic=italic)
    c.alignment = aln(h=h, v=v, wrap=wrap)
    if bg:  c.fill   = fill(bg)
    if border: c.border = brd()
    if num_format: c.number_format = num_format
    return c

# ── LOAD BASE WORKBOOK AND UPDATE TITLES ─────────────────────────────────────
shutil.copy2(BASE, OUTPUT)
wb = load_workbook(OUTPUT)

# Update title row in existing sheets where FECHA_DOC appears
for ws in wb.worksheets:
    for row in ws.iter_rows():
        for c in row:
            if c.value and isinstance(c.value, str) and FECHA_ANT in c.value:
                c.value = c.value.replace(
                    FECHA_ANT,
                    f"{FECHA_ANT} → ACTUALIZADO {FECHA_DOC}"
                )
            if c.value and isinstance(c.value, str) and "PRELIMINAR" in c.value:
                c.value = c.value.replace("PRELIMINAR", "ACTUALIZADO")

# ─────────────────────────────────────────────────────────────────────────────
# NUEVA HOJA 6 — ESTADO Y PENDIENTES (al 09-JUN-2026)
# ─────────────────────────────────────────────────────────────────────────────
C_RED  = "CC0000"; C_DRD = "880000"
C_BLU  = "2E75B6"; C_DBL = "1F4E79"
C_GRN  = "375623"; C_ORG = "E36C09"
C_PURP = "7030A0"; C_YEL = "FFEB9C"
C_WARN = "FF9900"; C_RLT = "FFE0E0"
C_BLT  = "DEEAF1"; C_GLT = "E2EFDA"
C_WRN  = "FFC7CE"; C_OK  = "C6EFCE"
C_WNT  = "9C5700"; C_OKT = "375623"
C_GRY  = "595959"; C_WH  = "FFFFFF"

ws6 = wb.create_sheet("6_Estado_JUN2026")
ws6.sheet_view.showGridLines = False
ws6.column_dimensions["A"].width = 12
ws6.column_dimensions["B"].width = 36
ws6.column_dimensions["C"].width = 20
ws6.column_dimensions["D"].width = 14
ws6.column_dimensions["E"].width = 14
ws6.column_dimensions["F"].width = 40

# Title
ws6.merge_cells("A1:F1")
c = ws6["A1"]
c.value = f"ESTADO DEL PROYECTO — INFORME DE AVANCE N°1  |  ABB  |  Período: Mayo 2026  |  Actualizado: {FECHA_DOC}"
c.font      = fnt(bold=True, color=C_WH, size=13)
c.alignment = aln("center","center")
c.fill      = fill(C_RED)
c.border    = brd("medium")
ws6.row_dimensions[1].height = 28

ws6.merge_cells("A2:F2")
c = ws6["A2"]
c.value = ("Fuente: R-2631036 Informe de Avance N°1 (04-JUN-2026)  |  "
           "Cronograma Oficial: 5155-00-0009-VE-CO-PG-001 Rev A  |  "
           "R-2631036 Cronograma de Certificaciones Rev A  |  "
           "PM ABB: Rodrigo Mack")
c.font      = fnt(color=C_WH, size=9)
c.alignment = aln("center","center")
c.fill      = fill(C_GRY)
c.border    = brd()
ws6.row_dimensions[2].height = 16

# ── SECTION A: RESUMEN ESTADO ─────────────────────────────────────────────────
r = 4
ws6.merge_cells(f"A{r}:F{r}")
cell(ws6, r, 1, "A. RESUMEN DE ESTADO — AVANCE MAYO 2026", bold=True,
     color=C_WH, bg=C_DRD, size=11, h="left", border=True)
ws6.row_dimensions[r].height = 20
r += 1

estado_items = [
    ("Ingeniería", "Sala #3 (BT) + Sala #4 (MT)", "EN CURSO", "✓",
     "Ingeniería Básica iniciada en todos los equipos. RevA en elaboración"),
    ("Ingeniería", "Básica PMS (800XA/IEC61850)", "EN CURSO", "✓",
     "Arquitectura+I/O+Planos Constructivos RevA iniciados"),
    ("Procura", "Pedidos internos / OC a 3ros", "PENDIENTE", "→",
     "Aguardan aprobación IB. OC Turquía/China/Finlandia no emitidas aún"),
    ("Fabricación", "CCMs + Tableros + Celdas MT", "NO INICIADA", "—",
     "Sin actividades de fabricación en el período Mayo 2026"),
    ("Calidad", "Planes de Inspección", "EN ELABORACIÓN", "→",
     "En preparación para CCMs, tableros MT/BT, VFDs"),
    ("Hitos Pago", "Certificación IB envío", "PREVISTO", "→",
     "Primer pago S3+S4: Jun 2026. Ver hoja 6_Cert_Mensual"),
    ("Cronograma", "PG-001 Rev A publicado", "COMPLETADO", "✓",
     "5155-00-0009-VE-CO-PG-001_A.mpp emitido 26-MAY-2026 por PM Rodrigo Mack"),
    ("Cronograma", "Informe de Avance N°1", "COMPLETADO", "✓",
     "R-2631036 Informe Avance N°1 emitido 04-JUN-2026"),
]

hdr_colors_e = {"COMPLETADO": C_GRN, "EN CURSO": C_BLU, "PENDIENTE": C_WARN,
                "EN ELABORACIÓN": C_ORG, "NO INICIADA": C_GRY, "PREVISTO": C_PURP}

for hdr_cols in [("Área","OC / Ítem","Estado","","Observación","")]:
    for i, h in enumerate(["Área","OC / Ítem","Estado","","Observación",""], 1):
        cell(ws6, r, i, h, bold=True, color=C_WH, bg=C_DBL, size=9)
    ws6.row_dimensions[r].height = 16
    r += 1

for area, item, estado, mark, obs in estado_items:
    status_bg = {"COMPLETADO":"C6EFCE","EN CURSO":"DEEAF1","PENDIENTE":C_YEL,
                 "EN ELABORACIÓN":"FDEBD0","NO INICIADA":"F2F2F2","PREVISTO":"EAD1DC"}.get(estado,"FFFFFF")
    cell(ws6, r, 1, area, size=9, h="left", bg="F8F8F8")
    cell(ws6, r, 2, item, size=9, h="left")
    cell(ws6, r, 3, estado, bold=True, size=9, bg=status_bg,
         color={"COMPLETADO":C_OKT,"EN CURSO":C_DBL,"PENDIENTE":C_WNT,
                "EN ELABORACIÓN":"7B4000","NO INICIADA":C_GRY,"PREVISTO":"5B0069"}.get(estado,"000000"))
    cell(ws6, r, 4, mark, bold=True, size=11,
         color=C_GRN if mark=="✓" else (C_WARN if mark=="→" else C_GRY))
    ws6.merge_cells(f"E{r}:F{r}")
    cell(ws6, r, 5, obs, size=9, h="left", wrap=True)
    ws6.row_dimensions[r].height = 22
    r += 1

r += 1

# ── SECTION B: PENDIENTES CRÍTICOS AESA ──────────────────────────────────────
ws6.merge_cells(f"A{r}:F{r}")
cell(ws6, r, 1, "B. PENDIENTES CRÍTICOS — INFORMACIÓN REQUERIDA A AESA", bold=True,
     color=C_WH, bg="8B0000", size=11, h="left")
ws6.row_dimensions[r].height = 20
r += 1

for i, h in enumerate(["Fecha Límite","Descripción Requerimiento","Resp.","Estado",
                        "Impacto en Cronograma","Ref. IB"], 1):
    cell(ws6, r, i, h, bold=True, color=C_WH, bg=C_DRD, size=9)
ws6.row_dimensions[r].height = 16
r += 1

pendientes_aesa = [
    (date(2026,6,5),
     "Definición de típicos funcionales para cada salida a motor (reservas equipadas incluidas)",
     "AESA", "VENCIDA" if date.today() > date(2026,6,5) else "PENDIENTE",
     "FP-S3-1 (03-JUL) y FP-S4-1 (26-JUN) requieren IB aprobada",
     "Esquemas CCM / Planos Constructivos"),
    (date(2026,6,18),
     "Detalles de acometida de transformadores MT y posiciones relativas para ducto de barras",
     "AESA", "PENDIENTE",
     "Inicio OC Ductos Barras Turquía (lead 6m). FP-S3-1 en riesgo",
     "DB-005/006/007 planos instalación"),
    (date(2026,6,30),
     "Típicos y definición de accesorios. Suministro equipos físicos según cronograma de cada sala",
     "AESA", "PENDIENTE",
     "Cierre Ingeniería Básica S3 y S4. Gatilla pagos Hito 1+2",
     "Planos constructivos tableros"),
    (date(2026,7,20),
     "Confirmación de equipos para sistema de control de acceso (canalizaciones, puertas sala)",
     "AESA", "PENDIENTE",
     "Ingeniería de detalle sala eléctrica (FP-S3-3 / FP-S4-3 → 28-31 AGO)",
     "Planos instalación eléctrica sala"),
]

for fec, desc, resp, estado, impacto, ref in pendientes_aesa:
    days_left = (fec - date.today()).days
    if days_left < 0:
        row_bg = C_WRN; status_txt = f"⚠ VENCIDA ({abs(days_left)}d)"
    elif days_left <= 7:
        row_bg = C_YEL; status_txt = f"⚡ URGENTE ({days_left}d)"
    elif days_left <= 21:
        row_bg = "FFF0CD"; status_txt = f"→ {days_left}d"
    else:
        row_bg = "F8F8F8"; status_txt = f"{days_left}d"

    cell(ws6, r, 1, fec, bold=(days_left <= 7), size=9, num_format="DD-MMM-YYYY",
         color="CC0000" if days_left < 0 else "000000")
    cell(ws6, r, 2, desc, size=9, h="left", wrap=True, bg=row_bg)
    cell(ws6, r, 3, resp, bold=True, size=9, bg=C_BLT, color=C_DBL)
    cell(ws6, r, 4, status_txt, bold=True, size=9,
         bg=C_WRN if days_left < 0 else (C_YEL if days_left <= 7 else None),
         color="CC0000" if days_left < 0 else C_WNT)
    cell(ws6, r, 5, impacto, size=9, h="left", wrap=True, bg=row_bg)
    cell(ws6, r, 6, ref, size=9, h="left", italic=True, wrap=True, bg=row_bg)
    ws6.row_dimensions[r].height = 32
    r += 1

r += 1

# ── SECTION C: ALCANCE CONFIRMADO (del Informe) ───────────────────────────────
ws6.merge_cells(f"A{r}:F{r}")
cell(ws6, r, 1, "C. ALCANCE CONFIRMADO — SALAS ELÉCTRICAS (según Informe N°1)", bold=True,
     color=C_WH, bg=C_DBL, size=11, h="left")
ws6.row_dimensions[r].height = 20
r += 1

alcance = [
    ("SALA #3 BT",
     "Sala eléctrica modular 37.000×9.000×3.700mm (LxAxA)",
     "Tableros 102-CCM-005/6/7 + VFDs 102-DP-VFD-24210A/B/C + Ductos 102-DB-005A/B / 006A/B / 007A/B",
     "14-MAY-2027", "OC 4508944971  USD 8.13M", C_RLT),
    ("SALA #4 MT",
     "Sala eléctrica modular 28.000×5.000×3.500mm (LxAxA)",
     "Tableros MT 102-TGMT-001/002 + VFDs MT 102-DP-VFD-23310A/B/C + Tablero auxiliar 102-DP-101",
     "08-ABR-2027", "OC 4508944973  USD 3.69M", C_BLT),
    ("PMS",
     "Sistema de Gestión de Energía 800XA + IEC61850 + PROFINET + MODBUS TCP/IP",
     "Tableros PMS-001 (SE#3) + PMS-101 (SE#4). FAT Integral Shelter 25-31-MAR-2027",
     "17-FEB-2027", "OC 4508945953  Incluido", C_YEL),
]

for i, h in enumerate(["Ítem","Dimensiones","Equipamiento Principal","Fecha Entrega",
                        "OC / Monto",""], 1):
    cell(ws6, r, i, h, bold=True, color=C_WH, bg=C_DBL, size=9)
ws6.row_dimensions[r].height = 16
r += 1

for item, dim, equip, fecha, oc, bg in alcance:
    cell(ws6, r, 1, item, bold=True, size=9, bg=bg)
    cell(ws6, r, 2, dim,  size=9, h="left", wrap=True, bg=bg)
    cell(ws6, r, 3, equip, size=9, h="left", wrap=True, bg=bg)
    cell(ws6, r, 4, fecha, bold=True, size=9, bg=bg)
    cell(ws6, r, 5, oc, size=9, bg=bg)
    cell(ws6, r, 6, "", bg=bg)
    ws6.row_dimensions[r].height = 30
    r += 1

ws6.freeze_panes = "A4"

# ─────────────────────────────────────────────────────────────────────────────
# NUEVA HOJA 7 — CRONOGRAMA CERTIFICACIONES MENSUAL
# ─────────────────────────────────────────────────────────────────────────────
ws7 = wb.create_sheet("7_Cert_Mensual_Hitos")
ws7.sheet_view.showGridLines = False

# Months from data
MONTHS = [
    date(2026,5,1), date(2026,6,1), date(2026,7,1), date(2026,8,1),
    date(2026,9,1), date(2026,10,1),date(2026,11,1),date(2026,12,1),
    date(2027,1,1), date(2027,2,1), date(2027,3,1), date(2027,4,1),
    date(2027,5,1), date(2027,6,1), date(2027,7,1),
]

# Column mapping: col 9 (idx 8) = May 2026
# Items data: (oc, pos, tag, desc, price, monthly_payments[15])
ITEMS_CERT = [
    # OC S3 (45089944971) - items 70-270
    # (oc, pos, tag, desc, price, [15 monthly amounts])
    ("S3","70","VFDs 102-DP-VFD-24210A/B/C","VFD BT Bombas",242508.12,
     [0,0,24250.81,0,0,0,0,133379.47,0,0,0,0,72752.44,12125.41,0]),
    ("S3","80","102-DP-001-TRA","Tab. Distribución Tracing",95632.73,
     [0,9563.27,9563.27,0,0,0,23908.18,19126.55,0,0,0,0,28689.82,4781.64,0]),
    ("S3","90","102-DP-001","Tab. Serv. Aux. SE#3",105885.06,
     [0,10588.51,10588.51,0,0,0,26471.27,21177.01,0,0,0,0,31765.52,5294.25,0]),
    ("S3","100","102-LP-002","Tab. Distribución UPS",63485.70,
     [0,6348.57,6348.57,0,0,0,15871.43,12697.14,0,0,0,0,19045.71,3174.29,0]),
    ("S3","110","102-LP-001","Tab. Distribución Serv. Aux.",96409.77,
     [0,9640.98,9640.98,0,0,0,24102.44,19281.95,0,0,0,0,28922.93,4820.49,0]),
    ("S3","120","102-DB-007A/B","Ducto Barras DB-007",169501.43,
     [0,0,16950.14,16950.14,0,0,0,33900.29,0,42375.36,0,0,50850.43,8475.07,0]),
    ("S3","130","102-DB-006A/B","Ducto Barras DB-006",182075.08,
     [0,0,18207.51,18207.51,0,0,0,36415.02,0,45518.77,0,0,54622.52,9103.75,0]),
    ("S3","140","102-DB-005A/B","Ducto Barras DB-005",139094.32,
     [0,0,13909.43,13909.43,0,0,0,27818.86,0,34773.58,0,0,41728.30,6954.72,0]),
    ("S3","150","102-CCM-007 (lote 1)","CCM 380/220V lote 1",451298.47,
     [0,0,0,0,0,0,0,293344.01,0,0,0,0,135389.54,22564.92,0]),
    ("S3","160","102-CCM-006 (lote 1)","CCM 380/220V lote 1",492942.64,
     [0,0,0,0,0,0,0,320412.72,0,0,0,0,147882.79,24647.13,0]),
    ("S3","170","102-CCM-005 (lote 1)","CCM 380/220V lote 1",552737.98,
     [0,0,0,0,0,0,0,359279.69,0,0,0,0,165821.39,27636.90,0]),
    ("S3","180","SE#3 MT (lote 1)","Sala Eléctrica N°3 BT",622817.07,
     [0,0,62281.71,62281.71,0,0,0,155704.27,0,0,0,124563.41,186845.12,31140.85,0]),
    ("S3","190","102-UPS-001A/B","Sistema UPS CA",550434.32,
     [0,55043.43,55043.43,0,0,0,137608.58,110086.86,0,0,0,0,165130.30,27521.72,0]),
    ("S3","200","102-VCC-001A/B","Sistema CC VCC",274833.91,
     [0,27483.39,27483.39,0,0,0,68708.48,54966.78,0,0,0,0,82450.17,13741.70,0]),
    ("S3","210","102-CCM-007 (lote 2)","CCM 380/220V lote 2",451298.47,
     [0,0,0,0,0,0,0,293344.01,0,0,0,0,135389.54,22564.92,0]),
    ("S3","220","102-CCM-006 (lote 2)","CCM 380/220V lote 2",492942.64,
     [0,0,0,0,0,0,0,320412.72,0,0,0,0,147882.79,24647.13,0]),
    ("S3","230","102-CCM-005 (lote 2)","CCM 380/220V lote 2",552737.98,
     [0,0,0,0,0,0,0,359279.69,0,0,0,0,165821.39,27636.90,0]),
    ("S3","240","SE#3 MT (lote 2)","Sala Eléctrica N°3 BT",622817.07,
     [0,0,62281.71,62281.71,0,0,0,155704.27,0,0,0,124563.41,186845.12,31140.85,0]),
    ("S3","250","SE#3 MT (lote 3)","Sala Eléctrica N°3 BT",622817.07,
     [0,0,62281.71,62281.71,0,0,0,155704.27,0,0,0,124563.41,186845.12,31140.85,0]),
    ("S3","260","SE#3 MT (lote 4)","Sala Eléctrica N°3 BT",622817.07,
     [0,0,62281.71,62281.71,0,0,0,155704.27,0,0,0,124563.41,186845.12,31140.85,0]),
    ("S3","270","SE#3 MT (lote 5)","Sala Eléctrica N°3 BT",622817.07,
     [0,0,62281.71,62281.71,0,0,0,155704.27,0,0,0,124563.41,186845.12,31140.85,0]),
    # OC S4 (45089944973)
    ("S4","60","102-DP-101UPS","Tab. Distribución UPS SE#4",25305.57,
     [0,2530.56,2530.56,0,0,0,6326.39,5061.11,0,0,0,7591.67,0,1265.28,0]),
    ("S4","70","102-DP-101VCC","Tab. Distribución VCC SE#4",29291.23,
     [0,2929.12,2929.12,0,0,0,7322.81,5858.25,0,0,0,8787.37,0,1464.56,0]),
    ("S4","80","102-DP-101","Tab. Serv. Aux. SE#4",66110.81,
     [0,6611.08,6611.08,0,0,0,16527.70,13222.16,0,0,19833.24,0,0,3305.54,0]),
    ("S4","90","VFDs 102-DP-VFD-23310A/B/C","VFD MT Bombas",396952.56,
     [0,0,39695.26,0,0,0,0,218323.91,0,0,0,119085.77,0,19847.63,0]),
    ("S4","100","102-TGMT-001/003 celdas","SWG MT 13.2/6.6kV Celdas",517484.58,
     [0,51748.46,51748.46,0,0,0,103496.92,129371.15,0,0,0,155245.37,0,25874.23,0]),
    ("S4","110","102-TGMT-001/003 tab.1","SWG MT Tableros lote 1",386434.73,
     [0,38643.47,38643.47,0,0,0,77286.95,96608.68,0,0,0,115930.42,0,19321.74,0]),
    ("S4","120","SE#4 MT (lote 1)","Sala Eléctrica N°4 MT",605581.31,
     [0,60558.13,60558.13,0,0,0,0,151395.33,0,0,121116.26,181674.39,0,30279.07,0]),
    ("S4","130","102-TGMT-001/003 tab.2","SWG MT Tableros lote 2",386434.73,
     [0,38643.47,38643.47,0,0,0,77286.95,96608.68,0,0,0,115930.42,0,19321.74,0]),
    ("S4","140","SE#4 MT (lote 2)","Sala Eléctrica N°4 MT",605581.31,
     [0,60558.13,60558.13,0,0,0,0,151395.33,0,0,121116.26,181674.39,0,30279.07,0]),
    ("S4","150","SE#4 MT (lote 3)","Sala Eléctrica N°4 MT",605581.31,
     [0,60558.13,60558.13,0,0,0,0,151395.33,0,0,121116.26,181674.39,0,30279.07,0]),
]

# Title
ws7.merge_cells("A1:T1")
c7 = ws7["A1"]
c7.value = ("CRONOGRAMA DE CERTIFICACIONES MENSUAL — OC S3 + OC S4  |  "
            f"Actualizado: {FECHA_DOC}  |  Fuente: R-2631036 Cron. Certificaciones Rev A")
c7.font      = fnt(bold=True, color=C_WH, size=13)
c7.alignment = aln("center","center")
c7.fill      = fill(C_DBL)
c7.border    = brd("medium")
ws7.row_dimensions[1].height = 28

ws7.merge_cells("A2:T2")
c7b = ws7["A2"]
c7b.value = ("Hito 1:10% envío IB  |  Hito 2:10% aprob. IB  |  Hito 3:20% FAT  |  "
             "Hito 4:25% Acopio  |  Hito 5:30% Entrega  |  Hito 6:5% DataBook  |  "
             "Hito 7:100% Repuestos (separado)  |  Totales excluyen repuestos (ítems 10-60)")
c7b.font      = fnt(color=C_WH, size=9)
c7b.alignment = aln("center","center")
c7b.fill      = fill("404040")
c7b.border    = brd()
ws7.row_dimensions[2].height = 16

# Column widths
ws7.column_dimensions["A"].width = 5
ws7.column_dimensions["B"].width = 22
ws7.column_dimensions["C"].width = 26
ws7.column_dimensions["D"].width = 12
for i in range(15):
    ws7.column_dimensions[get_column_letter(5+i)].width = 10
ws7.column_dimensions[get_column_letter(20)].width = 12

# Headers row 3 – MONTHS
cell(ws7, 3, 1, "OC",  bold=True, color=C_WH, bg=C_DBL, size=8)
cell(ws7, 3, 2, "Pos / TAG", bold=True, color=C_WH, bg=C_DBL, size=8)
cell(ws7, 3, 3, "Descripción", bold=True, color=C_WH, bg=C_DBL, size=8)
cell(ws7, 3, 4, "Precio OC\n(USD)", bold=True, color=C_WH, bg=C_DBL, size=8, wrap=True)
ws7.row_dimensions[3].height = 24

month_colors = {5:"8B0000",6:"CC0000",7:"C55A11",8:"E36C09",9:"375623",
                10:"1F4E79",11:"1F4E79",12:"1F3864",
                1:"375623",2:"375623",3:"375623",4:"375623",
                -1:"375623",-2:"375623",-3:"375623"}  # months

for i, m in enumerate(MONTHS):
    col = 5 + i
    mname = m.strftime("%b'%y").upper()
    mc = month_colors.get(m.month, C_DBL)
    cell(ws7, 3, col, mname, bold=True, color=C_WH, bg=mc, size=8)

cell(ws7, 3, 20, "TOTAL\nCert.", bold=True, color=C_WH, bg=C_GRN, size=8, wrap=True)

# Data rows
r7 = 4
prev_oc = None
s3_totals = [0]*15
s4_totals = [0]*15

for oc, pos, tag, desc, price, monthly in ITEMS_CERT:
    # OC separator
    if oc != prev_oc:
        oc_color = C_DRD if oc == "S3" else C_DBL
        ws7.merge_cells(f"A{r7}:T{r7}")
        oc_label = (f"OC S3 — 4508944971  |  Sala 3 BT  |  Total: USD {sum(it[4] for it in ITEMS_CERT if it[0]=='S3'):,.0f}"
                    if oc == "S3" else
                    f"OC S4 — 4508944973  |  Sala 4 MT  |  Total: USD {sum(it[4] for it in ITEMS_CERT if it[0]=='S4'):,.0f}")
        cell(ws7, r7, 1, oc_label, bold=True, color=C_WH, bg=oc_color, size=10, h="left")
        ws7.row_dimensions[r7].height = 20
        prev_oc = oc
        r7 += 1

    row_bg = "FFF5F5" if oc == "S3" else "F0F8FF"
    cell(ws7, r7, 1, oc,   bold=True, size=8, bg=row_bg,
         color=C_DRD if oc=="S3" else C_DBL)
    cell(ws7, r7, 2, f"Pos {pos}\n{tag}", size=8, h="left", wrap=True, bg=row_bg)
    cell(ws7, r7, 3, desc, size=8, h="left", bg=row_bg)
    cell(ws7, r7, 4, price, size=8, bg=row_bg, num_format='$#,##0')
    row_total = 0
    for i, amt in enumerate(monthly):
        col = 5 + i
        if amt:
            oc_color_bg = "FFE0E0" if oc=="S3" else "DEEAF1"
            cell(ws7, r7, col, amt, size=8, bg=oc_color_bg, num_format='$#,##0')
            row_total += amt
            if oc == "S3": s3_totals[i] += amt
            else:           s4_totals[i] += amt
        else:
            cell(ws7, r7, col, None, size=8, bg=None)
    cell(ws7, r7, 20, row_total if row_total else None, bold=True, size=8,
         bg="E2EFDA" if row_total else None, num_format='$#,##0')
    ws7.row_dimensions[r7].height = 24
    r7 += 1

# S3 subtotal
r7 += 1
ws7.merge_cells(f"A{r7}:C{r7}")
cell(ws7, r7, 1, "SUBTOTAL OC S3 — MENSUAL (USD)", bold=True, color=C_WH,
     bg=C_DRD, size=9, h="left")
cell(ws7, r7, 4, sum(s3_totals), bold=True, color=C_WH, bg=C_DRD,
     size=9, num_format='$#,##0')
for i, amt in enumerate(s3_totals):
    col = 5 + i
    cell(ws7, r7, col, amt if amt else None, bold=True, color=C_WH, bg=C_DRD,
         size=9, num_format='$#,##0')
cell(ws7, r7, 20, sum(s3_totals), bold=True, color=C_WH, bg=C_DRD,
     size=9, num_format='$#,##0')
ws7.row_dimensions[r7].height = 18

r7 += 1
ws7.merge_cells(f"A{r7}:C{r7}")
cell(ws7, r7, 1, "SUBTOTAL OC S4 — MENSUAL (USD)", bold=True, color=C_WH,
     bg=C_DBL, size=9, h="left")
cell(ws7, r7, 4, sum(s4_totals), bold=True, color=C_WH, bg=C_DBL,
     size=9, num_format='$#,##0')
for i, amt in enumerate(s4_totals):
    col = 5 + i
    cell(ws7, r7, col, amt if amt else None, bold=True, color=C_WH, bg=C_DBL,
         size=9, num_format='$#,##0')
cell(ws7, r7, 20, sum(s4_totals), bold=True, color=C_WH, bg=C_DBL,
     size=9, num_format='$#,##0')
ws7.row_dimensions[r7].height = 18

# Combined total
r7 += 1
comb_totals = [s3_totals[i] + s4_totals[i] for i in range(15)]
ws7.merge_cells(f"A{r7}:C{r7}")
cell(ws7, r7, 1, "TOTAL COMBINADO S3+S4 — MENSUAL (USD)", bold=True, color=C_WH,
     bg=C_GRN, size=10, h="left")
cell(ws7, r7, 4, sum(comb_totals), bold=True, color=C_WH, bg=C_GRN,
     size=10, num_format='$#,##0')
for i, amt in enumerate(comb_totals):
    col = 5 + i
    cell(ws7, r7, col, amt if amt else None, bold=True, color=C_WH, bg=C_GRN,
         size=10, num_format='$#,##0')
cell(ws7, r7, 20, sum(comb_totals), bold=True, color=C_WH, bg=C_GRN,
     size=10, num_format='$#,##0')
ws7.row_dimensions[r7].height = 20

# Acumulado row
r7 += 1
acum = 0
ws7.merge_cells(f"A{r7}:C{r7}")
cell(ws7, r7, 1, "ACUMULADO (Curva S) — USD", bold=True, size=9, bg="FFF8DC", h="left")
cell(ws7, r7, 4, "", bg="FFF8DC")
for i, amt in enumerate(comb_totals):
    acum += amt
    col = 5 + i
    cell(ws7, r7, col, acum, bold=True, size=9, bg="FFF8DC", num_format='$#,##0')
cell(ws7, r7, 20, acum, bold=True, size=9, bg="FFF8DC", num_format='$#,##0')
ws7.row_dimensions[r7].height = 18

ws7.freeze_panes = "E4"

# ── REORDER SHEETS ────────────────────────────────────────────────────────────
# Sheets should be: original 1-5, then 6, 7
desired = ["1_Gantt_General","2_PMS_Detalle","3_Compatibilidad_Shelter",
           "4_Freezing_Points","5_Hitos_Pago_por_OC",
           "6_Estado_JUN2026","7_Cert_Mensual_Hitos"]
for i, name in enumerate(desired):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i:
            wb.move_sheet(name, offset=i - idx)

# ── SAVE ──────────────────────────────────────────────────────────────────────
wb.save(OUTPUT)
print(f"✓ Guardado: {OUTPUT}")
print(f"\nHojas en el archivo:")
for s in wb.sheetnames:
    print(f"  - {s}")
print(f"\nNOVEDADES INCORPORADAS:")
print(f"  • Cronograma oficial PG-001 Rev A (26-MAY-2026)")
print(f"  • Informe de Avance N°1 (04-JUN-2026): estado actividades Mayo 2026")
print(f"  • Pendientes AESA: 4 ítems críticos (5/6, 18/6, 30/6, 20/7)")
print(f"  • Hoja 6: Estado y pendientes al {FECHA_DOC}")
print(f"  • Hoja 7: Cronograma Certificaciones mensual completo (15 meses)")
print(f"  • Entrega Sala 4 confirmada: 08-ABR-2027")
print(f"  • Entrega Sala 3 confirmada: 14-MAY-2027")
print(f"\nTOTALES CERTIFICACIÓN (excluyendo repuestos):")
print(f"  OC S3: USD {sum(s3_totals):>14,.0f}")
print(f"  OC S4: USD {sum(s4_totals):>14,.0f}")
print(f"  TOTAL: USD {sum(comb_totals):>14,.0f}")
