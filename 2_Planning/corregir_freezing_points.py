#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CORRECCIÓN FREEZING POINTS — PG-001 Rev A vs Cronograma Actualizado 09-JUN-2026
Fuente de verdad: 5155-00-0009-VE-CO-PG-001_A.mpp (ABB 26-MAY-2026)
20 FPs reales en el MPP vs 8 simplificados en el Excel → se corrigen y expanden
"""

import os, shutil
from datetime import date
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))
BASE   = os.path.join(PLANNING_DIR, "CRONOGRAMA_GENERAL_ABB_Actualizado_2026-06-09.xlsx")
OUTPUT = os.path.join(PLANNING_DIR, "CRONOGRAMA_GENERAL_ABB_Actualizado_2026-06-09.xlsx")

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

def fmt(d): return d.strftime("%d/%m/%Y")

# ── LOAD WORKBOOK ─────────────────────────────────────────────────────────────
wb = load_workbook(OUTPUT)

# ── REMOVE OLD SHEET 4 AND REBUILD ───────────────────────────────────────────
if "4_Freezing_Points" in wb.sheetnames:
    del wb["4_Freezing_Points"]

ws4 = wb.create_sheet("4_Freezing_Points")
ws4.sheet_view.showGridLines = False

# Column widths
ws4.column_dimensions["A"].width = 4
ws4.column_dimensions["B"].width = 8
ws4.column_dimensions["C"].width = 38
ws4.column_dimensions["D"].width = 14
ws4.column_dimensions["E"].width = 11
ws4.column_dimensions["F"].width = 11
ws4.column_dimensions["G"].width = 11
ws4.column_dimensions["H"].width = 30

C_RED="CC0000"; C_DRD="880000"; C_BLU="2E75B6"; C_DBL="1F4E79"
C_GRN="375623"; C_WH="FFFFFF"; C_GRY="595959"
C_YEL="FFEB9C"; C_ORG="E36C09"; C_WRN="FFC7CE"; C_OK="C6EFCE"
C_BLT="DEEAF1"; C_RLT="FFE0E0"; C_YEL2="FFF8DC"
C_WT="9C5700"; C_OKT="375623"; C_RKT="9C0006"

# Title
ws4.merge_cells("A1:H1")
ws4["A1"].value = ("FREEZING POINTS CORREGIDOS — PG-001 Rev A  |  "
                   "Fuente: 5155-00-0009-VE-CO-PG-001_A.mpp  |  ABB 26-MAY-2026  |  "
                   "Actualizado: 09-JUN-2026")
ws4["A1"].font      = fnt(bold=True, color=C_WH, size=13)
ws4["A1"].alignment = aln("center", "center")
ws4["A1"].fill      = fill(C_DRD)
ws4["A1"].border    = brd("medium")
ws4.row_dimensions[1].height = 28

ws4.merge_cells("A2:H2")
ws4["A2"].value = ("MPP contiene 20 FPs por equipo/nivel ingeniería. "
                   "Correcciones críticas vs Excel previo: "
                   "FP-S3-3 era 31-AGO → ahora 26-OCT (+56d) | "
                   "FP-S3/S4 Sala ID era AGO → ahora 06-OCT (+39-53d) | "
                   "★ ÚLTIMO FP ABB REAL: Ductos barras 26-OCT-2026")
ws4["A2"].font      = fnt(color=C_WH, size=9)
ws4["A2"].alignment = aln("center", "center")
ws4["A2"].fill      = fill(C_GRY)
ws4["A2"].border    = brd()
ws4.row_dimensions[2].height = 18

# Section headers helper
def section_hdr(ws, r, title, bg):
    ws.merge_cells(f"A{r}:H{r}")
    c = ws.cell(row=r, column=1, value=title)
    c.font      = fnt(bold=True, color=C_WH, size=10)
    c.alignment = aln("left", "center")
    c.fill      = fill(bg)
    c.border    = brd("medium")
    ws.row_dimensions[r].height = 20

def col_hdr(ws, r):
    for i, h in enumerate(["#","ID FP","Descripción FP","Fecha MPP","Δ vs Excel previo","Estado","OC / Sala","Impacto si demora"], 1):
        cell(ws, r, i, h, bold=True, color=C_WH, bg=C_DBL, size=9)
    ws.row_dimensions[r].height = 16

def fp_row(ws, r, num, fp_id, desc, fecha_mpp, fecha_old, estado, oc, impacto,
           critica=False, bg_row=None):
    diff = (fecha_mpp - fecha_old).days if fecha_old else None
    if bg_row is None:
        bg_row = C_RLT if critica else (C_BLT if "S4" in oc else C_YEL2)

    # status colors
    if estado == "PENDIENTE":
        st_bg, st_fc = C_YEL, C_WT
    elif estado == "VENCIDA":
        st_bg, st_fc = C_WRN, C_RKT
    else:
        st_bg, st_fc = C_OK, C_OKT

    cell(ws, r, 1, num, bold=critica, size=9, bg=bg_row,
         color=C_DRD if critica else C_DBL)
    cell(ws, r, 2, fp_id, bold=True, size=9, bg=bg_row,
         color=C_DRD if critica else C_DBL)
    cell(ws, r, 3, desc, size=9, h="left", wrap=True, bg=bg_row)
    cell(ws, r, 4, fmt(fecha_mpp), bold=True, size=9, bg=bg_row,
         color=C_DRD if critica else "000000")
    if diff is not None:
        dstr = f"+{diff}d" if diff > 0 else f"{diff}d"
        d_bg = C_WRN if abs(diff)>=30 else (C_YEL if abs(diff)>=7 else C_OK)
        d_fc = C_RKT if abs(diff)>=30 else (C_WT if abs(diff)>=7 else C_OKT)
        cell(ws, r, 5, dstr, bold=abs(diff)>=7, size=9, bg=d_bg, color=d_fc)
    else:
        cell(ws, r, 5, "—(nuevo)", size=8, bg=C_YEL, color=C_WT, italic=True)
    cell(ws, r, 6, estado, bold=True, size=9, bg=st_bg, color=st_fc)
    cell(ws, r, 7, oc, bold=critica, size=9, bg=bg_row)
    cell(ws, r, 8, impacto, size=8, h="left", wrap=True, bg=bg_row, italic=True)
    ws.row_dimensions[r].height = 28

# ─────────────────────────────────────────────────────────────────────────────
# SALA 3 — BT — EQUIPOS
# ─────────────────────────────────────────────────────────────────────────────
r = 4
section_hdr(ws4, r, "  SALA #3 BT — FREEZING POINTS POR EQUIPO (OC 4508944971)", C_DRD)
r += 1; col_hdr(ws4, r); r += 1

s3_fps = [
    # (fp_id, descripción, fecha_mpp, fecha_excel_previa, critica, oc, impacto)
    ("FP-S3/Tab.Aux-IB", "IB Tableros Auxiliares SE#3 (102-DP-001, LP-001/002, DP-001-TRA)",
     date(2026,6,24), date(2026,7,3), False, "S3",
     "Gatilla acopio materiales Tab.Aux → Fabricación Nov"),
    ("FP-S3/Dr.BT-IB",  "IB Drives BT — 102-DP-VFD-24210A/B/C",
     date(2026,6,29), date(2026,7,3), False, "S3",
     "OC VFDs Finlandia (lead ~4m) → Entrega Enero 2027"),
    ("FP-S3/CCM-IB",    "IB CCM-005 / CCM-006 / CCM-007 (todos el mismo día)",
     date(2026,7,6),  date(2026,7,3), True,  "S3",
     "OC ductos barras Turquía (lead 6m). CADENA CRÍTICA — 1d demora = impacto en despacho"),
    ("FP-S3/UPS-IB",    "IB Rectificadores/UPS/Baterías — 102-UPS-001A/B, VCC-001A/B",
     date(2026,7,7),  date(2026,7,3), False, "S3",
     "Gatilla compra y acopio materiales → Fabricación Nov"),
    ("FP-S3/Tab.Aux-ID","ID Tableros Auxiliares SE#3",
     date(2026,7,20), date(2026,8,14), False, "S3",
     "Inicio fabricación tableros auxiliares. Entrega shelterista: Ene 2027"),
    ("FP-S3/Sala3-IB",  "IB Sala Eléctrica #3 — BUILDING (Layout + dimensiones + accesos)",
     date(2026,7,21), date(2026,7,3), False, "S3",
     "Congela: dimensiones sala, acometidas trafo, disposición equipos S3"),
    ("FP-S3/CCM005-ID", "ID CCM-005 — Ingeniería de Detalle completa y aprobada",
     date(2026,8,12), date(2026,8,14), True,  "S3",
     "Inicio compra materiales secundarios CCM-005 → Fabricación Oct"),
    ("FP-S3/CCM006-ID", "ID CCM-006 — Ingeniería de Detalle completa y aprobada",
     date(2026,8,17), date(2026,8,14), True,  "S3",
     "Inicio compra materiales secundarios CCM-006 → Fabricación Nov"),
    ("FP-S3/CCM007-ID", "ID CCM-007 — Ingeniería de Detalle completa y aprobada",
     date(2026,8,19), date(2026,8,14), True,  "S3",
     "Último FP CCM → Fabricación CCM-007 comienza Nov 2026"),
    ("FP-S3/Sala3-ID",  "★ ID Sala Eléctrica #3 — BUILDING (Planos constructivos + instalación)",
     date(2026,10,6),  date(2026,8,31), True,  "S3",
     "Inicio fabricación sala → Construcción base Oct-Dic 2026. +36d vs Excel previo"),
    ("FP-S3/DB-FP",     "★★★ Ductos de Barras SE#3 — ÚLTIMO FP ABSOLUTO DEL PROYECTO ABB",
     date(2026,10,26), date(2026,8,31), True,  "S3",
     "ÚLTIMO FP REAL: Congela diseño ductos → Fabricación Tur/Oct-Ene → Despacho Ene 2027. +56d vs Excel"),
]

for idx, (fp_id, desc, f_mpp, f_old, crit, oc, imp) in enumerate(s3_fps, 1):
    estado = "VENCIDA" if f_mpp < date.today() else "PENDIENTE"
    fp_row(ws4, r, idx, fp_id, desc, f_mpp, f_old, estado, oc, imp, crit)
    r += 1

r += 1

# ─────────────────────────────────────────────────────────────────────────────
# SALA 4 — MT — EQUIPOS
# ─────────────────────────────────────────────────────────────────────────────
section_hdr(ws4, r, "  SALA #4 MT — FREEZING POINTS POR EQUIPO (OC 4508944973)", C_DBL)
r += 1; col_hdr(ws4, r); r += 1

s4_fps = [
    ("FP-S4/Celdas-IB",  "IB Celdas MT Unigear 13.2/6.6kV (TGMT-001/002)",
     date(2026,6,23), date(2026,6,26), False, "S4",
     "OC Celdas MT Turquía (lead 6m) → Entrega Dic 2026. -3d vs Excel ≈ OK"),
    ("FP-S4/Tab.Aux-IB", "IB Tableros Auxiliares SE#4 (102-DP-101, DP-101UPS, DP-101VCC)",
     date(2026,6,24), date(2026,6,26), False, "S4",
     "Gatilla acopio materiales Tab.Aux → Fabricación Nov"),
    ("FP-S4/Dr.MT-IB",   "IB Drives MT — 102-DP-VFD-23310A/B/C",
     date(2026,6,29), date(2026,6,26), False, "S4",
     "OC VFDs MT China (lead 4-5m) → Entrega Feb 2027"),
    ("FP-S4/Tab.Aux-ID", "ID Tableros Auxiliares SE#4",
     date(2026,7,20), date(2026,8,14), False, "S4",
     "Inicio fabricación tableros auxiliares → Entrega shelterista Ene 2027"),
    ("FP-S4/Sala4-IB",   "IB Sala Eléctrica #4 — BUILDING (Layout + dimensiones + accesos)",
     date(2026,7,21), date(2026,6,26), False, "S4",
     "Congela: dimensiones sala MT, acometidas celdas, disposición equipos. +25d vs Excel"),
    ("FP-S4/Celdas-ID",  "ID Celdas MT Unigear — Ingeniería de Detalle aprobada",
     date(2026,8,3),  date(2026,8,14), False, "S4",
     "Inicio compra materiales secundarios → Fabricación Sep-Nov. -11d vs Excel"),
    ("FP-S4/Sala4-ID",   "★ ID Sala Eléctrica #4 — BUILDING (Planos constructivos + instalación)",
     date(2026,10,6), date(2026,8,28), False, "S4",
     "Inicio fabricación sala → Construcción base Oct-Dic 2026. +39d vs Excel previo"),
]

for idx, (fp_id, desc, f_mpp, f_old, crit, oc, imp) in enumerate(s4_fps, 1):
    estado = "VENCIDA" if f_mpp < date.today() else "PENDIENTE"
    fp_row(ws4, r, idx, fp_id, desc, f_mpp, f_old, estado, oc, imp, crit)
    r += 1

r += 1

# ─────────────────────────────────────────────────────────────────────────────
# TABLA RESUMEN / COMPARACIÓN
# ─────────────────────────────────────────────────────────────────────────────
section_hdr(ws4, r,
    "  RESUMEN COMPARATIVO: FECHAS EXCEL PREVIO vs MPP PG-001 Rev A (SIMPLIFICADO 8 FPs principales)",
    "404040")
r += 1

headers_cmp = ["ID (Excel)", "FP Simplificado", "Fecha Excel previo", "Fecha MPP real",
               "Diferencia", "Veredicto"]
for i, h in enumerate(headers_cmp, 1):
    cell(ws4, r, i, h, bold=True, color=C_WH, bg=C_GRY, size=9)
ws4.row_dimensions[r].height = 16
r += 1

resumen_cmp = [
    ("FP-P1",  "Ing. Básica PMS",           date(2026,8,21),  None,               "PMS (otro MPP) — sin cambio"),
    ("FP-P2",  "Ing. Detalle PMS",           date(2026,10,23), None,               "PMS (otro MPP) — sin cambio"),
    ("FP-S4-1","IB Sala #4 (equipos)",       date(2026,6,26),  date(2026,6,23),    "≈OK — Celdas MT IB 23-JUN"),
    ("FP-S4-1","IB Sala #4 (building)",      date(2026,6,26),  date(2026,7,21),    "⚠ REVISAR — Sala IB FP 21-JUL"),
    ("FP-S4-2","ID Sala #4 (Celdas MT)",     date(2026,8,14),  date(2026,8,3),     "⚠ Celdas ID es 03-AGO (-11d)"),
    ("FP-S4-2","ID Sala #4 (building)",      date(2026,8,14),  date(2026,10,6),    "🔴 CRÍTICO: Sala ID es 06-OCT (+53d)"),
    ("FP-S4-3","IC Sala #4 (→ ID building)", date(2026,8,28),  date(2026,10,6),    "🔴 CRÍTICO: +39 días — usar 06-OCT"),
    ("FP-S3-1","IB Sala #3 (CCMs)",         date(2026,7,3),   date(2026,7,6),     "≈OK — CCMs IB FP 06-JUL (+3d)"),
    ("FP-S3-1","IB Sala #3 (building)",      date(2026,7,3),   date(2026,7,21),    "⚠ REVISAR — Sala IB FP 21-JUL (+18d)"),
    ("FP-S3-2","ID Sala #3 (CCMs)",         date(2026,8,14),  date(2026,8,12),    "≈OK — CCM-005 ID 12-AGO (-2d)"),
    ("FP-S3-2","ID Sala #3 (building)",      date(2026,8,14),  date(2026,10,6),    "🔴 CRÍTICO: Sala ID es 06-OCT (+53d)"),
    ("FP-S3-3","IC/Último FP S3 (Sala)",    date(2026,8,31),  date(2026,10,6),    "🔴 CRÍTICO: Sala ID es 06-OCT (+36d)"),
    ("FP-S3-3","★ ÚLTIMO FP REAL (Ductos)", date(2026,8,31),  date(2026,10,26),   "🔴 CRÍTICO: Ductos 26-OCT (+56 DÍAS)"),
]

for (id_, fp_name, f_old, f_mpp, veredicto) in resumen_cmp:
    if f_mpp:
        diff = (f_mpp - f_old).days
        diff_str = f"+{diff}d" if diff > 0 else f"{diff}d"
        d_bg = C_WRN if abs(diff)>=30 else (C_YEL if abs(diff)>=7 else C_OK)
        d_fc = C_RKT if abs(diff)>=30 else (C_WT if abs(diff)>=7 else C_OKT)
        cell(ws4, r, 3, fmt(f_old), size=9, bg="F8F8F8")
        cell(ws4, r, 4, fmt(f_mpp), bold=abs(diff)>=30, size=9,
             color=C_DRD if abs(diff)>=30 else "000000")
        cell(ws4, r, 5, diff_str, bold=abs(diff)>=7, size=9, bg=d_bg, color=d_fc)
    else:
        cell(ws4, r, 3, fmt(f_old), size=9, bg="F8F8F8")
        cell(ws4, r, 4, "sin cambio", size=9, italic=True, color=C_GRY)
        cell(ws4, r, 5, "—", size=9, color=C_GRY)
    cell(ws4, r, 1, id_, bold=True, size=9, bg="F0F0F0")
    cell(ws4, r, 2, fp_name, size=9, h="left")
    v_bg = C_WRN if "🔴" in veredicto else (C_YEL if "⚠" in veredicto else C_OK)
    v_fc = C_RKT if "🔴" in veredicto else (C_WT if "⚠" in veredicto else C_OKT)
    cell(ws4, r, 6, veredicto, size=8, h="left", bg=v_bg, color=v_fc, wrap=True)
    ws4.row_dimensions[r].height = 22
    r += 1

r += 1

# ── NOTA FINAL ────────────────────────────────────────────────────────────────
ws4.merge_cells(f"A{r}:H{r}")
ws4[f"A{r}"].value = (
    "⚠ ACCIÓN INMEDIATA: "
    "El ÚLTIMO FP real del proyecto es Ductos de barras el 26-OCT-2026 (no el 31-AGO-2026 que figuraba antes). "
    "El plazo hasta el último FP se extendió +56 días, lo que puede afectar las fechas de despacho si no se gestionan los proveedores de Turquía. "
    "| Sala eléctrica ID FPs (S3 y S4) = 06-OCT-2026. "
    "| Revisar lead times de fabricación con el nuevo cronograma."
)
ws4[f"A{r}"].font      = fnt(bold=True, color=C_WH, size=9)
ws4[f"A{r}"].alignment = aln("left", "center", wrap=True)
ws4[f"A{r}"].fill      = fill("8B0000")
ws4[f"A{r}"].border    = brd("medium")
ws4.row_dimensions[r].height = 44

ws4.freeze_panes = "A4"

# ── REORDER SHEETS ────────────────────────────────────────────────────────────
desired = ["1_Gantt_General","2_PMS_Detalle","3_Compatibilidad_Shelter",
           "4_Freezing_Points","5_Hitos_Pago_por_OC",
           "6_Estado_JUN2026","7_Cert_Mensual_Hitos"]
for i, name in enumerate(desired):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i:
            try:
                wb.move_sheet(name, offset=i - idx)
            except:
                pass

# ── SAVE ──────────────────────────────────────────────────────────────────────
wb.save(OUTPUT)
print(f"✓ Guardado: {OUTPUT}")
print(f"\nHojas:")
for s in wb.sheetnames:
    print(f"  - {s}")
print("\nCORRECCIONES CRÍTICAS APLICADAS EN HOJA 4_Freezing_Points:")
print("  🔴 FP-S3-3 (Último FP ABB): 31-AGO-2026 → 26-OCT-2026  (+56 días)")
print("  🔴 FP-S4-3 / Sala 4 ID:     28-AGO-2026 → 06-OCT-2026  (+39 días)")
print("  🔴 FP-S3-2 / Sala 3 ID:     14-AGO-2026 → 06-OCT-2026  (+53 días)")
print("  ⚠  FP-S4-1 / Sala 4 IB:     26-JUN-2026 → 21-JUL-2026  (+25 días) [building]")
print("  ⚠  FP-S3-1 / Sala 3 IB:     03-JUL-2026 → 21-JUL-2026  (+18 días) [building]")
print("  ✓  11 nuevos FPs por equipo agregados (total: 20 FPs de MPP)")
print("  ✓  Tabla comparativa completa agregada en Sheet 4")
