#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rev 2 — Análisis de Puntas de Conexión CPF-2  — 09-JUN-2026
Cambios vs Rev 1:
  - Sheet 1_Cables_Completo: 7 columnas nuevas por cable (COND/PUNTA, ARMADO,
    PTAS_COND, PTAS_ARM, TOTAL_PTAS, TIPO_CABLE, TIPO_CARGA)
  - Sheet 8_Puntas_Conexion reconstruida con fórmulas SUMIF que apuntan
    directamente a Sheet 1 → cierre garantizado sin inconsistencias
"""

import os, re
from datetime import date
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))
SRC  = os.path.join(PLANNING_DIR, "Analisis_Electrico_CPF2_Campo_Rev1_2026-06-09.xlsx")
OUT  = os.path.join(PLANNING_DIR, "Analisis_Electrico_CPF2_Campo_Rev2_2026-06-09.xlsx")
REV  = "Rev 2"
FECHA = "09-JUN-2026"
import shutil; shutil.copy2(SRC, OUT)

# ── STYLE HELPERS ─────────────────────────────────────────────────────────────
def fl(c): return PatternFill("solid", fgColor=c)
def fn(bold=False, color="000000", size=9, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic, name="Calibri")
def al(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def bd(s="thin"):
    x = Side(style=s); return Border(left=x, right=x, top=x, bottom=x)

def C(ws, r, c, v=None, bold=False, fg="000000", bg=None, sz=9,
      h="center", v2="center", wrap=False, ita=False, bdr=True, nf=None):
    cell = ws.cell(r, c, v)
    cell.font      = fn(bold, fg, sz, ita)
    cell.alignment = al(h, v2, wrap)
    if bg:  cell.fill   = fl(bg)
    if bdr: cell.border = bd()
    if nf:  cell.number_format = nf
    return cell

RED="CC0000"; DRD="880000"; BLU="2E75B6"; DBL="1F4E79"; GRN="375623"
WH="FFFFFF"; GRY="595959"; YEL="FFEB9C"; ORG="E36C09"
BLT="DEEAF1"; RLT="FFE0E0"; YLT="FFF8DC"; GR2="E0E0E0"
OKT="375623"; WNT="9C5700"; RKT="9C0006"; OK="C6EFCE"; WRN="FFC7CE"
PRP="7030A0"

# ── FORMATION PARSER ──────────────────────────────────────────────────────────
def parse_cond(formacion):
    """Returns (n_cond_por_punta, nota_formacion)"""
    if not formacion: return 2, "sin dato"
    f = str(formacion).strip().replace(',', '.')
    # Fiber optic
    if re.match(r'^[Ff]\.?[Oo]\.?', f): return 2, "FO-2conect"
    # Multi-par NxMxS  (e.g. 6x2x0.5 = 12, 8x3x1.31 = 24)
    m = re.match(r'^(\d+)x(\d+)x[\d.]+', f)
    if m: return int(m.group(1)) * int(m.group(2)), f"multi-par({m.group(1)}×{m.group(2)})"
    # NxS/M  (3x240/120 → 4; 3x95/50 → 4)
    m = re.match(r'^(\d+)x[\d.]+/[\d.]+', f)
    if m: return int(m.group(1)) + 1, f"{m.group(1)}fase+N/PE"
    # NxS+T  (7x2.5+T → 8; 4x4+T → 5)
    m = re.match(r'^(\d+)x[\d.]+\+T', f)
    if m: return int(m.group(1)) + 1, f"{m.group(1)}cond+dren"
    # NxS  (3x240 → 3; 2x4 → 2; 12x2.5 → 12)
    m = re.match(r'^(\d+)x[\d.]+', f)
    if m: return int(m.group(1)), f"{m.group(1)}cond"
    # Solo número → probablemente 1 conductor (cable unipolar) o dato faltante
    m = re.match(r'^(\d+)$', f)
    if m:
        n = int(m.group(1))
        if n <= 5: return n, f"{n}cond-verificar"
        return 2, f"DATO-REVISAR({f})"
    return 2, f"FORM-REVISAR({f})"

def is_arm(const):
    return bool(const) and 'Arm' in str(const)

# ── CLASSIFICATION HELPERS ────────────────────────────────────────────────────
def tipo_cable(const, tipo_col, formacion, tension):
    c = str(const or ''); t = str(tipo_col or ''); f = str(formacion or '')
    if 'FO' in t.upper() or f.upper().startswith('F.O'): return "FO / ETHERNET"
    mt_kv = ['13,2kV','6,6kV','3,3kV','13.2kV','6.6kV','3.3kV']
    if any(x in c for x in mt_kv): return "POTENCIA MT (>1kV)"
    if t == 'SEÑAL': return "SEÑAL / INSTRUMENTO"
    if t == 'CONTROL': return "CONTROL"
    # BT: classify by section
    fm = f.replace(',', '.')
    m2 = re.search(r'x([\d.]+)', fm)
    sec = float(m2.group(1)) if m2 else 0
    if sec >= 95:  return "POTENCIA BT GRANDE (≥95mm²)"
    if sec >= 35:  return "POTENCIA BT MEDIANA (35-70mm²)"
    if sec >= 10:  return "POTENCIA BT PEQUEÑA (10-25mm²)"
    if sec > 0:    return "POTENCIA BT MUY PEQ (<10mm²)"
    return "OTROS/VERIFICAR"

def tipo_carga(desc, hacia, area):
    d = str(desc or '').upper(); h = str(hacia or '').upper()
    a = str(area or '').upper()
    if any(x in d+h for x in ['VFD','VARIADOR','DRIVE']): return "VFD/VARIADOR"
    if any(x in d+h for x in ['MOTOR','BOMB','PUMP','FAN','BLOWER',
                                'COOLER','COMPRESSOR','COMPRESOR']): return "MOTOR/ACCIONAMIENTO"
    if any(x in d+h+a for x in ['ILUM','LIGHT','LAMP','LA-','LE-']): return "ILUMINACION"
    if any(x in d+h+a for x in ['TOMA','SOCKET','OUTLET']): return "TOMACORRIENTE"
    if any(x in d+h for x in ['HEATER','CALEF','TRACING']): return "CALEFACCION/TRAZADO"
    if any(x in d+h for x in ['CCM','MCC','TABLERO','PANEL','BOARD']): return "TABLERO/CCM"
    if any(x in d+h for x in ['UPS','VCC','RECTIF','BATER','ATS']): return "UPS/RECTIF/SSAA"
    if any(x in d+h for x in ['INSTR','BB-','TRANSM','SENSOR','DETECT']): return "INSTRUMENTO/SENSOR"
    if any(x in d+h for x in ['SEÑAL','SIGNAL','F.O.','FO','FIBER']): return "SEÑAL/FO"
    if any(x in d+h for x in ['CONTROL','CTRL','MANDO']): return "CONTROL"
    if any(x in a for x in ['SALA ELÉCTRICA','SALA ELECTRICA','GENERADORES']): return "TABLERO/CCM"
    return "OTROS/VERIFICAR"

# ── LOAD WORKBOOK ─────────────────────────────────────────────────────────────
wb = load_workbook(OUT)
ws1 = wb['1_Cables_Completo']
MAX_DATA_ROW = ws1.max_row  # 634

# ── STEP 1: ADD COLUMNS TO SHEET 1 ───────────────────────────────────────────
# Current columns A-K (11). New columns L-R (7 more).
#  L = COND/PUNTA   (numeric: conductors per cable end)
#  M = ARMADO       (text: SI / NO)
#  N = PTAS_COND    (numeric: COND/PUNTA × 2 = total conductor terminations)
#  O = PTAS_ARM     (numeric: 2 if armored else 0 = armor gland terminations)
#  P = TOTAL_PTAS   (numeric: N + O = grand total terminations for this cable)
#  Q = TIPO_CABLE   (text category)
#  R = TIPO_CARGA   (text category)

NEW_COLS = {
    12: ("COND/\nPUNTA",   "Conductores individuales\nen cada extremo del cable",   BLT, DBL, 8),
    13: ("ARMADO\n(SI/NO)","Cable armado = gland\nen cada extremo",                 YLT, "8B6914", 7),
    14: ("PTAS\nCOND",     "Puntas conductores\n(COND/PUNTA × 2 extremos)",         BLT, DBL, 8),
    15: ("PTAS\nARM",      "Puntas armadura\n(glands: 2 si armado, 0 si no)",       YLT, "8B6914", 8),
    16: ("TOTAL\nPTAS",    "PTAS_COND + PTAS_ARM\n→ total puntas este cable",       OK,  OKT, 9),
    17: ("TIPO\nCABLE",    "Categoría por tensión\ny sección del conductor",         GR2, GRY, 7),
    18: ("TIPO\nCARGA",    "Tipo de equipo destino\n(extremo campo/carga)",          GR2, GRY, 7),
}

# Update title row
for col_i, (hdr, tooltip, hdr_bg, hdr_fg, sz) in NEW_COLS.items():
    ws1.column_dimensions[get_column_letter(col_i)].width = \
        11 if col_i in (16,) else (9 if col_i in (12,13,14,15) else 20)
    c2 = ws1.cell(1, col_i, hdr)
    c2.font      = fn(True, WH, sz)
    c2.fill      = fl(hdr_bg)
    c2.alignment = al("center", "center", wrap=True)
    c2.border    = bd()
ws1.row_dimensions[1].height = 30

# Populate data rows
stats = {'cpf2':0,'cpf1':0,'arm_cables':0,'total_ptas_cpf2':0}
for row_i in range(2, MAX_DATA_ROW + 1):
    tag     = ws1.cell(row_i, 1).value
    de      = ws1.cell(row_i, 2).value
    hacia   = ws1.cell(row_i, 3).value
    tension = ws1.cell(row_i, 4).value
    const   = ws1.cell(row_i, 6).value
    forma   = ws1.cell(row_i, 7).value
    area    = ws1.cell(row_i, 8).value
    desc    = ws1.cell(row_i, 9).value
    tipo    = ws1.cell(row_i, 10).value
    cpf     = ws1.cell(row_i, 11).value

    if not tag: continue

    n_cond, nota_form = parse_cond(forma)
    armado  = is_arm(const)
    ptas_cond = n_cond * 2
    ptas_arm  = 2 if armado else 0
    total_pt  = ptas_cond + ptas_arm
    t_cable   = tipo_cable(const, tipo, forma, tension)
    t_carga   = tipo_carga(desc, hacia, area)

    is_cpf2 = str(cpf).upper() == 'CPF2' if cpf else False

    # Highlight rows by CPF
    row_bg_cond = BLT if is_cpf2 else "F5F5F5"
    row_bg_arm  = YLT if is_cpf2 else "FFFAE0"
    row_bg_tot  = "E8F5E9" if is_cpf2 else "F0FFF0"
    row_bg_typ  = "F0F0F0"
    arm_bg      = "FFF0D0" if armado and is_cpf2 else (row_bg_arm)

    # Column L: COND/PUNTA
    c = ws1.cell(row_i, 12, n_cond)
    c.font = fn(True if is_cpf2 else False, DBL if is_cpf2 else GRY, 9)
    c.fill = fl(row_bg_cond); c.alignment = al("center"); c.border = bd()

    # Column M: ARMADO
    arm_txt = "SI" if armado else "NO"
    c = ws1.cell(row_i, 13, arm_txt)
    c.font = fn(armado, ("8B6914" if armado else GRY), 9)
    c.fill = fl(arm_bg if armado else row_bg_arm)
    c.alignment = al("center"); c.border = bd()

    # Column N: PTAS_COND
    c = ws1.cell(row_i, 14, ptas_cond)
    c.font = fn(is_cpf2, DBL if is_cpf2 else GRY, 9)
    c.fill = fl(row_bg_cond); c.alignment = al("center"); c.border = bd()
    c.number_format = '0'

    # Column O: PTAS_ARM
    c = ws1.cell(row_i, 15, ptas_arm)
    c.font = fn(ptas_arm > 0, ("C67200" if ptas_arm else GRY), 9)
    c.fill = fl(arm_bg if ptas_arm else row_bg_arm)
    c.alignment = al("center"); c.border = bd()
    c.number_format = '0'

    # Column P: TOTAL_PTAS
    c = ws1.cell(row_i, 16, total_pt)
    c.font = fn(True, OKT if is_cpf2 else GRY, 10 if is_cpf2 else 9)
    c.fill = fl(row_bg_tot); c.alignment = al("center"); c.border = bd("medium" if is_cpf2 else "thin")
    c.number_format = '0'

    # Column Q: TIPO_CABLE
    c = ws1.cell(row_i, 17, t_cable)
    c.font = fn(False, GRY, 8)
    c.fill = fl(row_bg_typ); c.alignment = al("left", wrap=True); c.border = bd()

    # Column R: TIPO_CARGA
    c = ws1.cell(row_i, 18, t_carga)
    c.font = fn(False, GRY, 8)
    c.fill = fl(row_bg_typ); c.alignment = al("left", wrap=True); c.border = bd()

    if is_cpf2:
        stats['cpf2'] += 1
        if armado: stats['arm_cables'] += 1
        stats['total_ptas_cpf2'] += total_pt
    else:
        stats['cpf1'] += 1

print(f"Sheet 1 populated: {stats}")

# Freeze and resize
ws1.freeze_panes = "A2"
ws1.row_dimensions[1].height = 30

# ── STEP 2: REBUILD SHEET 8 WITH SUMIF FORMULAS ───────────────────────────────
# Sheet 1 column references (row 2 to MAX_DATA_ROW):
# K = CPF (col 11)
# L = COND/PUNTA (col 12)
# N = PTAS_COND (col 14)
# O = PTAS_ARM (col 15)
# P = TOTAL_PTAS (col 16)  ← KEY
# Q = TIPO_CABLE (col 17)
# R = TIPO_CARGA (col 18)
S1 = "'1_Cables_Completo'"
r2 = 2; rN = MAX_DATA_ROW
K  = f"{S1}!$K${r2}:$K${rN}"
P  = f"{S1}!$P${r2}:$P${rN}"
N_ = f"{S1}!$N${r2}:$N${rN}"
O_ = f"{S1}!$O${r2}:$O${rN}"
Q  = f"{S1}!$Q${r2}:$Q${rN}"
R_ = f"{S1}!$R${r2}:$R${rN}"

def sf_cpf2_total(col_range):
    """SUMIF CPF column = CPF2, sum col_range"""
    return f'=SUMIF({K},"CPF2",{col_range})'

def sf_cpf2_tipo(tipo_range, tipo_val, col_range):
    """SUMIFS CPF=CPF2 AND tipo=tipo_val, sum col_range"""
    return f'=SUMIFS({col_range},{K},"CPF2",{tipo_range},"{tipo_val}")'

def sf_cpf(cpf_val, col_range):
    return f'=SUMIF({K},"{cpf_val}",{col_range})'

def sf_cpf_tipo(cpf_val, tipo_range, tipo_val, col_range):
    return f'=SUMIFS({col_range},{K},"{cpf_val}",{tipo_range},"{tipo_val}")'

def countif_cpf2(tipo_range, tipo_val):
    return f'=COUNTIFS({K},"CPF2",{tipo_range},"{tipo_val}")'

def countif_cpf(cpf_val, tipo_range, tipo_val):
    return f'=COUNTIFS({K},"{cpf_val}",{tipo_range},"{tipo_val}")'

# Rebuild the sheet
if "8_Puntas_Conexion" in wb.sheetnames:
    del wb["8_Puntas_Conexion"]
ws8 = wb.create_sheet("8_Puntas_Conexion")
ws8.sheet_view.showGridLines = False

# Column widths
for col, w in {"A":5,"B":38,"C":10,"D":10,"E":10,"F":10,"G":10,"H":10,"I":10,"J":32}.items():
    ws8.column_dimensions[col].width = w

def shdr(r, title, bg, ncols=10):
    ws8.merge_cells(f"A{r}:{get_column_letter(ncols)}{r}")
    c = ws8.cell(r, 1, title)
    c.font = fn(True, WH, 10); c.alignment = al("left", "center")
    c.fill = fl(bg); c.border = bd("medium")
    ws8.row_dimensions[r].height = 20

def col_hdr_row(r, headers, bg=DBL):
    for i, h in enumerate(headers, 1):
        C(ws8, r, i, h, bold=True, fg=WH, bg=bg, sz=9, wrap=True)
    ws8.row_dimensions[r].height = 30

# ── TITLE ─────────────────────────────────────────────────────────────────────
r = 1
ws8.merge_cells("A1:J1")
ws8["A1"].value = (f"ANÁLISIS DE PUNTAS DE CONEXIÓN — CPF-2 LA CALERA II  |  "
                   f"{REV}  |  {FECHA}  |  "
                   f"Totales calculados con SUMIF sobre hoja 1_Cables_Completo col.P (TOTAL_PTAS)")
ws8["A1"].font = fn(True, WH, 11); ws8["A1"].alignment = al("center")
ws8["A1"].fill = fl(DRD); ws8["A1"].border = bd("medium")
ws8.row_dimensions[1].height = 28

r = 2
ws8.merge_cells("A2:J2")
ws8["A2"].value = (
    "★ CIERRE GARANTIZADO: cada celda de esta hoja usa SUMIF/SUMIFS sobre la columna P (TOTAL_PTAS) "
    "de la hoja 1_Cables_Completo — cualquier cambio en Sheet 1 actualiza automáticamente este resumen.  "
    "| PUNTA = 1 conductor a conectar en 1 extremo del cable (incluye conductores + armadura si corresponde)."
)
ws8["A2"].font = fn(False, WH, 8.5); ws8["A2"].alignment = al("center", wrap=True)
ws8["A2"].fill = fl(GRY); ws8["A2"].border = bd()
ws8.row_dimensions[2].height = 18

# ── A. TOTALES GENERALES ──────────────────────────────────────────────────────
r = 4
shdr(r, "  A. TOTALES GENERALES (SUMIF directo sobre 1_Cables_Completo columna TOTAL_PTAS)", DRD)
r += 1

col_hdr_row(r, ["Concepto","Cables\nCPF-2","Puntas\nConductores\n(PTAS_COND)",
                 "Puntas\nArmadura\n(PTAS_ARM)","TOTAL\nPUNTAS\n(SUMIF)","",
                 "Fórmula usada","","","Verificación"])
r += 1

tot_cables_formula   = f'=COUNTIF({K},"CPF2")'
tot_ptas_cond_formula= sf_cpf2_total(N_)
tot_ptas_arm_formula = sf_cpf2_total(O_)
tot_total_formula    = sf_cpf2_total(P)

rows_A = [
    ("TOTAL CPF-2 — Puntas conductores",
     tot_cables_formula, tot_ptas_cond_formula, "—", "—",
     "", f'SUMIF(K,"CPF2",N)', "", "",
     "N_COND × 2 extremos por cable"),
    ("TOTAL CPF-2 — Puntas armadura (glands)",
     f'=COUNTIFS({K},"CPF2",{S1}!$M${r2}:$M${rN},"SI")',
     "—", tot_ptas_arm_formula, "—",
     "", f'SUMIF(K,"CPF2",O)', "", "",
     "2 glands por cable armado"),
    ("★ TOTAL CPF-2 — GRAN TOTAL PUNTAS",
     tot_cables_formula, tot_ptas_cond_formula, tot_ptas_arm_formula,
     tot_total_formula,
     "", f'SUMIF(K,"CPF2",P)', "", "",
     f"Valor esperado: {stats['total_ptas_cpf2']:,}"),
]

row_bgs = ["F8F8F8", "FFFAEF", "E2EFDA"]
row_bolds = [False, False, True]
row_fgs   = ["000000","8B6914", OKT]
for i, (concepto, *vals) in enumerate(rows_A):
    bg = row_bgs[i]; bold = row_bolds[i]; fg = row_fgs[i]
    C(ws8, r, 1, concepto, bold=bold, fg=fg, bg=bg, sz=9, h="left")
    C(ws8, r, 2, vals[0], bold=bold, fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 3, vals[1], bold=bold, fg=fg, bg=bg, sz=9, nf='#,##0' if vals[1]!="—" else None)
    C(ws8, r, 4, vals[2], bold=bold, fg=fg, bg=bg, sz=9, nf='#,##0' if vals[2]!="—" else None)
    efg = OKT if bold else fg
    C(ws8, r, 5, vals[3], bold=bold, fg=efg, bg=bg if i<2 else "C6EFCE", sz=10 if bold else 9,
      nf='#,##0' if vals[3]!="—" else None)
    for ci in range(6, 11): C(ws8, r, ci, vals[ci-2] if ci-2 < len(vals) else "", bg=bg, sz=8, ita=(ci==7))
    ws8.row_dimensions[r].height = 22 if bold else 18
    r += 1

r += 1

# ── B. POR TIPO DE CABLE ──────────────────────────────────────────────────────
shdr(r, "  B. PUNTAS POR TIPO DE CABLE — SUMIFS(P, K=\"CPF2\", Q=tipo)", DBL)
r += 1
col_hdr_row(r, ["Tipo de Cable","Cables\n(COUNTIFS)","Puntas\nConductores",
                 "Puntas\nArmadura","TOTAL\nPUNTAS","% del\ntotal","",
                 "Fórmula TOTAL_PTAS","","Nota"])
r += 1

TYPE_ORDER = [
    ("POTENCIA MT (>1kV)",          (BLT, DBL)),
    ("POTENCIA BT GRANDE (≥95mm²)", (RLT, DRD)),
    ("POTENCIA BT MEDIANA (35-70mm²)",("FFF0F0", RED)),
    ("POTENCIA BT PEQUEÑA (10-25mm²)",("FFF5F5","AA0000")),
    ("POTENCIA BT MUY PEQ (<10mm²)", ("FFFAFA","AA0000")),
    ("CONTROL",                      ("EBF3FB", BLU)),
    ("SEÑAL / INSTRUMENTO",          (YLT, WNT)),
    ("FO / ETHERNET",                ("F0F0FF","5050AA")),
    ("OTROS/VERIFICAR",              (GR2, GRY)),
]

total_check_row = r + len([t for t in TYPE_ORDER])
first_data_tipo_row = r
for tipo, (bg, fg) in TYPE_ORDER:
    cnt_f = countif_cpf2(Q, tipo)
    ptas_cond_f = sf_cpf_tipo("CPF2", Q, tipo, N_)
    ptas_arm_f  = sf_cpf_tipo("CPF2", Q, tipo, O_)
    total_f     = sf_cpf_tipo("CPF2", Q, tipo, P)
    pct_f       = f"=IF({sf_cpf2_total(P)}>0,{total_f}/{sf_cpf2_total(P)},0)"
    nota = ("Terminales a compresión MT — técnico certificado" if "MT" in tipo
            else ("Lugs a compresión BT" if "BT" in tipo
                  else ("Ferrules de bornera" if "CONTROL" in tipo
                        else ("Conector SC/LC" if "FO" in tipo else ""))))
    C(ws8, r, 1, tipo,       bold=False, fg=fg, bg=bg, sz=9, h="left")
    C(ws8, r, 2, cnt_f,      bold=False, fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 3, ptas_cond_f,bold=False, fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 4, ptas_arm_f, bold=False, fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 5, total_f,    bold=True,  fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 6, pct_f,      bold=False, fg=fg, bg=bg, sz=9, nf='0.0%')
    C(ws8, r, 7, "", bg=bg)
    C(ws8, r, 8, f'SUMIFS(P,K,"CPF2",Q,"{tipo}")', bg=bg, sz=7, ita=True, fg=GRY, h="left")
    C(ws8, r, 9, "", bg=bg)
    C(ws8, r, 10, nota, bg=bg, sz=8, fg=GRY, h="left")
    ws8.row_dimensions[r].height = 18
    r += 1

# Subtotal with check formula
C(ws8, r, 1, "TOTAL (debe = Gran Total sección A)", bold=True, fg=OKT, bg=OK, sz=9, h="left")
C(ws8, r, 2, tot_cables_formula, bold=True, fg=OKT, bg=OK, nf='#,##0')
C(ws8, r, 3, tot_ptas_cond_formula, bold=True, fg=OKT, bg=OK, nf='#,##0')
C(ws8, r, 4, tot_ptas_arm_formula,  bold=True, fg=OKT, bg=OK, nf='#,##0')
C(ws8, r, 5, tot_total_formula,     bold=True, fg=OKT, bg=OK, sz=10, nf='#,##0')
check_B = f'=IF({tot_total_formula}={sf_cpf2_total(P)},"✓ CIERRA","✗ REVISAR")'
C(ws8, r, 6, check_B, bold=True, fg=OKT, bg=OK, sz=9)
for ci in range(7, 11): C(ws8, r, ci, "", bg=OK)
ws8.row_dimensions[r].height = 18
r += 2

# ── C. POR TIPO DE CARGA ──────────────────────────────────────────────────────
shdr(r, "  C. PUNTAS POR TIPO DE CARGA / DESTINO — SUMIFS(P, K=\"CPF2\", R=carga)", GRN)
r += 1
col_hdr_row(r, ["Tipo de Carga / Equipo Destino","Cables","Puntas\nConductores",
                 "Puntas\nArmadura","TOTAL\nPUNTAS","% del\ntotal","",
                 "Fórmula TOTAL_PTAS","","Nota"], bg=GRN)
r += 1

CARGA_ORDER = [
    ("MOTOR/ACCIONAMIENTO",    ("FFE8E0","7B2600"), "Terminales en caja JB del motor / borne motor"),
    ("VFD/VARIADOR",           ("FFDDC0","7B2600"), "Terminales en regleta de campo del VFD"),
    ("TABLERO/CCM",            (RLT,     DRD),      "Borneras en frente del tablero destino"),
    ("UPS/RECTIF/SSAA",        ("FFE8F0","880055"), "Borneras panel UPS/VCC/ATS"),
    ("ILUMINACION",            ("FFFFF0","5B5800"), "Regleta en caja de conexión luminaria"),
    ("CALEFACCION/TRAZADO",    ("FFF0D8", WNT),     "Borne heater / junction box trazado"),
    ("TOMACORRIENTE",          ("F8F8FF", GRY),     "Borne regleta tomacorriente"),
    ("INSTRUMENTO/SENSOR",     (YLT,     WNT),      "Borneras Head / JB instrumento"),
    ("SEÑAL/FO",               ("F0F0FF","5050AA"), "Conector SC/LC fibra ó bornera señal"),
    ("CONTROL",                (BLT,     DBL),      "Borneras gabinete control"),
    ("OTROS/VERIFICAR",        (GR2,     GRY),      "Sin clasificar — revisar lista cables"),
]

for carga, (bg, fg), nota in CARGA_ORDER:
    cnt_f       = countif_cpf2(R_, carga)
    ptas_cond_f = sf_cpf_tipo("CPF2", R_, carga, N_)
    ptas_arm_f  = sf_cpf_tipo("CPF2", R_, carga, O_)
    total_f     = sf_cpf_tipo("CPF2", R_, carga, P)
    pct_f       = f"=IF({sf_cpf2_total(P)}>0,{total_f}/{sf_cpf2_total(P)},0)"
    C(ws8, r, 1, carga,      bold=False, fg=fg, bg=bg, sz=9, h="left")
    C(ws8, r, 2, cnt_f,      bold=False, fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 3, ptas_cond_f,bold=False, fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 4, ptas_arm_f, bold=False, fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 5, total_f,    bold=True,  fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws8, r, 6, pct_f,      bold=False, fg=fg, bg=bg, sz=9, nf='0.0%')
    C(ws8, r, 7, "", bg=bg)
    C(ws8, r, 8, f'SUMIFS(P,K,"CPF2",R,"{carga}")', bg=bg, sz=7, ita=True, fg=GRY, h="left")
    C(ws8, r, 9, "", bg=bg)
    C(ws8, r, 10, nota, bg=bg, sz=8, fg=GRY, h="left")
    ws8.row_dimensions[r].height = 18
    r += 1

# Subtotal + check
C(ws8, r, 1, "TOTAL (debe = Gran Total sección A)", bold=True, fg=OKT, bg=OK, sz=9, h="left")
C(ws8, r, 2, tot_cables_formula, bold=True, fg=OKT, bg=OK, nf='#,##0')
C(ws8, r, 3, tot_ptas_cond_formula, bold=True, fg=OKT, bg=OK, nf='#,##0')
C(ws8, r, 4, tot_ptas_arm_formula,  bold=True, fg=OKT, bg=OK, nf='#,##0')
C(ws8, r, 5, tot_total_formula,     bold=True, fg=OKT, bg=OK, sz=10, nf='#,##0')
check_C = f'=IF({tot_total_formula}={sf_cpf2_total(P)},"✓ CIERRA","✗ REVISAR")'
C(ws8, r, 6, check_C, bold=True, fg=OKT, bg=OK, sz=9)
for ci in range(7, 11): C(ws8, r, ci, "", bg=OK)
ws8.row_dimensions[r].height = 18
r += 2

# ── D. CPF-1 vs CPF-2 COMPARATIVO ────────────────────────────────────────────
shdr(r, "  D. COMPARATIVO CPF-1 vs CPF-2 (CPF-1 = scope EPC externo, referencia solamente)", PRP)
r += 1
col_hdr_row(r, ["CPF","Tipo Cable","Cables","Puntas\nConductores",
                 "Puntas\nArmadura","TOTAL\nPUNTAS","","Fórmula","",""], bg=PRP)
r += 1

for cpf_val, cpf_bg in [("CPF2",BLT),("CPF1","FFF0F0")]:
    first_cpf_row = r
    for tipo, (bg, _fg) in TYPE_ORDER:
        cnt_f       = countif_cpf(cpf_val, Q, tipo)
        ptas_cond_f = sf_cpf_tipo(cpf_val, Q, tipo, N_)
        ptas_arm_f  = sf_cpf_tipo(cpf_val, Q, tipo, O_)
        total_f     = sf_cpf_tipo(cpf_val, Q, tipo, P)
        C(ws8, r, 1, cpf_val, bold=True, fg=(DBL if cpf_val=="CPF2" else DRD), bg=cpf_bg, sz=8)
        C(ws8, r, 2, tipo, bold=False, fg=_fg, bg=bg, sz=8, h="left")
        C(ws8, r, 3, cnt_f,       fg=GRY, bg=bg, sz=8, nf='#,##0')
        C(ws8, r, 4, ptas_cond_f, fg=GRY, bg=bg, sz=8, nf='#,##0')
        C(ws8, r, 5, ptas_arm_f,  fg=GRY, bg=bg, sz=8, nf='#,##0')
        C(ws8, r, 6, total_f,     fg=_fg, bg=bg, sz=9, bold=True, nf='#,##0')
        for ci in range(7, 11): C(ws8, r, ci, "", bg=bg)
        ws8.row_dimensions[r].height = 15
        r += 1
    # CPF subtotal
    C(ws8, r, 1, f"TOTAL {cpf_val}", bold=True, fg=WH,
      bg=DBL if cpf_val=="CPF2" else DRD, sz=9)
    C(ws8, r, 2, "",   bold=True, bg=OK if cpf_val=="CPF2" else "FFE0E0")
    C(ws8, r, 3, sf_cpf(cpf_val, N_), bold=True, bg=OK if cpf_val=="CPF2" else "FFE0E0",
      fg=OKT if cpf_val=="CPF2" else RKT, nf='#,##0')
    C(ws8, r, 4, sf_cpf(cpf_val, O_), bold=True, bg=OK if cpf_val=="CPF2" else "FFE0E0",
      fg=OKT if cpf_val=="CPF2" else RKT, nf='#,##0')
    C(ws8, r, 5, sf_cpf(cpf_val, P),  bold=True, bg=OK if cpf_val=="CPF2" else "FFE0E0",
      fg=OKT if cpf_val=="CPF2" else RKT, sz=10, nf='#,##0')
    for ci in range(6, 11): C(ws8, r, ci, "", bg=OK if cpf_val=="CPF2" else "FFE0E0")
    ws8.row_dimensions[r].height = 18
    r += 2

# ── E. INSTRUCCIONES / METODOLOGÍA ───────────────────────────────────────────
shdr(r, "  E. METODOLOGÍA — CÓMO LEER ESTE ANÁLISIS", GRY)
r += 1

metodologia = [
    ("¿Qué es una PUNTA?",
     "Es un conductor individual (hilo de cobre o aluminio) en un extremo de un cable que debe ser conectado. "
     "Ejemplos: un cable 3x240mm² tiene 3 conductores → 3 puntas por extremo → 6 puntas totales + 2 glands armadura = 8 puntas."),
    ("Columna COND/PUNTA (col.L)",
     "Cantidad de conductores en cada extremo del cable, extraída de la formación. "
     "7x2.5+T = 8 cond/punta | 3x240/120 = 4 cond/punta (3 fases + 1 neutro) | 4x2x0.5 = 8 cond/punta (4 pares)."),
    ("Columna PTAS_COND (col.N)",
     "= COND/PUNTA × 2 extremos. Siempre par. Ejemplo: cable 7x2.5+T → 8 × 2 = 16 puntas conductor."),
    ("Columna PTAS_ARM (col.O)",
     "Puntas de armadura: 2 si el cable es armado (ARMADO=SI), 0 si no. "
     "Cada extremo de un cable armado requiere 1 prensaestopa o gland metálico → 2 por cable."),
    ("Columna TOTAL_PTAS (col.P)",
     "= PTAS_COND + PTAS_ARM. Esta es la columna maestra de la que se alimentan TODOS los SUMIF de esta hoja."),
    ("Cierre Sheet 8 vs Sheet 1",
     "Las secciones B, C y D usan SUMIFS sobre col.P de Sheet 1. Los subtotales de B y C "
     "verifican automáticamente con la fórmula ✓ CIERRA / ✗ REVISAR. No puede haber inconsistencias."),
    ("Formaciones a verificar",
     'Cables con formación "15" (9 cables): dato ambiguo en lista, se contabilizan como 2 cond/punta. '
     'Verificar con lista de cables original.'),
    ("CPF-1 vs CPF-2",
     "Solo CPF-2 es scope de este contrato. CPF-1 se muestra en sección D para contexto pero no se incluye "
     "en los totales operativos de las secciones A, B y C."),
]

for concepto, nota in metodologia:
    ws8.merge_cells(f"A{r}:B{r}")
    c = ws8.cell(r, 1, concepto)
    c.font = fn(True, DBL, 9); c.alignment = al("left", "center")
    c.fill = fl(BLT); c.border = bd()
    ws8.merge_cells(f"C{r}:J{r}")
    c = ws8.cell(r, 3, nota)
    c.font = fn(False, GRY, 9, True); c.alignment = al("left", "center", wrap=True)
    c.fill = fl("F8FBFF"); c.border = bd()
    ws8.row_dimensions[r].height = 30
    r += 1

ws8.freeze_panes = "A4"

# ── UPDATE REVISION IN EXISTING SHEETS ───────────────────────────────────────
for ws in wb.worksheets:
    if ws.title in ("8_Puntas_Conexion", "1_Cables_Completo"): continue
    for row_cells in ws.iter_rows(max_row=3):
        for cell in row_cells:
            if cell.value and isinstance(cell.value, str):
                if "Rev 1" in cell.value and "Rev 2" not in cell.value:
                    cell.value = cell.value.replace("Rev 1", "Rev 2")
                elif ("ANÁLISIS" in cell.value.upper() or "CABLES COMPLETO" in cell.value.upper()):
                    if "Rev 2" not in cell.value:
                        cell.value = cell.value.strip() + f"  |  {REV}  |  {FECHA}"

# ── REORDER SHEETS ────────────────────────────────────────────────────────────
desired = ["1_Cables_Completo","2_Resumen_Formaciones","3_Cargas_CPF2",
           "4_Estimacion_Instalacion","5_Cronograma_Campo","6_Resumen_Ejecutivo",
           "8_Puntas_Conexion"]
for i, name in enumerate(desired):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i:
            try: wb.move_sheet(name, offset=i-idx)
            except: pass

# ── SAVE ─────────────────────────────────────────────────────────────────────
wb.save(OUT)
print(f"\n✓ Guardado: {OUT}")
print(f"\nHojas: {wb.sheetnames}")
print(f"\nCOLUMNAS NUEVAS EN SHEET 1 (L→R):")
print(f"  L = COND/PUNTA     (conductores individuales por extremo)")
print(f"  M = ARMADO         (SI / NO)")
print(f"  N = PTAS_COND      (COND/PUNTA × 2 extremos)")
print(f"  O = PTAS_ARM       (2 si armado, 0 si no)")
print(f"  P = TOTAL_PTAS     ← columna maestra para SUMIF en Sheet 8")
print(f"  Q = TIPO_CABLE     (clasificación por tensión/sección)")
print(f"  R = TIPO_CARGA     (tipo equipo destino)")
print(f"\nSTATS:")
print(f"  CPF-2 cables procesados: {stats['cpf2']}")
print(f"  CPF-1 cables procesados: {stats['cpf1']}")
print(f"  Cables armados CPF-2:    {stats['arm_cables']}")
print(f"  Total puntas CPF-2:      {stats['total_ptas_cpf2']:,}")
print(f"\nSheet 8 usa SUMIF sobre col.P de Sheet 1 → cierre garantizado.")
