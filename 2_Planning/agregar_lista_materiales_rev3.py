#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rev 3 — Agrega Sheet 9_Lista_Materiales desde ACAL-00102-LM-E-0012-0.xlsx
Cambios vs Rev 2:
  - Sheet 9_Lista_Materiales: listado oficial de materiales cables (LM) con
    cantidades por formación Rev0 + surplus + total, según doc ACAL-00102-LM-E-0012-0
  - Formación normalizada (NxSección) para cruce con Sheet 1_Cables_Completo
"""

import os, re, shutil
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))
DOC_DIR      = os.path.join(PLANNING_DIR, "..", "4_Documentos")
SRC  = os.path.join(PLANNING_DIR, "Analisis_Electrico_CPF2_Campo_Rev2_2026-06-09.xlsx")
LM   = os.path.join(DOC_DIR, "ACAL-00102-LM-E-0012-0.xlsx")
OUT  = os.path.join(PLANNING_DIR, "Analisis_Electrico_CPF2_Campo_Rev3_2026-06-13.xlsx")
REV  = "Rev 3"
FECHA = "13-JUN-2026"

shutil.copy2(SRC, OUT)

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

# Colours
RED="CC0000"; DRD="880000"; BLU="2E75B6"; DBL="1F4E79"; GRN="375623"
WH="FFFFFF"; GRY="595959"; YEL="FFEB9C"; ORG="E36C09"; PRP="7030A0"
BLT="DEEAF1"; RLT="FFE0E0"; YLT="FFF8DC"; GR2="E0E0E0"
OKT="375623"; WNT="9C5700"; OK="C6EFCE"; WRN="FFC7CE"
MT_BG="EBF5FF"; MT_FG=DBL
BTSN_BG="FFF0F0"; BTSN_FG=DRD
BTARM_BG="FFF8DC"; BTARM_FG="7B4A00"
BTRF_BG="F0FFF0"; BTRF_FG=GRN
CTRL_BG="F5F0FF"; CTRL_FG=PRP
SEC_BLDS = {
    "CABLES DE MEDIA TENSIÓN": (MT_BG, MT_FG, "MT"),
    "CABLES DE BAJA TENSIÓN SIN ARMAR": (BTSN_BG, BTSN_FG, "BT-SinArm"),
    "CABLES DE BAJA TENSIÓN ARMADOS": (BTARM_BG, BTARM_FG, "BT-Arm"),
    "CABLES DE BAJA TENSIÓN RESISTENTES AL FUEGO": (BTRF_BG, BTRF_FG, "BT-RF"),
    "CABLES DE COMANDO Y SEÑALES": (CTRL_BG, CTRL_FG, "Control/Senal"),
}

# ── PARSE LM SOURCE ───────────────────────────────────────────────────────────
def cond_count(config):
    """Return (n_conductores, cond_str) from config like '1Ph', '3Ph+N', etc."""
    if '3Ph+N+PE' in config: return 5, '3Ph+N+PE'
    if '3Ph+N'    in config: return 4, '3Ph+N'
    if '3Ph+PE'   in config: return 4, '3Ph+PE'
    if '3Ph'      in config: return 3, '3Ph'
    if '1Ph+N+PE' in config: return 3, '1Ph+N+PE'
    if '1Ph+N'    in config: return 2, '1Ph+N'
    if '1Ph'      in config: return 1, '1Ph'
    m = re.search(r'(\d+)\s+conductor', config)
    if m: return int(m.group(1)), f'{m.group(1)}cond'
    m = re.search(r'(\d+)\s*PAIR', config, re.I)
    if m: return int(m.group(1)) * 2, f'{m.group(1)}par'
    m = re.search(r'(\d+)\s*TRIAD', config, re.I)
    if m: return int(m.group(1)) * 3, f'{m.group(1)}tríadas'
    return None, config

def extract_section(desc):
    """Extract conductor section from description string."""
    # Section 2,5mm2 or Section: 2,5mm2 or Section 70mm2 or 3x25/16mm2
    m = re.search(r'[Ss]ection[:\s]+([0-9,./x]+)\s*mm', desc)
    if m: return m.group(1).replace(',','.')
    # Ø 0,52mm2
    m = re.search(r'Ø\s*([0-9,.]+)\s*mm', desc)
    if m: return m.group(1).replace(',','.')
    # 1,31mm2
    m = re.search(r'(\d+[,.]?\d*)\s*mm2', desc)
    if m: return m.group(1).replace(',','.')
    return ''

def make_formacion(n_cond, sec_raw, section_name, desc):
    """Build a normalised formation string."""
    if 'MEDIA TENSIÓN' in section_name or 'kV' in desc[:50]:
        arm = 'MT'
    elif 'RESISTENTES' in section_name:
        arm = 'RF'
    elif 'ARMADO' in section_name or 'armour' in desc.lower() or 'armor' in desc.lower():
        arm = 'Arm'
    else:
        arm = 'SinArm'

    if not n_cond:
        return f"REVISAR | {arm}"

    sec = sec_raw.strip() if sec_raw else '?'
    # For reduced-neutral cables like 3x95/50, show as NxS/Sn
    if '/' in sec and 'x' not in sec:
        # e.g. 3x95/50 → already has 3 phases included in notation
        return f"{n_cond}cond×{sec}mm² {arm}"
    # Clean section: remove trailing slash
    sec_clean = re.sub(r'x[\d.]+$', '', sec)  # 3x25/16 → remove /16? no keep
    return f"{n_cond}x{sec_clean}mm² {arm}"

def is_armado(desc):
    return 'armour' in desc.lower() or 'armor' in desc.lower() or 'wire arm' in desc.lower()

wb_lm = load_workbook(LM, read_only=True, data_only=True)
ws_lm = wb_lm['Listado']
rows_lm = list(ws_lm.iter_rows(min_row=1, values_only=True))

SECTION_HEADERS = set(SEC_BLDS.keys())
current_section = ""
current_bg = GR2; current_fg = GRY; current_tag = ""

LM_ITEMS = []

for i, row in enumerate(rows_lm[11:], 12):
    r = list(row) + [None]*200
    # Detect section header (col 1 non-null, col 4 item = None, col 7 code = None)
    col1 = r[0]; col4 = r[3]; col7 = r[6]
    if col1 and not col4 and not col7:
        sec_name = str(col1).strip().upper()
        for known in SECTION_HEADERS:
            if known.upper() in sec_name or sec_name in known.upper():
                current_section = known
                current_bg, current_fg, current_tag = SEC_BLDS[known]
                break
        continue

    item  = r[3]; code  = r[6]; desc  = r[15]; unit = r[41]
    rev0c = r[44]; rev0s = r[47]; total = r[68]; diff = r[73]

    if not item or not (code or desc): continue

    desc_str = str(desc or '')
    sec_raw  = extract_section(desc_str)

    # Detect conductor configuration
    config_pat = re.search(r'(3Ph\+N\+PE|3Ph\+N|3Ph\+PE|3Ph|1Ph\+N\+PE|1Ph\+N|1Ph'
                           r'|\d+\s+conductor[s]?'
                           r'|\d+\s+PAIR[S]?'
                           r'|\d+\s+TRIAD[S]?)', desc_str, re.I)
    config = config_pat.group(0).strip() if config_pat else ''
    n_cond, cond_lbl = cond_count(config)

    forma = make_formacion(n_cond, sec_raw, current_section, desc_str)
    armado_bool = is_armado(desc_str)

    rev0c_v = rev0c if isinstance(rev0c, (int, float)) else 0
    rev0s_v = rev0s if isinstance(rev0s, (int, float)) else 0
    total_v = total if isinstance(total, (int, float)) else (rev0c_v + rev0s_v)
    sur_pct = (rev0s_v / rev0c_v * 100) if rev0c_v else 0

    # Short description (first 80 chars before comma after main cable type)
    # Try to find tension and main specs
    desc_short_m = re.match(r'^(.+?),\s*rated voltage\s*([^\,]+)', desc_str)
    if desc_short_m:
        dtype = desc_short_m.group(1).strip()
        volt  = desc_short_m.group(2).strip().split(' ')[0]
        desc_short = f"{dtype} {volt}"
    else:
        desc_short = desc_str[:80]

    LM_ITEMS.append({
        'item': str(item),
        'code': str(code or ''),
        'desc_full': desc_str,
        'desc_short': desc_short,
        'section': current_section,
        'tag': current_tag,
        'bg': current_bg, 'fg': current_fg,
        'formacion': forma,
        'n_cond': n_cond,
        'sec_raw': sec_raw,
        'armado': armado_bool,
        'unit': str(unit or 'm'),
        'rev0c': rev0c_v,
        'rev0s': rev0s_v,
        'sur_pct': sur_pct,
        'total': total_v,
        'diff': diff or 0,
    })

wb_lm.close()
print(f"LM items loaded: {len(LM_ITEMS)}")

# ── BUILD SHEET 9 ─────────────────────────────────────────────────────────────
wb = load_workbook(OUT)

if "9_Lista_Materiales" in wb.sheetnames:
    del wb["9_Lista_Materiales"]
ws9 = wb.create_sheet("9_Lista_Materiales")
ws9.sheet_view.showGridLines = False

# Column widths
COLS = {
    'A': 5,   # #
    'B': 7,   # Item
    'C': 22,  # Código
    'D': 22,  # Formación
    'E': 50,  # Descripción resumida
    'F': 10,  # Tensión
    'G': 8,   # Armado
    'H': 10,  # Cant. Base (m)
    'I': 10,  # Surplus (m)
    'J': 8,   # Sur%
    'K': 12,  # Total c/Sur (m)
    'L': 28,  # Notas
}
for col, w in COLS.items():
    ws9.column_dimensions[col].width = w

NC = len(COLS)  # 12

def shdr9(r, title, bg, nc=NC):
    ws9.merge_cells(f"A{r}:{get_column_letter(nc)}{r}")
    c = ws9.cell(r, 1, title)
    c.font = fn(True, WH, 10)
    c.alignment = al("left", "center")
    c.fill = fl(bg)
    c.border = bd("medium")
    ws9.row_dimensions[r].height = 22

def col_hdrs9(r, headers, bg=DBL):
    for i, h in enumerate(headers, 1):
        C(ws9, r, i, h, bold=True, fg=WH, bg=bg, sz=8, wrap=True)
    ws9.row_dimensions[r].height = 36

# ── TITLE ─────────────────────────────────────────────────────────────────────
r = 1
ws9.merge_cells(f"A1:{get_column_letter(NC)}1")
ws9["A1"].value = (
    f"LISTADO DE MATERIALES — CABLES ELÉCTRICOS  |  "
    f"Ref. Doc: ACAL-00102-LM-E-0012-0  |  {REV}  |  {FECHA}  |  "
    f"CPF-2 LA CALERA II"
)
ws9["A1"].font = fn(True, WH, 11)
ws9["A1"].alignment = al("center")
ws9["A1"].fill = fl(DRD)
ws9["A1"].border = bd("medium")
ws9.row_dimensions[1].height = 26

r = 2
ws9.merge_cells(f"A2:{get_column_letter(NC)}2")
ws9["A2"].value = (
    "Documento oficial de materiales de cables para aprovisionamiento CPF-2. "
    "Las cantidades incluyen el surplus aprobado por formación. "
    "La 'FORMACIÓN' normalizada permite el cruce con la hoja 1_Cables_Completo. "
    "Revisión 0 del LM emitida el 11-JUN-2026 — sin cambios respecto a revisiones previas (Dif.=0)."
)
ws9["A2"].font = fn(False, WH, 8.5)
ws9["A2"].alignment = al("left", "center", wrap=True)
ws9["A2"].fill = fl(GRY)
ws9["A2"].border = bd()
ws9.row_dimensions[2].height = 18

r = 4
col_hdrs9(r, [
    "#", "Item\nLM", "Código\nABB/Cliente",
    "Formación\n(normalizada)", "Descripción del Material",
    "Tensión", "Armado\n(SI/NO)",
    "Cant.\nBase\n(m)", "Surplus\n(m)", "Sur\n%",
    "TOTAL\nc/Surplus\n(m)", "Notas"
])
r += 1

# ── DATA ROWS ─────────────────────────────────────────────────────────────────
section_order = list(SEC_BLDS.keys())
current_sec_printed = ""
row_counter = 0
section_totals = {s: {'base': 0.0, 'sur': 0.0, 'total': 0.0, 'n': 0} for s in section_order}
grand = {'base': 0.0, 'sur': 0.0, 'total': 0.0, 'n': 0}

for item in LM_ITEMS:
    sec = item['section']

    # Section header
    if sec != current_sec_printed:
        bg, fg, tag = SEC_BLDS.get(sec, (GR2, GRY, ""))
        shdr9(r, f"  {sec}", bg)
        r += 1
        current_sec_printed = sec

    bg = item['bg']; fg = item['fg']
    row_counter += 1

    # Tension from description
    volt_m = re.search(r'([\d.]+/[\d.]+)\s*kV', item['desc_full'])
    volt_str = (volt_m.group(0) if volt_m else
                ("13.2kV" if "13.2" in item['desc_full'] else
                 ("6.6kV" if "6.6" in item['desc_full'] else "0.6/1kV")))

    arm_str = "SI" if item['armado'] else "NO"
    arm_fg  = "8B4513" if item['armado'] else GRY
    arm_bg  = "FFF0D0" if item['armado'] else bg

    sur_pct_str = f"{item['sur_pct']:.0f}%"

    C(ws9, r, 1, row_counter, bold=False, fg=GRY, bg=bg, sz=8)
    C(ws9, r, 2, item['item'], bold=True, fg=fg, bg=bg, sz=9)
    C(ws9, r, 3, item['code'], bold=False, fg=fg, bg=bg, sz=8, h="left")
    C(ws9, r, 4, item['formacion'], bold=True, fg=fg, bg=bg, sz=9, h="left")
    C(ws9, r, 5, item['desc_short'], bold=False, fg=GRY, bg=bg, sz=8, h="left", wrap=True)
    C(ws9, r, 6, volt_str, bold=False, fg=fg, bg=bg, sz=8)
    C(ws9, r, 7, arm_str, bold=item['armado'], fg=arm_fg, bg=arm_bg, sz=8)
    C(ws9, r, 8, item['rev0c'] if item['rev0c'] else None, bold=False, fg=fg, bg=bg, sz=9, nf='#,##0')
    C(ws9, r, 9, item['rev0s'] if item['rev0s'] else None, bold=False, fg=arm_fg, bg=arm_bg, sz=9, nf='#,##0.0')
    C(ws9, r, 10, sur_pct_str, bold=False, fg=arm_fg, bg=arm_bg, sz=8)
    C(ws9, r, 11, item['total'] if item['total'] else None, bold=True, fg=fg, bg=OK if item['total'] else bg, sz=10, nf='#,##0.0')
    C(ws9, r, 12, "", bold=False, fg=GRY, bg=bg, sz=8, h="left")

    ws9.row_dimensions[r].height = 20
    r += 1

    # Accumulate
    if sec in section_totals:
        section_totals[sec]['base']  += item['rev0c']
        section_totals[sec]['sur']   += item['rev0s']
        section_totals[sec]['total'] += item['total']
        section_totals[sec]['n']     += 1
    grand['base']  += item['rev0c']
    grand['sur']   += item['rev0s']
    grand['total'] += item['total']
    grand['n']     += 1

    # Section subtotal after last item in section
    # Peek ahead
    next_sec = LM_ITEMS[LM_ITEMS.index(item) + 1]['section'] if LM_ITEMS.index(item) + 1 < len(LM_ITEMS) else None
    if next_sec != sec:
        bg, fg, tag = SEC_BLDS.get(sec, (GR2, GRY, ""))
        st = section_totals[sec]
        ws9.merge_cells(f"A{r}:G{r}")
        c = ws9.cell(r, 1, f"Subtotal {sec}  ({st['n']} tipos)")
        c.font = fn(True, WH, 9); c.alignment = al("left", "center")
        c.fill = fl(fg); c.border = bd()
        C(ws9, r, 8, st['base'], bold=True, fg=WH, bg=fg, sz=9, nf='#,##0')
        C(ws9, r, 9, st['sur'],  bold=True, fg=WH, bg=fg, sz=9, nf='#,##0.0')
        C(ws9, r, 10, f"{st['sur']/st['base']*100:.0f}%" if st['base'] else "", bold=True, fg=WH, bg=fg, sz=8)
        C(ws9, r, 11, st['total'], bold=True, fg=WH, bg=fg, sz=10, nf='#,##0.0')
        C(ws9, r, 12, "", bg=fg)
        ws9.row_dimensions[r].height = 18
        r += 1

r += 1

# ── GRAND TOTAL ───────────────────────────────────────────────────────────────
ws9.merge_cells(f"A{r}:{get_column_letter(NC-1)}{r}")
c = ws9.cell(r, 1, f"★  GRAN TOTAL — TODOS LOS TIPOS  ({grand['n']} ítems)  — Rev 0 del LM emitida 11-JUN-2026")
c.font = fn(True, WH, 11); c.alignment = al("left", "center")
c.fill = fl(DRD); c.border = bd("medium")
C(ws9, r, NC, "", bg=DRD)
ws9.row_dimensions[r].height = 22
r += 1

TOTAL_HEADERS = ["", "Item", "Código", "Formación", "Descripción", "Tensión", "Armado",
                 "Base (m)", "Surplus (m)", "Sur%", "TOTAL (m)", ""]
for ci, h in enumerate(TOTAL_HEADERS, 1):
    C(ws9, r, ci, h, bold=True, fg=WH, bg=DBL, sz=8)

r += 1
sub_bgs = [MT_BG, BTSN_BG, BTARM_BG, BTRF_BG, CTRL_BG]
sub_fgs = [MT_FG, BTSN_FG, BTARM_FG, BTRF_FG, CTRL_FG]
for (sec_name, st), sbg, sfg in zip(section_totals.items(), sub_bgs, sub_fgs):
    _, _, tag = SEC_BLDS.get(sec_name, (GR2, GRY, ""))
    ws9.merge_cells(f"A{r}:G{r}")
    c = ws9.cell(r, 1, sec_name)
    c.font = fn(True, sfg, 9); c.alignment = al("left", "center")
    c.fill = fl(sbg); c.border = bd()
    C(ws9, r, 8,  st['base'],  bold=True, fg=sfg, bg=sbg, sz=9, nf='#,##0')
    C(ws9, r, 9,  st['sur'],   bold=True, fg=sfg, bg=sbg, sz=9, nf='#,##0.0')
    pct = f"{st['sur']/st['base']*100:.0f}%" if st['base'] else "—"
    C(ws9, r, 10, pct, bold=False, fg=sfg, bg=sbg, sz=8)
    C(ws9, r, 11, st['total'], bold=True, fg=sfg, bg=OK, sz=10, nf='#,##0.0')
    C(ws9, r, 12, "", bg=sbg)
    ws9.row_dimensions[r].height = 18
    r += 1

# Final total row
ws9.merge_cells(f"A{r}:G{r}")
c = ws9.cell(r, 1, "TOTAL GENERAL — TODOS LOS CABLES CPF-2")
c.font = fn(True, WH, 10); c.alignment = al("left", "center")
c.fill = fl(DRD); c.border = bd("medium")
C(ws9, r, 8,  grand['base'],  bold=True, fg=WH, bg=DRD, sz=10, nf='#,##0')
C(ws9, r, 9,  grand['sur'],   bold=True, fg=WH, bg=DRD, sz=10, nf='#,##0.0')
g_pct = f"{grand['sur']/grand['base']*100:.0f}%" if grand['base'] else "—"
C(ws9, r, 10, g_pct, bold=True, fg=WH, bg=DRD, sz=9)
C(ws9, r, 11, grand['total'], bold=True, fg=WH, bg=DRD, sz=12, nf='#,##0.0')
C(ws9, r, 12, "", bg=DRD)
ws9.row_dimensions[r].height = 24
r += 2

# ── METHODOLOGY NOTE ──────────────────────────────────────────────────────────
shdr9(r, "  NOTAS Y METODOLOGÍA", GRY)
r += 1
notes = [
    ("Documento de referencia",
     "ACAL-00102-LM-E-0012-0 — Listado de Materiales Cables — Rev 0, emitida 11-JUN-2026. "
     "Código AESA de referencia: 5155-00-0000-IG-EL-LM-001."),
    ("¿Qué es el LM?",
     "El Listado de Materiales (LM) es el documento oficial de ingeniería que especifica "
     "los tipos y cantidades de cables a adquirir para el proyecto CPF-2. "
     "Incluye el surplus aprobado por cada tipo de formación para cubrir errores de medición, "
     "reparaciones y remanentes en obra."),
    ("Cant. Base vs Total",
     "Cant. Base = metros netos según planilla de cómputo de cables. "
     "Surplus = margen adicional por formación (entre 5% y 20% según tipo y sección). "
     "Total = Cant. Base + Surplus → cantidad a comprar/fabricar."),
    ("Formación normalizada",
     "La columna FORMACIÓN traduce la especificación técnica del LM al formato NxSección usado "
     "en la hoja 1_Cables_Completo. Permite cruzar código de material con cada tramo del cable schedule. "
     "Ej: 3Ph+N 6mm² armado → '4x6mm² Arm' | 7 cond 2.5mm² armado → '7x2.5mm² Arm'."),
    ("5 Secciones del LM",
     "1) MT: Media Tensión 6.6kV y 13.2kV  |  2) BT Sin Arm: Baja Tensión no armados  |  "
     "3) BT Armados: Baja Tensión con armadura de acero  |  "
     "4) BT RF: Baja Tensión resistentes al fuego  |  5) Control/Señal: cables de comando e instrumentación."),
    ("Versión del documento",
     f"Análisis Rev 3 | {FECHA}. Incorpora LM ACAL-00102-LM-E-0012-0 Rev 0 como referencia. "
     "La columna Dif. del LM = 0 en todos los ítems → esta es la primera emisión."),
]
for concepto, nota in notes:
    ws9.merge_cells(f"A{r}:C{r}")
    c = ws9.cell(r, 1, concepto)
    c.font = fn(True, DBL, 9); c.alignment = al("left", "center")
    c.fill = fl(BLT); c.border = bd()
    ws9.merge_cells(f"D{r}:{get_column_letter(NC)}{r}")
    c = ws9.cell(r, 4, nota)
    c.font = fn(False, GRY, 8.5, True); c.alignment = al("left", "center", wrap=True)
    c.fill = fl("F8FBFF"); c.border = bd()
    ws9.row_dimensions[r].height = 32
    r += 1

ws9.freeze_panes = "A5"

# ── UPDATE REVISION IN EXISTING SHEETS ───────────────────────────────────────
for ws in wb.worksheets:
    if ws.title in ("9_Lista_Materiales",): continue
    for row_cells in ws.iter_rows(max_row=3):
        for cell in row_cells:
            if cell.value and isinstance(cell.value, str):
                if "Rev 2" in cell.value and "Rev 3" not in cell.value:
                    cell.value = cell.value.replace("Rev 2", "Rev 3")
                    cell.value = cell.value.replace("09-JUN-2026", FECHA)

# ── REORDER SHEETS ────────────────────────────────────────────────────────────
desired = ["1_Cables_Completo", "2_Resumen_Formaciones", "3_Cargas_CPF2",
           "4_Estimacion_Instalacion", "5_Cronograma_Campo", "6_Resumen_Ejecutivo",
           "8_Puntas_Conexion", "9_Lista_Materiales"]
for i, name in enumerate(desired):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i:
            try: wb.move_sheet(name, offset=i - idx)
            except: pass

wb.save(OUT)
print(f"\n✓ Guardado: {OUT}")
print(f"Hojas: {wb.sheetnames}")
print(f"\nRESUMEN LISTADO DE MATERIALES:")
for sec, st in section_totals.items():
    _, _, tag = SEC_BLDS[sec]
    sur_avg = (st['sur']/st['base']*100) if st['base'] else 0
    print(f"  {tag:<18} {st['n']:>2} tipos  Base={st['base']:>10,.0f}m  Sur={st['sur']:>8,.1f}m ({sur_avg:.0f}%)  Total={st['total']:>10,.1f}m")
print(f"\n  {'TOTAL GENERAL':<18} {grand['n']:>2} tipos  Base={grand['base']:>10,.0f}m  Sur={grand['sur']:>8,.1f}m  Total={grand['total']:>10,.1f}m")
