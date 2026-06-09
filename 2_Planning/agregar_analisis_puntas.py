#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
ANÁLISIS DE PUNTAS DE CONEXIÓN — Rev 1 — 09-JUN-2026
Agrega hoja "8_Puntas_Conexion" al Analisis_Electrico_CPF2_Campo.xlsx
Puntas = extremos de conductores a conectar en cada cable (ambos extremos)
Incluye: tableros/salas eléctricas + equipos/cargas en campo
"""

import os, re
from datetime import date
from collections import defaultdict
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))
FILE = os.path.join(PLANNING_DIR, "Analisis_Electrico_CPF2_Campo.xlsx")
FECHA   = "09-JUN-2026"
REV     = "Rev 1"
HOY     = date(2026, 6, 9)

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

C_RED="CC0000"; C_DRD="880000"; C_BLU="2E75B6"; C_DBL="1F4E79"
C_GRN="375623"; C_WH="FFFFFF"; C_GRY="595959"; C_YEL="FFEB9C"
C_ORG="E36C09"; C_WRN="FFC7CE"; C_OK="C6EFCE"; C_BLT="DEEAF1"
C_RLT="FFE0E0"; C_YEL2="FFF8DC"; C_PURP="7030A0"
C_WNT="9C5700"; C_OKT="375623"; C_RKT="9C0006"; C_WHT="F8F8F8"
C_GR2="E0E0E0"

# ── PARSING HELPERS ───────────────────────────────────────────────────────────
def parse_conductores_punta(formacion):
    """
    Parse cable formation string → (n_cond_por_punta, tipo_formacion)
    Cuenta conductores individuales que se conectan en CADA PUNTA del cable.
    """
    if not formacion: return 2, "genérico"
    f = str(formacion).strip().replace(',', '.')

    # FO (fibra óptica) — se cuentan conectores, no conductores
    m = re.match(r'^[Ff]\.?[Oo]\.?', f)
    if m: return 2, "fibra_optica"

    # NxAxB multi-par (e.g. 2x2x0.5 → 4 conductores, 8x3x1.31 → 24)
    m = re.match(r'^(\d+)x(\d+)x[\d.]+', f)
    if m: return int(m.group(1)) * int(m.group(2)), "multi_par"

    # NxS/M (fase+neutro/PE separado, e.g. 3x240/120 → 3+1=4)
    m = re.match(r'^(\d+)x[\d.]+/[\d.]+', f)
    if m: return int(m.group(1)) + 1, "trifasico_neutro"

    # NxS+T (con drenaje/tierra, e.g. 7x2.5+T → 7+1=8)
    m = re.match(r'^(\d+)x[\d.]+\+T', f)
    if m: return int(m.group(1)) + 1, "con_drenaje"

    # NxS básico (e.g. 3x240 → 3, 7x2.5 → 7)
    m = re.match(r'^(\d+)x[\d.]+', f)
    if m: return int(m.group(1)), "basico"

    # Solo número → default 2
    return 2, "generico"

def get_seccion_mm2(formacion):
    """Extrae la sección principal en mm² de la formación."""
    if not formacion: return 0
    f = str(formacion).strip().replace(',', '.')
    m = re.search(r'x([\d.]+)', f)
    if m:
        try: return float(m.group(1))
        except: pass
    return 0

def is_armado(const):
    return bool(const) and 'Arm' in str(const)

def is_fo(const, tipo, formacion):
    if tipo and 'FO' in str(tipo).upper(): return True
    if const and 'F.O.' in str(const): return True
    if formacion and str(formacion).upper().startswith('F.O'): return True
    return False

def is_mt(const, tension):
    if const and any(x in str(const) for x in ['13,2kV', '6,6kV', '3,3kV',
                                                  '13.2kV', '6.6kV', '3.3kV']): return True
    if tension:
        t = str(tension).replace(',', '.')
        v = re.sub(r'[^\d.]', '', t)
        try:
            if float(v) >= 1000: return True
        except: pass
    return False

def classif_tipo_cable(row):
    const = row['const']; tipo = row['tipo']; formacion = row['formacion']
    tension = row['tension']
    if is_fo(const, tipo, formacion): return "FIBRA ÓPTICA / ETHERNET"
    if is_mt(const, tension): return "POTENCIA MT (>1kV)"
    if tipo == 'SEÑAL': return "SEÑAL / INSTRUMENTO"
    if tipo == 'CONTROL': return "CONTROL"
    # BT — clasificar por sección
    sec = get_seccion_mm2(formacion)
    if sec >= 95:   return "POTENCIA BT GRANDE (≥95mm²)"
    if sec >= 35:   return "POTENCIA BT MEDIANA (35-70mm²)"
    if sec >= 10:   return "POTENCIA BT PEQUEÑA (10-25mm²)"
    if sec > 0:     return "POTENCIA BT MUY PEQUEÑA (<10mm²)"
    return "OTROS"

def classif_ubicacion_de(de_tag):
    """End DE: ¿está en sala eléctrica?"""
    if not de_tag: return 'SALA_ELEC'
    t = str(de_tag).upper()
    sala_kw = ['CCM-', 'DP-', 'TGMT', 'TGBT', 'LP-', 'UPS-', 'VCC-', 'DB-',
               'PMS-', 'GEN-', 'TABLERO', 'PANEL', 'SALA', 'KBE', 'SSAA',
               'TCA', 'TCB', 'TC-', 'MDB', 'ATS']
    for kw in sala_kw:
        if kw in t: return 'SALA_ELEC'
    return 'CAMPO'

def classif_ubicacion_hacia(hacia_tag, area):
    """End HACIA: ¿sala eléctrica o campo/equipo?"""
    a = str(area).upper() if area else ''
    if 'SALA ELÉCTRICA' in a or 'SALA ELECTRICA' in a: return 'SALA_ELEC'
    if 'ÁREA GENERADORES' in a or 'AREA GENERADORES' in a: return 'SALA_ELEC'
    if not hacia_tag: return 'CAMPO'
    t = str(hacia_tag).upper()
    sala_kw = ['CCM-', 'DP-', 'TGMT', 'TGBT', 'LP-', 'UPS-', 'VCC-', 'DB-',
               'PMS-', 'GEN-', 'TABLERO', 'PANEL', 'KBE', 'TCA', 'TCB']
    for kw in sala_kw:
        if kw in t: return 'SALA_ELEC'
    return 'CAMPO'

def classif_tipo_carga(desc, hacia_tag, area):
    d = str(desc).upper() if desc else ''
    t = str(hacia_tag).upper() if hacia_tag else ''
    a = str(area).upper() if area else ''

    if any(x in d+t for x in ['VFD','VARIADOR','DRIVE','VFD-']): return 'VFD / VARIADOR DE FRECUENCIA'
    if any(x in d+t for x in ['MOTOR','BOMB','PUMP','FAN','BLOWER',
                                'COOLER','COMPRESSOR','COMPRESOR','COOLER']): return 'MOTOR / ACCIONAMIENTO'
    if any(x in d+t+a for x in ['ILUM','LIGHT','LAMP','LA-','LE-']): return 'ILUMINACIÓN'
    if any(x in d+t+a for x in ['TOMA','SOCKET','OUTLET']): return 'TOMACORRIENTE'
    if any(x in d+t for x in ['HEATER','CALEF','TRACING','HEAT MED']): return 'CALEFACCIÓN / TRAZADO'
    if any(x in d+t for x in ['CCM','MCC','TABLERO','BOARD','PANEL','ARMARIO']): return 'TABLERO / CCM'
    if any(x in d+t for x in ['UPS','VCC','RECTIF','BATER','ATS']): return 'UPS / RECTIFICADOR / SSAA'
    if any(x in d+t for x in ['INSTR','BB-','TRANSM','SENSOR','DETECT',
                                'ANALI','ANALIZ']): return 'INSTRUMENTO / SENSOR'
    if any(x in d+t for x in ['SEÑAL','SIGNAL','F.O.','FO','FIBER','FIBRA',
                                'ETHERNET']): return 'SEÑAL / FIBRA ÓPTICA'
    if any(x in d+t for x in ['CONTROL','CTRL','MANDO','COMANDO']): return 'CONTROL'
    if any(x in a for x in ['SALA ELÉCTRICA', 'SALA ELECTRICA',
                              'AREA GENERADORES', 'ÁREA GENERADORES']): return 'TABLERO / CCM'
    return 'OTROS / SIN CLASIFICAR'

def classif_sala_elec(area, de_tag, hacia_tag):
    """Determina a cuál sala eléctrica pertenece el cable."""
    a = str(area).upper() if area else ''
    d = str(de_tag).upper() if de_tag else ''
    h = str(hacia_tag).upper() if hacia_tag else ''
    combined = a + d + h

    if 'SALA ELÉCTRICA 3' in combined or 'SALA ELECTRICA 3' in combined: return 'SE-3 BT'
    if 'SALA ELÉCTRICA 4' in combined or 'SALA ELECTRICA 4' in combined: return 'SE-4 MT'
    if 'GENERACIÓN' in combined or 'GENERACION' in combined or 'GEN' in combined: return 'SE-GENERACIÓN'
    if '102-CCM-0' in combined or '102-DP-' in combined or '102-LP' in combined: return 'SE-3 BT'
    if '102-TGMT' in combined: return 'SE-4 MT'
    if 'KBE' in combined or 'TCA' in combined: return 'SE-GENERACIÓN'
    return 'CAMPO / EXTERNO'

# ── READ CABLES ───────────────────────────────────────────────────────────────
wb = load_workbook(FILE)
ws1 = wb['1_Cables_Completo']

cables_all = []
for row in ws1.iter_rows(min_row=2, values_only=True):
    if row[0]:
        cables_all.append({
            'tag': row[0], 'de': row[1], 'hacia': row[2],
            'tension': row[3], 'longitud': row[4], 'const': row[5],
            'formacion': row[6], 'area': row[7], 'desc': row[8],
            'tipo': row[9], 'cpf': row[10]
        })

# Only CPF-2 cables for the main analysis (the scope of this project)
cables_cpf2 = [c for c in cables_all if str(c.get('cpf', '')).upper() == 'CPF2']
cables_cpf1 = [c for c in cables_all if str(c.get('cpf', '')).upper() != 'CPF2']
print(f"Total: {len(cables_all)}  CPF-2: {len(cables_cpf2)}  CPF-1: {len(cables_cpf1)}")

# ── COMPUTE TERMINATIONS PER CABLE ────────────────────────────────────────────
class CableTermData:
    def __init__(self, c):
        self.c = c
        self.n_cond, self.tipo_form = parse_conductores_punta(c['formacion'])
        self.armado       = is_armado(c['const'])
        self.fo           = is_fo(c['const'], c['tipo'], c['formacion'])
        self.tipo_cable   = classif_tipo_cable(c)
        self.ub_de        = classif_ubicacion_de(c['de'])
        self.ub_hacia     = classif_ubicacion_hacia(c['hacia'], c['area'])
        self.tipo_carga   = classif_tipo_carga(c['desc'], c['hacia'], c['area'])
        self.sala         = classif_sala_elec(c['area'], c['de'], c['hacia'])

        # Puntas conductores (cada extremo = n_cond; total = n_cond × 2)
        self.puntas_cond_de    = self.n_cond   # terminations at DE end
        self.puntas_cond_hacia = self.n_cond   # terminations at HACIA end
        self.puntas_cond_total = self.n_cond * 2

        # Puntas de armadura (1 por extremo si armado)
        self.puntas_arm_de    = 1 if self.armado else 0
        self.puntas_arm_hacia = 1 if self.armado else 0
        self.puntas_arm_total = 2 if self.armado else 0

        # Total puntas (conductores + armadura)
        self.puntas_total = self.puntas_cond_total + self.puntas_arm_total

td_cpf2 = [CableTermData(c) for c in cables_cpf2]
td_all  = [CableTermData(c) for c in cables_all]

# ── AGGREGATE DATA ────────────────────────────────────────────────────────────
def aggregate(td_list, key_fn):
    d = defaultdict(lambda: {'n_cables': 0, 'puntas_cond_sala': 0, 'puntas_cond_campo': 0,
                              'puntas_arm_sala': 0, 'puntas_arm_campo': 0,
                              'puntas_total': 0, 'metros': 0})
    for td in td_list:
        k = key_fn(td)
        d[k]['n_cables'] += 1
        try:
            d[k]['metros'] += float(td.c['longitud']) if td.c['longitud'] else 0
        except (ValueError, TypeError):
            pass

        # conductors at DE end
        if td.ub_de == 'SALA_ELEC':
            d[k]['puntas_cond_sala'] += td.puntas_cond_de
            d[k]['puntas_arm_sala']  += td.puntas_arm_de
        else:
            d[k]['puntas_cond_campo'] += td.puntas_cond_de
            d[k]['puntas_arm_campo']  += td.puntas_arm_de

        # conductors at HACIA end
        if td.ub_hacia == 'SALA_ELEC':
            d[k]['puntas_cond_sala'] += td.puntas_cond_hacia
            d[k]['puntas_arm_sala']  += td.puntas_arm_hacia
        else:
            d[k]['puntas_cond_campo'] += td.puntas_cond_hacia
            d[k]['puntas_arm_campo']  += td.puntas_arm_hacia

        d[k]['puntas_total'] += td.puntas_total
    return d

# Aggregations we need
by_tipo    = aggregate(td_cpf2, lambda t: t.tipo_cable)
by_carga   = aggregate(td_cpf2, lambda t: t.tipo_carga)
by_form    = aggregate(td_cpf2, lambda t: (str(t.c['formacion']), t.tipo_cable))
by_sala_e  = aggregate(td_cpf2, lambda t: t.sala)
by_cpf_tipo= aggregate(td_all,  lambda t: (str(t.c.get('cpf','')), t.tipo_cable))

# Global totals CPF-2
tot_cables_cpf2 = len(td_cpf2)
tot_puntas_cond_sala  = sum(t.puntas_cond_de if t.ub_de == 'SALA_ELEC' else 0
                            for t in td_cpf2) + \
                        sum(t.puntas_cond_hacia if t.ub_hacia == 'SALA_ELEC' else 0
                            for t in td_cpf2)
tot_puntas_cond_campo = sum(t.puntas_cond_de if t.ub_de != 'SALA_ELEC' else 0
                             for t in td_cpf2) + \
                        sum(t.puntas_cond_hacia if t.ub_hacia != 'SALA_ELEC' else 0
                             for t in td_cpf2)
tot_puntas_arm        = sum(t.puntas_arm_total for t in td_cpf2)
tot_puntas_total      = sum(t.puntas_total for t in td_cpf2)
tot_metros_cpf2 = sum(float(t.c['longitud']) for t in td_cpf2 if t.c['longitud'] and str(t.c['longitud']).replace('.','').isdigit())

print(f"\nRESUMEN CPF-2:")
print(f"  Cables: {tot_cables_cpf2}")
print(f"  Puntas conductores en sala:  {tot_puntas_cond_sala:>6,}")
print(f"  Puntas conductores en campo: {tot_puntas_cond_campo:>6,}")
print(f"  Puntas armadura (glands):    {tot_puntas_arm:>6,}")
print(f"  TOTAL PUNTAS:                {tot_puntas_total:>6,}")

# ── BUILD NEW SHEET ───────────────────────────────────────────────────────────
if "8_Puntas_Conexion" in wb.sheetnames:
    del wb["8_Puntas_Conexion"]

ws8 = wb.create_sheet("8_Puntas_Conexion")
ws8.sheet_view.showGridLines = False

# Column widths
col_w = {"A":4,"B":34,"C":10,"D":10,"E":10,"F":10,"G":10,"H":10,"I":10,"J":30}
for col, w in col_w.items():
    ws8.column_dimensions[col].width = w

def shdr(ws, r, title, bg, ncols=9):
    ws.merge_cells(f"A{r}:{get_column_letter(ncols)}{r}")
    c = ws.cell(row=r, column=1, value=title)
    c.font      = fnt(bold=True, color=C_WH, size=10)
    c.alignment = aln("left","center")
    c.fill      = fill(bg)
    c.border    = brd("medium")
    ws.row_dimensions[r].height = 20

def col_hdr_row(ws, r, headers, bg=C_DBL):
    for i, h in enumerate(headers, 1):
        cell(ws, r, i, h, bold=True, color=C_WH, bg=bg, size=8.5, wrap=True)
    ws.row_dimensions[r].height = 28

def data_row(ws, r, vals, bg="F8F8F8", bold=False, color="000000", height=18):
    for i, v in enumerate(vals, 1):
        nf = '#,##0' if isinstance(v, (int, float)) and i > 1 else None
        cell(ws, r, i, v, bold=bold, bg=bg if bg else None, size=9,
             h="right" if isinstance(v, (int, float)) else "left",
             border=True, num_format=nf, color=color)
    ws.row_dimensions[r].height = height

# ─── TITLE ────────────────────────────────────────────────────────────────────
r = 1
ws8.merge_cells(f"A{r}:J{r}")
ws8["A1"].value = (f"ANÁLISIS DE PUNTAS DE CONEXIÓN — CABLES ELÉCTRICOS CPF-2 LA CALERA II  |  "
                   f"{REV}  |  {FECHA}  |  CPF-2: {tot_cables_cpf2} cables  |  "
                   f"Total puntas: {tot_puntas_total:,}")
ws8["A1"].font      = fnt(bold=True, color=C_WH, size=12)
ws8["A1"].alignment = aln("center","center")
ws8["A1"].fill      = fill(C_DRD)
ws8["A1"].border    = brd("medium")
ws8.row_dimensions[1].height = 28

r = 2
ws8.merge_cells(f"A{r}:J{r}")
ws8["A2"].value = (
    "PUNTA = extremo conductor individual a conectar en cada cable (2 extremos por cable: lado tablero + lado carga)  |  "
    "Armadura: gland de apriete en cada extremo de cable armado  |  "
    "Nota: cablesCPF-1 excluidos del análisis (scope EPC CPF-1)"
)
ws8["A2"].font      = fnt(color=C_WH, size=8.5)
ws8["A2"].alignment = aln("center","center")
ws8["A2"].fill      = fill(C_GRY)
ws8["A2"].border    = brd()
ws8.row_dimensions[2].height = 14

# ─── SECTION A: TOTALES GENERALES ─────────────────────────────────────────────
r = 4
shdr(ws8, r, "  A. TOTALES GENERALES — CPF-2 (513 cables, excluyendo CPF-1)", C_DRD)
r += 1

col_hdr_row(ws8, r, ["Concepto","Cables","Metros","Puntas\nSala Elec","Puntas\nCampo/Carga",
                      "Puntas\nArmadura","TOTAL\nPUNTAS","","","Observación"])
r += 1

totales_rows = [
    ("TOTAL CPF-2 — conductores (puntas lado tablero)",
     tot_cables_cpf2, f"{tot_metros_cpf2:,.0f} m",
     tot_puntas_cond_sala, "—", "—", "—", "","",
     "Extremos en CCMs, Tableros, Salas Eléctricas"),
    ("TOTAL CPF-2 — conductores (puntas lado carga/campo)",
     "—", "—",
     "—", tot_puntas_cond_campo, "—", "—", "","",
     "Extremos en Motores, VFDs, Instrumentos, Luminarias, etc."),
    ("TOTAL CPF-2 — armadura (glands metálicos)",
     sum(1 for t in td_cpf2 if t.armado), "—",
     sum(t.puntas_arm_de for t in td_cpf2 if t.ub_de == 'SALA_ELEC'),
     sum(t.puntas_arm_de for t in td_cpf2 if t.ub_de != 'SALA_ELEC'),
     tot_puntas_arm, "—", "","",
     "1 gland/extremo en cables armados (Arm)"),
]
t_sala = tot_puntas_cond_sala + sum(t.puntas_arm_de for t in td_cpf2 if t.ub_de == 'SALA_ELEC') + \
         sum(t.puntas_arm_hacia for t in td_cpf2 if t.ub_hacia == 'SALA_ELEC')
t_campo = tot_puntas_cond_campo + sum(t.puntas_arm_de for t in td_cpf2 if t.ub_de != 'SALA_ELEC') + \
          sum(t.puntas_arm_hacia for t in td_cpf2 if t.ub_hacia != 'SALA_ELEC')

totales_rows.append(
    ("★ TOTAL PUNTAS CPF-2 (conductores + armaduras)",
     tot_cables_cpf2, f"{tot_metros_cpf2:,.0f} m",
     t_sala, t_campo,
     tot_puntas_arm, tot_puntas_total, "","",
     f"GRAN TOTAL: {tot_puntas_total:,} puntas individuales a terminar")
)

for i, vals in enumerate(totales_rows):
    is_total = i == len(totales_rows) - 1
    bg = "E2EFDA" if is_total else C_WHT
    fc = C_OKT if is_total else "000000"
    for col_i, v in enumerate(vals, 1):
        bold = is_total
        num_f = '#,##0' if isinstance(v, int) and col_i > 1 else None
        h = "right" if isinstance(v, int) else "left"
        cell(ws8, r, col_i, v, bold=bold, bg=bg, size=9 if not is_total else 10,
             h=h, color=fc, num_format=num_f)
    ws8.row_dimensions[r].height = 22 if is_total else 18
    r += 1

r += 1

# ─── SECTION B: POR TIPO DE CABLE ─────────────────────────────────────────────
shdr(ws8, r, "  B. PUNTAS POR TIPO DE CABLE", C_DBL)
r += 1

col_hdr_row(ws8, r, ["Tipo de Cable","Cables","Metros\ntotal","Puntas\nSala",
                      "Puntas\nCampo","Puntas\nArmadura","TOTAL\nPUNTAS",
                      "% del total","","Observación"])
r += 1

TYPE_ORDER = [
    "POTENCIA MT (>1kV)",
    "POTENCIA BT GRANDE (≥95mm²)",
    "POTENCIA BT MEDIANA (35-70mm²)",
    "POTENCIA BT PEQUEÑA (10-25mm²)",
    "POTENCIA BT MUY PEQUEÑA (<10mm²)",
    "CONTROL",
    "SEÑAL / INSTRUMENTO",
    "FIBRA ÓPTICA / ETHERNET",
    "OTROS",
]
TYPE_BG = {
    "POTENCIA MT (>1kV)":              ("FFE0CC","E36C09"),
    "POTENCIA BT GRANDE (≥95mm²)":     (C_RLT, C_DRD),
    "POTENCIA BT MEDIANA (35-70mm²)":  ("FFF0F0", C_RED),
    "POTENCIA BT PEQUEÑA (10-25mm²)":  ("FFF5F5", C_RED),
    "POTENCIA BT MUY PEQUEÑA (<10mm²)":("FFFAFA","AA0000"),
    "CONTROL":                          (C_BLT,  C_DBL),
    "SEÑAL / INSTRUMENTO":              (C_YEL2, C_WNT),
    "FIBRA ÓPTICA / ETHERNET":          ("F0F0FF","5050AA"),
    "OTROS":                            (C_GR2,  C_GRY),
}

subtot_tipo_puntas = 0
for tipo in TYPE_ORDER:
    if tipo not in by_tipo: continue
    d = by_tipo[tipo]
    arm  = d['puntas_arm_sala'] + d['puntas_arm_campo']
    ptot = d['puntas_total']
    pct  = ptot / tot_puntas_total * 100 if tot_puntas_total else 0
    metros = d['metros']
    bg, fc = TYPE_BG.get(tipo, (C_WHT, "000000"))
    vals = [tipo, d['n_cables'], int(metros),
            d['puntas_cond_sala'] + d['puntas_arm_sala'],
            d['puntas_cond_campo'] + d['puntas_arm_campo'],
            arm, ptot, f"{pct:.1f}%", "", ""]
    for ci, v in enumerate(vals, 1):
        h = "right" if isinstance(v, int) else ("center" if ci == 8 else "left")
        cell(ws8, r, ci, v, bg=bg, color=fc, size=9,
             h=h, num_format='#,##0' if isinstance(v, int) and ci>1 else None)
    ws8.row_dimensions[r].height = 18
    subtot_tipo_puntas += ptot
    r += 1

# Subtotal
cell(ws8, r, 1, "TOTAL", bold=True, bg="E2EFDA", color=C_OKT, size=9)
cell(ws8, r, 2, tot_cables_cpf2, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 3, int(tot_metros_cpf2), bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 4, t_sala, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 5, t_campo, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 6, tot_puntas_arm, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 7, tot_puntas_total, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 8, "100%", bold=True, bg="E2EFDA", color=C_OKT, size=9)
for ci in range(9, 11): cell(ws8, r, ci, "", bg="E2EFDA")
ws8.row_dimensions[r].height = 18
r += 2

# ─── SECTION C: POR TIPO DE CARGA ─────────────────────────────────────────────
shdr(ws8, r, "  C. PUNTAS POR TIPO DE CARGA / EQUIPO DESTINO (extremo campo)", C_GRN)
r += 1
col_hdr_row(ws8, r, ["Tipo de Carga / Equipo","Cables","Puntas\nSala",
                      "Puntas\nCampo","Puntas\nArmadura","TOTAL\nPUNTAS",
                      "% del total","","","Observación"], bg=C_GRN)
r += 1

CARGA_ORDER = [
    'MOTOR / ACCIONAMIENTO',
    'VFD / VARIADOR DE FRECUENCIA',
    'TABLERO / CCM',
    'UPS / RECTIFICADOR / SSAA',
    'ILUMINACIÓN',
    'CALEFACCIÓN / TRAZADO',
    'TOMACORRIENTE',
    'INSTRUMENTO / SENSOR',
    'SEÑAL / FIBRA ÓPTICA',
    'CONTROL',
    'OTROS / SIN CLASIFICAR',
]
CARGA_BG = {
    'MOTOR / ACCIONAMIENTO':          ("FFE8E0", "7B2600"),
    'VFD / VARIADOR DE FRECUENCIA':   ("FFDDC0", "7B2600"),
    'TABLERO / CCM':                   (C_RLT,   C_DRD),
    'UPS / RECTIFICADOR / SSAA':       ("FFE8F0", "880055"),
    'ILUMINACIÓN':                      ("FFFFF0", "5B5800"),
    'CALEFACCIÓN / TRAZADO':           ("FFF0D8", C_WNT),
    'TOMACORRIENTE':                    ("F8F8FF", C_GRY),
    'INSTRUMENTO / SENSOR':            (C_YEL2,  C_WNT),
    'SEÑAL / FIBRA ÓPTICA':            ("F0F0FF", "5050AA"),
    'CONTROL':                          (C_BLT,   C_DBL),
    'OTROS / SIN CLASIFICAR':          (C_GR2,   C_GRY),
}

for carga in CARGA_ORDER:
    if carga not in by_carga: continue
    d  = by_carga[carga]
    arm = d['puntas_arm_sala'] + d['puntas_arm_campo']
    ptot = d['puntas_total']
    pct  = ptot / tot_puntas_total * 100 if tot_puntas_total else 0
    bg, fc = CARGA_BG.get(carga, (C_WHT, "000000"))
    vals = [carga, d['n_cables'],
            d['puntas_cond_sala'] + d['puntas_arm_sala'],
            d['puntas_cond_campo'] + d['puntas_arm_campo'],
            arm, ptot, f"{pct:.1f}%", "", "", ""]
    for ci, v in enumerate(vals, 1):
        h = "right" if isinstance(v, int) else ("center" if ci == 7 else "left")
        cell(ws8, r, ci, v, bg=bg, color=fc, size=9,
             h=h, num_format='#,##0' if isinstance(v, int) and ci>1 else None)
    ws8.row_dimensions[r].height = 18
    r += 1

# Total
tot_carga_puntas = sum(by_carga[k]['puntas_total'] for k in by_carga)
cell(ws8, r, 1, "TOTAL", bold=True, bg="E2EFDA", color=C_OKT, size=9)
cell(ws8, r, 2, tot_cables_cpf2, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 3, t_sala,  bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 4, t_campo, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 5, tot_puntas_arm, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 6, tot_puntas_total, bold=True, bg="E2EFDA", color=C_OKT, num_format='#,##0', size=9)
cell(ws8, r, 7, "100%", bold=True, bg="E2EFDA", color=C_OKT, size=9)
for ci in range(8, 11): cell(ws8, r, ci, "", bg="E2EFDA")
ws8.row_dimensions[r].height = 18
r += 2

# ─── SECTION D: POR SALA ELÉCTRICA ────────────────────────────────────────────
shdr(ws8, r, "  D. PUNTAS EN SALAS ELÉCTRICAS (tableros, CCMs, paneles — extremo sala)", C_RED)
r += 1
col_hdr_row(ws8, r, ["Sala Eléctrica","Cables\nasociados","Metros","Puntas\nconductor",
                      "Puntas\narmadura","TOTAL\nPUNTAS SALA","% del total","","","Obs."],
            bg=C_RED)
r += 1

SALA_ORDER = ['SE-3 BT', 'SE-4 MT', 'SE-GENERACIÓN', 'CAMPO / EXTERNO']
SALA_BG = {'SE-3 BT':(C_RLT,C_DRD),'SE-4 MT':(C_BLT,C_DBL),
           'SE-GENERACIÓN':("FFF8DC","8B6914"),'CAMPO / EXTERNO':(C_GR2,C_GRY)}
sala_totales = defaultdict(lambda: {'cables':0,'metros':0,'puntas_cond':0,'puntas_arm':0,'puntas_total':0})
for td in td_cpf2:
    s = td.sala
    sala_totales[s]['cables'] += 1
    try:
        sala_totales[s]['metros'] += float(td.c['longitud']) if td.c['longitud'] else 0
    except (ValueError, TypeError):
        pass
    # Puntas en sala for this cable: DE end if in sala + HACIA end if in sala
    if td.ub_de == 'SALA_ELEC':
        sala_totales[s]['puntas_cond'] += td.puntas_cond_de
        sala_totales[s]['puntas_arm']  += td.puntas_arm_de
    if td.ub_hacia == 'SALA_ELEC':
        sala_totales[s]['puntas_cond'] += td.puntas_cond_hacia
        sala_totales[s]['puntas_arm']  += td.puntas_arm_hacia
    sala_totales[s]['puntas_total'] += (
        (td.puntas_cond_de + td.puntas_arm_de) if td.ub_de=='SALA_ELEC' else 0
    ) + (
        (td.puntas_cond_hacia + td.puntas_arm_hacia) if td.ub_hacia=='SALA_ELEC' else 0
    )

for sala in SALA_ORDER:
    if sala not in sala_totales: continue
    d = sala_totales[sala]
    ptot = d['puntas_total']
    pct  = ptot / t_sala * 100 if t_sala else 0
    bg, fc = SALA_BG.get(sala, (C_WHT, "000000"))
    vals = [sala, d['cables'], int(d['metros']),
            d['puntas_cond'], d['puntas_arm'], ptot, f"{pct:.1f}%", "", "", ""]
    for ci, v in enumerate(vals, 1):
        h = "right" if isinstance(v, int) else ("center" if ci==7 else "left")
        cell(ws8, r, ci, v, bg=bg, color=fc, size=9,
             h=h, num_format='#,##0' if isinstance(v, int) and ci>1 else None)
    ws8.row_dimensions[r].height = 18
    r += 1
ws8.row_dimensions[r].height = 18
r += 2

# ─── SECTION E: POR FORMACIÓN ─────────────────────────────────────────────────
shdr(ws8, r, "  E. PUNTAS POR FORMACIÓN DE CABLE (detalle por sección y tipo)", "404040")
r += 1
col_hdr_row(ws8, r, ["Formación","Tipo Cable","Cables","Cond/\npunta","Metros",
                      "Puntas\nSala","Puntas\nCampo","Puntas\nArmadura","TOTAL\nPUNTAS","Obs."],
            bg="404040")
r += 1

# Sort by category then section size desc
form_sorted = sorted(
    [(k, v) for k, v in by_form.items()],
    key=lambda x: (x[0][1], -get_seccion_mm2(x[0][0]))
)
prev_tipo_form = None
for (form, tipo), d in form_sorted:
    if tipo != prev_tipo_form:
        ws8.merge_cells(f"A{r}:J{r}")
        ws8[f"A{r}"].value = f"  — {tipo} —"
        ws8[f"A{r}"].font      = fnt(bold=True, color=C_WH, size=8.5)
        ws8[f"A{r}"].alignment = aln("left","center")
        ws8[f"A{r}"].fill      = fill(TYPE_BG.get(tipo, (C_GR2,"000000"))[0])
        ws8.row_dimensions[r].height = 14
        prev_tipo_form = tipo
        r += 1

    n_cond, _ = parse_conductores_punta(form)
    arm  = d['puntas_arm_sala'] + d['puntas_arm_campo']
    ptot = d['puntas_total']
    bg, fc = TYPE_BG.get(tipo, (C_WHT, "000000"))
    vals = [form, tipo, d['n_cables'], n_cond, int(d['metros']),
            d['puntas_cond_sala'] + d['puntas_arm_sala'],
            d['puntas_cond_campo'] + d['puntas_arm_campo'],
            arm, ptot, ""]
    for ci, v in enumerate(vals, 1):
        h = "right" if isinstance(v, int) else ("left" if ci in (1,2,10) else "center")
        cell(ws8, r, ci, v, bg=bg, color=fc, size=9,
             h=h, num_format='#,##0' if isinstance(v, int) and ci>2 else None)
    ws8.row_dimensions[r].height = 16
    r += 1

r += 1

# ─── SECTION F: CPF-1 + CPF-2 COMBINED ───────────────────────────────────────
shdr(ws8, r, "  F. COMPARATIVO CPF-1 vs CPF-2 (solo referencia — scope CPF-1 es EPC externo)", C_PURP)
r += 1
col_hdr_row(ws8, r, ["CPF","Tipo Cable","Cables","Metros","Puntas\nSala",
                      "Puntas\nCampo","Puntas\nArmadura","TOTAL\nPUNTAS","",""], bg=C_PURP)
r += 1

for (cpf, tipo), d in sorted(by_cpf_tipo.items(), key=lambda x: (x[0][0], x[0][1])):
    arm  = d['puntas_arm_sala'] + d['puntas_arm_campo']
    bg = "DEEAF1" if cpf=='CPF2' else "FFF0F0"
    fc = C_DBL  if cpf=='CPF2' else C_DRD
    vals = [cpf, tipo, d['n_cables'], int(d['metros']),
            d['puntas_cond_sala'] + d['puntas_arm_sala'],
            d['puntas_cond_campo'] + d['puntas_arm_campo'],
            arm, d['puntas_total'], "", ""]
    for ci, v in enumerate(vals, 1):
        h = "right" if isinstance(v, int) else "left"
        cell(ws8, r, ci, v, bg=bg, color=fc, size=9,
             h=h, num_format='#,##0' if isinstance(v, int) and ci>2 else None)
    ws8.row_dimensions[r].height = 16
    r += 1

r += 2

# ─── SECTION G: NOTA METODOLÓGICA ─────────────────────────────────────────────
shdr(ws8, r, "  G. METODOLOGÍA Y SUPUESTOS", C_GRY)
r += 1
notas = [
    ("Definición de PUNTA:",
     "Cada conductor individual (hilo o filamento) en un extremo del cable que debe ser conectado, "
     "ya sea mediante terminal a compresión (lug), bornera, conector, o soldadura."),
    ("Armadura:",
     "Cable armado (CONST contiene 'Arm'): se contabiliza 1 gland o prensaestopa metálico por extremo. "
     "Total: 2 puntas armadura por cable armado."),
    ("Formaciones trifásicas (3xS):",
     "3 conductores/punta. Si lleva neutro o PE separado (3xS/N) → 4 conductores/punta."),
    ("Formaciones control (NxS+T):",
     "N conductores + 1 drenaje = N+1 conductores/punta."),
    ("Cables de señal multi-par (NxMxS):",
     "N×M conductores/punta (e.g. 7x2x0.5 = 14 conductores/punta por par de hilos)."),
    ("Fibra óptica:",
     "Se contabilizan como conectores (SC/APC, LC, etc.) — no como conductores de cobre."),
    ("Lado Sala Eléctrica:",
     "Extremo que conecta en CCM, Tablero, TGMT, TGBT, DP, LP, UPS, DB, PMS dentro de Sala Eléctrica."),
    ("Lado Campo/Carga:",
     "Extremo que conecta en el equipo de proceso: motor, VFD, instrumento, luminaria, tomacorriente, "
     "calefactor, tablero local en campo, etc."),
    ("CPF-1 vs CPF-2:",
     "Solo cables CPF-2 son scope de este proyecto. CPF-1 incluido en sección F solo como referencia. "
     "Total proyecto: " + f"{len(td_all)} cables / {sum(t.puntas_total for t in td_all):,} puntas totales."),
]
for concepto, nota in notas:
    ws8.merge_cells(f"A{r}:B{r}")
    ws8[f"A{r}"].value = concepto
    ws8[f"A{r}"].font  = fnt(bold=True, size=9, color=C_DBL)
    ws8[f"A{r}"].alignment = aln("left","center")
    ws8[f"A{r}"].fill  = fill(C_BLT)
    ws8.merge_cells(f"C{r}:J{r}")
    ws8[f"C{r}"].value = nota
    ws8[f"C{r}"].font  = fnt(size=9, italic=True)
    ws8[f"C{r}"].alignment = aln("left","center", wrap=True)
    ws8[f"C{r}"].fill  = fill("F8FBFF")
    ws8.row_dimensions[r].height = 28
    r += 1

ws8.freeze_panes = "A4"

# ── UPDATE TITLES IN EXISTING SHEETS WITH REV1 ───────────────────────────────
for ws in wb.worksheets:
    if ws.title == "8_Puntas_Conexion": continue
    for row_cells in ws.iter_rows():
        for c in row_cells:
            if c.value and isinstance(c.value, str):
                if "27-MAY-2026" in c.value and "Rev 1" not in c.value:
                    # Don't double-update
                    pass
                # Mark Rev1 on title cells that contain the project title
                if ("CRONOGRAMA" in c.value.upper() or
                    "ANÁLISIS" in c.value.upper() or
                    "ELÉCTRICO CPF2" in c.value.upper() or
                    "ELECTRICAL ANALYSIS" in c.value.upper()):
                    if REV not in c.value and FECHA not in c.value:
                        c.value = c.value.rstrip() + f"  |  {REV}  |  {FECHA}"

# Update Sheet 6 (Resumen Ejecutivo) to add puntas summary
if "6_Resumen_Ejecutivo" in wb.sheetnames:
    ws6e = wb["6_Resumen_Ejecutivo"]
    max_r = ws6e.max_row + 3
    ws6e.merge_cells(f"A{max_r}:G{max_r}")
    ws6e.cell(max_r, 1).value = (f"★ NUEVO {REV} {FECHA}: Análisis de Puntas de Conexión — "
                                  f"CPF-2: {tot_puntas_total:,} puntas totales "
                                  f"({tot_puntas_cond_sala + sum(t.puntas_arm_de for t in td_cpf2 if t.ub_de=='SALA_ELEC'):,} en salas / "
                                  f"{tot_puntas_cond_campo + sum(t.puntas_arm_de for t in td_cpf2 if t.ub_de!='SALA_ELEC'):,} en campo). "
                                  f"Ver hoja 8_Puntas_Conexion.")
    ws6e.cell(max_r, 1).font  = fnt(bold=True, color=C_WH, size=10)
    ws6e.cell(max_r, 1).alignment = aln("left", "center", wrap=True)
    ws6e.cell(max_r, 1).fill  = fill(C_GRN)
    ws6e.row_dimensions[max_r].height = 28

# ── REORDER SHEETS ────────────────────────────────────────────────────────────
desired = ["1_Cables_Completo","2_Resumen_Formaciones","3_Cargas_CPF2",
           "4_Estimacion_Instalacion","5_Cronograma_Campo","6_Resumen_Ejecutivo",
           "8_Puntas_Conexion"]
for i, name in enumerate(desired):
    if name in wb.sheetnames:
        idx = wb.sheetnames.index(name)
        if idx != i:
            try: wb.move_sheet(name, offset=i - idx)
            except: pass

# ── SAVE ──────────────────────────────────────────────────────────────────────
# Generate output with revision in filename
OUT = os.path.join(PLANNING_DIR, f"Analisis_Electrico_CPF2_Campo_Rev1_{HOY.strftime('%Y-%m-%d')}.xlsx")
wb.save(OUT)
print(f"\n✓ Guardado: {OUT}")
print(f"\nHojas:")
for s in wb.sheetnames:
    print(f"  - {s}")
print(f"\n{'='*60}")
print(f"RESUMEN ANÁLISIS DE PUNTAS — CPF-2")
print(f"{'='*60}")
print(f"  Cables CPF-2:                   {tot_cables_cpf2:>6,}")
print(f"  Metros CPF-2:                   {int(tot_metros_cpf2):>6,} m")
print(f"  Conductores / punta total:      {tot_puntas_cond_sala + tot_puntas_cond_campo:>6,}")
print(f"    Puntas en salas eléctricas:   {tot_puntas_cond_sala:>6,}")
print(f"    Puntas en campo/cargas:       {tot_puntas_cond_campo:>6,}")
print(f"  Puntas armadura (glands):       {tot_puntas_arm:>6,}")
print(f"  TOTAL PUNTAS CPF-2:             {tot_puntas_total:>6,}")
print(f"\nDETALLE POR TIPO:")
for tipo in TYPE_ORDER:
    if tipo not in by_tipo: continue
    d = by_tipo[tipo]
    print(f"  {tipo:<42} {d['n_cables']:>4} cables  {d['puntas_total']:>5,} puntas")
print(f"\nDETALLE POR CARGA:")
for c_tipo in CARGA_ORDER:
    if c_tipo not in by_carga: continue
    d = by_carga[c_tipo]
    print(f"  {c_tipo:<45} {d['n_cables']:>4} cables  {d['puntas_total']:>5,} puntas")
