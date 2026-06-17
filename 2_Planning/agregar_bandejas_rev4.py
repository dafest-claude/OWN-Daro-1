#!/usr/bin/env python3
"""
Rev4: Agrega hoja 10_Bandejas_Canalizaciones al análisis eléctrico.
Fuente: ACAL-00102-LM-E-0018 (LISTA DE MATERIALES - CANALIZACIONES)
"""

import shutil
from openpyxl import load_workbook
from openpyxl.styles import (
    Font, PatternFill, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

SRC = "Analisis_Electrico_CPF2_Campo_Rev3_2026-06-13.xlsx"
DST = "Analisis_Electrico_CPF2_Campo_Rev4_2026-06-17.xlsx"

REV  = "Rev 4"
DATE = "17-JUN-2026"

# ── colour palette ──────────────────────────────────────────────────────────
C_TITLE   = "1F3864"   # dark navy  (white font)
C_SEC_A   = "2E75B6"   # bandeja escalera  (white)
C_SEC_B   = "4472C4"   # tapas             (white)
C_SEC_C   = "70AD47"   # curvas horiz      (white)
C_SEC_D   = "ED7D31"   # curvas vert       (white)
C_SEC_E   = "FF0000"   # tees              (white)
C_SEC_F   = "7030A0"   # reducciones       (white)
C_SEC_G   = "375623"   # accesorios        (white)
C_SEC_H   = "833C00"   # caños             (white)
C_HEAD    = "D6E4F0"   # sub-header row    (dark font)
C_EVEN    = "EBF3FB"   # alternating row
C_TOTAL   = "FFF2CC"   # subtotal / total row

def fill(hex_color):
    return PatternFill("solid", fgColor=hex_color)

def font(bold=False, color="000000", size=10):
    return Font(bold=bold, color=color, size=size, name="Calibri")

thin = Side(style="thin", color="BFBFBF")
thick= Side(style="medium", color="000000")

def border_all():
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def border_thick_bottom():
    return Border(left=thin, right=thin, top=thin, bottom=thick)

def center(wrap=False):
    return Alignment(horizontal="center", vertical="center", wrap_text=wrap)

def left(wrap=False):
    return Alignment(horizontal="left", vertical="center", wrap_text=wrap)

# ── workbook ─────────────────────────────────────────────────────────────────
shutil.copy(SRC, DST)
wb = load_workbook(DST)

# ── update revision / date in existing sheets ────────────────────────────────
REV_KEYWORDS = ["Rev", "Revisión", "REVISION", "rev"]
DATE_KEYWORDS = ["Fecha", "FECHA", "fecha", "Date"]

def update_cell_if_rev(cell):
    v = str(cell.value) if cell.value else ""
    for kw in ["Rev 3", "Rev3", "REV3", "REV 3"]:
        if kw in v:
            cell.value = v.replace(kw, REV)
            return
    for kw in ["2026-06-13", "13-JUN-2026", "13/06/2026"]:
        if kw in v:
            cell.value = v.replace(kw, DATE)
            return

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                update_cell_if_rev(cell)

# ── create new sheet ──────────────────────────────────────────────────────────
if "10_Bandejas_Canalizaciones" in wb.sheetnames:
    del wb["10_Bandejas_Canalizaciones"]
ws = wb.create_sheet("10_Bandejas_Canalizaciones")

# column widths (A..K)
col_w = {
    "A": 6,   # Item
    "B": 22,  # Código
    "C": 52,  # Descripción
    "D": 8,   # Unidad
    "E": 10,  # Rev0
    "F": 8,   # Surplus
    "G": 10,  # Total cant.
    "H": 14,  # Long/tramo (m)
    "I": 14,  # Total metros
}
for col, w in col_w.items():
    ws.column_dimensions[col].width = w

ws.row_dimensions[1].height = 30
ws.row_dimensions[2].height = 18
ws.row_dimensions[3].height = 14

# ── helpers ──────────────────────────────────────────────────────────────────
_row = [1]

def cur_row():
    return _row[0]

def next_row():
    _row[0] += 1
    return _row[0]

def write_row(values, bg=None, bold=False, fgc="000000", num_fmt=None,
              align_fn=None, border_fn=None):
    r = _row[0]
    for ci, (col, val) in enumerate(zip("ABCDEFGHI", values), start=1):
        cell = ws.cell(row=r, column=ci, value=val)
        if bg:
            cell.fill = fill(bg)
        cell.font = font(bold=bold, color=fgc)
        if align_fn:
            cell.alignment = align_fn()
        else:
            cell.alignment = left()
        cell.border = border_all()
        if num_fmt and ci >= 5:
            cell.number_format = num_fmt
    _row[0] += 1
    return r

def merge_title(text, bg, fgc="FFFFFF", height=20):
    r = _row[0]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    cell = ws.cell(row=r, column=1, value=text)
    cell.fill = fill(bg)
    cell.font = font(bold=True, color=fgc, size=11)
    cell.alignment = center()
    cell.border = border_thick_bottom()
    ws.row_dimensions[r].height = height
    _row[0] += 1
    return r

def col_headers(bg=C_HEAD):
    hdrs = ["Item", "Código", "Descripción", "Unid.",
            "Rev.0", "Surplus", "Total", "m/tramo", "Total m"]
    r = _row[0]
    for ci, h in enumerate(hdrs, start=1):
        cell = ws.cell(row=r, column=ci, value=h)
        cell.fill = fill(bg)
        cell.font = font(bold=True)
        cell.alignment = center(wrap=True)
        cell.border = border_all()
    ws.row_dimensions[r].height = 30
    _row[0] += 1

def data_row(item, code, desc, unit, rev0, sur, total,
             m_tramo="", bg=None, bold=False):
    total_m = ""
    if isinstance(total, (int, float)) and isinstance(m_tramo, (int, float)):
        total_m = total * m_tramo
    values = [item, code, desc, unit, rev0, sur, total, m_tramo, total_m]
    r = _row[0]
    bg_use = bg if bg else ("FFFFFF" if r % 2 == 0 else C_EVEN)
    for ci, val in enumerate(values, start=1):
        cell = ws.cell(row=r, column=ci, value=val)
        cell.fill = fill(bg_use)
        cell.font = font(bold=bold)
        if ci in (1, 4, 5, 6, 7, 8, 9):
            cell.alignment = center()
        else:
            cell.alignment = left(wrap=True)
        cell.border = border_all()
        if ci in (5, 6, 7, 9) and isinstance(val, (int, float)):
            cell.number_format = "#,##0"
        if ci == 8 and isinstance(val, (int, float)):
            cell.number_format = "#,##0.00"
    _row[0] += 1
    return r

def subtotal_row(label, total_m_sum):
    r = _row[0]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c = ws.cell(row=r, column=1, value=label)
    c.fill = fill(C_TOTAL)
    c.font = font(bold=True)
    c.alignment = left()
    c.border = border_all()
    c8 = ws.cell(row=r, column=8, value="SUBTOTAL")
    c8.fill = fill(C_TOTAL)
    c8.font = font(bold=True)
    c8.alignment = center()
    c8.border = border_all()
    c9 = ws.cell(row=r, column=9, value=total_m_sum)
    c9.fill = fill(C_TOTAL)
    c9.font = font(bold=True)
    c9.alignment = center()
    c9.border = border_all()
    if isinstance(total_m_sum, (int, float)):
        c9.number_format = "#,##0.00"
    _row[0] += 1

def blank_row():
    r = _row[0]
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=9)
    _row[0] += 1

# ============================================================
# TITLE BLOCK
# ============================================================
ws.merge_cells("A1:I1")
t = ws["A1"]
t.value = "CÓMPUTO DE BANDEJAS Y CANALIZACIONES — CAMPO CPF-2"
t.fill = fill(C_TITLE)
t.font = font(bold=True, color="FFFFFF", size=13)
t.alignment = center()
t.border = border_all()

ws.merge_cells("A2:I2")
s = ws["A2"]
s.value = (f"Fuente: ACAL-00102-LM-E-0018 / 5155-00-0000-IG-EL-LM-003  |  "
           f"Rev 0 emitida 11-JUN-2026  |  {REV} — {DATE}")
s.fill = fill(C_HEAD)
s.font = font(bold=False, size=9)
s.alignment = center()
s.border = border_all()

ws.merge_cells("A3:I3")
n = ws["A3"]
n.value = ("Nota: tramos bandejas = 6 000 mm c/u  |  tapas = 3 000 mm c/u  |  "
           "curvas / tees / reducciones = unidades  |  caños = metros")
n.fill = fill("FFF9C4")
n.font = font(size=8, color="5C4600")
n.alignment = center(wrap=True)
n.border = border_all()
_row[0] = 4

# ============================================================
# A — BANDEJAS TIPO ESCALERA
# ============================================================
blank_row()
merge_title("A — BANDEJAS TIPO ESCALERA (tramo 6 000 mm)", C_SEC_A)
col_headers()

bandeja_data = [
    ("1.2",  "XEBTS_14030006", "Bandeja tipo escalera 150 mm × 6 000 mm", "tramo", 114, 6,  120, 6.0),
    ("1.23", "XEBTS_14030007", "Bandeja tipo escalera 300 mm × 6 000 mm", "tramo", 162, 8,  170, 6.0),
    ("1.44", "XEBTS_14030008", "Bandeja tipo escalera 450 mm × 6 000 mm", "tramo",  24, 1,   25, 6.0),
    ("1.65", "XEBTS_14030009", "Bandeja tipo escalera 600 mm × 6 000 mm", "tramo", 488, 24, 512, 6.0),
]
sub_a = 0.0
for row in bandeja_data:
    data_row(*row)
    sub_a += row[6] * row[7]
subtotal_row("Subtotal bandejas tipo escalera", sub_a)

# ============================================================
# B — TAPAS BANDEJAS
# ============================================================
blank_row()
merge_title("B — TAPAS BANDEJAS (tramo 3 000 mm)", C_SEC_B)
col_headers()

tapa_data = [
    ("1.3",  "XEBCD_14030830", "Tapa bandeja 150 mm × 3 000 mm", "tramo", 228, 11, 239, 3.0),
    ("1.24", "XEBCD_14030832", "Tapa bandeja 300 mm × 3 000 mm", "tramo", 181,  9, 190, 3.0),
    ("1.45", "XEBCD_14030834", "Tapa bandeja 450 mm × 3 000 mm", "tramo",  47,  2,  49, 3.0),
    ("1.66", "XEBCD_14030836", "Tapa bandeja 600 mm × 3 000 mm", "tramo", 472, 24, 496, 3.0),
]
sub_b = 0.0
for row in tapa_data:
    data_row(*row)
    sub_b += row[6] * row[7]
subtotal_row("Subtotal tapas bandejas", sub_b)

# ============================================================
# C — CURVAS HORIZONTALES
# ============================================================
blank_row()
merge_title("C — CURVAS HORIZONTALES (unidades)", C_SEC_C)
col_headers()

curva_h_data = [
    # 90°
    ("1.5",  "XEBCQ_15310393", "Curva horizontal 90° 150 mm",            "u",  16, 1, 17, ""),
    ("1.6",  "XEBCK_14030724", "Tapa curva horizontal 90° 150 mm",       "u",  16, 1, 17, ""),
    ("1.26", "XEBCQ_15310399", "Curva horizontal 90° 300 mm",            "u",  12, 1, 13, ""),
    ("1.27", "XEBCK_14030728", "Tapa curva horizontal 90° 300 mm",       "u",   7, 0,  7, ""),
    ("1.47", "XEBCQ_15310405", "Curva horizontal 90° 450 mm",            "u",   5, 0,  5, ""),
    ("1.48", "XEBCK_14030732", "Tapa curva horizontal 90° 450 mm",       "u",   5, 0,  5, ""),
    ("1.68", "XEBCQ_15310412", "Curva horizontal 90° 600 mm",            "u",  27, 1, 28, ""),
    ("1.69", "XEBCK_14030736", "Tapa curva horizontal 90° 600 mm",       "u",  20, 1, 21, ""),
    # 45°
    ("1.29", "XEBCQ_15310395", "Curva horizontal 45° 300 mm",            "u",   4, 0,  4, ""),
    ("1.30", "XEBCK_14030727", "Tapa curva horizontal 45° 300 mm",       "u",   4, 0,  4, ""),
    ("1.50", "XEBCQ_15310401", "Curva horizontal 45° 450 mm",            "u",   2, 0,  2, ""),
    ("1.51", "XEBCK_14030731", "Tapa curva horizontal 45° 450 mm",       "u",   2, 0,  2, ""),
    ("1.71", "XEBCQ_15310408", "Curva horizontal 45° 600 mm",            "u",   2, 0,  2, ""),
    ("1.72", "XEBCK_14030735", "Tapa curva horizontal 45° 600 mm",       "u",   2, 0,  2, ""),
]
for row in curva_h_data:
    data_row(*row)
# no metros sub-total for unit items — show "—"
subtotal_row("Subtotal curvas horizontales (unidades)", "—")

# ============================================================
# D — CURVAS VERTICALES
# ============================================================
blank_row()
merge_title("D — CURVAS VERTICALES (unidades)", C_SEC_D)
col_headers()

curva_v_data = [
    # Exteriores
    ("1.231", "XEBVF_14030390", "Curva vertical exterior 150 mm",        "u",  19, 0, 19, ""),
    ("1.232", "XEBCO_14031276", "Tapa curva vertical exterior 150 mm",   "u",  19, 0, 19, ""),
    ("1.233", "XEBVF_14030392", "Curva vertical exterior 300 mm",        "u",  12, 1, 13, ""),
    ("1.234", "XEBCO_14031275", "Tapa curva vertical exterior 300 mm",   "u",   7, 0,  7, ""),
    ("1.235", "XEBVF_14030394", "Curva vertical exterior 450 mm",        "u",   4, 0,  4, ""),
    ("1.236", "XEBCO_11843379", "Tapa curva vertical exterior 450 mm",   "u",   4, 0,  4, ""),
    ("1.237", "XEBVF_14030396", "Curva vertical exterior 600 mm",        "u",  34, 2, 36, ""),
    ("1.238", "XEBCO_11843381", "Tapa curva vertical exterior 600 mm",   "u",  17, 0, 17, ""),
    # Interiores
    ("1.223", "XEBVJ_14030475", "Curva vertical interior 150 mm",        "u",  17, 1, 18, ""),
    ("1.224", "XEBCV_14031277", "Tapa curva vertical interior 150 mm",   "u",  18, 0, 18, ""),
    ("1.225", "XEBVJ_14030477", "Curva vertical interior 300 mm",        "u",  11, 1, 12, ""),
    ("1.226", "XEBCV_14031273", "Tapa curva vertical interior 300 mm",   "u",   6, 0,  6, ""),
    ("1.227", "XEBVJ_14030479", "Curva vertical interior 450 mm",        "u",   4, 0,  4, ""),
    ("1.228", "XEBCV_11843341", "Tapa curva vertical interior 450 mm",   "u",   4, 0,  4, ""),
    ("1.229", "XEBVJ_14030481", "Curva vertical interior 600 mm",        "u",  34, 2, 36, ""),
    ("1.230", "XEBCV_11843343", "Tapa curva vertical interior 600 mm",   "u",  17, 0, 17, ""),
]
for row in curva_v_data:
    data_row(*row)
subtotal_row("Subtotal curvas verticales (unidades)", "—")

# ============================================================
# E — TEES HORIZONTALES
# ============================================================
blank_row()
merge_title("E — TEES HORIZONTALES (unidades)", C_SEC_E)
col_headers()

tee_data = [
    ("1.17", "XEBTQ_15100014", "Tee horizontal 150 mm",             "u", 1, 0, 1, ""),
    ("1.18", "XEBHT_14030822", "Tapa tee horizontal 150 mm",        "u", 1, 0, 1, ""),
    ("1.80", "XEBTQ_15100017", "Tee horizontal 600 mm",             "u", 5, 0, 5, ""),
    ("1.81", "XEBHT_14030828", "Tapa tee horizontal 600 mm",        "u", 3, 0, 3, ""),
]
for row in tee_data:
    data_row(*row)
subtotal_row("Subtotal tees horizontales (unidades)", "—")

# ============================================================
# F — REDUCCIONES
# ============================================================
blank_row()
merge_title("F — REDUCCIONES (unidades)", C_SEC_F)
col_headers()

reduc_data = [
    # Centrales
    ("1.86",  "XEBRB_15100031", "Reducción central 300→150 mm",         "u", 1, 0, 1, ""),
    ("1.87",  "XEBCJ_14030770", "Tapa reducción central 300→150 mm",    "u", 1, 0, 1, ""),
    ("1.92",  "XEBRB_15100033", "Reducción central 450→300 mm",         "u", 1, 0, 1, ""),
    ("1.93",  "XEBCJ_14030774", "Tapa reducción central 450→300 mm",    "u", 1, 0, 1, ""),
    ("1.95",  "XEBRB_15100034", "Reducción central 600→150 mm",         "u", 2, 0, 2, ""),
    ("1.96",  "XEBCJ_14030778", "Tapa reducción central 600→150 mm",    "u", 1, 0, 1, ""),
    ("1.98",  "XEBRB_15100035", "Reducción central 600→300 mm",         "u", 1, 0, 1, ""),
    ("1.99",  "XEBCJ_14030779", "Tapa reducción central 600→300 mm",    "u", 1, 0, 1, ""),
    ("1.101", "XEBRB_15100036", "Reducción central 600→450 mm",         "u", 1, 0, 1, ""),
    ("1.102", "XEBCJ_14030780", "Tapa reducción central 600→450 mm",    "u", 1, 0, 1, ""),
    # Lateral derecha
    ("1.134", "XEBRE_15100041", "Reducción lateral derecha 600→300 mm", "u", 2, 0, 2, ""),
    ("1.135", "XEBRR_14030803", "Tapa reducción lateral der. 600→300 mm","u",2, 0, 2, ""),
]
for row in reduc_data:
    data_row(*row)
subtotal_row("Subtotal reducciones (unidades)", "—")

# ============================================================
# G — ACCESORIOS
# ============================================================
blank_row()
merge_title("G — ACCESORIOS (unidades)", C_SEC_G)
col_headers()

acces_data = [
    ("1.141", "XEBFC_14064633", "Grampa de fijación bandeja",  "u", 2148, 107, 2255, ""),
    ("1.143", "XEBSP_14064553", "Cupla de unión bandeja",      "u", 2148, 107, 2255, ""),
]
for row in acces_data:
    data_row(*row)
subtotal_row("Subtotal accesorios (unidades)", "—")

# ============================================================
# H — CAÑOS (CONDUIT)
# ============================================================
blank_row()
merge_title("H — CAÑOS CONDUIT (metros)", C_SEC_H)
col_headers()

cano_data = [
    ("1.148", "XLBMBF14151901", "Caño AG 1\"  (25 mm)",           "m",  300, 15,  315, 1.0),
    ("1.152", "XLBMBF14151902", "Caño AG 1½\" (40 mm)",           "m",  300, 15,  315, 1.0),
    ("1.156", "XLBMBF14151906", "Caño AG 12\" (300 mm)",          "m",  300, 15,  315, 1.0),
    ("1.173", "XLBNMC12143074", "Caño PVC 6\" (160 mm OD)",       "m",  600, 30,  630, 1.0),
]
sub_h = 0.0
for row in cano_data:
    data_row(*row)
    sub_h += row[6] * row[7]
subtotal_row("Subtotal caños conduit", sub_h)

# ============================================================
# RESUMEN TOTAL
# ============================================================
blank_row()
r_res = _row[0]
ws.merge_cells(start_row=r_res, start_column=1, end_row=r_res, end_column=9)
c = ws.cell(row=r_res, column=1,
            value="RESUMEN — METROS LINEALES DE BANDEJA Y CAÑO")
c.fill = fill(C_TITLE)
c.font = font(bold=True, color="FFFFFF", size=11)
c.alignment = center()
c.border = border_all()
ws.row_dimensions[r_res].height = 22
_row[0] += 1

res_rows = [
    ("A", "Bandejas tipo escalera", sub_a),
    ("B", "Tapas bandejas",         sub_b),
    ("H", "Caños conduit",          sub_h),
]
grand_m = sub_a + sub_b + sub_h

for sec, desc, val in res_rows:
    r = _row[0]
    ws.cell(row=r, column=1, value=sec).fill   = fill(C_HEAD)
    ws.cell(row=r, column=1).font  = font(bold=True)
    ws.cell(row=r, column=1).alignment = center()
    ws.cell(row=r, column=1).border = border_all()
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    ws.cell(row=r, column=2, value=desc).fill  = fill(C_EVEN)
    ws.cell(row=r, column=2).alignment = left()
    ws.cell(row=r, column=2).border = border_all()
    cv = ws.cell(row=r, column=9, value=val)
    cv.fill = fill(C_EVEN)
    cv.font = font(bold=True)
    cv.alignment = center()
    cv.number_format = "#,##0.00"
    cv.border = border_all()
    _row[0] += 1

r_grand = _row[0]
ws.merge_cells(start_row=r_grand, start_column=1, end_row=r_grand, end_column=8)
cg = ws.cell(row=r_grand, column=1,
             value="TOTAL METROS LINEALES (bandejas + tapas + caños)")
cg.fill = fill(C_TOTAL)
cg.font = font(bold=True, size=11)
cg.alignment = left()
cg.border = border_thick_bottom()
cv_g = ws.cell(row=r_grand, column=9, value=grand_m)
cv_g.fill = fill(C_TOTAL)
cv_g.font = font(bold=True, size=11)
cv_g.alignment = center()
cv_g.number_format = "#,##0.00"
cv_g.border = border_thick_bottom()
ws.row_dimensions[r_grand].height = 22

# freeze header rows
ws.freeze_panes = "A4"

wb.save(DST)
print(f"Guardado: {DST}")
print(f"  Bandejas escalera : {sub_a:,.1f} m")
print(f"  Tapas bandejas    : {sub_b:,.1f} m")
print(f"  Caños conduit     : {sub_h:,.1f} m")
print(f"  TOTAL metros lin. : {grand_m:,.1f} m")
