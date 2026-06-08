#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cronograma de Certificación IB – Curva S de Pagos
OC S3 (4508944971) y OC S4 (4508944973)
Hito 1: 10% al envío RevA  |  Hito 2: 10% a la aprobación Rev0
"""

from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side, numbers
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.series import DataPoint
from datetime import date, timedelta
from collections import defaultdict
import os

# ── PALETTE ──────────────────────────────────────────────────────────────────
C = {
    "s3":       "CC0000",  # ABB red – OC S3
    "s3_dk":    "880000",
    "s3_lt":    "F4CCCC",
    "s4":       "2E75B6",  # blue – OC S4
    "s4_dk":    "1F4E79",
    "s4_lt":    "CFE2F3",
    "comb":     "375623",  # dark green – combined
    "comb_lt":  "D9EAD3",
    "rev0":     "7030A0",  # purple – Rev0 approvals
    "rev0_lt":  "EAD1DC",
    "orng":     "E36C09",
    "yell":     "FFEB9C",
    "warn":     "FF9900",
    "gray":     "595959",
    "lgray":    "D9D9D9",
    "llgray":   "F2F2F2",
    "white":    "FFFFFF",
    "black":    "000000",
    "hdr":      "1F4E79",
    "hdr_s3":   "8B0000",
    "hdr_s4":   "1F4E79",
}

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
    if bg:
        c.fill = fill(bg)
    if border:
        c.border = brd()
    if num_format:
        c.number_format = num_format
    return c

# ── DATA ─────────────────────────────────────────────────────────────────────
# (pos, tag, description, price_usd, revA_date, rev0_date)
S3_ITEMS = [
    (70,  "102-DP-VFD-24210A/B/C", "VFD Variador Frec. Bombas BT (×3)",
     242508.12, date(2026,6,15), date(2026,6,30)),
    (80,  "102-DP-001-TRA",        "Tablero Distribución Tracing",
     95632.73,  date(2026,6,8),  date(2026,7,10)),
    (90,  "102-DP-001",            "Tablero Serv. Aux. SE#3",
     105885.06, date(2026,6,8),  date(2026,7,10)),
    (100, "102-LP-002",            "Tablero Distribución UPS",
     63485.70,  date(2026,6,8),  date(2026,7,10)),
    (110, "102-LP-001",            "Tablero Distribución Serv. Aux. S#3",
     96409.77,  date(2026,6,8),  date(2026,7,10)),
    (120, "102-DB-007A/B",         "Ducto de Barras 3F+N  DB-007",
     169501.43, date(2026,6,26), date(2026,7,16)),
    (130, "102-DB-006A/B",         "Ducto de Barras 3F+N  DB-006",
     182075.08, date(2026,6,26), date(2026,7,16)),
    (140, "102-DB-005A/B",         "Ducto de Barras 3F+N  DB-005",
     139094.32, date(2026,6,26), date(2026,7,16)),
    (150, "102-CCM-007",           "CCM 380/220V lote 1",
     451298.47, date(2026,8,10), date(2026,8,25)),
    (160, "102-CCM-006",           "CCM 380/220V lote 1",
     492942.64, date(2026,8,10), date(2026,8,25)),
    (170, "102-CCM-005",           "CCM 380/220V lote 1",
     552737.98, date(2026,8,10), date(2026,8,25)),
    (180, "SE#3 MT",               "Sala Eléctrica N°3 BT lote 1",
     622817.07, date(2026,7,29), date(2026,8,17)),
    (190, "102-UPS-001A/B",        "Sistema UPS Corriente Alterna",
     550434.32, date(2026,6,12), date(2026,6,22)),
    (200, "102-VCC-001A/B",        "Sistema Corriente Continua VCC",
     274833.91, date(2026,6,12), date(2026,6,22)),
    (210, "102-CCM-007",           "CCM 380/220V lote 2",
     451298.47, date(2026,8,10), date(2026,8,25)),
    (220, "102-CCM-006",           "CCM 380/220V lote 2",
     492942.64, date(2026,8,10), date(2026,8,25)),
    (230, "102-CCM-005",           "CCM 380/220V lote 2",
     552737.98, date(2026,8,10), date(2026,8,25)),
    (240, "SE#3 MT",               "Sala Eléctrica N°3 BT lote 2",
     622817.07, date(2026,7,29), date(2026,8,17)),
    (250, "SE#3 MT",               "Sala Eléctrica N°3 BT lote 3",
     622817.07, date(2026,7,29), date(2026,8,17)),
    (260, "SE#3 MT",               "Sala Eléctrica N°3 BT lote 4",
     622817.07, date(2026,7,29), date(2026,8,17)),
    (270, "SE#3 MT",               "Sala Eléctrica N°3 BT lote 5",
     622817.07, date(2026,7,29), date(2026,8,17)),
]

S4_ITEMS = [
    (60,  "102-DP-101UPS",         "Tablero Distribución UPS SE#4",
     25305.57,  date(2026,6,12), date(2026,6,22)),
    (70,  "102-DP-101VCC",         "Tablero Distribución VCC SE#4",
     29291.23,  date(2026,6,12), date(2026,6,22)),
    (80,  "102-DP-101",            "Tablero Serv. Aux. SE#4",
     66110.81,  date(2026,6,8),  date(2026,7,6)),
    (90,  "102-DP-VFD-23310A/B/C", "VFD Variador Frec. Bombas MT (×3)",
     396952.56, date(2026,6,15), date(2026,6,30)),
    (100, "102-TGMT-001/003",      "SWG MT 13.2kV / 6.6kV – Celdas",
     517484.58, date(2026,6,12), date(2026,7,31)),
    (110, "102-TGMT-001/003",      "SWG MT – Tableros lote 1",
     386434.73, date(2026,6,12), date(2026,7,31)),
    (120, "SE#4 MT",               "Sala Eléctrica N°4 MT lote 1",
     605581.31, date(2026,7,21), date(2026,8,10)),
    (130, "102-TGMT-001/003",      "SWG MT – Tableros lote 2",
     386434.73, date(2026,6,12), date(2026,7,31)),
    (140, "SE#4 MT",               "Sala Eléctrica N°4 MT lote 2",
     605581.31, date(2026,7,21), date(2026,8,10)),
    (150, "SE#4 MT",               "Sala Eléctrica N°4 MT lote 3",
     605581.31, date(2026,7,21), date(2026,8,10)),
]

S3_TOTAL = sum(it[3] for it in S3_ITEMS)
S4_TOTAL = sum(it[3] for it in S4_ITEMS)

# ── WEEKLY TIMELINE ───────────────────────────────────────────────────────────
GANTT_START = date(2026, 6, 1)   # first Monday on or before first event
GANTT_END   = date(2026, 9, 7)   # cover all rev0 dates

def monday_of(d):
    return d - timedelta(days=d.weekday())

def build_weeks(start, end):
    wks, d = [], monday_of(start)
    while d <= end:
        wks.append(d)
        d += timedelta(weeks=1)
    return wks

WEEKS = build_weeks(GANTT_START, GANTT_END)

def to_week_idx(d):
    m = monday_of(d)
    try:
        return WEEKS.index(m)
    except ValueError:
        # find nearest week
        for i, w in enumerate(WEEKS):
            if w > m:
                return i
        return len(WEEKS) - 1

def aggregate_by_week(items):
    revA_wk = defaultdict(float)
    rev0_wk = defaultdict(float)
    for pos, tag, desc, price, revA, rev0 in items:
        revA_wk[monday_of(revA)] += price * 0.10
        rev0_wk[monday_of(rev0)] += price * 0.10
    return revA_wk, rev0_wk

s3_revA_wk, s3_rev0_wk = aggregate_by_week(S3_ITEMS)
s4_revA_wk, s4_rev0_wk = aggregate_by_week(S4_ITEMS)

def week_series(revA_wk, rev0_wk):
    """Returns lists of (revA_inc, rev0_inc, revA_cum, rev0_cum, total_inc, total_cum)"""
    revA_inc = [revA_wk.get(w, 0.0) for w in WEEKS]
    rev0_inc = [rev0_wk.get(w, 0.0) for w in WEEKS]
    revA_cum, rev0_cum, tot_cum = [], [], []
    ca = cr = 0.0
    for a, r in zip(revA_inc, rev0_inc):
        ca += a; cr += r
        revA_cum.append(ca)
        rev0_cum.append(cr)
        tot_cum.append(ca + cr)
    total_inc = [a + r for a, r in zip(revA_inc, rev0_inc)]
    return revA_inc, rev0_inc, revA_cum, rev0_cum, total_inc, tot_cum

s3_series = week_series(s3_revA_wk, s3_rev0_wk)
s4_series = week_series(s4_revA_wk, s4_rev0_wk)

# Combined series
comb_revA_wk = defaultdict(float)
comb_rev0_wk = defaultdict(float)
for w, v in s3_revA_wk.items(): comb_revA_wk[w] += v
for w, v in s4_revA_wk.items(): comb_revA_wk[w] += v
for w, v in s3_rev0_wk.items(): comb_rev0_wk[w] += v
for w, v in s4_rev0_wk.items(): comb_rev0_wk[w] += v
comb_series = week_series(comb_revA_wk, comb_rev0_wk)

# ── WORKBOOK ─────────────────────────────────────────────────────────────────
wb = Workbook()
ws_main = wb.active
ws_main.title = "1_Tabla_Certificacion"

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 1 – TABLA DETALLADA
# ─────────────────────────────────────────────────────────────────────────────
ws = ws_main
ws.sheet_view.showGridLines = False
ws.freeze_panes = "A4"

# Merge title
ws.merge_cells("A1:K1")
c = ws["A1"]
c.value = "CRONOGRAMA DE CERTIFICACIÓN – INGENIERÍA BÁSICA  |  OC S3 + OC S4"
c.font      = fnt(bold=True, color=C["white"], size=14)
c.alignment = aln("center", "center")
c.fill      = fill(C["hdr"])
c.border    = brd("medium")
ws.row_dimensions[1].height = 28

ws.merge_cells("A2:K2")
c = ws["A2"]
c.value = (f"Hito 1: 10% OC al envío RevA  |  Hito 2: 10% OC a la aprobación Rev0  |  "
           f"S3 Total elegible: USD {S3_TOTAL:,.0f}  |  "
           f"S4 Total elegible: USD {S4_TOTAL:,.0f}  |  "
           f"IB 20% combinado: USD {(S3_TOTAL+S4_TOTAL)*0.20:,.0f}")
c.font      = fnt(bold=False, color=C["white"], size=10)
c.alignment = aln("center", "center")
c.fill      = fill(C["gray"])
c.border    = brd()
ws.row_dimensions[2].height = 18

# Column widths
col_widths = [6, 20, 38, 14, 12, 12, 12, 12, 12, 12, 20]
col_names  = ["OC","Pos / TAG","Descripción","Precio (USD)",
              "RevA\n(Envío IB)","Monto RevA\n(10%)","Rev0\n(Aprobación)",
              "Monto Rev0\n(10%)","Total IB\n(20%)","% sobre\nOC total","Notas"]
for i, (w, n) in enumerate(zip(col_widths, col_names), 1):
    ws.column_dimensions[get_column_letter(i)].width = w
    cell(ws, 3, i, n, bold=True, color=C["white"], bg=C["hdr"],
         size=9, h="center", wrap=True)
ws.row_dimensions[3].height = 30

def write_oc_block(ws, items, oc_label, oc_total, start_row, hdr_bg, lt_bg):
    r = start_row
    # OC header
    ws.merge_cells(f"A{r}:K{r}")
    oc_color = C["s3"] if "S3" in oc_label else C["s4"]
    c = ws.cell(row=r, column=1)
    c.value = oc_label
    c.font      = fnt(bold=True, color=C["white"], size=11)
    c.alignment = aln("left", "center")
    c.fill      = fill(oc_color)
    c.border    = brd("medium")
    ws.row_dimensions[r].height = 22
    r += 1

    for pos, tag, desc, price, revA, rev0 in items:
        is_sala = "SE#" in tag
        row_bg  = lt_bg if is_sala else None
        cell(ws, r, 1,  oc_label.split()[0], bg=row_bg, size=9)
        cell(ws, r, 2,  f"Pos {pos}\n{tag}", bg=row_bg, size=8, h="left", wrap=True)
        cell(ws, r, 3,  desc, bg=row_bg, size=9, h="left", wrap=True)
        cell(ws, r, 4,  price, bg=row_bg, size=9, num_format='$#,##0.00')
        cell(ws, r, 5,  revA,  bg=row_bg, size=9, num_format='DD-MMM-YYYY')
        cell(ws, r, 6,  price*0.10, bg=row_bg, size=9, num_format='$#,##0.00')
        cell(ws, r, 7,  rev0,  bg=row_bg, size=9, num_format='DD-MMM-YYYY')
        cell(ws, r, 8,  price*0.10, bg=row_bg, size=9, num_format='$#,##0.00')
        cell(ws, r, 9,  price*0.20, bg=row_bg, size=9, num_format='$#,##0.00')
        cell(ws, r, 10, price*0.20/oc_total*100, bg=row_bg, size=9,
             num_format='0.00"%"')
        # notes: days between revA and rev0
        days = (rev0 - revA).days
        cell(ws, r, 11, f"Gap aprobación: {days}d", bg=row_bg, size=8,
             italic=True, color=C["gray"])
        ws.row_dimensions[r].height = 28
        r += 1

    # Subtotal row
    ws.merge_cells(f"A{r}:C{r}")
    cell(ws, r, 1, f"SUBTOTAL {oc_label.split()[0]}  ({len(items)} ítems)",
         bold=True, bg=hdr_bg, color=C["white"], h="center", size=10)
    cell(ws, r, 4, oc_total, bold=True, bg=hdr_bg, color=C["white"],
         num_format='$#,##0.00')
    cell(ws, r, 5, "→ Hito 1", bold=True, bg=hdr_bg, color=C["white"])
    cell(ws, r, 6, oc_total*0.10, bold=True, bg=hdr_bg, color=C["white"],
         num_format='$#,##0.00')
    cell(ws, r, 7, "→ Hito 2", bold=True, bg=hdr_bg, color=C["white"])
    cell(ws, r, 8, oc_total*0.10, bold=True, bg=hdr_bg, color=C["white"],
         num_format='$#,##0.00')
    cell(ws, r, 9, oc_total*0.20, bold=True, bg=hdr_bg, color=C["white"],
         num_format='$#,##0.00')
    cell(ws, r, 10, 20.0, bold=True, bg=hdr_bg, color=C["white"], num_format='0.00"%"')
    cell(ws, r, 11, "", bg=hdr_bg)
    ws.row_dimensions[r].height = 20
    return r + 2  # blank line after block

row = 4
row = write_oc_block(ws, S3_ITEMS,
                     "OC S3 – 4508944971  |  Sala 3 BT  |  ABB Brazil",
                     S3_TOTAL, row, C["hdr_s3"], C["s3_lt"])
row = write_oc_block(ws, S4_ITEMS,
                     "OC S4 – 4508944973  |  Sala 4 MT  |  ABB Brazil",
                     S4_TOTAL, row, C["hdr_s4"], C["s4_lt"])

# Grand total
ws.merge_cells(f"A{row}:C{row}")
gt_bg = C["comb"]
cell(ws, row, 1, "TOTAL COMBINADO S3 + S4", bold=True, bg=gt_bg,
     color=C["white"], size=11, h="center")
cell(ws, row, 4, S3_TOTAL+S4_TOTAL, bold=True, bg=gt_bg, color=C["white"],
     num_format='$#,##0.00')
cell(ws, row, 5, "", bg=gt_bg)
cell(ws, row, 6, (S3_TOTAL+S4_TOTAL)*0.10, bold=True, bg=gt_bg, color=C["white"],
     num_format='$#,##0.00')
cell(ws, row, 7, "", bg=gt_bg)
cell(ws, row, 8, (S3_TOTAL+S4_TOTAL)*0.10, bold=True, bg=gt_bg, color=C["white"],
     num_format='$#,##0.00')
cell(ws, row, 9, (S3_TOTAL+S4_TOTAL)*0.20, bold=True, bg=gt_bg, color=C["white"],
     num_format='$#,##0.00')
cell(ws, row, 10, 20.0, bold=True, bg=gt_bg, color=C["white"], num_format='0.00"%"')
cell(ws, row, 11, f"JUN-AGO 2026", bold=True, bg=gt_bg, color=C["white"])
ws.row_dimensions[row].height = 22

# ─────────────────────────────────────────────────────────────────────────────
# SHEET 5 – RESUMEN HITOS (SORTED BY DATE)
# ─────────────────────────────────────────────────────────────────────────────
ws5 = wb.create_sheet("5_Resumen_Hitos")
ws5.sheet_view.showGridLines = False

def mk_hitos_list(items, oc_label):
    events = []
    for pos, tag, desc, price, revA, rev0 in items:
        oc = oc_label
        events.append((revA, oc, f"Pos {pos}", tag, desc, "RevA – Envío IB",
                        price*0.10, "Hito 1 (10%)"))
        events.append((rev0, oc, f"Pos {pos}", tag, desc, "Rev0 – Aprobación IB",
                        price*0.10, "Hito 2 (10%)"))
    return events

all_events = (mk_hitos_list(S3_ITEMS, "S3-4508944971") +
              mk_hitos_list(S4_ITEMS, "S4-4508944973"))
all_events.sort(key=lambda x: x[0])

# Build running totals
s3_run = s4_run = tot_run = 0.0
enriched = []
for ev in all_events:
    d, oc, pos, tag, desc, tipo, amt, hito = ev
    if "S3" in oc:
        s3_run += amt
    else:
        s4_run += amt
    tot_run += amt
    enriched.append((d, oc, pos, tag, desc, tipo, amt, hito, s3_run, s4_run, tot_run))

# Header
ws5.merge_cells("A1:K1")
c5 = ws5["A1"]
c5.value = "RESUMEN CRONOLÓGICO DE HITOS – CERTIFICACIÓN INGENIERÍA BÁSICA"
c5.font = fnt(bold=True, color=C["white"], size=13)
c5.alignment = aln("center", "center")
c5.fill = fill(C["hdr"])
c5.border = brd("medium")
ws5.row_dimensions[1].height = 26

hdrs5 = ["Fecha","OC","Posición","TAG","Descripción","Tipo Hito",
         "Monto (USD)","Hito","Acum. S3","Acum. S4","Acum. Total"]
wdths5 = [12,16,8,24,36,22,14,12,14,14,14]
for i, (h, w) in enumerate(zip(hdrs5, wdths5), 1):
    ws5.column_dimensions[get_column_letter(i)].width = w
    cell(ws5, 2, i, h, bold=True, color=C["white"], bg=C["hdr"], size=9,
         h="center", wrap=True)
ws5.row_dimensions[2].height = 24

for r, ev in enumerate(enriched, 3):
    d, oc, pos, tag, desc, tipo, amt, hito, a3, a4, at = ev
    is_s3   = "S3" in oc
    row_bg  = C["s3_lt"] if is_s3 else C["s4_lt"]
    tipo_bg = C["s3"] if "RevA" in tipo else C["rev0"]
    cell(ws5, r, 1,  d,    bg=row_bg, size=9, num_format='DD-MMM-YYYY')
    cell(ws5, r, 2,  oc,   bg=row_bg if "S3" in oc else None,
         bold=True, color=C["s3_dk"] if is_s3 else C["s4_dk"], size=9)
    cell(ws5, r, 3,  pos,  bg=row_bg, size=9)
    cell(ws5, r, 4,  tag,  bg=row_bg, size=8, h="left")
    cell(ws5, r, 5,  desc, bg=row_bg, size=8, h="left")
    cell(ws5, r, 6,  tipo, bold=True,
         color=C["s3_dk"] if "RevA" in tipo else C["rev0"], size=9)
    cell(ws5, r, 7,  amt,  bold=True, size=9, num_format='$#,##0.00')
    cell(ws5, r, 8,  hito, size=9, italic=True)
    cell(ws5, r, 9,  a3,   size=9, num_format='$#,##0.00',
         bg=C["s3_lt"] if is_s3 else None)
    cell(ws5, r, 10, a4,   size=9, num_format='$#,##0.00',
         bg=C["s4_lt"] if not is_s3 else None)
    cell(ws5, r, 11, at,   bold=True, size=9, num_format='$#,##0.00',
         bg=C["comb_lt"])
    ws5.row_dimensions[r].height = 16

# Freeze
ws5.freeze_panes = "A3"

# ─────────────────────────────────────────────────────────────────────────────
# HELPER – build a Curva-S sheet
# ─────────────────────────────────────────────────────────────────────────────
def build_curva_sheet(ws, title, subtitle, series_tuple, oc_total, hdr_color, lt_color):
    """
    series_tuple = (revA_inc, rev0_inc, revA_cum, rev0_cum, total_inc, total_cum)
    """
    revA_inc, rev0_inc, revA_cum, rev0_cum, total_inc, total_cum = series_tuple
    ws.sheet_view.showGridLines = False

    # Title rows
    ws.merge_cells("A1:T1")
    c = ws["A1"]
    c.value = title
    c.font = fnt(bold=True, color=C["white"], size=13)
    c.alignment = aln("center", "center")
    c.fill = fill(hdr_color)
    c.border = brd("medium")
    ws.row_dimensions[1].height = 26

    ws.merge_cells("A2:T2")
    c = ws["A2"]
    c.value = subtitle
    c.font = fnt(bold=False, color=C["white"], size=10)
    c.alignment = aln("center", "center")
    c.fill = fill(C["gray"])
    c.border = brd()
    ws.row_dimensions[2].height = 18

    # ── DATA TABLE (rows 3 onward, weeks as columns) ──────────────────────
    # Col layout: A=label | B=month | C:... = weeks
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 9

    COL0 = 3  # first week column (1-indexed)
    # Header row 3 – MONTHS
    cell(ws, 3, 1, "CONCEPTO", bold=True, color=C["white"], bg=hdr_color, size=9)
    cell(ws, 3, 2, "TOTAL",    bold=True, color=C["white"], bg=hdr_color, size=9)
    prev_month = None
    month_start_col = COL0
    for i, w in enumerate(WEEKS):
        col = COL0 + i
        ws.column_dimensions[get_column_letter(col)].width = 7.5
        if w.month != prev_month:
            if prev_month is not None:
                # merge previous month header
                mc = month_start_col
                ec = col - 1
                if mc < ec:
                    ws.merge_cells(start_row=3, start_column=mc,
                                   end_row=3,   end_column=ec)
            mname = w.strftime("%B %Y").upper()
            cell(ws, 3, col, mname, bold=True, color=C["white"],
                 bg=C["hdr_s3"] if w.month == 6 else
                    (C["hdr_s4"] if w.month == 7 else C["comb"]),
                 size=9)
            prev_month = w.month
            month_start_col = col
    # merge last month
    if month_start_col < COL0 + len(WEEKS) - 1:
        ws.merge_cells(start_row=3, start_column=month_start_col,
                       end_row=3, end_column=COL0+len(WEEKS)-1)

    # Row 4 – week dates
    cell(ws, 4, 1, "SEMANA (LUNES)", bold=True, color=C["white"], bg=hdr_color, size=8)
    cell(ws, 4, 2, "", bg=hdr_color)
    ws.row_dimensions[3].height = 18
    ws.row_dimensions[4].height = 18
    for i, w in enumerate(WEEKS):
        col = COL0 + i
        cell(ws, 4, col, w.strftime("%d-%b"), bold=True, color=C["white"],
             bg=hdr_color, size=8)

    # Row labels and data
    ROWS = [
        (5,  "Hito 1 – Envío RevA (USD)",  revA_inc,  lt_color,   False, '$#,##0'),
        (6,  "Hito 2 – Aprobación Rev0 (USD)", rev0_inc, C["rev0_lt"], False, '$#,##0'),
        (7,  "PAGO SEMANAL TOTAL (USD)",    total_inc, C["yell"],   True,  '$#,##0'),
        (8,  "Acum. Hito 1 (USD)",          revA_cum,  lt_color,   False, '$#,##0'),
        (9,  "Acum. Hito 2 (USD)",          rev0_cum,  C["rev0_lt"],False, '$#,##0'),
        (10, "CURVA S – ACUMULADO TOTAL",   total_cum, C["comb_lt"],True,  '$#,##0'),
        (11, "% Avance IB (sobre 20% OC)",  [v/(oc_total*0.20)*100 if oc_total else 0
                                              for v in total_cum],
             C["llgray"], True, '0.0"%"'),
    ]
    for row_num, label, values, bg, bold_flag, nf in ROWS:
        cell(ws, row_num, 1, label, bold=bold_flag, size=9, h="left",
             bg=bg, color=C["black"])
        row_total = sum(v for v in values if v)
        cell(ws, row_num, 2, row_total if row_num < 8 else values[-1],
             bold=bold_flag, size=9, bg=bg, num_format=nf)
        for i, val in enumerate(values):
            col = COL0 + i
            bg_cell = bg if val else None
            cell(ws, row_num, col, val if val else None,
                 bold=bold_flag, size=9, bg=bg_cell, num_format=nf)
        ws.row_dimensions[row_num].height = 18

    # Separator
    ws.row_dimensions[12].height = 6

    # ── VISUAL BAR INDICATOR (rows 13-32 = 20 pixels tall) ───────────────
    # Each week column: bar height proportional to total_cum / total
    final_cum = total_cum[-1] if total_cum else 1
    cell(ws, 13, 1, "GRÁFICO SIMPLIFICADO – CURVA S ACUMULADA", bold=True,
         color=C["white"], bg=hdr_color, size=9)
    cell(ws, 13, 2, "", bg=hdr_color)
    ws.row_dimensions[13].height = 16

    # Draw simple bar chart using colored cells
    # 15 rows available for bars = rows 14-28
    BAR_ROWS = 15
    for i, (tot, cum) in enumerate(zip(total_inc, total_cum)):
        col = COL0 + i
        pct = cum / final_cum if final_cum else 0
        filled_rows = round(pct * BAR_ROWS)
        cell(ws, 13, col, f"{pct*100:.0f}%", bold=True, size=7,
             bg=hdr_color, color=C["white"])
        for br in range(BAR_ROWS):
            row_r = 28 - br   # draw from bottom up
            if BAR_ROWS - br <= filled_rows:
                bar_color = lt_color
                # darker at top = recent payments
                if tot > 0 and BAR_ROWS - br > filled_rows - 2:
                    bar_color = hdr_color
                c_bar = ws.cell(row=row_r, column=col)
                c_bar.fill = fill(bar_color)
                c_bar.border = brd()
            else:
                c_bar = ws.cell(row=row_r, column=col)
                c_bar.fill = fill(C["llgray"])
                c_bar.border = brd()
            ws.row_dimensions[row_r].height = 7

    # Left labels for bars
    for br in range(BAR_ROWS):
        row_r = 28 - br
        pct_lbl = f"{(br+1)/BAR_ROWS*100:.0f}%"
        cell(ws, row_r, 1, pct_lbl, size=7, italic=True, color=C["gray"],
             h="right", bg=C["llgray"])
        cell(ws, row_r, 2, "", bg=C["llgray"])

    # Legend row
    ws.row_dimensions[29].height = 16
    ws.merge_cells("A29:B29")
    cell(ws, 29, 1, "◀ CURVA S ACUMULADA", bold=True, color=C["white"],
         bg=hdr_color, size=9)
    for i in range(len(WEEKS)):
        col = COL0 + i
        cell(ws, 29, col, WEEKS[i].strftime("%d-%b"), size=7,
             color=C["gray"], bg=C["llgray"])
    ws.row_dimensions[30].height = 6

    # ── OPENPYXL CHART ────────────────────────────────────────────────────
    # Build a combo: bar (incremental) + line (cumulative)
    # Data rows: total_inc=row 7, total_cum=row 10
    # We'll place chart in rows 31+
    data_row_revA  = 5
    data_row_rev0  = 6
    data_row_cum   = 10

    # Create bar chart for incremental
    bar = BarChart()
    bar.type    = "col"
    bar.grouping = "stacked"
    bar.title   = title
    bar.style   = 10
    bar.y_axis.title = "Pago semanal (USD)"
    bar.x_axis.title = "Semana"
    bar.width  = 22
    bar.height = 12

    ref_revA = Reference(ws, min_col=COL0, max_col=COL0+len(WEEKS)-1,
                         min_row=data_row_revA, max_row=data_row_revA)
    ref_rev0 = Reference(ws, min_col=COL0, max_col=COL0+len(WEEKS)-1,
                         min_row=data_row_rev0, max_row=data_row_rev0)
    bar.add_data(ref_revA, from_rows=True, titles_from_data=False)
    bar.add_data(ref_rev0, from_rows=True, titles_from_data=False)
    from openpyxl.chart.series import SeriesLabel
    bar.series[0].title = SeriesLabel(v="Hito 1 – RevA")
    bar.series[1].title = SeriesLabel(v="Hito 2 – Rev0")
    bar.series[0].graphicalProperties.solidFill = hdr_color
    bar.series[1].graphicalProperties.solidFill = "7030A0"

    # Line chart for cumulative
    line = LineChart()
    line.grouping = "standard"
    ref_cum = Reference(ws, min_col=COL0, max_col=COL0+len(WEEKS)-1,
                        min_row=data_row_cum, max_row=data_row_cum)
    line.add_data(ref_cum, from_rows=True, titles_from_data=False)
    line.series[0].title = SeriesLabel(v="Curva S – Acumulado")
    line.series[0].graphicalProperties.line.solidFill = C["orng"]
    line.series[0].graphicalProperties.line.width = 25000
    line.y_axis.axId = 200
    line.y_axis.title = "Acumulado (USD)"
    bar.y_axis.crosses = "max"
    bar += line

    # Categories = week labels
    cats = Reference(ws, min_col=COL0, max_col=COL0+len(WEEKS)-1,
                     min_row=4, max_row=4)
    bar.set_categories(cats)

    ws.add_chart(bar, f"A{32}")

    ws.freeze_panes = "C4"
    return ws

# ─────────────────────────────────────────────────────────────────────────────
# BUILD SHEETS 2, 3, 4
# ─────────────────────────────────────────────────────────────────────────────
ws2 = wb.create_sheet("2_Curva_S_S3")
build_curva_sheet(
    ws2,
    "CURVA S – OC S3 (4508944971)  |  SALA 3 BT  |  USD 8.13M",
    f"Hito 1 – 10% RevA  |  Hito 2 – 10% Rev0  |  IB total OC S3: USD {S3_TOTAL*0.20:,.0f}",
    s3_series, S3_TOTAL, C["hdr_s3"], C["s3_lt"]
)

ws3 = wb.create_sheet("3_Curva_S_S4")
build_curva_sheet(
    ws3,
    "CURVA S – OC S4 (4508944973)  |  SALA 4 MT  |  USD 3.69M",
    f"Hito 1 – 10% RevA  |  Hito 2 – 10% Rev0  |  IB total OC S4: USD {S4_TOTAL*0.20:,.0f}",
    s4_series, S4_TOTAL, C["hdr_s4"], C["s4_lt"]
)

ws4 = wb.create_sheet("4_Curva_S_Combinada")
build_curva_sheet(
    ws4,
    "CURVA S COMBINADA – OC S3 + OC S4  |  USD 11.81M",
    (f"S3: USD {S3_TOTAL*0.20:,.0f} IB  |  S4: USD {S4_TOTAL*0.20:,.0f} IB  |  "
     f"Total IB (20%): USD {(S3_TOTAL+S4_TOTAL)*0.20:,.0f}"),
    comb_series, S3_TOTAL+S4_TOTAL, C["comb"], C["comb_lt"]
)

# ─────────────────────────────────────────────────────────────────────────────
# REORDER SHEETS
# ─────────────────────────────────────────────────────────────────────────────
wb.move_sheet("1_Tabla_Certificacion", offset=0)
for name in ["2_Curva_S_S3","3_Curva_S_S4","4_Curva_S_Combinada","5_Resumen_Hitos"]:
    wb.move_sheet(name, offset=len(wb.sheetnames))

# ─────────────────────────────────────────────────────────────────────────────
# SAVE
# ─────────────────────────────────────────────────────────────────────────────
OUT_PATH = "/home/user/OWN-Daro-1/2_Planning/Certificacion_IB_CurvaS_S3_S4.xlsx"
wb.save(OUT_PATH)
print(f"\n✓ Archivo guardado: {OUT_PATH}")
print(f"\nRESUMEN IB:")
print(f"  OC S3 total elegible : USD {S3_TOTAL:>14,.2f}")
print(f"  OC S4 total elegible : USD {S4_TOTAL:>14,.2f}")
print(f"  IB Hito 1 S3 (RevA)  : USD {S3_TOTAL*0.10:>14,.2f}")
print(f"  IB Hito 2 S3 (Rev0)  : USD {S3_TOTAL*0.10:>14,.2f}")
print(f"  IB Hito 1 S4 (RevA)  : USD {S4_TOTAL*0.10:>14,.2f}")
print(f"  IB Hito 2 S4 (Rev0)  : USD {S4_TOTAL*0.10:>14,.2f}")
print(f"  TOTAL IB COMBINADO   : USD {(S3_TOTAL+S4_TOTAL)*0.20:>14,.2f}")
print(f"\nFechas clave:")
for d_k, label in sorted([
    (date(2026,6,8),  "Primeros envíos RevA (DP/LP S3 y DP-101 S4)"),
    (date(2026,6,12), "RevA UPS/VCC S3 + TGMT/DP S4"),
    (date(2026,6,15), "RevA VFD S3 y S4"),
    (date(2026,6,22), "Rev0 UPS/VCC S3 + DP-101UPS/VCC S4"),
    (date(2026,6,26), "RevA Ductos Barras S3"),
    (date(2026,6,30), "Rev0 VFD S3 y S4"),
    (date(2026,7,6),  "Rev0 DP-101 S4"),
    (date(2026,7,10), "Rev0 DP/LP S3"),
    (date(2026,7,16), "Rev0 Ductos Barras S3"),
    (date(2026,7,21), "RevA Sala 4 MT"),
    (date(2026,7,29), "RevA Sala 3 BT"),
    (date(2026,7,31), "Rev0 TGMT S4"),
    (date(2026,8,10), "RevA CCMs S3 + Rev0 Sala 4 MT"),
    (date(2026,8,17), "Rev0 Sala 3 BT"),
    (date(2026,8,25), "Rev0 CCMs S3  ← ÚLTIMA CERTIFICACIÓN"),
]):
    print(f"  {d_k}  {label}")
