#!/usr/bin/env python3
"""
Rev 3 del Cronograma General ABB.
Fuentes:
  - Cronograma actualizado: 5155-00-0009-VE-CO-PG-001_A - (13-7).pdf  (15-JUL-2026)
  - Informe de Avance Nro 2: R-2631036 - Informe de Avance nro 2_A.pdf (13-JUL-2026)
Genera: CRONOGRAMA_GENERAL_ABB_Rev3_2026-07-21.xlsx
"""

import shutil
from datetime import date
from openpyxl import load_workbook
from openpyxl.styles import (Font, PatternFill, Alignment, Border, Side,
                              numbers)
from openpyxl.utils import get_column_letter

SRC = "CRONOGRAMA_GENERAL_ABB_Actualizado_2026-06-09.xlsx"
DST = "CRONOGRAMA_GENERAL_ABB_Rev3_2026-07-21.xlsx"
REV   = "Rev 3"
FECHA = "21-JUL-2026"

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
C_CRITICAL= "C00000"

# ── Step 1: copy base file ────────────────────────────────────────────────────
shutil.copy(SRC, DST)
wb = load_workbook(DST)

# ── Step 2: update revision header in every existing sheet ───────────────────
OLD_STRINGS = ["Actualizado_2026-06-09", "09-JUN-2026", "2026-06-09",
               "Rev 2", "REV 2", "Fecha: 27-MAY-2026 → ACTUALIZADO 09-JUN-2026"]
NEW_STRINGS = [f"Rev3_{FECHA}",         "21-JUL-2026", "2026-07-21",
               REV,     REV,     f"ACTUALIZADO {FECHA}"]

for ws in wb.worksheets:
    for row in ws.iter_rows():
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                v = cell.value
                for old, new in zip(OLD_STRINGS, NEW_STRINGS):
                    v = v.replace(old, new)
                # generic "09-JUN" → "21-JUL" for title blocks
                v = v.replace("Fecha: 27-MAY-2026 → ACTUALIZADO 09-JUN-2026",
                               f"Actualizado {FECHA}")
                cell.value = v

# Also update the title cell of sheet 1 specifically
ws1 = wb["1_Gantt_General"]
for row in ws1.iter_rows(max_row=3):
    for cell in row:
        if cell.value and "CPF-2" in str(cell.value):
            cell.value = cell.value.replace(
                "Fecha: 27-MAY-2026 → ACTUALIZADO 09-JUN-2026",
                f"Fecha: 27-MAY-2026 → {FECHA}")

# ── Step 3: update Freezing Points status ────────────────────────────────────
ws_fp = wb["4_Freezing_Points"]
# FPs by equipo — update status based on IA#2 (June 2026)
fp_updates = {
    "FP-S3/Tab.Aux-IB": ("COMPLETADO ✓", "IB emitida y aprobada con observaciones (Informe N°2 Jun-2026)"),
    "FP-S4/Tab.Aux-IB": ("COMPLETADO ✓", "IB emitida y aprobada con observaciones (Informe N°2)"),
    "FP-S3/CCM-IB":     ("COMPLETADO ✓", "IB CCMs emitida y aprobada con obs. Cambios cargas CCM-006/007 → nota cambio Jul"),
    "FP-S4/Celdas-IB":  ("COMPLETADO ✓", "IB Tableros MT emitida y aprobada con observaciones"),
    "FP-S3/UPS-IB":     ("COMPLETADO ✓", "IB UPS/Baterías emitida y aprobada con observaciones"),
}
for row in ws_fp.iter_rows():
    for cell in row:
        v = str(cell.value) if cell.value else ""
        for fp_id, (new_status, obs) in fp_updates.items():
            if fp_id in v:
                # Update status cell (usually 2 columns to the right)
                try:
                    sr = cell.row
                    for c in ws_fp[sr]:
                        if str(c.value) == "PENDIENTE":
                            c.value = new_status
                            c.fill = fill(C_GREEN)
                            c.font = font(bold=True, color="375623")
                except:
                    pass

# ── Step 4: Update sheet 6 title to reflect N°2 ──────────────────────────────
ws6 = wb["6_Estado_JUN2026"]
ws6.title = "6_Estado_JUN2026_IA1"  # preserve as historical

# ── Step 5: Add new sheet — Estado IA#2 (Julio 2026) ─────────────────────────
if "7_Estado_JUL2026_IA2" in wb.sheetnames:
    del wb["7_Estado_JUL2026_IA2"]
ws7 = wb.create_sheet("7_Estado_JUL2026_IA2")

# Column widths
for col, w in [("A",6),("B",30),("C",16),("D",8),("E",12),("F",12),("G",16)]:
    ws7.column_dimensions[col].width = w

_r = [1]

def row():  return _r[0]
def nrow(): _r[0]+=1; return _r[0]-1

def write_title(text, bg=C_TITLE, fgc="FFFFFF", span=7, height=24):
    r = _r[0]
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws7.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, color=fgc, size=11)
    c.alignment = center(); c.border = border_all()
    ws7.row_dimensions[r].height = height
    _r[0]+=1

def write_section(text, bg=C_HEAD):
    r = _r[0]
    ws7.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    c = ws7.cell(row=r, column=1, value=text)
    c.fill = fill(bg); c.font = font(bold=True, size=10)
    c.alignment = left(); c.border = border_thick()
    ws7.row_dimensions[r].height = 18
    _r[0]+=1

def write_hdr(cols, bg=C_HEAD):
    r = _r[0]
    for ci, (val, w) in enumerate(cols, 1):
        c = ws7.cell(row=r, column=ci, value=val)
        c.fill = fill(bg); c.font = font(bold=True, size=9)
        c.alignment = center(wrap=True); c.border = border_all()
    ws7.row_dimensions[r].height = 30
    _r[0]+=1

def write_data(values, bg_col=None, bold=False, aligns=None):
    r = _r[0]
    bg_use = bg_col if bg_col else ("FFFFFF" if r%2==0 else C_EVEN)
    for ci, val in enumerate(values, 1):
        c = ws7.cell(row=r, column=ci, value=val)
        c.fill = fill(bg_use); c.font = font(bold=bold, size=9)
        if aligns and ci <= len(aligns):
            c.alignment = aligns[ci-1]
        else:
            c.alignment = left(wrap=True) if ci in (2,5,7) else center()
        c.border = border_all()
    _r[0]+=1

def blank():
    _r[0]+=1

# ── HEADER ───────────────────────────────────────────────────────────────────
write_title(f"ESTADO DEL PROYECTO — INFORME DE AVANCE N°2  |  ABB  |  {FECHA}  |  {REV}", height=26)
r=_r[0]
ws7.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7)
c=ws7.cell(row=r,column=1,
    value="Fuente: R-2631036 Informe de Avance N°2 (13-JUL-2026) + Cronograma 5155-00-0009-VE-CO-PG-001_A (15-JUL-2026)")
c.fill=fill(C_HEAD); c.font=font(size=8); c.alignment=center(); c.border=border_all()
_r[0]+=1
blank()

# ── A. RESUMEN EJECUTIVO ──────────────────────────────────────────────────────
write_section("A. RESUMEN DE ESTADO — AVANCE JUNIO 2026 (Informe N°2, 13-JUL-2026)")
write_hdr([("Área",1),("OC / Ítem",2),("Estado",3),("✓",4),("Avance Global",5),
           ("Previsto",6),("Observación",7)])

status_data = [
    ("Ingeniería","CCMs 102-CCM-005/6/7","COMPLETADO","✓","20% global","20%",
     "IB emitida y aprobada con obs. Cambios cargas CCM-006/007 → nota cambio Jul"),
    ("Ingeniería","Tableros MT 102-TGMT-001/002","COMPLETADO","✓","30% global","30%",
     "IB emitida y aprobada con observaciones. Materiales críticos definidos"),
    ("Ingeniería","Tableros Aux SE#3 y SE#4","COMPLETADO","✓","30% global","30%",
     "IB emitida y aprobada con obs. Lista materiales (mecánica + interruptores) definida"),
    ("Ingeniería","UPS + Sistema CC","COMPLETADO","✓","23% global","23%",
     "IB emitida y aprobada con obs. Materiales críticos definidos"),
    ("Ingeniería","Sala #3 (global BT)","EN CURSO","→","3% global","3%",
     "Ingeniería 10% completada. Sub-OCs emitidas para reserva slots fabricación"),
    ("Ingeniería","Sala #4 (global MT)","EN CURSO","→","3% global","3%",
     "Ingeniería 10% completada. Sub-OCs emitidas para reserva slots fabricación"),
    ("Ingeniería","Ductos de Barras (DB-005/6/7)","⚠ SIN AVANCE","—","0%","0%",
     "⚠ ALERTA: 0% avance. Pendiente definición acometida trafos MT de AESA (vence 20/JUL)"),
    ("Ingeniería","Variadores BT 102-DP-VFD-24210","EN CURSO","→","0%","0%",
     "Sin avance Jun. IB esperada julio. OC sub-proveedor China emitida 26/5/2026"),
    ("Ingeniería","Variadores MT 102-DP-VFD-23310","EN CURSO","→","0%","0%",
     "Sin avance Jun. IB esperada julio. OC sub-proveedor China emitida 22/5/2026"),
    ("Procura","Sub-OCs a 3ros (salas+tableros+VFDs)","COMPLETADO","✓","—","—",
     "8 OCs emitidas a subproveedores Argentina/Brasil/China/Turquía entre 16-26 May 2026"),
    ("Fabricación","CCMs + Tableros + Celdas MT","NO INICIADA","—","—","—",
     "Sin actividades de fabricación. Inicio previsto Sep-Oct 2026"),
    ("Calidad","Planes de Inspección","EN ELABORACIÓN","→","—","—",
     "Sin actividades de calidad/inspección en el período"),
    ("Hitos Pago","Certificación N°1 OC S3","PRESENTADA","✓","USD 143,540","10%",
     "Presentada ~4/Jul. Fecha pago ~7/Ago. Aprobación IB como base"),
    ("Hitos Pago","Certificación N°1 OC S4","PRESENTADA","✓","USD 170,739","10%",
     "Presentada ~4/Jul. Fecha pago ~7/Ago. Aprobación IB como base"),
    ("Cronograma","Actualización cronograma ABB","COMPLETADO","✓","—","—",
     "5155-00-0009-VE-CO-PG-001_A (13-7).pdf actualizado 15-JUL-2026 por PM Rodrigo Mack"),
    ("Cronograma","Informe de Avance N°2","COMPLETADO","✓","—","—",
     "R-2631036 Informe Avance N°2 emitido 13-JUL-2026. Firmado R.Mack / aprobado M.Puga"),
]

status_colors = {
    "COMPLETADO": C_GREEN,
    "⚠ SIN AVANCE": C_RED,
    "EN CURSO": C_YELLOW,
    "PRESENTADA": C_GREEN,
    "NO INICIADA": "F2F2F2",
    "EN ELABORACIÓN": C_YELLOW,
}

for row_data in status_data:
    r = _r[0]
    bg_base = "FFFFFF" if r%2==0 else C_EVEN
    for ci, val in enumerate(row_data, 1):
        c = ws7.cell(row=r, column=ci, value=val)
        if ci == 3:
            bg = status_colors.get(str(val).split(" ")[0], bg_base)
            c.fill = fill(bg)
            c.font = font(bold=True, size=9)
        else:
            c.fill = fill(bg_base)
            c.font = font(size=9)
        c.alignment = left(wrap=True) if ci in (2,6,7) else center()
        c.border = border_all()
    _r[0]+=1

blank()

# ── B. AVANCES CUANTITATIVOS POR EQUIPO ──────────────────────────────────────
write_section("B. AVANCE CUANTITATIVO POR EQUIPO — Junio 2026 vs Julio 2026 (previsto)")
write_hdr([("Equipo/Ítem",2),("TAG",3),("Abr-May'26",4),
           ("Jun'26 Real",5),("Jun'26 Plan",6),("Jul'26 Prev",7)], bg="D0E0FF")

avance_data = [
    ("CCMs BT","102-CCM-005/6/7","0%","20%","20%","34%"),
    ("Tableros MT (TGMT)","102-TGMT-001/002","0%","30%","30%","30%"),
    ("Tableros Auxiliares","102-DP/LP-001/002+TRA","0%","30%","30%","30%"),
    ("UPS + Sistema CC","102-UPS-001 / VCC-001","0%","23%","23%","22%+"),
    ("Variadores BT (VFD)","102-DP-VFD-24210A/B/C","0%","0%","0%","10%"),
    ("Variadores MT (VFD)","102-DP-VFD-23310A/B/C","0%","0%","0%","10%"),
    ("Ductos de Barras","102-DB-005/006/007 A/B","0%","0%","0%","15% ⚠"),
    ("Sala Eléctrica #3","Global SE#3 (BT)","0%","3%","3%","17%"),
    ("Sala Eléctrica #4","Global SE#4 (MT)","0%","3%","3%","17%"),
]
col_colors = {4: C_RED, 5: C_YELLOW, 6: C_GREEN}
for eq, tag, m0, m_real, m_plan, m_prox in avance_data:
    r = _r[0]
    bg = "FFFFFF" if r%2==0 else C_EVEN
    # Check if "⚠" in m_prox → red
    if "⚠" in m_prox:
        bg_prox = C_ORANGE
    elif m_real == "0%" and m_plan == "0%":
        bg_prox = C_YELLOW
    else:
        bg_prox = C_GREEN

    vals = [eq, tag, m0, m_real, m_plan, m_prox]
    for ci, val in enumerate(vals, 2):  # start at col 2 (skip col A)
        c = ws7.cell(row=r, column=ci, value=val)
        if ci == 5 and m_real != m_plan:  # deviation
            c.fill = fill(C_ORANGE)
        elif ci == 7 and "⚠" in str(val):
            c.fill = fill(C_ORANGE)
        else:
            c.fill = fill(bg)
        c.font = font(size=9)
        c.alignment = center() if ci >= 4 else left(wrap=True)
        c.border = border_all()
    _r[0]+=1

blank()

# ── C. SUB-OCs EMITIDAS ───────────────────────────────────────────────────────
write_section("C. SUB-OCs EMITIDAS A TERCEROS (Sub-Proveedores)")
# Adjusted column widths for this section
write_hdr([("N° OC Sub",2),("Descripción",3),("Emisión OC",4),
           ("Origen/País",5),("Entrega Prevista",6),("Estado",7)])

sub_oc_data = [
    ("4500407392","Sala eléctrica modular N°3 y N°4 (estructura + instalaciones aux)",
     "18/5/2026","Argentina","14/5/27 (S3)\n8/4/27 (S4)","OC EMITIDA ✓"),
    ("4500407396","Tableros auxiliares distribución (locales Argentina)",
     "18/5/2026","Argentina","30/12/2026","OC EMITIDA ✓"),
    ("4500407468","UPS / Sistema cargador-baterías (local Argentina)",
     "19/5/2026","Argentina","27/11/2026","OC EMITIDA ✓"),
    ("4700201354","Ductos de Barras (Turquía — lead ~6 meses)",
     "16/5/2026","Turquía","En definición","⚠ RESERVA SLOT — Diseño pendiente AESA"),
    ("4700201355","Tableros de Media Tensión Unigear (Brasil)",
     "16/5/2026","Brasil","27/11/2026","OC EMITIDA ✓"),
    ("4700201478","Tableros BT / VFD BT (Brasil)",
     "19/5/2026","Brasil","25/1/2027","OC EMITIDA ✓"),
    ("4700201819","Variadores de Baja Tensión (China)",
     "26/5/2026","China","6/11/2026","OC EMITIDA ✓"),
    ("4700201626","Variadores de Media Tensión (China)",
     "22/5/2026","China","6/11/2026","OC EMITIDA ✓"),
]

for n_oc, desc, emision, pais, entrega, estado in sub_oc_data:
    r = _r[0]
    bg = "FFFFFF" if r%2==0 else C_EVEN
    if "⚠" in estado:
        bg = C_ORANGE
    for ci, val in enumerate([n_oc, desc, emision, pais, entrega, estado], 2):
        c = ws7.cell(row=r, column=ci, value=val)
        c.fill = fill(bg)
        c.font = font(size=9, bold=("⚠" in str(val) or "✓" in str(val)))
        c.alignment = left(wrap=True) if ci in (3, 6) else center()
        c.border = border_all()
    _r[0]+=1

blank()

# ── D. CERTIFICACIONES / FACTURACIÓN ─────────────────────────────────────────
write_section("D. ESTADO DE CERTIFICACIONES Y FACTURACIÓN")
write_hdr([("OC",2),("Concepto",3),("Hito %",4),("Monto USD",5),
           ("Presentación",6),("Pago Previsto",7)])

cert_data = [
    ("OC S3 — 4508944971","Certificación N°1 — Envío IB S3","10%","USD 143,540","~4/Jul/2026","~7/Ago/2026"),
    ("OC S4 — 4508944973","Certificación N°1 — Envío IB S4","10%","USD 170,739","~4/Jul/2026","~7/Ago/2026"),
    ("OC S3 — 4508944971","Certificación N°2 — Aprobación IB S3","10%","Aprox. USD 143,540","Previsto Jul-2026","Ago-2026"),
    ("OC S4 — 4508944973","Certificación N°2 — Aprobación IB S4","10%","Aprox. USD 170,739","Previsto Jul-2026","Ago-2026"),
    ("OC S3 — 4508944971","Certificación N°3 — Acopio Mat. S3","25%","USD ~802,000","Prev. Ene-2027","Feb-2027"),
    ("OC S4 — 4508944973","Certificación N°3 — Acopio Mat. S4","25%","USD ~320,000","Prev. Ene-2027","Feb-2027"),
]
for row_d in cert_data:
    r = _r[0]
    bg = C_GREEN if "Certificación N°1" in row_d[1] else (C_YELLOW if "N°2" in row_d[1] else C_EVEN)
    for ci, val in enumerate(row_d, 2):
        c = ws7.cell(row=r, column=ci, value=val)
        c.fill = fill(bg); c.font = font(size=9)
        c.alignment = left(wrap=True) if ci == 3 else center()
        c.border = border_all()
    _r[0]+=1

blank()

# ── E. PENDIENTES CRÍTICOS AESA (del IA#2) ───────────────────────────────────
write_section("E. PENDIENTES CRÍTICOS — INFORMACIÓN REQUERIDA A AESA (Informe N°2)", bg=C_ORANGE)
write_hdr([("Fecha Límite",2),("Descripción Requerimiento",3),("Resp.",4),
           ("Estado (Jul-2026)",5),("Impacto en Cronograma",6),("Ref.",7)])

pend_data = [
    ("20-JUL-2026",
     "Detalle de acometida de transformadores MT y posiciones relativas para ducto de barras",
     "AESA","⚠ PENDIENTE",
     "CRÍTICO: Bloquea inicio OC Ductos Barras Turquía (lead 6m). FP-DB 26/OCT en riesgo → impacta ENTREGA S#3 14-MAY-2027",
     "DB-005/6/7 planos"),
    ("20-JUL-2026",
     "Control de acceso — Definición equipos (gabinetes). Suministro equipos físicos per cronograma cada sala",
     "AESA","⚠ PENDIENTE",
     "Cierre planos constructivos sala → FP-S3-3/S4-3. Retraso afecta Ingeniería de Detalle",
     "Planos instalación sala"),
    ("VENCIDA (Jun-2026)",
     "Típicos funcionales para cada salida a motor (CCM). Ya resuelto parcialmente con IBs aprobadas",
     "AESA","RESUELTO ✓",
     "Resuelto: IB CCMs aprobada Jun-2026. Cambios cargas CCM-006/007 → nota cambio pendiente Jul",
     "Esquemas CCM"),
    ("VENCIDA (Jun-2026)",
     "Detalles acometida trafo MT y posiciones — original fecha límite 18/6. EXTENDIDA A 20/JUL",
     "AESA","⚠ EXTENDIDA",
     "Ver fila 1 — sigue bloqueando ductos barras",
     "DB planos instalación"),
]
for pd in pend_data:
    r = _r[0]
    if "CRÍTICO" in pd[4] or "⚠ PENDIENTE" in pd[3]:
        bg = C_RED
    elif "RESUELTO" in pd[3]:
        bg = C_GREEN
    elif "EXTENDIDA" in pd[3]:
        bg = C_ORANGE
    else:
        bg = C_YELLOW
    for ci, val in enumerate(pd, 2):
        c = ws7.cell(row=r, column=ci, value=val)
        c.fill = fill(bg); c.font = font(size=9, bold=("CRÍTICO" in str(val) or "⚠" in str(val)))
        c.alignment = left(wrap=True)
        c.border = border_all()
    _r[0]+=1

blank()

# ── F. EVENTOS NEXT MONTH (Jul-2026) ─────────────────────────────────────────
write_section("F. EVENTOS PLANIFICADOS — JULIO 2026 (según Informe N°2)")
write_hdr([("Área",2),("Evento / Entregable",3),("Equipos",4),
           ("Estado Previsto",5),("Impacto",6),("Nota",7)])

events_jul = [
    ("Ingeniería","Emisión IB Rev A — VFDs BT","102-DP-VFD-24210A/B/C",
     "PLANIFICADO","Gatilla revisión AESA para OC definitiva Finlandia","Primer avance VFDs BT"),
    ("Ingeniería","Emisión IB Rev A — VFDs MT","102-DP-VFD-23310A/B/C",
     "PLANIFICADO","Gatilla revisión AESA para OC definitiva China","Primer avance VFDs MT"),
    ("Ingeniería","Emisión IB Rev 0 (construcción) — Tableros Auxiliares",
     "102-DP-001/LP-001/002/DP-001-TRA","PLANIFICADO","Inicia compra materiales BOM final","Rev 0 incorpora obs. IB"),
    ("Ingeniería","Emisión IB Rev 0 (construcción) — CCMs",
     "102-CCM-005/6/7","PLANIFICADO","Inicia compra materiales BOM final + nota cambio 006/007","Cambios cargas → incluir en Rev 0"),
    ("Ingeniería","Emisión IB Rev 0 (construcción) — Tableros MT",
     "102-TGMT-001/002","PLANIFICADO","Confirma materiales críticos ya pedidos","Rev 0 incorpora obs. IB"),
    ("Ingeniería","Emisión IB Rev 0 — UPS + Sistema CC",
     "102-UPS-001/VCC-001","PLANIFICADO","Inicia fabricación componentes locales","Rev 0 incorpora obs. IB"),
    ("Ingeniería","Emisión ID — CCMs, Tableros MT, Tableros Aux, VFDs",
     "Todos","PLANIFICADO","Detalle constructivo. Gatilla inicio fabricación Sep-Oct","ID = primer plano para taller"),
    ("Certificación","Certificación Parcial Hito N°2 — Aprobación Ingeniería",
     "OC S3 + S4","PLANIFICADO","Segundo pago 10% por OC (~USD 143K + 171K)","Depende aprobación AESA Jul"),
    ("AESA","⚠ Recepción info acometida trafos MT","DB-005/6/7",
     "CRÍTICO PENDIENTE","Si no llega el 20/Jul → OC Turquía en riesgo → FP-DB oct. riesgo","Fecha límite: 20/JUL/2026"),
    ("Scope","Nota de Cambio — CCM-006 y CCM-007 (cambio cargas)",
     "102-CCM-006/007","PLANIFICADO","Potencial impacto en IB Rev 0 y materiales","Emitir nota cambio en julio"),
]
for ev in events_jul:
    r = _r[0]
    bg = C_RED if "CRÍTICO" in ev[3] else (C_YELLOW if "PLANIFICADO" in ev[3] else C_EVEN)
    for ci, val in enumerate(ev, 2):
        c = ws7.cell(row=r, column=ci, value=val)
        c.fill = fill(bg); c.font = font(size=9, bold=("⚠" in str(val) or "CRÍTICO" in str(val)))
        c.alignment = left(wrap=True)
        c.border = border_all()
    _r[0]+=1

ws7.freeze_panes = "B4"

# ── Step 6: Add deviation analysis sheet ─────────────────────────────────────
if "8_Desvios_Alertas" in wb.sheetnames:
    del wb["8_Desvios_Alertas"]
ws8 = wb.create_sheet("8_Desvios_Alertas")

for col, w in [("A",4),("B",9),("C",32),("D",14),("E",14),("F",14),("G",12),("H",36)]:
    ws8.column_dimensions[col].width = w

_r[0] = 1

def w8_title(text, bg=C_TITLE, fgc="FFFFFF", span=8, h=24):
    r = _r[0]
    ws8.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws8.cell(row=r, column=1, value=text)
    c.fill=fill(bg); c.font=font(bold=True, color=fgc, size=11)
    c.alignment=center(); c.border=border_all()
    ws8.row_dimensions[r].height=h
    _r[0]+=1

def w8_sec(text, bg=C_HEAD, span=8):
    r = _r[0]
    ws8.merge_cells(start_row=r, start_column=1, end_row=r, end_column=span)
    c = ws8.cell(row=r, column=1, value=text)
    c.fill=fill(bg); c.font=font(bold=True)
    c.alignment=left(); c.border=border_thick()
    ws8.row_dimensions[r].height=18
    _r[0]+=1

def w8_hdr(labels, bg=C_HEAD):
    r = _r[0]
    for ci, lab in enumerate(labels, 1):
        c = ws8.cell(row=r, column=ci, value=lab)
        c.fill=fill(bg); c.font=font(bold=True, size=9)
        c.alignment=center(wrap=True); c.border=border_all()
    ws8.row_dimensions[r].height=32
    _r[0]+=1

def w8_row(vals, severity="INFO"):
    r = _r[0]
    bg_map = {
        "CRÍTICO": C_RED,
        "ADVERTENCIA": C_ORANGE,
        "INFO": C_YELLOW,
        "OK": C_GREEN,
    }
    bg = bg_map.get(severity, C_EVEN)
    for ci, val in enumerate(vals, 1):
        c = ws8.cell(row=r, column=ci, value=val)
        c.fill=fill(bg); c.font=font(size=9, bold=(severity=="CRÍTICO"))
        c.alignment = left(wrap=True) if ci >= 3 else center()
        c.border=border_all()
    _r[0]+=1

w8_title(f"ANÁLISIS DE DESVÍOS Y ALERTAS TEMPRANAS — {REV}  |  {FECHA}", h=26)
r=_r[0]
ws8.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8)
c=ws8.cell(row=r,column=1,value=(
    "Comparativo: Línea Base PG-001 Rev A (26-MAY-2026) vs Estado Informe N°2 (13-JUL-2026)  |  "
    "Cronograma actualizado: 15-JUL-2026  |  RFSU objetivo: 31-ENE-2028"))
c.fill=fill(C_HEAD); c.font=font(size=8); c.alignment=center(); c.border=border_all()
_r[0]+=1
_r[0]+=1

# ── SEMÁFORO GLOBAL ──────────────────────────────────────────────────────────
w8_sec("SEMÁFORO GLOBAL — ESTADO PROYECTO", bg=C_HEAD)
sem_data = [
    ("🔴","ROJO","Ductos de Barras (DB-005/6/7)","OC Turquía bloqueada","AESA no entregó acometida trafos"),
    ("🟡","AMARILLO","Variadores BT y MT","0% en junio; inicio en julio","OC emitidas pero ING en 0%"),
    ("🟡","AMARILLO","CCMs - Cambio de cargas 006/007","Nota de cambio pendiente","Posible retraso IB Rev 0"),
    ("🟡","AMARILLO","Sala Eléctrica 3 y 4 (building)","Solo 3% global","Dimensiones/layout aún en definición"),
    ("🟢","VERDE","CCMs 005/6/7 Ingeniería","IB aprobada en junio","On plan: 20% global"),
    ("🟢","VERDE","Tableros MT Unigear","IB aprobada en junio","On plan: 30% global"),
    ("🟢","VERDE","Tableros Auxiliares","IB aprobada en junio","On plan: 30% global"),
    ("🟢","VERDE","Certificación Hito N°1","Presentada ~4-Jul","USD 143,540 + 170,739"),
    ("🟢","VERDE","Sub-OCs a 3ros","8 OCs emitidas May-Jun","Argentina/Brasil/China/Turquía"),
]
for sem, nivel, item, estado, obs in sem_data:
    r = _r[0]
    bg = C_RED if nivel == "ROJO" else (C_YELLOW if nivel == "AMARILLO" else C_GREEN)
    for ci, val in enumerate([sem, nivel, item, estado, obs], 1):
        c = ws8.cell(row=r, column=ci, value=val)
        c.fill=fill(bg); c.font=font(size=10, bold=(nivel=="ROJO"))
        c.alignment=center() if ci<=2 else left(wrap=True)
        c.border=border_all()
    _r[0]+=1

_r[0]+=1

# ── DESVÍOS CRONOGRAMA vs LÍNEA BASE ────────────────────────────────────────
w8_sec("DESVÍOS DE CRONOGRAMA vs LÍNEA BASE PG-001 Rev A", bg="FCE4D6")
w8_hdr(["#","ID Hito","Descripción Hito","Fecha Línea Base","Fecha Actualizada",
        "Δ Días","Severidad","Impacto / Acción Requerida"])

desvios = [
    (1,"FP-DB","Freezing Point Ductos Barras (ÚLTIMO FP REAL PROYECTO)",
     "26/10/2026","⚠ EN RIESGO","+0d (si AESA entrega 20/Jul)",
     "CRÍTICO","OC Turquía requiere acometidas trafos (AESA). Si no llega 20/Jul → FP-DB se corre → impacto directo S#3 14-MAY-2027"),
    (2,"FP-S3/CCM-IB","Freezing Point IB CCMs (todos)",
     "06/07/2026","COMPLETADO ✓","~0d (on time)",
     "OK","Completado en junio. Base para OC Ductos Barras Turquía. Sin desvío en este item"),
    (3,"FP-S3/Tab.Aux-IB","Freezing Point IB Tableros Aux S#3",
     "24/06/2026","COMPLETADO ✓","~+5d (aprobado con obs.)",
     "INFO","Aprobado con observaciones → Rev 0 en julio. Pequeño desvío absorbible"),
    (4,"FP-S4/Celdas-IB","Freezing Point IB Celdas MT",
     "23/06/2026","COMPLETADO ✓","~0d",
     "OK","On time. Materiales críticos MT pedidos"),
    (5,"S3-02","OC Ductos Barras Turquía (6 meses lead)",
     "06/07/2026","BLOQUEADA AESA","Indefinido hasta 20/Jul",
     "CRÍTICO","Bloqueada por pendiente AESA. Lead 6m = necesita OC antes de 26/Oct para que llegue a tiempo (FP-DB 26-Oct → Fabricación Nov-Dic → Argentina Ene-2027)"),
    (6,"S3-02","OC VFDs BT Finlandia",
     "16/07/2026 (est.)","EN PROCESO","Dentro de plazo si FP-S3/Dr-IB Jul",
     "ADVERTENCIA","OC sub-prov. ya emitida (4700201819). Necesita FP-IB julio. Lead 4m → entrega Nov 2026 vs acopio S3 Ene 2027 → OK si no se retrasa"),
    (7,"PAY-S01","Pago Hito 1 S3+S4 (10% Envío IB)",
     "12/06/2026","Presentado ~4/Jul/2026","+22d vs plan",
     "ADVERTENCIA","Retraso en presentación vs plan (12-Jun) pero dentro de tolerancia. IB fue aprobada en junio con obs. Pago esperado Ago 2026"),
    (8,"PP-IB (PMS)","Emisión IB PMS (Arquitectura+I/O+Variables)",
     "08/06/2026→17/07/2026","En curso","—",
     "INFO","No hay datos de avance PMS en IA#2. Verificar con PM en próxima reunión"),
    (9,"PP-MQ","FAT Maqueta PMS (AESA OBLIGATORIO)",
     "22/09/2026→25/09/2026","Sin cambio aparente","—",
     "INFO","Fecha sensible: si PMS se retrasa en IB, FAT Maqueta puede impactar"),
    (10,"Scope","Cambios Cargas CCM-006/007",
     "N/A (nuevo)","Nota Cambio — Jul 2026","Nuevo ítem",
     "ADVERTENCIA","Cambio de cargas detectado → nota de cambio en julio. Puede impactar IB Rev 0 CCMs y generar retraso en compra materiales CCMs"),
]

for d in desvios:
    sev = d[6]
    w8_row(list(d), severity=sev)

_r[0]+=1

# ── CADENA CRÍTICA ACTUALIZADA ────────────────────────────────────────────────
w8_sec("CADENA CRÍTICA S#3 — ESTADO ACTUALIZADO", bg="E6E0FF")
w8_hdr(["Paso","Hito/Actividad","Fecha Plan","Estado Jul-2026",
        "Semáforo","Slack Plan","Slack Real","Nota Crítica"])

cadena = [
    ("1","IB CCMs (FP-S3/CCM-IB)","06-Jul-2026","COMPLETADO","🟢","0d","0d","On time. Base crítica ✓"),
    ("2","OC Ductos Barras Turquía (licitación)","06-Jul-2026","⚠ BLOQUEADA","🔴","0d","INDEF.",
     "Bloqueada por acometidas trafos AESA. VENCE 20/Jul. Acción: FOLLOW-UP URGENTE AESA"),
    ("3","ID CCM-005 aprobada (FP-S3/CCM005-ID)","12-Ago-2026","PENDIENTE","🟡","0d","—",
     "Depende de IB Rev 0 en julio (con nota cambio 006/007)"),
    ("4","Último FP Sala #3 IC (FP-S3/DB-FP)","26-Oct-2026","⚠ EN RIESGO","🔴","0d","—",
     "FP REAL = Ductos Barras. Si OC Turquía no sale antes de 26-Oct → NO HAY MARGEN"),
    ("5","Fabricación 37 Tableros BT (Brasil)","14-Sep-2026","NO INICIADA","🟡","+0d","—",
     "Inicio previsto sep-2026. Depende ID CCMs. Sin novedad negativa aún"),
    ("6","FAT Tableros BT + CCMs (Brasil)","03-Nov-2026","NO INICIADA","⚪","0d","—","Fecha plan sin cambio"),
    ("7","Acopio materiales Argentina S#3","07-Dic-2026","NO INICIADA","⚪","0d","—",
     "Depende de llegada Ductos Turquía. CRÍTICO: si Turquía se retrasa, acopio se atrasa"),
    ("8","Montaje Shelter S#3 (Bottino, Mendoza)","01-Feb-2027","NO INICIADA","⚪","0d","—","Sin cambio"),
    ("9","FAT Sala Eléctrica #3 (Shelter)","02-Abr-2027","NO INICIADA","⚪","+35d","—","Holgura respecto P6: +35d"),
    ("10","★ ENTREGA S#3 — FECHA CONTRACTUAL","14-May-2027","NO INICIADA","⚪","+45d","—",
     "FECHA MÁXIMA CONTRACTUAL. Penalidad 1%/semana, tope 10%. OC 4508944971"),
]

for c_row in cadena:
    sev = "CRÍTICO" if "🔴" in c_row[4] else ("ADVERTENCIA" if "🟡" in c_row[4] else "INFO")
    w8_row(list(c_row), severity=sev)

_r[0]+=1

# ── RESUMEN EJECUTIVO DE RIESGOS ─────────────────────────────────────────────
w8_sec("RESUMEN EJECUTIVO DE RIESGOS PARA DECISIÓN", bg=C_ALERTA)
risk_texts = [
    "🔴 RIESGO 1 (URGENTE): AESA debe proveer planos de acometida de transformadores MT antes del 20/JUL/2026.",
    "   → SIN ESTA INFORMACIÓN ABB NO PUEDE FINALIZAR EL DISEÑO DE DUCTOS DE BARRAS (DB-005/6/7).",
    "   → El lead time de Turquía es 6 meses. Si OC no sale antes de 26/Oct/2026 → ENTREGA S#3 14-MAY/2027 EN RIESGO.",
    "   → Acción: confirmar recepción de información el 20/Jul/2026. Si no llega → elevar al PM del proyecto inmediatamente.",
    "",
    "🟡 RIESGO 2: Cambio de cargas CCM-006 y CCM-007 → nota de cambio en julio.",
    "   → Puede retrasar IB Rev 0 CCMs y la compra de materiales secundarios.",
    "   → Acción: ABB debe emitir nota de cambio en la 1ra semana de julio y AESA aprobarlo en 5 días hábiles.",
    "",
    "🟡 RIESGO 3: VFDs BT y MT (100% ingeniería) previsto para julio. Vigilar cumplimiento.",
    "   → OC a China/Finlandia ya emitida (reserva slot). FP-IB julio es condición para OC definitiva.",
    "   → Acción: verificar en IA#3 (agosto) que IB VFDs está emitida y aprobada.",
    "",
    "🟡 RIESGO 4: Todas las IBs aprobadas CON OBSERVACIONES → rework en Rev 0 julio.",
    "   → Si Rev 0 tiene más observaciones → segunda ronda de aprobaciones → retraso Hito 2.",
    "   → Acción: seguimiento en reuniones semanales ABB/AESA.",
    "",
    "📅 PRÓXIMO HITO CRÍTICO: 20-JUL-2026 — Recepción info acometida trafos MT de AESA.",
    "📅 SIGUIENTE HITO PAGO: Certificación N°2 (aprobación IB) — Previsto Julio 2026 → Pago Agosto 2026.",
    "📅 FECHA CONTRACTUAL MÁXIMA (CADENA CRÍTICA): 14-MAY-2027 → ENTREGA SALA #3 (OC 4508944971).",
]
for txt in risk_texts:
    r = _r[0]
    ws8.merge_cells(start_row=r, start_column=1, end_row=r, end_column=8)
    c = ws8.cell(row=r, column=1, value=txt)
    if "🔴" in txt:
        c.fill=fill(C_RED); c.font=font(bold=True, size=10)
    elif "🟡" in txt:
        c.fill=fill(C_ORANGE); c.font=font(bold=True, size=10)
    elif "📅" in txt:
        c.fill=fill("D0E8FF"); c.font=font(bold=True, size=10, color="00008B")
    elif txt.strip().startswith("→"):
        c.fill=fill("FFF0F0"); c.font=font(italic=True, size=9)
    elif txt == "":
        c.fill=fill("FFFFFF")
    else:
        c.fill=fill(C_EVEN); c.font=font(size=9)
    c.alignment=left(wrap=True)
    ws8.row_dimensions[r].height=16
    _r[0]+=1

ws8.freeze_panes = "C4"

# ── Save ─────────────────────────────────────────────────────────────────────
wb.save(DST)
print(f"\n✓ Guardado: {DST}")
print(f"  Hojas: {wb.sheetnames}")
print(f"  Revisión: {REV} | Fecha: {FECHA}")
