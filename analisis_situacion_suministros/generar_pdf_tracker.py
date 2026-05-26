#!/usr/bin/env python3
"""
Generador PDF – Tracker Suministros IN/EL La Calera II
Optimizado para NoteShelf 3 (iPad): orientación landscape, colores,
tablas limpias y márgenes amplios para anotaciones a mano.
"""

from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import cm, mm
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Table, TableStyle, Spacer,
    PageBreak, HRFlowable, KeepTogether
)
from reportlab.platypus.flowables import BalancedColumns
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
import os

# ─── PALETA ──────────────────────────────────────────────────────────────────
TITULO      = colors.HexColor('#1F3864')
SUBTITULO   = colors.HexColor('#2F5496')
HDR_IN      = colors.HexColor('#843C0C')
HDR_EL      = colors.HexColor('#375623')
HDR_GRP     = colors.HexColor('#4472C4')
HDR_COL     = colors.HexColor('#BDD7EE')
COMPLETADO  = colors.HexColor('#70AD47')
EN_PROCESO  = colors.HexColor('#FFD966')
ATRASADO    = colors.HexColor('#FF3333')
PENDIENTE   = colors.HexColor('#F2F2F2')
NO_APLICA   = colors.HexColor('#4472C4')
EN_TRANSITO = colors.HexColor('#00B0F0')
APROBADO    = colors.HexColor('#92D050')
SIN_RI      = colors.HexColor('#FF4500')
PLAN_REF    = colors.HexColor('#FFF2CC')
ROW_ALT     = colors.HexColor('#DDEEFF')
ROW_NRM     = colors.white
CRIT_LLI    = colors.HexColor('#FF7F27')
CRIT_CC     = colors.HexColor('#FF4444')
CRIT_MONTO  = colors.HexColor('#FFC000')
CRIT_HITO   = colors.HexColor('#4472C4')
CRIT_MEDIA  = colors.HexColor('#C6E0B4')
URGENTE_BG  = colors.HexColor('#FFE0E0')
URGENTE_FG  = colors.HexColor('#CC0000')

PAGE_W, PAGE_H = landscape(A4)
MARGIN = 1.2 * cm

STATUS_CLR = {
    'COMPLETADO':  (COMPLETADO,  colors.white,   True),
    'COMPLETADA':  (COMPLETADO,  colors.white,   True),
    'APROBADO':    (APROBADO,    colors.white,   True),
    'EN PROCESO':  (EN_PROCESO,  colors.black,   True),
    'EMITIDA':     (EN_PROCESO,  colors.black,   True),
    'ATRASADO':    (ATRASADO,    colors.white,   True),
    'PENDIENTE':   (PENDIENTE,   colors.black,   False),
    'NO APLICA':   (NO_APLICA,   colors.white,   True),
    'EN TRÁNSITO': (EN_TRANSITO, colors.white,   True),
    'SIN RI':      (SIN_RI,      colors.white,   True),
}

CRIT_CLR = {
    'LLI':        (CRIT_LLI,   colors.white),
    'C. CRÍTICO': (CRIT_CC,    colors.white),
    'MONTO':      (CRIT_MONTO, colors.black),
    'HITO 2':     (CRIT_HITO,  colors.white),
    'MEDIA':      (CRIT_MEDIA, colors.black),
}

STATUS_EMOJI = {
    'COMPLETADO':  '✔ COMPLETADO',
    'COMPLETADA':  '✔ COMPLETADA',
    'APROBADO':    '✔ APROBADO',
    'EN PROCESO':  '● EN PROCESO',
    'EMITIDA':     '● EMITIDA',
    'ATRASADO':    '✖ ATRASADO',
    'PENDIENTE':   '○ PENDIENTE',
    'NO APLICA':   '— NO APLICA',
    'EN TRÁNSITO': '» EN TRÁNSITO',
    'SIN RI':      '⊗ SIN RI',
}

styles = getSampleStyleSheet()

def sty(name, parent='Normal', **kw):
    return ParagraphStyle(name, parent=styles[parent], **kw)

TITLE_S  = sty('TitleS',  fontSize=16, textColor=colors.white,
                alignment=TA_CENTER, leading=20, fontName='Helvetica-Bold')
SUB_S    = sty('SubS',    fontSize=11, textColor=colors.white,
                alignment=TA_CENTER, leading=14, fontName='Helvetica-Bold')
META_L   = sty('MetaL',   fontSize=9,  textColor=colors.white,
                alignment=TA_RIGHT,  leading=12, fontName='Helvetica-Bold')
META_V   = sty('MetaV',   fontSize=9,  textColor=colors.black,
                alignment=TA_LEFT,   leading=12)
CELL_S   = sty('CellS',   fontSize=7.5, textColor=colors.black,
                alignment=TA_CENTER, leading=9,  fontName='Helvetica')
CELL_L   = sty('CellL',   fontSize=7.5, textColor=colors.black,
                alignment=TA_LEFT,   leading=9,  fontName='Helvetica')
CELL_B   = sty('CellB',   fontSize=7.5, textColor=colors.black,
                alignment=TA_CENTER, leading=9,  fontName='Helvetica-Bold')
OBS_S    = sty('ObsS',    fontSize=6.5, textColor=colors.black,
                alignment=TA_LEFT,   leading=8,  fontName='Helvetica')
ACCION_S = sty('AcS',     fontSize=6.5, textColor=URGENTE_FG,
                alignment=TA_LEFT,   leading=8,  fontName='Helvetica-Bold')
ACCION_N = sty('AcN',     fontSize=6.5, textColor=colors.black,
                alignment=TA_LEFT,   leading=8,  fontName='Helvetica')
HDR_C_S  = sty('HdrC',   fontSize=7,   textColor=TITULO,
                alignment=TA_CENTER, leading=8,  fontName='Helvetica-Bold',
                wordWrap='CJK')
SECT_S   = sty('SectS',  fontSize=10,  textColor=colors.white,
                alignment=TA_LEFT,   leading=13, fontName='Helvetica-Bold')
LEGEND_S = sty('LegS',   fontSize=8,   textColor=colors.black,
                alignment=TA_LEFT,   leading=10)
DASH_KPI = sty('DKpi',   fontSize=22,  textColor=colors.white,
                alignment=TA_CENTER, leading=26, fontName='Helvetica-Bold')
DASH_LBL = sty('DLbl',   fontSize=8,   textColor=colors.white,
                alignment=TA_CENTER, leading=10, fontName='Helvetica-Bold',
                wordWrap='CJK')


def p(txt, style=CELL_S):
    return Paragraph(str(txt), style)


def status_para(status):
    s = status.strip().upper()
    bg, fg, bold = STATUS_CLR.get(s, (PENDIENTE, colors.black, False))
    txt = STATUS_EMOJI.get(s, status)
    color_hex = fg.hexval() if hasattr(fg, 'hexval') else '#000000'
    fn = 'Helvetica-Bold' if bold else 'Helvetica'
    style = sty(f'st_{s}', fontSize=7, textColor=fg,
                alignment=TA_CENTER, leading=8, fontName=fn)
    return Paragraph(txt, style), bg


def crit_para(crit):
    bg, fg = CRIT_CLR.get(crit, (PENDIENTE, colors.black))
    style = sty(f'cr_{crit}', fontSize=7.5, textColor=fg,
                alignment=TA_CENTER, leading=9, fontName='Helvetica-Bold')
    return Paragraph(crit, style), bg


# ─── HEADER / FOOTER ─────────────────────────────────────────────────────────
def add_header_footer(canvas_obj, doc, specialty='', page_label=''):
    canvas_obj.saveState()
    w, h = landscape(A4)
    # Banda superior
    canvas_obj.setFillColor(TITULO)
    canvas_obj.rect(0, h - 1.1*cm, w, 1.1*cm, fill=1, stroke=0)
    canvas_obj.setFont('Helvetica-Bold', 9)
    canvas_obj.setFillColor(colors.white)
    canvas_obj.drawString(MARGIN, h - 0.75*cm,
        'TRACKER SUMINISTROS  ·  LA CALERA II – CPF2  ·  Rev.0  ·  22-May-2026')
    if specialty:
        canvas_obj.setFont('Helvetica-Bold', 9)
        canvas_obj.setFillColor(colors.HexColor('#FFD966'))
        canvas_obj.drawRightString(w - MARGIN, h - 0.75*cm, specialty)
    # Banda inferior
    canvas_obj.setFillColor(SUBTITULO)
    canvas_obj.rect(0, 0, w, 0.7*cm, fill=1, stroke=0)
    canvas_obj.setFont('Helvetica', 7.5)
    canvas_obj.setFillColor(colors.white)
    canvas_obj.drawString(MARGIN, 0.22*cm,
        f'RFSU Planificado: 03-Feb-2027  ·  Hito Ing.80%: 01-Jun-2026')
    canvas_obj.drawRightString(w - MARGIN, 0.22*cm,
        f'{page_label}  Pág. {canvas_obj.getPageNumber()}')
    canvas_obj.restoreState()


# ─── PORTADA ──────────────────────────────────────────────────────────────────
def build_portada():
    elems = []
    elems.append(Spacer(1, 1.5*cm))

    title_data = [[Paragraph(
        'TRACKER DE SEGUIMIENTO DE SUMINISTROS<br/>'
        '<font size=12 color="#FFD966">ESPECIALIDADES: INSTRUMENTACIÓN &amp; CONTROL  ·  ELECTRICIDAD</font>',
        sty('T', fontSize=18, textColor=colors.white, alignment=TA_CENTER,
            leading=26, fontName='Helvetica-Bold')
    )]]
    title_tbl = Table(title_data, colWidths=[PAGE_W - 2*MARGIN])
    title_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), TITULO),
        ('TOPPADDING', (0,0), (-1,-1), 14),
        ('BOTTOMPADDING', (0,0), (-1,-1), 14),
        ('LEFTPADDING', (0,0), (-1,-1), 16),
        ('RIGHTPADDING', (0,0), (-1,-1), 16),
        ('ROUNDEDCORNERS', [6]),
    ]))
    elems.append(title_tbl)
    elems.append(Spacer(1, 0.8*cm))

    # Metadatos
    meta = [
        ['Proyecto:',              'La Calera II – CPF2',
         'Contrato:',              'CPF2'],
        ['Plan P0 emitido:',       '06-May-2026',
         'RFSU Planificado:',      '03-Feb-2027'],
        ['Hito Ing. 80%:',         '01-Jun-2026',
         'Fecha de referencia:',   '22-May-2026'],
        ['Revisión:',              'Rev. 0',
         'Actualizado por:',       ''],
    ]
    meta_rows = []
    for row in meta:
        meta_rows.append([
            p(row[0], META_L), p(row[1], META_V),
            p(row[2], META_L), p(row[3], META_V),
        ])
    meta_tbl = Table(meta_rows,
                     colWidths=[3.5*cm, 5.5*cm, 4*cm, 6*cm])
    meta_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (0,-1), SUBTITULO),
        ('BACKGROUND', (1,0), (1,-1), colors.HexColor('#DEEAF1')),
        ('BACKGROUND', (2,0), (2,-1), SUBTITULO),
        ('BACKGROUND', (3,0), (3,-1), colors.HexColor('#DEEAF1')),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#8EA9C1')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]))
    elems.append(meta_tbl)
    elems.append(Spacer(1, 0.8*cm))

    # Leyenda
    leyenda_header = [[Paragraph('LEYENDA – ESTADOS', SECT_S)]]
    lt = Table(leyenda_header, colWidths=[PAGE_W - 2*MARGIN])
    lt.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), SUBTITULO),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
    ]))
    elems.append(lt)
    elems.append(Spacer(1, 0.3*cm))

    leyenda_items = [
        ('✔ COMPLETADO / APROBADO', COMPLETADO, colors.white,
         'Hito/etapa finalizada en tiempo o antes'),
        ('● EN PROCESO / EMITIDA',  EN_PROCESO, colors.black,
         'Actividad en curso dentro del plazo programado'),
        ('✖ ATRASADO',              ATRASADO,   colors.white,
         'Fecha programada superada sin completar la etapa'),
        ('○ PENDIENTE',             PENDIENTE,  colors.black,
         'No iniciado, dentro del plazo programado'),
        ('— NO APLICA',             NO_APLICA,  colors.white,
         'El ítem no requiere esta etapa'),
        ('» EN TRÁNSITO',           EN_TRANSITO,colors.white,
         'Material despachado, en camino a obra'),
        ('⊗ SIN RI',                SIN_RI,     colors.white,
         'Requisición de Ingeniería no localizada / no emitida'),
    ]
    leg_data = []
    for txt, bg, fg, desc in leyenda_items:
        sty_l = sty('ll', fontSize=8, textColor=fg,
                    alignment=TA_CENTER, leading=10, fontName='Helvetica-Bold')
        leg_data.append([
            Paragraph(txt, sty_l),
            p(desc, sty('ld', fontSize=8.5, textColor=colors.black,
                        alignment=TA_LEFT, leading=11)),
        ])
    leg_tbl = Table(leg_data, colWidths=[5*cm, 12*cm])
    leg_cmds = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#8EA9C1')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]
    for i, (_, bg, _, _) in enumerate(leyenda_items):
        leg_cmds.append(('BACKGROUND', (0,i), (0,i), bg))
        leg_cmds.append(('BACKGROUND', (1,i), (1,i),
                         colors.HexColor('#F5F5F5')))
    leg_tbl.setStyle(TableStyle(leg_cmds))
    elems.append(leg_tbl)

    elems.append(Spacer(1, 0.6*cm))

    # Criticidades
    crit_header = [[Paragraph('NIVELES DE CRITICIDAD', SECT_S)]]
    ct = Table(crit_header, colWidths=[PAGE_W - 2*MARGIN])
    ct.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), SUBTITULO),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
    ]))
    elems.append(ct)
    elems.append(Spacer(1, 0.3*cm))

    crit_items = [
        ('LLI',        CRIT_LLI,   colors.white, 'Long Lead Item – Equipment with extended lead time (>180 days)'),
        ('C. CRÍTICO', CRIT_CC,    colors.white, 'Ítem en ruta crítica del proyecto. Impacta directamente en RFSU.'),
        ('HITO 2',     CRIT_HITO,  colors.white, 'Asociado a hito contractual Hito 2 del cronograma.'),
        ('MONTO',      CRIT_MONTO, colors.black, 'Relevante por monto de inversión.'),
        ('MEDIA',      CRIT_MEDIA, colors.black, 'Criticidad media. Seguimiento estándar.'),
    ]
    crit_data = []
    for crit, bg, fg, desc in crit_items:
        cs = sty('cs', fontSize=9, textColor=fg, alignment=TA_CENTER,
                 leading=11, fontName='Helvetica-Bold')
        crit_data.append([
            Paragraph(crit, cs),
            p(desc, sty('cd', fontSize=8.5, textColor=colors.black,
                        alignment=TA_LEFT, leading=11)),
        ])
    crit_tbl = Table(crit_data, colWidths=[3.5*cm, 13.5*cm])
    crit_cmds = [
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor('#8EA9C1')),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
    ]
    for i, (_, bg, _, _) in enumerate(crit_items):
        crit_cmds.append(('BACKGROUND', (0,i), (0,i), bg))
        crit_cmds.append(('BACKGROUND', (1,i), (1,i),
                          colors.HexColor('#F5F5F5')))
    crit_tbl.setStyle(TableStyle(crit_cmds))
    elems.append(crit_tbl)

    elems.append(Spacer(1, 0.5*cm))
    note_data = [[p(
        'NOTA: Las fechas con fondo amarillo claro corresponden al Plan de Suministros P0 '
        '(baseline 06-May-2026). Las fechas sin fondo son valores reales o pronóstico (FC).',
        sty('note', fontSize=8.5, textColor=colors.HexColor('#595959'),
            alignment=TA_LEFT, leading=11))]]
    note_tbl = Table(note_data, colWidths=[PAGE_W - 2*MARGIN])
    note_tbl.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), PLAN_REF),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('BOX', (0,0), (-1,-1), 0.5, colors.HexColor('#8EA9C1')),
    ]))
    elems.append(note_tbl)

    return elems


# ─── DASHBOARD ────────────────────────────────────────────────────────────────
def build_dashboard(in_data, el_data):
    elems = []
    elems.append(Spacer(1, 0.4*cm))

    def cnt(data, field, val):
        return sum(1 for d in data if val.upper() in d[field].upper())

    def kpi_table(label, data, color):
        total = len(data)
        comp  = cnt(data, 'est_gen', 'COMPLETADO')
        proc  = cnt(data, 'est_gen', 'EN PROCESO')
        atr   = cnt(data, 'est_gen', 'ATRASADO')
        pend  = cnt(data, 'est_gen', 'PENDIENTE')

        header = [[Paragraph(label, sty('kh', fontSize=11, textColor=colors.white,
                                        alignment=TA_CENTER, leading=14,
                                        fontName='Helvetica-Bold'))]]
        ht = Table(header, colWidths=[PAGE_W/2 - MARGIN - 0.2*cm])
        ht.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), color),
            ('TOPPADDING', (0,0), (-1,-1), 8),
            ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ]))

        kpis = [
            ('Total RIs', str(total), SUBTITULO),
            ('✔ Completados', str(comp), COMPLETADO),
            ('● En Proceso', str(proc), EN_PROCESO),
            ('✖ Atrasados', str(atr), ATRASADO),
            ('○ Pendientes', str(pend), PENDIENTE),
        ]
        kpi_row_l = []
        kpi_row_v = []
        kpi_bgs_l = []
        kpi_bgs_v = []
        for lbl, val, bg in kpis:
            fg = colors.white if bg not in (EN_PROCESO, PENDIENTE) else colors.black
            kpi_row_l.append(Paragraph(lbl, sty('kl', fontSize=7.5,
                             textColor=fg, alignment=TA_CENTER,
                             leading=9, fontName='Helvetica-Bold',
                             wordWrap='CJK')))
            kpi_row_v.append(Paragraph(val, sty('kv', fontSize=20,
                             textColor=fg, alignment=TA_CENTER,
                             leading=24, fontName='Helvetica-Bold')))
            kpi_bgs_l.append(bg)
            kpi_bgs_v.append(bg)

        n = len(kpis)
        w = (PAGE_W/2 - MARGIN - 0.2*cm) / n
        kpi_t = Table([kpi_row_l, kpi_row_v],
                      colWidths=[w]*n, rowHeights=[0.6*cm, 1.2*cm])
        kpi_cmds = [
            ('GRID', (0,0), (-1,-1), 0.5, colors.white),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]
        for i, bg in enumerate(kpi_bgs_l):
            kpi_cmds.append(('BACKGROUND', (i,0), (i,0), bg))
            kpi_cmds.append(('BACKGROUND', (i,1), (i,1), bg))
        kpi_t.setStyle(TableStyle(kpi_cmds))
        return [ht, kpi_t]

    kpi_in = kpi_table('IN – INSTRUMENTACIÓN & CONTROL  (22 requisiciones)',
                        in_data, HDR_IN)
    kpi_el = kpi_table('EL – ELECTRICIDAD  (16 requisiciones)',
                        el_data, HDR_EL)

    w_half = PAGE_W/2 - MARGIN - 0.2*cm
    dash_cols = Table(
        [[Table([[k] for k in kpi_in], colWidths=[w_half]),
          Table([[k] for k in kpi_el], colWidths=[w_half])]],
        colWidths=[w_half + 0.3*cm, w_half + 0.3*cm]
    )
    dash_cols.setStyle(TableStyle([
        ('LEFTPADDING', (0,0), (-1,-1), 0),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 0),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    elems.append(dash_cols)
    elems.append(Spacer(1, 0.5*cm))

    # ── Tabla de ítems críticos ───────────────────────────────────────────────
    crit_hdr = [[Paragraph(
        'ÍTEMS CRÍTICOS / EN RIESGO PARA RFSU  03-FEB-2027',
        sty('ch', fontSize=10, textColor=colors.white, alignment=TA_CENTER,
            leading=13, fontName='Helvetica-Bold'))]]
    ct = Table(crit_hdr, colWidths=[PAGE_W - 2*MARGIN])
    ct.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), ATRASADO),
        ('TOPPADDING', (0,0), (-1,-1), 7),
        ('BOTTOMPADDING', (0,0), (-1,-1), 7),
    ]))
    elems.append(ct)
    elems.append(Spacer(1, 0.2*cm))

    col_names = ['Esp.', 'N° RI', 'Descripción', 'Crit.',
                 'Est. RI', 'Nec. OC (P0)', 'L.T.', 'Nec. Obra',
                 'Est. Gral.', 'Acción Clave']
    col_ws = [1.0*cm, 2.0*cm, 6.0*cm, 2.0*cm,
              2.2*cm, 2.2*cm, 1.5*cm, 2.2*cm,
              2.2*cm, 0]
    # last col fills remaining
    used = sum(col_ws[:-1])
    col_ws[-1] = PAGE_W - 2*MARGIN - used

    hdr_row = [Paragraph(h, HDR_C_S) for h in col_names]

    criticos = [
        ('IN', 'RI-016', 'Sistema de Seguridad SIS',          'LLI',
         'EN PROCESO', '21-JUN-26', '300d', '18-ABR-27',
         'EN PROCESO', 'OC CRÍTICA fc 21-Jun-26. AT urgente.'),
        ('IN', 'RI-015', 'Sistema de Control PCS (Inauco)',   'LLI',
         'EN PROCESO', '19-AGO-26', '300d', '24-JUN-27',
         'EN PROCESO', 'AT fc Jul-26. OC fc Feb-27. LT 300d.'),
        ('IN', 'RI-022', 'Materiales Montaje Instrumentos',   'C. CRÍTICO',
         'ATRASADO',   '01-NOV-26', '90d',  '03-FEB-27',
         'ATRASADO',   'Definir alcance y emitir RI URGENTE.'),
        ('IN', 'RI-020', 'Cables Instrumentación Rev.0',      'C. CRÍTICO',
         'ATRASADO',   '21-OCT-26', '150d', '29-MAR-27',
         'ATRASADO',   'Emitir RI fc May-26. Dep. plano cables.'),
        ('IN', 'RI-019', 'Bandejas Portacables IN',           'C. CRÍTICO',
         'ATRASADO',   '01-JUL-26', '90d',  '03-FEB-27',
         'ATRASADO',   'Verificar RI activa. Emitir urgente.'),
        ('EL', 'RI-EL-006', 'Tableros BT y CCM',             'C. CRÍTICO',
         'ATRASADO',   '01-JUN-26', '300d', '03-FEB-27',
         'ATRASADO',   'Emitir RI INMEDIATO. LT 10-12 meses.'),
        ('EL', 'RI-EL-007', 'Sistema UPS y Rectificadores',  'C. CRÍTICO',
         'ATRASADO',   '01-JUN-26', '300d', '03-FEB-27',
         'ATRASADO',   'Emitir RI INMEDIATO. LT ~10 meses.'),
        ('EL', 'RI-EL-008', 'Ducto de Barras',               'C. CRÍTICO',
         'ATRASADO',   '01-JUN-26', '240d', '03-FEB-27',
         'ATRASADO',   'Emitir RI INMEDIATO. LT 8-10 meses.'),
        ('EL', 'RI-EL-009', 'Cables Eléctricos BT/MT Rev.0', 'C. CRÍTICO',
         'ATRASADO',   '18-NOV-26', '150d', '26-ABR-27',
         'ATRASADO',   'Emitir RI fc Jun-26. Dep. plano Rev.0.'),
        ('EL', 'RI-EL-011', 'Canalizaciones Eléctricas',     'C. CRÍTICO',
         'ATRASADO',   '01-JUL-26', '120d', '01-OCT-26',
         'ATRASADO',   'Entrega fc Ene-27 vs nec. Oct-26 (+90d).'),
        ('EL', 'RI-EL-004', 'Shelter SE#5 (sin RI)',         'HITO 2',
         'SIN RI',     '01-JUN-26', '390d', '01-JUL-27',
         'ATRASADO',   'Confirmar alcance. Emitir RI si aplica.'),
        ('EL', 'RI-EL-005', 'Tableros MT + Transformadores', 'LLI',
         'EN PROCESO',  '01-ENE-27', '360d', '01-JUN-27',
         'EN PROCESO',  'Cerrar AT urgente. LT ~12 meses.'),
    ]

    rows = [hdr_row]
    for row in criticos:
        (esp, n_ri, desc, crit, est_ri, nec_oc,
         lt, nec_obra, est_gen, accion) = row
        sp, sbg = status_para(est_ri)
        sg, sgbg = status_para(est_gen)
        cp, cbg = crit_para(crit)
        esp_color = HDR_IN if esp == 'IN' else HDR_EL
        rows.append([
            Paragraph(esp, sty('ec', fontSize=8, textColor=colors.white,
                               alignment=TA_CENTER, fontName='Helvetica-Bold')),
            p(n_ri, CELL_S),
            p(desc, CELL_L),
            cp,
            sp,
            Paragraph(nec_oc, sty('nc', fontSize=7.5, textColor=colors.HexColor('#595959'),
                                   alignment=TA_CENTER, leading=9)),
            p(lt, CELL_S),
            Paragraph(nec_obra, sty('no', fontSize=7.5, textColor=colors.black,
                                     alignment=TA_CENTER, fontName='Helvetica-Bold',
                                     leading=9)),
            sg,
            Paragraph(accion, sty('ac', fontSize=6.5,
                                   textColor=URGENTE_FG if '🚨' in accion or 'INMEDIATO' in accion or 'URGENTE' in accion else colors.black,
                                   alignment=TA_LEFT, leading=8,
                                   fontName='Helvetica-Bold' if 'INMEDIATO' in accion else 'Helvetica')),
        ])

    tbl_cmds = [
        ('BACKGROUND', (0,0), (-1,0), HDR_COL),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor('#8EA9C1')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 3),
        ('RIGHTPADDING', (0,0), (-1,-1), 3),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0,1), (-1,-1),
         [ROW_ALT, ROW_NRM]),
    ]
    for i, row in enumerate(criticos):
        r = i + 1
        esp = row[0]
        tbl_cmds.append(('BACKGROUND', (0,r), (0,r),
                          HDR_IN if esp == 'IN' else HDR_EL))
        # Criticidad col
        crit = row[3]
        bg, _ = CRIT_CLR.get(crit, (PENDIENTE, colors.black))
        tbl_cmds.append(('BACKGROUND', (3,r), (3,r), bg))
        # Status cols
        est_ri = row[4]
        bg_ri, _, _ = STATUS_CLR.get(est_ri.upper(), (PENDIENTE, colors.black, False))
        tbl_cmds.append(('BACKGROUND', (4,r), (4,r), bg_ri))
        est_gen = row[8]
        bg_gen, _, _ = STATUS_CLR.get(est_gen.upper(), (PENDIENTE, colors.black, False))
        tbl_cmds.append(('BACKGROUND', (8,r), (8,r), bg_gen))

    crit_tbl = Table(rows, colWidths=col_ws, repeatRows=1)
    crit_tbl.setStyle(TableStyle(tbl_cmds))
    elems.append(crit_tbl)
    return elems


# ─── TABLA DETALLE IN / EL ────────────────────────────────────────────────────
def build_detail_section(data, specialty, color_header):
    """Genera páginas de detalle para una especialidad, agrupando cols."""
    all_elems = []

    # ── BLOQUE 1: Identificación + RI + SOLPED ──────────────────────────────
    block1_cols = ['N°', 'Crit.', 'N° RI', 'Descripción',
                   'F.RI\nPlan P0', 'F.RI\nReal/FC', 'Desv.\nRI(d)', 'Est.RI',
                   'N° SOLPED', 'F.SOLP\nPlan P0', 'F.SOLP\nReal/FC', 'Est.SOLP']
    b1_ws = [0.6*cm, 1.6*cm, 2.0*cm, 6.0*cm,
             1.8*cm, 1.8*cm, 1.5*cm, 2.2*cm,
             3.0*cm, 1.8*cm, 1.8*cm, 2.2*cm]
    used = sum(b1_ws)
    b1_ws[3] += PAGE_W - 2*MARGIN - used  # expand description col

    # ── BLOQUE 2: Rec.Of + AT + OC ──────────────────────────────────────────
    block2_cols = ['N°', 'N° RI', 'Descripción',
                   'F.RecOf\nPlan P0', 'F.RecOf\nReal/FC', 'Est.Rec.Of',
                   'F.AT\nPlan P0', 'F.AT\nReal/FC', 'Est.AT',
                   'N° OC', 'Proveedor', 'F.Nec.OC\nPlan P0', 'L.T.(d)', 'Est.OC']
    b2_ws = [0.6*cm, 2.0*cm, 4.5*cm,
             1.8*cm, 1.8*cm, 2.0*cm,
             1.8*cm, 1.8*cm, 2.0*cm,
             2.8*cm, 3.0*cm, 1.8*cm, 1.5*cm, 2.0*cm]
    used2 = sum(b2_ws)
    b2_ws[2] += PAGE_W - 2*MARGIN - used2

    # ── BLOQUE 3: Entrega + Estado + Observaciones ────────────────────────────
    block3_cols = ['N°', 'N° RI', 'Descripción',
                   'F.Entrega\nOC', 'F.Nec.\nObra', 'Desv.\nEnt.(d)', 'Est.\nEntrega',
                   'Estado\nGeneral', 'Observaciones', 'Acción Requerida']
    b3_ws = [0.6*cm, 2.0*cm, 4.0*cm,
             1.8*cm, 1.8*cm, 1.5*cm, 2.0*cm,
             2.0*cm, 0, 0]
    used3 = sum(b3_ws[:-2])
    rem = PAGE_W - 2*MARGIN - used3
    b3_ws[-2] = rem * 0.38   # observaciones
    b3_ws[-1] = rem * 0.62   # acción

    def make_header_row(col_names, h_color):
        return [Paragraph(cn.replace('\n', '<br/>'), HDR_C_S)
                for cn in col_names]

    def section_title(block_label):
        hd = [[Paragraph(
            f'{specialty}  ·  {block_label}',
            sty('sh', fontSize=9, textColor=colors.white, alignment=TA_LEFT,
                leading=11, fontName='Helvetica-Bold'))]]
        t = Table(hd, colWidths=[PAGE_W - 2*MARGIN])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), color_header),
            ('TOPPADDING', (0,0), (-1,-1), 5),
            ('BOTTOMPADDING', (0,0), (-1,-1), 5),
            ('LEFTPADDING', (0,0), (-1,-1), 8),
        ]))
        return t

    def desvio_cell(d, bg_row):
        if d is None:
            return p('', CELL_S), bg_row
        if d > 30:
            return (Paragraph(f'+{d}d',
                    sty('dv', fontSize=7.5, textColor=colors.white,
                        alignment=TA_CENTER, fontName='Helvetica-Bold')),
                    ATRASADO)
        elif d > 0:
            return (Paragraph(f'+{d}d',
                    sty('dv', fontSize=7.5, textColor=colors.black,
                        alignment=TA_CENTER, fontName='Helvetica-Bold')),
                    CRIT_MONTO)
        elif d < 0:
            return (Paragraph(f'{d}d',
                    sty('dv', fontSize=7.5, textColor=colors.white,
                        alignment=TA_CENTER, fontName='Helvetica-Bold')),
                    COMPLETADO)
        return (Paragraph('0d',
                sty('dv', fontSize=7.5, textColor=colors.white,
                    alignment=TA_CENTER, fontName='Helvetica-Bold')),
                COMPLETADO)

    def date_p(d, is_plan=False):
        if not d:
            return p('', CELL_S)
        style = sty('dp', fontSize=7.5,
                    textColor=colors.HexColor('#595959') if is_plan else colors.black,
                    alignment=TA_CENTER, leading=9)
        return Paragraph(d, style)

    def fmt(d):
        if not d: return ''
        try:
            from datetime import datetime
            return datetime.strptime(d, '%Y-%m-%d').strftime('%d-%b-%y').upper()
        except Exception:
            return d

    # ── Build rows for each block ─────────────────────────────────────────────
    rows_b1, rows_b2, rows_b3 = [], [], []
    cmds_b1 = base_cmds()
    cmds_b2 = base_cmds()
    cmds_b3 = base_cmds()

    for idx, item in enumerate(data):
        r = idx + 1
        bg = ROW_ALT if idx % 2 == 0 else ROW_NRM
        urgent = '🚨' in item['accion'] or 'INMEDIATO' in item['accion']

        sp_ri,  bg_ri  = status_para(item['est_ri'])
        sp_sp,  bg_sp  = status_para(item['est_solp'])
        sp_ro,  bg_ro  = status_para(item['est_recof'])
        sp_at,  bg_at  = status_para(item['est_at'])
        sp_oc,  bg_oc  = status_para(item['est_oc'])
        sp_en,  bg_en  = status_para(item['est_ent'])
        sp_gen, bg_gen = status_para(item['est_gen'])
        cp, cbg        = crit_para(item['crit'])
        dv_ri,  bg_dri = desvio_cell(item.get('desv_ri'), bg)
        dv_en,  bg_den = desvio_cell(item.get('desv_ent'), bg)

        # Block 1
        rows_b1.append([
            p(str(item['n']), CELL_B),
            cp,
            p(item['n_ri'], CELL_S),
            Paragraph(item['desc'], CELL_L),
            date_p(fmt(item['ri_prog']), True),
            date_p(fmt(item['ri_real'])),
            dv_ri,
            sp_ri,
            Paragraph(item['solped'], sty('sl', fontSize=6.5,
                      textColor=colors.black, alignment=TA_CENTER, leading=8)),
            date_p(fmt(item['solp_prog']), True),
            date_p(fmt(item['solp_real'])),
            sp_sp,
        ])
        cmds_b1 += [
            ('BACKGROUND', (0,r), (0,r), bg),
            ('BACKGROUND', (1,r), (1,r), cbg),
            ('BACKGROUND', (6,r), (6,r), bg_dri),
            ('BACKGROUND', (7,r), (7,r), bg_ri),
            ('BACKGROUND', (11,r), (11,r), bg_sp),
        ]
        for c in range(12):
            if c not in (1, 6, 7, 11):
                cmds_b1.append(('BACKGROUND', (c,r), (c,r), bg))

        # Block 2
        rows_b2.append([
            p(str(item['n']), CELL_B),
            p(item['n_ri'], CELL_S),
            Paragraph(item['desc'], CELL_L),
            date_p(fmt(item['recof_prog']), True),
            date_p(fmt(item['recof_real'])),
            sp_ro,
            date_p(fmt(item['at_prog']), True),
            date_p(fmt(item['at_real'])),
            sp_at,
            Paragraph(item['n_oc'], sty('oc', fontSize=6.5,
                      textColor=colors.black, alignment=TA_CENTER, leading=8)),
            Paragraph(item['prov'], CELL_L),
            date_p(fmt(item['nec_oc']), True),
            p(str(item['lt']) if item['lt'] else '', CELL_S),
            sp_oc,
        ])
        cmds_b2 += [
            ('BACKGROUND', (5,r),  (5,r),  bg_ro),
            ('BACKGROUND', (8,r),  (8,r),  bg_at),
            ('BACKGROUND', (13,r), (13,r), bg_oc),
        ]
        for c in range(14):
            if c not in (5, 8, 13):
                cmds_b2.append(('BACKGROUND', (c,r), (c,r), bg))

        # Block 3
        rows_b3.append([
            p(str(item['n']), CELL_B),
            p(item['n_ri'], CELL_S),
            Paragraph(item['desc'], CELL_L),
            date_p(fmt(item['ent_oc'])),
            date_p(fmt(item['nec_obra']), True),
            dv_en,
            sp_en,
            sp_gen,
            Paragraph(item['obs'], OBS_S),
            Paragraph(item['accion'],
                      ACCION_S if urgent else ACCION_N),
        ])
        cmds_b3 += [
            ('BACKGROUND', (5,r),  (5,r),  bg_den),
            ('BACKGROUND', (6,r),  (6,r),  bg_en),
            ('BACKGROUND', (7,r),  (7,r),  bg_gen),
            ('BACKGROUND', (9,r),  (9,r),  URGENTE_BG if urgent else bg),
        ]
        for c in range(10):
            if c not in (5, 6, 7, 9):
                cmds_b3.append(('BACKGROUND', (c,r), (c,r), bg))

    def make_table(rows, col_ws, cmds, hdr_names):
        hdr = [make_header_row(hdr_names, HDR_COL)]
        tbl = Table(hdr + rows, colWidths=col_ws, repeatRows=1)
        tbl.setStyle(TableStyle(cmds))
        return tbl

    # Block 1
    all_elems.append(section_title('BLOQUE 1 – IDENTIFICACIÓN · RI · SOLPED'))
    all_elems.append(Spacer(1, 0.15*cm))
    all_elems.append(make_table(rows_b1, b1_ws, cmds_b1, block1_cols))
    all_elems.append(PageBreak())

    # Block 2
    all_elems.append(section_title('BLOQUE 2 – REC.OFERTAS · ANÁLISIS TÉCNICO · ORDEN DE COMPRA'))
    all_elems.append(Spacer(1, 0.15*cm))
    all_elems.append(make_table(rows_b2, b2_ws, cmds_b2, block2_cols))
    all_elems.append(PageBreak())

    # Block 3
    all_elems.append(section_title('BLOQUE 3 – ENTREGA · ESTADO GENERAL · OBSERVACIONES · ACCIONES'))
    all_elems.append(Spacer(1, 0.15*cm))
    all_elems.append(make_table(rows_b3, b3_ws, cmds_b3, block3_cols))

    return all_elems


def base_cmds():
    return [
        ('BACKGROUND', (0,0), (-1,0), HDR_COL),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor('#8EA9C1')),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 2),
        ('RIGHTPADDING', (0,0), (-1,-1), 2),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE', (0,0), (-1,-1), 7.5),
        ('ROWHEIGHT', (0,1), (-1,-1), 0.75*cm),
    ]


# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    from generar_tracker_suministros import get_in_data, get_el_data
    in_data = get_in_data()
    el_data = get_el_data()

    output = '/home/user/OWN-Daro-1/Tracker_Suministros_IN_EL_LaCalera_II.pdf'

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=1.5*cm, bottomMargin=1.0*cm,
        title='Tracker Suministros IN/EL – La Calera II',
        author='La Calera II CPF2',
    )

    story = []

    # ── Page 1: Portada ───────────────────────────────────────────────────────
    story += build_portada()
    story.append(PageBreak())

    # ── Page 2: Dashboard ────────────────────────────────────────────────────
    story += build_dashboard(in_data, el_data)
    story.append(PageBreak())

    # ── Pages 3-5: IN detail (3 bloques) ─────────────────────────────────────
    story += build_detail_section(
        in_data,
        'IN – INSTRUMENTACIÓN & CONTROL  (22 RI)',
        HDR_IN)
    story.append(PageBreak())

    # ── Pages 6-8: EL detail (3 bloques) ─────────────────────────────────────
    story += build_detail_section(
        el_data,
        'EL – ELECTRICIDAD  (16 RI)',
        HDR_EL)

    # Build with header/footer
    current_specialty = ['']
    current_block     = ['']

    def on_page(canvas_obj, doc_obj):
        pn = canvas_obj.getPageNumber()
        if pn == 1:
            sp = ''
            bl = 'PORTADA'
        elif pn == 2:
            sp = ''
            bl = 'DASHBOARD'
        elif pn <= 5:
            sp = 'IN – INSTRUMENTACIÓN & CONTROL'
            bl = f'Bloque {pn-2}'
        else:
            sp = 'EL – ELECTRICIDAD'
            bl = f'Bloque {pn-5}'
        add_header_footer(canvas_obj, doc_obj, sp, bl)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    print(f'✅ PDF generado: {output}')
    print(f'   Páginas estimadas: ~8  (Portada + Dashboard + 3×IN + 3×EL)')


if __name__ == '__main__':
    main()
