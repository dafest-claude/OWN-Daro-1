#!/usr/bin/env python3
"""
Dashboard Gerencial – Análisis de Tiempos RI → OC
La Calera II CPF2 | Especialidades: Instrumentación (IN) y Electricidad (EL)
Fuente: Plan de Suministros (210826, Rev11) – hojas Cuadro resumen, Suministros críticos,
        RI y OC x mes.
Objetivo: Cuantificar el tiempo entre emisión de RI y colocación efectiva de OC,
          identificar dónde se concentran las demoras del circuito de suministros.
"""

import os
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference, Series
from datetime import date, datetime

import glob as _glob
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
def _find_plan(tag):
    hits = [f for f in _glob.glob(os.path.join(SCRIPT_DIR, '..', 'info_suministros', '*.xlsx'))
            if f'({tag})' in os.path.basename(f)]
    return hits[0] if hits else None
PLAN_FILE     = _find_plan('210826')
PLAN_FILE_OLD = _find_plan('140826')
OUT_XLSX   = os.path.join(SCRIPT_DIR, 'Dashboard_RI_OC_IN_EL_LaCalera_II_Rev11_210826.xlsx')
TODAY      = date(2026, 8, 21)
RFSU       = date(2027, 2, 3)
VERSION    = 'Rev11 · 210826'

# ── Paleta gerencial ─────────────────────────────────────────────────────────
C = {
    'azul_h':   '1F3864', 'azul_m':  '2F5496', 'azul_cl': 'D6E4F0',
    'el':       '2E6B2E', 'el_cl':   'D7E9D7',
    'in':       '9C4500', 'in_cl':   'F5E0D0',
    'gris':     '404040', 'gris_cl': 'F2F2F2',
    'blanco':   'FFFFFF',
    'ok':       '548235', 'ok_cl':   'C6EFCE',
    'warn':     'BF8F00', 'warn_cl': 'FFEB9C',
    'crit':     'C00000', 'crit_cl': 'FFC7CE',
    'kpi1':     '1F3864', 'kpi2':    '2E6B2E', 'kpi3':  '9C4500', 'kpi4': '7030A0',
    'borde':    'A6A6A6',
}

def F(c): return PatternFill('solid', fgColor=c)
def ft(b=False, c='000000', s=10, name='Calibri'): return Font(bold=b, color=c, size=s, name=name)
def al(h='center', v='center', wrap=False): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def bd(c=C['borde'], style='thin'):
    s = Side(style=style, color=c); return Border(left=s, right=s, top=s, bottom=s)
def bd_med(): return bd('000000', 'medium')

def dias(d1, d2):
    if isinstance(d1, date) and isinstance(d2, date): return (d2 - d1).days
    return None


# ═══════════════════════════════════════════════════════════════════════════
# DATOS (extraídos y verificados del Plan de Suministros 210826 · Rev11)
# ═══════════════════════════════════════════════════════════════════════════

# Pipeline actual por especialidad (hoja "Cuadro resumen" – plan 210826)
#   cant_ri, emitidas(en gestión), solped, ofertas, at, oc
PIPELINE = {
    'EL': {'nombre': 'ELECTRICIDAD',     'cant_ri': 19, 'emitidas': 15,
           'solped': 0, 'ofertas': 0, 'at': 7, 'oc': 8},
    'IN': {'nombre': 'INSTRUMENTACIÓN',  'cant_ri': 50, 'emitidas': 46,
           'solped': 0, 'ofertas': 2, 'at': 28, 'oc': 16},
}

# Pipeline del corte anterior (plan 140826) – comparativa
PIPELINE_OLD = {
    'EL': {'emitidas': 15, 'solped': 0, 'ofertas': 0, 'at': 7, 'oc': 8},
    'IN': {'emitidas': 45, 'solped': 0, 'ofertas': 1, 'at': 28, 'oc': 16},
}

# Evolución (Cuadro resumen) – IN+EL por etapa
#   cortes: 240726 → 300726 → 070826 → 140826 → 210826
TREND_CORTES = ['24/07', '30/07', '07/08', '14/08', '21/08']
TREND = {
    'EL': {'solped':   [0, 0, 0, 0, 0],
           'ofertas':  [3, 3, 0, 0, 0],
           'at':       [4, 4, 7, 7, 7],
           'oc':       [8, 8, 8, 8, 8],
           'emitidas': [15, 15, 15, 15, 15]},
    'IN': {'solped':   [0, 1, 0, 0, 0],
           'ofertas':  [4, 4, 2, 1, 2],
           'at':       [30, 30, 33, 28, 28],
           'oc':       [10, 10, 10, 16, 16],
           'emitidas': [44, 45, 45, 45, 46]},
}

# Ítems críticos con cadena RI→OC (hoja "Suministros críticos" – plan 210826)
# fechas: ri, solped, recof, at_cierre, necoc, oc_kom (None=no alcanzado)
CRITICOS = [
    # esp, criticidad, descripcion, ri, solped, recof, at_cierre, necoc, oc_kom, oc_num, proveedor, estado
    ('EL','HITO 2','Shelter Eléctrico SE#4 c/equipamiento',
     date(2026,1,9), date(2026,1,21), date(2026,2,9), date(2026,3,13), date(2026,5,20), date(2026,5,18),
     '4508944971','ABB','OC COLOCADA – KOM 18/5 – 129 d RI→OC. Fabricación avanza: Certificado n°2 emitido'),
    ('EL','HITO 2','Shelter Eléctrico SE#3 c/equipamiento',
     date(2026,1,9), date(2026,1,21), date(2026,2,9), date(2026,3,13), date(2026,5,20), date(2026,5,18),
     '4508944973','ABB','OC COLOCADA – KOM 18/5. Certificado n°2 emitido. OCA de USD 125k APROBADA – pendiente de emisión de OC'),
    ('EL','LLI','Sistema PMS (Power Management System)',
     date(2026,1,9), date(2026,1,21), date(2026,2,9), date(2026,3,13), date(2026,5,29), date(2026,5,22),
     '4508945953','ABB','OC COLOCADA – KOM 22/5 – 133 d RI→OC. OC (cambio de alcance) en proceso de emisión, incluye estudios en la provisión'),
    ('IN','LLI','Sistema de Control PCS',
     date(2026,2,3), date(2026,2,19), date(2026,3,25), date(2026,6,5), date(2026,8,19), date(2026,7,23),
     '4509023773','Inauco','OC (materiales) EMITIDA 4509023773 (23/07) – 170 d RI→OC. LD en confección. Pendiente enviar al proveedor info crítica (P&IDs, listado de señales, diagramas de JB)'),
    ('IN','LLI','Sistema de Seguridad SIS',
     date(2026,2,3), date(2026,2,19), date(2026,3,11), date(2026,6,5), date(2026,6,21), date(2026,7,3),
     '4509010341','HIMA','OC EMITIDA 4509010341 (03/07) – 150 d RI→OC. OC aceptada, proforma recibida, se gestiona SBL. KOM coordinada (entrega dic-26)'),
    ('IN','LLI','Válvulas de Control y Autorreguladoras',
     date(2026,4,23), date(2026,5,5), date(2026,5,21), None, date(2026,7,26), None,
     '','—','AT cerrado (válvulas de control) – cierre comercial. PENDIENTE revisión final de ofertas por ING con los comentarios de PP. Autorreguladoras: RECOTIZAR por circular ING. 120 d sin OC. NecOC 26/07 vencida'),
    ('IN','MONTO','Cables de Instrumentación',
     date(2026,5,15), date(2026,5,30), date(2026,6,23), date(2026,8,12), date(2026,10,21), None,
     '','—','AT cerrado – en proceso de presentar el LEGAJO para aprobación del cliente. 98 d desde RI. NecOC: 21/10/26'),
    ('EL','MONTO','Cables Eléctricos',
     date(2026,6,11), date(2026,7,1), date(2026,7,22), date(2026,8,4), date(2026,11,18), None,
     '','—','LEGAJO APROBADO por el cliente – OC EN PROCESO DE EMISIÓN (próxima OC). Aprovecha sobrantes de CPF 1.5. 71 d desde RI. NecOC: 18/11/26'),
]

# Demora promedio por segmento (cadena completa de los 3 adjudicados ABB)
SEGMENTOS = [
    ('RI → SOLPED',        12),
    ('SOLPED → Rec.Of.',   19),
    ('Rec.Of. → AT cierre', 32),
    ('AT cierre → OC',     66),
]


# ═══════════════════════════════════════════════════════════════════════════
# HOJA 1 – RESUMEN GERENCIAL
# ═══════════════════════════════════════════════════════════════════════════

def build_resumen(wb):
    ws = wb.create_sheet('RESUMEN GERENCIAL')
    ws.sheet_view.showGridLines = False
    for col, w in zip('ABCDEFGHIJKLMN',
                      [3,20,12,12,12,12,12,12,3,18,12,12,12,12]):
        ws.column_dimensions[col].width = w

    # Título
    ws.merge_cells('B2:M2')
    t = ws['B2']
    t.value = 'ANÁLISIS DE TIEMPOS RI → ORDEN DE COMPRA  |  La Calera II CPF2'
    t.fill = F(C['azul_h']); t.font = ft(True, C['blanco'], 16)
    t.alignment = al('center'); ws.row_dimensions[2].height = 32
    ws.merge_cells('B3:M3')
    s = ws['B3']
    s.value = (f'Especialidades: Instrumentación (IN) y Electricidad (EL)   ·   '
               f'Fuente: Plan 210826 (Rev11)   ·   Corte: 21/08/2026   ·   Evolución (24/07→21/08)   ·   RFSU: 03/02/2027')
    s.fill = F(C['azul_m']); s.font = ft(False, C['blanco'], 10)
    s.alignment = al('center'); ws.row_dimensions[3].height = 20

    # ── KPI cards ────────────────────────────────────────────────────────────
    tot_ri   = PIPELINE['EL']['cant_ri'] + PIPELINE['IN']['cant_ri']
    tot_oc   = PIPELINE['EL']['oc'] + PIPELINE['IN']['oc']
    tot_gest = PIPELINE['EL']['emitidas'] + PIPELINE['IN']['emitidas']
    tot_sin  = tot_ri - tot_gest

    kpis = [
        ('RIs TOTALES\nIN + EL',        tot_ri,           C['kpi1']),
        ('OCs COLOCADAS',               tot_oc,           C['kpi2']),
        ('EN GESTIÓN\nDE COMPRA',       tot_gest,         C['kpi3']),
        ('RIs SIN\nEMITIR',             tot_sin,          C['kpi4']),
    ]
    r = 5
    col = 2
    for label, val, color in kpis:
        ws.merge_cells(start_row=r, start_column=col, end_row=r, end_column=col+2)
        cl = ws.cell(r, col, label)
        cl.fill = F(color); cl.font = ft(True, C['blanco'], 10)
        cl.alignment = al('center', wrap=True); cl.border = bd_med()
        ws.merge_cells(start_row=r+1, start_column=col, end_row=r+2, end_column=col+2)
        cv = ws.cell(r+1, col, val)
        cv.fill = F(color); cv.font = ft(True, C['blanco'], 30)
        cv.alignment = al('center'); cv.border = bd_med()
        col += 3
    ws.row_dimensions[r].height = 30
    ws.row_dimensions[r+1].height = 26
    ws.row_dimensions[r+2].height = 26

    # ── Tabla pipeline por especialidad ──────────────────────────────────────
    r = 10
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    th = ws.cell(r, 2, 'EMBUDO DE GESTIÓN DE COMPRA POR ESPECIALIDAD')
    th.fill = F(C['azul_m']); th.font = ft(True, C['blanco'], 12)
    th.alignment = al('center'); ws.row_dimensions[r].height = 22
    r += 1

    headers = ['Especialidad', 'RIs\nTotales', 'En\nGestión', 'En\nSOLPED',
               'En Petición\nOfertas', 'En Análisis\nTécnico (AT)', 'Con OC\nColocada']
    for ci, h in enumerate(headers, start=2):
        c = ws.cell(r, ci, h)
        c.fill = F(C['gris']); c.font = ft(True, C['blanco'], 9)
        c.alignment = al('center', wrap=True); c.border = bd()
    ws.row_dimensions[r].height = 32
    r += 1

    for esp in ['EL', 'IN']:
        p = PIPELINE[esp]
        base = C['el_cl'] if esp == 'EL' else C['in_cl']
        hdrc = C['el'] if esp == 'EL' else C['in']
        vals = [p['nombre'], p['cant_ri'], p['emitidas'], p['solped'],
                p['ofertas'], p['at'], p['oc']]
        for ci, v in enumerate(vals, start=2):
            c = ws.cell(r, ci, v)
            if ci == 2:
                c.fill = F(hdrc); c.font = ft(True, C['blanco'], 10); c.alignment = al('left')
            else:
                c.fill = F(base); c.font = ft(ci == 8, C['gris'], 11); c.alignment = al('center')
            c.border = bd()
            # Resaltar el cuello de botella (AT) y OC
            if ci == 7 and v >= 20:
                c.fill = F(C['crit_cl']); c.font = ft(True, C['crit'], 11)
            if ci == 8 and v >= 5:
                c.fill = F(C['ok_cl']); c.font = ft(True, C['ok'], 11)
        ws.row_dimensions[r].height = 22
        r += 1

    # ── Gráfico de barras pipeline ───────────────────────────────────────────
    data_top = 11   # fila de headers
    chart = BarChart()
    chart.type = 'col'; chart.style = 10
    chart.title = 'Estado del circuito de compras por especialidad'
    chart.y_axis.title = 'Cantidad de RIs'
    chart.height = 8; chart.width = 20
    data = Reference(ws, min_col=4, max_col=8, min_row=data_top, max_row=data_top + 2)
    cats = Reference(ws, min_col=2, max_col=2, min_row=data_top + 1, max_row=data_top + 2)
    chart.add_data(data, titles_from_data=True, from_rows=True)
    chart.set_categories(cats)
    chart.x_axis.delete = False; chart.y_axis.delete = False
    ws.add_chart(chart, 'B16')

    # ── Hallazgos clave ──────────────────────────────────────────────────────
    r = 34
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    th = ws.cell(r, 2, 'HALLAZGOS CLAVE PARA LA DECISIÓN')
    th.fill = F(C['azul_m']); th.font = ft(True, C['blanco'], 12)
    th.alignment = al('center'); ws.row_dimensions[r].height = 22
    r += 1

    hallazgos = [
        ('[+]', 'PRÓXIMA OC – Cables Eléctricos: legajo APROBADO por el cliente, OC en emisión',
         'Cables Electricos: el LEGAJO fue APROBADO por el cliente y la OC esta EN PROCESO DE EMISION (proxima OC a colocar). '
         'Aprovecha sobrantes de CPF 1.5 (menor costo). 71 d desde RI, NecOC 18/11. '
         'Cables Instrumentacion: AT cerrado, en proceso de presentar el legajo para aprobacion del cliente (98 d, NecOC 21/10).'),
        ('[+]', 'AVANCE IN: +1 RI emitida (49 → 50) · OCs IN+EL estables en 24',
         'Instrumentacion sumo 1 nueva RI al circuito (49 → 50 RIs; emitidas 45 → 46). '
         'Las OCs IN+EL se mantienen en 24 (sin nuevas esta semana). KPIs: 69 RIs, 61 en gestion, 8 sin emitir. '
         'SIS: OC aceptada, KOM coordinada. PCS: OC de materiales emitida; pendiente enviar diagramas de JB al proveedor.'),
        ('[!]', 'PUNTO CRÍTICO – Válvulas de control: pendiente revisión final de ING',
         'Sigue el bloqueo: el AT esta cerrado pero PENDIENTE la revision final de ofertas por parte de ING, incorporando los comentarios de PP. '
         'Sin ese cierre no se coloca la OC (120 d desde RI, NecOC 26/07 ya vencida). '
         'Valvulas autorreguladoras: siguen pendientes de RECOTIZACION por la circular de ING. Escalar para no perder mas plazo.'),
        ('[i]', 'Shelters en fabricación · OCs modificatorias en emisión',
         'Shelters SE#3/SE#4 (ABB): Certificado n°2 emitido (fabricacion en marcha). '
         'SE#3: OCA USD 125k aprobada, pendiente de emision de OC. PMS: OC por cambio de alcance (estudios) en emision. '
         'La cadena RI->OC de los adjudicados ABB se mantiene en ~129-133 dias.'),
    ]
    for icon, titulo, texto in hallazgos:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=2)
        ic = ws.cell(r, 2, icon); ic.font = ft(False, s=14); ic.alignment = al('center', 'top')
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=8)
        tx = ws.cell(r, 3, '')
        tx.value = f'{titulo}\n{texto}'
        tx.fill = F(C['gris_cl']); tx.font = ft(False, C['gris'], 9)
        tx.alignment = al('left', 'center', wrap=True); tx.border = bd()
        # Negrita del título via rich text no soportado simple; dejamos título en línea
        ws.row_dimensions[r].height = 46
        r += 1

    # ── EVOLUCIÓN (Cuadro resumen) ───────────────────────────────────────────
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    th = ws.cell(r, 2, 'EVOLUCIÓN 24/07 → 21/08 – PIPELINE POR ESPECIALIDAD (Cuadro resumen)')
    th.fill = F(C['azul_m']); th.font = ft(True, C['blanco'], 12)
    th.alignment = al('center'); ws.row_dimensions[r].height = 22
    r += 1

    trend_hdr = ['Especialidad · Etapa'] + TREND_CORTES + ['Δ período']
    for ci, h in enumerate(trend_hdr, start=2):
        c = ws.cell(r, ci, h); c.fill = F(C['gris']); c.font = ft(True, C['blanco'], 9)
        c.alignment = al('center', wrap=True); c.border = bd()
    ws.row_dimensions[r].height = 22
    r += 1

    ETAPAS = [('solped', 'En SOLPED'), ('ofertas', 'En Petición Ofertas'),
              ('at', 'En AT'), ('oc', 'Con OC'), ('emitidas', 'En gestión (emitidas)')]
    for esp in ['EL', 'IN']:
        hdrc = C['el'] if esp == 'EL' else C['in']
        base = C['el_cl'] if esp == 'EL' else C['in_cl']
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
        ce = ws.cell(r, 2, PIPELINE[esp]['nombre']); ce.fill = F(hdrc)
        ce.font = ft(True, C['blanco'], 10); ce.alignment = al('left'); ce.border = bd()
        ws.row_dimensions[r].height = 18; r += 1
        for key, label in ETAPAS:
            serie = TREND[esp][key]
            lc = ws.cell(r, 2, label); lc.fill = F(base); lc.font = ft(False, C['gris'], 9)
            lc.alignment = al('left'); lc.border = bd()
            for ci, val in enumerate(serie, start=3):
                c = ws.cell(r, ci, val); c.fill = F(C['blanco']); c.font = ft(False, C['gris'], 9)
                c.alignment = al('center'); c.border = bd()
            d = serie[-1] - serie[0]
            dc = ws.cell(r, 7)
            if d > 0:
                dc.value = f'+{d}'; dc.fill = F(C['ok_cl']); dc.font = ft(True, C['ok'], 10)
            elif d < 0:
                dc.value = f'{d}'; dc.fill = F(C['azul_cl']); dc.font = ft(True, C['azul_h'], 10)
            else:
                dc.value = '='; dc.fill = F(C['blanco']); dc.font = ft(False, C['gris'], 10)
            dc.alignment = al('center'); dc.border = bd()
            ws.row_dimensions[r].height = 17; r += 1

    # Resumen IN+EL de las 3 preguntas clave (período 26/06 → 24/07)
    r += 1
    def t_ini(k): return TREND['EL'][k][0] + TREND['IN'][k][0]
    def t_fin(k): return TREND['EL'][k][-1] + TREND['IN'][k][-1]
    preguntas = [
        ('Última semana (14/08 → 21/08)',
         'IN sumó +1 RI (49→50); emitidas 45→46. OCs IN+EL sin cambio (24). '
         'Cables EL: legajo aprobado por el cliente, OC en emisión (próxima OC). Cables IN: legajo al cliente.'),
        ('OCs colocadas en el período (IN+EL)',
         f'{t_ini("oc")} → {t_fin("oc")}: +{t_fin("oc")-t_ini("oc")} OCs entre 24/07 y 21/08 '
         f'(IN 10→16, EL 8=8). Próxima OC en camino: Cables Eléctricos (legajo aprobado). '
         f'Total OCs colocadas IN+EL: {t_fin("oc")}.'),
        ('Puntos críticos / pendientes',
         f'En petición de ofertas: EL {PIPELINE["EL"]["ofertas"]} · IN {PIPELINE["IN"]["ofertas"]} '
         f'= {t_fin("ofertas")} ítems. VÁLVULAS de control: pendiente revisión final de ofertas por ING (con comentarios de PP) – sin OC (120 d). '
         f'Autorreguladoras: a recotizar. PCS: pendiente diagramas de JB.'),
    ]
    for titulo, texto in preguntas:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        ct = ws.cell(r, 2, titulo); ct.fill = F(C['azul_cl']); ct.font = ft(True, C['azul_h'], 9)
        ct.alignment = al('left', 'center', wrap=True); ct.border = bd()
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=8)
        cx = ws.cell(r, 4, texto); cx.fill = F(C['gris_cl']); cx.font = ft(False, C['gris'], 9)
        cx.alignment = al('left', 'center', wrap=True); cx.border = bd()
        ws.row_dimensions[r].height = 34
        r += 1

    ws.sheet_view.zoomScale = 100
    return ws


# ═══════════════════════════════════════════════════════════════════════════
# HOJA 2 – ANÁLISIS RI → OC (DEMORAS)
# ═══════════════════════════════════════════════════════════════════════════

def color_demora(d, umbral_ok=90, umbral_warn=130):
    if d is None: return C['gris_cl'], C['gris']
    if d <= umbral_ok:   return C['ok_cl'], C['ok']
    if d <= umbral_warn: return C['warn_cl'], C['warn']
    return C['crit_cl'], C['crit']

def build_analisis(wb):
    ws = wb.create_sheet('ANÁLISIS RI→OC')
    ws.sheet_view.showGridLines = False
    widths = [3,7,10,34,11,11,11,11,11,11,13,13,30]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.merge_cells('B2:M2')
    t = ws['B2']
    t.value = 'ANÁLISIS DETALLADO RI → OC  ·  Ítems Críticos IN / EL'
    t.fill = F(C['azul_h']); t.font = ft(True, C['blanco'], 15)
    t.alignment = al('center'); ws.row_dimensions[2].height = 30
    ws.merge_cells('B3:M3')
    s = ws['B3']
    s.value = ('Tiempo desde emisión de RI hasta OC efectiva (adjudicados) o días transcurridos sin OC (en gestión).  '
               'OC efectiva tomada en fecha KOM para los paquetes adjudicados.')
    s.fill = F(C['azul_m']); s.font = ft(False, C['blanco'], 9)
    s.alignment = al('center', wrap=True); ws.row_dimensions[3].height = 24

    # Encabezados
    r = 5
    headers = ['', 'Esp.', 'Criticidad', 'Suministro', 'RI\n(emisión)', 'SOLPED', 'Rec.\nOfertas',
               'AT\ncierre', 'OC\nefectiva', 'Nec.\nOC', 'RI→OC\n(días)', 'Estado\nactual', 'Observaciones / Acción']
    for ci, h in enumerate(headers, start=1):
        if ci == 1: continue
        c = ws.cell(r, ci, h)
        c.fill = F(C['gris']); c.font = ft(True, C['blanco'], 9)
        c.alignment = al('center', wrap=True); c.border = bd()
    ws.row_dimensions[r].height = 34
    r += 1

    def fd(d): return d.strftime('%d/%m/%y') if isinstance(d, date) else '—'

    for (esp, crit, desc, ri, solped, recof, atc, necoc, ockom, ocn, prov, estado) in CRITICOS:
        espc = C['el'] if esp == 'EL' else C['in']
        base = C['el_cl'] if esp == 'EL' else C['in_cl']

        # cálculo días
        if ockom:
            d_total = dias(ri, ockom); adjud = True
        elif isinstance(ri, date) and ri <= TODAY:
            d_total = dias(ri, TODAY); adjud = False
        else:
            d_total = None; adjud = False   # RI futura

        ws.cell(r, 2, esp).fill = F(espc)
        ws.cell(r, 2).font = ft(True, C['blanco'], 9); ws.cell(r, 2).alignment = al('center'); ws.cell(r, 2).border = bd()
        ws.cell(r, 3, crit).fill = F(base)
        ws.cell(r, 3).font = ft(True, C['gris'], 8); ws.cell(r, 3).alignment = al('center'); ws.cell(r, 3).border = bd()
        cd = ws.cell(r, 4, desc); cd.fill = F(base); cd.font = ft(False, C['gris'], 9)
        cd.alignment = al('left', wrap=True); cd.border = bd()

        for ci, dval in [(5, ri), (6, solped), (7, recof), (8, atc), (9, ockom), (10, necoc)]:
            c = ws.cell(r, ci, fd(dval))
            c.fill = F(C['blanco']); c.font = ft(False, C['gris'], 9)
            c.alignment = al('center'); c.border = bd()
            if ci == 9 and ockom:
                c.fill = F(C['ok_cl']); c.font = ft(True, C['ok'], 9)

        # RI→OC días
        bgc, fgc = color_demora(d_total)
        if d_total is None:
            txt = 'RI futura'
        elif adjud:
            txt = f'{d_total}'
        else:
            txt = f'{d_total} *'
        cdias = ws.cell(r, 11, txt)
        cdias.fill = F(bgc); cdias.font = ft(True, fgc, 11)
        cdias.alignment = al('center'); cdias.border = bd()

        # Estado
        if adjud:
            est_bg, est_fg, est_txt = C['ok'], C['blanco'], 'OC COLOCADA'
        elif atc or 'AT' in estado:
            est_bg, est_fg, est_txt = C['warn'], C['blanco'], 'EN AT'
        elif d_total is None:
            est_bg, est_fg, est_txt = C['gris_cl'], C['gris'], 'RI PROGRAM.'
        else:
            est_bg, est_fg, est_txt = C['azul_m'], C['blanco'], 'EN GESTIÓN'
        ce = ws.cell(r, 12, est_txt); ce.fill = F(est_bg); ce.font = ft(True, est_fg, 8)
        ce.alignment = al('center', wrap=True); ce.border = bd()

        co = ws.cell(r, 13, estado); co.fill = F(base); co.font = ft(False, C['gris'], 8)
        co.alignment = al('left', 'center', wrap=True); co.border = bd()

        ws.row_dimensions[r].height = 30
        r += 1

    # nota
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=13)
    n = ws.cell(r, 2, '(*) Para ítems sin OC, el valor es la cantidad de días transcurridos desde la RI hasta el corte (21/08/26), sin orden de compra colocada todavía.')
    n.font = ft(False, C['gris'], 8); n.alignment = al('left'); ws.row_dimensions[r].height = 16

    # ── Tabla auxiliar para gráfico días RI→OC ───────────────────────────────
    r += 2
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=8)
    th = ws.cell(r, 2, 'TIEMPO RI → OC POR ÍTEM (días)')
    th.fill = F(C['azul_m']); th.font = ft(True, C['blanco'], 11); th.alignment = al('center')
    ws.row_dimensions[r].height = 20
    r += 1
    gtop = r
    ws.cell(r, 2, 'Ítem'); ws.cell(r, 3, 'Días RI→OC / transcurridos')
    ws.cell(r, 2).font = ft(True, s=9); ws.cell(r, 3).font = ft(True, s=9)
    ws.cell(r, 2).fill = F(C['gris_cl']); ws.cell(r, 3).fill = F(C['gris_cl'])
    r += 1
    chart_rows = []
    for (esp, crit, desc, ri, solped, recof, atc, necoc, ockom, ocn, prov, estado) in CRITICOS:
        if ockom: d = dias(ri, ockom)
        elif isinstance(ri, date) and ri <= TODAY: d = dias(ri, TODAY)
        else: continue
        short = (desc[:26] + '…') if len(desc) > 27 else desc
        ws.cell(r, 2, f'{esp} · {short}').font = ft(False, s=9)
        ws.cell(r, 2).alignment = al('left')
        ws.cell(r, 3, d).font = ft(False, s=9); ws.cell(r, 3).alignment = al('center')
        chart_rows.append(r)
        r += 1
    gbot = r - 1

    chart = BarChart()
    chart.type = 'bar'; chart.style = 12
    chart.title = 'Días RI → OC (adjudicados) / días transcurridos (en gestión)'
    chart.x_axis.title = 'días'; chart.height = 8; chart.width = 22
    data = Reference(ws, min_col=3, max_col=3, min_row=gtop, max_row=gbot)
    cats = Reference(ws, min_col=2, max_col=2, min_row=gtop + 1, max_row=gbot)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    ws.add_chart(chart, 'E' + str(gtop))

    return ws


# ═══════════════════════════════════════════════════════════════════════════
# HOJA 3 – DÓNDE SE PIERDE EL TIEMPO (segmentos) + METODOLOGÍA
# ═══════════════════════════════════════════════════════════════════════════

def build_segmentos(wb):
    ws = wb.create_sheet('DÓNDE DEMORA + FUENTES')
    ws.sheet_view.showGridLines = False
    for col, w in zip('ABCDEFGH', [3,26,14,14,40,3,3,3]):
        ws.column_dimensions[col].width = w

    ws.merge_cells('B2:E2')
    t = ws['B2']
    t.value = '¿DÓNDE SE CONCENTRA LA DEMORA DEL CIRCUITO?'
    t.fill = F(C['azul_h']); t.font = ft(True, C['blanco'], 14)
    t.alignment = al('center'); ws.row_dimensions[2].height = 28
    ws.merge_cells('B3:E3')
    s = ws['B3']
    s.value = 'Descomposición del tiempo RI → OC en los 3 paquetes ya adjudicados (ABB) – cadena completa de 129 días'
    s.fill = F(C['azul_m']); s.font = ft(False, C['blanco'], 9)
    s.alignment = al('center', wrap=True); ws.row_dimensions[3].height = 22

    r = 5
    for ci, h in enumerate(['Tramo del circuito', 'Días', '% del total', 'Lectura'], start=2):
        c = ws.cell(r, ci, h); c.fill = F(C['gris']); c.font = ft(True, C['blanco'], 10)
        c.alignment = al('center'); c.border = bd()
    ws.row_dimensions[r].height = 20
    r += 1

    total = sum(d for _, d in SEGMENTOS)
    lecturas = {
        'RI → SOLPED':         'Ágil – generación de solicitud de pedido en SAP.',
        'SOLPED → Rec.Of.':    'Normal – tiempo de mercado para recibir ofertas.',
        'Rec.Of. → AT cierre': 'Análisis técnico; aceptable para equipos complejos.',
        'AT cierre → OC':      'CRÍTICO – el tramo más largo. Gestión comercial/aprobaciones post-AT.',
    }
    gtop = r
    for nombre, d in SEGMENTOS:
        pct = d / total
        es_crit = nombre.startswith('AT cierre')
        base = C['crit_cl'] if es_crit else C['azul_cl']
        c1 = ws.cell(r, 2, nombre); c1.fill = F(base); c1.font = ft(es_crit, C['gris'], 10)
        c1.alignment = al('left'); c1.border = bd()
        c2 = ws.cell(r, 3, d); c2.fill = F(base); c2.font = ft(True, C['crit'] if es_crit else C['gris'], 11)
        c2.alignment = al('center'); c2.border = bd()
        c3 = ws.cell(r, 4, pct); c3.number_format = '0%'; c3.fill = F(base)
        c3.font = ft(es_crit, C['gris'], 10); c3.alignment = al('center'); c3.border = bd()
        c4 = ws.cell(r, 5, lecturas[nombre]); c4.fill = F(base); c4.font = ft(False, C['gris'], 8)
        c4.alignment = al('left', wrap=True); c4.border = bd()
        ws.row_dimensions[r].height = 26
        r += 1
    gbot = r - 1
    # fila total
    ct = ws.cell(r, 2, 'TOTAL RI → OC'); ct.fill = F(C['azul_h']); ct.font = ft(True, C['blanco'], 10)
    ct.alignment = al('left'); ct.border = bd()
    cv = ws.cell(r, 3, total); cv.fill = F(C['azul_h']); cv.font = ft(True, C['blanco'], 11)
    cv.alignment = al('center'); cv.border = bd()
    cp = ws.cell(r, 4, 1.0); cp.number_format = '0%'; cp.fill = F(C['azul_h']); cp.font = ft(True, C['blanco'], 10)
    cp.alignment = al('center'); cp.border = bd()
    ws.cell(r, 5, '≈ 4,3 meses entre RI y OC').fill = F(C['azul_h'])
    ws.cell(r, 5).font = ft(True, C['blanco'], 9); ws.cell(r, 5).alignment = al('left'); ws.cell(r, 5).border = bd()

    # gráfico de segmentos
    chart = BarChart(); chart.type = 'bar'; chart.style = 11
    chart.title = 'Días por tramo del circuito RI → OC'
    chart.x_axis.title = 'días'; chart.height = 7; chart.width = 18
    data = Reference(ws, min_col=3, max_col=3, min_row=gtop - 1, max_row=gbot)
    cats = Reference(ws, min_col=2, max_col=2, min_row=gtop, max_row=gbot)
    chart.add_data(data, titles_from_data=True); chart.set_categories(cats); chart.legend = None
    ws.add_chart(chart, 'B12')

    # ── Metodología / fuentes ────────────────────────────────────────────────
    r = 28
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=5)
    th = ws.cell(r, 2, 'METODOLOGÍA Y FUENTES DE DATOS')
    th.fill = F(C['azul_m']); th.font = ft(True, C['blanco'], 11); th.alignment = al('center')
    ws.row_dimensions[r].height = 20
    r += 1
    notas = [
        ('Fuente', 'Plan de Suministros – La Calera II (210826).xlsx · Rev11'),
        ('Hojas usadas', 'Cuadro resumen (pipeline) · Suministros críticos (cadena por ítem) · RI y OC x mes.'),
        ('Fecha de corte', '21/08/2026 (evolución: 24/07 · 30/07 · 07/08 · 14/08 · 21/08)'),
        ('Universo', 'IN: 48 RIs · EL: 20 RIs (total 68). Cadena RI→OC detallada disponible para los 8 ítems críticos.'),
        ('Definición OC efectiva', 'Para paquetes adjudicados sin fecha formal de OC, se usa la fecha KOM como hito de OC efectiva.'),
        ('Días en gestión', 'Para ítems sin OC: días entre RI y fecha de corte (proceso aún abierto).'),
        ('Limitación', 'El Plan registra fechas reales de cadena solo para ítems críticos; el resto se reporta a nivel de pipeline agregado (Cuadro resumen).'),
    ]
    for lbl, val in notas:
        cl = ws.cell(r, 2, lbl); cl.fill = F(C['azul_cl']); cl.font = ft(True, C['azul_h'], 9)
        cl.alignment = al('left', 'center'); cl.border = bd()
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=5)
        cv = ws.cell(r, 3, val); cv.font = ft(False, C['gris'], 9)
        cv.alignment = al('left', 'center', wrap=True); cv.border = bd()
        ws.row_dimensions[r].height = 28
        r += 1

    return ws


def main():
    print('Generando Dashboard Gerencial RI→OC...')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    build_resumen(wb)
    build_analisis(wb)
    build_segmentos(wb)
    wb.save(OUT_XLSX)
    print(f'OK → {os.path.basename(OUT_XLSX)} ({os.path.getsize(OUT_XLSX)//1024} KB)')

if __name__ == '__main__':
    main()
