#!/usr/bin/env python3
"""
Tracker Suministros v3R1 – La Calera II CPF2
Especialidades IN (Instrumentación & Control) | EL (Electricidad)
Lee directamente del Plan de Suministros.
v3: Versión inicial con todas las requisiciones individuales, análisis demoras RI→OC,
    AT CT1/CT0, números SOLPED, items SC adicionales.
v3R1: Corrección de fechas – se reemplaza fechas P0 Preliminar (2024) en columnas
      "Prog. Actual" por fechas del Crono 4.11 (2025-2026).
"""

import os, sys, re
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import date, datetime

# ── Rutas ───────────────────────────────────────────────────────────────────
import glob as _glob
SCRIPT_DIR  = os.path.dirname(os.path.abspath(__file__))
PLANS_DIR   = os.path.join(SCRIPT_DIR, '..', 'info_suministros')

def find_plan(tag):
    """Resuelve el archivo de plan por su tag ddmmyy (tolerante al prefijo del nombre)."""
    hits = [f for f in _glob.glob(os.path.join(PLANS_DIR, '*.xlsx')) if f'({tag})' in os.path.basename(f)]
    if not hits:
        raise FileNotFoundError(f'No se encontró plan con tag ({tag}) en {PLANS_DIR}')
    return hits[0]

PLAN_FILE     = find_plan('280826')          # corte actual
PLAN_FILE_OLD = find_plan('210826')          # corte anterior (para diff de críticos)
# Cortes para la evolución (rolling window de 5) — se muestran como tendencia
PLAN_TREND    = [
    ('30/07', find_plan('300726')),
    ('07/08', find_plan('070826')),
    ('14/08', find_plan('140826')),
    ('21/08', find_plan('210826')),
    ('28/08', find_plan('280826')),
]
OUT_XLSX      = os.path.join(SCRIPT_DIR, 'Tracker_Suministros_IN_EL_LaCalera_II_Rev12_280826.xlsx')
TODAY         = date(2026, 8, 28)
RFSU          = date(2027, 2, 3)
VERSION       = 'Rev12_280826'

# ── Paleta ──────────────────────────────────────────────────────────────────
C = {
    'titulo':     '1F3864', 'subtitulo':  '2F5496',
    'hdr_in':     '843C0C', 'hdr_el':     '375623',
    'hdr_grp':    '2E4057', 'hdr_col':    'B8CCE4',
    'completado': '70AD47', 'en_proceso': 'FFD966',
    'atrasado':   'FF0000', 'pendiente':  'F2F2F2',
    'no_aplica':  '4472C4', 'en_transito':'00B0F0',
    'emitida':    'FFC000', 'adjudicado': '92D050',
    'sin_ri':     'FF4500',
    'row_alt':    'EBF3FB', 'row_norm':   'FFFFFF',
    'llt':        'FF7F27', 'crit':       'FF0000',
    'monto':      'FFC000', 'hito2':      '4472C4',
    'media':      'E2EFDA', 'plan_p0':    'FFF2CC',
    'demora_ok':  'C6EFCE', 'demora_warn':'FFEB9C', 'demora_crit':'FFC7CE',
    'border':     '8EA9C1',
}

STATUS_MAP = {
    'COMPLETADO':     (C['completado'],  'FFFFFF', True),
    'COMPLETADA':     (C['completado'],  'FFFFFF', True),
    'EN PROCESO':     (C['en_proceso'],  '000000', True),
    'EMITIDA':        (C['emitida'],     '000000', True),
    'EN LIBERACIÓN':  ('FF8C00',         'FFFFFF', True),   # ámbar oscuro = SOLPED en trámite
    'ADJUDICADO':     (C['adjudicado'],  'FFFFFF', True),
    'EN TRÁNSITO':    (C['en_transito'], 'FFFFFF', True),
    'ATRASADO':       (C['atrasado'],    'FFFFFF', True),
    'PENDIENTE':      (C['pendiente'],   '000000', False),
    'NO APLICA':      (C['no_aplica'],   'FFFFFF', True),
    'SIN RI':      (C['sin_ri'],      'FFFFFF', True),
}

CRIT_COLOR = {
    'LLI': C['llt'], 'C. CRÍTICO': C['crit'], 'MONTO': C['monto'],
    'HITO 2': C['hito2'], 'MEDIA': C['media'],
}

def F(col): return PatternFill('solid', fgColor=col)
def ft(bold=False, color='000000', sz=9): return Font(bold=bold, color=color, size=sz, name='Calibri')
def al(h='center', v='center', wrap=False): return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def bd(color=C['border']):
    s = Side(style='thin', color=color); return Border(left=s, right=s, top=s, bottom=s)
def bd_med():
    s = Side(style='medium', color='000000'); return Border(left=s, right=s, top=s, bottom=s)

def fmt_date(v):
    if v is None: return ''
    if isinstance(v, (date, datetime)): return v.strftime('%d/%m/%y')
    return str(v)[:10]

def to_date(v):
    if v is None: return None
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date): return v
    return None

def safe_date(v):
    """Convert to date, filtering bogus 1900-era Excel dates."""
    d = to_date(v)
    return d if (d and d.year > 1900) else None

def days_between(d1, d2):
    a, b = to_date(d1), to_date(d2)
    if a and b: return (b - a).days
    return None

def apply_status_cell(ws, row, col, text):
    cell = ws.cell(row=row, column=col, value=text)
    cfg = STATUS_MAP.get(text.upper() if text else '', None)
    if cfg:
        cell.fill = F(cfg[0]); cell.font = ft(bold=cfg[2], color=cfg[1], sz=9)
    else:
        cell.fill = F(C['pendiente']); cell.font = ft(sz=9)
    cell.alignment = al('center', 'center')
    cell.border = bd()
    return cell

def style_delay_cell(ws, row, col, days, threshold_warn=15, threshold_crit=30):
    cell = ws.cell(row=row, column=col, value=days if days is not None else '')
    if days is not None:
        if days <= threshold_warn:
            cell.fill = F(C['demora_ok'])
        elif days <= threshold_crit:
            cell.fill = F(C['demora_warn'])
        else:
            cell.fill = F(C['demora_crit'])
    else:
        cell.fill = F(C['pendiente'])
    cell.font = ft(sz=9); cell.alignment = al('center'); cell.border = bd()
    return cell


# ── Mapeo N° RI: descripción Preliminar → RI number (Datos P0) ─────────────
IN_RI_NUMBER = {
    'Transmisor de Presión Diferencial':     '5155-00-0000-IG-IN-RI-006 / 2000-IN-RI-006',
    'Transmisor de Presión':                 '5155-00-0000-IG-IN-RI-005',
    'Transmisor de Caudal Multivariable':    '5155-00-0000-IG-IN-RI-027',
    'Transmisor de Caudal Coriolis':         '5155-00-0000-IG-IN-RI-025',
    'Transmisor de Caudal Presión Dif':      '5155-00-0000-IG-IN-RI-026',
    'Transmisor de Nivel Dp':                '5155-00-0000-IG-IN-RI-010',
    'Transmisor de Nivel Radar':             '5155-00-0000-IG-IN-RI-012',
    'Transmisor de Nivel':                   '5155-00-0000-IG-IN-RI-009',
    'Transmisor de Temperatura':             '5155-00-0000-IG-IN-RI-020',
    'Interruptor de Presion':                '5155-00-0000-IG-IN-RI-030',
    'Interruptor de Nivel':                  '5155-00-0000-IG-IN-RI-030',
    'Interruptor de Vibraci':                '5155-00-0000-IG-IN-RI-030',
    'Manómetros':                            '5155-00-0000-IG-IN-RI-011 / 2000-IN-RI-003',
    'Indicador de Nivel':                    '5155-00-0000-IG-IN-RI-009',
    'Rotámetro':                             '5155-00-0000-IG-IN-RI-007 / 2000-IN-RI-004',
    'Termómetro':                            '5155-00-0000-IG-IN-RI-020',
    'Cupón de Corros':                       '5155-00-0000-IG-IN-RI-014',
    'Plcas Orificio':                        '5155-00-0000-IG-IN-RI-004',
    'Analizador':                            '5155-00-0000-IG-IN-RI-028 / RI-029',
    'Detector de Fugas':                     '5155-00-0000-IG-IN-RI-016',
    'Detector de Mecla Explosiva':           '5155-00-0000-IG-IN-RI-017 / RI-018',
    'Pulsador de Emergencia':                '5155-00-0000-IG-IN-RI-030',
    'Indicador de Paso de Scr':              '5155-00-0000-IG-IN-RI-008',
    'Detector de Llama':                     '5155-00-0000-IG-IN-RI-017 / RI-018',
    'Válvulas de Control':                   '5155-00-0000-IG-IN-RI-001/-003/-013/-019',
    'Válvulas Autor':                        '5155-00-2000-IG-IN-RI-001 / 2000-IN-RI-007',
    'Válvulas de Seguridad':                 '5155-00-0000-IG-IN-RI-002',
    'Válvulas BlowDown':                     '5155-00-0000-IG-IN-RI-001 / RI-003',
    'Válvulas ShutDown':                     '5155-00-0000-IG-IN-RI-013 / RI-019',
    'Válvulas Diluvio':                      '5155-00-0000-IG-IN-RI-021',
    'Cables de Instrumentos':                '5155-00-0000-IG-IN-RI-032 / 2000-IN-RI-009',
    'Cajas de Conexionado':                  '5155-00-0000-IG-IN-RI-033',
    'Requerimiento para Sistema de Comunicaciones': '5155-00-0000-IG-IN-RI-017 / RI-018',
    'CCTV':                                  '5155-00-0000-IG-IN-RI-017 / RI-018',
    'Materiales Mecánicos':                  '5155-00-0000-IG-IN-RI-039',
    'Materiales Eléctricos':                 '5155-00-0000-IG-IN-RI-040',
    # Suministros críticos IN
    'SISTEMA DE CONTROL PCS':               '5155-xx-IN-PCS (en gestión)',
    'SISTEMA DE SEGURIDAD SIS':             '5155-xx-IN-SIS (en gestión)',
}

EL_RI_NUMBER = {
    'Puesta a tierra':                       '5155-00-0000-IG-EL-RI-002 / RI-005',
    'Cables - Baja y media':                 '5155-00-0000-IG-EL-RI-006',
    'Canalizaciones':                        '5155-00-0000-IG-EL-RI-007',
    'Iluminación':                           '5155-00-0000-IG-EL-RI-003',
    'Prensacables':                          '5155-00-0000-IG-EL-RI-006 (accesorio)',
    'Cajas de conexionado':                  '5155-00-0000-IG-EL-RI-001',
    'Tracing eléctrico':                     '5155-00-0000-IG-EL-RI-004',
    'Protección catódica':                   '5155-00-0000-IG-EL-RI-002',
    'Materiales Misceláneos':                '5155-00-0000-IG-EL-RI-009',
    # Suministros críticos EL
    'SHELTER ELECTRICO SE#4':               '5155-00-2300-IG-EL-RI-006',
    'SHELTER ELECTRICO SE#3':               '5155-00-2300-IG-EL-RI-005',
    'SHELTER ELECTRICO SE#5':               '5155-00-2300-IG-EL-RI-004',
    'SISTEMA PMS':                           '5155-00-2300-IG-EL-RI-007',
    'TRANSFORMADORES SECOS':                 '5155-00-2300-IG-EL-RI-001 / RI-008',
    'GENERADOR DE EMERGENCIA':               '5155-00-2300-IG-EL-RI-003',
    'DUCTO DE BARRAS':                       '5155-00-2300-IG-EL-RI-002',
    'CABLES ELÉCTRICOS':                     '5155-00-0000-IG-EL-RI-006',
}

def get_ri_number(desc, ri_map):
    d = str(desc) if desc else ''
    for key, val in ri_map.items():
        if key.lower() in d.lower():
            return val
    return ''


# ── Mapeo Preliminar → Crono 4.11 ───────────────────────────────────────────
# Lista de (keyword en desc Prelim, keyword en desc Crono) ordenada por especificidad
IN_PRELIM_TO_CRONO = [
    ('válvulas de control',    'válvulas de control, shutdown'),
    ('válvulas autor',         'válvulas de control, shutdown'),
    ('válvulas blowdown',      'válvulas de control, shutdown'),
    ('válvulas shutdown',      'válvulas de control, shutdown'),
    ('válvulas diluvio',       'válvulas de control, shutdown'),
    ('válvulas de seguridad',  'válvulas de seguridad'),
    ('transmisor de caudal',   'medidores de caudal'),
    ('plcas orificio',         'medidores de caudal'),
    ('transmisor de nivel',    'transmisores'),
    ('transmisor de presión',  'transmisores'),
    ('transmisor de temperatura', 'transmisores'),
    ('analizador',             'analizadores'),
    ('tomamuestras',           'tomamuestras'),
    ('detector de llama',      'sistema fire & gas'),
    ('detector de fugas',      'sistema fire & gas'),
    ('detector de mecla',      'sistema fire & gas'),
    ('pulsador de emergencia', 'sistema fire & gas'),
    ('cables de instrument',   'cables de instrumentación rev. 0'),
    ('cajas de conexionado',   'cajas'),
    ('requerimiento para sistema', 'sistema contra incendio'),
    ('cctv',                   'sistema contra incendio'),
    ('indicador de paso',      'instrumentos comunes'),
    ('interruptor de',         'instrumentos comunes'),
    ('manómetro',              'instrumentos comunes'),
    ('indicador de nivel',     'instrumentos comunes'),
    ('rotámetro',              'instrumentos comunes'),
    ('termómetro',             'instrumentos comunes'),
    ('cupón de corros',        'instrumentos comunes'),
    ('materiales mecán',       'materiales para montaje'),
    ('materiales eléct',       'materiales para montaje'),
    ('sistema de control pcs', 'sistema de control, hardware'),
    ('sistema de seguridad sis', 'sistema contra incendio'),
]

EL_PRELIM_TO_CRONO = [
    ('cables - baja y media',  'cables rev. 0'),
    ('canalizaciones',         'canalizaciones eléctricas'),
    ('tracing eléctrico',      'tracing'),
    ('puesta a tierra',        'materiales para pat'),
    ('protección catódica',    'materiales para pat'),
    ('sistema ups',            'sistema ups y rectificadores'),
    ('ducto de barras',        'ducto de barras'),
    ('shelter',                'shelter y banco de carga'),
    # Iluminación, prensacables, cajas conexionado EL, materiales misceláneos
    # no tienen paquete equivalente en crono 4.11
]


def find_crono_item(prelim_desc, esp, crono_items):
    """Encuentra el paquete de crono 4.11 correspondiente a una descripción de Preliminar."""
    d = prelim_desc.lower()
    mapping = IN_PRELIM_TO_CRONO if esp == 'IN' else EL_PRELIM_TO_CRONO
    for prelim_kw, crono_kw in mapping:
        if prelim_kw in d:
            for key, item in crono_items.items():
                if item['esp'] == esp and crono_kw in key:
                    return item
    return None


# ── Leer datos del Plan ──────────────────────────────────────────────────────
def load_plan():
    wb = openpyxl.load_workbook(PLAN_FILE, data_only=True)

    # ─── Suministros críticos ──────────────────────────────────────────────
    ws_sc = wb['Suministros críticos']
    sc_in, sc_el = [], []
    for row in ws_sc.iter_rows(min_row=2, values_only=True):
        if not row[0]: continue
        esp = str(row[2]) if row[2] else ''
        if esp not in ('IN', 'EL'): continue
        item = {
            'esp':      esp,
            'crit':     str(row[1]) if row[1] else '',
            'desc_sc':  str(row[3]) if row[3] else '',
            'ri_real':  to_date(row[4]),
            'solped':   to_date(row[5]),
            'recof':    to_date(row[6]),
            'at_cierre':to_date(row[7]),
            'nec_oc':   to_date(row[8]),
            'oc_real_d':to_date(row[9]),
            'lt':       row[10],
            'oc_n':     str(row[11]) if row[11] else '',
            'prov':     str(row[12]) if row[12] else '',
            'kom':      to_date(row[13]),
            'ent_oc':   to_date(row[17]),
            'nec_ent':  to_date(row[18]),
            'status':   str(row[19]) if row[19] else '',
        }
        if esp == 'IN': sc_in.append(item)
        else:           sc_el.append(item)

    # ─── Crono 4.11 (fechas programa actual 2025-2026) ────────────────────
    ws_c = wb['La Calera II - crono 4.11']
    crono_items = {}  # keyed by description lowercase
    for row in ws_c.iter_rows(min_row=4, values_only=True):
        esp_c = str(row[0]) if row[0] else ''
        if not (esp_c.startswith('IN') or esp_c.startswith('EL')): continue
        desc = str(row[2]) if row[2] else ''
        if not desc: continue
        esp_short = 'IN' if esp_c.startswith('IN') else 'EL'
        crono_items[desc.lower()] = {
            'desc':     desc,
            'esp':      esp_short,
            'ri_prog':  safe_date(row[4]),
            'sp_prog':  safe_date(row[10]),
            'rec_prog': safe_date(row[18]),
            'at0_prog': safe_date(row[29]),
            'oc_prog':  safe_date(row[36]),
        }

    # ─── Preliminar (P0 baseline 2024 – solo se usa para P0 y datos reales) ─
    ws_p = wb['Preliminar']
    in_items, el_items = [], []
    for row in ws_p.iter_rows(min_row=4, values_only=True):
        esp = str(row[0]) if row[0] else ''
        if not esp: continue
        item = {
            'desc':          str(row[2]) if row[2] else '',
            'ri_p0':         to_date(row[4]),      # P0 baseline (2024)
            'ri_real_p':     to_date(row[6]),
            'ri_desv':       row[7],
            'solp_p0':       to_date(row[10]),     # P0 baseline (2024)
            'solp_real_p':   to_date(row[13]),
            'solp_n':        str(row[15]) if row[15] else '',
            'recof_p0':      to_date(row[18]),     # P0 baseline (2024)
            'recof_real_p':  to_date(row[21]),
            'at_reva_p0':    to_date(row[24]),     # P0 baseline (2024)
            'at_reva_real':  to_date(row[25]),
            'at_rev0_p0':    to_date(row[29]),     # P0 baseline (2024)
            'at_rev0_real':  to_date(row[32]),
            'oc_p0':         to_date(row[36]),     # P0 baseline (2024)
            'oc_real_p':     to_date(row[39]),
            'oc_n':          str(row[41]) if row[41] else '',
            'prov':          str(row[43]) if row[43] else '',
            'ent_oc':        to_date(row[44]),
            'nec_obra':      to_date(row[75]),
        }
        if esp.startswith('IN'):   in_items.append(item)
        elif esp.startswith('EL'): el_items.append(item)

    return sc_in, sc_el, in_items, el_items, crono_items


def load_sc_from_file(filepath):
    """Carga únicamente la hoja Suministros críticos de un Plan de Suministros dado."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws_sc = wb['Suministros críticos']
    sc = []
    for row in ws_sc.iter_rows(min_row=2, values_only=True):
        if not row[0]: continue
        esp = str(row[2]) if row[2] else ''
        if esp not in ('IN', 'EL'): continue
        sc.append({
            'esp':       esp,
            'crit':      str(row[1]) if row[1] else '',
            'desc_sc':   str(row[3]) if row[3] else '',
            'ri_real':   safe_date(row[4]),
            'solped':    safe_date(row[5]),
            'recof':     safe_date(row[6]),
            'at_cierre': safe_date(row[7]),
            'nec_oc':    safe_date(row[8]),
            'oc_real_d': safe_date(row[9]),
            'oc_n':      str(row[11]) if row[11] else '',
            'prov':      str(row[12]) if row[12] else '',
            'kom':       safe_date(row[13]),
            'ent_oc':    safe_date(row[17]),
            'nec_ent':   safe_date(row[18]),
            'status':    str(row[19]) if row[19] else '',
        })
    return sc


def load_cuadro(filepath):
    """Carga el pipeline EL/IN de la hoja 'Cuadro resumen' de un Plan de Suministros.
    Devuelve {'EL': {...}, 'IN': {...}} con cant_ri, emitidas, solped, ofertas, at, oc."""
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb['Cuadro resumen']
    nombre_map = {'electricidad': 'EL', 'instrumentos': 'IN'}
    out = {}
    for row in ws.iter_rows(min_row=8, max_row=15, values_only=True):
        esp_nombre = str(row[1]).strip().lower() if row[1] else ''
        code = nombre_map.get(esp_nombre)
        if not code:
            continue
        out[code] = {
            'cant_ri':  row[2] or 0,
            'emitidas': row[3] or 0,
            'solped':   row[4] or 0,
            'ofertas':  row[5] or 0,
            'at':       row[6] or 0,
            'oc':       row[7] or 0,
        }
    return out


# ── Cruzar Preliminar con Suministros críticos ───────────────────────────────
def match_sc(prelim_desc, sc_list):
    """Intenta casar descripción Preliminar con item Suministros críticos."""
    d = prelim_desc.lower()
    for s in sc_list:
        sc_d = s['desc_sc'].lower()
        if ('válvulas de control' in d or 'válvulas autor' in d or
                'válvulas blowdown' in d or 'válvulas shutdown' in d):
            if 'válvulas de control' in sc_d: return s
        if 'válvulas de seguridad' in d and 'válvulas de seguridad' in sc_d: return s
        if 'cables de instrument' in d and 'cables instrument' in sc_d: return s
        if 'cables - baja' in d and 'cables eléctricos' in sc_d: return s
    return None

def infer_at_state(item, sc_item=None):
    """Infiere estado AT y descripción CT1/CT0."""
    if sc_item:
        status = sc_item['status'].lower()
        if 'adjudicado' in status or 'kom' in status:
            return 'ADJUDICADO'
        if sc_item['at_cierre']:
            return 'AT CERRADO'
        if 'se envió a at' in status or 'en at' in status:
            if 'espera oferta final' in status or 'nueva oferta' in status:
                return 'AT - CT2 en curso'
            return 'AT - CT1 en curso'
        if 'proveedor nominado' in status and 'se espera' in status:
            return 'AT - EN PROCESO'
        if 'apertura de ofertas' in status:
            return 'REC.OF. EN EVALUACIÓN'
        if sc_item['recof']:
            return 'REC.OF. COMPLETA – EN AT'
        if 'solped liberada' in status or 'solpeds' in status or sc_item['solped']:
            return 'SOLPED LIBERADA'
    if item.get('at_rev0_real'): return 'AT COMPLETADO'
    if item.get('at_reva_real'): return 'CT0 EN CURSO'
    if item.get('recof_real_p'): return 'CT1 EN PROCESO'
    if item.get('solp_real_p'):  return 'EN RECEP.OF.'
    if item.get('ri_real_p'):    return 'EN SOLPED'
    if item.get('ri_p0'):        return 'PENDIENTE'
    return 'SIN RI'

def infer_est_general(item, sc_item=None):
    if sc_item:
        status = sc_item['status'].lower()
        if sc_item['oc_n'] or ('adjudicado' in status and sc_item['prov']):
            return 'OC / ADJUDICADO'
        if 'kom' in status:
            return 'OC / KOM'
        if 'espera oferta final' in status or 'nueva oferta' in status:
            return 'EN AT – CT2'
        if 'se envió a at' in status or 'en at' in status:
            return 'EN AT – CT1'
        if 'proveedor nominado' in status:
            return 'EN AT / NOMINADO'
        if 'apertura de ofertas' in status:
            return 'EVALUACIÓN OFERTAS'
        if sc_item['recof']:
            return 'EN AT'
        if 'solped liberada' in status or 'solpeds' in status or sc_item['solped']:
            return 'EN SOLPED / REC.OF.'
        if sc_item['ri_real']:
            return 'RI EMITIDA'
    if item.get('oc_real_p') or item.get('oc_n'): return 'OC EMITIDA'
    if item.get('at_rev0_real'):                  return 'AT COMPLETADO'
    if item.get('recof_real_p'):                  return 'EN AT'
    if item.get('solp_real_p'):                   return 'EN REC.OF.'
    if item.get('ri_real_p'):                     return 'EN SOLPED'
    if item.get('ri_p0'):                         return 'PENDIENTE'
    return 'SIN RI'

def get_solped_num(item, sc_item=None):
    """Extrae N° SOLPED del texto de status o del Preliminar."""
    if sc_item:
        nums = re.findall(r'\b2\d{7}\b', sc_item['status'])
        if nums:
            return ' / '.join(nums[:3])
    if item.get('solp_n'): return item['solp_n']
    return ''


def _dm_to_date(d, m, year=2026):
    """Convierte día/mes de un texto a date() del año indicado (default 2026)."""
    try:
        return date(year, int(m), int(d))
    except (ValueError, TypeError):
        return None


def parse_status_dates(status):
    """
    Extrae fechas de hitos REALES embebidas en el texto de 'Comentarios AT/Adjudicaciones'
    de Suministros Críticos. Sólo devuelve fechas con evidencia textual explícita.
    Devuelve dict con claves: recof_real, at_ct1_real, kom_real.
    """
    out = {}
    if not status:
        return out
    s = status.lower()

    # Recepción de ofertas: "Apertura de ofertas 21/5" → fecha de apertura/recepción
    m = re.search(r'apertura de ofertas\s+(?:el\s+)?(\d{1,2})[/-](\d{1,2})', s)
    if m:
        d = _dm_to_date(m.group(1), m.group(2))
        if d:
            out['recof_real'] = d

    # Inicio de Análisis Técnico (CT1): "El 13/3 se envió a AT" / "se envió a AT el 13/3"
    m = re.search(r'(?:el\s+)?(\d{1,2})[/-](\d{1,2})\s+se envió a at', s)
    if not m:
        m = re.search(r'se envió a at\s+(?:el\s+)?(\d{1,2})[/-](\d{1,2})', s)
    if m:
        d = _dm_to_date(m.group(1), m.group(2))
        if d:
            out['at_ct1_real'] = d

    # KOM (Kick-Off Meeting): "KOM 18/5" → inicio de fabricación post-adjudicación
    m = re.search(r'kom\s+(?:el\s+)?(\d{1,2})[/-](\d{1,2})', s)
    if m:
        d = _dm_to_date(m.group(1), m.group(2))
        if d:
            out['kom_real'] = d

    # Estado de liberación SOLPED (para diferenciar "liberada" de "en liberación")
    # "Solpeds ... liberadas" → ya emitida (hecho consumado)
    if re.search(r'solpeds?\s+[\d,/\s]+liberadas?', s):
        out['solped_estado'] = 'EMITIDA'
    # "Solpeds ... en liberación" → en trámite, NO emitida todavía
    elif re.search(r'solpeds?\s+[\d,/\s]+en liberaci', s):
        out['solped_estado'] = 'EN LIBERACIÓN'

    return out


# ═══════════════════════════════════════════════════════════════════════════════
# GENERACIÓN EXCEL
# ═══════════════════════════════════════════════════════════════════════════════

def write_header_rows(ws, esp_label, hdr_color, start_row=1):
    """
    Genera 3 filas de encabezado.
    Col 5 = P0 BASELINE (2024 – Preliminar)
    Col 6 = PROG. ACTUAL (Crono 4.11 – 2025/2026)
    Col 7 = REAL (Suministros críticos)
    """
    COLS = [
        # Info básica
        ('N°',              'INFO', '', 4),
        ('Criticidad',      'INFO', '', 12),
        ('N° RI',           'INFO', '', 28),
        ('Descripción RI',  'INFO', '', 42),
        # RI
        ('P0 Base',         'RI', 'P0', 11),
        ('Prog. Actual',    'RI', 'PA', 11),
        ('Real',            'RI', 'REAL', 11),
        ('Desvío (d)',       'RI', 'DESV', 9),
        ('Estado RI',       'RI', 'EST', 13),
        # SOLPED
        ('N° Solped',       'SOLPED', 'DOC', 20),
        ('Prog. Actual',    'SOLPED', 'PA', 11),
        ('Real',            'SOLPED', 'REAL', 11),
        ('Estado',          'SOLPED', 'EST', 13),
        # Rec. Ofertas
        ('Prog. Actual',    'REC.OFERTAS', 'PA', 11),
        ('Real',            'REC.OFERTAS', 'REAL', 11),
        ('Estado',          'REC.OFERTAS', 'EST', 13),
        # Análisis Técnico
        ('CT1 P0',          'ANÁLISIS TÉCNICO (AT)', 'CT1', 11),
        ('CT1 Real',        'ANÁLISIS TÉCNICO (AT)', 'CT1', 11),
        ('CT0 Prog.',       'ANÁLISIS TÉCNICO (AT)', 'CT0', 11),
        ('CT0 Real',        'ANÁLISIS TÉCNICO (AT)', 'CT0', 11),
        ('Estado AT',       'ANÁLISIS TÉCNICO (AT)', 'EST', 17),
        # OC
        ('N° OC',           'ORDEN DE COMPRA', 'DOC', 18),
        ('Proveedor',       'ORDEN DE COMPRA', 'INFO', 22),
        ('Nec. OC Prog',    'ORDEN DE COMPRA', 'FECHA', 11),
        ('OC Real',         'ORDEN DE COMPRA', 'FECHA', 11),
        ('L.T. (días)',     'ORDEN DE COMPRA', 'LT', 9),
        ('Fecha Entrega',   'ORDEN DE COMPRA', 'ENT', 11),
        ('Nec. Obra',       'ORDEN DE COMPRA', 'NEC', 11),
        # Demoras
        ('RI→Solped',       'ANÁLISIS DEMORAS (días)', 'D1', 10),
        ('Solped→RecOf',    'ANÁLISIS DEMORAS (días)', 'D2', 10),
        ('RecOf→AT',        'ANÁLISIS DEMORAS (días)', 'D3', 10),
        ('AT→OC',           'ANÁLISIS DEMORAS (días)', 'D4', 10),
        ('RI→OC Total',     'ANÁLISIS DEMORAS (días)', 'DTOT', 11),
        # Estado
        ('Est. General',    'ESTADO', 'GEN', 18),
        ('Observaciones',   'ESTADO', 'OBS', 45),
        ('Acciones',        'ESTADO', 'ACC', 35),
    ]

    for ci, (name, bloque, sub, w) in enumerate(COLS, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w

    # Fila 1: Bloques (merge grupos)
    BLOQUES = [
        ('INFO', 1, 4, hdr_color, 'FFFFFF'),
        ('RI - REQUISICIÓN DE INGENIERÍA', 5, 9, '4A235A', 'FFFFFF'),
        ('SOLPED', 10, 13, '1F4E79', 'FFFFFF'),
        ('RECEPCIÓN DE OFERTAS', 14, 16, '375623', 'FFFFFF'),
        ('ANÁLISIS TÉCNICO (AT)', 17, 21, '833C00', 'FFFFFF'),
        ('ORDEN DE COMPRA (OC)', 22, 28, '1F4E79', 'FFFFFF'),
        ('ANÁLISIS DEMORAS (días RI→OC)', 29, 33, '44546A', 'FFFFFF'),
        ('ESTADO', 34, 36, C['titulo'], 'FFFFFF'),
    ]
    for label, c1, c2, bg, fg in BLOQUES:
        cell = ws.cell(row=start_row, column=c1, value=label)
        cell.fill = F(bg); cell.font = ft(True, fg, 10)
        cell.alignment = al('center', 'center')
        cell.border = bd_med()
        if c2 > c1:
            ws.merge_cells(start_row=start_row, start_column=c1,
                           end_row=start_row, end_column=c2)

    # Fila 2: sub-bloques
    sub_style = [
        (1,  esp_label,      hdr_color),
        (5,  'P0 BASELINE',  '6D2077'),
        (6,  'PROG. ACTUAL', '6D2077'),
        (7,  'REAL',         '6D2077'),
        (8,  'DESVÍO',       '6D2077'),
        (9,  'ESTADO',       '6D2077'),
        (10, 'N° SOLPED',    '2C5F8A'),
        (11, 'PROG. ACTUAL', '2C5F8A'),
        (12, 'REAL',         '2C5F8A'),
        (13, 'ESTADO',       '2C5F8A'),
        (14, 'PROG. ACTUAL', '3D6B2F'),
        (15, 'REAL',         '3D6B2F'),
        (16, 'ESTADO',       '3D6B2F'),
        (17, 'CT1 (Rev.A)',  '9B3A00'),
        (18, '',             '9B3A00'),
        (19, 'CT0 (Rev.0)', '9B3A00'),
        (20, '',            '9B3A00'),
        (21, 'ESTADO AT',   '9B3A00'),
        (22, 'N° OC',       '2C5F8A'),
        (23, 'PROVEEDOR',   '2C5F8A'),
        (24, 'NEC. OC',     '2C5F8A'),
        (25, 'OC REAL',     '2C5F8A'),
        (26, 'L.T.',        '2C5F8A'),
        (27, 'ENTREGA OC',  '2C5F8A'),
        (28, 'NEC. OBRA',   '2C5F8A'),
        (29, 'RI→Solped',   '44546A'),
        (30, 'Solped→RecOf','44546A'),
        (31, 'RecOf→AT',    '44546A'),
        (32, 'AT→OC',       '44546A'),
        (33, 'RI→OC TOTAL', '44546A'),
        (34, 'GENERAL',     C['titulo']),
        (35, 'OBSERVACIONES AT/ADJUDICACIONES', C['titulo']),
        (36, 'ACCIONES REQUERIDAS', C['titulo']),
    ]
    r2 = start_row + 1
    for col, label, bg in sub_style:
        c = ws.cell(row=r2, column=col, value=label)
        c.fill = F(bg); c.font = ft(True, 'FFFFFF', 8)
        c.alignment = al('center', 'center', wrap=True); c.border = bd()

    # Fila 3: nombres de columna individuales
    col3_labels = [
        'N°', 'Criticidad', 'N° RI (Datos P0)', 'Descripción RI',
        'P0 Base (2024)', 'Prog. Actual (crono)', 'Real', 'Desvío (d)', 'Estado RI',
        'N° Solped', 'Prog. Actual (crono)', 'Real', 'Estado',
        'Prog. Actual (crono)', 'Real', 'Estado',
        'CT1 Prog (P0)', 'CT1 Real', 'CT0 Prog (crono)', 'CT0 Real/Cierre', 'Estado AT',
        'N° OC', 'Proveedor', 'Nec. OC Prog', 'OC Real/KOM', 'L.T. días',
        'Fecha Entrega OC', 'Nec. Obra',
        'RI→Solped', 'Solped→RecOf', 'RecOf→AT', 'AT→OC', 'RI→OC Total',
        'Est. General', 'Obs. / Comentarios AT-Adj.', 'Acciones Necesarias',
    ]
    r3 = start_row + 2
    for ci, label in enumerate(col3_labels, start=1):
        c = ws.cell(row=r3, column=ci, value=label)
        c.fill = F(C['hdr_col']); c.font = ft(True, '1F3864', 8)
        c.alignment = al('center', 'center', wrap=True); c.border = bd()

    ws.row_dimensions[start_row].height = 20
    ws.row_dimensions[r2].height = 24
    ws.row_dimensions[r3].height = 30
    return len(COLS)


def write_data_row(ws, row_n, seq, item, sc_item, ri_num, hdr_color, is_alt, crono_item=None):
    bg = C['row_alt'] if is_alt else C['row_norm']

    def cell(col, val, fmt=None, bold=False, wrap=False, h='center'):
        c = ws.cell(row=row_n, column=col, value=val)
        c.fill = F(bg); c.font = ft(bold=bold, sz=9)
        c.alignment = al(h, 'center', wrap=wrap); c.border = bd()
        if fmt: c.number_format = fmt
        return c

    # ── Fechas RI ─────────────────────────────────────────────────────────────
    ri_p0   = item.get('ri_p0')                         # P0 baseline (2024)
    ri_prog = crono_item['ri_prog'] if crono_item else None  # Prog. Actual (crono 2025/2026)
    ri_real = sc_item['ri_real'] if sc_item else item.get('ri_real_p')
    ri_desv = item.get('ri_desv')

    # ── SOLPED ────────────────────────────────────────────────────────────────
    solp_n    = get_solped_num(item, sc_item)
    solp_prog = crono_item['sp_prog'] if crono_item else None   # crono 2025/2026
    solp_real = sc_item['solped'] if sc_item else item.get('solp_real_p')

    # ── Rec. Ofertas ──────────────────────────────────────────────────────────
    recof_prog = crono_item['rec_prog'] if crono_item else None  # crono 2025/2026
    recof_real = sc_item['recof'] if sc_item else item.get('recof_real_p')

    # ── AT ────────────────────────────────────────────────────────────────────
    # CT1 (Rev.A): no está en Crono, se usa P0 Preliminar
    at_ct1_p0 = item.get('at_reva_p0')
    at_ct1_r  = item.get('at_reva_real')
    # CT0 (Rev.0): programa actual desde Crono
    at_ct0_prog = crono_item['at0_prog'] if crono_item else item.get('at_rev0_p0')
    at_ct0_r    = sc_item['at_cierre'] if sc_item else item.get('at_rev0_real')

    # ── OC ────────────────────────────────────────────────────────────────────
    oc_n_v    = sc_item['oc_n']      if sc_item else item.get('oc_n', '')
    prov_v    = sc_item['prov']      if sc_item else item.get('prov', '')
    # Nec. OC: SC > crono > P0 Preliminar
    nec_oc_v  = (sc_item['nec_oc']  if sc_item else None) or \
                (crono_item['oc_prog'] if crono_item else None) or \
                item.get('oc_p0')
    oc_real_v = sc_item['oc_real_d'] if sc_item else item.get('oc_real_p')
    lt_v      = sc_item['lt']        if sc_item else None
    ent_oc_v  = sc_item['ent_oc']    if sc_item else item.get('ent_oc')
    nec_obra_v= sc_item['nec_ent']   if sc_item else item.get('nec_obra')
    status_v  = sc_item['status']    if sc_item else ''
    kom_v     = sc_item.get('kom')   if sc_item else None

    # ── Completar fechas REALES desde el texto de estado (SC) ─────────────────
    # Sólo cuando hay evidencia textual explícita y el campo está vacío.
    parsed = parse_status_dates(status_v)
    if not recof_real and parsed.get('recof_real'):
        recof_real = parsed['recof_real']           # ej. "Apertura de ofertas 21/5"
    if not at_ct1_r and parsed.get('at_ct1_real'):
        at_ct1_r = parsed['at_ct1_real']            # ej. "El 13/3 se envió a AT"
    if not kom_v and parsed.get('kom_real'):
        kom_v = parsed['kom_real']                  # ej. "KOM 18/5"

    at_state  = infer_at_state(item, sc_item)
    est_gen   = infer_est_general(item, sc_item)
    crit      = sc_item['crit'] if sc_item else ''

    # ── Análisis demoras ──────────────────────────────────────────────────────
    ri_date_for_delay    = ri_real or ri_prog or ri_p0
    solp_date_for_delay  = solp_real or solp_prog
    recof_date_for_delay = recof_real or recof_prog
    at_date_for_delay    = at_ct0_r or at_ct1_r
    oc_date_for_delay    = oc_real_v or kom_v or nec_oc_v

    d_ri_solp  = days_between(ri_date_for_delay,    solp_date_for_delay)
    d_solp_rec = days_between(solp_date_for_delay,  recof_date_for_delay)
    d_rec_at   = days_between(recof_date_for_delay, at_date_for_delay)
    d_at_oc    = days_between(at_date_for_delay,    oc_date_for_delay)
    d_ri_oc    = days_between(ri_date_for_delay,    oc_date_for_delay)

    # ── Escribir celdas ───────────────────────────────────────────────────────
    # Col 1: N°
    c1 = ws.cell(row=row_n, column=1, value=seq)
    c1.fill = F(hdr_color); c1.font = ft(True, 'FFFFFF', 9)
    c1.alignment = al('center', 'center'); c1.border = bd()

    # Col 2: Criticidad
    c2 = ws.cell(row=row_n, column=2, value=crit)
    cc = CRIT_COLOR.get(crit, bg)
    c2.fill = F(cc); c2.font = ft(True, 'FFFFFF' if crit in ('LLI','C. CRÍTICO','HITO 2') else '000000', 8)
    c2.alignment = al('center', 'center'); c2.border = bd()

    cell(3, ri_num, h='left')

    desc_v = item.get('desc') or (sc_item['desc_sc'] if sc_item else '')
    for pre in ['Requisición de Ingeniería ', 'RI - ', 'RI-REQUISICIÓN DE INGENIERÍA - ',
                'REQUISICIÓN DE INGENIERÍA ']:
        if desc_v.upper().startswith(pre.upper()):
            desc_v = desc_v[len(pre):]
            break
    ws.cell(row=row_n, column=4, value=desc_v).fill = F(bg)
    ws.cell(row=row_n, column=4).font = ft(bold=False, sz=9)
    ws.cell(row=row_n, column=4).alignment = al('left', 'center', wrap=True)
    ws.cell(row=row_n, column=4).border = bd()

    def date_cell(col, val):
        c = ws.cell(row=row_n, column=col, value=fmt_date(val))
        c.fill = F(bg); c.font = ft(sz=9); c.alignment = al('center'); c.border = bd()
        return c

    # ── RI (col 5-9) ─────────────────────────────────────────────────────────
    date_cell(5, ri_p0)    # P0 baseline 2024
    date_cell(6, ri_prog)  # Prog. Actual (crono 2025/2026)
    date_cell(7, ri_real)  # Real (SC)

    c_desv = ws.cell(row=row_n, column=8, value=ri_desv if ri_desv is not None else '')
    c_desv.fill = F(bg); c_desv.font = ft(sz=9); c_desv.alignment = al('center'); c_desv.border = bd()
    if ri_desv and isinstance(ri_desv, (int, float)) and ri_desv > 15:
        c_desv.fill = F(C['demora_crit'])

    if ri_real:  est_ri = 'COMPLETADA'
    elif ri_p0:  est_ri = 'EMITIDA' if (ri_prog or TODAY >= ri_p0) else 'PENDIENTE'
    else:        est_ri = 'SIN RI'
    apply_status_cell(ws, row_n, 9, est_ri)

    # ── SOLPED (col 10-13) ────────────────────────────────────────────────────
    s_num = ws.cell(row=row_n, column=10, value=solp_n)
    s_num.fill = F(bg); s_num.font = ft(sz=8)
    s_num.alignment = al('left', 'center', wrap=True); s_num.border = bd()

    date_cell(11, solp_prog)  # Prog. Actual (crono)
    date_cell(12, solp_real)  # Real (SC)

    if solp_real:
        est_sp = 'COMPLETADA'
    elif parsed.get('solped_estado'):
        # Texto explícito del estado: "liberada" → EMITIDA, "en liberación" → EN LIBERACIÓN
        est_sp = parsed['solped_estado']
    elif solp_n:
        est_sp = 'EMITIDA'
    elif solp_prog:
        est_sp = 'PENDIENTE'
    else:
        est_sp = 'SIN RI'
    apply_status_cell(ws, row_n, 13, est_sp)

    # ── REC. OFERTAS (col 14-16) ──────────────────────────────────────────────
    date_cell(14, recof_prog)  # Prog. Actual (crono)
    date_cell(15, recof_real)  # Real (SC)

    if recof_real:  est_ro = 'COMPLETADA'
    elif recof_prog and TODAY > recof_prog: est_ro = 'ATRASADO'
    elif solp_real: est_ro = 'EN PROCESO'
    elif recof_prog: est_ro = 'PENDIENTE'
    else:           est_ro = 'SIN RI'
    apply_status_cell(ws, row_n, 16, est_ro)

    # ── AT (col 17-21) ────────────────────────────────────────────────────────
    date_cell(17, at_ct1_p0)    # CT1 P0 baseline (Preliminar)
    date_cell(18, at_ct1_r)     # CT1 Real (Preliminar)
    date_cell(19, at_ct0_prog)  # CT0 Prog. Actual (crono)
    date_cell(20, at_ct0_r)     # CT0 Real/Cierre (SC o Preliminar)

    at_cell = ws.cell(row=row_n, column=21, value=at_state)
    if at_state in ('ADJUDICADO', 'AT COMPLETADO', 'AT CERRADO', 'REC.OF. DONE'):
        at_cell.fill = F(C['completado']); at_cell.font = ft(True, 'FFFFFF', 8)
    elif 'CT' in at_state or 'AT' in at_state:
        at_cell.fill = F(C['en_proceso']); at_cell.font = ft(True, '000000', 8)
    else:
        at_cell.fill = F(C['pendiente']); at_cell.font = ft(sz=8)
    at_cell.alignment = al('center', 'center', wrap=True); at_cell.border = bd()

    # ── OC (col 22-28) ────────────────────────────────────────────────────────
    oc_nc = ws.cell(row=row_n, column=22, value=oc_n_v)
    if oc_n_v:
        oc_nc.fill = F(C['completado']); oc_nc.font = ft(True, 'FFFFFF', 9)
    else:
        oc_nc.fill = F(bg); oc_nc.font = ft(sz=9)
    oc_nc.alignment = al('center'); oc_nc.border = bd()

    prov_c = ws.cell(row=row_n, column=23, value=prov_v)
    prov_c.fill = F(bg); prov_c.font = ft(sz=9); prov_c.alignment = al('left'); prov_c.border = bd()

    date_cell(24, nec_oc_v)
    oc_or_kom = oc_real_v or kom_v
    date_cell(25, oc_or_kom)

    lt_c = ws.cell(row=row_n, column=26, value=lt_v if lt_v else '')
    lt_c.fill = F(bg); lt_c.font = ft(sz=9); lt_c.alignment = al('center'); lt_c.border = bd()

    date_cell(27, ent_oc_v)
    date_cell(28, nec_obra_v)

    # ── DEMORAS (col 29-33) ───────────────────────────────────────────────────
    style_delay_cell(ws, row_n, 29, d_ri_solp,  threshold_warn=7,  threshold_crit=21)
    style_delay_cell(ws, row_n, 30, d_solp_rec, threshold_warn=30, threshold_crit=60)
    style_delay_cell(ws, row_n, 31, d_rec_at,   threshold_warn=30, threshold_crit=60)
    style_delay_cell(ws, row_n, 32, d_at_oc,    threshold_warn=30, threshold_crit=60)
    style_delay_cell(ws, row_n, 33, d_ri_oc,    threshold_warn=90, threshold_crit=150)

    # ── ESTADO (col 34-36) ────────────────────────────────────────────────────
    est_c = ws.cell(row=row_n, column=34, value=est_gen)
    if 'OC' in est_gen or 'ADJUDICADO' in est_gen:
        est_c.fill = F(C['completado']); est_c.font = ft(True, 'FFFFFF', 9)
    elif 'AT' in est_gen or 'GESTIÓN' in est_gen:
        est_c.fill = F(C['en_proceso']); est_c.font = ft(True, '000000', 9)
    elif 'PENDIENTE' in est_gen or 'SIN RI' in est_gen:
        est_c.fill = F(C['pendiente']); est_c.font = ft(sz=9)
    else:
        est_c.fill = F(C['emitida']); est_c.font = ft(True, '000000', 9)
    est_c.alignment = al('center', 'center', wrap=True); est_c.border = bd()

    obs_txt = status_v[:300] if status_v else ''
    obs_c = ws.cell(row=row_n, column=35, value=obs_txt)
    obs_c.fill = F(bg); obs_c.font = ft(sz=8)
    obs_c.alignment = al('left', 'center', wrap=True); obs_c.border = bd()

    acc_v = ''
    if sc_item:
        s = sc_item['status'].lower()
        if sc_item['oc_n']:
            acc_v = 'Seguimiento fabricación / entrega OC'
        elif 'kom' in s:
            acc_v = 'Monitorear fabricación – verificar hitos KOM'
        elif 'se espera' in s and 'oferta' in s:
            acc_v = 'Seguimiento oferentes – aguardar oferta final'
        elif 'apertura de ofertas' in s:
            acc_v = 'Evaluar ofertas recibidas – iniciar AT'
        elif 'en at' in s or 'se envió a at' in s:
            acc_v = 'Completar Análisis Técnico – aguardar cierre CT'
        elif parsed.get('solped_estado') == 'EN LIBERACIÓN':
            acc_v = 'Completar liberación SOLPED – confirmar en SAP'
        elif 'liberadas' in s and not sc_item['recof']:
            acc_v = 'SOLPED liberada – coordinar apertura de ofertas'
        elif sc_item['solped']:
            acc_v = 'Coordinar apertura de ofertas'
        elif sc_item['ri_real']:
            acc_v = 'Liberar SOLPED'
    elif not ri_p0:
        acc_v = 'Emitir RI'
    elif not solp_real:
        acc_v = 'Liberar SOLPED'
    elif not recof_real:
        acc_v = 'Coordinar apertura de ofertas'
    elif not at_ct0_r:
        acc_v = 'Completar Análisis Técnico'
    elif not oc_n_v:
        acc_v = 'Colocar OC'

    acc_c = ws.cell(row=row_n, column=36, value=acc_v)
    acc_c.fill = F(bg); acc_c.font = ft(sz=8)
    acc_c.alignment = al('left', 'center', wrap=True); acc_c.border = bd()

    ws.row_dimensions[row_n].height = 32


def _build_pipeline_compare(ws, trend, start_row=3):
    """Renderiza la EVOLUCIÓN de 3 semanas del pipeline agregado (Cuadro resumen) por
    especialidad y etapa, respondiendo: SOLPED, OCs nuevas y ofertas en gestión.
    `trend` = lista de (label, cuadro_dict) ordenada por fecha. Devuelve la sig. fila libre."""
    labels = [t[0] for t in trend]
    cuadros = [t[1] for t in trend]
    cold, cnew = cuadros[0], cuadros[-1]
    ncols = len(labels)

    r = start_row
    ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=7)
    th = ws.cell(r, 1, f'EVOLUCIÓN 3 SEMANAS – PIPELINE POR ESPECIALIDAD (Cuadro resumen)  ·  {labels[0]} → {labels[-1]}')
    th.fill = F(C['subtitulo']); th.font = ft(True, 'FFFFFF', 11)
    th.alignment = al('center', 'center'); ws.row_dimensions[r].height = 20
    r += 1

    # Encabezado: Etapa | corte1..corteN | Δ
    hdrs = ['Especialidad · Etapa'] + labels + ['Δ 3 sem']
    for ci, h in enumerate(hdrs, start=1):
        c = ws.cell(r, ci, h); c.fill = F(C['hdr_grp']); c.font = ft(True, 'FFFFFF', 9)
        c.alignment = al('center', 'center', wrap=True); c.border = bd()
    ws.row_dimensions[r].height = 22
    r += 1

    ESP_NOMBRE = {'EL': 'ELECTRICIDAD', 'IN': 'INSTRUMENTACIÓN'}
    ESP_COLOR  = {'EL': C['hdr_el'], 'IN': C['hdr_in']}
    ETAPAS = [('solped', 'En SOLPED'), ('ofertas', 'En Petición Ofertas'),
              ('at', 'En AT'), ('oc', 'Con OC'), ('emitidas', 'En gestión (emitidas)')]
    for esp in ['EL', 'IN']:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=1+ncols+1)
        ce = ws.cell(r, 1, ESP_NOMBRE[esp]); ce.fill = F(ESP_COLOR[esp])
        ce.font = ft(True, 'FFFFFF', 9); ce.alignment = al('left', 'center'); ce.border = bd()
        ws.row_dimensions[r].height = 16; r += 1
        for campo, label in ETAPAS:
            lc = ws.cell(r, 1, label); lc.fill = F(C['row_norm']); lc.font = ft(False, sz=9)
            lc.alignment = al('left', 'center'); lc.border = bd()
            serie = [c.get(esp, {}).get(campo, 0) for c in cuadros]
            for ci, val in enumerate(serie, start=2):
                cc = ws.cell(r, ci, val); cc.fill = F(C['row_norm']); cc.font = ft(False, sz=9)
                cc.alignment = al('center'); cc.border = bd()
            d = serie[-1] - serie[0]
            dc = ws.cell(r, 2+ncols); dc.alignment = al('center'); dc.border = bd()
            if d > 0:   dc.value = f'+{d}'; dc.fill = F('C6EFCE'); dc.font = ft(True, '375623', 10)
            elif d < 0: dc.value = f'{d}';  dc.fill = F('DDEEFF'); dc.font = ft(True, '1F3864', 10)
            else:       dc.value = '=';     dc.fill = F(C['row_norm']); dc.font = ft(False, sz=10)
            ws.row_dimensions[r].height = 15; r += 1

    # Resumen de las 3 preguntas clave (IN+EL) sobre las 3 semanas
    def tot(cuadro, campo): return cuadro.get('EL', {}).get(campo, 0) + cuadro.get('IN', {}).get(campo, 0)
    d_oc = tot(cnew,'oc') - tot(cold,'oc')
    d_at = tot(cnew,'at') - tot(cold,'at')
    preguntas = [
        ('SOLPED en las 3 semanas (IN+EL)',
         f'{tot(cold,"solped")} → {tot(cnew,"solped")} SOLPED en proceso: las {tot(cold,"solped")} iniciales se '
         f'liberaron/avanzaron. Cables EL con 2 SOLPED liberadas (apertura 20/07).'),
        ('OCs nuevas (IN+EL)',
         f'{tot(cold,"oc")} → {tot(cnew,"oc")} OCs colocadas (+{d_oc} en 3 semanas): '
         f'IN {cold.get("IN",{}).get("oc",0)}→{cnew.get("IN",{}).get("oc",0)}, '
         f'EL {cold.get("EL",{}).get("oc",0)}→{cnew.get("EL",{}).get("oc",0)}. '
         f'Destacada: SIS OC 4509010341 (crítico resuelto).'),
        ('Ofertas para seguir la compra',
         f'En petición de ofertas: EL {cnew.get("EL",{}).get("ofertas",0)} · IN {cnew.get("IN",{}).get("ofertas",0)} '
         f'= {tot(cnew,"ofertas")} ítems. Ofertas recibidas Cables IN (23/06); próxima apertura Cables EL 20/07. '
         f'En AT: {tot(cold,"at")} → {tot(cnew,"at")} (+{d_at}).'),
    ]
    for titulo, texto in preguntas:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=2)
        ct = ws.cell(r, 1, titulo); ct.fill = F('FFF2CC'); ct.font = ft(True, '7F6000', 9)
        ct.alignment = al('left', 'center', wrap=True); ct.border = bd()
        ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
        cx = ws.cell(r, 3, texto); cx.fill = F(C['row_norm']); cx.font = ft(False, sz=9)
        cx.alignment = al('left', 'center', wrap=True); cx.border = bd()
        ws.row_dimensions[r].height = 30
        r += 1

    ws.row_dimensions[r].height = 6
    return r + 1


def build_cambios_semana(wb, sc_old, sc_new, trend=None):
    """
    Hoja de comparación: evolución de 3 semanas plan 250626 → 170726.
    Incluye (1) evolución del pipeline agregado por especialidad (Cuadro resumen):
    SOLPED, OCs nuevas y ofertas en gestión a lo largo de 4 cortes; y (2) los campos que
    cambiaron por ítem de Suministros Críticos (extremos del período).
    """
    def norm(s): return re.sub(r'\s+', ' ', s.strip().lower())
    def fmtv(v):
        if isinstance(v, date): return v.strftime('%d/%m/%y')
        return str(v or '').strip()

    old_by_desc = {norm(s['desc_sc']): s for s in sc_old}
    new_by_desc = {norm(s['desc_sc']): s for s in sc_new}

    CAMPOS = [
        ('ri_real',   'RI Real (emisión)'),
        ('solped',    'SOLPED'),
        ('recof',     'Rec. Ofertas'),
        ('at_cierre', 'AT Cierre / CT0'),
        ('nec_oc',    'Nec. OC'),
        ('oc_real_d', 'OC Real'),
        ('kom',       'KOM'),
        ('oc_n',      'N° OC'),
        ('prov',      'Proveedor'),
    ]

    rows = []  # (esp, desc, campo, v_old, v_new, tipo)
    for s_new in sc_new:
        key = norm(s_new['desc_sc'])
        s_old = old_by_desc.get(key)
        if s_old is None:
            rows.append((s_new['esp'], s_new['desc_sc'], 'ÍTEM NUEVO', '—', 'En plan 120626', 'NUEVO'))
            continue
        for field, label in CAMPOS:
            v_old = fmtv(s_old.get(field))
            v_new = fmtv(s_new.get(field))
            if v_old != v_new:
                if not v_old and v_new:
                    tipo = 'HITO ALCANZADO'
                elif v_old and not v_new:
                    tipo = 'DATO ELIMINADO'
                else:
                    tipo = 'ACTUALIZACIÓN'
                rows.append((s_new['esp'], s_new['desc_sc'], label,
                             v_old or '—', v_new or '—', tipo))
        # Estado/comentarios: comparar primeros 200 chars
        vs_old = (s_old.get('status') or '')[:200]
        vs_new = (s_new.get('status') or '')[:200]
        if vs_old.strip() != vs_new.strip():
            rows.append((s_new['esp'], s_new['desc_sc'], 'Estado/Comentarios',
                         (vs_old[:120] or '—'), (vs_new[:120] or '—'), 'ESTADO ACTUALIZADO'))

    for s_old in sc_old:
        if norm(s_old['desc_sc']) not in new_by_desc:
            rows.append((s_old['esp'], s_old['desc_sc'], 'ÍTEM ELIMINADO',
                         s_old['desc_sc'], '—', 'ELIMINADO'))

    ws = wb.create_sheet('CAMBIOS SEMANA')
    ws.sheet_view.showGridLines = False

    for col, w in zip('ABCDEFG', [4, 6, 36, 22, 50, 50, 20]):
        ws.column_dimensions[col].width = w

    ws.merge_cells('A1:G1')
    t = ws['A1']
    t.value = 'EVOLUCIÓN: Plan 30/07 → 07/08 → 14/08 → 21/08 → 28/08/2026 | CAMBIOS Y AVANCES'
    t.fill = F(C['titulo']); t.font = ft(True, 'FFFFFF', 13)
    t.alignment = al('center', 'center'); ws.row_dimensions[1].height = 26

    ws.merge_cells('A2:G2')
    s2 = ws['A2']
    s2.value = ('Especialidades: IN (Instrumentación & Control)  |  EL (Electricidad)  |  '
                'Proyecto: La Calera II CPF2  |  Corte actual: 28/08/2026 (Rev12)')
    s2.fill = F(C['subtitulo']); s2.font = ft(False, 'FFFFFF', 10)
    s2.alignment = al('center', 'center'); ws.row_dimensions[2].height = 18

    # ── Bloque 1: Evolución pipeline agregado 3 semanas (Cuadro resumen) ───────
    base = 3
    if trend:
        base = _build_pipeline_compare(ws, trend, start_row=3)

    # ── Bloque 2: Cambios por ítem de Suministros Críticos ────────────────────
    ws.merge_cells(start_row=base, start_column=1, end_row=base, end_column=7)
    sub = ws.cell(base, 1, 'DETALLE DE CAMBIOS POR ÍTEM – SUMINISTROS CRÍTICOS IN / EL (21/08 → 28/08)')
    sub.fill = F(C['subtitulo']); sub.font = ft(True, 'FFFFFF', 11)
    sub.alignment = al('center', 'center'); ws.row_dimensions[base].height = 20
    base += 1

    if not rows:
        ws.merge_cells(start_row=base, start_column=1, end_row=base, end_column=7)
        nc = ws.cell(base, 1, 'Sin cambios en los ítems críticos IN/EL respecto del plan anterior '
                              '(los avances de la semana se reflejan en el pipeline agregado, bloque superior).')
        nc.font = ft(sz=10); nc.alignment = al('center'); ws.row_dimensions[base].height = 18
        return ws

    n_hitos  = sum(1 for r in rows if r[5] == 'HITO ALCANZADO')
    n_estado = sum(1 for r in rows if r[5] == 'ESTADO ACTUALIZADO')
    n_otros  = len(rows) - n_hitos - n_estado
    ws.merge_cells(start_row=base, start_column=1, end_row=base, end_column=7)
    res = ws.cell(base, 1)
    res.value = (f'Resumen: {len(rows)} cambios detectados — '
                 f'{n_hitos} hitos nuevos alcanzados · '
                 f'{n_estado} estados actualizados · {n_otros} otros')
    res.fill = F('FFF2CC'); res.font = ft(True, '7F6000', 10)
    res.alignment = al('center', 'center'); ws.row_dimensions[base].height = 18
    base += 1

    r = base
    hdrs = ['N°', 'Esp', 'Suministro / Ítem', 'Campo', 'Plan 210826 (anterior)', 'Plan 280826 (actual)', 'Tipo de cambio']
    for ci, h in enumerate(hdrs, start=1):
        c = ws.cell(r, ci, h)
        c.fill = F(C['hdr_grp']); c.font = ft(True, 'FFFFFF', 9)
        c.alignment = al('center', 'center', wrap=True); c.border = bd()
    ws.row_dimensions[r].height = 28

    TIPO_COLOR = {
        'HITO ALCANZADO':    ('C6EFCE', '375623'),
        'ESTADO ACTUALIZADO':('DDEEFF', '1F3864'),
        'ACTUALIZACIÓN':     ('FFEB9C', '9C6500'),
        'NUEVO':             ('E2EFDA', '375623'),
        'DATO ELIMINADO':    ('FFC7CE', 'C00000'),
        'ELIMINADO':         ('FFC7CE', 'C00000'),
    }
    for i, (esp, desc, campo, v_old, v_new, tipo) in enumerate(rows, start=1):
        r += 1
        base_bg = C['row_alt'] if i % 2 == 0 else C['row_norm']
        tipo_bg, tipo_fg = TIPO_COLOR.get(tipo, (base_bg, '000000'))

        c1 = ws.cell(r, 1, i); c1.fill = F(C['hdr_col']); c1.font = ft(True, sz=8)
        c1.alignment = al('center'); c1.border = bd()

        espc = C['hdr_el'] if esp == 'EL' else C['hdr_in']
        c2 = ws.cell(r, 2, esp); c2.fill = F(espc); c2.font = ft(True, 'FFFFFF', 9)
        c2.alignment = al('center'); c2.border = bd()

        c3 = ws.cell(r, 3, desc); c3.fill = F(base_bg); c3.font = ft(sz=8)
        c3.alignment = al('left', 'center', wrap=True); c3.border = bd()

        c4 = ws.cell(r, 4, campo); c4.fill = F(base_bg); c4.font = ft(True, sz=8)
        c4.alignment = al('center', 'center', wrap=True); c4.border = bd()

        c5 = ws.cell(r, 5, v_old); c5.fill = F('FFF2CC'); c5.font = ft(sz=8)
        c5.alignment = al('left', 'center', wrap=True); c5.border = bd()

        c6 = ws.cell(r, 6, v_new); c6.fill = F(tipo_bg); c6.font = ft(True, tipo_fg, 8)
        c6.alignment = al('left', 'center', wrap=True); c6.border = bd()

        c7 = ws.cell(r, 7, tipo); c7.fill = F(tipo_bg); c7.font = ft(True, tipo_fg, 8)
        c7.alignment = al('center', 'center'); c7.border = bd()

        ws.row_dimensions[r].height = 40 if tipo == 'ESTADO ACTUALIZADO' else 22

    return ws


def build_detail_sheet(wb, sheet_name, esp, prelim_items, sc_list, hdr_color, ri_map, crono_items):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_view.showGridLines = False

    ws.merge_cells('A1:AJ1')
    t = ws['A1']; t.value = f'TRACKER SUMINISTROS – LA CALERA II CPF2 | {sheet_name} | {VERSION}'
    t.fill = F(C['titulo']); t.font = ft(True, 'FFFFFF', 12)
    t.alignment = al('center', 'center')
    ws.row_dimensions[1].height = 26

    write_header_rows(ws, esp, hdr_color, start_row=2)
    ws.freeze_panes = 'A5'

    current_row = 5
    seq = 1
    # sheet_name is 'IN' or 'EL' — used for crono matching (esp label is the long form)
    esp_code = sheet_name

    # ── Items Preliminar ────────────────────────────────────────────────────
    for item in prelim_items:
        sc_match    = match_sc(item['desc'], sc_list)
        crono_match = find_crono_item(item['desc'], esp_code, crono_items)
        ri_num      = get_ri_number(item['desc'], ri_map)
        is_alt      = (seq % 2 == 0)
        write_data_row(ws, current_row, seq, item, sc_match, ri_num, hdr_color, is_alt, crono_match)
        current_row += 1
        seq += 1

    # ── Items adicionales de Suministros críticos (no en Preliminar) ────────
    already_matched = set()
    for item in prelim_items:
        m = match_sc(item['desc'], sc_list)
        if m: already_matched.add(id(m))

    for sc_item in sc_list:
        if id(sc_item) in already_matched:
            continue
        empty_item = {k: None for k in ['desc','ri_p0','ri_real_p','ri_desv',
                                         'solp_real_p','solp_n',
                                         'recof_real_p',
                                         'at_reva_p0','at_reva_real','at_rev0_p0','at_rev0_real',
                                         'oc_real_p','oc_n','prov','ent_oc','nec_obra']}
        ri_num = get_ri_number(sc_item['desc_sc'], ri_map)
        # SC-only items may still match crono (e.g., shelter → shelter y banco de carga)
        crono_match = find_crono_item(sc_item['desc_sc'], esp_code, crono_items)
        is_alt = (seq % 2 == 0)
        write_data_row(ws, current_row, seq, empty_item, sc_item, ri_num, hdr_color, is_alt, crono_match)
        current_row += 1
        seq += 1

    ws.row_dimensions[current_row].height = 6


def build_dashboard(wb, sc_in, sc_el, in_items, el_items):
    ws = wb.create_sheet('DASHBOARD')
    ws.sheet_view.showGridLines = False

    ws.merge_cells('A1:R1')
    t = ws['A1']; t.value = f'DASHBOARD – PLAN DE SUMINISTROS LA CALERA II CPF2 | {VERSION} | RFSU: {RFSU.strftime("%d/%m/%Y")}'
    t.fill = F(C['titulo']); t.font = ft(True, 'FFFFFF', 13)
    t.alignment = al('center', 'center'); ws.row_dimensions[1].height = 28
    ws.row_dimensions[2].height = 8

    def kpi_block(ws, row, col, label, val, color):
        ws.merge_cells(start_row=row, start_column=col, end_row=row+1, end_column=col+1)
        c = ws.cell(row=row, column=col, value=label)
        c.fill = F(color); c.font = ft(True, 'FFFFFF', 9)
        c.alignment = al('center', 'top'); c.border = bd_med()
        ws.merge_cells(start_row=row+2, start_column=col, end_row=row+3, end_column=col+1)
        cv = ws.cell(row=row+2, column=col, value=val)
        cv.fill = F(color); cv.font = Font(bold=True, color='FFFFFF', size=22, name='Calibri')
        cv.alignment = al('center', 'center'); cv.border = bd_med()

    row0 = 3
    el_oc = sum(1 for s in sc_el if s['oc_n'])
    el_at = sum(1 for s in sc_el if not s['oc_n'] and (s['recof'] or s['solped']))
    kpi_block(ws, row0, 1,  f'EL – TOTAL RIs\n(Plan+Críticos)', len(el_items) + len(sc_el), C['hdr_el'])
    kpi_block(ws, row0, 4,  f'EL – Con OC',                 el_oc,             '375623')
    kpi_block(ws, row0, 7,  f'EL – En AT/Gestión',          el_at,             '833C00')
    in_oc = sum(1 for s in sc_in if s['oc_n'])
    in_at = sum(1 for s in sc_in if not s['oc_n'] and (s['recof'] or s['solped']))
    kpi_block(ws, row0, 10, f'IN – TOTAL RIs\n(Plan+Críticos)', len(in_items) + len(sc_in), C['hdr_in'])
    kpi_block(ws, row0, 13, f'IN – Con OC',                 in_oc,             '843C0C')
    kpi_block(ws, row0, 16, f'IN – En AT/Gestión',          in_at,             '9B3A00')

    row0 += 6
    ws.row_dimensions[row0].height = 8
    row0 += 1

    ws.merge_cells(start_row=row0, start_column=1, end_row=row0, end_column=18)
    th = ws.cell(row=row0, column=1, value='SUMINISTROS CRÍTICOS – ESTADO ACTUAL')
    th.fill = F(C['subtitulo']); th.font = ft(True, 'FFFFFF', 11)
    th.alignment = al('center'); ws.row_dimensions[row0].height = 20
    row0 += 1

    DASH_COLS = ['Esp.','Criticidad','Descripción','N° RI','RI Real','Solped','RecOf',
                 'At Cierre','Nec.OC','OC Real/KOM','LT días','N° OC','Proveedor',
                 'Fecha Entrega','Nec. Entrega','Estado AT','Est.General','Observaciones']
    widths = [5, 12, 38, 30, 11, 11, 11, 11, 11, 11, 8, 18, 22, 11, 11, 16, 16, 55]
    for ci, (lbl, w) in enumerate(zip(DASH_COLS, widths), 1):
        ws.column_dimensions[get_column_letter(ci)].width = w
        c = ws.cell(row=row0, column=ci, value=lbl)
        c.fill = F(C['titulo']); c.font = ft(True, 'FFFFFF', 9)
        c.alignment = al('center', 'center', wrap=True); c.border = bd_med()
    ws.row_dimensions[row0].height = 28
    row0 += 1

    all_sc = [('EL', C['hdr_el'], s) for s in sc_el] + [('IN', C['hdr_in'], s) for s in sc_in]
    for idx, (esp, esp_color, s) in enumerate(all_sc):
        bg = C['row_alt'] if idx % 2 == 0 else C['row_norm']
        ri_map = EL_RI_NUMBER if esp == 'EL' else IN_RI_NUMBER
        ri_num = get_ri_number(s['desc_sc'], ri_map)
        at_st  = infer_at_state({}, s)
        est_g  = infer_est_general({}, s)

        def dc(col, val, h='center', wrap=False, bold=False):
            c = ws.cell(row=row0, column=col, value=val)
            c.fill = F(bg); c.font = ft(bold=bold, sz=9)
            c.alignment = al(h, 'center', wrap=wrap); c.border = bd()

        ws.cell(row=row0, column=1, value=esp).fill = F(esp_color)
        ws.cell(row=row0, column=1).font = ft(True, 'FFFFFF', 9)
        ws.cell(row=row0, column=1).alignment = al('center')
        ws.cell(row=row0, column=1).border = bd()

        c2 = ws.cell(row=row0, column=2, value=s['crit'])
        cc = CRIT_COLOR.get(s['crit'], bg)
        c2.fill = F(cc); c2.font = ft(True, 'FFFFFF' if s['crit'] in ('LLI','C. CRÍTICO','HITO 2') else '000000', 8)
        c2.alignment = al('center'); c2.border = bd()

        # Enriquecer con fechas reales extraídas del texto de estado
        parsed = parse_status_dates(s['status'])
        recof_d   = s['recof']     or parsed.get('recof_real')
        atcierre_d= s['at_cierre'] or parsed.get('at_ct1_real')
        kom_d     = s.get('kom')   or parsed.get('kom_real')

        desc_sh = s['desc_sc']
        for pre in ['RI - ', 'RI-REQUISICIÓN DE INGENIERÍA - ']:
            if desc_sh.upper().startswith(pre.upper()): desc_sh = desc_sh[len(pre):]; break
        dc(3,  desc_sh,             h='left', wrap=True)
        dc(4,  ri_num,              h='left', wrap=True)
        dc(5,  fmt_date(s['ri_real']))
        dc(6,  fmt_date(s['solped']))
        dc(7,  fmt_date(recof_d))
        dc(8,  fmt_date(atcierre_d))
        dc(9,  fmt_date(s['nec_oc']))
        dc(10, fmt_date(s['oc_real_d'] or kom_d))
        dc(11, s['lt'])

        oc_c = ws.cell(row=row0, column=12, value=s['oc_n'])
        if s['oc_n']:
            oc_c.fill = F(C['completado']); oc_c.font = ft(True, 'FFFFFF', 9)
        else:
            oc_c.fill = F(bg); oc_c.font = ft(sz=9)
        oc_c.alignment = al('center'); oc_c.border = bd()

        dc(13, s['prov'])
        dc(14, fmt_date(s['ent_oc']))
        dc(15, fmt_date(s['nec_ent']))

        at_c = ws.cell(row=row0, column=16, value=at_st)
        if 'ADJUDICADO' in at_st or 'CERRADO' in at_st or 'COMPLETADO' in at_st:
            at_c.fill = F(C['completado']); at_c.font = ft(True, 'FFFFFF', 8)
        elif 'AT' in at_st or 'CT' in at_st:
            at_c.fill = F(C['en_proceso']); at_c.font = ft(True, '000000', 8)
        else:
            at_c.fill = F(C['pendiente']); at_c.font = ft(sz=8)
        at_c.alignment = al('center', 'center', wrap=True); at_c.border = bd()

        est_c = ws.cell(row=row0, column=17, value=est_g)
        if 'OC' in est_g or 'ADJUDICADO' in est_g:
            est_c.fill = F(C['completado']); est_c.font = ft(True, 'FFFFFF', 9)
        elif 'AT' in est_g or 'GESTIÓN' in est_g:
            est_c.fill = F(C['en_proceso']); est_c.font = ft(True, '000000', 9)
        else:
            est_c.fill = F(C['emitida']); est_c.font = ft(sz=9)
        est_c.alignment = al('center', 'center', wrap=True); est_c.border = bd()

        obs_c = ws.cell(row=row0, column=18, value=s['status'][:250] if s['status'] else '')
        obs_c.fill = F(bg); obs_c.font = ft(sz=8)
        obs_c.alignment = al('left', 'center', wrap=True); obs_c.border = bd()

        ws.row_dimensions[row0].height = 36
        row0 += 1

    ws.freeze_panes = 'A10'


def build_portada(wb):
    ws = wb.create_sheet('PORTADA')
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 60
    ws.column_dimensions['C'].width = 30

    def row_h(r, h): ws.row_dimensions[r].height = h

    row_h(1, 12); row_h(2, 60); row_h(3, 12)
    ws.merge_cells('B2:C2')
    t = ws['B2']
    t.value = 'TRACKER DE SITUACIÓN DE SUMINISTROS'
    t.fill = F(C['titulo']); t.font = Font(bold=True, color='FFFFFF', size=22, name='Calibri')
    t.alignment = al('center', 'center')

    data = [
        ('Proyecto',           'La Calera II – CPF2'),
        ('Especialidades',     'Instrumentación & Control (IN) | Electricidad (EL)'),
        ('Fuente plan',        '29.07.2026 – Plan de Suministros – La Calera II (280826).xlsx · Rev12'),
        ('Versión tracker',    VERSION),
        ('Fecha generación',   TODAY.strftime('%d/%m/%Y')),
        ('RFSU objetivo',      RFSU.strftime('%d/%m/%Y')),
        ('Días hasta RFSU',    str((RFSU - TODAY).days) + ' días'),
        ('Objetivo análisis',  'Identificar demoras en circuito RI → OC y acciones correctivas'),
        ('Fuente fechas prog.','Crono 4.11 (2025-2026) — columnas "Prog. Actual"'),
        ('Fuente fechas P0',   'Preliminar (2024) — columnas "P0 Base"'),
        ('Fuente fechas Real', 'Suministros Críticos — columnas "Real" y OC/KOM'),
    ]
    r = 4
    for lbl, val in data:
        row_h(r, 22)
        cl = ws.cell(row=r, column=2, value=lbl)
        cl.fill = F(C['hdr_col']); cl.font = ft(True, '1F3864', 10)
        cl.alignment = al('left', 'center')
        cv = ws.cell(row=r, column=3, value=val)
        cv.font = ft(sz=10); cv.alignment = al('left', 'center')
        r += 1

    row_h(r, 12); r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    leg = ws.cell(row=r, column=2,
                  value='LEYENDA DEMORAS: verde ≤ umbral | amarillo ≤ 2× umbral | rojo > 2× umbral')
    leg.fill = F(C['hdr_col']); leg.font = ft(True, '1F3864', 9)
    leg.alignment = al('center', 'center')
    row_h(r, 18)
    r += 1
    ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
    nota = ws.cell(row=r, column=2,
                   value='NOTA: Un campo "Real" vacío = hito aún no alcanzado (proceso en curso), '
                         'no es una omisión. Sólo los 8 ítems de Suministros Críticos tienen '
                         'fechas reales registradas en el Plan.')
    nota.fill = F('FFF2CC'); nota.font = ft(False, '7F6000', 8)
    nota.alignment = al('left', 'center', wrap=True)
    row_h(r, 40)


def build_cambios(wb):
    ws = wb.create_sheet('CAMBIOS v3')
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 4
    ws.column_dimensions['B'].width = 18
    ws.column_dimensions['C'].width = 60

    ws.merge_cells('A1:C1')
    t = ws['A1']; t.value = f'REGISTRO DE CAMBIOS – Versión {VERSION}'
    t.fill = F(C['titulo']); t.font = ft(True, 'FFFFFF', 12)
    t.alignment = al('center', 'center'); ws.row_dimensions[1].height = 24

    cambios = [
        ('1', 'Scope ampliado',
         'Se incorporan TODOS los ítems del Plan de Suministros (Preliminar: 36 IN + 9 EL). '
         'v2 tenía datos hardcodeados parciales (22 IN + 16 EL).'),
        ('2', 'Items SC nuevos EL',
         'Se agregan ítems de Suministros Críticos no en Preliminar: '
         'SHELTER SE#3 (OC 4508944973/ABB), SE#4 (OC 4508944971/ABB), '
         'SE#5, SISTEMA PMS (OC 4508945953/ABB), TRANSFORMADORES, GENERADOR, DUCTO DE BARRAS.'),
        ('3', 'Items SC nuevos IN',
         'Se agregan: SISTEMA DE CONTROL PCS (Inauco nominado, SOLPED 23256882), '
         'SISTEMA DE SEGURIDAD SIS (HIMA nominado, SOLPED 23256883).'),
        ('4', 'N° RI precisos',
         'Se muestra el número RI en formato 5155-00-XXXX-IG-XX-RI-XXX '
         'según Datos P0 Suministros cruzado con descripción.'),
        ('5', 'AT por CT1/CT0',
         'Se separa el Análisis Técnico en CT1 (Rev.A – primera vuelta) '
         'y CT0 (Rev.0 – segunda vuelta / cierre), mostrando fechas P0 y reales.'),
        ('6', 'N° SOLPED',
         'Se extraen los números de SOLPED del campo "Estatus" de Suministros Críticos '
         '(ej.: 23359919/23359920/23359922 para Válvulas; 23392215/16/13 para Cables IN).'),
        ('7', 'Análisis demoras',
         'Nueva sección ANÁLISIS DEMORAS: días entre RI→Solped, Solped→RecOf, '
         'RecOf→AT, AT→OC, y total RI→OC. Código de color: verde/amarillo/rojo.'),
        ('8', 'OC/KOM EL',
         'Para SE#3/SE#4/PMS ya adjudicados a ABB, se muestra fecha KOM '
         '(18/5 y 22/5/2026) como fecha de inicio de fabricación.'),
        ('9', 'Fuente de datos',
         'El tracker lee directamente del Plan de Suministros (210526). '
         'No hay datos hardcodeados – actualización automática al cambiar el plan.'),
        ('10', 'FIX FECHAS – R1',
         'CORRECCIÓN v3R1: Las columnas "Prog. Actual" en hojas IN y EL ahora usan '
         'fechas del CRONO 4.11 (2025-2026), no del Preliminar P0 (2024). '
         'La columna "P0 Base (2024)" conserva las fechas originales P0. '
         'Fuentes: P0 Base=Preliminar | Prog.Actual=Crono 4.11 | Real=Suministros Críticos. '
         'Se filtran fechas inválidas 1900-era provenientes de celdas vacías en Excel.'),
        ('11', 'REVISIÓN FECHAS REAL – R2',
         'v3R2: Revisión integral de los campos "Real" (RI/SOLPED/Rec.Of./AT/OC). '
         'Se completan fechas reales extraídas del TEXTO del campo Estatus de Suministros '
         'Críticos cuando hay evidencia explícita: '
         '• Válvulas Control → Rec.Ofertas Real 21/05/26 ("Apertura de ofertas 21/5"). '
         '• Sistema SIS → AT CT1 Real 13/03/26 ("El 13/3 se envió a AT"). '
         '• SE#3/SE#4 → KOM 18/05/26 | PMS → KOM 22/05/26 (inicio de fabricación). '
         'Los SE#3/SE#4/PMS muestran N° OC (adjudicados ABB) + KOM en columna OC Real/KOM.'),
        ('12', 'Criterio campos vacíos',
         'IMPORTANTE: Un campo "Real" VACÍO significa que ese hito AÚN NO se alcanzó '
         '(proceso en curso), NO es una omisión. El Plan de Suministros sólo registra '
         'fechas reales para los 8 ítems de "Suministros Críticos"; el resto de ítems '
         '(instrumentos por tipo, paquetes EL) no tienen ejecución real registrada todavía '
         'y muestran únicamente fechas de programa (Prog. Actual / Crono 4.11).'),
        ('13', 'SOLPED EN LIBERACIÓN – R3',
         'v3R3: Corrección del estado SOLPED para "Cables IN" (23392215/16/13): '
         'el texto indica "en liberación" (trámite en curso, NO emitida aún). '
         'Se diferencia explícitamente de "liberada" (hecho consumado). '
         'Nuevo estado EN LIBERACIÓN (ámbar oscuro) en columna Estado SOLPED. '
         'Se mejoran las Acciones para todos los ítems SC con lógica más precisa '
         '(seguimiento KOM, completar AT, apertura ofertas, confirmar SOLPED en SAP).'),
        ('14', 'ACTUALIZACIÓN A v4/v5/v6 – Historial fuentes',
         'v4: Plan 290526 + hoja CAMBIOS SEMANA. '
         'v5: Plan 040626 (04/06/2026), comparación 290526→040626. '
         'v6: Plan 120626 (12/06/2026), comparación 040626→120626.'),
        ('15', 'CAMBIOS DETECTADOS 040626 → 120626',
         'PCS: AT CERRADO – aguarda oferta comercial Inauco (HITO MAYOR). '
         'SIS: AT CERRADO – cierre comercial, SBL solicitado, pendiente oferta servicios HIMA-Arg (HITO MAYOR). '
         'Cables IN: Ofertas enviadas a AT el 12/06 (hito nuevo). '
         'SE#3/SE#4/PMS: Fechas entrega OC y NecEnt incorporadas al plan. '
         'Válvulas: Plazo ING 10/06 vencido sin respuesta – riesgo de demora.'),
        ('16', 'ACTUALIZACIÓN A v7 (190626)',
         'v7: Plan 180626 (18/06/2026), comparación 120626→180626. '
         'Avance IN: 3 nuevas OCs (1→4), AT 21→30, petición 18→6. '
         'PCS/SIS AT cerrado 05/06 (122 d). Cables IN SOLPED 30/05.'),
        ('17', 'ACTUALIZACIÓN A v8 (260626) – plan 250626',
         'v8: Plan 250626 (25/06/2026), comparación 180626→250626. '
         'NOVEDADES PIPELINE (Cuadro resumen) — SOLPED en proceso IN+EL: 5 (=); '
         'OCs nuevas: +1 en Instrumentación (IN 4→5), total IN+EL 10→11; '
         '+1 RI EL emitida (14→15); avances a AT: EL 2→5, IN 30→31; '
         'ofertas en gestión (petición): EL 2 · IN 4. '
         'CRÍTICO: SIS NecOC 21/06 VENCIDA al 26/06 (AT cerrado 05/06, 143 d sin OC). '
         'Nueva sección "Pipeline agregado" en hoja CAMBIOS SEMANA.'),
        ('25', 'REVISIÓN 12 (280826)',
         'Rev12: Plan 280826 (28/08/2026). Comparación 210826 → 280826. '
         'GRAN SEMANA DE OCs EN CABLES: Cables Eléctricos OC COLOCADA a Marlew (4509059424, 23/08); '
         'Cables Instrumentación OC ENVIADA (4509060019, 27/08). Ambas con FE parcial (material en stock + '
         'plazos de hasta 14-16 semanas); en 2da instancia se licitan parciales adicionales (20.000m en IN). '
         '+5 RIs emitidas (EL 15→18, IN 46→48); OCs IN+EL 24→25. Sin emitir 8→3. '
         'Shelters SE#3/SE#4 y PMS: OCs ADICIONALES enviadas (SE#3 adicional CCM + ductos de barras 2da instancia; PMS maqueta de pruebas). '
         'SIS: KOM efectuada el 14/8, codificando LD. PCS: FAT 29/3-26/4, reunión con PP el 28/8, servicios en 2da instancia. '
         'VÁLVULAS (único crítico sin OC, 127 d): se avanzará con FLETE AÉREO (160 d India + 30/40 Comex); autorreguladoras a recotizar.'),
        ('24', 'REVISIÓN 11 (210826)',
         'Rev11: Plan 210826 (21/08/2026). Comparación 140826 → 210826. '
         'IN: +1 RI emitida (49→50 RIs; emitidas 45→46). OCs IN+EL estables (24). '
         'Cables Eléctricos: LEGAJO APROBADO POR EL CLIENTE – OC EN PROCESO DE EMISIÓN (próxima OC). '
         'Cables Instrumentación: en proceso de presentar el legajo para aprobación del cliente. '
         'Válvulas de control: pendiente la revisión final de ofertas por parte de ING con los comentarios de PP; '
         'autorreguladoras siguen pendientes de recotización. Sin OC (120 d). '
         'PCS (Inauco): pendiente enviar diagramas de JB al proveedor.'),
        ('23', 'REVISIÓN 10 (140826)',
         'Rev10: Plan 140826 (14/08/2026). Comparación 070826 → 140826. '
         'FUERTE AVANCE EN OCs: Instrumentación 10 → 16 (+6 OCs, ítems bulk); total OCs IN+EL 18 → 24. '
         'AT IN 33 → 28 (ítems pasaron a OC). Cables Instrumentación: AT CERRADO, el 12/08 se emite revisión de RI con ajuste de cantidades. '
         'SIS (HIMA): KOM en coordinación para el 14/8 (sostiene entrega dic-26). '
         'PCS (Inauco): se envía info crítica al proveedor (P&IDs, listado de señales) para avanzar; LD en confección. '
         'Cables EL: AT cerrado – negociación final en proceso. '
         'PUNTO CRÍTICO: Válvulas de control – PP tiene comentarios sobre el AT (a discutir con ING); autorreguladoras siguen pendientes de recotización. Sin OC (113 d).'),
        ('22', 'REVISIÓN 9 (070826)',
         'Rev9: Plan 070826 (07/08/2026). Comparación 300726 → 070826. '
         'Semana de fuerte avance en AT: EL petición 3→0 y AT 4→7; IN petición 4→2 y AT 30→33 (total AT IN+EL 34→40). OCs IN+EL estables (18). '
         'PCS (Inauco): OC de materiales 4509023773 EMITIDA (23/07) – KOM el 7/8, LD en confección. '
         'SIS (HIMA): OC ACEPTADA, se recibe proforma y se gestiona SBL – KOM semana del 10/8 (entrega dic-26). '
         'VÁLVULAS: AT CERRADO para válvulas de control, avanza cierre comercial; autorreguladoras deben RECOTIZARSE por circular ING (nuevo punto). '
         'Cables EL: AT CERRADO 04/08 (aprovecha sobrantes CPF 1.5; revisar RI cables 120mm2). '
         'Shelters SE#3/SE#4: Certificado n°2 emitido (fabricación avanza). PMS: OC por cambio de alcance en emisión.'),
        ('21', 'REVISIÓN 8 (300726)',
         'Rev8: Plan 300726 (30/07/2026). Comparación 240726 → 300726. '
         'IN: +1 RI emitida (48→49 RIs, ahora en SOLPED); emitidas 44→45. OCs IN+EL sin cambio (18). '
         'RESUELTO: Shelter SE#3 (ABB) – OCA USD 125k APROBADA, pendiente de emisión de OC (decisión del 27/7 cerrada). '
         'PMS: OC en proceso de emisión por cambio de alcance (estudios en la provisión). '
         'PCS: KOM primera semana de agosto. SIS: OC emitida, faltan cerrar puntos por PP. '
         'VÁLVULAS: NecOC 26/07 VENCIDA al corte 30/07 (98 d sin OC) – en julio debe salir el AT para adjudicar al menos un parcial (RIESGO ALTO).'),
        ('20', 'REVISIÓN 7 (240726)',
         'Rev7: Plan 240726 (24/07/2026). Comparación 170726 → 240726. '
         'BALANCE IN+EL: +3 OCs (Instrumentación 7→10, ítems bulk); total OCs IN+EL 15→18. '
         'AT: IN 33→30, EL 5→4 (ítems avanzaron a OC). Petición EL 2→3. '
         'CRÍTICOS: Cables EL – ofertas enviadas a AT el 22/7 (avanzó a AT). '
         'Cables IN – esta semana saldría el AT + revisión de RI con ajuste de cantidades. '
         'Shelter SE#3 (ABB): OCA con impacto USD 125k requiere aprobación antes del 27/7 (DECISIÓN). '
         'VÁLVULAS: 92 d sin OC, NecOC 26/07 (2 d) – reuniones Emerson/Baker Hughes, circular n°2 (RIESGO ALTO). '
         'PCS: OC materiales en proceso; servicios de asistencia en campo en discusión con PP.'),
        ('18', 'REVISIÓN 6 (170726) – Evolución 3 semanas',
         'Rev6: Plan 170726 (17/07/2026). Evaluación de avances de 3 semanas: cortes '
         '26/06 → 03/07 → 10/07 → 17/07. Hoja CAMBIOS SEMANA convertida a TENDENCIA de 4 cortes. '
         'BALANCE: +4 OCs IN+EL (11→15: IN 5→7, EL 6→8); SOLPED 5→0 (liberadas); AT 36→38. '
         'RESUELTO: SIS OC EMITIDA 4509010341 (crítico NecOC 21/06 de la Rev5). '
         'PCS: OC materiales en proceso, servicios en discusión con PP. '
         'Shelters SE#3/SE#4: en fabricación (Certif. n°1 emitido). Cables EL: apertura ofertas 20/07. '
         'ESTANCADO: Válvulas de Control – sin avance en 3 semanas, 85 d sin OC, NecOC 26/07 (riesgo alto).'),
        ('19', 'Revisión nombre',
         f'Archivo: Tracker_Suministros_IN_EL_LaCalera_II_{VERSION}.xlsx'),
    ]

    for n, tipo, desc in cambios:
        r = int(n) + 1
        ws.row_dimensions[r].height = 50
        c1 = ws.cell(row=r, column=1, value=n)
        c1.fill = F(C['hdr_col']); c1.font = ft(True, sz=9)
        c1.alignment = al('center'); c1.border = bd()
        c2 = ws.cell(row=r, column=2, value=tipo)
        c2.fill = F(C['hdr_col']); c2.font = ft(True, sz=9)
        c2.alignment = al('left', 'center', wrap=True); c2.border = bd()
        c3 = ws.cell(row=r, column=3, value=desc)
        c3.fill = F('FFFFFF'); c3.font = ft(sz=9)
        c3.alignment = al('left', 'center', wrap=True); c3.border = bd()

    ws.freeze_panes = 'A2'


# ── MAIN ────────────────────────────────────────────────────────────────────
def main():
    print(f'[1/7] Leyendo Plan 280826 (actual, Rev12): {os.path.basename(PLAN_FILE)}')
    sc_in, sc_el, in_items, el_items, crono_items = load_plan()
    print(f'      IN Preliminar: {len(in_items)} items | EL Preliminar: {len(el_items)} items')
    print(f'      IN Suministros Críticos: {len(sc_in)} | EL: {len(sc_el)}')
    print(f'      Crono 4.11 paquetes: {len(crono_items)} (IN+EL)')

    print(f'[2/7] Leyendo Plan 210826 (corte anterior, para comparación): {os.path.basename(PLAN_FILE_OLD)}')
    sc_old = load_sc_from_file(PLAN_FILE_OLD)
    print(f'      SC plan anterior: {len(sc_old)} ítems')
    print(f'[2b/7] Leyendo evolución 3 semanas (Cuadro resumen)...')
    trend = []
    for lbl, fname in PLAN_TREND:
        fpath = os.path.join(SCRIPT_DIR, '..', 'info_suministros', fname)
        trend.append((lbl, load_cuadro(fpath)))
        print(f'      {lbl}: EL={trend[-1][1].get("EL")} IN={trend[-1][1].get("IN")}')

    matched_in = sum(1 for i in in_items if find_crono_item(i['desc'], 'IN', crono_items))
    matched_el = sum(1 for i in el_items if find_crono_item(i['desc'], 'EL', crono_items))
    print(f'      Crono match IN: {matched_in}/{len(in_items)} | EL: {matched_el}/{len(el_items)}')

    print('[3/7] Creando workbook...')
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    print('[4/7] Generando hoja CAMBIOS SEMANA (evolución 210826 → 280826)...')
    sc_new_all = sc_in + sc_el
    build_cambios_semana(wb, sc_old, sc_new_all, trend)

    print('[5/7] Generando hoja CAMBIOS v3 (historial de versiones)...')
    build_cambios(wb)

    print('[5b/7] Generando PORTADA...')
    build_portada(wb)

    print('[5c/7] Generando DASHBOARD...')
    build_dashboard(wb, sc_in, sc_el, in_items, el_items)

    print('[5d/7] Generando hoja IN...')
    build_detail_sheet(wb, 'IN', 'IN – INSTRUMENTACIÓN & CONTROL',
                       in_items, sc_in, C['hdr_in'], IN_RI_NUMBER, crono_items)

    print('[5e/7] Generando hoja EL...')
    build_detail_sheet(wb, 'EL', 'EL – ELECTRICIDAD',
                       el_items, sc_el, C['hdr_el'], EL_RI_NUMBER, crono_items)

    print(f'[6/7] Guardando: {OUT_XLSX}')
    wb.save(OUT_XLSX)
    size = os.path.getsize(OUT_XLSX) // 1024
    print(f'      OK → {os.path.basename(OUT_XLSX)} ({size} KB)')
    print()
    print('Resumen:')
    print(f'  - IN total filas: {len(in_items)} (Preliminar) + SC adicionales')
    print(f'  - EL total filas: {len(el_items)} (Preliminar) + SC adicionales')
    print(f'  - Items EL con OC: {sum(1 for s in sc_el if s["oc_n"])}')
    print(f'  - Items IN con OC: {sum(1 for s in sc_in if s["oc_n"])}')

if __name__ == '__main__':
    main()
