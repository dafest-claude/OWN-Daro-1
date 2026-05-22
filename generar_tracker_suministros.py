#!/usr/bin/env python3
"""
Generador Tracker de Suministros – La Calera II CPF2
Especialidades: Instrumentación & Control (IN) | Electricidad (EL)
Contraste con Plan de Suministros P0 (2026.04.06)
Una fila por requisición para máximo detalle de análisis.
"""

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_DATE_DDMMYY
from datetime import date, datetime
import os

# ─── PALETA DE COLORES ────────────────────────────────────────────────────────
C = {
    'titulo':        '1F3864',
    'subtitulo':     '2F5496',
    'header_in':     '843C0C',   # naranja oscuro IN
    'header_el':     '375623',   # verde oscuro EL
    'header_grup':   '4472C4',   # azul grupo columnas
    'header_col':    'BDD7EE',   # azul claro nombre columna
    'completado':    '70AD47',
    'en_proceso':    'FFD966',
    'atrasado':      'FF0000',
    'pendiente':     'F2F2F2',
    'no_aplica':     '4472C4',
    'en_transito':   '00B0F0',
    'aprobado':      '92D050',
    'emitida':       'FFD966',
    'row_alt':       'DDEEFF',
    'row_normal':    'FFFFFF',
    'crit_lc':       'FF7F27',   # LLI/Long Lead Item
    'crit_cc':       'FF0000',   # C.CRÍTICO
    'crit_m':        'FFC000',   # MONTO
    'crit_h2':       '4472C4',   # HITO 2
    'crit_med':      'E2EFDA',   # MEDIA
    'plan_ref':      'FFF2CC',   # referencia plan P0
    'border_col':    '8EA9C1',
}

STATUS_MAP = {
    'COMPLETADO':   (C['completado'],  'FFFFFF', True,  '🟢'),
    'COMPLETADA':   (C['completado'],  'FFFFFF', True,  '🟢'),
    'EN PROCESO':   (C['en_proceso'],  '000000', True,  '🟡'),
    'EMITIDA':      (C['en_proceso'],  '000000', True,  '🟡'),
    'APROBADO':     (C['aprobado'],    'FFFFFF', True,  '🟢'),
    'EN TRÁNSITO':  (C['en_transito'], 'FFFFFF', True,  '🚛'),
    'ATRASADO':     (C['atrasado'],    'FFFFFF', True,  '🔴'),
    'PENDIENTE':    (C['pendiente'],   '000000', False, '⚪'),
    'NO APLICA':    (C['no_aplica'],   'FFFFFF', True,  '🔵'),
    'SIN RI':       ('FF4500',         'FFFFFF', True,  '⛔'),
}

CRIT_COLOR = {
    'LLI':        C['crit_lc'],
    'C. CRÍTICO': C['crit_cc'],
    'MONTO':      C['crit_m'],
    'HITO 2':     C['crit_h2'],
    'MEDIA':      C['crit_med'],
    'MOVILIZACIÓN': 'E2EFDA',
    'SERV':       'DDEBF7',
}


def F(col): return PatternFill('solid', fgColor=col)
def ft(bold=False, color='000000', sz=10):
    return Font(bold=bold, color=color, size=sz, name='Calibri')
def al(h='center', v='center', wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def bd(color=C['border_col']):
    s = Side(style='thin', color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def bd_m():
    s = Side(style='medium', color='000000')
    return Border(left=s, right=s, top=s, bottom=s)


def sc(cell, status, short=False):
    """Apply status coloring to a cell."""
    s = status.strip().upper()
    if s in STATUS_MAP:
        bg, fg, bold, emoji = STATUS_MAP[s]
        cell.value = f"{emoji} {status}" if not short else emoji
        cell.fill = F(bg)
        cell.font = ft(bold=bold, color=fg, sz=9)
    else:
        cell.value = status
        cell.font = ft(sz=9)
    cell.alignment = al(h='center', wrap=False)
    cell.border = bd()


def fmt_date(d):
    """Format date string dd-mmm-yy."""
    if not d:
        return ''
    if isinstance(d, str):
        try:
            dt = datetime.strptime(d, '%Y-%m-%d')
            return dt.strftime('%d-%b-%y').upper()
        except Exception:
            return d
    return ''


def write_date(ws, row, col, d, planned=False):
    c = ws.cell(row=row, column=col)
    c.value = fmt_date(d)
    c.fill = F(C['plan_ref']) if planned else F(C['row_normal'])
    c.font = ft(sz=9, color='595959' if planned else '000000')
    c.alignment = al(h='center')
    c.border = bd()
    return c


def desvio_text(d):
    if d is None:
        return ''
    if d > 0:
        return f'+{d}d'
    elif d < 0:
        return f'{d}d'
    return '0d'


def write_desvio(ws, row, col, days):
    c = ws.cell(row=row, column=col)
    if days is None:
        c.value = ''
        c.fill = F(C['pendiente'])
    elif days > 30:
        c.value = f'+{days}d'
        c.fill = F(C['atrasado'])
        c.font = ft(bold=True, color='FFFFFF', sz=9)
    elif days > 0:
        c.value = f'+{days}d'
        c.fill = F('FFC000')
        c.font = ft(bold=True, color='000000', sz=9)
    elif days < 0:
        c.value = f'{days}d'
        c.fill = F(C['completado'])
        c.font = ft(bold=True, color='FFFFFF', sz=9)
    else:
        c.value = '0d'
        c.fill = F(C['completado'])
        c.font = ft(bold=True, color='FFFFFF', sz=9)
    c.alignment = al(h='center')
    c.border = bd()


# ─── DATOS IN ────────────────────────────────────────────────────────────────
def get_in_data():
    """
    Una fila por requisición de ingeniería (RI).
    Campos: n, crit, n_ri, desc, ref_plan, ref_crit,
            ri_prog, ri_real, desv_ri, est_ri,
            solped, solp_prog, solp_real, est_solp,
            recof_prog, recof_real, est_recof,
            at_prog, at_real, est_at,
            n_oc, prov, nec_oc, oc_real, lt, est_oc,
            ent_oc, nec_obra, desv_ent, est_ent,
            est_gen, obs, accion
    Fechas en formato 'YYYY-MM-DD', None si no disponible.
    """
    return [
        # ── VÁLVULAS ──────────────────────────────────────────────────────────
        dict(n=1, crit='LLI', n_ri='RI-001',
             desc='Válvulas de Control y Autorreguladoras',
             ref_plan='IN – Válvulas de Control / Autorreguladoras',
             ref_crit='RI – VÁLVULAS DE CONTROL Y AUTORREGULADORAS',
             ri_prog='2026-02-20', ri_real='2026-04-23', desv_ri=62, est_ri='COMPLETADO',
             solped='23359919 / 23359920 / 23359922',
             solp_prog='2026-03-15', solp_real='2026-05-05', est_solp='EMITIDA',
             recof_prog='2026-04-15', recof_real='2026-05-19', est_recof='EN PROCESO',
             at_prog='2026-06-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-07-26', oc_real=None, lt=300, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-05-31', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI-001 emitida 23-Abr (Val.Control/Autoreg.). Apertura ofertas en curso fc 19-May-26. AT fc Ago-26. '
                 'Plan P0: nec.OC 26-Jul-26, LT 300 días, nec.obra 31-May-27.',
             accion='Cerrar Rec.Ofertas fc 19-May. AT fc Ago-26. Necesidad OC: 26-Jul-26.'),

        dict(n=2, crit='LLI', n_ri='RI-002',
             desc='Válvulas BlowDown y ShutDown',
             ref_plan='IN – Válvulas BlowDown / ShutDown',
             ref_crit='',
             ri_prog='2026-03-01', ri_real='2026-03-19', desv_ri=18, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-01', solp_real=None, est_solp='EN PROCESO',
             recof_prog='2026-05-15', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-08-01', oc_real=None, lt=270, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-05-31', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI-003 emitida 19-Mar (leve desvío sobre P0). SOLPED en proceso. '
                 'Rec.Ofertas fc May-Jun. AT fc Jul-26.',
             accion='Emitir SOLPED. Cerrar Rec.Of fc May-26. AT fc Jul-26.'),

        dict(n=3, crit='LLI', n_ri='RI-003',
             desc='Válvulas de Seguridad (PSV)',
             ref_plan='IN – Válvulas de Seguridad',
             ref_crit='',
             ri_prog='2026-02-15', ri_real='2026-03-30', desv_ri=43, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-01', solp_real='2026-04-15', est_solp='EMITIDA',
             recof_prog='2026-04-20', recof_real='2026-04-30', est_recof='COMPLETADA',
             at_prog='2026-06-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-07-15', oc_real=None, lt=None, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-01-15', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI emitida 30-Mar. Rec.Ofertas completada 30-Abr. AT pendiente fc Jul-26. OC fc Ene-27.',
             accion='Completar AT fc Jul-26. OC prevista Ene-27.'),

        # ── INSTRUMENTOS – TRANSMISORES / MEDICIÓN ────────────────────────────
        dict(n=4, crit='MEDIA', n_ri='RI-004',
             desc='Transmisores de Presión y Presión Diferencial',
             ref_plan='IN – Transmisor de Presión / Transmisor de Presión Diferencial',
             ref_crit='',
             ri_prog='2026-03-01', ri_real='2026-04-10', desv_ri=40, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-20', solp_real=None, est_solp='EN PROCESO',
             recof_prog='2026-05-15', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-08-15', oc_real=None, lt=180, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RIs completadas (T.Pres y T.Pres.Dif) emitidas 10-Abr. SOLPED en proceso. '
                 'Rec.Ofertas fc May-Jun.',
             accion='SOLPED fc May-26. Rec.Of fc Jun-26. AT fc Jul-26.'),

        dict(n=5, crit='MEDIA', n_ri='RI-005',
             desc='Transmisores de Nivel (DP Cell y Radar)',
             ref_plan='IN – Transmisor de Nivel Dp Cell / Transmisor de Nivel Radar / Transmisor de Nivel',
             ref_crit='',
             ri_prog='2026-03-10', ri_real='2026-04-15', desv_ri=36, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-25', solp_real=None, est_solp='EN PROCESO',
             recof_prog='2026-05-20', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-08-15', oc_real=None, lt=180, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RIs completadas (T.Nivel Radar, DP Cell). Emitidas 15-Abr. SOLPED en proceso. '
                 'Rec.Ofertas fc May-Jun.',
             accion='SOLPED fc May-26. Rec.Of fc Jun-26.'),

        dict(n=6, crit='MEDIA', n_ri='RI-006',
             desc='Transmisores de Temperatura',
             ref_plan='IN – Transmisor de Temperatura',
             ref_crit='',
             ri_prog='2026-03-10', ri_real='2026-04-20', desv_ri=41, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-05-01', solp_real=None, est_solp='EN PROCESO',
             recof_prog='2026-05-25', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-09-01', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RIs completadas (T.Temp) emitidas ~20-Abr. SOLPED en proceso fc May. '
                 'Rec.Ofertas fc May-Jun.',
             accion='SOLPED fc May-26. Rec.Of fc Jun-26.'),

        dict(n=7, crit='MEDIA', n_ri='RI-007',
             desc='Manómetros, Manóm.Dif., Indicadores de Nivel y Rotámetros',
             ref_plan='IN – Manómetros y Manómetros Diferenciales / Indicador de Nivel / Rotámetro',
             ref_crit='',
             ri_prog='2026-03-15', ri_real='2026-04-20', desv_ri=36, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-05-01', solp_real=None, est_solp='EN PROCESO',
             recof_prog='2026-05-25', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-09-01', oc_real=None, lt=120, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RIs (Manóm., Ind.Nivel) emitidas 10-24 Abr. SOLPED en proceso. Rec.Of fc May-Jun.',
             accion='SOLPED fc May-26. Rec.Of fc Jun-26.'),

        dict(n=8, crit='MEDIA', n_ri='RI-008',
             desc='Interruptores de Presión, Nivel y Vibración',
             ref_plan='IN – Interruptor de Presión / Interruptor de Nivel / Interruptor de Vibración',
             ref_crit='',
             ri_prog='2026-03-15', ri_real='2026-04-20', desv_ri=36, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-05-01', solp_real=None, est_solp='EN PROCESO',
             recof_prog='2026-06-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-09-01', oc_real=None, lt=120, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RIs completadas (Interruptores P, N, Vib). Emitidas ~20 Abr. SOLPED en proceso.',
             accion='SOLPED fc May-26. Rec.Of fc Jun-26.'),

        dict(n=9, crit='MEDIA', n_ri='RI-009',
             desc='Medidores de Caudal (Magnético, Coriolis, Ultrasónico)',
             ref_plan='IN – Transmisor de Caudal Coriolis / Multivariable / Presión Diferencial',
             ref_crit='',
             ri_prog='2026-03-15', ri_real='2026-04-22', desv_ri=38, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-25', solp_real='2026-04-22', est_solp='EMITIDA',
             recof_prog='2026-04-22', recof_real='2026-04-22', est_recof='COMPLETADA',
             at_prog='2026-06-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-10-01', oc_real=None, lt=120, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-15', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI y Rec.Ofertas completadas 22-Abr (Caudal Mag, Coriolis, Ultrasónico). '
                 'AT fc Jun-26. OC Oct-Nov-26. Entrega Dic-26.',
             accion='AT fc Jun-26. OC fc Oct-26. Entrega fc Dic-26.'),

        dict(n=10, crit='MEDIA', n_ri='RI-010',
             desc='Termómetros, Cupones de Corrosión, Placas Orificio, Indicador Paso Scrapper',
             ref_plan='IN – Termómetro / Cupón de Corrosión / Placa Orificio / Ind.Paso Scrapper',
             ref_crit='',
             ri_prog='2026-04-01', ri_real='2026-04-20', desv_ri=19, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-05-01', solp_real=None, est_solp='EN PROCESO',
             recof_prog='2026-06-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-08-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-09-01', oc_real=None, lt=90, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RIs completadas. SOLPED en proceso. Sin ítems de largo lead time. Holgura disponible.',
             accion='SOLPED fc May-26. Sin urgencia crítica.'),

        dict(n=11, crit='MEDIA', n_ri='RI-011',
             desc='Analizadores (RVP y Corte de Agua)',
             ref_plan='IN – Analizador',
             ref_crit='',
             ri_prog='2026-03-15', ri_real='2026-04-01', desv_ri=17, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-10', solp_real='2026-04-17', est_solp='EMITIDA',
             recof_prog='2026-04-17', recof_real='2026-04-17', est_recof='COMPLETADA',
             at_prog='2026-06-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-07-01', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI y Rec.Ofertas completadas 17-Abr. AT fc Jun-26. OC/Entrega 2do sem 2026.',
             accion='AT fc Jun-26. OC y entrega 2do semestre 2026.'),

        dict(n=12, crit='MEDIA', n_ri='RI-012',
             desc='Tomamuestras y Detectores de Fugas',
             ref_plan='IN – Detector de Fugas',
             ref_crit='',
             ri_prog='2026-04-01', ri_real='2026-03-30', desv_ri=-2, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-15', solp_real=None, est_solp='EN PROCESO',
             recof_prog='2026-05-15', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-08-01', oc_real=None, lt=90, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-11-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI emitida 30-Mar (2 días antes del prog). SOLPED en proceso fc May. '
                 'Entrega fc Nov-26. Holgura ok.',
             accion='SOLPED fc May-26. Sin urgencia crítica.'),

        # ── SISTEMA F&G / CCI ─────────────────────────────────────────────────
        dict(n=13, crit='MEDIA', n_ri='RI-013',
             desc='Sistema Fire & Gas (FGS) – Detectores de Llama y Mezcla Explosiva',
             ref_plan='IN – Detector de Llama / Detector de Mezcla Explosiva',
             ref_crit='',
             ri_prog='2026-03-01', ri_real='2026-04-15', desv_ri=45, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-20', solp_real='2026-04-25', est_solp='EMITIDA',
             recof_prog='2026-04-30', recof_real='2026-04-30', est_recof='COMPLETADA',
             at_prog='2026-06-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-07-15', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI F&G y Detección Térmica Lineal completadas. Rec.Ofertas 30-Abr. '
                 'AT fc Jul-26. OC/Entrega Dic-26.',
             accion='AT fc Jul-26. OC y entrega previstas Dic-26.'),

        dict(n=14, crit='MEDIA', n_ri='RI-014',
             desc='Sistema Contra Incendio (CCI) – Válvulas Diluvio y Pulsadores',
             ref_plan='IN – Válvulas Diluvio / Pulsador de Emergencia',
             ref_crit='',
             ri_prog='2026-03-15', ri_real='2026-04-20', desv_ri=36, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-04-25', solp_real='2026-05-01', est_solp='EMITIDA',
             recof_prog='2026-05-06', recof_real='2026-05-06', est_recof='COMPLETADA',
             at_prog='2026-06-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-07-15', oc_real=None, lt=180, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-01-15', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI y Rec.Ofertas (Val.Diluvio) completadas 06-May. AT fc Jul-26. OC/Entrega Ene-27.',
             accion='AT fc Jul-26. OC/Entrega previstas Ene-27.'),

        # ── SISTEMAS DIGITALES ────────────────────────────────────────────────
        dict(n=15, crit='LLI', n_ri='RI-015',
             desc='Sistema de Control PCS (DCS/RTU/PLC)',
             ref_plan='IN – Sistema de Control PCS',
             ref_crit='RI – SISTEMA DE CONTROL PCS (Solped 23256882)',
             ri_prog='2026-02-03', ri_real='2026-02-03', desv_ri=0, est_ri='COMPLETADO',
             solped='23256882',
             solp_prog='2026-02-15', solp_real='2026-02-19', est_solp='EMITIDA',
             recof_prog='2026-03-15', recof_real='2026-03-25', est_recof='COMPLETADA',
             at_prog='2026-06-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='Inauco (nominado)',
             nec_oc='2026-08-19', oc_real=None, lt=300, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-06-24', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='PCS (Inauco nominado): RI 03-Feb, Rec.Of 25-Mar completadas. AT fc Jul-26. '
                 'OC fc Feb-27. Plan P0: nec.OC 19-Ago-26, LT 300 días, nec.obra 24-Jun-27. '
                 'Lead time crítico para RFSU Feb-27.',
             accion='⚠ Cerrar AT urgente fc Jul-26. OC CRÍTICA nec. antes 19-Ago-26. '
                    'LT 300d: sin OC en Ago no hay entrega en Jun-27. RFSU en riesgo.'),

        dict(n=16, crit='LLI', n_ri='RI-016',
             desc='Sistema de Seguridad SIS',
             ref_plan='IN – Sistema de Seguridad SIS',
             ref_crit='RI – SISTEMA DE SEGURIDAD SIS (Solped 23256883)',
             ri_prog='2026-02-03', ri_real='2026-02-03', desv_ri=0, est_ri='COMPLETADO',
             solped='23256883',
             solp_prog='2026-02-15', solp_real='2026-02-19', est_solp='EMITIDA',
             recof_prog='2026-03-05', recof_real='2026-03-11', est_recof='COMPLETADA',
             at_prog='2026-05-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='HIMA (nominado)',
             nec_oc='2026-06-21', oc_real=None, lt=300, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-04-18', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='SIS (HIMA nominado): RI 03-Feb, Rec.Of 11-Mar completadas. AT fc Jul-26. '
                 'Plan P0: nec.OC 21-Jun-26, LT 300 días, nec.obra 18-Abr-27.',
             accion='⚠ OC CRÍTICA fc 21-Jun-26. Cerrar AT inmediato. '
                    'LT 300d: nec.obra 18-Abr-27 vs RFSU Feb-27. Acción urgente.'),

        dict(n=17, crit='MEDIA', n_ri='RI-017',
             desc='Sistema de Comunicaciones y CCTV',
             ref_plan='IN – Requerimiento para Sistema de Comunicaciones / CCTV',
             ref_crit='',
             ri_prog='2026-04-01', ri_real=None, est_ri='EN PROCESO',
             solped='',
             solp_prog='2026-05-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-06-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-08-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-09-01', oc_real=None, lt=180, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI en proceso. Sin datos de avance registrados. Pendiente emisión. '
                 'Verificar estado real con ingeniería.',
             accion='Confirmar estado real de RI. Emitir fc May-26.'),

        # ── CAJAS / CABLES / MATERIALES ───────────────────────────────────────
        dict(n=18, crit='MEDIA', n_ri='RI-018',
             desc='Cajas de Conexionado (J-Box) Instrumentación',
             ref_plan='IN – Cajas de Conexionado',
             ref_crit='',
             ri_prog='2025-07-01', ri_real=None, desv_ri=325, est_ri='EN PROCESO',
             solped='',
             solp_prog='2025-08-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2025-09-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2025-11-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-06-01', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-11-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI prog Jul-25. Tendencia actual fc May-26 (+10 meses sobre P0). '
                 'Desvío importante. Emisión pendiente. Entrega fc Nov-26.',
             accion='Emitir RI urgente. Tendencia May-26, desvío ~10 meses sobre P0.'),

        dict(n=19, crit='C. CRÍTICO', n_ri='RI-019',
             desc='Bandejas Portacables (instrumentación)',
             ref_plan='IN – Canalizaciones / Materiales Mecánicos',
             ref_crit='',
             ri_prog='2025-08-01', ri_real=None, desv_ri=294, est_ri='ATRASADO',
             solped='',
             solp_prog='2025-09-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2025-10-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2025-12-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-07-01', oc_real=None, lt=90, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='ATRASADO',
             obs='⚠ RI prog Ago-25 sin emisión registrada al 22-May-26. '
                 'Sin RI activa en el plan. Requiere verificación y emisión urgente.',
             accion='🚨 URGENTE: Verificar si hay RI activa. Emitir inmediatamente. '
                    'Desvío +294 días sobre P0.'),

        dict(n=20, crit='C. CRÍTICO', n_ri='RI-020',
             desc='Cables de Instrumentación – Rev.0 (Ingeniería Inicial)',
             ref_plan='IN – Cables de Instrumentos (Rev.0)',
             ref_crit='RI – CABLES INSTRUMENTACIÓN',
             ri_prog='2025-09-01', ri_real=None, desv_ri=263, est_ri='ATRASADO',
             solped='',
             solp_prog='2025-10-15', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2025-12-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-02-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-10-21', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-03-29', desv_ent=None, est_ent='ATRASADO',
             est_gen='ATRASADO',
             obs='⚠ RI prog Sep-25. Tendencia actual May-26 (+8 meses). '
                 'Plan P0: nec.OC 21-Oct-26, LT 150 días, nec.obra 29-Mar-27. '
                 'Dependiente plano de cables Rev.0 (ingeniería).',
             accion='🚨 Acelerar ingeniería plano cables Rev.0. Emitir RI fc May-26. '
                    'Nec.OC: 21-Oct-26. Desvío +263d sobre P0.'),

        dict(n=21, crit='MONTO', n_ri='RI-021',
             desc='Cables de Instrumentación – Rev.1 (Hito Ing.80%)',
             ref_plan='IN – Cables de Instrumentos (Rev.1)',
             ref_crit='RI – CABLES INSTRUMENTACIÓN',
             ri_prog='2026-05-15', ri_real=None, desv_ri=None, est_ri='PENDIENTE',
             solped='',
             solp_prog='2026-06-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-07-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-08-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-08-15', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='PENDIENTE',
             obs='Fase Rev.1 sujeta a hito Ing.80% (Jun-26). '
                 'Tendencia RI May-26. OC prevista Ago-26. Entrega Dic-26.',
             accion='Pendiente hito Ing.80% Jun-26. Monitorear avance ingeniería.'),

        dict(n=22, crit='C. CRÍTICO', n_ri='RI-022',
             desc='Materiales para Montaje de Instrumentos (bandejas, soportes, conduits IN)',
             ref_plan='IN – Materiales Mecánicos / Materiales Eléctricos (montaje)',
             ref_crit='',
             ri_prog='2025-06-01', ri_real=None, desv_ri=356, est_ri='ATRASADO',
             solped='',
             solp_prog='2025-07-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2025-08-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2025-10-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-11-01', oc_real=None, lt=90, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=480, est_ent='ATRASADO',
             est_gen='ATRASADO',
             obs='⚠ CRÍTICO: RI prog Jun-25 sin emisión al 22-May-26. '
                 'Tendencia actual Oct-26 (+16 meses sobre P0). '
                 'Revisar alcance y estrategia de compra urgente.',
             accion='🚨 CRÍTICO: Definir alcance completo y emitir RI. '
                    'Tendencia Oct-26, desvío +356 días sobre P0.'),
    ]


# ─── DATOS EL ────────────────────────────────────────────────────────────────
def get_el_data():
    return [
        # ── SHELTER ELÉCTRICO / PMS ───────────────────────────────────────────
        dict(n=1, crit='HITO 2', n_ri='RI-EL-001',
             desc='Shelter Eléctrico SE#4 con Equipamiento',
             ref_plan='EL – Shelter Eléctrico SE#4',
             ref_crit='RI – SHELTER ELECTRICO SE#4 CON EQUIPAMIENTO',
             ri_prog='2026-01-09', ri_real='2026-01-09', desv_ri=0, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-01-15', solp_real='2026-01-21', est_solp='EMITIDA',
             recof_prog='2026-02-01', recof_real='2026-02-09', est_recof='COMPLETADA',
             at_prog='2026-03-01', at_real='2026-04-01', est_at='APROBADO',
             n_oc='4508944971', prov='ABB',
             nec_oc='2026-05-20', oc_real='2026-05-14', lt=390, est_oc='EMITIDA',
             ent_oc='2027-06-07', nec_obra='2027-06-23', desv_ent=None, est_ent='EN PROCESO',
             est_gen='EN PROCESO',
             obs='OC 4508944971 adjudicada ABB. KOM realizada 18-May-26. '
                 'Plan P0: nec.OC 20-May-26, LT 390 días, nec.obra 23-Jun-27. '
                 'Fabricación en curso. Entrega contractual fc 7-Jun-27.',
             accion='Seguimiento fabricación ABB. KOM 18-May-26 realizada. '
                    'Confirmar hitos intermedios (Ing., Acopio, Fab.).'),

        dict(n=2, crit='HITO 2', n_ri='RI-EL-002',
             desc='Shelter Eléctrico SE#3 con Equipamiento',
             ref_plan='EL – Shelter Eléctrico SE#3',
             ref_crit='RI – SHELTER ELECTRICO SE#3 CON EQUIPAMIENTO',
             ri_prog='2026-01-09', ri_real='2026-01-09', desv_ri=0, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-01-15', solp_real='2026-01-21', est_solp='EMITIDA',
             recof_prog='2026-02-01', recof_real='2026-02-09', est_recof='COMPLETADA',
             at_prog='2026-03-01', at_real='2026-04-01', est_at='APROBADO',
             n_oc='4508944973', prov='ABB',
             nec_oc='2026-05-20', oc_real='2026-05-14', lt=360, est_oc='EMITIDA',
             ent_oc='2027-05-09', nec_obra='2027-05-24', desv_ent=None, est_ent='EN PROCESO',
             est_gen='EN PROCESO',
             obs='OC 4508944973 adjudicada ABB. KOM realizada 18-May-26. '
                 'Plan P0: nec.OC 20-May-26, LT 360 días, nec.obra 24-May-27. '
                 'Fabricación en curso.',
             accion='Seguimiento fabricación ABB. Confirmar % ingeniería y acopio.'),

        dict(n=3, crit='LLI', n_ri='RI-EL-003',
             desc='Sistema PMS (Power Management System)',
             ref_plan='EL – Sistema PMS',
             ref_crit='RI – SISTEMA PMS',
             ri_prog='2026-01-09', ri_real='2026-01-09', desv_ri=0, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-01-15', solp_real='2026-01-21', est_solp='EMITIDA',
             recof_prog='2026-02-01', recof_real='2026-02-09', est_recof='COMPLETADA',
             at_prog='2026-03-15', at_real='2026-04-10', est_at='APROBADO',
             n_oc='4508945953', prov='ABB',
             nec_oc='2026-05-29', oc_real='2026-05-18', lt=360, est_oc='EMITIDA',
             ent_oc='2027-05-13', nec_obra='2027-06-02', desv_ent=None, est_ent='EN PROCESO',
             est_gen='EN PROCESO',
             obs='OC 4508945953 adjudicada ABB. KOM realizada 18-May-26. '
                 'Plan P0: nec.OC 29-May-26, LT 360 días, nec.obra 02-Jun-27.',
             accion='Seguimiento fabricación ABB. Monitorear % avance.'),

        dict(n=4, crit='HITO 2', n_ri='RI-EL-004',
             desc='Shelter Eléctrico SE#5 con Equipamiento',
             ref_plan='EL – Shelter Eléctrico SE#5',
             ref_crit='',
             ri_prog='2026-02-01', ri_real=None, desv_ri=None, est_ri='SIN RI',
             solped='',
             solp_prog='2026-03-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-04-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-05-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-06-01', oc_real=None, lt=390, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-07-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='ATRASADO',
             obs='⚠ SE#5 sin RI localizada al 22-May-26. No identificado en cronograma ni '
                 'en suministros críticos. Requiere verificación urgente si aplica al proyecto.',
             accion='🚨 URGENTE: Confirmar si SE#5 aplica al alcance del proyecto. '
                    'Si aplica, emitir RI inmediatamente. LT ~390 días.'),

        # ── TABLEROS / CCM / MT ───────────────────────────────────────────────
        dict(n=5, crit='LLI', n_ri='RI-EL-005',
             desc='Tableros y Celdas Media Tensión (MT) + Transformadores',
             ref_plan='EL – Cables - Baja y media tensión / Canalizaciones y montaje',
             ref_crit='',
             ri_prog='2026-01-15', ri_real='2026-02-01', desv_ri=17, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-02-15', solp_real='2026-03-01', est_solp='EMITIDA',
             recof_prog='2026-04-15', recof_real='2026-04-30', est_recof='COMPLETADA',
             at_prog='2026-05-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2027-01-01', oc_real=None, lt=360, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-06-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='EN PROCESO',
             obs='RI Transf.Secos/Aceite y Eq.Maniobra MT completadas. Rec.Of 30-Abr completada. '
                 'AT pendiente. Lead time crítico ~12 meses. OC fc Ene-27.',
             accion='⚠ Cerrar AT urgente. LT ~12 meses: OC necesaria antes Ene-27 '
                    'para nec.obra Jun-27. Riesgo RFSU Feb-27.'),

        dict(n=6, crit='C. CRÍTICO', n_ri='RI-EL-006',
             desc='Tableros BT y Centros de Control de Motores (CCM)',
             ref_plan='EL – Canalizaciones y montaje',
             ref_crit='',
             ri_prog='2025-07-01', ri_real=None, desv_ri=325, est_ri='ATRASADO',
             solped='',
             solp_prog='2025-08-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2025-09-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2025-11-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-06-01', oc_real=None, lt=300, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=300, est_ent='ATRASADO',
             est_gen='ATRASADO',
             obs='⚠ RI prog Jul-25 sin emisión confirmada al 22-May-26 (+325 días). '
                 'Requiere verificación urgente. Lead time CCM ~10-12 meses.',
             accion='🚨 CRÍTICO: Verificar y emitir RI INMEDIATAMENTE. '
                    'LT 10-12 meses compromete RFSU Feb-27. Desvío +325d.'),

        dict(n=7, crit='C. CRÍTICO', n_ri='RI-EL-007',
             desc='Sistema UPS y Rectificadores',
             ref_plan='EL – Materiales Misceláneos',
             ref_crit='',
             ri_prog='2025-11-01', ri_real=None, desv_ri=202, est_ri='ATRASADO',
             solped='',
             solp_prog='2025-12-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-01-15', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-03-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-06-01', oc_real=None, lt=300, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='ATRASADO',
             obs='⚠ RI prog Nov-25 sin real confirmada al 22-May-26 (+202 días). '
                 'Equipo de largo lead time (~10 meses). Requiere RI y adjudicación urgente.',
             accion='🚨 ACCIÓN INMEDIATA: Emitir RI. LT ~10 meses. '
                    'Sin OC urgente, compromete RFSU Feb-27. Desvío +202d.'),

        dict(n=8, crit='C. CRÍTICO', n_ri='RI-EL-008',
             desc='Ducto de Barras',
             ref_plan='EL – Cables - Baja y media tensión',
             ref_crit='',
             ri_prog='2025-10-01', ri_real=None, desv_ri=233, est_ri='ATRASADO',
             solped='',
             solp_prog='2025-11-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-01-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-03-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-06-01', oc_real=None, lt=240, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-02-03', desv_ent=None, est_ent='PENDIENTE',
             est_gen='ATRASADO',
             obs='⚠ Sin RI emitida. No identificado en cronograma al 22-May-26. '
                 'Lead time ducto de barras 8-10 meses. Acción inmediata requerida.',
             accion='🚨 ACCIÓN INMEDIATA: Emitir RI. LT 8-10 meses. '
                    'Riesgo crítico para RFSU Feb-27. Desvío +233d sobre P0.'),

        # ── CABLES ────────────────────────────────────────────────────────────
        dict(n=9, crit='C. CRÍTICO', n_ri='RI-EL-009',
             desc='Cables Eléctricos BT/MT – Rev.0 (Ingeniería Inicial)',
             ref_plan='EL – Cables - Baja y media tensión (Rev.0)',
             ref_crit='RI – CABLES ELÉCTRICOS',
             ri_prog='2026-03-01', ri_real=None, desv_ri=82, est_ri='ATRASADO',
             solped='',
             solp_prog='2026-04-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-05-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-11-18', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2027-04-26', desv_ent=60, est_ent='ATRASADO',
             est_gen='ATRASADO',
             obs='⚠ RI prog Mar-26, tendencia Jun-26 (+82 días). '
                 'Plan P0: nec.OC 18-Nov-26, LT 150 días, nec.obra 26-Abr-27. '
                 'Entrega fc Feb-27 vs necesidad Dic-26. Desvío ~60 días. '
                 'Dependiente plano cables Rev.0.',
             accion='🚨 Emitir RI fc Jun-26. Nec.OC: 18-Nov-26. LT 150d. '
                    'Nec.obra: 26-Abr-27. Dependiente plano Rev.0.'),

        dict(n=10, crit='MONTO', n_ri='RI-EL-010',
             desc='Cables Eléctricos BT/MT – Rev.1 (Hito Ing.80%)',
             ref_plan='EL – Cables - Baja y media tensión (Rev.1)',
             ref_crit='RI – CABLES ELÉCTRICOS',
             ri_prog='2026-05-15', ri_real=None, desv_ri=None, est_ri='PENDIENTE',
             solped='',
             solp_prog='2026-06-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-07-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-08-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-08-01', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='PENDIENTE',
             obs='Fase Rev.1 sujeta a hito Ing.80% (Jun-26). OC fc Ago-26. Entrega Dic-26.',
             accion='Pendiente hito Ing.80% Jun-26. Monitorear avance ingeniería.'),

        # ── CANALIZACIONES ────────────────────────────────────────────────────
        dict(n=11, crit='C. CRÍTICO', n_ri='RI-EL-011',
             desc='Canalizaciones Eléctricas (conduits, bandejas y accesorios)',
             ref_plan='EL – Canalizaciones y montaje',
             ref_crit='',
             ri_prog='2026-02-01', ri_real=None, desv_ri=110, est_ri='ATRASADO',
             solped='',
             solp_prog='2026-03-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-04-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-05-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-07-01', oc_real=None, lt=120, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-10-01', desv_ent=90, est_ent='ATRASADO',
             est_gen='ATRASADO',
             obs='⚠ RI prog Feb-26, tendencia Jun-26 (+110 días / +4 meses). '
                 'Entrega fc Ene-27 vs necesidad Oct-26. Desvío ~90 días.',
             accion='🚨 ATRASADO: Emitir RI fc Jun-26. Entrega fc Ene-27 vs '
                    'necesidad Oct-26, desvío +90 días.'),

        # ── PAT / ILUMINACIÓN / PROTECCIÓN ────────────────────────────────────
        dict(n=12, crit='C. CRÍTICO', n_ri='RI-EL-012',
             desc='Materiales PAT y SPCDA (Puesta a Tierra y Pararrayos)',
             ref_plan='EL – Puesta a tierra y SPCDA',
             ref_crit='',
             ri_prog='2026-02-01', ri_real='2026-02-10', desv_ri=9, est_ri='COMPLETADO',
             solped='',
             solp_prog='2026-02-15', solp_real='2026-03-01', est_solp='EMITIDA',
             recof_prog='2026-03-15', recof_real='2026-04-01', est_recof='COMPLETADA',
             at_prog='2026-04-01', at_real='2026-04-10', est_at='APROBADO',
             n_oc='', prov='',
             nec_oc='2026-04-15', oc_real='2026-04-20', lt=60, est_oc='EMITIDA',
             ent_oc='2026-06-20', nec_obra='2026-06-01', desv_ent=19, est_ent='EN TRÁNSITO',
             est_gen='EN PROCESO',
             obs='✅ Ítem más avanzado de EL: RI, Rec.Of, AT y OC completadas (Abr-26). '
                 'Entrega contractual Abr-26. FC entrega Jun-26. Confirmar estado físico.',
             accion='Confirmar estado físico y fecha estimada de llegada a obra.'),

        dict(n=13, crit='MEDIA', n_ri='RI-EL-013',
             desc='Iluminación – Columnas, Artefactos y Proyectores',
             ref_plan='EL – Iluminación - Columnas y artefactos',
             ref_crit='',
             ri_prog='2026-04-01', ri_real=None, desv_ri=None, est_ri='PENDIENTE',
             solped='',
             solp_prog='2026-05-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-06-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-08-15', oc_real=None, lt=120, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='PENDIENTE',
             obs='Sin información de avance registrada. Pendiente emisión de RI.',
             accion='Emitir RI fc May-Jun-26. Verificar estado real con ingeniería.'),

        dict(n=14, crit='MEDIA', n_ri='RI-EL-014',
             desc='Prensacables (cable glands) y Cajas de Conexionado EL',
             ref_plan='EL – Cables - Prensacables / Cajas de conexionado',
             ref_crit='',
             ri_prog='2026-04-15', ri_real=None, desv_ri=None, est_ri='PENDIENTE',
             solped='',
             solp_prog='2026-05-15', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-06-15', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-08-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-09-01', oc_real=None, lt=90, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-01', desv_ent=None, est_ent='PENDIENTE',
             est_gen='PENDIENTE',
             obs='Sin información de avance registrada. Pendiente emisión de RI.',
             accion='Emitir RI fc May-Jun-26.'),

        dict(n=15, crit='MEDIA', n_ri='RI-EL-015',
             desc='Tracing Eléctrico',
             ref_plan='EL – Tracing eléctrico',
             ref_crit='',
             ri_prog='2026-04-01', ri_real=None, desv_ri=51, est_ri='ATRASADO',
             solped='',
             solp_prog='2026-05-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-06-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-07-01', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-08-01', oc_real=None, lt=150, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-01', desv_ent=90, est_ent='ATRASADO',
             est_gen='ATRASADO',
             obs='RI prog Abr-26, tendencia Ago-26 (+51 días sobre P0). '
                 'Entrega fc Mar-27 vs necesidad Dic-26. Desvío ~3 meses. '
                 'Alinear con Ing. de detalle.',
             accion='Emitir RI fc Jun-Ago-26. Alinear con cronograma '
                    'ingeniería de detalle. Desvío ~3 meses.'),

        dict(n=16, crit='MEDIA', n_ri='RI-EL-016',
             desc='Protección Catódica y Materiales Misceláneos EL',
             ref_plan='EL – Protección catódica / Materiales Misceláneos',
             ref_crit='',
             ri_prog='2026-05-01', ri_real=None, desv_ri=None, est_ri='PENDIENTE',
             solped='',
             solp_prog='2026-06-01', solp_real=None, est_solp='PENDIENTE',
             recof_prog='2026-07-01', recof_real=None, est_recof='PENDIENTE',
             at_prog='2026-08-15', at_real=None, est_at='PENDIENTE',
             n_oc='', prov='',
             nec_oc='2026-09-15', oc_real=None, lt=90, est_oc='PENDIENTE',
             ent_oc=None, nec_obra='2026-12-15', desv_ent=None, est_ent='PENDIENTE',
             est_gen='PENDIENTE',
             obs='Sin información de avance registrada. Pendiente emisión de RI.',
             accion='Emitir RI fc Jun-26. Sin urgencia crítica a la fecha.'),
    ]


# ─── COLUMNAS DEFINICIÓN ──────────────────────────────────────────────────────
GRUPOS = [
    ('IDENTIFICACIÓN',          5,  C['subtitulo']),
    ('RI  –  REQUISICIÓN DE INGENIERÍA',  4,  C['header_grup']),
    ('SOLICITUD DE PEDIDO (SOLPED)',       4,  '00B050'),
    ('RECEPCIÓN DE OFERTAS',    3,  '7030A0'),
    ('ANÁLISIS TÉCNICO (AT)',   3,  'C55A11'),
    ('ORDEN DE COMPRA (OC)',    6,  '2F5496'),
    ('ENTREGA',                 4,  '375623'),
    ('ESTADO GENERAL',          3,  '843C0C'),
]

COLS = [
    # Identificación
    ('N°',              5,  'center'),
    ('Criticidad',      12, 'center'),
    ('N° RI',           14, 'center'),
    ('Descripción Suministro', 42, 'left'),
    ('Ref. Plan P0',    32, 'left'),
    # RI
    ('F.RI Plan (P0)',  14, 'center'),
    ('F.RI Real / FC',  14, 'center'),
    ('Desvío RI (d)',    12, 'center'),
    ('Estado RI',       16, 'center'),
    # SOLPED
    ('N° SOLPED',       20, 'center'),
    ('F.SOLP Plan (P0)', 14, 'center'),
    ('F.SOLP Real / FC', 14, 'center'),
    ('Estado SOLP',     16, 'center'),
    # Rec. Ofertas
    ('F.RecOf Plan (P0)', 14, 'center'),
    ('F.RecOf Real / FC', 14, 'center'),
    ('Estado Rec.Of',   16, 'center'),
    # AT
    ('F.AT Plan (P0)',  14, 'center'),
    ('F.AT Real / FC',  14, 'center'),
    ('Estado AT',       16, 'center'),
    # OC
    ('N° OC',           18, 'center'),
    ('Proveedor',       20, 'left'),
    ('F.Nec.OC (P0)',   14, 'center'),
    ('F.OC Real / FC',  14, 'center'),
    ('Lead Time (d)',   12, 'center'),
    ('Estado OC',       16, 'center'),
    # Entrega
    ('F.Entrega OC',    14, 'center'),
    ('F.Nec.Obra',      14, 'center'),
    ('Desvío Ent.(d)',  12, 'center'),
    ('Estado Entrega',  16, 'center'),
    # General
    ('Estado General',  18, 'center'),
    ('Observaciones',   60, 'left'),
    ('Acción Requerida', 50, 'left'),
]


def write_sheet(wb, sheet_name, title_color, data, specialty_label):
    ws = wb.create_sheet(title=sheet_name)
    ws.freeze_panes = 'A6'
    ws.sheet_view.showGridLines = True

    # ── Fila 1: Título principal ──────────────────────────────────────────────
    ws.row_dimensions[1].height = 28
    ncols = len(COLS) + 1
    ws.merge_cells(start_row=1, start_column=2,
                   end_row=1, end_column=ncols)
    c = ws.cell(row=1, column=2)
    c.value = 'TRACKER DE SEGUIMIENTO DE SUMINISTROS  ·  PROYECTO LA CALERA II – CPF2'
    c.fill = F(C['titulo'])
    c.font = ft(bold=True, color='FFFFFF', sz=13)
    c.alignment = al('center', 'center')

    # ── Fila 2: Especialidad ──────────────────────────────────────────────────
    ws.row_dimensions[2].height = 22
    ws.merge_cells(start_row=2, start_column=2,
                   end_row=2, end_column=ncols)
    c = ws.cell(row=2, column=2)
    c.value = f'ESPECIALIDAD:  {specialty_label}'
    c.fill = F(title_color)
    c.font = ft(bold=True, color='FFFFFF', sz=11)
    c.alignment = al('center', 'center')

    # ── Fila 3: Metadatos ─────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 18
    meta = [
        ('Proyecto:', 'La Calera II – CPF2', 4),
        ('Plan P0 emitido:', '06-May-2026', 3),
        ('RFSU Planificado:', '03-Feb-2027', 3),
        ('Fecha referencia:', '22-May-2026', 3),
        ('Rev.:', 'Rev. 0', 2),
    ]
    col = 2
    for lbl, val, span in meta:
        lc = ws.cell(row=3, column=col, value=lbl)
        lc.fill = F(C['subtitulo'])
        lc.font = ft(bold=True, color='FFFFFF', sz=9)
        lc.alignment = al('right', 'center')
        vc = ws.cell(row=3, column=col + 1, value=val)
        vc.fill = F('E9EFF7')
        vc.font = ft(sz=9)
        vc.alignment = al('left', 'center')
        col += span

    # ── Fila 4: Grupos de columnas ────────────────────────────────────────────
    ws.row_dimensions[4].height = 16
    col = 2
    for grp_name, grp_span, grp_color in GRUPOS:
        ws.merge_cells(start_row=4, start_column=col,
                       end_row=4, end_column=col + grp_span - 1)
        c = ws.cell(row=4, column=col)
        c.value = grp_name
        c.fill = F(grp_color)
        c.font = ft(bold=True, color='FFFFFF', sz=9)
        c.alignment = al('center', 'center')
        for cc in range(col, col + grp_span):
            ws.cell(row=4, column=cc).border = bd()
        col += grp_span

    # ── Fila 5: Cabeceras de columnas ─────────────────────────────────────────
    ws.row_dimensions[5].height = 40
    for i, (col_name, col_w, col_h) in enumerate(COLS):
        real_col = i + 2
        c = ws.cell(row=5, column=real_col)
        c.value = col_name
        c.fill = F(C['header_col'])
        c.font = ft(bold=True, color='1F3864', sz=9)
        c.alignment = al('center', 'center', wrap=True)
        c.border = bd()
        ws.column_dimensions[get_column_letter(real_col)].width = col_w

    # ── Columna A: estrecha para margen ──────────────────────────────────────
    ws.column_dimensions['A'].width = 1.5

    # ── Filas de datos ────────────────────────────────────────────────────────
    for idx, item in enumerate(data):
        row = 6 + idx
        ws.row_dimensions[row].height = 55
        row_fill = C['row_alt'] if idx % 2 == 0 else C['row_normal']

        def wr(col_offset, value, bold=False, h='center', wrap=False,
               sz=9, color='000000', bg=None):
            c = ws.cell(row=row, column=col_offset + 1)
            c.value = value
            c.fill = F(bg if bg else row_fill)
            c.font = ft(bold=bold, color=color, sz=sz)
            c.alignment = al(h=h, v='top', wrap=wrap)
            c.border = bd()
            return c

        col = 1
        # N°
        wr(col, item['n'], bold=True, sz=10); col += 1
        # Criticidad
        cc = ws.cell(row=row, column=col + 1)
        crit = item['crit']
        cc.value = crit
        cc.fill = F(CRIT_COLOR.get(crit, C['row_normal']))
        cc.font = ft(bold=True, color='FFFFFF' if crit in ['C. CRÍTICO', 'HITO 2'] else '000000', sz=9)
        cc.alignment = al('center', 'center', wrap=True)
        cc.border = bd()
        col += 1
        # N° RI
        wr(col, item['n_ri'], bold=True, h='center'); col += 1
        # Descripción
        wr(col, item['desc'], h='left', wrap=True, sz=9); col += 1
        # Ref. Plan P0
        rc = ws.cell(row=row, column=col + 1)
        rc.value = item['ref_plan']
        rc.fill = F(C['plan_ref'])
        rc.font = ft(sz=8, color='595959')
        rc.alignment = al(h='left', v='top', wrap=True)
        rc.border = bd()
        col += 1

        # RI
        write_date(ws, row, col + 1, item['ri_prog'], planned=True); col += 1
        write_date(ws, row, col + 1, item['ri_real']); col += 1
        write_desvio(ws, row, col + 1, item.get('desv_ri')); col += 1
        sc(ws.cell(row=row, column=col + 1), item['est_ri']); col += 1

        # SOLPED
        wr(col, item['solped'], h='center', sz=8); col += 1
        write_date(ws, row, col + 1, item['solp_prog'], planned=True); col += 1
        write_date(ws, row, col + 1, item['solp_real']); col += 1
        sc(ws.cell(row=row, column=col + 1), item['est_solp']); col += 1

        # Rec.Ofertas
        write_date(ws, row, col + 1, item['recof_prog'], planned=True); col += 1
        write_date(ws, row, col + 1, item['recof_real']); col += 1
        sc(ws.cell(row=row, column=col + 1), item['est_recof']); col += 1

        # AT
        write_date(ws, row, col + 1, item['at_prog'], planned=True); col += 1
        write_date(ws, row, col + 1, item['at_real']); col += 1
        sc(ws.cell(row=row, column=col + 1), item['est_at']); col += 1

        # OC
        wr(col, item['n_oc'], h='center', sz=8); col += 1
        wr(col, item['prov'], h='left', sz=9); col += 1
        write_date(ws, row, col + 1, item['nec_oc'], planned=True); col += 1
        write_date(ws, row, col + 1, item['oc_real']); col += 1
        lt_c = ws.cell(row=row, column=col + 1)
        lt_c.value = item['lt'] if item['lt'] else ''
        lt_c.fill = F(row_fill)
        lt_c.font = ft(sz=9)
        lt_c.alignment = al('center')
        lt_c.border = bd()
        col += 1
        sc(ws.cell(row=row, column=col + 1), item['est_oc']); col += 1

        # Entrega
        write_date(ws, row, col + 1, item['ent_oc']); col += 1
        write_date(ws, row, col + 1, item['nec_obra'], planned=True); col += 1
        write_desvio(ws, row, col + 1, item.get('desv_ent')); col += 1
        sc(ws.cell(row=row, column=col + 1), item['est_ent']); col += 1

        # Estado General
        sc(ws.cell(row=row, column=col + 1), item['est_gen']); col += 1
        # Observaciones
        wr(col, item['obs'], h='left', wrap=True, sz=8); col += 1
        # Acción
        ac = ws.cell(row=row, column=col + 1)
        is_urgent = '🚨' in item['accion'] or '⚠' in item['accion']
        ac.value = item['accion']
        ac.fill = F('FFE0E0' if is_urgent else row_fill)
        ac.font = ft(sz=8, color='CC0000' if is_urgent else '000000',
                     bold=is_urgent)
        ac.alignment = al(h='left', v='top', wrap=True)
        ac.border = bd()

    # ── Leyenda ───────────────────────────────────────────────────────────────
    last_row = 6 + len(data) + 2
    ws.row_dimensions[last_row].height = 14
    leyenda = [
        ('🟢 COMPLETADO / APROBADO', C['completado'], 'FFFFFF'),
        ('🟡 EN PROCESO / EMITIDA', C['en_proceso'], '000000'),
        ('🔴 ATRASADO', C['atrasado'], 'FFFFFF'),
        ('⚪ PENDIENTE', C['pendiente'], '000000'),
        ('🔵 NO APLICA', C['no_aplica'], 'FFFFFF'),
        ('🚛 EN TRÁNSITO', C['en_transito'], 'FFFFFF'),
        ('⛔ SIN RI', 'FF4500', 'FFFFFF'),
        ('  Fecha Plan P0 (referencia)', C['plan_ref'], '595959'),
    ]
    ws.merge_cells(start_row=last_row, start_column=2,
                   end_row=last_row, end_column=4)
    ws.cell(row=last_row, column=2).value = 'LEYENDA:'
    ws.cell(row=last_row, column=2).font = ft(bold=True, sz=9)

    for i, (txt, bg, fg) in enumerate(leyenda):
        lc = ws.cell(row=last_row + 1, column=2 + i * 2)
        ws.merge_cells(start_row=last_row + 1,
                       start_column=2 + i * 2,
                       end_row=last_row + 1,
                       end_column=2 + i * 2 + 1)
        lc.value = txt
        lc.fill = F(bg)
        lc.font = ft(sz=8, color=fg)
        lc.alignment = al('center', 'center')
        lc.border = bd()

    return ws


# ─── HOJA DASHBOARD ──────────────────────────────────────────────────────────
def write_dashboard(wb, in_data, el_data):
    ws = wb.create_sheet(title='DASHBOARD', index=1)
    ws.sheet_view.showGridLines = False

    def cnt(data, field, val):
        return sum(1 for d in data if val.upper() in d[field].upper())

    def kpi_block(start_row, start_col, label, data, color):
        total = len(data)
        comp = cnt(data, 'est_gen', 'COMPLETADO')
        proc = cnt(data, 'est_gen', 'EN PROCESO')
        atr  = cnt(data, 'est_gen', 'ATRASADO')
        pend = cnt(data, 'est_gen', 'PENDIENTE')

        ws.row_dimensions[start_row].height = 22
        ws.merge_cells(start_row=start_row, start_column=start_col,
                       end_row=start_row, end_column=start_col + 5)
        c = ws.cell(row=start_row, column=start_col, value=label)
        c.fill = F(color)
        c.font = ft(bold=True, color='FFFFFF', sz=12)
        c.alignment = al('center', 'center')

        kpis = [
            ('Total Ítems (RI)', total, C['subtitulo'], 'FFFFFF'),
            ('🟢 Completados',   comp,  C['completado'], 'FFFFFF'),
            ('🟡 En Proceso',    proc,  C['en_proceso'],  '000000'),
            ('🔴 Atrasados',     atr,   C['atrasado'],    'FFFFFF'),
            ('⚪ Pendientes',    pend,  C['pendiente'],   '000000'),
        ]
        for j, (lbl, val, bg, fg) in enumerate(kpis):
            r = start_row + 1
            ws.row_dimensions[r].height = 36
            lc = ws.cell(row=r, column=start_col + j)
            lc.value = lbl
            lc.fill = F(bg)
            lc.font = ft(bold=True, color=fg, sz=8)
            lc.alignment = al('center', 'center', wrap=True)
            lc.border = bd()
            vc = ws.cell(row=r + 1, column=start_col + j)
            ws.row_dimensions[r + 1].height = 30
            vc.value = val
            vc.fill = F(bg)
            vc.font = ft(bold=True, color=fg, sz=18)
            vc.alignment = al('center', 'center')
            vc.border = bd()

    ws.column_dimensions['A'].width = 2
    for i in range(12):
        ws.column_dimensions[get_column_letter(i + 2)].width = 18

    # Título DASHBOARD
    ws.row_dimensions[1].height = 32
    ws.merge_cells('B1:M1')
    c = ws.cell(row=1, column=2)
    c.value = 'DASHBOARD  ·  TRACKER SUMINISTROS  ·  LA CALERA II – CPF2'
    c.fill = F(C['titulo'])
    c.font = ft(bold=True, color='FFFFFF', sz=14)
    c.alignment = al('center', 'center')

    ws.row_dimensions[2].height = 18
    ws.merge_cells('B2:M2')
    c = ws.cell(row=2, column=2)
    c.value = ('Fecha referencia: 22-May-2026  |  '
               'RFSU Planificado: 03-Feb-2027  |  '
               'Hito Ing.80%: 01-Jun-2026')
    c.fill = F(C['subtitulo'])
    c.font = ft(bold=False, color='FFFFFF', sz=10)
    c.alignment = al('center', 'center')

    kpi_block(4,  2,  'IN – INSTRUMENTACIÓN & CONTROL', in_data,  C['header_in'])
    kpi_block(4,  8,  'EL – ELECTRICIDAD',              el_data,  C['header_el'])

    # ── Resumen de ítems críticos ─────────────────────────────────────────────
    r = 9
    ws.row_dimensions[r].height = 20
    ws.merge_cells(start_row=r, start_column=2,
                   end_row=r, end_column=13)
    c = ws.cell(row=r, column=2)
    c.value = 'ÍTEMS CRÍTICOS / EN RIESGO PARA RFSU 03-FEB-2027'
    c.fill = F(C['atrasado'])
    c.font = ft(bold=True, color='FFFFFF', sz=11)
    c.alignment = al('center', 'center')

    headers = ['Esp.', 'N° RI', 'Descripción', 'Criticidad',
               'Estado RI', 'Nec. OC (P0)', 'Lead Time', 'Nec. Obra',
               'Estado Gral.', 'Acción clave']
    widths   = [5, 14, 40, 12, 14, 14, 12, 14, 14, 45]
    ws.row_dimensions[r + 1].height = 30
    for j, (h, w) in enumerate(zip(headers, widths)):
        c = ws.cell(row=r + 1, column=2 + j)
        c.value = h
        c.fill = F(C['header_col'])
        c.font = ft(bold=True, color='1F3864', sz=9)
        c.alignment = al('center', 'center', wrap=True)
        c.border = bd()

    # Items críticos seleccionados manualmente
    criticos = [
        ('IN', 'RI-016', 'Sistema SIS (HIMA)', 'LLI',
         '🟡 EN PROCESO', '21-JUN-26', '300 d', '18-ABR-27',
         '🟡 EN PROCESO', 'OC CRÍTICA fc 21-Jun-26. LT 300d.'),
        ('IN', 'RI-015', 'Sistema PCS (Inauco)', 'LLI',
         '🟡 EN PROCESO', '19-AGO-26', '300 d', '24-JUN-27',
         '🟡 EN PROCESO', 'AT fc Jul-26. OC fc Feb-27. LT 300d.'),
        ('IN', 'RI-020', 'Cables Instrumentación Rev.0', 'C. CRÍTICO',
         '🔴 ATRASADO', '21-OCT-26', '150 d', '29-MAR-27',
         '🔴 ATRASADO', 'Emitir RI fc May-26. Dep. plano Rev.0.'),
        ('IN', 'RI-022', 'Materiales Montaje IN', 'C. CRÍTICO',
         '🔴 ATRASADO', '01-NOV-26', '90 d', '03-FEB-27',
         '🔴 ATRASADO', 'Definir alcance y emitir RI URGENTE.'),
        ('IN', 'RI-019', 'Bandejas portacables IN', 'C. CRÍTICO',
         '🔴 ATRASADO', '01-JUL-26', '90 d', '03-FEB-27',
         '🔴 ATRASADO', 'Verificar RI activa. Emitir urgente.'),
        ('EL', 'RI-EL-006', 'Tableros BT y CCM', 'C. CRÍTICO',
         '🔴 ATRASADO', '01-JUN-26', '300 d', '03-FEB-27',
         '🔴 ATRASADO', 'Emitir RI INMEDIATO. LT 10-12 meses.'),
        ('EL', 'RI-EL-007', 'Sistema UPS y Rectificadores', 'C. CRÍTICO',
         '🔴 ATRASADO', '01-JUN-26', '300 d', '03-FEB-27',
         '🔴 ATRASADO', 'Emitir RI INMEDIATO. LT ~10 meses.'),
        ('EL', 'RI-EL-008', 'Ducto de Barras', 'C. CRÍTICO',
         '🔴 ATRASADO', '01-JUN-26', '240 d', '03-FEB-27',
         '🔴 ATRASADO', 'Emitir RI INMEDIATO. LT 8-10 meses.'),
        ('EL', 'RI-EL-009', 'Cables Eléctricos Rev.0', 'C. CRÍTICO',
         '🔴 ATRASADO', '18-NOV-26', '150 d', '26-ABR-27',
         '🔴 ATRASADO', 'Emitir RI fc Jun-26. Dep. plano Rev.0.'),
        ('EL', 'RI-EL-011', 'Canalizaciones EL', 'C. CRÍTICO',
         '🔴 ATRASADO', '01-JUL-26', '120 d', '01-OCT-26',
         '🔴 ATRASADO', 'ATRASADO: entrega fc Ene-27 vs nec Oct-26.'),
        ('EL', 'RI-EL-004', 'Shelter SE#5 (sin RI)', 'HITO 2',
         '⛔ SIN RI', '01-JUN-26', '390 d', '01-JUL-27',
         '🔴 ATRASADO', 'Confirmar alcance. Emitir RI si aplica.'),
        ('EL', 'RI-EL-005', 'Tableros MT + Transf.', 'LLI',
         '🟡 EN PROCESO', '01-ENE-27', '360 d', '01-JUN-27',
         '🟡 EN PROCESO', 'Cerrar AT urgente. LT ~12 meses.'),
    ]

    STATUS_BG = {
        '🟢': C['completado'], '🟡': C['en_proceso'],
        '🔴': C['atrasado'],   '⚪': C['pendiente'],
        '⛔': 'FF4500',
    }

    for i, row_data in enumerate(criticos):
        rr = r + 2 + i
        ws.row_dimensions[rr].height = 36
        esp, n_ri, desc, crit, est_ri, nec_oc, lt, nec_obra, est_gen, accion = row_data
        bg_row = C['row_alt'] if i % 2 == 0 else C['row_normal']
        vals = [esp, n_ri, desc, crit, est_ri, nec_oc, lt, nec_obra, est_gen, accion]
        for j, v in enumerate(vals):
            c = ws.cell(row=rr, column=2 + j)
            c.value = v
            first_char = v[0] if v else ''
            if first_char in STATUS_BG:
                c.fill = F(STATUS_BG[first_char])
                is_dark = first_char in ('🟢', '🔴', '⛔')
                c.font = ft(sz=9, bold=(first_char != '⚪'),
                            color='FFFFFF' if is_dark else '000000')
            elif j == 3:  # Criticidad
                c.fill = F(CRIT_COLOR.get(v, bg_row))
                c.font = ft(sz=9, bold=True,
                            color='FFFFFF' if v in ('C. CRÍTICO', 'HITO 2', 'LLI') else '000000')
            else:
                c.fill = F(bg_row)
                c.font = ft(sz=9)
            c.alignment = al(h='left' if j in (2, 9) else 'center',
                              v='center', wrap=True)
            c.border = bd()

    return ws


# ─── HOJA PORTADA ─────────────────────────────────────────────────────────────
def write_portada(wb):
    ws = wb.create_sheet(title='PORTADA', index=0)
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 40

    r = 2
    ws.row_dimensions[r].height = 40
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = 'TRACKER DE SEGUIMIENTO DE SUMINISTROS'
    c.fill = F(C['titulo'])
    c.font = ft(bold=True, color='FFFFFF', sz=16)
    c.alignment = al('center', 'center')

    r += 1
    ws.row_dimensions[r].height = 24
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = 'ESPECIALIDADES: INSTRUMENTACIÓN & CONTROL  ·  ELECTRICIDAD'
    c.fill = F(C['titulo'])
    c.font = ft(bold=True, color='FFD966', sz=12)
    c.alignment = al('center', 'center')

    r += 2
    fields = [
        ('Proyecto:',             'La Calera II – CPF2'),
        ('Contrato:',             'CPF2'),
        ('Fecha de emisión Plan (P0):', '06-May-2026'),
        ('Revisión:',             'Rev. 0'),
        ('RFSU Planificado:',     '03-Feb-2027'),
        ('Hito Ing. 80%:',        '01-Jun-2026'),
        ('Fecha de referencia:',  '22-May-2026'),
        ('Actualizado por:',      ''),
    ]
    for lbl, val in fields:
        ws.row_dimensions[r].height = 18
        lc = ws.cell(row=r, column=2, value=lbl)
        lc.fill = F(C['subtitulo'])
        lc.font = ft(bold=True, color='FFFFFF', sz=10)
        lc.alignment = al('right', 'center')
        lc.border = bd()
        vc = ws.cell(row=r, column=3, value=val)
        vc.fill = F('DEEAF1')
        vc.font = ft(sz=10)
        vc.alignment = al('left', 'center')
        vc.border = bd()
        r += 1

    r += 1
    ws.row_dimensions[r].height = 18
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = 'LEYENDA – ESTADO GENERAL'
    c.fill = F(C['subtitulo'])
    c.font = ft(bold=True, color='FFFFFF', sz=11)
    c.alignment = al('center', 'center')

    r += 1
    legend = [
        ('🟢  COMPLETADO / APROBADO', C['completado'],  'FFFFFF',
         'Hito/etapa finalizada en tiempo o antes'),
        ('🟡  EN PROCESO / EMITIDA',  C['en_proceso'],  '000000',
         'Actividad en curso dentro del plazo programado'),
        ('🔴  ATRASADO',              C['atrasado'],    'FFFFFF',
         'Fecha programada superada sin completar'),
        ('⚪  PENDIENTE',             C['pendiente'],   '000000',
         'No iniciado, dentro del plazo programado'),
        ('🔵  NO APLICA',             C['no_aplica'],   'FFFFFF',
         'El ítem no requiere esta etapa'),
        ('🚛  EN TRÁNSITO',           C['en_transito'], 'FFFFFF',
         'Material despachado, en camino a obra'),
        ('⛔  SIN RI',                'FF4500',         'FFFFFF',
         'Requisición de Ingeniería no localizada / no emitida'),
    ]
    for emoji_txt, bg, fg, desc in legend:
        ws.row_dimensions[r].height = 18
        ec = ws.cell(row=r, column=2, value=emoji_txt)
        ec.fill = F(bg)
        ec.font = ft(bold=True, color=fg, sz=10)
        ec.alignment = al('left', 'center')
        ec.border = bd()
        dc = ws.cell(row=r, column=3, value=desc)
        dc.fill = F('F5F5F5')
        dc.font = ft(sz=10)
        dc.alignment = al('left', 'center')
        dc.border = bd()
        r += 1

    # Nota sobre columnas Plan P0
    r += 2
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = ('NOTA: Las celdas con fondo amarillo claro corresponden a fechas '
               'del Plan de Suministros P0 (referencia baseline 06-May-2026).')
    c.fill = F(C['plan_ref'])
    c.font = ft(sz=9, color='595959')
    c.alignment = al('left', 'center', wrap=True)
    c.border = bd()
    ws.row_dimensions[r].height = 24

    return ws


# ─── MAIN ────────────────────────────────────────────────────────────────────
def main():
    wb = openpyxl.Workbook()
    wb.remove(wb.active)   # remove default sheet

    in_data = get_in_data()
    el_data = get_el_data()

    write_portada(wb)
    write_dashboard(wb, in_data, el_data)

    write_sheet(wb, 'IN',
                title_color=C['header_in'],
                data=in_data,
                specialty_label='INSTRUMENTACIÓN & CONTROL  (22 requisiciones)')

    write_sheet(wb, 'EL',
                title_color=C['header_el'],
                data=el_data,
                specialty_label='ELECTRICIDAD  (16 requisiciones)')

    output = '/home/user/OWN-Daro-1/Tracker_Suministros_IN_EL_LaCalera_II.xlsx'
    wb.save(output)
    print(f'✅ Archivo generado: {output}')
    print(f'   IN: {len(in_data)} requisiciones')
    print(f'   EL: {len(el_data)} requisiciones')
    print(f'   Total: {len(in_data) + len(el_data)} filas de detalle')


if __name__ == '__main__':
    main()
