#!/usr/bin/env python3
"""
actualizar_tracker.py
Actualiza el Tracker de Suministros IN/EL – La Calera II CPF2
Versión 2 – Plan 210526 vs Versión Anterior

Genera:
  · Tracker_Suministros_IN_EL_LaCalera_II_v2.xlsx
  · Tracker_Suministros_IN_EL_LaCalera_II_v2.pdf

Cambios detectados entre versión anterior y plan 210526:
  1. EL / RI-EL-003  – Fecha KOM: 18-May-26 → 22-May-26
  2. EL / RI-EL-003  – Observación: "KOM 18/5" → "KOM 22/5"
  3. IN / RI-020      – F. RI Real: (vacío) → 15-May-26
  4. IN / RI-020      – N° SOLPED: (vacío) → "23392215 / 23392216 / 23392213 (en liberación)"
  5. IN / RI-020      – Observación: agrega solpeds y solicitud cables para Dic-26
  6. IN / RI-001      – Observación Rec.Of: apertura 19-May-26 → 21-May-26
  7. IN / RI-015      – Observación: PCS estado actualizado
  8. IN / RI-016      – Observación: SIS estado actualizado
"""

import sys
import os
import copy

import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ── Importar funciones originales ─────────────────────────────────────────────
sys.path.insert(0, '/home/user/OWN-Daro-1')
from generar_tracker_suministros import (
    get_in_data, get_el_data,
    write_portada, write_dashboard, write_sheet,
    C, F, ft, al, bd, bd_m, sc, fmt_date, write_date, write_desvio,
    COLS, GRUPOS, CRIT_COLOR, STATUS_MAP,
)

# ── Importar funciones PDF ────────────────────────────────────────────────────
from generar_pdf_tracker import (
    build_portada as pdf_build_portada,
    build_dashboard as pdf_build_dashboard,
    build_detail_section,
    add_header_footer,
    base_cmds,
    HDR_IN, HDR_EL,
    MARGIN,
)
from reportlab.lib.pagesizes import A4, landscape
from reportlab.platypus import SimpleDocTemplate, PageBreak

# ─────────────────────────────────────────────────────────────────────────────
# VERSIÓN ANTERIOR (hardcodeada) – valores previos al plan 210526
# ─────────────────────────────────────────────────────────────────────────────
VERSION_ANTERIOR = {
    # EL / RI-EL-003 Sistema PMS
    'RI-EL-003': {
        'kom_fecha': '18-May-26',
        'obs': ('OC 4508945953 adjudicada ABB. KOM realizada 18-May-26. '
                'Plan P0: nec.OC 29-May-26, LT 360 días, nec.obra 02-Jun-27.'),
        'accion': 'Seguimiento fabricación ABB. Monitorear % avance.',
    },
    # IN / RI-020 Cables Instrumentación Rev.0
    'RI-020': {
        'ri_real': None,
        'solped': '',
        'obs': ('⚠ RI prog Sep-25. Tendencia actual May-26 (+8 meses). '
                'Plan P0: nec.OC 21-Oct-26, LT 150 días, nec.obra 29-Mar-27. '
                'Dependiente plano de cables Rev.0 (ingeniería).'),
        'accion': ('🚨 Acelerar ingeniería plano cables Rev.0. Emitir RI fc May-26. '
                   'Nec.OC: 21-Oct-26. Desvío +263d sobre P0.'),
    },
    # IN / RI-001 Válvulas de Control
    'RI-001': {
        'obs': ('RI-001 emitida 23-Abr (Val.Control/Autoreg.). Apertura ofertas en curso fc 19-May-26. AT fc Ago-26. '
                'Plan P0: nec.OC 26-Jul-26, LT 300 días, nec.obra 31-May-27.'),
        'accion': 'Cerrar Rec.Ofertas fc 19-May. AT fc Ago-26. Necesidad OC: 26-Jul-26.',
    },
    # IN / RI-015 Sistema PCS
    'RI-015': {
        'obs': ('PCS (Inauco nominado): RI 03-Feb, Rec.Of 25-Mar completadas. AT fc Jul-26. '
                'OC fc Feb-27. Plan P0: nec.OC 19-Ago-26, LT 300 días, nec.obra 24-Jun-27. '
                'Lead time crítico para RFSU Feb-27. '
                'En análisis en conjunto con PP.'),
    },
    # IN / RI-016 Sistema SIS
    'RI-016': {
        'obs': ('SIS (HIMA nominado): RI 03-Feb, Rec.Of 11-Mar completadas. AT fc Jul-26. '
                'Plan P0: nec.OC 21-Jun-26, LT 300 días, nec.obra 18-Abr-27. '
                'En análisis en conjunto con PP.'),
    },
}

# ─────────────────────────────────────────────────────────────────────────────
# REGISTRO DE CAMBIOS – definición estructurada
# ─────────────────────────────────────────────────────────────────────────────
CAMBIOS = [
    {
        'n': 1,
        'esp': 'EL',
        'n_ri': 'RI-EL-003',
        'desc_ri': 'Sistema PMS (Power Management System)',
        'campo': 'Fecha KOM (en observación)',
        'valor_ant': 'KOM 18-May-26',
        'valor_new': 'KOM 22-May-26',
        'tipo': 'FECHA MODIFICADA',
        'impacto': 'BAJO',
        'impacto_detalle': 'KOM ya realizada (+4 días). No afecta RFSU ni hitos críticos.',
    },
    {
        'n': 2,
        'esp': 'EL',
        'n_ri': 'RI-EL-003',
        'desc_ri': 'Sistema PMS (Power Management System)',
        'campo': 'Observación',
        'valor_ant': '"KOM 18/5"',
        'valor_new': '"KOM 22/5"',
        'tipo': 'ACTUALIZACIÓN ESTADO',
        'impacto': 'BAJO',
        'impacto_detalle': 'Corrección de fecha KOM. Fabricación en curso sin desvíos.',
    },
    {
        'n': 3,
        'esp': 'IN',
        'n_ri': 'RI-020',
        'desc_ri': 'Cables de Instrumentación – Rev.0',
        'campo': 'F. RI Real',
        'valor_ant': '(vacío)',
        'valor_new': '15-May-26',
        'tipo': 'NUEVA INFO',
        'impacto': 'BAJO',
        'impacto_detalle': ('RI emitida 15-May-26. Confirmación de emisión. '
                            'Desvío sobre P0: +263 días. No cambia la situación crítica.'),
    },
    {
        'n': 4,
        'esp': 'IN',
        'n_ri': 'RI-020',
        'desc_ri': 'Cables de Instrumentación – Rev.0',
        'campo': 'N° SOLPED',
        'valor_ant': '(vacío)',
        'valor_new': '23392215 / 23392216 / 23392213 (en liberación)',
        'tipo': 'NUEVA INFO',
        'impacto': 'MEDIO',
        'impacto_detalle': ('SOLPEDs en proceso de liberación. Avance positivo: '
                            'acelera la gestión hacia Rec.Ofertas y eventual OC. '
                            'Proyecto solicita cables en obra para Dic-26.'),
    },
    {
        'n': 5,
        'esp': 'IN',
        'n_ri': 'RI-020',
        'desc_ri': 'Cables de Instrumentación – Rev.0',
        'campo': 'Observación',
        'valor_ant': 'Sin mención de SOLPEDs ni solicitud de entrega.',
        'valor_new': ('Agrega: "Solpeds 23392215/23392216/23392213 en liberación. '
                      'Proyecto solicita cables en obra para Dic-26."'),
        'tipo': 'NUEVA INFO',
        'impacto': 'MEDIO',
        'impacto_detalle': ('Hito positivo: SOLPEDs avanzando. La solicitud Dic-26 '
                            'define un deadline de entrega claro que presiona la OC.'),
    },
    {
        'n': 6,
        'esp': 'IN',
        'n_ri': 'RI-001',
        'desc_ri': 'Válvulas de Control y Autorreguladoras',
        'campo': 'Observación – Rec. Ofertas (apertura)',
        'valor_ant': 'Apertura ofertas en curso fc 19-May-26',
        'valor_new': 'Apertura ofertas en curso fc 21-May-26',
        'tipo': 'FECHA MODIFICADA',
        'impacto': 'BAJO',
        'impacto_detalle': ('Desvío de +2 días en apertura. Mínimo. '
                            'No afecta hitos de adjudicación ni RFSU.'),
    },
    {
        'n': 7,
        'esp': 'IN',
        'n_ri': 'RI-015',
        'desc_ri': 'Sistema de Control PCS (DCS/RTU/PLC)',
        'campo': 'Observación',
        'valor_ant': '"En análisis en conjunto con PP."',
        'valor_new': ('"Se espera nueva oferta de Inauco para próxima semana. '
                      'Con esta oferta, previo OK de ING, se avanzará con la adjudicación."'),
        'tipo': 'ACTUALIZACIÓN ESTADO',
        'impacto': 'MEDIO',
        'impacto_detalle': ('Señal positiva: adjudicación inminente de PCS (Inauco). '
                            'LLI crítico. Cierre esperado próxima semana. '
                            'Monitorear para no comprometer OC fc Ago-26.'),
    },
    {
        'n': 8,
        'esp': 'IN',
        'n_ri': 'RI-016',
        'desc_ri': 'Sistema de Seguridad SIS',
        'campo': 'Observación',
        'valor_ant': '"En análisis en conjunto con PP."',
        'valor_new': ('"Se espera oferta final de HIMA para cerrar la gestión. '
                      'Se hará validación técnica. En paralelo, oferta budgetaria '
                      'por servicios 2da instancia."'),
        'tipo': 'ACTUALIZACIÓN ESTADO',
        'impacto': 'MEDIO',
        'impacto_detalle': ('Más información sobre SIS: oferta final HIMA pendiente. '
                            'Sin cambio material pero avanza la gestión. '
                            'OC CRÍTICA fc 21-Jun-26 mantiene urgencia máxima.'),
    },
]

# ─────────────────────────────────────────────────────────────────────────────
# APLICAR CAMBIOS A LOS DATOS
# ─────────────────────────────────────────────────────────────────────────────

def apply_changes_to_data(in_data, el_data):
    """Aplica los cambios del plan 210526 sobre los datos del tracker."""
    changes_applied = []

    for item in in_data:
        if item['n_ri'] == 'RI-020':
            # Cambio 3: F. RI Real
            item['ri_real'] = '2026-05-15'
            item['desv_ri'] = (263 + 14)  # 9-sep-25 a 15-may-26 = ~247 días
            # Recalcular: sep 1 2025 a may 15 2026
            from datetime import date
            prog = date(2025, 9, 1)
            real = date(2026, 5, 15)
            item['desv_ri'] = (real - prog).days
            item['est_ri'] = 'COMPLETADO'

            # Cambio 4: N° SOLPED
            item['solped'] = '23392215 / 23392216 / 23392213 (en liberación)'
            item['est_solp'] = 'EN PROCESO'
            item['solp_real'] = '2026-05-15'

            # Cambio 5: Observación actualizada
            item['obs'] = (
                '⚠ RI emitida 15-May-26 (prog. Sep-25, desvío +257 días sobre P0). '
                'SOLPEDs 23392215 / 23392216 / 23392213 en proceso de liberación. '
                'Plan P0: nec.OC 21-Oct-26, LT 150 días, nec.obra 29-Mar-27. '
                'Proyecto solicita cables en obra para Dic-26.'
            )
            item['accion'] = (
                '🚨 SOLPEDs en liberación – gestionar apertura urgente. '
                'Nec.OC: 21-Oct-26. Solicitud entrega Dic-26 presiona cronograma OC. '
                'Desvío +257d sobre P0. Dependiente plano cables Rev.0.'
            )
            changes_applied.append('IN/RI-020: ri_real=15-May-26, solped actualizado, obs actualizada')

        elif item['n_ri'] == 'RI-001':
            # Cambio 6: apertura +2d
            item['obs'] = (
                'RI-001 emitida 23-Abr (Val.Control/Autoreg.). Apertura ofertas en curso fc 21-May-26. '
                'AT fc Ago-26. Plan P0: nec.OC 26-Jul-26, LT 300 días, nec.obra 31-May-27.'
            )
            item['accion'] = 'Cerrar Rec.Ofertas fc 21-May. AT fc Ago-26. Necesidad OC: 26-Jul-26.'
            changes_applied.append('IN/RI-001: apertura Rec.Of actualizada a 21-May-26')

        elif item['n_ri'] == 'RI-015':
            # Cambio 7: PCS actualización estado
            item['obs'] = (
                'PCS (Inauco nominado): RI 03-Feb, Rec.Of 25-Mar completadas. AT fc Jul-26. '
                'OC fc Feb-27. Plan P0: nec.OC 19-Ago-26, LT 300 días, nec.obra 24-Jun-27. '
                'Lead time crítico para RFSU Feb-27. '
                'Se espera nueva oferta de Inauco para próxima semana. '
                'Con esta oferta, previo OK de ING, se avanzará con la adjudicación.'
            )
            changes_applied.append('IN/RI-015: PCS observación actualizada – adjudicación inminente')

        elif item['n_ri'] == 'RI-016':
            # Cambio 8: SIS actualización estado
            item['obs'] = (
                'SIS (HIMA nominado): RI 03-Feb, Rec.Of 11-Mar completadas. AT fc Jul-26. '
                'Plan P0: nec.OC 21-Jun-26, LT 300 días, nec.obra 18-Abr-27. '
                'Se espera oferta final de HIMA para cerrar la gestión. Se hará validación técnica. '
                'En paralelo, oferta budgetaria por servicios 2da instancia.'
            )
            changes_applied.append('IN/RI-016: SIS observación actualizada – oferta HIMA pendiente')

    for item in el_data:
        if item['n_ri'] == 'RI-EL-003':
            # Cambios 1 y 2: KOM 22-May-26
            item['obs'] = (
                'OC 4508945953 adjudicada ABB. KOM realizada 22-May-26. '
                'Plan P0: nec.OC 29-May-26, LT 360 días, nec.obra 02-Jun-27.'
            )
            item['accion'] = (
                'Seguimiento fabricación ABB. KOM 22-May-26 realizada. '
                'Confirmar hitos intermedios (Ing., Acopio, Fab.).'
            )
            changes_applied.append('EL/RI-EL-003: KOM actualizado a 22-May-26')

    return changes_applied


# ─────────────────────────────────────────────────────────────────────────────
# HOJA CAMBIOS – EXCEL
# ─────────────────────────────────────────────────────────────────────────────

def write_hoja_cambios(wb):
    """Genera la hoja CAMBIOS al índice 0 del workbook."""
    ws = wb.create_sheet(title='CAMBIOS', index=0)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = 'A5'

    # Colores específicos de la hoja
    COLOR_TITULO_CAMBIOS = 'C00000'   # Rojo oscuro
    COLOR_SUB_CAMBIOS    = 'FF4500'   # Naranja rojo
    COLOR_IN_ROW         = 'FDE9D9'   # Naranja suave
    COLOR_EL_ROW         = 'EBF1DE'   # Verde suave
    COLOR_BAJO           = 'C6EFCE'   # Verde claro
    COLOR_MEDIO          = 'FFEB9C'   # Amarillo
    COLOR_ALTO           = 'FFC7CE'   # Rojo claro
    COLOR_FECHA          = 'DAEEF3'   # Azul muy suave
    COLOR_NUEVA          = 'EBF1DE'   # Verde suave
    COLOR_ACTUALIZACION  = 'FFF2CC'   # Amarillo suave
    HDR_CAMBIOS          = '7030A0'   # Violeta

    def F2(col): return PatternFill('solid', fgColor=col)
    def ft2(bold=False, color='000000', sz=10):
        return Font(bold=bold, color=color, size=sz, name='Calibri')
    def al2(h='center', v='center', wrap=False):
        return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
    def bds():
        s = Side(style='thin', color='8EA9C1')
        return Border(left=s, right=s, top=s, bottom=s)
    def bdm():
        s = Side(style='medium', color='000000')
        return Border(left=s, right=s, top=s, bottom=s)

    # ── Fila 1: Título principal ─────────────────────────────────────────────
    ws.row_dimensions[1].height = 32
    ws.merge_cells('A1:I1')
    c = ws.cell(row=1, column=1)
    c.value = 'REGISTRO DE CAMBIOS – Plan de Suministros 210526 vs Versión Anterior'
    c.fill = F2(COLOR_TITULO_CAMBIOS)
    c.font = ft2(bold=True, color='FFFFFF', sz=14)
    c.alignment = al2('center', 'center')

    # ── Fila 2: Subtítulo/metadatos ──────────────────────────────────────────
    ws.row_dimensions[2].height = 20
    ws.merge_cells('A2:I2')
    c = ws.cell(row=2, column=1)
    c.value = (
        'Proyecto: La Calera II – CPF2  |  '
        'Fecha revisión: 22-May-2026  |  '
        'Rev. anterior: Plan P0 (06-May-2026)  |  '
        'Nuevo plan: 210526 (22-May-2026)'
    )
    c.fill = F2(COLOR_SUB_CAMBIOS)
    c.font = ft2(bold=False, color='FFFFFF', sz=10)
    c.alignment = al2('center', 'center')

    # ── Fila 3: KPIs resumen ─────────────────────────────────────────────────
    ws.row_dimensions[3].height = 22
    kpis = [
        ('Total cambios', str(len(CAMBIOS)), '7030A0', 'FFFFFF'),
        ('EL', str(sum(1 for c2 in CAMBIOS if c2['esp'] == 'EL')), '375623', 'FFFFFF'),
        ('IN', str(sum(1 for c2 in CAMBIOS if c2['esp'] == 'IN')), '843C0C', 'FFFFFF'),
        ('Impacto BAJO',  str(sum(1 for c2 in CAMBIOS if c2['impacto'] == 'BAJO')),  '70AD47', 'FFFFFF'),
        ('Impacto MEDIO', str(sum(1 for c2 in CAMBIOS if c2['impacto'] == 'MEDIO')), 'FFC000', '000000'),
        ('Impacto ALTO',  str(sum(1 for c2 in CAMBIOS if c2['impacto'] == 'ALTO')),  'FF0000', 'FFFFFF'),
    ]
    for i, (lbl, val, bg, fg) in enumerate(kpis):
        col_start = i * 1 + 1
        lc = ws.cell(row=3, column=col_start, value=f'{lbl}: {val}')
        lc.fill = F2(bg)
        lc.font = ft2(bold=True, color=fg, sz=10)
        lc.alignment = al2('center', 'center')
        lc.border = bds()

    # Ajustar merge para kpis en 3 celdas cada uno si sobran
    # (la fila 3 queda como está – una celda por KPI)

    # ── Fila 4: Cabeceras de la tabla ────────────────────────────────────────
    ws.row_dimensions[4].height = 36
    cabeceras = [
        ('N°',             5),
        ('Especialidad',   14),
        ('N° RI',          16),
        ('Descripción RI', 38),
        ('Campo Modificado', 30),
        ('Valor Anterior', 42),
        ('Valor Nuevo',    42),
        ('Tipo de Cambio', 22),
        ('Impacto',        18),
    ]
    for i, (h, w) in enumerate(cabeceras):
        c = ws.cell(row=4, column=i + 1)
        c.value = h
        c.fill = F2(HDR_CAMBIOS)
        c.font = ft2(bold=True, color='FFFFFF', sz=10)
        c.alignment = al2('center', 'center', wrap=True)
        c.border = bds()
        ws.column_dimensions[get_column_letter(i + 1)].width = w

    # ── Filas de datos ────────────────────────────────────────────────────────
    for cambio in CAMBIOS:
        row = 4 + cambio['n']
        ws.row_dimensions[row].height = 52

        esp = cambio['esp']
        row_bg = COLOR_IN_ROW if esp == 'IN' else COLOR_EL_ROW

        # Colores por tipo de cambio
        tipo = cambio['tipo']
        tipo_bg = {
            'FECHA MODIFICADA':    COLOR_FECHA,
            'NUEVA INFO':          COLOR_NUEVA,
            'ACTUALIZACIÓN ESTADO': COLOR_ACTUALIZACION,
        }.get(tipo, 'FFFFFF')

        # Colores por impacto
        imp = cambio['impacto']
        imp_bg = {
            'BAJO': COLOR_BAJO,
            'MEDIO': COLOR_MEDIO,
            'ALTO': COLOR_ALTO,
        }.get(imp, 'FFFFFF')
        imp_fg = '000000'

        vals = [
            (str(cambio['n']),      True,  'center', False, row_bg, '000000'),
            (esp,                   True,  'center', False,
             '843C0C' if esp == 'IN' else '375623', 'FFFFFF'),
            (cambio['n_ri'],        True,  'center', False, row_bg, '000000'),
            (cambio['desc_ri'],     False, 'left',   True,  row_bg, '000000'),
            (cambio['campo'],       False, 'left',   True,  row_bg, '595959'),
            (cambio['valor_ant'],   False, 'left',   True,  'FFE0E0', 'CC0000'),
            (cambio['valor_new'],   False, 'left',   True,  'E2EFDA', '375623'),
            (tipo,                  True,  'center', True,  tipo_bg, '000000'),
            (imp,                   True,  'center', False, imp_bg,  imp_fg),
        ]

        for col_i, (val, bold, halign, wrap, bg, fg) in enumerate(vals):
            c = ws.cell(row=row, column=col_i + 1)
            c.value = val
            c.fill = F2(bg)
            c.font = ft2(bold=bold, color=fg, sz=9)
            c.alignment = Alignment(horizontal=halign, vertical='top',
                                     wrap_text=wrap)
            c.border = bds()

    # ── Separador ─────────────────────────────────────────────────────────────
    sep_row = 4 + len(CAMBIOS) + 2
    ws.row_dimensions[sep_row].height = 20
    ws.merge_cells(f'A{sep_row}:I{sep_row}')
    c = ws.cell(row=sep_row, column=1)
    c.value = 'ANÁLISIS DE IMPACTO – Conclusiones'
    c.fill = F2(COLOR_TITULO_CAMBIOS)
    c.font = ft2(bold=True, color='FFFFFF', sz=12)
    c.alignment = al2('left', 'center')

    # ── Análisis de impacto textual ──────────────────────────────────────────
    analisis_rows = [
        ('EL / RI-EL-003 – PMS KOM +4 días',
         'BAJO',
         COLOR_BAJO,
         ('La KOM del Sistema PMS (ABB, OC 4508945953) se realizó el 22-May-26 en lugar del 18-May-26. '
          'El desvío de 4 días es irrelevante para el cronograma: la fabricación ya está en curso, '
          'los hitos contractuales se mantienen (entrega 13-May-27, nec.obra 02-Jun-27) y '
          'no hay impacto sobre RFSU 03-Feb-2027.')),
        ('IN / RI-020 – Cables IN: RI emitida + SOLPEDs en liberación',
         'MEDIO',
         COLOR_MEDIO,
         ('La emisión de la RI el 15-May-26 y el avance de SOLPEDs (23392215/23392216/23392213) '
          'representan un hito positivo significativo frente a la situación crítica anterior. '
          'Sin embargo, el desvío acumulado (+257 días sobre P0) obliga a gestionar la OC '
          'antes del 21-Oct-26 para cumplir la solicitud de entrega en Dic-26. '
          'Riesgo residual: el proceso de adjudicación debe acelerarse.')),
        ('IN / RI-001 – Válvulas Control: apertura Rec.Of +2 días',
         'BAJO',
         COLOR_BAJO,
         ('El corrimiento de la apertura de ofertas de válvulas de control al 21-May-26 '
          'representa un desvío de sólo 2 días respecto a lo informado. '
          'Sin impacto en adjudicación prevista (Ago-26) ni en RFSU.')),
        ('IN / RI-015 – PCS: adjudicación inminente (Inauco)',
         'MEDIO',
         COLOR_MEDIO,
         ('La actualización indica que se espera nueva oferta de Inauco "para la próxima semana" '
          'con perspectiva de adjudicación inmediata. Esto es altamente positivo para el ítem '
          'más crítico del proyecto (LLI, LT 300 días, nec.OC 19-Ago-26). '
          'Monitorear cierre de semana para confirmar adjudicación.')),
        ('IN / RI-016 – SIS: oferta final HIMA pendiente',
         'MEDIO',
         COLOR_MEDIO,
         ('Se actualiza que la oferta final de HIMA está pendiente para cerrar la gestión. '
          'Aunque no hay cambio material en el estado, la información adicional sobre servicios '
          '2da instancia sugiere que el proceso está avanzando. '
          'OC CRÍTICA fc 21-Jun-26 (en ~30 días): se requiere acción inmediata.')),
        ('RESUMEN EJECUTIVO',
         '',
         COLOR_TITULO_CAMBIOS,
         ('Plan 210526 confirma avances positivos en 2 ítems clave (RI-020, RI-015) y '
          'mantiene sin cambio la situación crítica de los ítems ATRASADOS (RI-019, RI-022, '
          'RI-EL-006, RI-EL-007, RI-EL-008, RI-EL-009, RI-EL-011). '
          'El semáforo general del proyecto no cambia: RFSU 03-Feb-2027 sigue en riesgo '
          'por múltiples ítems sin RI emitida y lead times críticos sin cobertura OC.')),
    ]

    for i, (titulo, imp, bg_titulo, texto) in enumerate(analisis_rows):
        r_titulo = sep_row + 1 + i * 3
        r_texto  = r_titulo + 1
        r_sep2   = r_titulo + 2

        ws.row_dimensions[r_titulo].height = 18
        ws.merge_cells(f'A{r_titulo}:I{r_titulo}')
        ct = ws.cell(row=r_titulo, column=1)
        ct.value = f'  {titulo}{"  [Impacto: " + imp + "]" if imp else ""}'
        ct.fill = F2(bg_titulo)
        ct.font = ft2(bold=True,
                       color='FFFFFF' if bg_titulo in (COLOR_TITULO_CAMBIOS, '375623') else '000000',
                       sz=10)
        ct.alignment = al2('left', 'center')

        ws.row_dimensions[r_texto].height = 60
        ws.merge_cells(f'A{r_texto}:I{r_texto}')
        cx = ws.cell(row=r_texto, column=1)
        cx.value = texto
        cx.fill = F2('F8F8F8')
        cx.font = ft2(sz=9, color='000000')
        cx.alignment = Alignment(horizontal='left', vertical='top', wrap_text=True)
        cx.border = bds()

        ws.row_dimensions[r_sep2].height = 6

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# HOJA PORTADA ACTUALIZADA (v2)
# ─────────────────────────────────────────────────────────────────────────────

def write_portada_v2(wb):
    """Portada actualizada para v2 – Rev. 1."""
    ws = wb.create_sheet(title='PORTADA')
    ws.sheet_view.showGridLines = False
    ws.column_dimensions['A'].width = 3
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 40

    r = 2
    ws.row_dimensions[r].height = 40
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = 'TRACKER DE SEGUIMIENTO DE SUMINISTROS'
    c.fill = PatternFill('solid', fgColor=C['titulo'])
    c.font = ft(bold=True, color='FFFFFF', sz=16)
    c.alignment = al('center', 'center')

    r += 1
    ws.row_dimensions[r].height = 24
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = 'ESPECIALIDADES: INSTRUMENTACIÓN & CONTROL  ·  ELECTRICIDAD'
    c.fill = PatternFill('solid', fgColor=C['titulo'])
    c.font = ft(bold=True, color='FFD966', sz=12)
    c.alignment = al('center', 'center')

    r += 2
    fields = [
        ('Proyecto:',                   'La Calera II – CPF2'),
        ('Contrato:',                   'CPF2'),
        ('Fecha emisión Plan anterior:', '06-May-2026  (P0)'),
        ('Fecha nuevo Plan:',           '22-May-2026  (210526)'),
        ('Revisión:',                   'Rev. 1  –  Actualización Plan 210526'),
        ('RFSU Planificado:',           '03-Feb-2027'),
        ('Hito Ing. 80%:',              '01-Jun-2026'),
        ('Fecha de referencia:',        '22-May-2026'),
        ('Cambios incorporados:',       '8 cambios registrados (ver hoja CAMBIOS)'),
        ('Actualizado por:',            ''),
    ]
    for lbl, val in fields:
        ws.row_dimensions[r].height = 18
        lc = ws.cell(row=r, column=2, value=lbl)
        lc.fill = PatternFill('solid', fgColor=C['subtitulo'])
        lc.font = ft(bold=True, color='FFFFFF', sz=10)
        lc.alignment = al('right', 'center')
        lc.border = bd()
        vc = ws.cell(row=r, column=3, value=val)
        vc.fill = PatternFill('solid', fgColor='DEEAF1')
        vc.font = ft(sz=10)
        vc.alignment = al('left', 'center')
        vc.border = bd()
        r += 1

    # Nota sobre la actualización
    r += 1
    ws.row_dimensions[r].height = 30
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = (
        'NOTA Rev. 1: Este tracker incorpora los cambios del Plan de Suministros '
        'actualizado (210526, 22-May-2026) respecto a la versión anterior (P0, 06-May-2026). '
        'Los cambios detectados se detallan en la hoja CAMBIOS (índice 0).'
    )
    c.fill = PatternFill('solid', fgColor='FFF2CC')
    c.font = ft(sz=9, color='595959')
    c.alignment = al('left', 'center', wrap=True)
    c.border = bd()

    r += 2
    ws.row_dimensions[r].height = 18
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = 'LEYENDA – ESTADO GENERAL'
    c.fill = PatternFill('solid', fgColor=C['subtitulo'])
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
        ec.fill = PatternFill('solid', fgColor=bg)
        ec.font = ft(bold=True, color=fg, sz=10)
        ec.alignment = al('left', 'center')
        ec.border = bd()
        dc = ws.cell(row=r, column=3, value=desc)
        dc.fill = PatternFill('solid', fgColor='F5F5F5')
        dc.font = ft(sz=10)
        dc.alignment = al('left', 'center')
        dc.border = bd()
        r += 1

    r += 2
    ws.merge_cells(f'B{r}:H{r}')
    c = ws.cell(row=r, column=2)
    c.value = ('NOTA: Las celdas con fondo amarillo claro corresponden a fechas '
               'del Plan de Suministros P0 (referencia baseline 06-May-2026).')
    c.fill = PatternFill('solid', fgColor=C['plan_ref'])
    c.font = ft(sz=9, color='595959')
    c.alignment = al('left', 'center', wrap=True)
    c.border = bd()
    ws.row_dimensions[r].height = 24

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# GENERADOR EXCEL v2
# ─────────────────────────────────────────────────────────────────────────────

def generar_excel_v2(in_data, el_data):
    output = '/home/user/OWN-Daro-1/Tracker_Suministros_IN_EL_LaCalera_II_v2.xlsx'

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # Crear hojas en orden correcto.
    # write_dashboard tiene index=1 hardcoded, así que hay que gestionar
    # el orden manualmente reordenando al final.

    # 1. Hoja CAMBIOS
    write_hoja_cambios(wb)            # index 0

    # 2. Hoja PORTADA
    write_portada_v2(wb)              # index 1 en este momento

    # 3. Hoja DASHBOARD – la función original fuerza index=1,
    #    lo que desplaza PORTADA a index 2. Se corrige abajo.
    write_dashboard(wb, in_data, el_data)

    # 4. Hoja IN
    write_sheet(wb, 'IN',
                title_color=C['header_in'],
                data=in_data,
                specialty_label='INSTRUMENTACIÓN & CONTROL  (22 requisiciones)  ·  Rev.1 / 22-May-2026')

    # 5. Hoja EL
    write_sheet(wb, 'EL',
                title_color=C['header_el'],
                data=el_data,
                specialty_label='ELECTRICIDAD  (16 requisiciones)  ·  Rev.1 / 22-May-2026')

    # Reordenar hojas al orden deseado: CAMBIOS, PORTADA, DASHBOARD, IN, EL
    desired_order = ['CAMBIOS', 'PORTADA', 'DASHBOARD', 'IN', 'EL']
    # openpyxl permite mover hojas con _sheets list
    sheet_map = {ws.title: ws for ws in wb.worksheets}
    wb._sheets = [sheet_map[name] for name in desired_order if name in sheet_map]

    wb.save(output)
    return output


# ─────────────────────────────────────────────────────────────────────────────
# GENERADOR PDF v2
# ─────────────────────────────────────────────────────────────────────────────

def build_cambios_pdf(cambios_list):
    """Genera las páginas de registro de cambios para el PDF."""
    from reportlab.platypus import (Paragraph, Table, TableStyle, Spacer)
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.enums import TA_CENTER, TA_LEFT
    from reportlab.lib import colors
    from reportlab.lib.units import cm
    from reportlab.platypus.flowables import KeepTogether
    from generar_pdf_tracker import (
        sty, p, TITULO, SUBTITULO, HDR_COL, HDR_IN, HDR_EL,
        COMPLETADO, EN_PROCESO, ATRASADO, PENDIENTE, CRIT_MONTO,
        PLAN_REF, ROW_ALT, ROW_NRM, URGENTE_FG, URGENTE_BG,
        HDR_C_S, CELL_S, CELL_L, CELL_B, OBS_S,
        PAGE_W, MARGIN,
    )

    elems = []

    # Título hoja cambios
    titulo_data = [[Paragraph(
        'REGISTRO DE CAMBIOS – Plan de Suministros 210526 vs Versión Anterior',
        sty('T_cambios', fontSize=13, textColor=colors.white, alignment=TA_CENTER,
            leading=17, fontName='Helvetica-Bold'))]]
    tt = Table(titulo_data, colWidths=[PAGE_W - 2*MARGIN])
    tt.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#C00000')),
        ('TOPPADDING', (0, 0), (-1, -1), 10),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
    ]))
    elems.append(tt)

    sub_data = [[Paragraph(
        'Proyecto: La Calera II – CPF2  |  '
        'Fecha referencia: 22-May-2026  |  '
        'Rev. anterior: Plan P0 (06-May-2026)  |  '
        'Nuevo plan: 210526 (22-May-2026)',
        sty('Sub_c', fontSize=9, textColor=colors.white, alignment=TA_CENTER,
            leading=11, fontName='Helvetica'))]]
    st = Table(sub_data, colWidths=[PAGE_W - 2*MARGIN])
    st.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#FF4500')),
        ('TOPPADDING', (0, 0), (-1, -1), 5),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
    ]))
    elems.append(st)
    elems.append(Spacer(1, 0.3*cm))

    # Tabla de cambios
    col_names = ['N°', 'Esp.', 'N° RI', 'Descripción RI', 'Campo',
                 'Valor Anterior', 'Valor Nuevo', 'Tipo Cambio', 'Impacto']
    col_ws_ch = [0.5*cm, 1.0*cm, 1.8*cm, 4.5*cm, 3.0*cm, 0, 0, 2.8*cm, 1.5*cm]
    used_ch = sum(col_ws_ch[:-4]) + col_ws_ch[-2] + col_ws_ch[-1]
    rem_ch = PAGE_W - 2*MARGIN - used_ch - col_ws_ch[0] - col_ws_ch[1] - col_ws_ch[2] - col_ws_ch[7] - col_ws_ch[8]
    col_ws_ch[3] = 4.5*cm
    col_ws_ch[4] = 3.0*cm
    col_ws_ch[5] = rem_ch * 0.48
    col_ws_ch[6] = rem_ch * 0.52

    # Recalculate properly
    fixed = col_ws_ch[0]+col_ws_ch[1]+col_ws_ch[2]+col_ws_ch[3]+col_ws_ch[4]+col_ws_ch[7]+col_ws_ch[8]
    rem_ch2 = PAGE_W - 2*MARGIN - fixed
    col_ws_ch[5] = rem_ch2 * 0.47
    col_ws_ch[6] = rem_ch2 * 0.53

    hdr_row_ch = [Paragraph(h.replace('\n', '<br/>'), HDR_C_S) for h in col_names]

    COLOR_IN_PDF = colors.HexColor('#FDE9D9')
    COLOR_EL_PDF = colors.HexColor('#EBF1DE')
    COLOR_BAJO_PDF  = colors.HexColor('#C6EFCE')
    COLOR_MEDIO_PDF = colors.HexColor('#FFEB9C')
    COLOR_ALTO_PDF  = colors.HexColor('#FFC7CE')

    TIPO_BG = {
        'FECHA MODIFICADA':     colors.HexColor('#DAEEF3'),
        'NUEVA INFO':           colors.HexColor('#EBF1DE'),
        'ACTUALIZACIÓN ESTADO': colors.HexColor('#FFF2CC'),
    }

    rows_ch = [hdr_row_ch]
    cmds_ch = [
        ('BACKGROUND', (0, 0), (-1, 0), HDR_COL),
        ('GRID', (0, 0), (-1, -1), 0.4, colors.HexColor('#8EA9C1')),
        ('TOPPADDING', (0, 0), (-1, -1), 3),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 3),
        ('LEFTPADDING', (0, 0), (-1, -1), 3),
        ('RIGHTPADDING', (0, 0), (-1, -1), 3),
        ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
    ]

    for i, cambio in enumerate(cambios_list):
        r = i + 1
        esp = cambio['esp']
        bg = COLOR_IN_PDF if esp == 'IN' else COLOR_EL_PDF
        tipo = cambio['tipo']
        tipo_bg = TIPO_BG.get(tipo, colors.white)
        imp = cambio['impacto']
        imp_bg = {'BAJO': COLOR_BAJO_PDF, 'MEDIO': COLOR_MEDIO_PDF,
                  'ALTO': COLOR_ALTO_PDF}.get(imp, colors.white)

        esp_color = HDR_IN if esp == 'IN' else HDR_EL
        esp_s = sty(f'es_{esp}_{i}', fontSize=7.5, textColor=colors.white,
                    alignment=TA_CENTER, fontName='Helvetica-Bold')
        ant_s = sty(f'ant_{i}', fontSize=6.5, textColor=colors.HexColor('#CC0000'),
                    alignment=TA_LEFT, leading=8)
        new_s = sty(f'new_{i}', fontSize=6.5, textColor=colors.HexColor('#375623'),
                    alignment=TA_LEFT, leading=8)

        rows_ch.append([
            p(str(cambio['n']), CELL_B),
            Paragraph(esp, esp_s),
            p(cambio['n_ri'], CELL_S),
            Paragraph(cambio['desc_ri'], OBS_S),
            Paragraph(cambio['campo'], OBS_S),
            Paragraph(cambio['valor_ant'], ant_s),
            Paragraph(cambio['valor_new'], new_s),
            Paragraph(tipo, sty(f'tp_{i}', fontSize=6.5, textColor=colors.black,
                                alignment=TA_CENTER, leading=8, fontName='Helvetica-Bold')),
            Paragraph(imp, sty(f'im_{i}', fontSize=7, textColor=colors.black,
                               alignment=TA_CENTER, leading=9, fontName='Helvetica-Bold')),
        ])
        cmds_ch += [
            ('BACKGROUND', (1, r), (1, r), esp_color),
            ('BACKGROUND', (7, r), (7, r), tipo_bg),
            ('BACKGROUND', (8, r), (8, r), imp_bg),
        ]
        for c_i in [0, 2, 3, 4, 5, 6]:
            cmds_ch.append(('BACKGROUND', (c_i, r), (c_i, r), bg))

    ch_tbl = Table(rows_ch, colWidths=col_ws_ch, repeatRows=1)
    ch_tbl.setStyle(TableStyle(cmds_ch))
    elems.append(ch_tbl)
    elems.append(Spacer(1, 0.4*cm))

    # Análisis de impacto
    analisis_hdr = [[Paragraph('ANÁLISIS DE IMPACTO – Conclusiones',
                                sty('AH', fontSize=10, textColor=colors.white,
                                    alignment=TA_LEFT, leading=13,
                                    fontName='Helvetica-Bold'))]]
    ah_t = Table(analisis_hdr, colWidths=[PAGE_W - 2*MARGIN])
    ah_t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#C00000')),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
    ]))
    elems.append(ah_t)
    elems.append(Spacer(1, 0.2*cm))

    analisis_items = [
        ('EL / RI-EL-003 – PMS KOM +4 días', 'BAJO', COLOR_BAJO_PDF,
         ('La KOM del Sistema PMS (ABB) se realizó el 22-May-26 (+4 días). '
          'Sin impacto en fabricación ni RFSU.')),
        ('IN / RI-020 – Cables IN: RI emitida + SOLPEDs en liberación', 'MEDIO', COLOR_MEDIO_PDF,
         ('RI emitida 15-May-26 y SOLPEDs 23392215/23392216/23392213 avanzando. '
          'Hito positivo. Gestionar OC antes del 21-Oct-26 para cumplir entrega Dic-26.')),
        ('IN / RI-001 – Válvulas Control: apertura +2 días', 'BAJO', COLOR_BAJO_PDF,
         ('Apertura rec. ofertas corrida a 21-May-26. Desvío mínimo, sin impacto en adj. ni RFSU.')),
        ('IN / RI-015 – PCS: adjudicación inminente (Inauco)', 'MEDIO', COLOR_MEDIO_PDF,
         ('Nueva oferta de Inauco esperada próxima semana. Adjudicación inminente. '
          'LLI crítico: OC necesaria antes 19-Ago-26.')),
        ('IN / RI-016 – SIS: oferta final HIMA pendiente', 'MEDIO', COLOR_MEDIO_PDF,
         ('Oferta final HIMA en curso. OC CRÍTICA fc 21-Jun-26 (~30 días). Acción inmediata.')),
        ('RESUMEN EJECUTIVO', '', colors.HexColor('#C00000'),
         ('Plan 210526 confirma avances positivos en RI-020 y RI-015. '
          'La situación crítica global del proyecto se mantiene: RFSU 03-Feb-2027 en riesgo '
          'por múltiples ítems sin OC y lead times críticos sin cobertura.')),
    ]

    for titulo, imp, imp_bg, texto in analisis_items:
        imp_txt = f'  [{imp}]' if imp else ''
        titulo_data = [[Paragraph(
            f'{titulo}{imp_txt}',
            sty('at_h', fontSize=8.5, textColor=colors.white, alignment=TA_LEFT,
                leading=11, fontName='Helvetica-Bold'))]]
        at_t = Table(titulo_data, colWidths=[PAGE_W - 2*MARGIN])
        at_bg = imp_bg if imp else colors.HexColor('#C00000')
        fg_color = colors.black if imp in ('BAJO', 'MEDIO') else colors.white
        at_t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), at_bg),
            ('TOPPADDING', (0, 0), (-1, -1), 4),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 4),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
            ('TEXTCOLOR', (0, 0), (-1, -1), fg_color),
        ]))
        texto_data = [[Paragraph(texto, sty('at_txt', fontSize=8, textColor=colors.black,
                                             alignment=TA_LEFT, leading=10))]]
        tx_t = Table(texto_data, colWidths=[PAGE_W - 2*MARGIN])
        tx_t.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor('#F8F8F8')),
            ('TOPPADDING', (0, 0), (-1, -1), 5),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('BOX', (0, 0), (-1, -1), 0.4, colors.HexColor('#8EA9C1')),
        ]))
        elems.append(KeepTogether([at_t, tx_t, Spacer(1, 0.15*cm)]))

    return elems


def generar_pdf_v2(in_data, el_data):
    from reportlab.lib.units import cm as rl_cm
    output = '/home/user/OWN-Daro-1/Tracker_Suministros_IN_EL_LaCalera_II_v2.pdf'

    doc = SimpleDocTemplate(
        output,
        pagesize=landscape(A4),
        leftMargin=MARGIN, rightMargin=MARGIN,
        topMargin=1.5*rl_cm, bottomMargin=1.0*rl_cm,
        title='Tracker Suministros IN/EL – La Calera II v2',
        author='La Calera II CPF2',
    )

    story = []

    # Página 1: Portada
    story += pdf_build_portada()
    story.append(PageBreak())

    # Página 2: CAMBIOS
    story += build_cambios_pdf(CAMBIOS)
    story.append(PageBreak())

    # Página 3: Dashboard
    story += pdf_build_dashboard(in_data, el_data)
    story.append(PageBreak())

    # Páginas 4-6: IN detalle
    story += build_detail_section(
        in_data,
        'IN – INSTRUMENTACIÓN & CONTROL  (22 RI)  ·  Rev.1 / 22-May-2026',
        HDR_IN)
    story.append(PageBreak())

    # Páginas 7-9: EL detalle
    story += build_detail_section(
        el_data,
        'EL – ELECTRICIDAD  (16 RI)  ·  Rev.1 / 22-May-2026',
        HDR_EL)

    def on_page(canvas_obj, doc_obj):
        pn = canvas_obj.getPageNumber()
        if pn == 1:
            sp, bl = '', 'PORTADA'
        elif pn == 2:
            sp, bl = '', 'CAMBIOS  ·  Plan 210526'
        elif pn == 3:
            sp, bl = '', 'DASHBOARD'
        elif pn <= 6:
            sp = 'IN – INSTRUMENTACIÓN & CONTROL'
            bl = f'Bloque {pn-3}'
        else:
            sp = 'EL – ELECTRICIDAD'
            bl = f'Bloque {pn-6}'
        add_header_footer(canvas_obj, doc_obj, sp, bl)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)
    return output


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

def main():
    print('=' * 70)
    print('ACTUALIZADOR TRACKER SUMINISTROS – La Calera II CPF2')
    print('Plan 210526 vs Versión Anterior')
    print('=' * 70)

    # 1. Obtener datos base
    in_data = get_in_data()
    el_data = get_el_data()
    print(f'\n  Datos base cargados: {len(in_data)} IN | {len(el_data)} EL')

    # 2. Aplicar cambios
    print('\n  Aplicando cambios del plan 210526...')
    changes_applied = apply_changes_to_data(in_data, el_data)
    for ch in changes_applied:
        print(f'    ✓ {ch}')
    print(f'  Total cambios aplicados: {len(changes_applied)}')

    # 3. Generar Excel v2
    print('\n  Generando Excel v2...')
    xlsx_out = generar_excel_v2(in_data, el_data)
    print(f'  ✅ Excel generado: {xlsx_out}')

    # 4. Generar PDF v2
    print('\n  Generando PDF v2...')
    pdf_out = generar_pdf_v2(in_data, el_data)
    print(f'  ✅ PDF generado: {pdf_out}')

    # 5. Resumen final
    print('\n' + '=' * 70)
    print('RESUMEN DE CAMBIOS INCORPORADOS:')
    print('=' * 70)
    for cambio in CAMBIOS:
        print(f'  [{cambio["n"]:2d}] {cambio["esp"]:3s} | {cambio["n_ri"]:12s} | '
              f'{cambio["campo"][:30]:30s} | {cambio["tipo"]:22s} | Impacto: {cambio["impacto"]}')

    print('\n  Archivos generados:')
    print(f'    • {xlsx_out}')
    print(f'    • {pdf_out}')
    print('\n  Hojas Excel:')
    print('    • CAMBIOS (índice 0) – Registro de cambios con análisis de impacto')
    print('    • PORTADA (índice 1) – Portada actualizada Rev.1')
    print('    • DASHBOARD (índice 2) – KPIs y ítems críticos')
    print('    • IN (índice 3) – 22 requisiciones IN actualizadas')
    print('    • EL (índice 4) – 16 requisiciones EL actualizadas')
    print('=' * 70)


if __name__ == '__main__':
    main()
