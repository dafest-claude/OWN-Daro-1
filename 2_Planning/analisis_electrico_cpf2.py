#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Análisis Eléctrico CPF-2 La Calera II - Vaca Muerta
Genera Excel completo con análisis de instalación eléctrica
Autor: Especialista Eléctrico / Project Manager
Fecha: Junio 2026
"""

import zlib
import re
import os
from collections import defaultdict
from datetime import date, timedelta

import openpyxl
from openpyxl.styles import (PatternFill, Font, Alignment, Border, Side,
                              GradientFill)
from openpyxl.utils import get_column_letter
from openpyxl.styles.numbers import FORMAT_NUMBER_COMMA_SEPARATED1

# ─────────────────────────────────────────────
# RUTAS
# ─────────────────────────────────────────────
PDF_CABLES   = '/home/user/OWN-Daro-1/4_Documentos/ACAL-00102-LC-E-0001-0 VCO.pdf'
PDF_UNIFILAR = '/home/user/OWN-Daro-1/4_Documentos/ACAL-00100-EE-E-0001-1_VCO.pdf'
OUTPUT_XLSX  = '/home/user/OWN-Daro-1/2_Planning/Analisis_Electrico_CPF2_Campo.xlsx'

# ─────────────────────────────────────────────
# COLORES
# ─────────────────────────────────────────────
COL_HDR_FILL   = PatternFill("solid", fgColor="1F3864")   # Azul oscuro
COL_HDR_FONT   = Font(color="FFFFFF", bold=True, size=10)
COL_ALT_FILL   = PatternFill("solid", fgColor="DCE6F1")   # Azul muy claro
COL_SUBTOT     = PatternFill("solid", fgColor="BDC7C9")   # Gris
COL_TOTAL_FILL = PatternFill("solid", fgColor="FF6600")   # Naranja
COL_TOTAL_FONT = Font(color="FFFFFF", bold=True, size=10)
COL_CPF2_FILL  = PatternFill("solid", fgColor="E2EFDA")   # Verde claro (CPF2)
COL_CPF1_FILL  = PatternFill("solid", fgColor="FFF2CC")   # Amarillo claro (CPF1)
COL_ALERTA     = PatternFill("solid", fgColor="FF0000")   # Rojo
COL_OK_FILL    = PatternFill("solid", fgColor="70AD47")   # Verde

def thin_border():
    thin = Side(style='thin', color="AAAAAA")
    return Border(left=thin, right=thin, top=thin, bottom=thin)

# ─────────────────────────────────────────────
# PDF PARSER
# ─────────────────────────────────────────────
def get_all_text(path):
    """Extrae todo el texto del PDF usando descompresión zlib."""
    with open(path, 'rb') as f:
        data = f.read()
    streams = re.findall(b'stream\r?\n(.*?)\r?\nendstream', data, re.DOTALL)
    texts = []
    for s in streams:
        try:
            dec = zlib.decompress(s)
            for m in re.finditer(rb'\(([^)\n]{1,300})\)\s*Tj', dec):
                t = m.group(1).decode('latin-1', errors='ignore').strip()
                if t:
                    texts.append(t)
            for m in re.finditer(rb'\[([^\]]+)\]\s*TJ', dec):
                inner = m.group(1)
                parts = re.findall(rb'\(([^)\n]{1,300})\)', inner)
                chunk = ''.join(p.decode('latin-1', errors='ignore') for p in parts)
                if chunk.strip():
                    texts.append(chunk.strip())
        except Exception:
            pass
    return texts

def clean_text(t):
    """Limpia texto con secuencias de escape."""
    t = t.replace('\\323', 'Ó').replace('\\363', 'ó')
    t = t.replace('\\301', 'Á').replace('\\341', 'á')
    t = t.replace('\\311', 'É').replace('\\351', 'é')
    t = t.replace('\\315', 'Í').replace('\\355', 'í')
    t = t.replace('\\332', 'Ú').replace('\\372', 'ú')
    t = t.replace('\\321', 'Ñ').replace('\\361', 'ñ')
    t = t.replace('\\(', '(').replace('\\)', ')')
    t = t.replace('\\050', '(').replace('\\051', ')')
    return t

# ─────────────────────────────────────────────
# PARSEO DE CABLES
# ─────────────────────────────────────────────
# Palabras clave que indican encabezados de sección (no cables)
SECTION_KEYWORDS = {
    'ÁREA', 'AREA', 'SALA', 'SECCIÓN', 'SECCION', 'COMANDO', 'SEÑALES',
    'SENALES', 'OTROS', 'GABINETES', 'IDENTIFICACIÓN', 'IDENTIFICACION',
    'TENSIÓN', 'TENSION', 'LONG.', 'FORM.', 'FASE', 'CONST.', 'PROYECTO',
    'TITULO', 'TITULO:', 'LISTADO', 'CABLES', 'DOC', 'DESDE', 'HASTA',
    'TAG', 'ESPECIFICACIÓN', 'ESPECIFICACION', 'INTERCONEXIÓN', 'INTERCONEXION',
    'REVISIÓN', 'REVISION', 'FECHA', 'DESCRIPCION', 'DESCRIPCIÓN',
    'REEMPLAZA', 'CALERA', 'FACILITIES', 'ENGINEERING', 'MAJOR', 'PROJECTS',
    'PLUSPETROL', 'PERU', 'CORP', 'ESC', 'DOCUMENTO', 'AESA', 'VIBROSWITCH',
    'NOTA', 'PÁGINA', 'PAGINA', 'PAG', 'REV', 'APR', 'AZA', 'PLC', 'TDJ',
    'EMISION', 'EMISIÓN', 'APROBACIÓN', 'APROBACION', 'CONSTRUCCIÓN',
    'CONSTRUCCION', 'TODA', 'INFORMACIÓN', 'INFORMACION', 'CONTENIDA',
    'PRESENTE', 'DOCUMENTACIÓN', 'DOCUMENTACION', 'CONFIDENCIAL',
    'PROPIEDAD', 'REPRODUCCIÓN', 'REPRODUCCION', 'COPIA', 'TOTAL',
    'PARCIAL', 'AUTORIZACIÓN', 'AUTORIZACION', 'PREVIA',
}

# Constantes reconocidas de cable
CONST_PATTERNS = [
    '13,2kV', '6,6kV', '1kV-XLPE', '300 V', 'F.O.', 'UTP', 'CAT6',
    'XLPE', 'THHN', 'EPR'
]

# Formaciones reconocidas (patrones)
FORM_PATTERNS = re.compile(
    r'^(\d+x[\d,\.]+(?:/[\d,\.]+)?(?:\+T)?|F\.O\.|[\d,\.]+x[\d,\.]+x[\d,\.]+|[\d]+x[\d,\.]+|[\d,\.]+|UTP|CAT6)',
    re.IGNORECASE
)

def is_number_or_dash(s):
    s2 = s.strip()
    if s2 == '-':
        return True
    try:
        float(s2.replace(',', '.'))
        return True
    except ValueError:
        return False

def is_voltage(s):
    try:
        v = float(s.replace(',', '.'))
        return v in [110, 220, 380, 460, 6600, 13200] or (100 <= v <= 15000)
    except ValueError:
        return False

def is_length(s):
    try:
        v = float(s.replace(',', '.'))
        return 1 <= v <= 5000
    except ValueError:
        return False

def is_formation(s):
    s = s.strip()
    if s == 'F.O.':
        return True
    if FORM_PATTERNS.match(s):
        return True
    if 'UTP' in s.upper() or 'CAT6' in s.upper():
        return True
    return False

def is_construction(s):
    kw = ['kV', 'XLPE', 'EPR', 'THHN', 'F.O.', 'CU', 'AL', 'monomodo',
          'multimodo', '300 V', 'UTP', 'Arm', 'RF']
    return any(k.lower() in s.lower() for k in kw)

def is_tag_cable(s):
    """Detecta si un string parece un TAG de cable."""
    prefixes = ['P-', 'C-', 'FO-', 'MC-', 'ETH-', 'P-RC-', 'P-RCM-']
    return any(s.startswith(p) for p in prefixes)

def is_section_header(s):
    """Detecta si es un encabezado de sección/área."""
    su = s.strip().upper()
    # Páginas
    if re.match(r'^(P[ÁA]GINA|PAG)\s+\d+', su, re.IGNORECASE):
        return True
    # Encabezados de tabla
    if any(kw in su for kw in ['IDENTIFICACI', 'TENSI', 'FECHA DE EMISI',
                                'TAG DEL CABLE', 'ESPECIFICACI']):
        return True
    # Palabras clave de sección
    words = set(su.split())
    if words & SECTION_KEYWORDS:
        # No si tiene un prefijo de tag
        if not is_tag_cable(s) and not s.startswith('102-') and not s.startswith('670-'):
            return True
    return False

def determine_area(section_headers_seen):
    """Determina el área basada en los encabezados de sección vistos."""
    for h in reversed(section_headers_seen):
        h_up = h.upper()
        if 'SALA ELÉCTRICA 4' in h_up or 'SALA 4' in h_up:
            return 'SALA ELÉCTRICA 4'
        if 'SALA ELÉCTRICA 3' in h_up or 'SALA 3' in h_up:
            return 'SALA ELÉCTRICA 3'
        if 'SALA ELÉCTRICA GENERACIÓN' in h_up or 'GENERACIÓN' in h_up:
            return 'SALA ELÉCTRICA GENERACIÓN'
        if 'OTROS GABINETES' in h_up:
            return 'OTROS GABINETES'
        if 'COMANDO' in h_up:
            return 'COMANDO'
        if 'SEÑALE' in h_up or 'SENALE' in h_up:
            return 'SEÑALES'
        if 'VIBROSWITCH' in h_up:
            return 'SEÑALES'
        if 'ÁREA' in h_up or 'AREA' in h_up:
            return h
    return 'GENERAL'

def parse_cables(path):
    """
    Parsea el listado de cables del PDF.

    El PDF tiene dos patrones de renderizado:
    1. Normal: FASE | TAG | FROM | TO | V | m | CONST | FORM | FASE2 | DESC
    2. Page-break split: bloque numérico (FASE,V,m,FASE2) separado del bloque texto
       Ej: al inicio de una nueva página, primero aparecen todos los V/m de la página,
           luego los TAG/FROM/TO/CONST/FORM/DESC.

    Estrategia: detectar bloques de solo-números (FASE,V,m) sin tag asociado y
    crear una cola de datos numéricos para emparejar con los tags que vienen después.
    """
    texts = get_all_text(path)
    cables = []
    section_stack = []
    current_area = 'GENERAL'

    cleaned = [clean_text(t) for t in texts]
    n = len(cleaned)

    # Pre-scan: identificar bloques numéricos huérfanos (page-break split)
    # Un bloque huérfano es: secuencia de (FASE, V, m, FASE2) sin un TAG antes
    # Detectamos el patrón: 0|A + voltaje + longitud + 2 repetidos sin tags intermedios
    # Construimos una cola de (fase, tension, longitud) para emparejar luego

    numeric_queue = []  # lista de (fase, tension, longitud) para page-split cables

    def scan_numeric_orphans():
        """Escanea el texto buscando bloques numéricos sin tag."""
        orphans = []
        i2 = 0
        while i2 < n - 2:
            t2 = cleaned[i2]
            # Patrón: FASE (0 o A) seguido de TENSION, LONGITUD, y "2"
            if t2 in ['0', 'A', 'B', '1'] and i2+2 < n:
                if is_voltage(cleaned[i2+1]) and is_length(cleaned[i2+2]):
                    # Verificar que no hay un tag de cable cercano (±5 posiciones)
                    nearby = cleaned[max(0,i2-3):i2+4]
                    has_tag_near = any(is_tag_cable(x) for x in nearby)
                    # Si no hay tag cercano, es un huérfano potencial
                    if not has_tag_near:
                        orphans.append((i2, t2, cleaned[i2+1], cleaned[i2+2]))
            i2 += 1
        return orphans

    # Ahora parsear con lógica mejorada
    i = 0
    # Acumulador de datos numéricos para el split de páginas
    pending_numeric = []  # lista de (fase, tension, longitud) en orden de aparición

    while i < n:
        t = cleaned[i]

        # Skip encabezados de página
        if re.match(r'^(P[áa]gina|Pagina|P\\341gina)\s+\d+', t, re.IGNORECASE):
            i += 1
            continue

        # Detectar encabezados de sección
        if is_section_header(t):
            section_stack.append(t)
            current_area = determine_area(section_stack)
            i += 1
            continue

        # Detectar encabezados con número de área tipo "102-TGMT-001" solos
        if (re.match(r'^(102|670)-[A-Z]', t) and not is_tag_cable(t)
                and i + 1 < n and not is_tag_cable(cleaned[i+1])
                and not is_voltage(cleaned[i+1])):
            section_stack.append(t)
            i += 1
            continue

        # Detectar bloque numérico huérfano: FASE + TENSION + LONGITUD sin tag
        # Este patrón ocurre en los page-breaks donde el PDF renderiza las columnas
        # numéricas separadas de las columnas de texto
        if (t in ['0', 'A', 'B', '1'] and i+2 < n
                and is_voltage(cleaned[i+1]) and is_length(cleaned[i+2])
                and not (i > 0 and is_tag_cable(cleaned[i-1]))):
            # Verificar que no hay un tag de cable justo antes (que ya fue procesado)
            # Este es un dato huérfano - agregarlo a la cola
            # Pero solo si el próximo elemento no es inmediatamente un tag (caso normal)
            look_ahead = cleaned[i+3] if i+3 < n else ''
            if look_ahead in ['2', '1', 'A', 'B'] or look_ahead == '' or not is_tag_cable(look_ahead):
                # Patrón: FASE | V | m | FASE2(=2) -> datos huérfanos de page-break
                if i+3 < n and cleaned[i+3] in ['2', '1', 'A', 'B', '0']:
                    pending_numeric.append((t, cleaned[i+1], cleaned[i+2]))
                    i += 4  # skip FASE, V, m, FASE2
                    continue

        # Detectar inicio de entrada de cable
        if is_tag_cable(t):
            tag_cable = t.strip()
            fase = '0'
            if i > 0 and cleaned[i-1] in ['0', 'A', 'B', '1']:
                fase = cleaned[i-1]

            j = i + 1
            from_eq = ''
            to_eq = ''
            tension = ''
            longitud = ''
            const = ''
            formacion = ''
            desc = ''

            fields_found = []
            while j < n and len(fields_found) < 9:
                val = cleaned[j]
                if is_tag_cable(val) or is_section_header(val):
                    break
                if re.match(r'^(P[áa]gina|Pagina|P\\341gina)\s+\d+', val, re.IGNORECASE):
                    j += 1
                    continue
                fields_found.append(val)
                j += 1

            pos = 0
            # FROM
            if pos < len(fields_found):
                from_eq = fields_found[pos]
                pos += 1

            # TO
            if pos < len(fields_found):
                to_eq = fields_found[pos]
                pos += 1

            # TENSION
            if pos < len(fields_found) and (is_voltage(fields_found[pos])
                                             or fields_found[pos] == '-'):
                tension = fields_found[pos]
                pos += 1

            # LONGITUD
            if pos < len(fields_found) and is_length(fields_found[pos]):
                longitud = fields_found[pos]
                pos += 1

            # CONST
            if pos < len(fields_found) and is_construction(fields_found[pos]):
                const = fields_found[pos]
                pos += 1

            # FORMACION
            if pos < len(fields_found) and is_formation(fields_found[pos]):
                formacion = fields_found[pos]
                pos += 1

            # FASE2 (skip)
            if pos < len(fields_found) and fields_found[pos] in ['1', '2', '3', 'A', 'B']:
                pos += 1

            # DESCRIPCION
            desc_parts = fields_found[pos:]
            desc = ' '.join(desc_parts).strip()

            # Si falta tensión/longitud, intentar recuperar desde pending_numeric
            if (not tension or not longitud) and pending_numeric:
                pend = pending_numeric.pop(0)
                if not tension:
                    tension = pend[1]
                if not longitud:
                    longitud = pend[2]
                if not fase or fase == '0':
                    fase = pend[0]

            # Si aún falta const/formacion y solo hay datos numéricos disponibles en fields
            # (caso page-break: FROM/TO/CONST/FORM/DESC vienen sin V/m)
            if not const and pos <= 2 and pos < len(fields_found):
                # Reevaluar: puede que from_eq sea CONST y to_eq sea FORM
                # Verificar si el "from_eq" es en realidad construcción
                if is_construction(from_eq):
                    const = from_eq
                    formacion = to_eq if is_formation(to_eq) else ''
                    desc = ' '.join(fields_found[2:]).strip() if len(fields_found) > 2 else ''
                    from_eq = ''
                    to_eq = ''

            if tag_cable and (from_eq or to_eq or const):
                cables.append({
                    'TAG_CABLE': tag_cable,
                    'DE': from_eq,
                    'HACIA': to_eq,
                    'TENSION_V': tension,
                    'LONGITUD_m': longitud,
                    'CONST': const,
                    'FORMACION': formacion,
                    'AREA': current_area,
                    'DESCRIPCION': clean_text(desc),
                    'FASE': fase,
                })

            i = j
            continue

        i += 1

    return cables

# ─────────────────────────────────────────────
# CLASIFICACIÓN DE CABLES
# ─────────────────────────────────────────────
def determine_cpf(cable):
    """Determina si el cable es CPF1 o CPF2."""
    tag = cable['TAG_CABLE']
    de_ = cable['DE']
    hacia = cable['HACIA']
    combined = f"{tag} {de_} {hacia}"
    if '102-' in combined or '102-' in tag:
        return 'CPF2'
    if '670-' in combined:
        return 'CPF1'
    # Control cables que mencionan 102 en DE o HACIA
    if '102-' in de_ or '102-' in hacia:
        return 'CPF2'
    # Señales a VFDs con 102
    if 'CPF-CPC' in de_ or 'CPF-CPC' in hacia:
        return 'CPF2'  # Estos son del CPF2 (panel de control condensado)
    return 'CPF1'

def determine_tipo(cable):
    """Determina el tipo de cable."""
    tag = cable['TAG_CABLE'].upper()
    form = cable['FORMACION'].upper()
    const = cable['CONST'].upper()
    desc = cable['DESCRIPCION'].upper()

    # FO: fibra óptica
    if 'F.O.' in form or 'F.O.' in const or 'MONOMODO' in const or 'MULTIMODO' in const:
        return 'FO'

    # Ethernet / señales de comunicación
    if 'UTP' in form or 'CAT6' in form or 'UTP' in const:
        return 'ETHERNET'

    # Señales multiconductor (multiconductor analógico/digital)
    if re.match(r'^\d+x\d+x[\d,\.]+', form):  # e.g. 6x2x0,5
        return 'SEÑAL'
    if 'MC-' in tag or '300 V' in const:
        return 'SEÑAL'
    if 'BG' in const or 'BI' in const:
        return 'SEÑAL'

    # Tensión
    try:
        tension = float(cable['TENSION_V'].replace(',', '.'))
    except (ValueError, AttributeError):
        tension = 0

    # Media tensión
    if tension >= 6000:
        return 'POTENCIA_MT'

    # Control: empieza con C- o baja tensión de control
    if tag.startswith('C-') or tag.startswith('C-VS') or tag.startswith('C-M-'):
        return 'CONTROL'

    # Potencia BT
    if tag.startswith('P-') and tension >= 200:
        return 'POTENCIA_BT'

    # Potencia RC (resistencia de calentamiento / space heater)
    if tag.startswith('P-RC-') or tag.startswith('P-RCM-'):
        return 'CONTROL'  # Son calefactores auxiliares de motores

    # Default basado en voltaje
    if tension >= 380:
        return 'POTENCIA_BT'
    elif tension > 0:
        return 'CONTROL'
    return 'CONTROL'

def parse_formacion_mm2(formacion):
    """Extrae la sección en mm2 de la formación."""
    if not formacion:
        return 0
    # Patrones: 3x95, 3x240, 3x25/16, 2x4, 4x10+T, 7x2,5+T, 2x2,5, etc.
    m = re.match(r'(\d+)x([\d,\.]+)', formacion)
    if m:
        try:
            return float(m.group(2).replace(',', '.'))
        except ValueError:
            return 0
    return 0

def longitud_float(cable):
    try:
        return float(cable['LONGITUD_m'].replace(',', '.'))
    except (ValueError, AttributeError):
        return 0.0

# ─────────────────────────────────────────────
# CARGAS CPF-2
# ─────────────────────────────────────────────
def determine_tipo_carga(tag, desc, tension, formacion):
    """Clasifica el tipo de carga."""
    tag_u = tag.upper()
    desc_u = desc.upper()

    if 'VFD' in tag_u and ('DP-' in tag_u or 'DP VFD' in tag_u):
        try:
            t = float(tension.replace(',', '.'))
            return 'VFD_MT' if t >= 6000 else 'VFD_BT'
        except (ValueError, AttributeError):
            return 'VFD_BT'

    if 'CCM' in tag_u or 'TGBT' in tag_u:
        return 'SSAA'

    if 'TGMT' in tag_u:
        return 'SSAA'

    if 'UPS' in tag_u or 'VCC' in tag_u:
        return 'SSAA'

    if any(k in tag_u for k in ['M-PAY', 'M-PAZ', 'M-PBM', 'M-PBA', 'M-EAL']):
        try:
            t = float(tension.replace(',', '.'))
            return 'MOTOR_MT' if t >= 6000 else 'MOTOR_BT'
        except (ValueError, AttributeError):
            return 'MOTOR_BT'

    if tag_u.startswith('M-'):
        try:
            t = float(tension.replace(',', '.'))
            return 'MOTOR_MT' if t >= 6000 else 'MOTOR_BT'
        except (ValueError, AttributeError):
            return 'MOTOR_BT'

    if any(k in desc_u for k in ['BOMBA', 'MOTOR', 'COMPRESOR', 'VENTILADOR', 'COOLER']):
        try:
            t = float(tension.replace(',', '.'))
            return 'MOTOR_MT' if t >= 6000 else 'MOTOR_BT'
        except (ValueError, AttributeError):
            return 'MOTOR_BT'

    if any(k in tag_u for k in ['JBI', 'JBIE', 'LP-', 'LP ']) or 'ILUMINAC' in desc_u:
        return 'ILUMINACION'

    if 'CALEF' in desc_u or 'HEAT' in desc_u or 'JBT' in tag_u:
        return 'CALEFACCION'

    if 'SKID' in tag_u or 'SKZZ' in tag_u:
        return 'INSTRUMENTO'

    if any(k in tag_u for k in ['PMS', 'PLC', 'DP-', 'PC-', 'TR-', 'TRMT']):
        return 'CONTROL'

    if 'TM-' in tag_u or 'TOMA' in desc_u:
        return 'SSAA'

    return 'INSTRUMENTO'

# ─────────────────────────────────────────────
# ESTILOS Y UTILIDADES EXCEL
# ─────────────────────────────────────────────
def set_col_widths(ws, widths):
    for col_idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

def apply_header(ws, row, headers, fill=None, font=None):
    if fill is None:
        fill = COL_HDR_FILL
    if font is None:
        font = COL_HDR_FONT
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.fill = fill
        c.font = font
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = thin_border()

def write_row(ws, row, values, fill=None, bold=False, center=False):
    for col, val in enumerate(values, 1):
        c = ws.cell(row=row, column=col, value=val)
        if fill:
            c.fill = fill
        if bold:
            c.font = Font(bold=True)
        c.border = thin_border()
        if center:
            c.alignment = Alignment(horizontal='center', vertical='center')
        else:
            c.alignment = Alignment(vertical='center', wrap_text=True)

def freeze_and_autofilter(ws, freeze_cell, filter_range):
    ws.freeze_panes = freeze_cell
    ws.auto_filter.ref = filter_range

# ─────────────────────────────────────────────
# HOJA 1: CABLES COMPLETO
# ─────────────────────────────────────────────
def build_hoja1(wb, cables):
    ws = wb.create_sheet("1_Cables_Completo")
    ws.sheet_view.showGridLines = True
    ws.row_dimensions[1].height = 35

    headers = ['TAG_CABLE', 'DE (From)', 'HACIA (To)', 'TENSION_V',
               'LONGITUD_m', 'CONST', 'FORMACION', 'AREA', 'DESCRIPCION',
               'TIPO', 'CPF']
    apply_header(ws, 1, headers)

    for i, c in enumerate(cables, 2):
        cpf = c.get('CPF', 'CPF1')
        tipo = c.get('TIPO', '')
        fill = COL_CPF2_FILL if cpf == 'CPF2' else (COL_ALT_FILL if i % 2 == 0 else None)

        values = [
            c['TAG_CABLE'], c['DE'], c['HACIA'],
            c['TENSION_V'], c['LONGITUD_m'], c['CONST'],
            c['FORMACION'], c['AREA'], c['DESCRIPCION'],
            tipo, cpf
        ]
        write_row(ws, i, values, fill=fill)

    widths = [32, 28, 28, 10, 10, 22, 14, 22, 48, 14, 6]
    set_col_widths(ws, widths)
    freeze_and_autofilter(ws, 'A2', f'A1:K{len(cables)+1}')
    return ws

# ─────────────────────────────────────────────
# HOJA 2: RESUMEN FORMACIONES
# ─────────────────────────────────────────────
def build_hoja2(wb, cables):
    ws = wb.create_sheet("2_Resumen_Formaciones")
    ws.row_dimensions[1].height = 35

    headers = ['FORMACION', 'CONST', 'TIPO', 'CANT_CABLES',
               'METROS_TOTAL', 'CPF1_m', 'CPF2_m', '% CPF2']
    apply_header(ws, 1, headers)

    # Agrupar por formación
    grupos = defaultdict(lambda: {
        'const': '', 'tipo': '', 'total': 0, 'cpf1': 0, 'cpf2': 0, 'cant': 0
    })
    for c in cables:
        form = c.get('FORMACION', '') or 'SIN FORMACIÓN'
        lng = longitud_float(c)
        cpf = c.get('CPF', 'CPF1')
        tipo = c.get('TIPO', '')
        grupos[form]['cant'] += 1
        grupos[form]['total'] += lng
        grupos[form]['const'] = c.get('CONST', '')
        grupos[form]['tipo'] = tipo
        if cpf == 'CPF2':
            grupos[form]['cpf2'] += lng
        else:
            grupos[form]['cpf1'] += lng

    sorted_grupos = sorted(grupos.items(), key=lambda x: -x[1]['total'])

    row = 2
    total_cant = 0
    total_m = 0
    total_cpf1 = 0
    total_cpf2 = 0

    for i, (form, data) in enumerate(sorted_grupos):
        pct_cpf2 = (data['cpf2'] / data['total'] * 100) if data['total'] > 0 else 0
        fill = COL_ALT_FILL if i % 2 == 0 else None
        values = [
            form, data['const'], data['tipo'], data['cant'],
            round(data['total'], 0), round(data['cpf1'], 0),
            round(data['cpf2'], 0), round(pct_cpf2, 1)
        ]
        write_row(ws, row, values, fill=fill, center=False)
        # Formato numérico
        for col in [4, 5, 6, 7]:
            ws.cell(row=row, column=col).number_format = '#,##0'
        ws.cell(row=row, column=8).number_format = '0.0"%"'
        total_cant += data['cant']
        total_m += data['total']
        total_cpf1 += data['cpf1']
        total_cpf2 += data['cpf2']
        row += 1

    # Fila total
    pct_total = (total_cpf2 / total_m * 100) if total_m > 0 else 0
    total_vals = ['TOTAL', '', '', total_cant, round(total_m, 0),
                  round(total_cpf1, 0), round(total_cpf2, 0), round(pct_total, 1)]
    write_row(ws, row, total_vals, fill=COL_TOTAL_FILL, bold=True, center=False)
    for col_idx in range(1, 9):
        ws.cell(row=row, column=col_idx).font = COL_TOTAL_FONT
    for col in [4, 5, 6, 7]:
        ws.cell(row=row, column=col).number_format = '#,##0'
    ws.cell(row=row, column=8).number_format = '0.0"%"'

    widths = [20, 24, 14, 12, 14, 12, 12, 10]
    set_col_widths(ws, widths)
    freeze_and_autofilter(ws, 'A2', f'A1:H{row}')
    return ws, total_cant, total_m, total_cpf1, total_cpf2

# ─────────────────────────────────────────────
# HOJA 3: CARGAS CPF-2
# ─────────────────────────────────────────────
def build_hoja3(wb, cables):
    ws = wb.create_sheet("3_Cargas_CPF2")
    ws.row_dimensions[1].height = 35

    headers = ['TAG_EQUIPO', 'DESCRIPCION', 'TENSION_V',
               'CCM_ALIMENTADOR', 'FORMACION_POTENCIA', 'LONGITUD_m',
               'AREA', 'TIPO_CARGA']
    apply_header(ws, 1, headers)

    # Extraer cargas CPF-2: equipos destino de cables de potencia CPF-2
    cargas = {}
    for c in cables:
        if c.get('CPF') != 'CPF2':
            continue
        tipo = c.get('TIPO', '')
        if tipo not in ('POTENCIA_BT', 'POTENCIA_MT'):
            continue

        equipo = c['HACIA'].strip()
        if not equipo or equipo in cargas:
            # Si ya existe, actualizar si la longitud es mayor o el tipo es MT
            if equipo in cargas:
                if tipo == 'POTENCIA_MT':
                    cargas[equipo]['TIPO_CARGA'] = determine_tipo_carga(
                        equipo, c['DESCRIPCION'], c['TENSION_V'], c['FORMACION'])
                    cargas[equipo]['TENSION_V'] = c['TENSION_V']
            continue

        tipo_carga = determine_tipo_carga(
            equipo, c['DESCRIPCION'], c['TENSION_V'], c['FORMACION'])

        cargas[equipo] = {
            'TAG_EQUIPO': equipo,
            'DESCRIPCION': c['DESCRIPCION'],
            'TENSION_V': c['TENSION_V'],
            'CCM_ALIMENTADOR': c['DE'],
            'FORMACION_POTENCIA': c['FORMACION'],
            'LONGITUD_m': c['LONGITUD_m'],
            'AREA': c['AREA'],
            'TIPO_CARGA': tipo_carga,
        }

    sorted_cargas = sorted(cargas.values(), key=lambda x: x['TAG_EQUIPO'])

    for i, carga in enumerate(sorted_cargas, 2):
        fill = COL_ALT_FILL if i % 2 == 0 else None
        values = [
            carga['TAG_EQUIPO'], carga['DESCRIPCION'], carga['TENSION_V'],
            carga['CCM_ALIMENTADOR'], carga['FORMACION_POTENCIA'],
            carga['LONGITUD_m'], carga['AREA'], carga['TIPO_CARGA']
        ]
        write_row(ws, i, values, fill=fill)

    widths = [30, 50, 10, 24, 18, 10, 24, 16]
    set_col_widths(ws, widths)
    freeze_and_autofilter(ws, 'A2', f'A1:H{len(sorted_cargas)+1}')
    return ws, sorted_cargas

# ─────────────────────────────────────────────
# HOJA 4: ESTIMACIÓN DE INSTALACIÓN
# ─────────────────────────────────────────────
def build_hoja4(wb, cables, cargas):
    ws = wb.create_sheet("4_Estimacion_Instalacion")
    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 30

    # Título
    ws.merge_cells('A1:I1')
    title = ws['A1']
    title.value = 'ESTIMACIÓN DE TIEMPOS DE INSTALACIÓN ELÉCTRICA - CPF-2 LA CALERA II'
    title.fill = COL_HDR_FILL
    title.font = Font(color="FFFFFF", bold=True, size=12)
    title.alignment = Alignment(horizontal='center', vertical='center')

    ws.merge_cells('A2:I2')
    sub = ws['A2']
    sub.value = 'Base: 3 cuadrillas eléctricas x 4 personas | 8 hs/día efectivas | Rendimientos campo Patagonia'
    sub.fill = PatternFill("solid", fgColor="2E75B6")
    sub.font = Font(color="FFFFFF", bold=True, size=10)
    sub.alignment = Alignment(horizontal='center', vertical='center')

    headers = ['CATEGORIA', 'ITEM', 'CANTIDAD', 'UNIDAD',
               'RENDIMIENTO', 'DIAS_CUADRILLA', 'CUADRILLAS',
               'DIAS_PARALELO', 'OBSERVACION']
    apply_header(ws, 3, headers)

    # ── Calcular cantidades desde los datos reales ──
    cpf2_cables = [c for c in cables if c.get('CPF') == 'CPF2']

    # Por tipo
    pot_mt = [c for c in cpf2_cables if c.get('TIPO') == 'POTENCIA_MT']
    pot_bt_grande = [c for c in cpf2_cables
                     if c.get('TIPO') == 'POTENCIA_BT'
                     and parse_formacion_mm2(c.get('FORMACION', '')) >= 35]
    pot_bt_medio = [c for c in cpf2_cables
                    if c.get('TIPO') == 'POTENCIA_BT'
                    and 10 <= parse_formacion_mm2(c.get('FORMACION', '')) < 35]
    pot_bt_chico = [c for c in cpf2_cables
                    if c.get('TIPO') == 'POTENCIA_BT'
                    and 0 < parse_formacion_mm2(c.get('FORMACION', '')) < 10]
    ctrl = [c for c in cpf2_cables if c.get('TIPO') == 'CONTROL']
    fo = [c for c in cpf2_cables if c.get('TIPO') == 'FO']
    eth = [c for c in cpf2_cables if c.get('TIPO') == 'ETHERNET']
    senial = [c for c in cpf2_cables if c.get('TIPO') == 'SEÑAL']

    def total_m(lst):
        return sum(longitud_float(c) for c in lst)

    m_pot_mt = total_m(pot_mt)
    m_pot_bt_grande = total_m(pot_bt_grande)
    m_pot_bt_medio = total_m(pot_bt_medio)
    m_pot_bt_chico = total_m(pot_bt_chico)
    m_ctrl = total_m(ctrl) + total_m(senial)
    m_fo = total_m(fo)
    m_eth = total_m(eth)

    # Cargas por tipo
    from collections import Counter
    tipo_carga_cnt = Counter(c['TIPO_CARGA'] for c in cargas)

    # Calcular días de cuadrilla para cada item
    # 3 cuadrillas disponibles en paralelo
    CUADRILLAS = 3

    def dias_paralelo(dias_cuadrilla, cuadrillas=CUADRILLAS):
        return max(1, round(dias_cuadrilla / cuadrillas, 1))

    # ─ TENDIDO DE CABLES ─
    tend_mt = round(m_pot_mt / 30, 1) if m_pot_mt > 0 else 0      # MT más lento
    tend_bt_g = round(m_pot_bt_grande / 50, 1) if m_pot_bt_grande > 0 else 0
    tend_bt_m = round(m_pot_bt_medio / 80, 1) if m_pot_bt_medio > 0 else 0
    tend_bt_c = round(m_pot_bt_chico / 120, 1) if m_pot_bt_chico > 0 else 0
    tend_ctrl = round(m_ctrl / 100, 1) if m_ctrl > 0 else 0
    tend_fo = round(m_fo / 200, 1) if m_fo > 0 else 0
    tend_eth = round(m_eth / 150, 1) if m_eth > 0 else 0

    # ─ CONEXIONADO ─
    n_ext_mt = len(pot_mt) * 2          # 2 extremos por cable
    n_ext_bt_g = len(pot_bt_grande) * 2
    n_ext_bt_m = (len(pot_bt_medio) + len(pot_bt_chico)) * 2
    n_cables_ctrl = len(ctrl) + len(senial)

    conex_mt = round(n_ext_mt * 8 / 8, 1)         # 8 hs/extremo = 1 día/extremo
    conex_bt_g = round(n_ext_bt_g * 3 / 8, 1)     # 3 hs/extremo
    conex_bt_m = round(n_ext_bt_m * 1.5 / 8, 1)   # 1.5 hs/extremo promedio
    conex_ctrl = round(n_cables_ctrl * 4 / 8, 1)   # 4 hs/cable

    # ─ MONTAJE EQUIPOS ─
    # CCMs y tableros MT (de los datos - solo tableros físicos principales)
    # Contamos tableros FÍSICOS, no cada referencia de tag destino
    ccm_fisicos = set()     # Solo CCM-005, CCM-006, CCM-007 (3 unidades)
    tgmt_fisicos = set()    # Solo TGMT-001, TGMT-002, TRMT-001A/B
    motor_mt_tags = set()
    motor_bt_g_tags = set()  # >90kW → asumimos formación >=50mm2
    motor_bt_m_tags = set()  # 22-90kW
    motor_bt_c_tags = set()  # <22kW
    vfd_tags = set()

    for carga in cargas:
        t = carga['TIPO_CARGA']
        tag = carga['TAG_EQUIPO']
        mm2 = parse_formacion_mm2(carga.get('FORMACION_POTENCIA', ''))
        # CCMs físicos: solo los 3 principales
        if re.match(r'^102-CCM-0\d{2}$', tag):
            ccm_fisicos.add(tag)
        # TGMTs físicos
        elif re.match(r'^102-(TGMT|TRMT)-\d+[AB]?$', tag):
            tgmt_fisicos.add(tag)
        elif t in ('MOTOR_MT', 'VFD_MT'):
            motor_mt_tags.add(tag)
        elif t in ('VFD_BT',):
            vfd_tags.add(tag)
        elif t == 'MOTOR_BT':
            if mm2 >= 50:
                motor_bt_g_tags.add(tag)
            elif mm2 >= 10:
                motor_bt_m_tags.add(tag)
            else:
                motor_bt_c_tags.add(tag)

    # Valores realistas para campo CPF-2:
    # 3 CCMs grandes (>20 bahías): 5 días c/u con 2 cuadrillas = 15 días / 2 = 7.5 días
    n_ccm = max(len(ccm_fisicos), 3)  # Mínimo 3 CCMs (CCM-005/006/007)
    mont_ccm = n_ccm * 5
    # Celdas MT: TGMT-001 (2 barras), TGMT-002 (2 barras), TRMT-001A/B = ~6 equipos
    n_tgmt = max(len(tgmt_fisicos), 6)
    mont_tgmt = n_tgmt * 3
    mont_motor_mt = len(motor_mt_tags) * 2
    mont_motor_bt_g = len(motor_bt_g_tags) * 1
    mont_motor_bt_m = len(motor_bt_m_tags) * 0.5
    mont_motor_bt_c = len(motor_bt_c_tags) * 0.25

    mont_total = (mont_ccm + mont_tgmt + mont_motor_mt +
                  mont_motor_bt_g + mont_motor_bt_m + mont_motor_bt_c)

    # ─ PRUEBAS ─
    total_cables_cpf2 = len(cpf2_cables)
    prueb_mega = round(total_cables_cpf2 * 0.25, 1)  # 2 hs/cable = 0.25 días
    prueb_energiz = 10   # días fijos para energización progresiva
    prueb_precom = 15    # días precomisionado eléctrico
    prueb_sat = 30       # SAT PMS (oct-dic)

    rows_data = [
        # CATEGORIA, ITEM, CANT, UNIDAD, RENDIMIENTO, DIAS_CUAD, CUADR, DIAS_PAR, OBS
        # ── TENDIDO ──
        ('── TENDIDO DE CABLES ──', '', '', '', '', '', '', '', ''),
        ('TENDIDO', 'Cables Potencia MT (13.2/6.6kV)',
         round(m_pot_mt), 'm', '30 m/día/cuadrilla',
         tend_mt, 2, dias_paralelo(tend_mt, 2),
         f'{len(pot_mt)} cables | Req. maniobras especiales MT'),
        ('TENDIDO', 'Cables Potencia BT Grande (≥35mm²)',
         round(m_pot_bt_grande), 'm', '50 m/día/cuadrilla',
         tend_bt_g, CUADRILLAS, dias_paralelo(tend_bt_g),
         f'{len(pot_bt_grande)} cables | 3x240, 3x150, 3x95, 3x50mm²'),
        ('TENDIDO', 'Cables Potencia BT Mediano (10-35mm²)',
         round(m_pot_bt_medio), 'm', '80 m/día/cuadrilla',
         tend_bt_m, CUADRILLAS, dias_paralelo(tend_bt_m),
         f'{len(pot_bt_medio)} cables | 3x16, 3x25mm²'),
        ('TENDIDO', 'Cables Potencia BT Pequeño (<10mm²)',
         round(m_pot_bt_chico), 'm', '120 m/día/cuadrilla',
         tend_bt_c, CUADRILLAS, dias_paralelo(tend_bt_c),
         f'{len(pot_bt_chico)} cables | 3x4, 3x6, 4x4mm²'),
        ('TENDIDO', 'Cables Control + Señales',
         round(m_ctrl), 'm', '100 m/día/cuadrilla',
         tend_ctrl, CUADRILLAS, dias_paralelo(tend_ctrl),
         f'{len(ctrl)+len(senial)} cables | multiconductores'),
        ('TENDIDO', 'Cables Fibra Óptica',
         round(m_fo), 'm', '200 m/día/cuadrilla',
         tend_fo, 1, tend_fo,
         f'{len(fo)} cables FO'),
        ('TENDIDO', 'Cables Ethernet UTP',
         round(m_eth), 'm', '150 m/día/cuadrilla',
         tend_eth, 1, tend_eth,
         f'{len(eth)} cables ETH'),

        # ── CONEXIONADO ──
        ('── CONEXIONADO Y TERMINACIÓN ──', '', '', '', '', '', '', '', ''),
        ('CONEXIONADO', 'Terminales MT (6.6/13.2kV)',
         n_ext_mt, 'extremos', '8 hs/extremo = 1 día',
         conex_mt, 2, dias_paralelo(conex_mt, 2),
         'Req. técnicos especializados MT | ABB/AESA'),
        ('CONEXIONADO', 'Terminales BT Grande (≥35mm²)',
         n_ext_bt_g, 'extremos', '3 hs/extremo',
         conex_bt_g, CUADRILLAS, dias_paralelo(conex_bt_g),
         'Zapatas + precintos + termorretráctil'),
        ('CONEXIONADO', 'Terminales BT Mediano/Pequeño',
         n_ext_bt_m, 'extremos', '1.5 hs/extremo',
         conex_bt_m, CUADRILLAS, dias_paralelo(conex_bt_m),
         'Terminales preaislados'),
        ('CONEXIONADO', 'Cables de Control (peinado+etiq.+prueba)',
         n_cables_ctrl, 'cables', '4 hs/cable',
         conex_ctrl, CUADRILLAS, dias_paralelo(conex_ctrl),
         'Incluye control, señales y vibroswitches'),

        # ── MONTAJE EQUIPOS ──
        ('── MONTAJE DE EQUIPOS ──', '', '', '', '', '', '', '', ''),
        ('MONTAJE', 'CCMs BT (grandes >20 bahías)',
         n_ccm, 'unidades', '5 días/CCM',
         mont_ccm, 2, dias_paralelo(mont_ccm, 2),
         f'CCM-005, CCM-006, CCM-007 | Llegan en ABB S3/S4'),
        ('MONTAJE', 'Celdas MT (TGMT + TRMT)',
         n_tgmt, 'equipos', '3 días/equipo',
         mont_tgmt, 2, dias_paralelo(mont_tgmt, 2),
         'TGMT-001, TGMT-002, TRMT-001A/B | Izaje + nivelación'),
        ('MONTAJE', 'Motores MT (>90kW, 6.6kV)',
         len(motor_mt_tags), 'motores', '2 días/motor',
         mont_motor_mt, 2, dias_paralelo(mont_motor_mt, 2),
         'M-PAY-23310 A/B/C, M-PAZ-24210 A/B/C | Izaje + alineación'),
        ('MONTAJE', 'Motores BT Grande (>90kW)',
         len(motor_bt_g_tags), 'motores', '1 día/motor',
         mont_motor_bt_g, CUADRILLAS, dias_paralelo(mont_motor_bt_g),
         'PAY-23040, PAZ bombas transferencia'),
        ('MONTAJE', 'Motores BT Mediano (22-90kW)',
         len(motor_bt_m_tags), 'motores', '0.5 días/motor',
         mont_motor_bt_m, CUADRILLAS, dias_paralelo(mont_motor_bt_m),
         'PBM-24390, PBA-24420, etc.'),
        ('MONTAJE', 'Motores BT Pequeño (<22kW)',
         len(motor_bt_c_tags), 'motores', '0.25 días/motor',
         mont_motor_bt_c, CUADRILLAS, dias_paralelo(mont_motor_bt_c),
         'Auxiliares, vibroswitches, EAL-23310'),

        # ── PRUEBAS ──
        ('── PRUEBAS Y ENERGIZACIÓN ──', '', '', '', '', '', '', '', ''),
        ('PRUEBAS', 'Pruebas de Aislación (Megger) por circuito',
         total_cables_cpf2, 'cables', '2 hs/cable = 0.25 días',
         prueb_mega, CUADRILLAS, dias_paralelo(prueb_mega),
         'Incluye registro y documentación ITP'),
        ('PRUEBAS', 'Energización progresiva MT → BT',
         1, 'lote', '10 días fijos',
         prueb_energiz, 2, prueb_energiz,
         'Con presencia ABB + AESA + Cliente | HOLD POINT'),
        ('PRUEBAS', 'Precomisionado eléctrico funcional',
         1, 'lote', '15 días fijos',
         prueb_precom, 2, prueb_precom,
         'Loop check eléctrico + ajuste relés + secuencia arranque'),
        ('PRUEBAS', 'SAT PMS en sitio (HOLD POINT PIE)',
         1, 'evento', '30 días (oct-dic 2027)',
         prueb_sat, 1, prueb_sat,
         'Ventana oct-dic 2027 per contrato | ABB+Cliente'),
    ]

    row = 4
    total_dias_paralelo = 0
    subtotals = {}
    current_cat = None

    for data in rows_data:
        cat, item, cant, unidad, rend, dias_c, cuadr, dias_p, obs = data

        if item == '':  # Encabezado de categoría
            ws.merge_cells(f'A{row}:I{row}')
            c = ws.cell(row=row, column=1, value=cat)
            c.fill = PatternFill("solid", fgColor="2E75B6")
            c.font = Font(color="FFFFFF", bold=True, size=10)
            c.alignment = Alignment(horizontal='left', vertical='center')
            c.border = thin_border()
            ws.row_dimensions[row].height = 20
            current_cat = cat
            subtotals[cat] = 0
            row += 1
            continue

        fill = COL_ALT_FILL if row % 2 == 0 else None
        values = [cat, item, cant, unidad, rend, dias_c, cuadr, dias_p, obs]
        write_row(ws, row, values, fill=fill)

        # Marcar días paralelo como número
        for col_idx in [3, 6, 7, 8]:
            ws.cell(row=row, column=col_idx).number_format = '#,##0.0'

        if isinstance(dias_p, (int, float)):
            subtotals[current_cat] = subtotals.get(current_cat, 0) + dias_p

        ws.row_dimensions[row].height = 18
        row += 1

    # Total general (suma de máximos paralelo por categoría)
    # Consideramos que las fases se hacen secuencialmente
    phase_totals = {
        'TENDIDO': sum(d for k, d in subtotals.items() if 'TENDIDO' in k),
        'CONEXIONADO': sum(d for k, d in subtotals.items() if 'CONEXION' in k),
        'MONTAJE': sum(d for k, d in subtotals.items() if 'MONTAJE' in k),
        'PRUEBAS': sum(d for k, d in subtotals.items() if 'PRUEBA' in k),
    }

    # Usar días paralelo reales calculados
    # ─ Cálculo de días por ruta crítica (critical path) ─
    # Tendido y Conexionado se solapan parcialmente (conexionado empieza a 50% del tendido)
    # Montaje de tableros se hace en paralelo con tendido de campo
    # Secuencia crítica: Montaje tableros → Tendido → Conexionado → Pruebas → Energización → SAT → RFSU

    # Tendido: el más largo de los grupos con 3 cuadrillas en paralelo
    # Las cuadrillas pueden hacer 2 tipos en simultáneo pero no más
    # Agrupamos en dos bloques: MT/BT-Grande (pesados) y Control/señales
    tend_blq1 = max(dias_paralelo(tend_mt, 2), dias_paralelo(tend_bt_g),
                    dias_paralelo(tend_bt_m), dias_paralelo(tend_bt_c))  # paralelo con 3 cuad
    tend_blq2 = dias_paralelo(tend_ctrl)  # 3 cuadrillas
    # El bloque de control se hace en paralelo con BT pero empieza 2 semanas después
    tend_dias = round(max(tend_blq1, tend_blq2) * 1.1)  # 10% overhead secuencial

    # Conexionado se puede hacer mientras se sigue tendiendo (overlap 40%)
    conex_blq1 = dias_paralelo(conex_mt, 2)
    conex_blq2 = max(dias_paralelo(conex_bt_g), dias_paralelo(conex_bt_m))
    conex_blq3 = dias_paralelo(conex_ctrl)
    # Bloque 1 se hace en paralelo con bloque 2 y 3
    conex_dias = round(conex_blq1 + conex_blq3)  # MT primero, luego control

    # Montaje de tableros y equipos (se hace en paralelo con tendido inicial)
    mont_tableros = max(dias_paralelo(mont_ccm, 2), dias_paralelo(mont_tgmt, 2))
    mont_motores = max(dias_paralelo(mont_motor_mt, 2),
                       dias_paralelo(mont_motor_bt_g),
                       dias_paralelo(mont_motor_bt_m + mont_motor_bt_c, CUADRILLAS))
    # Los tableros están en la ruta crítica, pero los motores no (se montan en campo mientras el tendido avanza)
    mont_dias = round(mont_tableros)  # Solo tableros en ruta crítica

    # Pruebas son secuenciales pero megger se hace en paralelo con últimas conexiones
    prueb_dias = round(dias_paralelo(prueb_mega) * 0.5) + prueb_energiz + prueb_precom + prueb_sat

    # Total (secuencial pero con solapamientos)
    # Montaje tableros → (Tendido || Montaje motores) → Conexionado → Pruebas
    total_dias = mont_dias + tend_dias + conex_dias + prueb_dias

    row += 1
    ws.merge_cells(f'A{row}:E{row}')
    ws.cell(row=row, column=1).value = 'TOTAL DÍAS ESTIMADO (CON PARALELIZACIÓN)'
    ws.cell(row=row, column=1).fill = COL_TOTAL_FILL
    ws.cell(row=row, column=1).font = COL_TOTAL_FONT
    ws.cell(row=row, column=1).alignment = Alignment(horizontal='right', vertical='center')
    for col_idx in range(1, 10):
        ws.cell(row=row, column=col_idx).fill = COL_TOTAL_FILL
        ws.cell(row=row, column=col_idx).font = COL_TOTAL_FONT
        ws.cell(row=row, column=col_idx).border = thin_border()

    ws.cell(row=row, column=6).value = (f'MONTAJE TABLEROS: {round(mont_dias)} + '
                                          f'TENDIDO: {round(tend_dias)} + '
                                          f'CONEXION: {round(conex_dias)} + '
                                          f'PRUEBAS: {round(prueb_dias)}')
    ws.cell(row=row, column=8).value = round(total_dias)
    ws.cell(row=row, column=9).value = f'≈ {round(total_dias)} días calendario con 3 cuadrillas'

    ws.row_dimensions[row].height = 22

    widths = [18, 45, 10, 10, 24, 14, 10, 14, 55]
    set_col_widths(ws, widths)
    ws.freeze_panes = 'A4'

    return ws, round(total_dias)

# ─────────────────────────────────────────────
# HOJA 5: CRONOGRAMA CAMPO
# ─────────────────────────────────────────────
def build_hoja5(wb, total_dias):
    ws = wb.create_sheet("5_Cronograma_Campo")
    ws.row_dimensions[1].height = 35
    ws.row_dimensions[2].height = 20

    # Título
    ws.merge_cells('A1:H1')
    t = ws['A1']
    t.value = 'CRONOGRAMA INSTALACIÓN ELÉCTRICA CPF-2 LA CALERA II - MAY 2027 → ENE 2028'
    t.fill = COL_HDR_FILL
    t.font = Font(color="FFFFFF", bold=True, size=12)
    t.alignment = Alignment(horizontal='center', vertical='center')

    headers = ['#', 'ACTIVIDAD', 'INICIO', 'FIN', 'DUR_DÍAS',
               'RESPONSABLE', 'HITO/CONSTRAINT', 'OBSERVACION']
    apply_header(ws, 2, headers)

    # Fechas clave
    d_inicio = date(2027, 5, 3)     # Inicio campo
    d_s4     = date(2027, 4, 7)     # Llegada Sala 4
    d_s3     = date(2027, 5, 3)     # Llegada Sala 3
    d_com    = date(2027, 7, 16)    # Inicio comisionado
    d_rfmc   = date(2027, 8, 15)    # RFMC estimado
    d_rfsu   = date(2028, 1, 31)    # RFSU

    # Calcular semanas de trabajo disponibles
    sem_disponibles = (d_rfsu - d_inicio).days // 7
    dias_disponibles = (d_rfsu - d_inicio).days

    activities = [
        # n°, Actividad, Inicio, Fin, Resp, Hito, Obs
        (1, 'Recepción y verificación materiales eléctricos (CCMs S3/S4)',
         date(2027, 4, 7), date(2027, 4, 18), 'AESA/ABB',
         'LLEGADA S4: 07-ABR-2027',
         'Inspección visual + megger de recepción CCMs, celdas MT'),
        (2, 'Montaje celdas MT: TGMT-001/003 y TRMT-001A/B (Sala 4)',
         date(2027, 4, 14), date(2027, 5, 9), 'ABB+AESA',
         'S4 en sitio ABR-2027',
         'Nivelación, interconexión barras, cableado de control interno'),
        (3, 'Montaje CCM-005/006/007 en Sala 3',
         date(2027, 5, 3), date(2027, 5, 24), 'ABB+AESA',
         'S3 en sitio MAY-2027',
         'Alineación, anclaje, busbars, cableado auxiliares'),
        (4, 'Instalación bandejas portacables principales (campo)',
         date(2027, 4, 21), date(2027, 5, 14), 'AESA',
         None,
         'Bandeja MC1, MC2, MC3 - trayectos principales SE3→SE4→campo'),
        (5, 'Tendido cables potencia MT (13.2kV/6.6kV)',
         date(2027, 5, 10), date(2027, 6, 7), 'AESA',
         None,
         'Alimentadores TGMT-001A/B desde 670-TGMT-001 | Req. herram. MT'),
        (6, 'Tendido cables potencia BT - motores y cargas grandes',
         date(2027, 5, 17), date(2027, 7, 4), 'AESA',
         None,
         '3x240, 3x150 bombas inyección | Tendido pesado en canaletas'),
        (7, 'Tendido cables potencia BT - cargas medianas y pequeñas',
         date(2027, 6, 7), date(2027, 7, 19), 'AESA',
         None,
         'Motores BT auxiliares, iluminación, tomacorrientes, calefacción'),
        (8, 'Tendido cables control y señales',
         date(2027, 6, 14), date(2027, 7, 19), 'AESA',
         None,
         'Control CCMs→motores, vibroswitches, señales VFDs'),
        (9, 'Tendido FO y Ethernet',
         date(2027, 6, 21), date(2027, 7, 5), 'AESA/ABB',
         None,
         'Red PMS SE3-SE4, señales digitales CPF-CPC-001'),
        (10, 'Conexionado y terminación cables potencia MT',
         date(2027, 6, 7), date(2027, 7, 5), 'ABB+AESA',
         'TÉCNICOS MT CERTIFICADOS',
         'Terminales premoldeadas 13.2kV | Req. ABB field service'),
        (11, 'Conexionado y terminación cables BT (potencia)',
         date(2027, 6, 21), date(2027, 7, 26), 'AESA',
         None,
         'Zapatas, terminales preaislados, prueba continuidad'),
        (12, 'Conexionado cables control (peinado + etiquetado)',
         date(2027, 7, 5), date(2027, 8, 2), 'AESA',
         None,
         'Incluye señales VFD MT, vibroswitches, HOA locales'),
        (13, 'Pruebas de aislación (megger) circuito por circuito',
         date(2027, 7, 19), date(2027, 8, 9), 'AESA',
         'HOLD POINT - ITP ELÉCTRICO',
         'Registro conforme ITP | Mínimo 1M Ω a 1kV / 10M Ω a 5kV'),
        (14, 'Energización progresiva: MT → Transformadores → BT',
         date(2027, 8, 9), date(2027, 8, 24), 'ABB+AESA+Cliente',
         'RFMC: ~15-AGO-2027',
         'Pre-requis: Megger OK | Prot. relés ajustadas | Personal Cliente'),
        (15, 'Pruebas funcionales precomisionado eléctrico',
         date(2027, 8, 18), date(2027, 9, 7), 'AESA+Cliente',
         None,
         'Loop check eléctrico, secuencia arranque motores, VFD tuning'),
        (16, 'SAT PMS en sitio - HOLD POINT per PIE/Contrato',
         date(2027, 10, 6), date(2027, 12, 7), 'ABB+Cliente',
         'HOLD POINT: OCT-DIC 2027',
         'Factory accepted, ahora SAT campo | 30 días ventana'),
        (17, 'Comisionado eléctrico integrado (con proceso)',
         date(2027, 9, 7), date(2027, 12, 15), 'AESA+ABB+Cliente',
         'INICIO COMISIONADO: 16-JUL-2027',
         'Energizado progresivo de procesos | Coordinado con Ops'),
        (18, 'RFSU - Ready For Start Up',
         date(2028, 1, 31), date(2028, 1, 31), 'TODOS',
         'HITO CONTRACTUAL: 31-ENE-2028',
         'Entrega al cliente | Todos los sistemas eléctricos operativos'),
    ]

    for i, act in enumerate(activities, 3):
        num, titulo, inicio, fin, resp, hito, obs = act
        dur = (fin - inicio).days + 1
        fill = COL_ALT_FILL if i % 2 == 0 else None

        # Colorear hitos especiales
        if hito and ('HOLD' in hito or 'RFMC' in hito or 'RFSU' in hito):
            fill = PatternFill("solid", fgColor="FFD966")

        values = [
            num, titulo,
            inicio.strftime('%d-%b-%Y'),
            fin.strftime('%d-%b-%Y'),
            dur, resp, hito or '', obs
        ]
        write_row(ws, i, values, fill=fill)
        ws.row_dimensions[i].height = 22

    # Análisis de compatibilidad
    last_row = i + 2

    ws.merge_cells(f'A{last_row}:H{last_row}')
    anal_title = ws.cell(row=last_row, column=1)
    anal_title.value = 'ANÁLISIS DE COMPATIBILIDAD CON RFSU 31-ENE-2028'
    anal_title.fill = PatternFill("solid", fgColor="2E75B6")
    anal_title.font = Font(color="FFFFFF", bold=True, size=11)
    anal_title.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[last_row].height = 24

    last_row += 1
    margen = dias_disponibles - total_dias
    margen_pct = round(margen / dias_disponibles * 100)
    if margen > 20:
        margen_estado = 'OK'
    elif margen >= -10:
        margen_estado = 'RIESGO'
    else:
        margen_estado = 'CRITICO'

    concl_estado = 'OK' if margen >= -10 else 'CRITICO'
    if margen >= 0:
        concl_txt = (f'Cronograma AJUSTADO con RFSU 31-ENE-2028: {margen} días de margen ({margen_pct}%). '
                     f'Requiere inicio 03-MAY-2027 y 3 cuadrillas ininterrumpidas sin imprevistos.')
    else:
        concl_txt = (f'ALERTA: Cronograma COMPROMETIDO con RFSU 31-ENE-2028: déficit de {abs(margen)} días. '
                     f'ACCIÓN: aumentar a 4 cuadrillas para el tendido BT Grande (17,470m es la ruta crítica) '
                     f'o ampliar bandejas prefabricadas para acelerar tendido.')

    fechas_analisis = [
        ('Días disponibles MAY-2027 → ENE-2028',
         f'{dias_disponibles} días = {sem_disponibles} semanas (03-MAY-2027 → 31-ENE-2028)', 'INFO'),
        ('Días estimados instalación (paralelizado 3 cuadrillas)',
         f'~{total_dias} días | Ruta crítica: Tendido BT Grande (17,470m / 50m/día/cuadrilla)', 'INFO'),
        ('Margen disponible',
         f'{margen} días de margen ({margen_pct}%) {"ACEPTABLE" if margen >= 0 else "DÉFICIT - VER ACCIONES"}',
         margen_estado),
        ('Ruta Crítica (CP) identificada',
         'Tendido BT Grande 3x240/3x150 (104 cables, 17,470m) → Conexionado MT → SAT PMS → Comisionado',
         'CRITICO'),
        ('Riesgo 1: Volumen tendido BT Grande',
         '104 cables / 17,470m de 3x240/3x150/3x95mm² → Ruta crítica. Opción: agregar 1 cuadrilla extra en este frente',
         'CRITICO'),
        ('Riesgo 2: Llegada Salas Eléctricas ABB',
         'S4: ABR-2027 | S3: MAY-2027 → Solo 9 meses para instalar 81,231m de cables',
         'RIESGO'),
        ('Riesgo 3: Cables MT 13.2kV (3x95, 18 cables, 2,110m)',
         'Técnicos ABB field service certificados requeridos | Coordinación con 3 meses de antelación',
         'RIESGO'),
        ('Riesgo 4: SAT PMS - HOLD POINT Contractual',
         'Ventana OCT-DIC 2027 fija por contrato/PIE - No admite deslizamiento - Precondiciona comisionado',
         'CRITICO'),
        ('Riesgo 5: Lista de cables incompleta (Rev.0)',
         'Notas del revisor: faltan cables HOA, comunicación CCM↔PMS, señales hardwired DOL, M-94213B',
         'RIESGO'),
        ('Acción Recomendada',
         'Agregar 1 cuadrilla BT para tendido (total 4) en el bloque MAY-JUL-2027 para recuperar 30 días',
         'RIESGO'),
        ('CONCLUSIÓN GENERAL', concl_txt, concl_estado),
    ]

    for analisis, valor, estado in fechas_analisis:
        fill_map = {'OK': COL_OK_FILL, 'RIESGO': PatternFill("solid", fgColor="FF9900"),
                    'CRITICO': PatternFill("solid", fgColor="FF0000"),
                    'INFO': PatternFill("solid", fgColor="BDD7EE")}
        font_map = {'OK': Font(color="FFFFFF", bold=True),
                    'RIESGO': Font(color="FFFFFF", bold=True),
                    'CRITICO': Font(color="FFFFFF", bold=True),
                    'INFO': Font(color="000000")}

        ws.merge_cells(f'A{last_row}:C{last_row}')
        c1 = ws.cell(row=last_row, column=1, value=analisis)
        c1.fill = fill_map.get(estado, COL_ALT_FILL)
        c1.font = font_map.get(estado, Font())
        c1.alignment = Alignment(horizontal='left', vertical='center')
        c1.border = thin_border()

        ws.merge_cells(f'D{last_row}:G{last_row}')
        c2 = ws.cell(row=last_row, column=4, value=valor)
        c2.fill = fill_map.get(estado, COL_ALT_FILL)
        c2.font = font_map.get(estado, Font())
        c2.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
        c2.border = thin_border()

        c3 = ws.cell(row=last_row, column=8, value=estado)
        c3.fill = fill_map.get(estado, COL_ALT_FILL)
        c3.font = font_map.get(estado, Font())
        c3.alignment = Alignment(horizontal='center', vertical='center')
        c3.border = thin_border()

        ws.row_dimensions[last_row].height = 24
        last_row += 1

    widths = [4, 52, 14, 14, 8, 20, 36, 55]
    set_col_widths(ws, widths)
    ws.freeze_panes = 'A3'

    compatible = (dias_disponibles - total_dias) >= -10  # hasta 10 días de déficit = ajustado
    return ws, compatible

# ─────────────────────────────────────────────
# HOJA 6: RESUMEN EJECUTIVO
# ─────────────────────────────────────────────
def build_hoja6(wb, cables, cargas, total_dias, compatible,
                total_m_cpf1, total_m_cpf2, total_cant):
    ws = wb.create_sheet("6_Resumen_Ejecutivo")

    # Título
    ws.merge_cells('A1:F1')
    t = ws['A1']
    t.value = 'RESUMEN EJECUTIVO - ANÁLISIS ELÉCTRICO CPF-2 LA CALERA II'
    t.fill = COL_HDR_FILL
    t.font = Font(color="FFFFFF", bold=True, size=14)
    t.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 40

    ws.merge_cells('A2:F2')
    s = ws['A2']
    s.value = 'Documento: ACAL-00102-LC-E-0001 Rev.0 | Unifilar: ACAL-00100-EE-E-0001-1 | Análisis: Junio 2026'
    s.fill = PatternFill("solid", fgColor="2E75B6")
    s.font = Font(color="FFFFFF", size=10)
    s.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[2].height = 18

    row = 4

    def section_title(ws, row, text):
        ws.merge_cells(f'A{row}:F{row}')
        c = ws.cell(row=row, column=1, value=text)
        c.fill = PatternFill("solid", fgColor="2E75B6")
        c.font = Font(color="FFFFFF", bold=True, size=11)
        c.alignment = Alignment(horizontal='left', vertical='center')
        c.border = thin_border()
        ws.row_dimensions[row].height = 22
        return row + 1

    def kv_row(ws, row, key, value, fill=None):
        c1 = ws.cell(row=row, column=1, value=key)
        c1.font = Font(bold=True)
        c1.fill = fill or COL_ALT_FILL
        c1.border = thin_border()
        c1.alignment = Alignment(vertical='center')
        ws.merge_cells(f'B{row}:F{row}')
        c2 = ws.cell(row=row, column=2, value=value)
        c2.fill = fill or PatternFill("solid", fgColor="F2F2F2")
        c2.border = thin_border()
        c2.alignment = Alignment(vertical='center', wrap_text=True)
        ws.row_dimensions[row].height = 18
        return row + 1

    # ─ Cables ─
    cpf2_cables = [c for c in cables if c.get('CPF') == 'CPF2']
    cpf1_cables = [c for c in cables if c.get('CPF') == 'CPF1']
    n_cpf2 = len(cpf2_cables)
    n_cpf1 = len(cpf1_cables)
    m_cpf2 = sum(longitud_float(c) for c in cpf2_cables)
    m_cpf1 = sum(longitud_float(c) for c in cpf1_cables)

    # Por tipo CPF-2
    tipos = {}
    for tipo in ['POTENCIA_MT', 'POTENCIA_BT', 'CONTROL', 'FO', 'ETHERNET', 'SEÑAL']:
        lst = [c for c in cpf2_cables if c.get('TIPO') == tipo]
        tipos[tipo] = {'cant': len(lst), 'metros': sum(longitud_float(c) for c in lst)}

    # Cargas por tipo
    from collections import Counter
    tipo_carga_cnt = Counter(c['TIPO_CARGA'] for c in cargas)

    row = section_title(ws, row, '1. VOLUMEN DE CABLES')
    row = kv_row(ws, row, 'Total cables en listado',
                 f'{total_cant} cables ({n_cpf1} CPF-1 + {n_cpf2} CPF-2)')
    row = kv_row(ws, row, 'Total metros CPF-2',
                 f'{round(m_cpf2):,} m ({n_cpf2} cables)')
    row = kv_row(ws, row, 'Total metros CPF-1',
                 f'{round(m_cpf1):,} m ({n_cpf1} cables)')
    row = kv_row(ws, row, 'Potencia MT (6.6/13.2kV) - CPF-2',
                 f'{tipos["POTENCIA_MT"]["cant"]} cables | {round(tipos["POTENCIA_MT"]["metros"]):,} m')
    row = kv_row(ws, row, 'Potencia BT (≤1kV) - CPF-2',
                 f'{tipos["POTENCIA_BT"]["cant"]} cables | {round(tipos["POTENCIA_BT"]["metros"]):,} m')
    row = kv_row(ws, row, 'Control + Señales - CPF-2',
                 f'{tipos["CONTROL"]["cant"] + tipos["SEÑAL"]["cant"]} cables | '
                 f'{round(tipos["CONTROL"]["metros"] + tipos["SEÑAL"]["metros"]):,} m')
    row = kv_row(ws, row, 'Fibra Óptica - CPF-2',
                 f'{tipos["FO"]["cant"]} cables | {round(tipos["FO"]["metros"]):,} m')
    row = kv_row(ws, row, 'Ethernet UTP - CPF-2',
                 f'{tipos["ETHERNET"]["cant"]} cables | {round(tipos["ETHERNET"]["metros"]):,} m')

    row += 1
    row = section_title(ws, row, '2. CARGAS ELÉCTRICAS CPF-2')
    row = kv_row(ws, row, 'Total equipos identificados CPF-2',
                 f'{len(cargas)} cargas eléctricas')
    for tipo_c, cnt in sorted(tipo_carga_cnt.items()):
        row = kv_row(ws, row, f'  - {tipo_c}', f'{cnt} unidades')

    row += 1
    row = section_title(ws, row, '3. ESTIMACIÓN INSTALACIÓN CAMPO')
    row = kv_row(ws, row, 'Cuadrillas eléctricas',
                 '3 cuadrillas x 4 personas = 12 personas en campo')
    row = kv_row(ws, row, 'Tiempo total estimado (paralelizado)',
                 f'~{total_dias} días calendario')
    row = kv_row(ws, row, 'Período de trabajo',
                 'MAY-2027 → ENE-2028 = 274 días disponibles')
    row = kv_row(ws, row, 'Margen disponible',
                 f'{274 - total_dias} días ({round((274-total_dias)/274*100, 0)}% del total)')
    row = kv_row(ws, row, 'Fecha inicio mínima recomendada',
                 '03-MAY-2027 (con llegada Sala 3)')

    row += 1
    row = section_title(ws, row, '4. HITOS CLAVE CPF-2')
    hitos = [
        ('Llegada Sala Eléctrica 4 (ABB)', '~07-ABR-2027'),
        ('Llegada Sala Eléctrica 3 (ABB)', '~03-MAY-2027'),
        ('Inicio precomisionado / montaje campo', '03-MAY-2027'),
        ('Inicio comisionado', '16-JUL-2027'),
        ('RFMC (Ready For Mechanical Completion)', '~15-AGO-2027'),
        ('SAT PMS en sitio (ventana)', 'OCT-DIC 2027'),
        ('RFSU (Ready For Start Up)', '31-ENE-2028 [CONTRACTUAL]'),
    ]
    for h, v in hitos:
        row = kv_row(ws, row, h, v)

    row += 1
    row = section_title(ws, row, '5. ALERTAS Y RIESGOS CRÍTICOS')
    alertas = [
        ('CRÍTICO', 'Cables MT 13.2kV requieren técnicos ABB certificados - Coordinar con 3 meses de antelación'),
        ('CRÍTICO', 'SAT PMS es HOLD POINT contractual - Ventana OCT-DIC 2027 no puede deslizarse'),
        ('RIESGO', 'Llegada tardía de salas eléctricas ABB comprime ventana de instalación a <9 meses'),
        ('RIESGO', 'Lista de cables Rev.0 incompleta - Faltan cables de comunicación CCM→PMS y señales hardwired motores DOL'),
        ('RIESGO', 'Bombas de inyección (M-PAZ-24210) a 460V con VFDs 1kV - Verificar coordinación de protecciones'),
        ('RIESGO', 'Cables 3x240 mm² a 460V (bombas inyección): verificar formación final con memoria de cálculo'),
        ('ALERTA', 'Distancias en lista de cables a confirmar con ingeniería - Nota en Rev.0 del documento'),
        ('ALERTA', 'Falta M-94213B en listado (nota del revisor del documento)'),
        ('ALERTA', 'Calefactores de motores (EAL-23310): incluir en circuito de calefacción con control thermostat'),
    ]
    for nivel, texto in alertas:
        fill_map = {'CRÍTICO': PatternFill("solid", fgColor="FF0000"),
                    'RIESGO': PatternFill("solid", fgColor="FF6600"),
                    'ALERTA': PatternFill("solid", fgColor="FFD966")}
        font_map = {'CRÍTICO': Font(color="FFFFFF", bold=True),
                    'RIESGO': Font(color="FFFFFF", bold=True),
                    'ALERTA': Font(color="000000", bold=True)}
        c1 = ws.cell(row=row, column=1, value=nivel)
        c1.fill = fill_map.get(nivel, COL_ALT_FILL)
        c1.font = font_map.get(nivel, Font())
        c1.alignment = Alignment(horizontal='center', vertical='center')
        c1.border = thin_border()
        ws.merge_cells(f'B{row}:F{row}')
        c2 = ws.cell(row=row, column=2, value=texto)
        c2.fill = PatternFill("solid", fgColor="FFF2CC") if nivel == 'ALERTA' else PatternFill("solid", fgColor="F2F2F2")
        c2.alignment = Alignment(vertical='center', wrap_text=True)
        c2.border = thin_border()
        ws.row_dimensions[row].height = 22
        row += 1

    row += 1
    # Conclusión
    fill_conc = COL_OK_FILL if compatible else COL_ALERTA
    ws.merge_cells(f'A{row}:F{row}')
    c = ws.cell(row=row, column=1)
    margen_dias = 274 - total_dias
    if margen_dias >= 20:
        estado_txt = 'COMPATIBLE'
        detalle = f'Margen holgado de {margen_dias} días ({round(margen_dias/274*100)}%). Riesgo bajo.'
    elif margen_dias >= 0:
        estado_txt = 'AJUSTADO - SIN MARGEN'
        detalle = f'Solo {margen_dias} días de margen ({round(margen_dias/274*100)}%). ACCIÓN: 4 cuadrillas para tendido BT Grande (ruta crítica).'
    else:
        estado_txt = 'COMPROMETIDO'
        detalle = (f'Déficit de {abs(margen_dias)} días. ACCIÓN URGENTE: agregar 1 cuadrilla extra para tendido BT '
                   f'(reduce tendido_dias en ~{abs(margen_dias)+10} días).')
    c.value = (
        f'CONCLUSIÓN: Cronograma {estado_txt} con RFSU 31-ENE-2028. '
        f'Tiempo estimado: ~{total_dias} días vs {274} disponibles con 3 cuadrillas. {detalle} '
        f'RUTA CRÍTICA: Tendido cables BT Grande (104 cables, 17,470m). '
        f'RIESGOS: (1) Lista Rev.0 incompleta - estimado podría crecer 10-15%. '
        f'(2) Salas ABB deben llegar antes de MAY-2027. '
        f'(3) SAT PMS es HOLD POINT contractual OCT-DIC 2027.'
    )
    c.fill = fill_conc
    c.font = Font(color="FFFFFF", bold=True, size=11)
    c.alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    c.border = thin_border()
    ws.row_dimensions[row].height = 55

    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 70
    for col in ['C', 'D', 'E', 'F']:
        ws.column_dimensions[col].width = 10

    return ws

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    print("=" * 60)
    print("ANÁLISIS ELÉCTRICO CPF-2 LA CALERA II - VACA MUERTA")
    print("=" * 60)

    # ── 1. PARSEAR CABLES ──
    print("\n[1/6] Parseando lista de cables...")
    cables = parse_cables(PDF_CABLES)

    # Enriquecer con CPF y TIPO
    for c in cables:
        c['CPF']  = determine_cpf(c)
        c['TIPO'] = determine_tipo(c)

    cpf1_cables = [c for c in cables if c['CPF'] == 'CPF1']
    cpf2_cables = [c for c in cables if c['CPF'] == 'CPF2']

    print(f"  Total cables extraídos: {len(cables)}")
    print(f"  CPF-1: {len(cpf1_cables)} cables | {round(sum(longitud_float(c) for c in cpf1_cables)):,} m")
    print(f"  CPF-2: {len(cpf2_cables)} cables | {round(sum(longitud_float(c) for c in cpf2_cables)):,} m")

    # Tipos CPF-2
    from collections import Counter
    tipos_cpf2 = Counter(c['TIPO'] for c in cpf2_cables)
    print("  Tipos CPF-2:", dict(tipos_cpf2))

    # ── 2. GENERAR EXCEL ──
    print("\n[2/6] Creando workbook Excel...")
    wb = openpyxl.Workbook()
    # Eliminar hoja por defecto
    if 'Sheet' in wb.sheetnames:
        del wb['Sheet']

    # ── HOJA 1 ──
    print("[3/6] Generando Hoja 1: Cables Completo...")
    build_hoja1(wb, cables)

    # ── HOJA 2 ──
    print("[4/6] Generando Hoja 2: Resumen Formaciones...")
    _, total_cant, total_m_all, total_m_cpf1, total_m_cpf2 = build_hoja2(wb, cables)

    # ── HOJA 3 ──
    print("[5/6] Generando Hoja 3: Cargas CPF-2...")
    _, cargas = build_hoja3(wb, cables)
    print(f"  Cargas CPF-2 identificadas: {len(cargas)}")

    # ── HOJA 4 ──
    print("[6/6] Generando Hoja 4: Estimación instalación...")
    _, total_dias = build_hoja4(wb, cables, cargas)

    # ── HOJA 5 ──
    print("      Generando Hoja 5: Cronograma campo...")
    _, compatible = build_hoja5(wb, total_dias)

    # ── HOJA 6 ──
    print("      Generando Hoja 6: Resumen ejecutivo...")
    build_hoja6(wb, cables, cargas, total_dias, compatible,
                total_m_cpf1, total_m_cpf2, total_cant)

    # Ordenar hojas
    for i, name in enumerate(['6_Resumen_Ejecutivo', '1_Cables_Completo',
                               '2_Resumen_Formaciones', '3_Cargas_CPF2',
                               '4_Estimacion_Instalacion', '5_Cronograma_Campo']):
        if name in wb.sheetnames:
            idx = wb.sheetnames.index(name)
            wb.move_sheet(name, offset=i - idx)

    # Activa Hoja 6 al abrir
    wb.active = wb['6_Resumen_Ejecutivo']

    # ── GUARDAR ──
    os.makedirs(os.path.dirname(OUTPUT_XLSX), exist_ok=True)
    wb.save(OUTPUT_XLSX)

    # ── RESUMEN CONSOLA ──
    m_cpf2 = sum(longitud_float(c) for c in cpf2_cables)
    m_cpf1 = sum(longitud_float(c) for c in cpf1_cables)

    print("\n" + "=" * 60)
    print("RESUMEN ESTADÍSTICO")
    print("=" * 60)
    print(f"Total cables extraídos: {len(cables)}")
    print(f"  CPF-1: {len(cpf1_cables)} cables | {round(m_cpf1):,} m")
    print(f"  CPF-2: {len(cpf2_cables)} cables | {round(m_cpf2):,} m")
    print(f"Cargas CPF-2: {len(cargas)} equipos")
    print(f"Tiempo instalación estimado: ~{total_dias} días con 3 cuadrillas")
    margen = 274 - total_dias
    if margen >= 20:
        compat_str = f'SI (margen {margen} días)'
    elif margen >= 0:
        compat_str = f'AJUSTADO - solo {margen} días de margen - agregar 1 cuadrilla BT'
    else:
        compat_str = f'NO - déficit {abs(margen)} días - ACCIÓN REQUERIDA'
    print(f"Compatible con RFSU 31-ENE-2028: {compat_str}")
    print(f"\nGenerado: {OUTPUT_XLSX}")
    print("=" * 60)

if __name__ == '__main__':
    main()
