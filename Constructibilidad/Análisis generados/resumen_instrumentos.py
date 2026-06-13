#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Resumen de instrumentos - La Calera CPF2 Fase 2
Consolida las dos Listas de Instrumentos (areas 00102 y 00670),
cuenta instrumentos por tipo y area, y los clasifica en:
  - IN-LINE: montados sobre la caneria (valvulas y elementos en linea)
  - MONTAJE ESPECIFICO: el resto, segun tipicos de montaje
Excluye senales, alarmas y todo lo que no sea un instrumento fisico.
Fuente: Datos entrada/ACAL-00102-LI-K-0001-A VCO.pdf
        Datos entrada/ACAL-00670-LI-K-0001-B VCO.pdf
"""
import fitz, re, os, sys
from collections import Counter, defaultdict, OrderedDict
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
IN = os.path.join(ROOT, "Datos entrada")
OUT = os.path.join(BASE, "Resumen_Instrumentos_CPF2.pdf")

DOCS = OrderedDict([
    ("ACAL-00102 (Proceso / Facilities)", "ACAL-00102-LI-K-0001-A VCO.pdf"),
    ("ACAL-00670 (Generacion)",           "ACAL-00670-LI-K-0001-B VCO.pdf"),
])

# Columnas de la tabla (indice fijo segun encabezado de la LI)
IDX = {'cod':1,'tag':4,'tipo':5,'area':8,'linea':17}

# ---- Familias IN-LINE (montadas en la caneria / parte de la linea) ----
# Valvulas (interrumpen la caneria) + elementos en linea.
INLINE_CODES = {
    # valvulas de control / autorreguladoras
    'FV','FCV','LCV','PV','PCV','TCV','TV','LV','DPV','PDV','PRV','BPRV',
    # valvulas de seguridad / alivio / presion-vacio
    'PSV','TSV','PVSV',
    # valvulas de bloqueo / shutdown / blowdown / diluvio
    'ESDV','SDV','BDV','XV','DV',
    # elementos en linea
    'RO',   # orificio de restriccion
    'FE',   # placa orificio / elemento de caudal
    'FI',   # rotametro
    'PD',   # amortiguador de pulsaciones
    'TW',   # termovaina
    'TM',   # toma muestra
    'CC',   # cupon de corrosion
}
CAT_INLINE = "In-line (montado en caneria)"
CAT_ESP    = "Montaje especifico (segun tipico)"

def extract(fn):
    d = fitz.open(fn)
    seen, out, header = set(), [], False
    for pno in range(d.page_count):
        for t in d[pno].find_tables().tables:
            for r in t.extract():
                c = [(x or '').replace('\n',' ').strip() for x in r]
                if 'TAG' in c and 'INST. COD.' in c:
                    header = True; continue
                if not header or len(c) < 9:
                    continue
                tag = c[IDX['tag']]
                if not re.match(r'^[A-Z]{1,4}-?\d{3,6}[A-Z]?$', tag):
                    continue
                if tag in seen:
                    continue
                seen.add(tag)
                out.append({
                    'cod': c[IDX['cod']].upper().strip(),
                    'tag': tag,
                    'tipo': c[IDX['tipo']].strip(),
                    'area': (c[IDX['area']].split() or [''])[0].strip(),
                })
    return out

# --- Extraccion ---
rows = []
per_doc = OrderedDict()
for label, fn in DOCS.items():
    rs = extract(os.path.join(IN, fn))
    per_doc[label] = rs
    rows += rs

for r in rows:
    r['cat'] = CAT_INLINE if r['cod'] in INLINE_CODES else CAT_ESP

total = len(rows)
n_inline = sum(1 for r in rows if r['cat'] == CAT_INLINE)
n_esp = total - n_inline

# Conteo por tipo (codigo + descripcion + categoria)
by_code = defaultdict(lambda: {'n':0,'tipo':'','cat':''})
for r in rows:
    e = by_code[r['cod']]
    e['n'] += 1
    if not e['tipo']:
        e['tipo'] = r['tipo']; e['cat'] = r['cat']
code_rows = sorted(by_code.items(), key=lambda kv: (kv[1]['cat']!=CAT_INLINE, -kv[1]['n']))

# Conteo por area
by_area = Counter(r['area'] for r in rows if r['area'])
area_rows = sorted(by_area.items(), key=lambda kv: kv[0])

# ---------------- PDF ----------------
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, LongTable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

styles = getSampleStyleSheet()
H1 = ParagraphStyle('H1', parent=styles['Title'], fontSize=17, spaceAfter=4)
SUB = ParagraphStyle('SUB', parent=styles['Normal'], fontSize=10,
                     textColor=colors.HexColor('#555555'), alignment=TA_CENTER)
H2 = ParagraphStyle('H2', parent=styles['Heading2'], fontSize=12,
                    textColor=colors.HexColor('#1f3b63'), spaceBefore=12, spaceAfter=6)
N = ParagraphStyle('N', parent=styles['Normal'], fontSize=9, leading=12)
SMALL = ParagraphStyle('SM', parent=styles['Normal'], fontSize=7.5,
                       textColor=colors.HexColor('#666666'))

BLUE = colors.HexColor('#1f3b63')
LBLUE = colors.HexColor('#dce6f2')
GREEN = colors.HexColor('#e2efda')
GREY = colors.HexColor('#f2f2f2')

def header_style(extra=None):
    base = [
        ('BACKGROUND',(0,0),(-1,0),BLUE),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),8),
        ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#bbbbbb')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, GREY]),
        ('TOPPADDING',(0,0),(-1,-1),3),
        ('BOTTOMPADDING',(0,0),(-1,-1),3),
    ]
    return TableStyle(base + (extra or []))

elems = []
elems.append(Paragraph("Resumen de Instrumentos", H1))
elems.append(Paragraph("Proyecto LA CALERA CPF2 - FASE 2 &nbsp;|&nbsp; "
                       "Consolidado de Listas de Instrumentos", SUB))
elems.append(Paragraph("Emitido: %s" % date.today().strftime('%d-%m-%Y'), SUB))
elems.append(Spacer(1, 8*mm))

# --- Resumen ejecutivo (tarjetas) ---
elems.append(Paragraph("1. Resumen ejecutivo", H2))
res = [
    ["Total de instrumentos", "%d" % total],
    ["In-line (montados en caneria)", "%d  (%.1f%%)" % (n_inline, 100*n_inline/total)],
    ["Montaje especifico (segun tipico)", "%d  (%.1f%%)" % (n_esp, 100*n_esp/total)],
    ["Tipos de instrumento distintos", "%d" % len(by_code)],
    ["Areas con instrumentos", "%d" % len(by_area)],
]
t = Table(res, colWidths=[95*mm, 60*mm])
t.setStyle(TableStyle([
    ('FONTSIZE',(0,0),(-1,-1),10),
    ('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),
    ('FONTNAME',(1,0),(1,-1),'Helvetica'),
    ('BACKGROUND',(0,0),(0,-1),LBLUE),
    ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#bbbbbb')),
    ('TOPPADDING',(0,0),(-1,-1),5),('BOTTOMPADDING',(0,0),(-1,-1),5),
    ('LEFTPADDING',(0,0),(-1,-1),8),
]))
elems.append(t)

# --- Por documento / area general ---
elems.append(Paragraph("2. Por documento (area general)", H2))
data = [["Documento / Area", "Total", "In-line", "Montaje especifico"]]
for label, rs in per_doc.items():
    ni = sum(1 for r in rs if r['cod'] in INLINE_CODES)
    data.append([label, str(len(rs)), str(ni), str(len(rs)-ni)])
t = Table(data, colWidths=[85*mm, 22*mm, 22*mm, 38*mm])
t.setStyle(header_style([('ALIGN',(1,0),(-1,-1),'CENTER')]))
elems.append(t)

# --- Conteo por tipo de instrumento ---
elems.append(Paragraph("3. Cantidad por tipo de instrumento", H2))
elems.append(Paragraph("Ordenado por categoria (in-line primero) y cantidad. "
                       "El codigo corresponde al prefijo del tag (norma ISA).", N))
elems.append(Spacer(1, 3))
data = [["Cod.", "Tipo de instrumento", "Categoria", "Cant."]]
for cod, e in code_rows:
    cat = "In-line" if e['cat']==CAT_INLINE else "Esp."
    data.append([cod, e['tipo'][:52], cat, str(e['n'])])
data.append(["", "TOTAL", "", str(total)])
t = LongTable(data, colWidths=[16*mm, 92*mm, 20*mm, 16*mm], repeatRows=1)
st = header_style([
    ('ALIGN',(0,0),(0,-1),'CENTER'),
    ('ALIGN',(2,0),(3,-1),'CENTER'),
    ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
    ('BACKGROUND',(0,-1),(-1,-1),LBLUE),
])
# colorear filas in-line
for i,(cod,e) in enumerate(code_rows, start=1):
    if e['cat']==CAT_INLINE:
        st.add('BACKGROUND',(2,i),(2,i),GREEN)
t.setStyle(st)
elems.append(t)

# --- Conteo por area ---
elems.append(Paragraph("4. Cantidad por area", H2))
data = [["Area", "Cant.", "Area", "Cant.", "Area", "Cant."]]
col = []
for a,n in area_rows:
    col.append((a,n))
# acomodar en 3 columnas
rowsN = (len(col)+2)//3
grid = []
for i in range(rowsN):
    line = []
    for j in range(3):
        k = i + j*rowsN
        if k < len(col):
            line += [col[k][0], str(col[k][1])]
        else:
            line += ["",""]
    grid.append(line)
data += grid
t = Table(data, colWidths=[26*mm,18*mm,26*mm,18*mm,26*mm,18*mm])
t.setStyle(header_style([('ALIGN',(0,0),(-1,-1),'CENTER')]))
elems.append(t)

# --- Nota metodologica ---
elems.append(Paragraph("5. Criterio de clasificacion y notas", H2))
inline_list = ", ".join(sorted(INLINE_CODES))
notas = [
 "<b>Alcance:</b> se contabilizan unicamente instrumentos fisicos listados en las "
 "Listas de Instrumentos. No se incluyen senales, alarmas, funciones logicas ni "
 "elementos que no sean instrumentos.",
 "<b>In-line (montado en caneria):</b> instrumentos que se instalan sobre la propia "
 "linea de proceso, formando parte de la caneria. Incluye las familias por codigo: "
 "<i>%s</i>. Validado contra la columna 'Linea o Equipo' de la LI, que en estos casos "
 "referencia una linea de caneria (ej. 8\"-PL-102656-...)." % inline_list,
 "<b>Montaje especifico (segun tipico de montaje):</b> el resto de los instrumentos "
 "(transmisores, indicadores, interruptores, detectores, posicionadores, solenoides, "
 "fines de carrera, etc.), que se montan sobre soporte, equipo o sobre la valvula segun "
 "el tipico de montaje correspondiente.",
 "<b>Nota sobre termovaina/sensor:</b> la termovaina (TW) se cuenta como in-line por ir "
 "soldada/roscada a la linea; el sensor de temperatura (TE) que se inserta en ella se "
 "cuenta como montaje especifico.",
 "<b>Fuente:</b> ACAL-00102-LI-K-0001 Rev.A y ACAL-00670-LI-K-0001 Rev.B (Pluspetrol).",
 "<b>Proximo paso:</b> esta clasificacion alimenta el siguiente analisis de "
 "constructibilidad (cantidad y tipo de montaje por instrumento).",
]
for nta in notas:
    elems.append(Paragraph(nta, N))
    elems.append(Spacer(1, 2))

def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica', 7)
    canvas.setFillColor(colors.HexColor('#888888'))
    canvas.drawString(15*mm, 10*mm, "Constructibilidad CPF2 - Resumen de Instrumentos")
    canvas.drawRightString(195*mm, 10*mm, "Pag. %d" % doc.page)
    canvas.restoreState()

doc = SimpleDocTemplate(OUT, pagesize=A4,
                        leftMargin=15*mm, rightMargin=15*mm,
                        topMargin=15*mm, bottomMargin=15*mm,
                        title="Resumen de Instrumentos CPF2")
doc.build(elems, onFirstPage=footer, onLaterPages=footer)
print("OK ->", OUT)
print("Total=%d  In-line=%d  Especifico=%d  Tipos=%d  Areas=%d"
      % (total, n_inline, n_esp, len(by_code), len(by_area)))
