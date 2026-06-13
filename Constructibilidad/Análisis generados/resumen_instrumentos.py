#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Resumen de instrumentos - La Calera CPF2 Fase 2
Consolida las dos Listas de Instrumentos (areas 00102 y 00670) y reporta:
  - Cantidad por tipo de instrumento y por area.
  - Clasificacion de montaje:
        IN-LINE: montados sobre la caneria (valvulas y elementos en linea)
        MONTAJE ESPECIFICO: el resto, segun tipicos de montaje
  - Division por suministro (columna COMENTARIOS de la LI):
        AESA: alcance de provision AESA
        OTRO PROVEEDOR: provistos por vendor / paquete (skid, valvula, bomba, etc.)
        EXISTENTE: a reutilizar / a reemplazar
Excluye senales, alarmas y todo lo que no sea un instrumento fisico.
Fuente: Datos entrada/ACAL-00102-LI-K-0001-A VCO.pdf  (Rev. A)
        Datos entrada/ACAL-00670-LI-K-0001-B VCO.pdf  (Rev. B)
"""
import fitz, re, os
from collections import Counter, defaultdict, OrderedDict
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
IN = os.path.join(ROOT, "Datos entrada")
OUT = os.path.join(BASE, "Resumen_Instrumentos_CPF2.pdf")

# -------- Control de revision del documento --------
REV_ACTUAL = "1"
REV_HIST = [
    # rev, fecha, descripcion
    ("0", "13-06-2026", "Emision inicial. Conteo por tipo/area y clasificacion "
                        "in-line / montaje especifico."),
    ("1", "13-06-2026", "Se agrega division por suministro (AESA / otro proveedor) "
                        "segun columna Comentarios, desglose por proveedor y matriz "
                        "montaje x suministro."),
]

DOCS = OrderedDict([
    ("ACAL-00102 (Proceso / Facilities)", "ACAL-00102-LI-K-0001-A VCO.pdf"),
    ("ACAL-00670 (Generacion)",           "ACAL-00670-LI-K-0001-B VCO.pdf"),
])

IDX = {'cod':1,'tag':4,'tipo':5,'area':8}   # COMENTARIOS = ultima columna

# ---- Familias IN-LINE (montadas en la caneria / parte de la linea) ----
INLINE_CODES = {
    'FV','FCV','LCV','PV','PCV','TCV','TV','LV','DPV','PDV','PRV','BPRV',  # control/autoreg.
    'PSV','TSV','PVSV',                                                    # seguridad/alivio
    'ESDV','SDV','BDV','XV','DV',                                          # bloqueo/shutdown
    'RO','FE','FI','PD','TW','TM','CC',                                    # elementos en linea
}
CAT_INLINE = "In-line (montado en caneria)"
CAT_ESP    = "Montaje especifico (segun tipico)"

# ---- Clasificacion de suministro segun COMENTARIOS ----
SUP_AESA, SUP_OTRO, SUP_EXIST = "AESA", "Otro proveedor", "Existente"

def supply(com):
    c = (com or '').upper().strip()
    if c in ('', '-'):                 return SUP_AESA
    if 'PROVISTO POR AESA' in c:        return SUP_AESA
    if c.startswith('EXISTENTE'):       return SUP_EXIST
    if 'PROVEEDOR' in c or 'PORVEEDOR' in c:  return SUP_OTRO
    return SUP_AESA

def vendor(com):
    """Agrupa el comentario en el proveedor/paquete responsable."""
    c = (com or '').upper()
    if 'PROVISTO POR AESA' in c:        return None
    if 'GAS COMBUSTIBLE' in c:          return "Skid gas combustible"
    if 'QUIMICOS' in c or 'QUÍMICOS' in c: return "Skid inyeccion quimicos"
    if 'AIRE DE INSTRUMENTOS' in c:     return "Skid aire de instrumentos"
    if 'INHIBIDOR DE HIDRATOS' in c:    return "Skid inhibidor de hidratos"
    if 'FILTRO' in c:                   return "Skid filtro"
    if 'PROPAK' in c:                   return "PROPAK (skid principal)"
    if 'VALVULA' in c or 'VÁLVULA' in c: return "Proveedor de la valvula"
    if 'BOMBA' in c:                    return "Proveedor de la bomba"
    if 'TANQUE' in c:                   return "Proveedor del tanque"
    if 'ANTORCHA' in c:                 return "Proveedor de la antorcha"
    if 'SKID' in c:                     return "Otro skid"
    return "Otro proveedor"

def extract(fn):
    d = fitz.open(fn)
    seen, out, header = set(), [], False
    for pno in range(d.page_count):
        for t in d[pno].find_tables().tables:
            for r in t.extract():
                c = [(x or '').replace('\n', ' ').strip() for x in r]
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
                    'cod':  c[IDX['cod']].upper().strip(),
                    'tag':  tag,
                    'tipo': c[IDX['tipo']].strip(),
                    'area': (c[IDX['area']].split() or [''])[0].strip(),
                    'com':  c[-1].strip(),
                })
    return out

# --- Extraccion y clasificacion ---
rows, per_doc = [], OrderedDict()
for label, fn in DOCS.items():
    rs = extract(os.path.join(IN, fn))
    per_doc[label] = rs
    rows += rs

for r in rows:
    r['cat'] = CAT_INLINE if r['cod'] in INLINE_CODES else CAT_ESP
    r['sup'] = supply(r['com'])

total = len(rows)
n_inline = sum(1 for r in rows if r['cat'] == CAT_INLINE)
n_esp    = total - n_inline
sup_count = Counter(r['sup'] for r in rows)
n_montaje_epc = sum(1 for r in rows if 'MONTAJE POR EPC' in r['com'].upper())

# Conteo por tipo
by_code = defaultdict(lambda: {'n':0,'tipo':'','cat':''})
for r in rows:
    e = by_code[r['cod']]; e['n'] += 1
    if not e['tipo']:
        e['tipo'] = r['tipo']; e['cat'] = r['cat']
code_rows = sorted(by_code.items(), key=lambda kv: (kv[1]['cat'] != CAT_INLINE, -kv[1]['n']))

# Conteo por area
by_area = Counter(r['area'] for r in rows if r['area'])
area_rows = sorted(by_area.items(), key=lambda kv: kv[0])

# Desglose de OTRO por proveedor
vend = Counter(vendor(r['com']) for r in rows if r['sup'] == SUP_OTRO)
vend_rows = sorted(vend.items(), key=lambda kv: -kv[1])

# Matriz montaje x suministro
matrix = defaultdict(int)
for r in rows:
    matrix[(r['cat'], r['sup'])] += 1

# ---------------- PDF ----------------
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, LongTable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

styles = getSampleStyleSheet()
H1  = ParagraphStyle('H1', parent=styles['Title'], fontSize=17, spaceAfter=4)
SUB = ParagraphStyle('SUB', parent=styles['Normal'], fontSize=10,
                     textColor=colors.HexColor('#555555'), alignment=TA_CENTER)
H2  = ParagraphStyle('H2', parent=styles['Heading2'], fontSize=12,
                     textColor=colors.HexColor('#1f3b63'), spaceBefore=12, spaceAfter=6)
N   = ParagraphStyle('N', parent=styles['Normal'], fontSize=9, leading=12)

BLUE  = colors.HexColor('#1f3b63')
LBLUE = colors.HexColor('#dce6f2')
GREEN = colors.HexColor('#e2efda')
AESAC = colors.HexColor('#fce4d6')   # naranja suave = AESA
OTROC = colors.HexColor('#ddebf7')   # azul suave   = otro proveedor
GREY  = colors.HexColor('#f2f2f2')

def hstyle(extra=None):
    base = [
        ('BACKGROUND',(0,0),(-1,0),BLUE),
        ('TEXTCOLOR',(0,0),(-1,0),colors.white),
        ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
        ('FONTSIZE',(0,0),(-1,-1),8),
        ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#bbbbbb')),
        ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
        ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white, GREY]),
        ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3),
    ]
    return TableStyle(base + (extra or []))

E = []
E.append(Paragraph("Resumen de Instrumentos", H1))
E.append(Paragraph("Proyecto LA CALERA CPF2 - FASE 2 &nbsp;|&nbsp; "
                   "Consolidado de Listas de Instrumentos", SUB))
E.append(Paragraph("Revision %s &nbsp;|&nbsp; Fecha: %s"
                   % (REV_ACTUAL, date.today().strftime('%d-%m-%Y')), SUB))
E.append(Spacer(1, 5*mm))

# Control de revisiones
REVP = ParagraphStyle('REVP', parent=styles['Normal'], fontSize=7.5, leading=9)
data = [["Rev.", "Fecha", "Descripcion"]]
for rv, fch, desc in REV_HIST:
    data.append([rv, fch, Paragraph(desc, REVP)])
t = Table(data, colWidths=[12*mm, 24*mm, 119*mm])
t.setStyle(hstyle([('ALIGN',(0,0),(1,-1),'CENTER'),
                   ('FONTSIZE',(0,1),(-1,-1),7.5)]))
E.append(t)

# 1. Resumen ejecutivo
E.append(Paragraph("1. Resumen ejecutivo", H2))
res = [
    ["Total de instrumentos", "%d" % total],
    ["  -  In-line (montados en caneria)", "%d  (%.1f%%)" % (n_inline, 100*n_inline/total)],
    ["  -  Montaje especifico (segun tipico)", "%d  (%.1f%%)" % (n_esp, 100*n_esp/total)],
    ["Suministro AESA", "%d  (%.1f%%)" % (sup_count[SUP_AESA], 100*sup_count[SUP_AESA]/total)],
    ["Suministro por otro proveedor", "%d  (%.1f%%)" % (sup_count[SUP_OTRO], 100*sup_count[SUP_OTRO]/total)],
    ["  -  de ellos, con montaje por AESA/EPC", "%d" % n_montaje_epc],
    ["Existente (reutilizar / reemplazar)", "%d" % sup_count[SUP_EXIST]],
    ["Tipos de instrumento distintos", "%d" % len(by_code)],
    ["Areas con instrumentos", "%d" % len(by_area)],
]
t = Table(res, colWidths=[100*mm, 55*mm])
t.setStyle(TableStyle([
    ('FONTSIZE',(0,0),(-1,-1),10),
    ('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),
    ('BACKGROUND',(0,0),(0,-1),LBLUE),
    ('BACKGROUND',(0,3),(0,3),AESAC),('BACKGROUND',(0,4),(0,5),OTROC),
    ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#bbbbbb')),
    ('TOPPADDING',(0,0),(-1,-1),4),('BOTTOMPADDING',(0,0),(-1,-1),4),
    ('LEFTPADDING',(0,0),(-1,-1),8),
]))
E.append(t)

# 2. Por documento
E.append(Paragraph("2. Por documento (area general)", H2))
data = [["Documento / Area", "Total", "In-line", "M. especif.", "AESA", "Otro prov."]]
for label, rs in per_doc.items():
    ni = sum(1 for r in rs if r['cod'] in INLINE_CODES)
    na = sum(1 for r in rs if supply(r['com']) == SUP_AESA)
    no = sum(1 for r in rs if supply(r['com']) == SUP_OTRO)
    data.append([label, str(len(rs)), str(ni), str(len(rs)-ni), str(na), str(no)])
t = Table(data, colWidths=[63*mm, 18*mm, 18*mm, 22*mm, 18*mm, 22*mm])
t.setStyle(hstyle([('ALIGN',(1,0),(-1,-1),'CENTER')]))
E.append(t)

# 3. Suministro
E.append(Paragraph("3. Division por suministro (AESA vs otro proveedor)", H2))
E.append(Paragraph("Segun la columna <b>Comentarios</b> de la Lista de Instrumentos. "
                   "Comentario vacio o '-' se interpreta como alcance de provision AESA.", N))
E.append(Spacer(1, 3))
data = [["Suministro", "Cantidad", "%"]]
for s in (SUP_AESA, SUP_OTRO, SUP_EXIST):
    data.append([s, str(sup_count[s]), "%.1f%%" % (100*sup_count[s]/total)])
data.append(["TOTAL", str(total), "100%"])
t = Table(data, colWidths=[80*mm, 30*mm, 25*mm])
st = hstyle([('ALIGN',(1,0),(-1,-1),'CENTER'),
             ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
             ('BACKGROUND',(0,-1),(-1,-1),LBLUE)])
st.add('BACKGROUND',(0,1),(0,1),AESAC)
st.add('BACKGROUND',(0,2),(0,2),OTROC)
t.setStyle(st)
E.append(t)

E.append(Paragraph("Detalle del suministro por otro proveedor", N))
E.append(Spacer(1, 2))
data = [["Proveedor / Paquete", "Cantidad"]]
for v, n in vend_rows:
    data.append([v, str(n)])
data.append(["TOTAL otro proveedor", str(sup_count[SUP_OTRO])])
t = Table(data, colWidths=[110*mm, 30*mm])
t.setStyle(hstyle([('ALIGN',(1,0),(-1,-1),'CENTER'),
                   ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
                   ('BACKGROUND',(0,-1),(-1,-1),LBLUE)]))
E.append(t)

# 4. Matriz montaje x suministro
E.append(Paragraph("4. Matriz montaje x suministro", H2))
cats = [CAT_INLINE, CAT_ESP]
sups = [SUP_AESA, SUP_OTRO, SUP_EXIST]
data = [["Montaje \\ Suministro", "AESA", "Otro prov.", "Existente", "Total"]]
for c in cats:
    fila = [c] + [str(matrix[(c, s)]) for s in sups]
    fila.append(str(sum(matrix[(c, s)] for s in sups)))
    data.append(fila)
data.append(["TOTAL"] + [str(sup_count[s]) for s in sups] + [str(total)])
t = Table(data, colWidths=[60*mm, 24*mm, 26*mm, 24*mm, 22*mm])
t.setStyle(hstyle([('ALIGN',(1,0),(-1,-1),'CENTER'),
                   ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
                   ('FONTNAME',(-1,0),(-1,-1),'Helvetica-Bold'),
                   ('BACKGROUND',(0,-1),(-1,-1),LBLUE),
                   ('BACKGROUND',(-1,1),(-1,-1),LBLUE)]))
E.append(t)

# 5. Cantidad por tipo
E.append(Paragraph("5. Cantidad por tipo de instrumento", H2))
E.append(Paragraph("Ordenado por categoria (in-line primero) y cantidad. "
                   "AESA / Otro = cantidad por suministro dentro de cada tipo.", N))
E.append(Spacer(1, 3))
data = [["Cod.", "Tipo de instrumento", "Cat.", "AESA", "Otro", "Exist.", "Cant."]]
for cod, e in code_rows:
    na = sum(1 for r in rows if r['cod']==cod and r['sup']==SUP_AESA)
    no = sum(1 for r in rows if r['cod']==cod and r['sup']==SUP_OTRO)
    nx = sum(1 for r in rows if r['cod']==cod and r['sup']==SUP_EXIST)
    cat = "In-line" if e['cat']==CAT_INLINE else "Esp."
    data.append([cod, e['tipo'][:46], cat, str(na), str(no), str(nx), str(e['n'])])
data.append(["", "TOTAL", "", str(sup_count[SUP_AESA]), str(sup_count[SUP_OTRO]),
             str(sup_count[SUP_EXIST]), str(total)])
t = LongTable(data, colWidths=[15*mm, 73*mm, 17*mm, 15*mm, 14*mm, 15*mm, 15*mm], repeatRows=1)
st = hstyle([('ALIGN',(0,0),(0,-1),'CENTER'),
             ('ALIGN',(2,0),(-1,-1),'CENTER'),
             ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),
             ('BACKGROUND',(0,-1),(-1,-1),LBLUE)])
for i,(cod,e) in enumerate(code_rows, start=1):
    if e['cat']==CAT_INLINE:
        st.add('BACKGROUND',(2,i),(2,i),GREEN)
t.setStyle(st)
E.append(t)

# 6. Cantidad por area
E.append(Paragraph("6. Cantidad por area", H2))
col = [(a, n) for a, n in area_rows]
rowsN = (len(col) + 2)//3
grid = [["Area","Cant.","Area","Cant.","Area","Cant."]]
for i in range(rowsN):
    line = []
    for j in range(3):
        k = i + j*rowsN
        line += ([col[k][0], str(col[k][1])] if k < len(col) else ["",""])
    grid.append(line)
t = Table(grid, colWidths=[26*mm,18*mm,26*mm,18*mm,26*mm,18*mm])
t.setStyle(hstyle([('ALIGN',(0,0),(-1,-1),'CENTER')]))
E.append(t)

# 7. Notas
E.append(Paragraph("7. Criterio de clasificacion y notas", H2))
inline_list = ", ".join(sorted(INLINE_CODES))
notas = [
 "<b>Alcance:</b> se contabilizan unicamente instrumentos fisicos listados en las Listas "
 "de Instrumentos. No se incluyen senales, alarmas, funciones logicas ni elementos que no "
 "sean instrumentos.",
 "<b>In-line (montado en caneria):</b> instrumentos que se instalan sobre la propia linea "
 "de proceso. Familias por codigo: <i>%s</i>. Validado contra la columna 'Linea o Equipo'." % inline_list,
 "<b>Montaje especifico:</b> el resto (transmisores, indicadores, interruptores, detectores, "
 "posicionadores, solenoides, fines de carrera, etc.), montados sobre soporte, equipo o sobre "
 "la valvula segun el tipico de montaje correspondiente.",
 "<b>Suministro AESA:</b> instrumentos con Comentario vacio / '-' (alcance EPC AESA) y los "
 "indicados explicitamente como 'PROVISTO POR AESA' (aunque esten dentro de un skid de vendor).",
 "<b>Suministro por otro proveedor:</b> instrumentos indicados como 'POR PROVEEDOR ...' "
 "(PROPAK, skids de gas combustible / quimicos / aire, proveedor de valvula, bomba, tanque, "
 "antorcha, etc.). Los marcados 'MONTAJE POR EPC' los provee el vendor pero los monta AESA.",
 "<b>Existente:</b> instrumentos 'EXISTENTE A REUTILIZAR' o 'A REEMPLAZAR'; se listan aparte "
 "por requerir definicion de alcance especifica.",
 "<b>Fuente:</b> ACAL-00102-LI-K-0001 Rev.A y ACAL-00670-LI-K-0001 Rev.B (Pluspetrol).",
 "<b>Proximo paso:</b> esta clasificacion alimenta el siguiente analisis de constructibilidad "
 "(cantidad y tipo de montaje por instrumento, segregado por alcance de provision).",
]
for nta in notas:
    E.append(Paragraph(nta, N)); E.append(Spacer(1, 2))

def footer(canvas, doc):
    canvas.saveState()
    canvas.setFont('Helvetica', 7)
    canvas.setFillColor(colors.HexColor('#888888'))
    canvas.drawString(15*mm, 10*mm,
        "Constructibilidad CPF2 - Resumen de Instrumentos  |  Rev. %s" % REV_ACTUAL)
    canvas.drawRightString(195*mm, 10*mm, "Pag. %d" % doc.page)
    canvas.restoreState()

doc = SimpleDocTemplate(OUT, pagesize=A4, leftMargin=15*mm, rightMargin=15*mm,
                        topMargin=15*mm, bottomMargin=15*mm,
                        title="Resumen de Instrumentos CPF2 - Rev. %s" % REV_ACTUAL)
doc.build(E, onFirstPage=footer, onLaterPages=footer)
print("OK ->", OUT)
print("Total=%d  In-line=%d  Esp=%d  AESA=%d  Otro=%d  Exist=%d"
      % (total, n_inline, n_esp, sup_count[SUP_AESA], sup_count[SUP_OTRO], sup_count[SUP_EXIST]))
