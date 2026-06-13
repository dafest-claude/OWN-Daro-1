#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Analisis de Montaje de Instrumentos - Constructibilidad La Calera CPF2 Fase 2
Determina el TIPICO DE MONTAJE de cada instrumento y lo segrega por alcance de
provision (AESA / otro proveedor), como insumo de constructibilidad (soporteria,
materiales de montaje y horas de instalacion).

El tipico de montaje se deriva de:
  - la categoria de montaje (in-line vs especifico), y
  - la columna 'Linea o Equipo' de la LI, que indica a que se conecta el
    instrumento (linea de caneria, equipo, valvula u otro instrumento).

Fuente: Datos entrada/ACAL-00102-LI-K-0001-A VCO.pdf  (Rev. A)
        Datos entrada/ACAL-00670-LI-K-0001-B VCO.pdf  (Rev. B)
"""
import fitz; fitz.TOOLS.mupdf_display_errors(False)
import re, os
from collections import Counter, defaultdict, OrderedDict
from datetime import date

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
IN   = os.path.join(ROOT, "Datos entrada")
OUT  = os.path.join(BASE, "Analisis_Montaje_Instrumentos_CPF2.pdf")

REV_ACTUAL = "0"
REV_HIST = [
    ("0", "13-06-2026", "Emision inicial. Tipico de montaje por instrumento, "
                        "segregado por alcance de provision (AESA / otro proveedor)."),
]

DOCS = OrderedDict([
    ("ACAL-00102 (Proceso / Facilities)", "ACAL-00102-LI-K-0001-A VCO.pdf"),
    ("ACAL-00670 (Generacion)",           "ACAL-00670-LI-K-0001-B VCO.pdf"),
])

# --- Codigos ---
INLINE_CODES = {'FV','FCV','LCV','PV','PCV','TCV','TV','LV','DPV','PDV','PRV','BPRV',
    'PSV','TSV','PVSV','ESDV','SDV','BDV','XV','DV','RO','FE','FI','PD','TW','TM','CC'}
VALVE_CODES = {'FV','FCV','LCV','PV','PCV','TCV','TV','LV','DPV','PDV','PRV','BPRV',
    'PSV','TSV','PVSV','ESDV','SDV','BDV','XV','DV','HCV','HV'}
INLINE_VALVE = INLINE_CODES & VALVE_CODES
INLINE_ELEM  = INLINE_CODES - VALVE_CODES

# --- Tipicos de montaje (orden y descripcion) ---
M_VAL  = "En linea - valvula"
M_ELEM = "En linea - elemento"
M_STAND= "Soporte junto a linea (stand)"
M_EQUIP= "Sobre equipo (recipiente/maquina)"
M_VLVA = "Sobre valvula (accesorio)"
M_INST = "Sobre/junto a otro instrumento"
M_FG   = "En estructura / campo (F&G)"
MONT_ORDER = [M_VAL, M_ELEM, M_STAND, M_EQUIP, M_VLVA, M_INST, M_FG]
MONT_DESC = {
 M_VAL : "Valvulas bridadas/soldadas en la caneria (control, seguridad, bloqueo, "
         "regulacion). Forman parte del spool de caneria.",
 M_ELEM: "Elementos insertados en la linea: placa orificio, orificio de restriccion, "
         "termovaina, toma de muestra, cupon de corrosion, rotametro, amortiguador.",
 M_STAND:"Instrumentos con toma a proceso montados sobre soporte de pie (tipico stand "
         "2\") junto a la caneria: transmisores y manometros de presion / caudal, etc.",
 M_EQUIP:"Instrumentos montados sobre recipientes, tanques, bombas, compresores o "
         "intercambiadores (nivel, presion/temperatura sobre equipo, vibracion).",
 M_VLVA: "Accesorios montados sobre la valvula / actuador: posicionadores, valvulas "
         "solenoides, finales de carrera, actuadores.",
 M_INST: "Transmisores remotos asociados a un elemento primario (caudal sobre placa "
         "orificio, transmisor de analisis sobre sonda).",
 M_FG  : "Dispositivos de fuego y gas y afines, montados en estructura / soporte "
         "dedicado: detectores de gas y llama, sirenas, balizas, pulsadores.",
}
# in-line vs especifico (para subtotales)
MONT_GROUP = {M_VAL:'In-line', M_ELEM:'In-line',
              M_STAND:'Especifico', M_EQUIP:'Especifico', M_VLVA:'Especifico',
              M_INST:'Especifico', M_FG:'Especifico'}

# --- Suministro ---
SUP_AESA, SUP_OTRO, SUP_EXIST = "AESA", "Otro proveedor", "Existente"
def supply(com):
    c=(com or '').upper().strip()
    if c in ('','-'): return SUP_AESA
    if 'PROVISTO POR AESA' in c: return SUP_AESA
    if c.startswith('EXISTENTE'): return SUP_EXIST
    if 'PROVEEDOR' in c or 'PORVEEDOR' in c: return SUP_OTRO
    return SUP_AESA

def hdrmap(cells):
    m={}
    for i,c in enumerate(cells):
        u=c.upper()
        if 'INST. COD' in u: m['cod']=i
        elif u=='TAG': m['tag']=i
        elif 'TIPO DE INSTR' in u: m['tipo']=i
        elif u=='AREA': m['area']=i
        elif 'LINEA O EQUIPO' in u or 'LÍNEA O EQUIPO' in u: m['linea']=i
        elif 'COMENTARIOS' in u: m['com']=i
    return m

def extract(fn):
    d=fitz.open(fn); seen=set(); out=[]; H=None
    for pno in range(d.page_count):
        for t in d[pno].find_tables().tables:
            for r in t.extract():
                c=[(x or '').replace('\n',' ').strip() for x in r]
                if 'TAG' in c and 'INST. COD.' in c: H=hdrmap(c); continue
                if not H or 'linea' not in H: continue
                if len(c)<=max(H.values()): continue
                tag=c[H['tag']]
                if not re.match(r'^[A-Z]{1,4}-?\d{3,6}[A-Z]?$', tag): continue
                if tag in seen: continue
                seen.add(tag)
                out.append({'cod':c[H['cod']].upper(), 'tag':tag,
                            'tipo':c[H['tipo']].strip(),
                            'area':(c[H['area']].split() or [''])[0].strip(),
                            'linea':c[H['linea']].strip(), 'com':c[H['com']].strip()})
    return out

rows=[]
for label,fn in DOCS.items(): rows+=extract(os.path.join(IN,fn))
ALL_CODES={r['cod'] for r in rows}

def reftype(linea):
    s=(linea or '').strip()
    if s in ('','-'): return 'NONE'
    first=s.split('/')[0].strip()
    if '"' in first: return 'LINE'
    m=re.match(r'^([A-Z]{1,4})-?\d', first)
    if not m: return 'EQUIP'
    pfx=m.group(1)
    if pfx in VALVE_CODES: return 'VALVE'
    if pfx in ALL_CODES:   return 'INSTR'
    return 'EQUIP'

REFMAP={'LINE':M_STAND,'EQUIP':M_EQUIP,'VALVE':M_VLVA,'INSTR':M_INST,'NONE':M_FG}
def montaje(r):
    if r['cod'] in INLINE_VALVE: return M_VAL
    if r['cod'] in INLINE_ELEM:  return M_ELEM
    return REFMAP[reftype(r['linea'])]

for r in rows:
    r['mont']=montaje(r); r['sup']=supply(r['com'])

total=len(rows)
mont_count=Counter(r['mont'] for r in rows)
sup_count=Counter(r['sup'] for r in rows)
# matriz montaje x suministro
mat=defaultdict(int)
for r in rows: mat[(r['mont'],r['sup'])]+=1
# mapeo montaje -> codigos
mont_codes=defaultdict(Counter)
for r in rows: mont_codes[r['mont']][r['cod']]+=1
# matriz area x montaje
areas=sorted({r['area'] for r in rows if r['area']})
amat=defaultdict(int)
for r in rows:
    if r['area']: amat[(r['area'],r['mont'])]+=1

# ---------------- PDF ----------------
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                TableStyle, LongTable)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

styles=getSampleStyleSheet()
H1=ParagraphStyle('H1',parent=styles['Title'],fontSize=16,spaceAfter=4)
SUB=ParagraphStyle('SUB',parent=styles['Normal'],fontSize=10,
                   textColor=colors.HexColor('#555555'),alignment=TA_CENTER)
H2=ParagraphStyle('H2',parent=styles['Heading2'],fontSize=12,
                  textColor=colors.HexColor('#1f3b63'),spaceBefore=12,spaceAfter=6)
N=ParagraphStyle('N',parent=styles['Normal'],fontSize=9,leading=12)
CELL=ParagraphStyle('CELL',parent=styles['Normal'],fontSize=7.5,leading=9)
esc=lambda s: s.replace('&','&amp;')

BLUE=colors.HexColor('#1f3b63'); LBLUE=colors.HexColor('#dce6f2')
GREEN=colors.HexColor('#e2efda'); AESAC=colors.HexColor('#fce4d6')
OTROC=colors.HexColor('#ddebf7'); GREY=colors.HexColor('#f2f2f2')

def hstyle(extra=None):
    base=[('BACKGROUND',(0,0),(-1,0),BLUE),('TEXTCOLOR',(0,0),(-1,0),colors.white),
          ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),8),
          ('GRID',(0,0),(-1,-1),0.4,colors.HexColor('#bbbbbb')),
          ('VALIGN',(0,0),(-1,-1),'MIDDLE'),
          ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,GREY]),
          ('TOPPADDING',(0,0),(-1,-1),3),('BOTTOMPADDING',(0,0),(-1,-1),3)]
    return TableStyle(base+(extra or []))

E=[]
E.append(Paragraph("Analisis de Montaje de Instrumentos", H1))
E.append(Paragraph("Constructibilidad - Proyecto LA CALERA CPF2 - FASE 2", SUB))
E.append(Paragraph("Revision %s &nbsp;|&nbsp; Fecha: %s"
                   % (REV_ACTUAL, date.today().strftime('%d-%m-%Y')), SUB))
E.append(Spacer(1,5*mm))

REVP=ParagraphStyle('REVP',parent=styles['Normal'],fontSize=7.5,leading=9)
data=[["Rev.","Fecha","Descripcion"]]
for rv,fch,desc in REV_HIST: data.append([rv,fch,Paragraph(desc,REVP)])
t=Table(data,colWidths=[12*mm,24*mm,119*mm])
t.setStyle(hstyle([('ALIGN',(0,0),(1,-1),'CENTER')])); E.append(t)

# 1. Objeto
E.append(Paragraph("1. Objeto y alcance", H2))
E.append(Paragraph("Determinar el <b>tipico de montaje</b> de cada instrumento del proyecto "
 "y segregarlo por <b>alcance de provision</b> (AESA / otro proveedor), como insumo para el "
 "dimensionamiento de soporteria, materiales de montaje y horas de instalacion. Se parte del "
 "consolidado de %d instrumentos de las Listas de Instrumentos (areas 00102 y 00670). Se "
 "excluyen senales, alarmas y todo lo que no sea un instrumento fisico." % total, N))

# 2. Resumen ejecutivo
E.append(Paragraph("2. Resumen ejecutivo por tipico de montaje", H2))
data=[["Tipico de montaje","Grupo","Cantidad","%"]]
for m in MONT_ORDER:
    data.append([m, MONT_GROUP[m], str(mont_count[m]),
                 "%.1f%%"%(100*mont_count[m]/total)])
data.append(["TOTAL","",str(total),"100%"])
t=Table(data,colWidths=[78*mm,24*mm,28*mm,22*mm])
st=hstyle([('ALIGN',(2,0),(-1,-1),'CENTER'),('ALIGN',(1,0),(1,-1),'CENTER'),
           ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('BACKGROUND',(0,-1),(-1,-1),LBLUE)])
for i,m in enumerate(MONT_ORDER,start=1):
    if MONT_GROUP[m]=='In-line': st.add('BACKGROUND',(1,i),(1,i),GREEN)
t.setStyle(st); E.append(t)

# 3. Definicion
E.append(Paragraph("3. Definicion de los tipicos de montaje", H2))
data=[["Tipico de montaje","Descripcion"]]
for m in MONT_ORDER: data.append([Paragraph("<b>%s</b>"%esc(m),CELL), Paragraph(esc(MONT_DESC[m]),CELL)])
t=Table(data,colWidths=[45*mm,110*mm])
t.setStyle(hstyle([('VALIGN',(0,0),(-1,-1),'TOP')])); E.append(t)

# 4. Montaje x suministro
E.append(Paragraph("4. Tipico de montaje x alcance de provision", H2))
data=[["Tipico de montaje","AESA","Otro prov.","Existente","Total"]]
for m in MONT_ORDER:
    data.append([m, str(mat[(m,SUP_AESA)]), str(mat[(m,SUP_OTRO)]),
                 str(mat[(m,SUP_EXIST)]), str(mont_count[m])])
data.append(["TOTAL", str(sup_count[SUP_AESA]), str(sup_count[SUP_OTRO]),
             str(sup_count[SUP_EXIST]), str(total)])
t=Table(data,colWidths=[72*mm,22*mm,24*mm,22*mm,20*mm])
t.setStyle(hstyle([('ALIGN',(1,0),(-1,-1),'CENTER'),
    ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('FONTNAME',(-1,0),(-1,-1),'Helvetica-Bold'),
    ('BACKGROUND',(0,-1),(-1,-1),LBLUE),('BACKGROUND',(-1,1),(-1,-1),LBLUE),
    ('BACKGROUND',(1,0),(1,0),BLUE),('BACKGROUND',(2,0),(2,0),BLUE)]))
E.append(t)
E.append(Spacer(1,2))
E.append(Paragraph("Nota: de los instrumentos provistos por otro proveedor, los marcados "
 "'MONTAJE POR EPC' en la LI son montados por AESA aunque el suministro sea del vendor.", CELL))

# 5. Mapeo tipo de instrumento -> montaje
E.append(Paragraph("5. Codigos de instrumento por tipico de montaje", H2))
data=[["Tipico de montaje","Codigos de instrumento (cantidad)"]]
for m in MONT_ORDER:
    items=", ".join("%s(%d)"%(c,n) for c,n in mont_codes[m].most_common())
    data.append([Paragraph("<b>%s</b>"%esc(m),CELL), Paragraph(esc(items),CELL)])
t=Table(data,colWidths=[45*mm,110*mm])
t.setStyle(hstyle([('VALIGN',(0,0),(-1,-1),'TOP')])); E.append(t)

# 6. Area x montaje
E.append(Paragraph("6. Tipico de montaje por area", H2))
abbr=["Vlv.lin","Elem.lin","Stand","Equipo","s/Vlv","s/Instr","F&G","Total"]
data=[["Area"]+abbr]
for a in areas:
    fila=[a]+[str(amat[(a,m)]) for m in MONT_ORDER]
    fila.append(str(sum(amat[(a,m)] for m in MONT_ORDER)))
    data.append(fila)
data.append(["TOTAL"]+[str(mont_count[m]) for m in MONT_ORDER]+[str(sum(mont_count.values()))])
cw=[18*mm]+[17*mm]*7+[15*mm]
t=LongTable(data,colWidths=cw,repeatRows=1)
t.setStyle(hstyle([('ALIGN',(1,0),(-1,-1),'CENTER'),('FONTSIZE',(0,0),(-1,-1),7),
    ('FONTNAME',(0,-1),(-1,-1),'Helvetica-Bold'),('BACKGROUND',(0,-1),(-1,-1),LBLUE)]))
E.append(t)
E.append(Spacer(1,2))
E.append(Paragraph("Referencias: Vlv.lin = valvula en linea; Elem.lin = elemento en linea; "
 "Stand = soporte junto a linea; Equipo = sobre equipo; s/Vlv = sobre valvula; "
 "s/Instr = sobre otro instrumento; F&G = estructura / campo.", CELL))

# 7. Metodologia
E.append(Paragraph("7. Metodologia y notas", H2))
notas=[
 "<b>Derivacion del tipico de montaje:</b> los instrumentos in-line se separan en valvulas "
 "y elementos por su codigo. El resto se clasifica segun la columna 'Linea o Equipo' de la "
 "LI: si referencia una linea de caneria -> soporte junto a linea; un equipo -> sobre equipo; "
 "una valvula -> accesorio sobre valvula; otro instrumento -> transmisor remoto; sin "
 "referencia -> estructura / campo (F&G).",
 "<b>Alcance de provision:</b> AESA (comentario vacio/'-' o 'PROVISTO POR AESA') vs otro "
 "proveedor ('POR PROVEEDOR ...'); 'Existente' para reutilizar/reemplazar.",
 "<b>Uso:</b> insumo para soporteria (stands, abrazaderas, soportes a estructura), materiales "
 "de montaje y estimacion de horas de instalacion por tipico.",
 "<b>Fuente:</b> ACAL-00102-LI-K-0001 Rev.A y ACAL-00670-LI-K-0001 Rev.B (Pluspetrol).",
]
for nt in notas: E.append(Paragraph(nt,N)); E.append(Spacer(1,2))

def footer(canvas,doc):
    canvas.saveState(); canvas.setFont('Helvetica',7)
    canvas.setFillColor(colors.HexColor('#888888'))
    canvas.drawString(15*mm,10*mm,"Constructibilidad CPF2 - Analisis de Montaje  |  Rev. %s"%REV_ACTUAL)
    canvas.drawRightString(195*mm,10*mm,"Pag. %d"%doc.page)
    canvas.restoreState()

doc=SimpleDocTemplate(OUT,pagesize=A4,leftMargin=15*mm,rightMargin=15*mm,
                      topMargin=15*mm,bottomMargin=15*mm,
                      title="Analisis de Montaje de Instrumentos CPF2 - Rev. %s"%REV_ACTUAL)
doc.build(E,onFirstPage=footer,onLaterPages=footer)
print("OK ->",OUT)
for m in MONT_ORDER: print("  %-38s %4d"%(m,mont_count[m]))
print("  TOTAL %d | AESA %d Otro %d Exist %d"%(total,sup_count[SUP_AESA],sup_count[SUP_OTRO],sup_count[SUP_EXIST]))
