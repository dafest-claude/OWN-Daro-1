#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rev.6 sobre la Rev.5 MODIFICADA por el usuario (preserva sus ediciones de formato/índice/textos).
- Inserta en Anexo I un nuevo 'I.8 Montaje de bandejas y caños/conduits de instrumentación'
  (LM-K-0002) y renumera el 'Personal...' de I.8 -> I.9.
- Agrega la línea de montaje de bandejas al cómputo (C.2 Instrumentación).
- Actualiza la revisión (5 -> 6).
Base: Datos entrada/Pliego subcontrato E&I - Rev5 - 2026-07-02 (con Anexos).docx  (modificada por el usuario)
Salida: Análisis generados/Pliego subcontrato E&I - Rev6 - 2026-07-02 (con Anexos).docx
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

BASE="/home/user/OWN-Daro-1/Constructibilidad"
SRC=os.path.join(BASE,"Datos entrada","Pliego subcontrato E&I - Rev5 - 2026-07-02 (con Anexos).docx")
OUT=os.path.join(BASE,"Análisis generados","Pliego subcontrato E&I - Rev6 - 2026-07-02 (con Anexos).docx")
FECHA="02-07-2026"
doc=Document(SRC)
TBL_STYLE='Tabla Cuadro'
try: doc.styles[TBL_STYLE]
except KeyError: TBL_STYLE='Table Grid'

# ---- 1) actualizar revisión 5 -> 6 ----
for p in doc.paragraphs:
    if 'Revisión 5' in p.text or 'Rev. 5' in p.text or 'Rev.5' in p.text:
        for r in p.runs:
            r.text=r.text.replace('Revisión 5','Revisión 6').replace('Rev. 5','Rev. 6').replace('Rev.5','Rev.6')

# ---- 2) localizar el heading 'I.8  Personal ...' y renumerar a I.9 ----
ref=None
for p in doc.paragraphs:
    if p.text.strip().startswith('I.8') and 'Personal' in p.text:
        ref=p
        for r in p.runs:
            r.text=r.text.replace('I.8','I.9',1)
        break
assert ref is not None, "No se encontró el heading I.8 Personal"

# ---- 3) construir el nuevo bloque I.8 (bandejas) ----
def new_heading(txt):
    return doc.add_heading(txt,level=2)
def new_para(txt,note=False):
    p=doc.add_paragraph(); r=p.add_run(txt)
    if note: r.italic=True; r.font.size=Pt(9); r.font.color.rgb=RGBColor(0x55,0x55,0x55)
    return p
def new_table(headers,rows):
    t=doc.add_table(rows=1,cols=len(headers)); t.style=TBL_STYLE
    for j,htxt in enumerate(headers):
        cell=t.rows[0].cells[j]; cell.text=''
        rr=cell.paragraphs[0].add_run(htxt); rr.bold=True; rr.font.size=Pt(8.5)
    for row in rows:
        cells=t.add_row().cells
        for j,v in enumerate(row):
            cells[j].text=''; rr=cells[j].paragraphs[0].add_run(str(v)); rr.font.size=Pt(8.5)
            if j>0: cells[j].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    return t

nh=new_heading("I.8  Montaje de bandejas y caños/conduits de instrumentación")
intro=new_para("Actividades para el montaje de las bandejas portacables y caños/conduits destinados al cableado "
    "eléctrico de los instrumentos, según el listado de materiales ACAL-102-LM-K-0002 (Materiales para montaje "
    "eléctrico de instrumentos). Comprende el montaje de la soportería, la fijación de los tramos y accesorios "
    "(curvas horizontales y verticales, tees y reducciones) y la colocación de las tapas, como canalización previa "
    "al tendido de los cables de instrumentación.")
tbl=new_table(["Material principal (según LM-K-0002)","Unidad","Cantidad"],
 [["Bandeja portacables recta tipo escalera (SAE 1010, galvanizada)","U","797"],
  ["Curvas horizontales","U","85"],
  ["Curvas verticales (interior + exterior)","U","134"],
  ["Tees y reducciones","U","≈ 20"],
  ["Tapas de bandeja (tramos y accesorios)","U","1.806"],
  ["Subtotal tramos + accesorios de bandeja","U","1.032"]])
note=new_para("Material: bandeja tipo escalera SAE 1010 galvanizada. Incluye la soportería y los accesorios de "
    "fijación según LM-K-0002. Los caños/conduits son de alcance menor en este listado y se computan del listado "
    "completo. Los prensacables de las cajas de conexión se consideran en el ítem I.7.",note=True)

# mover los nuevos elementos ANTES del heading personal (ahora I.9), en orden
for el in (nh,intro,tbl,note):
    ref._p.addprevious(el._p if hasattr(el,'_p') else el._tbl)

# ---- 4) agregar línea al cómputo C.2 (Instrumentación) ----
in_comp=None
for t in doc.tables:
    txt=' '.join(c.text for c in t.rows[0].cells)+' '+' '.join(c.text for r in t.rows[1:] for c in r.cells)
    if 'Ítem de obra' in (t.rows[0].cells[0].text) and ('Loop check' in txt or 'instrumentación' in txt.lower()):
        in_comp=t
for t in doc.tables:
    if t.rows[0].cells[0].text.strip()=='Ítem de obra':
        # C.2 es la que contiene 'Loop check' o 'puntas (instrumentación)'
        body=' '.join(c.text for r in t.rows for c in r.cells)
        if 'Loop check' in body or 'puntas (instrumentación)' in body.lower() or 'Montaje de instrumentos en línea' in body:
            in_comp=t
if in_comp is not None:
    cells=in_comp.add_row().cells
    for j,v in enumerate(["Montaje de bandejas portacables de instrumentación (tramos+accesorios+tapas)","U","2.838","",""]):
        cells[j].text=''; rr=cells[j].paragraphs[0].add_run(str(v)); rr.font.size=Pt(8.5)
        if j>0: cells[j].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER

doc.save(OUT)
print("OK ->",OUT)
print("computo IN actualizado:", in_comp is not None)
