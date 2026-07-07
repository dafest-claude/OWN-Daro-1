#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rev. Alt 5 sobre la Rev. Alt 4.
Verificación: el MONTAJE de la caja de conexión (JB) en sí (fijación, prensacables/glands,
rotulado, PAT interna) estaba descrito en I.7 pero NO estaba cuantificado como HH; solo se
computaba el montaje del soporte. Se incorpora y se renombra el ítem para que cubra
CAJA + SOPORTE.
 - C.1: nuevo estándar 'Montaje de caja JB' = 4 HH/caja (además de los estándares de soporte).
 - C.3: ítem renombrado 'Montaje de cajas de conexión (JB): caja + soporte' = 320 (caja) + 240
        (soporte) = 560 HH; subtotal IN 18.806 -> 19.126.
 - C.4: frente IN de montaje pasa a 8.155 HH (incluye cajas + soportes JB).
 - TOTAL E&I 60.426 -> 60.746 HH (~345 persona-mes).
 - I.7 y nota de C.3: aclaran que caja + soporte se computan en el Anexo C.
No borra revisiones anteriores.
Base: Análisis generados/Pliego subcontrato E&I - Rev Alt 4 - 2026-07-07 (con Anexos).docx
Salida: Análisis generados/Pliego subcontrato E&I - Rev Alt 5 - 2026-07-07 (con Anexos).docx
"""
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
SRC=os.path.join(GEN,"Pliego subcontrato E&I - Rev Alt 4 - 2026-07-07 (con Anexos).docx")
OUT=os.path.join(GEN,"Pliego subcontrato E&I - Rev Alt 5 - 2026-07-07 (con Anexos).docx")
doc=Document(SRC)

def set_cell_text(cell,new):
    p=cell.paragraphs[0]
    if p.runs:
        p.runs[0].text=new
        for r in list(p.runs)[1:]: r._element.getparent().remove(r._element)
    else:
        p.add_run(new)

def insert_row_before(table,ref_idx,values,size=8.5):
    new=table.add_row()
    for j,v in enumerate(values):
        c=new.cells[j]; c.text=''
        r=c.paragraphs[0].add_run(str(v)); r.font.size=Pt(size); r.font.name='Calibri'
        if j>0 and len(str(v))<22: c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    table.rows[ref_idx]._tr.addprevious(new._tr)
    trPr=new._tr.get_or_add_trPr()
    if trPr.find(qn('w:cantSplit')) is None: trPr.append(OxmlElement('w:cantSplit'))
    return new

# ---------- C.1: nuevo estándar de montaje de caja JB (antes de los soportes) ----------
t17=doc.tables[17]
for i,r in enumerate(t17.rows):
    if 'soporte de JB sobre estructura existente' in r.cells[1].text:
        insert_row_before(t17,i,
            ["IN","Montaje de caja de conexión JB (fijación, prensacables, rotulado, PAT interna)",
             "4","HH/caja","2","8"])
        break

# ---------- C.3: renombrar y recomputar el ítem (caja + soporte) ----------
t19=doc.tables[19]
for r in t19.rows:
    if 'soporte' in r.cells[0].text.lower() and 'JB' in r.cells[0].text or \
       'cajas de conexión (JB)' in r.cells[0].text:
        set_cell_text(r.cells[0],"Montaje de cajas de conexión (JB): caja + soporte")
        set_cell_text(r.cells[1],"~80 cajas + ~80 soportes (48 estr. + 32 ped.)")
        set_cell_text(r.cells[2],"4 HH/caja + 2 y 4,5 HH/soporte")
        set_cell_text(r.cells[3],"560")   # caja 80·4=320 + soporte 240
        break
# subtotal IN 18.806 -> 19.126
for c in t19.rows[-1].cells:
    if '18.806' in c.text: set_cell_text(c,c.text.replace('18.806','19.126'))

# ---------- C.4: frente IN montaje incluye cajas + soportes JB ----------
t20=doc.tables[20]
for r in t20.rows:
    if 'Montaje instrumentos + bandejas' in r.cells[1].text:
        set_cell_text(r.cells[1],"Montaje instrumentos + bandejas + cajas/soportes JB")
        set_cell_text(r.cells[2],"8.155")   # 5.815 + 1.780 + 560
        set_cell_text(r.cells[4],"1.359")   # 8.155 / 6
        set_cell_text(r.cells[5],"~8")
        break

# ---------- TOTAL E&I 60.426 -> 60.746 ; 343 -> 345 persona-mes ----------
for p in doc.paragraphs:
    if 'TOTAL mano de obra estimada' in p.text:
        for r in p.runs:
            r.text=r.text.replace('60.426','60.746').replace('343 persona-mes','345 persona-mes')

# ---------- nota de C.3: incluir la caja ----------
for p in doc.paragraphs:
    if 'son indicativos' in p.text and 'JB' in p.text:
        for r in p.runs:
            r.text=r.text.replace(
                "Los soportes de cajas de conexión (JB) —~80, y las HH asociadas—",
                "El montaje de cajas de conexión (JB) —~80 cajas y sus soportes, y las HH asociadas—")

# ---------- I.7: aclarar que caja + soporte se computan en Anexo C ----------
for p in doc.paragraphs:
    if 'se computan aparte como actividad propia' in p.text:
        for r in p.runs:
            r.text=r.text.replace(
                "se computan aparte como actividad propia.",
                "se computan aparte como actividad propia (Anexo C: montaje de caja + soporte).")

# ---------- tipografía Calibri en runs nuevos ----------
def force_cal(run):
    rpr=run._element.get_or_add_rPr(); rf=rpr.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts'); rpr.insert(0,rf)
    for a in ('w:asciiTheme','w:hAnsiTheme','w:cstheme'):
        if rf.get(qn(a)) is not None: del rf.attrib[qn(a)]
    rf.set(qn('w:ascii'),'Calibri'); rf.set(qn('w:hAnsi'),'Calibri'); rf.set(qn('w:cs'),'Calibri')
for p in doc.paragraphs:
    for r in p.runs: force_cal(r)
for t in doc.tables:
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs: force_cal(r)

doc.save(OUT)
print("OK ->",OUT,"| tablas:",len(doc.tables))
