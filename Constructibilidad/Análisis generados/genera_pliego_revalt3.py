#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rev. Alt 3 sobre la Rev. Alt 2 corregida por el usuario (en 'Datos entrada').
Preserva TODAS las ediciones del usuario y agrega el ítem faltante de mano de obra:
INSTALACIÓN DE SOPORTES DE LAS CAJAS DE CONEXIÓN (JB) de instrumentos.
- Suma un estándar en C.1 (modificable).
- Suma el ítem en C.3 (HH Instrumentación) y actualiza el subtotal.
- Lo incorpora en C.4 (dotación / HH-mes) y en el TOTAL E&I.
- Lo describe como actividad en el Anexo I (I.7 — cajas de conexión).
No borra revisiones anteriores.
Base: Datos entrada/Pliego subcontrato E&I - Rev Alt 2 - 2026-07-07 (con Anexos).docx
Salida: Análisis generados/Pliego subcontrato E&I - Rev Alt 3 - 2026-07-07 (con Anexos).docx
"""
import os
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ENT="/home/user/OWN-Daro-1/Constructibilidad/Datos entrada"
GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
SRC=os.path.join(ENT,"Pliego subcontrato E&I - Rev Alt 2 - 2026-07-07 (con Anexos).docx")
OUT=os.path.join(GEN,"Pliego subcontrato E&I - Rev Alt 3 - 2026-07-07 (con Anexos).docx")
doc=Document(SRC)

# ---------- helpers ----------
def set_cell_text(cell,new):
    """Reemplaza el texto de una celda conservando el formato del primer run."""
    p=cell.paragraphs[0]
    if p.runs:
        p.runs[0].text=new
        for r in list(p.runs)[1:]: r._element.getparent().remove(r._element)
    else:
        p.add_run(new)

def insert_row_before(table,ref_idx,values,size=8.5,bold_idx=()):
    """Crea una fila y la mueve delante de la fila ref_idx."""
    new=table.add_row()  # se agrega al final
    for j,v in enumerate(values):
        c=new.cells[j]; c.text=''
        r=c.paragraphs[0].add_run(str(v)); r.font.size=Pt(size); r.font.name='Calibri'
        if j in bold_idx: r.bold=True
        if j>0 and len(str(v))<22: c.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    ref_tr=table.rows[ref_idx]._tr
    ref_tr.addprevious(new._tr)
    trPr=new._tr.get_or_add_trPr()
    if trPr.find(qn('w:cantSplit')) is None: trPr.append(OxmlElement('w:cantSplit'))
    return new

# ---------- 1) C.1 Estándares (Tabla 17): nuevo estándar de soporte de JB ----------
# insertar entre 'Montaje bandeja (tapa)' (idx 18) y 'Loop check' (idx 19)
insert_row_before(doc.tables[17],19,
 ["IN","Montaje de soporte / pedestal de caja de conexión (JB)","3","HH/soporte","2","8"])

# ---------- 2) C.3 HH Instrumentación (Tabla 19): nuevo ítem + subtotal ----------
t19=doc.tables[19]
# subtotal está en la última fila; insertar el ítem antes de ella
sub_idx=len(t19.rows)-1
insert_row_before(t19,sub_idx,
 ["Montaje de soportes de cajas de conexión (JB)","~80 soportes","3 HH/soporte","240"])
# actualizar subtotal 18.566 -> 18.806
for c in t19.rows[-1].cells:
    if 'HH' in c.text:
        set_cell_text(c,c.text.replace('18.566','18.806'))

# ---------- 3) C.4 Dotación (Tabla 20): frente IN montaje incluye soportes JB ----------
t20=doc.tables[20]
for r in t20.rows:
    if 'Montaje instrumentos + bandejas' in r.cells[1].text:
        set_cell_text(r.cells[1],"Montaje instrumentos + bandejas + soportes JB")
        set_cell_text(r.cells[2],"7.835")   # 7.595 + 240
        set_cell_text(r.cells[4],"1.306")   # 7.835 / 6
        set_cell_text(r.cells[5],"~7,5")
        break

# ---------- 4) TOTAL E&I en la nota (60.186 -> 60.426 ; 342 -> 343 persona-mes) ----------
for p in doc.paragraphs:
    if 'TOTAL mano de obra estimada' in p.text:
        for r in p.runs:
            r.text=r.text.replace('60.186','60.426').replace('342 persona-mes','343 persona-mes')

# ---------- 5) Anexo I — I.7: describir la actividad de soportes de JB ----------
anchor=None
for p in doc.paragraphs:
    if p.text.strip().startswith('•') and 'Montaje y fijación de la caja de conexión' in p.text:
        anchor=p; break
if anchor is not None:
    new_p=doc.add_paragraph()      # se agrega al final
    new_p.style=anchor.style
    run=new_p.add_run("•  Provisión de mano de obra e instalación del soporte / pedestal de la "
        "caja (fijación y anclaje a estructura o base de hormigón, nivelación y aplomado). "
        "Se computa ~1 soporte por caja (~80 soportes) como mano de obra en el Anexo C (C.1/C.3).")
    # heredar tamaño del run del anchor si lo tiene
    if anchor.runs and anchor.runs[0].font.size is not None:
        run.font.size=anchor.runs[0].font.size
    run.font.name='Calibri'
    anchor._p.addprevious(new_p._p)

# ---------- 6) tipografía Calibri en runs nuevos ----------
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
