#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rev. Alt 2 sobre Rev. Alt 1. Retoques al anexo de mano de obra:
- Tabla de ESTÁNDARES (rendimientos) usados para el cálculo de HH, modificables, separada EL/IN/Común.
- Estimación de HH separada por ELECTRICIDAD e INSTRUMENTACIÓN.
- Dotación / HH-mes separada por especialidad.
- Cuadros a completar por el Oferente (cotización y disponibilidad) en PESOS (ARS), separados EL/IN/Común.
- Se mantiene el Anexo R (Matriz de Responsabilidades) sin cambios.
Base: Rev. Alt 1. Salida: 'Pliego subcontrato E&I - Rev Alt 2 - 2026-07-07 (con Anexos).docx'
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
SRC=os.path.join(BASE,"Pliego subcontrato E&I - Rev Alt 1 - 2026-07-05 (con Anexos).docx")
OUT=os.path.join(BASE,"Pliego subcontrato E&I - Rev Alt 2 - 2026-07-07 (con Anexos).docx")
doc=Document(SRC)
TBL='Tabla Cuadro'
try: doc.styles[TBL]
except KeyError: TBL='Table Grid'

# revisión / fecha
for p in doc.paragraphs:
    if 'Revisión Alt 1' in p.text:
        for r in p.runs: r.text=r.text.replace('Revisión Alt 1','Revisión Alt 2').replace('05-07-2026','07-07-2026')

def edit_para(start,new):
    for p in doc.paragraphs:
        if p.text.strip().startswith(start):
            for r in list(p.runs)[1:]: r._element.getparent().remove(r._element)
            if p.runs: p.runs[0].text=new
            else: p.add_run(new)
            return p
def remove_paras(pred):
    for p in list(doc.paragraphs):
        if pred(p.text.strip()): p._element.getparent().remove(p._element)
def remove_tables(pred):
    for t in list(doc.tables):
        hdr=' | '.join(c.text for c in t.rows[0].cells)
        if pred(hdr): t._element.getparent().remove(t._element)

# quitar contenido mano de obra de Rev Alt 1 (headings C.1-C.4, notas, 4 tablas)
remove_tables(lambda h: any(k in h for k in ['HH estimadas','HH-mes','Precio hora-hombre','Cantidad disponible']))
remove_paras(lambda t: t[:3] in ('C.1','C.2','C.3','C.4'))
remove_paras(lambda t: t.startswith(('Estimación de referencia derivada','Equivalente ≈ 317','HH-mes = HH','El Oferente completa el precio','El Oferente indica la cantidad')))

# intro nueva
edit_para('El presente pliego tiene por objeto la contratación de personal especializado',
 "El presente pliego tiene por objeto la contratación de personal especializado de Electricidad e "
 "Instrumentación. La cotización se estructura sobre la NECESIDAD DE MANO DE OBRA, separada por especialidad "
 "(Electricidad e Instrumentación; los conceptos comunes se indican como 'Común'). Se incluye la tabla de "
 "estándares de rendimiento usados para el cálculo de horas-hombre (modificable), la estimación de HH por "
 "especialidad, la dotación mensual, y los cuadros a completar por el Oferente (cotización en PESOS y "
 "disponibilidad de personal) para facilitar la comparación entre proveedores.")

# ---- construir nuevos elementos (se insertan tras el intro, antes del Anexo R) ----
new_els=[]
def H(txt):
    p=doc.add_heading(txt,level=2); p.paragraph_format.keep_with_next=True; new_els.append(p._p); return p
def P(txt,note=False):
    p=doc.add_paragraph(); r=p.add_run(txt)
    if note: r.italic=True; r.font.size=Pt(9); r.font.color.rgb=RGBColor(0x55,0x55,0x55)
    new_els.append(p._p); return p
def T(headers,rows,bold_last=False):
    t=doc.add_table(rows=1,cols=len(headers)); t.style=TBL
    for j,h in enumerate(headers):
        c=t.rows[0].cells[j]; c.text=''; r=c.paragraphs[0].add_run(h); r.bold=True; r.font.size=Pt(8.5)
    for ri,row in enumerate(rows):
        cs=t.add_row().cells
        for j,v in enumerate(row):
            cs[j].text=''; rr=cs[j].paragraphs[0].add_run(str(v)); rr.font.size=Pt(8.5)
            if bold_last and ri==len(rows)-1: rr.bold=True
            if j>0 and len(str(v))<18: cs[j].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    for row in t.rows:
        tr=row._tr.get_or_add_trPr()
        if tr.find(qn('w:cantSplit')) is None: tr.append(OxmlElement('w:cantSplit'))
    new_els.append(t._tbl); return t

# C.1 Estándares
H("C.1  Estándares de rendimiento para el cálculo de HH (modificables)")
P("Base del cálculo de horas-hombre (HH). Modificando estos estándares se recalculan las HH de C.2 y C.3. "
  "Separados por especialidad; los parámetros de jornada son comunes a ambas.")
T(["Especialidad","Tarea / concepto","Estándar","Unidad","Cuadrilla (pers.)","Jornada (h/día)"],
 [["EL","Tendido cable Potencia MT","30","m/día·cuadrilla","4","8"],
  ["EL","Tendido cable BT Grande (≥35 mm²)","50","m/día·cuadrilla","4","8"],
  ["EL","Tendido cable BT Mediana","80","m/día·cuadrilla","4","8"],
  ["EL","Tendido cable BT Pequeña","120","m/día·cuadrilla","4","8"],
  ["EL","Tendido cable Control / Señales","100","m/día·cuadrilla","4","8"],
  ["EL","Terminal MT","8","HH/extremo","—","8"],
  ["EL","Terminal BT Grande","3","HH/extremo","—","8"],
  ["EL","Terminal BT Mediana / Pequeña","1,5","HH/extremo","—","8"],
  ["EL","Conexionado cable de control","4","HH/cable","—","8"],
  ["EL","Montaje de motor","0,5–1","día/motor","4","8"],
  ["IN","Tendido cable simple armado","180","m/día·cuadrilla","2","8"],
  ["IN","Tendido cable multipar armado","120","m/día·cuadrilla","2","8"],
  ["IN","Punta de conductor","7","min/punta","—","8"],
  ["IN","Punta de pantalla","5","min/punta","—","8"],
  ["IN","Punta de armadura (prensacable)","10","min/punta","—","8"],
  ["IN","Montaje de instrumento","5","HH/instrumento","2","8"],
  ["IN","Montaje bandeja (tramo / accesorio)","1,2","HH/unidad","2","8"],
  ["IN","Montaje bandeja (tapa)","0,3","HH/unidad","2","8"],
  ["IN","Loop check / prueba de señal","2","HH/lazo","2","8"],
  ["Común","Jornada","8","h/día","—","—"],
  ["Común","Semana laboral","6","días/semana","—","—"],
  ["Común","Persona-mes","176","HH/persona-mes","—","—"]])

# C.2 HH Electricidad
H("C.2  Estimación de mano de obra — ELECTRICIDAD")
T(["Tarea","Cantidad de referencia","Estándar aplicado","HH estimadas"],
 [["Tendido de cables (todos los tipos)","81.231 m","Rendimientos de tendido EL (C.1)","32.829"],
  ["Conexionado y terminaciones","5.178 puntas","8 / 3 / 1,5 HH·ext · 4 HH/cable","2.274"],
  ["Montaje de equipos (motores/celdas; salas vienen montadas)","79 mot. + 6 celdas","C.1","2.840"],
  ["Pruebas, energización y precomisionado","513 circuitos","—","3.677"],
  ["SUBTOTAL ELECTRICIDAD","—","—","≈ 41.620 HH"]],bold_last=True)

# C.3 HH Instrumentación
H("C.3  Estimación de mano de obra — INSTRUMENTACIÓN")
T(["Tarea","Cantidad de referencia","Estándar aplicado","HH estimadas"],
 [["Tendido de cables","65.315 m","180 / 120 m/día·cuadrilla","7.227"],
  ["Conexionado de puntas","15.752 puntas","7 / 5 / 10 min/punta","1.974"],
  ["Montaje de instrumentos (AESA)","1.163 instr.","5 HH/instrumento","5.815"],
  ["Montaje de bandejas / conduits","2.838 U","1,2 y 0,3 HH/unidad","1.780"],
  ["Loop check / prueba de señales","885 lazos","2 HH/lazo","1.770"],
  ["SUBTOTAL INSTRUMENTACIÓN","—","—","≈ 18.566 HH"]],bold_last=True)
P("TOTAL mano de obra estimada E&I ≈ 60.186 HH ≈ 342 persona-mes (1 persona-mes ≈ 176 HH). "
  "Valores de referencia; se recalculan al cambiar los estándares de C.1.",note=True)

# C.4 Dotación / HH-mes por especialidad
H("C.4  Dotación y dedicación mensual (HH-mes) por especialidad")
T(["Especialidad","Frente","HH estim.","Duración (meses)","HH-mes","Dotación media (pers.)"],
 [["EL","Tendido + conexionado","35.103","6","5.851","~33"],
  ["EL","Montaje + pruebas + precomisionado","6.517","4","1.629","~9"],
  ["IN","Tendido + conexionado","9.201","7","1.314","~7,5"],
  ["IN","Montaje instrumentos + bandejas","7.595","6","1.266","~7"],
  ["IN","Loop check / pruebas","1.770","3","590","~3,5"],
  ["Común","PICO SIMULTÁNEO (referencia)","—","—","—","~40-50"]],bold_last=True)
P("HH-mes = HH / duración; dotación media = HH-mes / 176. El pico se da en la etapa de tendido (FEB-MAR).",note=True)

# perfiles
PEL=["Supervisor / Capataz de Electricidad","Oficial electricista — montaje y tendido",
 "Oficial electricista — conexionado y terminaciones","Técnico MT certificado (terminaciones y energización)",
 "Técnico de pruebas eléctricas / precomisionado"]
PIN=["Supervisor / Capataz de Instrumentación","Oficial instrumentista — montaje y tendido",
 "Técnico instrumentista — conexionado y loop check","Técnico de calibración / banco de pruebas"]
PCO=["Ayudantes","Ingeniería / técnico de precom y comisionado","QA/QC — control de calidad E&I","Higiene y Seguridad"]

# C.5 Cotización (PESOS) a completar
H("C.5  Cuadro de cotización por perfil (en PESOS) — a completar por el Oferente")
rows=[["EL",p,"",""] for p in PEL]+[["IN",p,"",""] for p in PIN]+[["Común",p,"",""] for p in PCO]
T(["Especialidad","Perfil","Precio hora-hombre (ARS)","Precio mes-hombre (ARS)"],rows)
P("El Oferente completa el precio por hora-hombre y/o mes-hombre por perfil, en PESOS (ARS). Los perfiles "
  "'Común' aplican a ambas especialidades. Permite comparar ofertas de forma homogénea.",note=True)

# C.6 Disponibilidad a completar
H("C.6  Disponibilidad de personal por perfil — a completar por el Oferente")
rows=[["EL",p,"","",""] for p in PEL]+[["IN",p,"","",""] for p in PIN]+[["Común",p,"","",""] for p in PCO]
T(["Especialidad","Perfil","Cantidad disponible","Fecha de disponibilidad","Certificaciones / obs."],rows)
P("El Oferente indica la cantidad de personal por perfil que puede ofrecer y su fecha de disponibilidad, "
  "para evaluar la capacidad de movilización frente al pico de MAR-2027.",note=True)

# ---- reubicar los nuevos elementos tras el intro (antes del Anexo R) ----
intro=None
for p in doc.paragraphs:
    if p.text.strip().startswith('El presente pliego tiene por objeto la contratación de personal especializado'):
        intro=p; break
anchor=intro._p
for el in new_els:
    anchor.addnext(el); anchor=el

# ---- ANEXOS box ----
def set_cell(cell,lines,size=9):
    cell.text=''; p0=cell.paragraphs[0]
    for i,ln in enumerate(lines):
        p=p0 if i==0 else cell.add_paragraph(); p.add_run(ln).font.size=Pt(size)
for t in doc.tables:
    labs=[t.rows[r].cells[0].text.strip().upper() for r in range(len(t.rows))]
    if 'ANEXOS' in labs and 'OBJETIVO' in labs:
        for r in range(len(t.rows)):
            if t.rows[r].cells[0].text.strip().upper()=='ANEXOS':
                set_cell(t.rows[r].cells[1],[
                 "Anexo E — Electricidad (cables, conexionado, tableros, PAT, SPCDA, iluminación).",
                 "Anexo I — Instrumentación (cables, conexionado, instrumentos, tableros, cajas, bandejas/conduits).",
                 "Anexo P — Precomisionado, Comisionado y Puesta en Marcha.",
                 "Anexo C — Necesidad de mano de obra y cotización (estándares, HH por especialidad, dotación y cuadros a completar por el Oferente, en pesos).",
                 "Anexo R — Matriz de Responsabilidades AESA / Subcontratista.",
                ])
        break

# ---- tipografía ----
def fc(run):
    rpr=run._element.get_or_add_rPr(); rf=rpr.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts'); rpr.insert(0,rf)
    for a in ('w:asciiTheme','w:hAnsiTheme','w:cstheme'):
        if rf.get(qn(a)) is not None: del rf.attrib[qn(a)]
    rf.set(qn('w:ascii'),'Calibri'); rf.set(qn('w:hAnsi'),'Calibri'); rf.set(qn('w:cs'),'Calibri')
for p in doc.paragraphs:
    for r in p.runs: fc(r)
for t in doc.tables:
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs: fc(r)

doc.save(OUT); print("OK ->",OUT,"| tablas:",len(doc.tables))
