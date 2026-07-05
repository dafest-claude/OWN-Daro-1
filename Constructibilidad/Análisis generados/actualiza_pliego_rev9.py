#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Rev.9 sobre Rev.8:
- Reduce encabezado/pie y márgenes (más área de contenido); acorta texto de pie.
- Mejora paginado: recorta espaciadores de carátulas, 'keep with next' en títulos, filas de tabla sin dividir.
- Aclara 'LOTO' (Lock-Out/Tag-Out). Cambia 'ferrule' por 'terminales'.
- Matriz de responsabilidades: anchos de columna ajustados (concepto/observaciones anchas, X angostas)
  + agrega 'Jefatura de obra' (AESA) y 'Oficina técnica' (AESA).
- Revisión 8 -> 9, fecha 05-07-2026.
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
SRC=os.path.join(BASE,"Pliego subcontrato E&I - Rev8 - 2026-07-04 (con Anexos).docx")
OUT=os.path.join(BASE,"Pliego subcontrato E&I - Rev9 - 2026-07-05 (con Anexos).docx")
doc=Document(SRC)
TBL='Tabla Cuadro'
try: doc.styles[TBL]
except KeyError: TBL='Table Grid'

# ---------- 1) revisión / fecha ----------
for p in doc.paragraphs:
    if 'Revisión 8' in p.text:
        for r in p.runs:
            r.text=r.text.replace('Revisión 8','Revisión 9').replace('04-07-2026','05-07-2026')

# ---------- 2) márgenes / encabezado / pie ----------
for s in doc.sections:
    s.top_margin=Cm(2.2); s.bottom_margin=Cm(1.9)
    s.header_distance=Cm(0.7); s.footer_distance=Cm(0.7)
    # acortar texto descriptivo del pie
    for p in s.footer.paragraphs:
        for r in p.runs:
            if r.font.size is None or r.font.size>Pt(8): r.font.size=Pt(7)

# ---------- 3) LOTO + ferrule ----------
def repl_runs(find_map):
    def apply(paras):
        for p in paras:
            for r in p.runs:
                for a,b in find_map.items():
                    if a in r.text: r.text=r.text.replace(a,b)
    apply(doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells: apply(c.paragraphs)
repl_runs({
 "Especialistas en LOTO":"Especialistas en LOTO (Lock-Out / Tag-Out — bloqueo y etiquetado de energías peligrosas)",
 "ferrules":"terminales de bornera","Ferrules":"Terminales de bornera",
 "ferrule":"terminal de bornera","Ferrule":"Terminal de bornera",
})

# ---------- 4) mejoras de paginado ----------
# 4a) recortar espaciadores vacíos antes de las carátulas ANEXO
COVERS={'ANEXO E','ANEXO I','ANEXO P','ANEXO C','ANEXO R'}
for p in list(doc.paragraphs):
    if p.style.name=='Heading 1' and p.text.strip() in COVERS:
        prev=p._p.getprevious(); removed=0
        while prev is not None and removed<5 and prev.tag.endswith('}p'):
            txt=''.join(t.text or '' for t in prev.iter(qn('w:t')))
            hasbr=prev.find('.//'+qn('w:br')) is not None
            if txt.strip()=='' and not hasbr:
                nxt=prev.getprevious(); prev.getparent().remove(prev); prev=nxt; removed+=1
            else: break
# 4b) keep with next en títulos (no quedan solos al pie)
for p in doc.paragraphs:
    if p.style.name in ('Heading 1','Heading 2'):
        p.paragraph_format.keep_with_next=True
# 4c) filas de tabla sin dividir entre páginas
def no_split(t):
    for row in t.rows:
        trPr=row._tr.get_or_add_trPr()
        if trPr.find(qn('w:cantSplit')) is None:
            trPr.append(OxmlElement('w:cantSplit'))
for t in doc.tables: no_split(t)
# 4d) quitar párrafos vacíos al final del documento
while doc.paragraphs and doc.paragraphs[-1].text.strip()=='' and doc.paragraphs[-1].find(qn('w:br')) is None if False else False:
    pass
body=doc.element.body
for ch in list(body.iterchildren())[::-1]:
    if ch.tag.endswith('}p'):
        txt=''.join(t.text or '' for t in ch.iter(qn('w:t')))
        hasbr=ch.find('.//'+qn('w:br')) is not None
        if txt.strip()=='' and not hasbr: body.remove(ch)
        else: break
    else: break

# ---------- 5) matriz: reconstruir con anchos + Jefatura/Oficina técnica ----------
MATRIZ=[
 ["1","Gerenciamiento global del proyecto","X","",""],
 ["2","Jefatura de obra","X","",""],
 ["3","Oficina técnica","X","",""],
 ["4","Ingeniería de detalle y planos 'For Construction'","X","",""],
 ["5","Provisión de materiales principales (cables, bandejas, instrumentos, tableros, PAT/SPCDA, iluminación)","X","","Según listas de materiales"],
 ["6","Consumibles y misceláneos de montaje (terminales menores y de bornera, precintos, tornillería)","","X","Salvo indicación en contrario"],
 ["7","Herramientas y equipos de montaje","","X",""],
 ["8","Instrumental de medición y prueba (megger, multímetro, calibrador, comunicador HART, banco)","","X","Calibrado y certificado"],
 ["9","Mano de obra calificada (montaje, tendido, conexionado, precom, comisionado)","","X","CS, seguros, ART s/normativa"],
 ["10","Supervisión técnica de las tareas","","X",""],
 ["11","Movilización y desmovilización de personal y equipos","","X",""],
 ["12","Andamios y plataformas para trabajo en altura","X","","Según disponibilidad de obra"],
 ["13","Izaje / grúas para montaje de equipos pesados (tableros, motores, celdas)","X","",""],
 ["14","Energía, aire comprimido y servicios en obrador y frentes","X","",""],
 ["15","Permisos de trabajo (PTW) y habilitación de tareas","X","","Emisión AESA; cumplimiento Subcontratista"],
 ["16","EPP para todo el personal","","X","S/normativa vigente"],
 ["17","Vehículos del subcontratista y choferes","","X",""],
 ["18","Cumplimiento de Seguridad, Higiene y Medio Ambiente","","X","Supervisión AESA / Cliente"],
 ["19","Plan de calidad, ITP, protocolos y registros","","X","Aprobación AESA"],
 ["20","Inspección y liberación de puntos de espera (hold points)","X","","Ejecución y autocontrol Subcontratista"],
 ["21","Documentación conforme a obra (as-built) de sus tareas","","X",""],
 ["22","Partes diarios, reportes de avance y certificación","","X",""],
 ["23","Gestión administrativa de documentación laboral/legal","","X",""],
 ["24","Alojamiento, comida y traslado del personal (campamento)","X","","Campamento de obra"],
 ["25","Servicio médico en obra y emergencias","X","",""],
 ["26","Gestión de residuos de sus tareas","","X","Dispone donde AESA indique"],
 ["27","Instrumental y personal para precom/comisionado y pruebas","","X","Coordinación con AESA / vendors"],
 ["28","Coordinación con vendors (ABB, Inauco, HIMA) para FAT/SAT/iFAT","X","",""],
 ["29","Seguros del personal y equipos (ART, RC)","","X",""],
]
HDR=["Nº","Concepto","AESA","Subcon.","Observaciones"]
WIDTHS=[Cm(0.9),Cm(7.6),Cm(1.3),Cm(1.5),Cm(4.7)]
old=None
for t in doc.tables:
    if t.rows and t.rows[0].cells[0].text.strip()=='Nº' and 'AESA' in ' '.join(c.text for c in t.rows[0].cells):
        old=t; break
if old is not None:
    nt=doc.add_table(rows=1,cols=5); nt.style=TBL; nt.autofit=False
    nt.alignment=WD_ALIGN_PARAGRAPH.CENTER
    def fill(cell,txt,bold=False,center=False):
        cell.text=''; p=cell.paragraphs[0]
        if center: p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        r=p.add_run(txt); r.bold=bold; r.font.size=Pt(8.5); r.font.name='Calibri'
    for j,h in enumerate(HDR): fill(nt.rows[0].cells[j],h,bold=True,center=(j!=1))
    for row in MATRIZ:
        cs=nt.add_row().cells
        for j,v in enumerate(row): fill(cs[j],v,center=(j!=1))
    # anchos
    for j,w in enumerate(WIDTHS):
        nt.columns[j].width=w
        for row in nt.rows: row.cells[j].width=w
    # filas sin dividir
    no_split(nt)
    # mover a la posición de la vieja y eliminarla
    old._tbl.addprevious(nt._tbl)
    old._tbl.getparent().remove(old._tbl)

# ---------- 6) reforzar tipografía Calibri en runs nuevos ----------
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
print("OK ->",OUT)
