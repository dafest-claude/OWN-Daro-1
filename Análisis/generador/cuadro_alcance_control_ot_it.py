#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cuadro de Alcance y Responsabilidades — Sistemas de Control e IT/OT
EPC CPF2 La Calera (actores: AESA y PP=Pluspetrol)

Columnas: #, Sistema, Alcance, Breve descripción, Procura, Instalación/Montaje,
Observación/riesgo, Documentos de referencia, Páginas.
(P = principal "ANEXO A - EPC CPF 2"; Ad = "Adenda Dic25").

Uso:  python cuadro_alcance_control_ot_it.py [carpeta_salida]
"""
import os, sys
from datetime import datetime
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle

AZUL="#1F3864"; AZUL2="#2E5496"; GRIS="#F2F2F2"
VERDE="#C6EFCE"; AMAR="#FFEB9C"; ROJO="#FFC7CE"

META={
 "titulo":"Cuadro de Alcance y Responsabilidades — Sistemas de Control (PCS/SCADA/PSS/ESD/F&G) e IT/OT",
 "subtitulo":"EPC CPF2 — La Calera · Actores: AESA (Contratista) y PP (Pluspetrol)",
 "proyecto":"Legal PP Project LC",
 "fuentes":"P = ANEXO A - EPC CPF 2 (682 págs) · Ad = CPF 2 - AESA - Adenda Dic25 (221 págs)",
 "fecha":"2026-06-06","analista":"Daro",
}

COLS=["#","Sistema","Alcance (ítem)","Breve descripción","Procura","Instalación / Montaje",
      "Observación / riesgo","Documentos de referencia","Páginas"]

# [#, sistema, ítem, descripción, procura, instalación, observación/riesgo, doc.ref, páginas]
ROWS=[
 # ---------- PCS ----------
 ["1","PCS","Gabinetes de sistema PCS (CPF2-PCS-SC-07A/B, -08)","Tableros del sistema de control de proceso para las nuevas salas remotas INS-007/008. Rockwell/Allen Bradley (ControlLogix, Stratix).","AESA","AESA","Respetar marca/modelo del Pliego; reserva mín. 20% capacidad / 25% ocupación.","ACAL-100-PL-K-105 (Tablero PCS 003); ACAL-100-PL-K-100 (Arq. PCS)","P-144, P-408"],
 ["2","PCS","Gabinetes de marshalling PCS (CPF2-PCS-MC-07A/B, -08A)","Marshalling para conexionado de campo; cables FTA asociados.","AESA","AESA","Cantidad según densidad de señales; mantener diseño/arquitectura adoptada.","ACAL-100-PL-K-100; Listado de Señales","P-144/145, P-408"],
 ["3","PCS","Tableros de comunicaciones (equipos activos) en salas remotas","Switches/equipos activos de la red de control en cada sala remota nueva.","AESA","AESA","Mantener arquitectura COM existente.","ACAL-100-PL-K-109 (Layout COM-003); 226-ACAL-100-PL-K-305 (diag. bloques)","P-144"],
 ["4","PCS","Programación de lógicas de control PCS","Nuevas lógicas (tableros nuevos y señales a remotas existentes).","AESA","AESA","Subcontrato INAUCO obligatorio (lock-in de integrador).","GRAL-EI-K-002/003 (ET PCS/SSS); GRAL-ET-K-101","P-144/145, P-303"],
 ["5","PCS","Instrumentos de campo off-skid","Provisión y montaje/contraste de transmisores, válvulas de control, caudalímetros, etc.","AESA","AESA","Suministro reembolsable (Apéndice I); contraste en campo.","Ad 01.C.07.03; Hojas de Datos INS","Ad-67"],
 ["6","PCS","Instrumentos sobre equipos/skids provistos por PP","Vienen montados y cableados a JB sobre su propio skid (excepto P&ID de proveedores, p.ej. Propak).","PP","PP","Riesgo de interfaz skid↔campo; verificar P&ID de proveedores.","P&ID de proveedores (Propak); Apéndice A","P-408"],
 # ---------- SCADA proceso ----------
 ["7","SCADA proceso","Actualización e integración del SCADA de CPF","Nuevas pantallas, adecuación de existentes, base OPC, vínculos y documentación. Wonderware/OMI (AVEVA).","AESA","AESA","Mantener Wonderware/OMI + integrador INAUCO; actualizar base OPC.","ACAL-000-MO-K-100/101; AARC-100-TI-K-002","P-145, P-155"],
 ["8","SCADA proceso","Integración a nivel lógico / configuración al sistema existente","La solución de AESA es completa SALVO la integración a nivel lógico/configuración al sistema en operación.","PP","PP","🚩 Frontera difusa: definir quién garantiza la integración punta a punta.","Aclaración Anexo 3.1 §3.7 (consulta 42)","P-297"],
 ["9","SCADA proceso","Pruebas de integración (SIT) PCS↔SCADA anticipadas en taller","Para el PCS, dentro del subcontrato de AESA.","AESA","AESA","Anticipar SIT a taller (IEC-62381 lo ubica post-SAT) para reducir riesgo en sitio.","IEC 62381; ACAL-000-MO-K-101","P-303"],
 ["10","SCADA proceso","Ampliación de SCADA CPF1/CPF2","Incluida en montaje/programación/integración (sin precio separado).","AESA","AESA","🚩 Sin precio separado (incluida) → riesgo de subvaloración / disputa de alcance.","Ad 01.C.07.04.03","Ad-68, P-282"],
 # ---------- SCADA eléctrico ----------
 ["11","SCADA eléctrico (PMS)","Ampliación del SCADA Eléctrico / PMS (ABB)","Ampliación de HW (tableros ABB), nuevas pantallas y modificación de existentes; desconexión de cargas.","AESA","AESA","Mantener ABB; acordar nuevas cargas disponibles para desconexión.","021-ACAL-100-PS030-E-302 (Arq.); -E-303 (Topología)","P-144, P-281, Ad-67"],
 ["12","SCADA eléctrico (PMS)","Red redundante IEC 61850","Criterio de red redundante para SCADA eléctrico confirmado.","AESA","AESA","Confirmado por consulta de oferente.","IEC 61850; 021-ACAL-100-PS030-E-303","P-304"],
 # ---------- PSS ----------
 ["13","PSS","Gabinetes de seguridad PSS — hardware (CPF-PSS-SC-07/08)","Hardware del sistema de seguridad de proceso (HIMA/HIMAX). PP provee; AESA monta y cablea.","PP","AESA","🚩 En P-408 rotulado con código del PCS (inconsistencia a reconciliar). FAT si cambia matriz C&E.","ACAL-100-PL-K-562/564; ACAL-100-PL-K-551","P-144, P-408"],
 ["14","PSS","Ampliación de remota PSS","Módulos de I/O (según listado + reserva), connector boards, cables FTA y marshalling completo (sala remota 6).","AESA","AESA","Única ampliación de HW de seguridad; módulos en rack existente.","Listado de Señales; ACAL-100-PL-K-100","P-145, P-303"],
 ["15","PSS","Configuración del sistema PSS","Excluida del alcance del Contratista.","PP","PP","Cambios de campo de AESA impactan la configuración → coordinar con PP.","Aclaración §3.7 (consulta)","P-303"],
 # ---------- ESD ----------
 ["16","ESD","Gabinetes de seguridad ESD — hardware (CPF-ESD-SC-07/08)","Hardware del sistema de parada de emergencia (HIMA/HIMAX). PP provee; AESA monta y cablea.","PP","AESA","🚩 Misma inconsistencia de código en P-408.","ACAL-100-PL-K-552/554; 021-ACAL-100-HD-K-311 (Válvulas ESD)","P-144, P-408"],
 ["17","ESD","Gabinetes de marshalling ESD (CPF2-ESD-MC-07A/08A)","Marshalling del ESD (ingeniería, provisión y montaje).","AESA","AESA","A cargo íntegro de AESA.","ACAL-100-PL-K-552/554","P-144/145, P-408"],
 ["18","ESD","Configuración del sistema ESD","Excluida del Contratista; PP entrega un controlador demo para pruebas.","PP","PP","Dependencia de insumos de PP para la integración.","Aclaración §3.7","P-303"],
 ["19","ESD/PSS/USS","Controlador demo para pruebas de integración","Provisto por PP para las pruebas de integración al SCADA.","PP","PP","🚩 Verificar plazo de entrega de PP vs. cronograma de pruebas.","Aclaración (consulta integración)","P-303"],
 # ---------- F&G ----------
 ["20","F&G","Estudio de Fire & Gas (ISA-TR 84.00.07)","Determina cantidad, ubicación y orientación de detectores de mezcla explosiva y llama (votación NooM).","AESA","—","Define el alcance físico de la detección; condiciona partidas de campo.","ISA-TR 84.00.07","P-171, P-242"],
 ["21","F&G","Detectores de campo (gas/mezcla explosiva, llama, térmico lineal), pulsadores, sirenas-balizas","Sistema de detección de campo F&G.","AESA","AESA","Cantidad/ubicación según estudio F&G.","Ad 01.G.08; SED Plano de ubicación de detectores de gas y fuego","Ad-69, P-242/250"],
 ["22","F&G","Sistema Fire & Gas — Hardware y Software","Partida 01.G.08 (alarmas, detectores, pulsadores, sirenas).","AESA","AESA","Apertura de partida solicitada por consulta de oferente.","Ad 01.G.08 (solapa B. Detalles)","Ad-69, P-233"],
 ["23","F&G","Centrales / controladores F&G","Considerados dentro de los instrumentos de campo; parte provistos con el proveedor de shelter (subcontrato).","PP / Prov. shelter","AESA / Prov.","🚩 Zona gris de provisión (campo vs. shelter); cerrar en ingeniería.","P-232; subcontrato de shelter","P-232"],
 ["24","F&G","Cajas de conexión F&G + cableado resistente al fuego","Multitriadas (analógicas) y multipares (digitales) desde JB al shelter. ESD-F&G comparten bandeja.","AESA","AESA","Cables F&G resistentes al fuego (optimizable con la ingeniería).","330-CPF2 — Redefinición de Alcance Rev.1","Ad-61"],
 ["25","F&G","Habilitación de la red F&G / auditoría entes provinciales (red húmeda)","Se excluye la habilitación de la red de F&G; auditoría de habilitación a cargo de PPSA.","PP","PP","Riesgo de interfaz en puesta en servicio; coordinar con habilitación AESA.","P-234 (Excepciones-Ingeniería)","P-234"],
 # ---------- OT/IT ----------
 ["26","OT/IT","Sistemas IT/OT (general) — provisión y montaje","Incluye ingeniería constructiva y de detalle para el montaje.","AESA","AESA","Base del alcance OT/IT de AESA.","001-ACAL-ET-O-100 (ET Disciplina OT)","P-409"],
 ["27","OT/IT","Servicios de Seguridad y CCTV","Cobertura de seguridad física-patrimonial y operaciones (CCTV Bosch, radares Spotter).","AESA","AESA","Plano de cobertura es orientativo; optimizable en ingeniería.","001-ACAL-PL-O-200 (cobertura); 001-ACAL-LI-O-203 (equipos)","P-146, P-409"],
 ["28","OT/IT","Control de Acceso — ingeniería y canalizaciones","Solo desarrollo de ingeniería e instalación de canalizaciones para implementación futura por terceros.","AESA","AESA","No incluye el equipamiento ni la integración (ver fila 29).","001-ACAL-ET-O-100 §3.8.1","P-146"],
 ["29","OT/IT","Equipamiento de control de acceso + licencias + integración a seguridad patrimonial","Fuera del alcance de AESA.","PP","PP","Frontera de responsabilidad con seguridad patrimonial existente.","001-ACAL-ET-O-100 §3.8.1","P-146"],
 ["30","OT/IT","Servicios WIFI In/Out","Cobertura de áreas exteriores y salas/shelters interiores.","AESA","AESA","Licencias e integración a IT Local quedan fuera (ver fila 32).","001-ACAL-ET-O-100 §3.8.2","P-146/147, P-409"],
 ["31","OT/IT","Rack de Comunicaciones IT","Provisión e instalación del rack; FO/UTP certificado; alimentación redundante; espejos de FO.","AESA","AESA","Concentra todos los servicios IT en salas/shelters.","001-ACAL-ET-O-100 §3.8.3","P-147, P-409"],
 ["32","OT/IT","Licencias e integración a la infraestructura IT Local de CPF","Fuera del alcance de AESA (aplica a WIFI y Rack IT).","PP","PP","🚩 Definir el punto de demarcación IT↔OT y la matriz de interfaz.","001-ACAL-ET-O-100 §3.8.2/3.8.3","P-147"],
 ["33","OT/IT","Infraestructura OT - FO (anillo de fibra óptica)","Provisión, instalación, fusionado y certificación de FO; continuidad del anillo de 24h con caminos disjuntos; PRP + redbox / anillos DLR.","AESA","AESA","Mantener filosofía y topología actuales del anillo OT.","226-ACAL-102-PL-K-301 (Anillo FO); 001-ACAL-ET-O-100 §3.8.4","P-147, P-305, P-409"],
 ["34","OT/IT","Tendidos y canalizaciones I&C","Bandejas portacables, cañeros de PVC Ø160, tritubo PEAD para FO.","AESA","AESA","Usar reservas de canalizaciones existentes para acceso a salas.","226-ACAL-102-PL-K-301","P-409"],
 ["35","OT/IT","Provisión de equipos IT/OT (servidores, switches, firewalls, radios)","Según lista de vendors aprobados (Cisco, Fortinet, Hirschmann, Stratix, MPLS Hitachi-ABB, Dell, etc.).","AESA","AESA","Solo partners oficiales autorizados por las marcas.","Lista de vendors PPSA — §3.8.5","P-148, P-409"],
 # ---------- Ciberseguridad / red (eliminados) ----------
 ["36","Ciberseg./Red","Configuración de routers, switches y firewalls","ELIMINADO del alcance de AESA (según respuestas del Anexo 3.8).","PP","PP","🚩 AESA provee los equipos (fila 35) pero NO los configura → ¿quién asegura la red OT?","Anexo 3.8 (respuestas); P-409","P-409"],
 ["37","Ciberseg./Red","Diagramas de zonas y conductos de ciberseguridad (IEC 62443)","ELIMINADO del alcance de AESA.","PP","—","🚩 Entregable de ciberseguridad fuera de AESA pese a la cláusula 43.15.","Anexo 3.8; IEC 62443","P-409"],
 ["38","Ciberseg./Red","Cálculo de niveles de disponibilidad · Tabla de IPS","ELIMINADO del alcance de AESA.","PP","—","Verificar quién ejecuta estos análisis de red.","Anexo 3.8","P-409"],
 ["39","Ciberseguridad","Cláusula contractual 43.15 Ciberseguridad","Marco COBIT / IEC 62443 / RGPD — obligaciones del Contratista ('Ciberseguridad en AESA').","AESA","—","🚩 AESA sigue obligada contractualmente pese a la eliminación de los entregables técnicos.","Cláusula 43.15; COBIT / IEC 62443 / RGPD","P-118, P-521"],
 # ---------- General / contractual ----------
 ["40","General","Regla por defecto (Adenda cláusula 3.1.4)","AESA suministra todo lo no indicado expresamente como provisión de PP en el Apéndice A, necesario para la obra.","AESA","AESA","Catch-all: traslada a AESA todo lo no listado para PP.","Adenda cláusula 3.1.4; Apéndice A","Ad-5"],
 ["41","General","Equipos de proceso provistos por PP (Apéndice A)","Equipos suministrados por Pluspetrol; AESA ejecuta su montaje.","PP","AESA","Riesgo de interfaz: montaje AESA sobre equipo provisto por PP.","Apéndice A","P-408, Ad-5"],
 ["42","Pruebas","FAT / iFAT / SAT / SIT, sintonía de lazos, taller de pantallas","Pruebas y puesta a punto de tableros y sistemas (con participación/insumos de PP).","AESA","AESA","Notificar a PP con antelación; FAT de tableros PP si cambia la matriz Causa-Efecto.","ACAL-000-MO-K-101; Matriz Causa-Efecto","P-145, P-303"],
]

NOTA_INCONSIST=("Nota de inconsistencia documental: en P-408 los gabinetes de seguridad aparecen rotulados con el "
 "código del PCS (CPF2-PCS-SC-07A/B). Según §3.7 (P-144), PP provee los gabinetes de seguridad "
 "CPF-ESD-SC-07/08 y CPF-PSS-SC-07/08, y AESA provee los del PCS. Conviene reconciliar la codificación "
 "de gabinetes en la ingeniería de detalle.")

# índices de columnas de responsable (para color)
IDX_PROC, IDX_INST = 4, 5

def color_resp(v):
    v=(v or "").upper()
    if v=="AESA": return VERDE
    if v=="PP": return ROJO
    if v in ("—",""): return None
    return AMAR  # mixto / compartido

# ----------------------------------------------------------- EXCEL
def gen_xlsx(path):
    wb=Workbook(); ws=wb.active; ws.title="Cuadro de Alcance"
    hf=PatternFill("solid",fgColor=AZUL[1:]); hfont=Font(bold=True,color="FFFFFF",size=10)
    thin=Side(style="thin",color="BFBFBF"); border=Border(thin,thin,thin,thin)
    wrap=Alignment(wrap_text=True,vertical="top")
    center=Alignment(wrap_text=True,vertical="center",horizontal="center")
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(COLS))
    ws.cell(row=1,column=1,value=META["titulo"]).font=Font(bold=True,size=13,color=AZUL[1:])
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(COLS))
    ws.cell(row=2,column=1,value=META["subtitulo"]+"   |   "+META["fuentes"]).font=Font(italic=True,size=9,color=AZUL2[1:])
    hr=4
    for i,h in enumerate(COLS,1):
        c=ws.cell(row=hr,column=i,value=h); c.fill=hf; c.font=hfont; c.border=border; c.alignment=center
    for r,row in enumerate(ROWS,hr+1):
        for i,v in enumerate(row,1):
            c=ws.cell(row=r,column=i,value=v); c.border=border
            c.alignment=center if i in (1,IDX_PROC+1,IDX_INST+1) else wrap
            if i in (IDX_PROC+1,IDX_INST+1):
                fill=color_resp(v)
                if fill: c.fill=PatternFill("solid",fgColor=fill[1:]); c.font=Font(bold=True,size=10)
    widths=[4,16,28,40,9,13,40,30,18]
    for i,w in enumerate(widths,1):
        ws.column_dimensions[get_column_letter(i)].width=w
    ws.freeze_panes="A5"
    ws.auto_filter.ref=f"A{hr}:{get_column_letter(len(COLS))}{hr+len(ROWS)}"
    nr=hr+len(ROWS)+2
    ws.merge_cells(start_row=nr,start_column=1,end_row=nr,end_column=len(COLS))
    ws.cell(row=nr,column=1,value=NOTA_INCONSIST).font=Font(italic=True,size=9,color="9C5700")
    ws.cell(row=nr,column=1).alignment=wrap
    lr=nr+2
    ws.cell(row=lr,column=1,value="Leyenda:").font=Font(bold=True)
    ws.cell(row=lr,column=2,value="AESA").fill=PatternFill("solid",fgColor=VERDE[1:])
    ws.cell(row=lr,column=3,value="PP (Pluspetrol)").fill=PatternFill("solid",fgColor=ROJO[1:])
    ws.cell(row=lr,column=4,value="Mixto / compartido").fill=PatternFill("solid",fgColor=AMAR[1:])
    # resumen
    ws2=wb.create_sheet("Resumen por responsable")
    from collections import Counter
    cp=Counter((r[IDX_PROC] or "").strip() for r in ROWS)
    ci=Counter((r[IDX_INST] or "").strip() for r in ROWS)
    ws2.cell(row=1,column=1,value="Conteo de ítems por responsable").font=Font(bold=True,size=12,color=AZUL[1:])
    ws2.cell(row=3,column=1,value="PROCURA").font=Font(bold=True); r=4
    for k,v in cp.most_common(): ws2.cell(row=r,column=1,value=k); ws2.cell(row=r,column=2,value=v); r+=1
    r+=1; ws2.cell(row=r,column=1,value="INSTALACIÓN / MONTAJE").font=Font(bold=True); r+=1
    for k,v in ci.most_common(): ws2.cell(row=r,column=1,value=k); ws2.cell(row=r,column=2,value=v); r+=1
    ws2.column_dimensions["A"].width=28; ws2.column_dimensions["B"].width=10
    wb.save(path); return path

# ----------------------------------------------------------- WORD (apaisado)
def gen_docx(path):
    doc=Document()
    sec=doc.sections[0]; sec.orientation=WD_ORIENT.LANDSCAPE
    sec.page_width, sec.page_height = sec.page_height, sec.page_width
    sec.left_margin=sec.right_margin=Cm(1.0); sec.top_margin=sec.bottom_margin=Cm(1.0)
    doc.styles["Normal"].font.name="Calibri"; doc.styles["Normal"].font.size=Pt(7.5)
    t=doc.add_heading(level=0); rr=t.add_run(META["titulo"]); rr.font.size=Pt(13); rr.font.color.rgb=RGBColor(0x1F,0x38,0x64)
    p=doc.add_paragraph(); r2=p.add_run(META["subtitulo"]); r2.italic=True
    doc.add_paragraph(META["fuentes"]).runs[0].font.size=Pt(7.5)
    tab=doc.add_table(rows=1,cols=len(COLS)); tab.style="Light Grid Accent 1"
    for i,h in enumerate(COLS):
        cell=tab.rows[0].cells[i]; cell.text=h
        for pp in cell.paragraphs:
            for rn in pp.runs: rn.font.bold=True; rn.font.size=Pt(7.5)
    for row in ROWS:
        cells=tab.add_row().cells
        for i,v in enumerate(row):
            cells[i].text=str(v)
            for pp in cells[i].paragraphs:
                for rn in pp.runs: rn.font.size=Pt(7)
    doc.add_paragraph()
    n=doc.add_paragraph(); rn=n.add_run(NOTA_INCONSIST); rn.italic=True; rn.font.size=Pt(8); rn.font.color.rgb=RGBColor(0x9C,0x57,0x00)
    leg=doc.add_paragraph(); leg.add_run("Leyenda: ").bold=True; leg.add_run("AESA = Contratista · PP = Pluspetrol · '—' = no aplica · 'Mixto' = compartido.")
    pie=doc.add_paragraph(f"Documento generado automáticamente — {datetime.now():%Y-%m-%d %H:%M}")
    pie.alignment=WD_ALIGN_PARAGRAPH.CENTER
    for rn in pie.runs: rn.font.size=Pt(7); rn.font.color.rgb=RGBColor(0x80,0x80,0x80)
    doc.save(path); return path

# ----------------------------------------------------------- PDF (apaisado)
def gen_pdf(path):
    doc=SimpleDocTemplate(path,pagesize=landscape(A4),leftMargin=0.8*cm,rightMargin=0.8*cm,topMargin=1.0*cm,bottomMargin=0.8*cm)
    ss=getSampleStyleSheet()
    h0=ParagraphStyle("h0",parent=ss["Title"],textColor=colors.HexColor(AZUL),fontSize=12,leading=14)
    sub=ParagraphStyle("sub",parent=ss["Italic"],fontSize=8,textColor=colors.HexColor(AZUL2))
    cell=ParagraphStyle("c",parent=ss["BodyText"],fontSize=6.2,leading=7.4)
    cellb=ParagraphStyle("cb",parent=cell,alignment=1)
    small=ParagraphStyle("sm",parent=cell,fontSize=6.5,textColor=colors.HexColor("#9C5700"))
    el=[]; W=doc.width
    el.append(Paragraph(META["titulo"],h0)); el.append(Paragraph(META["subtitulo"]+" — "+META["fuentes"],sub)); el.append(Spacer(1,4))
    # anchos (fracción de W): #, Sistema, Alcance, Descripción, Proc, Inst, Obs, DocRef, Pág
    cw=[W*0.022,W*0.075,W*0.135,W*0.165,W*0.05,W*0.06,W*0.165,W*0.135,W*0.073]
    data=[[Paragraph(f"<b>{h}</b>",cellb) for h in COLS]]
    for row in ROWS:
        data.append([Paragraph(str(v),cellb if i in(0,IDX_PROC,IDX_INST) else cell) for i,v in enumerate(row)])
    t=Table(data,colWidths=cw,repeatRows=1)
    style=[("BACKGROUND",(0,0),(-1,0),colors.HexColor(AZUL)),("TEXTCOLOR",(0,0),(-1,0),colors.white),
           ("GRID",(0,0),(-1,-1),0.4,colors.HexColor("#CCCCCC")),("VALIGN",(0,0),(-1,-1),"TOP"),
           ("FONTSIZE",(0,0),(-1,-1),6.2),("LEFTPADDING",(0,0),(-1,-1),2),("RIGHTPADDING",(0,0),(-1,-1),2),
           ("TOPPADDING",(0,0),(-1,-1),1.5),("BOTTOMPADDING",(0,0),(-1,-1),1.5)]
    for r,row in enumerate(ROWS,1):
        for idx in (IDX_PROC,IDX_INST):
            f=color_resp(row[idx])
            if f: style.append(("BACKGROUND",(idx,r),(idx,r),colors.HexColor(f)))
    t.setStyle(TableStyle(style)); el.append(t); el.append(Spacer(1,5))
    el.append(Paragraph(NOTA_INCONSIST,small))
    el.append(Paragraph("Leyenda: verde = AESA · rojo = PP (Pluspetrol) · amarillo = mixto/compartido · '—' = no aplica.",cell))
    doc.build(el); return path

def main():
    out=sys.argv[1] if len(sys.argv)>1 else os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    os.makedirs(out,exist_ok=True)
    base=os.path.join(out,"Cuadro_Alcance_Control_OT-IT_CPF2")
    res=[gen_xlsx(base+".xlsx"), gen_docx(base+".docx"), gen_pdf(base+".pdf")]
    print("Generados:")
    for x in res: print("  -",x)

if __name__=="__main__":
    main()
