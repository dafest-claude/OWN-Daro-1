#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera un Word NUEVO = Pliego subcontrato E&I (preservado) + Anexos por especialidad.
Anexo E (Electricidad) y Anexo I (Instrumentación) con cantidades representativas:
 tendido (formaciones/metros/secciones), conexionado (puntas por formación/sección),
 montaje de tableros / instrumentos por tipo / en línea y no / cajas de conexión,
 y personal especializado para precomisionado y comisionado.
Salida: 'Pliego subcontrato E&I - Rev1 - 2026-06-30 (con Anexos E-I).docx'
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT

BASE="/home/user/OWN-Daro-1/Constructibilidad"
SRC=os.path.join(BASE,"Datos entrada","Pliego subcontrato E&I.docx")
OUT=os.path.join(BASE,"Análisis generados","Pliego subcontrato E&I - Rev1 - 2026-06-30 (con Anexos E-I).docx")
REV="1"; FECHA="30-06-2026"
NAVY=RGBColor(0x1F,0x3B,0x63)

doc=Document(SRC)

def h(txt,lvl=1):
    p=doc.add_heading(txt,level=lvl)
    return p
def para(txt,bold=False,size=10,italic=False,color=None):
    p=doc.add_paragraph(); r=p.add_run(txt); r.bold=bold; r.italic=italic; r.font.size=Pt(size)
    if color: r.font.color.rgb=color
    return p
def bullet(txt):
    try:
        p=doc.add_paragraph(txt,style='List Bullet')
    except KeyError:
        p=doc.add_paragraph(); p.add_run("•  "+txt); p.paragraph_format.left_indent=Inches(0.25)
    for r in p.runs: r.font.size=Pt(10)
    return p
def table(headers,rows,widths=None,total_row=False):
    t=doc.add_table(rows=1,cols=len(headers)); t.style='Table Grid'; t.alignment=WD_TABLE_ALIGNMENT.CENTER
    hc=t.rows[0].cells
    for j,htxt in enumerate(headers):
        hc[j].text=''; r=hc[j].paragraphs[0].add_run(htxt); r.bold=True; r.font.size=Pt(8.5); r.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
        sh=hc[j]._tc.get_or_add_tcPr()
        from docx.oxml.ns import qn; from docx.oxml import OxmlElement
        shd=OxmlElement('w:shd'); shd.set(qn('w:fill'),'1F3B63'); sh.append(shd)
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        for j,v in enumerate(row):
            cells[j].text=''; rr=cells[j].paragraphs[0].add_run(str(v)); rr.font.size=Pt(8.5)
            if total_row and ri==len(rows)-1: rr.bold=True
            if j>0: cells[j].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    return t

# ============ PORTADA DE ANEXOS ============
doc.add_page_break()
p=doc.add_paragraph(); r=p.add_run("ANEXOS TÉCNICOS POR ESPECIALIDAD"); r.bold=True; r.font.size=Pt(16); r.font.color.rgb=NAVY
p.alignment=WD_ALIGN_PARAGRAPH.CENTER
para("Subcontratación de personal especialista y herramientas — Electricidad e Instrumentación",bold=True,size=11)
para("Proyecto: CPF2 La Calera II (Vaca Muerta)  ·  Revisión %s  ·  Fecha: %s"%(REV,FECHA),size=10,color=RGBColor(0x55,0x55,0x55))
para("Estos anexos complementan el Pliego (cuerpo principal, sin modificaciones) e indican las CANTIDADES "
     "REPRESENTATIVAS a ejecutar por especialidad, como base para el dimensionamiento de personal especialista "
     "y herramientas para montaje, instalación, conexionado, precomisionado y comisionado, según los perfiles "
     "definidos en el Pliego. Cantidades de ingeniería vigente (sujetas a revisión).",size=9.5,italic=True)

# ===================== ANEXO E — ELECTRICIDAD =====================
h("ANEXO E — ELECTRICIDAD: CANTIDADES REPRESENTATIVAS",1)
para("Volumen global EL (CPF-2): 513 cables · 81.231 m · ~5.178 puntas de conexión · 243 cargas "
     "(79 motores [76 BT + 3 MT] · 6 VFD · 52 instrumentos · 37 SSAA · 28 iluminación · 29 control · 12 calefacción).",size=10)

h("E.1  Tendido de cables eléctricos — por tipo y sección",2)
table(["Tipo / sección","Cables","Metros","Observación"],
 [["Potencia MT (6,6 / 13,2 kV)","18","2.110","Terminaciones por técnico certificado"],
  ["Potencia BT Grande (≥ 35 mm²)","104","17.470","3x240 · 3x150 · 3x95 · 3x50 — frente crítico"],
  ["Potencia BT Mediana / Pequeña","198","33.125","3x16 a 3x4 · 4x4 · iluminación"],
  ["Control y Señales (multiconductor)","192","28.126","7x2,5+T · 2x2,5 · señales VFD/vibros."],
  ["Fibra Óptica / Ethernet","1","400","Red PMS SE#3↔SE#4"],
  ["TOTAL EL (CPF-2)","513","81.231","—"]],total_row=True)

h("E.2  Tendido — desglose por formación (principales)",2)
table(["Formación","Constitución / tipo","Cables","Metros"],
 [["7x2,5+T","1kV XLPE Arm-T Cu — Control","96","17.390"],
  ["2x4","1kV XLPE Arm Cu — Potencia BT","105","10.905"],
  ["2x2,5","1kV XLPE Arm Cu — Control","39","7.270"],
  ["2x6","1kV XLPE Arm Cu — Potencia BT","15","4.480"],
  ["3x240","1kV XLPE Arm Cu — Potencia BT","23","4.050"],
  ["3x120/70","1kV XLPE Arm Cu — Potencia BT","8","4.000"],
  ["4x4","1kV XLPE Cu — Potencia BT","20","3.860"],
  ["4x10+T","1kV XLPE Arm-T-RF Cu — Potencia BT","24","3.560"],
  ["3x35 / 3x50 / 3x95","1kV XLPE Arm Cu — Potencia BT","40","7.400"],
  ["8x3x1,31","300V Arm BG+BI — Señal","8","1.700"],
  ["Otras formaciones BT/Control/Señal/FO","resto","≈170","≈16.616"]])
para("Nota: el listado completo de 49 formaciones está disponible en el análisis de cables eléctricos (Rev.4). "
     "Las secciones gobiernan la herramienta de terminación (lugs a compresión por sección, prensacables/glands por armadura).",size=9,italic=True)

h("E.3  Conexionado — puntas / terminaciones por sección",2)
table(["Terminación por sección","Cantidad","Unidad","Herramienta / perfil"],
 [["Terminales MT (6,6 / 13,2 kV)","36","extremos","Premoldeadas — técnico MT certificado"],
  ["Terminales BT Grande (≥ 95 mm²)","208","extremos","Lugs a compresión + termocontraíble"],
  ["Terminales BT Mediana / Pequeña","396","extremos","Lugs preaislados / a compresión"],
  ["Cables de control (peinado+etiquet.+prueba)","192","cables","Ferrules de bornera"],
  ["TOTAL puntas EL (conductores + glands)","~5.178","puntas","3.036 en salas · 1.717 en campo"]],total_row=True)

h("E.4  Montaje de tableros y equipos eléctricos",2)
table(["Ítem","Cantidad","Observación"],
 [["Tableros BT / CCMs","37","Llegan en salas SE#3/SE#4 (ABB) — alineación, anclaje, busbars"],
  ["Celdas MT (TGMT / TRMT)","6","Izaje + nivelación + interconexión de barras"],
  ["Transformadores","s/proyecto","Salas eléctricas"],
  ["Variadores (VFD)","6","3 MT + 3 BT"],
  ["Motores (montaje/alineación)","79","76 BT + 3 MT (>90 kW, 6,6 kV)"],
  ["Salas eléctricas (armado de shelters en campo)","2","SE#3 (21 d) + SE#4 (14 d) — módulos"]])

h("E.5  Personal especializado para precomisionado y comisionado — Electricidad",2)
para("Provisión de personal especialista y herramientas para las pruebas y puesta en servicio eléctrica:",size=10)
bullet("Pruebas de aislación (Megger) circuito por circuito — 513 cables (HOLD POINT ITP). Registro conforme.")
bullet("Energización progresiva MT → Transformadores → BT (HOLD POINT, con presencia ABB/Cliente).")
bullet("Precomisionado funcional de tableros y CCMs: ajuste de relés de protección, verificación de enclavamientos.")
bullet("Precomisionado de motores: secuencias de arranque, sentido de giro, protecciones, tuning de VFD.")
bullet("Comisionado integrado EL + PMS (ABB): deslastre de cargas, lógicas DAG, interfaz con PCS.")
bullet("Personal: electricistas de pruebas, técnicos MT certificados (ABB field service), ingeniería de comisionado.")

# ===================== ANEXO I — INSTRUMENTACIÓN =====================
doc.add_page_break()
h("ANEXO I — INSTRUMENTACIÓN: CANTIDADES REPRESENTATIVAS",1)
para("Volumen global IN (CPF-2): 885 cables · 65.315 m · 15.752 puntas de conexión · 2.798 instrumentos "
     "(de los cuales AESA monta ~1.163: 897 provistos por AESA + 266 'montaje por EPC'). Sistemas: PCS · ESD/SIS · F&G · PSS · SCADA.",size=10)

h("I.1  Tendido de cables de instrumentación — por formación",2)
table(["Formación","Pantalla / armadura","Cables","Metros","Puntas"],
 [["1P #16AWG Sh-A","Sh · Armado","409","13.980","3.272"],
  ["24P #18AWG ShT-A","Sh T · Armado","27","5.035","2.700"],
  ["24P #18AWG Sh I/T-A","Sh I/T · Armado","20","3.755","2.000"],
  ["2P #16AWG ShT-A","Sh T · Armado","86","2.145","1.032"],
  ["1T #16AWG Sh I/T-A (FR)","Sh I/T FR · Armado","79","1.975","790"],
  ["8T #14AWG Sh I/T-A (FR)","Sh I/T FR · Armado","15","3.180","780"],
  ["12P #18AWG Sh I/T-A","Sh I/T · Armado","12","3.140","624"],
  ["8P #18AWG ShT-A","Sh T · Armado","17","4.035","612"],
  ["4P #18AWG Sh I/T-A","Sh I/T · Armado","26","6.135","520"],
  ["14C #14AWG-A · 7C #14AWG-A","— · Armado","40","4.245","878"],
  ["2C #2,5 mm²-A","— · Armado","23","5.200","138"],
  ["FO Monomodo 6H / 24H","— · Armado","13","3.220","290"],
  ["FTP Cat. 6 (4P)","Sh T · No arm.","3","80","6"],
  ["Otras formaciones (resto)","instr. apantallado armado","115","8.990","2.110"],
  ["TOTAL IN (CPF-2)","—","885","65.315","15.752"]],total_row=True)
para("Tipos predominantes: par/terna apantallado armado #16-18AWG (4-20 mA, HART, RTD, TC), multipar (24P/12P/8P), "
     "FO monomodo y FTP Cat.6. La armadura agrega prensacables/glands en cada extremo.",size=9,italic=True)

h("I.2  Conexionado — puntas por tipo",2)
table(["Concepto","Cantidad","Unidad","Observación"],
 [["Puntas de conductor","12.376","puntas","Ferrules / borneras en Head, JB y Sala INS"],
  ["Puntas de pantalla","1.612","puntas","Aterramiento de malla"],
  ["Puntas de armadura (glands)","1.764","puntas","Prensacables en cada extremo"],
  ["TOTAL puntas IN","15.752","puntas","≈ 1.974 horas-hombre de conexionado"]],total_row=True)
para("Recorrido típico: instrumento de campo → Junction Box (JB) → Shelter Sala INS → PCS / SIS. "
     "Conexionado en salas/Sala INS admite doble turno.",size=9,italic=True)

h("I.3  Montaje de instrumentos — por tipo y modalidad",2)
para("Total 2.798 instrumentos. Por modalidad de montaje:",size=10)
table(["Modalidad de montaje","Cantidad","Detalle"],
 [["EN LÍNEA — válvulas","485","Control · seguridad (PSV) · bloqueo (ESDV/SDV/BDV/XV) · regulación"],
  ["EN LÍNEA — elementos","435","Placa orificio (FE) · orificio restricción (RO) · termovaina (TW) · toma muestra · cupón"],
  ["Soporte junto a línea (stand)","610","Transmisores / manómetros de presión-caudal con toma a proceso"],
  ["Sobre equipo (recipiente/máquina)","640","Nivel · presión/temperatura sobre equipo · vibración"],
  ["Sobre válvula (accesorio)","483","Posicionadores · solenoides · finales de carrera · actuadores"],
  ["Sobre otro instrumento","25","Transmisores remotos (caudal sobre placa, analizador sobre sonda)"],
  ["Estructura / campo (F&G)","120","Detectores de gas/llama · sirenas · balizas · pulsadores"],
  ["TOTAL","2.798","En línea 920 · montaje específico 1.878"]],total_row=True)
para("Por tipo (principales): Manómetro (PI) 269 · Termovaina (TW) 221 · Sensor temp. (TE) 178 · Válvula seguridad (PSV) 162 · "
     "Transmisor presión (PIT) 160 · Orificio restricción (RO) 127 · Indicador nivel magnético (LG) 126 · "
     "Transmisor temp. (TIT) 87 · Termómetro (TI) 76 · Transmisor nivel (LIT) 66 · fines de carrera, solenoides, etc.",size=9)
para("Alcance de montaje AESA: ~1.163 instrumentos (897 provistos por AESA + 266 de proveedor con 'montaje por EPC'). "
     "Los premontados en skids de vendor (~1.635) los instala el vendor, pero el conexionado y loop check son alcance del subcontrato.",size=9,italic=True)

h("I.4  Montaje de tableros en Sala de Instrumentación",2)
table(["Ítem","Cantidad","Observación"],
 [["PCS — remotas 7A + 7B (Inauco)","2","Tableros de control en Sala INS (Sala 7)"],
  ["Tablero marshalling ESD/F&G/PSS (HIMA)","1","Hardware de seguridad — interconexión a SIS"],
  ["Gabinetes / racks de Sala INS","s/proyecto","Montaje, fijación, interconexión interna"]])

h("I.5  Montaje de instrumentos en línea vs montaje específico",2)
bullet("Instrumentos EN LÍNEA (sobre cañería, parte del spool): 920 — válvulas 485 + elementos 435. Se instalan con la cañería.")
bullet("Instrumentos de MONTAJE ESPECÍFICO (según típico de montaje): 1.878 — sobre soporte/stand, equipo, válvula o estructura.")
bullet("Cada modalidad define el típico de montaje, soportería y herramienta correspondiente.")

h("I.6  Montaje de cajas de conexión (Junction Boxes)",2)
para("Montaje, fijación y rotulado de cajas de conexión (JB) de campo (DCS y SIS) en los recorridos "
     "instrumento → JB → Sala INS. Cantidad NO indicada en la ingeniería disponible — a relevar/confirmar "
     "(se computa por relevamiento de planos de montaje y disposición de JB por área).",size=10)

h("I.7  Personal especializado para precomisionado y comisionado — Instrumentación",2)
para("Provisión de personal especialista y herramientas para las pruebas y puesta en servicio de instrumentación:",size=10)
bullet("Calibración en banco de instrumentos (pre-montaje) — patrones y banco de pruebas.")
bullet("Prueba punta a punta (continuidad) e identificación de cada cable / punta.")
bullet("Prueba de señales en campo: verificación de lazos (loop check) sensor → JB → Sala INS → PCS/SIS (~885 lazos).")
bullet("Precomisionado de lazos PCS / SIS / F&G: simulación de señales, matriz de causa-efecto (MCE).")
bullet("Comisionado integrado con PCS (Inauco), SIS (HIMA) y SCADA: pruebas funcionales y de seguridad.")
bullet("Personal: técnicos instrumentistas, especialistas en calibración, ingeniería de lazos y comisionado IN.")

# ===================== CIERRE =====================
doc.add_page_break()
h("NOTA SOBRE LAS CANTIDADES",1)
para("Las cantidades de estos anexos provienen de la ingeniería vigente del proyecto (listas de instrumentos "
     "ACAL-00102 / ACAL-00670, análisis de cables eléctricos Rev.4 y de cables de instrumentación Rev.0) y son "
     "REPRESENTATIVAS para el dimensionamiento de personal y herramientas. Están sujetas a la revisión de "
     "ingeniería y al relevamiento de obra. El cómputo final para certificación se ajustará a planos 'For Construction'.",size=10)
para("Documento: Pliego subcontrato E&I — Revisión %s — %s. Anexos E (Electricidad) e I (Instrumentación) agregados; "
     "cuerpo principal del Pliego sin modificaciones."%(REV,FECHA),size=9,italic=True,color=RGBColor(0x55,0x55,0x55))

doc.save(OUT)
print("OK ->",OUT)
print("Anexos agregados. Total párrafos:",len(doc.paragraphs),"| tablas:",len(doc.tables))
