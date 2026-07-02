#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Word = Pliego subcontrato E&I (cuerpo original PRESERVADO) + Anexos en continuidad de estilo.
Rev.2: carátula propia por anexo (hoja nueva) + detalle en hoja nueva; columna 'Herramienta'
(solo herramientas, sin perfil) en conexionado EL; tableros SE#3/#4 vienen montados (salvo los
que ABB indique por transporte); instrumentos diferenciados por provisión (AESA vs proveedor).
Salida: 'Pliego subcontrato E&I - Rev2 - 2026-06-30 (con Anexos).docx'
"""
import os
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK

BASE="/home/user/OWN-Daro-1/Constructibilidad"
SRC=os.path.join(BASE,"Datos entrada","Pliego subcontrato E&I.docx")
OUT=os.path.join(BASE,"Análisis generados","Pliego subcontrato E&I - Rev3 - 2026-07-02 (con Anexos).docx")
REV="3"; FECHA="02-07-2026"
NAVY=RGBColor(0x1F,0x3B,0x63)
doc=Document(SRC)
TBL_STYLE='Tabla Cuadro'
try: doc.styles[TBL_STYLE]
except KeyError: TBL_STYLE='Table Grid'

def h(txt,lvl=1): return doc.add_heading(txt,level=lvl)
def para(txt,note=False):
    p=doc.add_paragraph(); r=p.add_run(txt)
    if note: r.italic=True; r.font.size=Pt(9); r.font.color.rgb=RGBColor(0x55,0x55,0x55)
    return p
def bullet(txt):
    p=doc.add_paragraph(); p.paragraph_format.left_indent=Inches(0.3); p.add_run("•  "+txt); return p
def pagebreak():
    doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
def cover(title,subtitle=None):
    """Carátula del anexo en hoja propia (título = Heading 1, queda en el índice);
    el detalle arranca en la hoja siguiente."""
    pagebreak()
    for _ in range(7): doc.add_paragraph()
    hp=doc.add_heading(title,level=1); hp.alignment=WD_ALIGN_PARAGRAPH.CENTER
    for r in hp.runs: r.font.size=Pt(28); r.font.color.rgb=NAVY
    if subtitle:
        ps=doc.add_paragraph(); ps.alignment=WD_ALIGN_PARAGRAPH.CENTER
        rs=ps.add_run(subtitle); rs.font.size=Pt(13); rs.font.color.rgb=RGBColor(0x55,0x55,0x55)
    pagebreak()   # el detalle empieza en hoja nueva
def table(headers,rows,bold_total=False):
    t=doc.add_table(rows=1,cols=len(headers)); t.style=TBL_STYLE
    for j,htxt in enumerate(headers):
        cell=t.rows[0].cells[j]; cell.text=''
        r=cell.paragraphs[0].add_run(htxt); r.bold=True; r.font.size=Pt(8.5)
    for ri,row in enumerate(rows):
        cells=t.add_row().cells
        for j,v in enumerate(row):
            cells[j].text=''; rr=cells[j].paragraphs[0].add_run(str(v)); rr.font.size=Pt(8.5)
            if bold_total and ri==len(rows)-1: rr.bold=True
            if j>0: cells[j].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    return t

# ===== PORTADILLA GENERAL DE ANEXOS =====
pagebreak()
h("ANEXOS TÉCNICOS POR ESPECIALIDAD — CANTIDADES REPRESENTATIVAS",1)
para("Los presentes anexos forman parte integrante del presente Pliego y lo complementan, sin modificar su cuerpo "
     "principal. Indican las cantidades representativas de la ingeniería vigente, por especialidad (Electricidad e "
     "Instrumentación), como base para el dimensionamiento del personal especialista y las herramientas a proveer por "
     "el Subcontratista para las tareas de montaje, instalación, conexionado, precomisionado y comisionado, conforme a "
     "los perfiles definidos en este Pliego.")
para("Revisión %s — %s. Fuentes: listas de instrumentos ACAL-00102/00670, análisis de cables eléctricos Rev.4 y de "
     "cables de instrumentación Rev.0. Cantidades sujetas a revisión de ingeniería y a planos 'For Construction'."%(REV,FECHA),note=True)

# ============================ ANEXO E ============================
cover("ANEXO E","ELECTRICIDAD")
para("Volumen global de Electricidad (CPF-2): 513 cables · 81.231 m · aproximadamente 5.178 puntas de conexión · "
     "243 cargas (79 motores [76 BT + 3 MT] · 6 variadores · 52 instrumentos · 37 SS.AA. · 28 iluminación · 29 control · 12 calefacción/trazado).")

h("E.1  Tendido de cables — resumen por tipo y sección",2)
table(["Tipo / sección","Cables","Metros","Observación"],
 [["Potencia MT (6,6 / 13,2 kV)","18","2.110","Maniobras especiales MT; coordinación con ABB"],
  ["Potencia BT Grande (≥ 35 mm²)","104","17.470","3x240 · 3x150 · 3x95 · 3x50 — tendido pesado en bandeja"],
  ["Potencia BT Mediana / Pequeña","198","33.125","3x16 a 3x4 · 4x4 · iluminación"],
  ["Control y Señales (multiconductor)","192","28.126","7x2,5+T · 2x2,5 · señales VFD / vibroswitches"],
  ["Fibra Óptica / Ethernet","1","400","Red PMS SE#3 ↔ SE#4"],
  ["TOTAL ELECTRICIDAD (CPF-2)","513","81.231","—"]],bold_total=True)

h("E.2  Tendido de cables — detalle completo por formación",2)
EL_FORM=[["7x2,5+T","Control","96","17.390"],["2x4","Potencia BT","105","10.905"],["2x2,5","Control","39","7.270"],
 ["2x6","Potencia BT","15","4.480"],["3x240","Potencia BT","23","4.050"],["3x120/70","Potencia BT","8","4.000"],
 ["4x4","Potencia BT","20","3.860"],["4x10+T","Potencia BT","24","3.560"],["3x35","Potencia BT","16","2.960"],
 ["3x50","Potencia BT","10","2.700"],["3x6","Potencia BT","11","2.350"],["3x16","Potencia BT","9","2.290"],
 ["4x4+T","Potencia BT","12","1.820"],["1x240","Potencia BT","65","1.810"],["3x95","Potencia BT","14","1.740"],
 ["3x10","Potencia BT","10","1.740"],["8x3x1,31","Señal","8","1.700"],["3x25","Potencia BT","9","1.560"],
 ["4x25/16","Potencia BT","12","1.200"],["2x2,5+T","Señal","14","1.050"],["2x10","Control","6","1.040"],
 ["12x2,5","Señal","8","600"],["3x25/16","Potencia BT","6","590"],["3x50+T","Potencia BT","3","570"],
 ["3x25+T","Potencia BT","3","570"],["2x35","Potencia BT","3","540"],["F.O. multimodo","Fibra Óptica","3","480"],
 ["2x50","Control","2","410"],["3x35/16","Potencia BT","3","400"],["3x70","Potencia BT","2","400"],
 ["3x4","Potencia BT","2","300"],["7x2,5","Control","10","290"],["15 cond.","Control","9","216"],
 ["3x2,5","Control","6","160"],["3x95/50","Potencia BT","4","160"],["1x70 (13,2 kV)","Potencia MT","6","120"],
 ["2x4+T","Potencia BT","1","100"],["3x150","Potencia BT","5","100"],["3x150/70","Potencia BT","4","80"],
 ["2x25","Control","2","80"],["3x50/25","Potencia BT","1","60"],["2x95","Control","5","50"],
 ["3x2,5+T","Control","2","40"],["1x2x0,5","Señal","2","30"],["3x240/120","Potencia BT","1","20"],
 ["2x120","Potencia BT","2","20"],["2x16","Potencia BT","2","20"],["6x2x0,5","Señal","4","0"],
 ["2x2x0,5","Señal","6","0"],["TOTAL (incl. CPF-1)","—","633","85.881"]]
table(["Formación","Tipo","Cables","Metros"],EL_FORM,bold_total=True)
para("El total CPF-2 (alcance) es 513 cables / 81.231 m; la diferencia corresponde a cables CPF-1 de referencia. "
     "Las secciones gobiernan la herramienta de terminación.",note=True)

h("E.3  Conexionado — puntas / terminaciones por sección",2)
table(["Terminación por sección","Cantidad","Unidad","Herramienta"],
 [["Terminales MT (6,6 / 13,2 kV)","36","extremos","Herramienta hidráulica de compresión MT, kit de terminación premoldeada, pistola de calor, torquímetro, megóhmetro 5 kV"],
  ["Terminales BT Grande (≥ 95 mm²)","208","extremos","Prensa hidráulica de compresión con matrices, lugs, mangas termocontraíbles, soplete"],
  ["Terminales BT Mediana / Pequeña","396","extremos","Pinza/tenaza de compresión, lugs preaislados, peladora de cable"],
  ["Cables de control (peinado + etiquetado + prueba)","192","cables","Peladora, pinza de ferrules (crimp), rotuladora/identificador, multímetro de continuidad"],
  ["TOTAL puntas EL (conductores + glands)","~5.178","puntas","—"]],bold_total=True)

h("E.4  Montaje de tableros y equipos eléctricos",2)
para("Las salas eléctricas SE#3 (BT) y SE#4 (MT) se reciben con su equipamiento (tableros BT/CCMs, celdas MT y "
     "transformadores) MONTADO de fábrica por ABB. El alcance de campo del Subcontratista comprende el armado de los "
     "módulos del shelter, la interconexión entre módulos/celdas y el montaje en campo ÚNICAMENTE de los tableros o "
     "componentes que ABB indique a posteriori que se transportan desmontados por razones de transporte.")
table(["Ítem","Cantidad","Observación"],
 [["Armado de shelters en campo (salas en módulos)","2","SE#3 (21 d) + SE#4 (14 d): ensamble, fijación, interconexión interna"],
  ["Tableros / componentes a montar en campo","A definir por ABB","Sólo los transportados desmontados por razones de transporte"],
  ["Tableros BT / CCMs (referencia, vienen montados)","37","Montados de fábrica en las salas — verificación y reapriete"],
  ["Celdas MT (TGMT / TRMT) (vienen montadas)","6","Montadas de fábrica — interconexión de barras, nivelación"],
  ["Transformadores (vienen montados)","s/proyecto","En salas eléctricas"],
  ["Motores (montaje y alineación en campo)","79","76 BT + 3 MT (>90 kW, 6,6 kV) — equipo de campo"],
  ["Variadores (VFD)","6","3 MT + 3 BT (integrados en salas)"]])

h("E.5  Personal especializado para precomisionado y comisionado — Electricidad",2)
para("Provisión de personal especialista y herramientas para las pruebas y puesta en servicio eléctrica:")
for b in ["Pruebas de aislación (Megger) circuito por circuito — 513 cables (HOLD POINT ITP), con registro conforme.",
 "Energización progresiva MT → Transformadores → BT (HOLD POINT, con presencia ABB / Cliente).",
 "Precomisionado funcional de tableros y CCMs: ajuste de relés de protección y verificación de enclavamientos.",
 "Precomisionado de motores: secuencias de arranque, sentido de giro, protecciones y tuning de VFD.",
 "Comisionado integrado Electricidad + PMS (ABB): deslastre de cargas, lógicas DAG, interfaz con PCS.",
 "Perfiles: electricistas de pruebas, técnicos MT certificados (ABB field service) e ingeniería de comisionado."]:
    bullet(b)

h("E.6  Puesta a tierra (PAT)",2)
para("Actividades para completar la puesta a tierra según el listado de materiales ACAL-102-LM-E-0017 "
     "(Puesta a Tierra y SPCDA — Underground). Comprende el tendido de conductores, sus conexiones y el montaje "
     "de los electrodos y accesorios:")
for b in ["Tendido de conductor de PAT de cobre desnudo en distintas secciones (p. ej. 70 / 50 / 35 mm²) en zanja, "
          "bandeja y sobre estructura.",
 "Armado y ejecución de conexiones por terminal de compresión en frío cable-cable y cable-jabalina "
 "(según IRAM 2349 / IEEE 837) y/o soldadura exotérmica (tipo cadweld) donde corresponda.",
 "Montaje de jabalinas / electrodos de puesta a tierra y cámaras de inspección.",
 "Conexión de la malla de PAT a estructuras, equipos, tableros, motores, bandejas portacables y sistema de mallado.",
 "Provisión de herramientas de compresión hidráulica, moldes y cargas para soldadura exotérmica y consumibles."]:
    bullet(b)
table(["Actividad / material principal (según LM-E-0017)","Unidad","Cantidad"],
 [["Conductor de PAT — cobre desnudo (por sección)","m","según LM-E-0017"],
  ["Terminales de compresión en frío (cable-cable / cable-jabalina)","u","según LM-E-0017"],
  ["Soldaduras exotérmicas (cadweld)","u","según LM-E-0017"],
  ["Jabalinas / electrodos de PAT","u","según LM-E-0017"],
  ["Cámaras de inspección de PAT","u","según LM-E-0017"]])

h("E.7  Protección contra descargas atmosféricas (SPCDA)",2)
para("Actividades para completar el sistema de protección contra descargas atmosféricas, según ACAL-102-LM-E-0017:")
for b in ["Montaje de terminales aéreos / pararrayos y mástiles.",
 "Tendido de conductores de bajada y su fijación.",
 "Ejecución de tomas de tierra dedicadas del SPCDA y sus conexiones (compresión / exotérmica).",
 "Interconexión del SPCDA con la malla de puesta a tierra."]:
    bullet(b)
table(["Actividad / material principal (según LM-E-0017)","Unidad","Cantidad"],
 [["Terminales aéreos / pararrayos y mástiles","u","según LM-E-0017"],
  ["Conductores de bajada","m","según LM-E-0017"],
  ["Tomas de tierra dedicadas SPCDA + conexiones","u","según LM-E-0017"]])

h("E.8  Iluminación y tomas",2)
para("Actividades para completar la instalación de iluminación y tomacorrientes según el listado de materiales "
     "ACAL-102-LM-E-0020 (Iluminación y Tomas):")
for b in ["Montaje de artefactos de iluminación (luminarias) interiores y exteriores, incluida iluminación de emergencia.",
 "Montaje de tomacorrientes y sus cajas.",
 "Montaje de cajas de conexión de iluminación / tomas.",
 "Provisión y montaje de cañería / bandeja, tendido y conexionado de los circuitos de iluminación y tomas."]:
    bullet(b)
table(["Actividad / material principal (según LM-E-0020)","Unidad","Cantidad"],
 [["Artefactos de iluminación (luminarias)","u","según LM-E-0020"],
  ["Tomacorrientes","u","según LM-E-0020"],
  ["Cajas de conexión de iluminación / tomas","u","según LM-E-0020"],
  ["Cañería / bandeja + tendido y conexionado de circuitos","m / gl","según LM-E-0020"]])

# ============================ ANEXO I ============================
cover("ANEXO I","INSTRUMENTACIÓN")
para("Volumen global de Instrumentación (CPF-2): 885 cables · 65.315 m · 15.752 puntas de conexión · "
     "2.798 instrumentos. Sistemas involucrados: PCS · ESD/SIS · F&G · PSS · SCADA.")

h("I.1  Tendido de cables de instrumentación — detalle completo por formación",2)
IN_FORM=[["1P #16AWG Sh-A","Sh · Armado","409","13.980","3.272"],["24P #18AWG ShT-A","Sh T · Armado","27","5.035","2.700"],
 ["24P #18AWG Sh I/T-A","Sh I/T · Armado","20","3.755","2.000"],["2P #16AWG ShT-A","Sh T · Armado","86","2.145","1.032"],
 ["1T #16AWG Sh I/T-A (FR)","Sh I/T FR · Armado","79","1.975","790"],["8T #14AWG Sh I/T-A (FR)","Sh I/T FR · Armado","15","3.180","780"],
 ["12P #18AWG Sh I/T-A","Sh I/T · Armado","12","3.140","624"],["8P #18AWG ShT-A","Sh T · Armado","17","4.035","612"],
 ["12P #18AWG ShT-A","Sh T · Armado","10","2.070","520"],["4P #18AWG Sh I/T-A","Sh I/T · Armado","26","6.135","520"],
 ["8P #18AWG Sh I/T-A","Sh I/T · Armado","12","2.985","432"],["1P #16AWG Sh-A (FR)","Sh FR · Armado","48","1.375","384"],
 ["7C #14AWG-A","— · Armado","23","575","368"],["12P #18AWG ShT-A (FR)","Sh T FR · Armado","7","1.485","364"],
 ["1T #16AWG Sh-A","Sh · Armado","32","160","320"],["14C #14AWG-A","— · Armado","17","3.670","510"],
 ["2C #2,5 mm²-A","— · Armado","23","5.200","138"],["FO Monomodo 24H","— · Armado","3","1.000","150"],
 ["FO Monomodo 6H","— · Armado","10","2.220","140"],["4P #18AWG ShT-A","Sh T · Armado","3","865","60"],
 ["1T #16AWG ShT-A (FR)","Sh T FR · Armado","3","250","30"],["FTP Categoría 6 (4P)","Sh T · No arm.","3","80","6"],
 ["TOTAL INSTRUMENTACIÓN (CPF-2)","—","885","65.315","15.752"]]
table(["Formación","Pantalla / armadura","Cables","Metros","Puntas"],IN_FORM,bold_total=True)
para("Tipos predominantes: par/terna apantallado armado #16-18AWG (4-20 mA, HART, RTD, TC), multipar (24P/12P/8P), "
     "FO monomodo y FTP Cat.6. La armadura agrega prensacables/glands en cada extremo.",note=True)

h("I.2  Conexionado — puntas por tipo",2)
table(["Concepto","Cantidad","Unidad","Observación"],
 [["Puntas de conductor","12.376","puntas","Ferrules / borneras en Head, JB y Sala INS"],
  ["Puntas de pantalla","1.612","puntas","Aterramiento de malla"],
  ["Puntas de armadura (glands)","1.764","puntas","Prensacables en cada extremo"],
  ["TOTAL puntas IN","15.752","puntas","≈ 1.974 horas-hombre de conexionado"]],bold_total=True)

h("I.3  Instrumentos por provisión — AESA vs proveedores",2)
para("Diferenciación del total de instrumentos según quién los provee, para acotar el alcance de provisión y montaje:")
table(["Provisión","Cantidad","Detalle"],
 [["Provistos por AESA","897","Compra AESA. Montaje e instalación por el Subcontratista (cuenta AESA)"],
  ["Provistos por proveedores (en equipos / skids)","1.894","Vendor (PROPAK, skids de gas/químicos, válvulas, bombas, etc.)"],
  ["   — de ellos, 'montaje por EPC'","266","Provistos por vendor pero MONTADOS por AESA (alcance del Subcontratista)"],
  ["   — premontados en skid por el vendor","1.628","Llegan montados; el Subcontratista realiza conexionado y loop check"],
  ["Existentes (reutilizar / reemplazar)","7","A definir"],
  ["TOTAL INSTRUMENTOS","2.798","Montaje a cargo de AESA: ~1.163 (897 + 266)"]],bold_total=True)

h("I.4  Montaje de instrumentos — por tipo y modalidad",2)
para("Total 2.798 instrumentos. Distribución por modalidad de montaje:")
table(["Modalidad de montaje","Cantidad","Detalle"],
 [["EN LÍNEA — válvulas","485","Control · seguridad (PSV) · bloqueo (ESDV/SDV/BDV/XV) · regulación"],
  ["EN LÍNEA — elementos","435","Placa orificio (FE) · orificio restricción (RO) · termovaina (TW) · toma muestra · cupón"],
  ["Soporte junto a línea (stand)","610","Transmisores / manómetros de presión-caudal con toma a proceso"],
  ["Sobre equipo (recipiente / máquina)","640","Nivel · presión/temperatura sobre equipo · vibración"],
  ["Sobre válvula (accesorio)","483","Posicionadores · solenoides · finales de carrera · actuadores"],
  ["Sobre otro instrumento","25","Transmisores remotos (caudal sobre placa, analizador sobre sonda)"],
  ["Estructura / campo (F&G)","120","Detectores de gas/llama · sirenas · balizas · pulsadores"],
  ["TOTAL","2.798","En línea 920 · montaje específico 1.878"]],bold_total=True)
para("Por tipo (principales): Manómetro (PI) 269 · Termovaina (TW) 221 · Sensor temp. (TE) 178 · Válvula seguridad (PSV) 162 · "
     "Transmisor presión (PIT) 160 · Orificio restricción (RO) 127 · Indicador nivel magnético (LG) 126 · "
     "Transmisor temp. (TIT) 87 · Termómetro (TI) 76 · Transmisor nivel (LIT) 66 · fines de carrera, solenoides, etc.")

h("I.5  Montaje de tableros en Sala de Instrumentación",2)
table(["Ítem","Cantidad","Observación"],
 [["PCS — remotas 7A + 7B (Inauco)","2","Tableros de control en Sala INS (Sala 7)"],
  ["Tablero marshalling ESD / F&G / PSS (HIMA)","1","Hardware de seguridad — interconexión a SIS"],
  ["Gabinetes / racks de Sala INS","s/proyecto","Montaje, fijación e interconexión interna"]])

h("I.6  Montaje de instrumentos en línea vs montaje específico",2)
bullet("Instrumentos EN LÍNEA (sobre cañería, parte del spool): 920 — válvulas 485 + elementos 435. Se instalan con la cañería.")
bullet("Instrumentos de MONTAJE ESPECÍFICO (según típico de montaje): 1.878 — sobre soporte/stand, equipo, válvula o estructura.")
bullet("Cada modalidad define el típico de montaje, la soportería y la herramienta correspondiente.")

h("I.7  Montaje de cajas de conexión (Junction Boxes)",2)
para("Montaje, fijación y rotulado de las cajas de conexión (JB) de campo en los recorridos instrumento → JB → Sala INS, "
     "según ACAL-00102-LM-K-0002 (montaje eléctrico de instrumentos) y las hojas de datos ACAL-00102-HD-K-0022 / "
     "ACAL-00670-HD-K-0011. Las JB son de fundición de aluminio (IP66), con borneras de 28 / 40 / 80 bornes (y 40 para "
     "ternas), típicamente 350 × 300 × 200 mm, segregadas por sistema (PCS, ESD, PSS, USS, F&G).")
table(["Sistema","JB (cantidad)","Referencia de tag"],
 [["PCS (JBA / JBD)","20","JBA-… / JBD-…"],
  ["PSS (JBPA / JBPD)","32","JBPA-… / JBPD-…"],
  ["F&G (JBFA / JBFD)","22","JBFA-… / JBFD-… (ajustar a requerimientos HIMA)"],
  ["ESD (JBED)","4","JBED-…"],
  ["USS / otras (JBUD / JBSD)","2","JBUD-… / JBSD-…"],
  ["TOTAL (relevado en la documentación disponible)","~80","—"]],bold_total=True)
para("Actividad a computar como INSTRUMENTACIÓN, adicional al conexionado:")
for b in ["Montaje y fijación de la caja de conexión + rotulado / identificación.",
 "Colocación de los prensacables (glands) que correspondan según la cantidad y sección de los cables que ingresan y "
 "egresan de cada caja, con sus tapones ciegos en las entradas no utilizadas.",
 "Conexión de borneras y puentes de puesta a tierra internos de la caja.",
 "Nota: las puntas de conexión ya están computadas en el conteo de cables IN (I.2), dado que incluyen los tramos de "
 "cable; el MONTAJE de la caja y la colocación de prensacables se computan aparte como actividad propia."]:
    bullet(b)
para("El cómputo de JB podrá ampliarse: el LM-K-0002 indica que a la fecha incluye solamente canalizaciones troncales; "
     "el resto de los materiales se incorporará en próximas revisiones.",note=True)

h("I.8  Personal especializado para precomisionado y comisionado — Instrumentación",2)
para("Provisión de personal especialista y herramientas para las pruebas y puesta en servicio de instrumentación:")
for b in ["Calibración en banco de instrumentos (previa al montaje), con patrones y banco de pruebas.",
 "Prueba punta a punta (continuidad) e identificación de cada cable y punta.",
 "Prueba de señales en campo: verificación de lazos (loop check) sensor → JB → Sala INS → PCS/SIS (~885 lazos).",
 "Precomisionado de lazos PCS / SIS / F&G: simulación de señales y prueba de matriz de causa-efecto (MCE).",
 "Comisionado integrado con PCS (Inauco), SIS (HIMA) y SCADA: pruebas funcionales y de seguridad.",
 "Perfiles: técnicos instrumentistas, especialistas en calibración e ingeniería de lazos y comisionado."]:
    bullet(b)

# ============================ ANEXO P ============================
cover("ANEXO P","PRECOMISIONADO · COMISIONADO · PUESTA EN MARCHA")
para("Alcance transversal de precomisionado, comisionado y puesta en marcha (PEM) para ambas especialidades. El "
     "Subcontratista proveerá el personal especialista y TODO el instrumental de prueba, calibración y medición "
     "necesario, en cantidad y con certificados de calibración vigentes, para completar las pruebas, la liberación "
     "y la entrega de los sistemas.")

h("P.1  Secuencia y alcance",2)
for b in ["Precomisionado: pruebas estáticas y de subsistema (megger, continuidad, calibración en banco, loop check, "
          "medición de PAT, ajuste de protecciones) previas a la energización/puesta en tensión.",
 "Comisionado: pruebas funcionales dinámicas por sistema y su integración (EL + PMS + PCS + SIS + F&G + SCADA).",
 "Puesta en marcha (PEM): arranque asistido, verificación en condiciones de proceso y acompañamiento hasta la "
 "estabilización, con el personal especialista requerido.",
 "Las pruebas y aceptación de equipos (FAT / iFAT) deben incluir equipos de precomisionado para asegurar la "
 "liberación y entrega."]:
    bullet(b)

h("P.2  Instrumental de prueba — Electricidad",2)
for b in ["Telurímetro (medición de resistencia de puesta a tierra) y micro-ohmímetro (continuidad de mallado).",
 "Megóhmetro (aislación) BT y MT (5 kV), con registro conforme ITP.",
 "Equipo de inyección primaria/secundaria para ensayo y ajuste de relés de protección; maletín de pruebas.",
 "Secuencímetro, pinza amperométrica, multímetros, analizador de red / calidad de energía.",
 "Cámara termográfica para inspección de conexiones bajo carga."]:
    bullet(b)

h("P.3  Instrumental de prueba — Instrumentación",2)
for b in ["Calibrador / generador de lazo 4-20 mA y comunicador HART; fuente de lazo.",
 "Patrones de presión (bomba/comparador y patrón), de temperatura (baño/horno seco y patrón) y de caudal según aplique.",
 "Banco de calibración para prueba de instrumentos previa al montaje.",
 "Generador de señales (RTD/TC/mV) y equipo para prueba de lazos (loop check) y matriz de causa-efecto (MCE).",
 "Detector portátil de gas para verificación de sensores de F&G."]:
    bullet(b)

h("P.4  Herramientas y materiales misceláneos (general)",2)
para("El Subcontratista proveerá todas las herramientas de montaje (manuales, eléctricas, hidráulicas y de izaje menor) "
     "y los materiales misceláneos y consumibles necesarios para completar los trabajos de montaje, conexionado, "
     "precomisionado, comisionado y puesta en marcha, entre ellos: elementos de fijación y soportería menor, "
     "prensacables y tapones, terminales y ferrules, precintos e identificadores/rótulos, mangas termocontraíbles, "
     "cintas, selladores, y todo elemento menor requerido para la correcta terminación de las tareas, esté o no "
     "detallado explícitamente en los listados de materiales.")

# ============================ ANEXO C ============================
cover("ANEXO C","PLANILLA DE CÓMPUTO Y PRECIOS UNITARIOS")
para("El Oferente deberá completar los precios unitarios (columna P. Unitario) para cada ítem; el Subtotal y el Total "
     "se obtienen por producto de cantidad × precio unitario. Las cantidades son las de la ingeniería vigente y podrán "
     "ajustarse a planos 'For Construction'.")
h("C.1  Electricidad",2)
table(["Ítem de obra","Unidad","Cantidad","P. Unitario","Subtotal"],
 [["Tendido cables Potencia MT","m","2.110","",""],["Tendido cables Potencia BT Grande (≥35 mm²)","m","17.470","",""],
  ["Tendido cables Potencia BT Mediana / Pequeña","m","33.125","",""],["Tendido cables Control y Señales","m","28.126","",""],
  ["Conexionado terminales MT","extremo","36","",""],["Conexionado terminales BT Grande","extremo","208","",""],
  ["Conexionado terminales BT Mediana / Pequeña","extremo","396","",""],["Conexionado cables de control","cable","192","",""],
  ["Armado de shelters en campo (salas)","u","2","",""],["Montaje en campo de tableros (los que ABB indique)","u","a definir","",""],
  ["Montaje y alineación de motores","u","79","",""],
  ["Puesta a tierra — tendido de conductor Cu desnudo","m","s/LM-E-0017","",""],
  ["Puesta a tierra — conexiones (compresión / exotérmica) y jabalinas","u","s/LM-E-0017","",""],
  ["SPCDA — pararrayos/mástiles, bajadas y tomas de tierra","global","s/LM-E-0017","",""],
  ["Iluminación — montaje de luminarias","u","s/LM-E-0020","",""],
  ["Tomas — montaje de tomacorrientes y cajas","u","s/LM-E-0020","",""],
  ["Precomisionado, comisionado y puesta en marcha eléctrico","global","1","",""]])
h("C.2  Instrumentación",2)
table(["Ítem de obra","Unidad","Cantidad","P. Unitario","Subtotal"],
 [["Tendido cables de instrumentación","m","65.315","",""],["Conexionado de puntas (instrumentación)","punta","15.752","",""],
  ["Montaje de instrumentos en línea","u","920","",""],["Montaje de instrumentos de montaje específico","u","1.878","",""],
  ["Montaje de tableros en Sala INS (PCS / marshalling)","u","3","",""],
  ["Montaje de cajas de conexión (JB)","u","~80","",""],
  ["Colocación de prensacables (glands) en JB","u","s/cables por JB","",""],
  ["Loop check / prueba de señales en campo","lazo","885","",""],
  ["Precomisionado, comisionado y puesta en marcha de instrumentación","global","1","",""]])
para("Nota: la planilla es indicativa para la cotización; el cómputo final para certificación se ajustará a planos "
     "'For Construction' y al relevamiento de obra.",note=True)
para("Documento: Pliego subcontrato E&I — Revisión %s — %s. Anexos E (Electricidad), I (Instrumentación) y C (Planilla "
     "de cómputo) agregados como continuidad; cuerpo principal del Pliego sin modificaciones."%(REV,FECHA),note=True)

doc.save(OUT)
print("OK ->",OUT,"| estilo tabla:",TBL_STYLE,"| tablas:",len(doc.tables))
