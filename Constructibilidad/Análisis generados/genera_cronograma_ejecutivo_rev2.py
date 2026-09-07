#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CRONOGRAMA EJECUTIVO — Provisiones críticas CPF-2 (Rev2, 2026-09-07).
Cambios respecto de Rev1:
  1) ABB Salas SE#3 y SE#4: se incorporan los hitos de ENSAYOS/FAT EN FÁBRICA DE ORIGEN
     (Brasil) para tableros y celdas, entre el acopio y el montaje en el shelter, más la
     adecuación/despacho de fábrica y la entrega en Shelterista. Se reordena la secuencia.
  2) ACOPIO DE MATERIALES: se abre en hitos propios por provisión y se agrega una hoja
     dedicada "Acopio de Materiales" con el detalle por paquete (pedido del cliente).
  3) HIMA se separa en dos provisiones: HPH (HIMA Paul Hilpert — Alemania, tableros de
     marshalling) y HARG (HIMA Argentina — ingeniería de aplicación, programación,
     documentación y FAT/IFAT), según PSCH_CPF2_Rev3 (04-09-2026).
Corte de alcance: hasta la ENTREGA DE EQUIPOS / del alcance de cada provisión.
"""
import os
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
REV="Rev2"; FECHA="2026-09-07"; HOY=date(2026,9,7)
D=lambda y,m,d: date(y,m,d)

# (hito, inicio, fin, crítico, observación, es_acopio)
PROV=[
 ("ABB — Sala Eléctrica SE#3 (BT)","ABB","C62828",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"OC recibida en ABB",0),
  ("Kick-off Meeting (KOM) técnico",D(2026,5,15),D(2026,5,18),0,"",0),
  ("Confirmación de layout de sala y equipamiento",D(2026,5,18),D(2026,5,20),0,"Insumo AESA",0),
  ("Emisión de ingeniería básica",D(2026,5,20),D(2026,8,28),0,"",0),
  ("Aprobación de ingeniería básica",D(2026,8,28),D(2026,9,11),0,"Aprobación AESA",0),
  ("FREEZING POINT — ingeniería básica",D(2026,9,11),D(2026,9,11),1,"Hito explícito del cronograma ABB",0),
  ("Emisión de ingeniería de detalle",D(2026,5,22),D(2026,9,21),0,"",0),
  ("Aprobación de ingeniería de detalle",D(2026,9,21),D(2026,10,5),0,"Aprobación AESA",0),
  ("FREEZING POINT — ingeniería de detalle",D(2026,10,5),D(2026,10,5),1,"Cierra ingeniería del shelter",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,19),D(2026,5,19),0,"Compra y acopio sala BT",1),
  ("ACOPIO — Baterías / rectificadores / UPS",D(2026,6,9),D(2026,10,12),0,"Paquete de CC",1),
  ("ACOPIO — Magnéticos, semiconductores, interruptores y fusibles",D(2026,7,13),D(2026,10,5),0,"Componentes de tableros",1),
  ("ACOPIO — Gabinetes y accesorios",D(2026,7,13),D(2026,10,12),0,"",1),
  ("ACOPIO — Tableros auxiliares",D(2026,7,22),D(2026,11,9),0,"",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,12,10),D(2026,12,10),1,"Cierre de acopio de la sala BT",1),
  ("Inicio de fabricación en fábrica de origen (Brasil)",D(2026,10,12),D(2026,10,12),0,"Tableros BT / CCM — M&B Brasil",0),
  ("Fabricación de tableros en fábrica de origen (Brasil)",D(2026,10,12),D(2026,12,7),1,"CCM-005, CCM-007, auxiliares, rectif./UPS",0),
  ("Ensayos internos en fábrica de origen (Brasil)",D(2026,10,22),D(2026,12,4),0,"Previos a la FAT",0),
  ("FAT EN FÁBRICA DE ORIGEN (Brasil) — tableros y CCM",D(2026,11,9),D(2026,12,14),1,"Witness AESA — rectif./UPS 09-11/11 · aux. 07-10/12 · CCM 14/12",0),
  ("Adecuación y despacho de fábrica de origen",D(2026,12,9),D(2026,12,18),1,"Condición EXW en origen",0),
  ("Entrega en Shelterista (Mendoza)",D(2026,11,23),D(2027,1,12),1,"Arribo escalonado de los paquetes",0),
  ("Montaje e integración de equipos en sala",D(2027,2,19),D(2027,3,15),0,"Shelterista",0),
  ("Interconexión de equipos",D(2027,3,16),D(2027,3,23),0,"",0),
  ("FAT SALA ELÉCTRICA (shelter completo)",D(2027,3,25),D(2027,3,30),1,"Witness AESA",0),
  ("Ensayos PMS en sala (INTERFASE — contrato PMS)",D(2027,3,31),D(2027,4,6),0,"Alcance del contrato PMS",0),
  ("DESPACHO / ENTREGA EXW",D(2027,4,21),D(2027,4,23),1,"HITO FINAL — cierre del alcance de la sala",0),
 ]),
 ("ABB — Sala Eléctrica SE#4 (MT)","ABB","AD1457",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"OC recibida en ABB",0),
  ("Kick-off Meeting (KOM) técnico",D(2026,5,15),D(2026,5,18),0,"",0),
  ("Confirmación de layout de sala y equipamiento",D(2026,5,18),D(2026,5,20),0,"Insumo AESA",0),
  ("Emisión de ingeniería básica",D(2026,5,20),D(2026,8,28),0,"",0),
  ("Aprobación de ingeniería básica",D(2026,8,28),D(2026,9,11),0,"Aprobación AESA",0),
  ("FREEZING POINT — ingeniería básica",D(2026,9,11),D(2026,9,11),1,"Hito explícito del cronograma ABB",0),
  ("Emisión de ingeniería de detalle",D(2026,5,22),D(2026,9,21),0,"",0),
  ("Aprobación de ingeniería de detalle",D(2026,9,21),D(2026,10,5),0,"Aprobación AESA",0),
  ("FREEZING POINT — ingeniería de detalle",D(2026,10,5),D(2026,10,5),1,"Cierra ingeniería del shelter",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,19),D(2026,5,19),0,"Compra y acopio sala MT",1),
  ("ACOPIO — Celdas MT (fábrica de origen)",D(2026,8,18),D(2026,10,17),1,"Equipo de largo plazo",1),
  ("ACOPIO — Variadores / drives",D(2026,8,21),D(2026,10,20),0,"",1),
  ("ACOPIO — Tableros auxiliares",D(2026,7,22),D(2026,11,9),0,"",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,12,10),D(2026,12,10),1,"Cierre de acopio de la sala MT",1),
  ("Inicio de fabricación en fábrica de origen (Brasil)",D(2026,10,5),D(2026,10,5),0,"Tableros MT / CCM — M&B Brasil",0),
  ("Fabricación de tableros y celdas en fábrica de origen (Brasil)",D(2026,10,21),D(2026,12,7),1,"CCM 21/10-11/11 · auxiliares 09/11-07/12",0),
  ("Ensayos internos en fábrica de origen (Brasil)",D(2026,11,12),D(2026,11,23),0,"Previos a la FAT",0),
  ("FAT EN FÁBRICA DE ORIGEN (Brasil) — tableros y celdas MT",D(2026,11,24),D(2026,12,10),1,"Witness AESA — CCM 24-30/11 · auxiliares 07-10/12",0),
  ("Despacho de fábrica de origen (celdas MT y tableros)",D(2026,10,20),D(2026,12,15),1,"Celdas MT despachan 20-27/10",0),
  ("Entrega en Shelterista (Mendoza)",D(2026,12,15),D(2027,1,11),1,"Arribo escalonado de los paquetes",0),
  ("Montaje e integración de equipos en sala",D(2027,1,21),D(2027,2,22),0,"Shelterista",0),
  ("Interconexión de equipos",D(2027,2,11),D(2027,2,23),0,"",0),
  ("FAT SALA ELÉCTRICA (shelter completo)",D(2027,2,24),D(2027,3,1),1,"Witness AESA",0),
  ("Ensayos PMS en sala (INTERFASE — contrato PMS)",D(2027,3,2),D(2027,3,9),0,"Alcance del contrato PMS",0),
  ("DESPACHO / ENTREGA EXW",D(2027,3,19),D(2027,3,25),1,"HITO FINAL — cierre del alcance de la sala",0),
 ]),
 ("ABB — Ductos de barras","ABB","795548",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"Mismo contrato ABB (alcance separado)",0),
  ("Emisión layout instalación + detalle conexionado trafo",D(2026,7,20),D(2026,8,29),1,"INSUMO AESA — condiciona el inicio de diseño",0),
  ("FREEZING POINT — inicio de diseño",D(2026,8,29),D(2026,8,29),1,"Hito explícito del cronograma ABB",0),
  ("Diseño de ductos de barras",D(2026,8,31),D(2026,11,20),0,"",0),
  ("FREEZING POINT — cierre de diseño",D(2026,12,4),D(2026,12,4),1,"Habilita fabricación",0),
  ("ACOPIO / fabricación de ductos de barras",D(2026,12,4),D(2027,3,14),1,"100 días corridos (acopio integrado a fabricación)",1),
  ("FCA fábrica de origen",D(2027,3,15),D(2027,3,19),1,"Puesta a disposición en origen",0),
  ("Traslado internacional + aduana",D(2027,3,19),D(2027,6,7),1,"80 días corridos — tramo más largo",0),
  ("ENTREGA DE EQUIPOS EN SITIO",D(2027,6,8),D(2027,6,10),1,"HITO FINAL — la entrega más tardía del paquete ABB",0),
 ]),
 ("ABB — PMS (Power Management System)","ABB","EF6C00",[
  ("Inicio de contrato — Recepción / aceptación OC",D(2026,5,8),D(2026,5,11),0,"",0),
  ("Kick-off Meeting (KOM)",D(2026,5,22),D(2026,5,22),0,"",0),
  ("Emisión de ingeniería básica",D(2026,6,8),D(2026,7,31),0,"Hito cert. 10%",0),
  ("Aprobación de ingeniería básica",D(2026,7,6),D(2026,8,14),0,"Hito cert. 10%",0),
  ("Emisión de ingeniería Rev.0",D(2026,7,20),D(2026,8,28),0,"",0),
  ("Aprobación ingeniería Rev.0 — FREEZING POINT",D(2026,8,3),D(2026,9,11),1,"Cierra ingeniería",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,22),D(2026,5,22),0,"Switches de red, S800, tableros",1),
  ("ACOPIO — Material S800 y tableros",D(2026,5,22),D(2026,10,8),0,"",1),
  ("ACOPIO — Switches de red",D(2026,5,22),D(2026,11,5),1,"Ítem de mayor plazo",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,11,5),D(2026,11,5),1,"Hito cert. 25%",1),
  ("Inicio de armado de tableros",D(2026,9,11),D(2026,9,11),0,"",0),
  ("Configuración / programación del sistema",D(2026,8,17),D(2026,12,11),1,"I/O, comms IEC61850/PROFI/MODBUS, lógicas",0),
  ("Pruebas de Maqueta Integrada PMS/CCM/PCS",D(2026,9,22),D(2026,9,25),0,"Prueba temprana de integración",0),
  ("Aprobación procedimientos FAT Rev.0",D(2026,10,12),D(2026,10,23),0,"",0),
  ("Fin de armado de tableros",D(2026,12,17),D(2026,12,17),1,"",0),
  ("Pruebas internas (pre-FAT)",D(2026,12,18),D(2027,1,7),0,"",0),
  ("FAT Tableros + Sistema",D(2027,1,25),D(2027,1,29),1,"Hito cert. 20% · witness AESA",0),
  ("Ajustes post-FAT / liberación",D(2027,2,1),D(2027,2,3),0,"",0),
  ("Despachos",D(2027,2,4),D(2027,2,17),1,"",0),
  ("ENTREGA DE EQUIPAMIENTO EN SHELTER",D(2027,2,17),D(2027,2,17),1,"HITO FINAL — cert. 30%",0),
 ]),
 ("Inauco — PCS / SCADA / Comunicaciones","Inauco","1565C0",[
  ("Inicio de contrato — Recepción de OC",D(2026,7,28),D(2026,7,28),0,"",0),
  ("Recepción de ing. Rev.0 p/ tableros (insumo AESA)",D(2026,8,10),D(2026,8,10),1,"Requisito 1 — insumo de AESA",0),
  ("Emisión de ingeniería (Rev. A y 0)",D(2026,8,11),D(2026,11,4),0,"",0),
  ("Configuración PLC — Etapa 1",D(2026,8,26),D(2026,9,22),0,"",0),
  ("Configuración SCADA",D(2026,8,26),D(2027,3,23),1,"",0),
  ("Aprobación ing. p/ construcción — FREEZING POINT",D(2026,11,19),D(2026,11,19),1,"Requisito 3",0),
  ("ACOPIO — Recepción de hardware (Rockwell / Stratix)",D(2026,11,26),D(2026,11,26),1,"Requisito 6 — cierre de acopio",1),
  ("Inicio de construcción de tableros",D(2026,11,27),D(2026,11,27),0,"",0),
  ("Recepción tableros HIMA en base Inauco (insumo HPH)",D(2027,1,25),D(2027,1,25),1,"Requisito 2 — interfase HPH/HARG",0),
  ("Configuración PLC — Etapa 2",D(2027,1,25),D(2027,3,23),0,"",0),
  ("Fin de construcción de tableros",D(2027,2,16),D(2027,2,16),1,"",0),
  ("Aprobación de procedimientos FAT",D(2027,2,24),D(2027,2,24),0,"Requisito 4",0),
  ("FAT PCS + Comunicaciones",D(2027,3,29),D(2027,4,22),1,"Witness AESA",0),
  ("FAT SIS (conjunto con HARG/HPH)",D(2027,3,29),D(2027,4,26),1,"Requiere presencia HIMA",0),
  ("Configuraciones post-FAT",D(2027,4,27),D(2027,5,17),0,"",0),
  ("Pruebas de integración PCS↔PMS (Bs.As.)",D(2027,4,28),D(2027,5,4),1,"Interfase con ABB PMS",0),
  ("ENTREGA DE EQUIPOS (post-FAT)",D(2027,5,17),D(2027,5,17),1,"HITO FINAL — cierre post-FAT; habilita el inicio de campo",0),
 ]),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","HPH","6A1B9A",[
  ("Inicio de contrato / proyecto",D(2026,7,10),D(2026,7,10),0,"",0),
  ("Inicio de ingeniería",D(2026,8,14),D(2026,8,14),0,"",0),
  ("Ingeniería de detalle / especificación",D(2026,8,17),D(2026,9,11),0,"",0),
  ("Fin de ingeniería / CAD — FREEZING POINT",D(2026,9,18),D(2026,9,18),1,"Cierra ingeniería",0),
  ("ACOPIO / disposición de materiales",D(2026,9,14),D(2026,10,16),1,"Integrado al plan de fabricación HPH",1),
  ("Fabricación de tableros de marshalling",D(2026,9,21),D(2026,11,6),1,"ESD/F&G y PSS 07A/07B",0),
  ("Hardware fabricado y probado (test de fábrica)",D(2026,11,6),D(2026,11,6),1,"FAT integrada a fabricación",0),
  ("Preparación de despacho (packing)",D(2026,11,9),D(2026,11,20),0,"",0),
  ("Listo para despacho",D(2026,11,20),D(2026,11,20),1,"",0),
  ("Despacho / shipping",D(2026,11,23),D(2026,12,14),1,"",0),
  ("Documentación de entrega enviada",D(2026,12,4),D(2026,12,4),0,"",0),
  ("Entrega DAP Ezeiza (Argentina)",D(2026,12,15),D(2026,12,15),1,"",0),
  ("RECEPCIÓN DE TABLEROS EN BASE INAUCO",D(2027,1,25),D(2027,1,25),1,"HITO FINAL — nacionalización + transporte",0),
 ]),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","HARG","00838F",[
  ("Documentación HARG — Doc. 1 (HW y otros)",D(2026,9,14),D(2026,9,25),0,"2 semanas",0),
  ("Programación HARG — Recurso 1",D(2026,9,21),D(2027,1,29),1,"11 semanas",0),
  ("Documentación HPH — Doc. 1 (INTERFASE)",D(2026,9,28),D(2026,10,16),0,"Insumo de HPH (3,5 semanas)",0),
  ("Ingeniería de Aplicación SIS y documentación (INTERFASE HPH)",D(2026,11,9),D(2026,12,18),1,"6 semanas — a definir por HPH",0),
  ("Documentación HARG — Doc. 2 (Modbus)",D(2026,11,9),D(2026,11,13),0,"1 semana",0),
  ("Fabricación HPH — entrega DIC-2026 (INTERFASE)",D(2026,12,21),D(2026,12,21),1,"Insumo de HPH",0),
  ("Adecuación de tableros marshalling (INTERFASE Inauco)",D(2026,12,21),D(2027,1,29),1,"6 semanas — a confirmar por Inauco",0),
  ("Programación HARG — Recurso 2",D(2027,1,4),D(2027,1,29),0,"5 semanas",0),
  ("Documentación HARG — Doc. 3 (Proc. FAT e IFAT)",D(2027,1,25),D(2027,2,12),0,"3 semanas",0),
  ("FAT",D(2027,2,24),D(2027,2,26),1,"0,5 semana",0),
  ("IFAT (Recursos 1 y 2)",D(2027,3,1),D(2027,3,12),1,"2 semanas — integración",0),
  ("Documentación HARG — Doc. 5 (Proy. Doc.) y HPH Doc. 2",D(2027,3,18),D(2027,3,22),0,"Documentación de cierre",0),
  ("ENTREGA DEL ALCANCE HARG (IFAT + documentación)",D(2027,3,12),D(2027,3,22),1,"HITO FINAL — cierre del alcance HARG",0),
 ]),
]
NOTAS=[
 "Alcance: cronograma EJECUTIVO por hitos, con corte en la ENTREGA DE EQUIPOS / del alcance de cada provisión. Las tareas posteriores al despacho/entrega (montaje en planta, SAT, precomisionado, comisionado, PEM, CAO/DataBook) quedan FUERA.",
 "Rev2 — cambio 1: en ABB Salas SE#3 y SE#4 se incorporaron los hitos de fabricación, ensayos internos y FAT EN FÁBRICA DE ORIGEN (Brasil, M&B) para tableros y celdas, más la adecuación/despacho de fábrica y la entrega en Shelterista, entre el acopio y el montaje del shelter.",
 "Rev2 — cambio 2: el ACOPIO DE MATERIALES se abrió en hitos propios por paquete y se agregó la hoja 'Acopio de Materiales' con el detalle consolidado (pedido del cliente para medir avance).",
 "Rev2 — cambio 3: HIMA se separó en HPH (HIMA Paul Hilpert, Alemania — tableros de marshalling) y HARG (HIMA Argentina — ingeniería de aplicación, programación, documentación y FAT/IFAT), según PSCH_CPF2_Rev3 del 04-09-2026.",
 "Fuentes: ABB Salas y Ductos 5155-00-0009-VE-CO-PG-001_A (21-08-26) · ABB PMS 5155-00-0007-VE-CO-PG-001_REV B (.mpp) · Inauco 5155-00-0029-VE-CO-PG-001-A Rev.A (24-08-26) · HPH 103618-Schedule 16-08-26 · HARG PSCH_CPF2_Rev3 (04-09-26).",
 "ALERTA DE INTERFASE A CONCILIAR: HPH entrega los tableros en DIC-2026 (DAP Ezeiza 15-12-26); el cronograma HARG asume disponibilidad para adecuación el 21-12-26, mientras que Inauco fija la recepción en su base el 25-01-27 (Requisito 2). Conviene alinear estas tres fechas.",
 "Los hitos de FAT en fábrica de origen (Brasil) son witness de AESA y anteceden a la FAT del shelter completo; no deben confundirse entre sí.",
 "Estado (plan): se calcula contra la fecha de corte. 'Cumplido s/plan' = la fecha planificada ya venció y debe confirmarse el cumplimiento real con el proveedor; 'En curso' = en ejecución; 'Pendiente' = aún no inició.",
 "Las columnas Fecha real / % avance / Estado real / Observaciones quedan en blanco para la carga de seguimiento periódico (reporte a PPSA).",
 "Los hitos se ordenan cronológicamente dentro de cada provisión y se renumeran (H01, H02, ...).",
]
INTERFASES=[
 "HPH → recepción de tableros de marshalling en base Inauco (25/01/27): habilita la adecuación, la configuración PLC Etapa 2 y la FAT SIS.",
 "HARG ↔ HPH → ingeniería de aplicación SIS y documentación de HPH (nov-dic/26): condicionan la programación y el FAT/IFAT de HARG.",
 "HARG ↔ Inauco → adecuación de tableros marshalling (21/12/26 → 29/01/27) en base Inauco.",
 "ABB PMS → entrega de equipamiento en shelter (17/02/27): habilita el montaje del PMS en las salas y los ensayos PMS.",
 "ABB PMS ↔ Inauco → pruebas de integración PCS↔PMS en Bs.As. (28/04 → 04/05/27).",
 "AESA → layout de instalación + detalle de conexionado a trafo (29/08/26): condiciona el inicio de diseño de los DUCTOS DE BARRAS (entrega 10/06/27, la más tardía).",
 "AESA → entrega de ingeniería Rev.0 a Inauco (10/08/26) y aprobaciones de ingeniería a ABB/HPH: condicionan los freezing points.",
]

NAVY="1F3864"; HEAD="37474F"
thin=Side(style="thin",color="D9D9D9"); B=Border(thin,thin,thin,thin)
EST_FILL={"Cumplido s/plan":"2E7D32","En curso":"EF6C00","Pendiente":"90A4AE"}
def estado(ini,fin):
    if fin < HOY: return "Cumplido s/plan"
    if ini <= HOY <= fin: return "En curso"
    return "Pendiente"
def cell(ws,r,c,v,fill=None,color="000000",bold=False,size=9,center=False,wrap=False):
    cc=ws.cell(r,c,v); cc.font=Font(bold=bold,color=color,size=size)
    if fill: cc.fill=PatternFill("solid",fgColor=fill)
    cc.alignment=Alignment(horizontal="center" if center else "left",vertical="center",wrap_text=wrap)
    cc.border=B; return cc

# ordenar cronológicamente y renumerar
ORD=[]
for (name,ven,col,ms) in PROV:
    ms2=sorted(ms,key=lambda m:(m[1],m[2]))
    ORD.append((name,ven,col,[(f"H{i+1:02d}",)+m for i,m in enumerate(ms2)]))
PROV=ORD

def month_list(a,b):
    out=[];y,m=a.year,a.month
    while (y<b.year) or (y==b.year and m<=b.month):
        out.append((y,m)); m+=1
        if m>12: m=1;y+=1
    return out
MESES=month_list(D(2026,5,1),D(2027,6,1))
M3=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

wb=openpyxl.Workbook()

# ---------------- Hoja 1: Resumen Ejecutivo ----------------
ws=wb.active; ws.title="Resumen Ejecutivo"
ws.merge_cells("A1:K1"); cell(ws,1,1,f"CRONOGRAMA EJECUTIVO — PROVISIONES CRÍTICAS CPF-2  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,13,True); ws.row_dimensions[1].height=24
ws.merge_cells("A2:K2"); cell(ws,2,1,"Seguimiento para PPSA · Corte de alcance: hasta la ENTREGA DE EQUIPOS / del alcance (excluye montaje en planta, SAT, precom/comisionado y PEM)",None,"555555",False,9)
ws.merge_cells("A3:K3"); cell(ws,3,1,f"Fecha de corte del estado: {HOY.strftime('%d/%m/%Y')}",None,"C62828",True,9)
hdr=["Proveedor","Provisión","Inicio contrato","Entrega","Hitos","Cumplidos s/plan","En curso","Pendientes","% avance plan","Próximo hito","Fecha próx. hito"]
for j,h in enumerate(hdr): cell(ws,5,1+j,h,HEAD,"FFFFFF",True,9,True,True)
ws.row_dimensions[5].height=30
r=6
for (name,ven,col,ms) in PROV:
    tot=len(ms); cum=sum(1 for m in ms if estado(m[2],m[3])=="Cumplido s/plan")
    enc=sum(1 for m in ms if estado(m[2],m[3])=="En curso"); pen=tot-cum-enc
    entrega=max(m[3] for m in ms); prox=next((m for m in ms if m[3]>=HOY), None)
    cell(ws,r,1,ven,col,"FFFFFF",True,9,True)
    cell(ws,r,2,name,None,"000000",True,9,False,True)
    cell(ws,r,3,min(m[2] for m in ms).strftime("%d/%m/%Y"),None,"000000",False,9,True)
    cell(ws,r,4,entrega.strftime("%d/%m/%Y"),"1F3864","FFFFFF",True,9,True)
    cell(ws,r,5,tot,None,"000000",False,9,True)
    cell(ws,r,6,cum,"2E7D32","FFFFFF",True,9,True)
    cell(ws,r,7,enc,"EF6C00","FFFFFF",True,9,True)
    cell(ws,r,8,pen,"90A4AE","FFFFFF",True,9,True)
    cell(ws,r,9,f"{100*cum/tot:.0f}%",None,"000000",True,10,True)
    cell(ws,r,10,(prox[1] if prox else "—"),None,"000000",False,8.5,False,True)
    cell(ws,r,11,(prox[3].strftime("%d/%m/%Y") if prox else "—"),None,"C62828",True,9,True)
    ws.row_dimensions[r].height=28; r+=1
for c,w in zip("ABCDEFGHIJK",[9,46,14,13,7,15,10,12,13,44,15]): ws.column_dimensions[c].width=w
r+=1
cell(ws,r,1,"Interfases clave entre provisiones",None,NAVY,True,11); r+=1
for t in INTERFASES:
    cell(ws,r,1,"• "+t,None,"000000",False,9,False,True); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=11); ws.row_dimensions[r].height=16; r+=1

# ---------------- Hoja 2: Cronograma Ejecutivo ----------------
cw=wb.create_sheet("Cronograma Ejecutivo")
cw.merge_cells("A1:J1"); cell(cw,1,1,f"CRONOGRAMA EJECUTIVO POR HITOS  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); cw.row_dimensions[1].height=22
cw.merge_cells("A2:J2"); cell(cw,2,1,"Hitos ordenados cronológicamente. Columnas de seguimiento (Fecha real / % avance / Estado real / Observaciones) para carga periódica.",None,"555555",False,9)
H2=["Nº","Hito ejecutivo","Inicio","Fin","Crítico","Estado (plan)","Fecha real","% avance","Estado real","Observaciones"]
r=4
for (name,ven,col,ms) in PROV:
    cw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=10)
    cell(cw,r,1,f"▶  {name}   ·   {ven}   ·   Entrega: {max(m[3] for m in ms).strftime('%d/%m/%Y')}",col,"FFFFFF",True,10)
    cw.row_dimensions[r].height=18; r+=1
    for j,h in enumerate(H2): cell(cw,r,1+j,h,HEAD,"FFFFFF",True,8.5,True,True)
    cw.row_dimensions[r].height=26; r+=1
    for (cod,hito,ini,fin,crit,obs,acop) in ms:
        est=estado(ini,fin)
        cell(cw,r,1,cod,None,"333333",True,8.5,True)
        cell(cw,r,2,hito,("FFF3E0" if acop else None),"000000",crit==1,9,False,True)
        cell(cw,r,3,ini.strftime("%d/%m/%Y"),None,"1F3864",False,8.5,True)
        cell(cw,r,4,fin.strftime("%d/%m/%Y"),None,"1F3864",crit==1,8.5,True)
        cell(cw,r,5,"SÍ" if crit else "",("C62828" if crit else None),"FFFFFF" if crit else "000000",True,8.5,True)
        cell(cw,r,6,est,EST_FILL[est],"FFFFFF",True,8.5,True)
        for k in (7,8,9): cell(cw,r,k,"",None,"000000",False,9,True)
        cell(cw,r,10,obs,None,"555555",False,8.5,False,True)
        r+=1
    r+=1
for c,w in zip("ABCDEFGHIJ",[6,58,12,12,8,16,12,10,14,44]): cw.column_dimensions[c].width=w
cw.freeze_panes="A4"

# ---------------- Hoja 3: Acopio de Materiales ----------------
aw=wb.create_sheet("Acopio de Materiales")
aw.merge_cells("A1:I1"); cell(aw,1,1,f"ACOPIO DE MATERIALES — seguimiento consolidado  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); aw.row_dimensions[1].height=22
aw.merge_cells("A2:I2"); cell(aw,2,1,"Detalle de los paquetes de acopio por provisión. El grado de avance del acopio es el mejor indicador temprano de cumplimiento de la entrega.",None,"555555",False,9)
AH=["Proveedor","Provisión","Paquete de acopio","Inicio","Fin","Crítico","Estado (plan)","% acopiado (real)","Observaciones"]
for j,h in enumerate(AH): cell(aw,4,1+j,h,HEAD,"FFFFFF",True,9,True,True)
aw.row_dimensions[4].height=28
r=5
for (name,ven,col,ms) in PROV:
    for (cod,hito,ini,fin,crit,obs,acop) in ms:
        if not acop: continue
        est=estado(ini,fin)
        cell(aw,r,1,ven,col,"FFFFFF",True,9,True)
        cell(aw,r,2,name,None,"000000",False,8.5,False,True)
        paq=hito[len("ACOPIO — "):] if hito.startswith("ACOPIO — ") else hito
        cell(aw,r,3,paq,None,"000000",crit==1,9,False,True)
        cell(aw,r,4,ini.strftime("%d/%m/%Y"),None,"1F3864",False,8.5,True)
        cell(aw,r,5,fin.strftime("%d/%m/%Y"),None,"1F3864",crit==1,8.5,True)
        cell(aw,r,6,"SÍ" if crit else "",("C62828" if crit else None),"FFFFFF" if crit else "000000",True,8.5,True)
        cell(aw,r,7,est,EST_FILL[est],"FFFFFF",True,8.5,True)
        cell(aw,r,8,"",None,"000000",False,9,True)
        cell(aw,r,9,obs,None,"555555",False,8.5,False,True)
        r+=1
for c,w in zip("ABCDEFGHI",[9,44,46,12,12,8,16,15,34]): aw.column_dimensions[c].width=w
aw.freeze_panes="A5"
r+=1
cell(aw,r,1,"Nota: 'Fin' del paquete = fecha planificada de acopio completo en fábrica. La columna '% acopiado (real)' se completa en cada reporte de seguimiento.",None,"555555",False,9)

# ---------------- Hoja 4: Timeline ----------------
tw=wb.create_sheet("Timeline (Gantt hitos)")
tw.merge_cells(start_row=1,start_column=1,end_row=1,end_column=3+len(MESES))
cell(tw,1,1,f"TIMELINE DE HITOS EJECUTIVOS  ·  {REV} · {FECHA}  ·  corte: entrega de equipos",NAVY,"FFFFFF",True,12,True); tw.row_dimensions[1].height=22
r=3
cell(tw,r,1,"Provisión / Hito",HEAD,"FFFFFF",True,8.5,False,True); cell(tw,r,2,"Inicio",HEAD,"FFFFFF",True,8.5,True); cell(tw,r,3,"Fin",HEAD,"FFFFFF",True,8.5,True)
for j,(y,m) in enumerate(MESES):
    cell(tw,r,4+j,f"{M3[m]}-{str(y)[2:]}",HEAD,"FFFFFF",True,7.5,True,True)
    tw.column_dimensions[get_column_letter(4+j)].width=6.5
tw.row_dimensions[r].height=26; r+=1
def midx(d):
    for j,(y,m) in enumerate(MESES):
        if d.year==y and d.month==m: return j
    return None
for (name,ven,col,ms) in PROV:
    tw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=3+len(MESES))
    cell(tw,r,1,f"▶  {name}",col,"FFFFFF",True,9.5); tw.row_dimensions[r].height=16; r+=1
    for (cod,hito,ini,fin,crit,obs,acop) in ms:
        cell(tw,r,1,f"{cod} · {hito}",("FFF3E0" if acop else None),"000000",crit==1,8.3,False,True)
        cell(tw,r,2,ini.strftime("%d/%m/%y"),None,"1F3864",False,8,True)
        cell(tw,r,3,fin.strftime("%d/%m/%y"),None,"1F3864",crit==1,8,True)
        a,b=midx(ini),midx(fin)
        for j in range(len(MESES)):
            v=""; f=None
            if a is not None and b is not None and a<=j<=b:
                f = "C62828" if crit else col
                if a==b or j==b: v="◆" if crit else "■"
            cell(tw,r,4+j,v,f,"FFFFFF",True,8,True)
        r+=1
    r+=1
for c,w in zip("ABC",[60,10,10]): tw.column_dimensions[c].width=w
tw.freeze_panes="D4"

# ---------------- Hoja 5: Notas ----------------
nw=wb.create_sheet("Notas y criterios")
cell(nw,1,1,"NOTAS Y CRITERIOS",NAVY,"FFFFFF",True,12)
for i,n in enumerate(NOTAS):
    cell(nw,3+i,1,"• "+n,None,("C62828" if n.startswith("ALERTA") else "000000"),n.startswith("ALERTA"),9,False,True)
    nw.row_dimensions[3+i].height=32
nw.column_dimensions['A'].width=165

out=os.path.join(GEN,f"Cronograma Ejecutivo Provisiones Criticas - {REV} - {FECHA}.xlsx")
wb.save(out)
print("OK ->",out)
for (name,ven,col,ms) in PROV:
    cum=sum(1 for m in ms if estado(m[2],m[3])=="Cumplido s/plan")
    ac=sum(1 for m in ms if m[6])
    print(f"  {name[:52]:52} hitos={len(ms):2} acopio={ac} cumpl={cum:2} entrega={max(m[3] for m in ms).strftime('%d/%m/%Y')}")
