#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CRONOGRAMA EJECUTIVO — Provisiones críticas CPF-2 (Rev5, 2026-09-23).
Cambios respecto de Rev4:
  1) ABB: se adopta la actualización del cronograma 5155-00-0009-VE-CO-PG-001_A del
     21-09-2026 (reemplaza la del 21-08-2026). 127 diferencias respecto de la versión
     anterior, detectadas comparando ambos PDF tarea por tarea. Lo esencial:
       · DUCTOS DE BARRAS: se atrasan 4 semanas. La entrega en sitio pasa del 10-06-27
         al 08-07-27 y se convierte en el hito más tardío de TODO el paquete. La causa es
         el corrimiento del insumo de AESA (layout de instalación + detalle de conexionado
         a trafo) del 29-08 al 18-09-26, que arrastró diseño, fabricación y traslado.
       · SALAS SE#3 y SE#4: los FREEZING POINTS de ingeniería se corren 2 semanas
         (básica 11-09 → 25-09-26 · detalle 05-10 → 13-10-26).
       · SE#3 (BT): la entrega EXW se ADELANTA del 23-04-27 al 20-04-27, con la FAT de
         sala del 23 al 30-03-27. La fabricación del shelter se atrasa 3 semanas.
       · SE#4 (MT): mantiene la entrega EXW del 25-03-27; la FAT en fábrica de origen de
         las celdas MT se corre al 26-11 / 03-12-26.
       · PMS: sin cambios (entrega en shelter 17-02-27).
  2) Se agregan los hitos de FABRICACIÓN DEL SHELTER en ambas salas y de APROBACIÓN DE
     DISEÑO en los ductos de barras, que pasaron a ser determinantes.
  3) PUNTOS DE ATENCIÓN: se suman PA-07 (atraso de los ductos de barras) y PA-08
     (freezing points de las salas corridos, el de ingeniería básica vence el 25-09-26).
Cambios de la Rev4 respecto de Rev3:
  1) INAUCO: se incorpora el cronograma de fabricación 5155-00-0029-VE-CA-PG-002 Rev.0
     (14-09-2026), con el detalle por tablero (CPF2-PCS-SC-07a y 07b, CPF-COM-07 y
     CPF-COM-SE3), sus PreFAT internos y liberación de pendientes, y las aclaraciones
     del pie del cronograma: Pre iFAT 10-02 → 26-03-27, FAT con el cliente 29-03 →
     22-04-27 y LIBERACIÓN DE TABLEROS el 04-05-27 (entrega en camión de AESA), que
     adelanta la entrega de Inauco respecto de la Rev3 (17-05-27).
  2) HARG (HIMA Argentina): al 23-09-2026 NO se colocó la orden de compra, por lo que el
     inicio de los trabajos se corre a principios de noviembre (02-11-2026). Se aplica un
     corrimiento de 49 días corridos a todas las actividades propias de HARG; los hitos
     que son insumo de HPH (ya contratada) NO se mueven.
  3) CAMINO CRÍTICO: recalculado. El cierre ya no lo marca Inauco sino HARG (10-05-27).
  4) PUNTOS DE ATENCIÓN: PA-01 y PA-02 se reformulan a la luz del corrimiento y se suman
     PA-04 (OC de HARG sin colocar) y PA-05 (superposición de la adecuación con la Pre
     iFAT de Inauco y cierre de HARG posterior a la liberación de tableros).
Corte de alcance: hasta la ENTREGA DE EQUIPOS / del alcance de cada provisión.
"""
import os
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
REV="Rev5"; FECHA="2026-09-23"; HOY=date(2026,9,23)
SHIFT_HARG=49   # días corridos de corrimiento por OC no colocada (14-09-26 -> 02-11-26)
D=lambda y,m,d: date(y,m,d)
from datetime import timedelta
SH=lambda d: d+timedelta(SHIFT_HARG)   # corrimiento de las actividades propias de HARG

# (hito, inicio, fin, crítico, observación, es_acopio)
PROV=[
 ("ABB — Sala Eléctrica SE#3 (BT)","ABB","C62828",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"OC recibida en ABB",0),
  ("Kick-off Meeting (KOM) técnico",D(2026,5,15),D(2026,5,18),0,"",0),
  ("Confirmación de layout de sala y equipamiento",D(2026,5,18),D(2026,5,20),0,"Insumo AESA",0),
  ("Emisión de ingeniería básica",D(2026,5,20),D(2026,9,11),0,"",0),
  ("Aprobación de ingeniería básica",D(2026,9,11),D(2026,9,25),0,"Aprobación AESA — 90% al corte",0),
  ("FREEZING POINT — ingeniería básica",D(2026,9,25),D(2026,9,25),1,"CORRIDO 2 semanas (antes 11-09-26)",0),
  ("Emisión de ingeniería de detalle",D(2026,5,22),D(2026,9,29),0,"95% al corte",0),
  ("Aprobación de ingeniería de detalle",D(2026,9,29),D(2026,10,13),0,"Aprobación AESA",0),
  ("FREEZING POINT — ingeniería de detalle",D(2026,10,13),D(2026,10,13),1,"CORRIDO 8 días (antes 05-10-26)",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,19),D(2026,5,19),0,"Compra y acopio sala BT",1),
  ("ACOPIO — Baterías / rectificadores / UPS",D(2026,6,9),D(2026,10,12),0,"Paquete de CC — 85%",1),
  ("ACOPIO — Magnéticos, semiconductores, interruptores y fusibles",D(2026,7,13),D(2026,10,5),0,"Componentes de tableros — 85%",1),
  ("ACOPIO — Gabinetes y accesorios",D(2026,7,13),D(2026,10,12),0,"80%",1),
  ("ACOPIO — Tableros auxiliares",D(2026,7,22),D(2026,10,30),0,"62% — adelantado (antes 09-11-26)",1),
  ("ACOPIO — Materiales de los CCM (críticos y secundarios)",D(2026,6,19),D(2026,10,16),0,"84% — críticos al 02-10-26",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,12,10),D(2026,12,10),1,"Cierre de acopio de la sala BT",1),
  ("Inicio de fabricación en fábrica de origen (Brasil)",D(2026,10,6),D(2026,10,6),0,"ADELANTADO (antes 12-10-26) — M&B Brasil",0),
  ("Fabricación de tableros en fábrica de origen (Brasil)",D(2026,10,6),D(2026,12,10),1,"CCM-005/006/007 06/10-20/11 · auxiliares 30/10-10/12 · rectif./UPS 12-21/10",0),
  ("Ensayos internos en fábrica de origen (Brasil)",D(2026,10,22),D(2026,12,1),0,"Previos a la FAT",0),
  ("FAT EN FÁBRICA DE ORIGEN (Brasil) — tableros y CCM",D(2026,11,9),D(2026,12,14),1,"Witness AESA — rectif./UPS 09-11/11 · CCM 20/11-14/12 · aux. 10-14/12",0),
  ("Adecuación y despacho de fábrica de origen",D(2026,11,30),D(2026,12,18),1,"ADELANTADO (antes 09-12-26) — condición EXW en origen",0),
  ("Entrega en Shelterista (Mendoza)",D(2026,11,23),D(2027,1,11),1,"Arribo escalonado: rectif./UPS 23-26/11 · aux. 17-22/12 · drives 05/01 · CCM 24/12-11/01",0),
  ("Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)",D(2026,11,23),D(2027,3,5),1,"ATRASADO 3 semanas (antes 02-11-26 → 26-02-27)",0),
  ("Montaje e integración de equipos en sala",D(2027,2,26),D(2027,3,17),0,"Shelterista — atrasado 1 semana",0),
  ("Interconexión de equipos",D(2027,3,9),D(2027,3,23),0,"",0),
  ("FAT SALA ELÉCTRICA (shelter completo)",D(2027,3,23),D(2027,3,30),1,"Witness AESA — adelantada 2 días",0),
  ("Ensayos PMS en sala (INTERFASE — contrato PMS)",D(2027,3,30),D(2027,4,6),0,"Alcance del contrato PMS",0),
  ("Desconexión, embalaje y despacho",D(2027,4,6),D(2027,4,20),1,"ADELANTADO (antes 07-04 → 23-04-27)",0),
  ("DESPACHO / ENTREGA EXW",D(2027,4,20),D(2027,4,20),1,"HITO FINAL — ADELANTADO 3 días (antes 23-04-27)",0),
 ]),
 ("ABB — Sala Eléctrica SE#4 (MT)","ABB","AD1457",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"OC recibida en ABB",0),
  ("Kick-off Meeting (KOM) técnico",D(2026,5,15),D(2026,5,18),0,"",0),
  ("Confirmación de layout de sala y equipamiento",D(2026,5,18),D(2026,5,20),0,"Insumo AESA",0),
  ("Emisión de ingeniería básica",D(2026,5,20),D(2026,9,11),0,"",0),
  ("Aprobación de ingeniería básica",D(2026,9,11),D(2026,9,25),0,"Aprobación AESA — 90% al corte",0),
  ("FREEZING POINT — ingeniería básica",D(2026,9,25),D(2026,9,25),1,"CORRIDO 2 semanas (antes 11-09-26)",0),
  ("Emisión de ingeniería de detalle",D(2026,5,22),D(2026,9,29),0,"95% al corte",0),
  ("Aprobación de ingeniería de detalle",D(2026,9,29),D(2026,10,13),0,"Aprobación AESA",0),
  ("FREEZING POINT — ingeniería de detalle",D(2026,10,13),D(2026,10,13),1,"CORRIDO 8 días (antes 05-10-26)",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,19),D(2026,5,19),0,"Compra y acopio sala MT",1),
  ("ACOPIO — Celdas MT (materiales principales, fábrica de origen)",D(2026,6,22),D(2026,10,30),1,"70% — equipo de largo plazo",1),
  ("ACOPIO — Variadores / drives MT",D(2026,8,18),D(2026,10,9),0,"70% — adelantado (antes 20-10-26)",1),
  ("ACOPIO — Tableros auxiliares",D(2026,7,22),D(2026,10,30),0,"62% — adelantado (antes 09-11-26)",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,12,10),D(2026,12,10),1,"Cierre de acopio de la sala MT",1),
  ("Inicio de fabricación en fábrica de origen (Brasil)",D(2026,10,16),D(2026,10,16),0,"Celdas MT / auxiliares — M&B Brasil",0),
  ("Fabricación de tableros y celdas en fábrica de origen (Brasil)",D(2026,10,16),D(2026,12,10),1,"Celdas MT 16/10-13/11 · auxiliares 30/10-10/12",0),
  ("Ensayos internos en fábrica de origen (Brasil)",D(2026,11,16),D(2026,11,25),0,"Previos a la FAT",0),
  ("FAT EN FÁBRICA DE ORIGEN (Brasil) — tableros y celdas MT",D(2026,11,26),D(2026,12,14),1,"Witness AESA — celdas MT 26/11-03/12 · auxiliares 10-14/12",0),
  ("Despacho de fábrica de origen (celdas MT y tableros)",D(2026,12,3),D(2026,12,17),1,"Acond./embalaje 03-09/12 · FCA 09-11/12",0),
  ("Entrega en Shelterista (Mendoza)",D(2026,12,17),D(2027,1,6),1,"Arribo escalonado: aux. 17-22/12 · celdas 04-05/01 · drives 06/01",0),
  ("Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)",D(2026,10,6),D(2027,1,29),1,"Atrasado 8 días (antes 05-10-26 → 21-01-27)",0),
  ("Montaje e integración de equipos en sala",D(2027,1,29),D(2027,2,19),0,"Shelterista — adelantado 3 días",0),
  ("Interconexión de equipos",D(2027,2,17),D(2027,2,23),0,"",0),
  ("FAT SALA ELÉCTRICA (shelter completo)",D(2027,2,25),D(2027,3,1),1,"Witness AESA",0),
  ("Ensayos PMS en sala (INTERFASE — contrato PMS)",D(2027,3,3),D(2027,3,9),0,"Alcance del contrato PMS",0),
  ("Desconexión, embalaje y despacho",D(2027,3,10),D(2027,3,25),1,"",0),
  ("DESPACHO / ENTREGA EXW",D(2027,3,25),D(2027,3,25),1,"HITO FINAL — sin cambio respecto de la Rev4",0),
 ]),
 ("ABB — Ductos de barras","ABB","795548",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"Mismo contrato ABB (alcance separado)",0),
  ("Emisión layout instalación + detalle conexionado trafo",D(2026,7,20),D(2026,9,18),1,"INSUMO AESA — se corrió del 29-08 al 18-09-26 y arrastró todo el bloque",0),
  ("FREEZING POINT — inicio de diseño",D(2026,9,18),D(2026,9,18),1,"CORRIDO 20 días (antes 29-08-26)",0),
  ("Diseño de ductos de barras",D(2026,9,18),D(2026,12,14),0,"5% al corte — antes 31-08 → 20-11-26",0),
  ("Aprobación de diseño",D(2026,12,14),D(2026,12,30),0,"Aprobación AESA",0),
  ("FREEZING POINT — cierre de diseño",D(2026,12,30),D(2026,12,30),1,"CORRIDO 26 días (antes 04-12-26) — habilita fabricación",0),
  ("ACOPIO / fabricación de ductos de barras",D(2026,12,30),D(2027,4,9),1,"100 días corridos (acopio integrado a fabricación)",1),
  ("FCA fábrica de origen",D(2027,4,9),D(2027,4,16),1,"Puesta a disposición en origen",0),
  ("Traslado internacional + aduana",D(2027,4,16),D(2027,7,5),1,"80 días corridos — tramo más largo",0),
  ("ENTREGA DE EQUIPOS EN SITIO",D(2027,7,5),D(2027,7,8),1,"HITO FINAL — ATRASADO 4 SEMANAS (antes 10-06-27); es la entrega más tardía de todo el paquete",0),
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
  ("Pruebas de integración PCS↔PMS (Bs.As.) — ANTICIPADA",D(2026,10,1),D(2026,11,30),1,"Ventana OCT-NOV 2026 · FECHA A CONVENIR — interfase con Inauco (PCS)",0),
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
  ("Pruebas de integración PCS↔PMS (Bs.As.) — ANTICIPADA",D(2026,10,1),D(2026,11,30),1,"Ventana OCT-NOV 2026 · FECHA A CONVENIR — interfase con ABB PMS",0),
  ("Aprobación ing. p/ construcción — FREEZING POINT",D(2026,11,19),D(2026,11,19),1,"Requisito 3",0),
  ("ACOPIO — Recepción de hardware (Rockwell / Stratix)",D(2026,11,26),D(2026,11,26),1,"Requisito 6 — cierre de acopio",1),
  ("INICIO DE ARMADO DE TABLEROS — estudio de ingeniería y recepción de materiales (taller)",D(2026,11,27),D(2026,11,30),1,"Crono. fabricación PG-002 Rev.0 — 50 días laborables de fabricación",0),
  ("Fabricación tablero CPF2-PCS-SC-07a (PCS + marshalling · 3 gabinetes)",D(2026,12,1),D(2027,1,19),1,"30 días — adecuación de gabinetes, montaje de elementos y tendido",0),
  ("Fabricación tablero CPF2-PCS-SC-07b (PCS + marshalling · 2 gabinetes)",D(2026,12,1),D(2027,1,5),1,"20 días",0),
  ("PreFAT interno y liberación de pendientes — tablero 07b",D(2027,1,6),D(2027,1,15),0,"Pruebas y verificación interna de taller (3 d) + pendientes (5 d)",0),
  ("PreFAT interno y liberación de pendientes — tablero 07a",D(2027,1,20),D(2027,1,29),0,"Pruebas y verificación interna de taller (3 d) + pendientes (5 d)",0),
  ("Fabricación racks de comunicaciones CPF-COM-07 y CPF-COM-SE3",D(2027,1,22),D(2027,2,1),0,"Racks 19\" 40U · la fabricación del CPF-COM-SE3 está A CONFIRMAR",0),
  ("PreFAT y liberación de pendientes — racks de comunicaciones",D(2027,2,2),D(2027,2,3),0,"",0),
  ("Recepción tableros HIMA en base Inauco (insumo HPH)",D(2027,1,25),D(2027,1,25),1,"Requisito 2 — interfase HPH/HARG",0),
  ("Configuración PLC — Etapa 2",D(2027,1,25),D(2027,3,23),0,"",0),
  ("FIN DE ARMADO DE TABLEROS",D(2027,2,16),D(2027,2,16),1,"Cierre de armado s/ crono. de fabricación PG-002 Rev.0",0),
  ("PRE iFAT (integración interna Inauco)",D(2027,2,10),D(2027,3,26),1,"Aclaración del crono. PG-002 Rev.0",0),
  ("Aprobación de procedimientos FAT",D(2027,2,24),D(2027,2,24),0,"Requisito 4",0),
  ("FAT CON EL CLIENTE — PCS + Comunicaciones",D(2027,3,29),D(2027,4,22),1,"Ventana confirmada por el crono. PG-002 Rev.0 · witness AESA/PPSA",0),
  ("FAT SIS (conjunto con HARG/HPH)",D(2027,3,29),D(2027,4,26),1,"Requiere presencia HIMA",0),
  ("Configuraciones post-FAT",D(2027,4,23),D(2027,5,4),0,"",0),
  ("LIBERACIÓN / ENTREGA DE TABLEROS",D(2027,5,4),D(2027,5,4),1,"HITO FINAL — entrega en camión de AESA (Rev3: 17-05-27)",0),
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
  ("Nacionalización / despacho de aduana + transporte a base Inauco",D(2026,12,15),D(2027,1,25),1,"Mínimo 3-4 semanas + receso de fin de año",0),
  ("RECEPCIÓN DE TABLEROS EN BASE INAUCO",D(2027,1,25),D(2027,1,25),1,"HITO FINAL — habilita la adecuación de HARG",0),
 ]),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","HARG","00838F",[
  # H = actividad propia de HARG -> corrida SHIFT_HARG días por la OC no colocada
  # I = insumo de HPH (ya contratada) -> NO se mueve
  ("Documentación HARG — Doc. 1 (HW y otros)",SH(D(2026,9,14)),SH(D(2026,9,25)),0,"2 semanas · corrido +49 d (OC)",0),
  ("Programación HARG — Recurso 1",SH(D(2026,9,21)),SH(D(2027,1,29)),1,"11 semanas · corrido +49 d (OC)",0),
  ("Documentación HPH — Doc. 1 (INTERFASE)",D(2026,9,28),D(2026,10,16),0,"Insumo de HPH (3,5 semanas) — NO se corre",0),
  ("Ingeniería de Aplicación SIS y documentación (INTERFASE HPH)",D(2026,11,9),D(2026,12,18),1,"6 semanas — a definir por HPH; NO se corre",0),
  ("Documentación HARG — Doc. 2 (Modbus)",SH(D(2026,11,9)),SH(D(2026,11,13)),0,"1 semana · cae en el receso de fin de año — reprogramar",0),
  ("Fabricación HPH — entrega DIC-2026 (INTERFASE)",D(2026,12,21),D(2026,12,21),1,"Insumo de HPH — NO se corre",0),
  ("Adecuación de tableros marshalling (INTERFASE Inauco)",SH(D(2026,12,21)),SH(D(2027,1,29)),1,"6 semanas en base Inauco · corrido +49 d (OC)",0),
  ("Programación HARG — Recurso 2",SH(D(2027,1,4)),SH(D(2027,1,29)),0,"5 semanas · corrido +49 d (OC)",0),
  ("Documentación HARG — Doc. 3 (Proc. FAT e IFAT)",SH(D(2027,1,25)),SH(D(2027,2,12)),0,"3 semanas · corrido +49 d (OC)",0),
  ("FAT",SH(D(2027,2,24)),SH(D(2027,2,26)),1,"0,5 semana · cae dentro de la ventana de FAT de Inauco",0),
  ("IFAT (Recursos 1 y 2)",SH(D(2027,3,1)),SH(D(2027,3,12)),1,"2 semanas · termina 4 días DESPUÉS del cierre de la FAT SIS",0),
  ("Documentación HARG — Doc. 5 (Proy. Doc.) y HPH Doc. 2",SH(D(2027,3,18)),SH(D(2027,3,22)),0,"Documentación de cierre · corrido +49 d (OC)",0),
  ("ENTREGA DEL ALCANCE HARG (IFAT + documentación)",SH(D(2027,3,12)),SH(D(2027,3,22)),1,"HITO FINAL — cierra DESPUÉS de la liberación de tableros de Inauco",0),
 ]),
]

# ---------------- Puntos de atención (Rev3) ----------------
# clave: (provisión, hito) -> código de PA
PA_MARK={
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Entrega DAP Ezeiza (Argentina)"):"PA-01",
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Nacionalización / despacho de aduana + transporte a base Inauco"):"PA-01",
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","RECEPCIÓN DE TABLEROS EN BASE INAUCO"):"PA-02",
 ("Inauco — PCS / SCADA / Comunicaciones","Recepción tableros HIMA en base Inauco (insumo HPH)"):"PA-02",
 ("Inauco — PCS / SCADA / Comunicaciones","Pruebas de integración PCS↔PMS (Bs.As.) — ANTICIPADA"):"PA-03",
 ("Inauco — PCS / SCADA / Comunicaciones","Fabricación racks de comunicaciones CPF-COM-07 y CPF-COM-SE3"):"PA-06",
 ("ABB — Ductos de barras","Emisión layout instalación + detalle conexionado trafo"):"PA-07",
 ("ABB — Ductos de barras","FREEZING POINT — inicio de diseño"):"PA-07",
 ("ABB — Ductos de barras","FREEZING POINT — cierre de diseño"):"PA-07",
 ("ABB — Ductos de barras","ENTREGA DE EQUIPOS EN SITIO"):"PA-07",
 ("ABB — Sala Eléctrica SE#3 (BT)","FREEZING POINT — ingeniería básica"):"PA-08",
 ("ABB — Sala Eléctrica SE#3 (BT)","FREEZING POINT — ingeniería de detalle"):"PA-08",
 ("ABB — Sala Eléctrica SE#4 (MT)","FREEZING POINT — ingeniería básica"):"PA-08",
 ("ABB — Sala Eléctrica SE#4 (MT)","FREEZING POINT — ingeniería de detalle"):"PA-08",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Documentación HARG — Doc. 1 (HW y otros)"):"PA-04",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Programación HARG — Recurso 1"):"PA-04",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Documentación HARG — Doc. 2 (Modbus)"):"PA-04",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Adecuación de tableros marshalling (INTERFASE Inauco)"):"PA-05",
 ("Inauco — PCS / SCADA / Comunicaciones","PRE iFAT (integración interna Inauco)"):"PA-05",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","IFAT (Recursos 1 y 2)"):"PA-05",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","ENTREGA DEL ALCANCE HARG (IFAT + documentación)"):"PA-05",
 ("Inauco — PCS / SCADA / Comunicaciones","LIBERACIÓN / ENTREGA DE TABLEROS"):"PA-05",
}
PUNTOS=[
 dict(cod="PA-07", tit="Los ductos de barras se atrasan 4 semanas y pasan a ser la entrega más tardía del proyecto",
      inter="AESA → ABB",
      plan="La actualización del cronograma ABB del 21-09-26 corre la entrega en sitio del 10-06-27 al 08-07-27. La causa es el insumo de AESA —layout de instalación + detalle de conexionado a trafo— que se emitió el 18-09-26 en lugar del 29-08-26.",
      riesgo="El corrimiento de 20 días del insumo se amplificó a 28 días en la entrega: freezing point de inicio de diseño 29-08 → 18-09-26, diseño hasta el 14-12-26, freezing point de cierre 04-12 → 30-12-26, fabricación hasta el 09-04-27 y traslado internacional + aduana hasta el 05-07-27. El tramo de traslado (80 días corridos) no admite compresión.",
      impacto="La entrega en sitio del 08-07-27 es ahora el hito más tardío de TODO el paquete de provisiones críticas, por delante del cierre del alcance HARG (10-05-27). Condiciona el montaje de la vinculación a transformadores y, con ello, la energización.",
      real="08-07-2027, sin holgura declarada. Cada día de demora adicional en la aprobación de diseño (14-12 → 30-12-26) se traslada íntegro a la entrega.",
      accion="Proteger las dos aprobaciones que quedan: el diseño de ductos cierra el 14-12-26 y la aprobación de AESA debe estar el 30-12-26, en pleno receso. Asignar el revisor con anticipación. Evaluar con ABB si el traslado internacional admite una alternativa más rápida para recuperar parte de las 4 semanas.",
      sev="ALTA"),
 dict(cod="PA-08", tit="Freezing points de ingeniería de las salas corridos dos semanas — el de ingeniería básica vence el 25-09-26",
      inter="AESA → ABB (Salas SE#3 y SE#4)",
      plan="En ambas salas, el freezing point de ingeniería básica pasa del 11-09 al 25-09-26 y el de ingeniería de detalle del 05-10 al 13-10-26. Al corte, la aprobación de la ingeniería básica está al 90% y la emisión de la de detalle al 95%.",
      riesgo="El freezing point de ingeniería básica vence dos días después de la fecha de corte de este informe. Un freezing point no cumplido no corre solo: congela fabricación, acopio y slots de fábrica en Brasil que ya están comprometidos.",
      impacto="Pese al corrimiento, ABB mantiene —y hasta adelanta— las entregas de las salas (SE#3 pasa del 23-04 al 20-04-27; SE#4 se mantiene el 25-03-27), lo que significa que absorbió el atraso comprimiendo su propio plan. Ese margen ya se consumió: un nuevo corrimiento sí impactaría las entregas.",
      real="25-09-2026 (ingeniería básica) y 13-10-2026 (ingeniería de detalle), ambos en ambas salas.",
      accion="Cerrar la aprobación de la ingeniería básica dentro de la semana y asegurar el circuito de revisión de la de detalle para el 13-10-26. Comunicar a ABB que no hay margen para un tercer corrimiento.",
      sev="ALTA"),
 dict(cod="PA-04", tit="Orden de compra de HIMA Argentina (HARG) no colocada al 23-09-2026",
      inter="AESA → HARG",
      plan="El cronograma PSCH_CPF2_Rev3 arrancaba los trabajos de HARG el 14-09-2026. A la fecha de corte la OC no está colocada, por lo que el inicio se corre a principios de noviembre (02-11-2026).",
      riesgo="Corrimiento de 49 días corridos (7 semanas) sobre TODAS las actividades propias de HARG: documentación, programación de los dos recursos, adecuación de los tableros de marshalling, FAT e IFAT. Cada semana adicional sin OC se traslada íntegra al cierre del alcance.",
      impacto="El cierre del alcance HARG pasa del 22-03-27 al 10-05-27 y se convierte en el nuevo FIN DEL CAMINO CRÍTICO, por delante de la liberación de tableros de Inauco (04-05-27). La FAT de HARG (14-16/04/27) y el IFAT (19-30/04/27) se corren hacia la ventana de FAT del cliente.",
      real="Inicio realista de HARG: 02-11-2026, sólo si la OC se coloca dentro de octubre. Cada semana de demora adicional corre el cierre una semana más.",
      accion="Acelerar la colocación de la OC de HARG como prioridad 1. Evaluar con HIMA una orden de inicio anticipada o carta de intención para largar la ingeniería de aplicación y la documentación mientras se formaliza la OC; los insumos de HPH ya estarán disponibles, de modo que parte del corrimiento es recuperable resecuenciando.",
      sev="ALTA"),
 dict(cod="PA-05", tit="La adecuación de HARG cae dentro de la Pre iFAT de Inauco y su cierre queda después de la liberación de tableros",
      inter="HARG ↔ Inauco",
      plan="Con el corrimiento, HARG adecúa los tableros de marshalling del 08-02 al 19-03-27, en la base de Inauco. El cronograma PG-002 Rev.0 de Inauco fija la Pre iFAT del 10-02 al 26-03-27 y la liberación de tableros el 04-05-27. El IFAT de HARG termina el 30-04-27 y su alcance cierra el 10-05-27.",
      riesgo="Tres solapamientos: (a) la adecuación de HARG ocupa físicamente los tableros durante toda la Pre iFAT de Inauco; (b) el IFAT de HARG (19-30/04/27) termina 4 días después del cierre de la FAT SIS conjunta (26-04-27); (c) el cierre del alcance HARG (10-05-27) es posterior a la liberación de los tableros de Inauco (04-05-27).",
      impacto="Inauco liberaría los tableros antes de que HIMA cierre su IFAT y su documentación, con lo que el sistema de seguridad viajaría a obra sin el alcance HARG terminado. Además compiten por el mismo tablero y el mismo espacio de taller durante seis semanas.",
      real="Se requiere una secuencia acordada entre ambos: adecuación HARG terminada antes del inicio de la Pre iFAT, o Pre iFAT desdoblada (PCS primero, SIS después).",
      accion="Convocar a Inauco y HIMA a acordar la secuencia y el uso del taller; evaluar adelantar la adecuación de HARG (depende de PA-04) o correr la liberación de tableros al cierre real del alcance HARG. Definir si la FAT SIS conjunta se extiende para cubrir el IFAT de HARG.",
      sev="ALTA"),
 dict(cod="PA-01", tit="Nacionalización de los tableros de marshalling entregados por HPH",
      inter="HPH (Alemania) → HARG / Inauco",
      plan="HPH entrega DAP Ezeiza el 15-12-26 y la recepción en base Inauco está prevista para el 25-01-27. Con el corrimiento de HARG, la adecuación ya no arranca el 21-12-26 sino el 08-02-27.",
      riesgo="El despacho de aduana de un embarque de tableros importados demanda como mínimo 3 a 4 semanas y el arribo cae en el receso de fin de año. El riesgo sobre el plazo de nacionalización no cambió; lo que cambió es que ahora hay margen aguas abajo.",
      impacto="Severidad REBAJADA de ALTA a MEDIA respecto de la Rev3: entre la recepción prevista (25-01-27) y el inicio de la adecuación de HARG (08-02-27) quedan 2 semanas de holgura, que antes no existían. Un atraso de nacionalización mayor a 2 semanas vuelve a impactar directamente.",
      real="Recepción en base Inauco entre el 25-01-27 y la primera semana de febrero de 2027.",
      accion="Mantener el seguimiento semanal del hito de nacionalización; designar despachante y anticipar la documentación de embarque con HPH. La holgura ganada NO debe consumirse para absorber otros atrasos.",
      sev="MEDIA"),
 dict(cod="PA-02", tit="Inauco recibe los tableros de HIMA sin adecuar, contra lo previsto en su Requisito 2",
      inter="HPH ↔ HARG ↔ Inauco",
      plan="Inauco fija en su Requisito 2 la recepción de los tableros HIMA en su base el 25-01-27, asumiendo el equipo prácticamente terminado. Con el corrimiento, HARG recién adecúa entre el 08-02 y el 19-03-27.",
      riesgo="La premisa de Inauco de recibir el tablero «casi terminado» no se cumple: recibe el tablero de HPH sin adecuar y la adecuación se ejecuta después, en su propia base.",
      impacto="Inauco debe prever espacio de taller, energía y acceso para que HIMA trabaje sobre los tableros durante seis semanas, en paralelo con su propia Pre iFAT (ver PA-05). La Configuración PLC Etapa 2 (25-01 → 23-03-27) convive con la adecuación.",
      real="Tablero disponible ya adecuado: 19-03-2027.",
      accion="Reconfirmar con Inauco el alcance real del Requisito 2 y acordar las condiciones de trabajo de HIMA en su base. Evaluar, con HPH, si parte de la adecuación puede ejecutarse en origen antes del despacho.",
      sev="ALTA"),
 dict(cod="PA-03", tit="Prueba de integración PCS↔PMS en Buenos Aires — fecha a convenir",
      inter="Inauco (PCS) ↔ ABB (PMS)",
      plan="Anticipada respecto de la Rev2: ventana OCT-NOV 2026, fecha exacta sin acordar.",
      riesgo="La fecha dentro de la ventana aún no está acordada entre AESA, Inauco y ABB, y requiere disponibilidad simultánea de la configuración PMS y del PCS. La ventana se está agotando.",
      impacto="Anticiparla es favorable: detecta tempranamente incompatibilidades de comunicación (IEC61850 / Modbus / Profibus) antes de las FAT. Sin fecha acordada el ítem no puede planificarse ni asignarse recursos, y si se pierde la ventana vuelve a caer después de las FAT.",
      real="Ventana OCT-NOV 2026; se sugiere fijarla luego del cierre del acopio de switches de red del PMS (05-11-26).",
      accion="Convocar la reunión de coordinación AESA / Inauco / ABB para fijar la fecha exacta y el alcance de la prueba, y emitir el protocolo. Es el punto de definición más próximo en el tiempo.",
      sev="MEDIA"),
 dict(cod="PA-06", tit="Fabricación del rack de comunicaciones CPF-COM-SE3 a confirmar",
      inter="AESA → Inauco",
      plan="El cronograma PG-002 Rev.0 lista el rack CPF-COM-SE3 (19\" 40U) con fechas de fabricación 22-01 → 03-02-27, pero marcado «A Confirmar Fabricación».",
      riesgo="Si la confirmación llega tarde, el rack no entra en la ventana de fabricación de 50 días laborables y queda fuera del armado que cierra el 16-02-27.",
      impacto="El rack de comunicaciones de la Sala Eléctrica 3 quedaría fuera del paquete de la FAT y de la liberación del 04-05-27, y habría que tratarlo como una provisión separada.",
      real="Definición requerida antes de la 2.ª semana de enero de 2027 (recepción de materiales 25-01-27).",
      accion="Cerrar con ingeniería y con Inauco el alcance del CPF-COM-SE3 y emitir la confirmación de fabricación antes de fin de diciembre de 2026.",
      sev="MEDIA"),
]

# ---------------- Camino crítico (Rev3) ----------------
# hitos marcados como ruta crítica en el cronograma por hitos
CP_MARK={
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Inicio de contrato / proyecto"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Ingeniería de detalle / especificación"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Fin de ingeniería / CAD — FREEZING POINT"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Fabricación de tableros de marshalling"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Hardware fabricado y probado (test de fábrica)"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Preparación de despacho (packing)"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Listo para despacho"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Despacho / shipping"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Entrega DAP Ezeiza (Argentina)"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Nacionalización / despacho de aduana + transporte a base Inauco"),
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","RECEPCIÓN DE TABLEROS EN BASE INAUCO"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Documentación HARG — Doc. 1 (HW y otros)"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Programación HARG — Recurso 1"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Adecuación de tableros marshalling (INTERFASE Inauco)"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","FAT"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","IFAT (Recursos 1 y 2)"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","ENTREGA DEL ALCANCE HARG (IFAT + documentación)"),
 ("Inauco — PCS / SCADA / Comunicaciones","Recepción tableros HIMA en base Inauco (insumo HPH)"),
 ("Inauco — PCS / SCADA / Comunicaciones","Configuración PLC — Etapa 2"),
 ("Inauco — PCS / SCADA / Comunicaciones","FAT SIS (conjunto con HARG/HPH)"),
 ("Inauco — PCS / SCADA / Comunicaciones","LIBERACIÓN / ENTREGA DE TABLEROS"),
}
# cadena crítica explícita: (Nº, proveedor, actividad, ini, fin, predecesor, nota)
CADENA=[
 (1,"HPH","Inicio de contrato / proyecto",D(2026,7,10),D(2026,7,10),"—","Arranque de la provisión que gobierna la cadena"),
 (2,"HPH","Ingeniería de detalle / especificación",D(2026,8,17),D(2026,9,11),"1",""),
 (3,"HPH","Fin de ingeniería / CAD — FREEZING POINT",D(2026,9,18),D(2026,9,18),"2","Cierra ingeniería; habilita fabricación"),
 (4,"HPH","Fabricación de tableros de marshalling",D(2026,9,21),D(2026,11,6),"3","ESD/F&G y PSS 07A/07B"),
 (5,"HPH","Hardware fabricado y probado (test de fábrica)",D(2026,11,6),D(2026,11,6),"4","FAT integrada a fabricación"),
 (6,"HPH","Preparación de despacho / listo para despacho",D(2026,11,9),D(2026,11,20),"5",""),
 (7,"HPH","Despacho / shipping (Alemania → Argentina)",D(2026,11,23),D(2026,12,14),"6",""),
 (8,"HPH","Entrega DAP Ezeiza (Argentina)",D(2026,12,15),D(2026,12,15),"7","⚠ PA-01"),
 (9,"HPH","Nacionalización / despacho de aduana + transporte a base Inauco",D(2026,12,15),D(2027,1,25),"8","⚠ PA-01 — mín. 3-4 semanas; hoy con 2 semanas de holgura"),
 (10,"HPH→Inauco","RECEPCIÓN DE TABLEROS EN BASE INAUCO",D(2027,1,25),D(2027,1,25),"9","⚠ PA-02 — el tablero llega SIN adecuar"),
 (11,"AESA→HARG","COLOCACIÓN DE LA OC DE HIMA ARGENTINA",D(2026,9,23),D(2026,10,30),"—","⚠ PA-04 — PENDIENTE al 23-09-26; gobierna todo lo que sigue"),
 (12,"HARG","Inicio de trabajos HARG — documentación y programación (R1 y R2)",SH(D(2026,9,14)),SH(D(2027,1,29)),"11","⚠ PA-04 — corrido +49 días"),
 (13,"HARG","Adecuación de tableros de marshalling en base Inauco (6 semanas)",SH(D(2026,12,21)),SH(D(2027,1,29)),"10, 12","⚠ PA-05 — se superpone con la Pre iFAT de Inauco"),
 (14,"Inauco","Configuración PLC — Etapa 2",D(2027,1,25),D(2027,3,23),"10",""),
 (15,"Inauco","PRE iFAT (integración interna Inauco)",D(2027,2,10),D(2027,3,26),"14","⚠ PA-05 — convive con la adecuación de HARG"),
 (16,"HARG","FAT",SH(D(2027,2,24)),SH(D(2027,2,26)),"13","Cae dentro de la ventana de FAT del cliente"),
 (17,"Inauco","FAT CON EL CLIENTE — PCS + Comunicaciones",D(2027,3,29),D(2027,4,22),"15","Ventana confirmada por el crono. PG-002 Rev.0"),
 (18,"Inauco","FAT SIS (conjunta HARG / HPH / Inauco)",D(2027,3,29),D(2027,4,26),"15, 16","La FAT cierra con el ítem más tardío"),
 (19,"HARG","IFAT (Recursos 1 y 2)",SH(D(2027,3,1)),SH(D(2027,3,12)),"16, 18","⚠ PA-05 — termina 4 días después de la FAT SIS"),
 (20,"Inauco","Configuraciones post-FAT",D(2027,4,23),D(2027,5,4),"17, 18",""),
 (21,"Inauco","LIBERACIÓN / ENTREGA DE TABLEROS (camión de AESA)",D(2027,5,4),D(2027,5,4),"20","⚠ PA-05 — anterior al cierre del alcance HARG"),
 (22,"HARG","ENTREGA DEL ALCANCE HARG (IFAT + documentación)",SH(D(2027,3,12)),SH(D(2027,3,22)),"19","HITO FINAL DEL CAMINO CRÍTICO — 10-05-2027"),
]
NOTAS=[
 "Alcance: cronograma EJECUTIVO por hitos, con corte en la ENTREGA DE EQUIPOS / del alcance de cada provisión. Las tareas posteriores al despacho/entrega (montaje en planta, SAT, precomisionado, comisionado, PEM, CAO/DataBook) quedan FUERA.",
 "Rev5 — cambio 1: se adopta la actualización del cronograma ABB 5155-00-0009-VE-CO-PG-001_A del 21-09-2026, que reemplaza la del 21-08-2026. Se compararon ambos documentos tarea por tarea: 127 diferencias sobre 218 tareas.",
 "Rev5 — cambio 2: DUCTOS DE BARRAS — la entrega en sitio se atrasa 4 semanas (10-06-27 → 08-07-27) y pasa a ser el hito más tardío de todo el paquete. Origen: el insumo de AESA (layout + detalle de conexionado a trafo) se emitió el 18-09-26 en lugar del 29-08-26. Ver PA-07.",
 "Rev5 — cambio 3: SALAS SE#3 y SE#4 — los freezing points de ingeniería se corren (básica 11-09 → 25-09-26 · detalle 05-10 → 13-10-26), pero ABB mantiene las entregas: SE#3 ADELANTA su despacho EXW del 23-04 al 20-04-27 y SE#4 mantiene el 25-03-27. Ver PA-08.",
 "Rev5 — cambio 4: se agregan los hitos de FABRICACIÓN DEL SHELTER en ambas salas (SE#3 atrasada 3 semanas, SE#4 8 días) y de APROBACIÓN DE DISEÑO en los ductos de barras, que pasaron a ser determinantes. El PMS no registra cambios.",
 "Rev4 — cambio 1: se incorpora el cronograma de fabricación de INAUCO 5155-00-0029-VE-CA-PG-002 Rev.0 (14-09-2026), con el detalle por tablero (CPF2-PCS-SC-07a de 3 gabinetes, CPF2-PCS-SC-07b de 2 gabinetes, y los racks CPF-COM-07 y CPF-COM-SE3), sus PreFAT internos de taller y la liberación de pendientes.",
 "Rev4 — cambio 2: se adoptan las ACLARACIONES del pie del cronograma de Inauco: PRE iFAT del 10-02 al 26-03-27 (hito nuevo), FAT con el cliente del 29-03 al 22-04-27 (confirma la ventana de la Rev3) y LIBERACIÓN DE TABLEROS el 04-05-27 en camión de AESA — la entrega de Inauco se adelanta 13 días respecto de la Rev3 (17-05-27).",
 "Rev4 — cambio 3: HARG (HIMA Argentina) — al 23-09-2026 NO está colocada la orden de compra, por lo que el inicio de los trabajos se corre a principios de noviembre (02-11-2026). Se aplicó un corrimiento de 49 días corridos a todas las actividades PROPIAS de HARG; los hitos que son insumo de HPH (ya contratada) se mantienen en su fecha original.",
 "Rev4 — cambio 4: el CAMINO CRÍTICO se recalculó. Ya no cierra con la entrega de Inauco sino con la ENTREGA DEL ALCANCE HARG el 10-05-2027, que queda 6 días DESPUÉS de la liberación de los tableros de Inauco (04-05-27). Se incorporó la colocación de la OC de HARG como eslabón explícito de la cadena.",
 "ALERTA — PA-04: la OC de HIMA Argentina es hoy el ítem de mayor impacto del paquete: cada semana sin colocar corre una semana el cierre del sistema de seguridad. PA-05: la adecuación de HARG (08-02 → 19-03-27) se superpone con la Pre iFAT de Inauco y su IFAT termina después de la FAT SIS conjunta.",
 "Nota favorable: el corrimiento de HARG deja 2 semanas de holgura entre la recepción de los tableros de HPH en base Inauco (25-01-27) y el inicio de la adecuación (08-02-27), holgura que en la Rev3 no existía. Por eso PA-01 baja de severidad ALTA a MEDIA. Esa holgura no debe consumirse para absorber otros atrasos.",
 "Fechas no laborables declaradas por Inauco para la fabricación: 07 y 08-12-26 · 24 y 25-12-26 · 31-12-26 · 01-01-27 · 08 y 09-02-27. Días laborables de fabricación: 50.",
 "Reserva del emisor (crono. PG-002 Rev.0): las fechas son una estimación sobre las condiciones actuales y quedan sujetas a reprogramación por demoras en insumos, disponibilidad de stock de tarjetas/equipamiento de instrumentación y tiempos de importación o logística de componentes comerciales de alta complejidad.",
 "Rev3 — cambio: hito explícito de NACIONALIZACIÓN / DESPACHO DE ADUANA en HPH y anticipación de la prueba de integración PCS↔PMS a la ventana OCT-NOV 2026 (fecha a convenir).",
 "Rev2 — cambios: FAT EN FÁBRICA DE ORIGEN (Brasil, M&B) en las Salas SE#3 y SE#4; apertura del ACOPIO DE MATERIALES en hitos propios con hoja dedicada; separación de HIMA en HPH (Alemania) y HARG (Argentina).",
 "Fuentes: ABB Salas y Ductos 5155-00-0009-VE-CO-PG-001_A (21-09-26) · ABB PMS 5155-00-0007-VE-CO-PG-001_REV B (.mpp) · Inauco 5155-00-0029-VE-CO-PG-001-A Rev.A (24-08-26) y 5155-00-0029-VE-CA-PG-002 Rev.0 (14-09-26) · HPH 103618-Schedule 16-08-26 · HARG PSCH_CPF2_Rev3 (04-09-26) con corrimiento +49 d.",
 "Los hitos de FAT en fábrica de origen (Brasil) son witness de AESA y anteceden a la FAT del shelter completo; no deben confundirse entre sí.",
 "Estado (plan): se calcula contra la fecha de corte. 'Cumplido s/plan' = la fecha planificada ya venció y debe confirmarse el cumplimiento real con el proveedor; 'En curso' = en ejecución; 'Pendiente' = aún no inició.",
 "Las columnas Fecha real / % avance / Estado real / Observaciones quedan en blanco para la carga de seguimiento periódico (reporte a PPSA).",
 "Los hitos se ordenan cronológicamente dentro de cada provisión y se renumeran (H01, H02, ...). Marcas: ◆ = hito del camino crítico · ⚠ = punto de atención · fondo naranja = acopio de materiales.",
]
INTERFASES=[
 "CAMINO CRÍTICO DEL PROYECTO: OC de HARG → HPH (Alemania) → nacionalización → HARG (adecuación, programación, FAT e IFAT) → Inauco (FAT SIS y liberación). Cierra el 10-05-27 con la entrega del alcance HARG. Ver hoja 'Camino crítico'.",
 "AESA → HARG: la COLOCACIÓN DE LA OC está PENDIENTE al 23-09-26. Es el ítem que gobierna la fecha final del sistema de seguridad.  ⚠ PA-04",
 "HPH → entrega DAP Ezeiza (15/12/26) + nacionalización → recepción en base Inauco (25/01/27): habilita la adecuación de HARG, que ahora arranca el 08/02/27 (2 semanas de holgura).  ⚠ PA-01 / PA-02",
 "HARG ↔ Inauco → adecuación de tableros de marshalling (08/02 → 19/03/27) en base Inauco, superpuesta con la Pre iFAT de Inauco (10/02 → 26/03/27).  ⚠ PA-05",
 "HARG ↔ Inauco → el IFAT de HARG (19-30/04/27) termina 4 días después del cierre de la FAT SIS conjunta (26/04/27), y el alcance HARG cierra el 10/05/27, después de la liberación de tableros de Inauco (04/05/27).  ⚠ PA-05",
 "Inauco → FAT con el cliente 29/03 → 22/04/27 y LIBERACIÓN DE TABLEROS el 04/05/27 en camión de AESA (crono. PG-002 Rev.0): habilita el inicio de los trabajos de campo.",
 "ABB PMS ↔ Inauco → pruebas de integración PCS↔PMS en Bs.As., ventana OCT-NOV 2026, fecha a convenir.  ⚠ PA-03",
 "ABB PMS → entrega de equipamiento en shelter (17/02/27): habilita el montaje del PMS en las salas y los ensayos PMS.",
 "AESA → layout de instalación + detalle de conexionado a trafo: se emitió el 18/09/26 en lugar del 29/08/26 y atrasó 4 semanas los DUCTOS DE BARRAS, cuya entrega en sitio (08/07/27) es hoy la más tardía de todo el paquete.  ⚠ PA-07",
 "AESA → aprobación de ingeniería de las SALAS: los freezing points se corrieron a 25/09/26 (básica) y 13/10/26 (detalle); ABB absorbió el atraso sin mover las entregas, pero ya sin margen.  ⚠ PA-08",
 "AESA → confirmación de fabricación del rack CPF-COM-SE3, hoy marcado 'a confirmar' en el cronograma de Inauco.  ⚠ PA-06",
]
# HARG no tiene contrato: la columna de inicio no debe leerse como fecha de OC
INICIO_OVR={"HARG — HIMA Argentina · ingeniería, programación y FAT":"OC PENDIENTE"}
NAVY="1F3864"; HEAD="37474F"; CPFILL="FFEBEE"; PAFILL="FFF9C4"
thin=Side(style="thin",color="D9D9D9"); B=Border(thin,thin,thin,thin)
EST_FILL={"Cumplido s/plan":"2E7D32","En curso":"EF6C00","Pendiente":"90A4AE"}
SEV_FILL={"ALTA":"C62828","MEDIA":"EF6C00","BAJA":"2E7D32"}
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
MESES=month_list(D(2026,5,1),D(2027,7,1))
M3=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

wb=openpyxl.Workbook()

# ---------------- Hoja 1: Resumen Ejecutivo ----------------
ws=wb.active; ws.title="Resumen Ejecutivo"
ws.merge_cells("A1:K1"); cell(ws,1,1,f"CRONOGRAMA EJECUTIVO — PROVISIONES CRÍTICAS CPF-2  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,13,True); ws.row_dimensions[1].height=24
ws.merge_cells("A2:K2"); cell(ws,2,1,"Seguimiento para PPSA · Corte de alcance: hasta la ENTREGA DE EQUIPOS / del alcance (excluye montaje en planta, SAT, precom/comisionado y PEM)",None,"555555",False,9)
ws.merge_cells("A3:K3"); cell(ws,3,1,f"Fecha de corte del estado: {HOY.strftime('%d/%m/%Y')}   ·   CAMINO CRÍTICO (sistemas): OC HARG → HPH → HARG → Inauco, cierre 10/05/2027   ·   ÚLTIMA ENTREGA DEL PAQUETE: ductos de barras 08/07/2027   ·   {len(PUNTOS)} puntos de atención abiertos (PA-01 a PA-08)",None,"C62828",True,9)
hdr=["Proveedor","Provisión","Inicio contrato","Entrega","Hitos","Cumplidos s/plan","En curso","Pendientes","% avance plan","Próximo hito","Fecha próx. hito"]
for j,h in enumerate(hdr): cell(ws,5,1+j,h,HEAD,"FFFFFF",True,9,True,True)
ws.row_dimensions[5].height=30
r=6
for (name,ven,col,ms) in PROV:
    tot=len(ms); cum=sum(1 for m in ms if estado(m[2],m[3])=="Cumplido s/plan")
    enc=sum(1 for m in ms if estado(m[2],m[3])=="En curso"); pen=tot-cum-enc
    entrega=max(m[3] for m in ms); prox=next((m for m in ms if m[3]>=HOY), None)
    en_cp=any((name,m[1]) in CP_MARK for m in ms)
    cell(ws,r,1,ven,col,"FFFFFF",True,9,True)
    cell(ws,r,2,("◆ "if en_cp else "")+name,(CPFILL if en_cp else None),"000000",True,9,False,True)
    ini_txt=INICIO_OVR.get(name)
    if ini_txt: cell(ws,r,3,ini_txt,"FFF9C4","C62828",True,9,True)
    else: cell(ws,r,3,min(m[2] for m in ms).strftime("%d/%m/%Y"),None,"000000",False,9,True)
    cell(ws,r,4,entrega.strftime("%d/%m/%Y"),"1F3864","FFFFFF",True,9,True)
    cell(ws,r,5,tot,None,"000000",False,9,True)
    cell(ws,r,6,cum,"2E7D32","FFFFFF",True,9,True)
    cell(ws,r,7,enc,"EF6C00","FFFFFF",True,9,True)
    cell(ws,r,8,pen,"90A4AE","FFFFFF",True,9,True)
    cell(ws,r,9,f"{100*cum/tot:.0f}%",None,"000000",True,10,True)
    cell(ws,r,10,(prox[1] if prox else "—"),None,"000000",False,8.5,False,True)
    cell(ws,r,11,(prox[3].strftime("%d/%m/%Y") if prox else "—"),None,"C62828",True,9,True)
    ws.row_dimensions[r].height=28; r+=1
for c,w in zip("ABCDEFGHIJK",[9,48,14,13,7,15,10,12,13,44,15]): ws.column_dimensions[c].width=w
r+=1
cell(ws,r,1,"◆  Provisiones que integran el CAMINO CRÍTICO del proyecto (ver hoja 'Camino crítico')",CPFILL,"C62828",True,9,False,True)
ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=11); r+=2
cell(ws,r,1,"Interfases clave entre provisiones",None,NAVY,True,11); r+=1
for t in INTERFASES:
    cell(ws,r,1,"• "+t,None,"000000",t.startswith("CAMINO"),9,False,True); ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=11); ws.row_dimensions[r].height=16; r+=1

# ---------------- Hoja 2: Cronograma Ejecutivo ----------------
cw=wb.create_sheet("Cronograma Ejecutivo")
cw.merge_cells("A1:K1"); cell(cw,1,1,f"CRONOGRAMA EJECUTIVO POR HITOS  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); cw.row_dimensions[1].height=22
cw.merge_cells("A2:K2"); cell(cw,2,1,"Hitos ordenados cronológicamente.  ◆ = camino crítico · ⚠ PA-xx = punto de atención (ver hoja) · fondo naranja = acopio.  Columnas de seguimiento para carga periódica.",None,"555555",False,9)
H2=["Nº","Hito ejecutivo","Inicio","Fin","Crítico","Camino crítico","Estado (plan)","Fecha real","% avance","Estado real","Observaciones"]
r=4
for (name,ven,col,ms) in PROV:
    cw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=11)
    cell(cw,r,1,f"▶  {name}   ·   {ven}   ·   Entrega: {max(m[3] for m in ms).strftime('%d/%m/%Y')}",col,"FFFFFF",True,10)
    cw.row_dimensions[r].height=18; r+=1
    for j,h in enumerate(H2): cell(cw,r,1+j,h,HEAD,"FFFFFF",True,8.5,True,True)
    cw.row_dimensions[r].height=26; r+=1
    for (cod,hito,ini,fin,crit,obs,acop) in ms:
        est=estado(ini,fin); cp=(name,hito) in CP_MARK; pa=PA_MARK.get((name,hito))
        base = CPFILL if cp else ("FFF3E0" if acop else None)
        cell(cw,r,1,cod,None,"333333",True,8.5,True)
        cell(cw,r,2,("◆ " if cp else "")+hito,base,"000000",crit==1,9,False,True)
        cell(cw,r,3,ini.strftime("%d/%m/%Y"),None,"1F3864",False,8.5,True)
        cell(cw,r,4,fin.strftime("%d/%m/%Y"),None,"1F3864",crit==1,8.5,True)
        cell(cw,r,5,"SÍ" if crit else "",("C62828" if crit else None),"FFFFFF" if crit else "000000",True,8.5,True)
        cell(cw,r,6,"◆" if cp else "",(CPFILL if cp else None),"C62828",True,10,True)
        cell(cw,r,7,est,EST_FILL[est],"FFFFFF",True,8.5,True)
        for k in (8,9,10): cell(cw,r,k,"",None,"000000",False,9,True)
        txt=(f"⚠ {pa} · " if pa else "")+obs
        cell(cw,r,11,txt,(PAFILL if pa else None),("C62828" if pa else "555555"),bool(pa),8.5,False,True)
        r+=1
    r+=1
for c,w in zip("ABCDEFGHIJK",[6,60,12,12,8,13,16,12,10,14,50]): cw.column_dimensions[c].width=w
cw.freeze_panes="A4"

# ---------------- Hoja 3: Camino crítico ----------------
kw=wb.create_sheet("Camino crítico")
kw.merge_cells("A1:H1"); cell(kw,1,1,f"CAMINO CRÍTICO DEL PROYECTO  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); kw.row_dimensions[1].height=22
kw.merge_cells("A2:H2"); cell(kw,2,1,"Cadena crítica armada con la colocación de la OC de HARG y las provisiones HPH (HIMA Paul Hilpert, Alemania) → HARG (HIMA Argentina) → Inauco (PCS/SIS). Es la secuencia que gobierna la fecha de entrega del sistema de control y seguridad.",None,"555555",False,9,False,True)
kw.row_dimensions[2].height=26
KH=["Nº","Proveedor","Actividad / hito de la cadena","Inicio","Fin","Días","Predecesor","Observación / riesgo"]
for j,h in enumerate(KH): cell(kw,4,1+j,h,HEAD,"FFFFFF",True,9,True,True)
kw.row_dimensions[4].height=26
VENCOL={"HPH":"6A1B9A","HARG":"00838F","Inauco":"1565C0","HPH→Inauco":"4527A0","AESA→HARG":"C62828"}
r=5
for (n,ven,act,ini,fin,pred,nota) in CADENA:
    est=estado(ini,fin); dur=(fin-ini).days+1
    warn = nota.startswith("⚠") or "HITO FINAL" in nota
    cell(kw,r,1,n,None,"333333",True,9,True)
    cell(kw,r,2,ven,VENCOL.get(ven,"37474F"),"FFFFFF",True,9,True)
    cell(kw,r,3,act,(PAFILL if nota.startswith("⚠") else CPFILL),"000000",True,9.5,False,True)
    cell(kw,r,4,ini.strftime("%d/%m/%Y"),None,"1F3864",False,9,True)
    cell(kw,r,5,fin.strftime("%d/%m/%Y"),None,"1F3864",True,9,True)
    cell(kw,r,6,dur,None,"000000",False,9,True)
    cell(kw,r,7,pred,None,"555555",False,9,True)
    cell(kw,r,8,nota,None,("C62828" if warn else "555555"),warn,8.5,False,True)
    kw.row_dimensions[r].height=20; r+=1
for c,w in zip("ABCDEFGH",[6,13,62,13,13,8,13,60]): kw.column_dimensions[c].width=w
r+=1
for txt in [
 "Duración total de la cadena crítica: 10-07-2026 → 10-05-2027 (304 días corridos desde el inicio de HPH).",
 "El eslabón 11 (COLOCACIÓN DE LA OC DE HIMA ARGENTINA) está PENDIENTE al 23-09-2026 y es hoy el que gobierna la fecha final: cada semana sin colocar corre una semana el cierre del alcance HARG.",
 "El cierre del camino crítico pasó de la entrega de Inauco (17-05-27 en la Rev3) a la ENTREGA DEL ALCANCE HARG el 10-05-2027, por el corrimiento de 49 días de HARG. La liberación de tableros de Inauco se adelantó al 04-05-27.",
 "La cadena NO tiene holgura declarada, salvo las 2 semanas ganadas entre la recepción de los tableros de HPH (25-01-27) y el inicio de la adecuación de HARG (08-02-27).",
 "Los eslabones 11, 13, 19 y 22 concentran el riesgo del proyecto (PA-04 y PA-05): OC de HARG, adecuación de los tableros de marshalling, IFAT y cierre del alcance HARG.",
 "Los eslabones 14, 15 y 17 (configuración PLC Etapa 2, Pre iFAT y FAT con el cliente de Inauco) hoy no gobiernan la fecha final, pero comparten taller y tableros con HARG durante seis semanas.",
 "El resto de las provisiones ABB (Salas SE#3 y SE#4, PMS y Ductos de barras) corren por caminos propios. ATENCIÓN: el DUCTO DE BARRAS entrega el 08-07-27 — 4 semanas más tarde que en la Rev4 y casi 2 meses después del cierre de esta cadena. Es un equipo de montaje, no de sistema, pero es hoy la última entrega del paquete y condiciona la energización.",
]:
    cell(kw,r,1,"• "+txt,None,"000000",False,9,False,True); kw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8); kw.row_dimensions[r].height=16; r+=1

# ---------------- Hoja 4: Puntos de atención ----------------
pw=wb.create_sheet("Puntos de atención")
pw.merge_cells("A1:B1"); cell(pw,1,1,f"PUNTOS DE ATENCIÓN  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); pw.row_dimensions[1].height=22
pw.merge_cells("A2:B2"); cell(pw,2,1,"Fechas de los cronogramas de los proveedores que, contrastadas entre sí, resultan poco probables o requieren definición. Se informan a PPSA para su seguimiento.",None,"555555",False,9)
CAMPOS=[("Interfase","inter"),("Fecha / premisa del plan","plan"),("Riesgo identificado","riesgo"),
        ("Impacto sobre el camino crítico","impacto"),("Fecha realista estimada","real"),("Acción recomendada","accion")]
r=4
for p in PUNTOS:
    pw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=2)
    cell(pw,r,1,f"{p['cod']}  ·  {p['tit']}",SEV_FILL[p['sev']],"FFFFFF",True,11,False,True)
    pw.row_dimensions[r].height=22; r+=1
    cell(pw,r,1,"Severidad",None,"333333",True,9); cell(pw,r,2,p['sev'],SEV_FILL[p['sev']],"FFFFFF",True,9,True); r+=1
    for lab,k in CAMPOS:
        cell(pw,r,1,lab,"ECEFF1","333333",True,9,False,True)
        cell(pw,r,2,p[k],(PAFILL if k in ("riesgo","impacto") else None),"000000",k in ("riesgo","impacto"),9,False,True)
        pw.row_dimensions[r].height=max(16,14*(1+len(p[k])//95)); r+=1
    r+=1
pw.column_dimensions['A'].width=32; pw.column_dimensions['B'].width=125

# ---------------- Hoja 5: Acopio de Materiales ----------------
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

# ---------------- Hoja 6: Timeline ----------------
tw=wb.create_sheet("Timeline (Gantt hitos)")
tw.merge_cells(start_row=1,start_column=1,end_row=1,end_column=3+len(MESES))
cell(tw,1,1,f"TIMELINE DE HITOS EJECUTIVOS  ·  {REV} · {FECHA}  ·  corte: entrega de equipos  ·  ◆ = camino crítico",NAVY,"FFFFFF",True,12,True); tw.row_dimensions[1].height=22
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
        cp=(name,hito) in CP_MARK
        base = CPFILL if cp else ("FFF3E0" if acop else None)
        cell(tw,r,1,f"{cod} · "+("◆ " if cp else "")+hito,base,"000000",crit==1,8.3,False,True)
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
for c,w in zip("ABC",[62,10,10]): tw.column_dimensions[c].width=w
tw.freeze_panes="D4"

# ---------------- Hoja 7: Notas ----------------
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
    ac=sum(1 for m in ms if m[6]); cp=sum(1 for m in ms if (name,m[1]) in CP_MARK)
    print(f"  {name[:50]:50} hitos={len(ms):2} acopio={ac} CC={cp:2} cumpl={cum:2} entrega={max(m[3] for m in ms).strftime('%d/%m/%Y')}")
# control: todas las claves de CP_MARK y PA_MARK deben existir
todos={(n,m[1]) for (n,v,c,ms) in PROV for m in ms}
falt=[k for k in list(CP_MARK)+list(PA_MARK) if k not in todos]
print("claves sin correspondencia:",falt if falt else "ninguna")
