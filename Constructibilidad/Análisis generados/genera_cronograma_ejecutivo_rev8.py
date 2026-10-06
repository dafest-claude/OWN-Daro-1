#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CRONOGRAMA EJECUTIVO — Provisiones críticas CPF-2 (Rev8, 2026-10-06).
Cambios respecto de Rev7:
  1) ABB SALAS ELÉCTRICAS — se adopta el cronograma 5155-00-0009-VE-CO-PG-001_A del
     02-10-2026 y el INFORME DE AVANCE Nº 5 (R-2631036, 05-10-2026). Lo esencial:
       · Se incorporan las FECHAS CONTRACTUALES de entrega, que hasta ahora no teníamos:
         SE#3 contractual 14-05-27 contra objetivo 20/23-04-27 (21 días de margen) y
         SE#4 contractual 08-04-27 contra objetivo 25-03-27 (14 días de margen).
       · APROBADA la ingeniería de DETALLE de todos los EQUIPOS de ambas salas (CCM, MT,
         auxiliares, UPS/baterías y variadores). Queda pendiente la aprobación de la
         ingeniería de las SALAS propiamente dichas: PA-08 se acota a ese alcance.
       · Se incorpora la PREVISIÓN DE FECHAS Y LUGARES DE FAT por equipo: UPS en Santa Fe,
         TGMT y CCM en Sorocaba (Brasil), auxiliares en AMBA y las salas en Mendoza.
       · DUCTOS DE BARRAS: ABB informa que no pudieron diseñarse por falta de información
         sobre los gabinetes de los transformadores de MT y sus posiciones relativas; el
         diseño arrancó el 21-09-26. ABB evalúa plazo e IMPACTO ECONÓMICO y pide a AESA
         adecuar la fecha contractual de las posiciones 120, 130 y 140 de la OC
         45089944971 para evitar penalidades. El subproveedor EAE (Turquía) figura con
         fecha de entrega en origen "tbd". PA-07 se reescribe con esto.
  2) La hoja "Avance a SEP-2026" se amplía con el avance ponderado de los nueve paquetes
     de las salas, la previsión de FAT, las órdenes de compra a subproveedores y el estado
     de certificación de ambas OC.
  3) PUNTOS DE ATENCIÓN: se suma PA-13 (suministro de equipos de control de acceso, a
     cargo de AESA) y se refuerzan PA-07, PA-08 y PA-12.
Cambios de la Rev7 respecto de Rev6:
  1) ABB PMS — se adopta el cronograma 5155-00-0007-VE-CO-PG-001 SEP-2026 (REV 0, emitido
     01-10-2026) y el INFORME DE AVANCE Nº 4 (E-2611027, 30-09-2026). Lo esencial:
       · El ARMADO DE TABLEROS se corre 4 semanas: del 11-09 al 09-10-2026 (50 días
         hábiles). El fin de armado se mantiene el 17-12-26 y la ENTREGA EN SHELTER el
         17-02-27: ABB absorbió el corrimiento.
       · La GESTIÓN DE COMPRA DE MATERIALES pasa al 100%: switches de red, material S800
         y tableros. La recepción de materiales cerró en septiembre (100% real contra 70%
         previsto).
       · Las pruebas de maqueta integrada se desdoblan en dos: MAQUETA CPF1 (19-23/10/26)
         y MAQUETA CPF2 (30/11 → 03/12/26), esta última con INAUCO por la comunicación con
         el PCS. Queda así fijada la prueba de integración PCS↔PMS que en la Rev6 figuraba
         "a convenir" (PA-03).
       · Los procedimientos de FAT y SAT se corren unas 6 semanas (emisión Rev.A desde el
         19-10-26, aprobación de la Rev.0 el 11-12-26).
  2) NUEVA HOJA "Avance a SEP-2026": avance ponderado por especialidad del informe de ABB
     (53% real contra 44% previsto, +9 puntos), eventos cumplidos y próximos, temas
     pendientes, acuerdos del período y estado de certificación; más el avance declarado
     por cada proveedor en su propio cronograma.
  3) PUNTOS DE ATENCIÓN: se suman PA-11 (temas pendientes prioritarios que ABB reclama a
     AESA) y PA-12 (una sola FAT integral de 5 días frente a las dos ventanas de ensayos
     PMS del cronograma de salas). PA-03 baja a severidad BAJA: ya tiene fecha.
Cambios de la Rev6 respecto de Rev5:
  1) NUEVA PROVISIÓN — CAT MIRON · TRANSFORMADORES. Se incorpora el cronograma de
     fabricación 5155-00-0022-VE-CO-PG-001 / ACAL-00102-EMBR-TRSEC-PP-X-0001 Rev.A
     (emitido 01-09-2026, actualización del 29-09-2026): 13 transformadores en 5 órdenes
     de trabajo, con entregas escalonadas entre el 13-11 y el 11-12-2026.
       · OT 5209 — 4 × 2000 kVA 13,2/0,4 kV .......... entrega 13-11-26
       · OT 5210 — 2 × 1600 kVA 13,2/0,4 kV .......... entrega 13-11-26
       · OT 5212 — 2 × 2000 kVA 0,4/13,86 kV ......... entrega 27-11-26
       · OT 5213 — 3 × 500 kVA 0,38/0,48 kV .......... entrega 04-12-26
       · OT 5211 — 2 × 2500 kVA 13,2/6,9 kV .......... entrega 11-12-26
  2) PUNTOS DE ATENCIÓN: se suman PA-09 (las pruebas FAT de los transformadores NO están
     incluidas en el cronograma del proveedor y dependen de la agenda de un laboratorio
     externo) y PA-10 (los transformadores llegan 7 meses antes que los ductos de barras
     que los vinculan, lo que obliga a prever almacenamiento y conservación).
  3) Se incorpora la interfase CAT Miron → ABB Ductos de barras: el detalle de conexionado
     a transformador es insumo del diseño de los ductos.
Cambios de la Rev5 respecto de Rev4:
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
REV="Rev8"; FECHA="2026-10-06"; HOY=date(2026,10,6)
SHIFT_HARG=49   # días corridos de corrimiento por OC no colocada (14-09-26 -> 02-11-26)
D=lambda y,m,d: date(y,m,d)
from datetime import timedelta
SH=lambda d: d+timedelta(SHIFT_HARG)   # corrimiento de las actividades propias de HARG

# (hito, inicio, fin, crítico, observación, es_acopio)
PROV=[
 ("ABB — Sala Eléctrica SE#3 (BT)","ABB","C62828",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"OC 45089944971 — pólizas de fiel cumplimiento y fondo de reparo aprobadas",0),
  ("Kick-off Meeting (KOM) técnico",D(2026,5,15),D(2026,5,18),0,"",0),
  ("Confirmación de layout de sala y equipamiento",D(2026,5,18),D(2026,5,20),0,"Insumo AESA",0),
  ("Emisión de ingeniería básica",D(2026,5,20),D(2026,9,11),0,"100%",0),
  ("Aprobación de ingeniería básica",D(2026,9,11),D(2026,9,25),0,"90% — VENCIDA al corte; certificación del hito 2 prevista para octubre",0),
  ("FREEZING POINT — ingeniería básica",D(2026,9,25),D(2026,9,25),1,"0% — VENCIDO. Sólo alcanza a la SALA: la ingeniería de los equipos ya está aprobada",0),
  ("Emisión de ingeniería de detalle",D(2026,5,22),D(2026,9,29),0,"95%",0),
  ("Aprobación de ingeniería de detalle — EQUIPOS",D(2026,9,1),D(2026,9,30),0,"APROBADA en septiembre: CCM, auxiliares, UPS/baterías y variadores BT",0),
  ("Aprobación de ingeniería de detalle — SALA",D(2026,9,29),D(2026,10,13),0,"Pendiente — alcance del shelter",0),
  ("FREEZING POINT — ingeniería de detalle",D(2026,10,13),D(2026,10,13),1,"Cierra ingeniería del shelter",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,19),D(2026,5,19),0,"Compra y acopio sala BT",1),
  ("ACOPIO — Materiales de los CCM (críticos y secundarios)",D(2026,6,19),D(2026,10,20),0,"93% — críticos cerrados el 30-09-26",1),
  ("ACOPIO — Rectificadores / UPS / baterías",D(2026,6,9),D(2026,10,12),0,"93% — fusibles y ventiladores 100%, interruptores 98%, semiconductores 63%",1),
  ("ACOPIO — Tableros auxiliares",D(2026,7,22),D(2026,10,30),0,"72% — acopio parcial de equipamiento mecánico y eléctrico",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,12,10),D(2026,12,10),1,"Cierre de acopio de la sala BT",1),
  ("Inicio de fabricación en fábrica de origen",D(2026,10,6),D(2026,10,6),0,"CCM: inicio temprano de columna y módulos extraíbles",0),
  ("Fabricación de tableros en fábrica de origen",D(2026,10,6),D(2026,12,11),1,"CCM 06/10-19/11 (Sorocaba) · rectif./UPS 14-23/10 (Santa Fe) · auxiliares 30/10-11/12 (AMBA)",0),
  ("Ensayos internos en fábrica de origen",D(2026,10,26),D(2026,12,3),0,"Previos a la FAT",0),
  ("FAT EN FÁBRICA DE ORIGEN — equipos",D(2026,11,12),D(2026,12,16),1,"Witness AESA · UPS/baterías 12-13/11 Santa Fe · CCM 03-11/12 Sorocaba · auxiliares 14-16/12 AMBA",0),
  ("Adecuación y despacho de fábrica de origen",D(2026,11,27),D(2026,12,17),1,"Condición EXW en origen",0),
  ("Entrega en Shelterista (Mendoza)",D(2026,11,26),D(2027,1,10),1,"Arribo escalonado: rectif./UPS 26-30/11 · aux. 21-28/12 · drives 23-29/12 · CCM 23/12-10/01",0),
  ("Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)",D(2026,11,23),D(2027,3,5),1,"Shelterista Bottino — inspección a proveedor prevista en octubre",0),
  ("Montaje e integración de equipos en sala",D(2027,2,26),D(2027,3,17),0,"Shelterista",0),
  ("Interconexión de equipos",D(2027,3,9),D(2027,3,23),0,"",0),
  ("FAT SALA ELÉCTRICA (shelter completo)",D(2027,3,29),D(2027,4,2),1,"Mendoza · previsión del informe Nº5 (el cronograma del 02-10 la ubica 23-30/03)",0),
  ("Ensayos PMS en sala (INTERFASE — contrato PMS)",D(2027,3,30),D(2027,4,6),0,"Alcance del contrato PMS",0),
  ("Desconexión, embalaje y despacho",D(2027,4,6),D(2027,4,20),1,"",0),
  ("DESPACHO / ENTREGA EXW",D(2027,4,20),D(2027,4,20),1,"HITO FINAL — objetivo ABB. Fecha CONTRACTUAL 14-05-27: 21 días de margen",0),
 ]),
 ("ABB — Sala Eléctrica SE#4 (MT)","ABB","AD1457",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"OC 45089944973 — pólizas de fiel cumplimiento y fondo de reparo aprobadas",0),
  ("Kick-off Meeting (KOM) técnico",D(2026,5,15),D(2026,5,18),0,"",0),
  ("Confirmación de layout de sala y equipamiento",D(2026,5,18),D(2026,5,20),0,"Insumo AESA",0),
  ("Emisión de ingeniería básica",D(2026,5,20),D(2026,9,11),0,"100%",0),
  ("Aprobación de ingeniería básica",D(2026,9,11),D(2026,9,25),0,"90% — VENCIDA al corte; certificación del hito 2 prevista para octubre",0),
  ("FREEZING POINT — ingeniería básica",D(2026,9,25),D(2026,9,25),1,"0% — VENCIDO. Sólo alcanza a la SALA: la ingeniería de los equipos ya está aprobada",0),
  ("Emisión de ingeniería de detalle",D(2026,5,22),D(2026,9,29),0,"95%",0),
  ("Aprobación de ingeniería de detalle — EQUIPOS",D(2026,9,1),D(2026,9,30),0,"APROBADA en septiembre: tableros MT, auxiliares y variadores MT",0),
  ("Aprobación de ingeniería de detalle — SALA",D(2026,9,29),D(2026,10,13),0,"Pendiente — alcance del shelter",0),
  ("FREEZING POINT — ingeniería de detalle",D(2026,10,13),D(2026,10,13),1,"Cierra ingeniería del shelter",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,19),D(2026,5,19),0,"Compra y acopio sala MT",1),
  ("ACOPIO — Celdas MT (materiales principales y secundarios)",D(2026,6,22),D(2026,10,30),1,"78% — principales 78%, secundarios 95%",1),
  ("ACOPIO — Variadores MT (China) y BT (Finlandia)",D(2026,8,18),D(2026,9,30),0,"Fabricación y ensayos FINALIZADOS — despacho de origen 30/09-09/10/26",1),
  ("ACOPIO — Tableros auxiliares",D(2026,7,22),D(2026,10,30),0,"72% — acopio parcial de equipamiento mecánico y eléctrico",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,12,10),D(2026,12,10),1,"Cierre de acopio de la sala MT",1),
  ("Inicio de fabricación en fábrica de origen",D(2026,10,8),D(2026,10,8),0,"Celdas MT — inicio de gabinetes de comando (Sorocaba)",0),
  ("Fabricación de tableros y celdas en fábrica de origen",D(2026,10,8),D(2026,12,11),1,"Celdas MT 08/10-16/11 (Sorocaba) · auxiliares 30/10-11/12 (AMBA)",0),
  ("Ensayos internos en fábrica de origen",D(2026,11,17),D(2026,11,26),0,"Previos a la FAT",0),
  ("FAT EN FÁBRICA DE ORIGEN — equipos",D(2026,11,27),D(2026,12,16),1,"Witness AESA · TGMT 27/11-03/12 Sorocaba · auxiliares 14-16/12 AMBA",0),
  ("Despacho de fábrica de origen (celdas MT y tableros)",D(2026,12,4),D(2026,12,21),1,"Acond./embalaje 04-10/12 · FCA 10-14/12",0),
  ("Entrega en Shelterista (Mendoza)",D(2026,12,28),D(2027,1,6),1,"Arribo escalonado: aux. 21-28/12 · drives 23-30/12 · celdas 04-06/01",0),
  ("Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)",D(2026,10,6),D(2027,1,29),1,"Shelterista Bottino — inicio de fabricación previsto en octubre",0),
  ("Montaje e integración de equipos en sala",D(2027,1,29),D(2027,2,19),0,"Shelterista",0),
  ("Interconexión de equipos",D(2027,2,17),D(2027,2,23),0,"",0),
  ("FAT SALA ELÉCTRICA (shelter completo)",D(2027,2,22),D(2027,2,26),1,"Mendoza · previsión del informe Nº5 (el cronograma del 02-10 la ubica 25/02-01/03)",0),
  ("Ensayos PMS en sala (INTERFASE — contrato PMS)",D(2027,3,3),D(2027,3,9),0,"Alcance del contrato PMS",0),
  ("Desconexión, embalaje y despacho",D(2027,3,10),D(2027,3,25),1,"",0),
  ("DESPACHO / ENTREGA EXW",D(2027,3,25),D(2027,3,25),1,"HITO FINAL — objetivo ABB. Fecha CONTRACTUAL 08-04-27: 14 días de margen",0),
 ]),
 ("ABB — Ductos de barras","ABB","795548",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"Mismo contrato ABB (alcance separado)",0),
  ("Emisión layout instalación + detalle conexionado trafo",D(2026,7,20),D(2026,9,18),1,"INSUMO AESA — se corrió del 29-08 al 18-09-26 y arrastró todo el bloque",0),
  ("FREEZING POINT — inicio de diseño",D(2026,9,18),D(2026,9,18),1,"CORRIDO 20 días (antes 29-08-26)",0),
  ("Diseño de ductos de barras",D(2026,9,18),D(2026,12,14),1,"15% — el diseño arrancó EFECTIVAMENTE el 21-09-26, al recibir la información de gabinetes de trafos MT",0),
  ("Aprobación de diseño",D(2026,12,14),D(2026,12,30),0,"Aprobación AESA",0),
  ("FREEZING POINT — cierre de diseño",D(2026,12,30),D(2026,12,30),1,"CORRIDO 26 días (antes 04-12-26) — habilita fabricación",0),
  ("ACOPIO / fabricación de ductos de barras",D(2026,12,30),D(2027,4,9),1,"100 días corridos (acopio integrado a fabricación)",1),
  ("FCA fábrica de origen",D(2027,4,9),D(2027,4,16),1,"Puesta a disposición en origen",0),
  ("Traslado internacional + aduana",D(2027,4,16),D(2027,7,5),1,"80 días corridos — tramo más largo",0),
  ("ENTREGA DE EQUIPOS EN SITIO",D(2027,7,5),D(2027,7,8),1,"HITO FINAL — la entrega más tardía del paquete. ABB pide adecuar la fecha contractual de las posiciones 120/130/140 de la OC 45089944971 y evalúa impacto económico",0),
 ]),
 ("CAT Miron — Transformadores (13 unidades)","CAT Miron","33691E",[
  ("Inicio de ingeniería (primer lote OT 5209 / 5210)",D(2026,7,13),D(2026,7,13),0,"Fecha de OC no consignada en el cronograma del proveedor",0),
  ("Ingeniería — cálculo y diseño (5 órdenes de trabajo)",D(2026,7,13),D(2026,8,21),0,"15 días hábiles por OT — 100% en las 5 OT",0),
  ("Ingeniería — proyecto de ingeniería en detalle",D(2026,7,27),D(2026,9,4),1,"15 días hábiles por OT — 100% salvo OT 5211 (75%); insumo del detalle de conexionado a trafo",0),
  ("ACOPIO — Conductores BT/MT",D(2026,7,27),D(2026,9,25),1,"30 días hábiles por OT — avance 40% (OT 5211) a 90% (OT 5209/5210)",1),
  ("ACOPIO — Materiales gruesos (incluye gabinetes de los transformadores)",D(2026,8,3),D(2026,10,16),1,"45 días hábiles por OT — avance 30% a 50%",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,10,16),D(2026,10,16),1,"Cierre del acopio de la última OT (5211)",1),
  ("Fabricación — núcleo magnético",D(2026,9,7),D(2026,11,6),0,"30 días hábiles por OT",0),
  ("Fabricación — bobinado BT/MT",D(2026,9,7),D(2026,11,27),1,"40 días hábiles por OT — actividad de mayor plazo",0),
  ("Fabricación — montaje final y montaje de gabinete",D(2026,11,2),D(2026,12,4),1,"3 + 5 días hábiles por OT",0),
  ("Ensayos internos de laboratorio",D(2026,11,2),D(2026,12,4),1,"2 días hábiles por OT — NO incluyen las pruebas FAT",0),
  ("PRUEBAS FAT — A COORDINAR (no incluidas en el cronograma)",D(2026,11,2),D(2026,12,11),1,"Sujetas a disponibilidad de agenda de laboratorio externo",0),
  ("Acondicionamiento y carga — OT 5209 (4×2000 kVA) y OT 5210 (2×1600 kVA)",D(2026,11,9),D(2026,11,13),1,"2 días de acondicionamiento + 1 de carga",0),
  ("ENTREGA lote 1 — OT 5209 (4 un.) + OT 5210 (2 un.)",D(2026,11,13),D(2026,11,13),1,"6 transformadores 13,2/0,4 kV",0),
  ("Acondicionamiento y carga — OT 5212 (2×2000 kVA)",D(2026,11,23),D(2026,11,27),1,"",0),
  ("ENTREGA lote 2 — OT 5212 (2 un.)",D(2026,11,27),D(2026,11,27),1,"2 transformadores 0,4/13,86 kV",0),
  ("Acondicionamiento y carga — OT 5213 (3×500 kVA)",D(2026,11,30),D(2026,12,4),1,"",0),
  ("ENTREGA lote 3 — OT 5213 (3 un.)",D(2026,12,4),D(2026,12,4),1,"3 transformadores 0,38/0,48 kV",0),
  ("Acondicionamiento y carga — OT 5211 (2×2500 kVA)",D(2026,12,7),D(2026,12,11),1,"",0),
  ("ENTREGA FINAL — OT 5211 (2 un.) · cierre de la provisión",D(2026,12,11),D(2026,12,11),1,"HITO FINAL — 2 transformadores 13,2/6,9 kV; completa las 13 unidades",0),
 ]),
 ("ABB — PMS (Power Management System)","ABB","EF6C00",[
  ("Inicio de contrato — Recepción / aceptación OC",D(2026,5,8),D(2026,5,11),0,"OC 4508945953 — 100%",0),
  ("Kick-off Meeting (KOM) y planificación detallada",D(2026,5,22),D(2026,6,5),0,"100%",0),
  ("Emisión de ingeniería básica (documentación para aprobación)",D(2026,6,8),D(2026,7,31),0,"100% · Hito cert. 10% el 31-07-26",0),
  ("Aprobación de ingeniería básica (AESA)",D(2026,7,6),D(2026,8,14),0,"100% · Hito cert. 10% el 14-08-26",0),
  ("Emisión de ingeniería Rev.0",D(2026,7,20),D(2026,8,28),0,"100%",0),
  ("Aprobación ingeniería Rev.0 — FREEZING POINT",D(2026,8,3),D(2026,9,11),1,"100% — cierra ingeniería",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,22),D(2026,5,22),0,"Switches de red, S800, tableros",1),
  ("ACOPIO — Material S800 y tableros",D(2026,5,22),D(2026,10,8),0,"GESTIÓN DE COMPRA AL 100% (Rev6: en curso)",1),
  ("ACOPIO — Switches de red",D(2026,5,22),D(2026,11,5),1,"GESTIÓN DE COMPRA AL 100% — ítem de mayor plazo",1),
  ("Configuración / programación del sistema",D(2026,8,17),D(2026,12,11),1,"32% — entorno virtual 100%, pantallas 50%, I/O 50%",0),
  ("INICIO DE ARMADO DE TABLEROS",D(2026,10,9),D(2026,10,9),1,"CORRIDO 4 SEMANAS (Rev6: 11-09-26) — 50 días hábiles",0),
  ("Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF1",D(2026,10,19),D(2026,10,23),0,"Maqueta de CPF1 — prueba temprana",0),
  ("Emisión de procedimientos de FAT y SAT (Rev.A)",D(2026,10,19),D(2026,10,30),0,"Corrido ~6 semanas respecto de la Rev6",0),
  ("Aprobación de procedimientos FAT / SAT Rev.A (AESA)",D(2026,11,2),D(2026,11,13),0,"",0),
  ("Configuración de variables de comunicación y lógicas del sistema",D(2026,11,2),D(2026,12,11),1,"IEC61850 / PROFINET / MODBUS — requiere los mapas pendientes de AESA",0),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,11,5),D(2026,11,5),1,"Hito cert. 25%",1),
  ("Emisión de procedimientos de FAT y SAT Rev.0",D(2026,11,16),D(2026,11,27),0,"",0),
  ("Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF2 (integración con PCS Inauco)",D(2026,11,30),D(2026,12,3),1,"ABB/AESA/PPSA + INAUCO — reemplaza la prueba 'a convenir' de la Rev6",0),
  ("Aprobación de procedimientos FAT / SAT Rev.0 (AESA)",D(2026,11,30),D(2026,12,11),0,"",0),
  ("FIN DE ARMADO DE TABLEROS",D(2026,12,17),D(2026,12,17),1,"Se sostiene pese al corrimiento del inicio",0),
  ("Pruebas internas (pre-FAT) de tableros y sistema",D(2026,12,18),D(2027,1,7),0,"15 días hábiles",0),
  ("FAT Tableros + Sistema",D(2027,1,25),D(2027,1,29),1,"Hito cert. 20% · witness AESA",0),
  ("Ajustes post-FAT / liberación",D(2027,2,1),D(2027,2,3),0,"",0),
  ("Despachos",D(2027,2,4),D(2027,2,17),1,"",0),
  ("ENTREGA DE EQUIPAMIENTO EN SHELTER",D(2027,2,17),D(2027,2,17),1,"HITO FINAL — cert. 30%. Fuera del corte: FAT INTEGRAL del shelter 25-31/03/27 y DataBook 31/03/27",0),
 ]),
 ("Inauco — PCS / SCADA / Comunicaciones","Inauco","1565C0",[
  ("Inicio de contrato — Recepción de OC",D(2026,7,28),D(2026,7,28),0,"",0),
  ("Recepción de ing. Rev.0 p/ tableros (insumo AESA)",D(2026,8,10),D(2026,8,10),1,"Requisito 1 — insumo de AESA",0),
  ("Emisión de ingeniería (Rev. A y 0)",D(2026,8,11),D(2026,11,4),0,"",0),
  ("Configuración PLC — Etapa 1",D(2026,8,26),D(2026,9,22),0,"",0),
  ("Configuración SCADA",D(2026,8,26),D(2027,3,23),1,"",0),
  ("Pruebas de integración PCS↔PMS — Maqueta CPF2 (Bs.As.)",D(2026,11,30),D(2026,12,3),1,"FECHA ACORDADA (informe ABB Nº4): ABB/AESA/PPSA + Inauco",0),
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
 ("Inauco — PCS / SCADA / Comunicaciones","Pruebas de integración PCS↔PMS — Maqueta CPF2 (Bs.As.)"):"PA-03",
 ("ABB — PMS (Power Management System)","Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF2 (integración con PCS Inauco)"):"PA-03",
 ("ABB — PMS (Power Management System)","Configuración de variables de comunicación y lógicas del sistema"):"PA-11",
 ("ABB — PMS (Power Management System)","ENTREGA DE EQUIPAMIENTO EN SHELTER"):"PA-12",
 ("ABB — Sala Eléctrica SE#3 (BT)","Ensayos PMS en sala (INTERFASE — contrato PMS)"):"PA-12",
 ("ABB — Sala Eléctrica SE#4 (MT)","Ensayos PMS en sala (INTERFASE — contrato PMS)"):"PA-12",
 ("Inauco — PCS / SCADA / Comunicaciones","Fabricación racks de comunicaciones CPF-COM-07 y CPF-COM-SE3"):"PA-06",
 ("ABB — Ductos de barras","Emisión layout instalación + detalle conexionado trafo"):"PA-07",
 ("CAT Miron — Transformadores (13 unidades)","PRUEBAS FAT — A COORDINAR (no incluidas en el cronograma)"):"PA-09",
 ("CAT Miron — Transformadores (13 unidades)","Ensayos internos de laboratorio"):"PA-09",
 ("CAT Miron — Transformadores (13 unidades)","ENTREGA lote 1 — OT 5209 (4 un.) + OT 5210 (2 un.)"):"PA-10",
 ("CAT Miron — Transformadores (13 unidades)","ENTREGA FINAL — OT 5211 (2 un.) · cierre de la provisión"):"PA-10",
 ("ABB — Ductos de barras","ENTREGA DE EQUIPOS EN SITIO"):"PA-07",
 ("ABB — Ductos de barras","FREEZING POINT — inicio de diseño"):"PA-07",
 ("ABB — Ductos de barras","FREEZING POINT — cierre de diseño"):"PA-07",
 ("ABB — Sala Eléctrica SE#3 (BT)","FREEZING POINT — ingeniería básica"):"PA-08",
 ("ABB — Sala Eléctrica SE#3 (BT)","Aprobación de ingeniería básica"):"PA-08",
 ("ABB — Sala Eléctrica SE#3 (BT)","Aprobación de ingeniería de detalle — SALA"):"PA-08",
 ("ABB — Sala Eléctrica SE#3 (BT)","FREEZING POINT — ingeniería de detalle"):"PA-08",
 ("ABB — Sala Eléctrica SE#4 (MT)","FREEZING POINT — ingeniería básica"):"PA-08",
 ("ABB — Sala Eléctrica SE#4 (MT)","Aprobación de ingeniería básica"):"PA-08",
 ("ABB — Sala Eléctrica SE#4 (MT)","Aprobación de ingeniería de detalle — SALA"):"PA-08",
 ("ABB — Sala Eléctrica SE#4 (MT)","FREEZING POINT — ingeniería de detalle"):"PA-08",
 ("ABB — Sala Eléctrica SE#3 (BT)","Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)"):"PA-13",
 ("ABB — Sala Eléctrica SE#4 (MT)","Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)"):"PA-13",
 ("ABB — Ductos de barras","Diseño de ductos de barras"):"PA-07",
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
 dict(cod="PA-13", tit="Suministro de los equipos físicos de control de acceso de ambas salas, a cargo de AESA",
      inter="AESA → ABB Salas",
      plan="El informe Nº 5 lista como ÚNICO tema pendiente prioritario: «Control de acceso. Suministro de equipos físicos de acuerdo con cronograma de cada sala — Resp. AESA». El cronograma fija el suministro el 04-01-2027 para SE#4 y el 13-02-2027 para SE#3.",
      riesgo="Es un suministro de AESA que se integra dentro de la fabricación del shelter, antes de la instalación eléctrica. Si llega tarde obliga a reabrir trabajos ya cerrados en el shelterista o a entregar la sala incompleta.",
      impacto="SE#4 instala en enero y SE#3 en febrero de 2027; ambas fechas caen sobre la instalación eléctrica del shelter. Un atraso aquí golpea directamente la FAT de sala y, por lo tanto, el despacho EXW.",
      real="04-01-2027 (SE#4) y 13-02-2027 (SE#3): la compra debe estar colocada con antelación suficiente.",
      accion="Confirmar con Suministros el estado de la compra de los equipos de control de acceso y su plazo de entrega, y comprometer con ABB la fecha de puesta a disposición en el shelterista. Es el único pendiente que ABB nos imputa en este informe.",
      sev="ALTA"),
 dict(cod="PA-11", tit="Temas pendientes prioritarios que ABB reclama a AESA para poder configurar el PMS",
      inter="AESA Ingeniería → ABB PMS",
      plan="El informe de avance Nº 4 de ABB (30-09-2026) lista tres definiciones pendientes: (a) señales cableadas de cada equipo con indicación de las borneras de vinculación al PMS (transformadores, generadores, seccionador bajo carga); (b) mapa Modbus y definición de señales de los equipos a integrar (UPS, generadores, multimedidores); (c) mapas Profinet y definición de señales de UMC y VFD.",
      riesgo="La configuración de variables de comunicación (IEC61850 / PROFINET / MODBUS) y de las lógicas del sistema arranca el 02-11-2026 y dura hasta el 11-12-26. Sin esas definiciones ABB no puede configurar ni dejar listo el sistema para las pruebas internas del 18-12-26.",
      impacto="Es un insumo de AESA, no un atraso del proveedor. Si no se cierra en octubre, se comprime la configuración y se pone en riesgo la cadena pruebas internas (18-12) → FAT (25-29/01/27) → despacho (17-02-27), que hoy no tiene holgura.",
      real="Las definiciones deberían estar cerradas antes del 02-11-2026, fecha de inicio de la configuración de comunicaciones.",
      accion="Asignar responsable en Ingeniería por cada uno de los tres paquetes y acordar fecha de entrega con ABB en la reunión semanal de seguimiento. Priorizar el mapa Modbus y los mapas Profinet, que dependen de información de terceros (UPS, generadores, UMC, VFD).",
      sev="ALTA"),
 dict(cod="PA-12", tit="Una sola FAT integral de 5 días frente a las dos ventanas de ensayos PMS del cronograma de salas",
      inter="ABB PMS ↔ ABB Salas ↔ AESA",
      plan="El informe del PMS indica que ABB y AESA acordaron una FAT INTEGRAL de las salas en planta shelterista, de 5 días, y el cronograma del PMS la ubica del 25 al 31-03-2027. El informe Nº 5 de las salas, en cambio, prevé DOS FAT de sala separadas en Mendoza: SE#4 del 22 al 26-02-2027 y SE#3 del 29-03 al 02-04-2027, más dos ventanas de «Ensayos PMS en sala» (SE#4 03-09/03 y SE#3 30/03-06/04/27).",
      riesgo="La FAT integral del PMS (25-31/03/27) coincide con la FAT de SE#3 en Mendoza (29/03-02/04/27), de modo que para la sala BT ambas lecturas convergen. Pero NO para SE#4: su FAT es del 22 al 26-02-27 y su despacho EXW el 25-03-27, un mes ANTES de la FAT integral del PMS.",
      impacto="Si prevalece una única FAT integral del 25 al 31-03-27, la sala MT ya estaría despachada y no podría participar. Si prevalecen dos campañas, ABB PMS debe rever el alcance y la duración de 5 días acordados, y prever una asistencia en febrero que hoy no tiene cargada.",
      real="A conciliar entre ABB PMS, ABB Salas y AESA antes de que ABB emita el protocolo de ensayos.",
      accion="Llevar el punto a la reunión semanal de seguimiento ABB/AESA y fijar el criterio antes de la emisión del protocolo: una campaña integral o dos por sala, y su encaje con los despachos EXW de cada shelter.",
      sev="ALTA"),
 dict(cod="PA-09", tit="Las pruebas FAT de los transformadores no están incluidas en el cronograma de CAT Miron",
      inter="CAT Miron → AESA / PPSA",
      plan="El cronograma 5155-00-0022-VE-CO-PG-001 Rev.A aclara textualmente: «No incluye pruebas FAT, las cuales se deberán coordinar según disponibilidad de agenda de Laboratorio cercano a esa fecha». El plan sólo contempla 2 días hábiles de ensayos internos por orden de trabajo, entre el 02-11 y el 04-12-26.",
      riesgo="La FAT se intercala entre los ensayos internos y el acondicionamiento para despacho, pero su fecha no está fijada y depende de la agenda de un laboratorio externo al fabricante. Con 5 órdenes de trabajo que se liberan escalonadamente en 5 semanas, se necesitan 5 ventanas de laboratorio en noviembre y diciembre.",
      impacto="Si el laboratorio no tiene disponibilidad en la ventana prevista, el despacho de cada lote se corre sin que el cronograma lo muestre. Es el único hito de esta provisión sin fecha comprometida y está inmediatamente antes de la entrega.",
      real="A definir con CAT Miron y el laboratorio. Se requiere una fecha por cada una de las 5 OT.",
      accion="Pedir a CAT Miron que reserve las ventanas de laboratorio y emita una revisión del cronograma con las fechas de FAT incorporadas. Definir el alcance del witness (AESA / PPSA) y emitir los protocolos de ensayo con anticipación.",
      sev="ALTA"),
 dict(cod="PA-10", tit="Los transformadores llegan siete meses antes que los ductos de barras que los vinculan",
      inter="CAT Miron ↔ ABB Ductos de barras ↔ AESA Construcciones",
      plan="CAT Miron entrega los 13 transformadores entre el 13-11 y el 11-12-2026. Los ductos de barras que los vinculan a las salas eléctricas entregan en sitio el 08-07-2027 (ver PA-07).",
      riesgo="Desfase de aproximadamente 7 meses entre ambas entregas. Los transformadores quedan almacenados en obra sin poder vincularse, durante el verano y buena parte del otoño.",
      impacto="Requiere prever superficie de almacenamiento, fundaciones o playa de acopio, y sobre todo un plan de conservación de transformadores almacenados (control de humedad y nivel de aceite o del gas de presurización según el tipo, verificación periódica y registro). Sin ese plan hay riesgo de observaciones en la recepción y de ensayos de re-verificación antes de energizar.",
      real="Almacenamiento efectivo: de nov-dic 2026 a jul 2027 como mínimo.",
      accion="Definir con Construcciones el sitio y las condiciones de acopio; solicitar a CAT Miron el procedimiento de almacenamiento prolongado y su garantía en esa condición; incorporar las verificaciones periódicas al plan de calidad. Evaluar si conviene escalonar la recepción en lugar de recibir las 13 unidades entre noviembre y diciembre.",
      sev="MEDIA"),
 dict(cod="PA-07", tit="Ductos de barras: entrega más tardía del paquete, plazo sin confirmar y pedido de ABB de adecuar la fecha contractual",
      inter="AESA → ABB",
      plan="La entrega en sitio pasó del 10-06-27 al 08-07-27. El informe de avance Nº 5 de ABB (05-10-26) explicita la causa: los ductos NO PUDIERON DISEÑARSE porque no se disponía de información suficiente sobre los gabinetes de los transformadores de media tensión y sus posiciones relativas. El diseño arrancó efectivamente el 21-09-26, al confirmarse esa información.",
      riesgo="ABB declara que está trabajando en la confirmación del plazo de entrega y del POSIBLE IMPACTO ECONÓMICO, y pide formalmente que AESA ADECUE LA FECHA DE ENTREGA CONTRACTUAL de las posiciones 120, 130 y 140 de la OC 45089944971 para evitar penalidades por incumplimiento. Además, el subproveedor de los ductos —EAE, Turquía, documento de compra 4700201354— figura con fecha de entrega en origen «tbd», es decir sin comprometer.",
      impacto="La entrega en sitio del 08-07-27 es el hito más tardío de TODO el paquete, por delante del cierre del alcance HARG (10-05-27), y hoy ni siquiera está firme: el plazo definitivo depende de la confirmación de ABB y de EAE. Hay además una exposición contractual y económica abierta. El avance del paquete es del 5% (ingeniería 15%, fabricación y transporte en 0%).",
      real="A confirmar por ABB. El transporte marítimo pesa el 34% del paquete y el traslado + aduana insume 80 días corridos, sin margen de compresión.",
      accion="(1) Responder formalmente el pedido de adecuación de la fecha contractual de las posiciones 120/130/140 y acotar la exposición económica. (2) Exigir a ABB la fecha comprometida de EAE, hoy «tbd». (3) Proteger las dos aprobaciones que quedan: el diseño cierra el 14-12-26 y la aprobación de AESA vence el 30-12-26, en pleno receso — asignar el revisor con anticipación. (4) Trazar la información de gabinetes de transformadores con CAT Miron para que no se repita.",
      sev="ALTA"),
 dict(cod="PA-08", tit="Aprobación de la ingeniería de las SALAS vencida: el freezing point de ingeniería básica venció el 25-09-26 y sigue sin cerrar",
      inter="AESA → ABB (Salas SE#3 y SE#4)",
      plan="BUENA NOTICIA PRIMERO: el informe Nº 5 confirma que en septiembre se APROBÓ la ingeniería básica y de detalle de TODOS LOS EQUIPOS de ambas salas (CCM, tableros MT, auxiliares, sistema UPS/baterías y variadores de MT y BT). El punto se acota ahora al alcance de la SALA propiamente dicha (el shelter): su ingeniería básica figura emitida al 100% pero aprobada sólo al 90%, con freezing point vencido el 25-09-26, y la de detalle emitida al 95% con aprobación a cerrar el 13-10-26.",
      riesgo="Al corte hay 11 días de atraso sobre el freezing point de ingeniería básica de las salas, y la certificación del hito 2 —que ABB ata a esa aprobación— está prevista recién para octubre. La fabricación del shelter de SE#4 arranca en octubre y la de SE#3 el 23-11-26.",
      impacto="Acotado: la ingeniería de los equipos, que gobierna la fabricación en Brasil y los slots de fábrica, ya está cerrada. Lo que queda en riesgo es el arranque de la fabricación del shelter en el shelterista y la certificación del hito 2.",
      real="Aprobación de ingeniería básica de salas: vencida el 25-09-26, a cerrar de inmediato. Aprobación de la de detalle: 13-10-2026.",
      accion="Cerrar esta semana la aprobación de la ingeniería básica de ambas salas para liberar la certificación del hito 2 y el arranque de la fabricación del shelter, y asegurar el circuito de revisión de la ingeniería de detalle para el 13-10-26.",
      sev="ALTA"),
 dict(cod="PA-04", tit="Orden de compra de HIMA Argentina (HARG) no colocada al 29-09-2026",
      inter="AESA → HARG",
      plan="El cronograma PSCH_CPF2_Rev3 arrancaba los trabajos de HARG el 14-09-2026. A la fecha de corte (29-09-2026) la OC sigue sin colocarse, por lo que el inicio se mantiene corrido a principios de noviembre (02-11-2026). Transcurrió una semana más desde la Rev4 sin novedades.",
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
 dict(cod="PA-03", tit="Prueba de integración PCS↔PMS en Buenos Aires — fecha ya fijada, resta confirmar duración y protocolo",
      inter="Inauco (PCS) ↔ ABB (PMS)",
      plan="Anticipada en la Rev3 a la ventana OCT-NOV 2026 sin fecha. El informe de ABB del 30-09-26 ya la fija.",
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
 (11,"AESA→HARG","COLOCACIÓN DE LA OC DE HIMA ARGENTINA",D(2026,9,29),D(2026,10,30),"—","⚠ PA-04 — PENDIENTE al 29-09-26; gobierna todo lo que sigue"),
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
# ---------------- Informe de avance a SEP-2026 (Rev7) ----------------
# ABB PMS — informe Nº 4 (E-2611027, 30-09-2026): avance físico ponderado por especialidad
PMSK="ABB — PMS (Power Management System)"
# ABB SALAS — informe de avance Nº5 (R-2631036, 05-10-2026): avance ponderado por paquete
AVANCE_SALAS=[  # (paquete, ago-26, sep-26, oct-26 previsto)
 ("Tableros CCM MNS III — 102-CCM-005/6/7","0%","69%","84%"),
 ("Tableros de media tensión UNIGEAR — 102-TGMT-001/2","0%","40%","48%"),
 ("Tableros auxiliares de distribución","0%","66%","75%"),
 ("Sistema UPS y baterías","0%","49%","61%"),
 ("Variadores de velocidad de BT — 102-DP-VFD-24210A/B/C","0%","68%","71%"),
 ("Variadores de velocidad de MT — 102-DP-VFD-23310A/B/C","0%","68%","71%"),
 ("Sala eléctrica 3 (shelter)","0%","36%","40%"),
 ("Sala eléctrica 4 (shelter)","0%","36%","43%"),
 ("Ductos de barras — 102-DB-005/006/007 A/B","0%","5%","30%"),
]
# previsión de fechas y lugares de FAT (informe Nº5)
FAT_PREV=[
 ("Sistema UPS y baterías","Argentina","12/11/2026","13/11/2026","Santa Fe — Argentina"),
 ("Tableros TGMT (media tensión)","Brasil","27/11/2026","03/12/2026","Sorocaba — Brasil"),
 ("Tableros CCM","Brasil","03/12/2026","11/12/2026","Sorocaba — Brasil"),
 ("Tableros auxiliares","Argentina","14/12/2026","16/12/2026","AMBA — Buenos Aires"),
 ("Sala eléctrica #4 (shelter completo)","Argentina","22/02/2027","26/02/2027","Mendoza — Argentina"),
 ("Sala eléctrica #3 (shelter completo)","Argentina","29/03/2027","02/04/2027","Mendoza — Argentina"),
]
# órdenes de compra de ABB a subproveedores (informe Nº5)
SUBPROV=[
 ("4500407392","Salas eléctricas 3 y 4","18/05/2026","Bottino","Argentina","14/05/27 y 08/04/27"),
 ("4500407396","Tableros auxiliares","18/05/2026","Mehcco","Argentina","28/01/2027"),
 ("4500407468","KMT","19/05/2026","KMT","Argentina","27/11/2026"),
 ("4700201354","EAE (ductos de barras)","16/05/2026","EAE","Turquía","tbd — SIN COMPROMETER"),
 ("4700201355","Tableros de media tensión","16/05/2026","BRABB","Brasil","27/11/2026"),
 ("4700201478","Tableros CCM","19/05/2026","BRABB","Brasil","25/01/2027"),
 ("4700201819","Variadores de baja tensión","26/05/2026","FIABB","Finlandia","06/11/2026"),
 ("4700201626","Variadores de media tensión","22/05/2026","CNABB","China","06/11/2026"),
]
SALAS_TXT=[
 ("Hito del período","Se consiguió la APROBACIÓN DE LA INGENIERÍA DE DETALLE de todos los equipos de las salas eléctricas: tableros auxiliares, tableros CCM, tableros de MT, sistema UPS y baterías, y variadores de media y baja tensión. Queda pendiente la aprobación de la ingeniería de las salas propiamente dichas (shelter). Ver PA-08."),
 ("Fechas contractuales vs objetivo","ABB presentó la versión optimista del cronograma: SALA 3 fecha contractual 14-05-2027 contra objetivo 23-04-2027 (mejora de 21 días); SALA 4 fecha contractual 08-04-2027 contra objetivo 25-03-2027 (mejora de 14 días). El cronograma del 02-10-26 ubica la entrega EXW de SE#3 el 20-04-27 — tres días antes del objetivo informado."),
 ("Ductos de barras","ABB verifica un corrimiento de la fecha contractual de entrega de los ductos de la sala #3: no pudieron diseñarse por falta de información sobre los gabinetes de los transformadores de MT y sus posiciones relativas. El diseño arrancó el 21-09-26. ABB evalúa plazo e impacto económico y pide a AESA adecuar la fecha contractual de las posiciones 120, 130 y 140 de la OC 45089944971 para evitar penalidades. Ver PA-07."),
 ("Situación contractual","Pólizas de fiel cumplimiento y fondo de reparo de ambas OC emitidas por ABB y aprobadas por AESA. Acordadas las aclaraciones sobre los hitos de certificación y revisados los documentos de compra. Nota de cambio Nº 1 (R-2631036-NC1_A) aprobada por AESA, con nota de pedido recibida en agosto; ABB avanzó con la implementación técnica."),
 ("Fabricación iniciada","Variadores de media tensión: fabricación y ensayos INICIADOS Y FINALIZADOS. Tableros CCM: inicio temprano de fabricación de columna y módulos extraíbles. Tableros de MT: inicio de fabricación de gabinetes de comando."),
 ("Acopio del sistema UPS y baterías","Fusibles 100% · ventiladores 100% · interruptores y fusibles 98% · magnéticos y semiconductores 95% · baterías 95% · disipadores y capacitores 90% · gabinetes y accesorios 85% · breakers 80% · TI 75% · placas electrónicas 75% · semiconductores 63%."),
 ("Calidad e inspección","Emitida la documentación de calidad: 16 protocolos de FAT APROBADOS (salas 3 y 4, CCM, TGMT, sistema CC, UPS y tableros auxiliares), 6 planes de inspección y ensayos (4 aprobados, 2 aprobados con comentarios) y los procedimientos de END, pintura, FAT por equipo y soldadura. Para octubre: inspección al proveedor de salas eléctricas."),
 ("Eventos previstos para octubre","Certificación parcial del hito 2 por aprobación de la ingeniería básica de salas eléctricas · certificación del hito 3 por acopio del sistema UPS/baterías y tableros auxiliares · emisión y aprobación de la nota de cambio por equipos adicionales de CCTV y Wifi · inspección al proveedor de salas · avance de fabricación de los CCM · inicio de fabricación de los tableros de MT · inicio de fabricación de la sala eléctrica 4."),
 ("Tareas adicionales detectadas","Se modificaron los listados de cargas de los tableros 102-CCM-006 y 102-CCM-007: nota de cambio emitida y aprobada en julio. Se solicitó ampliar la provisión de equipos de CCTV y Wifi: ABB evalúa el impacto con sus subproveedores."),
 ("Temas pendientes prioritarios","Único pendiente imputado a AESA: suministro de los equipos físicos de CONTROL DE ACCESO de acuerdo con el cronograma de cada sala. Ver PA-13."),
 ("Certificación y cobranzas","OC 45089944971: certificaciones 1 y 2 (USD 143.589 c/u) PAGADAS; certificación 3 (USD 29.343, vto. 23-10-26) y certificación 4 (USD 376.804, vto. 28-10-26) pendientes de cobro. Total USD 693.326. OC 45089944973: certificaciones 1 y 2 (USD 170.739 c/u) PAGADAS; certificación 3 (USD 29.343) y certificación 4 (USD 48.031) pendientes. Total USD 418.853."),
]
AVANCE_PMS=[  # (especialidad, peso, ago-26 real, sep-26 previsto, sep-26 real, oct-26 previsto, oct-26 pronóstico)
 ("Ingeniería",            "20%","60%","90%","90%","95%","95%"),
 ("Colocación de subórdenes","5%","100%","100%","100%","100%","100%"),
 ("Recepción de materiales","30%","40%","70%","100%","100%","100%"),
 ("Fabricación",           "20%","0%","0%","0%","0%","0%"),
 ("Configuración",         "10%","0%","0%","0%","0%","0%"),
 ("Tests / Inspección",    "10%","0%","0%","0%","0%","0%"),
 ("Despacho",               "5%","0%","0%","0%","0%","0%"),
 ("TOTAL GLOBAL",         "100%","29%","44%","53%","54%","54%"),
]
# avance declarado por cada proveedor en su propio cronograma
AVANCE_PROV=[  # (provisión, fuente y fecha, avance declarado, observación)
 ("ABB — Sala Eléctrica SE#3 (BT)","Informe de avance Nº5 · 05-10-2026","36%","Shelter. Ingeniería 100%, recepción de materiales 35%, fabricación 0%. Previsto oct-26: 40%. (El cronograma declara 57% de tareas completas, criterio distinto.)"),
 ("ABB — Sala Eléctrica SE#4 (MT)","Informe de avance Nº5 · 05-10-2026","36%","Shelter. Previsto oct-26: 43%. (El cronograma declara 50%.) Los equipos van mucho más adelantados: CCM 69%, variadores 68%, auxiliares 66%."),
 ("ABB — Ductos de barras","Informe de avance Nº5 · 05-10-2026","5%","Ingeniería 15%, fabricación y transporte marítimo en 0%. Previsto oct-26: 30%. Plazo de entrega SIN CONFIRMAR — ver PA-07."),
 ("ABB — PMS (Power Management System)","Informe de avance Nº4 · 30-09-2026","53%","Avance físico ponderado. Previsto 44% → +9 puntos. (El cronograma declara 66% de tareas completas, criterio distinto.)"),
 ("CAT Miron — Transformadores","Cronograma 29-09-2026","parcial","Ingeniería 100% (OT 5211 al 75%); conductores 40-90%; materiales gruesos 30-50%; fabricación 0%"),
 ("Inauco — PCS / SCADA / Comunicaciones","Crono. fabricación PG-002 Rev.0 · 14-09-2026","0%","Todas las tareas de taller en estado «Pendiente»; el armado inicia el 27-11-26"),
 ("HPH — HIMA Paul Hilpert (Alemania)","Schedule 16-08-2026","sin dato","El cronograma no declara porcentajes de avance"),
 ("HARG — HIMA Argentina","PSCH_CPF2_Rev3 · 04-09-2026","0%","OC no colocada — los trabajos no iniciaron"),
]
# avance real por hito al 30-09-2026 (informe ABB Nº4 + cronograma SEP-2026)
SE3K="ABB — Sala Eléctrica SE#3 (BT)"; SE4K="ABB — Sala Eléctrica SE#4 (MT)"; DUCK="ABB — Ductos de barras"
AVANCE_REAL={
 # --- ABB Salas: informe Nº5 (05-10-2026) ---
 (SE3K,"Emisión de ingeniería básica"):("100%","Cumplido"),
 (SE3K,"Aprobación de ingeniería básica"):("90%","VENCIDA — no cumplida"),
 (SE3K,"FREEZING POINT — ingeniería básica"):("0%","VENCIDO — no cumplido"),
 (SE3K,"Emisión de ingeniería de detalle"):("95%","En curso"),
 (SE3K,"Aprobación de ingeniería de detalle — EQUIPOS"):("100%","APROBADA en septiembre"),
 (SE3K,"Aprobación de ingeniería de detalle — SALA"):("0%","Pendiente — vence 13-10-26"),
 (SE3K,"ACOPIO — Materiales de los CCM (críticos y secundarios)"):("93%","En curso"),
 (SE3K,"ACOPIO — Rectificadores / UPS / baterías"):("93%","En curso"),
 (SE3K,"ACOPIO — Tableros auxiliares"):("72%","En curso"),
 (SE3K,"Inicio de fabricación en fábrica de origen"):("2%","Iniciada — CCM"),
 (SE3K,"Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)"):("0%","Pendiente — shelter al 36%"),
 (SE4K,"Emisión de ingeniería básica"):("100%","Cumplido"),
 (SE4K,"Aprobación de ingeniería básica"):("90%","VENCIDA — no cumplida"),
 (SE4K,"FREEZING POINT — ingeniería básica"):("0%","VENCIDO — no cumplido"),
 (SE4K,"Emisión de ingeniería de detalle"):("95%","En curso"),
 (SE4K,"Aprobación de ingeniería de detalle — EQUIPOS"):("100%","APROBADA en septiembre"),
 (SE4K,"Aprobación de ingeniería de detalle — SALA"):("0%","Pendiente — vence 13-10-26"),
 (SE4K,"ACOPIO — Celdas MT (materiales principales y secundarios)"):("78%","En curso"),
 (SE4K,"ACOPIO — Variadores MT (China) y BT (Finlandia)"):("100%","Fabricación y ensayos finalizados"),
 (SE4K,"ACOPIO — Tableros auxiliares"):("72%","En curso"),
 (SE4K,"Inicio de fabricación en fábrica de origen"):("0%","Gabinetes de comando iniciados"),
 (SE4K,"Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)"):("0%","Inicio previsto en octubre"),
 (DUCK,"Emisión layout instalación + detalle conexionado trafo"):("100%","Cumplido"),
 (DUCK,"Diseño de ductos de barras"):("15%","En curso — arrancó el 21-09-26"),
 # --- ABB PMS: informe Nº4 (30-09-2026) ---
 (PMSK,"Inicio de contrato — Recepción / aceptación OC"):("100%","Cumplido"),
 (PMSK,"Kick-off Meeting (KOM) y planificación detallada"):("100%","Cumplido"),
 (PMSK,"Emisión de ingeniería básica (documentación para aprobación)"):("100%","Cumplido · certificado"),
 (PMSK,"Aprobación de ingeniería básica (AESA)"):("100%","Cumplido · certificado"),
 (PMSK,"Emisión de ingeniería Rev.0"):("100%","Cumplido"),
 (PMSK,"Aprobación ingeniería Rev.0 — FREEZING POINT"):("100%","Cumplido"),
 (PMSK,"ACOPIO — Inicio de compra de materiales"):("100%","Cumplido"),
 (PMSK,"ACOPIO — Material S800 y tableros"):("100%","Compra cerrada"),
 (PMSK,"ACOPIO — Switches de red"):("100%","Compra cerrada"),
 (PMSK,"Configuración / programación del sistema"):("32%","En curso"),
 (PMSK,"INICIO DE ARMADO DE TABLEROS"):("0%","Pendiente"),
 (PMSK,"Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF1"):("0%","Pendiente"),
 (PMSK,"Emisión de procedimientos de FAT y SAT (Rev.A)"):("0%","Pendiente — previsto octubre"),
 (PMSK,"Configuración de variables de comunicación y lógicas del sistema"):("0%","Pendiente — requiere insumos AESA"),
}
AVANCE_TXT=[
 ("Avance del período (ABB PMS)","Avance físico real de septiembre 53% contra 44% previsto: 9 puntos POR ENCIMA del plan. El salto lo explica la recepción de materiales, que cerró al 100% cuando estaba prevista al 70%."),
 ("Eventos cumplidos en septiembre","Reunión integral de revisión y análisis del alcance entre Pluspetrol, ABB y AESA. Definiciones sobre los vínculos de fibra óptica para vincular las nuevas salas SE#3 y SE#4 con las salas existentes. Emisión de ingeniería: descripción funcional Rev.1, topología del sistema Rev.0, listado de señales y comunicaciones Rev.0, cálculo de consumo de potencia Rev.0, cálculo de carga de controladores Rev.0, layout y diagrama de conexionado y listado de partes de los tableros de control/IO de SE#3 y SE#4 Rev.0, especificaciones técnicas de equipos ABB y de terceros Rev.1, y planes de inspección y ensayos 102-PMS-001 y 102-PMS-101 Rev.0."),
 ("Eventos previstos para octubre","Desarrollo de la ingeniería básica restante: protocolo de aceptación en fábrica (FAT) y su registro de pruebas, y protocolo de aceptación en sitio (SAT) y su registro. Reuniones semanales de seguimiento técnico AESA / ABB / PPSA."),
 ("Acuerdos del período","(1) ABB y AESA acordaron realizar la FAT INTEGRAL de las salas en planta shelterista, con una duración de 5 días y verificando una muestra representativa de señales; el protocolo lo emitirá ABB en los próximos meses. (2) Se acordaron pruebas anticipadas con la maqueta de CPF1 y los materiales/cubicles de UMC propios de CPF2, en conjunto ABB / AESA / PLUSPETROL e INAUCO por la comunicación con el PCS, estimadas para la semana del 30 de noviembre de 2026 con 2 días de duración."),
 ("Temas pendientes prioritarios (insumos de AESA)","(1) Definiciones de señales cableadas en cada equipo, con indicación de las borneras de vinculación al PMS: transformadores, generadores, seccionador bajo carga. (2) Mapa Modbus y definición de señales de los equipos a integrar: UPS, generadores, multimedidores. (3) Mapas Profinet y definición de señales de UMC y VFD.  Ver PA-11."),
 ("Calidad / Inspección y Fabricación","Sin actividades en el período y sin actividades previstas para octubre: la fabricación arranca con el armado de tableros el 09-10-2026."),
 ("Factores potenciales de retraso","ABB señala posibles cambios de alcance, indefiniciones y otros factores no detectados al momento de redactar el informe. No se detectaron tareas adicionales dentro del alcance de la OC."),
 ("Situación contractual y certificación","Sin avances en la certificación ni cambios contractuales en el período. Certificación nº 1 presentada el 27-08-2026 por USD 176.044,58 + impuestos; sin certificaciones por avance de obra al cierre del informe. Plan de certificación: envío de ingeniería básica 10% (31-07-26, cumplido) · aprobación de ingeniería básica 10% (14-08-26, cumplido) · acopio de materiales 25% (05-11-26) · FAT 20% (29-01-27) · entrega de equipamiento en shelter 30% (17-02-27) · entrega del DataBook 5% (31-03-27)."),
 ("Fuera del alcance","Módulo adicional de entradas (oferta OPP-26-8808080 Rev.0, típico DOL_TIPICO-01A): AESA aceptó el 23-07-26 avanzar con la gestión interna de ABB para no impactar las pruebas de integración de nov/dic. Su gestión y ejecución no forman parte de esta OC — verificar que esté disponible para la maqueta CPF2."),
 ("Provisiones sin informe de avance","Hoy sólo ABB presenta informe mensual de avance, y lo hace por separado para Salas Eléctricas (Nº 5) y para el PMS (Nº 4). Para INAUCO, CAT MIRON y HIMA (HPH y HARG) se toma el porcentaje declarado en el propio cronograma del proveedor, con criterios de medición distintos y no comparables entre sí. Conviene pedirles informe de avance mensual con el mismo formato ponderado que usa ABB."),
]

NOTAS=[
 "Alcance: cronograma EJECUTIVO por hitos, con corte en la ENTREGA DE EQUIPOS / del alcance de cada provisión. Las tareas posteriores al despacho/entrega (montaje en planta, SAT, precomisionado, comisionado, PEM, CAO/DataBook) quedan FUERA.",
 "Rev7 — cambio 1: ABB PMS — se adopta el cronograma 5155-00-0007-VE-CO-PG-001 SEP-2026 (REV 0, 01-10-2026) y el informe de avance Nº 4 (E-2611027, 30-09-2026). El ARMADO DE TABLEROS se corre 4 semanas (11-09 → 09-10-26) pero el fin de armado (17-12-26) y la ENTREGA EN SHELTER (17-02-27) se sostienen: ABB absorbió el corrimiento.",
 "Rev7 — cambio 2: la GESTIÓN DE COMPRA DE MATERIALES del PMS pasa al 100% (switches de red, material S800 y tableros). La recepción de materiales cerró en septiembre al 100% contra un 70% previsto, y es lo que explica que el avance del mes supere al plan.",
 "Rev7 — cambio 3: las pruebas de maqueta integrada se desdoblan en MAQUETA CPF1 (19-23/10/26) y MAQUETA CPF2 (30/11 → 03/12/26, con INAUCO por la comunicación con el PCS). Queda fijada así la prueba de integración PCS↔PMS que en revisiones anteriores figuraba «a convenir»: PA-03 baja a severidad BAJA.",
 "Rev7 — cambio 4: nueva hoja «Avance a SEP-2026» con el avance ponderado del informe de ABB y el avance declarado por cada proveedor en su cronograma. Se suman PA-11 (definiciones que ABB reclama a AESA) y PA-12 (una FAT integral de 5 días frente a las dos ventanas de ensayos PMS del cronograma de salas).",
 "Rev6 — cambio 1: se incorpora la provisión CAT MIRON — TRANSFORMADORES (13 unidades en 5 órdenes de trabajo), según el cronograma de fabricación 5155-00-0022-VE-CO-PG-001 / ACAL-00102-EMBR-TRSEC-PP-X-0001 Rev.A (emitido 01-09-26, actualización del 29-09-26). Entregas escalonadas: OT 5209 (4×2000 kVA) y OT 5210 (2×1600 kVA) el 13-11-26 · OT 5212 (2×2000 kVA) el 27-11-26 · OT 5213 (3×500 kVA) el 04-12-26 · OT 5211 (2×2500 kVA) el 11-12-26.",
 "Rev6 — cambio 2: se suman PA-09 (las pruebas FAT de los transformadores NO están incluidas en el cronograma del proveedor y dependen de la agenda de un laboratorio externo) y PA-10 (los transformadores llegan 7 meses antes que los ductos de barras que los vinculan).",
 "CAT Miron — aclaraciones del emisor: los días del programa son DÍAS HÁBILES, no corridos; el acopio de «materiales gruesos» incluye los gabinetes de los transformadores; el cronograma se basa en compromisos con proveedores externos que pueden variar, en cuyo caso MIRON informará y emitirá una nueva revisión.",
 "CAT Miron — la fecha de la orden de compra no está consignada en el cronograma del proveedor: se toma como inicio el arranque de la ingeniería del primer lote (13-07-2026). A confirmar con Suministros.",
 "Rev5 — cambio 1: se adopta la actualización del cronograma ABB 5155-00-0009-VE-CO-PG-001_A del 21-09-2026, que reemplaza la del 21-08-2026. Se compararon ambos documentos tarea por tarea: 127 diferencias sobre 218 tareas.",
 "Rev5 — cambio 2: DUCTOS DE BARRAS — la entrega en sitio se atrasa 4 semanas (10-06-27 → 08-07-27) y pasa a ser el hito más tardío de todo el paquete. Origen: el insumo de AESA (layout + detalle de conexionado a trafo) se emitió el 18-09-26 en lugar del 29-08-26. Ver PA-07.",
 "Rev5 — cambio 3: SALAS SE#3 y SE#4 — los freezing points de ingeniería se corren (básica 11-09 → 25-09-26 · detalle 05-10 → 13-10-26), pero ABB mantiene las entregas: SE#3 ADELANTA su despacho EXW del 23-04 al 20-04-27 y SE#4 mantiene el 25-03-27. Ver PA-08.",
 "Rev5 — cambio 4: se agregan los hitos de FABRICACIÓN DEL SHELTER en ambas salas (SE#3 atrasada 3 semanas, SE#4 8 días) y de APROBACIÓN DE DISEÑO en los ductos de barras, que pasaron a ser determinantes. El PMS no registra cambios.",
 "Rev4 — cambio 1: se incorpora el cronograma de fabricación de INAUCO 5155-00-0029-VE-CA-PG-002 Rev.0 (14-09-2026), con el detalle por tablero (CPF2-PCS-SC-07a de 3 gabinetes, CPF2-PCS-SC-07b de 2 gabinetes, y los racks CPF-COM-07 y CPF-COM-SE3), sus PreFAT internos de taller y la liberación de pendientes.",
 "Rev4 — cambio 2: se adoptan las ACLARACIONES del pie del cronograma de Inauco: PRE iFAT del 10-02 al 26-03-27 (hito nuevo), FAT con el cliente del 29-03 al 22-04-27 (confirma la ventana de la Rev3) y LIBERACIÓN DE TABLEROS el 04-05-27 en camión de AESA — la entrega de Inauco se adelanta 13 días respecto de la Rev3 (17-05-27).",
 "Rev4 — cambio 3: HARG (HIMA Argentina) — al 29-09-2026 NO está colocada la orden de compra, por lo que el inicio de los trabajos se corre a principios de noviembre (02-11-2026). Se aplicó un corrimiento de 49 días corridos a todas las actividades PROPIAS de HARG; los hitos que son insumo de HPH (ya contratada) se mantienen en su fecha original.",
 "Rev4 — cambio 4: el CAMINO CRÍTICO se recalculó. Ya no cierra con la entrega de Inauco sino con la ENTREGA DEL ALCANCE HARG el 10-05-2027, que queda 6 días DESPUÉS de la liberación de los tableros de Inauco (04-05-27). Se incorporó la colocación de la OC de HARG como eslabón explícito de la cadena.",
 "ALERTA — PA-04: la OC de HIMA Argentina es hoy el ítem de mayor impacto del paquete: cada semana sin colocar corre una semana el cierre del sistema de seguridad. PA-05: la adecuación de HARG (08-02 → 19-03-27) se superpone con la Pre iFAT de Inauco y su IFAT termina después de la FAT SIS conjunta.",
 "Nota favorable: el corrimiento de HARG deja 2 semanas de holgura entre la recepción de los tableros de HPH en base Inauco (25-01-27) y el inicio de la adecuación (08-02-27), holgura que en la Rev3 no existía. Por eso PA-01 baja de severidad ALTA a MEDIA. Esa holgura no debe consumirse para absorber otros atrasos.",
 "Fechas no laborables declaradas por Inauco para la fabricación: 07 y 08-12-26 · 24 y 25-12-26 · 31-12-26 · 01-01-27 · 08 y 09-02-27. Días laborables de fabricación: 50.",
 "Reserva del emisor (crono. PG-002 Rev.0): las fechas son una estimación sobre las condiciones actuales y quedan sujetas a reprogramación por demoras en insumos, disponibilidad de stock de tarjetas/equipamiento de instrumentación y tiempos de importación o logística de componentes comerciales de alta complejidad.",
 "Rev3 — cambio: hito explícito de NACIONALIZACIÓN / DESPACHO DE ADUANA en HPH y anticipación de la prueba de integración PCS↔PMS a la ventana OCT-NOV 2026 (fecha a convenir).",
 "Rev2 — cambios: FAT EN FÁBRICA DE ORIGEN (Brasil, M&B) en las Salas SE#3 y SE#4; apertura del ACOPIO DE MATERIALES en hitos propios con hoja dedicada; separación de HIMA en HPH (Alemania) y HARG (Argentina).",
 "Fuentes: ABB PMS 5155-00-0007-VE-CO-PG-001 SEP-2026 REV 0 (01-10-26) + informe de avance Nº4 E-2611027 (30-09-26) · CAT Miron 5155-00-0022-VE-CO-PG-001 Rev.A (act. 29-09-26) · ABB Salas y Ductos 5155-00-0009-VE-CO-PG-001_A (21-09-26) · ABB PMS 5155-00-0007-VE-CO-PG-001_REV B (.mpp) · Inauco 5155-00-0029-VE-CO-PG-001-A Rev.A (24-08-26) y 5155-00-0029-VE-CA-PG-002 Rev.0 (14-09-26) · HPH 103618-Schedule 16-08-26 · HARG PSCH_CPF2_Rev3 (04-09-26) con corrimiento +49 d.",
 "Los hitos de FAT en fábrica de origen (Brasil) son witness de AESA y anteceden a la FAT del shelter completo; no deben confundirse entre sí.",
 "Estado (plan): se calcula contra la fecha de corte. 'Cumplido s/plan' = la fecha planificada ya venció y debe confirmarse el cumplimiento real con el proveedor; 'En curso' = en ejecución; 'Pendiente' = aún no inició.",
 "Las columnas Fecha real / % avance / Estado real / Observaciones quedan en blanco para la carga de seguimiento periódico (reporte a PPSA).",
 "Los hitos se ordenan cronológicamente dentro de cada provisión y se renumeran (H01, H02, ...). Marcas: ◆ = hito del camino crítico · ⚠ = punto de atención · fondo naranja = acopio de materiales.",
]
INTERFASES=[
 "CAMINO CRÍTICO DEL PROYECTO: OC de HARG → HPH (Alemania) → nacionalización → HARG (adecuación, programación, FAT e IFAT) → Inauco (FAT SIS y liberación). Cierra el 10-05-27 con la entrega del alcance HARG. Ver hoja 'Camino crítico'.",
 "AESA → HARG: la COLOCACIÓN DE LA OC está PENDIENTE al 29-09-26. Es el ítem que gobierna la fecha final del sistema de seguridad.  ⚠ PA-04",
 "HPH → entrega DAP Ezeiza (15/12/26) + nacionalización → recepción en base Inauco (25/01/27): habilita la adecuación de HARG, que ahora arranca el 08/02/27 (2 semanas de holgura).  ⚠ PA-01 / PA-02",
 "HARG ↔ Inauco → adecuación de tableros de marshalling (08/02 → 19/03/27) en base Inauco, superpuesta con la Pre iFAT de Inauco (10/02 → 26/03/27).  ⚠ PA-05",
 "HARG ↔ Inauco → el IFAT de HARG (19-30/04/27) termina 4 días después del cierre de la FAT SIS conjunta (26/04/27), y el alcance HARG cierra el 10/05/27, después de la liberación de tableros de Inauco (04/05/27).  ⚠ PA-05",
 "Inauco → FAT con el cliente 29/03 → 22/04/27 y LIBERACIÓN DE TABLEROS el 04/05/27 en camión de AESA (crono. PG-002 Rev.0): habilita el inicio de los trabajos de campo.",
 "ABB PMS ↔ Inauco → prueba de integración PCS↔PMS (maqueta CPF2) en Bs.As. del 30/11 al 03/12/26, con ABB, AESA, PPSA e Inauco. FECHA YA ACORDADA.  ⚠ PA-03",
 "AESA Ingeniería → ABB PMS: definiciones de señales cableadas con borneras de vinculación, mapa Modbus y mapas Profinet. Son insumo de la configuración de comunicaciones, que arranca el 02/11/26.  ⚠ PA-11",
 "ABB PMS ↔ ABB Salas → FAT INTEGRAL del shelter (25-31/03/27, 5 días acordados) frente a las dos ventanas de ensayos PMS del cronograma de salas (SE#4 03-09/03 y SE#3 30/03-06/04/27).  ⚠ PA-12",
 "ABB PMS → entrega de equipamiento en shelter (17/02/27): habilita el montaje del PMS en las salas y los ensayos PMS.",
 "AESA → layout de instalación + detalle de conexionado a trafo: se emitió el 18/09/26 en lugar del 29/08/26 y atrasó 4 semanas los DUCTOS DE BARRAS, cuya entrega en sitio (08/07/27) es hoy la más tardía de todo el paquete.  ⚠ PA-07",
 "AESA → aprobación de ingeniería de las SALAS: los freezing points se corrieron a 25/09/26 (básica) y 13/10/26 (detalle); ABB absorbió el atraso sin mover las entregas, pero ya sin margen.  ⚠ PA-08",
 "AESA → confirmación de fabricación del rack CPF-COM-SE3, hoy marcado 'a confirmar' en el cronograma de Inauco.  ⚠ PA-06",
 "CAT MIRON → ABB Ductos de barras: la ingeniería de detalle de los transformadores (cerrada el 04/09/26) es insumo del detalle de conexionado a trafo que AESA emitió el 18/09/26 y que gobierna el diseño de los ductos.",
 "CAT MIRON → AESA Construcciones: 13 transformadores entregados entre el 13/11 y el 11/12/26, siete meses antes de los ductos de barras (08/07/27). Requiere plan de almacenamiento y conservación.  ⚠ PA-10",
 "CAT MIRON → laboratorio externo: las pruebas FAT no están en el cronograma y dependen de la agenda del laboratorio.  ⚠ PA-09",
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
ws.merge_cells("A3:K3"); cell(ws,3,1,f"Fecha de corte del estado: {HOY.strftime('%d/%m/%Y')}   ·   CAMINO CRÍTICO (sistemas): OC HARG → HPH → HARG → Inauco, cierre 10/05/2027   ·   ÚLTIMA ENTREGA DEL PAQUETE: ductos de barras 08/07/2027   ·   AVANCE a SEP-26: PMS 53% (vs 44% prev.) · Salas SE#3 y SE#4 36% · Ductos 5%   ·   {len(PUNTOS)} puntos de atención abiertos (PA-01 a PA-13)",None,"C62828",True,9)
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
        av=AVANCE_REAL.get((name,hito))
        cell(cw,r,8,"",None,"000000",False,9,True)                       # fecha real
        if av:
            pc=av[0]; ok = pc=="100%"
            cell(cw,r,9,pc,("2E7D32" if ok else ("EF6C00" if pc!="0%" else None)),
                 ("FFFFFF" if pc!="0%" else "555555"),True,9,True)        # % avance real
            cell(cw,r,10,av[1],None,("2E7D32" if ok else "555555"),ok,8.5,True)
        else:
            for k in (9,10): cell(cw,r,k,"",None,"000000",False,9,True)
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
 "El eslabón 11 (COLOCACIÓN DE LA OC DE HIMA ARGENTINA) está PENDIENTE al 29-09-2026 y es hoy el que gobierna la fecha final: cada semana sin colocar corre una semana el cierre del alcance HARG.",
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

# ---------------- Hoja: Avance a SEP-2026 ----------------
vw=wb.create_sheet("Avance a SEP-2026")
vw.merge_cells("A1:G1"); cell(vw,1,1,f"AVANCE A SEPTIEMBRE 2026  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); vw.row_dimensions[1].height=22
vw.merge_cells("A2:G2"); cell(vw,2,1,"Informes de avance de ABB: Nº 5 de Salas Eléctricas (R-2631036, 05-10-2026) y Nº 4 del PMS (E-2611027, 30-09-2026), más el avance declarado por cada proveedor en su propio cronograma.",None,"555555",False,9)
r=4
cell(vw,r,1,"1 · ABB SALAS ELÉCTRICAS — avance ponderado por paquete (informe Nº 5, 05-10-2026)",None,NAVY,True,11); r+=1
SH=["Paquete","ago-26","sep-26 (alcanzado)","oct-26 (previsto)"]
for j,h in enumerate(SH): cell(vw,r,1+j,h,HEAD,"FFFFFF",True,9,True,True)
vw.row_dimensions[r].height=26; r+=1
for (pk,a,sep,oct_) in AVANCE_SALAS:
    bajo = int(sep.rstrip('%'))<40
    cell(vw,r,1,pk,None,"000000",False,9,False,True)
    cell(vw,r,2,a,None,"555555",False,9,True)
    cell(vw,r,3,sep,("C62828" if bajo else "1565C0"),"FFFFFF",True,9.5,True)
    cell(vw,r,4,oct_,None,"555555",False,9,True)
    r+=1
cell(vw,r,1,"ABB rotula la columna de sep-26 como «previsto»; corresponde al avance del mes informado. El total de ago-26 figura en 0% por un artefacto de la planilla de origen.",
     None,"555555",False,8.5,False,True); vw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7); r+=2

cell(vw,r,1,"2 · ABB PMS — avance físico ponderado por especialidad (informe Nº 4, 30-09-2026)",None,NAVY,True,11); r+=1
AH=["Especialidad","Peso","ago-26 real","sep-26 previsto","sep-26 REAL","oct-26 previsto","oct-26 pronóstico"]
for j,h in enumerate(AH): cell(vw,r,1+j,h,HEAD,"FFFFFF",True,9,True,True)
vw.row_dimensions[r].height=28; r+=1
for (esp,peso,a,pv,re,op,opr) in AVANCE_PMS:
    tot = esp.startswith("TOTAL")
    fill = "1F3864" if tot else None; col="FFFFFF" if tot else "000000"
    cell(vw,r,1,esp,fill,col,tot,9); cell(vw,r,2,peso,fill,col,tot,9,True)
    cell(vw,r,3,a,fill,col,tot,9,True); cell(vw,r,4,pv,fill,col,tot,9,True)
    # el real de septiembre se resalta: verde si supera el previsto, rojo si no llega
    try: mejor = int(re.rstrip('%'))>=int(pv.rstrip('%'))
    except Exception: mejor=True
    cell(vw,r,5,re,("2E7D32" if mejor else "C62828"),"FFFFFF",True,9,True)
    cell(vw,r,6,op,fill,col,tot,9,True); cell(vw,r,7,opr,fill,col,tot,9,True)
    r+=1
r+=1
cell(vw,r,1,"Avance real de septiembre 53% contra 44% previsto: 9 puntos POR ENCIMA del plan, por el cierre de la recepción de materiales al 100%.",
     "E8F5E9","2E7D32",True,9.5,False,True); vw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7); r+=2

cell(vw,r,1,"3 · Avance declarado por provisión",None,NAVY,True,11); r+=1
PH=["Provisión","Fuente y fecha","Avance declarado","Observaciones"]
for j,h in enumerate(PH): cell(vw,r,1+j,h,HEAD,"FFFFFF",True,9,True,True)
vw.row_dimensions[r].height=26; r+=1
for (nm,fte,av,obs) in AVANCE_PROV:
    cell(vw,r,1,nm,None,"000000",True,9,False,True)
    cell(vw,r,2,fte,None,"555555",False,8.5,False,True)
    cell(vw,r,3,av,("90A4AE" if av in ("0%","sin dato","parcial") else "1565C0"),"FFFFFF",True,9.5,True)
    cell(vw,r,4,obs,None,"000000",False,8.5,False,True)
    vw.merge_cells(start_row=r,start_column=4,end_row=r,end_column=7)
    vw.row_dimensions[r].height=max(16,13*(1+len(obs)//78)); r+=1
r+=1
cell(vw,r,1,"4 · Previsión de fechas y lugares de FAT — ABB Salas (informe Nº 5)",None,NAVY,True,11); r+=1
FH=["Equipo","Origen","Inicio FAT","Fin FAT","Lugar de la FAT"]
for j,h in enumerate(FH): cell(vw,r,1+j,h,HEAD,"FFFFFF",True,9,True,True)
r+=1
for (eq,org,ini,fin,lug) in FAT_PREV:
    cell(vw,r,1,eq,None,"000000",True,9,False,True); cell(vw,r,2,org,None,"555555",False,9,True)
    cell(vw,r,3,ini,None,"1F3864",True,9,True); cell(vw,r,4,fin,None,"1F3864",True,9,True)
    cell(vw,r,5,lug,None,"000000",False,9,False,True); vw.merge_cells(start_row=r,start_column=5,end_row=r,end_column=7)
    r+=1
cell(vw,r,1,"Son los eventos de witness de AESA (y de PPSA donde corresponda): dos viajes a Sorocaba (Brasil), uno a Santa Fe, uno a AMBA y dos a Mendoza. Conviene reservar agenda y pasajes con anticipación.",
     "FFF9C4","B8860B",True,9,False,True); vw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7); r+=2

cell(vw,r,1,"5 · Órdenes de compra de ABB a subproveedores (informe Nº 5)",None,NAVY,True,11); r+=1
PH2=["Doc. de compra","Descripción","Emisión","Proveedor","Origen","Entrega en origen"]
for j,h in enumerate(PH2): cell(vw,r,1+j,h,HEAD,"FFFFFF",True,9,True,True)
r+=1
for (doc,desc,em,prov,org,ent) in SUBPROV:
    tbd = "tbd" in ent
    cell(vw,r,1,doc,None,"000000",False,9,True); cell(vw,r,2,desc,None,"000000",True,9,False,True)
    cell(vw,r,3,em,None,"555555",False,9,True); cell(vw,r,4,prov,None,"000000",False,9,True)
    cell(vw,r,5,org,None,"555555",False,9,True)
    cell(vw,r,6,ent,("C62828" if tbd else None),("FFFFFF" if tbd else "1F3864"),True,9,True)
    r+=1
r+=1
cell(vw,r,1,"6 · Informe de avance ABB SALAS — síntesis del período",None,NAVY,True,11); r+=1
for (tit,txt) in SALAS_TXT:
    cell(vw,r,1,tit,"ECEFF1","333333",True,9,False,True)
    cell(vw,r,2,txt,None,"000000",False,8.5,False,True)
    vw.merge_cells(start_row=r,start_column=2,end_row=r,end_column=7)
    vw.row_dimensions[r].height=max(16,12*(1+len(txt)//118)); r+=1
r+=1
cell(vw,r,1,"7 · Informe de avance ABB PMS — síntesis del período",None,NAVY,True,11); r+=1
for (tit,txt) in AVANCE_TXT:
    cell(vw,r,1,tit,"ECEFF1","333333",True,9,False,True)
    cell(vw,r,2,txt,None,"000000",False,8.5,False,True)
    vw.merge_cells(start_row=r,start_column=2,end_row=r,end_column=7)
    vw.row_dimensions[r].height=max(16,12*(1+len(txt)//118)); r+=1
for c,w in zip("ABCDEFG",[34,17,15,17,15,17,17]): vw.column_dimensions[c].width=w

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
