#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CRONOGRAMA EJECUTIVO — Provisiones críticas CPF-2 (Rev13, 2026-11-07).

REV13 (07-11-2026) — qué cambió respecto de la Rev12:
  1) HARG (HIMA ARGENTINA): AESA informa que la orden de compra no sale esta semana
     sino la siguiente. Se adopta el 16-11-2026 y el corrimiento de las actividades
     propias pasa de 49 a 63 días corridos. EL CIERRE DEL CAMINO CRÍTICO SE MUEVE DEL
     10-05-2027 AL 24-05-2027.
  2) ABB PMS: se incorpora el informe de avance Nº 5 (E-2611027, 30-10-2026). Avance
     54% contra 54% previsto. El armado de tableros NO ARRANCÓ en octubre y pasa a
     noviembre en MEHCCO S.A.; el acopio quedó verificado al 100% el 22-10-26 con
     inspector de AESA y certificado. Nuevos PA-18, PA-20 y PA-21.
  3) ABB SALAS: se incorpora el informe de avance Nº 6 (R-2631036_A, 04-11-2026).
     Aprobada la ingeniería de detalle de TODOS los equipos para fabricación; la del
     shelter sigue con aprobación parcial. Acopio de auxiliares 100% y de UPS 98%.
     Fechas objetivo ratificadas: SE#3 23-04-27 y SE#4 25-03-27.
  4) Nueva previsión de FAT y nuevo PA-19 por la superposición del 30-11-2026.
Cambios respecto de Rev11:
  1) CAT MIRON — se incorpora el ACTA DE INSPECCIÓN IN-FC-CAT-19082-261021 del 21-10-2026
     (primera visita a planta, inspector Ing. Felice Corrao; por el proveedor Fernando
     Jelicich). Resultado: fase «proceso de fabricación», SIN NO CONFORMIDADES, provisión
     NO liberada y orden NO completa — como corresponde a una visita de seguimiento.
     Datos nuevos que aporta:
       · ORDEN DE COMPRA 4509019082 (en la Rev11 figuraba «no consignada»).
       · El alcance incluye además 3 LOTES DE REPUESTOS para 2 años de operación
         (ítems 10, 120 y 160) que no estaban en el cronograma ejecutivo.
       · Los 2 transformadores de 2.000 kVA 0,4/13,86 kV (OT 5212) son ELEVADORES EN
         BAÑO DE ACEITE; los otros 11 son secos cabinados. Cambia los requisitos de
         conservación durante el almacenamiento (PA-10).
       · La fabricación de NÚCLEOS está TERCERIZADA en una empresa de Villa Martelli,
         todavía no inspeccionada (nuevo PA-17).
       · Las bobinas de los transformadores de distribución (1,6 / 2 / 2,5 MVA) están en
         stock permanente, lo que agiliza las OT 5209, 5210 y 5211.
       · Se verificó el acopio de foils de aluminio y resina bicomponente con certificados
         de calidad de origen (FeSi Legnano Teknoelectric · resina Krempel DMR 9230 RF).
       · Se adjunta el PLAN DE INSPECCIÓN Y ENSAYOS 5155-00-0022-VE-CA-PE-001 Rev.A, con
         14 puntos de retención obligatorios y 10 días de aviso previo. Reformula PA-09.
     El cronograma de fabricación adjunto es el mismo Rev.A ya incorporado en la Rev6: se
     compararon las barras y los porcentajes y son idénticos. NO hay cambios de fecha.
  2) PUNTOS DE ATENCIÓN: se suma PA-17 (subproveedor de núcleos sin inspeccionar) y se
     reformulan PA-09 (ya hay plan de ensayos, falta la fecha) y PA-10 (dos unidades en
     baño de aceite).
Cambios de la Rev11 respecto de Rev10:
  1) INAUCO — se adopta el cronograma de ingeniería y suministro
     5155-00-0029-VE-CA-PG-001 REV 0 «PARA CONSTRUCCIÓN» del 23-10-2026 (reemplaza la
     Rev A «Para aprobación» del 24-08-2026). OC AESA 4509023773, obra Inauco 129-008.
     NO HAY CORRIMIENTOS: todas las fechas de hito y la entrega se sostienen, y la
     duración total baja de 340 a 334 días. Lo que cambió:
       · La PRUEBA DE INTEGRACIÓN ABB/PMS EN BUENOS AIRES se ADELANTA CINCO MESES, del
         28-04/04-05-27 al 01 y 02-12-2026 (2 días). Cierra PA-03.
       · Aparece la PRE-iFAT INTERNA (16-02 → 23-03-27), que antes no existía, y el hito
         de LIBERACIÓN Y ENTREGA DE TABLEROS el 04-05-27 queda incorporado al cronograma.
       · El REQUISITO 0 se desdobla en tres órdenes de compra: MO+MAT (100%), SEGUROS
         (0%) y MO DE CAMPO (0%); las dos últimas están VENCIDAS desde el 28-07-26.
       · El REQUISITO 1 pasa de un hito único a cinco entregas escalonadas de ingeniería
         de AESA, dos de ellas vencidas e incompletas (66% y 50%).
       · El rack CPF-COM-SE3 YA NO FIGURA en el alcance de construcción: PA-06 se cierra.
       · Avance global declarado: 6%. Configuración base 11% (PLC/comms 12%, SCADA 12%).
  2) NUEVA HOJA «Comparativo entre provisiones»: entregas, avance, informes de avance,
     insumos de AESA pendientes y estado de la documentación de las ocho provisiones.
  3) PUNTOS DE ATENCIÓN: se suma PA-16 (insumos de AESA vencidos a Inauco), se CIERRA
     PA-03 y se CIERRA PA-06.
Cambios de la Rev10 respecto de Rev9:
  1) HPH (HIMA, Alemania) — se incorpora el SUPPLIER PROGRESS REPORT Nº 01
     (AC-PR-2026-002 Rev.3, emitido el 15-10-2026 por Frédéric Plounévez), período
     01-08-26 a 30-10-26, PO/NP 4509010341. Lo esencial:
       · LA FABRICACIÓN DE LOS TRES TABLEROS DE MARSHALLING ESTÁ COMPLETADA
         (CPF-ESD-MC-07A, CPF-PSS-MC-07A y CPF-PSS-MC-07B) y el material de la orden
         principal fue entregado. El plan preveía fabricar hasta el 06-11-26: hay unas
         3 semanas de adelanto.
       · La ENTREGA SE CONFIRMA el 15-12-2026, DAP Ezeiza AIRPORT, sin desvío respecto
         de la fecha contractual en ninguno de los tres tableros.
       · Ingeniería al 80%: no hay documentos vencidos ni devueltos con comentarios;
         quedan 7 documentos a emitir entre el 21 y el 30-10-26.
       · INSPECCIÓN de los tres tableros prevista para la SEMANA 44 o 45 (26-10 al
         08-11-26), más un "remote HW check" con fecha a definir.
       · Único pendiente imputado a AESA: la LETTER OF CREDIT de la orden ampliada, con
         vencimiento el 30-10-2026.
     Se corrige además el nombre del proveedor: HIMA Paul HILDEBRANDT GmbH (veníamos
     escribiendo "Paul Hilpert").
  2) PUNTOS DE ATENCIÓN: se suman PA-14 (Letter of Credit de la orden ampliada) y PA-15
     (el informe no reporta liberación, despacho ni plan de certificación). PA-01 se
     matiza: la entrega es en EZEIZA AEROPUERTO, lo que apunta a carga aérea.
Cambios de la Rev9 respecto de Rev8:
  1) ABB SALAS ELÉCTRICAS — se adopta el cronograma 5155-00-0009-VE-CO-PG-001 REV 0,
     "PARA CONSTRUCCIÓN", del 14-10-2026 (reemplaza la Rev A "Para aprobación").
     IMPORTANTE: se compararon las seis páginas del Gantt contra la emisión del 02-10-26
     y son IDÉNTICAS — mismas fechas y mismos porcentajes de avance. Lo que cambió es el
     ESTADO del documento: el cronograma deja de ser un documento para aprobación y pasa
     a ser la LÍNEA BASE PARA CONSTRUCCIÓN. No hay, por lo tanto, actualización de avance
     de octubre: los porcentajes siguen siendo los del corte de septiembre.
  2) HARG (HIMA Argentina) — la orden de compra SIGUE SIN COLOCARSE al 21-10-2026. Se fija
     la colocación de la OC el 01-11-2026 y se confirma el inicio de los trabajos el
     02-11-2026, primer día hábil siguiente. Con esa premisa las fechas de HARG NO se
     mueven respecto de la Rev8 (el corrimiento de 49 días ya suponía ese arranque), pero
     el margen se agotó: quedan 11 días para colocar la OC sin mover el plan.
     Se incorpora la COLOCACIÓN DE LA OC como hito propio de la provisión y una TABLA DE
     SENSIBILIDAD que muestra, día por día, qué pasa si la OC se corre más.
  3) PA-04 se reescribe con la situación al 21-10 y la sensibilidad.
Cambios de la Rev8 respecto de Rev7:
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
REV="Rev13"; FECHA="2026-11-07"; HOY=date(2026,11,7)
OC_HARG=date(2026,11,16)  # Rev13: AESA estima colocar la OC la semana del 16-11-26 (+2 semanas sobre la Rev12)
SHIFT_HARG=63   # días corridos de corrimiento por OC no colocada (14-09-26 -> 16-11-26)
D=lambda y,m,d: date(y,m,d)
from datetime import timedelta
SH=lambda d: d+timedelta(SHIFT_HARG)   # corrimiento de las actividades propias de HARG

# (hito, inicio, fin, crítico, observación, es_acopio)
PROV=[
 ("ABB — Sala Eléctrica SE#3 (BT)","ABB","C62828",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"OC 45089944971 — pólizas aprobadas. Cronograma REV 0 «Para construcción» del 14-10-26",0),
  ("Kick-off Meeting (KOM) técnico",D(2026,5,15),D(2026,5,18),0,"",0),
  ("Confirmación de layout de sala y equipamiento",D(2026,5,18),D(2026,5,20),0,"Insumo AESA",0),
  ("Emisión de ingeniería básica",D(2026,5,20),D(2026,9,11),0,"100%",0),
  ("Aprobación de ingeniería básica",D(2026,9,11),D(2026,9,25),0,"APROBACIÓN PARCIAL al 04-11-26 (informe Nº6): nuevos comentarios en sucesivas revisiones impiden la aprobación final. Certificación del hito 2 corrida a noviembre",0),
  ("FREEZING POINT — ingeniería básica",D(2026,9,25),D(2026,9,25),1,"0% — VENCIDO HACE 6 SEMANAS. Sólo alcanza a la SALA: la ingeniería de TODOS los equipos quedó aprobada en octubre",0),
  ("Emisión de ingeniería de detalle",D(2026,5,22),D(2026,9,29),0,"95%",0),
  ("Aprobación de ingeniería de detalle — EQUIPOS",D(2026,9,1),D(2026,10,31),0,"APROBADA EN REVISIÓN NUMÉRICA PARA FABRICACIÓN (informe Nº6, hito del período): CCM, auxiliares, UPS/baterías y variadores BT",0),
  ("Aprobación de ingeniería de detalle — SALA",D(2026,9,29),D(2026,10,13),0,"APROBACIÓN PARCIAL — ABB declara en el informe Nº6 que los comentarios sucesivos impiden la aprobación final y la emisión para fabricación. VENCIDA hace 25 días",0),
  ("FREEZING POINT — ingeniería de detalle",D(2026,10,13),D(2026,10,13),1,"VENCIDO — ABB lo señala como factor potencial de retraso y pide no generar nuevos comentarios sobre documentos ya emitidos para fabricación",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,19),D(2026,5,19),0,"Compra y acopio sala BT",1),
  ("ACOPIO — Materiales de los CCM (críticos y secundarios)",D(2026,6,19),D(2026,10,20),0,"85% de recepción de materiales (informe Nº6, peso 53%); cierre al 100% previsto en noviembre",1),
  ("ACOPIO — Rectificadores / UPS / baterías",D(2026,6,9),D(2026,10,12),0,"98% — breakers, fusibles, ventiladores, TI, placas electrónicas y semiconductores al 100%; gabinetes y racks al 90% (informe Nº6)",1),
  ("ACOPIO — Tableros auxiliares",D(2026,7,22),D(2026,10,30),0,"100% — acopio de equipamiento mecánico y eléctrico COMPLETO (informe Nº6). Certificación del hito 4 prevista para noviembre",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,12,10),D(2026,12,10),1,"Cierre de acopio de la sala BT",1),
  ("Inicio de fabricación en fábrica de origen",D(2026,10,6),D(2026,10,6),0,"CUMPLIDO — CCM: fabricación de columnas y módulos extraíbles iniciada en octubre (informe Nº6)",0),
  ("Fabricación de tableros en fábrica de origen",D(2026,10,6),D(2026,12,11),1,"CCM 06/10-19/11 (Sorocaba) · rectif./UPS 14-23/10 (Santa Fe) · auxiliares 30/10-11/12 (AMBA)",0),
  ("Ensayos internos en fábrica de origen",D(2026,10,26),D(2026,12,3),0,"Previos a la FAT",0),
  ("FAT EN FÁBRICA DE ORIGEN — equipos",D(2026,11,18),D(2026,12,16),1,"Witness AESA — INVITACIONES ENVIADAS · UPS/baterías 18-20/11 Santa Fe · CCM 30/11-12/12 Sorocaba · auxiliares 14-16/12 AMBA. Se superponen con la maqueta CPF2 — PA-19",0),
  ("Adecuación y despacho de fábrica de origen",D(2026,11,27),D(2026,12,17),1,"Condición EXW en origen",0),
  ("Entrega en Shelterista (Mendoza)",D(2026,11,26),D(2027,1,10),1,"Arribo escalonado: rectif./UPS 26-30/11 · aux. 21-28/12 · drives 23-29/12 · CCM 23/12-10/01",0),
  ("Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)",D(2026,11,23),D(2027,3,5),1,"Shelterista Bottino — inspección a proveedor prevista en octubre",0),
  ("Montaje e integración de equipos en sala",D(2027,2,26),D(2027,3,17),0,"Shelterista",0),
  ("Interconexión de equipos",D(2027,3,9),D(2027,3,23),0,"",0),
  ("FAT SALA ELÉCTRICA (shelter completo)",D(2027,3,29),D(2027,4,2),1,"Mendoza · ratificada en el informe Nº6 (el cronograma del 02-10 la ubica 23-30/03)",0),
  ("Ensayos PMS en sala (INTERFASE — contrato PMS)",D(2027,3,30),D(2027,4,6),0,"Alcance del contrato PMS",0),
  ("Desconexión, embalaje y despacho",D(2027,4,7),D(2027,4,23),1,"",0),
  ("DESPACHO / ENTREGA EXW",D(2027,4,23),D(2027,4,23),1,"HITO FINAL — fecha objetivo RATIFICADA por ABB en el informe Nº6 (23-04-27). Fecha CONTRACTUAL 14-05-27: 21 días de mejora. Rev12 la ubicaba el 20-04-27",0),
 ]),
 ("ABB — Sala Eléctrica SE#4 (MT)","ABB","AD1457",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"OC 45089944973 — pólizas aprobadas. Cronograma REV 0 «Para construcción» del 14-10-26",0),
  ("Kick-off Meeting (KOM) técnico",D(2026,5,15),D(2026,5,18),0,"",0),
  ("Confirmación de layout de sala y equipamiento",D(2026,5,18),D(2026,5,20),0,"Insumo AESA",0),
  ("Emisión de ingeniería básica",D(2026,5,20),D(2026,9,11),0,"100%",0),
  ("Aprobación de ingeniería básica",D(2026,9,11),D(2026,9,25),0,"APROBACIÓN PARCIAL al 04-11-26 (informe Nº6): nuevos comentarios en sucesivas revisiones impiden la aprobación final. Certificación del hito 2 corrida a noviembre",0),
  ("FREEZING POINT — ingeniería básica",D(2026,9,25),D(2026,9,25),1,"0% — VENCIDO HACE 6 SEMANAS. Sólo alcanza a la SALA: la ingeniería de TODOS los equipos quedó aprobada en octubre",0),
  ("Emisión de ingeniería de detalle",D(2026,5,22),D(2026,9,29),0,"95%",0),
  ("Aprobación de ingeniería de detalle — EQUIPOS",D(2026,9,1),D(2026,10,31),0,"APROBADA EN REVISIÓN NUMÉRICA PARA FABRICACIÓN (informe Nº6, hito del período): tableros MT, auxiliares y variadores MT",0),
  ("Aprobación de ingeniería de detalle — SALA",D(2026,9,29),D(2026,10,13),0,"APROBACIÓN PARCIAL — ABB declara en el informe Nº6 que los comentarios sucesivos impiden la aprobación final y la emisión para fabricación. VENCIDA hace 25 días",0),
  ("FREEZING POINT — ingeniería de detalle",D(2026,10,13),D(2026,10,13),1,"VENCIDO — ABB lo señala como factor potencial de retraso y pide no generar nuevos comentarios sobre documentos ya emitidos para fabricación",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,19),D(2026,5,19),0,"Compra y acopio sala MT",1),
  ("ACOPIO — Celdas MT (materiales principales y secundarios)",D(2026,6,22),D(2026,10,30),1,"Recepción de materiales 40% real a sep-26 y 80% previsto a oct-26 (informe Nº6); cierre al 100% previsto en noviembre",1),
  ("ACOPIO — Variadores MT (China) y BT (Finlandia)",D(2026,8,18),D(2026,9,30),0,"Fabricación y ensayos FINALIZADOS (confirmado en el informe Nº6). Entrega en origen comprometida el 06-11-26 en ambos; transporte marítimo al 10%",1),
  ("ACOPIO — Tableros auxiliares",D(2026,7,22),D(2026,10,30),0,"100% — acopio de equipamiento mecánico y eléctrico COMPLETO (informe Nº6). Certificación del hito 4 prevista para noviembre",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,12,10),D(2026,12,10),1,"Cierre de acopio de la sala MT",1),
  ("Inicio de fabricación en fábrica de origen",D(2026,10,8),D(2026,10,8),0,"CUMPLIDO — celdas MT: fabricación de gabinetes de comando iniciada en octubre (informe Nº6)",0),
  ("Fabricación de tableros y celdas en fábrica de origen",D(2026,10,8),D(2026,12,11),1,"Celdas MT 08/10-16/11 (Sorocaba) · auxiliares 30/10-11/12 (AMBA)",0),
  ("Ensayos internos en fábrica de origen",D(2026,11,17),D(2026,11,26),0,"Previos a la FAT",0),
  ("FAT EN FÁBRICA DE ORIGEN — equipos",D(2026,11,30),D(2026,12,16),1,"Witness AESA — INVITACIONES ENVIADAS · TGMT 30/11-04/12 Sorocaba · auxiliares 14-16/12 AMBA. Arrancan el mismo día que la maqueta CPF2 en Bs.As. — PA-19",0),
  ("Despacho de fábrica de origen (celdas MT y tableros)",D(2026,12,4),D(2026,12,21),1,"Acond./embalaje 04-10/12 · FCA 10-14/12",0),
  ("Entrega en Shelterista (Mendoza)",D(2026,12,28),D(2027,1,6),1,"Arribo escalonado: aux. 21-28/12 · drives 23-30/12 · celdas 04-06/01",0),
  ("Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)",D(2026,10,6),D(2027,1,29),1,"Shelterista Bottino — inicio de fabricación previsto en octubre",0),
  ("Montaje e integración de equipos en sala",D(2027,1,29),D(2027,2,19),0,"Shelterista",0),
  ("Interconexión de equipos",D(2027,2,17),D(2027,2,23),0,"",0),
  ("FAT SALA ELÉCTRICA (shelter completo)",D(2027,2,22),D(2027,2,26),1,"Mendoza · ratificada en el informe Nº6 (el cronograma del 02-10 la ubica 25/02-01/03)",0),
  ("Ensayos PMS en sala (INTERFASE — contrato PMS)",D(2027,3,3),D(2027,3,9),0,"Alcance del contrato PMS",0),
  ("Desconexión, embalaje y despacho",D(2027,3,10),D(2027,3,25),1,"",0),
  ("DESPACHO / ENTREGA EXW",D(2027,3,25),D(2027,3,25),1,"HITO FINAL — objetivo ABB. Fecha CONTRACTUAL 08-04-27: 14 días de margen",0),
 ]),
 ("ABB — Ductos de barras","ABB","795548",[
  ("Inicio de contrato — Recepción de OC",D(2026,5,8),D(2026,5,8),0,"Mismo contrato ABB (alcance separado)",0),
  ("Emisión layout instalación + detalle conexionado trafo",D(2026,7,20),D(2026,9,18),1,"INSUMO AESA — se corrió del 29-08 al 18-09-26 y arrastró todo el bloque",0),
  ("FREEZING POINT — inicio de diseño",D(2026,9,18),D(2026,9,18),1,"CORRIDO 20 días (antes 29-08-26)",0),
  ("Diseño de ductos de barras",D(2026,9,18),D(2026,12,14),1,"15% real a sep-26 · 50% previsto a oct-26 y 100% a nov-26 (informe Nº6). El diseño arrancó el 21-09-26 al recibir la información de gabinetes de trafos MT",0),
  ("Aprobación de diseño",D(2026,12,14),D(2026,12,30),0,"Aprobación AESA",0),
  ("FREEZING POINT — cierre de diseño",D(2026,12,30),D(2026,12,30),1,"CORRIDO 26 días (antes 04-12-26) — habilita fabricación",0),
  ("ACOPIO / fabricación de ductos de barras",D(2026,12,30),D(2027,4,9),1,"100 días corridos (acopio integrado a fabricación)",1),
  ("FCA fábrica de origen",D(2027,4,9),D(2027,4,16),1,"Puesta a disposición en origen",0),
  ("Traslado internacional + aduana",D(2027,4,16),D(2027,7,5),1,"80 días corridos — tramo más largo",0),
  ("ENTREGA DE EQUIPOS EN SITIO",D(2027,7,5),D(2027,7,8),1,"HITO FINAL — la entrega más tardía del paquete. PLAZO SIN CONFIRMAR: ABB reitera en el informe Nº6 el pedido de adecuar la fecha contractual de las posiciones 120/130/140 de la OC 45089944971, sin respuesta de AESA desde el 05-10-26, y sigue evaluando el impacto económico",0),
 ]),
 ("CAT Miron — Transformadores (13 unidades)","CAT Miron","33691E",[
  ("Recepción de la orden de compra 4509019082",D(2026,7,13),D(2026,7,13),0,"Obra 5155 · la FECHA de la OC no consta; se toma el inicio de ingeniería",0),
  ("Ingeniería — cálculo y diseño (5 órdenes de trabajo)",D(2026,7,13),D(2026,8,21),0,"15 días hábiles por OT — 100% en las 5 OT",0),
  ("Ingeniería — proyecto de ingeniería en detalle",D(2026,7,27),D(2026,9,4),1,"15 días hábiles por OT — 100% salvo OT 5211 (75%); insumo del detalle de conexionado a trafo",0),
  ("ACOPIO — Conductores BT/MT (foils de aluminio)",D(2026,7,27),D(2026,9,25),1,"30 días hábiles por OT · acopio VERIFICADO en inspección con certificados de origen",1),
  ("Emisión del Plan de Inspección y Ensayos (PIE) Rev.A",D(2026,9,1),D(2026,9,1),1,"5155-00-0022-VE-CA-PE-001 · 14 puntos de retención con 10 días de aviso",0),
  ("ACOPIO — Materiales gruesos (gabinetes y resina bicomponente)",D(2026,8,3),D(2026,10,16),1,"45 días hábiles por OT · resina Krempel DMR 9230 RF lote 019/26 aprobada",1),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,10,16),D(2026,10,16),1,"Cierre del acopio de la última OT (5211)",1),
  ("Fabricación de núcleos — TERCERIZADA (Villa Martelli)",D(2026,9,7),D(2026,11,6),1,"Corte de FeSi y conformado de núcleos ventana fuera de la planta de CAT — subproveedor AÚN NO INSPECCIONADO",0),
  ("Fabricación — bobinado BT/MT",D(2026,9,7),D(2026,11,27),1,"40 días hábiles por OT · bobinas de 1,6/2/2,5 MVA en STOCK PERMANENTE: agiliza OT 5209, 5210 y 5211",0),
  ("INSPECCIÓN Nº 1 EN PLANTA CAT",D(2026,10,21),D(2026,10,21),1,"Acta IN-FC-CAT-19082-261021 · SIN no conformidades · provisión NO liberada (visita de seguimiento)",0),
  ("Fabricación — montaje final y montaje de gabinete",D(2026,11,2),D(2026,12,4),1,"3 + 5 días hábiles por OT · 11 unidades secas cabinadas",0),
  ("Ensayos internos de laboratorio",D(2026,11,2),D(2026,12,4),1,"2 días hábiles por OT — NO incluyen las pruebas con testigo del cliente",0),
  ("ENSAYOS CON PUNTO DE RETENCIÓN — A COORDINAR",D(2026,11,2),D(2026,12,11),1,"PIE Rev.A: 9 ensayos eléctricos + control dimensional + despacho. SIN FECHA en el cronograma",0),
  ("Acondicionamiento y carga — OT 5209 (4×2000 kVA) y OT 5210 (2×1600 kVA)",D(2026,11,9),D(2026,11,13),1,"Máquinas 3489-3492 y 3493-3494",0),
  ("ENTREGA lote 1 — OT 5209 (4 un.) + OT 5210 (2 un.)",D(2026,11,13),D(2026,11,13),1,"6 transformadores secos cabinados 13,2/0,4 kV",0),
  ("Acondicionamiento y carga — OT 5212 (2×2000 kVA elevadores)",D(2026,11,23),D(2026,11,27),1,"Máquinas 3497-3498 · EN BAÑO DE ACEITE",0),
  ("ENTREGA lote 2 — OT 5212 (2 un.) EN BAÑO DE ACEITE",D(2026,11,27),D(2026,11,27),1,"2 elevadores 0,4/13,86 kV YND11 — conservación distinta a los secos",0),
  ("Acondicionamiento y carga — OT 5213 (3×500 kVA)",D(2026,11,30),D(2026,12,4),1,"Máquinas A7546-A7548",0),
  ("ENTREGA lote 3 — OT 5213 (3 un.)",D(2026,12,4),D(2026,12,4),1,"3 transformadores secos cabinados 0,38/0,48 kV",0),
  ("Acondicionamiento y carga — OT 5211 (2×2500 kVA)",D(2026,12,7),D(2026,12,11),1,"Máquinas 3495-3496",0),
  ("ENTREGA FINAL — OT 5211 (2 un.) · cierre de la provisión",D(2026,12,11),D(2026,12,11),1,"HITO FINAL — 2 secos cabinados 13,2/6,9 kV. Completa 13 transformadores + 3 lotes de repuestos para 2 años",0),
 ]),
 ("ABB — PMS (Power Management System)","ABB","EF6C00",[
  ("Inicio de contrato — Recepción / aceptación OC",D(2026,5,8),D(2026,5,11),0,"OC 4508945953 — 100%",0),
  ("Kick-off Meeting (KOM) y planificación detallada",D(2026,5,22),D(2026,6,5),0,"100%",0),
  ("Emisión de ingeniería básica (documentación para aprobación)",D(2026,6,8),D(2026,7,31),0,"100% · Hito cert. 10% el 31-07-26",0),
  ("Aprobación de ingeniería básica (AESA)",D(2026,7,6),D(2026,8,14),0,"100% · Hito cert. 10% el 14-08-26",0),
  ("Emisión de ingeniería Rev.0",D(2026,7,20),D(2026,8,28),0,"100%",0),
  ("Aprobación ingeniería Rev.0 — FREEZING POINT",D(2026,8,3),D(2026,9,11),1,"100% — cierra ingeniería",0),
  ("ACOPIO — Inicio de compra de materiales",D(2026,5,22),D(2026,5,22),0,"Switches de red, S800, tableros",1),
  ("ACOPIO — Material S800 y tableros",D(2026,5,22),D(2026,10,8),0,"100% — verificado en stock ABB/DHL el 22-10-26",1),
  ("ACOPIO — Switches de red",D(2026,5,22),D(2026,10,22),1,"100% — verificado en stock ABB/DHL el 22-10-26 (era el ítem de mayor plazo)",1),
  ("Configuración / programación del sistema",D(2026,8,17),D(2026,12,11),1,"0% de la especialidad «Configuración» en el informe Nº5 (peso 10%): ABB la pronostica recién al 25% en nov-26. Requiere los mapas pendientes de AESA — PA-11",0),
  ("INSPECCIÓN DE ACOPIO EN STOCK ABB/DHL (witness AESA)",D(2026,10,22),D(2026,10,22),1,"Acta IN-SAV-ABB-45953-261022 — se corroboró el 100% de los materiales necesarios para el proyecto",1),
  ("INICIO DE ARMADO DE TABLEROS",D(2026,11,2),D(2026,11,2),1,"NO ARRANCÓ EN OCTUBRE: el informe Nº5 declara «sin actividades» en Fabricación y lleva el inicio a noviembre, en MEHCCO S.A. (San Martín, Bs.As.). Corrimiento acumulado de 7,5 semanas sobre la Rev6 (11-09-26) — PA-18",0),
  ("Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF1",D(2026,10,19),D(2026,10,23),0,"El informe Nº5 NO registra su ejecución en octubre: confirmar con ABB si se realizó",0),
  ("Emisión de procedimientos de FAT y SAT (Rev.A)",D(2026,11,2),D(2026,11,30),0,"CORRIDA A NOVIEMBRE — el informe Nº5 ubica el desarrollo de los protocolos y registros de FAT y SAT en los «eventos del próximo mes». La Rev12 la preveía 19-30/10-26 — PA-21",0),
  ("Aprobación de procedimientos FAT / SAT Rev.A (AESA)",D(2026,12,1),D(2026,12,11),0,"Corrida 4 semanas por el arrastre de la emisión",0),
  ("Configuración de variables de comunicación y lógicas del sistema",D(2026,11,2),D(2026,12,11),1,"IEC61850 / PROFINET / MODBUS — requiere los mapas pendientes de AESA",0),
  ("ACOPIO DE MATERIALES COMPLETO",D(2026,10,22),D(2026,10,22),1,"CUMPLIDO Y VERIFICADO el 22-10-26 con inspector de AESA. Hito de certificación C-2 (25%, USD 220.055,72) realizado en oct-26",1),
  ("Emisión de procedimientos de FAT y SAT Rev.0",D(2026,12,14),D(2026,12,23),0,"Cae sobre el receso de fin de año",0),
  ("Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF2 (integración con PCS Inauco)",D(2026,11,30),D(2026,12,1),1,"CONFIRMADA en el informe Nº5: semana del 30-11-26, 2 días. Maqueta de CPF1 + cubicles de UMC propios de CPF2 · ABB/AESA/PPSA + INAUCO. Choca con las FAT de TGMT y CCM en Sorocaba — PA-19",0),
  ("Aprobación de procedimientos FAT / SAT Rev.0 (AESA)",D(2027,1,4),D(2027,1,15),0,"Quedaría aprobada SÓLO 6 días hábiles antes de la FAT del 25-01-27 — PA-21",0),
  ("FIN DE ARMADO DE TABLEROS",D(2026,12,17),D(2026,12,17),1,"ABB sostiene la fecha, pero 50 días hábiles desde el 02-11-26 terminan el 12-01-27: EN RIESGO de 4 semanas — PA-18",0),
  ("Pruebas internas (pre-FAT) de tableros y sistema",D(2026,12,18),D(2027,1,7),0,"15 días hábiles",0),
  ("FAT Tableros + Sistema",D(2027,1,25),D(2027,1,29),1,"Hito cert. 20% · witness AESA. El plan de certificación acordado ubica el hito «contra finalización de la FAT» en feb-27 — PA-18",0),
  ("Ajustes post-FAT / liberación",D(2027,2,1),D(2027,2,3),0,"",0),
  ("Despachos",D(2027,2,4),D(2027,2,17),1,"",0),
  ("ENTREGA DE EQUIPAMIENTO EN SHELTER",D(2027,2,17),D(2027,2,17),1,"HITO FINAL — cert. 30%. El plan de certificación acordado (abr-26) ubica «contra entrega de equipos» en mar-27. Fuera del corte: FAT INTEGRAL del shelter en planta shelterista (5 días, acordada con AESA) y DataBook",0),
 ]),
 ("Inauco — PCS / SCADA / Comunicaciones","Inauco","1565C0",[
  ("Inicio del proyecto — Recepción de OC de MO + MATERIALES (AESA)",D(2026,7,28),D(2026,7,28),0,"OC AESA 4509023773 · obra Inauco 129-008 — REQ 0A, 100%",0),
  ("Recepción de OC de SEGUROS (AESA)",D(2026,7,28),D(2026,7,28),1,"REQ 0B — 0% al 23-10-26: VENCIDO hace 3 meses",0),
  ("Recepción de OC de MANO DE OBRA DE CAMPO (AESA)",D(2026,7,28),D(2026,7,28),1,"REQ 0C — 0% al 23-10-26: VENCIDO hace 3 meses",0),
  ("REQ 1 — Recepción de ingeniería de AESA (5 entregas escalonadas)",D(2026,8,14),D(2026,12,11),1,"14/08 100% · 21/08 66% · 09/10 50% · 30/10 0% (incluye doc. HIMA) · 11/12 0%",0),
  ("Emisión de ingeniería Rev.A — documentos de tableros",D(2026,8,24),D(2026,10,16),0,"5 de 6 entregas al 100%; la del 16/10 (13 doc. CA-PR) al 77%",0),
  ("Configuración base — PLC y comunicaciones",D(2026,8,26),D(2027,3,23),1,"12% — antes dividido en Etapa 1 y Etapa 2",0),
  ("Configuración base — SCADA",D(2026,8,26),D(2027,3,23),1,"12%",0),
  ("Emisión de ingeniería Rev.0 — documentos de tableros",D(2026,9,21),D(2026,11,13),0,"21/09 67% · 02/10 100% · 09/10 100% · 16/10 67% · 23/10 0% · 30/10 y 13/11 pendientes",0),
  ("REQ 3 — FREEZING POINT: aprobación de ingeniería para construcción",D(2026,11,19),D(2026,11,19),1,"Sin cambio respecto de la Rev.A",0),
  ("ACOPIO — REQ 6: Recepción de HW (materiales críticos)",D(2026,11,26),D(2026,11,26),1,"10/12 semanas desde la OC — sin cambio",1),
  ("INICIO DE CONSTRUCCIÓN DE TABLEROS",D(2026,11,27),D(2026,11,27),1,"Según cronograma de fabricación PG-002 Rev.1",0),
  ("Pruebas de integración con ABB (PMS) en Buenos Aires",D(2026,12,1),D(2026,12,2),1,"ADELANTADA 5 MESES (Rev.A: 28/04-04/05/27). 2 días · con ABB, AESA, PPSA",0),
  ("Fabricación tablero CPF2-PCS-SC-07b (PCS + marshalling · 2 gabinetes)",D(2026,12,1),D(2027,1,5),1,"20 días",0),
  ("Fabricación tablero CPF2-PCS-SC-07a (PCS + marshalling · 3 gabinetes)",D(2026,12,1),D(2027,1,19),1,"30 días",0),
  ("PreFAT interno y liberación de pendientes — tablero 07b",D(2027,1,6),D(2027,1,15),0,"3 días de PreFAT + 5 de pendientes",0),
  ("PreFAT interno y liberación de pendientes — tablero 07a",D(2027,1,20),D(2027,1,29),0,"3 días de PreFAT + 5 de pendientes",0),
  ("Fabricación rack de comunicaciones CPF-COM-07",D(2027,1,22),D(2027,2,3),0,"El rack CPF-COM-SE3 YA NO FIGURA en el alcance de construcción",0),
  ("REQ 2 — Recepción tableros HIMA en base Inauco (insumo HPH)",D(2027,1,25),D(2027,1,25),1,"Sin cambio — interfase HPH/HARG",0),
  ("CIERRE DE CONSTRUCCIÓN DE TABLEROS",D(2027,2,16),D(2027,2,16),1,"Sin cambio",0),
  ("PRE-iFAT INTERNA (pruebas integradas previas a la FAT con el cliente)",D(2027,2,16),D(2027,3,23),1,"HITO NUEVO — no existía en la Rev.A",0),
  ("Emisión de documentos de configuración (Rev.A y Rev.0)",D(2027,2,19),D(2027,4,9),0,"12 documentos de configuración",0),
  ("REQ 4 — Aprobación de procedimientos FAT Rev.0",D(2027,2,24),D(2027,2,24),0,"Sin cambio",0),
  ("FAT PCS + COMUNICACIONES (Inauco Neuquén)",D(2027,3,29),D(2027,4,22),1,"18 días · witness AESA/PPSA · LUGAR: Neuquén",0),
  ("FAT SIS — integración con HIMA (Inauco Neuquén)",D(2027,3,29),D(2027,4,26),1,"20 días · requiere presencia de HIMA",0),
  ("FIN DE FAT INTEGRAL",D(2027,4,26),D(2027,4,26),1,"HITO NUEVO — cierra con la FAT más tardía",0),
  ("Configuraciones post-FAT",D(2027,4,27),D(2027,5,17),0,"Se extiende 13 días más allá de la entrega de tableros",0),
  ("LIBERACIÓN Y ENTREGA DE TABLEROS (camión de AESA, Inauco Neuquén)",D(2027,5,4),D(2027,5,4),1,"HITO FINAL — sin cambio. Fuera del corte: SAT y comisionado en campo hasta el 26-10-27 y cierre de documentación el 07-12-27",0),
 ]),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","HPH","6A1B9A",[
  ("Inicio de contrato / proyecto",D(2026,7,10),D(2026,7,10),0,"PO/NP 4509010341 · CPF-ESD-MC-07A, CPF-PSS-MC-07A y CPF-PSS-MC-07B",0),
  ("Inicio de ingeniería",D(2026,8,14),D(2026,8,14),0,"",0),
  ("Ingeniería de detalle / especificación",D(2026,8,17),D(2026,9,11),0,"Estado general de ingeniería: 80% al 15-10-26",0),
  ("ACOPIO / disposición de materiales",D(2026,9,14),D(2026,10,16),1,"MATERIAL DE LA ORDEN PRINCIPAL ENTREGADO. Orden ampliada: material reservado en ERP",1),
  ("Fin de ingeniería / CAD — FREEZING POINT",D(2026,9,18),D(2026,9,18),1,"Cierra ingeniería",0),
  ("FABRICACIÓN DE LOS TABLEROS DE MARSHALLING",D(2026,9,21),D(2026,10,15),1,"COMPLETADA para los 3 tableros — plan 06-11-26, unas 3 semanas de adelanto",0),
  ("Emisión de documentación de ingeniería (diagramas de conexionado)",D(2026,10,21),D(2026,10,21),0,"3 documentos DC-001/002/003 Rev.0, uno por tablero",0),
  ("LETTER OF CREDIT de la orden ampliada (AESA)",D(2026,10,30),D(2026,10,30),1,"INSUMO DE AESA — único pendiente que HPH imputa al EPC",0),
  ("Emisión de documentación de ingeniería (cálculos térmicos y procedimientos FAT/SAT)",D(2026,10,30),D(2026,10,30),0,"Cronograma de ing. y suministro, cálculos de consumo y disipación, proc. FAT y SAT",0),
  ("Inspección / HW check de los 3 tableros",D(2026,10,26),D(2026,11,8),1,"Semana 44 o 45 según el informe · más un «remote HW check» con fecha A DEFINIR",0),
  ("Preparación de despacho (packing)",D(2026,11,9),D(2026,11,20),0,"",0),
  ("Listo para despacho",D(2026,11,20),D(2026,11,20),1,"",0),
  ("Despacho / shipping",D(2026,11,23),D(2026,12,14),1,"Con la fabricación terminada, evaluar adelantar el despacho",0),
  ("Documentación de entrega enviada",D(2026,12,4),D(2026,12,4),0,"",0),
  ("Entrega DAP Ezeiza AEROPUERTO (Argentina)",D(2026,12,15),D(2026,12,15),1,"CONFIRMADA sin desvío: fecha contractual = fecha objetivo en los 3 tableros",0),
  ("Nacionalización / despacho de aduana + transporte a base Inauco",D(2026,12,15),D(2027,1,25),1,"Entrega en aeropuerto: a confirmar modo de transporte y plazo real de aduana",0),
  ("RECEPCIÓN DE TABLEROS EN BASE INAUCO",D(2027,1,25),D(2027,1,25),1,"HITO FINAL — habilita la adecuación de HARG",0),
 ]),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","HARG","00838F",[
  # H = actividad propia de HARG -> corrida SHIFT_HARG días por la OC no colocada
  # I = insumo de HPH (ya contratada) -> NO se mueve
  ("COLOCACIÓN DE LA OC POR AESA",OC_HARG,OC_HARG,1,"PENDIENTE al 07-11-26 — AESA informa que la OC NO sale esta semana sino la siguiente: se asume el 16-11-26. DOS SEMANAS MÁS de postergación sobre la Rev12 (01-11-26). Gobierna todas las fechas de esta provisión",0),
  ("Documentación HARG — Doc. 1 (HW y otros)",SH(D(2026,9,14)),SH(D(2026,9,25)),0,"2 semanas · corrido +63 d desde el plan original (+14 d sobre la Rev12)",0),
  ("Programación HARG — Recurso 1",SH(D(2026,9,21)),SH(D(2027,1,29)),1,"11 semanas · corrido +63 d desde el plan original (+14 d sobre la Rev12)",0),
  ("Documentación HPH — Doc. 1 (INTERFASE)",D(2026,9,28),D(2026,10,16),0,"Insumo de HPH (3,5 semanas) — NO se corre",0),
  ("Ingeniería de Aplicación SIS y documentación (INTERFASE HPH)",D(2026,11,9),D(2026,12,18),1,"6 semanas — a definir por HPH; NO se corre",0),
  ("Documentación HARG — Doc. 2 (Modbus)",SH(D(2026,11,9)),SH(D(2026,11,13)),0,"1 semana · cae en el receso de fin de año — reprogramar",0),
  ("Fabricación HPH — entrega DIC-2026 (INTERFASE)",D(2026,12,21),D(2026,12,21),1,"Insumo de HPH — NO se corre",0),
  ("Adecuación de tableros marshalling (INTERFASE Inauco)",SH(D(2026,12,21)),SH(D(2027,1,29)),1,"6 semanas en base Inauco · corrido +63 d (OC). Arranca el 22-02-27: ya NO hay holgura contra la recepción de los tableros de HPH (25-01-27), pero se come la ventana de la Pre iFAT",0),
  ("Programación HARG — Recurso 2",SH(D(2027,1,4)),SH(D(2027,1,29)),0,"5 semanas · corrido +63 d (OC)",0),
  ("Documentación HARG — Doc. 3 (Proc. FAT e IFAT)",SH(D(2027,1,25)),SH(D(2027,2,12)),0,"3 semanas · corrido +63 d (OC)",0),
  ("FAT",SH(D(2027,2,24)),SH(D(2027,2,26)),1,"0,5 semana · 28-30/04/27: cae DESPUÉS del cierre de la FAT SIS conjunta de Inauco (26-04-27)",0),
  ("IFAT (Recursos 1 y 2)",SH(D(2027,3,1)),SH(D(2027,3,12)),1,"2 semanas · 03-14/05/27: termina 18 días DESPUÉS del cierre de la FAT SIS y 10 días después de la liberación de tableros de Inauco",0),
  ("Documentación HARG — Doc. 5 (Proy. Doc.) y HPH Doc. 2",SH(D(2027,3,18)),SH(D(2027,3,22)),0,"Documentación de cierre · corrido +63 d (OC)",0),
  ("ENTREGA DEL ALCANCE HARG (IFAT + documentación)",SH(D(2027,3,12)),SH(D(2027,3,22)),1,"HITO FINAL — 24-05-2027. Cierra 20 días DESPUÉS de la liberación de tableros de Inauco (04-05-27) y pasa a ser la fecha más tardía del camino crítico",0),
 ]),
]

# ---------------- Puntos de atención (Rev3) ----------------
# clave: (provisión, hito) -> código de PA
PA_MARK={
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Entrega DAP Ezeiza AEROPUERTO (Argentina)"):"PA-01",
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","LETTER OF CREDIT de la orden ampliada (AESA)"):"PA-14",
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Listo para despacho"):"PA-15",
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Despacho / shipping"):"PA-15",
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Nacionalización / despacho de aduana + transporte a base Inauco"):"PA-01",
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","RECEPCIÓN DE TABLEROS EN BASE INAUCO"):"PA-02",
 ("Inauco — PCS / SCADA / Comunicaciones","REQ 2 — Recepción tableros HIMA en base Inauco (insumo HPH)"):"PA-02",
 ("Inauco — PCS / SCADA / Comunicaciones","Recepción de OC de SEGUROS (AESA)"):"PA-16",
 ("Inauco — PCS / SCADA / Comunicaciones","Recepción de OC de MANO DE OBRA DE CAMPO (AESA)"):"PA-16",
 ("Inauco — PCS / SCADA / Comunicaciones","REQ 1 — Recepción de ingeniería de AESA (5 entregas escalonadas)"):"PA-16",
 ("Inauco — PCS / SCADA / Comunicaciones","Pruebas de integración con ABB (PMS) en Buenos Aires"):"PA-03",
 ("ABB — PMS (Power Management System)","Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF2 (integración con PCS Inauco)"):"PA-03",
 ("ABB — PMS (Power Management System)","Configuración de variables de comunicación y lógicas del sistema"):"PA-11",
 ("ABB — PMS (Power Management System)","Configuración / programación del sistema"):"PA-11",
 ("ABB — PMS (Power Management System)","INICIO DE ARMADO DE TABLEROS"):"PA-18",
 ("ABB — PMS (Power Management System)","FIN DE ARMADO DE TABLEROS"):"PA-18",
 ("ABB — PMS (Power Management System)","FAT Tableros + Sistema"):"PA-18",
 ("ABB — PMS (Power Management System)","Emisión de procedimientos de FAT y SAT (Rev.A)"):"PA-21",
 ("ABB — PMS (Power Management System)","Aprobación de procedimientos FAT / SAT Rev.0 (AESA)"):"PA-21",
 ("ABB — PMS (Power Management System)","Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF1"):"PA-20",
 ("ABB — Sala Eléctrica SE#3 (BT)","FAT EN FÁBRICA DE ORIGEN — equipos"):"PA-19",
 ("ABB — Sala Eléctrica SE#4 (MT)","FAT EN FÁBRICA DE ORIGEN — equipos"):"PA-19",
 ("ABB — PMS (Power Management System)","ENTREGA DE EQUIPAMIENTO EN SHELTER"):"PA-12",
 ("ABB — Sala Eléctrica SE#3 (BT)","Ensayos PMS en sala (INTERFASE — contrato PMS)"):"PA-12",
 ("ABB — Sala Eléctrica SE#4 (MT)","Ensayos PMS en sala (INTERFASE — contrato PMS)"):"PA-12",
 ("ABB — Ductos de barras","Emisión layout instalación + detalle conexionado trafo"):"PA-07",
 ("CAT Miron — Transformadores (13 unidades)","ENSAYOS CON PUNTO DE RETENCIÓN — A COORDINAR"):"PA-09",
 ("CAT Miron — Transformadores (13 unidades)","Ensayos internos de laboratorio"):"PA-09",
 ("CAT Miron — Transformadores (13 unidades)","Fabricación de núcleos — TERCERIZADA (Villa Martelli)"):"PA-17",
 ("CAT Miron — Transformadores (13 unidades)","ENTREGA lote 1 — OT 5209 (4 un.) + OT 5210 (2 un.)"):"PA-10",
 ("CAT Miron — Transformadores (13 unidades)","ENTREGA lote 2 — OT 5212 (2 un.) EN BAÑO DE ACEITE"):"PA-10",
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
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","COLOCACIÓN DE LA OC POR AESA"):"PA-04",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Documentación HARG — Doc. 1 (HW y otros)"):"PA-04",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Programación HARG — Recurso 1"):"PA-04",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Documentación HARG — Doc. 2 (Modbus)"):"PA-04",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Adecuación de tableros marshalling (INTERFASE Inauco)"):"PA-05",
 ("Inauco — PCS / SCADA / Comunicaciones","PRE-iFAT INTERNA (pruebas integradas previas a la FAT con el cliente)"):"PA-05",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","IFAT (Recursos 1 y 2)"):"PA-05",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","ENTREGA DEL ALCANCE HARG (IFAT + documentación)"):"PA-05",
 ("Inauco — PCS / SCADA / Comunicaciones","LIBERACIÓN Y ENTREGA DE TABLEROS (camión de AESA, Inauco Neuquén)"):"PA-05",
}
PUNTOS=[
 dict(cod="PA-18", tit="El armado de los tableros del PMS no arrancó en octubre y ABB sigue sosteniendo la fecha de fin",
      inter="ABB PMS → MEHCCO S.A. (taller) → FAT → entrega en shelter",
      plan="La Rev12 tomaba el 09-10-2026 como inicio del armado de tableros, con 50 días hábiles de duración y fin el 17-12-2026. El informe de avance Nº5 (30-10-2026) declara «sin actividades en el periodo» en el área de Fabricación y ubica el inicio de fabricación —en las instalaciones de MEHCCO S.A., Pasaje Ricardo Güiraldes 4567, San Martín— entre los eventos a cumplir el próximo mes. La especialidad Fabricación sigue en 0% con 25% pronosticado para noviembre.",
      riesgo="El inicio se corre al menos cuatro semanas más (acumula 7,5 semanas sobre la Rev6) y ABB no reemitió el cronograma: la fecha de fin de armado del 17-12-2026 ya no se sostiene aritméticamente. Cincuenta días hábiles desde el 02-11-2026, con el receso de fin de año, terminan alrededor del 12-01-2027.",
      impacto="Arrastra las pruebas internas pre-FAT, la FAT del 25 al 29-01-2027 y, con ella, la entrega de equipamiento en shelter del 17-02-2027, que es predecesora del montaje del PMS en ambas salas. El plan de certificación acordado ya ubica el hito «contra finalización de la FAT» en feb-27 y «contra entrega de equipos» en mar-27: el propio esquema comercial de ABB está un mes por detrás del cronograma técnico.",
      real="El armado debería estar en marcha. Al 07-11-2026 no hay confirmación de que haya arrancado ni de la fecha concreta.",
      accion="Pedir a ABB la fecha real de inicio del armado en MEHCCO y la reemisión del cronograma del PMS con la curva de fabricación recalculada, indicando si sostiene el 17-12-26 y con qué medida (turnos adicionales, solape con las pruebas internas). Definir antes del 30-11-26 si la FAT del 25-01-27 se mantiene.",
      sev="ALTA"),
 dict(cod="PA-19", tit="Tres FAT de ABB y la prueba de integración PCS↔PMS se superponen entre el 18-11 y el 16-12-2026",
      inter="AESA Calidad / PPSA → ABB Salas (Santa Fe, Sorocaba, AMBA) + ABB PMS / Inauco (Buenos Aires)",
      plan="La previsión de FAT del informe Nº6 fija: sistema UPS y baterías 18 al 20-11-2026 en Santa Fe (invitación enviada) · tableros TGMT 30-11 al 04-12-2026 en Sorocaba, Brasil (invitación enviada) · tableros CCM 30-11 al 12-12-2026 en Sorocaba (invitación enviada) · tableros auxiliares 14 al 16-12-2026 en AMBA. En paralelo, el informe Nº5 del PMS fija la prueba de integración con la maqueta de CPF2 para la semana del 30-11-2026 en Buenos Aires, con participación de ABB, AESA, PLUSPETROL e Inauco.",
      riesgo="El 30-11 y el 01-12 hay tres eventos simultáneos en dos países: dos FAT en Sorocaba y la prueba de integración en Buenos Aires. Las tres requieren witness de AESA y, en el caso de la integración, también de PPSA y de personal de Inauco. Con una sola dotación de Calidad se cae alguna de las tres o se acepta sin presenciar.",
      impacto="Una FAT sin witness obliga a repetirla o a liberar contra protocolo documental, lo que traslada el riesgo al SAT en sitio. Postergar la prueba de integración empuja la configuración de comunicaciones del PMS, que ya es el eslabón más atrasado de esa provisión.",
      real="Las invitaciones de las tres FAT de noviembre y diciembre ya fueron enviadas por ABB.",
      accion="Armar antes del 14-11-2026 el calendario único de witness del paquete, con nombre y asignación por evento, y resolver la colisión del 30-11: adelantar la FAT de TGMT, correr la maqueta CPF2 a la semana del 07-12 o asignar dos equipos. Incluir los viajes a Brasil en la planificación de Calidad.",
      sev="ALTA"),
 dict(cod="PA-20", tit="El módulo adicional de entradas del PMS está fuera del alcance de la orden de compra y la prueba de integración lo necesita",
      inter="AESA Suministros → ABB (oferta OPP-26-8808080 Rev.0, típico DOL_TIPICO-01A)",
      plan="El informe Nº5 deja constancia de que el 23-07-2026 AESA aceptó y solicitó avanzar con la gestión interna de ABB por el módulo adicional de entradas, para no impactar y poder realizar las pruebas de integración entre sistemas en los plazos solicitados (noviembre / diciembre). ABB aclara expresamente que «la gestión de esa PO y la ejecución de la misma no forma parte del presente alcance».",
      riesgo="La prueba de integración de la semana del 30-11-2026 se hace con la maqueta de CPF1 más los cubicles de UMC propios de CPF2. Si el módulo adicional no está comprado y disponible, la prueba se hace sobre una configuración incompleta y no cubre las señales que el módulo aporta.",
      impacto="Una integración validada parcialmente obliga a repetir el ensayo o a descubrir diferencias recién en la FAT de enero, con el armado ya cerrado.",
      real="Al 07-11-2026 no consta en el expediente la emisión de la orden de compra por la oferta OPP-26-8808080 Rev.0.",
      accion="Verificar en Suministros el estado de esa orden de compra. Si no está colocada, decidir antes del 20-11-2026 si se coloca o si la prueba de integración se acota formalmente, dejando registrado el alcance que no se verifica.",
      sev="MEDIA"),
 dict(cod="PA-21", tit="Los procedimientos de FAT y SAT del PMS se corren a noviembre y quedarían aprobados a días de la FAT",
      inter="ABB PMS → AESA Ingeniería / Calidad",
      plan="La Rev12 preveía la emisión de los procedimientos de FAT y SAT Rev.A entre el 19 y el 30-10-2026. El informe Nº5 los ubica en los eventos del próximo mes, dentro del desarrollo de ingeniería básica: protocolo de aceptación en fábrica y su registro, y protocolo de aceptación en sitio y su registro. Arrastrando los plazos de la cadena, la Rev.0 quedaría aprobada hacia el 15-01-2027.",
      riesgo="La FAT está prevista del 25 al 29-01-2027: el protocolo con el que se va a ensayar quedaría aprobado seis días hábiles antes. No hay margen para discutir criterios de aceptación ni para que AESA incorpore requisitos de PPSA.",
      impacto="Un protocolo aprobado sobre la fecha suele derivar en observaciones durante la propia FAT y en ensayos que se repiten. Además, ABB informa que el protocolo de la FAT integral de las salas en planta shelterista se emitirá «en los próximos meses», sin fecha.",
      real="Al 07-11-2026 no se recibieron los procedimientos Rev.A.",
      accion="Pedir a ABB fecha comprometida de emisión de la Rev.A antes del 30-11-2026 y fijar un plazo de revisión de AESA de diez días hábiles. Reclamar en la misma nota la fecha de emisión del protocolo de la FAT integral del shelter.",
      sev="MEDIA"),
 dict(cod="PA-17", tit="La fabricación de los núcleos de los transformadores está tercerizada y el subproveedor no fue inspeccionado",
      inter="CAT Miron → subproveedor (Villa Martelli) → AESA Calidad",
      plan="El acta de inspección del 21-10-2026 informa que TODO el proceso de fabricación de núcleos a partir de láminas de FeSi está tercerizado en una empresa afín a CAT ubicada en Villa Martelli: allí se cortan las chapas y se conforman los núcleos ventana, que recién después se llevan a la planta de CAT para el armado de bobinas, yugos y cabinados. El acta deja constancia de que ese taller «será objeto de inspección en futuras visitas».",
      riesgo="Es el único eslabón de la cadena que está fuera de la planta de CAT y el único todavía sin inspeccionar. El núcleo gobierna las pérdidas en vacío y la corriente de excitación, que son dos de los ensayos con punto de retención del PIE: un defecto detectado recién en el ensayo final obliga a rehacer el núcleo, con el despacho ya comprometido para noviembre y diciembre.",
      impacto="La fabricación de núcleos corre del 07-09 al 06-11-2026 y es predecesora del montaje final. Si hay que corregir algo del núcleo después de esa fecha, el impacto cae directamente sobre las cuatro ventanas de entrega escalonadas.",
      real="La ventana útil para inspeccionar el taller de Villa Martelli se cierra con el fin de la fabricación de núcleos, el 06-11-2026.",
      accion="Programar con Calidad la visita al taller de Villa Martelli dentro de las próximas dos semanas, antes de que termine la fabricación de núcleos. Pedir a CAT los datos del subproveedor y su plan de control de recepción de FeSi, y verificar la trazabilidad de los certificados Legnano Teknoelectric contra los núcleos efectivamente conformados.",
      sev="MEDIA"),
 dict(cod="PA-16", tit="Insumos de AESA a Inauco vencidos: dos órdenes de compra sin emitir y dos entregas de ingeniería incompletas",
      inter="AESA → Inauco",
      plan="El cronograma Rev 0 de Inauco desdobla el Requisito 0 en tres órdenes de compra, todas con fecha 28-07-2026: MO + materiales (100%, recibida), SEGUROS (0%) y MANO DE OBRA DE CAMPO (0%). Las dos últimas llevan TRES MESES vencidas. En el Requisito 1, la entrega de ingeniería del 21-08-26 (diagramas de conexionado de JB, diagrama de bloques y patcheo de fibra óptica) está al 66% y la del 09-10-26 (filosofía de operación y matriz causa-efecto) al 50%; la del 30-10-26 vence esta semana e incluye la documentación de HIMA (tabla Modbus y print de configuración).",
      riesgo="Las dos OC pendientes no frenan hoy la fabricación en taller, pero sí condicionan la cobertura de seguros y la contratación de la mano de obra de campo, que arranca el 18-05-27. Las entregas de ingeniería incompletas sí impactan de inmediato: el freezing point de aprobación de ingeniería para construcción es el 19-11-26, en tres semanas, y la documentación de HIMA que vence el 30-10 depende a su vez de HPH y de HARG.",
      impacto="Inauco declara un avance global del 6% y de configuración del 11% a cuatro meses de iniciado el contrato. La provisión no muestra corrimientos de entrega, pero su margen depende de que los insumos de AESA se normalicen antes del freezing point del 19-11-26.",
      real="Las dos OC: inmediato, llevan 3 meses de atraso. Ingeniería: completar las entregas del 21-08 y 09-10 antes del 19-11-26 y cerrar la del 30-10-26.",
      accion="(1) Emitir las OC de seguros y de mano de obra de campo. (2) Cerrar las dos entregas de ingeniería incompletas antes del freezing point. (3) Destrabar la documentación de HIMA (tabla Modbus y print de configuración) con HPH y HARG, que vence el 30-10-26 — se cruza con PA-11, donde ABB reclama los mismos mapas Modbus y Profinet.",
      sev="ALTA"),
 dict(cod="PA-14", tit="Letter of Credit de la orden de compra ampliada de HPH — a cargo de AESA, VENCIÓ el 30-10-2026",
      inter="AESA → HPH (HIMA Alemania)",
      plan="El Supplier Progress Report Nº 01 de HPH (15-10-2026) lista un ÚNICO pendiente prioritario y lo imputa a AESA: «Letter of Credit for the amended PO», con plazo solicitado el 30-10-2026. HPH ya procesó la orden ampliada y reservó el material adicional en su ERP.",
      riesgo="EL PLAZO VENCIÓ. Al 07-11-2026 no consta la emisión de la carta de crédito y HPH no emitió informe de avance posterior al Nº1. Sin ella HPH no libera el material adicional de la orden ampliada, con la fabricación de los tres tableros principales ya terminada y la inspección de la semana 44 o 45 sin resultado reportado.",
      impacto="Compromete la liberación y el despacho del alcance ampliado, y por lo tanto la entrega DAP del 15-12-26 que hoy está confirmada sin desvío. Es un trámite administrativo del EPC, no un problema técnico ni del proveedor.",
      real="30-10-2026, plazo pedido por HPH: VENCIDO hace ocho días sin confirmación de emisión.",
      accion="Confirmar con Finanzas y Suministros si la carta de crédito se emitió. Si no, gestionarla esta semana y pedir a HPH el informe de avance de octubre, que no se recibió, con el resultado de la inspección de la semana 44/45 y el estado del material de la orden ampliada.",
      sev="ALTA"),
 dict(cod="PA-15", tit="HPH no reporta estado de liberación, despacho ni plan de certificación",
      inter="HPH → AESA",
      plan="En el informe Nº 01 de HPH, la sección 4.4 «Release and Dispatch» está en blanco y la sección 9 «Certification» no tiene ningún hito cargado. Tampoco se declara ningún factor potencial de retraso (sección 6 vacía).",
      riesgo="La fabricación está terminada y el despacho debería prepararse en noviembre, pero no hay información sobre el plan de liberación, el packing, el modo de transporte ni los hitos de certificación de la orden. El «remote HW check» además figura con fecha a definir.",
      impacto="Sin esos datos no se puede planificar el witness de la liberación, ni anticipar la documentación de embarque para la nacionalización (ver PA-01), ni prever el flujo de certificaciones. Es el tramo donde la provisión pasa de fabricación a logística, justamente el que concentra el riesgo.",
      real="A pedir en el informe del próximo período, con el cierre del 30-10-2026.",
      accion="Pedir a HPH que complete en el próximo informe: fecha del remote HW check, plan de liberación e inspección final de despacho, modo de transporte y agente de carga, y el cuadro de hitos de certificación con montos. Aprovechar para pedir que evalúe ADELANTAR EL DESPACHO: con la fabricación terminada hay unas 3 semanas de holgura entre el fin de fabricación y el despacho previsto.",
      sev="MEDIA"),
 dict(cod="PA-13", tit="Suministro de los equipos físicos de control de acceso de ambas salas, a cargo de AESA",
      inter="AESA → ABB Salas",
      plan="El informe Nº 6 lo reitera entre los tres temas pendientes prioritarios: «Control de acceso. Suministro de equipos físicos de acuerdo con cronograma de cada sala — Resp. AESA». El cronograma fija el suministro el 04-01-2027 para SE#4 y el 13-02-2027 para SE#3. Es la segunda vez consecutiva que ABB lo reclama.",
      riesgo="Es un suministro de AESA que se integra dentro de la fabricación del shelter, antes de la instalación eléctrica. Si llega tarde obliga a reabrir trabajos ya cerrados en el shelterista o a entregar la sala incompleta.",
      impacto="SE#4 instala en enero y SE#3 en febrero de 2027; ambas fechas caen sobre la instalación eléctrica del shelter. Un atraso aquí golpea directamente la FAT de sala y, por lo tanto, el despacho EXW.",
      real="04-01-2027 (SE#4) y 13-02-2027 (SE#3): la compra debe estar colocada con antelación suficiente.",
      accion="Confirmar con Suministros el estado de la compra de los equipos de control de acceso y su plazo de entrega, y comprometer con ABB la fecha de puesta a disposición en el shelterista. Quedan ocho semanas para la fecha de SE#4.",
      sev="ALTA"),
 dict(cod="PA-11", tit="Temas pendientes prioritarios que ABB reclama a AESA para poder configurar el PMS",
      inter="AESA Ingeniería → ABB PMS",
      plan="El informe de avance Nº 5 de ABB (30-10-2026) REITERA sin cambios las tres definiciones pendientes: (a) señales cableadas de cada equipo con indicación de las borneras de vinculación al PMS (transformadores, generadores, seccionador bajo carga); (b) mapa Modbus y definición de señales de los equipos a integrar (UPS, generadores, multimedidores); (c) mapas Profinet y definición de señales de UMC y VFD. Es el segundo mes consecutivo que figuran como pendientes.",
      riesgo="La configuración de variables de comunicación (IEC61850 / PROFINET / MODBUS) debía arrancar el 02-11-2026 y la especialidad Configuración sigue en 0%, con apenas 25% pronosticado para noviembre. Sin esas definiciones ABB no puede configurar ni dejar listo el sistema, y la prueba de integración del 30-11 se haría sobre una configuración parcial.",
      impacto="Es un insumo de AESA, no un atraso del proveedor. Si no se cierra en octubre, se comprime la configuración y se pone en riesgo la cadena pruebas internas (18-12) → FAT (25-29/01/27) → despacho (17-02-27), que hoy no tiene holgura.",
      real="VENCIDAS: debían estar cerradas antes del 02-11-2026 y al 07-11-2026 siguen abiertas.",
      accion="Asignar responsable en Ingeniería por cada uno de los tres paquetes y acordar fecha de entrega con ABB en la reunión semanal de seguimiento. Priorizar el mapa Modbus y los mapas Profinet, que dependen de información de terceros (UPS, generadores, UMC, VFD).",
      sev="ALTA"),
 dict(cod="PA-12", tit="Una sola FAT integral de 5 días frente a las dos ventanas de ensayos PMS del cronograma de salas",
      inter="ABB PMS ↔ ABB Salas ↔ AESA",
      plan="El informe Nº5 del PMS ratifica que ABB y AESA acordaron una FAT INTEGRAL de las salas EN PLANTA SHELTERISTA, de 5 días, verificando una muestra representativa de señales, y aclara que el protocolo de ensayos «será definido por ABB y emitido en los próximos meses», sin fecha. El cronograma del PMS la ubica del 25 al 31-03-2027. El informe Nº 6 de las salas, en cambio, mantiene DOS FAT de sala separadas en Mendoza: SE#4 del 22 al 26-02-2027 y SE#3 del 29-03 al 02-04-2027, más dos ventanas de «Ensayos PMS en sala» (SE#4 03-09/03 y SE#3 30/03-06/04/27).",
      riesgo="La FAT integral del PMS (25-31/03/27) coincide con la FAT de SE#3 en Mendoza (29/03-02/04/27), de modo que para la sala BT ambas lecturas convergen. Pero NO para SE#4: su FAT es del 22 al 26-02-27 y su despacho EXW el 25-03-27, un mes ANTES de la FAT integral del PMS.",
      impacto="Si prevalece una única FAT integral del 25 al 31-03-27, la sala MT ya estaría despachada y no podría participar. Si prevalecen dos campañas, ABB PMS debe rever el alcance y la duración de 5 días acordados, y prever una asistencia en febrero que hoy no tiene cargada.",
      real="Sin conciliar al 07-11-2026. El protocolo de ensayos no tiene fecha de emisión comprometida.",
      accion="Llevar el punto a la reunión semanal de seguimiento ABB/AESA y fijar el criterio antes de la emisión del protocolo: una campaña integral o dos por sala, y su encaje con los despachos EXW de cada shelter.",
      sev="ALTA"),
 dict(cod="PA-09", tit="Los ensayos con punto de retención de los transformadores no tienen fecha en el cronograma de CAT Miron",
      inter="CAT Miron → AESA / PPSA",
      plan="AVANCE RESPECTO DE LA REV11: el acta de inspección del 21-10-2026 adjunta el PLAN DE INSPECCIÓN Y ENSAYOS 5155-00-0022-VE-CA-PE-001 Rev.A, que ya define QUÉ se ensaya y con qué nivel de intervención. Son 14 puntos de retención obligatorios para el cliente, con 10 días de aviso previo: control dimensional (3 puntos), nueve ensayos eléctricos (resistencia óhmica, resistencia de aislación, tensión aplicada, tensión inducida, relación de transformación y grupo de conexión, pérdidas en vacío y corriente de excitación, pérdidas y tensión de cortocircuito, ensayo del conmutador, control de instrumentos), despacho (marcación de puntos de izaje, control de marcación y preparación) y la documentación final.",
      riesgo="LO QUE SIGUE FALTANDO ES LA FECHA. El cronograma de fabricación mantiene la aclaración de que no incluye las pruebas y que se coordinan según disponibilidad de laboratorio, y sólo contempla 2 días hábiles de ensayos internos por orden de trabajo entre el 02-11 y el 04-12-26. Con 5 órdenes de trabajo liberándose en cinco semanas hacen falta 5 ventanas, y cada una exige avisar con 10 días de anticipación.",
      impacto="Para cumplir el preaviso de 10 días del primer lote (entrega 13-11-26) la convocatoria tendría que estar cursada esta semana. Es el único hito de la provisión sin fecha comprometida y está inmediatamente antes de cada entrega.",
      real="A definir con CAT Miron: una fecha por cada una de las 5 OT, cada una con 10 días de aviso previo.",
      accion="Pedir a CAT Miron una revisión del cronograma con las fechas de ensayo incorporadas y las convocatorias formales de los puntos de retención. Confirmar quién asiste por AESA y por PPSA, y verificar que los protocolos de ensayo estén aprobados antes de la primera ventana.",
      sev="ALTA"),
 dict(cod="PA-10", tit="Los transformadores llegan siete meses antes que los ductos de barras que los vinculan",
      inter="CAT Miron ↔ ABB Ductos de barras ↔ AESA Construcciones",
      plan="CAT Miron entrega los 13 transformadores entre el 13-11 y el 11-12-2026. Los ductos de barras que los vinculan a las salas eléctricas entregan en sitio el 08-07-2027 (ver PA-07). DATO NUEVO del acta del 21-10-26: 11 unidades son SECAS CABINADAS y 2 —los elevadores de 2.000 kVA 0,4/13,86 kV de la OT 5212, máquinas 3497 y 3498— son EN BAÑO DE ACEITE.",
      riesgo="Desfase de aproximadamente 7 meses entre ambas entregas. Los transformadores quedan almacenados en obra sin poder vincularse, durante el verano y buena parte del otoño. Los dos elevadores en baño de aceite tienen requisitos de conservación DISTINTOS de los secos: control de nivel y de estanqueidad, y ensayo de rigidez dieléctrica del aceite antes de energizar.",
      impacto="Requiere prever superficie de almacenamiento, fundaciones o playa de acopio, y DOS planes de conservación distintos: para las 11 unidades secas, control de humedad ambiente y limpieza; para las 2 en baño de aceite, control de nivel, estanqueidad y verificación del aceite. Sin esos planes hay riesgo de observaciones en la recepción y de ensayos de re-verificación antes de energizar.",
      real="Almacenamiento efectivo: de nov-dic 2026 a jul 2027 como mínimo.",
      accion="Definir con Construcciones el sitio y las condiciones de acopio; solicitar a CAT Miron el procedimiento de almacenamiento prolongado y su garantía en esa condición; incorporar las verificaciones periódicas al plan de calidad. Evaluar si conviene escalonar la recepción en lugar de recibir las 13 unidades entre noviembre y diciembre.",
      sev="MEDIA"),
 dict(cod="PA-07", tit="Ductos de barras: entrega más tardía del paquete, plazo sin confirmar y pedido de ABB de adecuar la fecha contractual",
      inter="AESA → ABB",
      plan="La entrega en sitio pasó del 10-06-27 al 08-07-27. Los informes de avance Nº 5 (05-10-26) y Nº 6 (04-11-26) de ABB explicitan la causa: los ductos NO PUDIERON DISEÑARSE porque no se disponía de información suficiente sobre los gabinetes de los transformadores de media tensión y sus posiciones relativas. El diseño arrancó efectivamente el 21-09-26, al confirmarse esa información, y la tabla de subproveedores del Nº6 muestra que la OC 4700201354 a EAE (Turquía) sigue con fecha de entrega en origen «tbd».",
      riesgo="ABB declara que está trabajando en la confirmación del plazo de entrega y del POSIBLE IMPACTO ECONÓMICO, y pide formalmente que AESA ADECUE LA FECHA DE ENTREGA CONTRACTUAL de las posiciones 120, 130 y 140 de la OC 45089944971 para evitar penalidades por incumplimiento. Además, el subproveedor de los ductos —EAE, Turquía, documento de compra 4700201354— figura con fecha de entrega en origen «tbd», es decir sin comprometer.",
      impacto="La entrega en sitio del 08-07-27 es el hito más tardío de TODO el paquete, por delante del cierre del alcance HARG (10-05-27), y hoy ni siquiera está firme: el plazo definitivo depende de la confirmación de ABB y de EAE. Hay además una exposición contractual y económica abierta. El avance del paquete es del 5% (ingeniería 15%, fabricación y transporte en 0%).",
      real="A confirmar por ABB. El transporte marítimo pesa el 34% del paquete y el traslado + aduana insume 80 días corridos, sin margen de compresión.",
      accion="(1) Responder formalmente el pedido de adecuación de la fecha contractual de las posiciones 120/130/140 y acotar la exposición económica. (2) Exigir a ABB la fecha comprometida de EAE, hoy «tbd». (3) Proteger las dos aprobaciones que quedan: el diseño cierra el 14-12-26 y la aprobación de AESA vence el 30-12-26, en pleno receso — asignar el revisor con anticipación. (4) Trazar la información de gabinetes de transformadores con CAT Miron para que no se repita.",
      sev="ALTA"),
 dict(cod="PA-08", tit="Aprobación de la ingeniería de las SALAS vencida: el freezing point de ingeniería básica venció el 25-09-26 y sigue sin cerrar",
      inter="AESA → ABB (Salas SE#3 y SE#4)",
      plan="El informe Nº 6 (04-11-2026) registra como HITO DEL PERÍODO la aprobación de la ingeniería de detalle de TODOS LOS EQUIPOS de ambas salas —tableros auxiliares, CCM, tableros de MT, sistema UPS y baterías y variadores de MT y BT—, emitida y aprobada en revisión numérica para fabricación. El punto queda acotado a la SALA propiamente dicha: ABB declara que tanto la ingeniería básica como la de detalle del shelter tienen APROBACIÓN PARCIAL y que «se observan nuevos comentarios en las sucesivas revisiones que impiden la aprobación final y la emisión de ingeniería para fabricación».",
      riesgo="Al corte hay SEIS SEMANAS de atraso sobre el freezing point de ingeniería básica (25-09-26) y veinticinco días sobre el de detalle (13-10-26). ABB lo incluye expresamente entre los factores potenciales de retraso y pide «evitar la generación de nuevos comentarios sobre documentos ya emitidos para fabricación»: el ciclo de revisión de AESA se volvió el problema, no la emisión del proveedor. La certificación del hito 2 se corrió de octubre a noviembre.",
      impacto="Acotado pero ya con consecuencia visible: la ingeniería de los equipos está cerrada y la fabricación en Brasil avanza según plan, pero la sala eléctrica 3 es el paquete de menor avance de toda la provisión (34% contra 37% previsto para noviembre) y su recepción de materiales está en 18%. La fabricación del shelter de SE#3 arranca el 23-11-26.",
      real="Al 07-11-2026 ninguna de las dos aprobaciones está cerrada: ambas con aprobación parcial y comentarios abiertos.",
      accion="Congelar el alcance de la revisión: una única reunión de cierre con ABB para consolidar TODOS los comentarios pendientes en un solo listado, con compromiso de no abrir observaciones nuevas sobre documentos ya emitidos para fabricación, y aprobación formal antes del 20-11-26 para no comprometer el arranque de la fabricación del shelter de SE#3 del 23-11-26.",
      sev="ALTA"),
 dict(cod="PA-04", tit="Orden de compra de HIMA Argentina (HARG): se posterga otras dos semanas y el corrimiento ya es de nueve",
      inter="AESA → HARG",
      plan="El cronograma PSCH_CPF2_Rev3 arrancaba los trabajos de HARG el 14-09-2026. La Rev12 asumía la colocación de la OC el 01-11-2026. Esa fecha pasó sin colocarse y AESA informa que la orden NO sale esta semana sino la siguiente: se adopta el 16-11-2026, con inicio de trabajos el 17-11-2026. Son DOS SEMANAS ADICIONALES sobre la premisa anterior y nueve semanas corridas (63 días) sobre el plan original del proveedor.",
      riesgo="Ya no hay margen que consumir: el corrimiento se traslada día por día al cierre del alcance. Esta postergación mueve TODAS las fechas propias de HARG catorce días y empuja el cierre del alcance del 10-05-2027 al 24-05-2027. Es el primer corrimiento de esta revisión que se traduce en una fecha final distinta.",
      impacto="El cierre del alcance HARG pasa al 24-05-2027 y sigue siendo el FIN DEL CAMINO CRÍTICO, ahora veinte días DESPUÉS de la liberación de los tableros de Inauco (04-05-27). La FAT de HARG cae el 28-30/04/27, después del cierre de la FAT SIS conjunta del 26-04-27, y el IFAT se corre al 03-14/05/27. Cada semana adicional sigue sumando una semana: 31-05-27 con una semana más, 07-06-27 con dos.",
      real="Colocación de la OC: semana del 16-11-2026 según lo informado por AESA. Inicio de trabajos: 17-11-2026. Al 07-11-2026 la orden no está emitida.",
      accion="PRIORIDAD 1: colocar la OC en la semana del 16-11-26 sin nueva postergación. En paralelo, y para no perder las dos semanas que ya se fueron, emitir de inmediato una orden de inicio anticipada o carta de intención que habilite a HIMA a largar la ingeniería de aplicación y la documentación mientras se formaliza la OC: los insumos de HPH estarán disponibles, de modo que parte del corrimiento es recuperable resecuenciando — pero sólo si el equipo de HARG está trabajando. Cada semana adicional suma una semana al cierre del paquete.",
      sev="ALTA"),
 dict(cod="PA-05", tit="La adecuación de HARG cae dentro de la Pre iFAT de Inauco y su cierre queda después de la liberación de tableros",
      inter="HARG ↔ Inauco",
      plan="Con el corrimiento de 63 días, HARG adecúa los tableros de marshalling del 22-02 al 02-04-27 en la base de Inauco. El cronograma PG-002 Rev.0 de Inauco fija la Pre iFAT del 10-02 al 26-03-27 y la liberación de tableros el 04-05-27. La FAT de HARG pasa al 28-30/04/27, el IFAT al 03-14/05/27 y el cierre del alcance al 24-05-27.",
      riesgo="Los tres solapamientos EMPEORAN con la nueva postergación: (a) la adecuación de HARG ocupa físicamente los tableros desde el 22-02 hasta el 02-04-27, es decir durante casi toda la Pre iFAT de Inauco; (b) la FAT de HARG (28-30/04/27) y su IFAT (03-14/05/27) quedan ahora ENTERAMENTE DESPUÉS del cierre de la FAT SIS conjunta del 26-04-27, con lo que la FAT SIS se haría sin el alcance HARG validado; (c) el cierre del alcance HARG (24-05-27) queda veinte días después de la liberación de los tableros de Inauco.",
      impacto="Inauco liberaría los tableros casi tres semanas antes de que HIMA cierre su IFAT y su documentación, con lo que el sistema de seguridad viajaría a obra sin el alcance HARG terminado. La secuencia que la Rev12 describía como ajustada pasa a ser inconsistente: ya no es un solapamiento, es una inversión del orden lógico.",
      real="Al 07-11-2026 no hay secuencia acordada entre Inauco y HIMA, y la premisa de fechas de HARG volvió a moverse.",
      accion="Convocar a Inauco y HIMA a acordar la secuencia y el uso del taller; evaluar adelantar la adecuación de HARG (depende de PA-04) o correr la liberación de tableros al cierre real del alcance HARG. Definir si la FAT SIS conjunta se extiende para cubrir el IFAT de HARG.",
      sev="ALTA"),
 dict(cod="PA-01", tit="Nacionalización de los tableros de marshalling entregados por HPH",
      inter="HPH (Alemania) → HARG / Inauco",
      plan="HPH confirma la entrega el 15-12-26 en condición DAP EZEIZA AEROPUERTO, sin desvío respecto de la fecha contractual en ninguno de los tres tableros. La recepción en base Inauco está prevista para el 25-01-27 y, con el nuevo corrimiento de HARG, la adecuación arranca recién el 22-02-27.",
      riesgo="El despacho de aduana de un embarque de tableros importados demanda típicamente 3 a 4 semanas y el arribo cae en el receso de fin de año. DATO NUEVO: la entrega es en EZEIZA AEROPUERTO, lo que apunta a carga aérea y a un despacho más rápido que un embarque marítimo — pero el modo de transporte no está confirmado en el informe (sección 4.4 en blanco, ver PA-15).",
      impacto="Severidad MEDIA: entre la recepción prevista (25-01-27) y el inicio de la adecuación de HARG (22-02-27) hay ahora CUATRO semanas de holgura, el doble que en la Rev12. Pero esa holgura la pagó la postergación de la OC de HARG: no es una mejora del plan sino el reflejo de un atraso aguas arriba, y no debe consumirse para absorber otros atrasos.",
      real="Recepción en base Inauco entre el 25-01-27 y la primera semana de febrero de 2027.",
      accion="Confirmar con HPH el modo de transporte y el agente de carga, designar despachante y anticipar la documentación de embarque. Con la fabricación ya terminada, evaluar ADELANTAR EL DESPACHO para ganar semanas sobre la nacionalización. La holgura ganada NO debe consumirse para absorber otros atrasos.",
      sev="MEDIA"),
 dict(cod="PA-02", tit="Inauco recibe los tableros de HIMA sin adecuar, contra lo previsto en su Requisito 2",
      inter="HPH ↔ HARG ↔ Inauco",
      plan="Inauco fija en su Requisito 2 la recepción de los tableros HIMA en su base el 25-01-27, asumiendo el equipo prácticamente terminado. Con el nuevo corrimiento, HARG recién adecúa entre el 22-02 y el 02-04-27.",
      riesgo="La premisa de Inauco de recibir el tablero «casi terminado» no se cumple: recibe el tablero de HPH sin adecuar y la adecuación se ejecuta después, en su propia base.",
      impacto="Inauco debe prever espacio de taller, energía y acceso para que HIMA trabaje sobre los tableros durante seis semanas, en paralelo con su propia Pre iFAT (ver PA-05). Además, los tableros quedan cuatro semanas almacenados en su base antes de que HIMA los toque. La Configuración PLC Etapa 2 (25-01 → 23-03-27) convive con la adecuación.",
      real="Tablero disponible ya adecuado: 02-04-2027 (Rev12: 19-03-2027).",
      accion="Reconfirmar con Inauco el alcance real del Requisito 2 y acordar las condiciones de trabajo de HIMA en su base. Evaluar, con HPH, si parte de la adecuación puede ejecutarse en origen antes del despacho.",
      sev="ALTA"),
 dict(cod="PA-03", tit="CERRADO — Prueba de integración PCS↔PMS en Buenos Aires: fecha y duración confirmadas por ambas partes",
      inter="Inauco (PCS) ↔ ABB (PMS)",
      plan="Anticipada en la Rev3 a la ventana OCT-NOV 2026 sin fecha; fijada por el informe Nº4 y RATIFICADA por el informe Nº5 (30-10-2026): semana del 30 de noviembre de 2026, dos días de duración, con la maqueta de CPF1 más los cubicles de UMC propios de CPF2, en conjunto ABB / AESA / PLUSPETROL e INAUCO por la comunicación con el PCS.",
      riesgo="La fecha está acordada por ambas partes. Quedan dos condiciones abiertas que se siguen en otros puntos: la disponibilidad del módulo adicional de entradas, que está fuera del alcance de la OC (PA-20), y la colisión con las FAT de TGMT y CCM en Sorocaba, que arrancan el mismo 30-11 (PA-19).",
      impacto="Anticiparla es favorable: detecta tempranamente incompatibilidades de comunicación (IEC61850 / Modbus / Profibus) antes de las FAT. Sin fecha acordada el ítem no puede planificarse ni asignarse recursos, y si se pierde la ventana vuelve a caer después de las FAT.",
      real="Semana del 30-11-2026, 2 días. El acopio de materiales del PMS, condición previa, quedó verificado al 100% el 22-10-2026.",
      accion="Emitir el protocolo de la prueba y confirmar la nómina de participantes de AESA, PPSA e Inauco. Resolver la colisión de agenda con Sorocaba (PA-19) y el estado del módulo adicional (PA-20).",
      sev="MEDIA"),
 dict(cod="PA-06", tit="CERRADO — El rack de comunicaciones CPF-COM-SE3 salió del alcance de construcción de Inauco",
      inter="AESA → Inauco",
      plan="El cronograma PG-002 Rev.0 lista el rack CPF-COM-SE3 (19\" 40U) con fechas de fabricación 22-01 → 03-02-27, pero marcado «A Confirmar Fabricación».",
      riesgo="RESUELTO POR ELIMINACIÓN. El cronograma Rev 0 del 23-10-26 ya no incluye el CPF-COM-SE3 entre los tableros a construir: sólo figuran CPF2-PCS-SC-07a, CPF2-PCS-SC-07b y CPF-COM-07.",
      impacto="El rack de comunicaciones de la Sala Eléctrica 3 queda FUERA del alcance de Inauco. Hay que confirmar quién lo provee —si pasa al alcance de ABB dentro de la sala o si se contrata por separado— para que no quede huérfano.",
      real="Cerrado en el cronograma de Inauco. Pendiente de reasignación del alcance.",
      accion="Confirmar con Ingeniería y Suministros a qué provisión pasa el CPF-COM-SE3, o si se da de baja definitivamente. Mientras no se defina, el rack no está en el alcance de nadie.",
      sev="MEDIA"),
]

# ---------------- Camino crítico (Rev3) ----------------
# hitos marcados como ruta crítica en el cronograma por hitos
CP_MARK={
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Inicio de contrato / proyecto"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Ingeniería de detalle / especificación"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Fin de ingeniería / CAD — FREEZING POINT"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","FABRICACIÓN DE LOS TABLEROS DE MARSHALLING"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Inspección / HW check de los 3 tableros"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Preparación de despacho (packing)"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Listo para despacho"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Despacho / shipping"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Entrega DAP Ezeiza AEROPUERTO (Argentina)"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","Nacionalización / despacho de aduana + transporte a base Inauco"),
 ("HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling","RECEPCIÓN DE TABLEROS EN BASE INAUCO"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","COLOCACIÓN DE LA OC POR AESA"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Documentación HARG — Doc. 1 (HW y otros)"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Programación HARG — Recurso 1"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Adecuación de tableros marshalling (INTERFASE Inauco)"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","FAT"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","IFAT (Recursos 1 y 2)"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","ENTREGA DEL ALCANCE HARG (IFAT + documentación)"),
 ("Inauco — PCS / SCADA / Comunicaciones","REQ 2 — Recepción tableros HIMA en base Inauco (insumo HPH)"),
 ("Inauco — PCS / SCADA / Comunicaciones","Configuración base — PLC y comunicaciones"),
 ("Inauco — PCS / SCADA / Comunicaciones","FAT SIS — integración con HIMA (Inauco Neuquén)"),
 ("Inauco — PCS / SCADA / Comunicaciones","LIBERACIÓN Y ENTREGA DE TABLEROS (camión de AESA, Inauco Neuquén)"),
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
 (11,"AESA→HARG","COLOCACIÓN DE LA OC DE HIMA ARGENTINA",D(2026,11,7),OC_HARG,"—","⚠ PA-04 — PENDIENTE al 07-11-26; AESA la estima para la semana del 16-11-26 (+2 semanas sobre la Rev12)"),
 (12,"HARG","Inicio de trabajos HARG — documentación y programación (R1 y R2)",SH(D(2026,9,14)),SH(D(2027,1,29)),"11","⚠ PA-04 — corrido +63 días (+14 d sobre la Rev12)"),
 (13,"HARG","Adecuación de tableros de marshalling en base Inauco (6 semanas)",SH(D(2026,12,21)),SH(D(2027,1,29)),"10, 12","⚠ PA-05 — 22/02 al 02/04/27: ocupa casi toda la Pre iFAT de Inauco"),
 (14,"Inauco","Configuración PLC — Etapa 2",D(2027,1,25),D(2027,3,23),"10",""),
 (15,"Inauco","PRE iFAT (integración interna Inauco)",D(2027,2,10),D(2027,3,26),"14","⚠ PA-05 — convive con la adecuación de HARG"),
 (16,"HARG","FAT",SH(D(2027,2,24)),SH(D(2027,2,26)),"13","⚠ PA-05 — 28-30/04/27: cae DESPUÉS del cierre de la FAT SIS conjunta"),
 (17,"Inauco","FAT CON EL CLIENTE — PCS + Comunicaciones",D(2027,3,29),D(2027,4,22),"15","Ventana confirmada por el crono. PG-002 Rev.0"),
 (18,"Inauco","FAT SIS (conjunta HARG / HPH / Inauco)",D(2027,3,29),D(2027,4,26),"15","⚠ PA-05 — se cerraría SIN el alcance HARG validado: la FAT de HARG es posterior"),
 (19,"HARG","IFAT (Recursos 1 y 2)",SH(D(2027,3,1)),SH(D(2027,3,12)),"16","⚠ PA-05 — 03 al 14/05/27: termina 18 días después de la FAT SIS"),
 (20,"Inauco","Configuraciones post-FAT",D(2027,4,23),D(2027,5,4),"17, 18",""),
 (21,"Inauco","LIBERACIÓN / ENTREGA DE TABLEROS (camión de AESA)",D(2027,5,4),D(2027,5,4),"20","⚠ PA-05 — 20 días ANTES del cierre del alcance HARG"),
 (22,"HARG","ENTREGA DEL ALCANCE HARG (IFAT + documentación)",SH(D(2027,3,12)),SH(D(2027,3,22)),"19","HITO FINAL DEL CAMINO CRÍTICO — 24-05-2027 (Rev12: 10-05-2027)"),
]
# ---------------- Informe de avance a SEP-2026 (Rev7) ----------------
# ABB PMS — informe Nº 5 (E-2611027, 30-10-2026): avance físico ponderado por especialidad
PMSK="ABB — PMS (Power Management System)"
# ABB SALAS — informe de avance Nº6 (R-2631036_A, 04-11-2026): avance ponderado por paquete
AVANCE_SALAS=[  # (paquete, sep-26 real, oct-26 previsto, nov-26 previsto) — informe Nº6
 ("Tableros CCM MNS III — 102-CCM-005/6/7","85% recep.","69%","85%"),
 ("Tableros de media tensión UNIGEAR — 102-TGMT-001/2","40% recep.","45%","62%"),
 ("Tableros auxiliares de distribución","80% recep.","75%","91%"),
 ("Sistema UPS y baterías","80% recep.","54%","90%"),
 ("Variadores de velocidad de BT — 102-DP-VFD-24210A/B/C","100% acopio","71%","98%"),
 ("Variadores de velocidad de MT — 102-DP-VFD-23310A/B/C","20% acopio","71%","92%"),
 ("Sala eléctrica 3 (shelter)","18% recep.","34%","37%"),
 ("Sala eléctrica 4 (shelter)","18% recep.","49%","55%"),
 ("Ductos de barras — 102-DB-005/006/007 A/B","15% ing.","15%","30%"),
]
# previsión de fechas y lugares de FAT (informe Nº6 + prueba de integración del informe PMS Nº5)
FAT_PREV=[
 ("Sistema UPS y baterías","Argentina","18/11/2026","20/11/2026","Santa Fe — Argentina · invitación enviada"),
 ("Tableros TGMT (media tensión)","Brasil","30/11/2026","04/12/2026","Sorocaba — Brasil · invitación enviada"),
 ("Tableros CCM","Brasil","30/11/2026","12/12/2026","Sorocaba — Brasil · invitación enviada"),
 ("Tableros auxiliares","Argentina","14/12/2026","16/12/2026","AMBA — Buenos Aires"),
 ("Sala eléctrica #4 (shelter completo)","Argentina","22/02/2027","26/02/2027","Mendoza — Argentina"),
 ("Sala eléctrica #3 (shelter completo)","Argentina","29/03/2027","02/04/2027","Mendoza — Argentina"),
 ("Maqueta integrada PMS / CCM / PCS — CPF2","Argentina","30/11/2026","01/12/2026","Buenos Aires · ABB + AESA + PPSA + Inauco (informe PMS Nº5)"),
]
# órdenes de compra de ABB a subproveedores (informe Nº6)
SUBPROV=[
 ("4500407392","Salas eléctricas 3 y 4","18/05/2026","Bottino","Argentina","14/05/27 y 08/04/27"),
 ("4500407396","Tableros auxiliares","18/05/2026","Mehcco","Argentina","28/01/2027"),
 ("4500407468","Sistema UPS y baterías","19/05/2026","KMT","Argentina","27/11/2026"),
 ("4700201354","EAE (ductos de barras)","16/05/2026","EAE","Turquía","tbd — SIN COMPROMETER"),
 ("4700201355","Tableros de media tensión","16/05/2026","BRABB","Brasil","27/11/2026"),
 ("4700201478","Tableros CCM","19/05/2026","BRABB","Brasil","25/01/2027"),
 ("4700201819","Variadores de baja tensión","26/05/2026","FIABB","Finlandia","06/11/2026"),
 ("4700201626","Variadores de media tensión","22/05/2026","CNABB","China","06/11/2026"),
]
CAT_TXT=[
 ("Identificación","Acta de inspección IN-FC-CAT-19082-261021 del 21-10-2026 (formulario AESA FGE-CA-039 Rev.2). Proveedor CAT Transformadores, obra 5155, orden de compra 4509019082. Inspector: Ing. Felice Corrao; por el proveedor, Fernando Jelicich. El cliente final no firmó el acta."),
 ("Resultado","Primera visita de inspección a planta. Fase: PROCESO DE FABRICACIÓN. Estado: «Otros» — no se emite dictamen de liberación. SIN informe de no conformidades. Provisión NO liberada y orden NO completa, como corresponde a una visita de seguimiento con la fabricación en curso."),
 ("Alcance confirmado","13 transformadores más 3 LOTES DE REPUESTOS para 2 años de operación (ítems 10, 120 y 160), que no figuraban en el cronograma ejecutivo. Once unidades son SECAS CABINADAS y dos —los elevadores de 2.000 kVA 0,4/13,86 kV YND11 de la OT 5212, máquinas 3497 y 3498— son EN BAÑO DE ACEITE. Números de máquina: 3489-3492 (OT 5209), 3493-3494 (OT 5210), 3495-3496 (OT 5211), 3497-3498 (OT 5212) y A7546-A7548 (OT 5213)."),
 ("Acopio verificado","En el depósito de materias primas se observó el acopio de foils de aluminio y de resina de dos componentes ya asignados a esta orden de compra. A pedido de la inspección se obtuvieron los certificados de calidad de origen de las tres materias primas principales: chapa de FeSi (Legnano Teknoelectric, según EN10107 e IEC60404), foils de aluminio y resina bicomponente (Krempel DMR 9230 RF, lote 019/26, aprobada en laboratorio)."),
 ("Fabricación de núcleos tercerizada","Todo el proceso de fabricación de núcleos a partir de láminas de FeSi se terceriza en una empresa afín a CAT ubicada en Villa Martelli: allí se cortan las chapas y se conforman los núcleos ventana, que luego se llevan a la planta de CAT para el armado de bobinas, yugos de cierre y cabinados. El acta deja constancia de que ese taller será objeto de inspección en futuras visitas. Ver PA-17."),
 ("Factor favorable","Las bobinas de los transformadores de distribución de 1,6 / 2 y 2,5 MVA están en buena medida en condición de stock permanente, lo que agiliza los tiempos de fabricación de las OT 5209, 5210 y 5211 — los tres lotes que concentran 8 de los 13 transformadores."),
 ("Plan de inspección y ensayos","Se adjunta el PIE 5155-00-0022-VE-CA-PE-001 Rev.A (01-09-2026) con 14 PUNTOS DE RETENCIÓN OBLIGATORIOS para el cliente, con 10 días de aviso previo: control dimensional, verificación de accesorios y dimensión final documentada; nueve ensayos eléctricos (control de instrumentos, resistencia óhmica, resistencia de aislación, tensión aplicada, tensión inducida, relación de transformación y grupo de conexión, pérdidas en vacío y corriente de excitación, pérdidas y tensión de cortocircuito, y ensayo del conmutador); marcación de puntos de izaje, control de marcación y preparación para el despacho; y los protocolos de ensayo de la documentación final. Ver PA-09."),
 ("Cronograma adjunto","El acta acompaña el cronograma de fabricación 5155-00-0022-VE-CO-PG-001 Rev.A. Se compararon sus barras y porcentajes contra la emisión del 29-09-2026 ya incorporada en la Rev6 y son IDÉNTICOS: no hay cambios de fecha en ninguna de las cinco órdenes de trabajo."),
]
HPH_TXT=[
 ("Identificación","Supplier Progress Report Nº 01 (plantilla AESA AC-PR-2026-002 Rev.3), emitido el 15-10-2026 por Frédéric Plounévez, Project Manager. Proveedor: HIMA Paul Hildebrandt GmbH. PO/NP 4509010341. Período informado: 01-08-2026 a 30-10-2026. Alcance: tableros de marshalling CPF-ESD-MC-07A, CPF-PSS-MC-07A y CPF-PSS-MC-07B."),
 ("Hito del período","FABRICACIÓN DE LOS TRES TABLEROS DE MARSHALLING COMPLETADA. El material de la orden de compra principal fue entregado. La orden ampliada se procesó y el material adicional quedó reservado en el ERP de HPH. El plan preveía fabricar hasta el 06-11-26: hay unas 3 semanas de adelanto."),
 ("Estado contractual","Entrega el 15-12-2026 en condición DAP (Incoterms 2020) EZEIZA AEROPUERTO, Argentina. Los tres tableros tienen fecha contractual igual a fecha objetivo: SIN DESVÍO declarado. La entrega en aeropuerto apunta a carga aérea, pero el informe no confirma el modo de transporte."),
 ("Ingeniería y documentación","Estado general 80%. No hay documentos vencidos de emisión ni documentos devueltos por AESA con comentarios a reemitir. A emitir en el próximo período: los tres diagramas de conexionado (DC-001/002/003 Rev.0) el 21-10-26, y el cronograma de ingeniería y suministro, los dos cálculos de consumo y disipación térmica (ESD&FGS y PSS) y los procedimientos de FAT y SAT el 30-10-26. En revisión de AESA desde el 23-10-26: cronograma de ingeniería y suministro Rev.A y cálculo térmico ESD&FGS Rev.A."),
 ("Inspecciones","Inspección de los tres tableros prevista para la SEMANA 44 o 45 de 2026 (26-10 al 08-11-26), sin fecha confirmada ni resultado. Se suma un «remote HW check» con fecha a definir."),
 ("Actividades del próximo período","Finalizar la emisión de los documentos de ingeniería y definir la fecha del remote HW check."),
 ("Factores potenciales de retraso","HPH no declara ninguno (sección 6 del informe en blanco)."),
 ("Temas pendientes prioritarios","Único pendiente, imputado a AESA: LETTER OF CREDIT de la orden de compra ampliada, con plazo solicitado el 30-10-2026. Ver PA-14."),
 ("Lo que el informe NO reporta","La sección «Release and Dispatch» está en blanco: no hay plan de liberación, packing, modo de transporte ni agente de carga. La tabla de certificaciones tampoco tiene hitos ni montos cargados. Ver PA-15."),
 ("Oportunidad","Con la fabricación terminada unas 3 semanas antes de lo previsto, hay holgura entre el fin de fabricación y el despacho planificado (23-11 al 14-12-26). Adelantar el despacho ganaría semanas sobre la nacionalización, que es el tramo de mayor riesgo de esta provisión (PA-01) y el que condiciona la adecuación de HARG (PA-02 y PA-05)."),
]
SALAS_TXT=[
 ("Identificación","Informe de avance ABB Nº 6 (R-2631036 - Informe de Avance nro 6_A), revisión A del 04-11-2026, realizado por R. Mack y aprobado por M. Puga. Período informado: octubre 2026. Alcance: órdenes de compra 45089944971 (SE#3) y 45089944973 (SE#4)."),
 ("Hito del período","APROBACIÓN DE LA INGENIERÍA DE DETALLE DE TODOS LOS EQUIPOS de las salas eléctricas: tableros auxiliares, tableros CCM, tableros de MT, sistema UPS y baterías, y variadores de media y baja tensión, todos emitidos y aprobados en revisión numérica PARA FABRICACIÓN. Es el hito que ABB destaca del mes."),
 ("Lo que NO se aprobó","La ingeniería de las SALAS propiamente dichas (el shelter) sigue con APROBACIÓN PARCIAL, tanto la básica como la de detalle. ABB es explícito: «se observan nuevos comentarios en las sucesivas revisiones que impiden la aprobación final y emisión de ingeniería para fabricación», y lo incluye entre los factores potenciales de retraso pidiendo «evitar la generación de nuevos comentarios sobre documentos ya emitidos para fabricación». Ver PA-08."),
 ("Fechas contractuales vs objetivo","ABB RATIFICA la versión optimista del cronograma: SALA 3 fecha contractual 14-05-2027 contra objetivo 23-04-2027 (mejora de 21 días); SALA 4 fecha contractual 08-04-2027 contra objetivo 25-03-2027 (mejora de 14 días). La Rev12 tomaba 20-04-27 para SE#3 a partir del cronograma del 02-10; se adopta el 23-04-27 que es la fecha que el proveedor declara."),
 ("Ductos de barras","ABB reitera el corrimiento de la fecha contractual de los ductos de la sala #3: no pudieron diseñarse por falta de información sobre los gabinetes de los transformadores de MT y sus posiciones relativas, y el diseño arrancó recién el 21-09-26. ABB «se encuentra trabajando en la confirmación del plazo de entrega y el posible impacto económico» y vuelve a pedir que AESA adecúe la fecha contractual de las posiciones 120, 130 y 140 de la OC 45089944971 para evitar penalidades. Es el segundo mes consecutivo sin respuesta. Ver PA-07."),
 ("Situación contractual","Pólizas de fiel cumplimiento y fondo de reparo de ambas OC emitidas y aprobadas. Acordadas las aclaraciones sobre los hitos de certificación y revisados los documentos de compra. Nota de cambio Nº 1 (R-2631036-NC1_A) aprobada por AESA, con nota de pedido recibida en agosto; ABB avanzó con la implementación técnica. Se avanzó con la certificación parcial del hito 1 contra emisión de las ingenierías básicas de los variadores y de las salas."),
 ("Fabricación del período","Variadores de media tensión: fabricación y ensayos INICIADOS Y FINALIZADOS. Tableros CCM: inicio de fabricación de columnas y módulos extraíbles. Tableros de MT: inicio de fabricación de gabinetes de comando. Para noviembre ABB prevé finalizar la fabricación de los CCM y de los tableros de MT e iniciar la de los tableros auxiliares."),
 ("Acopio","Tableros auxiliares: acopio de equipamiento mecánico y eléctrico al 100%. Sistema UPS y baterías: breakers, fusibles, ventiladores, transformadores de intensidad, placas electrónicas y semiconductores al 100%; gabinetes y racks al 90%. Son mejoras sustanciales respecto del informe Nº5 (auxiliares 72%, UPS con semiconductores al 63%)."),
 ("Órdenes a subproveedores","El informe publica por primera vez las fechas de entrega en origen comprometidas: salas eléctricas 3 y 4 con Bottino (14-05-27 y 08-04-27) · tableros auxiliares con Mehcco (28-01-27) · sistema UPS y baterías con KMT (27-11-26) · tableros de MT con BRABB Brasil (27-11-26) · tableros CCM con BRABB Brasil (25-01-27) · variadores de BT con FIABB Finlandia y de MT con CNABB China (06-11-26 ambos) · DUCTOS DE BARRAS con EAE Turquía: «tbd», sin fecha comprometida."),
 ("Previsión de FAT","Sistema UPS y baterías 18 al 20-11-26 en Santa Fe · tableros TGMT 30-11 al 04-12-26 en Sorocaba · tableros CCM 30-11 al 12-12-26 en Sorocaba · tableros auxiliares 14 al 16-12-26 en AMBA · sala eléctrica #4 22 al 26-02-27 y sala eléctrica #3 29-03 al 02-04-27, ambas en Mendoza. Las tres primeras ya tienen invitación enviada. Las de Sorocaba arrancan el mismo día que la prueba de integración PCS↔PMS en Buenos Aires. Ver PA-19."),
 ("Calidad e inspección","Documentación de calidad emitida y aprobada: procedimientos de END, de soldadura (WPS/WPQ/PQR) y de pintura; protocolos FAT de las salas 3 y 4; planes de inspección y ensayos de la sala 3, de los tableros CCM y de los TGMT; procedimientos de FAT de CCM y TGMT; protocolos FAT de los tres CCM. Los welding maps de ambos shelters quedaron en «aceptado con comentarios». Para noviembre: inspección al proveedor de salas eléctricas para verificar avance de fabricación y al de tableros auxiliares para verificar acopio."),
 ("Avance por paquete","Sala eléctrica 3 al 34% (previsto 37% para noviembre) es el paquete MÁS ATRASADO de la provisión, con recepción de materiales en 18% y fabricación en 0%. Sala eléctrica 4 al 49% (previsto 55%). Los equipos van muy por delante: variadores de BT 71% (98% en noviembre), variadores de MT 71% (92%), tableros auxiliares 75% (91%), CCM 69% (85%), UPS 54% (90%), tableros de MT 45% (62%). Ductos de barras al 15% (30% en noviembre)."),
 ("Temas pendientes prioritarios","Tres, todos con componente de AESA: (1) control de acceso — suministro de los equipos físicos según el cronograma de cada sala, responsabilidad de AESA (PA-13); (2) definición de las conexiones del ducto de barras con los transformadores de media tensión; (3) aprobación de la ingeniería constructiva de las salas (PA-08)."),
 ("Certificación y cobranzas","OC 45089944971: certificaciones 1 y 2 (USD 143.589 c/u) y 3 (USD 29.343) PAGADAS; certificación 4 (USD 376.804, vto. 28-10-26) y certificación 5 (USD 1.191.654,46, vto. 28-11-26) pendientes de cobro. Total USD 1.884.979,46. OC 45089944973: certificaciones 1 y 2 (USD 170.739 c/u) y 3 (USD 29.343) PAGADAS; certificación 4 (USD 48.031) y certificación 5 (USD 566.080,57, vto. 28-11-26) pendientes. Total USD 984.932,57. La facturación de octubre multiplica por cinco la acumulada: el grueso del compromiso económico entra ahora."),
]
AVANCE_PMS=[  # (especialidad, peso, sep-26 real, oct-26 previsto, oct-26 real, nov-26 previsto, nov-26 pronóstico)
 ("Ingeniería",            "20%","90%","95%","95%","95%","95%"),
 ("Colocación de subórdenes","5%","100%","100%","100%","100%","100%"),
 ("Recepción de materiales","30%","100%","100%","100%","100%","100%"),
 ("Fabricación",           "20%","0%","0%","0%","25%","25%"),
 ("Configuración",         "10%","0%","0%","0%","0%","25%"),
 ("Tests / Inspección",    "10%","0%","0%","0%","0%","0%"),
 ("Despacho",               "5%","0%","0%","0%","0%","0%"),
 ("TOTAL GLOBAL",         "100%","53%","54%","54%","59%","62%"),
]
# avance declarado por cada proveedor en su propio cronograma
AVANCE_PROV=[  # (provisión, fuente y fecha, avance declarado, observación)
 ("ABB — Sala Eléctrica SE#3 (BT)","Informe de avance Nº6 · 04-11-2026","34%","Shelter. Ingeniería 80% real a sep-26, recepción de materiales 18%, fabricación 0%. Previsto nov-26: 37%. ES EL PAQUETE DE MENOR AVANCE DE LA PROVISIÓN y en la Rev12 figuraba en 36%."),
 ("ABB — Sala Eléctrica SE#4 (MT)","Informe de avance Nº6 · 04-11-2026","49%","Shelter. Recepción de materiales 18% real y 50% previsto a oct-26; fabricación 30% prevista. Previsto nov-26: 55%. Mejora 13 puntos respecto de la Rev12 (36%)."),
 ("ABB — Ductos de barras","Informe de avance Nº6 · 04-11-2026","15%","Ingeniería 15% real a sep-26 y 50% prevista a oct-26; fabricación y transporte marítimo en 0%. Previsto nov-26: 30%. Plazo de entrega SIN CONFIRMAR — ver PA-07."),
 ("ABB — PMS (Power Management System)","Informe de avance Nº5 · 30-10-2026","54%","Avance físico ponderado, exactamente igual al previsto. Ingeniería 95%, recepción de materiales 100% y verificada con inspección. FABRICACIÓN 0%: no arrancó en octubre — ver PA-18."),
 ("CAT Miron — Transformadores","Acta de inspección Nº1 · 21-10-2026","parcial","Ingeniería 100% (OT 5211 al 75%); conductores 40-90%; materiales gruesos 30-50%; fabricación en curso. Sin informe de avance posterior. La primera entrega está comprometida para el 13-11-2026: confirmar."),
 ("Inauco — PCS / SCADA / Comunicaciones","Cronograma PG-001 REV 0 · 23-10-2026","6%","Avance global declarado. Configuración base 11% (PLC/comunicaciones 12%, SCADA 12%). La construcción de tableros arranca el 27-11-26 y concentra el grueso de la curva. Sin novedades en el período."),
 ("HPH — HIMA Paul Hildebrandt (Alemania)","Supplier Progress Report Nº01 · 15-10-2026","fabricación 100%","Sin informe nuevo. Fabricación de los 3 tableros completada; ingeniería 80%. La inspección prevista para la semana 44 o 45 (26-10 al 08-11-26) no tiene resultado reportado y la carta de crédito venció el 30-10 — ver PA-14."),
 ("HARG — HIMA Argentina","PSCH_CPF2_Rev3 · 04-09-2026","0%","OC no colocada — los trabajos no iniciaron. Segunda postergación: AESA la estima ahora para la semana del 16-11-2026."),
]
# avance real por hito al 31-10-2026 (informes ABB Nº5 PMS y Nº6 Salas)
HPHK="HPH — HIMA Paul Hildebrandt (Alemania) · tableros de marshalling"
INAK="Inauco — PCS / SCADA / Comunicaciones"
CMK="CAT Miron — Transformadores (13 unidades)"
SE3K="ABB — Sala Eléctrica SE#3 (BT)"; SE4K="ABB — Sala Eléctrica SE#4 (MT)"; DUCK="ABB — Ductos de barras"
AVANCE_REAL={
 # --- CAT Miron: acta de inspección IN-FC-CAT-19082-261021 (21-10-2026) ---
 (CMK,"Ingeniería — cálculo y diseño (5 órdenes de trabajo)"):("100%","Cumplido en las 5 OT"),
 (CMK,"Ingeniería — proyecto de ingeniería en detalle"):("95%","OT 5211 al 75%"),
 (CMK,"ACOPIO — Conductores BT/MT (foils de aluminio)"):("90%","Verificado en planta con certificados"),
 (CMK,"ACOPIO — Materiales gruesos (gabinetes y resina bicomponente)"):("50%","Resina aprobada en origen"),
 (CMK,"Emisión del Plan de Inspección y Ensayos (PIE) Rev.A"):("100%","Emitido 01-09-26"),
 (CMK,"Fabricación de núcleos — TERCERIZADA (Villa Martelli)"):("0%","Subproveedor sin inspeccionar"),
 (CMK,"Fabricación — bobinado BT/MT"):("0%","Bobinas de distribución en stock"),
 (CMK,"INSPECCIÓN Nº 1 EN PLANTA CAT"):("100%","Sin no conformidades"),
 (CMK,"ENSAYOS CON PUNTO DE RETENCIÓN — A COORDINAR"):("0%","PIE emitido — falta la fecha"),
 # --- Inauco: cronograma PG-001 Rev 0 (23-10-2026), avance global 6% ---
 (INAK,"Inicio del proyecto — Recepción de OC de MO + MATERIALES (AESA)"):("100%","Cumplido"),
 (INAK,"Recepción de OC de SEGUROS (AESA)"):("0%","VENCIDO — 3 meses de atraso"),
 (INAK,"Recepción de OC de MANO DE OBRA DE CAMPO (AESA)"):("0%","VENCIDO — 3 meses de atraso"),
 (INAK,"REQ 1 — Recepción de ingeniería de AESA (5 entregas escalonadas)"):("parcial","2 entregas vencidas al 66% y 50%"),
 (INAK,"Emisión de ingeniería Rev.A — documentos de tableros"):("77%","Última entrega incompleta"),
 (INAK,"Emisión de ingeniería Rev.0 — documentos de tableros"):("67%","3 entregas vencidas"),
 (INAK,"Configuración base — PLC y comunicaciones"):("12%","En curso"),
 (INAK,"Configuración base — SCADA"):("12%","En curso"),
 # --- HPH: Supplier Progress Report Nº01 (15-10-2026) ---
 (HPHK,"Ingeniería de detalle / especificación"):("80%","En curso — sin documentos vencidos"),
 (HPHK,"ACOPIO / disposición de materiales"):("100%","Material de la orden principal ENTREGADO"),
 (HPHK,"FABRICACIÓN DE LOS TABLEROS DE MARSHALLING"):("100%","COMPLETADA — los 3 tableros"),
 (HPHK,"Emisión de documentación de ingeniería (diagramas de conexionado)"):("—","A CONFIRMAR — previsto 21-10-26"),
 (HPHK,"Emisión de documentación de ingeniería (cálculos térmicos y procedimientos FAT/SAT)"):("0%","Previsto 30-10-26"),
 (HPHK,"LETTER OF CREDIT de la orden ampliada (AESA)"):("0%","PENDIENTE — vence 30-10-26"),
 (HPHK,"Inspección / HW check de los 3 tableros"):("0%","Semana 44 o 45 — remote HW check a definir"),
 (HPHK,"Entrega DAP Ezeiza AEROPUERTO (Argentina)"):("0%","Confirmada sin desvío"),
 # --- ABB Salas: informe Nº5 (05-10-2026) ---
 (SE3K,"Emisión de ingeniería básica"):("100%","Cumplido"),
 (SE3K,"Aprobación de ingeniería básica"):("parcial","APROBACIÓN PARCIAL — comentarios abiertos"),
 (SE3K,"FREEZING POINT — ingeniería básica"):("0%","VENCIDO hace 6 semanas"),
 (SE3K,"Emisión de ingeniería de detalle"):("95%","En curso"),
 (SE3K,"Aprobación de ingeniería de detalle — EQUIPOS"):("100%","APROBADA para fabricación — hito del período"),
 (SE3K,"Aprobación de ingeniería de detalle — SALA"):("parcial","VENCIDA el 13-10-26 — aprobación parcial"),
 (SE3K,"ACOPIO — Materiales de los CCM (críticos y secundarios)"):("85%","Recepción de materiales — 100% en nov"),
 (SE3K,"ACOPIO — Rectificadores / UPS / baterías"):("98%","Sólo gabinetes y racks al 90%"),
 (SE3K,"ACOPIO — Tableros auxiliares"):("100%","COMPLETO — certificación hito 4 en nov"),
 (SE3K,"Inicio de fabricación en fábrica de origen"):("100%","CCM: columnas y módulos extraíbles"),
 (SE3K,"Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)"):("0%","Pendiente — shelter al 34%, el más bajo"),
 (SE4K,"Emisión de ingeniería básica"):("100%","Cumplido"),
 (SE4K,"Aprobación de ingeniería básica"):("parcial","APROBACIÓN PARCIAL — comentarios abiertos"),
 (SE4K,"FREEZING POINT — ingeniería básica"):("0%","VENCIDO hace 6 semanas"),
 (SE4K,"Emisión de ingeniería de detalle"):("95%","En curso"),
 (SE4K,"Aprobación de ingeniería de detalle — EQUIPOS"):("100%","APROBADA para fabricación — hito del período"),
 (SE4K,"Aprobación de ingeniería de detalle — SALA"):("parcial","VENCIDA el 13-10-26 — aprobación parcial"),
 (SE4K,"ACOPIO — Celdas MT (materiales principales y secundarios)"):("80%","Previsto a oct-26 — 100% en nov"),
 (SE4K,"ACOPIO — Variadores MT (China) y BT (Finlandia)"):("100%","Fabricación y ensayos finalizados"),
 (SE4K,"ACOPIO — Tableros auxiliares"):("100%","COMPLETO — certificación hito 4 en nov"),
 (SE4K,"Inicio de fabricación en fábrica de origen"):("100%","Celdas MT: gabinetes de comando"),
 (SE4K,"Fabricación del shelter (base, estructura, revestimiento, pintura, inst. eléctrica)"):("30%","Fabricación prevista al 30% en oct-26"),
 (DUCK,"Emisión layout instalación + detalle conexionado trafo"):("100%","Cumplido"),
 (DUCK,"Diseño de ductos de barras"):("15%","50% previsto a oct-26, 100% a nov-26"),
 # --- ABB PMS: informe Nº4 (30-09-2026) ---
 (PMSK,"Inicio de contrato — Recepción / aceptación OC"):("100%","Cumplido"),
 (PMSK,"Kick-off Meeting (KOM) y planificación detallada"):("100%","Cumplido"),
 (PMSK,"Emisión de ingeniería básica (documentación para aprobación)"):("100%","Cumplido · certificado"),
 (PMSK,"Aprobación de ingeniería básica (AESA)"):("100%","Cumplido · certificado"),
 (PMSK,"Emisión de ingeniería Rev.0"):("100%","Cumplido"),
 (PMSK,"Aprobación ingeniería Rev.0 — FREEZING POINT"):("100%","Cumplido"),
 (PMSK,"ACOPIO — Inicio de compra de materiales"):("100%","Cumplido"),
 (PMSK,"ACOPIO — Material S800 y tableros"):("100%","Verificado en stock el 22-10-26"),
 (PMSK,"ACOPIO — Switches de red"):("100%","Verificado en stock el 22-10-26"),
 (PMSK,"Configuración / programación del sistema"):("0%","Especialidad «Configuración» en 0% — PA-11"),
 (PMSK,"INSPECCIÓN DE ACOPIO EN STOCK ABB/DHL (witness AESA)"):("100%","100% de los materiales verificados"),
 (PMSK,"ACOPIO DE MATERIALES COMPLETO"):("100%","Cumplido y certificado en oct-26"),
 (PMSK,"INICIO DE ARMADO DE TABLEROS"):("0%","NO ARRANCÓ en octubre — PA-18"),
 (PMSK,"Pruebas de Maqueta Integrada PMS/CCM/PCS — CPF1"):("—","Sin registro de ejecución en el informe Nº5"),
 (PMSK,"Emisión de procedimientos de FAT y SAT (Rev.A)"):("0%","Corrida a noviembre — PA-21"),
 (PMSK,"Configuración de variables de comunicación y lógicas del sistema"):("0%","Debía arrancar el 02-11-26 — PA-11"),
}
AVANCE_TXT=[
 ("Identificación (ABB PMS)","Informe de avance ABB Nº 5 (E-2611027 — Informe de Avance CPF2 LA CALERA, OCT 2026), revisión A del 30-10-2026, realizado por Pablo Kalis y aprobado por V. Toledo. Período informado: octubre 2026. Orden de compra 4508945953."),
 ("Avance del período (ABB PMS)","Avance físico real de octubre 54% contra 54% previsto: EXACTAMENTE EN PLAN, sin el colchón de septiembre (53% contra 44%). Ingeniería 95%, colocación de subórdenes 100%, recepción de materiales 100%. Fabricación, configuración, tests y despacho siguen en 0%. Pronóstico de noviembre: 62%."),
 ("La fabricación no arrancó","El informe declara «sin actividades en el periodo» en el área de Fabricación y lleva el inicio de fabricación a noviembre, en las instalaciones de MEHCCO S.A. (Pasaje Ricardo Güiraldes 4567, San Martín, Buenos Aires). La Rev12 tomaba el 09-10-2026 como inicio del armado de tableros con 50 días hábiles de duración: el corrimiento acumulado llega a 7,5 semanas sobre la Rev6 y la fecha de fin del 17-12-2026 ya no cierra aritméticamente. Ver PA-18."),
 ("Acopio verificado con inspección","El 22-10-2026 se realizó la inspección de los materiales en el stock de ABB/DHL con presencia del inspector de AESA. Se corroboró la existencia del 100% de los materiales necesarios para el proyecto (acta IN-SAV-ABB-45953-261022). Es el único hito de Calidad del período y habilitó la certificación C-2 de acopio por USD 220.055,72 en octubre."),
 ("Acuerdos del período","(1) ABB y AESA acordaron realizar la FAT INTEGRAL de las salas EN PLANTA SHELTERISTA, con una duración de 5 días y verificando una muestra representativa de señales; el protocolo de ensayos lo definirá y emitirá ABB «en los próximos meses», sin fecha comprometida. Ver PA-12. (2) Se acordaron pruebas anticipadas con la maqueta de CPF1 y los materiales y cubicles de UMC propios de CPF2, en conjunto ABB / AESA / PLUSPETROL, con participación de INAUCO por la comunicación con el PCS: semana del 30 de noviembre de 2026, duración aproximada de 2 días. Ver PA-03 y PA-19. (3) Definiciones sobre los vínculos de fibra óptica para vincular las nuevas salas SE#3 y SE#4 con las salas existentes en sitio."),
 ("Ingeniería emitida en octubre","Descripción funcional Rev.1 · topología del sistema Rev.0 · listado de señales y comunicaciones Rev.0 · cálculo de consumo de potencia Rev.0 · cálculo de carga de controladores Rev.0 · layout, diagrama de conexionado y listado de partes de los tableros de control / IO de SE#3 y SE#4 Rev.0 · especificaciones técnicas de equipos ABB y de terceros Rev.1 · planes de inspección y ensayos 102-PMS-001 y 102-PMS-101 Rev.0."),
 ("Eventos previstos para noviembre","Desarrollo de la ingeniería básica restante: protocolo de aceptación en fábrica (FAT) y su registro de pruebas, y protocolo de aceptación en sitio (SAT) y su registro — los mismos que la Rev12 ubicaba en octubre. Ver PA-21. Inicio de fabricación en MEHCCO. Sin actividades previstas de Calidad / Inspección. Reuniones semanales de seguimiento técnico AESA / ABB / PPSA."),
 ("Temas pendientes prioritarios (insumos de AESA)","Se repiten sin cambios respecto del informe Nº4: (1) definiciones de señales cableadas en cada equipo, con indicación de las borneras de vinculación al PMS (transformadores, generadores, seccionador bajo carga); (2) mapa Modbus y definición de señales de los equipos a integrar (UPS, generadores, multimedidores); (3) mapas Profinet y definición de señales de UMC y VFD. Segundo mes consecutivo pendientes, con la especialidad Configuración en 0%. Ver PA-11."),
 ("Alcance de los tableros","Tablero PMS de control de la SALA 3: tres cuerpos de 2200×800×600 mm. Tablero PMS de control de la SALA 4: un cuerpo de 2200×800×600 mm. Ambos alojan controladores redundantes, módulos de E/S, placas de comunicaciones IEC 61850, Profinet y Modbus, y switches IEC 61850, Profinet y Modbus TCP."),
 ("Fuera del alcance","Módulo adicional de entradas (oferta OPP-26-8808080 Rev.0, típico DOL_TIPICO-01A): AESA aceptó el 23-07-26 avanzar con la gestión interna de ABB para no impactar las pruebas de integración de nov/dic, pero ABB aclara que la gestión y la ejecución de esa PO no forman parte del presente alcance. Verificar que esté disponible para la maqueta CPF2. Ver PA-20."),
 ("Situación contractual y certificación","Sin avances en la certificación ni cambios contractuales en el período. Certificación nº 1 presentada el 27-08-2026 por USD 176.044,58 + impuestos y certificación nº 2 el 27-10-2026 por USD 220.055,72 + impuestos. Plan de certificación acordado sobre USD 907.315,10: emisión de ingeniería básica 9,70% y aprobación 9,70% (ago-26, realizados) · acopio de materiales principales en Argentina 24,25% (oct-26, realizado) · contra finalización de la FAT 19,40% (feb-27) · contra entrega de equipos 29,10% y DataBook 4,85% (mar-27) · contra entrega de repuestos 2,99% (may-27). El esquema comercial ubica la entrega de equipos en marzo, un mes después del cronograma técnico."),
 ("Vigencia de estos datos","Se incorporan los dos informes de octubre de ABB: PMS Nº5 (30-10-2026) y Salas Eléctricas Nº6 (04-11-2026). Ninguno de los dos reemite el cronograma: las fechas base siguen siendo las del REV 0 «Para construcción» del 14-10-26 y del SEP-2026 REV 0 del 01-10-26, con los ajustes que los propios informes declaran. No se recibieron informes nuevos de HPH, Inauco, CAT Miron ni HARG."),
 ("Provisiones sin informe de avance","Presentan informe mensual de avance ABB —por separado para Salas Eléctricas (Nº 6) y para el PMS (Nº 5)— y HPH / HIMA Alemania (Nº 1, con la plantilla AC-PR-2026-002 de AESA, sin continuidad desde octubre). Faltan INAUCO, CAT MIRON y HARG, para los que se toma el porcentaje declarado en el propio cronograma del proveedor, con criterios de medición distintos y no comparables entre sí. Conviene pedirles informe de avance mensual con el mismo formato ponderado que usa ABB."),
]
# ---------------- Comparativo entre provisiones (Rev11) ----------------
# (provisión, última revisión del cronograma, estado del documento, entrega, cambio vs. rev. anterior,
#  avance declarado, informe de avance, insumo de AESA pendiente)
COMPARATIVO=[
 ("ABB — Sala Eléctrica SE#3 (BT)","REV 0 · 14-10-26","Para construcción","23/04/2027","Fecha objetivo RATIFICADA (23-04-27, 21 d mejor que la contractual)","34% shelter · equipos 69-98%","Nº 6 · 04-11-26","Aprobación de ing. de la SALA (parcial, vencida 25-09) · control de acceso 13-02-27"),
 ("ABB — Sala Eléctrica SE#4 (MT)","REV 0 · 14-10-26","Para construcción","25/03/2027","Fecha objetivo RATIFICADA (25-03-27, 14 d mejor que la contractual)","49% shelter · equipos 45-92%","Nº 6 · 04-11-26","Aprobación de ing. de la SALA (parcial, vencida 25-09) · control de acceso 04-01-27"),
 ("ABB — Ductos de barras","REV 0 · 14-10-26","Para construcción","08/07/2027","Plazo SIN CONFIRMAR; ABB reitera el pedido de adecuación contractual","15% (50% previsto a oct-26)","Nº 6 · 04-11-26","Responder el pedido de adecuación contractual (pos. 120/130/140) — sin respuesta desde el 05-10"),
 ("CAT Miron — Transformadores","Act. 29-09-26","Para información","11/12/2026","Sin cambios · primera entrega el 13-11-26","Ingeniería 100% · fabricación en curso","Acta de insp. Nº1 · 21-10-26","Convocar los puntos de retención (10 d de aviso) · inspeccionar el taller de núcleos"),
 ("ABB — PMS","SEP-2026 REV 0 · 01-10-26","Emitido","17/02/2027","Armado NO iniciado en octubre: 4 semanas más de corrimiento","54% (vs 54% previsto)","Nº 5 · 30-10-26","Señales cableadas, mapa Modbus y mapas Profinet (PA-11) — 2.º mes pendientes"),
 ("Inauco — PCS / SCADA / Comunicaciones","REV 0 · 23-10-26","Para construcción","04/05/2027","SIN CORRIMIENTOS · prueba de integración confirmada 30-11/01-12","6% global · configuración 11%","— sin informe","OC de seguros y de MO de campo (vencidas) · 2 entregas de ingeniería incompletas"),
 ("HPH — HIMA Paul Hildebrandt (Alemania)","Schedule 16-08-26","Vigente","25/01/2027","Sin novedades desde el informe Nº1 (15-10-26)","Fabricación 100% · ingeniería 80%","Nº 1 · 15-10-26","Letter of Credit de la orden ampliada — VENCIDA el 30-10-26"),
 ("HARG — HIMA Argentina","PSCH_CPF2_Rev3 · 04-09-26","Previo a la OC","24/05/2027","CORRIDA 2 SEMANAS MÁS: +63 días; OC estimada el 16-11-26","0% — no iniciado","— sin informe","COLOCAR LA ORDEN DE COMPRA (postergada por segunda vez)"),
]
COMPARATIVO_OBS=[
 ("Qué cambió en esta revisión","Dos informes de ABB de octubre y una segunda postergación de la OC de HIMA Argentina. Los informes traen buenas noticias en las salas (toda la ingeniería de equipos aprobada para fabricación, acopio de auxiliares al 100%, fabricación en marcha en Brasil) y una mala en el PMS (el armado de tableros no arrancó). La postergación de la OC de HARG mueve, por primera vez en tres revisiones, la fecha final del paquete: el camino crítico pasa del 10-05-2027 al 24-05-2027."),
 ("Quién cumple y quién no","De las ocho provisiones, CINCO sostienen su fecha de entrega sin corrimientos: las dos salas —que incluso mejoran la contractual en 21 y 14 días—, los transformadores, Inauco y HPH. Las tres excepciones son los DUCTOS DE BARRAS, cuyo plazo ABB sigue declarando sin confirmar; HARG, que acumula nueve semanas de corrimiento por una orden de compra que no se coloca; y ahora el PMS, cuyo armado de tableros no arrancó en octubre y pone en duda la cadena FAT-entrega."),
 ("El patrón que se repite","En SIETE de las ocho provisiones el pendiente crítico sigue siendo un INSUMO DE AESA, no un problema del proveedor: aprobación de la ingeniería de las salas, respuesta al pedido contractual de los ductos, mapas de señales para el PMS, órdenes de compra e ingeniería para Inauco, carta de crédito para HPH y la orden de compra de HARG. En esta revisión el patrón se agrava: la carta de crédito de HPH venció el 30-10 sin emitirse, los mapas del PMS llevan dos meses pendientes y el pedido de adecuación de los ductos lleva un mes sin respuesta."),
 ("El ciclo de revisión se volvió el cuello de botella","ABB lo dice con todas las letras en el informe Nº6: pide «evitar la generación de nuevos comentarios sobre documentos ya emitidos para fabricación». La ingeniería de los equipos se cerró; la del shelter no, porque cada revisión abre observaciones nuevas. No es un problema de plazo de emisión del proveedor sino del circuito de aprobación del EPC. Ver PA-08."),
 ("Madurez documental","Tres cronogramas están en estado PARA CONSTRUCCIÓN: ABB Salas y Ductos (14-10) e Inauco (23-10). Ninguno de los dos informes de octubre reemitió cronograma, de modo que las fechas base no cambiaron formalmente aunque los informes declaren desvíos: el del armado del PMS no está reflejado en ningún cronograma vigente. CAT Miron sigue en «para información» y HARG ni siquiera tiene contrato."),
 ("Dispersión del avance","El rango va del 0% de HARG al 100% de fabricación de HPH. El punto bajo del paquete es ahora doble: los ductos de barras al 15%, con el plazo sin confirmar, y la sala eléctrica 3 al 34%, con la recepción de materiales en 18% y la fabricación del shelter arrancando el 23-11. El PMS, con 54%, está exactamente en plan en el papel, pero con el 20% de peso de Fabricación todavía en cero."),
 ("Convergencia de noviembre y diciembre de 2026","Cuatro eventos con witness se apilan en cuatro semanas: FAT del sistema UPS 18-20/11 en Santa Fe, FAT de TGMT 30/11-04/12 y de CCM 30/11-12/12 en Sorocaba, prueba de integración PCS↔PMS 30/11-01/12 en Buenos Aires y FAT de tableros auxiliares 14-16/12 en AMBA. El 30 de noviembre hay tres eventos simultáneos en dos países. Ver PA-19."),
 ("Convergencia de febrero a mayo de 2027","Cinco provisiones concentran sus FAT y entregas entre febrero y mayo de 2027: PMS (entrega 17/02), SE#4 (FAT 22-26/02, entrega 25/03), SE#3 (FAT 29/03-02/04, entrega 23/04), Inauco (FAT 29/03-26/04, entrega 04/05) y HARG (FAT 28-30/04, IFAT 03-14/05, cierre 24/05). La ventana se estiró a catorce semanas con witness en Mendoza, Neuquén y Buenos Aires."),
 ("Los dos extremos del paquete","HPH ya terminó de fabricar y entrega el 15-12-26 en Ezeiza; el ducto de barras entrega el 08-07-27, casi siete meses después. Entre ambos hay siete meses de almacenamiento de transformadores (PA-10) y un camino crítico de sistemas que ahora cierra el 24-05-27."),
]

NOTAS=[
 "Rev13 — cambio 1: HARG (HIMA ARGENTINA) — AESA informa que la orden de compra no sale esta semana sino la siguiente. Se adopta el 16-11-2026 como fecha de colocación (la Rev12 asumía el 01-11-26) y el corrimiento de las actividades propias de HARG pasa de 49 a 63 días corridos. Consecuencia: el CIERRE DEL CAMINO CRÍTICO se mueve del 10-05-2027 al 24-05-2027. Es el primer corrimiento del proyecto que se traduce en una fecha final distinta.",
 "Rev13 — cambio 2: ABB PMS — se incorpora el informe de avance Nº 5 (E-2611027, 30-10-2026). Avance 54% contra 54% previsto. El ARMADO DE TABLEROS NO ARRANCÓ EN OCTUBRE («sin actividades en el periodo») y se lleva a noviembre, en MEHCCO S.A. (San Martín); el acopio quedó verificado al 100% con inspección de AESA el 22-10-26 (IN-SAV-ABB-45953-261022) y certificado en octubre. Se suman PA-18 (armado corrido con fecha de fin sostenida), PA-20 (módulo adicional fuera de alcance) y PA-21 (procedimientos de FAT y SAT corridos a noviembre).",
 "Rev13 — cambio 3: ABB SALAS — se incorpora el informe de avance Nº 6 (R-2631036, 04-11-2026). Hito del período: APROBADA LA INGENIERÍA DE DETALLE DE TODOS LOS EQUIPOS para fabricación. La ingeniería del SHELTER sigue con aprobación parcial y ABB pide no abrir comentarios nuevos sobre documentos ya emitidos (PA-08). Acopio de tableros auxiliares al 100% y de UPS al 98%. Se ratifican las fechas objetivo 23-04-27 (SE#3, antes 20-04 en el cronograma) y 25-03-27 (SE#4).",
 "Rev13 — cambio 4: se actualiza la previsión de FAT con las fechas del informe Nº6 (UPS 18-20/11 Santa Fe · TGMT 30/11-04/12 y CCM 30/11-12/12 Sorocaba, con invitación enviada · auxiliares 14-16/12 AMBA) y se confirma la prueba de integración PCS↔PMS para la semana del 30-11-26, 2 días. Se suma PA-19 por la superposición de tres eventos con witness el 30 de noviembre en dos países.",
 "Rev13 — cambio 5: DUCTOS DE BARRAS — sin cambio de fecha pero con el pedido de adecuación contractual de las posiciones 120/130/140 reiterado por segundo mes sin respuesta de AESA. El proveedor EAE (Turquía) sigue sin fecha de entrega en origen comprometida («tbd»).",
 "ALERTA — PA-04: la OC de HIMA Argentina se postergó por segunda vez y ya mueve la fecha final del paquete. PA-18: el armado de los tableros del PMS no arrancó y ABB sostiene el 17-12-26 sin sustento aritmético. PA-14: la carta de crédito de HPH venció el 30-10-26 sin emitirse. Los tres son decisiones de AESA o reclamos a resolver con el proveedor, no imprevistos técnicos.",
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
 "Fuentes Rev13: ABB PMS informe de avance Nº5 E-2611027 (30-10-26) · ABB Salas informe de avance Nº6 R-2631036_A (04-11-26). Fuentes anteriores vigentes: CAT Miron acta de inspección IN-FC-CAT-19082-261021 (21-10-26) y PIE 5155-00-0022-VE-CA-PE-001 Rev.A · ABB PMS 5155-00-0007-VE-CO-PG-001 SEP-2026 REV 0 (01-10-26) + informe de avance Nº4 E-2611027 (30-09-26) · CAT Miron 5155-00-0022-VE-CO-PG-001 Rev.A (act. 29-09-26) · ABB Salas y Ductos 5155-00-0009-VE-CO-PG-001 REV 0 «Para construcción» (14-10-26) · ABB PMS 5155-00-0007-VE-CO-PG-001_REV B (.mpp) · Inauco 5155-00-0029-VE-CO-PG-001-A Rev.A (24-08-26) y 5155-00-0029-VE-CA-PG-002 Rev.0 (14-09-26) · HPH 103618-Schedule 16-08-26 · HARG PSCH_CPF2_Rev3 (04-09-26) con corrimiento +63 d.",
 "Los hitos de FAT en fábrica de origen (Brasil) son witness de AESA y anteceden a la FAT del shelter completo; no deben confundirse entre sí.",
 "Estado (plan): se calcula contra la fecha de corte. 'Cumplido s/plan' = la fecha planificada ya venció y debe confirmarse el cumplimiento real con el proveedor; 'En curso' = en ejecución; 'Pendiente' = aún no inició.",
 "Las columnas Fecha real / % avance / Estado real / Observaciones quedan en blanco para la carga de seguimiento periódico (reporte a PPSA).",
 "Los hitos se ordenan cronológicamente dentro de cada provisión y se renumeran (H01, H02, ...). Marcas: ◆ = hito del camino crítico · ⚠ = punto de atención · fondo naranja = acopio de materiales.",
]
INTERFASES=[
 "CAMINO CRÍTICO DEL PROYECTO: OC de HARG → HPH (Alemania) → nacionalización → HARG (adecuación, programación, FAT e IFAT) → Inauco (FAT SIS y liberación). Cierra el 24-05-27 con la entrega del alcance HARG, catorce días más tarde que en la Rev12. Ver hoja 'Camino crítico'.",
 "AESA → HARG: la COLOCACIÓN DE LA OC sigue PENDIENTE al 07-11-26 y se posterga a la semana del 16-11-26. Es el ítem que gobierna la fecha final del sistema de seguridad y ya movió el cierre del paquete.  ⚠ PA-04",
 "HPH → entrega DAP Ezeiza AEROPUERTO (15/12/26, confirmada sin desvío) + nacionalización → recepción en base Inauco (25/01/27). Con el nuevo corrimiento de HARG, la adecuación arranca recién el 22/02/27: cuatro semanas de espera del tablero ya nacionalizado.  ⚠ PA-01 / PA-02",
 "AESA → HPH: Letter of Credit de la orden de compra ampliada. El plazo pedido por HPH era el 30/10/26 y VENCIÓ sin confirmación de emisión.  ⚠ PA-14",
 "HARG ↔ Inauco → adecuación de tableros de marshalling (22/02 → 02/04/27) en base Inauco, superpuesta con la Pre iFAT de Inauco (10/02 → 26/03/27).  ⚠ PA-05",
 "HARG ↔ Inauco → la FAT de HARG (28-30/04/27) y su IFAT (03-14/05/27) quedan ENTERAMENTE DESPUÉS del cierre de la FAT SIS conjunta (26/04/27), y el alcance HARG cierra el 24/05/27, veinte días después de la liberación de tableros de Inauco (04/05/27).  ⚠ PA-05",
 "Inauco → FAT con el cliente 29/03 → 22/04/27 y LIBERACIÓN DE TABLEROS el 04/05/27 en camión de AESA (crono. PG-002 Rev.0): habilita el inicio de los trabajos de campo.",
 "ABB PMS ↔ Inauco → prueba de integración PCS↔PMS (maqueta CPF1 + cubicles de UMC de CPF2) en Bs.As. la semana del 30/11/26, 2 días, con ABB, AESA, PPSA e Inauco. CONFIRMADA en el informe Nº5.  ⚠ PA-03 / PA-19 / PA-20",
 "AESA Calidad → ABB Salas: FAT de UPS 18-20/11 en Santa Fe, de TGMT 30/11-04/12 y de CCM 30/11-12/12 en Sorocaba (invitaciones enviadas) y de auxiliares 14-16/12 en AMBA. El 30/11 coinciden con la prueba de integración en Buenos Aires.  ⚠ PA-19",
 "AESA Ingeniería → ABB PMS: definiciones de señales cableadas con borneras de vinculación, mapa Modbus y mapas Profinet. Segundo mes consecutivo pendientes; la configuración de comunicaciones debía arrancar el 02/11/26 y sigue en 0%.  ⚠ PA-11",
 "ABB PMS ↔ ABB Salas → FAT INTEGRAL del shelter EN PLANTA SHELTERISTA (5 días acordados, protocolo sin fecha de emisión) frente a las dos ventanas de ensayos PMS del cronograma de salas (SE#4 03-09/03 y SE#3 30/03-06/04/27).  ⚠ PA-12",
 "ABB PMS → entrega de equipamiento en shelter (17/02/27): habilita el montaje del PMS en las salas y los ensayos PMS. El armado de tableros no arrancó en octubre y la fecha está en riesgo.  ⚠ PA-18",
 "AESA Suministros → ABB PMS: módulo adicional de entradas (oferta OPP-26-8808080 Rev.0), fuera del alcance de la OC y necesario para la prueba de integración de noviembre.  ⚠ PA-20",
 "AESA → ABB Ductos: respuesta al pedido de adecuación de la fecha contractual de las posiciones 120/130/140 de la OC 45089944971, reiterado en el informe Nº6 y sin respuesta desde el 05/10/26. La entrega en sitio (08/07/27) sigue siendo la más tardía del paquete.  ⚠ PA-07",
 "AESA → aprobación de ingeniería de las SALAS: freezing points vencidos el 25/09/26 (básica) y el 13/10/26 (detalle). ABB pide no abrir comentarios nuevos sobre documentos ya emitidos para fabricación.  ⚠ PA-08",
 "AESA → ABB Salas: suministro de los equipos físicos de CONTROL DE ACCESO, 04/01/27 para SE#4 y 13/02/27 para SE#3. Reclamado por segundo informe consecutivo.  ⚠ PA-13",
 "CAT MIRON → ABB Ductos de barras: la ingeniería de detalle de los transformadores (cerrada el 04/09/26) es insumo del detalle de conexionado a trafo que AESA emitió el 18/09/26 y que gobierna el diseño de los ductos.",
 "CAT MIRON → AESA Construcciones: 13 transformadores entregados entre el 13/11 y el 11/12/26, siete meses antes de los ductos de barras (08/07/27). Requiere plan de almacenamiento y conservación.  ⚠ PA-10",
 "CAT MIRON → AESA Calidad: 14 puntos de retención del PIE Rev.A con 10 días de aviso previo, todavía sin fecha en el cronograma. La primera entrega está comprometida para el 13/11/26.  ⚠ PA-09",
 "CAT MIRON → subproveedor de núcleos (Villa Martelli): proceso tercerizado y no inspeccionado; la ventana se cerró el 06-11-26.  ⚠ PA-17",
]
# HARG no tiene contrato: la columna de inicio no debe leerse como fecha de OC
INICIO_OVR={"HARG — HIMA Argentina · ingeniería, programación y FAT":"OC PENDIENTE"}
# la entrega de una provisión no siempre es el fin más tardío: en Inauco la configuración
# post-FAT se extiende 13 días más allá de la liberación de los tableros
ENTREGA_OVR={"Inauco — PCS / SCADA / Comunicaciones":date(2027,5,4)}
def entrega_de(name,ms): return ENTREGA_OVR.get(name) or max(m[3] for m in ms)
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
ws.merge_cells("A3:K3"); cell(ws,3,1,f"Fecha de corte del estado: {HOY.strftime('%d/%m/%Y')}   ·   CAMINO CRÍTICO (sistemas): OC HARG → HPH → HARG → Inauco, cierre 24/05/2027 (Rev12: 10/05/2027)   ·   ÚLTIMA ENTREGA DEL PAQUETE: ductos de barras 08/07/2027   ·   AVANCE a OCT-26: PMS 54% (en plan, pero fabricación 0%) · Salas SE#3 34% y SE#4 49% · ductos 15%   ·   OC de HARG POSTERGADA por segunda vez a la semana del 16-11-26   ·   {len(PUNTOS)-2} puntos de atención ABIERTOS de PA-01 a PA-21 (13 en severidad alta; PA-03 y PA-06 cerrados)",None,"C62828",True,9)
hdr=["Proveedor","Provisión","Inicio contrato","Entrega","Hitos","Cumplidos s/plan","En curso","Pendientes","% avance plan","Próximo hito","Fecha próx. hito"]
for j,h in enumerate(hdr): cell(ws,5,1+j,h,HEAD,"FFFFFF",True,9,True,True)
ws.row_dimensions[5].height=30
r=6
for (name,ven,col,ms) in PROV:
    tot=len(ms); cum=sum(1 for m in ms if estado(m[2],m[3])=="Cumplido s/plan")
    enc=sum(1 for m in ms if estado(m[2],m[3])=="En curso"); pen=tot-cum-enc
    entrega=entrega_de(name,ms); prox=next((m for m in ms if m[3]>=HOY), None)
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
    cell(cw,r,1,f"▶  {name}   ·   {ven}   ·   Entrega: {entrega_de(name,ms).strftime('%d/%m/%Y')}",col,"FFFFFF",True,10)
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
kw.merge_cells("A2:H2"); cell(kw,2,1,"Cadena crítica armada con la colocación de la OC de HARG y las provisiones HPH (HIMA Paul Hildebrandt, Alemania) → HARG (HIMA Argentina) → Inauco (PCS/SIS). Es la secuencia que gobierna la fecha de entrega del sistema de control y seguridad.",None,"555555",False,9,False,True)
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
r+=2
# --- sensibilidad al corrimiento de la OC de HARG ---
cell(kw,r,1,"Sensibilidad: qué pasa si la OC de HIMA Argentina se corre más allá del 01-11-2026",None,NAVY,True,11)
kw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8); r+=1
cell(kw,r,1,"La provisión no tiene holgura: cada día de demora en la OC se traslada día por día a todas sus actividades y al cierre del camino crítico.",None,"555555",False,9,False,True)
kw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8); r+=1
SH2=["Escenario","Colocación de la OC","Inicio de trabajos","Adecuación de tableros","FAT HARG","IFAT HARG","CIERRE DEL ALCANCE HARG"]
for j,h in enumerate(SH2): cell(kw,r,1+j,h,HEAD,"FFFFFF",True,8.5,True,True)
kw.row_dimensions[r].height=26; r+=1
import datetime as _dt
_B=D(2026,9,14)
for extra,lab in [(0,"Base — OC el 01-11-26"),(7,"OC +1 semana"),(14,"OC +2 semanas"),(28,"OC +4 semanas")]:
    sh=SHIFT_HARG+extra
    f=lambda d: (d+_dt.timedelta(sh)).strftime("%d/%m/%Y")
    base = extra==0
    fill = "E8F5E9" if base else ("FFF3E0" if extra<=14 else "FFCDD2")
    cell(kw,r,1,lab,fill,("2E7D32" if base else "C62828"),True,9,False,True)
    cell(kw,r,2,(OC_HARG+_dt.timedelta(extra)).strftime("%d/%m/%Y"),fill,"000000",True,9,True)
    cell(kw,r,3,f(D(2026,9,14)),fill,"000000",False,9,True)
    cell(kw,r,4,f(D(2026,12,21))+" → "+f(D(2027,1,29)),fill,"000000",False,8.5,True)
    cell(kw,r,5,f(D(2027,2,24))+" → "+f(D(2027,2,26)),fill,"000000",False,8.5,True)
    cell(kw,r,6,f(D(2027,3,1))+" → "+f(D(2027,3,12)),fill,"000000",False,8.5,True)
    cell(kw,r,7,f(D(2027,3,22)),fill,("2E7D32" if base else "C62828"),True,9.5,True)
    r+=1
cell(kw,r,1,"Referencia: la liberación de tableros de Inauco es el 04-05-2027. Ya con una semana de corrimiento adicional el cierre de HARG (17-05-27) la supera en dos semanas.",
     "FFF9C4","B8860B",True,9,False,True); kw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=8); r+=2
for txt in [
 "Duración total de la cadena crítica: 10-07-2026 → 10-05-2027 (304 días corridos desde el inicio de HPH).",
 "El eslabón 11 (COLOCACIÓN DE LA OC DE HIMA ARGENTINA) sigue PENDIENTE al 21-10-2026. Se asume su colocación el 01-11-26, con inicio de trabajos el 02-11-26. Quedan 11 días para colocarla sin mover el plan: a partir de ahí, cada día de demora corre un día el cierre del alcance HARG. Ver la tabla de sensibilidad al pie.",
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

# ---------------- Hoja: Comparativo entre provisiones ----------------
cw2=wb.create_sheet("Comparativo provisiones")
cw2.merge_cells("A1:H1"); cell(cw2,1,1,f"COMPARATIVO ENTRE PROVISIONES  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); cw2.row_dimensions[1].height=22
cw2.merge_cells("A2:H2"); cell(cw2,2,1,"Estado de las ocho provisiones críticas: revisión vigente del cronograma, fecha de entrega, corrimientos, avance declarado, informe de avance e insumo de AESA pendiente.",None,"555555",False,9)
CH=["Provisión","Cronograma vigente","Estado del documento","Entrega","Cambio respecto de la revisión anterior","Avance declarado","Informe de avance","Insumo de AESA pendiente"]
for j,h in enumerate(CH): cell(cw2,4,1+j,h,HEAD,"FFFFFF",True,8.5,True,True)
cw2.row_dimensions[4].height=30
r=5
for (nm,cro,est,ent,cam,av,inf,ins) in COMPARATIVO:
    sinf = inf.startswith("—")
    cell(cw2,r,1,nm,None,"000000",True,9,False,True)
    cell(cw2,r,2,cro,None,"1F3864",False,8.5,True)
    cell(cw2,r,3,est,("2E7D32" if "construcción" in est else None),("FFFFFF" if "construcción" in est else "555555"),True,8.5,True)
    cell(cw2,r,4,ent,"1F3864","FFFFFF",True,9,True)
    cell(cw2,r,5,cam,None,("C62828" if ("SIN CONFIRMAR" in cam or "no" in cam[:3]) else "000000"),False,8.5,False,True)
    cell(cw2,r,6,av,None,"000000",True,8.5,False,True)
    cell(cw2,r,7,inf,("C62828" if sinf else "2E7D32"),"FFFFFF",True,8.5,True)
    cell(cw2,r,8,ins,"FFF9C4","B8860B",False,8.5,False,True)
    cw2.row_dimensions[r].height=max(26,12*(1+len(ins)//46)); r+=1
r+=1
cell(cw2,r,1,"Observaciones del comparativo",None,NAVY,True,11); r+=1
for (tit,txt) in COMPARATIVO_OBS:
    cell(cw2,r,1,tit,"ECEFF1","333333",True,9,False,True)
    cell(cw2,r,2,txt,None,"000000",False,8.5,False,True)
    cw2.merge_cells(start_row=r,start_column=2,end_row=r,end_column=8)
    cw2.row_dimensions[r].height=max(16,12*(1+len(txt)//120)); r+=1
for c,w in zip("ABCDEFGH",[34,19,17,13,38,24,16,44]): cw2.column_dimensions[c].width=w
cw2.freeze_panes="A5"

# ---------------- Hoja: Avance a OCT-2026 ----------------
vw=wb.create_sheet("Avance a OCT-2026")
vw.merge_cells("A1:G1"); cell(vw,1,1,f"AVANCE A OCTUBRE 2026  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); vw.row_dimensions[1].height=22
vw.merge_cells("A2:G2"); cell(vw,2,1,"Informes del período: ABB PMS Nº 5 (E-2611027, 30-10-2026) y ABB Salas Eléctricas Nº 6 (R-2631036_A, 04-11-2026). Vigentes de períodos anteriores: HPH / HIMA Alemania Nº 1 (AC-PR-2026-002, 15-10-2026) y CAT Miron acta de inspección Nº 1 (IN-FC-CAT-19082-261021, 21-10-2026). Más el avance declarado por cada proveedor en su propio cronograma.",None,"555555",False,9)
r=4
cell(vw,r,1,"1 · ABB SALAS ELÉCTRICAS — avance ponderado por paquete (informe Nº 6, 04-11-2026)",None,NAVY,True,11); r+=1
SH=["Paquete","sep-26 (real por especialidad)","oct-26 (previsto)","nov-26 (previsto)"]
for j,h in enumerate(SH): cell(vw,r,1+j,h,HEAD,"FFFFFF",True,9,True,True)
vw.row_dimensions[r].height=26; r+=1
for (pk,a,sep,oct_) in AVANCE_SALAS:
    bajo = int(sep.rstrip('%'))<40
    cell(vw,r,1,pk,None,"000000",False,9,False,True)
    cell(vw,r,2,a,None,"555555",False,9,True)
    cell(vw,r,3,sep,("C62828" if bajo else "1565C0"),"FFFFFF",True,9.5,True)
    cell(vw,r,4,oct_,None,"555555",False,9,True)
    r+=1
cell(vw,r,1,"ABB informa el avance REAL sólo del mes previo (sep-26) y por especialidad; las columnas de octubre y noviembre son PREVISTOS del propio proveedor. La fila TOTAL GLOBAL de septiembre figura en 0% en las nueve tablas de origen: es un artefacto de la planilla de ABB. En la tabla de tableros de MT los pesos por especialidad suman 67% y no 100%, también error de origen.",
     None,"555555",False,8.5,False,True); vw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7); r+=2

cell(vw,r,1,"2 · ABB PMS — avance físico ponderado por especialidad (informe Nº 5, 30-10-2026)",None,NAVY,True,11); r+=1
AH=["Especialidad","Peso","sep-26 real","oct-26 previsto","oct-26 REAL","nov-26 previsto","nov-26 pronóstico"]
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
cell(vw,r,1,"Avance real de octubre 54% contra 54% previsto: EXACTAMENTE EN PLAN. Pero la especialidad Fabricación (20% de peso) sigue en 0%: el armado de tableros no arrancó en octubre y se corre a noviembre, en MEHCCO S.A. Ver PA-18.",
     "FFF3E0","C62828",True,9.5,False,True); vw.merge_cells(start_row=r,start_column=1,end_row=r,end_column=7); r+=2

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
cell(vw,r,1,"4 · Previsión de fechas y lugares de FAT — ABB Salas (informe Nº 6) y prueba de integración (informe PMS Nº 5)",None,NAVY,True,11); r+=1
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
cell(vw,r,1,"7 · CAT Miron — acta de inspección Nº 1 (21-10-2026)",None,NAVY,True,11); r+=1
for (tit,txt) in CAT_TXT:
    cell(vw,r,1,tit,"ECEFF1","333333",True,9,False,True)
    cell(vw,r,2,txt,None,"000000",False,8.5,False,True)
    vw.merge_cells(start_row=r,start_column=2,end_row=r,end_column=7)
    vw.row_dimensions[r].height=max(16,12*(1+len(txt)//118)); r+=1
r+=1
cell(vw,r,1,"8 · Informe de avance HPH (HIMA Alemania) — síntesis del período",None,NAVY,True,11); r+=1
for (tit,txt) in HPH_TXT:
    cell(vw,r,1,tit,"ECEFF1","333333",True,9,False,True)
    cell(vw,r,2,txt,None,"000000",False,8.5,False,True)
    vw.merge_cells(start_row=r,start_column=2,end_row=r,end_column=7)
    vw.row_dimensions[r].height=max(16,12*(1+len(txt)//118)); r+=1
r+=1
cell(vw,r,1,"9 · Informe de avance ABB PMS — síntesis del período",None,NAVY,True,11); r+=1
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
    print(f"  {name[:50]:50} hitos={len(ms):2} acopio={ac} CC={cp:2} cumpl={cum:2} entrega={entrega_de(name,ms).strftime('%d/%m/%Y')}")
# control: todas las claves de CP_MARK y PA_MARK deben existir
todos={(n,m[1]) for (n,v,c,ms) in PROV for m in ms}
falt=[k for k in list(CP_MARK)+list(PA_MARK) if k not in todos]
print("claves sin correspondencia:",falt if falt else "ninguna")
