#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CRONOGRAMA EJECUTIVO — Provisiones críticas CPF-2 (Rev3, 2026-09-07).
Cambios respecto de Rev2:
  1) PUNTOS DE ATENCIÓN: nueva hoja dedicada. Se explicita que los tableros de HPH,
     entregados DAP Ezeiza el 15-12-26, requieren nacionalización / despacho de aduana
     (mínimo 3-4 semanas, agravado por el receso de fin de año), por lo que la adecuación
     prevista para el 21-12-26 es muy poco probable; y que la recepción en base Inauco el
     25-01-27 "ya casi terminada" también es poco probable. Se agrega el hito explícito de
     nacionalización en la provisión HPH.
  2) La prueba de integración PCS↔PMS en Bs.As. se ANTICIPA a la ventana OCT-NOV 2026
     (fecha exacta a convenir), en lugar de abr-may/2027. Se refleja en Inauco y en ABB PMS.
  3) CAMINO CRÍTICO: nueva hoja con la cadena crítica del proyecto, armada con las
     provisiones de HPH (Alemania) → HARG (Argentina) → Inauco, y marcado de los hitos
     de la ruta crítica en el cronograma por hitos.
Corte de alcance: hasta la ENTREGA DE EQUIPOS / del alcance de cada provisión.
"""
import os
from datetime import date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
REV="Rev3"; FECHA="2026-09-07"; HOY=date(2026,9,7)
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
  ("Inicio de construcción de tableros",D(2026,11,27),D(2026,11,27),0,"",0),
  ("Recepción tableros HIMA en base Inauco (insumo HPH)",D(2027,1,25),D(2027,1,25),1,"Requisito 2 — interfase HPH/HARG",0),
  ("Configuración PLC — Etapa 2",D(2027,1,25),D(2027,3,23),0,"",0),
  ("Fin de construcción de tableros",D(2027,2,16),D(2027,2,16),1,"",0),
  ("Aprobación de procedimientos FAT",D(2027,2,24),D(2027,2,24),0,"Requisito 4",0),
  ("FAT PCS + Comunicaciones",D(2027,3,29),D(2027,4,22),1,"Witness AESA",0),
  ("FAT SIS (conjunto con HARG/HPH)",D(2027,3,29),D(2027,4,26),1,"Requiere presencia HIMA",0),
  ("Configuraciones post-FAT",D(2027,4,27),D(2027,5,17),0,"",0),
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
  ("Nacionalización / despacho de aduana + transporte a base Inauco",D(2026,12,15),D(2027,1,25),1,"Mínimo 3-4 semanas + receso de fin de año",0),
  ("RECEPCIÓN DE TABLEROS EN BASE INAUCO",D(2027,1,25),D(2027,1,25),1,"HITO FINAL — habilita la adecuación de HARG",0),
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

# ---------------- Puntos de atención (Rev3) ----------------
# clave: (provisión, hito) -> código de PA
PA_MARK={
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Entrega DAP Ezeiza (Argentina)"):"PA-01",
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","Nacionalización / despacho de aduana + transporte a base Inauco"):"PA-01",
 ("HPH — HIMA Paul Hilpert (Alemania) · tableros marshalling","RECEPCIÓN DE TABLEROS EN BASE INAUCO"):"PA-02",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Fabricación HPH — entrega DIC-2026 (INTERFASE)"):"PA-01",
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Adecuación de tableros marshalling (INTERFASE Inauco)"):"PA-02",
 ("Inauco — PCS / SCADA / Comunicaciones","Recepción tableros HIMA en base Inauco (insumo HPH)"):"PA-02",
 ("Inauco — PCS / SCADA / Comunicaciones","Pruebas de integración PCS↔PMS (Bs.As.) — ANTICIPADA"):"PA-03",
 ("ABB — PMS (Power Management System)","Pruebas de integración PCS↔PMS (Bs.As.) — ANTICIPADA"):"PA-03",
}
PUNTOS=[
 dict(cod="PA-01", tit="Nacionalización de los tableros de marshalling entregados por HPH",
      inter="HPH (Alemania) → HARG / Inauco",
      plan="HPH entrega DAP Ezeiza el 15-12-26. El cronograma de HARG (PSCH_CPF2_Rev3) inicia la adecuación de los tableros el 21-12-26, es decir 6 días corridos después del arribo.",
      riesgo="El despacho de aduana de un embarque de tableros importados demanda como mínimo 3 a 4 semanas (documentación de embarque, intervención, liberación y transporte a la base de Inauco), y el arribo cae de lleno en el receso de fin de año, cuando la operatoria aduanera y logística se reduce.",
      impacto="Iniciar la adecuación el 21-12-26 es MUY POCO PROBABLE. Corrimiento estimado del inicio: 4 a 6 semanas.",
      real="Disponibilidad realista de los tableros en base Inauco: entre la 3.ª semana de enero y la 1.ª semana de febrero de 2027.",
      accion="Designar despachante y anticipar la documentación de embarque con HPH; adelantar el despacho de origen si es posible; evaluar despacho parcial / aéreo de los componentes que gobiernan el inicio de la adecuación; incorporar el hito de nacionalización al seguimiento semanal.",
      sev="ALTA"),
 dict(cod="PA-02", tit="Recepción en base Inauco el 25-01-27 con el tablero 'ya casi terminado'",
      inter="HPH ↔ HARG ↔ Inauco",
      plan="Inauco fija en su Requisito 2 la recepción de los tableros HIMA en su base el 25-01-27, asumiendo el equipo prácticamente terminado. HARG planifica 6 semanas de adecuación (21-12-26 → 29-01-27).",
      riesgo="Las dos premisas son incompatibles entre sí y con PA-01: si la nacionalización recién concluye alrededor del 25-01-27, la adecuación de 6 semanas terminaría a comienzos de marzo/27, no el 29-01-27. Recibir el 25-01-27 un tablero 'casi terminado' es POCO PROBABLE.",
      impacto="Se comprime la Configuración PLC Etapa 2 de Inauco (25-01 → 23-03-27) y quedan en riesgo la FAT de HARG (24-26/02/27), el IFAT (01-12/03/27) y la FAT SIS conjunta (29-03 → 26-04-27), que es el cierre del camino crítico.",
      real="Fin de adecuación estimado: entre la 1.ª y la 2.ª semana de marzo de 2027.",
      accion="Redefinir con HPH / HARG / Inauco la secuencia, el lugar y la duración de la adecuación; evaluar ejecutar la adecuación en origen (HPH) antes del despacho, o solapar adecuación con la programación de HARG; reconfirmar las fechas de FAT/IFAT y de la FAT SIS conjunta.",
      sev="ALTA"),
 dict(cod="PA-03", tit="Prueba de integración PCS↔PMS en Buenos Aires — fecha a convenir",
      inter="Inauco (PCS) ↔ ABB (PMS)",
      plan="Se anticipa respecto de la Rev2: pasa de abr-may/2027 a la ventana OCT-NOV 2026.",
      riesgo="La fecha exacta dentro de la ventana aún no está acordada entre AESA, Inauco y ABB, y requiere disponibilidad simultánea de la configuración PMS (en curso desde 17-08-26) y del PCS (Configuración PLC Etapa 1 / SCADA).",
      impacto="Anticiparla es favorable: detecta tempranamente incompatibilidades de comunicación (IEC61850 / Modbus / Profibus) antes de las FAT. Sin fecha acordada, el ítem no puede planificarse ni asignarse recursos.",
      real="Ventana OCT-NOV 2026; se sugiere fijarla luego del cierre del acopio de switches de red del PMS (05-11-26).",
      accion="Convocar reunión de coordinación AESA / Inauco / ABB para fijar la fecha exacta y el alcance de la prueba, y emitir el protocolo correspondiente.",
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
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Programación HARG — Recurso 1"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","Adecuación de tableros marshalling (INTERFASE Inauco)"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","FAT"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","IFAT (Recursos 1 y 2)"),
 ("HARG — HIMA Argentina · ingeniería, programación y FAT","ENTREGA DEL ALCANCE HARG (IFAT + documentación)"),
 ("Inauco — PCS / SCADA / Comunicaciones","Recepción tableros HIMA en base Inauco (insumo HPH)"),
 ("Inauco — PCS / SCADA / Comunicaciones","Configuración PLC — Etapa 2"),
 ("Inauco — PCS / SCADA / Comunicaciones","FAT SIS (conjunto con HARG/HPH)"),
 ("Inauco — PCS / SCADA / Comunicaciones","Configuraciones post-FAT"),
 ("Inauco — PCS / SCADA / Comunicaciones","ENTREGA DE EQUIPOS (post-FAT)"),
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
 (9,"HPH","Nacionalización / despacho de aduana + transporte a base Inauco",D(2026,12,15),D(2027,1,25),"8","⚠ PA-01 — mín. 3-4 semanas + receso de fin de año"),
 (10,"HPH→Inauco","RECEPCIÓN DE TABLEROS EN BASE INAUCO",D(2027,1,25),D(2027,1,25),"9","⚠ PA-02 — hito bisagra HPH → HARG / Inauco"),
 (11,"HARG","Programación HARG — Recursos 1 y 2",D(2026,9,21),D(2027,1,29),"—","Ruta paralela; converge en la FAT"),
 (12,"HARG","Adecuación de tableros de marshalling (6 semanas)",D(2026,12,21),D(2027,1,29),"10","⚠ PA-02 — el plan la inicia ANTES de la recepción"),
 (13,"Inauco","Configuración PLC — Etapa 2",D(2027,1,25),D(2027,3,23),"10",""),
 (14,"HARG","FAT",D(2027,2,24),D(2027,2,26),"11, 12",""),
 (15,"HARG","IFAT (Recursos 1 y 2)",D(2027,3,1),D(2027,3,12),"14","Cierre del alcance HARG: 22-03-27"),
 (16,"Inauco","FAT SIS (conjunta HARG / HPH / Inauco)",D(2027,3,29),D(2027,4,26),"13, 15","La FAT cierra con el ítem más tardío"),
 (17,"Inauco","Configuraciones post-FAT",D(2027,4,27),D(2027,5,17),"16",""),
 (18,"Inauco","ENTREGA DE EQUIPOS (post-FAT)",D(2027,5,17),D(2027,5,17),"17","HITO FINAL DEL CAMINO CRÍTICO"),
]

NOTAS=[
 "Alcance: cronograma EJECUTIVO por hitos, con corte en la ENTREGA DE EQUIPOS / del alcance de cada provisión. Las tareas posteriores al despacho/entrega (montaje en planta, SAT, precomisionado, comisionado, PEM, CAO/DataBook) quedan FUERA.",
 "Rev3 — cambio 1: se agregó la hoja 'Puntos de atención' (PA-01 a PA-03) y el hito explícito de NACIONALIZACIÓN / DESPACHO DE ADUANA en la provisión HPH, entre la entrega DAP Ezeiza y la recepción en base Inauco.",
 "Rev3 — cambio 2: la PRUEBA DE INTEGRACIÓN PCS↔PMS en Bs.As. se ANTICIPA a la ventana OCT-NOV 2026 (fecha exacta a convenir); antes figuraba en abr-may/2027. Se refleja en las provisiones Inauco y ABB PMS.",
 "Rev3 — cambio 3: se agregó la hoja 'Camino crítico' con la cadena HPH (Alemania) → HARG (Argentina) → Inauco, y los hitos de esa ruta quedan marcados con ◆ en el cronograma por hitos.",
 "ALERTA DE INTERFASE A CONCILIAR (ver PA-01 y PA-02): HPH entrega DAP Ezeiza el 15-12-26; el cronograma HARG asume disponibilidad para adecuación el 21-12-26 sin considerar la nacionalización; e Inauco fija la recepción en su base el 25-01-27 con el tablero casi terminado (Requisito 2). Las tres fechas no son compatibles entre sí.",
 "Rev2 — cambio 1: en ABB Salas SE#3 y SE#4 se incorporaron los hitos de fabricación, ensayos internos y FAT EN FÁBRICA DE ORIGEN (Brasil, M&B) para tableros y celdas, más la adecuación/despacho de fábrica y la entrega en Shelterista, entre el acopio y el montaje del shelter.",
 "Rev2 — cambio 2: el ACOPIO DE MATERIALES se abrió en hitos propios por paquete y se agregó la hoja 'Acopio de Materiales' con el detalle consolidado (pedido del cliente para medir avance).",
 "Rev2 — cambio 3: HIMA se separó en HPH (HIMA Paul Hilpert, Alemania — tableros de marshalling) y HARG (HIMA Argentina — ingeniería de aplicación, programación, documentación y FAT/IFAT), según PSCH_CPF2_Rev3 del 04-09-2026.",
 "Fuentes: ABB Salas y Ductos 5155-00-0009-VE-CO-PG-001_A (21-08-26) · ABB PMS 5155-00-0007-VE-CO-PG-001_REV B (.mpp) · Inauco 5155-00-0029-VE-CO-PG-001-A Rev.A (24-08-26) · HPH 103618-Schedule 16-08-26 · HARG PSCH_CPF2_Rev3 (04-09-26).",
 "Los hitos de FAT en fábrica de origen (Brasil) son witness de AESA y anteceden a la FAT del shelter completo; no deben confundirse entre sí.",
 "Estado (plan): se calcula contra la fecha de corte. 'Cumplido s/plan' = la fecha planificada ya venció y debe confirmarse el cumplimiento real con el proveedor; 'En curso' = en ejecución; 'Pendiente' = aún no inició.",
 "Las columnas Fecha real / % avance / Estado real / Observaciones quedan en blanco para la carga de seguimiento periódico (reporte a PPSA).",
 "Los hitos se ordenan cronológicamente dentro de cada provisión y se renumeran (H01, H02, ...). Marcas: ◆ = hito del camino crítico · ⚠ = punto de atención · fondo naranja = acopio de materiales.",
]
INTERFASES=[
 "CAMINO CRÍTICO DEL PROYECTO: HPH (Alemania) → nacionalización → HARG (adecuación y programación) → Inauco (config. PLC Etapa 2, FAT SIS y entrega). Cierra el 17-05-27. Ver hoja 'Camino crítico'.",
 "HPH → entrega DAP Ezeiza (15/12/26) + nacionalización → recepción en base Inauco (25/01/27): habilita la adecuación, la configuración PLC Etapa 2 y la FAT SIS.  ⚠ PA-01 / PA-02",
 "HARG ↔ HPH → ingeniería de aplicación SIS y documentación de HPH (nov-dic/26): condicionan la programación y el FAT/IFAT de HARG.",
 "HARG ↔ Inauco → adecuación de tableros marshalling (plan 21/12/26 → 29/01/27) en base Inauco.  ⚠ PA-02",
 "ABB PMS ↔ Inauco → pruebas de integración PCS↔PMS en Bs.As. ANTICIPADAS a la ventana OCT-NOV 2026, fecha a convenir.  ⚠ PA-03",
 "ABB PMS → entrega de equipamiento en shelter (17/02/27): habilita el montaje del PMS en las salas y los ensayos PMS.",
 "AESA → layout de instalación + detalle de conexionado a trafo (29/08/26): condiciona el inicio de diseño de los DUCTOS DE BARRAS (entrega 10/06/27, la más tardía).",
 "AESA → entrega de ingeniería Rev.0 a Inauco (10/08/26) y aprobaciones de ingeniería a ABB/HPH: condicionan los freezing points.",
]

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
MESES=month_list(D(2026,5,1),D(2027,6,1))
M3=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

wb=openpyxl.Workbook()

# ---------------- Hoja 1: Resumen Ejecutivo ----------------
ws=wb.active; ws.title="Resumen Ejecutivo"
ws.merge_cells("A1:K1"); cell(ws,1,1,f"CRONOGRAMA EJECUTIVO — PROVISIONES CRÍTICAS CPF-2  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,13,True); ws.row_dimensions[1].height=24
ws.merge_cells("A2:K2"); cell(ws,2,1,"Seguimiento para PPSA · Corte de alcance: hasta la ENTREGA DE EQUIPOS / del alcance (excluye montaje en planta, SAT, precom/comisionado y PEM)",None,"555555",False,9)
ws.merge_cells("A3:K3"); cell(ws,3,1,f"Fecha de corte del estado: {HOY.strftime('%d/%m/%Y')}   ·   CAMINO CRÍTICO: HPH → HARG → Inauco, cierre 17/05/2027   ·   3 puntos de atención abiertos (PA-01 a PA-03)",None,"C62828",True,9)
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
kw.merge_cells("A2:H2"); cell(kw,2,1,"Cadena crítica armada con las provisiones HPH (HIMA Paul Hilpert, Alemania) → HARG (HIMA Argentina) → Inauco (PCS/SIS). Es la secuencia que gobierna la fecha de entrega del sistema de control y seguridad.",None,"555555",False,9,False,True)
kw.row_dimensions[2].height=26
KH=["Nº","Proveedor","Actividad / hito de la cadena","Inicio","Fin","Días","Predecesor","Observación / riesgo"]
for j,h in enumerate(KH): cell(kw,4,1+j,h,HEAD,"FFFFFF",True,9,True,True)
kw.row_dimensions[4].height=26
VENCOL={"HPH":"6A1B9A","HARG":"00838F","Inauco":"1565C0","HPH→Inauco":"4527A0"}
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
 "Duración total de la cadena crítica: 10-07-2026 → 17-05-2027 (312 días corridos).",
 "La cadena NO tiene holgura declarada en los cronogramas de los proveedores: cada corrimiento de HPH se traslada íntegro a la fecha de entrega de Inauco.",
 "Los eslabones 8, 9, 10 y 12 concentran el riesgo del proyecto (PA-01 y PA-02): entrega DAP Ezeiza, nacionalización, recepción en base Inauco y adecuación de los tableros de marshalling.",
 "Los eslabones 11 y 13 (programación HARG y configuración PLC Etapa 2 de Inauco) son rutas paralelas que convergen en las FAT; hoy no gobiernan la fecha final, pero pasan a ser críticas si la adecuación se corre más de 4 semanas.",
 "El resto de las provisiones ABB (Salas SE#3 y SE#4, PMS y Ductos de barras) corren por caminos propios; el Ducto de barras entrega el 10-06-27, posterior al cierre de esta cadena, pero es un equipo de montaje, no de sistema.",
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
