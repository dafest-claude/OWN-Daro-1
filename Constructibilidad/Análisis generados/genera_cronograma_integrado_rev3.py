#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cronograma INTEGRADO CPF-2 — Rev3 (27-10-2026).
Actualiza la Rev2 (06-08-2026) con TODA la información incorporada desde entonces, que está
consolidada en el Cronograma Ejecutivo de Provisiones Críticas Rev11:

  · ABB SALAS SE#3 y SE#4 — cronograma 5155-00-0009-VE-CO-PG-001 REV 0 "Para construcción"
    (14-10-26) + informe de avance Nº 5 (05-10-26). Entregas EXW: SE#4 25-03-27 y SE#3
    20-04-27 (contractuales 08-04-27 y 14-05-27). FAT con lugar y fecha por equipo.
  · ABB DUCTOS DE BARRAS — provisión separada. Entrega 08-07-27, PLAZO SIN CONFIRMAR: ABB
    pide adecuar la fecha contractual y el subproveedor EAE (Turquía) figura con "tbd".
  · ABB PMS — cronograma SEP-2026 REV 0 (01-10-26) + informe Nº 4. Entrega 17-02-27.
  · CAT MIRON — TRANSFORMADORES (nueva provisión): 13 unidades en 5 OT, entregas
    escalonadas del 13-11 al 11-12-2026.
  · INAUCO — cronograma 5155-00-0029-VE-CA-PG-001 REV 0 "Para construcción" (23-10-26).
    Sin corrimientos. La prueba de integración con ABB se adelantó al 01-02/12/2026.
  · HIMA se separa en HPH (Alemania, marshalling — fabricación COMPLETADA, entrega DAP
    Ezeiza 15-12-26) y HARG (Argentina — OC AÚN NO COLOCADA, cierre 10-05-27).
  · INTEGRACIÓN EPC — reencajada: el montaje de los ductos de barras no puede iniciar
    antes del 08-07-27, que es hoy la última entrega del paquete.

Genera: HTML interactivo (Gantt + Flujo temporal + Tabla + Evaluación), Gantt PDF A3,
Flujo PDF A3 y XLSX. No borra revisiones anteriores.
"""
import os, re, json
from datetime import date

OUT="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
BASE=os.path.join(OUT,"Cronograma_Integrado_CPF2_Rev2_2026-08-06.html")
REV="Rev3"; FECHA="2026-10-27"
D=lambda y,m,d: date(y,m,d)

# ---------------------------------------------------------------- provisiones
PROV=[  # (clave, nombre corto, nombre HTML, var CSS, hex)
 ("ABBS","ABB · Salas","ABB · Salas SE#3/SE#4","--abbs","#C62828"),
 ("ABBD","ABB · Ductos","ABB · Ductos de barras","--abbd","#795548"),
 ("ABBP","ABB · PMS","ABB · PMS","--abbp","#EF6C00"),
 ("MIRON","CAT Miron","CAT Miron · Transf.","--miron","#33691E"),
 ("INA","Inauco · PCS","Inauco · PCS/SCADA","--inauco","#1565C0"),
 ("HPH","HPH · Marsh.","HPH · Marshalling (DE)","--hph","#6A1B9A"),
 ("HARG","HARG · SIS","HARG · SIS (AR)","--harg","#00838F"),
 ("EPC","Integr. EPC","Integración EPC","--epc","#2E7D32"),
]
PHEX={k:h for (k,_,_,_,h) in PROV}

# ---------------------------------------------------------------- tareas
# (clave, texto, inicio, fin|None=hito, marca)
T=[
 # ---- ABB SALAS ----
 ("ABBS","Ingeniería básica y de detalle (equipos APROBADOS · sala pendiente)",D(2026,5,20),D(2026,10,13),""),
 ("ABBS","◆ Freezing point ing. básica de salas — VENCIDO 25-09-26",D(2026,9,25),None,"◆"),
 ("ABBS","Acopio de materiales (CCM 93% · UPS 93% · auxiliares 72%)",D(2026,5,19),D(2026,12,10),""),
 ("ABBS","Fabricación de tableros en origen (Sorocaba · Santa Fe · AMBA)",D(2026,10,6),D(2026,12,11),""),
 ("ABBS","FAT en fábrica de origen (UPS · TGMT · CCM · auxiliares)",D(2026,11,12),D(2026,12,16),""),
 ("ABBS","Entrega en Shelterista (Mendoza)",D(2026,11,26),D(2027,1,10),""),
 ("ABBS","Fabricación de los shelters (Bottino)",D(2026,10,6),D(2027,3,5),""),
 ("ABBS","Montaje e integración de equipos en salas",D(2027,1,29),D(2027,3,17),""),
 ("ABBS","FAT Salas en Mendoza (SE#4 22-26/02 · SE#3 29/03-02/04)",D(2027,2,22),D(2027,4,2),""),
 ("ABBS","Ensayos PMS en sala (interfase contrato PMS)",D(2027,3,3),D(2027,4,6),""),
 ("ABBS","★ Entrega EXW SE#4 (25-03-27) · contractual 08-04-27",D(2027,3,25),None,"★"),
 ("ABBS","★ Entrega EXW SE#3 (20-04-27) · contractual 14-05-27",D(2027,4,20),None,"★"),
 # ---- ABB DUCTOS ----
 ("ABBD","Insumo AESA: layout + detalle de conexionado a trafo",D(2026,7,20),D(2026,9,18),""),
 ("ABBD","Diseño de ductos (arrancó efectivamente el 21-09-26) — 15%",D(2026,9,18),D(2026,12,14),""),
 ("ABBD","Aprobación de diseño + freezing point (30-12-26)",D(2026,12,14),D(2026,12,30),""),
 ("ABBD","Acopio y fabricación (EAE Turquía — entrega en origen «tbd»)",D(2026,12,30),D(2027,4,9),""),
 ("ABBD","FCA + traslado internacional + aduana (80 d)",D(2027,4,9),D(2027,7,5),""),
 ("ABBD","★ Entrega en sitio (08-07-27) — PLAZO SIN CONFIRMAR",D(2027,7,8),None,"★"),
 # ---- ABB PMS ----
 ("ABBP","Ingeniería básica y Rev.0 + aprobación (freezing 11-09-26)",D(2026,6,8),D(2026,9,11),""),
 ("ABBP","Procura y acopio — gestión de compra al 100%",D(2026,5,22),D(2026,11,5),""),
 ("ABBP","Configuración / programación del sistema (32%)",D(2026,8,17),D(2026,12,11),""),
 ("ABBP","Armado de tableros (inicio corrido al 09-10-26)",D(2026,10,9),D(2026,12,17),""),
 ("ABBP","Maqueta integrada CPF1",D(2026,10,19),D(2026,10,23),""),
 ("ABBP","Maqueta CPF2 — integración PCS↔PMS con Inauco (Bs.As.)",D(2026,11,30),D(2026,12,3),""),
 ("ABBP","Pruebas internas + FAT tableros y sistema (witness AESA)",D(2026,12,18),D(2027,1,29),""),
 ("ABBP","Despachos",D(2027,2,4),D(2027,2,17),""),
 ("ABBP","★ Entrega de equipamiento en Shelter (17-02-27)",D(2027,2,17),None,"★"),
 ("ABBP","FAT integral del shelter (PMS) + DataBook",D(2027,3,25),D(2027,3,31),""),
 # ---- CAT MIRON ----
 ("MIRON","Ingeniería: cálculo, diseño y proyecto de detalle (5 OT)",D(2026,7,13),D(2026,9,4),""),
 ("MIRON","Acopio: conductores BT/MT y materiales gruesos (gabinetes)",D(2026,7,27),D(2026,10,16),""),
 ("MIRON","Fabricación: núcleo magnético, bobinado y montaje",D(2026,9,7),D(2026,12,4),""),
 ("MIRON","Ensayos internos · FAT NO incluidas en el cronograma",D(2026,11,2),D(2026,12,11),""),
 ("MIRON","★ Entregas escalonadas de los 13 transformadores (13-11 → 11-12-26)",D(2026,12,11),None,"★"),
 ("MIRON","Almacenamiento y conservación en obra (7 meses hasta los ductos)",D(2026,12,11),D(2027,7,8),""),
 # ---- INAUCO ----
 ("INA","◆ Recepción de OC — inicio del proyecto (28-07-26)",D(2026,7,28),None,"◆"),
 ("INA","Recepción de ingeniería de AESA — 5 entregas (2 vencidas)",D(2026,8,14),D(2026,12,11),""),
 ("INA","Emisión de ingeniería Rev.A y Rev.0 de tableros",D(2026,8,24),D(2026,11,13),""),
 ("INA","Configuración base — PLC/comunicaciones y SCADA (12%)",D(2026,8,26),D(2027,3,23),""),
 ("INA","◆ Freezing point: aprobación de ing. para construcción (19-11)",D(2026,11,19),None,"◆"),
 ("INA","Pruebas de integración con ABB PMS (Bs.As.) — ADELANTADA",D(2026,12,1),D(2026,12,2),""),
 ("INA","Construcción de tableros 07a / 07b / COM-07",D(2026,11,27),D(2027,2,16),""),
 ("INA","Pre-iFAT interna (previa a la FAT con el cliente)",D(2027,2,16),D(2027,3,23),""),
 ("INA","FAT PCS + Comunicaciones (Inauco Neuquén)",D(2027,3,29),D(2027,4,22),""),
 ("INA","FAT SIS — conjunto con HIMA (Inauco Neuquén)",D(2027,3,29),D(2027,4,26),""),
 ("INA","★ Fin de FAT integral (26-04-27)",D(2027,4,26),None,"★"),
 ("INA","Configuraciones post-FAT",D(2027,4,27),D(2027,5,17),""),
 ("INA","★ Liberación y entrega de tableros (04-05-27)",D(2027,5,4),None,"★"),
 ("INA","SAT y comisionado PCS — campo",D(2027,5,18),D(2027,8,3),""),
 ("INA","SAT y comisionado SCADA/SIS — campo",D(2027,5,18),D(2027,10,5),""),
 ("INA","Prueba MCE + cierre de documentación",D(2027,10,6),D(2027,12,7),""),
 # ---- HPH ----
 ("HPH","Ingeniería de detalle (80%) + acopio — material ENTREGADO",D(2026,8,17),D(2026,10,16),""),
 ("HPH","Fabricación de los 3 tableros de marshalling — COMPLETADA",D(2026,9,21),D(2026,10,15),""),
 ("HPH","Inspección / HW check de los 3 tableros (semana 44-45)",D(2026,10,26),D(2026,11,8),""),
 ("HPH","◆ Letter of Credit de la orden ampliada (AESA) — 30-10-26",D(2026,10,30),None,"◆"),
 ("HPH","Packing y despacho desde Alemania",D(2026,11,9),D(2026,12,14),""),
 ("HPH","★ Entrega DAP Ezeiza aeropuerto (15-12-26) — CONFIRMADA",D(2026,12,15),None,"★"),
 ("HPH","Nacionalización + transporte a base Inauco",D(2026,12,15),D(2027,1,25),""),
 ("HPH","★ Recepción de tableros en base Inauco (25-01-27)",D(2027,1,25),None,"★"),
 # ---- HARG ----
 ("HARG","◆ COLOCACIÓN DE LA OC (AESA) — PENDIENTE, asumida 01-11-26",D(2026,11,1),None,"◆"),
 ("HARG","Documentación y programación HARG (recursos 1 y 2)",D(2026,11,2),D(2027,3,19),""),
 ("HARG","Ingeniería de aplicación SIS (insumo de HPH)",D(2026,11,9),D(2026,12,18),""),
 ("HARG","Adecuación de tableros de marshalling en base Inauco",D(2027,2,8),D(2027,3,19),""),
 ("HARG","FAT HARG",D(2027,4,14),D(2027,4,16),""),
 ("HARG","IFAT — integración (recursos 1 y 2)",D(2027,4,19),D(2027,4,30),""),
 ("HARG","★ Entrega del alcance HARG (10-05-27) — fin del camino crítico",D(2027,5,10),None,"★"),
 # ---- INTEGRACIÓN EPC ----
 ("EPC","Montaje en planta de salas y equipos",D(2027,4,20),D(2027,9,30),""),
 ("EPC","Montaje de los ductos de barras (no puede iniciar antes del 08-07)",D(2027,7,8),D(2027,10,15),""),
 ("EPC","Precomisionado",D(2027,6,1),D(2027,12,15),""),
 ("EPC","iFAT — integración en sitio (shelters + PMS + PCS + SIS)",D(2027,8,2),D(2027,8,31),""),
 ("EPC","Comisionado / Puesta en marcha (PEM)",D(2027,10,1),D(2028,1,31),""),
 ("EPC","★ RFSU — Ready For Start Up (31-01-2028)",D(2028,1,31),None,"★"),
]

# ---------------------------------------------------------------- fases (flujo)
# (clave, texto, inicio, fin|None, crítico 0/1, ancla opcional)
PH=[
 ("ABBS","Ingeniería + acopio",D(2026,5,20),D(2026,12,10),0),
 ("ABBS","Fabricación y FAT en origen",D(2026,10,6),D(2026,12,16),0),
 ("ABBS","Shelters + montaje + FAT de sala (Mendoza)",D(2026,10,6),D(2027,4,2),1),
 ("ABBS","★ Entregas EXW SE#4 / SE#3",D(2027,4,20),None,1,"se3"),
 ("ABBD","Insumo AESA + diseño",D(2026,7,20),D(2026,12,30),1),
 ("ABBD","Fabricación (EAE) + traslado y aduana",D(2026,12,30),D(2027,7,5),1),
 ("ABBD","★ Entrega en sitio (s/confirmar)",D(2027,7,8),None,1,"duc"),
 ("ABBP","Ingeniería + procura + configuración",D(2026,5,22),D(2026,12,11),0),
 ("ABBP","Armado + maquetas + FAT PMS",D(2026,10,9),D(2027,1,29),0),
 ("ABBP","★ Entrega en Shelter",D(2027,2,17),None,0,"pms"),
 ("MIRON","Ingeniería + acopio + fabricación",D(2026,7,13),D(2026,12,4),0),
 ("MIRON","★ Entregas 13-11 → 11-12-26",D(2026,12,11),None,0,"trf"),
 ("MIRON","Almacenamiento en obra (espera a los ductos)",D(2026,12,11),D(2027,7,8),0),
 ("INA","Ingeniería + configuración base",D(2026,7,28),D(2027,3,23),0),
 ("INA","Construcción tableros + Pre-iFAT",D(2026,11,27),D(2027,3,23),1),
 ("INA","FAT PCS + FAT SIS (Neuquén)",D(2027,3,29),D(2027,4,26),1),
 ("INA","★ Liberación de tableros",D(2027,5,4),None,1,"ina"),
 ("INA","Campo: SAT / comisionado + MCE",D(2027,5,18),D(2027,12,7),0),
 ("HPH","Ing. + fabricación (COMPLETADA)",D(2026,8,17),D(2026,10,15),1),
 ("HPH","Inspección + despacho",D(2026,10,26),D(2026,12,14),1),
 ("HPH","★ DAP Ezeiza",D(2026,12,15),None,1,"hph"),
 ("HPH","Nacionalización",D(2026,12,15),D(2027,1,25),1),
 ("HARG","◆ OC pendiente (01-11-26)",D(2026,11,1),None,1,"oc"),
 ("HARG","Documentación + progr.",D(2026,11,2),D(2027,3,19),1),
 ("HARG","Adecuación marshalling",D(2027,2,8),D(2027,3,19),1),
 ("HARG","FAT + IFAT",D(2027,4,14),D(2027,4,30),1),
 ("HARG","★ Cierre HARG",D(2027,5,10),None,1,"harg"),
 ("EPC","Montaje en planta",D(2027,4,20),D(2027,10,15),0),
 ("EPC","Precomisionado",D(2027,6,1),D(2027,12,15),0),
 ("EPC","iFAT en sitio",D(2027,8,2),D(2027,8,31),1,"ifat"),
 ("EPC","Comisionado / PEM",D(2027,10,1),D(2028,1,31),0),
 ("EPC","★ RFSU",D(2028,1,31),None,1,"rfsu"),
]

KPIS=[
 ("Transformadores (último lote)","11-DIC-2026","--miron"),
 ("HPH marshalling · DAP Ezeiza","15-DIC-2026","--hph"),
 ("Entrega PMS en shelter","17-FEB-2027","--abbp"),
 ("Entregas EXW SE#4 / SE#3","25-MAR / 20-ABR-2027","--abbs"),
 ("Entrega Inauco (tableros)","04-MAY-2027","--inauco"),
 ("Cierre HARG — camino crítico","10-MAY-2027","--harg"),
 ("Ductos de barras — última entrega","08-JUL-2027","--abbd"),
 ("RFSU","31-ENE-2028","--epc"),
]
VLINES=[(D(2026,12,15),"DAP Ezeiza HPH","#6A1B9A"),(D(2027,2,17),"Entrega PMS","#EF6C00"),
        (D(2027,4,20),"Entrega SE#3","#C62828"),(D(2027,5,4),"Entrega Inauco","#1565C0"),
        (D(2027,5,10),"Cierre HARG","#00838F"),(D(2027,7,8),"Entrega ductos","#795548"),
        (D(2028,1,31),"RFSU","#2E7D32")]
SUB=("8 provisiones · Salas SE#3/SE#4, Ductos de barras y PMS (ABB) · Transformadores (CAT Miron) · PCS/SCADA (Inauco) · "
     "marshalling (HPH, Alemania) · SIS (HARG, Argentina) &nbsp;|&nbsp; Camino crítico de sistemas: <b>HARG 10-MAY-2027</b> · "
     "Última entrega del paquete: <b>ductos de barras 08-JUL-2027</b> · RFSU 31-ENE-2028")

D0,D1=D(2026,5,1),D(2028,3,31)
MES=['','Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
iso=lambda d: d.strftime("%Y-%m-%d")

# ================================================================= HTML
def build_html():
    src=open(BASE,encoding="utf-8").read()
    # --- colores nuevos en :root y en el bloque dark ---
    src=src.replace("  --abbs:#C62828; --abbp:#EF6C00; --inauco:#1565C0; --hima:#6A1B9A; --epc:#2E7D32;",
      "  --abbs:#C62828; --abbd:#795548; --abbp:#EF6C00; --miron:#33691E; --inauco:#1565C0;\n"
      "  --hph:#6A1B9A; --harg:#00838F; --epc:#2E7D32; --hima:#6A1B9A;")
    # --- cabecera ---
    src=re.sub(r"<h1>Cronograma Integrado.*?</h1>",
      f'<h1>Cronograma Integrado — CPF-2 La Calera II · {REV} ({FECHA[8:]}-{FECHA[5:7]}-{FECHA[:4]}) '
      f'<span style="color:var(--muted);font-weight:400;font-size:15px">(Vaca Muerta)</span></h1>',src,flags=re.S)
    src=re.sub(r'<div class="sub">.*?</div>\n</header>', f'<div class="sub">{SUB}</div>\n</header>', src, flags=re.S)
    # --- PROV ---
    provjs="const PROV={\n"+ "".join(
      f' {k:6}:{{n:"{nm}",c:"var({var})",hex:"{hx}"}},\n' for (k,_,nm,var,hx) in PROV)+"};"
    src=re.sub(r"const PROV=\{.*?\n\};", provjs, src, flags=re.S)
    # --- T ---
    tjs="const T=[\n"+"".join(
      ' ["%s","%s","%s",%s%s],\n'%(k,t.replace('"','\\"'),iso(s),
            (f'"{iso(e)}"' if e else "null"), (f',"{m}"' if m else ""))
      for (k,t,s,e,m) in T)+"];"
    src=re.sub(r"const T=\[.*?\n\];", tjs, src, flags=re.S)
    # --- PH ---
    phjs="const PH=[\n"+"".join(
      ' ["%s","%s","%s",%s,%d%s],\n'%(p[0],p[1].replace('"','\\"'),iso(p[2]),
            (f'"{iso(p[3])}"' if p[3] else "null"), p[4], (f',"{p[5]}"' if len(p)>5 else ""))
      for p in PH)+"];"
    src=re.sub(r"const PH=\[.*?\n\];", phjs, src, flags=re.S)
    src=re.sub(r'const ORDER=\[.*?\];', "const ORDER=["+",".join(f'"{k}"' for (k,_,_,_,_) in PROV)+"];", src)
    # --- ventana temporal y KPIs ---
    src=src.replace('const D0=new Date("2026-05-01"), D1=new Date("2028-02-29");',
                    f'const D0=new Date("{iso(D0)}"), D1=new Date("{iso(D1)}");')
    kpijs="document.getElementById(\"kpis\").innerHTML=[\n"+"".join(
      f' ["{a}","{b}","var({c})"],\n' for (a,b,c) in KPIS)+"]"
    # el bloque termina en un ']' a principio de línea, antes de .map(
    src=re.sub(r'document\.getElementById\("kpis"\)\.innerHTML=\[\n(?:.*?\n)*?\](?=\.map)', kpijs, src, count=1)
    # --- leyendas ---
    src=src.replace("◆/★ = hito. Líneas de las provisiones filtradas se ocultan. Fechas de referencia (sujetas a OC y freezing points).",
      "◆/★ = hito. Usá los chips para aislar una provisión. Fechas según los cronogramas vigentes de cada proveedor al 27-10-2026 "
      "(ABB Salas y Ductos REV 0 «Para construcción» 14-10 · ABB PMS SEP-2026 REV 0 · Inauco REV 0 «Para construcción» 23-10 · "
      "HPH informe Nº1 15-10 · CAT Miron 29-09 · HARG con OC pendiente).")
    src=re.sub(r'Banda gris = <b>período FAT</b>.*?convergencia hacia iFAT → RFSU\.',
      "Banda gris = <b>período FAT</b> (FAT PCS 29-MAR → cierre HARG 10-MAY-2027). Banda verde = <b>iFAT en sitio</b>. "
      "Rojo = ruta crítica: OC de HARG → HPH → adecuación → FAT SIS → cierre HARG. El <b>ducto de barras</b> corre por fuera "
      "y es la última entrega del paquete (08-JUL-2027).", src, flags=re.S)
    # el carril del flujo necesita más ancho: nombres y barras nuevas
    src=src.replace("const VW=1180,LW=150,PX=VW-LW-16,AX=56,ROW=22,LPAD=8;",
                    "const VW=1560,LW=205,PX=VW-LW-16,AX=56,ROW=24,LPAD=10;")
    # bandas: el período FAT cierra con HARG y el iFAT se corre a agosto
    src=src.replace('s+=band("2027-03-29","2027-05-04","#607080",.15,"período FAT");',
                    's+=band("2027-03-29","2027-05-10","#607080",.15,"período FAT");')
    src=src.replace('s+=band("2027-06-25","2027-07-14","#2E7D32",.14,"iFAT");',
                    's+=band("2027-08-02","2027-08-31","#2E7D32",.14,"iFAT en sitio");')
    # con el lienzo más ancho entran más rótulos dentro de la barra
    src=src.replace("if(bw>96)s+=","if(bw>esc(r[1]).length*4.7+12)s+=")
    # el empaquetado por filas debe reservar el ancho del RÓTULO, no sólo el de la barra:
    # si no, dos barras cortas consecutivas caen en la misma fila y sus textos se pisan
    src=src.replace(
      "      const s=new Date(r[2]), e=r[3]?new Date(r[3]):new Date(new Date(r[2]).getTime()+10*864e5);",
      "      const s=new Date(r[2]);\n"
      "      let e=r[3]?new Date(r[3]):new Date(new Date(r[2]).getTime()+2*864e5);\n"
      "      const _bw=r[3]?PX*((new Date(r[3]))-s)/SPAN:0, _lab=String(r[1]).length*4.9+22;\n"
      "      if(_bw<_lab) e=new Date(e.getTime()+(_lab+10)*SPAN/PX);")
    src=src.replace('style="min-width:1000px','style="min-width:1320px')
    src=src.replace("cpf2_eval_v1","cpf2_eval_v3")
    path=os.path.join(OUT,f"Cronograma_Integrado_CPF2_{REV}_{FECHA}.html")
    open(path,"w",encoding="utf-8").write(src); return path

# ================================================================= PDFs
def _pdf_common():
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib import colors
    from reportlab.pdfgen import canvas
    return A3, landscape, colors, canvas

def build_gantt():
    A3,landscape,colors,canvas=_pdf_common()
    W,H=landscape(A3); path=os.path.join(OUT,f"Cronograma_Integrado_CPF2_Gantt_{REV}_{FECHA}.pdf")
    c=canvas.Canvas(path,pagesize=(W,H)); GREY=colors.HexColor('#555555')
    months=[]; y,m=D0.year,D0.month
    while (y<D1.year) or (y==D1.year and m<=D1.month):
        months.append((y,m)); m+=1
        if m>12: m=1;y+=1
    x0=330; x1=W-24; top=H-92; span=(D1-D0).days
    X=lambda dt: x0+(x1-x0)*((dt-D0).days/span)
    SHORT={k:s for (k,s,_,_,_) in PROV}
    c.setFont("Helvetica-Bold",15); c.setFillColor(colors.black)
    c.drawString(24,H-34,f"CRONOGRAMA INTEGRADO — CPF-2 La Calera II   ·   {REV} · {FECHA}")
    c.setFont("Helvetica",8.6); c.setFillColor(GREY)
    c.drawString(24,H-48,"8 provisiones · ABB Salas y Ductos REV 0 «Para construcción» (14-10-26) · ABB PMS SEP-2026 REV 0 · "
                         "CAT Miron (29-09-26) · Inauco REV 0 «Para construcción» (23-10-26) · HPH informe Nº1 (15-10-26) · HARG con OC pendiente")
    c.drawString(24,H-60,"Camino crítico de sistemas: cierre HARG 10-MAY-2027  ·  Última entrega del paquete: ductos de barras 08-JUL-2027  ·  RFSU 31-ENE-2028")
    yr=top+20; c.setFont("Helvetica-Bold",8); prev=None
    for (yy,mm) in months:
        mx=X(D(yy,mm,1)); mxn=X(D(yy+(mm//12),(mm%12)+1,1))
        c.setStrokeColor(colors.HexColor('#DDDDDD')); c.setLineWidth(0.5); c.line(mx,top+8,mx,42)
        c.setFillColor(GREY); c.setFont("Helvetica",6.4); c.drawCentredString((mx+mxn)/2,top+2,MES[mm])
        if yy!=prev: c.setFillColor(colors.black); c.setFont("Helvetica-Bold",9); c.drawString(mx+2,yr,str(yy)); prev=yy
    for dt,lab,hx in VLINES:
        vx=X(dt); col=colors.HexColor(hx)
        c.setStrokeColor(col); c.setLineWidth(1); c.setDash(2,2); c.line(vx,42,vx,top+8); c.setDash()
        c.setFillColor(col); c.setFont("Helvetica-Bold",6.2)
        c.saveState(); c.translate(vx,top+28); c.rotate(90); c.drawString(0,2,lab); c.restoreState()
    rowh=(top-52)/len(T); yy=top
    from math import cos,sin,pi
    def star(cx,cy,r,col):
        p=c.beginPath()
        for i in range(10):
            a=pi/2+i*pi/5; rr=r if i%2==0 else r*0.45
            (p.moveTo if i==0 else p.lineTo)(cx+rr*cos(a),cy+rr*sin(a))
        p.close(); c.setFillColor(col); c.drawPath(p,fill=1,stroke=0)
    prevk=None
    for (k,label,s,e,mark) in T:
        yy-=rowh; col=colors.HexColor(PHEX[k])
        if k!=prevk:
            c.setFillColor(col); c.rect(6,yy,10,rowh-2,fill=1,stroke=0)
            c.setFillColor(colors.black); c.setFont("Helvetica-Bold",6.8); c.drawString(20,yy+rowh/2-2.4,SHORT[k][:18])
            prevk=k
        c.setFillColor(colors.HexColor('#333333')); c.setFont("Helvetica",5.9); c.drawString(112,yy+rowh/2-2.2,label[:62])
        c.setStrokeColor(colors.HexColor('#EEEEEE')); c.setLineWidth(0.3); c.line(x0,yy,x1,yy)
        if e is None: star(X(s),yy+rowh/2,3.6,col)
        else:
            bx=X(s); bw=max(X(e)-X(s),2.0); c.setFillColor(col); c.roundRect(bx,yy+1.6,bw,rowh-4,1.6,fill=1,stroke=0)
    c.setStrokeColor(colors.HexColor('#BBBBBB')); c.setLineWidth(0.6); c.rect(x0,44,x1-x0,top-44,fill=0,stroke=1)
    c.setFont("Helvetica-Bold",6.6); ly=30; lx=24
    for (k,sn,_,_,hx) in PROV:
        c.setFillColor(colors.HexColor(hx)); c.rect(lx,ly,9,7,fill=1,stroke=0)
        c.setFillColor(colors.black); c.drawString(lx+11,ly+1,sn); lx+=112
    c.setFont("Helvetica-Oblique",6.2); c.setFillColor(GREY)
    c.drawRightString(x1,20,"Fechas de los cronogramas vigentes de cada proveedor al 27-10-2026. El ducto de barras tiene el plazo SIN CONFIRMAR y la OC de HARG está PENDIENTE.")
    c.showPage(); c.save(); return path

def build_flow():
    A3,landscape,colors,canvas=_pdf_common()
    W,H=landscape(A3); path=os.path.join(OUT,f"Cronograma_Integrado_CPF2_Flujo_{REV}_{FECHA}.pdf")
    c=canvas.Canvas(path,pagesize=(W,H)); GREY=colors.HexColor('#555555')
    c.setFont("Helvetica-Bold",15); c.setFillColor(colors.black)
    c.drawString(24,H-34,f"FLUJO INTEGRADO DE PROVISIONES — CPF-2 La Calera II · {REV} · {FECHA}")
    c.setFont("Helvetica",8.6); c.setFillColor(GREY)
    c.drawString(24,H-48,"Convergencia de las 8 provisiones hacia el montaje en planta → iFAT en sitio → comisionado → RFSU. En rojo, el camino crítico de sistemas.")
    # escala temporal
    x0=170; x1=W-150; top=H-78; span=(D1-D0).days
    X=lambda dt: x0+(x1-x0)*((dt-D0).days/span)
    months=[]; y,m=D0.year,D0.month
    while (y<D1.year) or (y==D1.year and m<=D1.month):
        months.append((y,m)); m+=1
        if m>12: m=1;y+=1
    c.setFont("Helvetica",6); prev=None
    for (yy,mm) in months:
        mx=X(D(yy,mm,1))
        c.setStrokeColor(colors.HexColor('#E8E8E8')); c.setLineWidth(0.4); c.line(mx,top-6,mx,60)
        c.setFillColor(GREY); c.drawCentredString(mx+6,top-2,MES[mm][0])
        if yy!=prev: c.setFillColor(colors.black); c.setFont("Helvetica-Bold",7.5); c.drawString(mx+1,top+8,str(yy)); c.setFont("Helvetica",6); prev=yy
    # banda FAT y banda iFAT
    fx0,fx1=X(D(2027,3,29)),X(D(2027,5,10))
    c.setFillColor(colors.HexColor('#ECEFF1')); c.rect(fx0,62,fx1-fx0,top-70,fill=1,stroke=0)
    ix0,ix1=X(D(2027,8,2)),X(D(2027,8,31))
    c.setFillColor(colors.HexColor('#E8F5E9')); c.rect(ix0,62,ix1-ix0,top-70,fill=1,stroke=0)
    c.setFillColor(GREY); c.setFont("Helvetica-Bold",6)
    c.drawCentredString((fx0+fx1)/2,top-16,"PERÍODO FAT"); c.drawCentredString((ix0+ix1)/2,top-16,"iFAT")
    rowh=(top-86)/len(PH); yy=top-22; prevk=None
    for ph in PH:
        k,txt,s,e,crit=ph[0],ph[1],ph[2],ph[3],ph[4]
        yy-=rowh; col=colors.HexColor(PHEX[k])
        if k!=prevk:
            c.setFillColor(col); c.setFont("Helvetica-Bold",7.4)
            c.drawString(24,yy+rowh/2-2.4,[s2 for (k2,s2,_,_,_) in PROV if k2==k][0][:13]); prevk=k
        c.setFillColor(colors.HexColor('#333333')); c.setFont("Helvetica",6.1); c.drawRightString(x0-6,yy+rowh/2-2.2,txt[:46])
        if e is None:
            from math import cos,sin,pi
            p=c.beginPath(); cx,cy,r=X(s),yy+rowh/2,4.0
            for i in range(10):
                a=pi/2+i*pi/5; rr=r if i%2==0 else r*0.45
                (p.moveTo if i==0 else p.lineTo)(cx+rr*cos(a),cy+rr*sin(a))
            p.close(); c.setFillColor(col); c.drawPath(p,fill=1,stroke=0)
        else:
            bx=X(s); bw=max(X(e)-X(s),2.0)
            c.setFillColor(col); c.roundRect(bx,yy+2,bw,rowh-5,2,fill=1,stroke=0)
            if crit:
                c.setStrokeColor(colors.HexColor('#C62828')); c.setLineWidth(1.1); c.roundRect(bx,yy+2,bw,rowh-5,2,fill=0,stroke=1)
    c.setStrokeColor(colors.HexColor('#BBBBBB')); c.setLineWidth(0.6); c.rect(x0,60,x1-x0,top-22-60,fill=0,stroke=1)
    c.setFont("Helvetica-Oblique",6.6); c.setFillColor(GREY)
    c.drawString(24,46,"Ruta crítica de sistemas: OC de HARG (pendiente) → HPH (fabricación completada, DAP Ezeiza 15-DIC) → nacionalización → adecuación en base Inauco → FAT SIS → cierre HARG 10-MAY-2027.")
    c.drawString(24,34,"Los ductos de barras corren por fuera del camino crítico de sistemas pero son la ÚLTIMA ENTREGA del paquete (08-JUL-2027) y condicionan el montaje de la vinculación a transformadores.")
    c.drawString(24,22,"Los 13 transformadores llegan entre nov y dic de 2026 y quedan 7 meses almacenados a la espera de los ductos.")
    c.showPage(); c.save(); return path

def build_xlsx():
    import openpyxl
    from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
    path=os.path.join(OUT,f"Cronograma_Integrado_CPF2_{REV}_{FECHA}.xlsx")
    thin=Side(style="thin",color="D9D9D9"); B=Border(thin,thin,thin,thin)
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Cronograma Integrado"
    ws["A1"]=f"CRONOGRAMA INTEGRADO CPF-2 — {REV} · {FECHA}"; ws.merge_cells("A1:E1")
    ws["A1"].font=Font(bold=True,size=12,color="FFFFFF"); ws["A1"].fill=PatternFill("solid",fgColor="1F3864")
    ws["A2"]="8 provisiones · fechas de los cronogramas vigentes de cada proveedor al 27-10-2026"; ws.merge_cells("A2:E2")
    ws["A2"].font=Font(italic=True,size=9,color="555555")
    NAMES={k:n for (k,_,n,_,_) in PROV}
    for j,h in enumerate(["Provisión","Actividad / Hito","Inicio","Fin","Tipo"]):
        cc=ws.cell(4,1+j,h); cc.font=Font(bold=True,color="FFFFFF"); cc.fill=PatternFill("solid",fgColor="37474F")
        cc.alignment=Alignment(horizontal="center"); cc.border=B
    r=5
    for (k,label,s,e,mark) in T:
        ws.cell(r,1,NAMES[k]); ws.cell(r,2,label)
        ws.cell(r,3,s.strftime("%d-%m-%Y")); ws.cell(r,4,e.strftime("%d-%m-%Y") if e else "—")
        ws.cell(r,5,"Hito" if e is None else "Tarea")
        ws.cell(r,1).fill=PatternFill("solid",fgColor=PHEX[k][1:]); ws.cell(r,1).font=Font(color="FFFFFF",bold=True,size=8)
        if e is None:
            for j in range(2,6): ws.cell(r,j).font=Font(bold=True,size=9)
        for j in range(1,6): ws.cell(r,j).border=B
        r+=1
    for j,w in zip("ABCDE",[32,62,12,12,8]): ws.column_dimensions[j].width=w
    ws.freeze_panes="A5"; ws.auto_filter.ref=f"A4:E{ws.max_row}"
    wb.save(path); return path

h=build_html(); g=build_gantt(); f=build_flow(); x=build_xlsx()
print(f"tareas={len(T)} · fases={len(PH)} · provisiones={len(PROV)}")
print("OK\n ",h,"\n ",g,"\n ",f,"\n ",x)
