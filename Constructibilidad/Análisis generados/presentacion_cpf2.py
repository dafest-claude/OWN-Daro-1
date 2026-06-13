#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Presentacion ejecutiva - Constructibilidad CPF2 La Calera II
Instrumentacion & Electricidad: hitos, volumenes, sistemas, secuencia y
precomisionado. Sintesis de los analisis previos. RFSU 20-DIC-2027.
Salida: Presentacion_CPF2_INS_EL.pptx
"""
import os
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
THREED = os.path.join(ROOT, "Datos entrada", "3D LC 120626.pdf")
OUT = os.path.join(BASE, "Presentacion_CPF2_INS_EL.pptx")

# Paleta
NAVY=RGBColor(0x1F,0x3B,0x63); IN=RGBColor(0x2E,0x75,0xB6); EL=RGBColor(0xC5,0x5A,0x11)
GREEN=RGBColor(0x54,0x82,0x35); GOLD=RGBColor(0xBF,0x90,0x00); RED=RGBColor(0xC0,0x00,0x00)
GREY=RGBColor(0x59,0x59,0x59); LGREY=RGBColor(0xF2,0xF2,0xF2); WHITE=RGBColor(0xFF,0xFF,0xFF)
LBLUE=RGBColor(0xDC,0xE6,0xF2)

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
SW,SH=prs.slide_width,prs.slide_height
BLANK=prs.slide_layouts[6]

def slide(): return prs.slides.add_slide(BLANK)
def rect(s,x,y,w,h,color,line=None):
    sp=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,x,y,w,h)
    sp.fill.solid(); sp.fill.fore_color.rgb=color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb=line; sp.line.width=Pt(0.75)
    sp.shadow.inherit=False
    return sp
def txt(s,x,y,w,h,text,size=18,color=NAVY,bold=False,align=PP_ALIGN.LEFT,
        anchor=MSO_ANCHOR.TOP,font='Calibri'):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True
    tf.vertical_anchor=anchor; tf.margin_left=Pt(2); tf.margin_right=Pt(2)
    p=tf.paragraphs[0]; p.alignment=align
    for i,line in enumerate(text.split('\n')):
        par=p if i==0 else tf.add_paragraph(); par.alignment=align
        r=par.add_run(); r.text=line; f=r.font
        f.size=Pt(size); f.bold=bold; f.color.rgb=color; f.name=font
    return tb
def header(s,title,accent=IN,kicker="CPF2 LA CALERA II · CONSTRUCTIBILIDAD"):
    rect(s,0,0,SW,Inches(1.12),NAVY)
    rect(s,0,Inches(1.12),SW,Inches(0.06),accent)
    txt(s,Inches(0.45),Inches(0.12),Inches(11),Inches(0.3),kicker,11,RGBColor(0x9D,0xB8,0xD8),True)
    txt(s,Inches(0.45),Inches(0.40),Inches(12.4),Inches(0.7),title,26,WHITE,True)
    txt(s,Inches(11.0),Inches(0.12),Inches(2.1),Inches(0.3),"RFSU 20-DIC-2027",11,GOLD,True,PP_ALIGN.RIGHT)
def footer(s,n):
    txt(s,Inches(0.45),Inches(7.08),Inches(9),Inches(0.3),
        "Constructibilidad CPF2 · Instrumentación & Electricidad · Rev. 0 · 13-06-2026",9,GREY)
    txt(s,Inches(12.2),Inches(7.08),Inches(0.9),Inches(0.3),str(n),9,GREY,align=PP_ALIGN.RIGHT)

def table(s,x,y,w,rows,colw,header_fill=NAVY,fs=12,hfs=12,rowh=Inches(0.34)):
    nr=len(rows); nc=len(rows[0])
    gf=s.shapes.add_table(nr,nc,x,y,w,rowh*nr); tb=gf.table
    tb.first_row=False; tb.horz_banding=False
    tot=sum(colw)
    for j,cw in enumerate(colw): tb.columns[j].width=Emu(int(w*cw/tot))
    for i,row in enumerate(rows):
        tb.rows[i].height=rowh
        for j,val in enumerate(row):
            cell=tb.cell(i,j); cell.margin_left=Pt(5); cell.margin_right=Pt(4)
            cell.margin_top=Pt(1); cell.margin_bottom=Pt(1)
            cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            tfc=cell.text_frame; tfc.word_wrap=True
            p=tfc.paragraphs[0]; r=p.add_run(); r.text=str(val); f=r.font; f.name='Calibri'
            if i==0:
                cell.fill.solid(); cell.fill.fore_color.rgb=header_fill
                f.size=Pt(hfs); f.bold=True; f.color.rgb=WHITE
            else:
                cell.fill.solid(); cell.fill.fore_color.rgb=WHITE if i%2 else LGREY
                f.size=Pt(fs); f.color.rgb=RGBColor(0x22,0x22,0x22)
                if j>0: p.alignment=PP_ALIGN.LEFT
    return tb

def card(s,x,y,w,h,big,small,accent):
    rect(s,x,y,w,h,WHITE,line=RGBColor(0xCC,0xCC,0xCC))
    rect(s,x,y,w,Inches(0.10),accent)
    txt(s,x,y+Inches(0.18),w,Inches(0.62),big,30,accent,True,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
    txt(s,x,y+h-Inches(0.62),w,Inches(0.55),small,11.5,GREY,False,PP_ALIGN.CENTER,MSO_ANCHOR.TOP)

def bullets(s,x,y,w,h,items,size=14,color=RGBColor(0x22,0x22,0x22),gap=6):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True
    for i,(lvl,t,c) in enumerate(items):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph()
        p.space_after=Pt(gap); p.level=lvl
        r=p.add_run(); r.text=("• " if lvl==0 else "– ")+t
        r.font.size=Pt(size- (1 if lvl else 0)); r.font.color.rgb=c or color; r.font.name='Calibri'
        r.font.bold=(lvl==0 and c is not None)
    return tb

# render 3D thumbnail
THUMB=None
try:
    import fitz; fitz.TOOLS.mupdf_display_errors(False)
    THUMB=os.path.join(BASE,"_3d_thumb.png")
    fitz.open(THREED)[0].get_pixmap(dpi=120).save(THUMB)
except Exception:
    THUMB=None

# ============ SLIDE 1 — PORTADA ============
s=slide(); rect(s,0,0,SW,SH,NAVY)
rect(s,0,Inches(4.55),SW,Inches(0.07),GOLD)
txt(s,Inches(0.7),Inches(1.5),Inches(12),Inches(0.4),"CPF2 LA CALERA II · VACA MUERTA",16,RGBColor(0x9D,0xB8,0xD8),True)
txt(s,Inches(0.7),Inches(2.0),Inches(12),Inches(1.6),
    "Plan de Construcción\nInstrumentación & Electricidad",40,WHITE,True)
txt(s,Inches(0.7),Inches(4.7),Inches(12),Inches(0.9),
    "Constructibilidad · Hitos, volúmenes, sistemas, secuencia y precomisionado",18,LBLUE)
txt(s,Inches(0.7),Inches(5.4),Inches(12),Inches(0.5),"Hito gobernante: RFSU 20-DIC-2027",18,GOLD,True)
txt(s,Inches(0.7),Inches(6.7),Inches(8),Inches(0.4),"Rev. 0 · 13-06-2026 · Documento preliminar para depuración",11,RGBColor(0x9D,0xB8,0xD8))
if THUMB:
    s.shapes.add_picture(THUMB,Inches(9.0),Inches(5.7),height=Inches(1.5))

# ============ SLIDE 2 — CIFRAS CLAVE ============
s=slide(); header(s,"Cifras clave del proyecto",IN)
cy=Inches(1.7); cw=Inches(4.0); ch=Inches(1.7); gx=Inches(0.30); x0=Inches(0.45)
data=[("81.231 m","Cable eléctrico (513 cables)",EL),
      ("65.315 m","Cable instrumentación (885 cables)",IN),
      ("~20.930","Puntas a interconectar (EL+IN)",NAVY),
      ("~1.163","Instrumentos a montar AESA (de 2.798)",GREEN),
      ("243","Cargas eléctricas (79 motores · 6 VFD)",EL),
      ("20-DIC-27","RFSU — Ready For Start Up",RED)]
for i,(b,sm,ac) in enumerate(data):
    r,c=divmod(i,3); x=x0+c*(cw+gx); y=cy+r*(ch+Inches(0.3))
    card(s,x,y,cw,ch,b,sm,ac)
txt(s,Inches(0.45),Inches(6.4),Inches(12.4),Inches(0.5),
    "Alcance: ampliación CPF2 — sistemas PCS · PSS · ESD · F&G · SCADA · PMS ABB. AESA monta instrumentos propios + los de proveedor marcados 'montaje por EPC'.",12,GREY)
footer(s,2)

# ============ SLIDE 3 — HITOS CLAVE ============
s=slide(); header(s,"Hitos clave del proyecto",GOLD)
rows=[["Hito","Fecha","Nota / restricción"],
 ["OC Cables IN (target)","01-AGO-26","LT 150 d → sitio 28-NOV-26"],
 ["OC Cables EL (target)","01-OCT-26","LT 150 d → sitio 28-FEB-27"],
 ["Llegada cables IN a sitio","28-NOV-26","Inicio canalizaciones IN"],
 ["Inicio campo PCS — Sala INS (Sala 7)","17-MAR-27","⛔ Cables IN tendidos ANTES"],
 ["Llegada Sala Eléctrica SE#4 (PMS-101)","08-ABR-27","Ancla EL · armado 14 d"],
 ["Llegada Sala Eléctrica SE#3 (PMS-001)","14-MAY-27","Ancla EL · ruta crítica · armado 21 d"],
 ["Megger EL (HOLD POINT)","26-JUL→15-AGO-27","Circuito por circuito"],
 ["Energización progresiva MT→BT","16-AGO→01-SEP-27","HOLD POINT"],
 ["SAT PMS (HOLD POINT contractual)","06-OCT→05-DIC-27","Ventana fija OCT-DIC"],
 ["★ RFSU — Ready For Start Up","20-DIC-27","Hito gobernante"]]
table(s,Inches(0.45),Inches(1.45),Inches(12.45),rows,[5,2.3,4.2],fs=12.5,rowh=Inches(0.45))
footer(s,3)

# ============ SLIDE 4 — SALAS ELÉCTRICAS Y SALA INS ============
s=slide(); header(s,"Entrega de salas eléctricas y Sala INS",EL)
rows=[["Sala / Sistema","Llegada / inicio","Armado en campo","Conexionado desde"],
 ["SE#4 (con PMS-101)","08-ABR-2027","14 días (08→22-ABR)","23-ABR-2027"],
 ["SE#3 (con PMS-001) ◄ crítica","14-MAY-2027","21 días (14-MAY→03-JUN)","05-JUN-2027"],
 ["Sala INS — PCS (Sala 7)","17-MAR-2027","Instalación PCS Inauco","17-MAR-2027 (con cables IN)"]]
table(s,Inches(0.45),Inches(1.5),Inches(12.45),rows,[3.6,2.6,3.2,3.2],fs=12.5,rowh=Inches(0.5))
bullets(s,Inches(0.45),Inches(3.7),Inches(12.4),Inches(3),[
 (0,"Las salas SE#3 y SE#4 llegan en MÓDULOS → requieren armado en campo antes del conexionado.",NAVY),
 (0,"PIVOTE clave:",EL),
 (1,"Parte del tendido de cables EL se hace ANTES de cerrar las salas, para conectar sala → motores/equipos.",None),
 (1,"El conexionado pivota sobre el armado de cada shelter (SE#4 desde 23-ABR; SE#3 desde 05-JUN).",None),
 (0,"Sala INS (control):",IN),
 (1,"Cables IN deben estar TENDIDOS antes del 17-MAR para instalar el PCS; control (PCS) y seguridad (ESD/SIS) se montan y comisionan allí.",None),
])
footer(s,4)

# ============ SLIDE 5 — CABLES A TENDER ============
s=slide(); header(s,"Cables a tender y puntas a interconectar",EL)
rows=[["Especialidad / tipo","Cables","Metros","Puntas / nota"],
 ["EL — Potencia MT (6,6/13,2 kV)","18","2.110","Técnicos cert. ABB"],
 ["EL — Potencia BT Grande ≥35 mm² ◄ crítica","104","17.470","Ruta crítica EL"],
 ["EL — Potencia BT Med/Peq","198","33.125",""],
 ["EL — Control y Señales","192","28.126","Multiconductores"],
 ["EL — FO / Ethernet","1","400","Red PMS SE3-SE4"],
 ["SUBTOTAL ELÉCTRICO","513","81.231","~5.178 puntas (3.036 sala / 1.717 campo)"],
 ["IN — Instrumentación PCS/SIS/F&G","885","65.315","15.752 puntas"],
 ["TOTAL EL + IN","1.398","146.546","~20.930 puntas"]]
table(s,Inches(0.45),Inches(1.5),Inches(12.45),rows,[5.2,1.4,1.8,4.0],fs=12,rowh=Inches(0.46))
txt(s,Inches(0.45),Inches(6.45),Inches(12.4),Inches(0.5),
    "Tendido IN completo ANTES del 17-MAR (PCS). Tendido EL completo ANTES de llegada de salas (SE#3 14-MAY). Pico de tendido FEB-MAR-2027.",12,RED,True)
footer(s,5)

# ============ SLIDE 6 — INSTRUMENTOS A MONTAR ============
s=slide(); header(s,"Instrumentos a montar (alcance AESA)",GREEN)
card(s,Inches(0.45),Inches(1.55),Inches(2.9),Inches(1.5),"2.798","Instrumentos totales",NAVY)
card(s,Inches(3.55),Inches(1.55),Inches(2.9),Inches(1.5),"~1.163","Monta AESA (897 + 266 EPC)",GREEN)
card(s,Inches(6.65),Inches(1.55),Inches(2.9),Inches(1.5),"920","En línea (Fase 1)",IN)
card(s,Inches(9.75),Inches(1.55),Inches(2.9),Inches(1.5),"1.878","Montaje específico (Fase 2)",EL)
bullets(s,Inches(0.45),Inches(3.35),Inches(7.2),Inches(3.4),[
 (0,"Entrega PROGRESIVA de instrumentos:",GREEN),
 (1,"Fase 1 — válvulas y todo instrumento en línea (920): se entregan y montan PRIMERO, junto con la cañería.",None),
 (1,"Fase 2 — resto de montaje específico (1.878): entrega progresiva desde FEB-2027.",None),
 (0,"Prueba en BANCO (calibración) previa al montaje de cada instrumento.",NAVY),
 (0,"Montaje de campo AESA: ENE→JUN-2027.",NAVY),
],size=13.5)
rows=[["Típico de montaje","Cant."],
 ["En línea — válvula","485"],["En línea — elemento","435"],
 ["Soporte junto a línea (stand)","610"],["Sobre equipo","640"],
 ["Sobre válvula (accesorio)","483"],["Otro instrumento / F&G","145"]]
table(s,Inches(8.0),Inches(3.35),Inches(4.9),rows,[3.6,1.1],fs=11.5,hfs=11.5,rowh=Inches(0.36))
footer(s,6)

# ============ SLIDE 7 — SISTEMAS A INSTALAR Y AMPLIAR ============
s=slide(); header(s,"Sistemas a instalar y ampliar",IN)
rows=[["Sistema","Descripción","Proveedor","Hito / ancla"],
 ["PCS","Sistema de Control de Procesos (DCS) — Sala INS / Sala 7","Inauco","Inicio campo 17-MAR-27 · SAT 55 d"],
 ["PSS","Sistema de seguridad de procesos (parada de planta)","Inauco / HIMA","Integrado con ESD/SIS"],
 ["ESD","Parada de emergencia (Emergency Shutdown / SIS)","HIMA","SAT SIS 100 d (17-MAR→03-AGO)"],
 ["F&G","Fuego y Gas — 120 dispositivos de campo (AESA)","AESA / Inauco","Detección + matriz C&E"],
 ["SCADA","Supervisión y adquisición de datos","Inauco","Prueba Matriz C&E AGO-27"],
 ["PMS","Power Management System — PMS-001 (SE#3) + PMS-101 (SE#4)","ABB","SAT PMS HOLD POINT OCT-DIC"]]
table(s,Inches(0.45),Inches(1.5),Inches(12.45),rows,[1.2,5.0,2.0,4.3],fs=12,hfs=12.5,rowh=Inches(0.62))
txt(s,Inches(0.45),Inches(6.4),Inches(12.4),Inches(0.5),
    "iFAT — integración SE#3+SE#4+PMS+PCS+SIS: 25-JUN→14-JUL-2027 (todos los vendors).",12,GREY)
footer(s,7)

# ============ SLIDE 8 — SECUENCIA DE TRABAJO ============
s=slide(); header(s,"Secuencia de trabajo (fechas tentativas)",NAVY)
bullets(s,Inches(0.45),Inches(1.45),Inches(12.5),Inches(5.4),[
 (0,"1 · INSTRUMENTACIÓN (IN) — NOV-2026 → AGO-2027",IN),
 (1,"Canalizaciones (NOV-26→ENE-27) → Tendido cables IN (FEB→14-MAR, antes PCS) → Conexionado campo→JB→Sala INS (17-MAR→31-JUL) → Loop check (JUN→AGO).",None),
 (1,"Montaje instrumentos: Fase 1 en línea (ENE→ABR) · Fase 2 específico (MAR→JUN), con banco previo.",None),
 (0,"2 · ELECTRICIDAD (EL) — MAR-2027 → JUL-2027",EL),
 (1,"Bandejas (MAR→ABR) → Tendido EL por tipo (MAR→MAY, BT Grande crítico) → Armado shelters SE#4/SE#3 → Conexionado SE#4 (23-ABR→12-JUN) y SE#3 (05-JUN→25-JUL).",None),
 (0,"3 · PRUEBAS Y ENERGIZACIÓN — JUL-2027 → SEP-2027",GREEN),
 (1,"Megger EL (26-JUL→15-AGO) → Energización progresiva MT→BT (16-AGO→01-SEP).",None),
 (0,"4 · PRECOMISIONADO Y COMISIONADO — SEP-2027 → DIC-2027",RED),
 (1,"Precom EL e IN (SEP) → Comisionado integrado EL+PCS+PMS+SIS (23-SEP→19-DIC) → SAT PMS (OCT-DIC) → RFSU 20-DIC.",None),
],size=13.5,gap=7)
footer(s,8)

# ============ SLIDE 9 — PRECOMISIONADO ============
s=slide(); header(s,"Precomisionado: arranque y completamiento",GREEN)
bullets(s,Inches(0.45),Inches(1.5),Inches(12.5),Inches(2.4),[
 (0,"¿CUÁNDO ARRANCA?",NAVY),
 (1,"Instrumentación: prueba en BANCO (calibración) desde DIC-2026, previa al montaje. Loop check de lazos (sensor→JB→PCS/SIS) desde JUN-2027, en paralelo con conexionado IN.",None),
 (1,"Eléctrico: precomisionado funcional arranca tras la energización, el 02-SEP-2027 (loop check eléctrico, ajuste de relés, secuencia de arranque de motores).",None),
],size=13)
bullets(s,Inches(0.45),Inches(3.9),Inches(12.5),Inches(2.9),[
 (0,"¿CÓMO SE COMPLETA?",NAVY),
 (1,"Secuencia EL: Megger (26-JUL→15-AGO) → Energización (16-AGO→01-SEP) → Precom funcional (02→22-SEP).",None),
 (1,"Secuencia IN: Loop check completo PCS/SIS (JUN→AGO) → Precom instrumentación / lazos (SEP).",None),
 (1,"Cierre: Comisionado integrado EL+PCS+PMS+SIS (23-SEP→19-DIC) y SAT PMS (HOLD POINT OCT-DIC) → RFSU 20-DIC-2027.",None),
 (0,"Condición: precom no puede iniciar sin instrumentos montados + chequeados en banco y cables conexionados/megger OK.",RED),
],size=13)
footer(s,9)

# ============ SLIDE 10 — RUTA CRÍTICA Y RIESGOS ============
s=slide(); header(s,"Ruta crítica, restricciones y riesgos",RED)
bullets(s,Inches(0.45),Inches(1.5),Inches(12.5),Inches(2.2),[
 (0,"RUTA CRÍTICA:",RED),
 (1,"SE#3 llega 14-MAY → armado 21 d → conexionado EL 51 d → megger 21 d → energización 17 d → precom 21 d → comisionado 88 d ≈ 19-DIC. Frente EL crítico: BT Grande (104 c / 17.470 m).",None),
],size=13)
bullets(s,Inches(0.45),Inches(3.4),Inches(6.2),Inches(3.4),[
 (0,"RESTRICCIONES DURAS:",NAVY),
 (1,"⛔ Cables IN tendidos antes del 17-MAR (PCS).",None),
 (1,"⛔ Cables EL tendidos antes del 14-MAY (SE#3).",None),
 (1,"Salas en módulos → armado en campo previo.",None),
 (1,"SAT PMS: ventana fija OCT-DIC (no desliza).",None),
],size=12.5)
bullets(s,Inches(6.8),Inches(3.4),Inches(6.1),Inches(3.4),[
 (0,"RIESGOS SENSIBLES:",NAVY),
 (1,"Volumen BT Grande → +1 cuadrilla EL.",None),
 (1,"Cables MT 13,2 kV: técnicos cert. ABB (coordinar 3 meses).",None),
 (1,"Tendido IN: ventana 42 d exige 5-6 cuadrillas.",None),
 (1,"Pico de personal ~40-50 en MAR-2027.",None),
 (1,"OC cables IN ≤01-AGO-26 / EL ≤01-OCT-26.",None),
],size=12.5)
footer(s,10)

prs.save(OUT)
if THUMB and os.path.exists(THUMB):
    try: os.remove(THUMB)
    except OSError: pass
print("OK ->",OUT,"| slides:",len(prs.slides._sldIdLst))
