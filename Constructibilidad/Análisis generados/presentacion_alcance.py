#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Presentacion 2 - Alcance por especialidad - CPF2 La Calera II
Enfoque: visibilizar el alcance de ELECTRICIDAD e INSTRUMENTACION por separado
(sistemas, equipos, cargas, cables, salas, suministro, magnitudes y tiempos),
distincion AESA vs proveedor en instrumentos, precomisionado (banco / campo /
punta-punta) y cronograma reducido. Eje en RFSU 20-DIC-2027.
Salida: Alcance_Especialidades_CPF2.pptx
"""
import os
from datetime import date
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

BASE=os.path.dirname(os.path.abspath(__file__))
OUT=os.path.join(BASE,"Alcance_Especialidades_CPF2.pptx")

NAVY=RGBColor(0x1F,0x3B,0x63); IN=RGBColor(0x2E,0x75,0xB6); EL=RGBColor(0xC5,0x5A,0x11)
GREEN=RGBColor(0x54,0x82,0x35); GOLD=RGBColor(0xBF,0x90,0x00); RED=RGBColor(0xC0,0x00,0x00)
GREY=RGBColor(0x59,0x59,0x59); LGREY=RGBColor(0xF2,0xF2,0xF2); WHITE=RGBColor(0xFF,0xFF,0xFF)
LBLUE=RGBColor(0xDC,0xE6,0xF2); DK=RGBColor(0x22,0x22,0x22)

prs=Presentation(); prs.slide_width=Inches(13.333); prs.slide_height=Inches(7.5)
SW,SH=prs.slide_width,prs.slide_height
BLANK=prs.slide_layouts[6]
def slide(): return prs.slides.add_slide(BLANK)
def rect(s,x,y,w,h,color,line=None):
    sp=s.shapes.add_shape(MSO_SHAPE.RECTANGLE,x,y,w,h); sp.fill.solid()
    sp.fill.fore_color.rgb=color
    if line is None: sp.line.fill.background()
    else: sp.line.color.rgb=line; sp.line.width=Pt(0.75)
    sp.shadow.inherit=False; return sp
def txt(s,x,y,w,h,text,size=18,color=NAVY,bold=False,align=PP_ALIGN.LEFT,anchor=MSO_ANCHOR.TOP):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True
    tf.vertical_anchor=anchor; tf.margin_left=Pt(2); tf.margin_right=Pt(2)
    for i,line in enumerate(text.split('\n')):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.alignment=align
        r=p.add_run(); r.text=line; f=r.font; f.size=Pt(size); f.bold=bold
        f.color.rgb=color; f.name='Calibri'
    return tb
def header(s,title,accent,kick="ALCANCE DE ESPECIALIDADES · CPF2 LA CALERA II"):
    rect(s,0,0,SW,Inches(1.12),NAVY); rect(s,0,Inches(1.12),SW,Inches(0.06),accent)
    txt(s,Inches(0.45),Inches(0.12),Inches(11),Inches(0.3),kick,11,RGBColor(0x9D,0xB8,0xD8),True)
    txt(s,Inches(0.45),Inches(0.40),Inches(12.4),Inches(0.7),title,25,WHITE,True)
    txt(s,Inches(10.9),Inches(0.12),Inches(2.2),Inches(0.3),"RFSU 20-DIC-2027",11,GOLD,True,PP_ALIGN.RIGHT)
def footer(s,n):
    txt(s,Inches(0.45),Inches(7.08),Inches(9),Inches(0.3),
        "Constructibilidad CPF2 · Alcance por especialidad · Rev. 0 · 13-06-2026",9,GREY)
    txt(s,Inches(12.2),Inches(7.08),Inches(0.9),Inches(0.3),str(n),9,GREY,align=PP_ALIGN.RIGHT)
def table(s,x,y,w,rows,colw,hfill=NAVY,fs=11.5,hfs=11.5,rowh=Inches(0.34),accent_rows=None):
    nr=len(rows); nc=len(rows[0]); gf=s.shapes.add_table(nr,nc,x,y,w,rowh*nr); tb=gf.table
    tb.first_row=False; tb.horz_banding=False; tot=sum(colw)
    for j,cw in enumerate(colw): tb.columns[j].width=Emu(int(w*cw/tot))
    for i,row in enumerate(rows):
        tb.rows[i].height=rowh
        for j,val in enumerate(row):
            cell=tb.cell(i,j); cell.margin_left=Pt(5); cell.margin_right=Pt(4)
            cell.margin_top=Pt(1); cell.margin_bottom=Pt(1); cell.vertical_anchor=MSO_ANCHOR.MIDDLE
            p=cell.text_frame.paragraphs[0]; cell.text_frame.word_wrap=True
            r=p.add_run(); r.text=str(val); f=r.font; f.name='Calibri'
            if i==0:
                cell.fill.solid(); cell.fill.fore_color.rgb=hfill
                f.size=Pt(hfs); f.bold=True; f.color.rgb=WHITE
            else:
                cell.fill.solid()
                if accent_rows and i in accent_rows:
                    cell.fill.fore_color.rgb=LBLUE; f.bold=(j==0)
                else:
                    cell.fill.fore_color.rgb=WHITE if i%2 else LGREY
                f.size=Pt(fs); f.color.rgb=DK
            if j>0: p.alignment=PP_ALIGN.CENTER if len(str(val))<14 else PP_ALIGN.LEFT
    return tb
def card(s,x,y,w,h,big,small,accent):
    rect(s,x,y,w,h,WHITE,line=RGBColor(0xCC,0xCC,0xCC)); rect(s,x,y,w,Inches(0.10),accent)
    txt(s,x,y+Inches(0.16),w,Inches(0.58),big,26,accent,True,PP_ALIGN.CENTER,MSO_ANCHOR.MIDDLE)
    txt(s,x,y+h-Inches(0.66),w,Inches(0.6),small,11,GREY,False,PP_ALIGN.CENTER)
def bullets(s,x,y,w,h,items,size=14,gap=6):
    tb=s.shapes.add_textbox(x,y,w,h); tf=tb.text_frame; tf.word_wrap=True
    for i,(lvl,t,c) in enumerate(items):
        p=tf.paragraphs[0] if i==0 else tf.add_paragraph(); p.space_after=Pt(gap); p.level=lvl
        r=p.add_run(); r.text=("• " if lvl==0 else "– ")+t
        r.font.size=Pt(size-(1 if lvl else 0)); r.font.color.rgb=c or DK; r.font.name='Calibri'
        r.font.bold=(lvl==0 and c is not None)
    return tb
def divider(s,name,accent,sub):
    rect(s,0,0,SW,SH,NAVY); rect(s,Inches(0.0),Inches(3.3),SW,Inches(0.08),accent)
    txt(s,Inches(0.7),Inches(2.5),Inches(12),Inches(1.0),name,46,WHITE,True)
    txt(s,Inches(0.7),Inches(3.55),Inches(12),Inches(0.7),sub,18,LBLUE)

# ---- mini gantt ----
MES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
def D(s): y,m,dd=map(int,s.split('-')); return date(y,m,dd)
T0=date(2026,6,1); T1=date(2028,1,15)
def fx(x0,w,dt): return x0+int(w*((dt-T0).days)/((T1-T0).days))
def mini_gantt(s,x,y,w,h,phases):
    span=(T1-T0).days
    # grid meses
    yy=date(T0.year,T0.month,1)
    while yy<=T1:
        gx=fx(x,w,yy)
        ln=rect(s,Emu(gx),y,Emu(900),h,RGBColor(0xEE,0xEE,0xEE))
        if yy.month in (1,7):
            txt(s,Emu(gx-Emu(1200)),y-Inches(0.28),Inches(0.9),Inches(0.25),
                ("%d "%yy.year if yy.month==1 else "")+MES[yy.month-1],8,GREY,True)
        m=yy.month+1; yr=yy.year+(1 if m>12 else 0); m=1 if m>12 else m; yy=date(yr,m,1)
    rh=int((h-Inches(0.1))/len(phases))
    for i,(nm,a,b,col,crit) in enumerate(phases):
        ry=y+Inches(0.05)+Emu(i*rh)
        txt(s,x-Inches(3.05),ry-Inches(0.02),Inches(3.0),Inches(0.3),nm,9.5,DK,False,PP_ALIGN.RIGHT)
        xa=fx(x,w,D(a)); xb=fx(x,w,D(b))
        bar=s.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,Emu(xa),ry,Emu(max(xb-xa,30000)),Emu(int(rh*0.6)))
        bar.fill.solid(); bar.fill.fore_color.rgb=col
        if crit: bar.line.color.rgb=RED; bar.line.width=Pt(1.5)
        else: bar.line.fill.background()
        bar.shadow.inherit=False
    # linea RFSU
    rx=fx(x,w,date(2027,12,20))
    rl=rect(s,Emu(rx),y-Inches(0.05),Emu(20000),h+Inches(0.1),RED)
    txt(s,Emu(rx-Emu(380000)),y+h+Inches(0.02),Inches(1.2),Inches(0.25),"RFSU",8,RED,True)

# ============ 1 PORTADA ============
s=slide(); rect(s,0,0,SW,SH,NAVY); rect(s,0,Inches(4.4),SW,Inches(0.07),GOLD)
txt(s,Inches(0.7),Inches(1.4),Inches(12),Inches(0.4),"CPF2 LA CALERA II · VACA MUERTA",16,RGBColor(0x9D,0xB8,0xD8),True)
txt(s,Inches(0.7),Inches(1.95),Inches(12),Inches(1.6),"Alcance de Especialidades\nElectricidad · Instrumentación",38,WHITE,True)
txt(s,Inches(0.7),Inches(4.55),Inches(12),Inches(0.6),"Sistemas · equipos · cargas · cables · salas · suministro · magnitudes y tiempos",17,LBLUE)
txt(s,Inches(0.7),Inches(5.2),Inches(12),Inches(0.5),"Eje en fechas para cumplir RFSU 20-DIC-2027",17,GOLD,True)
txt(s,Inches(0.7),Inches(6.7),Inches(9),Inches(0.4),"Rev. 0 · 13-06-2026 · Primera revisión para depuración",11,RGBColor(0x9D,0xB8,0xD8))
footer(s,1)

# ============ 2 ENFOQUE ============
s=slide(); header(s,"Enfoque de esta presentación",NAVY)
bullets(s,Inches(0.55),Inches(1.5),Inches(12.3),Inches(5),[
 (0,"Se presenta el alcance SEPARADO por especialidad, con datos de rápida visualización:",NAVY),
 (1,"ELECTRICIDAD — sistemas y equipos, cargas, cables y salas; magnitudes y tiempos de ejecución.",EL),
 (1,"INSTRUMENTACIÓN — sistemas, provisión (AESA vs proveedor), montaje, cableado y Sala de Instrumentos.",IN),
 (0,"Premisa transversal: PRIMERO el suministro y compras (equipos, salas, cables, instrumentos).",GOLD),
 (0,"Precomisionado — pruebas en banco, en campo y punta a punta.",GREEN),
 (0,"Cronograma reducido (Gantt + hitos) con fechas tentativas para asegurar el RFSU 20-DIC-2027.",RED),
],size=15,gap=10)
footer(s,2)

# ============ 3 DIVIDER EL ============
s=slide(); divider(s,"ELECTRICIDAD",EL,"Sistemas y equipos · cargas · cables · salas · suministro · tiempos")

# ============ 4 EL SUMINISTRO ============
s=slide(); header(s,"Electricidad · Suministro y compras (primero)",EL)
rows=[["Hito de suministro","Fecha","Nota"],
 ["OC ABB Salas SE#3 + SE#4 — aceptada","12-MAY-26","Inicio oficial ABB"],
 ["OC ABB PMS — aceptada","22-MAY-26","PMS-001 + PMS-101"],
 ["★ OC Cables EL (target)","01-OCT-26","LT 150 d"],
 ["Fabricación tableros (Brasil/Bs.As.)","SEP-NOV-26","37 tableros BT + PMS"],
 ["FAT S#3 + S#4 (Brasil)","23-NOV-26","Witness AESA"],
 ["Llegada Cables EL a sitio","28-FEB-27","513 cables / 81.231 m"],
 ["Llegada SE#4 / SE#3 a CPF2","08-ABR / 14-MAY-27","En módulos → armado en campo"]]
table(s,Inches(0.45),Inches(1.5),Inches(12.45),rows,[5.5,2.6,4.0],hfill=EL,rowh=Inches(0.5),accent_rows=[3,6])
txt(s,Inches(0.45),Inches(6.45),Inches(12.4),Inches(0.5),
    "El suministro gobierna el inicio de obra: sin salas, cables y equipos en sitio no arranca el montaje. OC Cables EL ≤ 01-OCT-26 es condición para tender antes de mayo.",12,GREY)
footer(s,4)

# ============ 5 EL SISTEMAS Y EQUIPOS ============
s=slide(); header(s,"Electricidad · Sistemas y equipos involucrados",EL)
for i,(b,sm) in enumerate([("2","Salas eléctricas SE#3 + SE#4 (ABB, en módulos)"),
        ("37","Tableros BT / CCMs"),("PMS","Power Mgmt PMS-001 + PMS-101 (ABB)"),
        ("MT","Celdas MT TGMT / TRMT + transformadores"),("6","Variadores VFD (3 MT + 3 BT)"),
        ("79","Motores (76 BT + 3 MT)")]):
    r,c=divmod(i,3); card(s,Inches(0.45)+c*Inches(4.25),Inches(1.6)+r*Inches(2.0),
        Inches(3.95),Inches(1.7),b,sm,EL)
txt(s,Inches(0.45),Inches(6.35),Inches(12.4),Inches(0.6),
    "Sistemas: distribución MT/BT, PMS (gestión de energía ABB), fuerza motriz, SSAA, iluminación, calefacción/trazado. Ampliación de la red de control PMS SE#3↔SE#4 (FO).",12,GREY)
footer(s,5)

# ============ 6 EL CARGAS Y CABLES ============
s=slide(); header(s,"Electricidad · Cargas y cables (magnitudes)",EL)
txt(s,Inches(0.45),Inches(1.3),Inches(6),Inches(0.3),"243 CARGAS ELÉCTRICAS",13,EL,True)
rows=[["Tipo de carga","Cant."],["Motores BT","76"],["Motores MT","3"],["Instrumentos","52"],
 ["SSAA","37"],["Control","29"],["Iluminación","28"],["Calefacción","12"],["VFD (BT+MT)","6"]]
table(s,Inches(0.45),Inches(1.65),Inches(5.7),rows,[4.0,1.2],hfill=EL,fs=11,hfs=11,rowh=Inches(0.40))
txt(s,Inches(6.6),Inches(1.3),Inches(6),Inches(0.3),"513 CABLES · 81.231 m · ~5.178 PUNTAS",13,EL,True)
rows2=[["Tipo de cable","Cables","Metros"],["Potencia MT","18","2.110"],
 ["BT Grande ≥35mm² ◄","104","17.470"],["BT Med/Peq","198","33.125"],
 ["Control y Señales","192","28.126"],["FO / Ethernet","1","400"],["TOTAL","513","81.231"]]
table(s,Inches(6.6),Inches(1.65),Inches(6.3),rows2,[3.2,1.3,1.6],hfill=EL,fs=11,hfs=11,rowh=Inches(0.40),accent_rows=[6])
txt(s,Inches(6.6),Inches(4.85),Inches(6.3),Inches(0.8),"Puntas: 3.036 en salas + 1.717 en campo.\nBT Grande (104 c/17.470 m) = frente de ruta crítica EL.",11.5,GREY)
footer(s,6)

# ============ 7 EL TIEMPOS ============
s=slide(); header(s,"Electricidad · Tiempos de ejecución",EL)
rows=[["Frente de trabajo","Inicio","Fin","Días"],
 ["Bandejas portacables","01-MAR-27","05-ABR-27","36"],
 ["Tendido EL — BT Grande ◄ crítico","01-MAR-27","18-ABR-27","49"],
 ["Tendido EL — Med/Peq + Control","15-MAR-27","13-MAY-27","60"],
 ["Tendido EL — Potencia MT","10-MAY-27","07-JUN-27","29"],
 ["Armado shelters SE#4 / SE#3","08-ABR / 14-MAY-27","22-ABR / 03-JUN-27","14 / 21"],
 ["Conexionado EL SE#4 → MT → SE#3","23-ABR-27","25-JUL-27","~94"],
 ["Megger (HOLD POINT)","26-JUL-27","15-AGO-27","21"],
 ["Energización progresiva MT→BT","16-AGO-27","01-SEP-27","17"]]
table(s,Inches(0.45),Inches(1.5),Inches(12.45),rows,[5.6,2.5,2.5,1.4],hfill=EL,fs=12,rowh=Inches(0.48),accent_rows=[2,7])
txt(s,Inches(0.45),Inches(6.45),Inches(12.4),Inches(0.5),
    "~263 días con 3-4 cuadrillas. Tendido pivota antes de la llegada de salas; conexionado pivota sobre el armado en campo de cada shelter.",12,GREY)
footer(s,7)

# ============ 8 DIVIDER IN ============
s=slide(); divider(s,"INSTRUMENTACIÓN",IN,"Sistemas · provisión AESA vs proveedor · montaje · cableado · Sala INS")

# ============ 9 IN SISTEMAS Y SUMINISTRO ============
s=slide(); header(s,"Instrumentación · Sistemas y suministro",IN)
rows=[["Sistema","Descripción","Provee"],
 ["PCS","Control de procesos (DCS) — Sala INS / Sala 7","Inauco"],
 ["ESD / SIS","Parada de emergencia (seguridad)","HIMA"],
 ["PSS","Sistema de seguridad/parada de procesos","HIMA"],
 ["F&G","Fuego y Gas — detectores de campo + hardware","AESA (det.) / HIMA (HW)"],
 ["SCADA","Supervisión y adquisición de datos","Inauco"],
 ["PMS","Power Management System (interfase EL)","ABB"]]
table(s,Inches(0.45),Inches(1.5),Inches(7.3),rows,[1.4,4.3,2.0],hfill=IN,fs=11.5,rowh=Inches(0.52))
bullets(s,Inches(8.0),Inches(1.55),Inches(4.9),Inches(4.2),[
 (0,"Suministro (primero):",GOLD),
 (1,"OC Cables IN ≤ 01-AGO-26 → sitio 28-NOV-26.",None),
 (1,"Instrumentos: entrega PROGRESIVA (Fase 1 válvulas/en línea; Fase 2 resto).",None),
 (0,"Sala de Instrumentos (Sala INS / Sala 7):",IN),
 (1,"Aloja PCS + nodos SIS; Shelters INS005/006/007.",None),
 (1,"Inicio instalación PCS 17-MAR-27.",None),
],size=12.5)
footer(s,9)

# ============ 10 IN PROVISION AESA vs PROVEEDOR ============
s=slide(); header(s,"Instrumentación · Provisión: AESA vs proveedor",IN)
card(s,Inches(0.45),Inches(1.5),Inches(3.0),Inches(1.5),"897","Provistos por AESA",GREEN)
card(s,Inches(3.6),Inches(1.5),Inches(3.0),Inches(1.5),"1.894","Provistos por proveedor",EL)
card(s,Inches(6.75),Inches(1.5),Inches(3.0),Inches(1.5),"7","Existentes (reutil.)",GREY)
card(s,Inches(9.9),Inches(1.5),Inches(3.0),Inches(1.5),"2.798","Instrumentos totales",NAVY)
txt(s,Inches(0.45),Inches(3.25),Inches(6),Inches(0.3),"Provisión por proveedor (1.894)",12,EL,True)
rows=[["Proveedor / paquete","Cant."],["PROPAK (skid principal)","1.505"],
 ["Skid gas combustible","165"],["Proveedor de válvula","85"],["Proveedor de bomba","60"],
 ["Skid inyección químicos","32"],["Tanque / filtro / otros","41"]]
table(s,Inches(0.45),Inches(3.6),Inches(6.0),rows,[4.2,1.2],hfill=EL,fs=11,hfs=11,rowh=Inches(0.36))
bullets(s,Inches(6.9),Inches(3.4),Inches(6.0),Inches(3.4),[
 (0,"Lectura para mensurar el alcance:",IN),
 (1,"AESA provee 1/3 de los instrumentos (897); el resto llega con paquetes de vendor (PROPAK domina).",None),
 (1,"F&G (120) es 100% AESA. ZSO/ZSC (fines de carrera) mayormente AESA.",None),
 (1,"Termovainas, TE, PSV, RO, LG provienen mayormente de proveedor (skids/cañería).",None),
],size=12.5)
footer(s,10)

# ============ 11 IN QUIEN MONTA ============
s=slide(); header(s,"Instrumentación · ¿Quién monta? (mensurar el trabajo)",IN)
card(s,Inches(0.45),Inches(1.5),Inches(4.0),Inches(1.5),"1.163","Monta AESA (897 + 266 'EPC')",GREEN)
card(s,Inches(4.7),Inches(1.5),Inches(4.0),Inches(1.5),"1.635","Premontados por vendor (skids)",EL)
card(s,Inches(8.95),Inches(1.5),Inches(3.9),Inches(1.5),"~42%","Del total lo monta AESA",IN)
rows=[["Típico de montaje","Monta AESA","Premont. vendor","Total"],
 ["Válvula en línea","145","340","485"],
 ["Elemento en línea","81","354","435"],
 ["Soporte junto a línea (stand)","243","367","610"],
 ["Sobre equipo","125","515","640"],
 ["Sobre válvula (accesorio)","183","300","483"],
 ["Otro instrumento","0","25","25"],
 ["Estructura / campo (F&G)","120","0","120"],
 ["TOTAL","1.163","1.635","2.798"]]
table(s,Inches(0.45),Inches(3.3),Inches(12.45),rows,[4.6,2.2,2.4,1.4],hfill=IN,fs=11.5,rowh=Inches(0.355),accent_rows=[8])
txt(s,Inches(0.45),Inches(6.55),Inches(12.4),Inches(0.4),
    "AESA monta sus instrumentos + los de proveedor 'montaje por EPC' (266). Los premontados en skids los instala el vendor, pero los LOOP CHECKS son alcance AESA.",11.5,GREY)
footer(s,11)

# ============ 12 IN POR TIPO ============
s=slide(); header(s,"Instrumentación · Discriminación por tipo",IN)
txt(s,Inches(0.45),Inches(1.3),Inches(6),Inches(0.3),"Top tipos (AESA / proveedor)",12,IN,True)
rows=[["Cód.","Tipo","AESA","Prov."],["PI","Manómetro","136","133"],
 ["TW","Termovaina","23","198"],["TE","Sensor temp. (RTD/TC)","21","157"],
 ["PSV","Válvula de seguridad","22","140"],["PIT","Transmisor de presión","81","79"],
 ["RO","Orificio de restricción","9","118"],["LG","Indicador nivel magnético","22","104"],
 ["ZSO/ZSC","Fines de carrera","107","10"]]
table(s,Inches(0.45),Inches(1.65),Inches(6.6),rows,[1.2,3.4,1.1,1.1],hfill=IN,fs=11,hfs=11,rowh=Inches(0.40))
txt(s,Inches(7.4),Inches(1.3),Inches(5.4),Inches(0.3),"Por tipo de montaje (7 típicos)",12,IN,True)
rows2=[["Típico","Cant."],["Sobre equipo","640"],["Soporte junto a línea","610"],
 ["Válvula en línea","485"],["Sobre válvula","483"],["Elemento en línea","435"],
 ["F&G (estructura)","120"],["Otro instrumento","25"]]
table(s,Inches(7.4),Inches(1.65),Inches(5.4),rows2,[3.6,1.1],hfill=IN,fs=11,hfs=11,rowh=Inches(0.40))
txt(s,Inches(0.45),Inches(6.2),Inches(12.4),Inches(0.7),
    "En línea (válvulas + elementos) = 920 → primeros en entregar/montar. Montaje específico = 1.878 → entrega progresiva. La distribución por tipo dimensiona típicos de montaje y materiales.",12,GREY)
footer(s,12)

# ============ 13 IN CABLEADO Y SALA ============
s=slide(); header(s,"Instrumentación · Cableado y Sala de Instrumentos",IN)
card(s,Inches(0.45),Inches(1.5),Inches(3.0),Inches(1.5),"885","Cables IN",IN)
card(s,Inches(3.6),Inches(1.5),Inches(3.0),Inches(1.5),"65.315 m","Metros a tender",IN)
card(s,Inches(6.75),Inches(1.5),Inches(3.0),Inches(1.5),"15.752","Puntas a interconectar",IN)
card(s,Inches(9.9),Inches(1.5),Inches(3.0),Inches(1.5),"3","Shelters INS (005/06/07)",IN)
bullets(s,Inches(0.45),Inches(3.35),Inches(12.4),Inches(3.2),[
 (0,"Recorrido: instrumento de campo → Junction Box (JB) → Shelter Sala INS → PCS / SIS.",IN),
 (0,"Tendido: ~150 días con 3 cuadrillas; ventana 01-FEB→14-MAR exige 5-6 cuadrillas (antes del PCS 17-MAR).",None),
 (0,"Conexionado: ~247 días-hombre (15.752 puntas); arranca con la instalación del PCS (17-MAR→31-JUL).",None),
 (0,"Tipos de cable: par/terna apantallado armado (1P/2P/multipar #16-18AWG), FO monomodo y FTP Cat6.",None),
],size=13.5,gap=9)
footer(s,13)

# ============ 14 PRECOMISIONADO ============
s=slide(); header(s,"Precomisionado · Tareas a contemplar",GREEN)
rows=[["Tarea","Cuándo","Descripción"],
 ["Pruebas en BANCO (calibración)","DIC-26 → MAR-27","Calibración/seteo de instrumentos ANTES del montaje (por fase de entrega)"],
 ["Prueba PUNTA a PUNTA (continuidad)","MAR → AGO-27","Continuidad cable extremo a extremo + identificación, por cada punta"],
 ["Megger EL (aislación)","26-JUL → 15-AGO-27","Circuito por circuito · HOLD POINT ITP"],
 ["Loop check IN (lazo completo)","JUN → AGO-27","Sensor → JB → Sala INS → PCS/SIS"],
 ["Energización progresiva","16-AGO → 01-SEP-27","MT → transformadores → BT · HOLD POINT"],
 ["Precomisionado funcional","SEP-27","EL (02→22-SEP) + IN (lazos PCS/SIS)"]]
table(s,Inches(0.45),Inches(1.5),Inches(12.45),rows,[3.4,2.6,6.2],hfill=GREEN,fs=11.5,rowh=Inches(0.6),accent_rows=[1,2])
txt(s,Inches(0.45),Inches(6.55),Inches(12.4),Inches(0.4),
    "Arranca con el banco (pre-montaje) y se completa con el comisionado integrado (23-SEP→19-DIC) y SAT PMS (OCT-DIC) → RFSU 20-DIC.",12,RED,True)
footer(s,14)

# ============ 15 CRONOGRAMA REDUCIDO ============
s=slide(); header(s,"Cronograma reducido · etapas e hitos hacia RFSU",NAVY)
phases=[("Suministro y compras","2026-06-01","2027-05-14",GOLD,False),
 ("Tendido cables IN","2027-02-01","2027-03-14",IN,True),
 ("Montaje instrumentos","2027-01-15","2027-06-30",IN,False),
 ("Tendido cables EL","2027-03-01","2027-05-13",EL,True),
 ("Armado salas + conexionado EL","2027-04-08","2027-07-25",EL,True),
 ("Conexionado IN + loop check","2027-03-17","2027-08-20",IN,False),
 ("Megger + energización","2027-07-26","2027-09-01",GREEN,True),
 ("Precom + comisionado + SAT PMS","2027-09-01","2027-12-19",GREEN,True)]
mini_gantt(s,Inches(3.6),Inches(1.85),Inches(9.0),Inches(3.2),phases)
txt(s,Inches(0.45),Inches(5.35),Inches(4),Inches(0.3),"Hitos para llegar al objetivo:",12,NAVY,True)
rows=[["Etapa lista","Fecha"],["Cables IN en sitio","28-NOV-26"],["Tendido IN completo / PCS","14→17-MAR-27"],
 ["Cables EL en sitio","28-FEB-27"],["SE#4 / SE#3 en sitio","08-ABR / 14-MAY-27"],
 ["Conexionado completo","25-JUL-27"],["Energización","01-SEP-27"],
 ["Precom completo","30-SEP-27"],["★ RFSU","20-DIC-27"]]
table(s,Inches(0.45),Inches(5.7),Inches(12.45),
      [rows[0]]+[[a,b] for a,b in rows[1:]],[3.0,1.4],hfill=NAVY,fs=10.5,hfs=10.5,rowh=Inches(0.0)) if False else None
# planilla compacta horizontal
flat=[["Cables IN sitio","Tendido IN/PCS","Cables EL sitio","SE#4 / SE#3","Conexionado fin","Energización","Precom fin","RFSU"],
      ["28-NOV-26","14→17-MAR-27","28-FEB-27","08-ABR / 14-MAY","25-JUL-27","01-SEP-27","30-SEP-27","20-DIC-27"]]
table(s,Inches(0.45),Inches(5.75),Inches(12.45),flat,[1,1,1,1.1,1,1,1,1],hfill=NAVY,fs=10,hfs=10,rowh=Inches(0.42))
footer(s,15)

# ============ 16 CONTEXTO 3D ============
import fitz as _fitz; _fitz.TOOLS.mupdf_display_errors(False)
ROOT=os.path.dirname(BASE)
_3d=os.path.join(ROOT,"Datos entrada","3D LC 120626.pdf")
_gp=os.path.join(BASE,"Cronograma_INS_EL_CPF2.pdf")
tmp3d=os.path.join(BASE,"_t3d.png"); tmpg=os.path.join(BASE,"_tg.png")
_pg=_fitz.open(_3d)[0]; _pg.get_pixmap(dpi=160).save(tmp3d)
iw,ih=_pg.rect.width,_pg.rect.height
s=slide(); header(s,"Anexo · Contexto de planta (vista 3D)",NAVY)
ph=Inches(5.3); pw=ph*iw/ih
s.shapes.add_picture(tmp3d,int(SW/2-pw/2),Inches(1.45),height=ph)
txt(s,Inches(0.45),Inches(6.85),Inches(12.4),Inches(0.3),
    "Modelo 3D CPF2 — escala de volumen y congestión: sustento físico de los volúmenes (cables, puntas, instrumentos) y de los factores de productividad y pico de cuadrillas.",10.5,GREY)
footer(s,16)

# ============ 17 GANTT DETALLADO ============
if os.path.exists(_gp):
    _g=_fitz.open(_gp)[0]; _g.get_pixmap(dpi=170).save(tmpg)
    gw,gh=_g.rect.width,_g.rect.height
    s=slide(); header(s,"Anexo · Cronograma detallado (Gantt INS+EL)",NAVY)
    iw2=Inches(12.5); ih2=iw2*gh/gw
    if ih2>Inches(5.4): ih2=Inches(5.4); iw2=ih2*gw/gh
    s.shapes.add_picture(tmpg,int(SW/2-iw2/2),Inches(1.4),width=int(iw2),height=int(ih2))
    txt(s,Inches(0.45),Inches(6.95),Inches(12.4),Inches(0.3),
        "Detalle en Cronograma_INS_EL_CPF2.pdf (Gantt + tabla de tareas + bases). RFSU 20-DIC-2027.",10.5,GREY)
    footer(s,17)

prs.save(OUT)
for _t in (tmp3d,tmpg):
    try: os.remove(_t)
    except OSError: pass
print("OK ->",OUT,"| slides:",len(prs.slides._sldIdLst))
