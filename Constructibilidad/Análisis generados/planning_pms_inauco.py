#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Planning de Provisiones de Sistemas - PMS (ABB) + PCS/SCADA/SIS (Inauco)
CPF2 La Calera II. Encadena ambas provisiones, sus puntos en comun (FAT de cada
una, integracion PCS<->PMS, iFAT) y la convergencia a precomisionado y
comisionado para cumplir el RFSU 20-DIC-2027.
Fuentes: CRONOGRAMA_GENERAL_ABB_Actualizado_2026-06-09 (PMS detalle) +
         2026-04 Cronograma Tentativo CPF2 LCA (Inauco, arranque JUL-2026).
Salidas: Planning_PMS_Inauco_CPF2.pdf (Gantt + flujo + planilla) y .xlsx
"""
import os
from datetime import date
from collections import OrderedDict

BASE=os.path.dirname(os.path.abspath(__file__))
OUT_PDF=os.path.join(BASE,"Planning_PMS_Inauco_CPF2.pdf")
OUT_XLS=os.path.join(BASE,"Planning_PMS_Inauco_CPF2.xlsx")
REV="0"; HOY=date(2026,6,13); RFSU=date(2027,12,20)
T0=date(2026,5,1); T1=date(2028,1,31)
def d(s): y,m,dd=map(int,s.split('-')); return date(y,m,dd)

C_PMS='#c55a11'; C_INA='#2e75b6'; C_INT='#548235'; C_MILE='#bf9000'; C_CRIT='#c00000'

# (lane, id, nombre, inicio, fin|None, color, critico, nota)
T=[
 ('PMS','##','PMS — PROVISIÓN ABB  (OC 4508945953)',None,None,None,False,''),
 ('PMS','◆','KOM PMS','2026-05-22',None,C_MILE,False,'arranque ABB'),
 ('PMS','PP-E','Ingeniería PMS (IB→ID, ciclo AESA)','2026-06-08','2026-10-23',C_PMS,False,'Arq/IO/Variables/Func.'),
 ('PMS','◆','FP Ing. Básica PMS','2026-08-21',None,C_MILE,False,'freezing point'),
 ('PMS','◆','FP Ing. Detalle PMS (último)','2026-10-23',None,C_MILE,False,'freezing point'),
 ('PMS','PP-PR','Procura (switches · AC800M · tableros)','2026-05-22','2026-11-05',C_PMS,False,'acopio 25% 05-NOV'),
 ('PMS','PP-SW','Software y config 800XA','2026-08-03','2026-11-27',C_PMS,False,'IEC61850/PROFINET/MODBUS'),
 ('PMS','PP-MQ','FAT Maqueta (prueba temprana, AESA)','2026-09-22','2026-09-25',C_PMS,False,'witness AESA'),
 ('PMS','PP-FAB','Armado tableros PMS-001 + PMS-101','2026-09-11','2026-12-17',C_PMS,False,'BsAs'),
 ('PMS','PP-TI','Pruebas internas ABB','2026-12-18','2027-01-07',C_PMS,False,'IV per PIE'),
 ('PMS','★FAT','★ FAT PMS (tableros + sistema, witness)','2027-01-25','2027-01-29',C_CRIT,True,'FAT · 16% OC'),
 ('PMS','PP-DES','Despacho PMS → Shelter Mendoza','2027-02-04','2027-02-17',C_PMS,False,'hito 30%'),
 ('PMS','PP-SH','Montaje PMS en Shelter','2027-02-18','2027-03-17',C_PMS,False,'PMS-001+PMS-101'),
 ('PMS','★FATi','★ FAT Integral Shelter PMS','2027-03-25','2027-03-31',C_CRIT,True,'FAT integral'),
 ('PMS','◆','DataBook PMS aprobado','2027-03-31',None,C_MILE,False,'5% OC'),
 ('PMS','LLEG','Llegada PMS a sitio (con SE#4 / SE#3)','2027-04-08','2027-05-14',C_PMS,False,'PMS-101 / PMS-001'),
 ('PMS','SATP','SAT PMS en sitio (HOLD POINT)','2027-10-06','2027-12-05',C_PMS,True,'ventana fija OCT-DIC'),

 ('INA','##','INAUCO — PCS / SCADA / SIS  (OC pendiente · arranque JUL-2026)',None,None,None,False,''),
 ('INA','◆','INICIO Inauco (OC + Ing. Rev0 + HW HIMA)','2026-07-01',None,C_MILE,True,'arranque JUL · OC pend.'),
 ('INA','ING','Ingeniería Rev A y 0','2026-07-01','2026-11-09',C_INA,False,'aprob. construcción 09-NOV'),
 ('INA','◆','Recepción HW (tableros HIMA + HW)','2026-08-21',None,C_MILE,False,'+10 sem.'),
 ('INA','PLC','Configuración PLC','2026-07-01','2026-08-28',C_INA,False,''),
 ('INA','SCADA','Configuración SCADA','2026-07-01','2027-01-26',C_INA,False,'140 d'),
 ('INA','TAB','Construcción tableros PCS (7A + 7B)','2026-11-10','2027-01-26',C_INA,False,'50 d'),
 ('INA','★FATp','★ FAT PCS + Comunicaciones','2027-01-27','2027-02-19',C_CRIT,True,'FAT · 18 d'),
 ('INA','★FATs','★ FAT SIS','2027-01-27','2027-02-23',C_CRIT,True,'FAT · 20 d'),
 ('INA','POST','Configuraciones post-FAT','2027-02-24','2027-03-16',C_INA,False,''),
 ('INA','SATpcs','SAT y Comisionado PCS (campo)','2027-03-17','2027-06-01',C_INA,False,'55 d'),
 ('INA','SATsis','SAT y Comisionado SCADA-SIS (campo)','2027-03-17','2027-08-03',C_INA,True,'100 d'),
 ('INA','MCE','Prueba Matriz C&E (MCE)','2027-08-04','2027-08-24',C_INA,False,'15 d'),
 ('INA','CAO','CAO / As-Built Inauco','2027-08-25','2027-09-21',C_INA,False,''),

 ('INT','##','INTEGRACIÓN Y CONVERGENCIA → RFSU',None,None,None,False,''),
 ('INT','★INT','★ Integración PCS↔PMS — Pruebas ABB Bs.As.','2027-02-24','2027-03-02',C_CRIT,True,'PUNTO EN COMÚN'),
 ('INT','◆','Inicio campo PCS — Sala INS (Sala 7)','2027-03-17',None,C_CRIT,True,'ancla campo IN'),
 ('INT','★iFAT','★ iFAT — SE#3+SE#4+PMS+PCS+SIS','2027-06-25','2027-07-14',C_CRIT,True,'PRUEBA INTEGRADA'),
 ('INT','PREC','Precomisionado (EL + IN)','2027-09-02','2027-09-30',C_INT,True,'banco + campo + punta-punta'),
 ('INT','COM','Comisionado integrado (EL+PCS+PMS+SIS)','2027-09-23','2027-12-19',C_INT,True,'88 d'),
 ('INT','★RFSU','★★★ RFSU — READY FOR START UP','2027-12-20',None,C_CRIT,True,'objetivo'),
]

# ---------------- PDF ----------------
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (BaseDocTemplate,PageTemplate,Frame,Paragraph,Spacer,
        Table,TableStyle,LongTable,Flowable,PageBreak)
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.enums import TA_CENTER
PAGE=landscape(A4)
st=getSampleStyleSheet()
H1=ParagraphStyle('H1',parent=st['Title'],fontSize=15,spaceAfter=2)
SUB=ParagraphStyle('SUB',parent=st['Normal'],fontSize=9,textColor=colors.HexColor('#555'),alignment=TA_CENTER)
H2=ParagraphStyle('H2',parent=st['Heading2'],fontSize=11,textColor=colors.HexColor('#1f3b63'),spaceBefore=6,spaceAfter=4)
N=ParagraphStyle('N',parent=st['Normal'],fontSize=8.5,leading=11)
CELL=ParagraphStyle('CELL',parent=st['Normal'],fontSize=7.5,leading=9)
MES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
def months(a,b):
    out=[];y,m=a.year,a.month
    while (y,m)<=(b.year,b.month):
        out.append(date(y,m,1)); m+=1
        if m>12:m=1;y+=1
    return out

class Gantt(Flowable):
    def __init__(s,tasks,w,h,lw=78*mm): s.t=tasks;s.w=w;s.h=h;s.lw=lw;Flowable.__init__(s)
    def wrap(s,a,b): return (s.w,s.h)
    def x(s,dt): return s.lw+(s.w-s.lw)*((dt-T0).days)/((T1-T0).days)
    def draw(s):
        c=s.canv;W=s.w;H=s.h;lw=s.lw;head=10*mm;n=len(s.t);rh=(H-head)/n
        ly=None
        for mdt in months(T0,T1):
            xx=s.x(mdt)
            c.setStrokeColor(colors.HexColor('#e8e8e8'));c.setLineWidth(0.3);c.line(xx,0,xx,H-head)
            nxt=date(mdt.year+(mdt.month//12),(mdt.month%12)+1,1)
            c.setFillColor(colors.HexColor('#666'));c.setFont('Helvetica',5)
            c.drawCentredString((xx+s.x(nxt))/2,H-head+1.5,MES[mdt.month-1])
            if mdt.year!=ly:
                c.setFillColor(colors.HexColor('#1f3b63'));c.setFont('Helvetica-Bold',7)
                c.drawString(xx+1,H-head+5.3,str(mdt.year));ly=mdt.year
                c.setStrokeColor(colors.HexColor('#b0b0b0'));c.setLineWidth(0.6);c.line(xx,0,xx,H-head+5)
        c.setStrokeColor(colors.HexColor('#bbb'));c.setLineWidth(0.5);c.rect(lw,0,W-lw,H-head,stroke=1,fill=0)
        for dt,col,lbl in [(HOY,'#1f9d55','HOY'),(RFSU,'#c00000','RFSU 20-DIC-27')]:
            xx=s.x(dt);c.setStrokeColor(colors.HexColor(col));c.setLineWidth(1);c.setDash(2,2)
            c.line(xx,0,xx,H-head);c.setDash()
            c.saveState();c.setFillColor(colors.HexColor(col));c.setFont('Helvetica-Bold',5.5)
            c.translate(xx,H-head-1);c.rotate(90);c.drawString(2,1.2,lbl);c.restoreState()
        for i,t in enumerate(s.t):
            lane,tid,name,ini,fin,color,crit,note=t; y=H-head-(i+1)*rh
            if tid=='##':
                c.setFillColor(colors.HexColor('#1f3b63'));c.rect(0,y,W,rh,stroke=0,fill=1)
                c.setFillColor(colors.white);c.setFont('Helvetica-Bold',6.6);c.drawString(2,y+rh*0.28,name);continue
            if i%2==0: c.setFillColor(colors.HexColor('#f6f6f6'));c.rect(0,y,W,rh,stroke=0,fill=1)
            c.setFillColor(colors.black);c.setFont('Helvetica',5.8);c.drawString(3,y+rh*0.28,name[:54])
            c.setFillColor(colors.HexColor('#999'));c.setFont('Helvetica',4.8);c.drawRightString(lw-2,y+rh*0.28,note[:24])
            if fin is None and ini:
                xx=s.x(d(ini));sz=rh*0.30;c.setFillColor(colors.HexColor(color))
                c.saveState();c.translate(xx,y+rh/2);c.rotate(45);c.rect(-sz,-sz,2*sz,2*sz,stroke=0,fill=1);c.restoreState()
            elif ini and fin:
                x0=s.x(d(ini));x1=s.x(d(fin));bh=rh*0.5;yy=y+(rh-bh)/2
                c.setFillColor(colors.HexColor(color));c.roundRect(x0,yy,max(x1-x0,1),bh,1,stroke=0,fill=1)
                if crit: c.setStrokeColor(colors.HexColor('#c00000'));c.setLineWidth(0.8);c.roundRect(x0,yy,max(x1-x0,1),bh,1,stroke=1,fill=0)

def box(c,x,y,w,h,lines,fill,fc=colors.white,fs=7.5,bold0=True):
    c.setFillColor(colors.HexColor(fill));c.roundRect(x,y,w,h,2,stroke=0,fill=1)
    c.setFillColor(fc)
    th=fs+1.5;ty=y+h-th
    for i,ln in enumerate(lines):
        c.setFont('Helvetica-Bold' if (i==0 and bold0) else 'Helvetica', fs if i==0 else fs-0.8)
        c.drawCentredString(x+w/2,ty,ln);ty-=th-0.5
def arrow(c,x0,y0,x1,y1,col='#555'):
    c.setStrokeColor(colors.HexColor(col));c.setLineWidth(1.2);c.line(x0,y0,x1,y1)
    import math;a=math.atan2(y1-y0,x1-x0);L=2.2*mm
    c.setFillColor(colors.HexColor(col))
    c.saveState();c.translate(x1,y1);c.rotate(math.degrees(a))
    p=c.beginPath();p.moveTo(0,0);p.lineTo(-L,L*0.5);p.lineTo(-L,-L*0.5);p.close();c.drawPath(p,fill=1,stroke=0);c.restoreState()

class Flow(Flowable):
    def __init__(s,w,h): s.w=w;s.h=h;Flowable.__init__(s)
    def wrap(s,a,b): return (s.w,s.h)
    def draw(s):
        c=s.canv;W=s.w;H=s.h
        bw=W/5-6*mm; bh=24*mm; gap=(W-5*bw)/4
        r1=H-bh-6*mm; r2=r1-bh-18*mm
        B=[]
        row1=[("Ingeniería + Freezing\nPoints (PMS + Inauco)\nJUN→NOV-26",C_PMS),
              ("Fabricación tableros\nPMS-001/101 · PCS 7A/7B\nSEP-26→ENE-27",C_INA),
              ("FAT individuales\nPMS 29-ENE · PCS 19-FEB\nSIS 23-FEB",C_MILE),
              ("★ Integración PCS↔PMS\nPruebas ABB Bs.As.\n24-FEB→02-MAR-27",C_CRIT),
              ("Montaje Shelter Mendoza\n+ FAT Integral PMS\n31-MAR-27",C_PMS)]
        row2=[("Campo: llegada salas\n08-ABR / 14-MAY + SAT\nPCS · SIS · PMS",C_INA),
              ("★ iFAT integrado\nSE#3+SE#4+PMS+PCS+SIS\n25-JUN→14-JUL-27",C_CRIT),
              ("MCE + Precomisionado\n(banco·campo·punta-punta)\nAGO → SEP-27",C_INT),
              ("Comisionado integrado\nEL+PCS+PMS+SIS\n23-SEP→19-DIC-27",C_INT),
              ("★★★ RFSU\nREADY FOR START UP\n20-DIC-2027",C_CRIT)]
        xs=[6*mm+i*(bw+gap) for i in range(5)]
        for i,(txt,col) in enumerate(row1):
            box(c,xs[i],r1,bw,bh,txt.split('\n'),col); B.append((xs[i],r1))
        for i,(txt,col) in enumerate(row2):
            box(c,xs[i],r2,bw,bh,txt.split('\n'),col)
        # flechas row1
        for i in range(4): arrow(c,xs[i]+bw,r1+bh/2,xs[i+1],r1+bh/2)
        # baja de row1[4] a row2[0]... snake: row2 va de derecha a izq visual? lo dejamos L->R con bajada del ultimo
        arrow(c,xs[4]+bw/2,r1,xs[4]+bw/2,r2+bh, '#555')  # baja
        arrow(c,xs[4]+bw/2,r2+bh,xs[0]+bw/2,r2+bh)       # conector largo a inicio row2 (visual)
        for i in range(4): arrow(c,xs[i]+bw,r2+bh/2,xs[i+1],r2+bh/2)
        # nota
        c.setFillColor(colors.HexColor('#777'));c.setFont('Helvetica-Oblique',6.5)
        c.drawString(6*mm,2*mm,"Puntos en común: FAT de cada provisión → Integración PCS↔PMS (Bs.As.) → iFAT integrado → comisionado. Convergen al RFSU 20-DIC-2027.")

def footer(cv,doc):
    cv.saveState();cv.setFont('Helvetica',7);cv.setFillColor(colors.HexColor('#888'))
    cv.drawString(12*mm,7*mm,"Constructibilidad CPF2 · Planning Provisiones PMS (ABB) + Inauco · Rev. %s · RFSU 20-DIC-2027"%REV)
    cv.drawRightString(285*mm,7*mm,"Pag. %d"%doc.page);cv.restoreState()

doc=BaseDocTemplate(OUT_PDF,pagesize=PAGE,leftMargin=12*mm,rightMargin=12*mm,topMargin=12*mm,bottomMargin=12*mm)
fr=Frame(doc.leftMargin,doc.bottomMargin,doc.width,doc.height,id='f')
doc.addPageTemplates([PageTemplate(id='m',frames=[fr],onPage=footer)])
E=[]
E.append(Paragraph("Planning de Provisiones · PMS (ABB) + PCS/SCADA/SIS (Inauco)",H1))
E.append(Paragraph("Constructibilidad CPF2 La Calera II · Encadenamiento, puntos en común (FAT · integración · iFAT) y convergencia a precom/comisionado · Rev. %s · %s · RFSU 20-DIC-2027"%(REV,HOY.strftime('%d-%m-%Y')),SUB))
E.append(Spacer(1,3*mm))
E.append(Gantt(T,doc.width,doc.height-34*mm))
leg=[["■ PMS (ABB)","■ Inauco PCS/SCADA/SIS","■ Integración/Campo","◆ Hito / FAT","▭ borde rojo = crítico/FAT","Líneas: HOY · RFSU"]]
tl=Table(leg,colWidths=[doc.width/6]*6)
tl.setStyle(TableStyle([('FONTSIZE',(0,0),(-1,-1),7),
  ('TEXTCOLOR',(0,0),(0,0),colors.HexColor(C_PMS)),('TEXTCOLOR',(1,0),(1,0),colors.HexColor(C_INA)),
  ('TEXTCOLOR',(2,0),(2,0),colors.HexColor(C_INT)),('TEXTCOLOR',(3,0),(3,0),colors.HexColor(C_MILE)),
  ('TEXTCOLOR',(4,0),(4,0),colors.HexColor(C_CRIT))]))
E.append(tl); E.append(PageBreak())

# Pagina 2 - flujo
E.append(Paragraph("Diagrama de flujo — encadenamiento e integración de provisiones",H2))
E.append(Flow(doc.width,doc.height-20*mm)); E.append(PageBreak())

# Pagina 3 - planilla
E.append(Paragraph("Planilla de tareas e hitos",H2))
BL={'PMS':'PMS (ABB)','INA':'Inauco','INT':'Integración'}
data=[["Sist.","ID","Tarea / Hito","Inicio","Fin","Días","Nota"]]
for t in T:
    lane,tid,name,ini,fin,color,crit,note=t
    if tid=='##':
        data.append(['',' ',name,'','','',''])  # subheader
        continue
    if ini is None and fin is None: continue
    if fin is None: ini_s=d(ini).strftime('%d-%m-%y');fin_s='';dur='hito'
    else: ini_s=d(ini).strftime('%d-%m-%y');fin_s=d(fin).strftime('%d-%m-%y');dur=str((d(fin)-d(ini)).days+1)
    nm=(name+("  ◄ crítico" if crit and fin else "")).replace('&','&amp;')
    data.append([BL.get(lane,lane),tid,Paragraph(nm,CELL),ini_s,fin_s,dur,note])
tb=LongTable(data,colWidths=[16*mm,14*mm,108*mm,20*mm,20*mm,14*mm,81*mm],repeatRows=1)
sty=TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1f3b63')),('TEXTCOLOR',(0,0),(-1,0),colors.white),
  ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7.5),
  ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#ccc')),('VALIGN',(0,0),(-1,-1),'MIDDLE'),
  ('ALIGN',(3,0),(5,-1),'CENTER'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f4f4f4')])])
ri=1
for t in T:
    if t[1]=='##':
        sty.add('SPAN',(2,ri),(6,ri));sty.add('BACKGROUND',(0,ri),(-1,ri),colors.HexColor('#dce6f2'))
        sty.add('FONTNAME',(0,ri),(-1,ri),'Helvetica-Bold');ri+=1;continue
    if t[3] is None and t[4] is None: continue
    if t[6]: sty.add('TEXTCOLOR',(2,ri),(2,ri),colors.HexColor('#c00000'))
    ri+=1
tb.setStyle(sty);E.append(tb);E.append(PageBreak())

# Pagina 4 - notas
E.append(Paragraph("Puntos en común, supuestos y objetivo",H2))
def kv(title,rows):
    dd=[[title,""]]+[[Paragraph("<b>%s</b>"%k,CELL),Paragraph(v,CELL)] for k,v in rows]
    t=Table(dd,colWidths=[60*mm,213*mm])
    t.setStyle(TableStyle([('SPAN',(0,0),(1,0)),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1f3b63')),
      ('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),9),
      ('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#ccc')),('VALIGN',(0,0),(-1,-1),'TOP'),
      ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f4f4f4')])]))
    return t
E.append(kv("PUNTOS EN COMÚN (deben comunicarse, probarse e integrarse)",[
  ("FAT individuales","FAT PMS (25-29-ENE-27) · FAT PCS+Comunicaciones (27-ENE→19-FEB) · FAT SIS (27-ENE→23-FEB). Cada provisión valida en fábrica antes de integrar."),
  ("Integración PCS↔PMS","Pruebas en ABB Bs.As. (24-FEB→02-MAR-27): se prueba la comunicación PCS (Inauco) ↔ PMS (ABB) por Modbus TCP/IP. Punto de convergencia temprano."),
  ("iFAT (prueba integrada)","Integración SE#3+SE#4+PMS+PCS+SIS (25-JUN→14-JUL-27) con todos los vendors (ABB+Inauco+HIMA+AESA). Valida el conjunto antes del comisionado."),
  ("Campo / convergencia","SAT PCS (17-MAR→01-JUN) · SAT SCADA-SIS (17-MAR→03-AGO) · MCE (04-24-AGO) · SAT PMS (06-OCT→05-DIC) → Comisionado integrado (23-SEP→19-DIC) → RFSU."),
]))
E.append(Spacer(1,3*mm))
E.append(kv("SUPUESTOS Y OBJETIVO",[
  ("Inauco — arranque JUL-2026","La OC Inauco aún no se emitió; el planning se ancla en INICIO 01-JUL-2026 (OC + ingeniería Rev0 + recepción tableros HIMA). Si la OC slipa, todo el encadenamiento Inauco se desplaza."),
  ("PMS — base","Cronograma ABB actualizado 09-JUN-2026 (OC 4508945953). FAT PMS 29-ENE-27, FAT Integral Shelter 31-MAR-27, SAT PMS en sitio OCT-DIC-27 (HOLD POINT)."),
  ("RFSU objetivo","20-DIC-2027 (planning integrado). El cronograma ABB contempla RFSU contractual 31-ENE-2028 como respaldo (~6 semanas de margen)."),
  ("Riesgo de convergencia","La integración PCS↔PMS depende de FAT PMS y FAT PCS/SIS completados; el iFAT depende de salas en sitio (SE#3 14-MAY) y PMS integrado. Cualquier atraso de una provisión arrastra a la otra."),
]))
doc.build(E)
print("OK PDF ->",OUT_PDF)

# ---------------- XLSX ----------------
import openpyxl
from openpyxl.styles import Font,PatternFill
wb=openpyxl.Workbook();ws=wb.active;ws.title="Planning_PMS_Inauco"
ws.append(["Sistema","ID","Tarea / Hito","Inicio","Fin","Días","Crítico","Nota"])
for c in ws[1]: c.font=Font(bold=True,color="FFFFFF");c.fill=PatternFill("solid",fgColor="1F3B63")
for t in T:
    lane,tid,name,ini,fin,color,crit,note=t
    if tid=='##':
        ws.append([name,"","","","","","",""])
        for c in ws[ws.max_row]: c.font=Font(bold=True)
        continue
    if ini is None and fin is None: continue
    dur=((d(fin)-d(ini)).days+1) if fin else ""
    ws.append([BL.get(lane,lane),tid,name,d(ini) if ini else "",d(fin) if fin else "",dur,"SÍ" if crit else "",note])
for col,w in zip("ABCDEFGH",[12,8,52,12,12,7,8,42]): ws.column_dimensions[col].width=w
wb.save(OUT_XLS)
print("OK XLSX ->",OUT_XLS,"| tareas:",sum(1 for t in T if t[1]!='##' and not(t[3] is None and t[4] is None)))
