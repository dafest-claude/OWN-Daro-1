#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cronograma REALISTA CPF2 La Calera II - EL + IN
Escenario conservador anclado en el PLAN DE SUMINISTROS (18-06-26) y el TRACKER (19-06-26):
 - Entregas de equipo crítico hasta ~20-jul-2027 (skids combustible es la más tardía).
 - Calendario 5 d/sem · 8 h. Recursos acotados (escenario conservador).
 - Precomisionado y Comisionado SECUENCIALES (sin el solapamiento forzado del programa imagen).
 - Planta completa (RFSU único).
Objetivo: estimar la fecha REALISTA de RFSU vs la base contractual 31-ENE-2028.
Salidas: Cronograma_Realista_CPF2.pdf (Gantt + tabla + notas) y .xlsx
"""
import os
from datetime import date
from collections import OrderedDict

BASE=os.path.dirname(os.path.abspath(__file__))
OUT_PDF=os.path.join(BASE,"Cronograma_Realista_CPF2.pdf")
OUT_XLS=os.path.join(BASE,"Cronograma_Realista_CPF2.xlsx")
HOY=date(2026,6,19)
RFSU_BASE=date(2028,1,31)      # contractual
RFSU_REAL=date(2028,5,31)      # estimado realista
T0=date(2027,1,1); T1=date(2028,9,30)
def d(s): y,m,dd=map(int,s.split('-')); return date(y,m,dd)

C_PROC='#bf9000'; C_IN='#2e75b6'; C_EL='#c55a11'; C_TEST='#548235'; C_COM='#7030a0'; C_CRIT='#c00000'

# (bloque,id,nombre,inicio,fin|None,color,critico,nota)
T=[
 ('H','##','SUMINISTROS / ENTREGAS DE EQUIPO CRÍTICO (Plan 18-06-26)',None,None,None,False,''),
 ('H','S1','▼ Cables IN en obra','2027-03-29',None,C_PROC,False,'plan (proyecto pedía dic-26)'),
 ('H','S2','▼ SIS (HIMA)','2027-04-18',None,C_PROC,False,'entrega'),
 ('H','S3','▼ Cables EL en obra','2027-04-26',None,C_PROC,False,'plan (proyecto pedía dic-26)'),
 ('H','S4','▼ Sala SE#3 en sitio','2027-05-24',None,C_PROC,False,'ABB'),
 ('H','S5','▼ Válvulas de control','2027-05-31',None,C_PROC,False,'entrega'),
 ('H','S6','▼ Sistema PMS (ABB)','2027-06-02',None,C_PROC,False,'entrega'),
 ('H','S7','▼ Sala SE#4 en sitio','2027-06-23',None,C_PROC,False,'ABB'),
 ('H','S8','★ Sistema de Control PCS (Inauco)','2027-06-24',None,C_CRIT,True,'entrega · ancla control'),
 ('H','S9','▼ Bombas / calentadores / desgasificadores','2027-06-29',None,C_PROC,False,'entrega'),
 ('H','S10','★ Skids de combustible','2027-07-20',None,C_CRIT,True,'ÚLTIMA entrega crítica'),

 ('IN','##','INSTRUMENTACIÓN (IN)',None,None,None,False,''),
 ('IN','I1','Canalizaciones / conduit IN','2027-02-01','2027-04-15',C_IN,False,''),
 ('IN','I2','Tendido cables IN','2027-04-05','2027-08-31',C_IN,True,'~65.000 m · recursos acotados'),
 ('IN','I3','Montaje de instrumentos (AESA)','2027-06-01','2027-11-30',C_IN,False,'desde entrega válvulas/instr.'),
 ('IN','I4','Conexionado IN (campo→JB→Sala INS)','2027-07-15','2027-12-23',C_IN,True,'~15.752 puntas'),
 ('IN','I5','Loop check IN (PCS/SIS)','2027-11-01','2028-02-15',C_IN,True,'tras PCS instalado'),

 ('EL','##','ELECTRICIDAD (EL)',None,None,None,False,''),
 ('EL','E1','Bandejas / canalizaciones EL','2027-03-01','2027-05-15',C_EL,False,''),
 ('EL','E2','Tendido cables EL','2027-05-01','2027-09-30',C_EL,True,'~81.231 m · recursos acotados'),
 ('EL','E3','Armado salas SE#3 / SE#4 (campo)','2027-05-24','2027-07-21',C_EL,False,'módulos → armado'),
 ('EL','E4','Montaje tableros / celdas / CCMs','2027-07-01','2027-09-15',C_EL,False,'37 tableros'),
 ('EL','E5','Montaje motores / equipos de proceso','2027-08-01','2027-10-31',C_EL,False,'tras entregas jun-jul'),
 ('EL','E6','Conexionado EL (sectores SE#3 / SE#4)','2027-06-22','2027-11-30',C_EL,True,'~5.178 puntas'),
 ('EL','E7','Pruebas Megger EL (HOLD POINT)','2027-11-01','2027-11-30',C_EL,True,'circuito por circuito'),

 ('T','##','ENERGIZACIÓN · PRECOMISIONADO · COMISIONADO (secuencial)',None,None,None,False,''),
 ('T','P1','Energización progresiva MT→BT','2027-12-01','2027-12-31',C_TEST,True,'HOLD POINT'),
 ('T','P2','Precomisionado eléctrico funcional','2027-12-15','2028-02-15',C_TEST,True,'loop check EL + relés'),
 ('T','P3','Precomisionado instrumentación (PCS/SIS)','2027-12-15','2028-02-29',C_TEST,True,'lazos completos'),
 ('T','P4','Comisionado integrado (EL+PCS+PMS+SIS)','2028-02-15','2028-05-15',C_COM,True,'secuencial tras precom'),
 ('T','P5','SAT PMS (HOLD POINT)','2028-03-01','2028-04-30',C_COM,False,'ventana'),
 ('T','RBASE','◇ RFSU BASE (contractual)','2028-01-31',None,C_CRIT,False,'objetivo previo (forzado)'),
 ('T','RFSU','★★★ RFSU REALISTA (estimado)','2028-05-31',None,C_CRIT,True,'≈ +4 meses'),
]

# ---------------- PDF ----------------
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (BaseDocTemplate,PageTemplate,Frame,Paragraph,Spacer,
        Table,TableStyle,LongTable,Flowable,PageBreak)
from reportlab.lib.styles import getSampleStyleSheet,ParagraphStyle
from reportlab.lib.enums import TA_CENTER
PAGE=landscape(A4); st=getSampleStyleSheet()
H1=ParagraphStyle('H1',parent=st['Title'],fontSize=15,spaceAfter=2)
SUB=ParagraphStyle('SUB',parent=st['Normal'],fontSize=9,textColor=colors.HexColor('#555'),alignment=TA_CENTER)
H2=ParagraphStyle('H2',parent=st['Heading2'],fontSize=11,textColor=colors.HexColor('#1f3b63'),spaceBefore=6,spaceAfter=4)
N=ParagraphStyle('N',parent=st['Normal'],fontSize=8.5,leading=11)
CELL=ParagraphStyle('CELL',parent=st['Normal'],fontSize=7.5,leading=9)
MES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
def months(a,b):
    o=[];y,m=a.year,a.month
    while (y,m)<=(b.year,b.month):
        o.append(date(y,m,1)); m+=1
        if m>12:m=1;y+=1
    return o
class Gantt(Flowable):
    def __init__(s,tasks,w,h,lw=86*mm): s.t=tasks;s.w=w;s.h=h;s.lw=lw;Flowable.__init__(s)
    def wrap(s,a,b): return (s.w,s.h)
    def x(s,dt): return s.lw+(s.w-s.lw)*((dt-T0).days)/((T1-T0).days)
    def draw(s):
        c=s.canv;W=s.w;H=s.h;lw=s.lw;head=10*mm;n=len(s.t);rh=(H-head)/n;ly=None
        for mdt in months(T0,T1):
            xx=s.x(mdt); c.setStrokeColor(colors.HexColor('#e8e8e8'));c.setLineWidth(0.3);c.line(xx,0,xx,H-head)
            nxt=date(mdt.year+(mdt.month//12),(mdt.month%12)+1,1)
            c.setFillColor(colors.HexColor('#666'));c.setFont('Helvetica',5);c.drawCentredString((xx+s.x(nxt))/2,H-head+1.5,MES[mdt.month-1])
            if mdt.year!=ly:
                c.setFillColor(colors.HexColor('#1f3b63'));c.setFont('Helvetica-Bold',7);c.drawString(xx+1,H-head+5.3,str(mdt.year));ly=mdt.year
                c.setStrokeColor(colors.HexColor('#b0b0b0'));c.setLineWidth(0.6);c.line(xx,0,xx,H-head+5)
        c.setStrokeColor(colors.HexColor('#bbb'));c.setLineWidth(0.5);c.rect(lw,0,W-lw,H-head,stroke=1,fill=0)
        for dt,col,lbl in [(HOY,'#1f9d55','HOY'),(RFSU_BASE,'#999999','RFSU base 31-ENE-28'),(RFSU_REAL,'#c00000','RFSU real 31-MAY-28')]:
            xx=s.x(dt);c.setStrokeColor(colors.HexColor(col));c.setLineWidth(1);c.setDash(2,2);c.line(xx,0,xx,H-head);c.setDash()
            c.saveState();c.setFillColor(colors.HexColor(col));c.setFont('Helvetica-Bold',5.2);c.translate(xx,H-head-1);c.rotate(90);c.drawString(2,1.2,lbl);c.restoreState()
        for i,t in enumerate(s.t):
            blk,tid,name,ini,fin,color,crit,note=t; y=H-head-(i+1)*rh
            if tid=='##':
                c.setFillColor(colors.HexColor('#1f3b63'));c.rect(0,y,W,rh,stroke=0,fill=1)
                c.setFillColor(colors.white);c.setFont('Helvetica-Bold',6.6);c.drawString(2,y+rh*0.28,name);continue
            if i%2==0: c.setFillColor(colors.HexColor('#f6f6f6'));c.rect(0,y,W,rh,stroke=0,fill=1)
            c.setFillColor(colors.black);c.setFont('Helvetica',5.8);c.drawString(3,y+rh*0.28,name[:58])
            c.setFillColor(colors.HexColor('#999'));c.setFont('Helvetica',4.8);c.drawRightString(lw-2,y+rh*0.28,note[:26])
            if fin is None and ini:
                xx=s.x(d(ini));sz=rh*0.30;c.setFillColor(colors.HexColor(color))
                c.saveState();c.translate(xx,y+rh/2);c.rotate(45);c.rect(-sz,-sz,2*sz,2*sz,stroke=0,fill=1);c.restoreState()
            elif ini and fin:
                x0=s.x(d(ini));x1=s.x(d(fin));bh=rh*0.5;yy=y+(rh-bh)/2
                c.setFillColor(colors.HexColor(color));c.roundRect(x0,yy,max(x1-x0,1),bh,1,stroke=0,fill=1)
                if crit: c.setStrokeColor(colors.HexColor('#c00000'));c.setLineWidth(0.8);c.roundRect(x0,yy,max(x1-x0,1),bh,1,stroke=1,fill=0)
def footer(cv,doc):
    cv.saveState();cv.setFont('Helvetica',7);cv.setFillColor(colors.HexColor('#888'))
    cv.drawString(12*mm,7*mm,"Constructibilidad CPF2 · Cronograma REALISTA EL+IN · Rev.0 · base RFSU 31-ENE-28 → estimado 31-MAY-28")
    cv.drawRightString(285*mm,7*mm,"Pag. %d"%doc.page);cv.restoreState()
doc=BaseDocTemplate(OUT_PDF,pagesize=PAGE,leftMargin=12*mm,rightMargin=12*mm,topMargin=12*mm,bottomMargin=12*mm)
fr=Frame(doc.leftMargin,doc.bottomMargin,doc.width,doc.height,id='f')
doc.addPageTemplates([PageTemplate(id='m',frames=[fr],onPage=footer)])
E=[]
E.append(Paragraph("Cronograma REALISTA — Instrumentación y Electricidad",H1))
E.append(Paragraph("CPF2 La Calera II · escenario conservador anclado en el Plan de Suministros (18-06-26) · "
                   "calendario 5 d/sem · 8 h · precom→comisionado secuencial · RFSU base 31-ENE-2028 → "
                   "<b>estimado realista 31-MAY-2028 (≈ +4 meses)</b>",SUB))
E.append(Spacer(1,3*mm))
E.append(Gantt(T,doc.width,doc.height-34*mm))
leg=[["■ Suministros/entrega","■ Instrumentación","■ Electricidad","■ Energiz./Precom","■ Comisionado","▭ rojo=crítico · ◇ RFSU base"]]
tl=Table(leg,colWidths=[doc.width/6]*6)
tl.setStyle(TableStyle([('FONTSIZE',(0,0),(-1,-1),7),
  ('TEXTCOLOR',(0,0),(0,0),colors.HexColor(C_PROC)),('TEXTCOLOR',(1,0),(1,0),colors.HexColor(C_IN)),
  ('TEXTCOLOR',(2,0),(2,0),colors.HexColor(C_EL)),('TEXTCOLOR',(3,0),(3,0),colors.HexColor(C_TEST)),
  ('TEXTCOLOR',(4,0),(4,0),colors.HexColor(C_COM)),('TEXTCOLOR',(5,0),(5,0),colors.HexColor(C_CRIT))]))
E.append(tl); E.append(PageBreak())

# tabla de tareas
E.append(Paragraph("Detalle de tareas e hitos",H2))
BL={'H':'Suministro','IN':'Instrum.','EL':'Electric.','T':'Energ/Comis'}
data=[["Bloque","ID","Tarea / Hito","Inicio","Fin","Días","Nota"]]
for t in T:
    blk,tid,name,ini,fin,color,crit,note=t
    if tid=='##':
        data.append(['','',name,'','','','']);continue
    if fin is None: ini_s=d(ini).strftime('%d-%m-%y');fin_s='';dur='hito'
    else: ini_s=d(ini).strftime('%d-%m-%y');fin_s=d(fin).strftime('%d-%m-%y');dur=str((d(fin)-d(ini)).days+1)
    nm=(name+("  ◄ crítico" if crit and fin else "")).replace('&','&amp;')
    data.append([BL.get(blk,blk),tid,Paragraph(nm,CELL),ini_s,fin_s,dur,note])
tb=LongTable(data,colWidths=[18*mm,12*mm,108*mm,20*mm,20*mm,14*mm,81*mm],repeatRows=1)
sty=TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1f3b63')),('TEXTCOLOR',(0,0),(-1,0),colors.white),
  ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),7.5),('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#ccc')),
  ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(3,0),(5,-1),'CENTER'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f4f4f4')])])
ri=1
for t in T:
    if t[1]=='##':
        sty.add('SPAN',(2,ri),(6,ri));sty.add('BACKGROUND',(0,ri),(-1,ri),colors.HexColor('#dce6f2'));sty.add('FONTNAME',(0,ri),(-1,ri),'Helvetica-Bold');ri+=1;continue
    if t[6]: sty.add('TEXTCOLOR',(2,ri),(2,ri),colors.HexColor('#c00000'))
    ri+=1
tb.setStyle(sty);E.append(tb);E.append(PageBreak())

# notas
E.append(Paragraph("Premisas, cadena crítica y conclusión",H2))
def kv(title,rows):
    dd=[[title,""]]+[[Paragraph("<b>%s</b>"%k,CELL),Paragraph(v,CELL)] for k,v in rows]
    t=Table(dd,colWidths=[55*mm,218*mm])
    t.setStyle(TableStyle([('SPAN',(0,0),(1,0)),('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1f3b63')),('TEXTCOLOR',(0,0),(-1,0),colors.white),
      ('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),9),('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#ccc')),
      ('VALIGN',(0,0),(-1,-1),'TOP'),('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f4f4f4')])]))
    return t
E.append(kv("PREMISAS DEL ESCENARIO REALISTA",[
  ("Anclaje en suministros","Fechas de entrega del Plan de Suministros (18-06-26) y Tracker v7 (19-06-26). Última entrega crítica: skids de combustible 20-jul-2027."),
  ("Calendario","5 días/semana · 8 h efectivas (conservador). Sin doble turno."),
  ("Recursos","Acotados / a confirmar → durations conservadoras (sin sumar cuadrillas extra)."),
  ("Secuencia precom→comisionado","SECUENCIAL: el comisionado integrado arranca al completar el precomisionado (no se fuerza el solapamiento del programa-imagen)."),
  ("Alcance","RFSU único de planta completa."),
]))
E.append(Spacer(1,3*mm))
E.append(kv("CADENA CRÍTICA Y CONCLUSIÓN",[
  ("Cadena crítica","Entregas (skids 20-jul / PCS 24-jun) → montaje equipos → conexionado EL/IN → megger → energización (dic-27) → precom (dic-27→feb-28) → comisionado integrado (feb→may-28) → RFSU."),
  ("RFSU realista","≈ 31-MAY-2028, ~4 meses después de la base contractual 31-ENE-2028. El 31-ENE-2028 sólo se sostiene forzando el solapamiento precom+comisionado."),
  ("Desfasaje de cables","El proyecto pide cables en obra dic-26, pero el plan muestra entrega mar-abr-27. Si se confirma esa demora, el tendido se corre y agrava la cadena."),
  ("Palancas para recuperar","(1) doble turno / +cuadrillas en tendido y conexionado; (2) adelantar OC de cables y skids; (3) solapar precom-comisionado por subsistema; (4) RFSU por fases."),
  ("Estado","Preliminar para depuración. Ajustable al confirmar recursos, calendario y fechas firmes de entrega."),
]))
doc.build(E)
print("OK PDF ->",OUT_PDF)

# ---------------- XLSX ----------------
import openpyxl
from openpyxl.styles import Font,PatternFill
wb=openpyxl.Workbook();ws=wb.active;ws.title="Cronograma_Realista"
ws.append(["Bloque","ID","Tarea / Hito","Inicio","Fin","Días","Crítico","Nota"])
for c in ws[1]: c.font=Font(bold=True,color="FFFFFF");c.fill=PatternFill("solid",fgColor="1F3B63")
for t in T:
    blk,tid,name,ini,fin,color,crit,note=t
    if tid=='##':
        ws.append([name,"","","","","","",""])
        for c in ws[ws.max_row]: c.font=Font(bold=True)
        continue
    dur=((d(fin)-d(ini)).days+1) if fin else ""
    ws.append([BL.get(blk,blk),tid,name,d(ini) if ini else "",d(fin) if fin else "",dur,"SÍ" if crit else "",note])
for col,w in zip("ABCDEFGH",[12,8,52,12,12,7,8,46]): ws.column_dimensions[col].width=w
wb.save(OUT_XLS)
print("OK XLSX ->",OUT_XLS)
