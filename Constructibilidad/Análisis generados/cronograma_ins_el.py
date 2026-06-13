#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cronograma de Construccion - Instrumentacion y Electricidad - CPF2 La Calera II
Construye un cronograma (Gantt) detallado de la ejecucion de campo EL + IN,
partiendo del tendido de cables, basado en los volumenes y rendimientos de:
  - Gantt_Integrado_CPF2_Planning2027_Rev1 (hitos, secuencia, restricciones)
  - Analisis_Electrico_CPF2_Campo_Rev3 (513 cables / 81.231 m / ~5.178 puntas)
  - Analisis_Cables_Instrumentacion_Rev0 (885 cables / 65.315 m / 15.752 puntas)
  - Listas de Instrumentos 00102/00670 (2.798 instr; AESA monta ~1.163)
Hito gobernante: RFSU 20-DIC-2027.
Salidas: PDF (Gantt + tablas) y XLSX (lista de tareas).
"""
import os
from datetime import date, timedelta
from collections import OrderedDict

BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
IN_DIR = os.path.join(ROOT, "Datos entrada")
THREED = os.path.join(IN_DIR, "3D LC 120626.pdf")
OUT_PDF = os.path.join(BASE, "Cronograma_INS_EL_CPF2.pdf")
OUT_XLS = os.path.join(BASE, "Cronograma_INS_EL_CPF2.xlsx")
REV = "1"

HOY   = date(2026, 6, 13)
RFSU  = date(2027, 12, 20)
T0    = date(2026, 6, 1)     # inicio timeline
T1    = date(2028, 1, 15)    # fin timeline

def d(s):
    y,m,dd = map(int, s.split('-')); return date(y,m,dd)

# Colores por bloque
C_PROV = '#7f7f7f'; C_IN='#2e75b6'; C_EL='#c55a11'; C_COM='#548235'
C_CRIT = '#c00000'; C_MILE='#bf9000'

# ---- Definicion de tareas ----
# (bloque, id, nombre, inicio, fin|None(milestone), color, critico, base/cantidad)
TASKS = [
 ('H','##','PROVISIÓN Y ENTREGAS DE MATERIALES (hitos ancla)',None,None,None,False,''),
 ('H','M1','★ OC Cables IN (target)','2026-08-01',None,C_MILE,False,'LT 150 d'),
 ('H','M2','★ OC Cables EL (target)','2026-10-01',None,C_MILE,False,'LT 150 d'),
 ('H','M3','▼ Llegada Cables IN a sitio','2026-11-28',None,C_MILE,False,'885 c / 65.315 m'),
 ('H','M4a','▼ Entrega instrum. Fase 1 (válvulas + en línea)','2026-12-15',None,C_MILE,False,'920 en línea (válv 485+elem 435)'),
 ('H','M4b','▼ Entrega instrum. Fase 2 (resto, progresiva)','2027-02-15',None,C_MILE,False,'1.878 montaje específico'),
 ('H','M5','▼ Llegada Cables EL a sitio','2027-02-28',None,C_MILE,False,'513 c / 81.231 m'),
 ('H','M6','★ Inicio campo PCS — Sala INS (Sala 7)','2027-03-17',None,C_CRIT,True,'ancla IN'),
 ('H','M7','★ Llegada SE#4 (con PMS-101)','2027-04-08',None,C_MILE,False,'ancla EL'),
 ('H','M8','★ Llegada SE#3 (con PMS-001)','2027-05-14',None,C_CRIT,True,'ancla EL · ruta crítica'),

 ('IN','##','INSTRUMENTACIÓN (IN) — PCS / SIS / F&G',None,None,None,False,''),
 ('IN','I1','Canalizaciones / conduit instrumentación','2026-11-28','2027-01-31',C_IN,False,'bandeja+conduit IN'),
 ('IN','I2','Tendido cables IN — señal 1P/1T (≈609 c)','2027-02-01','2027-02-28',C_IN,True,'~31.600 m'),
 ('IN','I3','Tendido cables IN — multipar/terna (≈260 c)','2027-02-10','2027-03-14',C_IN,True,'~30.000 m · 5-6 cuad.'),
 ('IN','I4','Tendido cables IN — FO / FTP (16 c)','2027-03-01','2027-03-14',C_IN,False,'~3.300 m'),
 ('IN','I5a','Recepción + banco Fase 1 (válvulas/PSV + en línea)','2026-12-15','2027-02-28',C_IN,False,'en línea · pre-montaje'),
 ('IN','I6a','Montaje instrum. en línea — AESA (con cañería)','2027-01-15','2027-04-30',C_IN,False,'válvulas + elementos'),
 ('IN','I5b','Recepción + banco Fase 2 (resto instrumentos)','2027-02-15','2027-05-15',C_IN,False,'montaje específico'),
 ('IN','I6b','Montaje instrum. montaje específico — AESA','2027-03-15','2027-06-30',C_IN,False,'transmisores/sw/F&G'),
 ('IN','I7','Conexionado IN — campo→JB→Sala INS','2027-03-17','2027-07-31',C_IN,True,'15.752 puntas'),
 ('IN','I8','Loop check IN (sensor→JB→PCS/SIS)','2027-06-01','2027-08-20',C_IN,False,'~885 lazos'),

 ('EL','##','ELECTRICIDAD (EL) — Potencia / Control',None,None,None,False,''),
 ('EL','E1','Bandejas portacables principales','2027-03-01','2027-04-05',C_EL,False,'MC1/MC2/MC3'),
 ('EL','E2','Tendido EL — Potencia BT Grande ≥35mm²','2027-03-01','2027-04-18',C_CRIT,True,'104 c / 17.470 m ★'),
 ('EL','E3','Tendido EL — Potencia BT Med/Peq','2027-03-15','2027-04-26',C_EL,False,'198 c / 33.125 m'),
 ('EL','E4','Tendido EL — Control y Señales','2027-03-15','2027-05-13',C_EL,False,'192 c / 28.126 m'),
 ('EL','E5','Tendido EL — Potencia MT 6,6/13,2kV','2027-05-10','2027-06-07',C_EL,False,'18 c / 2.110 m · cert. ABB'),
 ('EL','E6','Armado shelter SE#4 (módulos en campo)','2027-04-08','2027-04-22',C_EL,False,'14 d'),
 ('EL','E7','Montaje CCMs / celdas MT / motores','2027-04-14','2027-05-24',C_EL,False,'37 tabl · 79 motores'),
 ('EL','E8','Armado shelter SE#3 (módulos en campo)','2027-05-14','2027-06-03',C_CRIT,True,'21 d · ruta crítica'),
 ('EL','E9','Conexionado EL — Sector SE#4','2027-04-23','2027-06-12',C_EL,False,'puntas BT+control S4'),
 ('EL','E10','Conexionado EL — Terminaciones MT (36 ext.)','2027-06-07','2027-07-05',C_EL,False,'técnicos cert. ABB'),
 ('EL','E11','Conexionado EL — Sector SE#3','2027-06-05','2027-07-25',C_CRIT,True,'puntas S3 · ruta crítica'),

 ('C','##','PRUEBAS · PRECOMISIONADO · COMISIONADO',None,None,None,False,''),
 ('C','P1','Pruebas Megger EL circuito×circuito','2027-07-26','2027-08-15',C_COM,True,'513 c · HOLD POINT'),
 ('C','P2','Energización progresiva MT→Transf→BT','2027-08-16','2027-09-01',C_COM,True,'HOLD POINT'),
 ('C','P3','Precomisionado eléctrico funcional','2027-09-02','2027-09-22',C_COM,True,'loop check EL + relés'),
 ('C','P4','Precomisionado instrumentación (lazos PCS/SIS)','2027-09-01','2027-09-30',C_COM,False,'loop check final IN'),
 ('C','P5','SAT/Comisionado PCS — Sala INS (Inauco)','2027-03-17','2027-06-01',C_COM,False,'55 d'),
 ('C','P6','SAT/Comisionado SIS — HIMA','2027-03-17','2027-08-03',C_COM,False,'100 d'),
 ('C','P7','Comisionado integrado (EL+PCS+PMS+SIS)','2027-09-23','2027-12-19',C_COM,True,'88 d'),
 ('C','P8','SAT PMS (HOLD POINT contractual)','2027-10-06','2027-12-05',C_COM,False,'ventana fija OCT-DIC'),
 ('C','RFSU','★★★ RFSU — READY FOR START UP','2027-12-20',None,C_CRIT,True,'hito gobernante'),
]

# ---------------- PDF ----------------
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.platypus import (BaseDocTemplate, PageTemplate, Frame, Paragraph,
        Spacer, Table, TableStyle, LongTable, Flowable, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER

PAGE = landscape(A4)            # 297 x 210 mm
styles = getSampleStyleSheet()
H1=ParagraphStyle('H1',parent=styles['Title'],fontSize=15,spaceAfter=2)
SUB=ParagraphStyle('SUB',parent=styles['Normal'],fontSize=9,
                   textColor=colors.HexColor('#555555'),alignment=TA_CENTER)
H2=ParagraphStyle('H2',parent=styles['Heading2'],fontSize=11,
                  textColor=colors.HexColor('#1f3b63'),spaceBefore=8,spaceAfter=4)
N=ParagraphStyle('N',parent=styles['Normal'],fontSize=8.5,leading=11)
CELL=ParagraphStyle('CELL',parent=styles['Normal'],fontSize=7,leading=8.5)

MES_ABBR=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

def months_between(a,b):
    out=[]; y,m=a.year,a.month
    while (y,m)<=(b.year,b.month):
        out.append(date(y,m,1))
        m+=1
        if m>12: m=1; y+=1
    return out

class Gantt(Flowable):
    def __init__(self, tasks, t0, t1, width, height, lw=82*mm):
        Flowable.__init__(self); self.tasks=tasks; self.t0=t0; self.t1=t1
        self.width=width; self.height=height; self.lw=lw
    def wrap(self,aw,ah): return (self.width,self.height)
    def x(self,dt):
        span=(self.t1-self.t0).days
        return self.lw + (self.width-self.lw)*((dt-self.t0).days)/span
    def draw(self):
        c=self.canv; W=self.width; H=self.height; lw=self.lw
        head=11*mm
        n=len(self.tasks)
        rh=(H-head)/n
        # --- grid de meses ---
        months=months_between(self.t0,self.t1)
        c.setFont('Helvetica',5.5)
        last_year=None
        for mdt in months:
            xx=self.x(mdt)
            c.setStrokeColor(colors.HexColor('#e0e0e0')); c.setLineWidth(0.3)
            c.line(xx, 0, xx, H-head)
            c.setFillColor(colors.HexColor('#666666'))
            c.drawCentredString(xx+ (self.x(date(mdt.year+(mdt.month//12), (mdt.month%12)+1,1))-xx)/2,
                                 H-head+1.5, MES_ABBR[mdt.month-1])
            if mdt.year!=last_year:
                c.setFillColor(colors.HexColor('#1f3b63')); c.setFont('Helvetica-Bold',7)
                c.drawString(xx+1, H-head+5.5, str(mdt.year)); c.setFont('Helvetica',5.5)
                last_year=mdt.year
                c.setStrokeColor(colors.HexColor('#b0b0b0')); c.setLineWidth(0.6)
                c.line(xx,0,xx,H-head+5)
        # marco timeline
        c.setStrokeColor(colors.HexColor('#bbbbbb')); c.setLineWidth(0.5)
        c.rect(lw,0,W-lw,H-head, stroke=1, fill=0)
        # HOY y RFSU
        for dt,col,lbl in [(HOY,'#1f9d55','HOY'),(RFSU,'#c00000','RFSU 20-DIC-27')]:
            xx=self.x(dt); c.setStrokeColor(colors.HexColor(col)); c.setLineWidth(1)
            c.setDash(2,2); c.line(xx,0,xx,H-head); c.setDash()
            c.setFillColor(colors.HexColor(col)); c.setFont('Helvetica-Bold',5.5)
            c.saveState(); c.translate(xx,H-head-1); c.rotate(90)
            c.drawString(2,1.2,lbl); c.restoreState()
        # --- filas ---
        for i,t in enumerate(self.tasks):
            blk,tid,name,ini,fin,color,crit,base=t
            y=H-head-(i+1)*rh
            is_hdr = (tid=='##')
            if is_hdr:
                c.setFillColor(colors.HexColor('#1f3b63'))
                c.rect(0,y,W,rh,stroke=0,fill=1)
                c.setFillColor(colors.white); c.setFont('Helvetica-Bold',6.8)
                c.drawString(2, y+rh*0.28, name);
                continue
            # zebra
            if i%2==0:
                c.setFillColor(colors.HexColor('#f6f6f6')); c.rect(0,y,W,rh,stroke=0,fill=1)
            # etiqueta
            c.setFillColor(colors.black); c.setFont('Helvetica',6)
            label=name if len(name)<58 else name[:57]
            c.drawString(3,y+rh*0.28,label)
            c.setFillColor(colors.HexColor('#888888')); c.setFont('Helvetica',5)
            c.drawRightString(lw-2,y+rh*0.28, base[:26])
            # barra o hito
            if fin is None and ini:    # milestone
                xx=self.x(d(ini)); s=rh*0.32
                c.setFillColor(colors.HexColor(color))
                c.saveState(); c.translate(xx,y+rh/2); c.rotate(45)
                c.rect(-s,-s,2*s,2*s,stroke=0,fill=1); c.restoreState()
            elif ini and fin:
                x0=self.x(d(ini)); x1=self.x(d(fin)); bh=rh*0.52
                yy=y+(rh-bh)/2
                c.setFillColor(colors.HexColor(color))
                c.roundRect(x0,yy,max(x1-x0,1.0),bh,1,stroke=0,fill=1)
                if crit:
                    c.setStrokeColor(colors.HexColor('#c00000')); c.setLineWidth(0.8)
                    c.roundRect(x0,yy,max(x1-x0,1.0),bh,1,stroke=1,fill=0)

def footer(canvas,doc):
    canvas.saveState(); canvas.setFont('Helvetica',7)
    canvas.setFillColor(colors.HexColor('#888888'))
    canvas.drawString(12*mm,7*mm,"Constructibilidad CPF2 — Cronograma INS+EL  |  Rev. %s  |  RFSU 20-DIC-2027"%REV)
    canvas.drawRightString(285*mm,7*mm,"Pag. %d"%doc.page)
    canvas.restoreState()

doc=BaseDocTemplate(OUT_PDF,pagesize=PAGE,leftMargin=12*mm,rightMargin=12*mm,
                    topMargin=12*mm,bottomMargin=12*mm,
                    title="Cronograma INS+EL CPF2 - Rev.0")
frame=Frame(doc.leftMargin,doc.bottomMargin,doc.width,doc.height,id='f')
doc.addPageTemplates([PageTemplate(id='main',frames=[frame],onPage=footer)])

E=[]
E.append(Paragraph("Cronograma de Construcción — Instrumentación y Electricidad", H1))
E.append(Paragraph("Constructibilidad · Proyecto LA CALERA CPF2 - FASE 2 (Vaca Muerta) · "
                   "Rev. %s · %s · Hito gobernante RFSU 20-DIC-2027"%(REV,HOY.strftime('%d-%m-%Y')), SUB))
E.append(Spacer(1,3*mm))
# Gantt ocupa el resto de la pagina
g_h = doc.height - 36*mm
E.append(Gantt(TASKS, T0, T1, doc.width, g_h))
# leyenda
leg=[["■ Provisión/hito","■ Instrumentación (IN)","■ Electricidad (EL)",
      "■ Pruebas/Comis.","◆ Hito","▭ borde rojo = ruta crítica"]]
tl=Table(leg,colWidths=[doc.width/6]*6)
tl.setStyle(TableStyle([('FONTSIZE',(0,0),(-1,-1),7),
   ('TEXTCOLOR',(0,0),(0,0),colors.HexColor(C_MILE)),
   ('TEXTCOLOR',(1,0),(1,0),colors.HexColor(C_IN)),
   ('TEXTCOLOR',(2,0),(2,0),colors.HexColor(C_EL)),
   ('TEXTCOLOR',(3,0),(3,0),colors.HexColor(C_COM)),
   ('TEXTCOLOR',(4,0),(4,0),colors.HexColor(C_MILE)),
   ('TEXTCOLOR',(5,0),(5,0),colors.HexColor(C_CRIT))]))
E.append(tl)
E.append(PageBreak())

# ---- Pagina 2: tabla detallada de tareas ----
E.append(Paragraph("Detalle de tareas e hitos", H2))
BLK={'H':'Provisión','IN':'Instrumentación','EL':'Electricidad','C':'Pruebas/Comis.'}
data=[["Bloque","ID","Tarea","Inicio","Fin","Días","Base / cantidad"]]
for t in TASKS:
    blk,tid,name,ini,fin,color,crit,base=t
    if ini is None and fin is None:
        continue
    if fin is None:
        dur="—"; fin_s=""; ini_s=d(ini).strftime('%d-%m-%y')
    else:
        dur=str((d(fin)-d(ini)).days+1); ini_s=d(ini).strftime('%d-%m-%y'); fin_s=d(fin).strftime('%d-%m-%y')
    nm = name + ("  ◄ CRÍTICA" if crit and fin else "")
    data.append([BLK.get(blk,blk),tid,Paragraph(nm,CELL),ini_s,fin_s,dur,base])
t=LongTable(data,colWidths=[24*mm,14*mm,118*mm,20*mm,20*mm,14*mm,63*mm],repeatRows=1)
st=TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1f3b63')),
   ('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
   ('FONTSIZE',(0,0),(-1,-1),7.5),('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#cccccc')),
   ('VALIGN',(0,0),(-1,-1),'MIDDLE'),('ALIGN',(3,0),(5,-1),'CENTER'),
   ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f4f4f4')])])
ri=1
for tt in TASKS:
    if tt[3] is None and tt[4] is None: continue
    if tt[6]:  # critico
        st.add('TEXTCOLOR',(2,ri),(2,ri),colors.HexColor('#c00000'))
    ri+=1
t.setStyle(st); E.append(t)
E.append(PageBreak())

# ---- Pagina 3: bases, secuencia y notas ----
E.append(Paragraph("Bases de cálculo, secuencia y supuestos", H2))
def kvtable(title, rows, w1=70*mm, w2=205*mm):
    dd=[[title,""]]+[[Paragraph("<b>%s</b>"%k,CELL),Paragraph(v,CELL)] for k,v in rows]
    tb=Table(dd,colWidths=[w1,w2])
    tb.setStyle(TableStyle([('SPAN',(0,0),(1,0)),
       ('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1f3b63')),
       ('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),
       ('FONTSIZE',(0,0),(-1,0),9),('GRID',(0,0),(-1,-1),0.3,colors.HexColor('#cccccc')),
       ('VALIGN',(0,0),(-1,-1),'TOP'),
       ('ROWBACKGROUNDS',(0,1),(-1,-1),[colors.white,colors.HexColor('#f4f4f4')])]))
    return tb
E.append(kvtable("VOLÚMENES BASE",[
  ("Cables EL","513 cables / 81.231 m (MT 18·2.110m; BT grande 104·17.470m ★; BT med/peq 198·33.125m; control/señales 192·28.126m; FO 1·400m)"),
  ("Puntas EL","~5.178 puntas (3.036 salas / 1.717 campo)"),
  ("Cargas EL","243 (79 motores BT+MT · 6 VFD · 52 instrumentos · 28 ilum. · 37 SSAA · 12 calef. · 29 control)"),
  ("Cables IN","885 cables / 65.315 m (procesos 850 · generación 35)"),
  ("Puntas IN","15.752 puntas (12.376 conductor + 1.612 pantalla + 1.764 armadura)"),
  ("Instrumentos","2.798 total · AESA monta ~1.163 (897 AESA + 266 'montaje por EPC')"),
]))
E.append(Spacer(1,3*mm))
E.append(kvtable("RENDIMIENTOS APLICADOS (cuadrilla, jornada 8 h)",[
  ("Tendido IN","simple armado 180 m/d · multipar ≥4P 120 m/d · simple no arm. 250 m/d · FO 200 m/d (3 cuad. → ~150 d; ventana 01-FEB→14-MAR exige 5-6 cuad.)"),
  ("Conexionado IN","7 min/punta cond · 5 pantalla · 10 armadura → 15.752 puntas ≈ 1.974 h-h ≈ 247 días-hombre"),
  ("Tendido EL","MT 30 m/d · BT grande 50 m/d ★ · BT med 80 · BT peq 120 · control 100 (3-4 cuad.)"),
  ("Conexionado EL","term. MT 8 h/ext · BT grande 3 h/ext · BT med/peq 1,5 h/ext · control 4 h/cable"),
  ("Montaje instrumentos","~1.163 instr AESA, 7 típicos de montaje; banco de pruebas (calibración) previo al montaje"),
]))
E.append(Spacer(1,3*mm))
E.append(kvtable("SECUENCIA Y PIVOTES (restricciones duras)",[
  ("⛔ Cables IN","Tendido completo ANTES del 17-MAR-2027 (instalación PCS Sala INS / Sala 7). Conexionado IN arranca con el PCS."),
  ("⛔ Cables EL","Tendido completo ANTES de llegada de salas: SE#4 08-ABR · SE#3 14-MAY. Parte del tendido se hace antes de cerrar las salas para conectar sala→motores/equipos."),
  ("Salas eléctricas 3 y 4","Llegan en módulos → armado en campo (SE#4 14 d; SE#3 21 d) antes del conexionado. Conexionado pivota sobre el armado: SE#4 desde 23-ABR; SE#3 desde 05-JUN."),
  ("Sala INS (control)","Sistemas de control (PCS) y seguridad (SIS) se montan/comisionan en la Sala INS; cables IN deben estar tendidos para iniciar PCS el 17-MAR."),
  ("Ruta crítica","SE#3 14-MAY → armado 21 d → conexionado EL 51 d → megger 21 d → energización 17 d → precom 21 d → comisionado 88 d ≈ 19-DIC. Frente EL crítico: BT Grande (104 c/17.470 m)."),
]))
E.append(Spacer(1,3*mm))
E.append(Paragraph("Notas: (1) Fechas alineadas al Gantt Integrado Rev1 (RFSU 20-DIC-2027); el análisis "
  "eléctrico Rev3 contempla RFSU contractual 31-ENE-2028 como respaldo (margen ~6 semanas). "
  "(2) El alcance de montaje de instrumentos AESA incluye los suministrados por AESA más los de "
  "proveedor marcados 'MONTAJE POR EPC'; los instrumentos en skids de vendor llegan premontados. "
  "(3) <b>Entrega progresiva de instrumentos (Rev.1):</b> Fase 1 = válvulas y todo instrumento en "
  "línea de cañería (920), entregados/montados primero junto con la cañería; Fase 2 = resto de "
  "instrumentos de montaje específico (1.878), entrega progresiva a partir de FEB-2027. "
  "(4) Productividades referenciales O&G Patagonia; aplicar factor de contingencia 1,20. "
  "(5) Documento preliminar para depuración conjunta.", CELL))

# ---- Pagina 4: contexto de planta (vista 3D) ----
from reportlab.platypus import Image as RLImage
import fitz as _fitz; _fitz.TOOLS.mupdf_display_errors(False)
_png = os.path.join(BASE, "_3d_contexto.png")
_pg = _fitz.open(THREED)[0]
_pg.get_pixmap(dpi=200).save(_png)
E.append(PageBreak())
E.append(Paragraph("Contexto de planta — Vista 3D (volumen y congestión)", H2))
_iw, _ih = _pg.rect.width, _pg.rect.height
_h = 138*mm; _w = _h * _iw / _ih
img = RLImage(_png, width=_w, height=_h)
img.hAlign = 'CENTER'
E.append(img)
E.append(Spacer(1,2*mm))
E.append(Paragraph("Modelo 3D CPF2 La Calera II (3D LC 12-06-26). Da escala del <b>volumen y la "
  "congestión</b> de la planta: footprint, densidad de áreas de proceso, recorridos de bandejas/"
  "canalizaciones y accesos. Es el sustento físico de los volúmenes que gobiernan el cronograma "
  "(81.231 m de cable EL · 65.315 m de cable IN · ~20.930 puntas · 2.798 instrumentos) y de los "
  "<b>factores de productividad</b> (trabajo en altura, congestión de canalización, accesos) y del "
  "<b>dimensionamiento y pico de cuadrillas</b> (~40-50 personas en MAR-2027). Uso previsto: "
  "validar secuencia por sectores/áreas y planificar logística de acceso e izaje.", CELL))

doc.build(E)
try: os.remove(_png)
except OSError: pass
print("OK PDF ->",OUT_PDF)

# ---------------- XLSX ----------------
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
wb=openpyxl.Workbook(); ws=wb.active; ws.title="Cronograma_INS_EL"
hdr=["Bloque","ID","Tarea","Inicio","Fin","Días","Critica","Base/Cantidad"]
ws.append(hdr)
for c in ws[1]:
    c.font=Font(bold=True,color="FFFFFF"); c.fill=PatternFill("solid",fgColor="1F3B63")
for t in TASKS:
    blk,tid,name,ini,fin,color,crit,base=t
    if ini is None and fin is None:
        r=[name,"","","","","","",""]; ws.append(r)
        for c in ws[ws.max_row]: c.font=Font(bold=True)
        continue
    dur = ((d(fin)-d(ini)).days+1) if fin else ""
    ws.append([BLK.get(blk,blk),tid,name,
               d(ini) if ini else "", d(fin) if fin else "",
               dur, "SÍ" if crit else "", base])
for col,w in zip("ABCDEFGH",[14,8,56,12,12,7,8,40]):
    ws.column_dimensions[col].width=w
wb.save(OUT_XLS)
print("OK XLSX ->",OUT_XLS)
print("Tareas:",sum(1 for t in TASKS if not(t[3] is None and t[4] is None)),
      "| Hitos:",sum(1 for t in TASKS if t[4] is None and t[3]))
