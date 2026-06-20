#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Flujo VERTICAL TIME-SCALED del programa de conexionado EL+IN - CPF2 La Calera 2
Eje temporal vertical (JUN-26 -> FEB-28) a la izquierda: cada actividad se ubica
en su fecha. Ramas IN / EL en columnas, shelters desglosados (SE#3, SE#4, Sala INS),
confluencia en precom/energización/comisionado y RUTA CRÍTICA en rojo. RFSU 31-ENE-2028.
Salida: Flujo_Conexionado_EL_IN_CPF2_Rev1.pdf
"""
import os, math
from datetime import date
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors

BASE=os.path.dirname(os.path.abspath(__file__))
OUT=os.path.join(BASE,"Flujo_Conexionado_EL_IN_CPF2_Rev1.pdf")
W,H=A4
NAVY=colors.HexColor('#1f3b63'); IN=colors.HexColor('#2e75b6'); EL=colors.HexColor('#c55a11')
GREEN=colors.HexColor('#548235'); GOLD=colors.HexColor('#bf9000'); RED=colors.HexColor('#c00000')
YEL=colors.HexColor('#ffc000'); GREY=colors.HexColor('#777777'); WHITE=colors.white
def D(s): y,m,dd=map(int,s.split('-')); return date(y,m,dd)
T0=D('2026-06-01'); T1=D('2028-02-29')
ytop=H-26*mm; ybot=20*mm
def Y(dt): return ytop-(dt-T0).days/(T1-T0).days*(ytop-ybot)
MES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

c=canvas.Canvas(OUT,pagesize=A4)
# Titulo
c.setFillColor(NAVY); c.rect(0,H-18*mm,W,18*mm,stroke=0,fill=1)
c.setFillColor(WHITE); c.setFont('Helvetica-Bold',13)
c.drawCentredString(W/2,H-10*mm,"Flujo de Conexionado EL+IN — vertical con escala temporal")
c.setFont('Helvetica',8); c.setFillColor(colors.HexColor('#9db8d8'))
c.drawCentredString(W/2,H-15*mm,"CPF2 La Calera 2 · cada actividad ubicada en su fecha · ruta crítica en rojo · RFSU 31-ENE-2028")

# Eje temporal (meses)
axx=24*mm
yy=date(T0.year,T0.month,1); ly=None
while yy<=T1:
    y=Y(yy); c.setStrokeColor(colors.HexColor('#e8e8e8')); c.setLineWidth(0.3); c.line(axx,y,W-8*mm,y)
    c.setFillColor(GREY); c.setFont('Helvetica',6); c.drawRightString(axx-1.5*mm,y-1.4,MES[yy.month-1])
    if yy.year!=ly:
        c.setFillColor(NAVY); c.setFont('Helvetica-Bold',7.5); c.drawRightString(axx-1.5*mm,y+1.5,str(yy.year)); ly=yy.year
        c.setStrokeColor(colors.HexColor('#b8b8b8')); c.setLineWidth(0.6); c.line(axx,y,W-8*mm,y)
    m=yy.month+1; yr=yy.year+(1 if m>12 else 0); m=1 if m>12 else m; yy=date(yr,m,1)
c.setStrokeColor(NAVY); c.setLineWidth(0.8); c.line(axx,ybot,axx,ytop)

# Encabezados de columna
def colhdr(x,txt,col): c.setFillColor(col); c.setFont('Helvetica-Bold',9); c.drawCentredString(x,ytop+4*mm,txt)
colhdr(44*mm,"INSTRUMENTACIÓN",IN); colhdr(92*mm,"ELECTRICIDAD",EL); colhdr(152*mm,"CONFLUENCIA",GREEN)

def bar(xmm,a,b,label,fill,crit=False):
    x=xmm; bw=7.5*mm; ya=Y(D(a)); yb=Y(D(b)); h=ya-yb
    c.setFillColor(fill); c.roundRect(x-bw/2,yb,bw,h,1.5,stroke=0,fill=1)
    if crit: c.setStrokeColor(RED); c.setLineWidth(1.6); c.roundRect(x-bw/2,yb,bw,h,1.5,stroke=1,fill=0)
    # etiqueta rotada centrada
    c.saveState(); c.setFillColor(WHITE if fill not in (YEL,) else colors.HexColor('#333333'))
    fs=6.2 if h>26*mm else 5.4
    c.setFont('Helvetica-Bold',fs); c.translate(x,(ya+yb)/2); c.rotate(90)
    c.drawCentredString(0,-fs*0.35,label[:34]); c.restoreState()
def milestone(xmm,a,label):
    x=xmm; y=Y(D(a)); s=2.0*mm
    c.setFillColor(RED); c.saveState(); c.translate(x,y); c.rotate(45); c.rect(-s,-s,2*s,2*s,stroke=0,fill=1); c.restoreState()
    c.setFillColor(RED); c.setFont('Helvetica-Bold',8); c.drawString(x+3.5*mm,y-1.2,label)

# IN
bar(40*mm,'2026-11-01','2027-01-31','Canalizaciones IN (ref)',YEL)
bar(40*mm,'2027-02-10','2027-10-28','Montaje de Instrumentos',IN)
bar(52*mm,'2027-04-03','2027-12-23','Cables IN (tendido+conex.)',IN,crit=True)
# EL
bar(74*mm,'2026-10-01','2026-12-31','Canalizaciones EL (ref)',YEL)
bar(74*mm,'2027-02-17','2027-03-29','Montaje Sala INS',EL)
bar(86*mm,'2027-05-14','2027-06-22','Shelter SE#4',EL)
bar(98*mm,'2027-05-28','2027-07-10','Shelter SE#3',EL,crit=True)
bar(86*mm,'2027-07-12','2027-09-11','Montaje tableros / equipos',EL)
bar(98*mm,'2027-07-23','2027-08-14','Transf / Gen Emerg',EL)
bar(110*mm,'2027-05-30','2027-11-29','Cables EL (tendido+conex.)',EL,crit=True)
# Confluencia
bar(134*mm,'2027-05-03','2028-01-13','Precomisionado',GREEN,crit=True)
bar(148*mm,'2027-10-01','2027-10-31','Energiz. P1',GOLD)
bar(148*mm,'2027-11-01','2027-11-30','Energiz. P2',GOLD,crit=True)
bar(162*mm,'2027-07-16','2028-01-31','Comisionado',GREEN,crit=True)
milestone(162*mm,'2028-01-31','RFSU 31-ENE-2028')

# Flechas de confluencia (cables -> comisionado/precom)
c.setDash(2,2)
def carr(x0,a,x1,b,col):
    y0=Y(D(a)); y1=Y(D(b)); c.setStrokeColor(col); c.setLineWidth(0.9); c.line(x0,y0,x1,y1)
    ang=math.atan2(y1-y0,x1-x0); L=2*mm; c.setFillColor(col)
    c.saveState(); c.translate(x1,y1); c.rotate(math.degrees(ang)); p=c.beginPath(); p.moveTo(0,0); p.lineTo(-L,L*0.5); p.lineTo(-L,-L*0.5); p.close(); c.drawPath(p,fill=1,stroke=0); c.restoreState()
carr(52*mm,'2027-12-23',134*mm,'2027-12-15',IN)   # cables IN -> precom
carr(110*mm,'2027-11-29',162*mm,'2027-11-29',EL)  # cables EL -> comisionado
c.setDash()

# Leyenda
ly0=14*mm
items=[("IN",IN),("EL",EL),("Confluencia / precom-comis",GREEN),("Energización",GOLD),("Referencia (línea base)",YEL),("◆/▭ rojo = ruta crítica · RFSU",RED)]
lx=axx
for txt,col in items:
    c.setFillColor(col); c.rect(lx,ly0,3.2*mm,3.2*mm,stroke=0,fill=1)
    c.setFillColor(colors.HexColor('#333333')); c.setFont('Helvetica',6.3); c.drawString(lx+4*mm,ly0+0.6*mm,txt); lx+=len(txt)*1.45*mm+9*mm
c.setFillColor(GREY); c.setFont('Helvetica-Oblique',6.3)
c.drawString(axx,9*mm,"Ruta crítica: cables → conexionado → energización (P2) → comisionado → RFSU. Cronograma muy ajustado; ver 'adelantar actividades y gestionar PHs'.")
c.setFont('Helvetica',6.5); c.setFillColor(GREY); c.drawRightString(W-8*mm,5*mm,"Constructibilidad CPF2 · Flujo conexionado vertical time-scaled · Rev. 1")
c.showPage(); c.save()
print("OK ->",OUT)
