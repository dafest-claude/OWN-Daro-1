#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diagrama de flujo VERTICAL time-scaled del CRONOGRAMA REALISTA - CPF2 La Calera II
Carriles: SUMINISTROS (entregas) · INSTRUMENTACIÓN · ELECTRICIDAD · CONFLUENCIA (precom→comis).
Eje temporal vertical ENE-27 → JUL-28. Ruta crítica en rojo. RFSU base 31-ENE-28 vs realista 31-MAY-28.
Salida: Flujo_Realista_CPF2.pdf
"""
import os, math
from datetime import date
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors
BASE=os.path.dirname(os.path.abspath(__file__))
OUT=os.path.join(BASE,"Flujo_Realista_CPF2.pdf")
W,H=A4
NAVY=colors.HexColor('#1f3b63'); GOLD=colors.HexColor('#bf9000'); IN=colors.HexColor('#2e75b6')
EL=colors.HexColor('#c55a11'); GREEN=colors.HexColor('#548235'); PURP=colors.HexColor('#7030a0')
RED=colors.HexColor('#c00000'); GREY=colors.HexColor('#888888'); WHITE=colors.white
def D(s): y,m,dd=map(int,s.split('-')); return date(y,m,dd)
T0=D('2027-01-01'); T1=D('2028-07-31'); ytop=H-26*mm; ybot=22*mm
def Y(dt): return ytop-(dt-T0).days/(T1-T0).days*(ytop-ybot)
MES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']
c=canvas.Canvas(OUT,pagesize=A4)
c.setFillColor(NAVY); c.rect(0,H-18*mm,W,18*mm,stroke=0,fill=1)
c.setFillColor(WHITE); c.setFont('Helvetica-Bold',13)
c.drawCentredString(W/2,H-9.5*mm,"Flujo Realista de Ejecución — EL + IN (anclado en suministros)")
c.setFont('Helvetica',8); c.setFillColor(colors.HexColor('#9db8d8'))
c.drawCentredString(W/2,H-14.5*mm,"CPF2 La Calera II · escala temporal · ruta crítica en rojo · RFSU base 31-ENE-28 → realista 31-MAY-28")
axx=22*mm
yy=date(T0.year,T0.month,1); ly=None
while yy<=T1:
    y=Y(yy); c.setStrokeColor(colors.HexColor('#eaeaea')); c.setLineWidth(0.3); c.line(axx,y,W-8*mm,y)
    c.setFillColor(GREY); c.setFont('Helvetica',6); c.drawRightString(axx-1.5*mm,y-1.3,MES[yy.month-1])
    if yy.year!=ly:
        c.setFillColor(NAVY); c.setFont('Helvetica-Bold',7.5); c.drawRightString(axx-1.5*mm,y+1.6,str(yy.year)); ly=yy.year
        c.setStrokeColor(colors.HexColor('#b8b8b8')); c.setLineWidth(0.6); c.line(axx,y,W-8*mm,y)
    m=yy.month+1; yr=yy.year+(1 if m>12 else 0); m=1 if m>12 else m; yy=date(yr,m,1)
c.setStrokeColor(NAVY); c.setLineWidth(0.8); c.line(axx,ybot,axx,ytop)
# lineas RFSU
for dt,col,lbl in [(D('2028-01-31'),GREY,'RFSU base'),(D('2028-05-31'),RED,'RFSU real')]:
    y=Y(dt); c.setStrokeColor(col); c.setLineWidth(1); c.setDash(2,2); c.line(axx,y,W-8*mm,y); c.setDash()
    c.setFillColor(col); c.setFont('Helvetica-Bold',6); c.drawString(axx+1,y+0.8,lbl)
def hdr(x,t,col): c.setFillColor(col); c.setFont('Helvetica-Bold',8.5); c.drawCentredString(x,ytop+4*mm,t)
hdr(38*mm,"SUMINISTROS",GOLD); hdr(70*mm,"INSTRUM.",IN); hdr(104*mm,"ELECTRIC.",EL); hdr(150*mm,"CONFLUENCIA",GREEN)
def bar(x,a,b,label,fill,crit=False):
    bw=7*mm; ya=Y(D(a)); yb=Y(D(b)); h=ya-yb
    c.setFillColor(fill); c.roundRect(x-bw/2,yb,bw,h,1.4,stroke=0,fill=1)
    if crit: c.setStrokeColor(RED); c.setLineWidth(1.5); c.roundRect(x-bw/2,yb,bw,h,1.4,stroke=1,fill=0)
    c.saveState(); c.setFillColor(WHITE); fs=6 if h>24*mm else 5.2; c.setFont('Helvetica-Bold',fs)
    c.translate(x,(ya+yb)/2); c.rotate(90); c.drawCentredString(0,-fs*0.35,label[:32]); c.restoreState()
def ms(x,a,label,fill=GOLD,crit=False):
    y=Y(D(a)); s=1.8*mm; col=RED if crit else fill
    c.setFillColor(col); c.saveState(); c.translate(x,y); c.rotate(45); c.rect(-s,-s,2*s,2*s,stroke=0,fill=1); c.restoreState()
    c.setFillColor(colors.HexColor('#333')); c.setFont('Helvetica',5.4); c.drawString(x+2.6*mm,y-1.0,label)
# SUMINISTROS (entregas)
for a,lbl,cr in [('2027-03-29','Cables IN',False),('2027-04-18','SIS',False),('2027-04-26','Cables EL',False),
 ('2027-05-24','SE#3',False),('2027-06-02','PMS',False),('2027-06-23','SE#4',False),
 ('2027-06-24','PCS (Inauco)',True),('2027-07-20','Skids comb.',True)]:
    ms(38*mm,a,lbl,GOLD,cr)
# IN
bar(64*mm,'2027-02-01','2027-04-15','Canalizaciones IN',IN)
bar(64*mm,'2027-06-01','2027-11-30','Montaje instrumentos',IN)
bar(76*mm,'2027-04-05','2027-08-31','Tendido IN',IN,crit=True)
bar(76*mm,'2027-07-15','2027-12-23','Conexionado IN',IN,crit=True)
bar(64*mm,'2027-11-01','2028-02-15','Loop check IN',IN,crit=True)
# EL
bar(98*mm,'2027-03-01','2027-05-15','Bandejas EL',EL)
bar(98*mm,'2027-05-24','2027-07-21','Armado salas SE#3/4',EL)
bar(110*mm,'2027-05-01','2027-09-30','Tendido EL',EL,crit=True)
bar(98*mm,'2027-07-01','2027-10-31','Montaje tableros/motores',EL)
bar(110*mm,'2027-06-22','2027-11-30','Conexionado EL',EL,crit=True)
bar(110*mm,'2027-11-01','2027-11-30','Megger',EL,crit=True)
# CONFLUENCIA
bar(140*mm,'2027-12-01','2027-12-31','Energización',GREEN,crit=True)
bar(140*mm,'2027-12-15','2028-02-29','Precomisionado',GREEN,crit=True)
bar(154*mm,'2028-02-15','2028-05-15','Comisionado',PURP,crit=True)
ms(154*mm,'2028-01-31','RFSU base (forzado)',GREY,False)
ms(168*mm,'2028-05-31','★ RFSU realista',RED,True)
# flechas confluencia
c.setDash(2,2)
def carr(x0,a,x1,b,col):
    y0=Y(D(a)); y1=Y(D(b)); c.setStrokeColor(col); c.setLineWidth(0.9); c.line(x0,y0,x1,y1)
    ang=math.atan2(y1-y0,x1-x0); L=2*mm; c.setFillColor(col)
    c.saveState(); c.translate(x1,y1); c.rotate(math.degrees(ang)); p=c.beginPath(); p.moveTo(0,0); p.lineTo(-L,L*0.5); p.lineTo(-L,-L*0.5); p.close(); c.drawPath(p,fill=1,stroke=0); c.restoreState()
carr(76*mm,'2027-12-23',140*mm,'2027-12-20',IN)
carr(110*mm,'2027-11-30',140*mm,'2027-12-05',EL)
carr(140*mm,'2028-02-29',154*mm,'2028-02-20',GREEN)
c.setDash()
# leyenda + nota
ly0=14*mm; items=[("Suministros",GOLD),("IN",IN),("EL",EL),("Energiz./Precom",GREEN),("Comisionado",PURP),("rojo=crítico",RED)]
lx=axx
for t,col in items:
    c.setFillColor(col); c.rect(lx,ly0,3.2*mm,3.2*mm,stroke=0,fill=1); c.setFillColor(colors.HexColor('#333')); c.setFont('Helvetica',6.3); c.drawString(lx+4*mm,ly0+0.6*mm,t); lx+=len(t)*1.5*mm+10*mm
c.setFillColor(GREY); c.setFont('Helvetica-Oblique',6.3)
c.drawString(axx,9.5*mm,"Última entrega crítica skids 20-jul-27 → montaje → conexionado → energización (dic-27) → precom (dic-27/feb-28) → comisionado (feb/may-28) → RFSU.")
c.drawString(axx,6*mm,"El RFSU base 31-ENE-28 sólo se cumple forzando el solapamiento precom+comisionado. Realista ≈ 31-MAY-28 (~+4 meses). Preliminar para depuración.")
c.showPage(); c.save(); print("OK ->",OUT)
