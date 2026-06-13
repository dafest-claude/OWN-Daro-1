#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diagrama de flujo por LÍNEA DE PROVISIÓN (carriles temporales) - CPF2 La Calera II
Carriles: ABB (PMS) · Inauco (PCS/SCADA/SIS, hasta entrega) · HIMA (ESD/F&G/PSS) ·
PPSA (config ESD/F&G/PSS, A DEFINIR) · AESA (campo/integración/comisionado, fechas AESA).
Muestra cuándo ocurre cada actividad y dónde confluyen (FATs, integración, iFAT, RFSU).
Nota: los trabajos de campo post-entrega de Inauco se ejecutan por AESA en fechas propias.
Salida: Flujo_Provisiones_CPF2.pdf
"""
import os, math
from datetime import date
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors

BASE=os.path.dirname(os.path.abspath(__file__))
OUT=os.path.join(BASE,"Flujo_Provisiones_CPF2.pdf")
REV="0"; HOY=date(2026,6,13); RFSU=date(2027,12,20)
T0=date(2026,6,1); T1=date(2028,1,31)
def D(s): y,m,dd=map(int,s.split('-')); return date(y,m,dd)

NAVY=colors.HexColor('#1f3b63'); C_PMS=colors.HexColor('#c55a11'); C_INA=colors.HexColor('#2e75b6')
C_HIMA=colors.HexColor('#7030a0'); C_DEF=colors.HexColor('#808080'); C_AESA=colors.HexColor('#548235')
GOLD=colors.HexColor('#bf9000'); RED=colors.HexColor('#c00000'); GREY=colors.HexColor('#666666')
LRED=colors.HexColor('#fbe4e4'); LGREY=colors.HexColor('#f4f4f4')

W,Hh=landscape(A4)
c=canvas.Canvas(OUT,pagesize=landscape(A4))

LM=9*mm; RM=7*mm; LW=42*mm
x0=LM+LW; x1=W-RM
def X(dt): return x0+(x1-x0)*((dt-T0).days)/((T1-T0).days)
MES=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic']

# --- Título ---
c.setFillColor(NAVY); c.rect(0,Hh-15*mm,W,15*mm,stroke=0,fill=1)
c.setFillColor(colors.white); c.setFont('Helvetica-Bold',15)
c.drawString(LM,Hh-9*mm,"Flujo por Línea de Provisión — ABB (PMS) · Inauco · HIMA · PPSA · AESA")
c.setFont('Helvetica',8.5); c.setFillColor(colors.HexColor('#9db8d8'))
c.drawString(LM,Hh-13.5*mm,"CPF2 La Calera II · Carriles temporales: cuándo ocurre cada actividad y dónde confluyen · Rev. %s · %s"%(REV,HOY.strftime('%d-%m-%Y')))
c.setFillColor(GOLD); c.setFont('Helvetica-Bold',9); c.drawRightString(x1,Hh-13.5*mm,"RFSU 20-DIC-2027")

yT=Hh-30*mm; yB=24*mm
lanes=[("ABB\nPMS",C_PMS),("INAUCO\nPCS/SCADA/SIS",C_INA),("HIMA\nESD·F&G·PSS",C_HIMA),
       ("PPSA\nConfig ESD/F&G/PSS\n[A DEFINIR]",C_DEF),("AESA\nCampo·Integr.·Comis.",C_AESA)]
nL=len(lanes); laneH=(yT-yB)/nL

# --- Bandas de confluencia (fondo) ---
# Banda FATs + integración fábrica (ENE-MAR-27)
c.setFillColor(LGREY); c.rect(X(D('2027-01-15')),yB,X(D('2027-03-05'))-X(D('2027-01-15')),yT-yB,stroke=0,fill=1)
# Banda iFAT (confluencia total)
c.setFillColor(LRED); c.rect(X(D('2027-06-25')),yB,X(D('2027-07-14'))-X(D('2027-06-25')),yT-yB,stroke=0,fill=1)
# etiquetas de banda
c.setFillColor(GREY); c.setFont('Helvetica-BoldOblique',6.5)
c.drawCentredString((X(D('2027-01-15'))+X(D('2027-03-05')))/2,yT+1.5*mm,"FATs + Integración (fábrica)")
c.setFillColor(RED); c.setFont('Helvetica-Bold',7)
c.drawCentredString((X(D('2027-06-25'))+X(D('2027-07-14')))/2,yT+1.5*mm,"iFAT — CONFLUENCIA TOTAL")

# --- Eje de meses ---
yy=date(T0.year,T0.month,1); ly=None
while yy<=T1:
    xx=X(yy); c.setStrokeColor(colors.HexColor('#e3e3e3')); c.setLineWidth(0.3); c.line(xx,yB,xx,yT)
    nxt=date(yy.year+(yy.month//12),(yy.month%12)+1,1)
    c.setFillColor(GREY); c.setFont('Helvetica',5.2); c.drawCentredString((xx+X(nxt))/2,yT+5.5*mm,MES[yy.month-1])
    if yy.year!=ly:
        c.setFillColor(NAVY); c.setFont('Helvetica-Bold',7); c.drawString(xx+1,yT+8.5*mm,str(yy.year)); ly=yy.year
        c.setStrokeColor(colors.HexColor('#b0b0b0')); c.setLineWidth(0.6); c.line(xx,yB,xx,yT+8*mm)
    yy=nxt
c.setStrokeColor(colors.HexColor('#bbb')); c.setLineWidth(0.5); c.rect(x0,yB,x1-x0,yT-yB,stroke=1,fill=0)

# lineas HOY y RFSU
for dt,col,lbl in [(HOY,colors.HexColor('#1f9d55'),'HOY'),(RFSU,RED,'RFSU')]:
    xx=X(dt); c.setStrokeColor(col); c.setLineWidth(1); c.setDash(2,2); c.line(xx,yB,xx,yT); c.setDash()
    c.saveState(); c.setFillColor(col); c.setFont('Helvetica-Bold',5.5); c.translate(xx,yB); c.rotate(90); c.drawString(1,1.5,lbl); c.restoreState()

# --- Carriles ---
def laneY(i): return yT-(i+1)*laneH
for i,(name,col) in enumerate(lanes):
    y=laneY(i)
    if i%2==0: c.setFillColor(colors.HexColor('#fafafa')); c.rect(x0,y,x1-x0,laneH,stroke=0,fill=1)
    c.setStrokeColor(colors.HexColor('#dddddd')); c.setLineWidth(0.4); c.line(x0,y,x1,y)
    # etiqueta carril
    c.setFillColor(col); c.rect(LM,y+1,LW-3*mm,laneH-2,stroke=0,fill=1)
    c.setFillColor(colors.white)
    lines=name.split('\n')
    fs=8 if len(lines)<=2 else 6.8
    ty=y+laneH/2+(len(lines)-1)*fs*0.5*0.9
    for ln in lines:
        c.setFont('Helvetica-Bold',fs); c.drawCentredString(LM+(LW-3*mm)/2,ty-fs*0.35,ln); ty-=fs*1.0

def bar(i,a,b,label,slot,fill,crit=False,defin=False):
    y=laneY(i); cy=y+laneH*(0.62 if slot==0 else 0.30)
    xa=X(D(a)); xb=X(D(b)); bh=5.2*mm
    c.setFillColor(fill); c.roundRect(xa,cy-bh/2,max(xb-xa,1.2*mm),bh,1.2,stroke=0,fill=1)
    if crit: c.setStrokeColor(RED); c.setLineWidth(1); c.roundRect(xa,cy-bh/2,max(xb-xa,1.2*mm),bh,1.2,stroke=1,fill=0)
    c.setFillColor(colors.HexColor('#222222')); c.setFont('Helvetica',5.4)
    c.drawString(xa+0.5*mm,cy+bh/2+0.6*mm,label)
def ms(i,a,label,slot,fill=GOLD):
    y=laneY(i); cy=y+laneH*(0.62 if slot==0 else 0.30); xx=X(D(a)); s=1.7*mm
    c.setFillColor(fill); c.saveState(); c.translate(xx,cy); c.rotate(45); c.rect(-s,-s,2*s,2*s,stroke=0,fill=1); c.restoreState()
    c.setFillColor(colors.HexColor('#222222')); c.setFont('Helvetica',5.4); c.drawString(xx+2.2*mm,cy-0.8*mm,label)

# ABB-PMS
bar(0,'2026-06-08','2026-10-23','Ingeniería + FP + FAT Maqueta (cliente)',0,C_PMS)
bar(0,'2026-09-11','2026-12-17','Armado tableros PMS-001/101',1,C_PMS)
ms(0,'2027-01-27','FAT PMS',0,RED)
bar(0,'2027-02-04','2027-03-31','Despacho + Montaje Shelter + FAT Integral',1,C_PMS,crit=True)
ms(0,'2027-05-14','Entrega PMS (con SE#4 / SE#3)',0,GOLD)
# Inauco (hasta entrega)
bar(1,'2026-07-01','2027-01-26','Ingeniería Rev0 + Config PLC / SCADA',0,C_INA)
bar(1,'2026-11-10','2027-01-26','Construcción tableros PCS 7A/7B',1,C_INA)
ms(1,'2027-02-13','FAT PCS / FAT SIS',0,RED)
bar(1,'2027-02-24','2027-03-20','Config post-FAT + Entrega a sitio',1,C_INA,crit=True)
# HIMA
bar(2,'2026-07-01','2026-12-23','Tablero marshalling HIMA (25 semanas)',0,C_HIMA)
ms(2,'2026-12-23','Marshalling listo  →  FAT SIS / iFAT',1,GOLD)
# PPSA
bar(3,'2026-08-01','2027-06-24','Config / readecuación ESD·F&G·PSS   [A DEFINIR]',0,C_DEF,crit=True,defin=True)
ms(3,'2027-06-24','FAT ESD/F&G/PSS completo (al iFAT)',1,RED)
# AESA (campo/integración/comisionado — fechas AESA)
ms(4,'2027-02-26','Integración PCS↔PMS (Bs.As., witness)',0,GOLD)
ms(4,'2027-03-17','Inicio campo PCS — Sala INS',1,GOLD)
bar(4,'2027-03-17','2027-08-03','SAT / Comisionado PCS y SIS (campo) — fechas AESA',0,C_AESA)
bar(4,'2027-06-25','2027-07-14','iFAT (integración total + PCS↔PMS EN SITIO)',1,C_AESA,crit=True)
bar(4,'2027-08-04','2027-12-19','MCE + Precom + Comisionado integrado',0,C_AESA,crit=True)
ms(4,'2027-12-20','★ RFSU',1,RED)

# --- Flechas de confluencia hacia el iFAT ---
xi=X(D('2027-06-25'));
def conf_arrow(i, fromdate):
    y=laneY(i)+laneH*0.30; xf=X(D(fromdate))
    c.setStrokeColor(colors.HexColor('#b04040')); c.setLineWidth(0.8); c.setDash(1.5,1.5)
    yi=laneY(4)+laneH*0.5
    c.line(xf,y,xi-1*mm,yi); c.setDash()
for i,fd in [(0,'2027-05-14'),(1,'2027-03-20'),(2,'2026-12-23'),(3,'2027-06-24')]:
    conf_arrow(i,fd)

# --- Leyenda y notas ---
c.setStrokeColor(colors.HexColor('#cccccc')); c.setLineWidth(0.4); c.line(LM,yB-3*mm,x1,yB-3*mm)
items=[("ABB · PMS",C_PMS),("Inauco · PCS/SCADA/SIS",C_INA),("HIMA · ESD/F&G/PSS",C_HIMA),
       ("PPSA · config [A DEFINIR]",C_DEF),("AESA · campo/comis.",C_AESA),("◆ Hito / FAT",GOLD),("▭ rojo = crítico",RED)]
lx=LM
for txt,col in items:
    c.setFillColor(col); c.rect(lx,yB-7*mm,3.5*mm,3.5*mm,stroke=0,fill=1)
    c.setFillColor(colors.HexColor('#333333')); c.setFont('Helvetica',6.5); c.drawString(lx+4.5*mm,yB-6.2*mm,txt)
    lx+=len(txt)*1.5*mm+12*mm
c.setFillColor(GREY); c.setFont('Helvetica-Oblique',6.3)
c.drawString(LM,yB-11*mm,"Inauco: alcance hasta la ENTREGA; los trabajos de campo (SAT/comisionado PCS-SIS) los ejecuta AESA en fechas propias. PPSA: config ESD/F&G/PSS A DEFINIR (arranca AGO, completa al iFAT).")
c.drawString(LM,yB-14*mm,"Confluencia: las provisiones se prueban individualmente (FATs ENE-FEB), se integran PCS↔PMS en fábrica (Bs.As. FEB-MAR) y confluyen TOTALMENTE en el iFAT (JUN-JUL) → comisionado → RFSU 20-DIC-2027.")
c.setFont('Helvetica',7); c.setFillColor(GREY)
c.drawRightString(x1,8*mm,"Constructibilidad CPF2 · Flujo por línea de provisión · Rev. %s"%REV)
c.showPage(); c.save()
print("OK ->",OUT)
