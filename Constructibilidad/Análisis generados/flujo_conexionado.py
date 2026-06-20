#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Diagrama de flujo VERTICAL del programa de conexionado EL + IN - CPF2 La Calera 2
Vista top-down: habilitaciones -> dos ramas paralelas (Instrumentación / Electricidad)
-> confluencia en Precomisionado -> Energización -> Comisionado -> RFSU 31-ENE-2028.
Basado en Programa_Conexionado_EL_IN_CPF2 (imagen original reconstruida).
Salida: Flujo_Conexionado_EL_IN_CPF2.pdf
"""
import os, math
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib import colors

BASE=os.path.dirname(os.path.abspath(__file__))
OUT=os.path.join(BASE,"Flujo_Conexionado_EL_IN_CPF2.pdf")
W,H=A4  # portrait 210x297

NAVY=colors.HexColor('#1f3b63'); IN=colors.HexColor('#2e75b6'); EL=colors.HexColor('#c55a11')
GREEN=colors.HexColor('#548235'); GOLD=colors.HexColor('#bf9000'); RED=colors.HexColor('#c00000')
GREY=colors.HexColor('#7f7f7f'); YEL=colors.HexColor('#bf9000'); WHITE=colors.white
LG=colors.HexColor('#666666')

c=canvas.Canvas(OUT,pagesize=A4)

def box(cx,cy,w,h,title,sub,fill,fc=WHITE,tsz=9,ssz=7):
    c.setFillColor(fill); c.roundRect(cx-w/2,cy-h/2,w,h,2.2,stroke=0,fill=1)
    c.setFillColor(fc); c.setFont('Helvetica-Bold',tsz)
    c.drawCentredString(cx,cy+ (h/2-tsz-0.5 if sub else -tsz*0.35),title)
    if sub:
        c.setFont('Helvetica',ssz)
        for i,ln in enumerate(sub.split('\n')):
            c.drawCentredString(cx,cy+h/2-tsz-3-(i+1)*(ssz+1.2),ln)
def arrow(x0,y0,x1,y1,col=GREY,wd=1.4):
    c.setStrokeColor(col); c.setLineWidth(wd); c.line(x0,y0,x1,y1)
    a=math.atan2(y1-y0,x1-x0); L=2.4*mm
    c.setFillColor(col); c.saveState(); c.translate(x1,y1); c.rotate(math.degrees(a))
    p=c.beginPath(); p.moveTo(0,0); p.lineTo(-L,L*0.5); p.lineTo(-L,-L*0.5); p.close(); c.drawPath(p,fill=1,stroke=0); c.restoreState()

# --- Título ---
c.setFillColor(NAVY); c.rect(0,H-20*mm,W,20*mm,stroke=0,fill=1)
c.setFillColor(WHITE); c.setFont('Helvetica-Bold',14)
c.drawCentredString(W/2,H-11*mm,"Flujo de Conexionado — Electricidad e Instrumentación")
c.setFont('Helvetica',8.5); c.setFillColor(colors.HexColor('#9db8d8'))
c.drawCentredString(W/2,H-16*mm,"CPF2 La Calera 2 · vista vertical (top-down) · basado en el programa de conexionado · RFSU 31-ENE-2028")

cxL=58*mm; cxR=W-58*mm; cxC=W/2
bw=80*mm; bh=15*mm

# --- Habilitaciones (top, full) ---
yh=H-32*mm
box(cxC,yh,176*mm,12*mm,"HABILITACIONES / LLEGADAS (predecesoras)",
    "Canalizaciones · cables · instrumentos · fundaciones · EE.MM.  (línea base original — referencia)",NAVY,tsz=9,ssz=7)

# encabezados de rama
c.setFillColor(IN); c.setFont('Helvetica-Bold',11); c.drawCentredString(cxL,yh-9*mm,"INSTRUMENTACIÓN (IN)")
c.setFillColor(EL); c.drawCentredString(cxR,yh-9*mm,"ELECTRICIDAD (EL)")

# filas de cada rama (4)
rowy=[yh-22*mm, yh-44*mm, yh-66*mm, yh-88*mm]
IN_boxes=[("Canalizaciones IN","≈ nov-26 → ene-27  (ref)"),
 ("Tendido cables IN","abr-27 → dic-27  ·  58.713 m / 1.816 ptas"),
 ("Montaje de Instrumentos","10-feb → 28-oct-27  ·  Válvulas 362 + Instr. 838"),
 ("Conexionado IN","→ 23-dic-27  ·  ~12.684 / 1.308 ptas")]
EL_boxes=[("Canalizaciones EL","≈ oct-26 → dic-26  (ref)"),
 ("Tendido cables EL","30-may → nov-27  ·  93.555 m / 4.175 ptas"),
 ("Montaje salas / tableros","Sala INS feb · SE#4/SE#3 may-jul · tableros jul-sep"),
 ("Conexionado EL","→ 29-nov-27  ·  ~7.137 / 1.062 ptas")]
for i in range(4):
    box(cxL,rowy[i],bw,bh,IN_boxes[i][0],IN_boxes[i][1],IN)
    box(cxR,rowy[i],bw,bh,EL_boxes[i][0],EL_boxes[i][1],EL)
    if i==0:
        arrow(cxL,yh-6*mm,cxL,rowy[0]+bh/2); arrow(cxR,yh-6*mm,cxR,rowy[0]+bh/2)
    else:
        arrow(cxL,rowy[i-1]-bh/2,cxL,rowy[i]+bh/2); arrow(cxR,rowy[i-1]-bh/2,cxR,rowy[i]+bh/2)

# --- Confluencia ---
yP=rowy[3]-24*mm
box(cxC,yP,150*mm,15*mm,"PRECOMISIONADO",
    "03-may-27 → 13-ene-28  ·  eléctrico (bomba-motor-PMS) + SCADA / PCS / PSS / ESD / F&G",GREEN,tsz=10,ssz=7.5)
arrow(cxL,rowy[3]-bh/2,cxC-30*mm,yP+15*mm/2,IN)
arrow(cxR,rowy[3]-bh/2,cxC+30*mm,yP+15*mm/2,EL)

yE=yP-22*mm
box(cxC,yE,150*mm,13*mm,"ENERGIZACIÓN",
    "Propuesta 1 (~oct-27)   ·   Propuesta 2 (~nov-27)",GOLD,tsz=10,ssz=7.5)
arrow(cxC,yP-15*mm/2,cxC,yE+13*mm/2)

yM=yE-22*mm
box(cxC,yM,150*mm,13*mm,"COMISIONADO",
    "16-jul-27 → 31-ene-28  ·  EL + PCS + PMS + SIS / SCADA",GREEN,tsz=10,ssz=7.5)
arrow(cxC,yE-13*mm/2,cxC,yM+13*mm/2)

yR=yM-22*mm
box(cxC,yR,110*mm,15*mm,"★ RFSU — READY FOR START UP","31-ENE-2028",RED,tsz=12,ssz=9)
arrow(cxC,yM-13*mm/2,cxC,yR+15*mm/2,RED,1.8)

# --- nota ---
c.setFillColor(LG); c.setFont('Helvetica-Oblique',7)
c.drawCentredString(W/2,16*mm,"Cronograma MUY ajustado. Anotación del programa: 'adelantar actividades y gestionar PHs → re-instalar instrumentos'.")
c.drawCentredString(W/2,12*mm,"Las habilitaciones (canalizaciones/cables/instrumentos) son la línea base ORIGINAL (referencia). Fechas según Programa_Conexionado_EL_IN_CPF2.")
c.setFont('Helvetica',7); c.setFillColor(LG)
c.drawRightString(W-12*mm,7*mm,"Constructibilidad CPF2 · Flujo de conexionado (vertical) · Rev. 0")
c.showPage(); c.save()
print("OK ->",OUT)
