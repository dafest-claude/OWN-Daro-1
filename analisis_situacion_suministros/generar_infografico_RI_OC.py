#!/usr/bin/env python3
"""
Infográfico gerencial (PNG) – Análisis de tiempos RI → OC
La Calera II CPF2 | Instrumentación (IN) y Electricidad (EL)
Complemento visual del dashboard Excel para presentación.
"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, Rectangle
import matplotlib.font_manager as fm

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_PNG = os.path.join(SCRIPT_DIR, 'Infografico_RI_OC_IN_EL_LaCalera_II_Rev14_110926.png')

# Paleta
AZUL='#1F3864'; AZUL_M='#2F5496'; VERDE='#2E6B2E'; NARANJA='#9C4500'
VERDE_OK='#548235'; AMBAR='#BF8F00'; ROJO='#C00000'; GRIS='#404040'; GRIS_CL='#F2F2F2'
VERDE_CL='#D7E9D7'; NARANJA_CL='#F5E0D0'

plt.rcParams['font.family'] = 'DejaVu Sans'

fig = plt.figure(figsize=(16, 10), dpi=130)
fig.patch.set_facecolor('white')
gs = fig.add_gridspec(100, 100, left=0.03, right=0.97, top=0.97, bottom=0.03)

# ── Banda de título ──────────────────────────────────────────────────────────
axT = fig.add_subplot(gs[0:9, :]); axT.axis('off')
axT.add_patch(Rectangle((0,0),1,1, transform=axT.transAxes, color=AZUL))
axT.text(0.015, 0.62, 'ANÁLISIS DE TIEMPOS  RI → ORDEN DE COMPRA',
         color='white', fontsize=23, fontweight='bold', va='center')
axT.text(0.015, 0.22, 'La Calera II CPF2  ·  Instrumentación (IN) y Electricidad (EL)  ·  '
                      'Plan 110926 (Rev14)  ·  corte 11/09/2026  ·  evolución 14/08→11/09',
         color='#D6E4F0', fontsize=11, va='center')

# ── KPIs ─────────────────────────────────────────────────────────────────────
kpis = [('69','RIs TOTALES\nIN + EL', AZUL, '='),
        ('32','OCs\nCOLOCADAS', VERDE, '+7'),
        ('66','EN GESTIÓN\nDE COMPRA', NARANJA, '='),
        ('3','RIs SIN\nEMITIR', '#7030A0', '=')]
for i,(num,lbl,col,delta) in enumerate(kpis):
    ax = fig.add_subplot(gs[11:24, i*25+1:i*25+23]); ax.axis('off')
    box = FancyBboxPatch((0.02,0.05),0.96,0.9, boxstyle='round,pad=0.02,rounding_size=0.06',
                         transform=ax.transAxes, facecolor=col, edgecolor='none')
    ax.add_patch(box)
    ax.text(0.5,0.64,num, transform=ax.transAxes, color='white', fontsize=42,
            fontweight='bold', ha='center', va='center')
    ax.text(0.5,0.22,lbl, transform=ax.transAxes, color='white', fontsize=11,
            ha='center', va='center')
    if delta:
        ax.text(0.93,0.90,delta, transform=ax.transAxes, color='white', fontsize=12,
                fontweight='bold', ha='right', va='center')

# ── Novedades: semana 17/07 → 24/07 ──────────────────────────────────────────
axN = fig.add_subplot(gs[24:28, 1:99]); axN.axis('off')
axN.text(0.0,0.5,'SEMANA 04/09→11/09:', fontsize=10.5, fontweight='bold', color=AZUL,
         transform=axN.transAxes, va='center')
axN.text(0.155,0.5,'+7 OCs IN+EL (25→32: IN 16→22, EL 9→10)   ·   AT IN 29→23   ·   '
                   'Cables: FO y Rev.1 por comprar   ·   Válvulas: único sin OC (141 d)',
         fontsize=10, color=GRIS, transform=axN.transAxes, va='center')

# ── Embudo pipeline por especialidad ─────────────────────────────────────────
axP = fig.add_subplot(gs[29:62, 1:48])
etapas = ['En\nSOLPED','Petición\nOfertas','Análisis\nTécnico','Con OC']
el = [0,3,5,10]; inn = [0,3,23,22]
x = range(len(etapas)); w=0.38
b1=axP.bar([i-w/2 for i in x], el, w, label='ELECTRICIDAD (19 RIs)', color=VERDE)
b2=axP.bar([i+w/2 for i in x], inn, w, label='INSTRUMENTACIÓN (48 RIs)', color=NARANJA)
axP.set_xticks(list(x)); axP.set_xticklabels(etapas, fontsize=10)
axP.set_ylabel('Cantidad de RIs', fontsize=10)
axP.set_title('Embudo de gestión de compra por especialidad', fontsize=13, fontweight='bold', color=AZUL, pad=10)
axP.legend(fontsize=9, loc='upper left')
for bars in (b1,b2):
    for b in bars:
        h=b.get_height()
        if h > 0:
            axP.text(b.get_x()+b.get_width()/2, h+0.3, str(int(h)), ha='center', fontsize=9, fontweight='bold')
# resaltar cuello de botella AT en IN
axP.annotate('CUELLO DE BOTELLA', xy=(2.19,23), xytext=(1.3,30),
             fontsize=10, fontweight='bold', color=ROJO,
             arrowprops=dict(arrowstyle='->', color=ROJO, lw=2))
axP.set_ylim(0,40); axP.spines['top'].set_visible(False); axP.spines['right'].set_visible(False)

# ── Descomposición de la demora (segmentos) ──────────────────────────────────
axS = fig.add_subplot(gs[29:62, 53:99])
segs = ['RI →\nSOLPED','SOLPED →\nRec.Of.','Rec.Of. →\nAT cierre','AT cierre →\nOC']
dias_seg = [12,19,32,66]
cols = [AZUL_M,AZUL_M,AZUL_M,ROJO]
bars = axS.barh(range(len(segs)), dias_seg, color=cols)
axS.set_yticks(range(len(segs))); axS.set_yticklabels(segs, fontsize=10)
axS.invert_yaxis()
axS.set_xlabel('días', fontsize=10)
axS.set_title('¿Dónde se concentra la demora?  (cadena de 129 días – adjudicados ABB)',
              fontsize=12.5, fontweight='bold', color=AZUL, pad=10)
for i,(b,d) in enumerate(zip(bars,dias_seg)):
    pct = d/sum(dias_seg)*100
    axS.text(d+1.2, b.get_y()+b.get_height()/2, f'{d} d  ({pct:.0f}%)',
             va='center', fontsize=10, fontweight='bold',
             color=ROJO if i==3 else GRIS)
axS.set_xlim(0,82); axS.spines['top'].set_visible(False); axS.spines['right'].set_visible(False)
axS.text(0.99,-0.16,'El tramo AT cierre → OC concentra la mitad del tiempo total',
         transform=axS.transAxes, ha='right', fontsize=10, style='italic', color=ROJO)

# ── Tabla de ítems críticos RI→OC ────────────────────────────────────────────
axI = fig.add_subplot(gs[66:99, 1:99]); axI.axis('off')
axI.text(0.0,1.02,'Ítems críticos – tiempo RI → OC', fontsize=13, fontweight='bold', color=AZUL,
         transform=axI.transAxes)
items = [
    ('EL','Shelter SE#4 / SE#3 (ABB)','09/01','18/05 (OC)','129 d','OC · FABRICAND.', VERDE_OK),
    ('EL','Sistema PMS (ABB)','09/01','22/05 (OC)','133 d','OC COLOCADA', VERDE_OK),
    ('IN','Sistema Seguridad SIS','03/02','03/07 (OC)','150 d','OC · KOM 14/8', VERDE_OK),
    ('IN','Sistema Control PCS','03/02','23/07 (OC mat.)','170 d','OC EMITIDA', VERDE_OK),
    ('EL','Cables Eléctricos','11/06','23/08 (OC Marlew)','73 d','OC COLOCADA', VERDE_OK),
    ('IN','Cables Instrumentación','15/05','27/08 (OC)','104 d','OC · 2da inst.', VERDE_OK),
    ('IN','Válvulas Control','23/04','flete aéreo','141 d*','ÚNICO SIN OC', ROJO),
]
cols_x = [0.0, 0.07, 0.42, 0.55, 0.71, 0.81, 0.99]
hdrs = ['Esp','Suministro','RI','OC efectiva','RI→OC','Estado']
y0=0.92
axI.add_patch(Rectangle((0,y0-0.005),1,0.07, transform=axI.transAxes, color=GRIS))
for hx,h in zip(cols_x,hdrs):
    axI.text(hx+0.005, y0+0.028, h, fontsize=10, fontweight='bold', color='white',
             transform=axI.transAxes, va='center')
for r,(esp,desc,ri,oc,d,est,ecol) in enumerate(items):
    yy = y0-0.04-(r+1)*0.115
    bg = VERDE_CL if esp=='EL' else NARANJA_CL
    axI.add_patch(Rectangle((0,yy-0.045),1,0.095, transform=axI.transAxes, color=bg, alpha=0.5))
    espc = VERDE if esp=='EL' else NARANJA
    axI.text(cols_x[0]+0.005, yy, esp, fontsize=10, fontweight='bold', color=espc, transform=axI.transAxes, va='center')
    axI.text(cols_x[1]+0.005, yy, desc, fontsize=10, color=GRIS, transform=axI.transAxes, va='center')
    axI.text(cols_x[2]+0.005, yy, ri, fontsize=10, color=GRIS, transform=axI.transAxes, va='center')
    axI.text(cols_x[3]+0.005, yy, oc, fontsize=10, color=GRIS, transform=axI.transAxes, va='center')
    axI.text(cols_x[4]+0.005, yy, d, fontsize=11, fontweight='bold', color=ecol, transform=axI.transAxes, va='center')
    # badge estado
    axI.add_patch(FancyBboxPatch((cols_x[5]-0.001, yy-0.028),0.16,0.056,
                  boxstyle='round,pad=0.005,rounding_size=0.02', transform=axI.transAxes,
                  facecolor=ecol, edgecolor='none'))
    axI.text(cols_x[5]+0.079, yy, est, fontsize=8.5, fontweight='bold', color='white',
             ha='center', transform=axI.transAxes, va='center')
axI.text(0.0,-0.04,'(*) al corte 11/09/26. +7 OCs esta semana (IN+EL 25→32). 6 de 7 críticos con OC. Válvulas: único sin OC (141 d) – pendiente validación final para adjudicar.  '
                   'OC efectiva = KOM/emisión en adjudicados.',
         fontsize=8.5, style='italic', color=GRIS, transform=axI.transAxes)

plt.savefig(OUT_PNG, dpi=130, bbox_inches='tight', facecolor='white')
print(f'OK -> {os.path.basename(OUT_PNG)}')
