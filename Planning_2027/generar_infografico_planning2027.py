#!/usr/bin/env python3
"""
Infográfico PNG — Gantt Integrado CPF2 La Calera II  |  Planning 2027
RFSU: 20-DIC-2027  |  ABB + Inauco + Cables EL/IN
"""
import os
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
from matplotlib.patches import FancyBboxPatch, Rectangle, FancyArrowPatch
import matplotlib.dates as mdates
from datetime import date, datetime, timedelta
import numpy as np

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_PNG    = os.path.join(SCRIPT_DIR, 'Gantt_Integrado_CPF2_Planning2027_Rev1_120626.png')

# Paleta
AZ   = '#1F3864'; AZM  = '#2F5496'; AZC  = '#DEEAF1'
S3C  = '#C00000'; S3B  = '#880000'; S3L  = '#FFE8E8'
S4C  = '#2E75B6'; S4B  = '#1F4E79'; S4L  = '#DEEAF1'
PMSC = '#B8860B'; PMSB = '#7B5B00'; PMSL = '#FFF8DC'
PCSC = '#548235'; PCSB = '#375623'; PCSL = '#E2EFDA'
ELC  = '#E36C09'; ELB  = '#833C11'; ELL  = '#FFF0E0'
INC  = '#9E4BB5'; INB  = '#7030A0'; INL  = '#F0E8FF'
CMPC = '#7030A0'; CMPL = '#EDE8F5'
RFSU_C = '#C00000'; RFSU_L = '#FFE0E0'
REST_C = '#FF8000'; GRIS = '#595959'
HOJE = '#FFCC00'

TODAY = date(2026, 6, 12)
RFSU  = date(2027, 12, 20)

def d2dt(d): return datetime(d.year, d.month, d.day)

# ── Actividades del Gantt (nombre, ini, fin, color_barra, color_texto, y_pos, tipo) ──
ROWS = [
    # GRUPO LABEL
    ("PROVISIÓN",               None,               None,               AZ,    'white', 0,  'grp'),
    ("OC ABB SE#3+SE#4 [OK]",   date(2026,5,12),    date(2026,5,12),    S4B,   'white', 1,  'hito'),
    ("OC ABB PMS [OK]",          date(2026,5,22),    date(2026,5,22),    PMSB,  'white', 2,  'hito'),
    ("OC Inauco PCS [pend.]",    date(2026,6,9),     date(2026,6,9),     PCSB,  'white', 3,  'hito'),
    ("* OC Cables IN TARGET",    date(2026,8,1),     date(2026,8,1),     INB,   'white', 4,  'fp'),
    ("* OC Cables EL TARGET",    date(2026,10,1),    date(2026,10,1),    ELB,   'white', 5,  'fp'),

    ("ABB — FABRICACIÓN",       None,               None,               S3B,   'white', 6,  'grp'),
    ("◆ FP-S3-3 Último FP ABB", date(2026,8,31),    date(2026,8,31),    S3B,   'white', 7,  'fp'),
    ("Fabricación SE#3+SE#4 (Brasil)", date(2026,9,1), date(2026,11,23), S3C, 'white', 8,  'act'),
    ("FAT+Acopio → Argentina",  date(2026,11,3),    date(2027,1,29),    AZM,   'white', 9,  'act'),
    ("Fabricación PMS (BsAs)",  date(2026,9,11),    date(2027,1,29),    PMSC,  'white', 10, 'act'),

    ("ABB — SHELTER MENDOZA",   None,               None,               S4B,   'white', 11, 'grp'),
    ("Montaje Shelter SE#4",    date(2027,1,18),    date(2027,3,5),     S4C,   'white', 12, 'act'),
    ("Montaje Shelter SE#3 ← CRÍTICA", date(2027,2,1), date(2027,3,26), S3C,  'white', 13, 'act'),
    ("Montaje PMS + FAT Integral", date(2027,2,18), date(2027,3,31),    PMSC,  'white', 14, 'act'),
    ("FAT SE#4 Shelter",        date(2027,3,8),     date(2027,3,16),    S4B,   'white', 15, 'fat'),
    ("FAT SE#3 Shelter",        date(2027,4,2),     date(2027,4,19),    S3B,   'white', 16, 'fat'),

    ("ABB — ENTREGA A CPF2",    None,               None,               S3B,   'white', 17, 'grp'),
    ("★ Llegada SE#4 a CPF2",   date(2027,4,8),     date(2027,4,8),     S4B,   'white', 18, 'hito'),
    ("★ Llegada SE#3 a CPF2 ← ANCLA EL", date(2027,5,14), date(2027,5,14), S3B, 'white', 19, 'hito'),

    ("INAUCO — PCS SALA 7",     None,               None,               PCSB,  'white', 20, 'grp'),
    ("◆ Freezing Point PCS",    date(2026,9,30),    date(2026,9,30),    PCSB,  'white', 21, 'fp'),
    ("Fabricación tableros PCS",date(2026,11,10),   date(2027,1,26),    PCSC,  'white', 22, 'act'),
    ("FAT PCS Neuquén",         date(2027,1,27),    date(2027,2,19),    PCSC,  'white', 23, 'fat'),
    ("Integración PCS↔PMS ABB", date(2027,2,24),    date(2027,3,2),     PCSC,  'white', 24, 'act'),
    ("★ INICIO PCS SALA 7 ← ANCLA IN", date(2027,3,17), date(2027,3,17), PCSB, 'white', 25, 'hito'),
    ("SAT PCS + SIS campo",     date(2027,3,17),    date(2027,8,3),     PCSC,  'white', 26, 'act'),

    ("CABLES INSTRUMENTACIÓN (IN)", None,           None,               INB,   'white', 27, 'grp'),
    ("Llegada cables IN",        date(2026,11,28),  date(2026,11,28),   INB,   'white', 28, 'hito'),
    ("Canaletas IN",             date(2026,11,28),  date(2027,1,31),    INC,   'white', 29, 'act'),
    ("[!] TENDIDO IN → antes 17-MAR (alcance IN PENDIENTE, estim. 300-500 cab.)", date(2027,2,1), date(2027,3,14), REST_C,'white', 30, 'rest'),
    ("Conexionado IN (inicia CON PCS)", date(2027,3,17), date(2027,7,31), INC, 'white', 31, 'act'),

    ("CABLES ELÉCTRICOS (EL)",  None,               None,               ELB,   'white', 32, 'grp'),
    ("Llegada cables EL",       date(2027,2,28),    date(2027,2,28),    ELB,   'white', 33, 'hito'),
    ("Bandejas portacables",    date(2027,3,1),     date(2027,4,5),     ELC,   'white', 34, 'act'),
    ("[!] TENDIDO EL → antes SE#3", date(2027,3,1),  date(2027,5,13),    REST_C,'white', 35, 'rest'),
    ("Conexionado EL S4 (tras armado 14d, inicia 23-ABR)", date(2027,4,23), date(2027,6,12), ELC, 'white', 36, 'act'),
    ("Conexionado EL S3 (tras armado 21d, inicia 05-JUN) <- CRITICA", date(2027,6,5), date(2027,7,25), ELC, 'white', 37, 'act'),

    ("CAMPO — COMISIONADO",     None,               None,               CMPC,  'white', 38, 'grp'),
    ("Armado shelter SE#4 en campo (14d min.)", date(2027,4,8), date(2027,4,22), S4C, 'white', 39, 'act'),
    ("Armado shelter SE#3 en campo (21d) <- CRITICA", date(2027,5,14), date(2027,6,3), S3C, 'white', 40, 'act'),
    ("Megger ITP (26-JUL→15-AGO)", date(2027,7,26), date(2027,8,15),    CMPC,  'white', 41, 'fat'),
    ("Energizacion MT→BT (16-AGO→01-SEP)", date(2027,8,16), date(2027,9,1), CMPC, 'white', 42, 'fat'),
    ("iFAT integracion TODOS",  date(2027,6,25),    date(2027,7,14),    AZ,    'white', 43, 'fat'),
    ("Precomisionado electrico (02-SEP→22-SEP)", date(2027,9,2), date(2027,9,22), CMPC, 'white', 44, 'act'),
    ("Comisionado integrado (23-SEP→19-DIC)", date(2027,9,23), date(2027,12,19), CMPC, 'white', 45, 'act'),
    ("SAT PMS (HOLD POINT PIE OCT-DIC)", date(2027,10,6), date(2027,12,5), PMSC, 'white', 46, 'fat'),

    ("*** RFSU — 20-DIC-2027",  date(2027,12,20),   date(2027,12,20),   RFSU_C,'white', 47, 'hito'),
]

N_ROWS   = 48
BAR_H    = 0.55
GRP_H    = 0.35
Y_SCALE  = 1.0

fig_h = N_ROWS * 0.36 + 2.5
fig_w = 24
fig, ax = plt.subplots(figsize=(fig_w, fig_h))
fig.patch.set_facecolor('white')
ax.set_facecolor('#F9F9F9')

# Timeline range
x_min = d2dt(date(2026, 5, 1))
x_max = d2dt(date(2028, 1, 15))
ax.set_xlim(x_min, x_max)
ax.set_ylim(-1, N_ROWS + 0.5)
ax.invert_yaxis()

# Grid vertical por mes
cursor = date(2026, 5, 1)
while cursor <= date(2028, 2, 1):
    ax.axvline(d2dt(cursor), color='#DDDDDD', lw=0.5, zorder=0)
    if cursor.month == 1:
        ax.axvline(d2dt(cursor), color='#BBBBBB', lw=1.2, zorder=0)
    cursor = date(cursor.year + (1 if cursor.month == 12 else 0),
                  1 if cursor.month == 12 else cursor.month + 1, 1)

# Línea HOY
ax.axvline(d2dt(TODAY), color=HOJE, lw=2.5, zorder=5, alpha=0.8)
ax.text(d2dt(TODAY), -0.5, 'HOY\n12-JUN-26', ha='center', va='bottom',
        fontsize=7, color='#CC6600', fontweight='bold')

# Línea RFSU
ax.axvline(d2dt(RFSU), color=RFSU_C, lw=2.5, zorder=5, alpha=0.8, linestyle='--')
ax.text(d2dt(RFSU), -0.5, '★RFSU\n20-DIC-27', ha='center', va='bottom',
        fontsize=7, color=RFSU_C, fontweight='bold')

# Líneas de restricción (anclas)
for ancla_d, ancla_label, ancla_col in [
    (date(2027, 3, 17), '← ANCLA IN\nPCS 17-MAR', INB),
    (date(2027, 5, 14), '← ANCLA EL\nSE#3 14-MAY', S3B),
]:
    ax.axvline(d2dt(ancla_d), color=ancla_col, lw=1.5, zorder=4, alpha=0.6, linestyle=':')
    ax.text(d2dt(ancla_d), N_ROWS - 0.5, ancla_label, ha='center', va='top',
            fontsize=6.5, color=ancla_col, fontweight='bold',
            bbox=dict(boxstyle='round,pad=0.2', facecolor='white', alpha=0.8, edgecolor=ancla_col))

# Dibujar filas
for (nombre, ini, fin, bar_col, txt_col, y_pos, tipo) in ROWS:
    y = y_pos

    if tipo == 'grp':
        # Fondo de grupo
        rect = Rectangle((mdates.date2num(x_min), y - GRP_H/2),
                          mdates.date2num(x_max) - mdates.date2num(x_min),
                          GRP_H,
                          transform=ax.get_xaxis_transform() if False else ax.transData,
                          facecolor=bar_col, alpha=0.9, zorder=2)
        ax.add_patch(rect)
        ax.barh(y, mdates.date2num(x_max) - mdates.date2num(x_min),
                left=mdates.date2num(x_min), height=GRP_H,
                color=bar_col, alpha=1.0, zorder=2)
        ax.text(d2dt(date(2026, 5, 8)), y, f'  ▶  {nombre}',
                va='center', ha='left', fontsize=8.5, fontweight='bold',
                color='white', zorder=3)
        continue

    # Fondo de fila alterno
    ax.axhspan(y - 0.5, y + 0.5,
               facecolor='#F5F5F5' if y % 2 == 0 else '#EFEFEF',
               alpha=0.3, zorder=0)

    # Etiqueta
    ax.text(d2dt(date(2026, 5, 3)), y, nombre,
            va='center', ha='left', fontsize=7.2,
            color=bar_col if tipo not in ('grp',) else 'white',
            fontweight='bold' if tipo in ('hito', 'fp', 'rest') else 'normal',
            zorder=3)

    if ini is None: continue

    # Barra o hito
    if tipo == 'hito':
        ax.plot(d2dt(ini), y, marker='D', color=bar_col, markersize=9,
                zorder=5, markeredgecolor='white', markeredgewidth=0.8)
        fecha_str = ini.strftime('%d/%m/%y')
        ax.text(d2dt(ini) + timedelta(days=3), y, fecha_str,
                va='center', ha='left', fontsize=6.5, color=bar_col, fontweight='bold', zorder=4)

    elif tipo == 'fp':
        ax.plot(d2dt(ini), y, marker='^', color=bar_col, markersize=9,
                zorder=5, markeredgecolor='white', markeredgewidth=0.8)
        fecha_str = ini.strftime('%d/%m/%y')
        ax.text(d2dt(ini) + timedelta(days=3), y, fecha_str,
                va='center', ha='left', fontsize=6.5, color=bar_col, fontweight='bold', zorder=4)

    else:
        dur = (fin - ini).days
        alpha = 0.95
        lw = 1.0
        ec  = 'white'
        if tipo == 'rest':
            # Barras de restricción con patrón
            ax.barh(y, dur, left=mdates.date2num(d2dt(ini)),
                    height=BAR_H, color=bar_col, alpha=0.85, zorder=3,
                    edgecolor=ec, linewidth=lw, hatch='///')
        elif tipo == 'fat':
            ax.barh(y, dur, left=mdates.date2num(d2dt(ini)),
                    height=BAR_H * 0.8, color=bar_col, alpha=alpha, zorder=3,
                    edgecolor=ec, linewidth=lw)
            # Bordes más marcados para FAT
            for side in ['left', 'right']:
                x_edge = mdates.date2num(d2dt(ini if side=='left' else fin))
                ax.plot([x_edge, x_edge], [y - BAR_H*0.4, y + BAR_H*0.4],
                        color=bar_col, lw=2, zorder=4)
        else:
            ax.barh(y, dur, left=mdates.date2num(d2dt(ini)),
                    height=BAR_H, color=bar_col, alpha=alpha, zorder=3,
                    edgecolor=ec, linewidth=lw)

# Eje X — fechas
ax.xaxis_date()
ax.xaxis.set_major_locator(mdates.MonthLocator())
ax.xaxis.set_major_formatter(mdates.DateFormatter('%b\n%Y'))
ax.tick_params(axis='x', labelsize=7.5, pad=2)
ax.xaxis.set_tick_params(which='major', gridOn=True, grid_color='#DDDDDD', grid_lw=0.5)

# Ocultar eje Y
ax.set_yticks([])
ax.spines['top'].set_visible(False)
ax.spines['right'].set_visible(False)
ax.spines['left'].set_visible(False)

# Título
fig.text(0.5, 0.99,
         'GANTT INTEGRADO — CPF2 La Calera II (Vaca Muerta)  ·  Planning 2027  ·  Rev1_120626',
         ha='center', va='top', fontsize=14, fontweight='bold', color=AZ)
fig.text(0.5, 0.975,
         'ABB (SE#3 · SE#4 · PMS)  +  INAUCO (PCS/SCADA Sala 7)  +  Cables EL/IN  |  '
         'RFSU: 20-DIC-2027  |  Shelters en modulos: armado in situ SE#4 (14d) + SE#3 (21d)  |  Corte: 12-JUN-2026',
         ha='center', va='top', fontsize=9, color=GRIS)

# Restricciones en banner
fig.text(0.5, 0.955,
         '[!] RESTRICCION IN: Alcance PENDIENTE (~300-500 cab./40,000-70,000m estimado, AT en proceso). '
         'Tendido ANTES 17-MAR-2027. Requiere 5-6 cuadrillas.  ·  '
         '[!] RESTRICCION EL: Tendido ANTES 14-MAY-2027 (llegada SE#3).  ·  '
         'Conexionado EL inicia TRAS armado shelters en campo (SE#4: +14d, SE#3: +21d)',
         ha='center', va='top', fontsize=8, color='#7B4000',
         bbox=dict(boxstyle='round,pad=0.4', facecolor='#FFF0CC', alpha=0.95, edgecolor='#E36C09'))

# Leyenda
legend_items = [
    mpatches.Patch(color=S3C, label='ABB SE#3 (cadena crítica)'),
    mpatches.Patch(color=S4C, label='ABB SE#4'),
    mpatches.Patch(color=PMSC, label='ABB PMS'),
    mpatches.Patch(color=PCSC, label='Inauco PCS/SCADA'),
    mpatches.Patch(color=INC, label='Cables IN'),
    mpatches.Patch(color=ELC, label='Cables EL'),
    mpatches.Patch(color=CMPC, label='Campo / Comisionado'),
    mpatches.Patch(color=REST_C, label='[!] Restricción de secuencia (hatch)'),
    plt.Line2D([0],[0], marker='D', color='w', markerfacecolor=AZ, markersize=8, label='Hito'),
    plt.Line2D([0],[0], marker='^', color='w', markerfacecolor=S3B, markersize=8, label='Freezing Point'),
    plt.Line2D([0],[0], color=HOJE, lw=2, label='Hoy (12-JUN-26)'),
    plt.Line2D([0],[0], color=RFSU_C, lw=2, linestyle='--', label='RFSU (20-DIC-27)'),
]
ax.legend(handles=legend_items, loc='lower right', fontsize=6.5,
          ncol=4, framealpha=0.95, edgecolor='#CCCCCC',
          bbox_to_anchor=(1.0, -0.12))

plt.tight_layout(rect=[0, 0.05, 1, 0.94])
plt.savefig(OUT_PNG, dpi=140, bbox_inches='tight', facecolor='white')
print(f'OK -> {os.path.basename(OUT_PNG)}')
