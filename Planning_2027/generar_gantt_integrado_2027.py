#!/usr/bin/env python3
"""
Gantt Integrado + Diagrama de Flujo — CPF2 La Calera II  |  Planning 2027
Integra: ABB (SE#3, SE#4, PMS) + Inauco (PCS Sala 7) + Cables EL/IN
RFSU objetivo: 20-DIC-2027
Fuentes:
  - ABB branch: CRONOGRAMA_GENERAL_ABB_Actualizado_2026-06-09.xlsx
  - ABB branch: Analisis_Electrico_CPF2_Campo_Rev2_2026-06-09.xlsx
  - Inauco branch: Dashboard_Inauco_Seguimiento.xlsx
Corte: 12-JUN-2026
"""
import os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_XLSX   = os.path.join(SCRIPT_DIR, 'Gantt_Integrado_CPF2_Planning2027.xlsx')

# ── PALETA ────────────────────────────────────────────────────────────────────
C = {
    # Headers/titles
    "azul":   "1F3864", "azul_m": "2F5496", "azul_cl": "DEEAF1",
    # ABB colores
    "s3_hdr": "880000", "s3_bar": "C00000", "s3_bg": "FFE8E8",
    "s4_hdr": "1F4E79", "s4_bar": "2E75B6", "s4_bg": "DEEAF1",
    "pms_hdr":"7B5B00", "pms_bar":"B8860B", "pms_bg": "FFF8DC",
    # Inauco / PCS
    "pcs_hdr":"375623", "pcs_bar":"548235", "pcs_bg": "E2EFDA",
    # Cables
    "el_hdr": "833C11", "el_bar": "E36C09", "el_bg":  "FFF0E0",
    "in_hdr": "7030A0", "in_bar": "9E4BB5", "in_bg":  "F0E8FF",
    # Campo / Comisionado
    "cmp_hdr":"4A0050","cmp_bar": "7030A0","cmp_bg":  "EDE8F5",
    # Hitos especiales
    "rfsu":   "C00000", "rfsu_bg":"FFE0E0",
    "fp":     "C00000", "hito":   "CC0000",
    "ok":     "375623", "ok_bg":  "C6EFCE",
    "warn":   "9C5700", "warn_bg":"FFEB9C",
    "risk":   "9C0006", "risk_bg":"FFC7CE",
    # Restricciones
    "rest_bg":"FFF0CC", "rest_t": "7B4000",
    # General
    "gray":   "595959", "gray_cl":"F2F2F2", "white": "FFFFFF",
    "sep":    "404040",
}

TODAY = date(2026, 6, 12)
RFSU  = date(2027, 12, 20)

def fl(c):   return PatternFill("solid", fgColor=c)
def fn(bold=False, color="000000", size=9, italic=False, name="Calibri"):
    return Font(bold=bold, color=color, size=size, italic=italic, name=name)
def al(h="center", v="center", wrap=False):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)
def thin():  s = Side(style="thin", color="CCCCCC"); return Border(left=s,right=s,top=s,bottom=s)
def medium():s = Side(style="medium",color="888888"); return Border(left=s,right=s,top=s,bottom=s)

def cell(ws, r, c_, v=None, bold=False, fg="000000", bg=None, sz=9,
         h="center", wrap=False, ita=False, bdr=True, nf=None):
    x = ws.cell(r, c_, v)
    x.font = fn(bold, fg, sz, ita)
    x.alignment = al(h, "center", wrap)
    if bg:  x.fill = fl(bg)
    if bdr: x.border = thin()
    if nf:  x.number_format = nf
    return x

def hdr(ws, r, c_, v, bg, fg="FFFFFF", bold=True, sz=9, h="center", wrap=False):
    x = ws.cell(r, c_, v)
    x.fill = fl(bg); x.font = fn(bold, fg, sz)
    x.alignment = al(h, "center", wrap); x.border = thin()
    return x

MESES = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

def fmt(d): return d.strftime("%d/%m/%y") if d else "—"

# ── TIMELINE ─────────────────────────────────────────────────────────────────
INICIO  = date(2026, 5, 1)
FIN     = date(2028, 1, 31)
SEMANAS = []
d = INICIO
while d <= FIN:
    SEMANAS.append(d)
    d += timedelta(weeks=1)
N_INFO  = 5     # columnas de texto (Grupo, ID, Actividad, Inicio, Fin)
COL_OFF = N_INFO + 1

def sem_rango(ini, fin):
    ci = cf = None
    for i, s in enumerate(SEMANAS):
        sw = s + timedelta(days=6)
        if ci is None and sw >= ini: ci = i
        if s <= fin: cf = i
    return ci, cf

# ── ACTIVIDADES ───────────────────────────────────────────────────────────────
# (grupo_key, id, nombre, ini, fin, bar_color, bg_color, tipo, nota)
# tipo: act | fp | hito | fat | desp | pago | rest
ACT = [
    # ── PROVISIÓN (OC ya colocadas) ───────────────────────────────────────────
    ("PROV","OC-S3S4","OC ABB Salas Eléctricas #3+#4 — aceptada",
     date(2026,5,12),date(2026,5,12),"hito","1F4E79","FFFFFF","hito",
     "OC 4508944971 (S3 USD 8.1M) + OC 4508944973 (S4 USD 3.7M). Aceptación Christian Bombace"),

    ("PROV","OC-PMS","OC ABB PMS — aceptada",
     date(2026,5,22),date(2026,5,22),"hito","B8860B","FFFFFF","hito",
     "OC 4508945953. KOM 22-MAY. PM Pablo Kalis"),

    ("PROV","OC-PCS","OC Inauco PCS/SCADA/Comms — pendiente",
     date(2026,6,9),date(2026,6,9),"hito","548235","FFFFFF","hito",
     "OFT-128-008 Rev3. Pendiente OC formal. Evaluación RECOMENDA ADJUDICAR"),

    ("PROV","AT-IN","AT Cables Instrumentación — en proceso",
     date(2026,6,12),date(2026,7,25),"act","9E4BB5","F0E8FF","act",
     "Ofertas enviadas a AT 12-JUN-2026. Meta: adjudicación ≤01-AGO-2026 para habilitar tendido antes 17-MAR-2027"),

    ("PROV","OC-IN","★ OC Cables Instrumentación — TARGET",
     date(2026,8,1),date(2026,8,1),"fp","7030A0","FFFFFF","fp",
     "TARGET: 01-AGO-2026. LT=150d → entrega en sitio 28-NOV-2026. CRITICO para tendido antes PCS 17-MAR-2027"),

    ("PROV","AT-EL","AT Cables Eléctricos — inicio",
     date(2026,6,12),date(2026,9,15),"act","E36C09","FFF0E0","act",
     "RI emitida 11-JUN. Sin SOLPED aún. Meta: OC ≤01-OCT-2026"),

    ("PROV","OC-EL","★ OC Cables Eléctricos — TARGET",
     date(2026,10,1),date(2026,10,1),"fp","833C11","FFFFFF","fp",
     "TARGET: 01-OCT-2026. LT=150d → entrega en sitio 28-FEB-2027. Para tendido antes llegada SE#4 (08-ABR)"),

    # ── ABB FABRICACIÓN (fuera del sitio) ────────────────────────────────────
    ("ABB_FAB","FP-S3","◆ FREEZING POINT — Ing. Constructiva S#3 ← ÚLTIMO FP ABB",
     date(2026,8,31),date(2026,8,31),"fp",C["s3_hdr"],"FFFFFF","fp",
     "★ 31-AGO-2026: ÚLTIMO FP DE TODO EL PROYECTO ABB. Cierra ingeniería S#3"),

    ("ABB_FAB","FAB-S4","Fabricación tableros S#4 (Brasil M&B)",
     date(2026,9,1),date(2026,11,2),"act",C["s4_bar"],C["s4_bg"],"act",
     "25 columnas ZS1 630A 25kA MT. Simultáneo con S#3"),

    ("ABB_FAB","FAB-S3","Fabricación 37 tableros BT S#3 (Brasil M&B) ← CADENA CRÍTICA",
     date(2026,9,14),date(2026,11,2),"act",C["s3_bar"],C["s3_bg"],"act",
     "37 CCMs 4000A 65kA. CADENA CRÍTICA — cualquier demora impacta RFSU"),

    ("ABB_FAB","FAB-PMS","Fabricación tableros PMS-001+PMS-101 (Bs.As.)",
     date(2026,9,11),date(2026,12,17),"act",C["pms_bar"],C["pms_bg"],"act",
     "PMS-001 para SE#3, PMS-101 para SE#4. AC800M redundante. IP54 Rittal"),

    ("ABB_FAB","FAT-BR","FAT S#3+S#4 en fábrica Brasil",
     date(2026,11,3),date(2026,11,23),"fat",C["s3_bar"],C["s3_bg"],"fat",
     "20 días. AESA inspecciona. Witness Point. Ambas salas simultáneo"),

    ("ABB_FAB","ACOP","Acopio S#3+S#4 en Argentina",
     date(2026,12,7),date(2027,1,29),"act",C["azul_m"],C["azul_cl"],"act",
     "Ductos barras Turquía + VFDs BT/MT + 37 CCMs BT + celdas MT en Argentina"),

    ("ABB_FAB","FAT-PMS","FAT tableros PMS (Bs.As.) + Pruebas Internas",
     date(2026,12,18),date(2027,1,29),"fat",C["pms_bar"],C["pms_bg"],"fat",
     "Pruebas internas 18-DIC → 07-ENE. FAT tableros 25-29-ENE-2027 (Witness AESA)"),

    # ── ABB SHELTER — MENDOZA ────────────────────────────────────────────────
    ("ABB_SH","MONT-S4","Montaje e integración Shelter S#4 (Bottino — Mendoza)",
     date(2027,1,18),date(2027,3,5),"act",C["s4_bar"],C["s4_bg"],"act",
     "Celdas MT + VFDs + tableros BT en Shelter. Simultáneo con inicio S#3"),

    ("ABB_SH","DESP-PMS","Despacho PMS — Bs.As. → Shelter Mendoza",
     date(2027,2,4),date(2027,2,17),"desp",C["pms_bar"],C["pms_bg"],"desp",
     "Hito 30% pago PMS. Equipamiento llega al Shelter 17-FEB"),

    ("ABB_SH","MONT-S3","Montaje e integración Shelter S#3 (Bottino — Mendoza) ← CRÍTICA",
     date(2027,2,1),date(2027,3,26),"act",C["s3_bar"],C["s3_bg"],"act",
     "37 CCMs + ductos barras + VFDs BT. CADENA CRÍTICA. Simultáneo con PMS montaje (18-FEB/17-MAR)"),

    ("ABB_SH","MONT-PMS","Montaje PMS en Shelter (PMS-001 en S#3 + PMS-101 en S#4)",
     date(2027,2,18),date(2027,3,17),"act",C["pms_bar"],C["pms_bg"],"act",
     "20 días. Simultáneo S#3 montaje. Coordinar con Bottino. PMS-001 → SE#3, PMS-101 → SE#4"),

    ("ABB_SH","FAT-S4SH","FAT Sala Eléctrica #4 (Shelter completo — Mendoza)",
     date(2027,3,8),date(2027,3,16),"fat",C["s4_bar"],C["s4_bg"],"fat",
     "8 días. AESA + Cliente. Termina 9 días antes FAT Integral PMS (25-MAR). ✅ Sin superposición"),

    ("ABB_SH","FAT-PMSI","★ FAT INTEGRAL SHELTER PMS",
     date(2027,3,25),date(2027,3,31),"fat",C["pms_bar"],C["pms_bg"],"fat",
     "5 días. Hito 20% (4% OC) PMS. Termina 2 días antes FAT S#3 → BUFFER AJUSTADO ⚠"),

    ("ABB_SH","FAT-S3SH","★ FAT Sala Eléctrica #3 ← FAT MÁS TARDÍO ABB",
     date(2027,4,2),date(2027,4,19),"fat",C["s3_bar"],C["s3_bg"],"fat",
     "17 días. CADENA CRÍTICA. PMS-001 integrado y validado. +35d holgura vs P6"),

    # ── ABB DESPACHO → CAMPO ─────────────────────────────────────────────────
    ("ABB_CAMPO","DESP-S4","▶ Despacho Sala #4 → CPF2 La Calera",
     date(2027,4,2),date(2027,4,8),"desp",C["s4_bar"],C["s4_bg"],"desp",
     "OC deadline 08-ABR-2027. Holgura +81d vs P6 montaje campo"),

    ("ABB_CAMPO","LLEG-S4","★ LLEGADA SE#4 a CPF2 (con PMS-101)",
     date(2027,4,8),date(2027,4,8),"hito",C["s4_hdr"],"FFFFFF","hito",
     "08-ABR-2027: SE#4 + PMS integrado en sitio. Habilita montaje celdas MT + conexionado EL (sector SE#4)"),

    ("ABB_CAMPO","DESP-S3","▶ Despacho Sala #3 → CPF2 La Calera ← DESPACHO CRÍTICO",
     date(2027,5,10),date(2027,5,14),"desp",C["s3_bar"],C["s3_bg"],"desp",
     "OC deadline 14-MAY-2027. CADENA CRÍTICA. Holgura +45d vs P6"),

    ("ABB_CAMPO","LLEG-S3","★ LLEGADA SE#3 a CPF2 (con PMS-001) ← HITO ANCLA EL",
     date(2027,5,14),date(2027,5,14),"hito",C["s3_hdr"],"FFFFFF","hito",
     "14-MAY-2027: SE#3 + PMS integrado en sitio. HITO ANCLA: todos los cables EL deben estar tendidos para este momento"),

    # ── INAUCO — PCS SALA 7 ──────────────────────────────────────────────────
    ("INAUCO","FRZING","◆ Freezing Point Ingeniería PCS",
     date(2026,9,30),date(2026,9,30),"fp",C["pcs_hdr"],"FFFFFF","fp",
     "30-SEP-2026: Congela I/O list, P&ID, MCE, filosofía. 16 semanas antes iFAT"),

    ("INAUCO","FAB-PCS","Construcción tableros PCS Remota 7A+7B (Neuquén/Inauco)",
     date(2026,11,10),date(2027,1,26),"act",C["pcs_bar"],C["pcs_bg"],"act",
     "50 días. ControlLogix 1756 + I/O HART + EN2TP PRP + Marshalling. Simultáneo config SCADA"),

    ("INAUCO","FAT-PCS","FAT PCS + Comunicaciones (18 días — Neuquén)",
     date(2027,1,27),date(2027,2,19),"fat",C["pcs_bar"],C["pcs_bg"],"fat",
     "27-ENE → 19-FEB-2027. Incluye Stratix, VLANs OT, Modbus paquetizados. AESA asiste"),

    ("INAUCO","INT-ABB","Integración PCS↔PMS ABB (Bs.As.) — 5 días RIESGO CONCENTRADO",
     date(2027,2,24),date(2027,3,2),"act",C["pcs_bar"],C["pcs_bg"],"act",
     "24-FEB → 02-MAR-2027. 5 días. Validar mapeo Modbus TCP. R7: pedir mini-FAT virtual anticipada AGO-2026"),

    ("INAUCO","ADJ-PCS","Ajustes post-FAT + preparativos campo",
     date(2027,2,24),date(2027,3,16),"act",C["pcs_bar"],C["pcs_bg"],"act",
     "15 días. Paralelismo con integración ABB. Finaliza 16-MAR-2027"),

    ("INAUCO","LLEG-PCS","★ INICIO CAMPO — PCS instalación Sala 7 ← HITO ANCLA IN",
     date(2027,3,17),date(2027,3,17),"hito",C["pcs_hdr"],"FFFFFF","hito",
     "17-MAR-2027: HITO ANCLA. Cables IN deben estar TENDIDOS. Conexionado IN empieza CON esta instalación"),

    ("INAUCO","SAT-PCS","SAT y Comisionado PCS — Sala 7 (55 días)",
     date(2027,3,17),date(2027,6,1),"act",C["pcs_bar"],C["pcs_bg"],"act",
     "17-MAR → 01-JUN-2027. 55 días. Sala 7 (Remota 7A+7B). ~10 señales/día en condiciones ideales"),

    ("INAUCO","SAT-SIS","SAT y Comisionado SCADA SIS — HIMA (100 días)",
     date(2027,3,17),date(2027,8,3),"act",C["pcs_bar"],C["pcs_bg"],"act",
     "17-MAR → 03-AGO-2027. 100 días. Requiere instrumentos energizados. HIMA presencia en campo"),

    ("INAUCO","MCE","Prueba Matriz C&E — SCADA completo",
     date(2027,8,4),date(2027,8,24),"fat",C["pcs_bar"],C["pcs_bg"],"fat",
     "04-AGO → 24-AGO-2027. Validación final. Requiere personal SIS presente"),

    ("INAUCO","CAO","CAO / As-Built — Inauco",
     date(2027,8,25),date(2027,9,21),"act",C["pcs_bar"],C["pcs_bg"],"act",
     "25-AGO → 21-SEP-2027. Finalización proyecto Inauco"),

    # ── CABLES INSTRUMENTACIÓN (IN) ──────────────────────────────────────────
    ("CABLES_IN","LLG-IN","Llegada cables IN a sitio",
     date(2026,11,28),date(2026,11,28),"hito","7030A0","FFFFFF","hito",
     "28-NOV-2026 (LT=150d desde OC 01-AGO). Habilitante para instalación canaletas y tendido"),

    ("CABLES_IN","CAN-IN","Instalación canaletas/conduit instrumentación",
     date(2026,11,28),date(2027,1,31),"act",C["in_bar"],C["in_bg"],"act",
     "65 días. Parallel con entrega cables. Canaletas IN desde campo → Sala 7. Habilitante tendido"),

    ("CABLES_IN","TEND-IN","◀ TENDIDO cables IN — ANTES instalación PCS ▶",
     date(2027,2,1),date(2027,3,14),"act",C["in_bar"],C["in_bg"],"rest",
     "42 días | 2 cuadrillas | ~28,126m / 192 cables | "
     "RESTRICCIÓN: completar antes de 17-MAR-2027 (inicio PCS Sala 7)"),

    ("CABLES_IN","REST-IN","⛔ RESTRICCIÓN: Tendido IN completo → PCS instalado",
     date(2027,3,14),date(2027,3,17),"hito","7030A0","FFFFFF","rest",
     "GAP 3 días de buffer. Tendido IN DEBE completarse antes del 17-MAR. "
     "Depende: OC ≤01-AGO-2026 y LT 150 días cumplido"),

    ("CABLES_IN","CONX-IN","Conexionado IN — comienza CON instalación PCS",
     date(2027,3,17),date(2027,7,31),"act",C["in_bar"],C["in_bg"],"act",
     "17-MAR → 31-JUL-2027. 136 días. Peinado+etiquetado+prueba continuidad. "
     "Simultáneo con SAT PCS. Pivota sobre instalación PCS 17-MAR"),

    # ── CABLES ELÉCTRICOS (EL) ───────────────────────────────────────────────
    ("CABLES_EL","LLG-EL","Llegada cables EL a sitio",
     date(2027,2,28),date(2027,2,28),"hito","833C11","FFFFFF","hito",
     "28-FEB-2027 (LT=150d desde OC 01-OCT-2026). 513 cables CPF-2 / 81,231m"),

    ("CABLES_EL","CAN-EL","Instalación bandejas portacables principales",
     date(2027,3,1),date(2027,4,5),"act",C["el_bar"],C["el_bg"],"act",
     "35 días. Bandejas MC1/MC2/MC3 trayectos SE#3→SE#4→campo. Habilitante tendido EL"),

    ("CABLES_EL","TEND-EL","◀ TENDIDO cables EL — ANTES montaje salas en campo ▶",
     date(2027,3,1),date(2027,5,13),"act",C["el_bar"],C["el_bg"],"rest",
     "74 días | 3 cuadrillas | 81,231m (tendido potencia MT+BT+control) | "
     "RESTRICCIÓN: completar antes de montaje final SE#4 (~09-MAY) y SE#3 (~24-MAY)"),

    ("CABLES_EL","REST-EL","⛔ RESTRICCIÓN: Cables EL tendidos → Salas instaladas en campo",
     date(2027,5,13),date(2027,5,14),"hito","833C11","FFFFFF","rest",
     "Buffer 1 día. Tendido EL DEBE estar completo cuando lleguen las salas (SE#4: 08-ABR, SE#3: 14-MAY). "
     "Ruta crítica EL: cables BT grande 3x240/3x150 (104 cables, 17,470m)"),

    ("CABLES_EL","CONX-EL-S4","Conexionado EL — Sector SE#4 (inicia con llegada SE#4)",
     date(2027,4,14),date(2027,5,31),"act",C["el_bar"],C["el_bg"],"act",
     "14-ABR → 31-MAY-2027. 48 días. Terminales MT 13.2kV + BT sector SE#4. "
     "Pivota sobre llegada SE#4 (08-ABR) → inicio conexionado 14-ABR"),

    ("CABLES_EL","CONX-EL-S3","Conexionado EL — Sector SE#3 (inicia con llegada SE#3)",
     date(2027,5,18),date(2027,7,6),"act",C["el_bar"],C["el_bg"],"act",
     "18-MAY → 06-JUL-2027. 50 días. Terminales BT + control sector SE#3. "
     "Pivota sobre llegada SE#3 (14-MAY) → inicio conexionado 18-MAY"),

    # ── MONTAJE Y CAMPO ──────────────────────────────────────────────────────
    ("CAMPO","MONT-S4C","Montaje celdas MT + CCMs S#4 en CPF2",
     date(2027,4,14),date(2027,5,9),"act",C["s4_bar"],C["s4_bg"],"act",
     "14-ABR → 09-MAY-2027. 26 días. ABB+AESA. TGMT-001/003, TRMT-001A/B. Nivelación, barras, control"),

    ("CAMPO","MONT-S3C","Montaje 37 CCMs S#3 en CPF2 ← CADENA CRÍTICA",
     date(2027,5,14),date(2027,6,4),"act",C["s3_bar"],C["s3_bg"],"act",
     "14-MAY → 04-JUN-2027. 22 días. CADENA CRÍTICA. 37 CCMs + ductos barras + VFDs"),

    ("CAMPO","MEGGER","Pruebas aislación (megger) circuito por circuito — ITP ELÉCTRICO",
     date(2027,7,7),date(2027,7,28),"fat",C["cmp_bar"],C["cmp_bg"],"fat",
     "07-JUL → 28-JUL-2027. 22 días. HOLD POINT ITP. Registro conforme ITP. ≥1MΩ/1kV y ≥10MΩ/5kV"),

    ("CAMPO","ENERG","Energización progresiva MT → Transformadores → BT",
     date(2027,7,29),date(2027,8,13),"fat",C["cmp_bar"],C["cmp_bg"],"fat",
     "29-JUL → 13-AGO-2027. 16 días. ABB + AESA + Cliente. Pre-requis: Megger OK + Protecciones ajustadas"),

    ("CAMPO","PRECOM","Precomisionado eléctrico funcional",
     date(2027,8,14),date(2027,9,3),"act",C["cmp_bar"],C["cmp_bg"],"act",
     "14-AGO → 03-SEP-2027. 21 días. Loop check eléctrico + ajuste relés + secuencia arranque motores"),

    ("CAMPO","SAT-PMS","SAT PMS en sitio — HOLD POINT contractual (PIE)",
     date(2027,10,6),date(2027,12,5),"fat",C["pms_bar"],C["pms_bg"],"fat",
     "06-OCT → 05-DIC-2027. Ventana OCT-DIC fija por contrato/PIE. NO puede deslizarse. "
     "ABB field service + Cliente. 100% E/S ensayadas"),

    ("CAMPO","COM","Comisionado eléctrico integrado (EL+PCS+PMS+SIS)",
     date(2027,9,4),date(2027,12,19),"act",C["cmp_bar"],C["cmp_bg"],"act",
     "04-SEP → 19-DIC-2027. 107 días. Energizado progresivo de procesos. "
     "Coordinado con Inauco (SIS), ABB (PMS), Operaciones"),

    ("CAMPO","IFAT","iFAT — Integración SE#3+SE#4+PMS+PCS+HIMA",
     date(2027,6,25),date(2027,7,14),"fat","1F3864","BDD7EE","fat",
     "25-JUN → 14-JUL-2027. 20 días. Todos los vendors integran en campo. "
     "ABB + Inauco + HIMA + AESA + Pluspetrol"),

    # ── RFSU ─────────────────────────────────────────────────────────────────
    ("RFSU","RFSU","★★★ RFSU — READY FOR START UP",
     date(2027,12,20),date(2027,12,20),"hito",C["rfsu"],"FFFFFF","hito",
     "20-DIC-2027. Entrega al cliente. Todos los sistemas operativos. "
     "ABB garantía extendida 20-DIC-2028"),
]

# ── GRUPOS ────────────────────────────────────────────────────────────────────
GRUPO_META = {
    "PROV":       ("PROVISIÓN DE SUMINISTROS  |  OC y Gestión de Compra",          C["azul"],    "DEEAF1"),
    "ABB_FAB":    ("ABB — FABRICACIÓN (Brasil + Bs.As.)  |  SE#3 · SE#4 · PMS",   C["s3_hdr"],  "FFE8E8"),
    "ABB_SH":     ("ABB — MONTAJE SHELTER MENDOZA (Bottino)  |  SE#3 · SE#4 · PMS",C["s4_hdr"], "DEEAF1"),
    "ABB_CAMPO":  ("ABB — DESPACHO Y ENTREGA A CPF2  |  Hitos ancla Cables EL",    C["azul"],   "E8F0FF"),
    "INAUCO":     ("INAUCO — PCS/SCADA/COMMS  |  Sala 7 (Control Room) · Hito ancla Cables IN", C["pcs_hdr"], "E2EFDA"),
    "CABLES_IN":  ("CABLES INSTRUMENTACIÓN (IN)  |  ⛔ Tendido ANTES de instalación PCS 17-MAR-2027", C["in_hdr"], "F0E8FF"),
    "CABLES_EL":  ("CABLES ELÉCTRICOS (EL)  |  ⛔ Tendido ANTES de llegada Salas SE#4 08-ABR / SE#3 14-MAY", C["el_hdr"], "FFF0E0"),
    "CAMPO":      ("CAMPO — MONTAJE, PRUEBAS Y COMISIONADO  |  CPF2 La Calera II", C["cmp_hdr"], "EDE8F5"),
    "RFSU":       ("★ RFSU  |  READY FOR START UP — 20-DIC-2027",                  C["rfsu"],    "FFE0E0"),
}

BAR_TIPO = {
    "hito":  ("★", 7),
    "fp":    ("◆", 7),
    "fat":   (None, None),
    "desp":  (None, None),
    "act":   (None, None),
    "rest":  ("!", 7),
    "pago":  ("§", 7),
}

# ── HOJA 1: GANTT GENERAL ─────────────────────────────────────────────────────
def sheet_gantt(wb):
    ws = wb.create_sheet("1_Gantt_Integrado_2027")
    ws.sheet_view.showGridLines = False

    # Anchos info
    for i, w in enumerate([14, 8, 44, 10, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # Anchos semanas
    for i in range(len(SEMANAS)):
        ws.column_dimensions[get_column_letter(COL_OFF + i)].width = 1.8

    last_col = get_column_letter(COL_OFF + len(SEMANAS) - 1)

    # Fila 1: Título principal
    ws.merge_cells(f"A1:{last_col}1"); ws.row_dimensions[1].height = 34
    hdr(ws, 1, 1,
        "GANTT INTEGRADO — PLANNING 2027  |  CPF2 La Calera II (Vaca Muerta)  |  "
        "ABB (SE#3 · SE#4 · PMS) + INAUCO (PCS Sala 7) + CABLES EL/IN  |  "
        "RFSU: 20-DIC-2027  |  Corte: 12-JUN-2026",
        C["azul"], sz=12)

    # Fila 2: Subtítulo con restricciones clave
    ws.merge_cells(f"A2:{last_col}2"); ws.row_dimensions[2].height = 16
    sub = (
        "RESTRICCIONES:  ⛔ Cables IN tendidos ANTES de 17-MAR-2027 (instalación PCS Sala 7)  ·  "
        "⛔ Cables EL tendidos ANTES de llegada SE#4 (08-ABR) y SE#3 (14-MAY-2027)  ·  "
        "Conexionado EL pivota sobre llegada salas  ·  Conexionado IN comienza CON instalación PCS"
    )
    x = ws.cell(2, 1, sub)
    x.fill = fl(C["rest_bg"]); x.font = fn(True, C["rest_t"], 9)
    x.alignment = al("left", "center"); x.border = thin()

    # Fila 3: Años
    ws.row_dimensions[3].height = 14
    anio_g = {}
    for i, s in enumerate(SEMANAS):
        a = s.year
        if a not in anio_g: anio_g[a] = [i, i]
        else: anio_g[a][1] = i
    for a, (ci, cf) in anio_g.items():
        try: ws.merge_cells(f"{get_column_letter(COL_OFF+ci)}3:{get_column_letter(COL_OFF+cf)}3")
        except: pass
        hdr(ws, 3, COL_OFF + ci, str(a), C["azul"], sz=9)

    # Fila 4: Meses
    ws.row_dimensions[4].height = 11
    mes_g = {}
    for i, s in enumerate(SEMANAS):
        k = (s.year, s.month)
        if k not in mes_g: mes_g[k] = [i, i]
        else: mes_g[k][1] = i
    for (a, m), (ci, cf) in mes_g.items():
        try: ws.merge_cells(f"{get_column_letter(COL_OFF+ci)}4:{get_column_letter(COL_OFF+cf)}4")
        except: pass
        hdr(ws, 4, COL_OFF + ci, MESES[m], "2E75B6" if m % 2 == 0 else "404040", sz=7)

    # Fila 5: Encabezados columnas info
    ws.row_dimensions[5].height = 20
    for i, h_ in enumerate(["Bloque", "ID", "Actividad / Hito", "Inicio", "Fin"], 1):
        hdr(ws, 5, i, h_, C["azul"], sz=9)

    # Marcar hoy en la timeline (solo en fila 5 que no está mergeada)
    hoy_ci, _ = sem_rango(TODAY, TODAY)
    if hoy_ci is not None:
        c_hoy = ws.cell(5, COL_OFF + hoy_ci)
        c_hoy.fill = fl("FFFF00")
        c_hoy.value = "HOY"
        c_hoy.font = fn(True, "CC0000", 6)
        c_hoy.alignment = al("center", "center")

    row = 6
    grupo_ant = None

    for rec in ACT:
        (grupo_key, id_, nombre, ini, fin, bar_c_key, bar_col, bg_col, tipo, nota) = rec

        # Separador de grupo
        if grupo_key != grupo_ant:
            g_titulo, g_hdr_c, g_bg = GRUPO_META[grupo_key]
            ws.merge_cells(f"A{row}:{last_col}{row}")
            ws.row_dimensions[row].height = 16
            x = ws.cell(row, 1, f"  ▶  {g_titulo}")
            x.fill = fl(g_hdr_c)
            x.font = fn(True, "FFFFFF", 9)
            x.alignment = al("left", "center")
            x.border = thin()
            row += 1
            grupo_ant = grupo_key

        g_titulo, g_hdr_c, g_bg = GRUPO_META[grupo_key]
        ws.row_dimensions[row].height = 13

        es_hito = tipo in ("hito", "fp")
        es_rest = tipo == "rest"
        row_bg  = ("FFE0E0" if es_rest else
                   ("FFF8C0" if es_hito else g_bg))

        # Determinar color de barra real
        tipo_col_map = {
            "fat":   "#" + bar_col if bar_col.startswith("FF") else bar_col,
            "desp":  "7030A0",
            "hito":  bar_col,
            "fp":    bar_col,
            "rest":  "FF8000",
            "act":   bar_col,
        }
        barra_color = tipo_col_map.get(tipo, bar_col)

        # Texto celdas info
        cell(ws, row, 1, grupo_key, bold=False, fg=g_hdr_c, bg=row_bg, sz=7, h="center")
        cell(ws, row, 2, id_,       bold=True,  fg=g_hdr_c, bg=row_bg, sz=7.5, h="center")
        cell(ws, row, 3, nombre,    bold=(es_hito or es_rest), fg=(C["rfsu"] if "RFSU" in nombre else ("C00000" if es_rest else g_hdr_c)),
             bg=row_bg, sz=9, h="left")
        cell(ws, row, 4, fmt(ini),  bold=es_hito, fg=g_hdr_c, bg=row_bg, sz=8, h="center")
        cell(ws, row, 5, fmt(fin),  bold=es_hito, fg=g_hdr_c, bg=row_bg, sz=8, h="center")

        # Barra Gantt
        ci_s, cf_s = sem_rango(ini, fin)
        for i in range(len(SEMANAS)):
            s_d = SEMANAS[i]
            c2 = ws.cell(row, COL_OFF + i)

            # Marcar HITO SE#3 llegada como línea vertical especial
            if hoy_ci is not None and i == hoy_ci:
                c2.border = Border(
                    left=Side(style="medium", color="FFCC00"),
                    right=Side(style="thin", color="DDDDDD"),
                    top=Side(style="thin", color="DDDDDD"),
                    bottom=Side(style="thin", color="DDDDDD"))
            else:
                lw = "medium" if s_d.day <= 7 else "thin"
                c2.border = Border(
                    left=Side(style=lw, color="888888" if s_d.day <= 7 else "DDDDDD"),
                    right=Side(style="thin", color="DDDDDD"),
                    top=Side(style="thin", color="DDDDDD"),
                    bottom=Side(style="thin", color="DDDDDD"))

            in_bar = ci_s is not None and cf_s is not None and ci_s <= i <= cf_s
            if in_bar:
                c2.fill = fl(barra_color)
                sym, sym_sz = BAR_TIPO.get(tipo, (None, None))
                if sym and i == ci_s:
                    c2.value = sym
                    c2.font = fn(True, "FFFFFF", sym_sz or 7)
                    c2.alignment = al("center", "center")
            else:
                c2.fill = fl("F8F8F8" if i % 2 == 0 else "F3F3F3")

        row += 1

    # Línea RFSU
    ws.merge_cells(f"A{row}:{last_col}{row}")
    ws.row_dimensions[row].height = 18
    hdr(ws, row, 1,
        "★ RFSU 20-DIC-2027  |  Cadena crítica: SE#3 llega 14-MAY → montaje → megger → energización → comisionado  |  "
        "Cables IN: tendido completo ANTES 17-MAR  |  Cables EL: tendido completo ANTES 14-MAY  |  "
        "SAT PMS: OCT-DIC-2027 (HOLD POINT contractual)",
        C["rfsu"], sz=8, h="left")
    ws.freeze_panes = "A6"


# ── HOJA 2: HITOS CLAVE ───────────────────────────────────────────────────────
def sheet_hitos(wb):
    ws = wb.create_sheet("2_Hitos_Clave")
    ws.sheet_view.showGridLines = False

    for col, w in {"A": 4, "B": 5, "C": 38, "D": 14, "E": 12, "F": 14, "G": 30, "H": 4}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("B1:G1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:G2"); ws.row_dimensions[2].height = 28
    hdr(ws, 2, 2,
        "HITOS CLAVE — INTEGRADO CPF2 La Calera II  |  RFSU: 20-DIC-2027  |  Corte: 12-JUN-2026",
        C["azul"], sz=13)
    ws.merge_cells("B3:G3"); ws.row_dimensions[3].height = 14
    hdr(ws, 3, 2,
        "Fuentes: CRONOGRAMA_GENERAL_ABB_Actualizado_2026-06-09  ·  "
        "Analisis_Electrico_CPF2_Campo_Rev2_2026-06-09  ·  Dashboard_Inauco_Seguimiento",
        C["azul_m"], sz=9)
    ws.row_dimensions[4].height = 8

    ws.row_dimensions[5].height = 20
    for i, h_ in enumerate(["#", "Hito", "Fecha", "Proveedor", "Restricción / Nota"], 2):
        hdr(ws, 5, i + 1, h_, C["azul"], sz=9)

    HITOS_TABLA = [
        # (nro, desc, fecha, prov, bg, fg, nota)
        (1,  "OC ABB Salas #3+#4 aceptada",             date(2026,5,12), "ABB",    C["s4_bg"],  C["s4_hdr"],
         "OC 4508944971 (S3) + 4508944973 (S4). Inicio oficial ABB"),
        (2,  "OC ABB PMS aceptada",                      date(2026,5,22), "ABB",    C["pms_bg"], C["pms_hdr"],
         "OC 4508945953. PM Pablo Kalis. KOM 22-MAY"),
        (3,  "OC Inauco PCS/SCADA/Comms",                date(2026,6,9),  "Inauco", C["pcs_bg"], C["pcs_hdr"],
         "OFT-128-008 Rev3. PENDIENTE firma OC formal"),
        (4,  "★ OC Cables IN — TARGET",                  date(2026,8,1),  "Cables", C["in_bg"],  C["in_hdr"],
         "⛔ CRÍTICO: LT 150d → entrega 28-NOV-2026 → tendido FEB-MAR → completo ANTES 17-MAR-2027"),
        (5,  "★ OC Cables EL — TARGET",                  date(2026,10,1), "Cables", C["el_bg"],  C["el_hdr"],
         "⛔ LT 150d → entrega 28-FEB-2027 → tendido MAR-MAY → completo ANTES llegada SE#3 14-MAY-2027"),
        (6,  "◆ Freezing Point Ing. PCS (Inauco)",       date(2026,9,30), "Inauco", C["pcs_bg"], C["pcs_hdr"],
         "Congela I/O list, P&ID, MCE, filosofía. 16 semanas antes iFAT. RIESGO R2"),
        (7,  "◆ Freezing Point Ing. Constructiva S#3 ← ÚLTIMO FP ABB", date(2026,8,31), "ABB", C["s3_bg"], C["s3_hdr"],
         "★ ÚLTIMO FP ABB. 1 semana late → Despacho S#3 late → PENALIDAD 1%/semana"),
        (8,  "FAT S#3+S#4 en fábrica Brasil",            date(2026,11,23),"ABB",    C["s3_bg"],  C["s3_hdr"],
         "Witness Point AESA. Ambas salas simultáneo. Finaliza 23-NOV-2026"),
        (9,  "FAT PMS tableros + sistema (Bs.As.)",       date(2027,1,29), "ABB",    C["pms_bg"], C["pms_hdr"],
         "25-29-ENE-2027. 16% OC certifica. Witness AESA"),
        (10, "FAT PCS + Comms Inauco (Neuquén)",          date(2027,2,19), "Inauco", C["pcs_bg"], C["pcs_hdr"],
         "27-ENE → 19-FEB-2027. 18 días. Stratix, VLANs OT, Modbus paquetizados"),
        (11, "FAT Shelter SE#4 (Mendoza)",                date(2027,3,16), "ABB",    C["s4_bg"],  C["s4_hdr"],
         "08-16-MAR-2027. 8 días. AESA + Cliente"),
        (12, "FAT Integral PMS Shelter",                  date(2027,3,31), "ABB",    C["pms_bg"], C["pms_hdr"],
         "25-31-MAR-2027. BUFFER AJUSTADO: 2 días antes FAT S#3"),
        (13, "FAT Shelter SE#3 ← FAT MÁS TARDÍO ABB",    date(2027,4,19), "ABB",    C["s3_bg"],  C["s3_hdr"],
         "02-19-ABR-2027. CADENA CRÍTICA. +35d holgura vs P6"),
        (14, "★ LLEGADA SE#4 a CPF2 (con PMS-101)",       date(2027,4,8),  "ABB",    C["s4_bg"],  C["s4_hdr"],
         "HITO ANCLA EL: habilita conexionado EL sector SE#4 (a partir 14-ABR)"),
        (15, "★ INICIO CAMPO PCS — Sala 7 (Inauco)",      date(2027,3,17), "Inauco", C["pcs_bg"], C["pcs_hdr"],
         "⛔ HITO ANCLA IN: Cables IN deben estar TENDIDOS. Conexionado IN comienza con esta instalación"),
        (16, "★ LLEGADA SE#3 a CPF2 (con PMS-001) ← ANCLA EL", date(2027,5,14), "ABB", C["s3_bg"], C["s3_hdr"],
         "⛔ HITO ANCLA EL: Cables EL deben estar TENDIDOS. CADENA CRÍTICA. Habilita montaje CCMs + conexionado EL"),
        (17, "iFAT Integración SE#3+SE#4+PMS+PCS+SIS",   date(2027,7,14), "TODOS",  C["azul_cl"],C["azul"],
         "25-JUN → 14-JUL-2027. Todos los vendors. ABB+Inauco+HIMA+AESA+Pluspetrol"),
        (18, "SAT PMS — HOLD POINT contractual (PIE)",    date(2027,12,5), "ABB",    C["pms_bg"], C["pms_hdr"],
         "06-OCT → 05-DIC-2027. Ventana FIJA OCT-DIC. No admite deslizamiento"),
        (19, "Fin MCE / Inauco field complete",           date(2027,9,21), "Inauco", C["pcs_bg"], C["pcs_hdr"],
         "CAO finalización proyecto Inauco 21-SEP-2027"),
        (20, "★★★ RFSU — Ready For Start Up",             date(2027,12,20),"TODOS",  C["rfsu_bg"],C["rfsu"],
         "20-DIC-2027. Todos los sistemas operativos. ABB garantía extendida 20-DIC-2028"),
    ]

    for r_off, (nro, desc, fecha, prov, bg, fg, nota) in enumerate(HITOS_TABLA):
        r = 6 + r_off
        ws.row_dimensions[r].height = 20
        dias = (fecha - TODAY).days
        estado_col = (C["ok_bg"] if dias < 0 else
                      (C["warn_bg"] if dias < 90 else
                       ("E8F0FF" if dias < 270 else C["gray_cl"])))
        cell(ws, r, 2, nro,   bold=True,  fg=fg, bg=bg, sz=9, h="center")
        cell(ws, r, 3, desc,  bold=True,  fg=fg, bg=bg, sz=9, h="left")
        cell(ws, r, 4, fmt(fecha), bold=True, fg=fg, bg=bg, sz=9, h="center")
        cell(ws, r, 5, prov,  bold=False, fg=fg, bg=bg, sz=9, h="center")
        cell(ws, r, 6, f"+{dias}d" if dias >= 0 else "PASADO",
             bold=True, fg=C["rfsu"] if dias > 500 else fg,
             bg=estado_col, sz=9, h="center")
        cell(ws, r, 7, nota,  bold=False, fg=C["gray"], bg=bg, sz=8.5,
             h="left", wrap=True, ita=True)

    # Separador restricciones
    r2 = 6 + len(HITOS_TABLA) + 2
    ws.merge_cells(f"B{r2}:G{r2}"); ws.row_dimensions[r2].height = 22
    hdr(ws, r2, 2,
        "RESTRICCIONES CRÍTICAS DE SECUENCIA — CABLES",
        C["rest_bg"], fg=C["rest_t"], sz=11)

    REST = [
        ("⛔ R1: Cables IN tendidos ANTES de 17-MAR-2027",
         "OC Cables IN ≤01-AGO-2026 → LT 150d → sitio 28-NOV-2026 → "
         "canaletas NOV-ENE → tendido 01-FEB→14-MAR → buffer 3 días → PCS llega 17-MAR",
         "Si OC Cables IN se demora, imposible tender antes de PCS → conexionado IN retrasa → "
         "SCADA SIS 100 días retrasa → RFSU impactado"),

        ("⛔ R2: Cables EL tendidos ANTES de 14-MAY-2027 (llegada SE#3)",
         "OC Cables EL ≤01-OCT-2026 → LT 150d → sitio 28-FEB-2027 → "
         "bandejas MAR → tendido 01-MAR→13-MAY → buffer 1 día → SE#3 llega 14-MAY",
         "Si cables EL llegan tarde, tendido se extiende más allá de llegada SE#3 → "
         "montaje CCMs en campo no puede completarse → energización retrasa → RFSU impactado"),

        ("⚠ R3: SE#3 cadena crítica — cualquier demora impacta RFSU",
         "FP-S3-3 (31-AGO-2026) → fabricación → FAT Brasil → FAT Shelter → Despacho 14-MAY → "
         "montaje campo → megger → energización → comisionado → RFSU 20-DIC-2027",
         "Holgura vs P6: Despacho S#3 +45d. Holgura vs RFSU 20-DIC: "
         "14-MAY + montaje 22d + megger 22d + energ 16d + precom 21d + com 107d = 20-DIC ✅"),

        ("⚠ R4: SAT PMS es HOLD POINT contractual OCT-DIC 2027",
         "Ventana fija OCT-DIC 2027 per PIE/contrato. ABB field service certificados. "
         "Debe estar completada ANTES del RFSU",
         "Si SAT PMS no completa en ventana → RFSU no puede certificarse. "
         "Energización debe completarse antes de OCT-2027"),
    ]

    for i, (titulo, cadena, impacto) in enumerate(REST):
        r3 = r2 + 1 + i
        ws.row_dimensions[r3].height = 36
        ws.merge_cells(f"B{r3}:C{r3}")
        x = ws.cell(r3, 2, titulo)
        x.fill = fl(C["rest_bg"]); x.font = fn(True, C["rest_t"], 9)
        x.alignment = al("left", "center", wrap=True); x.border = thin()
        ws.merge_cells(f"D{r3}:E{r3}")
        x2 = ws.cell(r3, 4, cadena)
        x2.fill = fl("FFFFF0"); x2.font = fn(False, C["gray"], 8, italic=True)
        x2.alignment = al("left", "center", wrap=True); x2.border = thin()
        ws.merge_cells(f"F{r3}:G{r3}")
        x3 = ws.cell(r3, 6, impacto)
        x3.fill = fl(C["risk_bg"]); x3.font = fn(False, C["risk"], 8)
        x3.alignment = al("left", "center", wrap=True); x3.border = thin()


# ── HOJA 3: DIAGRAMA DE FLUJO DE DEPENDENCIAS ─────────────────────────────────
def sheet_flujo(wb):
    ws = wb.create_sheet("3_Flujo_Dependencias")
    ws.sheet_view.showGridLines = False

    # Layout 60 cols x 90 rows
    for col in range(1, 65):
        ws.column_dimensions[get_column_letter(col)].width = 3.5
    for row in range(1, 95):
        ws.row_dimensions[row].height = 16

    def box(r, c, cr, cc, texto, bg, fg="FFFFFF", sz=8, bold=True, wrap=True):
        try: ws.merge_cells(start_row=r, start_column=c, end_row=r+cr-1, end_column=c+cc-1)
        except: pass
        cell_ = ws.cell(r, c, texto)
        cell_.fill = fl(bg)
        cell_.font = fn(bold, fg, sz)
        cell_.alignment = al("center", "center", wrap)
        cell_.border = medium()

    def arr(r, c, txt="↓", fg="595959"):
        x = ws.cell(r, c, txt)
        x.font = fn(True, fg, 11)
        x.alignment = al("center", "center")

    def sep_line(r, c1, c2, txt, bg, fg="FFFFFF"):
        try: ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
        except: pass
        x = ws.cell(r, c1, txt)
        x.fill = fl(bg); x.font = fn(True, fg, 10)
        x.alignment = al("center", "center"); x.border = medium()

    # Título
    try: ws.merge_cells("A1:BK1")
    except: pass
    hdr(ws, 1, 1,
        "DIAGRAMA DE FLUJO DE DEPENDENCIAS — CPF2 La Calera II  |  "
        "RFSU: 20-DIC-2027  |  Restricciones de secuencia cables IN/EL",
        C["azul"], sz=12)
    ws.row_dimensions[1].height = 28

    # Leyenda
    try: ws.merge_cells("A2:BK2")
    except: pass
    hdr(ws, 2, 1,
        "Lectura: de arriba hacia abajo = orden temporal  ·  "
        "⛔ = restricción de secuencia dura  ·  ◆ = Freezing Point  ·  "
        "★ = Hito ancla  ·  ▼ = dependencia  ·  RFSU pivota sobre cadena SE#3",
        C["azul_m"], sz=8)
    ws.row_dimensions[2].height = 16

    # ── COLUMNA 1: PROVISIÓN (izq) ─────────────────────────────────────────
    c1 = 2
    box(4,  c1, 2, 10, "OC ABB SE#3+SE#4\n12-MAY-2026\n✅ Aceptada",     C["s4_hdr"])
    arr(6,  c1+4, "▼", C["s4_hdr"])
    box(7,  c1, 2, 10, "OC ABB PMS\n22-MAY-2026\n✅ Aceptada",           C["pms_hdr"])
    arr(9,  c1+4, "▼", C["pms_hdr"])
    box(10, c1, 2, 10, "OC Inauco PCS\n09-JUN-2026\n⏳ Pendiente",       C["pcs_hdr"])
    arr(12, c1+4, "▼", C["in_hdr"])
    box(13, c1, 2, 10, "★ OC Cables IN\nTARGET: 01-AGO-2026\n⛔ CRÍTICO", "9E4BB5")
    arr(15, c1+4, "▼", C["el_hdr"])
    box(16, c1, 2, 10, "★ OC Cables EL\nTARGET: 01-OCT-2026\n⛔ CRITICO", C["el_hdr"])

    # Separador
    sep_line(19, c1, c1+9, "▼ FABRICACIÓN ▼", C["sep"])

    # Fabricación ABB
    box(20, c1, 2, 10, "Fabricación\nSE#3+SE#4 (Brasil)\nSEP→NOV 2026",   C["s3_hdr"])
    arr(22, c1+4, "▼", C["s3_hdr"])
    box(23, c1, 2, 10, "FAT Brasil\nS3+S4\n03-23-NOV-2026",               C["s3_hdr"])
    arr(25, c1+4, "▼", C["s3_hdr"])
    box(26, c1, 2, 10, "Acopio Argentina\nS3+S4\nDIC-2026→ENE-2027",      C["azul_m"])
    arr(28, c1+4, "▼", C["s3_hdr"])
    box(29, c1, 2, 10, "Montaje Shelter\nS#3 (Mendoza)\n01-FEB→26-MAR-27",C["s3_hdr"])
    arr(31, c1+4, "▼", C["s3_hdr"])
    box(32, c1, 2, 10, "FAT Shelter S#3\n02-19-ABR-2027\n+35d vs P6",     C["s3_hdr"])
    arr(34, c1+4, "▼", C["s3_hdr"])
    box(35, c1, 2, 10, "Despacho S#3\n10-14-MAY-2027\n+45d vs P6",        C["s3_hdr"])
    arr(37, c1+4, "▼", C["s3_hdr"])

    # HITO ANCLA SE#3
    box(38, c1, 2, 10,
        "★ LLEGADA SE#3\nCPF2 — 14-MAY-2027\n⛔ ANCLA CABLES EL",
        C["rfsu"])

    # ── COLUMNA 2: CABLES IN + PCS (centro-izq) ───────────────────────────
    c2 = 14
    box(4,  c2, 2, 10, "AT Cables IN\n12-JUN-2026\nEn proceso",           C["in_hdr"])
    arr(6,  c2+4, "▼", C["in_hdr"])
    box(7,  c2, 2, 10, "OC Cables IN\nTARGET: 01-AGO-2026\n⛔ LT=150d",  "7030A0")
    arr(9,  c2+4, "▼", C["in_hdr"])
    box(10, c2, 2, 10, "Fabricación+Entrega\nCables IN\n28-NOV-2026",     C["in_hdr"])
    arr(12, c2+4, "▼", C["in_hdr"])
    box(13, c2, 2, 10, "Canaletas IN\nNOV-2026→ENE-2027\n(65 días)",      C["in_hdr"])
    arr(15, c2+4, "▼", C["in_hdr"])

    # RESTRICCIÓN
    box(16, c2, 2, 10,
        "⛔ TENDIDO CABLES IN\n01-FEB→14-MAR-2027\n42d | 2 cuadrillas",
        "C75000", fg="FFFFFF")
    arr(18, c2+4, "▼", "C75000")

    box(19, c2, 2, 10,
        "★ HITO ANCLA IN:\nCables IN COMPLETOS\nantes 17-MAR-2027",
        "7030A0")
    arr(21, c2+4, "▼", C["pcs_hdr"])

    # PCS
    box(22, c2, 2, 10, "FAT PCS Neuquén\n27-ENE→19-FEB-2027\n(18 días)",  C["pcs_hdr"])
    arr(24, c2+4, "▼", C["pcs_hdr"])
    box(25, c2, 2, 10, "Integración PCS↔PMS\nBsAs 24-FEB→02-MAR\n⚠ 5 días", C["pcs_hdr"])
    arr(27, c2+4, "▼", C["pcs_hdr"])
    box(28, c2, 2, 10,
        "★ INICIO CAMPO\nPCS Sala 7\n17-MAR-2027",
        C["pcs_hdr"])
    arr(30, c2+4, "▼", C["pcs_hdr"])
    box(31, c2, 2, 10, "SAT PCS\nSala 7\n17-MAR→01-JUN-27",               C["pcs_hdr"])
    arr(33, c2+4, "▼", C["pcs_hdr"])
    box(34, c2, 2, 10, "⬅ Conexionado IN\ncomienza CON\nPCS instalado",   "7030A0")
    arr(36, c2+4, "▼", C["pcs_hdr"])
    box(37, c2, 2, 10, "SAT SCADA SIS\n17-MAR→03-AGO-27\n(HIMA+Inauco)", C["pcs_hdr"])
    arr(39, c2+4, "▼", C["pcs_hdr"])
    box(40, c2, 2, 10, "MCE + CAO\nAGO→SEP-2027\nInauco fin: 21-SEP",     C["pcs_hdr"])

    # ── COLUMNA 3: CABLES EL (centro-der) ─────────────────────────────────
    c3 = 26
    box(4,  c3, 2, 10, "AT Cables EL\n12-JUN-2026\nRI emitida",           C["el_hdr"])
    arr(6,  c3+4, "▼", C["el_hdr"])
    box(7,  c3, 2, 10, "OC Cables EL\nTARGET: 01-OCT-2026\n⛔ LT=150d",  C["el_hdr"])
    arr(9,  c3+4, "▼", C["el_hdr"])
    box(10, c3, 2, 10, "Fabricación+Entrega\nCables EL\n28-FEB-2027",     C["el_hdr"])
    arr(12, c3+4, "▼", C["el_hdr"])
    box(13, c3, 2, 10, "Bandejas principales\nMAR-2027\n(35 días)",       C["el_hdr"])
    arr(15, c3+4, "▼", C["el_hdr"])

    box(16, c3, 2, 10,
        "⛔ TENDIDO CABLES EL\n01-MAR→13-MAY-2027\n74d | 3 cuadrillas | 81,231m",
        "C75000", fg="FFFFFF")
    arr(18, c3+4, "▼", "C75000")

    box(19, c3, 2, 10,
        "★ HITO ANCLA EL:\nCables EL COMPLETOS\nantes 14-MAY-2027",
        C["el_hdr"])
    arr(21, c3+4, "▼", C["el_hdr"])

    box(22, c3, 2, 10, "Llegada SE#4\n08-ABR-2027\nHabilita conex. S4",   C["s4_hdr"])
    arr(24, c3+4, "▼", C["el_hdr"])
    box(25, c3, 2, 10, "Conexionado EL S4\n14-ABR→31-MAY\n⬅ Pivota SE#4",C["el_hdr"])
    arr(27, c3+4, "▼", C["el_hdr"])
    box(28, c3, 2, 10, "Llegada SE#3\n14-MAY-2027\n⛔ ANCLA",             C["s3_hdr"])
    arr(30, c3+4, "▼", C["el_hdr"])
    box(31, c3, 2, 10, "Conexionado EL S3\n18-MAY→06-JUL\n⬅ Pivota SE#3",C["el_hdr"])
    arr(33, c3+4, "▼", C["el_hdr"])
    box(34, c3, 2, 10, "Montaje CCMs campo\nS4: ABR | S3: MAY\nABB+AESA", C["s3_hdr"])

    # ── COLUMNA 4: CAMPO / COMISIONADO (der) ──────────────────────────────
    c4 = 38
    sep_line(4, c4, c4+9, "CAMPO — PRUEBAS Y COMISIONADO", C["cmp_hdr"])

    box(5,  c4, 2, 10, "Pruebas Megger\n(ITP HOLD POINT)\n07-28-JUL-2027", C["cmp_hdr"])
    arr(7,  c4+4, "▼", C["cmp_hdr"])
    box(8,  c4, 2, 10, "Energización MT→BT\nABB+AESA+Cliente\n29-JUL→13-AGO-27",C["cmp_hdr"])
    arr(10, c4+4, "▼", C["cmp_hdr"])
    box(11, c4, 2, 10, "Precomisionado\nEléctrico\n14-AGO→03-SEP-2027",   C["cmp_hdr"])
    arr(13, c4+4, "▼", C["cmp_hdr"])
    box(14, c4, 2, 10, "iFAT Integración\nTODOS vendors\n25-JUN→14-JUL-27",C["azul"])
    arr(16, c4+4, "▼", C["cmp_hdr"])
    box(17, c4, 2, 10, "Comisionado\nIntegrado EL+PCS+PMS\n04-SEP→19-DIC-27",C["cmp_hdr"])
    arr(19, c4+4, "▼", C["cmp_hdr"])
    box(20, c4, 2, 10, "SAT PMS\n(HOLD POINT PIE)\nOCT-DIC 2027",         C["pms_hdr"])
    arr(22, c4+4, "▼", C["rfsu"])

    # RFSU — span toda la anchura
    try: ws.merge_cells(start_row=23, start_column=2, end_row=25, end_column=47)
    except: pass
    rfsu_cell = ws.cell(23, 2,
        "★★★  RFSU — READY FOR START UP  ·  20-DIC-2027  ·  "
        "Todos los sistemas operativos: ABB (SE#3 · SE#4 · PMS) + Inauco (PCS/SCADA/SIS) + Cables EL/IN  ★★★")
    rfsu_cell.fill = fl(C["rfsu"])
    rfsu_cell.font = fn(True, "FFFFFF", 14)
    rfsu_cell.alignment = al("center", "center")
    rfsu_cell.border = medium()
    ws.row_dimensions[23].height = 30
    ws.row_dimensions[24].height = 30
    ws.row_dimensions[25].height = 30

    # Flechas entre columnas (texto indicativo)
    # Flecha IN → PCS
    arr(19, c2-1, "←", "7030A0")
    ws.cell(19, c2-1).value = "⛔ IN\ntend. ✓"
    ws.cell(19, c2-1).font = fn(True, "7030A0", 7)
    ws.cell(19, c2-1).alignment = al("center", "center", True)

    # Flecha EL → llegada salas
    arr(19, c3-1, "←", C["el_hdr"])
    ws.cell(19, c3-1).value = "⛔ EL\ntend. ✓"
    ws.cell(19, c3-1).font = fn(True, C["el_hdr"], 7)
    ws.cell(19, c3-1).alignment = al("center", "center", True)


# ── HOJA 4: VOLUMEN DE CABLES ─────────────────────────────────────────────────
def sheet_cables(wb):
    ws = wb.create_sheet("4_Volumen_Cables")
    ws.sheet_view.showGridLines = False

    for col, w in {"A":4,"B":34,"C":12,"D":12,"E":12,"F":14,"G":14,"H":14,"I":28,"J":4}.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("B1:I1"); ws.row_dimensions[1].height = 28
    hdr(ws, 1, 2,
        "VOLUMEN Y SECUENCIA DE CABLES — CPF2 La Calera II  |  "
        "Rev2 09-JUN-2026  |  Fuente: Analisis_Electrico_CPF2_Campo_Rev2_2026-06-09.xlsx",
        C["azul"], sz=12)

    ws.merge_cells("B2:I2"); ws.row_dimensions[2].height = 16
    hdr(ws, 2, 2,
        "513 cables CPF-2  ·  81,231 m totales  ·  "
        "~263 días calendario (3 cuadrillas)  ·  "
        "Ruta crítica: cables BT grande 3x240/3x150 (104 cables, 17,470m)",
        C["azul_m"], sz=9)

    ws.row_dimensions[3].height = 8
    ws.row_dimensions[4].height = 20
    for i, h_ in enumerate(["Categoría","Cant.\nCables","Metros\nCPF-2",
                              "Días\nTendido*","Días\nConexionado*",
                              "Inicio\nTendido","Fin\nTendido","Notas"], 2):
        hdr(ws, 4, i, h_, C["azul"], sz=9, wrap=True)

    DATA_CABLES = [
        # Instrumentación (IN)
        ("INSTRUMENTACIÓN — 192 cables | 28,126m | Tendido: ANTES 17-MAR-2027", None, None,
         None, None, None, None, "", C["in_hdr"], "FFFFFF"),
        ("Cables Control + Señales IN (7x2.5+T, mul.par)", 192, 28126,
         42, 106, date(2027,2,1), date(2027,3,14),
         "⛔ Tendido completo ANTES 17-MAR. Cone xionado inicia CON PCS 17-MAR",
         C["in_bg"], C["in_hdr"]),

        # Eléctrica (EL)
        ("ELÉCTRICA — 321 cables | 53,105m | Tendido: ANTES 14-MAY-2027", None, None,
         None, None, None, None, "", C["el_hdr"], "FFFFFF"),
        ("Potencia MT (13.2/6.6kV) — 18 cables", 18, 2110,
         29, 29, date(2027,5,10), date(2027,6,7),
         "Req. técnicos ABB certificados. Coordinación con 3 meses antelación",
         C["el_bg"], C["el_hdr"]),
        ("Potencia BT Grande ≥35mm² — 104 cables ← RUTA CRÍTICA", 104, 17470,
         49, 26, date(2027,3,1), date(2027,4,18),
         "★ RUTA CRÍTICA. 3x240, 3x150, 3x95, 3x50mm². Tendido pesado en canaletas",
         C["s3_bg"], C["s3_hdr"]),
        ("Potencia BT Mediano 10-35mm² — 71 cables", 71, 11250,
         43, 25, date(2027,3,15), date(2027,4,26),
         "3x16, 3x25mm². Motores BT auxiliares, calefacción, tomacorrientes",
         C["el_bg"], C["el_hdr"]),
        ("Potencia BT Pequeño <10mm² — 127 cables", 127, 21875,
         43, 25, date(2027,3,15), date(2027,4,26),
         "3x4, 3x6, 4x4mm². Iluminación, instrumentos, EAL",
         C["el_bg"], C["el_hdr"]),

        # FO/ETH
        ("FO / Ethernet — 1 cable", 1, 400,
         2, 5, date(2027,6,21), date(2027,7,5),
         "Red PMS SE#3-SE#4. ABB/AESA",
         C["pcs_bg"], C["pcs_hdr"]),

        # Totales
        ("TOTAL CPF-2", 513, 81231, None, None, date(2027,3,1), date(2027,5,13),
         "~263 días calendario (paralelizado 3 cuadrillas). Margen vs RFSU 20-DIC: 10 días",
         C["ok_bg"], C["ok"]),
    ]

    row = 5
    for rec in DATA_CABLES:
        if rec[1] is None:
            # Separador
            ws.merge_cells(f"B{row}:I{row}"); ws.row_dimensions[row].height = 18
            hdr(ws, row, 2, rec[0], rec[8], fg=rec[9], sz=9, h="left")
            row += 1
            continue
        (desc, cant, metros, d_tend, d_conx, ini_tend, fin_tend, nota, bg, fg) = rec
        ws.row_dimensions[row].height = 22
        cell(ws, row, 2, desc,     bold=(desc=="TOTAL CPF-2"), fg=fg, bg=bg, sz=9, h="left")
        cell(ws, row, 3, cant,     bold=False, fg=fg, bg=bg, sz=9, h="center", nf="#,##0")
        cell(ws, row, 4, metros,   bold=False, fg=fg, bg=bg, sz=9, h="center", nf="#,##0")
        cell(ws, row, 5, d_tend,   bold=False, fg=fg, bg=bg, sz=9, h="center")
        cell(ws, row, 6, d_conx,   bold=False, fg=fg, bg=bg, sz=9, h="center")
        cell(ws, row, 7, fmt(ini_tend) if ini_tend else "—",
             bold=False, fg=fg, bg=bg, sz=9)
        cell(ws, row, 8, fmt(fin_tend) if fin_tend else "—",
             bold=False, fg=fg, bg=bg, sz=9)
        cell(ws, row, 9, nota, bold=False, fg=C["gray"], bg=bg, sz=8,
             h="left", wrap=True, ita=True)
        row += 1

    # Notas metodología
    r2 = row + 2
    ws.merge_cells(f"B{r2}:I{r2}"); ws.row_dimensions[r2].height = 18
    hdr(ws, r2, 2, "NOTAS DE INSTALACIÓN", C["sep"], sz=9)

    notas = [
        "(*) Días paralelizados. Tendido: 3 cuadrillas x 4 personas. Conexionado: 3 cuadrillas. Base: 8hs/día efectivas, Patagonia.",
        "Ruta crítica EL: Cables BT Grande (104 cables, 17,470m) → 49 días tendido → mayor restricción para RFSU 20-DIC-2027.",
        "Cables IN: tendido debe completar ANTES del 17-MAR-2027 (inicio PCS Sala 7). Requiere OC cables IN ≤01-AGO-2026.",
        "SAT PMS: HOLD POINT contractual. Ventana OCT-DIC 2027 fija. Energización debe completar antes de OCT.",
        "Total puntas CPF-2 (Rev2): ~5,178 puntas (3,036 salas / 1,717 campo). Ver Analisis_Electrico_CPF2_Campo_Rev2.",
        "Cables MT 13.2kV: requieren técnicos ABB field service certificados. Coordinar con mínimo 3 meses de antelación.",
    ]
    for i, nota in enumerate(notas):
        r3 = r2 + 1 + i
        ws.row_dimensions[r3].height = 20
        ws.merge_cells(f"B{r3}:I{r3}")
        x = ws.cell(r3, 2, nota)
        x.fill = fl(C["gray_cl"]); x.font = fn(False, C["gray"], 8.5, italic=True)
        x.alignment = al("left", "center", wrap=True); x.border = thin()


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    wb = Workbook()
    wb.remove(wb.active)

    print("Generando Hoja 1: Gantt Integrado...")
    sheet_gantt(wb)

    print("Generando Hoja 2: Hitos Clave...")
    sheet_hitos(wb)

    print("Generando Hoja 3: Flujo de Dependencias...")
    sheet_flujo(wb)

    print("Generando Hoja 4: Volumen de Cables...")
    sheet_cables(wb)

    wb.save(OUT_XLSX)
    print(f"\n✅ Generado: {os.path.basename(OUT_XLSX)}")
    print(f"   Ruta: {OUT_XLSX}")
    print(f"\nHojas:")
    for i, ws in enumerate(wb.worksheets, 1):
        print(f"  {i}. {ws.title}")
    print(f"\nActividades en Gantt: {len(ACT)}")
    print(f"Semanas timeline: {len(SEMANAS)} ({SEMANAS[0]} → {SEMANAS[-1]})")
    print(f"\nRestricciones críticas:")
    print(f"  ⛔ Cables IN tendidos: 01-FEB → 14-MAR-2027 (antes PCS 17-MAR)")
    print(f"  ⛔ Cables EL tendidos: 01-MAR → 13-MAY-2027 (antes SE#3 14-MAY)")
    print(f"  ★  RFSU objetivo: 20-DIC-2027")


if __name__ == "__main__":
    main()
