"""
CRONOGRAMA GENERAL ABB — PRELIMINAR
Integración: Salas #3 (BT), #4 (MT) y PMS — CPF-2 / La Calera II — Vaca Muerta
Fecha: 27-MAY-2026
Fuentes: KOM 18/22-MAY | Cronograma ABB RevA (MS Project E-2611027) | ET ACAL-00102-ET-E-0005 | PIE ACAL-110-XX-X-XXX
"""
import os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))
FECHA_DOC    = "27-MAY-2026"
OUTPUT = os.path.join(PLANNING_DIR, "CRONOGRAMA_GENERAL_ABB_Preliminar_2026-05-27.xlsx")

C = {
    "red":    "CC0000", "dark":   "880000", "blue":  "2E75B6",
    "fat":    "E36C09", "desp":   "7030A0", "integ": "375623",
    "hito":   "C00000", "pms":    "B8860B", "aprov": "C55A11",
    "procur": "833C11", "soft":   "00819A", "maq":   "E36C09",
    "pay":    "1F4E79", "hdr":    "1F3864", "sep":   "404040",
    "s3l":    "FFE0E0", "s4l":    "DEEAF1", "pmsl":  "FFF8DC",
    "aprovl": "FFF2CC", "softl":  "E8F5E9", "intl":  "E2EFDA",
    "payl":   "BDD7EE", "ok":     "C6EFCE", "warn":  "FFEB9C",
    "risk":   "FFC7CE", "gray":   "F2F2F2", "white": "FFFFFF",
    "ok_t":   "375623", "warn_t": "9C5700", "risk_t":"9C0006",
}

def fill(c): return PatternFill("solid", fgColor=c)
def bd():
    t = Side(style="thin", color="CCCCCC")
    return Border(left=t, right=t, top=t, bottom=t)

def dat(ws, r, c, v, bg=None, bold=False, sz=9, h="center", col="000000", italic=False, wrap=False):
    x = ws.cell(row=r, column=c, value=v)
    x.font = Font(name="Calibri", bold=bold, size=sz, color=col, italic=italic)
    if bg: x.fill = fill(bg)
    x.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    return x

def hdr(ws, r, c, v, bg, fg="FFFFFF", bold=True, sz=9, h="center", wrap=False):
    x = ws.cell(row=r, column=c, value=v)
    x.fill = fill(bg); x.font = Font(name="Calibri", bold=bold, size=sz, color=fg)
    x.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    return x

def tborder(ws, r1, c1, r2, c2):
    t = Side(style="thin", color="BBBBBB")
    b = Border(left=t, right=t, top=t, bottom=t)
    for r in range(r1, r2+1):
        for c in range(c1, c2+1):
            ws.cell(r, c).border = b

MESES = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

def fmt(d): return d.strftime("%d/%m/%Y") if d else "—"

# ─────────────────────────────────────────────────────────────────────────────
# GANTT RESUMEN — actividades principales
# ─────────────────────────────────────────────────────────────────────────────
ACT = [
  # (grupo, id, nombre, ini, fin, tipo, critica, oc, holg_p6, nota)
  # tipo: actividad|fp|fat|despacho|integracion|pago|hito_contrato|aprobacion|procura|software|maqueta

  # ─ CONTRATO ─────────────────────────────────────────────────────────────────
  ("CONTRATO","HC-01","◆ Aceptación OC — Inicio oficial ABB",
   date(2026,5,12),date(2026,5,12),"hito_contrato",True,"S3+S4+PMS",None,
   "Christian Bombace. 3 OC aceptadas"),
  ("CONTRATO","HC-02","KOM Salas Eléctricas #3 y #4",
   date(2026,5,18),date(2026,5,18),"hito_contrato",True,"S3+S4",None,
   "PM Rodrigo Mack. Ref R-2631036"),
  ("CONTRATO","HC-03","KOM PMS (Power Management System)",
   date(2026,5,22),date(2026,5,22),"hito_contrato",True,"PMS",None,
   "PM Pablo Kalis. OC 4508945953. Ref E-2611027"),
  ("CONTRATO","HC-04","◆ Cronograma final ABB con Hold Points",
   date(2026,6,1),date(2026,6,1),"hito_contrato",True,"S3+S4+PMS",None,
   "COMPROMETIDO KOM. Con OC a 3ros y hitos ensayos"),
  ("CONTRATO","HC-05","Primer Informe de Avance ABB",
   date(2026,6,5),date(2026,6,5),"hito_contrato",True,"S3+S4+PMS",None,
   "Formato ABB. Reuniones semanales seguimiento + quincenales técnicas"),

  # ─ PMS ──────────────────────────────────────────────────────────────────────
  ("PMS","PP-PLN","Planificación Detallada PMS",
   date(2026,5,25),date(2026,6,5),"actividad",False,"PMS",None,
   "10 días. Organización interna ABB arranque ingeniería"),
  ("PMS","PP-IB","Emisión Ingeniería Básica (Arquitectura+I/O+Variables+PlConst.) RevA",
   date(2026,6,8),date(2026,7,17),"actividad",False,"PMS",None,
   "5 documentos RevA. Hito 10% envío IB = 17-JUL-2026"),
  ("PMS","PP-AI","▷ AESA — Aprobación IB (RevA → Rev0) ciclo completo",
   date(2026,6,22),date(2026,8,14),"aprobacion",False,"PMS",None,
   "10d hábiles por doc. Hito 10% aprobación IB = 14-AGO-2026"),
  ("PMS","FP-P1","◆ FREEZING POINT — Ingeniería Básica PMS",
   date(2026,8,21),date(2026,8,21),"fp",False,"PMS",None,
   "21-AGO: Planos Constructivos Rev0 aprobados. Congela: arq. 800XA + I/O 30% reserva + Modbus TCP/IP"),
  ("PMS","PP-ID","Emisión ID + Procedimientos FAT/SAT (RevA y Rev0)",
   date(2026,8,3),date(2026,10,9),"actividad",False,"PMS",None,
   "Variables+Esp.Funcional+Proc.FAT+Proc.SAT. Ver detalle hoja PMS_Detalle"),
  ("PMS","PP-AD","▷ AESA — Aprobación ID + Procedimientos (Rev0)",
   date(2026,8,17),date(2026,10,23),"aprobacion",False,"PMS",None,
   "Últ. aprobación: Proc.FAT/SAT Rev0 → 23-OCT"),
  ("PMS","FP-P2","◆ FREEZING POINT — Ing. Detalle + Procedimientos PMS",
   date(2026,10,23),date(2026,10,23),"fp",False,"PMS",None,
   "23-OCT: ÚLTIMO FP PMS. Proc.FAT+SAT Rev0 aprobados → habilita fabricación definitiva"),
  ("PMS","PP-PR","Procura: Switches Red + AC800M + Tableros",
   date(2026,5,22),date(2026,11,5),"procura",False,"PMS",None,
   "Switches 120d → 05-NOV (trigger hito acopio 25%)"),
  ("PMS","PP-SW","Software y Configuración 800XA (IEC61850 / PROFINET / MODBUS)",
   date(2026,8,3),date(2026,11,27),"software",False,"PMS",None,
   "Pantallas+I/O+Comunicaciones+Lógicas deslastre+DAG"),
  ("PMS","PP-MQ","FAT de Maqueta — Pruebas Tempranas (AESA OBLIGATORIO)",
   date(2026,9,22),date(2026,9,25),"maqueta",False,"PMS",None,
   "4 días fábrica BsAs. Mock-up PMS+CCM+AC800M+PROFINET+MODBUS TCP. Witness Point (W) PIE"),
  ("PMS","PP-FAB","Armado Tableros PMS-001 (SE#3) y PMS-101 (SE#4)",
   date(2026,9,11),date(2026,12,17),"actividad",False,"PMS",None,
   "70 días. Fabricación Bs.As. IP54 Rittal. AC800M redundante"),
  ("PMS","PP-TI","Pruebas Internas ABB (FAT Interna = IV per PIE)",
   date(2026,12,18),date(2027,1,7),"actividad",False,"PMS",None,
   "15 días. Sin cliente. Resultado: protocolo con observaciones"),
  ("PMS","PP-FAT","★ FAT Tableros + Sistema PMS (WITNESS — AESA asiste)",
   date(2027,1,25),date(2027,1,29),"fat",False,"PMS",222,
   "5 días BsAs. W per PIE. Hito 20% FAT → 80% del ítem = 16% OC. Holgura +222d vs P6"),
  ("PMS","PP-DES","▶ DESPACHO PMS — Bs.As. → Shelter (Mendoza)",
   date(2027,2,4),date(2027,2,17),"despacho",False,"PMS",None,
   "10 días. Hito 30% Entrega Shelter = 17-FEB-2027"),
  ("PMS","PP-SH1","Montaje PMS en Shelter Mendoza",
   date(2027,2,18),date(2027,3,17),"actividad",False,"PMS",None,
   "20 días. PMS-001 en SE#3 + PMS-101 en SE#4. Simultáneo con S3-montaje y S4-FAT ⚠"),
  ("PMS","PP-SH3","★ FAT INTEGRAL SHELTER PMS",
   date(2027,3,25),date(2027,3,31),"fat",False,"PMS",None,
   "5 días. Hito 20% FAT → 20% restante = 4% OC. Termina 2 días antes FAT S#3 ✅"),
  ("PMS","PP-DB","★ DataBook PMS aprobado — 31-MAR-2027",
   date(2027,3,31),date(2027,3,31),"hito_contrato",False,"PMS",None,
   "Hito 5% OC PMS. Planos conforme+protocolos+certificados+manuales"),

  # ─ SALA #4 (MT) ──────────────────────────────────────────────────────────────
  ("Sala #4 (MT)","S4-01","Emisión / Aprobación IB Sala #4",
   date(2026,5,22),date(2026,6,26),"actividad",False,"S4",None,
   "Layout + single line 13.2kV/6.6kV. 10d hábiles aprobación AESA"),
  ("Sala #4 (MT)","FP-S4-1","◆ FREEZING POINT — Ing. Básica Sala #4",
   date(2026,6,26),date(2026,6,26),"fp",False,"S4",None,
   "Inicio OC celdas MT Turquía (lead 6m). CRÍTICO logístico"),
  ("Sala #4 (MT)","S4-02","OC a 3ros: Celdas MT Turquía + Variadores MT China",
   date(2026,6,29),date(2026,9,15),"procura",False,"S4",None,
   "Celdas: lead 6m → Dic 2026. VFDs: lead 4-5m → Feb 2027"),
  ("Sala #4 (MT)","S4-03","Emisión / Aprobación ID + IC Sala #4",
   date(2026,6,29),date(2026,8,28),"actividad",False,"S4",None,
   "Esquemas detalle + planos constructivos. 10d hábiles c/aprobación AESA"),
  ("Sala #4 (MT)","FP-S4-2","◆ FREEZING POINT — Ing. Detalle Sala #4",
   date(2026,8,14),date(2026,8,14),"fp",False,"S4",None,
   "Inicio fabricación tableros auxiliares Bs.As."),
  ("Sala #4 (MT)","FP-S4-3","◆ FREEZING POINT — Ing. Constructiva S#4 ← CIERRA INGENIERÍA S4",
   date(2026,8,28),date(2026,8,28),"fp",False,"S4",None,
   "ÚLTIMO FP S4. Sin hito equivalente en P6 ⚠"),
  ("Sala #4 (MT)","S4-04","Fabricación Tableros S#4 (Brasil M&B)",
   date(2026,9,1),date(2026,11,2),"actividad",False,"S4",None,
   "25 columnas ZS1 630A 25kA"),
  ("Sala #4 (MT)","S4-05","FAT Celdas y Tableros MT S#4 (Brasil)",
   date(2026,11,3),date(2026,11,23),"fat",False,"S4",None,
   "20 días. AESA inspecciona en Brasil. Witness Point"),
  ("Sala #4 (MT)","S4-06","Acopio materiales Argentina S#4",
   date(2026,12,7),date(2027,1,15),"actividad",False,"S4",None,
   "Hito 25% acopio S4"),
  ("Sala #4 (MT)","S4-07","Montaje e integración Shelter S#4 (Bottino — Mendoza)",
   date(2027,1,18),date(2027,3,5),"actividad",False,"S4",None,
   "Incluye PMS-101 (llega 17-FEB) ← verificar espacio Shelter"),
  ("Sala #4 (MT)","S4-08","FAT Sala Eléctrica #4 (Shelter completo)",
   date(2027,3,8),date(2027,3,16),"fat",False,"S4",None,
   "8 días. Asiste AESA + Cliente. Termina antes FAT Integral PMS (25-MAR) ✅"),
  ("Sala #4 (MT)","S4-09","▶ DESPACHO Sala #4 → Vaca Muerta",
   date(2027,4,2),date(2027,4,8),"despacho",False,"S4",81,
   "OC deadline: 08-ABR-2027. Holgura +81d vs P6 montaje campo"),
  ("Sala #4 (MT)","S4-10","★ ENTREGA S#4 — OC 4508944973",
   date(2027,4,8),date(2027,4,8),"hito_contrato",False,"S4",81,
   "CONTRACTUAL: 08-ABR-2027 (11 meses)"),

  # ─ SALA #3 (BT) ★ CADENA CRÍTICA ────────────────────────────────────────────
  ("Sala #3 (BT) ★ CRÍTICA","S3-W","⚠ OPEN ISSUE: CCM7 (3x22kW → 2x45kW+VFD)",
   date(2026,5,26),date(2026,6,5),"actividad",True,"S3",None,
   "Confirmar antes 05-JUN. VFDs DP-VFD-24210 A/B/C en scope PMS vía PROFINET. AESA"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-01","Emisión / Aprobación IB Sala #3",
   date(2026,5,22),date(2026,6,26),"actividad",True,"S3",None,
   "37 CCMs 4000A 65kA + 6 ductos barras. Trigger hito 10% IB"),
  ("Sala #3 (BT) ★ CRÍTICA","FP-S3-1","◆ FREEZING POINT — Ing. Básica Sala #3",
   date(2026,7,3),date(2026,7,3),"fp",True,"S3",None,
   "Inicio OC ductos barras Turquía (lead 6m). CRÍTICO ← cadena crítica"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-02","OC a 3ros: Ductos Barras Turquía + VFDs BT Finlandia",
   date(2026,7,6),date(2026,7,17),"procura",True,"S3",None,
   "Ductos: lead 6m. VFDs: lead 4-5m. Entrega Dic 2026 - Ene 2027"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-03","Emisión / Aprobación ID + IC Sala #3",
   date(2026,6,29),date(2026,9,12),"actividad",True,"S3",None,
   "Esquemas 37 CCMs + planos constructivos. Pendiente: CCM7 ⚠"),
  ("Sala #3 (BT) ★ CRÍTICA","FP-S3-2","◆ FREEZING POINT — Ing. Detalle Sala #3",
   date(2026,8,14),date(2026,8,14),"fp",True,"S3",None,
   "Inicio fabricación 37 CCMs Brasil"),
  ("Sala #3 (BT) ★ CRÍTICA","FP-S3-3","◆ FREEZING POINT — Ing. Constructiva S#3 ← ÚLTIMO FP ABB",
   date(2026,8,31),date(2026,8,31),"fp",True,"S3",None,
   "★ 31-AGO-2026: ÚLTIMO FP DE TODO EL PROYECTO ABB"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-04","Fabricación 37 Tableros BT S#3 (Brasil M&B)",
   date(2026,9,14),date(2026,11,2),"actividad",True,"S3",None,
   "37 CCMs 4000A 65kA"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-05","FAT Celdas y Tableros BT S#3 (Brasil)",
   date(2026,11,3),date(2026,11,23),"fat",True,"S3",None,
   "20 días. AESA inspecciona. 10 días antelación"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-06","Acopio materiales Argentina S#3",
   date(2026,12,7),date(2027,1,29),"actividad",True,"S3",None,
   "Hito 25% acopio S3. Ductos Turquía + VFDs BT + tableros BT"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-07","Montaje e integración Shelter S#3 (Bottino — Mendoza)",
   date(2027,2,1),date(2027,3,26),"actividad",True,"S3",None,
   "37 CCMs + ductos + VFDs. Simultáneo con PMS montaje (18-FEB/17-MAR) ⚠"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-08","FAT Sala Eléctrica #3 ← FAT MÁS TARDÍO ABB",
   date(2027,4,2),date(2027,4,19),"fat",True,"S3",35,
   "17 días. PMS-001 integrado post FAT Integral PMS (31-MAR). Holgura +35d vs P6"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-09","▶ DESPACHO Sala #3 ← DESPACHO MÁS TARDÍO ABB",
   date(2027,5,10),date(2027,5,14),"despacho",True,"S3",45,
   "OC deadline: 14-MAY-2027. Holgura +45d vs P6 montaje campo"),
  ("Sala #3 (BT) ★ CRÍTICA","S3-10","★ ENTREGA S#3 — OC 4508944971",
   date(2027,5,14),date(2027,5,14),"hito_contrato",True,"S3",45,
   "CONTRACTUAL: 14-MAY-2027 (12 meses). CADENA CRÍTICA"),

  # ─ HITOS DE PAGO ─────────────────────────────────────────────────────────────
  ("HITOS DE PAGO","PAY-P01","§ PMS 10% — Envío IB (17-JUL-2026)",
   date(2026,7,17),date(2026,7,17),"pago",False,"PMS",None,
   "Planos: Arquitectura + Listado I/O + Planos Constructivos. Póliza Caución"),
  ("HITOS DE PAGO","PAY-S01","§ S3+S4 10% — Envío IB (12-JUN-2026)",
   date(2026,6,12),date(2026,6,12),"pago",False,"S3+S4",None,
   "Emisión IB S3 y S4. Póliza Caución Avance Fabricación"),
  ("HITOS DE PAGO","PAY-P02","§ PMS 10% — Aprobación IB (14-AGO-2026)",
   date(2026,8,14),date(2026,8,14),"pago",False,"PMS",None,
   "Aprobación I/O + Variables RevA aprobados por AESA"),
  ("HITOS DE PAGO","PAY-S02","§ S3+S4 10% — Aprobación IB (03-JUL-2026)",
   date(2026,7,3),date(2026,7,3),"pago",False,"S3+S4",None,
   "Aprobación final IB S3 y S4"),
  ("HITOS DE PAGO","PAY-P03","§ PMS 25% — Acopio (05-NOV-2026)",
   date(2026,11,5),date(2026,11,5),"pago",False,"PMS",None,
   "HW ABB: PMS + Tableros + Switches en Argentina. Póliza Caución Acopio"),
  ("HITOS DE PAGO","PAY-S03","§ S3+S4 25% — Acopio (29-ENE-2027)",
   date(2027,1,29),date(2027,1,29),"pago",False,"S3+S4",None,
   "Celdas MT + tableros BT + ductos + VFDs + UPS en Argentina"),
  ("HITOS DE PAGO","PAY-P04","§ PMS 20% — FAT Fábrica 80% (29-ENE-2027) = 16% OC",
   date(2027,1,29),date(2027,1,29),"pago",False,"PMS",None,
   "FAT Tableros+Sistema. Propuesta: 80% hito = 16% OC. Póliza Caución"),
  ("HITOS DE PAGO","PAY-P05","§ PMS 30% — Entrega Shelter (17-FEB-2027)",
   date(2027,2,17),date(2027,2,17),"pago",False,"PMS",None,
   "Equipamiento PMS en Shelter Mendoza. Liberación Calidad Comprador"),
  ("HITOS DE PAGO","PAY-P06","§ PMS 20% — FAT Integral Shelter 20% (31-MAR-2027) = 4% OC",
   date(2027,3,31),date(2027,3,31),"pago",False,"PMS",None,
   "FAT Integral Shelter. 20% restante del hito = 4% OC"),
  ("HITOS DE PAGO","PAY-S04","§ S3 20% — FAT (19-ABR-2027)",
   date(2027,4,19),date(2027,4,19),"pago",False,"S3",None,
   "FAT más tardío (S#3). Póliza Caución Avance Fabricación"),
  ("HITOS DE PAGO","PAY-S05","§ S3+S4 30% — Contra Entrega (14-MAY-2027)",
   date(2027,5,14),date(2027,5,14),"pago",False,"S3+S4",None,
   "Contra entrega S#3 (última). Liberación Calidad Comprador"),
  ("HITOS DE PAGO","PAY-P07","§ PMS 5% — DataBook (31-MAR-2027)",
   date(2027,3,31),date(2027,3,31),"pago",False,"PMS",None,
   "DataBook PMS aprobado por AESA"),
  ("HITOS DE PAGO","PAY-S06","§ S3+S4 5% — DataBook (01-JUL-2027 est.)",
   date(2027,7,1),date(2027,7,1),"pago",False,"S3+S4",None,
   "~6-8 semanas post entrega S#3"),

  # ─ INTEGRACIÓN EPC ──────────────────────────────────────────────────────────
  ("INTEGRACIÓN EPC","INT-01","iFAT — Integración Shelters + PMS + INAUCO + HIMA",
   date(2027,6,25),date(2027,7,14),"fat",True,"S3+S4+PMS",3,
   "P6 TA1124=28-JUN. ABB 25-JUN. Holgura: 3 días. Todos vendors integran"),
  ("INTEGRACIÓN EPC","INT-02","Inicio Precomisionado / Montaje campo",
   date(2027,5,3),date(2028,1,13),"integracion",False,"EPC",None,
   "CDW1000. Fecha idéntica P6 y ABB ✅"),
  ("INTEGRACIÓN EPC","INT-03","Inicio Comisionado",
   date(2027,7,16),date(2028,1,31),"integracion",True,"EPC",None,
   "CDW1010. Fecha idéntica P6 y ABB ✅"),
  ("INTEGRACIÓN EPC","INT-04","SAT PMS en Sitio — HOLD POINT (HP) per PIE",
   date(2027,10,1),date(2027,12,15),"actividad",False,"PMS",None,
   "HP per PIE ACAL-110: NO puede continuar sin AESA + Pluspetrol. 100% E/S ensayados"),
  ("INTEGRACIÓN EPC","INT-05","★ RFSU — Ready For Start Up",
   date(2028,1,31),date(2028,1,31),"hito_contrato",True,"PROYECTO",None,
   "A10290. Fecha idéntica P6 y ABB ✅. Garantía extendida: 31-ENE-2029"),
  ("INTEGRACIÓN EPC","INT-06","Capacitación PMS — 1ra Instancia (Ingeniería + Mantenimiento)",
   date(2028,2,5),date(2028,2,9),"actividad",False,"PMS",None,
   "En sitio. Castellano. Per ET sec.23"),
  ("INTEGRACIÓN EPC","INT-07","Capacitación PMS — 2da Instancia (Operación)",
   date(2028,3,3),date(2028,3,7),"actividad",False,"PMS",None,
   "2da fecha en sitio per ET sec.23. Personal apto sin asistencia ABB"),
]

# ─────────────────────────────────────────────────────────────────────────────
# ACTIVIDADES PMS DETALLADAS (para hoja PMS_Detalle)
# ─────────────────────────────────────────────────────────────────────────────
ACT_PMS = [
  # (id, nombre, ini, fin, tipo, nota)
  ("PP-00","Planificación Detallada",date(2026,5,25),date(2026,6,5),"actividad","10 días"),
  ("PP-E1","Arquitectura/Topología de Red — RevA",date(2026,6,8),date(2026,6,19),"actividad","10d. 800XA. Modbus TCP/IP"),
  ("PP-E2","Planos Constructivos Tableros — RevA",date(2026,6,8),date(2026,7,10),"actividad","25d. PMS-001+PMS-101"),
  ("PP-E3","Listado de I/O — RevA",date(2026,6,22),date(2026,7,17),"actividad","20d. 30% reserva señales"),
  ("PP-E4","Variables por Comunicación — RevA",date(2026,6,22),date(2026,7,31),"actividad","30d. IEC61850+PROFINET+MODBUS"),
  ("PP-E5","Especificación Funcional — RevA",date(2026,6,22),date(2026,7,31),"actividad","30d. Deslastre+DAG"),
  ("PP-A1","▷ AESA aprueba Arquitectura RevA",date(2026,6,22),date(2026,7,3),"aprobacion","10d hábiles"),
  ("PP-A2","▷ AESA aprueba Planos Constructivos RevA",date(2026,7,13),date(2026,7,24),"aprobacion","10d hábiles"),
  ("PP-A3","▷ AESA aprueba I/O RevA",date(2026,7,20),date(2026,7,31),"aprobacion","10d hábiles → hito IB envío"),
  ("PP-A4","▷ AESA aprueba Variables RevA",date(2026,8,3),date(2026,8,14),"aprobacion","10d hábiles"),
  ("PP-A5","▷ AESA aprueba Esp.Funcional RevA",date(2026,8,3),date(2026,8,14),"aprobacion","10d hábiles → hito IB aprob."),
  ("PP-R1","Arquitectura Rev0",date(2026,7,6),date(2026,7,17),"actividad","10d. Incorpora comentarios"),
  ("PP-R2","Planos Constructivos Rev0",date(2026,7,27),date(2026,8,7),"actividad","10d"),
  ("PP-R3","I/O Rev0",date(2026,8,3),date(2026,8,14),"actividad","10d"),
  ("PP-R4","Variables Rev0",date(2026,8,17),date(2026,8,28),"actividad","10d"),
  ("PP-R5","Especificación Funcional Rev0",date(2026,8,17),date(2026,8,28),"actividad","10d"),
  ("PP-B1","▷ AESA aprueba Arquitectura Rev0",date(2026,7,20),date(2026,7,31),"aprobacion","10d"),
  ("PP-B2","▷ AESA aprueba Planos Constructivos Rev0",date(2026,8,10),date(2026,8,21),"aprobacion","10d → FP-P1"),
  ("FP-P1","◆ FREEZING POINT Ing. Básica PMS",date(2026,8,21),date(2026,8,21),"fp","21-AGO: arq+I/O+Modbus TCP/IP"),
  ("PP-B3","▷ AESA aprueba I/O Rev0",date(2026,8,17),date(2026,8,28),"aprobacion","10d"),
  ("PP-B4","▷ AESA aprueba Variables Rev0",date(2026,8,31),date(2026,9,11),"aprobacion","10d"),
  ("PP-B5","▷ AESA aprueba Esp.Funcional Rev0",date(2026,8,31),date(2026,9,11),"aprobacion","10d"),
  ("PP-SW1","Entorno Virtual / Instalación 800XA",date(2026,8,3),date(2026,8,14),"software","10d. Verificar licencias CPF1"),
  ("PP-SW2","Config Pantallas (1ra ronda)",date(2026,8,17),date(2026,9,4),"software","15d. Mímicos unifilares IEC"),
  ("PP-SW3","Config Pantallas (2da ronda)",date(2026,9,8),date(2026,9,28),"software","15d"),
  ("PP-SW4","Config I/O",date(2026,8,31),date(2026,10,16),"software","35d. SOE 1ms"),
  ("PP-SW5","Config Comunicaciones (IEC61850/PROFINET/MODBUS)",date(2026,10,19),date(2026,11,27),"software","30d"),
  ("PP-SW6","Config Lógicas Sistema (deslastre+DAG)",date(2026,10,19),date(2026,11,27),"software","30d"),
  ("PP-MQ","FAT Maqueta — Pruebas Tempranas AESA OBLIGATORIO",date(2026,9,22),date(2026,9,25),"maqueta","4d. BsAs. Witness(W)"),
  ("PP-E6","Procedimiento FAT — RevA",date(2026,8,31),date(2026,9,11),"actividad","10d"),
  ("PP-E7","Procedimiento SAT — RevA",date(2026,8,31),date(2026,9,11),"actividad","10d. SAT=HP per PIE"),
  ("PP-A6","▷ AESA aprueba Proc.FAT RevA",date(2026,9,14),date(2026,9,25),"aprobacion","10d"),
  ("PP-A7","▷ AESA aprueba Proc.SAT RevA",date(2026,9,14),date(2026,9,25),"aprobacion","10d"),
  ("PP-R6","Procedimiento FAT Rev0",date(2026,9,28),date(2026,10,9),"actividad","10d"),
  ("PP-R7","Procedimiento SAT Rev0",date(2026,9,28),date(2026,10,9),"actividad","10d"),
  ("PP-B6","▷ AESA aprueba Proc.FAT Rev0",date(2026,10,12),date(2026,10,23),"aprobacion","10d → habilita FAT 29-ENE"),
  ("PP-B7","▷ AESA aprueba Proc.SAT Rev0",date(2026,10,12),date(2026,10,23),"aprobacion","10d → HP SAT habilitado"),
  ("FP-P2","◆ FREEZING POINT Ing. Detalle + Procedimientos PMS",date(2026,10,23),date(2026,10,23),"fp","23-OCT: ÚLTIMO FP PMS"),
  ("PP-PR1","Procura Switches de Red",date(2026,5,22),date(2026,11,5),"procura","120d → trigger acopio 25%"),
  ("PP-PR2","Procura Material S800/AC800M",date(2026,5,22),date(2026,10,8),"procura","100d"),
  ("PP-PR3","Procura Tableros (carcasas)",date(2026,5,22),date(2026,10,8),"procura","100d. Rittal IP54"),
  ("PP-FAB","Armado Tableros PMS-001+PMS-101",date(2026,9,11),date(2026,12,17),"actividad","70d. BsAs"),
  ("PP-TI","Pruebas Internas ABB (IV per PIE)",date(2026,12,18),date(2027,1,7),"actividad","15d. Sin cliente"),
  ("PP-FAT","FAT Tableros + Sistema (WITNESS — W)",date(2027,1,25),date(2027,1,29),"fat","5d. 16% OC certifica"),
  ("PP-ADJ","Ajustes post-FAT",date(2027,2,1),date(2027,2,3),"actividad","3d"),
  ("PP-DES","▶ DESPACHO PMS → Shelter Mendoza",date(2027,2,4),date(2027,2,17),"despacho","Hito 30% = 17-FEB"),
  ("PP-SH1","Montaje PMS en Shelter",date(2027,2,18),date(2027,3,17),"actividad","20d. PMS-001+PMS-101"),
  ("PP-SH2","Supervisión + Preparativos FAT Integral",date(2027,3,18),date(2027,3,24),"actividad","5d"),
  ("PP-SH3","★ FAT INTEGRAL SHELTER PMS",date(2027,3,25),date(2027,3,31),"fat","5d. 4% OC certifica"),
  ("PP-DB","★ DataBook PMS aprobado",date(2027,3,31),date(2027,3,31),"hito_contrato","Hito 5% OC PMS"),
]

# ─────────────────────────────────────────────────────────────────────────────
# TIMELINE
# ─────────────────────────────────────────────────────────────────────────────
INICIO = date(2026, 5, 1)
FIN    = date(2028, 3, 31)
SEMANAS = []
d = INICIO
while d <= FIN:
    SEMANAS.append(d)
    d += timedelta(weeks=1)

N_INFO  = 6
COL_OFF = N_INFO + 1

def sem_rango(ini, fin):
    ci = cf = None
    for i, s in enumerate(SEMANAS):
        sw = s + timedelta(days=6)
        if ci is None and sw >= ini: ci = i
        if s <= fin: cf = i
    return ci, cf

def color_barra(tipo, critica):
    m = {"fp":"C00000","fat":"E36C09","despacho":"7030A0",
         "integracion":"375623","pago":"1F4E79","hito_contrato":"CC0000",
         "aprobacion":"C55A11","procura":"833C11","software":"00819A",
         "maqueta":"FF6600"}
    if tipo in m: return m[tipo]
    return "CC0000" if critica else "2E75B6"

# ─────────────────────────────────────────────────────────────────────────────
# HOJA 1: GANTT GENERAL
# ─────────────────────────────────────────────────────────────────────────────
GRUPO_BG = {
    "CONTRATO":                "CC0000",
    "PMS":                     "B8860B",
    "Sala #4 (MT)":            "2E75B6",
    "Sala #3 (BT) ★ CRÍTICA": "880000",
    "HITOS DE PAGO":           "1F4E79",
    "INTEGRACIÓN EPC":         "375623",
}
FILA_BG = {
    "CONTRATO":                "FFF0F0",
    "PMS":                     "FFF8DC",
    "Sala #4 (MT)":            "DEEAF1",
    "Sala #3 (BT) ★ CRÍTICA": "FFE0E0",
    "HITOS DE PAGO":           "BDD7EE",
    "INTEGRACIÓN EPC":         "E2EFDA",
}

def sheet_gantt(wb):
    ws = wb.create_sheet("1_Gantt_General")
    ws.sheet_view.showGridLines = False

    for i, w in enumerate([5, 40, 10, 11, 11, 10], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for i in range(len(SEMANAS)):
        ws.column_dimensions[get_column_letter(COL_OFF+i)].width = 2.0

    last_col = get_column_letter(COL_OFF + len(SEMANAS) - 1)

    # Título
    ws.merge_cells(f"A1:{last_col}1"); ws.row_dimensions[1].height = 32
    hdr(ws,1,1,
        f"CRONOGRAMA GENERAL ABB — PRELIMINAR  |  CPF-2 La Calera II (Vaca Muerta)  |  "
        f"Fecha: {FECHA_DOC}  |  OC: 4508944971 (S3) | 4508944973 (S4) | 4508945953 (PMS)  |  RFSU: 31/01/2028",
        C["red"], sz=11)

    # Leyenda
    ws.merge_cells(f"A2:{last_col}2"); ws.row_dimensions[2].height = 13
    ley = ("■ Cadena Crítica S#3   ■ Actividades Paralelas   ■ FAT/Ensayos   "
           "■ Despacho   ■ Integración   ■ ◆ Freezing Point   ■ § Pago   "
           "■ Aprobación AESA   ■ Procura   ■ Software   ■ FAT Maqueta")
    c = ws.cell(row=2, column=1, value=ley)
    c.font = Font(name="Calibri", size=7, color="444444")
    c.alignment = Alignment(horizontal="left", vertical="center")

    # Años
    ws.row_dimensions[3].height = 14
    anio_g = {}
    for i, s in enumerate(SEMANAS):
        a = s.year
        if a not in anio_g: anio_g[a] = [i, i]
        else: anio_g[a][1] = i
    for a, (ci, cf) in anio_g.items():
        try: ws.merge_cells(f"{get_column_letter(COL_OFF+ci)}3:{get_column_letter(COL_OFF+cf)}3")
        except: pass
        hdr(ws,3,COL_OFF+ci,str(a),C["hdr"],sz=9)

    # Meses
    ws.row_dimensions[4].height = 12
    mes_g = {}
    for i, s in enumerate(SEMANAS):
        k = (s.year, s.month)
        if k not in mes_g: mes_g[k] = [i, i]
        else: mes_g[k][1] = i
    for (a, m), (ci, cf) in mes_g.items():
        try: ws.merge_cells(f"{get_column_letter(COL_OFF+ci)}4:{get_column_letter(COL_OFF+cf)}4")
        except: pass
        hdr(ws,4,COL_OFF+ci,MESES[m],"2E75B6" if m%2==0 else "404040",sz=7)

    # Encabezados
    ws.row_dimensions[5].height = 20
    for c2, t in enumerate(["ID","Actividad / Hito","OC","Inicio","Fin","Holgura P6"],1):
        hdr(ws,5,c2,t,C["hdr"],sz=9)

    grupo_ant = None
    row = 6
    for (grupo,id_,nombre,ini,fin,tipo,critica,oc,holg,nota) in ACT:
        if grupo != grupo_ant:
            ws.merge_cells(f"A{row}:{last_col}{row}")
            ws.row_dimensions[row].height = 15
            bg_sep = GRUPO_BG.get(grupo, C["sep"])
            lbl = f"  ▶  {grupo.upper()}"
            if "CRÍTICA" in grupo:
                lbl += "   ← CADENA CRÍTICA  |  OC 4508944971  |  12 meses  |  Entrega: 14-MAY-2027"
            hdr(ws,row,1,lbl,bg_sep,sz=9,h="left")
            row += 1
            grupo_ant = grupo

        bg = FILA_BG.get(grupo, "FFFFFF")
        es_fp   = tipo == "fp"
        es_hito = tipo == "hito_contrato"
        es_warn = "WARN" in id_ or id_ == "S3-W"
        id_col  = C["dark"] if critica else C["hdr"]

        dat(ws,row,1,id_,bg=bg,bold=True,sz=8,col=id_col)
        dat(ws,row,2,nombre,bg=bg,bold=(critica or es_fp or es_hito),sz=9,h="left",
            col=C["dark"] if critica else ("000000"))
        dat(ws,row,3,oc,bg=bg,sz=8)
        dat(ws,row,4,fmt(ini),bg=bg,sz=8,bold=(es_fp or critica),
            col=C["dark"] if critica else "000000")
        dat(ws,row,5,fmt(fin),bg=bg,sz=8,bold=(es_fp or es_hito or critica),
            col=C["dark"] if critica else "000000")
        if holg is not None:
            h_bg  = C["ok"] if holg>=30 else (C["warn"] if holg>=7 else C["risk"])
            h_col = C["ok_t"] if holg>=30 else (C["warn_t"] if holg>=7 else C["risk_t"])
            dat(ws,row,6,f"+{holg}d",bg=h_bg,bold=True,sz=8,col=h_col)
        else:
            dat(ws,row,6,"—",bg=bg,sz=7,col="999999")

        for col_ in range(1,N_INFO+1):
            ws.cell(row,col_).border = bd()

        ci_s, cf_s = sem_rango(ini, fin)
        cb = color_barra(tipo, critica)
        if es_warn: cb = "FF0000"

        for i in range(len(SEMANAS)):
            s = SEMANAS[i]
            c2 = ws.cell(row, COL_OFF+i)
            izq = "medium" if s.day <= 7 else "thin"
            c2.border = Border(
                left=Side(style=izq, color="888888" if s.day<=7 else "DDDDDD"),
                right=Side(style="thin",color="DDDDDD"),
                top=Side(style="thin",color="DDDDDD"),
                bottom=Side(style="thin",color="DDDDDD"))
            in_bar = ci_s is not None and cf_s is not None and ci_s<=i<=cf_s
            if in_bar:
                c2.fill = fill(cb)
                if tipo in ("fp","hito_contrato") and i==ci_s:
                    c2.value = "◆" if tipo=="fp" else "★"
                    c2.font = Font(bold=True, color="FFFFFF", size=7)
                    c2.alignment = Alignment(horizontal="center",vertical="center")
                elif tipo=="pago" and i==ci_s:
                    c2.value = "§"
                    c2.font = Font(bold=True, color="FFFFFF", size=7)
                    c2.alignment = Alignment(horizontal="center",vertical="center")
                elif es_warn and i==ci_s:
                    c2.value = "⚠"
                    c2.font = Font(bold=True, color="FFFFFF", size=7)
                    c2.alignment = Alignment(horizontal="center",vertical="center")
            else:
                c2.fill = fill("F8F8F8" if i%2==0 else "F3F3F3")

        ws.row_dimensions[row].height = 13
        row += 1

    ws.merge_cells(f"A{row}:{last_col}{row}"); ws.row_dimensions[row].height = 20
    hdr(ws,row,1,
        "★ CADENA CRÍTICA S#3: FP-IC 31/AGO/2026 → FAT 19/ABR/2027 → Despacho 14/MAY/2027  |  "
        "PMS: Despacho 17/FEB/2027 → FAT Shelter 31/MAR/2027 → SIN IMPACTO EN S3 ✅  |  "
        "Holguras: FAT S3 +35d | Despacho S3 +45d | iFAT +3d",
        C["hdr"],sz=8,h="left")
    ws.freeze_panes = "G6"

# ─────────────────────────────────────────────────────────────────────────────
# HOJA 2: PMS DETALLE
# ─────────────────────────────────────────────────────────────────────────────
def sheet_pms_detalle(wb):
    ws = wb.create_sheet("2_PMS_Detalle")
    ws.sheet_view.showGridLines = False

    for i, w in enumerate([5, 38, 12, 12, 28], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    for i in range(len(SEMANAS)):
        ws.column_dimensions[get_column_letter(6+i)].width = 2.0

    last_col = get_column_letter(6 + len(SEMANAS) - 1)

    ws.merge_cells(f"A1:{last_col}1"); ws.row_dimensions[1].height = 28
    hdr(ws,1,1,
        f"PMS — CRONOGRAMA DETALLADO | OC 4508945953 | PM: Pablo Kalis | "
        f"Ref: E-2611027 | {FECHA_DOC}",
        C["pms"],fg="000000",sz=11)

    # Años y meses
    ws.row_dimensions[2].height = 13
    anio_g = {}
    for i, s in enumerate(SEMANAS):
        a = s.year
        if a not in anio_g: anio_g[a] = [i, i]
        else: anio_g[a][1] = i
    for a, (ci, cf) in anio_g.items():
        try: ws.merge_cells(f"{get_column_letter(6+ci)}2:{get_column_letter(6+cf)}2")
        except: pass
        hdr(ws,2,6+ci,str(a),C["hdr"],sz=9)

    ws.row_dimensions[3].height = 11
    mes_g = {}
    for i, s in enumerate(SEMANAS):
        k = (s.year,s.month)
        if k not in mes_g: mes_g[k]=[i,i]
        else: mes_g[k][1]=i
    for (a,m),(ci,cf) in mes_g.items():
        try: ws.merge_cells(f"{get_column_letter(6+ci)}3:{get_column_letter(6+cf)}3")
        except: pass
        hdr(ws,3,6+ci,MESES[m],"2E75B6" if m%2==0 else "404040",sz=7)

    ws.row_dimensions[4].height = 18
    for c2,t in enumerate(["ID","Actividad","Inicio","Fin","Nota"],1):
        hdr(ws,4,c2,t,C["hdr"],sz=9)

    TIPO_COLOR = {
        "actividad":  ("EBF3FB","2E75B6"),
        "aprobacion": (C["aprovl"],"C55A11"),
        "fp":         ("FFE0E0","C00000"),
        "fat":        ("FFF0E0","E36C09"),
        "despacho":   ("F3E8FF","7030A0"),
        "maqueta":    ("FFF0E0","E36C09"),
        "hito_contrato":("FFF0F0","CC0000"),
        "procura":    ("F9ECE8","833C11"),
        "software":   (C["softl"],"00819A"),
    }

    row = 5
    for (id_, nombre, ini, fin, tipo, nota) in ACT_PMS:
        bg, bar_c = TIPO_COLOR.get(tipo, ("FFFFFF","444444"))
        es_fp = tipo == "fp"
        es_hito = tipo == "hito_contrato"

        dat(ws,row,1,id_,bg=bg,bold=True,sz=8,col=bar_c)
        dat(ws,row,2,nombre,bg=bg,bold=(es_fp or es_hito),sz=9,h="left",
            col="880000" if es_fp else ("CC0000" if es_hito else bar_c))
        dat(ws,row,3,fmt(ini),bg=bg,sz=8,bold=es_fp)
        dat(ws,row,4,fmt(fin),bg=bg,sz=8,bold=(es_fp or es_hito),
            col="880000" if es_fp else "000000")
        dat(ws,row,5,nota,bg=bg,sz=8,h="left",italic=True)

        for col_ in range(1,6):
            ws.cell(row,col_).border = bd()

        ci_s, cf_s = sem_rango(ini, fin)
        for i in range(len(SEMANAS)):
            s2 = SEMANAS[i]
            c2 = ws.cell(row, 6+i)
            c2.border = Border(
                left=Side(style="medium" if s2.day<=7 else "thin",
                          color="888888" if s2.day<=7 else "DDDDDD"),
                right=Side(style="thin",color="DDDDDD"),
                top=Side(style="thin",color="DDDDDD"),
                bottom=Side(style="thin",color="DDDDDD"))
            in_bar = ci_s is not None and cf_s is not None and ci_s<=i<=cf_s
            if in_bar:
                c2.fill = fill(bar_c)
                if tipo in ("fp","hito_contrato") and i==ci_s:
                    c2.value = "◆" if tipo=="fp" else "★"
                    c2.font = Font(bold=True,color="FFFFFF",size=7)
                    c2.alignment = Alignment(horizontal="center",vertical="center")
            else:
                c2.fill = fill("F8F8F8" if i%2==0 else "F3F3F3")

        ws.row_dimensions[row].height = 13
        row += 1

    ws.freeze_panes = "F5"

# ─────────────────────────────────────────────────────────────────────────────
# HOJA 3: COMPATIBILIDAD SHELTER
# ─────────────────────────────────────────────────────────────────────────────
def sheet_compat(wb):
    ws = wb.create_sheet("3_Compatibilidad_Shelter")
    ws.sheet_view.showGridLines = False
    for k,v in {"A":4,"B":12,"C":40,"D":14,"E":14,"F":12,"G":30,"H":4}.items():
        ws.column_dimensions[k].width = v

    ws.merge_cells("B1:G1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:G2"); ws.row_dimensions[2].height = 28
    hdr(ws,2,2,"COMPATIBILIDAD DE FECHAS — SHELTER BOTTINO (MENDOZA) — INTEGRACIÓN PMS / S3 / S4",
        C["hdr"],sz=13)
    ws.merge_cells("B3:G3"); ws.row_dimensions[3].height = 18
    hdr(ws,3,2,
        f"Análisis de convivencia en el Shelter y verificación de impacto en despacho S#3  |  {FECHA_DOC}",
        C["blue"],sz=10)
    ws.row_dimensions[4].height = 8

    # Tabla timeline Shelter
    ws.row_dimensions[5].height = 20
    for i,h in enumerate(["#","Período","Actividad en Shelter","OC/Alcance","Fechas","Riesgo","Observación"],2):
        hdr(ws,5,i,h,C["hdr"],sz=9)

    SHELTER = [
        (1,"ENE-2027","Montaje Shelter S#4","S4 (OC 4508944973)",
         "18-ENE → 05-MAR",C["ok"],"Único alcance activo. Sin conflicto"),
        (2,"FEB-2027","Montaje Shelter S#3 (inicio)","S3 (OC 4508944971)",
         "01-FEB → 26-MAR",C["warn"],"S3 y PMS coinciden en Shelter desde 18-FEB"),
        (3,"FEB-2027","LLEGADA PMS al Shelter","PMS (OC 4508945953)",
         "17-FEB (llegada)","FFEB9C","PMS despachado desde BsAs. Comienza montaje 18-FEB"),
        (4,"FEB-MAR-2027","Montaje PMS (PMS-001+PMS-101)","PMS",
         "18-FEB → 17-MAR",C["warn"],"Simultáneo con S3 montaje. Coordinar con Bottino"),
        (5,"MAR-2027","FAT Sala Eléctrica #4 (Shelter S4 completo)","S4",
         "08-MAR → 16-MAR",C["warn"],"En paralelo con PMS montaje (hasta 17-MAR). Coordinar acceso"),
        (6,"MAR-2027","Supervisión + Preparativos FAT Integral PMS","PMS",
         "18-MAR → 24-MAR",C["ok"],"S4 ya despachado (08-ABR). S3 montaje finaliza 26-MAR"),
        (7,"MAR-2027","★ FAT INTEGRAL SHELTER PMS","PMS",
         "25-MAR → 31-MAR",C["ok"],"S4 despacho 02-ABR. S3 FAT comienza 02-ABR. BUFFER: 2 días ⚠"),
        (8,"ABR-2027","★ FAT Sala Eléctrica #3 (Shelter S3 completo)","S3",
         "02-ABR → 19-ABR",C["ok"],"PMS-001 integrado y validado. Sin conflicto. +35d vs P6"),
        (9,"ABR-2027","DESPACHO Sala #4 → Vaca Muerta","S4",
         "02-ABR → 08-ABR",C["ok"],"OC deadline 08-ABR ✅. Holgura +81d vs P6"),
        (10,"MAY-2027","DESPACHO Sala #3 → Vaca Muerta","S3",
         "10-MAY → 14-MAY",C["ok"],"OC deadline 14-MAY ✅. Sin impacto de PMS. Holgura +45d vs P6"),
    ]

    for row_off, (num,per,act,oc2,fec,bg2,obs) in enumerate(SHELTER):
        r = 6 + row_off
        ws.row_dimensions[r].height = 28
        dat(ws,r,2,num,bg=bg2,bold=True,sz=11,col=C["hdr"])
        dat(ws,r,3,per,bg=bg2,bold=True,sz=9)
        dat(ws,r,4,act,bg=bg2,bold=True,sz=9,h="left",wrap=True)
        dat(ws,r,5,oc2,bg=bg2,sz=8)
        dat(ws,r,6,fec,bg=bg2,bold=True,sz=9)
        riesgo = "✅ OK" if bg2==C["ok"] else "⚠ COORDINAR"
        dat(ws,r,7,riesgo,bg=bg2,bold=True,sz=9,
            col=C["ok_t"] if bg2==C["ok"] else C["warn_t"])
        dat(ws,r,8,obs,bg=bg2,sz=8,h="left",wrap=True,italic=True)

    tborder(ws,5,2,6+len(SHELTER)-1,8)

    # Resumen de compatibilidad
    nr = 6 + len(SHELTER) + 2
    ws.merge_cells(f"B{nr}:G{nr}"); ws.row_dimensions[nr].height = 22
    hdr(ws,nr,2,"CONCLUSIONES — ANÁLISIS DE COMPATIBILIDAD DE FECHAS",C["hdr"],sz=11)

    conclusiones = [
        (C["ok"],C["ok_t"],"✅ COMPATIBLE",
         "FAT INTEGRAL PMS (31-MAR) completa 2 días antes del FAT S#3 (02-ABR). PMS-001 integrado y validado previo al FAT S#3."),
        (C["ok"],C["ok_t"],"✅ COMPATIBLE",
         "DESPACHO PMS (17-FEB) llega al Shelter con 44 días de margen antes del FAT S#3 (02-ABR). Tiempo suficiente para montaje PMS-001+PMS-101."),
        (C["ok"],C["ok_t"],"✅ COMPATIBLE",
         "FAT S4 (8-16/MAR) termina 9 días antes de FAT INTEGRAL PMS (25-MAR). Sin superposición de FATs."),
        (C["ok"],C["ok_t"],"✅ SIN IMPACTO",
         "PMS DataBook (31-MAR) y todas las actividades PMS completan ANTES del deadline contractual S#3 (14-MAY-2027)."),
        (C["warn"],C["warn_t"],"⚠ COORDINAR",
         "MARZO 2027: Shelter simultáneamente con S3-montaje (1-FEB/26-MAR), PMS-montaje (18-FEB/17-MAR) y S4-FAT (8-16/MAR). Requiere coordinación Bottino/ABB/AESA para acceso físico."),
        (C["warn"],C["warn_t"],"⚠ BUFFER AJUSTADO",
         "Solo 2 días de buffer entre FAT INTEGRAL PMS (31-MAR) y FAT S#3 (02-ABR). Cualquier demora en FAT Integral PMS impacta directamente el FAT S#3."),
        (C["risk"],C["risk_t"],"🔴 ACCIÓN REQUERIDA",
         "Confirmar con Bottino (shelterista) la disponibilidad simultánea del Shelter para S3+PMS en FEB-MAR 2027. Necesario en cronograma final (SEM 01-JUN)."),
    ]

    for i, (bg2, fc2, estado, texto) in enumerate(conclusiones):
        r = nr + 1 + i
        ws.row_dimensions[r].height = 28
        ws.merge_cells(f"C{r}:D{r}")
        hdr(ws,r,3,estado,bg=bg2,fg=fc2,sz=9)
        ws.merge_cells(f"E{r}:G{r}")
        dat(ws,r,5,texto,bg=bg2,sz=9,h="left",wrap=True,col=fc2)

    tborder(ws,nr,2,nr+len(conclusiones),7)

# ─────────────────────────────────────────────────────────────────────────────
# HOJA 4: FREEZING POINTS (ACTUALIZADO)
# ─────────────────────────────────────────────────────────────────────────────
def sheet_fp(wb):
    ws = wb.create_sheet("4_Freezing_Points")
    ws.sheet_view.showGridLines = False
    for k,v in {"A":4,"B":8,"C":40,"D":14,"E":12,"F":14,"G":30,"H":4}.items():
        ws.column_dimensions[k].width = v

    ws.merge_cells("B1:G1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:G2"); ws.row_dimensions[2].height = 28
    hdr(ws,2,2,f"FREEZING POINTS — REGISTRO Y SEGUIMIENTO  |  {FECHA_DOC}",C["dark"],sz=14)
    ws.merge_cells("B3:G3"); ws.row_dimensions[3].height = 18
    hdr(ws,3,2,"Un cambio post-FP requiere Nota de Cambio + impacto costo/plazo. FP-P1 y FP-P2 actualizados per Cronograma ABB RevA (MS Project 27-MAY-2026)",
        "880000",sz=9)
    ws.row_dimensions[4].height = 8
    ws.row_dimensions[5].height = 22
    for i,h in enumerate(["#","ID","Freezing Point","Fecha Plan","Estado","Impacto si demora"],2):
        hdr(ws,5,i,h,C["dark"],sz=9)

    FPS = [
        ("FP-P1","◆ Ing. Básica PMS",date(2026,8,21),False,
         "Planos Constructivos Rev0 aprobados (PP-B2). Congela: arquitectura 800XA + I/O 30% reserva + protocolo Modbus TCP/IP",
         "Demora FP-P1 → retrasa fabricación tableros + FAT Maqueta y FAT Tableros. Tiene holgura amplia pero buffer FAT Shelter se comprime",
         "ACTUALIZADO: era 19-JUN, ahora 21-AGO per cronograma ABB RevA"),
        ("FP-P2","◆ Ing. Detalle + Procedimientos PMS",date(2026,10,23),False,
         "Proc.FAT+SAT Rev0 aprobados (PP-B6/B7). Habilita fabricación definitiva tableros PMS y ejecución FAT 29-ENE-2027",
         "Demora FP-P2 → FAT Tableros (29-ENE) se corre → Despacho (17-FEB) tarde → Montaje en Shelter tarde → FAT Integral Shelter tarde → impacto sobre FAT S#3",
         "ACTUALIZADO: era 31-JUL, ahora 23-OCT per cronograma ABB RevA"),
        ("FP-S4-1","◆ Ing. Básica Sala #4",date(2026,6,26),False,
         "Inicio OC celdas MT Turquía (lead 6m). Entrega celdas: Dic 2026",
         "Demora OC Turquía → celdas llegan tarde → FAT S4 tarde → Despacho S4 tarde (holgura 81d actual)",
         "Sin cambios vs cronograma previo"),
        ("FP-S4-2","◆ Ing. Detalle Sala #4",date(2026,8,14),False,
         "Inicio fabricación tableros auxiliares Bs.As.",
         "Demora → montaje Shelter S4 tarde → FAT S4 tarde",
         "Sin cambios"),
        ("FP-S4-3","◆ Ing. Constructiva Sala #4  ← CIERRA ING. S4",date(2026,8,28),False,
         "ÚLTIMO FP S4. Cierre total ingeniería. Sin hito P6 ⚠",
         "Demora directa sobre montaje Shelter S4 → FAT S4 tarde. Margen S4: 81d sobre P6",
         "Sin cambios"),
        ("FP-S3-1","◆ Ing. Básica Sala #3",date(2026,7,3),True,
         "Inicio OC ductos barras Turquía (lead 6m). CADENA CRÍTICA",
         "1 día de demora → OC Turquía tarde → ductos llegan tarde → FAT S3 tarde → PENALIDAD",
         "Sin cambios"),
        ("FP-S3-2","◆ Ing. Detalle Sala #3",date(2026,8,14),True,
         "Inicio fabricación 37 CCMs Brasil. CADENA CRÍTICA",
         "Demora → fabricación CCMs tarde → ensayos tarde → FAT S#3 tarde → Despacho tarde → PENALIDAD",
         "Sin cambios"),
        ("FP-S3-3","◆ Ing. Constructiva S#3  ★ ÚLTIMO FP ABB",date(2026,8,31),True,
         "31-AGO-2026: ÚLTIMO FP DE TODO EL PROYECTO ABB. Cierra toda la ingeniería ABB",
         "1 semana late → montaje Shelter S3 late → FAT 1 semana late → Despacho 1 semana late → PENALIDAD 1%/semana",
         "Sin cambios — sigue siendo el hito más crítico del proyecto"),
    ]

    for idx,(id_,nombre,fecha,critica,que_congela,impacto,nota) in enumerate(FPS):
        r = 6 + idx
        ws.row_dimensions[r].height = 36
        bg = C["s3l"] if critica else (C["pmsl"] if "P" in id_ else C["s4l"])
        fc = C["dark"] if critica else ("B8860B" if "P" in id_ else C["blue"])
        dat(ws,r,2,idx+1,bg=bg,bold=True,sz=11,col=C["hdr"])
        dat(ws,r,3,id_,bg=bg,bold=True,sz=9,col=fc)
        dat(ws,r,4,nombre,bg=bg,bold=critica,sz=9,h="left",wrap=True)
        dat(ws,r,5,fmt(fecha),bg=bg,bold=True,sz=9,col=C["dark"] if critica else fc)
        dat(ws,r,6,C["warn"] if not critica else C["risk"],bg=bg,sz=8,
            col=C["warn_t"] if not critica else C["risk_t"])
        hdr(ws,r,6,"Pendiente ⚠",bg=C["warn"],fg=C["warn_t"],bold=True,sz=8)
        dat(ws,r,7,impacto,bg=bg,sz=8,h="left",wrap=True,
            col=C["risk_t"] if critica else C["warn_t"],italic=False)
        ws.cell(r,7).value = impacto

    tborder(ws,5,2,6+len(FPS)-1,7)

    # nota actualización
    nr = 6+len(FPS)+1
    ws.merge_cells(f"B{nr}:G{nr}"); ws.row_dimensions[nr].height = 22
    hdr(ws,nr,2,
        "⚠ FP-P1 y FP-P2 actualizados: FP-P1 era 19-JUN → ahora 21-AGO  |  FP-P2 era 31-JUL → ahora 23-OCT  |  Per Cronograma MS Project ABB RevA (E-2611027, 27-MAY-2026)",
        C["risk"],sz=9,h="left")

# ─────────────────────────────────────────────────────────────────────────────
# HOJA 5: HITOS Y PAGOS POR OC
# ─────────────────────────────────────────────────────────────────────────────
def sheet_pagos(wb):
    ws = wb.create_sheet("5_Hitos_Pago_por_OC")
    ws.sheet_view.showGridLines = False
    for k,v in {"A":4,"B":8,"C":36,"D":14,"E":16,"F":12,"G":22,"H":4}.items():
        ws.column_dimensions[k].width = v

    ws.merge_cells("B1:G1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:G2"); ws.row_dimensions[2].height = 28
    hdr(ws,2,2,f"HITOS DE PAGO POR OC — CRONOGRAMA ABB GENERAL PRELIMINAR  |  {FECHA_DOC}",C["pay"],sz=13)
    ws.merge_cells("B3:G3"); ws.row_dimensions[3].height = 18
    hdr(ws,3,2,
        "Penalidad por atraso: 1%/semana, tope 10%  |  3 pólizas: Avance Fabricación + Fiel Cumplimiento (10%) + Fondo de Reparo (10%)",
        "2E75B6",sz=9)
    ws.row_dimensions[4].height = 8

    PAGOS_POR_OC = [
        # OC PMS
        ("OC PMS",4508945953,"USD ~(ver oferta)","14-MAY-2027 (12 meses)","B8860B","FFF8DC",[
            ("10%","Envío IB","17-JUL-2026","Arquitectura + I/O + Planos Constructivos RevA. Póliza Caución"),
            ("10%","Aprobación IB","14-AGO-2026","Documentos IB aprobados Rev0 por AESA"),
            ("25%","Acopio Materiales","05-NOV-2026","HW ABB: PMS + Tableros + Switches en Argentina. Póliza Caución Acopio"),
            ("16%  (80% del 20%)","FAT Fábrica","29-ENE-2027","FAT Tableros+Sistema en fábrica BsAs. Póliza Caución"),
            ("30%","Entrega Shelter","17-FEB-2027","Equipamiento en Shelter Mendoza. Liberación Calidad Comprador"),
            ("4%  (20% del 20%)","FAT Integral Shelter","31-MAR-2027","FAT Integral PMS en Shelter completo"),
            ("5%","DataBook","31-MAR-2027","DataBook aprobado por AESA (manuales + protocolos + planos conforme)"),
        ]),
        # OC S4
        ("OC S4 (MT)",4508944973,"USD 3,686,311.55","08-ABR-2027 (11 meses)","2E75B6","DEEAF1",[
            ("10%","Envío IB","12-JUN-2026","Emisión IB S4. Póliza Caución"),
            ("10%","Aprobación IB","03-JUL-2026","Aprobación IB S4 por AESA"),
            ("25%","Acopio Materiales","~15-ENE-2027","Celdas MT + VFDs + tableros en Argentina. Póliza Caución"),
            ("20%","FAT Shelter S4","16-MAR-2027","FAT Shelter completo S4 en Mendoza. Póliza Caución"),
            ("30%","Contra Entrega","08-ABR-2027","Entrega S4. Liberación Calidad Comprador"),
            ("5%","DataBook S4","~JUN-2027","DataBook S4 aprobado"),
        ]),
        # OC S3
        ("OC S3 (BT) ★ CRÍTICA",4508944971,"USD 8,127,765.50","14-MAY-2027 (12 meses)","880000","FFE0E0",[
            ("10%","Envío IB","12-JUN-2026","Emisión IB S3. Póliza Caución"),
            ("10%","Aprobación IB","03-JUL-2026","Aprobación IB S3 por AESA"),
            ("25%","Acopio Materiales","29-ENE-2027","Ductos Turquía + VFDs + 37 CCMs BT en Argentina. Póliza Caución"),
            ("20%","FAT Shelter S3","19-ABR-2027","FAT Shelter completo S3. Hito más tardío. Póliza Caución"),
            ("30%","Contra Entrega","14-MAY-2027","FECHA CONTRACTUAL. Penalidad si excede"),
            ("5%","DataBook S3","~JUL-2027","DataBook S3 aprobado"),
        ]),
    ]

    row = 5
    for (titulo, oc_num, monto, plazo, hdr_col, bg_col, pagos) in PAGOS_POR_OC:
        ws.row_dimensions[row].height = 8; row+=1

        ws.merge_cells(f"B{row}:G{row}"); ws.row_dimensions[row].height = 22
        hdr(ws,row,2,
            f"  {titulo}  |  OC: {oc_num}  |  Monto: {monto}  |  Plazo: {plazo}",
            hdr_col,sz=10,h="left")
        row+=1

        ws.row_dimensions[row].height = 18
        for i,h in enumerate(["%","Hito","Fecha","Condición / Nota"],2):
            hdr(ws,row,i+1,h,C["hdr"],sz=9)
        row+=1

        for (pct,hito,fecha,cond) in pagos:
            ws.row_dimensions[row].height = 22
            dat(ws,row,3,pct,bg=bg_col,bold=True,sz=10,col=hdr_col)
            dat(ws,row,4,hito,bg=bg_col,bold=True,sz=9,h="left")
            dat(ws,row,5,fecha,bg=bg_col,bold=True,sz=9,col=hdr_col)
            dat(ws,row,7,cond,bg=bg_col,sz=8,h="left",wrap=True,italic=True)
            for c2 in range(3,8):
                ws.cell(row,c2).border = bd()
            row+=1

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    wb = Workbook()
    wb.remove(wb.active)
    sheet_gantt(wb)
    sheet_pms_detalle(wb)
    sheet_compat(wb)
    sheet_fp(wb)
    sheet_pagos(wb)
    wb.save(OUTPUT)
    print(f"Generado: {OUTPUT}")
    fps = sum(1 for a in ACT if a[5]=="fp")
    crits = sum(1 for a in ACT if a[6])
    print(f"Actividades Gantt General: {len(ACT)}  |  FPs: {fps}  |  Cadena crítica: {crits}")
    print(f"Actividades PMS Detalle: {len(ACT_PMS)}")

if __name__ == "__main__":
    main()
