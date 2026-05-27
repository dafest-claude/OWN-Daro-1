"""
Cronograma Maestro de Seguimiento — Provisión ABB (CPF-2 / La Calera II)
Salas Eléctricas #3, #4 y PMS
Incluye: Gantt maestro | Freezing Points | Camino crítico | Holguras vs P6 | Dashboard alertas
"""

import os
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import (
    PatternFill, Font, Alignment, Border, Side, numbers
)
from openpyxl.utils import get_column_letter

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))
OUTPUT = os.path.join(PLANNING_DIR, "Cronograma_Seguimiento_ABB.xlsx")

# ── Paleta ────────────────────────────────────────────────────────────────────
C = {
    "abb_red":    "CC0000",
    "abb_dark":   "880000",
    "critico":    "CC0000",
    "paralelo":   "2E75B6",
    "fat":        "E36C09",
    "despacho":   "7030A0",
    "integracion":"375623",
    "hito":       "C00000",
    "pms":        "B8860B",
    "pms_light":  "FFF8DC",
    "s3_light":   "FFE0E0",
    "s4_light":   "DEEAF1",
    "int_light":  "E2EFDA",
    "pay":        "1F4E79",
    "pay_light":  "BDD7EE",
    "hdr_dark":   "1F3864",
    "hdr_mid":    "2E75B6",
    "hdr_sep":    "404040",
    "ok":         "C6EFCE",
    "warn":       "FFEB9C",
    "risk":       "FFC7CE",
    "na":         "D9D9D9",
    "ok_txt":     "375623",
    "warn_txt":   "9C5700",
    "risk_txt":   "9C0006",
    "gray":       "F2F2F2",
    "white":      "FFFFFF",
}

def fill(c): return PatternFill("solid", fgColor=c)
def bd(thin="BBBBBB", med=None):
    t = Side(style="thin", color=thin)
    m = Side(style="medium", color=med or thin)
    return Border(left=t, right=t, top=t, bottom=t)

# ─────────────────────────────────────────────────────────────────────────────
# DATOS MAESTROS
# ─────────────────────────────────────────────────────────────────────────────
# Campos de cada actividad:
# (grupo, id_act, nombre, inicio_plan, fin_plan, tipo, critica, oc, holgura_p6, nota_tecnica)
#
# tipo: "fase" | "actividad" | "fp" | "fat" | "despacho" | "integracion" | "pago" | "hito_contrato"
# critica: True = cadena crítica (S#3)
# holgura_p6: días de holgura vs cronograma P6 (None si no hay hito P6 equivalente)

ACTIVIDADES = [

    # ══════════════════════════════════════════════════════════════════════════
    # HITOS CONTRACTUALES
    # ══════════════════════════════════════════════════════════════════════════
    ("CONTRATO", "HC-01",
     "◆ Aceptación OC — Inicio oficial ABB",
     date(2026,  5, 12), date(2026,  5, 12), "hito_contrato", True,
     "S3+S4+PMS", None,
     "Christian Bombace acepta las 3 OC. Ref: R-2631036 / E-2611027"),

    ("CONTRATO", "HC-02",
     "Envío cronograma preliminar ABB",
     date(2026,  5, 12), date(2026,  5, 12), "hito_contrato", True,
     "S3+S4+PMS", None,
     "Enviado 12-MAY-2026. PDF R-263100X en carpeta 02_Planning"),

    ("CONTRATO", "HC-03",
     "KOM Salas Eléctricas #3 y #4",
     date(2026,  5, 18), date(2026,  5, 18), "hito_contrato", True,
     "S3+S4", None,
     "Lanzamiento oficial. ABB Ref: R-2631036. PM: Rodrigo Mack"),

    ("CONTRATO", "HC-04",
     "KOM PMS (Power Management System)",
     date(2026,  5, 22), date(2026,  5, 22), "hito_contrato", True,
     "PMS", None,
     "Lanzamiento PMS. ABB Ref: E-2611027. PM: Pablo Kalis"),

    ("CONTRATO", "HC-05",
     "◆ Cronograma final ABB con Hold Points",
     date(2026,  6,  1), date(2026,  6,  1), "hito_contrato", True,
     "S3+S4+PMS", None,
     "COMPROMETIDO EN KOM. Incluye hitos de certificación y OC a 3ros"),

    ("CONTRATO", "HC-06",
     "Primer Informe de Avance ABB",
     date(2026,  6,  5), date(2026,  6,  5), "hito_contrato", True,
     "S3+S4+PMS", None,
     "Formato ABB para primera presentación. Luego mensual o quincenal"),

    # ══════════════════════════════════════════════════════════════════════════
    # PMS — OC 4508945953 | Ref E-2611027 | PM: Pablo Kalis
    # ══════════════════════════════════════════════════════════════════════════
    ("PMS", "P-01",
     "Emisión Ingeniería Básica PMS (Tablero control)",
     date(2026,  5, 22), date(2026,  6,  5), "actividad", False,
     "PMS", None,
     "Planos principales: Single line, layout, arquitectura PMS. Trigger pago 10%"),

    ("PMS", "P-02",
     "Aprobación IB PMS (revisión AESA — 10 días hábiles)",
     date(2026,  6,  8), date(2026,  6, 19), "actividad", False,
     "PMS", None,
     "AESA 10 días hábiles. Comentarios → ABB 10 días hábiles para responder"),

    ("PMS", "FP-P1",
     "◆ FREEZING POINT — Ing. Básica PMS",
     date(2026,  6, 19), date(2026,  6, 19), "fp", False,
     "PMS", None,
     "CONGELAR: arquitectura 800XA + lista señales (30% reserva mín.) + protocolo Modbus TCP/IP (acordado Oct-2025, reunión AESA-INAUCO-ABB). Cambios → NC + impacto costo/plazo"),

    ("PMS", "P-03",
     "Emisión Ingeniería de Detalle PMS",
     date(2026,  6, 22), date(2026,  7, 17), "actividad", False,
     "PMS", None,
     "Planos detalle: esquemas de control, BOM definitivo, documentos de fabricación"),

    ("PMS", "P-04",
     "Aprobación ID PMS (AESA — 10 días hábiles)",
     date(2026,  7, 20), date(2026,  7, 31), "actividad", False,
     "PMS", None,
     "Segunda vuelta aprobación. Trigger inicio fabricación tablero PMS"),

    ("PMS", "FP-P2",
     "◆ FREEZING POINT — Ing. Detalle PMS",
     date(2026,  7, 31), date(2026,  7, 31), "fp", False,
     "PMS", None,
     "CONGELAR DISEÑO ID. Inicio oficial de fabricación tablero de control PMS"),

    ("PMS", "P-05",
     "Procura / Fabricación Tablero PMS (local BsAs)",
     date(2026,  8,  3), date(2027,  1, 29), "actividad", False,
     "PMS", None,
     "Fabricación en ABB Buenos Aires. Gabinetes: 102-PMS-001 (SE#3) y 102-PMS-101 (SE#4). Controlador AC800M redundante. FAT Maqueta Sep-Oct 2026 / FAT Tableros Ene 2027"),

    ("PMS", "P-06",
     "FAT Maqueta PMS — Pruebas Tempranas (ET sec.20)",
     date(2026,  9,  1), date(2026,  9, 30), "fat", False,
     "PMS", None,
     "Sep-Oct 2026. Fábrica ABB Buenos Aires. Pruebas sobre mock-up PMS+CCM+AC800M red.+PROFINET. AESA + Contratista SCADA participan OBLIGATORIAMENTE (ET ACAL-00102-ET-E-0005 sec.20). ⚠ Planificar viaje BsAs"),

    ("PMS", "P-07",
     "Ensayos y FAT PMS (tablero control integrado)",
     date(2026, 10, 19), date(2026, 10, 23), "fat", False,
     "PMS", 222,
     "4 días. FAT tablero PMS integrado con señales. Holgura 222d vs demanda P6 (02/JUN/2027)"),

    ("PMS", "P-08",
     "FAT Tableros PMS — Enero 2027",
     date(2027,  1,  4), date(2027,  1, 29), "fat", False,
     "PMS", None,
     "FAT final tableros PMS en fábrica ABB Buenos Aires antes de integración en Shelter"),

    ("PMS", "P-09",
     "Integración PMS en Shelter (prueba integrada solicitada por AESA)",
     date(2027,  4, 20), date(2027,  5,  9), "fat", False,
     "PMS", None,
     "AESA solicita prueba integrada (PMS + gabinetes) en Shelter antes del despacho. ABB analizando factibilidad"),

    ("PMS", "P-10",
     "▶ DESPACHO Tablero PMS — 14-MAY-2027",
     date(2027,  5, 14), date(2027,  5, 14), "despacho", False,
     "PMS", 222,
     "OC deadline: 14-MAY-2027 (alineado S#3). Gabinetes 102-PMS-001 y 102-PMS-101 despa chados desde BsAs hacia Vaca Muerta"),

    # ══════════════════════════════════════════════════════════════════════════
    # SALA #4 (MT) — OC 4508944973 | 11 meses | Entrega: 08-ABR-2027
    # ══════════════════════════════════════════════════════════════════════════
    ("Sala #4 (MT)", "S4-01",
     "Emisión Ingeniería Básica Sala #4",
     date(2026,  5, 22), date(2026,  6, 12), "actividad", False,
     "S4", None,
     "Planos principales MT: layout, single line 13.2kV/6.6kV, 25 cols ZS1 630A 25kA. Trigger pago 10%"),

    ("Sala #4 (MT)", "S4-02",
     "Aprobación IB Sala #4 (AESA — 10 días hábiles)",
     date(2026,  6, 15), date(2026,  6, 26), "actividad", False,
     "S4", None,
     "AESA aprueba o comenta. Pendiente: AESA debe confirmar ductos de barra y unifilares CCMs"),

    ("Sala #4 (MT)", "FP-S4-1",
     "◆ FREEZING POINT — Ing. Básica Sala #4",
     date(2026,  6, 26), date(2026,  6, 26), "fp", False,
     "S4", None,
     "CONGELAR. Inicio procura celdas MT (Turquía — lead time ~6 meses). CRÍTICO para timing logístico"),

    ("Sala #4 (MT)", "S4-03",
     "OC a 3ros: Celdas MT Turquía (post FP-IB)",
     date(2026,  6, 29), date(2026,  7, 10), "actividad", False,
     "S4", None,
     "Lead time ~6 meses. Entrega prevista Dic 2026 / Ene 2027. Fabricación Turquía"),

    ("Sala #4 (MT)", "S4-04",
     "Emisión Ingeniería de Detalle Sala #4",
     date(2026,  6, 29), date(2026,  7, 17), "actividad", False,
     "S4", None,
     "Esquemas detalle, cableado, BOM definitivo"),

    ("Sala #4 (MT)", "S4-05",
     "Aprobación ID Sala #4 (AESA — 10 días hábiles)",
     date(2026,  7, 20), date(2026,  8, 14), "actividad", False,
     "S4", None,
     "Aprobación engineering de detalle"),

    ("Sala #4 (MT)", "FP-S4-2",
     "◆ FREEZING POINT — Ing. Detalle Sala #4",
     date(2026,  8, 14), date(2026,  8, 14), "fp", False,
     "S4", None,
     "CONGELAR diseño detalle. Inicio fabricación tableros auxiliares BsAs"),

    ("Sala #4 (MT)", "S4-06",
     "Emisión Ingeniería Constructiva Sala #4",
     date(2026,  8, 10), date(2026,  8, 21), "actividad", False,
     "S4", None,
     "Planos constructivos: assembly, wiring, interconexiones en Shelter"),

    ("Sala #4 (MT)", "S4-07",
     "Aprobación IC Sala #4 (AESA — 10 días hábiles)",
     date(2026,  8, 24), date(2026,  8, 28), "actividad", False,
     "S4", None,
     "Aprobación planos constructivos finales"),

    ("Sala #4 (MT)", "FP-S4-3",
     "◆ FREEZING POINT — Ing. Constructiva Sala #4  ← CIERRA INGENIERÍA S4",
     date(2026,  8, 28), date(2026,  8, 28), "fp", False,
     "S4", None,
     "ÚLTIMO FP DE S4 (24-AGO). Cierre total ingeniería Sala #4. Sin hito equivalente en P6 ⚠"),

    ("Sala #4 (MT)", "S4-08",
     "OC a 3ros: Variadores MT China (post FP-IC)",
     date(2026,  9,  1), date(2026,  9, 15), "actividad", False,
     "S4", None,
     "3 VFDs MT. Lead time ~4-5 meses. Entrega prevista Feb 2027"),

    ("Sala #4 (MT)", "S4-09",
     "Fabricación Tableros S#4 (Brasil — M&B)",
     date(2026,  9,  1), date(2026, 11,  2), "actividad", False,
     "S4", None,
     "25 columnas ZS1. Ensayos en Brasil previo despacho a Argentina"),

    ("Sala #4 (MT)", "S4-10",
     "Ensayos Celdas y Tableros MT Sala #4",
     date(2026, 11,  3), date(2026, 11, 23), "fat", False,
     "S4", None,
     "20 días. Pruebas en Brasil — ABB + AESA inspeccionan"),

    ("Sala #4 (MT)", "S4-11",
     "Acopio materiales en Argentina (celdas + VFDs + ductos S#4)",
     date(2026, 12,  7), date(2027,  1, 15), "actividad", False,
     "S4", None,
     "HITO DE PAGO 25%. Trigger: llegada de materiales importados a Argentina"),

    ("Sala #4 (MT)", "S4-12",
     "Montaje e integración en Shelter S#4 (Bottino — Mendoza)",
     date(2027,  1, 18), date(2027,  3,  5), "actividad", False,
     "S4", None,
     "Montaje tableros, VFDs, HVAC, CCTV, racks en Shelter 28x5m"),

    ("Sala #4 (MT)", "S4-13",
     "FAT Sala Eléctrica #4 (Shelter completo)",
     date(2027,  3,  8), date(2027,  3, 16), "fat", False,
     "S4", None,
     "8 días. FAT integrado del Shelter completo. Asiste AESA + Cliente"),

    ("Sala #4 (MT)", "S4-14",
     "▶ DESPACHO Sala #4 → Vaca Muerta",
     date(2027,  4,  2), date(2027,  4,  8), "despacho", False,
     "S4", 81,
     "OC deadline: 08-ABR-2027. P6 inicio montaje campo: 28-JUN-2027. Holgura: 81 días"),

    ("Sala #4 (MT)", "S4-15",
     "★ ENTREGA S#4 — Cumplimiento OC 4508944973",
     date(2027,  4,  8), date(2027,  4,  8), "hito_contrato", False,
     "S4", 81,
     "FECHA CONTRACTUAL: 08-ABR-2027 (11 meses). Penalidad si excede: 1%/semana, tope 10%"),

    # ══════════════════════════════════════════════════════════════════════════
    # SALA #3 (BT) — OC 4508944971 | 12 meses | Entrega: 14-MAY-2027 ← CADENA CRÍTICA
    # ══════════════════════════════════════════════════════════════════════════
    ("Sala #3 (BT) ★ CRÍTICA", "S3-01",
     "Emisión Ingeniería Básica Sala #3",
     date(2026,  5, 22), date(2026,  6, 12), "actividad", True,
     "S3", None,
     "37 CCMs 4000A 65kA + 6 ductos barras. Layout 37x6m + sala baterías 10.5x3m. Trigger pago 10%"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-02",
     "Aprobación IB Sala #3 (AESA — 10 días hábiles)",
     date(2026,  6, 15), date(2026,  6, 26), "actividad", True,
     "S3", None,
     "Pendiente: AESA confirmar ductos de barra (posición transfo). OPEN ISSUE ⚠"),

    ("Sala #3 (BT) ★ CRÍTICA", "FP-S3-1",
     "◆ FREEZING POINT — Ing. Básica Sala #3",
     date(2026,  7,  3), date(2026,  7,  3), "fp", True,
     "S3", None,
     "1 semana después que S4 (26-JUN). CONGELAR. Inicio procura ductos barras (Turquía)"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-03",
     "OC a 3ros: Ductos de Barras Turquía (post FP-IB S3)",
     date(2026,  7,  6), date(2026,  7, 17), "actividad", True,
     "S3", None,
     "6 ductos de barras. Lead time ~5-6 meses. Entrega prevista Dic 2026. CRÍTICO para fabricación"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-04",
     "OC a 3ros: Variadores BT Finlandia (post FP-IB S3)",
     date(2026,  7,  6), date(2026,  7, 17), "actividad", True,
     "S3", None,
     "3 VFDs BT. Lead time ~4-5 meses. Entrega prevista Dic 2026 - Ene 2027"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-05",
     "Emisión Ingeniería de Detalle Sala #3",
     date(2026,  6, 29), date(2026,  7, 31), "actividad", True,
     "S3", None,
     "Esquemas detalle, cableado 37 CCMs, BOM definitivo. Pendiente: cambio CCM7 ⚠"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-WARN",
     "⚠ OPEN ISSUE: Confirmar cambio CCM7 (3x22kW → 2x45kW+VFD)",
     date(2026,  5, 26), date(2026,  6,  5), "actividad", True,
     "S3", None,
     "AESA debe confirmar antes del 05-JUN-2026. Impacta IB/ID S3. VFDs 102-DP-VFD-24210 A/B/C (SE#3) y 102-DP-VFD-23310 A/B/C (SE#4) en scope PMS vía PROFINET → decisión afecta señales del sistema. Responsable: AESA"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-06",
     "Aprobación ID Sala #3 (AESA — 10 días hábiles)",
     date(2026,  8,  3), date(2026,  8, 14), "actividad", True,
     "S3", None,
     "Aprobación engineering detalle. 5-6 sem después de la IB"),

    ("Sala #3 (BT) ★ CRÍTICA", "FP-S3-2",
     "◆ FREEZING POINT — Ing. Detalle Sala #3",
     date(2026,  8, 14), date(2026,  8, 14), "fp", True,
     "S3", None,
     "CONGELAR diseño detalle S3. Inicio fabricación tableros BT (Brasil). Sin hito P6 ⚠"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-07",
     "Emisión Ingeniería Constructiva Sala #3",
     date(2026,  8, 17), date(2026,  8, 28), "actividad", True,
     "S3", None,
     "Planos constructivos: assembly 37 CCMs, wiring, interconexiones Shelter 37x6m"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-08",
     "Aprobación IC Sala #3 (AESA — 10 días hábiles)",
     date(2026,  9,  1), date(2026,  9, 12), "actividad", True,
     "S3", None,
     "AESA revisa planos constructivos finales"),

    ("Sala #3 (BT) ★ CRÍTICA", "FP-S3-3",
     "◆ FREEZING POINT — Ing. Constructiva Sala #3  ← ÚLTIMO FP ABB",
     date(2026,  8, 31), date(2026,  8, 31), "fp", True,
     "S3", None,
     "★ HITO CLAVE: ÚLTIMO FREEZING POINT DE TODO EL PROYECTO ABB — 31-AGO-2026. Cierra toda la ingeniería. Sin hito en P6 ⚠"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-09",
     "Fabricación Tableros S#3 (37 CCMs — Brasil M&B)",
     date(2026,  9, 14), date(2026, 11,  2), "actividad", True,
     "S3", None,
     "37 tableros CCM 4000A 65kA. Fabricación en Brasil. Ensayos en origen"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-10",
     "Ensayos Celdas y Tableros BT Sala #3",
     date(2026, 11,  3), date(2026, 11, 23), "fat", True,
     "S3", None,
     "20 días. Ensayos de rutina + tipo en Brasil. AESA inspecciona (10 días de antelación)"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-11",
     "Acopio materiales en Argentina (celdas + VFDs + ductos S#3)",
     date(2026, 12,  7), date(2027,  1, 29), "actividad", True,
     "S3", None,
     "HITO DE PAGO 25%. Ductos Turquía + VFDs BT Finlandia + tableros BT Brasil llegan a Argentina"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-12",
     "Montaje e integración en Shelter S#3 (Bottino — Mendoza)",
     date(2027,  2,  1), date(2027,  3, 26), "actividad", True,
     "S3", None,
     "37 CCMs + 6 ductos barras + 3 VFDs BT + HVAC + baterías KMT + tableros aux Mehcco"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-13",
     "FAT Sala Eléctrica #3  ← FAT MÁS TARDÍO DE ABB",
     date(2027,  4,  2), date(2027,  4, 19), "fat", True,
     "S3", 35,
     "17 días. FAT integrado Shelter S#3 completo. P6 cierra FDE1103 el 24-MAY-2027. Holgura: 35 días"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-14",
     "▶ DESPACHO Sala #3 ← DESPACHO MÁS TARDÍO DE ABB",
     date(2027,  5, 10), date(2027,  5, 14), "despacho", True,
     "S3", 45,
     "ÚLTIMO DESPACHO ABB. OC deadline: 14-MAY-2027. P6 inicio montaje campo: 28-JUN-2027. Holgura: 45 días"),

    ("Sala #3 (BT) ★ CRÍTICA", "S3-15",
     "★ ENTREGA S#3 — Cumplimiento OC 4508944971",
     date(2027,  5, 14), date(2027,  5, 14), "hito_contrato", True,
     "S3", 45,
     "FECHA CONTRACTUAL: 14-MAY-2027 (12 meses). CADENA CRÍTICA — determina todo el cronograma ABB"),

    # ══════════════════════════════════════════════════════════════════════════
    # HITOS DE PAGO — Aplican a las 3 OC
    # ══════════════════════════════════════════════════════════════════════════
    ("HITOS DE PAGO", "PAY-01",
     "§ 10% — Envío Ingeniería Básica (S3 + S4 + PMS)",
     date(2026,  6, 12), date(2026,  6, 12), "pago", False,
     "S3+S4+PMS", None,
     "Emisión IB de los tres alcances. Póliza Caución Avance Fabricación requerida previamente"),

    ("HITOS DE PAGO", "PAY-02",
     "§ 10% — Aprobación Ingeniería Básica (AESA aprueba)",
     date(2026,  7,  3), date(2026,  7,  3), "pago", False,
     "S3+S4+PMS", None,
     "Aprobación final IB por AESA. Póliza Caución Avance Fabricación vigente"),

    ("HITOS DE PAGO", "PAY-03",
     "§ 20% — FAT completado (S3 más tardío: 19-ABR-2027)",
     date(2027,  4, 19), date(2027,  4, 19), "pago", False,
     "S3", None,
     "FAT de integración S#3. Póliza Caución Avance Fabricación requerida. PMS: ver acuerdo FAT en Shelter"),

    ("HITOS DE PAGO", "PAY-04",
     "§ 25% — Acopio Materiales Principales en Argentina",
     date(2027,  1, 29), date(2027,  1, 29), "pago", False,
     "S3+S4+PMS", None,
     "Celdas MT + tableros BT arco interno + ductos + VFDs MT/BT + PMS + UPS + aux. Póliza Caución Acopio"),

    ("HITOS DE PAGO", "PAY-05",
     "§ 30% — Contra Entrega (Shelter listo en Mendoza)",
     date(2027,  5, 14), date(2027,  5, 14), "pago", False,
     "S3+S4+PMS", None,
     "30 días desde Despacho S#3 (última entrega). Liberación por Calidad del COMPRADOR"),

    ("HITOS DE PAGO", "PAY-06",
     "§ 5% — Entrega Databook aprobado",
     date(2027,  7,  1), date(2027,  7,  1), "pago", False,
     "S3+S4+PMS", None,
     "Databook completo y aprobado por AESA. Fecha estimada ~6-8 semanas post entrega"),

    # ══════════════════════════════════════════════════════════════════════════
    # INTEGRACIÓN EPC (campo)
    # ══════════════════════════════════════════════════════════════════════════
    ("INTEGRACIÓN EPC", "INT-01",
     "Inicio Precomisionado",
     date(2027,  5,  3), date(2028,  1, 13), "integracion", False,
     "EPC", None,
     "CDW1000. FECHA IDÉNTICA en P6 y Gantt ABB ✅ Responsabilidad EPC / AESA"),

    ("INTEGRACIÓN EPC", "INT-02",
     "iFAT — Integración Shelters + PMS + INAUCO + HIMA",
     date(2027,  6, 25), date(2027,  7, 14), "fat", True,
     "S3+S4+PMS", 3,
     "P6: TA1124 = 28-JUN-2027. ABB arranca 25-JUN. Todos los vendors integran. Holgura: 3 días"),

    ("INTEGRACIÓN EPC", "INT-03",
     "Inicio montaje Shelters en campo (S#3 y S#4)",
     date(2027,  6, 28), date(2027,  8, 21), "integracion", False,
     "EPC", None,
     "CE014050 / CE018050. Campo Vaca Muerta. Responsabilidad EPC. Shelter S#3 hasta 10-AGO, S#4 hasta 21-AGO"),

    ("INTEGRACIÓN EPC", "INT-04",
     "Inicio Comisionado",
     date(2027,  7, 16), date(2028,  1, 31), "integracion", True,
     "EPC", None,
     "CDW1010. FECHA IDÉNTICA en P6 y Gantt ABB ✅"),

    ("INTEGRACIÓN EPC", "INT-05",
     "★ RFSU — Ready For Start Up Planta Total",
     date(2028,  1, 31), date(2028,  1, 31), "hito_contrato", True,
     "PROYECTO", None,
     "A10290. FECHA IDÉNTICA P6 y ABB ✅ Objetivo firme. Fin garantía extendida: 31-ENE-2029"),

    ("INTEGRACIÓN EPC", "INT-06",
     "SAT PMS — Pruebas en Sitio (100% puntos E/S)",
     date(2027, 10,  1), date(2027, 12, 15), "actividad", False,
     "PMS", None,
     "100% puntos E/S ensayados en campo (ET ACAL-00102-ET-E-0005 sec.20). Asistencia ABB para IEDs de distintos fabricantes. Protocolo SAT aprobado por PP antes del inicio. Lista puntos pendientes debe resolverse antes del RFSU"),

    ("INTEGRACIÓN EPC", "INT-07",
     "Capacitación PMS — 1ra Instancia (Ingeniería y Mantenimiento)",
     date(2028,  2,  5), date(2028,  2,  9), "actividad", False,
     "PMS", None,
     "Capacitación en sitio per ET sec.23. Cursos en castellano para Ingeniería y Mantenimiento. ABB debe presentar propuesta con lista de cursos, duración y cantidad de personas"),

    ("INTEGRACIÓN EPC", "INT-08",
     "Capacitación PMS — 2da Instancia (Operación)",
     date(2028,  3,  3), date(2028,  3,  7), "actividad", False,
     "PMS", None,
     "2da fecha en sitio per ET sec.23. Grupo: Operación. Personal debe quedar apto para operar, mantener y hacer ingeniería del PMS sin asistencia del proveedor"),
]

# ─────────────────────────────────────────────────────────────────────────────
# TIMELINE DEL GANTT
# ─────────────────────────────────────────────────────────────────────────────
INICIO_GANTT = date(2026,  5,  1)
FIN_GANTT    = date(2028,  2, 28)
SEMANAS = []
d = INICIO_GANTT
while d <= FIN_GANTT:
    SEMANAS.append(d)
    d += timedelta(weeks=1)

N_INFO   = 6   # ID | Actividad | OC | Inicio | Fin | Holgura P6
COL_OFF  = N_INFO + 1

# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────
def sem_idx(fecha):
    for i, s in enumerate(SEMANAS):
        if s <= fecha < s + timedelta(days=7):
            return i
    return None

def rango_sem(ini, fin):
    if ini == fin:
        idx = sem_idx(ini)
        return (idx, idx) if idx is not None else (None, None)
    ci = cf = None
    for i, s in enumerate(SEMANAS):
        sw = s + timedelta(days=6)
        if ci is None and sw >= ini:
            ci = i
        if s <= fin:
            cf = i
    return ci, cf

def color_barra(tipo, critica):
    m = {
        "fp":             C["hito"],
        "fat":            C["fat"],
        "despacho":       C["despacho"],
        "integracion":    C["integracion"],
        "pago":           C["pay"],
        "hito_contrato":  C["abb_red"],
    }
    if tipo in m:
        return m[tipo]
    return C["critico"] if critica else C["paralelo"]

def fmt_date(d):
    if d is None:
        return "—"
    return d.strftime("%d/%m/%Y")

MESES_ES = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

def hdr(ws, row, col, txt, bg, fg="FFFFFF", bold=True, sz=9, wrap=False, h="center"):
    c = ws.cell(row=row, column=col, value=txt)
    c.fill = fill(bg)
    c.font = Font(name="Calibri", bold=bold, size=sz, color=fg)
    c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    return c

def dat(ws, row, col, val, bg=None, bold=False, sz=9, h="center",
        color="000000", italic=False, wrap=False):
    c = ws.cell(row=row, column=col, value=val)
    c.font = Font(name="Calibri", bold=bold, size=sz, color=color, italic=italic)
    if bg:
        c.fill = fill(bg)
    c.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    return c

def thin_border(ws, r1, c1, r2, c2, color="BBBBBB"):
    t = Side(style="thin", color=color)
    bd_ = Border(left=t, right=t, top=t, bottom=t)
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).border = bd_

# ─────────────────────────────────────────────────────────────────────────────
# HOJA 1: GANTT MAESTRO
# ─────────────────────────────────────────────────────────────────────────────
def sheet_gantt(wb):
    ws = wb.create_sheet("1_Gantt_Maestro")
    ws.sheet_view.showGridLines = False

    # ancho columnas de info
    widths_info = [5, 42, 10, 11, 11, 10]
    for i, w in enumerate(widths_info, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # ancho columnas semana
    for i in range(len(SEMANAS)):
        ws.column_dimensions[get_column_letter(COL_OFF + i)].width = 2.2

    last_col = get_column_letter(COL_OFF + len(SEMANAS) - 1)
    last_ci  = COL_OFF + len(SEMANAS) - 1

    # ── Título ────────────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    ws.row_dimensions[1].height = 30
    hdr(ws, 1, 1,
        "GANTT MAESTRO DE SEGUIMIENTO — PROVISIÓN ABB  |  CPF-2 LA CALERA II (VACA MUERTA)  |  RFSU: 31/01/2028",
        C["abb_red"], sz=12)

    # ── Leyenda ────────────────────────────────────────────────────────────────
    ws.merge_cells(f"A2:{last_col}2")
    ws.row_dimensions[2].height = 13
    leyenda_items = [
        (C["critico"],   "★ Cadena Crítica (S#3)"),
        (C["paralelo"],  "Actividades Paralelas (S#4/PMS)"),
        (C["fat"],       "FAT / Ensayos"),
        (C["despacho"],  "Despacho"),
        (C["integracion"],"Integración/Comisionado"),
        (C["hito"],      "◆ Freezing Point"),
        (C["pay"],       "§ Hito de Pago"),
        (C["abb_red"],   "★ Hito Contractual"),
    ]
    leg = "    ".join(f"■ {t}" for _, t in leyenda_items)
    c2 = ws.cell(row=2, column=1, value=leg)
    c2.font = Font(name="Calibri", size=7, color="444444")
    c2.alignment = Alignment(horizontal="left", vertical="center")

    # ── Fila año ──────────────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 14
    anio_g = {}
    for i, s in enumerate(SEMANAS):
        a = s.year
        if a not in anio_g:
            anio_g[a] = [i, i]
        else:
            anio_g[a][1] = i
    for a, (ci, cf) in anio_g.items():
        c1s = get_column_letter(COL_OFF + ci)
        cfs = get_column_letter(COL_OFF + cf)
        try:
            ws.merge_cells(f"{c1s}3:{cfs}3")
        except Exception:
            pass
        hdr(ws, 3, COL_OFF + ci, str(a), C["hdr_dark"], sz=9)

    # ── Fila mes ──────────────────────────────────────────────────────────────
    ws.row_dimensions[4].height = 13
    mes_g = {}
    for i, s in enumerate(SEMANAS):
        k = (s.year, s.month)
        if k not in mes_g:
            mes_g[k] = [i, i]
        else:
            mes_g[k][1] = i
    for (a, m), (ci, cf) in mes_g.items():
        c1s = get_column_letter(COL_OFF + ci)
        cfs = get_column_letter(COL_OFF + cf)
        try:
            ws.merge_cells(f"{c1s}4:{cfs}4")
        except Exception:
            pass
        bg = C["hdr_mid"] if m % 2 == 0 else C["hdr_sep"]
        hdr(ws, 4, COL_OFF + ci, MESES_ES[m], bg, sz=7)

    # ── Encabezados info ──────────────────────────────────────────────────────
    ws.row_dimensions[5].height = 20
    for col, txt in enumerate(["ID", "Actividad / Hito", "OC", "Inicio Plan", "Fin Plan", "Holgura P6"], 1):
        hdr(ws, 5, col, txt, C["hdr_dark"], sz=9)

    # ── Filas de datos ────────────────────────────────────────────────────────
    grupo_ant = None
    row = 6

    GRUPO_BG = {
        "CONTRATO":                 C["abb_red"],
        "PMS":                      "B8860B",
        "Sala #4 (MT)":             C["hdr_mid"],
        "Sala #3 (BT) ★ CRÍTICA":  C["abb_dark"],
        "HITOS DE PAGO":            C["pay"],
        "INTEGRACIÓN EPC":          C["integracion"],
    }
    FILA_BG = {
        "CONTRATO":                 "FFF0F0",
        "PMS":                      C["pms_light"],
        "Sala #4 (MT)":             C["s4_light"],
        "Sala #3 (BT) ★ CRÍTICA":  C["s3_light"],
        "HITOS DE PAGO":            C["pay_light"],
        "INTEGRACIÓN EPC":          C["int_light"],
    }

    for (grupo, id_act, nombre, ini, fin, tipo, critica, oc, holgura, nota) in ACTIVIDADES:

        # separador de grupo
        if grupo != grupo_ant:
            ws.merge_cells(f"A{row}:{last_col}{row}")
            ws.row_dimensions[row].height = 16
            bg_sep = GRUPO_BG.get(grupo, C["hdr_sep"])
            lbl = f"  ▶  {grupo.upper()}"
            if "CRÍTICA" in grupo:
                lbl += "   ← CADENA CRÍTICA  |  OC 4508944971  |  12 meses  |  Entrega: 14-MAY-2027"
            hdr(ws, row, 1, lbl, bg_sep, sz=9, h="left")
            row += 1
            grupo_ant = grupo

        bg_fila = FILA_BG.get(grupo, "FFFFFF")
        es_fp   = tipo == "fp"
        es_pago = tipo == "pago"
        es_hito = tipo == "hito_contrato"
        es_warn = "WARN" in id_act

        # columnas info
        id_color = C["abb_dark"] if critica else C["hdr_mid"]
        dat(ws, row, 1, id_act,  bg=bg_fila, bold=True, sz=8, color=id_color)
        dat(ws, row, 2, nombre,  bg=bg_fila, bold=(critica or es_fp or es_hito),
            sz=9, h="left", wrap=False,
            color=(C["abb_dark"] if critica else ("000000" if not es_pago else C["pay"])))
        dat(ws, row, 3, oc,      bg=bg_fila, sz=8, h="center")
        dat(ws, row, 4, fmt_date(ini), bg=bg_fila, sz=8, h="center",
            bold=es_fp or es_hito or critica,
            color=(C["abb_dark"] if critica else "000000"))
        dat(ws, row, 5, fmt_date(fin), bg=bg_fila, sz=8, h="center",
            bold=es_fp or es_hito or critica,
            color=(C["abb_dark"] if (critica and not es_hito) else "000000"))

        # columna holgura P6
        if holgura is not None:
            h_txt   = f"+{holgura}d"
            h_bg    = C["ok"] if holgura >= 30 else (C["warn"] if holgura >= 7 else C["risk"])
            h_color = C["ok_txt"] if holgura >= 30 else (C["warn_txt"] if holgura >= 7 else C["risk_txt"])
            dat(ws, row, 6, h_txt, bg=h_bg, bold=True, sz=8, color=h_color)
        else:
            dat(ws, row, 6, "—", bg=bg_fila, sz=7, color="999999")

        # bordes info
        for col in range(1, N_INFO + 1):
            t = Side(style="thin", color="CCCCCC")
            ws.cell(row, col).border = Border(left=t, right=t, top=t, bottom=t)

        # ── barras Gantt ───────────────────────────────────────────────────────
        ci_s, cf_s = rango_sem(ini, fin)
        cb = color_barra(tipo, critica)
        if es_warn:
            cb = "FF0000"

        for i in range(len(SEMANAS)):
            sem = SEMANAS[i]
            c = ws.cell(row, COL_OFF + i)
            # borde de mes
            izq = "medium" if sem.day <= 7 else "thin"
            izq_c = "888888" if sem.day <= 7 else "DDDDDD"
            c.border = Border(
                left=Side(style=izq, color=izq_c),
                right=Side(style="thin", color="DDDDDD"),
                top=Side(style="thin", color="DDDDDD"),
                bottom=Side(style="thin", color="DDDDDD"),
            )

            in_bar = ci_s is not None and cf_s is not None and ci_s <= i <= cf_s
            if in_bar:
                c.fill = fill(cb)
                if (tipo == "fp" or tipo == "hito_contrato") and i == ci_s:
                    c.value = "◆" if tipo == "fp" else "★"
                    c.font = Font(bold=True, color="FFFFFF", size=7)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif tipo == "pago" and i == ci_s:
                    c.value = "§"
                    c.font = Font(bold=True, color="FFFFFF", size=7)
                    c.alignment = Alignment(horizontal="center", vertical="center")
                elif es_warn and i == ci_s:
                    c.value = "⚠"
                    c.font = Font(bold=True, color="FFFFFF", size=7)
                    c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.fill = fill("F8F8F8" if i % 2 == 0 else "F3F3F3")

        ws.row_dimensions[row].height = 14 if not (es_fp or es_hito) else 13
        row += 1

    # ── Nota final ────────────────────────────────────────────────────────────
    ws.merge_cells(f"A{row}:{last_col}{row}")
    ws.row_dimensions[row].height = 22
    nota = ("★ CADENA CRÍTICA S#3: 37 CCMs BT → FP-IC 31/AGO/2026 → FAT 19/ABR/2027 → Despacho 14/MAY/2027  |  "
            "Holguras: FAT S3 +35d | Despacho S3 +45d | iFAT +3d | RFSU 31/ENE/2028 = IGUAL EN P6 ✅")
    hdr(ws, row, 1, nota, C["hdr_dark"], sz=8, h="left")

    ws.freeze_panes = "G6"
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# HOJA 2: FREEZING POINTS — tabla de seguimiento
# ─────────────────────────────────────────────────────────────────────────────
def sheet_freezing(wb):
    ws = wb.create_sheet("2_Freezing_Points")
    ws.sheet_view.showGridLines = False

    cols = {"A": 4, "B": 8, "C": 40, "D": 14, "E": 14, "F": 14,
            "G": 18, "H": 22, "I": 28, "J": 4}
    for k, v in cols.items():
        ws.column_dimensions[k].width = v

    ws.merge_cells("B1:I1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:I2"); ws.row_dimensions[2].height = 28
    hdr(ws, 2, 2, "FREEZING POINTS — REGISTRO Y SEGUIMIENTO", C["abb_dark"], sz=14)
    ws.merge_cells("B3:I3"); ws.row_dimensions[3].height = 18
    hdr(ws, 3, 2,
        "Congelamiento de diseño: un cambio post-FP requiere Nota de Cambio + impacto costo/plazo",
        "880000", sz=10)
    ws.row_dimensions[4].height = 8

    hdrs = ["#", "ID", "Freezing Point", "Fecha Plan", "Fecha Real",
            "Días δ", "Estado", "Impacto si se demora", "Acciones / Notas"]
    ws.row_dimensions[5].height = 22
    for i, h in enumerate(hdrs):
        hdr(ws, 5, i + 2, h, C["abb_dark"], sz=9)

    fps = [a for a in ACTIVIDADES if a[5] == "fp"]
    state_def = {
        "Completado": (C["ok"],   C["ok_txt"]),
        "Pendiente":  (C["warn"], C["warn_txt"]),
        "Riesgo":     (C["risk"], C["risk_txt"]),
        "No aplica":  (C["na"],   "555555"),
    }

    # estado actual estimado (today = 26-MAY-2026)
    hoy = date(2026, 5, 26)
    fp_estados = {
        "FP-P1":   "Pendiente",
        "FP-P2":   "Pendiente",
        "FP-S4-1": "Pendiente",
        "FP-S4-2": "Pendiente",
        "FP-S4-3": "Pendiente",
        "FP-S3-1": "Pendiente",
        "FP-S3-2": "Pendiente",
        "FP-S3-3": "Pendiente",
    }
    impacto = {
        "FP-P1":   "Arquitectura 800XA + lista señales (30% reserva) + protocolo Modbus TCP/IP no se pueden cerrar → FAT Maqueta Sep-Oct en riesgo. PMS tiene holgura amplia (222d vs P6) pero FP-P1 tardío comprime FAT Tableros",
        "FP-P2":   "Retrasa inicio fabricación tablero de control PMS. Holgura absorbida",
        "FP-S4-1": "Demora OC celdas MT Turquía (lead time 6m) → FAT S#4 tarde → Despacho S#4 tarde",
        "FP-S4-2": "Retrasa fabricación tableros auxiliares. Impacto en montaje Shelter S#4",
        "FP-S4-3": "Cierra ingeniería S#4. Demora → montaje e integración Shelter late → FAT S#4 tarde",
        "FP-S3-1": "Demora OC ductos barras Turquía (lead 6m) + VFDs BT Finlandia → FAT S#3 tardío ★ CRÍTICO",
        "FP-S3-2": "Retrasa fabricación 37 CCMs Brasil. Demora en ensayos → FAT S#3 tarde → Despacho late",
        "FP-S3-3": "★ ÚLTIMO FP: demora directa sobre montaje Shelter S#3 → FAT tarde → ENTREGA TARDE → PENALIDAD",
    }

    for idx, (grupo, id_act, nombre, ini, fin, tipo, critica, oc, holgura, nota) in enumerate(fps):
        r = 6 + idx
        ws.row_dimensions[r].height = 28
        estado = fp_estados.get(id_act, "Pendiente")
        bg_e, fc_e = state_def.get(estado, (C["na"], "555555"))
        bg_row = C["s3_light"] if critica else (C["s4_light"] if "S4" in id_act else C["pms_light"])

        dat(ws, r, 2, idx + 1,      bg=bg_row, bold=True, sz=10, color=C["abb_dark"])
        dat(ws, r, 3, id_act,       bg=bg_row, bold=True, sz=9,  color=C["abb_dark"] if critica else C["hdr_mid"])
        dat(ws, r, 4, nombre,       bg=bg_row, bold=critica, sz=9, h="left", wrap=True)
        dat(ws, r, 5, fmt_date(fin), bg=bg_row, bold=True, sz=9,
            color=C["abb_dark"] if critica else "000000")
        dat(ws, r, 6, "—",          bg=C["gray"], sz=9, color="999999")   # Fecha Real (vacía para completar)
        dat(ws, r, 7, "—",          bg=C["gray"], sz=9, color="999999")   # Días δ
        hdr(ws, r, 8, estado, bg=bg_e, fg=fc_e, bold=True, sz=9)
        dat(ws, r, 9, impacto.get(id_act, ""), bg=bg_row, sz=8, h="left", wrap=True,
            color=(C["risk_txt"] if critica else C["warn_txt"]))
        dat(ws, r, 10, nota, bg=bg_row, sz=7, h="left", wrap=True, italic=True)

    thin_border(ws, 2, 2, 6 + len(fps) - 1, 10)

    # Nota crítica
    nr = 6 + len(fps) + 1
    ws.merge_cells(f"B{nr}:I{nr}")
    ws.row_dimensions[nr].height = 22
    hdr(ws, nr, 2,
        "★ ACCIÓN INMEDIATA: AESA debe confirmar antes del 05-JUN-2026 el cambio CCM7 (3×22kW → 2×45kW+VFD) — impacta IB y FP-S3-1",
        C["risk"], sz=9, h="left")

    return ws


# ─────────────────────────────────────────────────────────────────────────────
# HOJA 3: CAMINO CRÍTICO
# ─────────────────────────────────────────────────────────────────────────────
def sheet_critico(wb):
    ws = wb.create_sheet("3_Camino_Critico")
    ws.sheet_view.showGridLines = False

    cols = {"A": 4, "B": 6, "C": 6, "D": 42, "E": 14, "F": 14,
            "G": 12, "H": 18, "I": 28, "J": 4}
    for k, v in cols.items():
        ws.column_dimensions[k].width = v

    ws.merge_cells("B1:I1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:I2"); ws.row_dimensions[2].height = 28
    hdr(ws, 2, 2,
        "CAMINO CRÍTICO — PROVISIÓN ABB  |  CPF-2 / La Calera II",
        C["abb_red"], sz=14)
    ws.merge_cells("B3:I3"); ws.row_dimensions[3].height = 18
    hdr(ws, 3, 2,
        "Sala #3 (BT) determina el despacho más tardío: 14-MAY-2027. Toda demora en S#3 impacta el proyecto sin holgura",
        C["abb_dark"], sz=10)
    ws.row_dimensions[4].height = 8

    # ── Tabla camino crítico ──────────────────────────────────────────────────
    hdrs = ["#", "ID", "Actividad", "Inicio Plan", "Fin Plan", "Duración", "Tipo", "Consecuencia si demora"]
    ws.row_dimensions[5].height = 22
    for i, h in enumerate(hdrs):
        hdr(ws, 5, i + 2, h, C["abb_dark"], sz=9)

    criticas = [(g, id_, n, ini, fin, t, cr, oc, hol, nota)
                for (g, id_, n, ini, fin, t, cr, oc, hol, nota) in ACTIVIDADES if cr]

    consecuencias = {
        "HC-01": "Sin OC no hay proyecto. Completado ✅",
        "HC-02": "Sin cronograma preliminar no hay baseline. Completado ✅",
        "HC-03": "Sin KOM no comienza ingeniería. Completado ✅",
        "HC-04": "Sin KOM PMS no comienza ingeniería PMS. Completado ✅",
        "HC-05": "Sin cronograma final no hay hitos certificables ni control contractual",
        "HC-06": "Retraso en primer informe genera desconfianza del cliente",
        "S3-01": "Demora en IB → FP-S3-1 tarde → OC ductos barras tarde → FAT tarde → Entrega tarde",
        "S3-02": "Demora en aprobación AESA → FP-S3-1 tarde. Dependiente de AESA (ductos barras ⚠)",
        "FP-S3-1": "1 día de demora → OC Turquía tarde → ductos llegan tarde → FAT S#3 tarde → PENALIDAD",
        "S3-03": "Lead time 6 meses. Si se retrasa la OC en julio, los ductos llegan en Feb-2027 → montaje late",
        "S3-04": "VFDs BT lead time 4-5 meses. Retraso en OC → llegan tarde → integración en Shelter tarde",
        "S3-05": "Demora ID → FP-S3-2 tarde → fabricación 37 CCMs en Brasil tarde → montaje Shelter late",
        "S3-WARN": "AESA no confirma CCM7 → ABB no puede cerrar ID S3. VFDs DP-VFD-24210 A/B/C (SE#3) y DP-VFD-23310 A/B/C (SE#4) en scope PMS vía PROFINET — la decisión afecta señales del sistema",
        "S3-06": "Aprobación AESA: 10 días hábiles. Comentarios múltiples pueden demorar FP-S3-2",
        "FP-S3-2": "1 día late → OC fabricación 37 CCMs en Brasil tarde → ensayos tarde → FAT S#3 tarde",
        "S3-07": "Planos constructivos definen montaje Shelter. Errores → retrabajo en Mendoza",
        "S3-08": "AESA 10 días hábiles. Comentarios en IC generan ciclos adicionales",
        "FP-S3-3": "★ HITO MÁS CRÍTICO: 1 semana late → montaje Shelter S#3 late → FAT 1 semana late → despacho 1 semana late → PENALIDAD 1%/semana",
        "S3-09": "Fabricación 37 CCMs Brasil: ~7 semanas. Retraso → ensayos tardan → todo late",
        "S3-10": "Ensayos 20 días. Si falla algún tablero → retrabajo → FAT S#3 se corre",
        "S3-11": "Acopio trigger del pago 25%. Retraso en llegada materiales → pago demorado + montaje late",
        "S3-12": "Montaje Shelter: 8 semanas. Retraso → FAT S#3 se corre → Despacho late → PENALIDAD",
        "S3-13": "FAT S#3 es el más tardío (19-ABR). 35d de holgura vs P6. Cada semana que se corre reduce margen",
        "S3-14": "★ DESPACHO MÁS TARDÍO: si sale después del 14-MAY → penalidad. 45d holgura hasta montaje campo",
        "S3-15": "FECHA LÍMITE CONTRACTUAL. Superarla activa penalidad 1%/semana, tope 10%",
        "INT-02": "iFAT: holgura de solo 3 días. Si ABB S#3 llega tarde al iFAT, impacta comisionado",
        "INT-04": "Comisionado: si se corre la fecha, impacta RFSU directo",
        "INT-05": "RFSU: objetivo firme 31-ENE-2028. Alineado P6 y ABB. No tiene flotante conocido",
    }

    for idx, (grupo, id_act, nombre, ini, fin, tipo, critica, oc, holgura, nota) in enumerate(criticas):
        r = 6 + idx
        ws.row_dimensions[r].height = 28
        dur = (fin - ini).days
        es_fp = tipo == "fp"
        es_hito_c = tipo == "hito_contrato"
        es_warn = "WARN" in id_act

        bg_r = C["risk"] if es_warn else (C["s3_light"] if not es_hito_c else "FFF0F0")
        if es_fp:
            bg_r = "FFE0E0"

        dat(ws, r, 2, idx + 1, bg=bg_r, bold=True, sz=10, color=C["abb_dark"])
        dat(ws, r, 3, id_act, bg=bg_r, bold=True, sz=8, color=C["abb_dark"])
        dat(ws, r, 4, nombre, bg=bg_r, bold=(es_fp or es_hito_c), sz=9, h="left", wrap=True)
        dat(ws, r, 5, fmt_date(ini), bg=bg_r, sz=8)
        dat(ws, r, 6, fmt_date(fin), bg=bg_r, bold=True, sz=8,
            color=C["abb_dark"] if es_fp or es_hito_c else "000000")
        dat(ws, r, 7, "◆" if es_fp else ("★" if es_hito_c else f"{dur}d"),
            bg=bg_r, bold=es_fp or es_hito_c, sz=9,
            color=C["abb_red"] if es_fp or es_hito_c else "000000")
        tipo_str = {"fp":"Freezing Point","fat":"FAT","despacho":"Despacho",
                    "actividad":"Actividad","hito_contrato":"Hito contractual",
                    "integracion":"Integración"}.get(tipo, tipo)
        dat(ws, r, 8, tipo_str, bg=bg_r, sz=8)
        dat(ws, r, 9, consecuencias.get(id_act, nota),
            bg=bg_r if not es_warn else C["risk"],
            sz=8, h="left", wrap=True,
            color=C["risk_txt"] if (es_warn or es_fp) else "000000")

    thin_border(ws, 2, 2, 6 + len(criticas) - 1, 9)

    # ── Resumen holguras ──────────────────────────────────────────────────────
    sr = 6 + len(criticas) + 2
    ws.merge_cells(f"B{sr}:I{sr}")
    ws.row_dimensions[sr].height = 22
    hdr(ws, sr, 2, "HOLGURAS DISPONIBLES vs CRONOGRAMA P6 (COMERCIAL)", C["hdr_dark"], sz=11)

    holguras = [
        ("PMS — FAT (23-OCT-2026)",         "P6 necesita PMS el 02-JUN-2027",       "+222 días", C["ok"],   C["ok_txt"]),
        ("FAT Sala #3 (19-ABR-2027)",        "P6 cierra FDE1103 el 24-MAY-2027",      "+35 días",  C["ok"],   C["ok_txt"]),
        ("Despacho Sala #3 (14-MAY-2027)",   "P6 inicio montaje campo 28-JUN-2027",  "+45 días",  C["ok"],   C["ok_txt"]),
        ("Despacho Sala #4 (08-ABR-2027)",   "P6 inicio montaje campo 28-JUN-2027",  "+81 días",  C["ok"],   C["ok_txt"]),
        ("iFAT Integración (25-JUN-2027)",   "P6 TA1124 = 28-JUN-2027",              "+3 días",   C["risk"], C["risk_txt"]),
        ("FDE1104 cierre S#4 (23-JUN-2027)", "Inicio montaje campo 28-JUN-2027",     "+5 días en P6", C["risk"], C["risk_txt"]),
        ("RFSU (31-ENE-2028)",               "IDÉNTICO en P6 y ABB",                 "= 0 días",  "E2EFDA",  C["ok_txt"]),
        ("Freezing Points (Jun-Ago 2026)",   "SIN hito equivalente en P6",           "SIN CONTROL P6", C["warn"], C["warn_txt"]),
    ]

    ws.row_dimensions[sr + 1].height = 18
    for i, h in enumerate(["Provisión ABB", "Referencia P6", "Holgura", "Estado"], 2):
        hdr(ws, sr + 1, i + 1, h, C["hdr_mid"], sz=9)

    for i, (prov, ref, hol, bg, fc) in enumerate(holguras):
        r = sr + 2 + i
        ws.row_dimensions[r].height = 20
        dat(ws, r, 3, prov, bg=bg, bold=True, sz=9, h="left", color=fc)
        dat(ws, r, 4, ref, bg=bg, sz=9, h="left")
        hdr(ws, r, 5, hol, bg=bg, fg=fc, bold=True, sz=9)
        st = "✅ OK" if bg == C["ok"] else ("⚠ AJUSTADO" if bg == C["warn"] else "🔴 SIN MARGEN")
        hdr(ws, r, 6, st, bg=bg, fg=fc, sz=9)

    thin_border(ws, sr, 2, sr + 2 + len(holguras), 6)
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# HOJA 4: REGISTRO HITOS — planilla de seguimiento semanal
# ─────────────────────────────────────────────────────────────────────────────
def sheet_hitos(wb):
    ws = wb.create_sheet("4_Seguimiento_Hitos")
    ws.sheet_view.showGridLines = False

    cols = {"A": 4, "B": 5, "C": 38, "D": 8, "E": 14, "F": 14,
            "G": 14, "H": 12, "I": 16, "J": 28, "K": 4}
    for k, v in cols.items():
        ws.column_dimensions[k].width = v

    ws.merge_cells("B1:J1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:J2"); ws.row_dimensions[2].height = 28
    hdr(ws, 2, 2, "REGISTRO DE HITOS — PLANILLA DE SEGUIMIENTO SEMANAL", C["hdr_dark"], sz=14)
    ws.merge_cells("B3:J3"); ws.row_dimensions[3].height = 18
    hdr(ws, 3, 2,
        "Completar las columnas en AZUL en cada reunión de seguimiento semanal",
        C["hdr_mid"], sz=10)
    ws.row_dimensions[4].height = 8

    hdrs5 = ["ID", "Hito / Entregable", "OC", "Fecha Plan",
             "Fecha Real", "Estado", "% Avance", "Desvío (d)", "Acción / Observación"]
    ws.row_dimensions[5].height = 20
    for i, h in enumerate(hdrs5):
        hdr(ws, 5, i + 2, h, C["hdr_dark"], sz=9)

    # Seleccionar hitos clave para seguimiento
    hitos_seguimiento = [
        # hitos contractuales + FPs + FATs + despachos + pagos
        a for a in ACTIVIDADES
        if a[5] in ("fp", "fat", "despacho", "hito_contrato", "pago")
        and "WARN" not in a[1]
    ]

    # estados iniciales (hoy = 26-MAY-2026)
    completados = {"HC-01", "HC-02", "HC-03", "HC-04"}

    for idx, (grupo, id_act, nombre, ini, fin, tipo, critica, oc, holgura, nota) in enumerate(hitos_seguimiento):
        r = 6 + idx
        ws.row_dimensions[r].height = 22
        estado = "Completado" if id_act in completados else "Pendiente"
        bg_e = C["ok"] if estado == "Completado" else C["warn"]
        fc_e = C["ok_txt"] if estado == "Completado" else C["warn_txt"]
        bg_r = C["s3_light"] if critica else (C["pms_light"] if "PMS" in grupo else
               (C["s4_light"] if "S4" in grupo else
                (C["pay_light"] if "PAGO" in grupo else "FFF8F8")))
        if tipo == "fp":
            bg_r = C["s3_light"] if critica else C["s4_light"]

        dat(ws, r, 2, id_act,  bg=bg_r, bold=True, sz=8, color=C["abb_dark"] if critica else C["hdr_mid"])
        dat(ws, r, 3, nombre,  bg=bg_r, bold=critica or tipo == "fp", sz=9, h="left", wrap=True,
            color=C["abb_dark"] if critica else "000000")
        dat(ws, r, 4, oc,      bg=bg_r, sz=8)
        dat(ws, r, 5, fmt_date(fin), bg=bg_r, bold=True, sz=8,
            color=C["abb_dark"] if critica else "000000")
        # columnas a completar (azul claro)
        dat(ws, r, 6, fmt_date(fin) if id_act in completados else "—",
            bg="DEEAF1" if id_act not in completados else C["ok"], sz=8)
        hdr(ws, r, 7, estado, bg=bg_e, fg=fc_e, bold=True, sz=8)
        dat(ws, r, 8, "100%" if id_act in completados else "0%",
            bg="DEEAF1" if id_act not in completados else C["ok"], sz=8, bold=True)
        dat(ws, r, 9, "0" if id_act in completados else "—",
            bg="DEEAF1" if id_act not in completados else C["ok"], sz=8)
        dat(ws, r, 10, nota, bg=bg_r, sz=7, h="left", wrap=True, italic=True)

    thin_border(ws, 2, 2, 6 + len(hitos_seguimiento) - 1, 10)

    # nota
    nr = 6 + len(hitos_seguimiento) + 1
    ws.merge_cells(f"B{nr}:J{nr}")
    ws.row_dimensions[nr].height = 22
    hdr(ws, nr, 2,
        "INSTRUCCIONES: Fecha Real = fecha real de ocurrencia | Estado: Completado / Pendiente / En proceso / En riesgo | Desvío = Fecha Real - Fecha Plan (días)",
        C["hdr_mid"], sz=8, h="left")
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# HOJA 5: ALERTAS Y RIESGOS
# ─────────────────────────────────────────────────────────────────────────────
def sheet_alertas(wb):
    ws = wb.create_sheet("5_Alertas_Riesgos")
    ws.sheet_view.showGridLines = False

    cols = {"A": 4, "B": 5, "C": 38, "D": 14, "E": 18, "F": 20,
            "G": 22, "H": 20, "I": 4}
    for k, v in cols.items():
        ws.column_dimensions[k].width = v

    ws.merge_cells("B1:H1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:H2"); ws.row_dimensions[2].height = 28
    hdr(ws, 2, 2, "ALERTAS Y RIESGOS ACTIVOS — CPF-2 / La Calera II — ABB", C["abb_red"], sz=14)
    ws.merge_cells("B3:H3"); ws.row_dimensions[3].height = 18
    hdr(ws, 3, 2, f"Estado al: 26-MAY-2026  |  Ref: R-2631036 / E-2611027", C["abb_dark"], sz=10)
    ws.row_dimensions[4].height = 8

    hdrs5 = ["#", "Alerta / Riesgo", "Responsable", "Fecha límite", "Impacto", "Mitigación", "Estado"]
    ws.row_dimensions[5].height = 22
    for i, h in enumerate(hdrs5):
        hdr(ws, 5, i + 2, h, C["abb_dark"], sz=9)

    alertas = [
        # (severidad, alerta, responsable, fecha_limite, impacto, mitigacion, estado)
        ("🔴 CRÍTICO",
         "Confirmar cambio CCM7: 3×22kW → 2×45kW+VFD en S#3",
         "AESA (Dario Stirparo)",
         "05-JUN-2026",
         "Si no se confirma antes del 05-JUN: IB de S#3 no puede cerrarse → FP-S3-1 se corre → OC Turquía tarde → despacho S#3 tarde",
         "AESA contacting client. ABB debe hacer follow-up semanal",
         "Pendiente ⚠"),

        ("🔴 CRÍTICO",
         "Cronograma final con Hold Points comprometido para SEM 01-JUN",
         "ABB (Rodrigo Mack)",
         "01-JUN-2026",
         "Sin cronograma final no hay baseline contractual. No se pueden planificar hitos de pago ni Pólizas",
         "ABB debe emitir cronograma con OCs a 3ros confirmadas",
         "Pendiente ⚠"),

        ("🟡 IMPORTANTE",
         "Acceso a portal ShareFile (Gustavo Pantolini / Juan Palma / Rodrigo Mack)",
         "ABB (interno)",
         "SEM 26-MAY",
         "Sin acceso al portal no se pueden emitir documentos vía transmittal oficial",
         "ABB gestionando habilitación interna de acceso",
         "En proceso"),

        ("🟡 IMPORTANTE",
         "Información de ductos de barra: posición transfo respecto a Salas",
         "AESA (con cliente Pluspetrol)",
         "Antes del FP-S4-1 (26-JUN)",
         "Sin esta información ABB no puede finalizar diseño ductos de barras (Turquía, lead 6 meses)",
         "AESA en contacto con el cliente. ABB debe confirmar fecha de respuesta",
         "Pendiente ⚠"),

        ("🟡 IMPORTANTE",
         "Confirmación unifilares tableros CCMs (posibles cambios de cargas)",
         "AESA (Tulio De La Torre)",
         "Antes del FP-S3-1 (03-JUL)",
         "Cambio de cargas en CCMs impacta el diseño BT. Si se confirma después del FP, genera Nota de Cambio",
         "Tema en revisión por AESA con cliente. Seguimiento en reunión técnica quincenal",
         "Pendiente ⚠"),

        ("🟡 IMPORTANTE",
         "Pólizas de Caución: Avance Fabricación (requerida antes del Hito 10% IB)",
         "ABB (Dpto Legal/Seguros)",
         "Antes de PAY-01 (12-JUN)",
         "Sin Póliza aprobada por AESA: no se puede certificar el Hito 10% IB → retraso en cobro",
         "ABB debe iniciar gestión de pólizas con aseguradora. Póliza debe estar aprobada por AESA Seguros",
         "Pendiente"),

        ("🟠 ATENCIÓN",
         "Freezing Points NO representados en P6 comercial del proyecto",
         "ABB + AESA (PM)",
         "SEM 01-JUN (próxima reunión)",
         "Desvíos en FPs no son visibles en el P6 hasta que impactan el despacho. Sin control temprano",
         "Solicitar a AESA incorporar FP-IC S#3 (31-AGO), FP-IC S#4 (24-AGO) y FATs como hitos P6",
         "Pendiente ⚠"),

        ("🟠 ATENCIÓN",
         "FDE1104 (S#4 en P6): margen de solo 5 días entre cierre P6 e inicio montaje campo",
         "ABB (seguimiento S#4) + AESA PM",
         "Continuo",
         "Si S#4 se retrasa vs ABB plan, el montaje en campo no tiene flotante en P6",
         "Monitorear S#4 mensualmente. S#4 tiene 81 días de holgura desde despacho ABB hasta montaje P6",
         "Bajo control — monitorear"),

        ("🟠 ATENCIÓN",
         "iFAT Integración: holgura de solo 3 días vs P6 (ABB 25-JUN vs P6 28-JUN-2027)",
         "ABB + EPC",
         "JUN-2027",
         "Si algún vendor llega tarde al iFAT (ABB, INAUCO, HIMA), el comisionado se corre",
         "Incluir en cronograma final los hitos de todos los vendors para el iFAT",
         "Futuro — bajo seguimiento"),

        ("🟢 INFO",
         "Prueba integrada PMS en Shelter: solicitada por AESA, factibilidad pendiente",
         "ABB (Pablo Kalis)",
         "SEM 08-JUN",
         "Si se acepta prueba en Shelter: puede impactar timing despacho PMS. Si no: FAT PMS queda en BsAs",
         "ABB evalúa si la prueba integrada se realiza antes del despacho. Aclarar en cronograma final",
         "En análisis"),

        ("🟡 IMPORTANTE",
         "AESA obligatorio en FAT Maqueta PMS — Sep/Oct 2026 — planificar viaje a Buenos Aires",
         "AESA (Equipo Técnico / TDJ-PLC)",
         "AGO-2026 (con 4 semanas de anticipación)",
         "La ET ACAL-00102-ET-E-0005 sec.20 establece participación obligatoria de AESA en Pruebas Tempranas (mock-up PMS+CCM+AC800M). Sin AESA presente no se certifica el resultado de la prueba",
         "Planificar viaje a BsAs para SEP-OCT 2026. Confirmar disponibilidad con ABB (Pablo Kalis) al menos 4 semanas antes del inicio",
         "Pendiente — planificar"),

        ("🟢 INFO",
         "Protocolo PMS cambiado: Ethernet TCP/IP → Modbus TCP/IP (acuerdo Oct 2025)",
         "AESA + ABB + INAUCO",
         "Vigente desde Oct-2025",
         "Acordado en reunión AESA-INAUCO-ABB del 02-OCT-2025. Incorporado en ET Rev.0 (25-NOV-2025) sec.23. Afecta diseño de red FP-P1 y comunicación PLC-PMS con SCADA de Procesos",
         "Verificar que ABB refleja el cambio en la IB PMS (P-01). Confirmar protocolo definitivo en revisión FP-P1",
         "Acordado ✅ — verificar en IB"),
    ]

    sev_bg = {
        "🔴 CRÍTICO":   C["risk"],
        "🟡 IMPORTANTE": C["warn"],
        "🟠 ATENCIÓN":  "FFE0B2",
        "🟢 INFO":      C["ok"],
    }
    sev_fc = {
        "🔴 CRÍTICO":   C["risk_txt"],
        "🟡 IMPORTANTE": C["warn_txt"],
        "🟠 ATENCIÓN":  "7B3F00",
        "🟢 INFO":      C["ok_txt"],
    }

    for idx, (sev, alerta, resp, fecha, impacto, mitigacion, estado) in enumerate(alertas):
        r = 6 + idx
        ws.row_dimensions[r].height = 50
        bg_s = sev_bg.get(sev, C["warn"])
        fc_s = sev_fc.get(sev, C["warn_txt"])
        hdr(ws, r, 2, sev, bg=bg_s, fg=fc_s, bold=True, sz=9)
        dat(ws, r, 3, alerta, bg=bg_s, bold=True, sz=9, h="left", wrap=True, color=fc_s)
        dat(ws, r, 4, resp, bg=bg_s, sz=8, h="left", wrap=True)
        dat(ws, r, 5, fecha, bg=bg_s, bold=True, sz=9, color=C["abb_dark"])
        dat(ws, r, 6, impacto, bg=bg_s, sz=8, h="left", wrap=True, italic=True)
        dat(ws, r, 7, mitigacion, bg=bg_s, sz=8, h="left", wrap=True)
        dat(ws, r, 8, estado, bg=bg_s, bold=True, sz=8, color=fc_s)

    thin_border(ws, 2, 2, 6 + len(alertas) - 1, 8)
    return ws


# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────
def main():
    wb = Workbook()
    wb.remove(wb.active)

    sheet_gantt(wb)
    sheet_freezing(wb)
    sheet_critico(wb)
    sheet_hitos(wb)
    sheet_alertas(wb)

    wb.save(OUTPUT)
    print(f"Generado: {OUTPUT}")
    print(f"Actividades totales: {len(ACTIVIDADES)}")
    print(f"Freezing Points: {sum(1 for a in ACTIVIDADES if a[5]=='fp')}")
    print(f"Actividades cadena crítica: {sum(1 for a in ACTIVIDADES if a[6])}")

if __name__ == "__main__":
    main()
