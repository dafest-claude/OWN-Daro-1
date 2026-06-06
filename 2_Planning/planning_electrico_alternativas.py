"""
PLANNING ELÉCTRICO CPF-2 — DOS ALTERNATIVAS
Gantt comparativo: Alternativa 1 (Pre-tendido) vs Alternativa 2 (Post-arribo salas)
Fecha: 06-JUN-2026
Fuente: ACAL-00102-LC-E-0001 (613 cables, 81,231m CPF-2) | Cronograma ABB RevA
"""
from datetime import date, timedelta
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import os

OUTPUT = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                      "Planning_Electrico_CPF2_Alternativas.xlsx")

# ─── PALETTE ──────────────────────────────────────────────────────────────────
C = {
    "hdr":    "1F3864", "hdr2":   "2E75B6", "sep":    "404040",
    "civil":  "A9D18E", "tendido":"5B9BD5", "conex":  "7030A0",
    "testing":"00B0F0", "comis":  "375623", "energia":"FF9900",
    "hito":   "CC0000", "sala":   "E36C09", "sat":    "FF0000",
    "rfmc":   "C00000", "rfsu":   "880000",
    "pre":    "BDD7EE", "post":   "FCE4D6",
    "ok":     "C6EFCE", "warn":   "FFEB9C", "rsk":    "FFC7CE",
    "ok_t":   "375623", "warn_t": "9C5700", "rsk_t":  "9C0006",
    "white":  "FFFFFF", "gray":   "F2F2F2", "lgray":  "D9D9D9",
    "alt1_bg":"EBF3FB", "alt2_bg":"FFF3E2",
    "crit":   "FF0000",
}

MESES = ["","Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"]

def fill(c):   return PatternFill("solid", fgColor=c)
def bd(color="CCCCCC", style="thin"):
    s = Side(style=style, color=color)
    return Border(left=s, right=s, top=s, bottom=s)
def fmt(d):    return d.strftime("%d/%m/%y") if d else "—"

def cel(ws, r, c, v="", bg=None, bold=False, sz=9, h="center", col="000000",
        italic=False, wrap=False, border=True):
    x = ws.cell(row=r, column=c, value=v)
    x.font = Font(name="Calibri", bold=bold, size=sz, color=col, italic=italic)
    if bg: x.fill = fill(bg)
    x.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    if border: x.border = bd()
    return x

def hcl(ws, r, c, v, bg, fg="FFFFFF", sz=9, bold=True, h="center", wrap=False):
    x = ws.cell(row=r, column=c, value=v)
    x.fill = fill(bg); x.font = Font(name="Calibri", bold=bold, size=sz, color=fg)
    x.alignment = Alignment(horizontal=h, vertical="center", wrap_text=wrap)
    x.border = bd()
    return x

# ─── SEMANAS ──────────────────────────────────────────────────────────────────
INICIO_GANTT = date(2027, 1, 4)   # primer lunes ENE-2027
FIN_GANTT    = date(2028, 2, 25)  # fin FEB-2028
SEMANAS = []
d = INICIO_GANTT
while d <= FIN_GANTT:
    SEMANAS.append(d)
    d += timedelta(weeks=1)

N_INFO  = 7     # columnas de info: Fase, Actividad, Resp, Inicio, Fin, Días, Notas
COL_OFF = N_INFO + 1

def sem_cols(ini, fin):
    """Devuelve (ci, cf) columnas Gantt para rango de fechas."""
    ci = cf = None
    for i, s in enumerate(SEMANAS):
        sw = s + timedelta(days=6)
        if ci is None and sw >= ini: ci = i
        if s <= fin: cf = i
    return ci, cf

def draw_bar(ws, row, ini, fin, color, marker="", bg_off="F5F5F5"):
    ci, cf = sem_cols(ini, fin)
    for i, s in enumerate(SEMANAS):
        c = ws.cell(row=row, column=COL_OFF + i)
        izq = Side(style="medium" if s.day <= 7 else "thin",
                   color="888888" if s.day <= 7 else "E0E0E0")
        c.border = Border(left=izq, right=Side(style="thin", color="E0E0E0"),
                          top=Side(style="thin", color="E0E0E0"),
                          bottom=Side(style="thin", color="E0E0E0"))
        if ci is not None and cf is not None and ci <= i <= cf:
            c.fill = fill(color)
            if marker and i == ci:
                c.value = marker
                c.font = Font(bold=True, color="FFFFFF", size=7)
                c.alignment = Alignment(horizontal="center", vertical="center")
        else:
            c.fill = fill("F8F8F8" if i % 2 == 0 else "F2F2F2")

# ─── ACTIVIDADES ──────────────────────────────────────────────────────────────
# (fase, nombre, responsable, inicio, fin, tipo, critica, nota)
# tipo → color_barra
TIPO_COLOR = {
    "civil":    "A9D18E",
    "tendido":  "5B9BD5",
    "conex":    "7030A0",
    "testing":  "00B0F0",
    "energia":  "FF9900",
    "comis":    "375623",
    "hito":     "CC0000",
    "sala":     "E36C09",
    "sat":      "C00000",
    "rfmc":     "880000",
    "rfsu":     "880000",
    "sep":      None,
}

# ──────────────────────────────────────────────────────────────────────────────
#  ALTERNATIVA 1 — PRE-TENDIDO
#  Todos los cables tendidos ANTES del arribo de salas.
#  Cuando llegan las salas: 2 semanas armado → inicio conexionado.
# ──────────────────────────────────────────────────────────────────────────────
ACT1 = [
    # (fase, actividad, resp, ini, fin, tipo, critica, nota)

    # ── PRE-ARRIBO (trabajos civiles y tendido anticipado) ────────────────────
    ("PRE-ARRIBO", None, None, None, None, "sep", False, ""),

    ("Pre-Arribo","Preparación bandejas portacables principales",
     "EPC", date(2027,1,4), date(2027,1,22), "civil", False,
     "Montaje soportería + bandejas antes del tendido. Condición previa."),

    ("Pre-Arribo","Tendido cables potencia MT (18 cables, 2.110m)",
     "AESA/EPC", date(2027,1,25), date(2027,2,14), "tendido", False,
     "6.6kV/13.2kV | 3 cuadrillas | 60m/día → 12 días paralelo"),

    ("Pre-Arribo","Tendido cables potencia BT grande (104 cables, 17.470m) ← RUTA CRÍTICA",
     "AESA/EPC", date(2027,1,25), date(2027,4,4), "tendido", True,
     "3x95 a 3x240mm² | 3 cuadrillas | 50m/día → 79 días paralelo | ¡ARRANCAR 25-ENE!"),

    ("Pre-Arribo","Tendido cables potencia BT mediano/pequeño (198 cables, 33.125m)",
     "AESA/EPC", date(2027,2,1), date(2027,3,28), "tendido", False,
     "3x6 a 3x70mm² | 3 cuadrillas | 80m/día → 45 días paralelo"),

    ("Pre-Arribo","Tendido cables control multiconductor (100 cables, 12.000m)",
     "AESA/EPC", date(2027,2,15), date(2027,4,4), "tendido", False,
     "12x2,5 y similares | 3 cuadrillas | 100m/día → 40 días paralelo"),

    ("Pre-Arribo","Tendido cables señal BT y auxiliares (92 cables, 16.126m)",
     "AESA/EPC", date(2027,3,1), date(2027,4,11), "tendido", False,
     "2x2,5+T, 2x4 | 3 cuadrillas | 100m/día → 35 días paralelo"),

    ("Pre-Arribo","Tendido FO (400m) y Ethernet UTP CPF-2",
     "ABB/EPC", date(2027,3,15), date(2027,3,26), "tendido", False,
     "FO: 200m/día | UTP: 150m/día → 5 días"),

    # ── ARRIBO Y ARMADO SALAS ─────────────────────────────────────────────────
    ("LLEGADA SALAS", None, None, None, None, "sep", False, ""),

    ("Sala #4","◆ LLEGADA Sala Eléctrica #4 (MT) — HITO CONTRACTUAL",
     "ABB", date(2027,4,8), date(2027,4,8), "hito", True,
     "OC 4508944973 | deadline 08-ABR-2027 | USD 3.69M"),

    ("Sala #4","Armado y habilitación Sala #4 en sitio (2 semanas)",
     "ABB+EPC", date(2027,4,9), date(2027,4,22), "sala", True,
     "Posicionamiento, fijación, conexión interna, verificación"),

    ("Sala #3","◆ LLEGADA Sala Eléctrica #3 (BT) ← CADENA CRÍTICA ABB",
     "ABB", date(2027,5,14), date(2027,5,14), "hito", True,
     "OC 4508944971 | deadline 14-MAY-2027 | USD 8.13M | HITO MÁS TARDÍO"),

    ("Sala #3","Armado y habilitación Sala #3 en sitio (2 semanas)",
     "ABB+EPC", date(2027,5,15), date(2027,5,28), "sala", True,
     "37 CCMs 4000A 65kA | 3 ductos de barras | PMS-001 integrado"),

    # ── CONEXIONADO Y TERMINACIÓN ─────────────────────────────────────────────
    ("CONEXIONADO", None, None, None, None, "sep", False, ""),

    ("Conexionado","Conexionado cables MT (18 cables × 2 extrem. × 8hs = 36 días)",
     "AESA", date(2027,4,23), date(2027,6,13), "conex", True,
     "Personal especialista terminales MT | 1 cuadrilla MT"),

    ("Conexionado","Conexionado BT grande (3x95 a 3x240mm² — 3hs/extremo)",
     "AESA", date(2027,4,23), date(2027,6,6), "conex", False,
     "104 cables × 2 extrem. × 3hs / (3cuad × 8hs) → 26 días"),

    ("Conexionado","Conexionado BT mediano/pequeño (2hs / 1hs por extremo)",
     "AESA", date(2027,5,10), date(2027,6,21), "conex", False,
     "198+98 cables → ~30 días paralelo con 3 cuadrillas"),

    ("Conexionado","Conexionado control + señales (4hs/cable — peinado + etiquetado)",
     "AESA", date(2027,5,29), date(2027,7,12), "conex", True,
     "192 cables × 4hs / (3cuad × 8hs) → 32 días. INICIA post armado S3"),

    ("Conexionado","Conexionado FO — fusión / conectorización ABB",
     "ABB", date(2027,5,29), date(2027,6,5), "conex", False,
     "1 cable FO + patch de red | ABB es responsable"),

    # ── PRUEBAS Y ENERGIZACIÓN ────────────────────────────────────────────────
    ("PRUEBAS", None, None, None, None, "sep", False, ""),

    ("Pruebas","Megger e inspección visual — 100% circuitos",
     "AESA", date(2027,7,13), date(2027,8,1), "testing", True,
     "Prueba aislación por circuito previo energización | 20 días"),

    ("Pruebas","iFAT Integración (ABB+INAUCO+HIMA) — integración en sitio",
     "ABB/AESA", date(2027,6,25), date(2027,7,14), "hito", False,
     "Según cronograma EPC P6 | puede ocurrir en paralelo con conexionado final"),

    ("Energización","Energización progresiva MT (13.2kV → 6.6kV)",
     "AESA", date(2027,8,2), date(2027,8,13), "energia", True,
     "Requiere 100% cables MT terminados + megger aprobado | 10 días"),

    ("Energización","Energización progresiva BT — carga por carga",
     "AESA", date(2027,8,14), date(2027,9,3), "energia", True,
     "Secuencia: CCMs → motores → instrumentación | 15 días"),

    # ── RFMC ──────────────────────────────────────────────────────────────────
    ("HITOS EPC", None, None, None, None, "sep", False, ""),

    ("Hito EPC","★ RFMC — Ready For Mechanical Completion (ALT.1)",
     "EPC/AESA", date(2027,9,3), date(2027,9,3), "rfmc", True,
     "✅ FACTIBLE ALT.1: 03-SEP-2027 | margen: +149 días vs RFSU"),

    ("Comisionado","Precomisionado eléctrico + funcional",
     "AESA", date(2027,9,4), date(2027,10,3), "comis", False,
     "Pruebas funcionales por subsistema | 30 días"),

    ("Comisionado","SAT PMS — HOLD POINT (AESA + Pluspetrol per PIE)",
     "ABB/AESA", date(2027,10,1), date(2027,12,15), "sat", True,
     "100% E/S ensayados | HP obligatorio | 75 días (OCT-DIC 2027)"),

    ("Comisionado","Comisionado integrado EPC",
     "EPC", date(2027,7,16), date(2028,1,31), "comis", False,
     "Paralelo a SAT y precomisionado | arranque escalonado"),

    ("Hito EPC","★ RFSU — Ready For Start Up (31-ENE-2028)",
     "PROYECTO", date(2028,1,31), date(2028,1,31), "rfsu", True,
     "✅ COMPATIBLE ALT.1 — margen confortable"),

    ("Post-RFSU","Capacitación PMS — 2 instancias en sitio",
     "ABB", date(2028,2,5), date(2028,3,7), "comis", False,
     "Per ET sec.23 | FEB+MAR 2028"),
]

# ──────────────────────────────────────────────────────────────────────────────
#  ALTERNATIVA 2 — POST-ARRIBO SALAS
#  Todo comienza al arribo de las salas. Sin trabajo previo de tendido.
# ──────────────────────────────────────────────────────────────────────────────
ACT2 = [
    # ── LLEGADA SALAS ─────────────────────────────────────────────────────────
    ("LLEGADA SALAS", None, None, None, None, "sep", False, ""),

    ("Sala #4","◆ LLEGADA Sala Eléctrica #4 (MT) — HITO CONTRACTUAL",
     "ABB", date(2027,4,8), date(2027,4,8), "hito", True,
     "OC 4508944973 | deadline 08-ABR-2027"),

    ("Sala #4","Armado y habilitación Sala #4 (2 semanas)",
     "ABB+EPC", date(2027,4,9), date(2027,4,22), "sala", True,
     "Posicionamiento, fijación, verificación"),

    ("Sala #3","◆ LLEGADA Sala Eléctrica #3 (BT) ← CADENA CRÍTICA ABB",
     "ABB", date(2027,5,14), date(2027,5,14), "hito", True,
     "OC 4508944971 | deadline 14-MAY-2027"),

    ("Sala #3","Armado y habilitación Sala #3 (2 semanas)",
     "ABB+EPC", date(2027,5,15), date(2027,5,28), "sala", True,
     "37 CCMs 4000A 65kA + PMS-001"),

    # ── PRE-TRABAJO MÍNIMO (bandejas pueden ir en paralelo al armado) ─────────
    ("PRE-TRABAJO", None, None, None, None, "sep", False, ""),

    ("Pre-Trabajo","Bandejas portacables (inicia con S4, termina con S3)",
     "EPC", date(2027,4,9), date(2027,5,14), "civil", False,
     "Puede avanzar durante armado salas | 25 días"),

    # ── TENDIDO — TODO POST ARRIBO ────────────────────────────────────────────
    ("TENDIDO", None, None, None, None, "sep", False, ""),

    ("Tendido","Tendido cables potencia MT (18 cables, 2.110m)",
     "AESA/EPC", date(2027,5,15), date(2027,5,30), "tendido", False,
     "3 cuadrillas | 60m/día → 12 días paralelo"),

    ("Tendido","Tendido cables BT grande (104 cables, 17.470m) ← CUELLO DE BOTELLA",
     "AESA/EPC", date(2027,5,29), date(2027,8,23), "tendido", True,
     "3x95 a 3x240mm² | 3 cuadrillas | 50m/día → 79 días | CRÍTICO: termina 23-AGO"),

    ("Tendido","Tendido cables BT mediano/pequeño (198 cables, 33.125m)",
     "AESA/EPC", date(2027,6,12), date(2027,8,7), "tendido", False,
     "3 cuadrillas | 80m/día → 45 días"),

    ("Tendido","Tendido cables control y señales (192 cables, 28.126m)",
     "AESA/EPC", date(2027,7,1), date(2027,8,24), "tendido", False,
     "3 cuadrillas | 100m/día → 38 días"),

    ("Tendido","Tendido FO + Ethernet",
     "ABB/EPC", date(2027,8,10), date(2027,8,18), "tendido", False,
     "5 días"),

    # ── CONEXIONADO ───────────────────────────────────────────────────────────
    ("CONEXIONADO", None, None, None, None, "sep", False, ""),

    ("Conexionado","Conexionado cables MT (36 días — terminalistas MT)",
     "AESA", date(2027,5,31), date(2027,7,21), "conex", True,
     "18 cables × 2 extrem. × 8hs | 1 cuadrilla MT especialista"),

    ("Conexionado","Conexionado BT grande (26 días)",
     "AESA", date(2027,8,24), date(2027,9,26), "conex", True,
     "Inicia post-tendido BT grande | 104 cables → 26 días paralelo"),

    ("Conexionado","Conexionado BT mediano/pequeño (30 días)",
     "AESA", date(2027,8,8), date(2027,9,17), "conex", False,
     "198+98 cables → 30 días con 3 cuadrillas"),

    ("Conexionado","Conexionado control + señales (32 días) ← RUTA CRÍTICA",
     "AESA", date(2027,8,25), date(2027,10,6), "conex", True,
     "192 cables × 4hs | 32 días con 3 cuadrillas | TERMINA 06-OCT"),

    # ── PRUEBAS Y ENERGIZACIÓN ────────────────────────────────────────────────
    ("PRUEBAS", None, None, None, None, "sep", False, ""),

    ("Pruebas","Megger + inspección visual — 100% circuitos",
     "AESA", date(2027,10,7), date(2027,10,30), "testing", True,
     "20 días | INICIA post-conexionado control"),

    ("Pruebas","iFAT Integración (ABB+INAUCO+HIMA)",
     "ABB/AESA", date(2027,6,25), date(2027,7,14), "hito", False,
     "Según P6 | puede adelantarse aunque tendido no esté completo"),

    ("Energización","Energización progresiva MT (10 días)",
     "AESA", date(2027,10,31), date(2027,11,13), "energia", True,
     "⚠ TARDE vs Alt.1 | Post megger aprobado"),

    ("Energización","Energización progresiva BT (15 días)",
     "AESA", date(2027,11,14), date(2027,12,5), "energia", True,
     "⚠ MUY AJUSTADO — 57 días antes RFSU"),

    # ── RFMC ──────────────────────────────────────────────────────────────────
    ("HITOS EPC", None, None, None, None, "sep", False, ""),

    ("Hito EPC","★ RFMC — Ready For Mechanical Completion (ALT.2)",
     "EPC/AESA", date(2027,12,5), date(2027,12,5), "rfmc", True,
     "⚠ RIESGO ALT.2: 05-DIC-2027 | solo 57 días para comisionado"),

    ("Comisionado","SAT PMS — HOLD POINT (HP per PIE) — COMPRIMIDO",
     "ABB/AESA", date(2027,10,1), date(2027,11,30), "sat", True,
     "⚠ SUPERPUESTO con energización — CONFLICTO CRÍTICO"),

    ("Comisionado","Comisionado integrado EPC — RIESGO ALTO",
     "EPC", date(2027,12,6), date(2028,1,31), "comis", True,
     "⚠ Solo 57 días para comisionar completo — MARGEN CERO"),

    ("Hito EPC","★ RFSU — Ready For Start Up (31-ENE-2028) — EN RIESGO",
     "PROYECTO", date(2028,1,31), date(2028,1,31), "rfsu", True,
     "🔴 RIESGO REAL ALT.2 — cualquier demora en tendido o conexionado → RFSU tarde"),

    ("Post-RFSU","Capacitación PMS (en riesgo de posponerse)",
     "ABB", date(2028,2,5), date(2028,3,7), "comis", False,
     "Per ET | FEB-MAR 2028"),
]

# ──────────────────────────────────────────────────────────────────────────────
#  GENERADOR DE HOJA GANTT
# ──────────────────────────────────────────────────────────────────────────────
def make_gantt_sheet(wb, nombre_hoja, actividades, alt_num, bg_titulo):
    ws = wb.create_sheet(nombre_hoja)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = f"{get_column_letter(COL_OFF)}6"

    # anchos columnas info
    anchos = [4, 14, 40, 10, 11, 11, 6, 30]
    for i, w in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    # anchos semanas (2.1 por semana ≈ 1 caracter)
    for i in range(len(SEMANAS)):
        ws.column_dimensions[get_column_letter(COL_OFF + i)].width = 2.1

    last_col = get_column_letter(COL_OFF + len(SEMANAS) - 1)

    # ── TÍTULO ────────────────────────────────────────────────────────────────
    ws.merge_cells(f"A1:{last_col}1")
    ws.row_dimensions[1].height = 36
    alt_txt = "PRE-TENDIDO" if alt_num == 1 else "POST-ARRIBO SALAS"
    status  = "✅ COMPATIBLE CON RFSU 31-ENE-2028" if alt_num == 1 else "⚠ RIESGO — MARGEN MÍNIMO PARA RFSU 31-ENE-2028"
    t = ws.cell(row=1, column=1,
        value=f"ALTERNATIVA {alt_num} — {alt_txt}  |  CPF-2 La Calera II (Vaca Muerta)  |  {status}  |  06-JUN-2026")
    t.fill = fill(bg_titulo)
    t.font = Font(name="Calibri", bold=True, size=12, color="FFFFFF")
    t.alignment = Alignment(horizontal="left", vertical="center")

    # ── LEYENDA ───────────────────────────────────────────────────────────────
    ws.merge_cells(f"A2:{last_col}2")
    ws.row_dimensions[2].height = 13
    ley = ("■ Civil/Bandejas   ■ Tendido cables   ■ Conexionado/Terminación   "
           "■ Pruebas/Megger   ■ Energización   ■ Comisionado   "
           "■ ◆ Hito/Arribo   ■ ★ RFMC/RFSU   ■ SAT PMS (Hold Point)")
    l = ws.cell(row=2, column=1, value=ley)
    l.font = Font(name="Calibri", size=7, color="444444")
    l.alignment = Alignment(horizontal="left", vertical="center")

    # ── CABECERAS AÑO ─────────────────────────────────────────────────────────
    ws.row_dimensions[3].height = 13
    anio_g = {}
    for i, s in enumerate(SEMANAS):
        a = s.year
        if a not in anio_g: anio_g[a] = [i, i]
        else: anio_g[a][1] = i
    for a, (ci, cf) in anio_g.items():
        try: ws.merge_cells(f"{get_column_letter(COL_OFF+ci)}3:{get_column_letter(COL_OFF+cf)}3")
        except: pass
        hcl(ws, 3, COL_OFF+ci, str(a), C["hdr"])

    # ── CABECERAS MES ─────────────────────────────────────────────────────────
    ws.row_dimensions[4].height = 11
    mes_g = {}
    for i, s in enumerate(SEMANAS):
        k = (s.year, s.month)
        if k not in mes_g: mes_g[k] = [i, i]
        else: mes_g[k][1] = i
    for (a, m), (ci, cf) in mes_g.items():
        try: ws.merge_cells(f"{get_column_letter(COL_OFF+ci)}4:{get_column_letter(COL_OFF+cf)}4")
        except: pass
        hcl(ws, 4, COL_OFF+ci, MESES[m], C["hdr2"] if m % 2 == 0 else "404040", sz=7)

    # ── CABECERAS COLUMNAS INFO ───────────────────────────────────────────────
    ws.row_dimensions[5].height = 20
    for c2, t in enumerate(["#","Fase","Actividad","Resp.","Inicio","Fin","Días","Notas / Condición"], 1):
        hcl(ws, 5, c2, t, C["hdr"])

    # ── DATOS ─────────────────────────────────────────────────────────────────
    row = 6
    num = 0
    for (fase, nombre, resp, ini, fin, tipo, critica, nota) in actividades:

        # ── Separador de grupo ────────────────────────────────────────────────
        if tipo == "sep":
            ws.merge_cells(f"A{row}:{last_col}{row}")
            ws.row_dimensions[row].height = 14
            lbl = f"  ▶  {fase}"
            hcl(ws, row, 1, lbl, C["sep"], sz=9, h="left")
            row += 1
            continue

        num += 1
        es_hito = tipo in ("hito", "rfmc", "rfsu")
        es_crit = critica
        dias = (fin - ini).days + 1 if ini and fin else 0

        # fondo de fila según criticidad
        if tipo == "rfsu":
            row_bg = C["rfmc"]
            row_fc = "FFFFFF"
        elif tipo == "rfmc":
            row_bg = "880000"
            row_fc = "FFFFFF"
        elif tipo == "sat":
            row_bg = C["rsk"]
            row_fc = C["rsk_t"]
        elif es_crit:
            row_bg = "FFF0F0"
            row_fc = "880000"
        else:
            row_bg = "FFFFFF"
            row_fc = "000000"

        cel(ws, row, 1, num, bg=row_bg, bold=True, sz=8, col=row_fc)
        cel(ws, row, 2, fase, bg=row_bg, bold=es_hito, sz=8, h="left", col=row_fc)
        cel(ws, row, 3, nombre or "", bg=row_bg, bold=(es_hito or es_crit),
            sz=9, h="left", col=row_fc, wrap=True)
        cel(ws, row, 4, resp or "", bg=row_bg, sz=8, col=row_fc)
        cel(ws, row, 5, fmt(ini), bg=row_bg, bold=es_hito, sz=8, col=row_fc)
        cel(ws, row, 6, fmt(fin), bg=row_bg, bold=es_hito, sz=8, col=row_fc)
        cel(ws, row, 7, dias if not es_hito else "—", bg=row_bg, sz=8, col=row_fc)
        cel(ws, row, 8, nota, bg=row_bg, sz=7, h="left", italic=True, col=row_fc, wrap=True)

        # barra Gantt
        bar_c = TIPO_COLOR.get(tipo, "2E75B6")
        marker = "◆" if tipo == "hito" else ("★" if tipo in ("rfmc","rfsu") else "")
        if ini and fin and bar_c:
            draw_bar(ws, row, ini, fin, bar_c, marker)

        ws.row_dimensions[row].height = 14
        row += 1

    # ── RESUMEN FINAL ─────────────────────────────────────────────────────────
    ws.merge_cells(f"A{row}:{last_col}{row}")
    ws.row_dimensions[row].height = 20
    if alt_num == 1:
        msg = ("✅ ALT.1 PRE-TENDIDO — RFMC: 03-SEP-2027 | RFSU: 31-ENE-2028  |  "
               "Tendido inicia 25-ENE-2027 | 79 días ruta crítica BT grande | "
               "Margen RFSU: ~149 días post-RFMC | RECOMENDADA")
        bg_r = C["ok"]
        fg_r = C["ok_t"]
    else:
        msg = ("⚠ ALT.2 POST-ARRIBO — RFMC: 05-DIC-2027 | RFSU: 31-ENE-2028  |  "
               "Solo 57 días comisionado | SAT PMS superpuesto con energización | "
               "Riesgo real de atraso RFSU | NO RECOMENDADA sin recursos adicionales")
        bg_r = C["rsk"]
        fg_r = C["rsk_t"]
    hcl(ws, row, 1, msg, bg_r, fg=fg_r, sz=9, h="left")

# ──────────────────────────────────────────────────────────────────────────────
#  HOJA DE COMPARACIÓN
# ──────────────────────────────────────────────────────────────────────────────
def make_comparacion(wb):
    ws = wb.create_sheet("0_Comparacion")
    ws.sheet_view.showGridLines = False
    for k, v in {"A":4,"B":30,"C":38,"D":28,"E":4}.items():
        ws.column_dimensions[k].width = v

    ws.merge_cells("B1:D1"); ws.row_dimensions[1].height = 8
    ws.merge_cells("B2:D2"); ws.row_dimensions[2].height = 32
    hcl(ws, 2, 2, "COMPARACIÓN DE ALTERNATIVAS — INSTALACIÓN ELÉCTRICA CPF-2  |  06-JUN-2026",
        C["hdr"], sz=14)
    ws.merge_cells("B3:D3"); ws.row_dimensions[3].height = 20
    hcl(ws, 3, 2,
        "Fuente: ACAL-00102-LC-E-0001 Rev.0 (513 cables CPF-2 / 81,231m) | Cronograma ABB RevA",
        C["hdr2"], sz=10)
    ws.row_dimensions[4].height = 8

    # Tabla volúmenes
    ws.row_dimensions[5].height = 18
    for c2, t in enumerate(["Parámetro","Alternativa 1 — Pre-Tendido","Alternativa 2 — Post-Arribo"], 2):
        hcl(ws, 5, c2, t, C["hdr"])

    TABLA = [
        # (parámetro, alt1, alt2, color_alt1, color_alt2)
        ("Inicio trabajos campo","25-ENE-2027 (tendido)","09-ABR-2027 (solo armado salas)", C["ok"], C["warn"]),
        ("Tendido cables BT grande (ruta crítica)","25-ENE → 04-ABR-2027 (79d)","29-MAY → 23-AGO-2027 (79d)", C["ok"], C["rsk"]),
        ("Llegada Sala #4 (MT)","08-ABR-2027","08-ABR-2027", C["ok"], C["ok"]),
        ("Inicio conexionado S4","23-ABR-2027 (post armado)","24-ABR-2027 (post armado)", C["ok"], C["ok"]),
        ("Llegada Sala #3 (BT)","14-MAY-2027","14-MAY-2027", C["ok"], C["ok"]),
        ("Inicio conexionado S3","29-MAY-2027","29-MAY-2027 (solo MT)", C["ok"], C["warn"]),
        ("Fin conexionado control+señales","12-JUL-2027","06-OCT-2027", C["ok"], C["rsk"]),
        ("Fin megger — apto energización","01-AGO-2027","30-OCT-2027", C["ok"], C["rsk"]),
        ("Energización MT","02-AGO-2027","31-OCT-2027", C["ok"], C["rsk"]),
        ("Energización BT completa","14-AGO-2027","14-NOV-2027", C["ok"], C["rsk"]),
        ("★ RFMC target","03-SEP-2027","05-DIC-2027", C["ok"], C["rsk"]),
        ("SAT PMS (HOLD POINT per PIE)","01-OCT → 15-DIC-2027","01-OCT → 30-NOV-2027 ⚠SOLAPADO", C["warn"], C["rsk"]),
        ("★ RFSU contractual","31-ENE-2028","31-ENE-2028", C["ok"], C["ok"]),
        ("Margen RFMC → RFSU","149 días ✅","57 días ⚠", C["ok"], C["rsk"]),
        ("Cuadrillas necesarias","3 cuadrillas (4p c/u)","3 cuadrillas + HORAS EXTRA", C["ok"], C["rsk"]),
        ("Riesgo cumplimiento RFSU","BAJO — margen amplio","ALTO — margen cero", C["ok"], C["rsk"]),
        ("VEREDICTO","✅ RECOMENDADA","⚠ SOLO CON RECURSOS ADICIONALES", C["ok"], C["rsk"]),
    ]

    for i, (param, a1, a2, c1, c2_) in enumerate(TABLA):
        r = 6 + i
        ws.row_dimensions[r].height = 22
        bold_row = "VEREDICTO" in param or "RFMC" in param or "RFSU" in param
        cel(ws, r, 2, param, bg=C["gray"], bold=bold_row, sz=9, h="left",
            col=C["hdr"] if bold_row else "000000")
        cel(ws, r, 3, a1, bg=c1,
            bold=bold_row, sz=9, h="left",
            col=C["ok_t"] if c1==C["ok"] else (C["warn_t"] if c1==C["warn"] else C["rsk_t"]))
        cel(ws, r, 4, a2, bg=c2_,
            bold=bold_row, sz=9, h="left",
            col=C["ok_t"] if c2_==C["ok"] else (C["warn_t"] if c2_==C["warn"] else C["rsk_t"]))

    # Volúmenes de trabajo
    r = 6 + len(TABLA) + 2
    ws.merge_cells(f"B{r}:D{r}"); ws.row_dimensions[r].height = 18
    hcl(ws, r, 2, "VOLUMEN DE TRABAJO CPF-2 — BASE CÁLCULO", C["hdr2"])
    r += 1

    VOLS = [
        ("Cables totales CPF-2","513 cables","/","81,231 m totales"),
        ("Cables potencia MT (6.6/13.2kV)","18 cables","/","2,110 m"),
        ("Cables potencia BT grande (≥95mm²)","104 cables — CRÍTICO","/","17,470 m"),
        ("Cables potencia BT mediano/pequeño","198 cables","/","33,125 m"),
        ("Cables control + señales","192 cables","/","28,126 m"),
        ("Fibra óptica + Ethernet","1 FO + varios UTP","/","400m FO"),
        ("Cargas CPF-2 identificadas","243 equipos","/","63 motores BT + 3 MT + 3 VFD-MT"),
        ("Cuadrillas estimadas","3 cuadrillas de 4 personas","/","8 hs/día efectivas"),
        ("Ruta crítica tendido BT grande","79 días paralelo","/","50 m/día/cuadrilla × 3 = 150m/día"),
    ]
    for param, v1, sep, v2 in VOLS:
        ws.row_dimensions[r].height = 18
        cel(ws, r, 2, param, bg=C["gray"], sz=9, h="left")
        cel(ws, r, 3, v1, bg="FFFFFF", bold=True, sz=9)
        cel(ws, r, 4, v2, bg="FFFFFF", sz=9, h="left", italic=True)
        r += 1

    # Alertas
    r += 1
    ws.merge_cells(f"B{r}:D{r}"); ws.row_dimensions[r].height = 18
    hcl(ws, r, 2, "ALERTAS Y ACCIONES REQUERIDAS", C["rfmc"])
    r += 1

    ALERTAS = [
        ("ACCIÓN INMEDIATA","ALT.1: Confirmar inicio tendido BT grande el 25-ENE-2027. "
         "Requiere que civil y bandejas estén listos para esa fecha."),
        ("ACCIÓN INMEDIATA","Confirmar con EPC disponibilidad de 3 cuadrillas eléctricas "
         "de 4 personas a partir de ENE-2027."),
        ("RIESGO ALTO","ALT.2: Tendido BT grande termina 23-AGO-2027 → conexionado control "
         "termina 06-OCT → megger 30-OCT → energización NOV → RFMC 05-DIC. "
         "Solo 57 días para comisionado completo."),
        ("RIESGO ALTO","SAT PMS (HOLD POINT per PIE) no puede solaparse con energización. "
         "En ALT.2 ambas actividades colisionan OCT-NOV 2027."),
        ("RECOMENDACIÓN","Adoptar ALT.1. Si por restricciones contractuales o civiles no es "
         "posible iniciar tendido en ENE-2027, agregar 1 cuadrilla extra en ALT.2 "
         "para acortar la ruta crítica de 79 → ~55 días."),
        ("NOTA TÉCNICA","Los 18 cables MT requieren personal especialista en terminales de MT. "
         "Confirmar subcontratista con ABB antes del inicio de comisionado."),
    ]
    for tipo_alerta, texto in ALERTAS:
        ws.row_dimensions[r].height = 30
        bg_a = C["rsk"] if "RIESGO" in tipo_alerta else (C["warn"] if "ACCIÓN" in tipo_alerta else C["ok"])
        fc_a = C["rsk_t"] if "RIESGO" in tipo_alerta else (C["warn_t"] if "ACCIÓN" in tipo_alerta else C["ok_t"])
        cel(ws, r, 2, tipo_alerta, bg=bg_a, bold=True, sz=9, col=fc_a)
        ws.merge_cells(f"C{r}:D{r}")
        cel(ws, r, 3, texto, bg="FFFFFF", sz=9, h="left", wrap=True, col=fc_a)
        r += 1

# ──────────────────────────────────────────────────────────────────────────────
#  MAIN
# ──────────────────────────────────────────────────────────────────────────────
def main():
    wb = Workbook()
    wb.remove(wb.active)

    make_comparacion(wb)
    make_gantt_sheet(wb, "1_Alt1_PreTendido",   ACT1, 1, "1F3864")
    make_gantt_sheet(wb, "2_Alt2_PostArribo",   ACT2, 2, "833C11")

    wb.save(OUTPUT)
    print(f"Generado: {OUTPUT}")
    print(f"\nResumen:")
    print(f"  ALT.1 (Pre-Tendido): Tendido inicia 25-ENE-2027 | RFMC: 03-SEP-2027 | "
          f"Margen RFSU: 149 días ✅")
    print(f"  ALT.2 (Post-Arribo): Todo inicia 09-ABR-2027   | RFMC: 05-DIC-2027 | "
          f"Margen RFSU:  57 días ⚠")
    print(f"\n  Semanas en Gantt: {len(SEMANAS)} | "
          f"Actividades ALT1: {sum(1 for a in ACT1 if a[5]!='sep')} | "
          f"Actividades ALT2: {sum(1 for a in ACT2 if a[5]!='sep')}")

if __name__ == "__main__":
    main()
