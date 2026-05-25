"""
Análisis comparativo definitivo: Cronograma ABB (Gantt interno) vs P6 Comercial CPF-2.
Basado en los códigos reales identificados en el XER.
"""

import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from analizar_cronogramas import (
    parse_xer, extraer_proyecto, extraer_wbs, extraer_tareas, PLANNING_DIR
)

BASE_DIR     = os.path.dirname(os.path.abspath(__file__))
PLANNING_DIR = os.path.join(BASE_DIR, "02_Planning")

# ── Estilos ───────────────────────────────────────────────────────────────────
def thin_border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(ws, row, cols, bg="CC0000", fg="FFFFFF", sz=10, height=24):
    ws.row_dimensions[row].height = height
    for c in range(1, cols + 1):
        cell = ws.cell(row=row, column=c)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(bold=True, color=fg, size=sz)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border()

def dc(ws, row, col, val, bg=None, bold=False, align="left", sz=10):
    cell = ws.cell(row=row, column=col, value=val)
    cell.border = thin_border()
    cell.alignment = Alignment(vertical="center", horizontal=align, wrap_text=True)
    cell.font = Font(bold=bold, size=sz)
    if bg:
        cell.fill = PatternFill("solid", fgColor=bg)
    ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 0, 20)
    return cell

def title_row(ws, text, row, cols, bg="CC0000", sz=13):
    ws.merge_cells(f"A{row}:{chr(64+cols)}{row}")
    cell = ws.cell(row=row, column=1, value=text)
    cell.fill = PatternFill("solid", fgColor=bg)
    cell.font = Font(bold=True, color="FFFFFF", size=sz)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[row].height = 34

def semaforo(delta, tol=14):
    if delta is None:      return "Sin datos en P6",   "D9D9D9"
    if abs(delta) <= tol:  return f"Compatible ({delta:+}d)", "C6EFCE"
    if abs(delta) <= 30:   return f"Desvío leve ({delta:+}d)", "FFEB9C"
    return f"{'Atraso' if delta>0 else 'Adelanto'} {abs(delta)}d", "FFC7CE"


# ── Tabla maestra de comparación ─────────────────────────────────────────────
# Columnas: Hito ABB, Fecha ABB, Código P6, Actividad P6, Fecha P6, Delta, Nota
COMPARACION = [
    # ── INGENIERÍA
    ("ING", "Gestión de compra — Shelter S#3",
     "GCE1100", "Gestión de compras - Shelters Sala Eléctrica # 3",
     date(2026,1,13), date(2026,5,28),   date(2026,5,28),
     "OC cerrada. P6 muestra mismo plazo fin."),

    ("ING", "Gestión de compra — Shelter S#4",
     "GCE1102", "Gestión de compras - Shelter Sala Eléctrica # 4",
     date(2026,1,13), date(2026,5,28),   date(2026,5,28),
     "OC cerrada. P6 muestra mismo plazo fin."),

    ("ING", "Gestión de compra — PMS",
     "GCE1070", "Gestión de compras - Power Manager System (PMS)",
     date(2026,1,13), date(2026,6,19),   date(2026,6,19),
     "Fecha Freezing Point IB PMS ABB = 19/06/2026. Coincide."),

    ("ING", "RI Shelters Sala #3 (inc. Tableros BT)",
     "RIE1040", "RI - Shelters Salas Eléctricas # 3 (Incl. Tableros BT)",
     date(2026,1,12), date(2026,1,12),   date(2026,1,12),
     "Hito emisión de Requerimiento Interno. Fecha pasada."),

    ("ING", "RI Shelters Sala #4 (inc. Tableros MT)",
     "RIE1042", "RI - Shelters Salas Eléctricas # 4 (Incl. Tableros MT)",
     date(2026,1,12), date(2026,1,12),   date(2026,1,12),
     "Hito emisión de RI. Fecha pasada."),

    ("ING", "RI Power Manager System (PMS)",
     "RIE1041", "RI - Power Manager System (PMS)",
     date(2026,1,12), date(2026,1,12),   date(2026,1,12),
     "Hito RI PMS. Fecha pasada."),

    # ── FABRICACIÓN Y DESPACHO
    ("FAB", "Fabricación y despacho — Shelter Sala #3",
     "FDE1103", "Fabricación y despacho - Shelter Sala Eléctrica # 3",
     date(2026,5,29), date(2027,5,24),   date(2027,5,24),
     "ABB Gantt: Despacho S#3 = 14/05/2027. P6 cierra 24/05/27 (+10d). Compatible."),

    ("FAB", "Fabricación y despacho — Shelter Sala #4",
     "FDE1104", "Fabricación y despacho - Shelter Sala Eléctrica # 4",
     date(2026,5,29), date(2027,6,23),   date(2027,4,8),
     "ABB Gantt: Despacho S#4 = 08/04/2027. P6 cierra 23/06/27 (+76d). ANALIZAR."),

    ("FAB", "Fabricación y despacho — PMS",
     "FDE1114", "Fabricación y despacho - Power Manager System (PMS)",
     date(2026,6,30), date(2027,6,2),    date(2026,10,23),
     "ABB Gantt: FAT PMS = 23/10/2026. P6 despacho hasta 02/06/27. P6 incluye el FAT interno de ABB como parte del proceso."),

    # ── FAT Y ENTREGABLES
    ("FAT", "Inicio entregas en Sala Eléctrica #3",
     "HITALER115", "Inicio entregas en sala eléctrica # 3",
     date(2027,3,1),  date(2027,3,1),    date(2027,4,2),
     "ABB Gantt: FAT S#3 inicio = 02/04/2027. P6 hito inicio entregas = 01/03/27 (-32d). P6 prevé llegada de materiales antes del FAT ABB."),

    ("FAT", "FAT Integrado — Integración Shelters y PMS",
     "TA1124", "Integracion Shelters y PMS (FAT)",
     date(2027,6,28), date(2027,6,28),   date(2027,6,25),
     "ABB Gantt: iFAT Integrado EPC 25/06-14/07/2027. P6 hito FAT = 28/06/27. Perfectamente alineado."),

    # ── CONSTRUCCIÓN EN CAMPO
    ("CAMPO", "Montaje Shelter Sala #3",
     "CE014050", "Montaje de shelter - Sala Eléctrica # 3",
     date(2027,6,28), date(2027,8,10),   None,
     "P6: Montaje en campo 28/06-10/08/2027. No hay fecha específica en Gantt ABB para esta actividad de campo (es responsabilidad EPC)."),

    ("CAMPO", "Montaje Shelter Sala #4",
     "CE018050", "Montaje de shelter - Sala Eléctrica # 4 - AJUST",
     date(2027,6,28), date(2027,8,21),   None,
     "P6: Montaje en campo 28/06-21/08/2027. Responsabilidad EPC. No en Gantt ABB."),

    ("CAMPO", "Montaje Transformadores SE #3 y #4",
     "CE014060", "Montaje de Transformadores y Generador de Emergencia - Salas Eléc.",
     date(2027,7,30), date(2027,8,21),   None,
     "P6: Actividad de campo posterior a entrega ABB. Responsabilidad EPC."),

    # ── PRECOMISIONADO Y RFSU
    ("COM", "Precomisionado (inicio)",
     "CDW1000", "Precomisionado",
     date(2027,5,3),  date(2028,1,13),   date(2027,5,3),
     "Fecha inicio idéntica al Gantt ABB. PERFECTAMENTE ALINEADO."),

    ("COM", "Comisionado (inicio)",
     "CDW1010", "Comisionado",
     date(2027,7,16), date(2028,1,31),   date(2027,7,16),
     "Fecha inicio idéntica al Gantt ABB. PERFECTAMENTE ALINEADO."),

    ("COM", "RFSU — Planta Total",
     "A10290", "Hito 11 - RFSU Planta Total",
     date(2028,1,31), date(2028,1,31),   date(2028,1,31),
     "RFSU 31/01/2028. PERFECTAMENTE ALINEADO. Fecha objetivo confirmada en ambos cronogramas."),
]


def generar():
    wb = Workbook()

    # ══════════════════════════════════════════════════════════════
    # HOJA 1: COMPARATIVA PRINCIPAL
    # ══════════════════════════════════════════════════════════════
    ws = wb.active
    ws.title = "Comparativa ABB vs P6"

    title_row(ws, "ANÁLISIS COMPARATIVO: Cronograma ABB (Gantt interno) vs Cronograma Global P6 (LB Comercial Rev.0)", 1, 9)
    ws["A2"] = f"Generado: {datetime.today().strftime('%d/%m/%Y %H:%M')}  |  Data Date P6: 05/12/2025  |  Proyecto: CPF-2 Vaca Muerta  |  RFSU Objetivo: 31/01/2028"
    ws["A2"].font = Font(italic=True, size=9, color="888888")
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 14

    headers = ["Fase", "Hito / Actividad ABB", "Cód. P6", "Actividad en P6",
               "Inicio P6", "Fin P6", "Fecha Ref. ABB", "Δ días", "Estado / Observación"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=3, column=col, value=h)
    hdr(ws, 3, len(headers), height=28)

    # Colores por fase
    fase_color = {
        "ING":   "EBF3FB",
        "FAB":   "FFF2CC",
        "FAT":   "E2EFDA",
        "CAMPO": "FCE4D6",
        "COM":   "E2EFDA",
    }
    fase_label = {
        "ING":   "Ingeniería / Compras",
        "FAB":   "Fabricación / Despacho",
        "FAT":   "FAT / Ensayos",
        "CAMPO": "Montaje en Campo (EPC)",
        "COM":   "Precomisionado / RFSU",
    }

    prev_fase = None
    data_row = 4
    for (fase, hito_abb, cod_p6, nombre_p6, ini_p6, fin_p6, fecha_abb, obs) in COMPARACION:
        # Separador de fase
        if fase != prev_fase:
            ws.merge_cells(f"A{data_row}:I{data_row}")
            ws.cell(row=data_row, column=1, value=f"  ▶  {fase_label[fase]}")
            ws.cell(row=data_row, column=1).fill = PatternFill("solid", fgColor="333333")
            ws.cell(row=data_row, column=1).font = Font(bold=True, color="FFFFFF", size=10)
            ws.cell(row=data_row, column=1).alignment = Alignment(vertical="center")
            ws.row_dimensions[data_row].height = 20
            data_row += 1
            prev_fase = fase

        bg_fase = fase_color.get(fase, "F5F5F5")

        # Delta y semáforo
        if fecha_abb and fin_p6:
            delta = (fin_p6 - fecha_abb).days
        elif fecha_abb and ini_p6:
            delta = (ini_p6 - fecha_abb).days
        else:
            delta = None
        estado_txt, sem_color = semaforo(delta)

        dc(ws, data_row, 1, fase,                      bg=bg_fase, bold=True, align="center")
        dc(ws, data_row, 2, hito_abb,                  bg=bg_fase, bold=True)
        dc(ws, data_row, 3, cod_p6,                    bg=bg_fase, align="center")
        dc(ws, data_row, 4, nombre_p6,                 bg=bg_fase)
        dc(ws, data_row, 5, ini_p6.strftime("%d/%m/%Y") if ini_p6 else "—", align="center")
        dc(ws, data_row, 6, fin_p6.strftime("%d/%m/%Y") if fin_p6 else "—", align="center")
        dc(ws, data_row, 7, fecha_abb.strftime("%d/%m/%Y") if fecha_abb else "—", align="center")
        dc(ws, data_row, 8, delta if delta is not None else "—", align="center", bold=True)
        dc(ws, data_row, 9, estado_txt, bg=sem_color)
        ws.row_dimensions[data_row].height = 22
        data_row += 1

    # Leyenda
    data_row += 1
    ws.merge_cells(f"A{data_row}:I{data_row}")
    ws.cell(row=data_row, column=1, value="REFERENCIA DE COLORES")
    hdr(ws, data_row, 9, bg="555555", sz=9, height=18)
    data_row += 1
    leyenda = [
        ("C6EFCE", "Compatible: desvío ≤ 14 días"),
        ("FFEB9C", "Desvío leve: entre 15 y 30 días"),
        ("FFC7CE", "Desvío significativo: > 30 días"),
        ("D9D9D9", "Sin datos comparables en P6"),
    ]
    for col, (color, texto) in enumerate(leyenda, 1):
        cell = ws.cell(row=data_row, column=col * 2 - 1, value=f"  {texto}")
        cell.fill = PatternFill("solid", fgColor=color)
        cell.font = Font(size=9)
        ws.merge_cells(f"{chr(64+col*2-1)}{data_row}:{chr(64+col*2)}{data_row}")
    ws.row_dimensions[data_row].height = 16

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 38
    ws.column_dimensions["C"].width = 13
    ws.column_dimensions["D"].width = 48
    ws.column_dimensions["E"].width = 13
    ws.column_dimensions["F"].width = 13
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 9
    ws.column_dimensions["I"].width = 40
    ws.freeze_panes = ws["A4"]


    # ══════════════════════════════════════════════════════════════
    # HOJA 2: CONCLUSIONES Y ALERTAS
    # ══════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Conclusiones y Alertas")
    title_row(ws2, "CONCLUSIONES DEL ANÁLISIS COMPARATIVO — CPF-2 / PROYECTO ABB", 1, 5)
    ws2["A2"] = f"Fecha: {datetime.today().strftime('%d/%m/%Y')}"
    ws2["A2"].font = Font(italic=True, size=9, color="888888")
    ws2.merge_cells("A2:E2")

    secciones = [
        ("ALINEACIÓN CONFIRMADA ✅", "C6EFCE", "333333", [
            ("RFSU Planta Total",
             "31/01/2028 en ambos cronogramas. Fecha objetivo firme y consistente."),
            ("Precomisionado",
             "Inicio 03/05/2027 idéntico en P6 (CDW1000) y Gantt ABB."),
            ("Comisionado",
             "Inicio 16/07/2027 idéntico en P6 (CDW1010) y Gantt ABB."),
            ("FAT Integrado EPC",
             "P6 hito TA1124 = 28/06/2027. Gantt ABB iFAT = 25/06-14/07/2027. Diferencia de 3 días — compatible."),
            ("Montaje Shelters en campo",
             "P6: S#3 = 28/06-10/08/2027, S#4 = 28/06-21/08/2027. Alineado con la secuencia post-despacho ABB."),
            ("Despacho Sala #3",
             "ABB Gantt: 14/05/2027 — P6 (FDE1103): 24/05/2027. Diferencia de +10 días. Margen aceptable."),
        ]),
        ("PUNTOS A ANALIZAR / ACLARAR ⚠", "FFEB9C", "555500", [
            ("Despacho Sala Eléctrica #4 — Desvío de ~76 días",
             "ABB Gantt: Despacho S#4 = 08/04/2027. P6 (FDE1104): fin 23/06/2027. "
             "Diferencia de +76 días. POSIBLE EXPLICACIÓN: El P6 incluye el FAT interno de ABB, "
             "integración del shelter completo y logística. El Gantt ABB refleja solo el FAT de equipos. "
             "ACCIÓN: Confirmar con AESA si la fecha 23/06 en P6 es el despacho físico a Vaca Muerta "
             "o incluye el FAT de shelter completo."),
            ("FAT PMS — Fechas muy diferentes",
             "ABB Gantt: FAT PMS = 19-23/10/2026. P6 (FDE1114): despacho PMS hasta 02/06/2027. "
             "El P6 no registra el FAT interno de ABB como hito separado. El FAT de PMS (Oct 2026) "
             "es un evento interno de ABB que ocurre antes del cierre de la tarea P6 de fabricación. "
             "ACCIÓN: Solicitar a AESA que incluya el hito FAT PMS en el P6 global."),
            ("FAT Sala #4 — No hay hito específico en P6",
             "ABB Gantt: FAT Sala #4 = 08-16/03/2027. En el P6 no existe un hito 'FAT Sala #4' "
             "separado. Solo aparece 'Inicio entregas en sala eléctrica #3' (01/03/2027). "
             "ACCIÓN: Solicitar incorporar hitos FAT individuales de S#3 y S#4 al P6 global."),
        ]),
        ("RIESGOS IDENTIFICADOS 🔴", "FFC7CE", "880000", [
            ("Riesgo 1 — Cadena crítica FAT → Despacho → Montaje S#4",
             "La diferencia entre la fecha de despacho ABB (08/04/2027) y la fecha P6 (23/06/2027) "
             "implica que si ABB cumple su Gantt, hay 76 días de flotante. Pero si ABB se demora, "
             "el montaje de shelter (28/06) queda sin margen. MONITOREAR mensualmente."),
            ("Riesgo 2 — FAT PMS no visible en P6 como hito controlable",
             "Al no tener el FAT PMS (Oct 2026) como hito en el P6 comercial, cualquier desvío "
             "en el PMS no activará alertas en el cronograma EPC. Riesgo de detección tardía."),
            ("Riesgo 3 — Freezing Points no representados en P6",
             "Los 11 Freezing Points de ingeniería de ABB (Jun-Ago 2026) no tienen equivalente "
             "en el P6 global. Si alguno se demora, el impacto no es visible en el P6 hasta que "
             "afecte fechas de fabricación."),
        ]),
        ("RECOMENDACIONES PARA PRÓXIMA REUNIÓN 📋", "EBF3FB", "003366", [
            ("1. Solicitar actualización del P6",
             "Pedir a AESA que incorpore como hitos controlables: FAT PMS (Oct 2026), "
             "FAT Sala #3 (Abr 2027), FAT Sala #4 (Mar 2027) y Freezing Points clave."),
            ("2. Aclarar FDE1104 vs Despacho ABB",
             "Confirmar si FDE1104 (fin 23/06/2027) representa el despacho a Vaca Muerta "
             "o el cierre del proceso FAT completo del shelter."),
            ("3. Establecer reporte mensual de estado",
             "Con este análisis como base, proponer un reporte mensual comparando avance real "
             "de ABB contra el P6. Foco en: Despacho S#3 y S#4, FAT PMS, iFAT Integrado."),
            ("4. Confirmar fecha KOM técnico",
             "El P6 no registra el KOM técnico de ABB (15/05/2026). Confirmar realización "
             "y solicitar hito en P6."),
        ]),
    ]

    row = 4
    for titulo, bg_titulo, fg_titulo, items in secciones:
        ws2.merge_cells(f"A{row}:E{row}")
        ws2.cell(row=row, column=1, value=titulo)
        ws2.cell(row=row, column=1).fill = PatternFill("solid", fgColor=bg_titulo)
        ws2.cell(row=row, column=1).font = Font(bold=True, color=fg_titulo, size=11)
        ws2.cell(row=row, column=1).alignment = Alignment(vertical="center")
        ws2.row_dimensions[row].height = 26
        row += 1

        for subtitulo, detalle in items:
            dc(ws2, row, 1, subtitulo, bold=True, bg="F5F5F5")
            ws2.merge_cells(f"B{row}:E{row}")
            cell = ws2.cell(row=row, column=2, value=detalle)
            cell.alignment = Alignment(wrap_text=True, vertical="top")
            cell.border = thin_border()
            ws2.row_dimensions[row].height = 60
            row += 1
        row += 1

    ws2.column_dimensions["A"].width = 38
    ws2.column_dimensions["B"].width = 20
    ws2.column_dimensions["C"].width = 20
    ws2.column_dimensions["D"].width = 20
    ws2.column_dimensions["E"].width = 20
    # Ajustar columna B-E como merged
    for col in "BCDE":
        ws2.column_dimensions[col].width = 25


    # ══════════════════════════════════════════════════════════════
    # HOJA 3: TABLA DE HITOS P6 COMPLETA (ABB/Shelters/RFSU)
    # ══════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Hitos P6 ABB-Eléctrico")
    title_row(ws3, "HITOS P6 RELACIONADOS CON SALAS ELÉCTRICAS, SHELTERS Y RFSU", 1, 6)

    hdr(ws3, 2, 6, bg="880000", height=24)
    for col, h in enumerate(["Código P6", "Actividad", "WBS", "Inicio", "Fin", "Observación"], 1):
        ws3.cell(row=2, column=col, value=h)

    hitos_p6 = [
        ("RIE1040",   "RI - Shelters Salas Eléctricas # 3 (Incl. Tableros BT)",          "RI´s",              date(2026,1,12),  date(2026,1,12),  "Hito pasado"),
        ("RIE1042",   "RI - Shelters Salas Eléctricas # 4 (Incl. Tableros MT)",           "RI´s",              date(2026,1,12),  date(2026,1,12),  "Hito pasado"),
        ("RIE1041",   "RI - Power Manager System (PMS)",                                   "RI´s",              date(2026,1,12),  date(2026,1,12),  "Hito pasado"),
        ("GCE1100",   "Gestión de compras - Shelters Sala Eléctrica # 3",                  "ELECTRICIDAD",      date(2026,1,13),  date(2026,5,28),  "OC emitida"),
        ("GCE1102",   "Gestión de compras - Shelter Sala Eléctrica # 4",                   "ELECTRICIDAD",      date(2026,1,13),  date(2026,5,28),  "OC emitida"),
        ("GCE1070",   "Gestión de compras - Power Manager System (PMS)",                   "LB Rev.1",          date(2026,1,13),  date(2026,6,19),  "Cierra al Freezing Point IB PMS ABB"),
        ("FDE1103",   "Fabricación y despacho - Shelter Sala Eléctrica # 3",               "ELECTRICIDAD",      date(2026,5,29),  date(2027,5,24),  "Incluye FAT interno ABB + despacho"),
        ("FDE1104",   "Fabricación y despacho - Shelter Sala Eléctrica # 4",               "ELECTRICIDAD",      date(2026,5,29),  date(2027,6,23),  "ANALIZAR vs Despacho ABB 08/04/27"),
        ("FDE1114",   "Fabricación y despacho - Power Manager System (PMS)",               "LB Rev.1",          date(2026,6,30),  date(2027,6,2),   "Cubre FAT PMS (Oct-26) hasta entrega"),
        ("HITALER115","Inicio entregas en sala eléctrica # 3",                              "LB Rev.1",          date(2027,3,1),   date(2027,3,1),   "Hito recepción de materiales en S#3"),
        ("TA1124",    "Integracion Shelters y PMS (FAT)",                                   "LB Rev.1",          date(2027,6,28),  date(2027,6,28),  "iFAT Integrado. ALINEADO con Gantt ABB"),
        ("CE014050",  "Montaje de shelter - Sala Eléctrica # 3",                            "Montaje Shelters",  date(2027,6,28),  date(2027,8,10),  "Responsabilidad EPC en campo"),
        ("CE018050",  "Montaje de shelter - Sala Eléctrica # 4",                            "Montaje Shelters",  date(2027,6,28),  date(2027,8,21),  "Responsabilidad EPC en campo"),
        ("CE014060",  "Montaje de Transformadores y Generador de Emergencia - Salas Eléc.", "Montaje Shelters",  date(2027,7,30),  date(2027,8,21),  "Actividad de campo post-entrega ABB"),
        ("CDW1000",   "Precomisionado",                                                     "PRECOMISIONADO",    date(2027,5,3),   date(2028,1,13),  "INICIO IDÉNTICO al Gantt ABB"),
        ("CDW1010",   "Comisionado",                                                        "COMISIONADO",       date(2027,7,16),  date(2028,1,31),  "INICIO IDÉNTICO al Gantt ABB"),
        ("A10290",    "Hito 11 - RFSU Planta Total",                                        "HITOS DE TRABAJOS", date(2028,1,31),  date(2028,1,31),  "RFSU 31/01/2028 — PERFECTAMENTE ALINEADO"),
    ]

    color_tipo = {
        "RI´s":             "EBF3FB",
        "ELECTRICIDAD":     "FFF2CC",
        "LB Rev.1":         "FFF2CC",
        "Montaje Shelters": "FCE4D6",
        "PRECOMISIONADO":   "E2EFDA",
        "COMISIONADO":      "E2EFDA",
        "HITOS DE TRABAJOS":"C6EFCE",
    }
    for row_i, (cod, nombre, wbs, ini, fin, obs) in enumerate(hitos_p6, 3):
        bg = color_tipo.get(wbs, "F5F5F5")
        dc(ws3, row_i, 1, cod,   bg=bg, bold=True, align="center")
        dc(ws3, row_i, 2, nombre, bg=bg)
        dc(ws3, row_i, 3, wbs,   bg=bg)
        dc(ws3, row_i, 4, ini.strftime("%d/%m/%Y"), align="center")
        dc(ws3, row_i, 5, fin.strftime("%d/%m/%Y"), align="center")
        dc(ws3, row_i, 6, obs)
        ws3.row_dimensions[row_i].height = 22

    ws3.column_dimensions["A"].width = 14
    ws3.column_dimensions["B"].width = 55
    ws3.column_dimensions["C"].width = 22
    ws3.column_dimensions["D"].width = 13
    ws3.column_dimensions["E"].width = 13
    ws3.column_dimensions["F"].width = 45
    ws3.freeze_panes = ws3["A3"]

    out = os.path.join(PLANNING_DIR, "Analisis_Comparativo_ABB_vs_P6.xlsx")
    wb.save(out)
    return out


if __name__ == "__main__":
    print("\n=== Generando Análisis Comparativo ABB vs P6 ===\n")
    out = generar()
    print(f"✅ Reporte guardado: {out}")
    print("\nHojas generadas:")
    print("  1. Comparativa ABB vs P6  — tabla detallada con semáforos")
    print("  2. Conclusiones y Alertas — resumen ejecutivo con recomendaciones")
    print("  3. Hitos P6 ABB-Eléctrico — todos los hitos eléctricos del P6")
