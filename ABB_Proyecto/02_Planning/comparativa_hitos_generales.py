"""
Comparativa de alto nivel: Provisiones ABB (fechas más tardías) vs Cronograma P6.
Nivel de apertura equivalente al P6 comercial.
"""

import os
from datetime import datetime, date
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

PLANNING_DIR = os.path.dirname(os.path.abspath(__file__))

def S(style="thin", color="CCCCCC"):
    return Side(style=style, color=color)

def border(color="CCCCCC"):
    s = S(color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def fill(color):
    return PatternFill("solid", fgColor=color)

def font(bold=False, size=10, color="000000", italic=False):
    return Font(bold=bold, size=size, color=color, italic=italic)

def align(h="left", v="center", wrap=True):
    return Alignment(horizontal=h, vertical=v, wrap_text=wrap)

def cell(ws, row, col, val, bg=None, bold=False, sz=10, h="left", color="000000",
         italic=False, height=22):
    c = ws.cell(row=row, column=col, value=val)
    c.border   = border()
    c.alignment = align(h)
    c.font      = font(bold=bold, size=sz, color=color, italic=italic)
    if bg:
        c.fill = fill(bg)
    ws.row_dimensions[row].height = max(ws.row_dimensions[row].height or 0, height)
    return c

def title(ws, text, row, end_col, bg="CC0000", sz=13, height=36):
    ws.merge_cells(f"A{row}:{chr(64+end_col)}{row}")
    c = ws.cell(row=row, column=1, value=text)
    c.fill = fill(bg)
    c.font = font(bold=True, size=sz, color="FFFFFF")
    c.alignment = align("center")
    ws.row_dimensions[row].height = height

def section(ws, text, row, end_col, bg="333333", height=22):
    ws.merge_cells(f"A{row}:{chr(64+end_col)}{row}")
    c = ws.cell(row=row, column=1, value=f"  {text}")
    c.fill = fill(bg)
    c.font = font(bold=True, size=10, color="FFFFFF")
    c.alignment = align("left")
    ws.row_dimensions[row].height = height

def header_row(ws, row, headers, bg="880000", sz=9, height=26):
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=row, column=col, value=h)
        c.fill = fill(bg)
        c.font = font(bold=True, size=sz, color="FFFFFF")
        c.alignment = align("center")
        c.border = border()
    ws.row_dimensions[row].height = height


# ═══════════════════════════════════════════════════════════════════════════════
# DATOS DEL ANÁLISIS
# ═══════════════════════════════════════════════════════════════════════════════

# ── ABB: Fechas más tardías por provisión (peor caso de cada entregable)
# Fuente: CPF2 Gantt EPC Integrado.xlsx
ABB_PROVISIONES = [
    # (Provisión, Descripción, Fecha más tardía ABB, Qué representa)
    ("OC / Inicio",
     "Recepción OC en ABB — Inicio del proyecto",
     date(2026, 5, 8),
     "Fecha de inicio: KOM técnico 15/05/2026"),

    ("Ingeniería — Freezing Points",
     "Último Freezing Point de Ingeniería Constructiva\n(Sala #3 — más tardía de las dos salas)",
     date(2026, 8, 31),
     "Cierre total de ingeniería ABB. S#4 cierra antes: 24/08/2026"),

    ("FAT — PMS",
     "Factory Acceptance Test Power Management System\n(único equipo — Salas #3 y #4 compartido)",
     date(2026, 10, 23),
     "Duración FAT PMS: 19-23/10/2026 (4 días)"),

    ("FAT — Salas Eléctricas",
     "FAT Sala Eléctrica más tardío\n(Sala #3 — más tardía: S#4 FAT termina 16/03/2027)",
     date(2027, 4, 19),
     "FAT S#3: 02/04-19/04/2027. FAT S#4: 08/03-16/03/2027"),

    ("Despacho — Provisión ABB",
     "Último despacho de las dos salas\n(Sala #3 — más tardía de las dos)",
     date(2027, 5, 14),
     "Despacho S#3: 10-14/05/2027. Despacho S#4: 02-08/04/2027"),
]

# ── P6: Hitos equivalentes a nivel general
# Fuente: Cronograma CPF2 Reanudación Rev.0 LB Comercial (XER)
P6_HITOS = [
    # (Código, Actividad P6, Inicio P6, Fin P6, Qué representa en el proyecto)
    ("GCE1100/1102/1070",
     "Gestión de compras Shelters S#3, S#4 y PMS",
     date(2026, 1, 13), date(2026, 6, 19),
     "OC emitida antes del inicio ABB. P6 cierra con Freezing PMS 19/06"),

    ("—",
     "Sin hito específico de Freezing Points en P6",
     None, None,
     "Los Freezing Points de ABB no están representados en el P6 comercial"),

    ("FDE1114",
     "Fabricación y despacho — Power Manager System (PMS)",
     date(2026, 6, 30), date(2027, 6, 2),
     "P6 agrupa FAT + entrega PMS en una sola actividad. Cierra 02/06/2027"),

    ("FDE1103 / FDE1104",
     "Fabricación y despacho — Shelter Sala #3 y Sala #4",
     date(2026, 5, 29), date(2027, 6, 23),
     "P6 agrupa Ing + FAB + FAT + Despacho. La más tardía (S#4) cierra 23/06/2027"),

    ("FDE1103 / FDE1104",
     "Fabricación y despacho — Shelter Sala #3 (más tardía)",
     date(2026, 5, 29), date(2027, 5, 24),
     "S#3 cierra 24/05/2027 en P6. S#4 cierra 23/06/2027"),
]

# ── Hitos de demanda: ¿cuándo NECESITA el proyecto las entregas de ABB?
DEMANDA_P6 = [
    ("TA1124",    "Integración Shelters y PMS (iFAT)",  date(2027, 6, 28)),
    ("CE014050",  "Montaje Shelter Sala #3 en campo",   date(2027, 6, 28)),
    ("CE018050",  "Montaje Shelter Sala #4 en campo",   date(2027, 6, 28)),
    ("CDW1000",   "Inicio Precomisionado",               date(2027, 5, 3)),
    ("CDW1010",   "Inicio Comisionado",                  date(2027, 7, 16)),
    ("A10290",    "RFSU — Planta Total",                 date(2028, 1, 31)),
]

# ── Tabla central: ABB provisión más tardía vs fecha que el P6 la necesita
TABLA_CENTRAL = [
    # (Provisión ABB, Fecha ABB peor caso, Fecha P6 de necesidad, Actividad P6 que la necesita, Cod P6)
    ("OC / Inicio del proyecto",
     date(2026, 5, 8),
     date(2026, 5, 29),
     "Inicio fabricación P6 (FDE1103/1104)",
     "FDE1103/1104",
     "P6 prevé arranque de fabricación 21 días después del inicio ABB. Holgura positiva."),

    ("Freezing Points de ingeniería (último: S#3)",
     date(2026, 8, 31),
     None,
     "Sin hito equivalente en P6 — no controlado",
     "—",
     "RIESGO: Los Freezing Points no están en el P6. Un desvío aquí no es visible en el cronograma global hasta que impacte el despacho."),

    ("FAT PMS (único FAT de PMS)",
     date(2026, 10, 23),
     date(2027, 6, 2),
     "Cierre tarea FDE1114 (Fab. y despacho PMS)",
     "FDE1114",
     "P6 da plazo hasta 02/06/2027 para cerrar PMS. FAT ABB es 23/10/2026. Holgura de 222 días — amplia."),

    ("FAT Salas Eléctricas (más tardío: Sala #3)",
     date(2027, 4, 19),
     date(2027, 5, 24),
     "Cierre FDE1103 (Fab. y despacho Shelter S#3)",
     "FDE1103",
     "FAT S#3 termina 19/04 y P6 cierra la tarea el 24/05. Holgura de 35 días para embalaje, logística y despacho."),

    ("Despacho final ABB (más tardío: Sala #3)",
     date(2027, 5, 14),
     date(2027, 6, 28),
     "Inicio montaje shelters en campo (CE014050/CE018050)",
     "CE014050",
     "Despacho S#3 el 14/05 y campo necesita el shelter el 28/06. Holgura de 45 días para transporte + recepción en Vaca Muerta."),
]

COLORES = {
    "ok":     "C6EFCE",
    "warn":   "FFEB9C",
    "riesgo": "FFC7CE",
    "info":   "D9E1F2",
    "na":     "D9D9D9",
}


def holgura_color(dias):
    if dias is None:    return "N/A", COLORES["na"]
    if dias >= 30:      return f"+{dias}d  HOLGURA AMPLIA",  COLORES["ok"]
    if dias >= 7:       return f"+{dias}d  HOLGURA AJUSTADA", COLORES["warn"]
    if dias >= 0:       return f"+{dias}d  SIN MARGEN",      COLORES["riesgo"]
    return f"{dias}d  ATRASO",                               COLORES["riesgo"]


def generar():
    wb = Workbook()

    # ════════════════════════════════════════════════════════════
    # HOJA 1: COMPARATIVA GENERAL
    # ════════════════════════════════════════════════════════════
    ws = wb.active
    ws.title = "Comparativa General"

    title(ws, "COMPARATIVA DE ALTO NIVEL — PROVISIONES ABB vs CRONOGRAMA GLOBAL P6 (CPF-2)", 1, 7)
    ws.cell(row=2, column=1,
            value="Criterio: se toma la fecha MÁS TARDÍA de cada provisión ABB (peor caso) y "
                  "se compara contra la fecha en que el P6 la necesita disponible en campo.").font = font(italic=True, size=9, color="666666")
    ws.merge_cells("A2:G2")
    ws.row_dimensions[2].height = 16

    header_row(ws, 3,
               ["Provisión ABB", "Fecha ABB\n(peor caso)", "Fecha P6\n(necesidad)",
                "Holgura\n(días)", "Estado", "Actividad P6 de referencia", "Observación"], height=32)

    for i, (prov, fecha_abb, fecha_p6, act_p6, cod_p6, obs) in enumerate(TABLA_CENTRAL, 4):
        delta = (fecha_p6 - fecha_abb).days if fecha_p6 and fecha_abb else None
        estado_txt, col_estado = holgura_color(delta)

        # color de fila por estado
        bg_row = "F5F5F5" if i % 2 == 0 else "FFFFFF"

        cell(ws, i, 1, prov,   bg=bg_row, bold=True, height=40)
        cell(ws, i, 2, fecha_abb.strftime("%d/%m/%Y") if fecha_abb else "—",
             bg=bg_row, h="center", height=40)
        cell(ws, i, 3, fecha_p6.strftime("%d/%m/%Y") if fecha_p6 else "Sin hito en P6",
             bg=bg_row, h="center", height=40)
        cell(ws, i, 4, delta if delta is not None else "—",
             bg=bg_row, h="center", bold=True, height=40)
        cell(ws, i, 5, estado_txt, bg=col_estado, h="center", bold=True, height=40)
        cell(ws, i, 6, f"[{cod_p6}] {act_p6}", bg=bg_row, height=40)
        cell(ws, i, 7, obs, bg=bg_row, height=40)

    # Línea resumen RFSU
    r = 4 + len(TABLA_CENTRAL) + 1
    ws.merge_cells(f"A{r}:G{r}")
    c = ws.cell(row=r, column=1,
                value="  ★  RFSU PLANTA TOTAL: 31/01/2028 — ALINEADO EN AMBOS CRONOGRAMAS   "
                      "|   Precomisionado inicio: 03/05/2027 — IDÉNTICO   |   Comisionado inicio: 16/07/2027 — IDÉNTICO")
    c.fill  = fill("1F4E79")
    c.font  = font(bold=True, size=11, color="FFFFFF")
    c.alignment = align("center")
    ws.row_dimensions[r].height = 28

    ws.column_dimensions["A"].width = 32
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 10
    ws.column_dimensions["E"].width = 22
    ws.column_dimensions["F"].width = 38
    ws.column_dimensions["G"].width = 50
    ws.freeze_panes = ws["A4"]


    # ════════════════════════════════════════════════════════════
    # HOJA 2: CADENA DE HITOS — vista cronológica
    # ════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Cadena Cronológica")
    title(ws2, "CADENA CRONOLÓGICA DE HITOS GENERALES — ABB ↔ PROYECTO CPF-2", 1, 6)

    CADENA = [
        # (Fecha, Fuente, Hito, Tipo, Observación)
        (date(2026, 5,  8),  "ABB",       "OC recibida — Inicio proyecto ABB",                      "Arranque",    "KOM técnico previsto 15/05/2026"),
        (date(2026, 5, 29),  "P6",        "Inicio tarea Fabricación y despacho Shelters S#3 y S#4",  "P6",          "FDE1103/1104. P6 prevé inicio 21d después del inicio ABB"),
        (date(2026, 6, 19),  "P6",        "Cierre OC/Compras PMS",                                   "P6",          "GCE1070 — coincide con Freezing Point IB PMS de ABB"),
        (date(2026, 6, 30),  "P6",        "Inicio tarea Fabricación y despacho PMS",                 "P6",          "FDE1114. Arranca post-freezing ingeniería PMS"),
        (date(2026, 8, 31),  "ABB",       "Último Freezing Point ingeniería (Sala #3)",               "Freezing",    "Cierre total ingeniería ABB. Sin hito equivalente en P6 ⚠"),
        (date(2026,10, 23),  "ABB",       "FAT PMS — Power Management System",                       "FAT interno", "FAT interno ABB, no visible en P6 como hito separado"),
        (date(2027, 3, 16),  "ABB",       "FAT Sala Eléctrica #4 (finaliza)",                        "FAT interno", "Duración: 08-16/03/2027. No tiene hito en P6"),
        (date(2027, 4, 19),  "ABB",       "FAT Sala Eléctrica #3 — MÁS TARDÍO (finaliza)",           "FAT interno", "Duración: 02-19/04/2027. No tiene hito en P6"),
        (date(2027, 5,  3),  "P6/ABB",    "Inicio Precomisionado",                                   "Comisionado", "CDW1000. FECHA IDÉNTICA en P6 y Gantt ABB ✅"),
        (date(2027, 5, 14),  "ABB",       "Despacho Sala #3 — MÁS TARDÍO (sale de fábrica)",         "Despacho",    "Último despacho ABB. 45 días antes de montaje en campo"),
        (date(2027, 5, 24),  "P6",        "Cierre FDE1103 — Shelter Sala #3",                        "P6",          "P6 cierra esta tarea 10d después del despacho ABB S#3"),
        (date(2027, 6,  2),  "P6",        "Cierre FDE1114 — PMS",                                    "P6",          "P6 cierra PMS 222d después del FAT interno ABB (Oct-26)"),
        (date(2027, 6, 23),  "P6",        "Cierre FDE1104 — Shelter Sala #4",                        "P6",          "Solo 5 días antes del inicio de montaje en campo ⚠"),
        (date(2027, 6, 28),  "P6/ABB",    "iFAT — Integración Shelters y PMS",                       "iFAT",        "TA1124. Gantt ABB: 25/06-14/07/2027. ALINEADO ✅"),
        (date(2027, 6, 28),  "P6",        "Inicio montaje Shelters en campo (S#3 y S#4)",             "Campo EPC",   "CE014050/CE018050. Responsabilidad EPC"),
        (date(2027, 7, 16),  "P6/ABB",    "Inicio Comisionado",                                      "Comisionado", "CDW1010. FECHA IDÉNTICA en P6 y Gantt ABB ✅"),
        (date(2027, 8, 10),  "P6",        "Fin montaje Shelter Sala #3",                              "Campo EPC",   "CE014050"),
        (date(2027, 8, 21),  "P6",        "Fin montaje Shelter Sala #4",                              "Campo EPC",   "CE018050"),
        (date(2028, 1, 31),  "P6/ABB",    "RFSU — Ready For Start Up Planta Total",                  "RFSU",        "A10290. FECHA IDÉNTICA ✅ Objetivo firme"),
    ]

    header_row(ws2, 2, ["Fecha", "Fuente", "Hito / Actividad", "Tipo", "Observación"], height=26)

    color_fuente = {
        "ABB":     "FFE0E0",
        "P6":      "EBF3FB",
        "P6/ABB":  "E2EFDA",
    }
    color_tipo = {
        "Arranque":    "D9E1F2",
        "Freezing":    "FCE4D6",
        "FAT interno": "FFF2CC",
        "Despacho":    "FFEB9C",
        "iFAT":        "C6EFCE",
        "Comisionado": "E2EFDA",
        "Campo EPC":   "F2F2F2",
        "RFSU":        "C6EFCE",
        "P6":          "EBF3FB",
    }

    for i, (fec, fuente, hito, tipo, obs) in enumerate(CADENA, 3):
        bg_f = color_fuente.get(fuente, "FFFFFF")
        bg_t = color_tipo.get(tipo,    "FFFFFF")
        cell(ws2, i, 1, fec.strftime("%d/%m/%Y"), bg=bg_f, h="center", bold=True, height=22)
        cell(ws2, i, 2, fuente, bg=bg_f, h="center", bold=True, height=22)
        cell(ws2, i, 3, hito,   bg=bg_f, bold=(fuente=="P6/ABB"), height=22)
        cell(ws2, i, 4, tipo,   bg=bg_t, h="center", height=22)
        cell(ws2, i, 5, obs,    height=22)

    ws2.column_dimensions["A"].width = 14
    ws2.column_dimensions["B"].width = 10
    ws2.column_dimensions["C"].width = 52
    ws2.column_dimensions["D"].width = 16
    ws2.column_dimensions["E"].width = 55
    ws2.freeze_panes = ws2["A3"]

    # Leyenda
    lr = 3 + len(CADENA) + 1
    ws2.merge_cells(f"A{lr}:E{lr}")
    ws2.cell(row=lr, column=1, value="REFERENCIAS DE COLOR").fill = fill("333333")
    ws2.cell(row=lr, column=1).font = font(bold=True, color="FFFFFF", size=9)
    ws2.row_dimensions[lr].height = 16
    leyenda = [("FFE0E0","Hito ABB"), ("EBF3FB","Hito P6"), ("E2EFDA","Coincide ABB y P6")]
    for col, (c, t) in enumerate(leyenda, 1):
        ws2.cell(row=lr+1, column=col, value=t).fill = fill(c)
        ws2.cell(row=lr+1, column=col).font = font(size=9)
        ws2.cell(row=lr+1, column=col).border = border()


    # ════════════════════════════════════════════════════════════
    # HOJA 3: CONCLUSIÓN EJECUTIVA
    # ════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Conclusión Ejecutiva")
    title(ws3, "CONCLUSIÓN EJECUTIVA — ¿EL CRONOGRAMA ABB ES COMPATIBLE CON EL P6?", 1, 4)

    BLOQUES = [
        ("VEREDICTO GENERAL", "1F4E79", "FFFFFF",
         "LAS FECHAS DE PROVISIÓN ABB SON COMPATIBLES CON EL CRONOGRAMA P6 DEL PROYECTO CPF-2.\n"
         "Tomando el peor caso de cada entregable, ABB entrega antes de que el proyecto los necesite "
         "en todos los casos. El RFSU de 31/01/2028 es consistente en ambos cronogramas."),

        ("HOLGURAS POR PROVISIÓN", "C6EFCE", "1A5C1A", None),
        ("ALERTA — 1 punto crítico", "FFC7CE", "880000", None),
        ("RIESGO — 1 punto de atención", "FFEB9C", "7D4B00", None),
        ("RECOMENDACIÓN PARA PRÓXIMA REUNIÓN", "EBF3FB", "1A3A5C", None),
    ]

    DETALLE_BLOQUES = {
        "HOLGURAS POR PROVISIÓN": [
            ("PMS (FAT 23/10/2026):",       "P6 lo necesita el 02/06/2027 → HOLGURA 222 DÍAS  ✅"),
            ("FAT Sala #3 (19/04/2027):",   "P6 cierra FDE1103 el 24/05/2027 → 35 días para despacho  ✅"),
            ("Despacho Sala #3 (14/05/27):","P6 empieza montaje el 28/06/2027 → 45 días de transporte  ✅"),
            ("Despacho Sala #4 (08/04/27):","P6 empieza montaje el 28/06/2027 → 81 días de transporte  ✅"),
            ("iFAT Integrado:",             "ABB 25/06, P6 28/06/2027 → diferencia de 3 días  ✅"),
            ("RFSU:",                        "31/01/2028 en ambos cronogramas → EXACTAMENTE IGUAL  ✅"),
        ],
        "ALERTA — 1 punto crítico": [
            ("FDE1104 cierra 23/06/2027:",
             "El P6 cierra la tarea de fabricación Sala #4 apenas 5 días antes del inicio "
             "del montaje en campo (28/06/2027). NO HAY MARGEN en el P6 para S#4. "
             "Si ABB se demora en S#4, el montaje en campo no tiene flotante.\n"
             "→ Monitorear mensualmente el avance de S#4."),
        ],
        "RIESGO — 1 punto de atención": [
            ("Freezing Points no están en P6:",
             "Los 11 Freezing Points de ingeniería de ABB (Jun-Ago 2026) no tienen "
             "representación en el P6 comercial. Un desvío en ellos no es detectable "
             "en el cronograma del proyecto hasta que impacta el despacho.\n"
             "→ Solicitar incorporar al menos los Freezing Points de Ing. Constructiva "
             "como hitos controlables en el P6."),
        ],
        "RECOMENDACIÓN PARA PRÓXIMA REUNIÓN": [
            ("1.", "Pedir a AESA que incorpore en el P6: FAT PMS (Oct-26), FAT Sala #4 (Mar-27), FAT Sala #3 (Abr-27) y los Freezing Points de Ing. Constructiva."),
            ("2.", "Aclarar el significado de la fecha 23/06/2027 de FDE1104: ¿es despacho de ABB a Vaca Muerta o cierre del proceso FAT completo?"),
            ("3.", "Establecer reporte mensual con estas 5 fechas clave de ABB como indicadores de estado."),
        ],
    }

    row = 3
    for titulo_bloque, bg, fg, texto_directo in BLOQUES:
        ws3.merge_cells(f"A{row}:D{row}")
        c = ws3.cell(row=row, column=1, value=titulo_bloque)
        c.fill = fill(bg); c.font = font(bold=True, size=11, color=fg)
        c.alignment = align("left"); ws3.row_dimensions[row].height = 26
        row += 1

        if texto_directo:
            ws3.merge_cells(f"A{row}:D{row}")
            c = ws3.cell(row=row, column=1, value=texto_directo)
            c.alignment = Alignment(wrap_text=True, vertical="top")
            c.font = font(bold=True, size=11)
            c.fill = fill(bg)
            ws3.row_dimensions[row].height = 60
            row += 1

        elif titulo_bloque in DETALLE_BLOQUES:
            for label, detalle in DETALLE_BLOQUES[titulo_bloque]:
                c1 = ws3.cell(row=row, column=1, value=label)
                c1.font = font(bold=True, size=10); c1.fill = fill(bg)
                c1.border = border(); c1.alignment = align()
                ws3.merge_cells(f"B{row}:D{row}")
                c2 = ws3.cell(row=row, column=2, value=detalle)
                c2.alignment = Alignment(wrap_text=True, vertical="top")
                c2.font = font(size=10)
                c2.border = border()
                ws3.row_dimensions[row].height = 55
                row += 1

        row += 1

    ws3.column_dimensions["A"].width = 30
    ws3.column_dimensions["B"].width = 28
    ws3.column_dimensions["C"].width = 28
    ws3.column_dimensions["D"].width = 28

    out = os.path.join(PLANNING_DIR, "Hitos_Generales_ABB_vs_P6.xlsx")
    wb.save(out)
    return out


if __name__ == "__main__":
    print("\n=== Comparativa Hitos Generales ABB vs P6 ===\n")
    out = generar()
    print(f"✅ Reporte: {out}")
    print("\nHojas:")
    print("  1. Comparativa General   — 5 provisiones ABB vs fecha de demanda P6")
    print("  2. Cadena Cronológica    — todos los hitos ordenados por fecha")
    print("  3. Conclusión Ejecutiva  — veredicto + alertas + recomendaciones")

    print("\n── RESUMEN EJECUTIVO ──────────────────────────────────────────")
    print("  Provisión           │ Fecha ABB (peor)  │ Fecha P6 necesita │ Holgura")
    print("  ─────────────────────────────────────────────────────────────────────")
    for prov, fa, fp, _, _, _ in [
        ("OC / Inicio",          date(2026,5,8),  date(2026,5,29), None,None,None),
        ("Freezing Ing. (S#3)",  date(2026,8,31), None,            None,None,None),
        ("FAT PMS",              date(2026,10,23),date(2027,6,2),  None,None,None),
        ("FAT Salas (S#3 tardío)",date(2027,4,19),date(2027,5,24), None,None,None),
        ("Despacho (S#3 tardío)",date(2027,5,14), date(2027,6,28), None,None,None),
    ]:
        delta = (fp - fa).days if fp else None
        hol = f"+{delta}d" if delta is not None else "Sin hito P6"
        print(f"  {prov:<25} │ {fa.strftime('%d/%m/%Y'):17} │ {fp.strftime('%d/%m/%Y') if fp else 'Sin hito en P6':17} │ {hol}")
    print()
    print("  ★ RFSU: 31/01/2028 — IDÉNTICO EN AMBOS CRONOGRAMAS")
