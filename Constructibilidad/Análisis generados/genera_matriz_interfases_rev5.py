#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MATRIZ DE INTERFASES — Provisiones críticas CPF-2 (Rev5, 2026-09-24).

Cambios respecto de Rev4:
 1) PROVISIONES
    · Se agrega ABB — DUCTOS DE BARRAS como provisión propia (hoy la última entrega del
      paquete: 08-07-27). Estaba fuera de la matriz pese a tener cronograma e interfase
      propia con AESA (layout de instalación + detalle de conexionado a trafo).
    · HIMA se separa en HPH (HIMA Paul Hilpert, Alemania — tableros de marshalling) y
      HARG (HIMA Argentina — adecuación, programación y FAT/IFAT), igual que en el
      cronograma ejecutivo. HARG sigue EN ADJUDICACIÓN: la OC no está colocada.
 2) TAREAS  (ver CAMBIOS_TAREAS)
    · ALTA   — «FAT en fábrica de origen / pruebas internas del proveedor» (se desdobla de
               la antigua «Ejecución de FAT»; en las Salas son la FAT de Brasil y la FAT
               del shelter, dos eventos distintos con responsables distintos).
    · ALTA   — «Nacionalización / despacho de aduana e internación»: es la interfase que
               concentra el riesgo de HPH y de los ductos y no figuraba en la matriz.
    · ALTA   — «Recepción y custodia en destino intermedio (shelterista / base de
               integración)»: punto de transferencia de custodia que hoy nadie tenía
               asignado formalmente.
    · ALTA   — «Despacho desde fábrica de origen» se separa de «Entrega en sitio»:
               en la Rev4 eran una sola fila con Incoterms y responsables distintos.
    · BAJA   — «SAT / pruebas en sitio» y «Precomisionado» se FUSIONAN: tenían códigos
               RACI idénticos y rangos superpuestos; en un resumen no aportaban distinción.
 3) FECHAS
    · Todos los rangos se rehacen contra el cronograma ejecutivo Rev5 (24-09-2026):
      ABB 5155-00-0009-VE-CO-PG-001_A (21-09-26), ABB PMS .mpp REV B, Inauco
      5155-00-0029-VE-CA-PG-002 Rev.0 (14-09-26) y HARG PSCH_CPF2_Rev3 con el
      corrimiento de +49 días por la OC no colocada.
 4) COHERENCIA
    · Se incorpora un VALIDADOR de precedencias que se ejecuta al generar el documento y
      cuyo resultado se publica en el propio documento (hoja/sección «Verificación de
      coherencia»). Comprueba 11 reglas de secuencia obligatoria por provisión.
Genera Excel, PDF (A3) y HTML.
"""
import os, json, textwrap
GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
REV="Rev5"; FECHA="2026-09-24"

ACT_SHORT=["Proveedor","Activación suministros","AESA Proy/Contrato","AESA Ingeniería",
           "AESA Construcciones","AESA Precom/Com","AESA Calidad (QA/QC)","Cliente PPSA"]
SECTOR_COLORS=["#455A64","#00838F","#1565C0","#5E35B1","#8D6E63","#2E7D32","#C62828","#EF6C00"]
RACI={'E':('1F3864','Ejecuta / responsable primario'),'A':('C62828','Aprueba / libera'),
      'R':('EF6C00','Revisa / comenta'),'P':('2E7D32','Participa / Asiste'),'I':('90A4AE','Informado')}
WEIGHT={'E':4,'A':2,'R':2,'P':3,'I':0}

# codes = [Proveedor, Activación, Proyecto, Ingeniería, Construcciones, Precom&Com, Calidad, Cliente]
ACTIVITIES=[
 ("Contrato","Gestión de la OC / contrato y coordinación",                       ['P','I','E','I','I','I','I','I']),
 ("Ingeniería","Emisión de documentación (IB / ID / IC)",                        ['E','I','I','R','I','I','I','I']),
 ("Ingeniería","Revisión y aprobación de documentación",                         ['P','I','I','A','R','I','R','I']),
 ("Ingeniería","Freezing / hold points · cierre de ingeniería",                  ['E','I','P','A','R','I','R','I']),
 ("Procura / Fabricación","Procura, acopio de materiales y fabricación",         ['E','P','R','I','I','I','R','I']),
 ("Procura / Fabricación","Certificación de hitos de pago / avance",             ['P','E','E','I','I','I','I','I']),
 ("Pruebas de fábrica","Procedimientos de FAT / pruebas (emisión y aprobación)",  ['E','I','I','R','I','P','A','I']),
 ("Pruebas de fábrica","FAT en fábrica de origen / pruebas internas del proveedor",['E','P','P','R','I','P','A','I']),
 ("Pruebas de fábrica","FAT de conjunto con el cliente (asistencia)",             ['E','P','P','P','I','P','A','P']),
 ("Pruebas de fábrica","iFAT / pruebas de integración entre provisiones",         ['E','I','P','P','P','P','R','P']),
 ("Logística","Despacho desde fábrica de origen (EXW / FCA / DAP)",               ['E','E','R','I','I','I','A','I']),
 ("Logística","Nacionalización / despacho de aduana e internación",               ['P','E','R','I','I','I','I','I']),
 ("Logística","Recepción y custodia en destino intermedio",                       ['E','P','R','I','P','I','R','I']),
 ("Despacho / Sitio","Entrega en sitio / recepción en obra",                      ['E','P','R','I','A','I','R','I']),
 ("Despacho / Sitio","Montaje / integración en sitio",                            ['P','I','R','R','E','I','R','I']),
 ("Precom. / PEM","SAT y precomisionado en sitio",                                ['P','I','P','P','P','E','R','P']),
 ("Precom. / PEM","Comisionado / Puesta en marcha (PEM)",                         ['P','I','P','P','I','E','R','P']),
 ("Precom. / PEM","Capacitación (operación / mantenimiento)",                     ['E','I','P','I','I','P','I','P']),
 ("Cierre","Documentación final / DataBook / As-built",                           ['E','I','I','A','P','P','A','I']),
]
ADJ_ROW=("Adjudicación","Proceso de adjudicación / cierre de contrato",['P','I','E','R','I','I','I','I'])
NACT=len(ACTIVITIES)   # 19

CAMBIOS_TAREAS=[
 ("ALTA","Desdoblada","«Ejecución de FAT / pruebas en fábrica» se separa en «FAT en fábrica de origen / pruebas internas del proveedor» y «FAT de conjunto con el cliente». En las Salas son dos eventos distintos (FAT de tableros en Brasil, nov-dic/26, y FAT del shelter completo, feb-mar/27) y el Cliente sólo participa en el segundo."),
 ("ALTA","Agregada","«Nacionalización / despacho de aduana e internación». Es la interfase que concentra el riesgo de HPH (PA-01 del cronograma ejecutivo) y de los ductos de barras (80 días de traslado + aduana), y no figuraba en la matriz. Ejecuta Activación de suministros."),
 ("ALTA","Agregada","«Recepción y custodia en destino intermedio (shelterista / base de integración)». Punto de transferencia de custodia — Mendoza para las salas, base de Inauco para los tableros de HIMA — que no tenía responsable asignado."),
 ("ALTA","Desdoblada","«Despacho y entrega en sitio» se separa en «Despacho desde fábrica de origen (EXW/FCA/DAP)» y «Entrega en sitio / recepción en obra». Tienen Incoterm, fecha y responsable de recepción distintos: en la entrega en sitio quien aprueba es AESA Construcciones."),
 ("MEDIA","Renombrada","«Procura y fabricación (seguimiento)» pasa a «Procura, acopio de materiales y fabricación», para que el acopio —que el cliente sigue como indicador de avance— quede explícito."),
 ("BAJA","Fusionadas","«SAT / pruebas en sitio» y «Precomisionado» se unifican en «SAT y precomisionado en sitio»: tenían códigos RACI idénticos y rangos superpuestos, de modo que en un resumen de tareas no aportaban distinción."),
 ("—","Sin cambio","Se mantienen sin cambios las filas de Contrato, Emisión / Revisión / Freezing de ingeniería, Certificación de hitos de pago, Procedimientos de FAT, iFAT, Montaje en sitio, PEM, Capacitación y Documentación final."),
]

PROV=[
 ("SALAS","ABB — Salas Eléctricas SE#3 y SE#4","ABB","Adjudicado — OC vigente + planificación","C62828",False),
 ("DUCTOS","ABB — Ductos de barras","ABB","Adjudicado — alcance separado del mismo contrato","795548",False),
 ("PMS","ABB — PMS (Power Management System)","ABB","Adjudicado — OC vigente + planificación","EF6C00",False),
 ("INAUCO","Inauco — PCS / SCADA / Comunicaciones","Inauco","Adjudicado — Obra 129-008","1565C0",False),
 ("HPH","HPH — HIMA Paul Hilpert · tableros de marshalling (ESD/F&G/PSS)","HIMA (Alemania)","Adjudicado","6A1B9A",False),
 ("HARG","HARG — HIMA Argentina · adecuación, programación y FAT/IFAT","HIMA (Argentina)","EN ADJUDICACIÓN — OC NO COLOCADA al 24-09-26","00838F",True),
]

# ---- Rangos por actividad (19 por provisión), anclados al cronograma ejecutivo Rev5 ----
NA=["",""]   # no aplica
RANGES={
 # 0 Contrato | 1 Emisión | 2 Revisión | 3 Freezing | 4 Procura/fabr | 5 Certificación
 # 6 Proced.FAT | 7 FAT origen | 8 FAT conjunto | 9 iFAT | 10 Despacho fábrica
 # 11 Nacionalización | 12 Recepción destino interm. | 13 Entrega en sitio | 14 Montaje
 # 15 SAT/precom | 16 PEM | 17 Capacitación | 18 Doc. final
 "SALAS":[["2026-05","2027-12"],["2026-05","2026-09"],["2026-09","2026-10"],["2026-09","2026-10"],
          ["2026-05","2027-03"],["2026-05","2027-04"],["2026-09","2026-11"],["2026-11","2026-12"],
          ["2027-02","2027-03"],["2027-03","2027-04"],["2026-11","2026-12"],["2026-12","2027-01"],
          ["2026-11","2027-01"],["2027-03","2027-04"],["2027-04","2027-07"],["2027-06","2027-10"],
          ["2027-09","2027-12"],["2027-10","2027-11"],["2027-05","2027-12"]],
 "DUCTOS":[["2026-05","2027-12"],["2026-07","2026-12"],["2026-09","2026-12"],["2026-09","2026-12"],
          ["2026-12","2027-04"],["2026-05","2027-07"],["2027-02","2027-03"],["2027-03","2027-04"],
          NA,NA,["2027-04","2027-04"],["2027-04","2027-07"],
          NA,["2027-07","2027-07"],["2027-07","2027-09"],["2027-09","2027-11"],
          ["2027-10","2027-12"],NA,["2027-07","2027-12"]],
 "PMS":  [["2026-05","2027-12"],["2026-06","2026-08"],["2026-07","2026-09"],["2026-09","2026-09"],
          ["2026-05","2026-12"],["2026-05","2027-02"],["2026-10","2026-10"],["2026-12","2027-01"],
          ["2027-01","2027-01"],["2026-09","2027-04"],["2027-02","2027-02"],["2027-02","2027-02"],
          ["2027-02","2027-02"],["2027-04","2027-07"],["2027-04","2027-08"],["2027-07","2027-11"],
          ["2027-10","2027-12"],["2027-11","2027-12"],["2027-06","2027-12"]],
 "INAUCO":[["2026-07","2027-12"],["2026-08","2026-11"],["2026-09","2026-11"],["2026-11","2026-11"],
          ["2026-08","2027-02"],["2026-07","2027-05"],["2027-02","2027-02"],["2027-01","2027-03"],
          ["2027-03","2027-04"],["2026-10","2027-04"],["2027-05","2027-05"],NA,
          ["2027-01","2027-03"],["2027-05","2027-05"],["2027-05","2027-08"],["2027-07","2027-11"],
          ["2027-10","2027-12"],["2027-11","2027-12"],["2027-06","2027-12"]],
 "HPH":  [["2026-07","2027-12"],["2026-08","2026-09"],["2026-09","2026-09"],["2026-09","2026-09"],
          ["2026-09","2026-11"],["2026-07","2027-01"],["2026-09","2026-10"],["2026-11","2026-11"],
          ["2027-03","2027-04"],["2027-04","2027-04"],["2026-11","2026-12"],["2026-12","2027-01"],
          ["2027-01","2027-01"],["2027-05","2027-05"],["2027-05","2027-08"],["2027-07","2027-11"],
          ["2027-10","2027-12"],["2027-11","2027-12"],["2027-06","2027-12"]],
 "HARG": [["2026-11","2027-12"],["2026-11","2026-12"],["2026-11","2026-12"],["2026-12","2026-12"],
          ["2026-11","2027-03"],["2026-11","2027-05"],["2027-03","2027-04"],NA,
          ["2027-04","2027-04"],["2027-04","2027-04"],NA,NA,
          ["2027-02","2027-03"],["2027-05","2027-05"],["2027-05","2027-08"],["2027-07","2027-11"],
          ["2027-10","2027-12"],["2027-11","2027-12"],["2027-06","2027-12"]],
}
ADJ_RANGE={"HARG":["2026-09","2026-10"]}

# ---- Validador de precedencias -------------------------------------------------------
# (idx_predecesora, idx_sucesora, modo)  modo: 'S'=inicio · 'F'=fin · 'SF'=ambos
PRECEDENCIAS=[
 (1,2,'S',"La revisión no puede iniciar antes de la emisión de documentación"),
 (2,3,'SF',"El freezing point no puede cerrar antes de la aprobación"),
 (3,4,'F',"La fabricación no puede terminar antes del cierre de ingeniería"),
 (6,8,'S',"La FAT con el cliente no puede iniciar antes de aprobar sus procedimientos"),
 (7,10,'F',"El despacho de fábrica no puede cerrar antes de las pruebas de fábrica"),
 (10,11,'S',"La nacionalización no puede iniciar antes del despacho"),
 (8,9,'F',"La integración entre provisiones cierra con la FAT más tardía"),
 (10,13,'SF',"La entrega en sitio no puede ocurrir antes del despacho de fábrica"),
 (13,14,'S',"El montaje no puede iniciar antes de la entrega en sitio"),
 (14,15,'S',"El SAT / precomisionado no puede iniciar antes del montaje"),
 (15,16,'SF',"La puesta en marcha no puede iniciar ni cerrar antes del precomisionado"),
 (16,18,'F',"La documentación final cierra después de la puesta en marcha"),
]
def _ok(a,b): return a<=b
def validar():
    """Devuelve (filas_resultado, incidencias). Verifica precedencias y envolvente de contrato."""
    filas=[]; inc=[]
    for (key,name,ven,estado,col,adj) in PROV:
        R=RANGES[key]; fallas=0; verif=0
        for (pi,si,modo,desc) in PRECEDENCIAS:
            p,s=R[pi],R[si]
            if not p[0] or not s[0]: continue    # alguna no aplica
            verif+=1
            bad=[]
            if 'S' in modo and not _ok(p[0],s[0]): bad.append(f"inicio {s[0]} < {p[0]}")
            if 'F' in modo and not _ok(p[1],s[1]): bad.append(f"fin {s[1]} < {p[1]}")
            if bad:
                fallas+=1
                inc.append((name,ACTIVITIES[pi][1],ACTIVITIES[si][1],desc,"; ".join(bad)))
        # envolvente: el contrato debe cubrir todas las actividades
        act=[r for i,r in enumerate(R) if i!=0 and r[0]]
        mn=min(r[0] for r in act); mx=max(r[1] for r in act); verif+=1
        if R[0][0]>mn or R[0][1]<mx:
            fallas+=1
            inc.append((name,"Gestión de la OC / contrato","(envolvente)",
                        "El rango del contrato debe cubrir todas las actividades de la provisión",
                        f"contrato {R[0][0]}→{R[0][1]} vs actividades {mn}→{mx}"))
        na=sum(1 for r in R if not r[0])
        filas.append((name,verif,fallas,na))
    return filas,inc

NOTAS=[
 "Rangos Inicio/Fin por actividad tomados del CRONOGRAMA EJECUTIVO Rev5 (24-09-2026), que consolida: ABB Salas y Ductos 5155-00-0009-VE-CO-PG-001_A (21-09-26), ABB PMS .mpp REV B, Inauco 5155-00-0029-VE-CA-PG-002 Rev.0 (14-09-26), HPH 103618-Schedule (16-08-26) y HARG PSCH_CPF2_Rev3 con corrimiento de +49 días.",
 "Hasta la fila «Entrega en sitio» las fechas provienen de los cronogramas de los proveedores. Desde «Montaje / integración en sitio» en adelante son la VENTANA DE OBRA de AESA (estimación EPC): los cronogramas de provisión cortan en la entrega.",
 "HARG (HIMA Argentina) sigue EN ADJUDICACIÓN: la OC no está colocada al 24-09-2026. Sus fechas incluyen el corrimiento de 49 días y se revisarán al colocar la orden.",
 "«—» indica actividad NO APLICABLE a esa provisión (p. ej. nacionalización en Inauco, de fabricación nacional; o FAT de conjunto e iFAT en los ductos de barras).",
 "La fila «iFAT / pruebas de integración entre provisiones» puede INICIAR antes de la FAT con el cliente: hay pruebas de integración anticipadas (maqueta PMS/CCM/PCS sep-26 y PCS↔PMS en Bs.As. oct-nov-26). Lo que sí se verifica es que CIERRE con la FAT más tardía.",
 "Activación de suministros: E en Certificación de hitos de pago, en Despacho desde fábrica y en Nacionalización; P (asiste) en Procura/Fabricación, FAT de fábrica y recepción.",
 "AESA — Precom&Comisionado y AESA — Calidad (QA/QC) pertenecen al Depto. de Calidad Integral, separados por función.",
 "AESA — Calidad (QA/QC): revisa/aprueba/libera según estadios de fabricación y despacho; aprueba/revisa la documentación de ensayos.",
 "Cliente PPSA: informado en general; participa en la FAT de conjunto, en el iFAT y en precomisionado/PEM. En la FAT en fábrica de origen (Brasil / Alemania) figura como informado — a confirmar si PPSA desea asistir.",
 "Fechas de referencia a nivel mensual; se refinarán con nuevas versiones de los cronogramas por contrato.",
]
MES3={'01':'Ene','02':'Feb','03':'Mar','04':'Abr','05':'May','06':'Jun','07':'Jul','08':'Ago','09':'Sep','10':'Oct','11':'Nov','12':'Dic'}
def fmt(iso):
    if not iso: return "—"
    y,m=iso.split('-'); return f"{MES3[m]}-{y[2:]}"
def rng(key,adj,idx):
    if adj and idx==-1: return ADJ_RANGE.get(key,["",""])
    r=RANGES.get(key,[])
    return r[idx] if 0<=idx<len(r) else ["",""]

# ---------- EXCEL ----------
def build_xlsx():
    import openpyxl
    from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
    thin=Side(style="thin",color="D9D9D9"); B=Border(thin,thin,thin,thin); NAVY="1F3864"; HEAD="37474F"
    def cell(ws,r,c,v,fill=None,color="000000",bold=False,size=9,center=False,wrap=False):
        cc=ws.cell(r,c,v); cc.font=Font(bold=bold,color=color,size=size)
        if fill: cc.fill=PatternFill("solid",fgColor=fill)
        cc.alignment=Alignment(horizontal="center" if center else "left",vertical="center",wrap_text=wrap); cc.border=B; return cc
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Matriz de Interfases"; NC=12
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=NC); cell(ws,1,1,f"MATRIZ DE INTERFASES / RESPONSABILIDADES — PROVISIONES CRÍTICAS CPF-2 · {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); ws.row_dimensions[1].height=22
    r=4; cell(ws,r,1,"Leyenda:",None,"000000",True,9)
    for i,(k,(col,desc)) in enumerate(RACI.items()): cell(ws,r,2+i,f"{k} = {desc}",col,"FFFFFF" if k!='I' else "000000",True,8,True,True)
    r=6
    def header(r):
        for j,h in enumerate(["Fase","Actividad / interfase","Inicio","Fin"]): cell(ws,r,1+j,h,HEAD,"FFFFFF",True,8.5,True)
        for j,a in enumerate(ACT_SHORT): cell(ws,r,5+j,a,HEAD,"FFFFFF",True,7.6,True,True)
        ws.row_dimensions[r].height=30
    for (key,name,ven,estado,col,adj) in PROV:
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC); cell(ws,r,1,f"▶  {name}   ·   {ven}   ·   {estado}",col,"FFFFFF",True,10); ws.row_dimensions[r].height=18; r+=1
        header(r); r+=1
        acts=([ADJ_ROW]+ACTIVITIES) if adj else ACTIVITIES
        for ai,(fase,act,codes) in enumerate(acts):
            idx=(ai-1) if adj else ai; ri = -1 if (adj and ai==0) else idx
            rr=rng(key,adj,ri);
            cell(ws,r,1,fase,None,"333333",False,8); cell(ws,r,2,act,None,"000000",False,8.5,False,True)
            cell(ws,r,3,fmt(rr[0]),None,"1F3864",False,8,True); cell(ws,r,4,fmt(rr[1]),None,"1F3864",False,8,True)
            na = not rr[0]          # actividad no aplicable a esta provisión
            for j,code in enumerate(codes):
                if na: cell(ws,r,5+j,"—","F2F2F2","BDBDBD",False,9,True)
                else:
                    fc,_=RACI[code]; cell(ws,r,5+j,code,fc,"FFFFFF" if code!='I' else "000000",True,9,True)
            r+=1
        r+=1
    for c,w in zip("ABCDEFGHIJKL",[18,42,10,10,11,12,12,12,12,12,12,11]): ws.column_dimensions[c].width=w
    ws.freeze_panes="A7"
    # --- hoja: cambios de tareas Rev4 -> Rev5 ---
    cs=wb.create_sheet("Cambios de tareas")
    cell(cs,1,1,f"AJUSTES DE TAREAS  Rev4 → {REV}",NAVY,"FFFFFF",True,12); cs.merge_cells("A1:C1")
    cell(cs,2,1,"Revisión del set de actividades del resumen: qué se agregó, se desdobló, se fusionó o se renombró.",None,"555555",False,9); cs.merge_cells("A2:C2")
    for j,h in enumerate(["Impacto","Tipo de cambio","Detalle"]): cell(cs,4,1+j,h,HEAD,"FFFFFF",True,9,True,True)
    IMPF={"ALTA":"C62828","MEDIA":"EF6C00","BAJA":"2E7D32","—":"90A4AE"}
    for i,(imp,tipo,det) in enumerate(CAMBIOS_TAREAS):
        rr=5+i
        cell(cs,rr,1,imp,IMPF[imp],"FFFFFF",True,9,True)
        cell(cs,rr,2,tipo,None,"000000",True,9,True)
        cell(cs,rr,3,det,None,"000000",False,9,False,True)
        cs.row_dimensions[rr].height=max(16,13*(1+len(det)//110))
    for c,w in zip("ABC",[11,15,140]): cs.column_dimensions[c].width=w

    # --- hoja: verificación de coherencia ---
    vs=wb.create_sheet("Verificación de coherencia")
    cell(vs,1,1,f"VERIFICACIÓN DE COHERENCIA DE RANGOS  ·  {REV} · {FECHA}",NAVY,"FFFFFF",True,12); vs.merge_cells("A1:D1")
    cell(vs,2,1,"Se comprueba que cada actividad respete el rango de sus antecesoras. Las actividades no aplicables a una provisión se omiten de la verificación.",None,"555555",False,9); vs.merge_cells("A2:D2")
    filas,inc=validar()
    for j,h in enumerate(["Provisión","Reglas verificadas","Incoherencias","Actividades N/A"]): cell(vs,4,1+j,h,HEAD,"FFFFFF",True,9,True,True)
    for i,(nm,ver,fal,na) in enumerate(filas):
        rr=5+i
        cell(vs,rr,1,nm,None,"000000",False,9,False,True)
        cell(vs,rr,2,ver,None,"000000",False,9,True)
        cell(vs,rr,3,fal,("C62828" if fal else "2E7D32"),"FFFFFF",True,9,True)
        cell(vs,rr,4,na,None,"555555",False,9,True)
    rr=6+len(filas)
    tot=sum(f[1] for f in filas); bad=sum(f[2] for f in filas)
    cell(vs,rr,1,f"RESULTADO: {tot} reglas verificadas · {bad} incoherencias",("C62828" if bad else "2E7D32"),"FFFFFF",True,10); vs.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=4)
    rr+=2
    cell(vs,rr,1,"Reglas de precedencia aplicadas",None,NAVY,True,10); rr+=1
    for (pi,si,modo,desc) in PRECEDENCIAS:
        cell(vs,rr,1,f"• {ACTIVITIES[pi][1]}  →  {ACTIVITIES[si][1]}   [{ 'inicio y fin' if modo=='SF' else ('inicio' if modo=='S' else 'fin') }]",None,"000000",False,8.5,False,True)
        vs.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=4); rr+=1
        cell(vs,rr,1,f"     {desc}",None,"555555",False,8,False,True)
        vs.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=4); rr+=1
    if inc:
        rr+=1; cell(vs,rr,1,"Incoherencias detectadas",None,"C62828",True,10); rr+=1
        for (nm,a,b,desc,det) in inc:
            cell(vs,rr,1,f"• {nm} — {a} → {b}: {det}",None,"C62828",False,8.5,False,True)
            vs.merge_cells(start_row=rr,start_column=1,end_row=rr,end_column=4); rr+=1
    for c,w in zip("ABCD",[62,18,15,16]): vs.column_dimensions[c].width=w

    ns=wb.create_sheet("Notas")
    for i,n in enumerate(NOTAS): cell(ns,2+i,1,f"• {n}",None,"000000",False,9,False,True); ns.row_dimensions[2+i].height=30
    ns.column_dimensions['A'].width=155
    path=os.path.join(GEN,f"Matriz de Interfases - Provisiones criticas - {REV} - {FECHA}.xlsx"); wb.save(path); return path

# ---------- PDF ----------
def build_pdf():
    from reportlab.lib.pagesizes import A3, landscape
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    path=os.path.join(GEN,f"Matriz de Interfases - Provisiones criticas - {REV} - {FECHA}.pdf")
    W,H=landscape(A3); c=canvas.Canvas(path,pagesize=(W,H)); NAVY=colors.HexColor('#1F3864'); x0=10*mm; x1=W-10*mm
    def banner():
        c.setFillColor(NAVY); c.rect(0,H-13*mm,W,13*mm,fill=1,stroke=0); c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold",13); c.drawString(x0,H-9*mm,"MATRIZ DE INTERFASES / RESPONSABILIDADES — PROVISIONES CRÍTICAS CPF-2")
        c.setFont("Helvetica",9); c.drawRightString(x1,H-9*mm,f"{REV} · {FECHA}")
    banner(); y=H-20*mm
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",11); c.drawString(x0,y,"Leyenda RACI ampliado"); y-=6*mm; lx=x0
    for k,(col,desc) in RACI.items():
        c.setFillColor(colors.HexColor('#'+col)); c.rect(lx,y-4.2*mm,5.2*mm,5.2*mm,fill=1,stroke=0); c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold",8.5); c.drawString(lx+6.5*mm,y-2.6*mm,k); c.setFont("Helvetica",7.6); c.drawString(lx+10.5*mm,y-2.6*mm,desc); lx+=62*mm
    y-=8*mm; c.setFillColor(colors.black); c.setFont("Helvetica",8)
    for n in NOTAS:
        for j,ln in enumerate(textwrap.wrap(n,235)): c.drawString(x0+(0 if j==0 else 3*mm),y,("• " if j==0 else "")+ln); y-=4*mm
        y-=0.6*mm
    aw=[19*mm,58*mm,22*mm,22*mm]+[ (x1-x0-121*mm)/8 ]*8
    FSHORT={"Procura / Fabricación":"Procura/Fabric.","Pruebas de fábrica":"Pruebas FAT","Precom. / PEM":"Precom./PEM","Despacho / Sitio":"Despacho/Sitio"}
    ASHORT={"Emisión de documentación (IB / ID / IC)":"Emisión de doc. (IB/ID/IC)","Revisión y aprobación de documentación":"Revisión y aprobación de doc.","Freezing / hold points · cierre de ingeniería":"Freezing/hold points · cierre ing.","Procura, acopio de materiales y fabricación":"Procura, acopio y fabricación","Certificación de hitos de pago / avance":"Certif. hitos de pago / avance","Procedimientos de FAT / pruebas (emisión y aprobación)":"Procedimientos FAT (emisión/aprob.)","FAT en fábrica de origen / pruebas internas del proveedor":"FAT en fábrica de origen / pruebas int.","FAT de conjunto con el cliente (asistencia)":"FAT de conjunto con el cliente","iFAT / pruebas de integración entre provisiones":"iFAT / integración entre provisiones","Despacho desde fábrica de origen (EXW / FCA / DAP)":"Despacho de fábrica (EXW/FCA/DAP)","Nacionalización / despacho de aduana e internación":"Nacionalización / aduana","Recepción y custodia en destino intermedio":"Recepción en destino intermedio","Entrega en sitio / recepción en obra":"Entrega en sitio / recepción en obra","Montaje / integración en sitio":"Montaje / integración en sitio","SAT y precomisionado en sitio":"SAT y precomisionado en sitio","Comisionado / Puesta en marcha (PEM)":"Comisionado / PEM","Capacitación (operación / mantenimiento)":"Capacitación (op./mant.)","Documentación final / DataBook / As-built":"Doc. final / DataBook / As-built","Gestión de la OC / contrato y coordinación":"Gestión de OC / contrato"}
    def mhead(y):
        cx=x0; c.setFillColor(colors.HexColor('#37474F')); c.rect(x0,y-9*mm,sum(aw),9*mm,fill=1,stroke=0); c.setFillColor(colors.white); c.setFont("Helvetica-Bold",6.8)
        for l,w in zip(["Fase","Actividad / interfase","Inicio","Fin"]+ACT_SHORT,aw):
            for j,ln in enumerate(textwrap.wrap(l,int(w/1.45))[:2]): c.drawString(cx+1.1*mm,y-3.6*mm-j*3.1*mm,ln)
            cx+=w
        return y-9*mm
    def ensure(y,need):
        if y-need<12*mm: c.showPage(); banner(); return H-18*mm,True
        return y,False
    c.showPage(); banner(); y=H-18*mm
    for (key,name,ven,estado,col,adj) in PROV:
        y,_=ensure(y,26*mm)
        c.setFillColor(colors.HexColor('#'+col)); c.rect(x0,y-6*mm,sum(aw),6*mm,fill=1,stroke=0); c.setFillColor(colors.white); c.setFont("Helvetica-Bold",9)
        c.drawString(x0+2*mm,y-4.3*mm,f"{name}   ·   {ven}   ·   {estado}"); y-=6*mm; y=mhead(y)
        acts=([ADJ_ROW]+ACTIVITIES) if adj else ACTIVITIES
        for ai,(fase,act,codes) in enumerate(acts):
            y,newp=ensure(y,6*mm)
            if newp: y=mhead(y)
            ri = -1 if (adj and ai==0) else ((ai-1) if adj else ai)
            rr=rng(key,adj,ri); rowh=6.2*mm; cx=x0
            c.setFillColor(colors.HexColor('#F7F9FB') if ai%2 else colors.white); c.rect(x0,y-rowh,sum(aw),rowh,fill=1,stroke=0)
            c.setFillColor(colors.HexColor('#444444')); c.setFont("Helvetica",6.6); c.drawString(cx+1.2*mm,y-4*mm,FSHORT.get(fase,fase)[:22]); cx+=aw[0]
            c.setFillColor(colors.black); c.setFont("Helvetica",6.8); c.drawString(cx+1.2*mm,y-4*mm,ASHORT.get(act,act)[:44]); cx+=aw[1]
            c.setFillColor(colors.HexColor('#1F3864')); c.setFont("Helvetica",6.6)
            c.drawCentredString(cx+aw[2]/2,y-4*mm,fmt(rr[0])); cx+=aw[2]
            c.drawCentredString(cx+aw[3]/2,y-4*mm,fmt(rr[1])); cx+=aw[3]
            na = not rr[0]
            for code in codes:
                if na:
                    c.setFillColor(colors.HexColor('#F2F2F2')); c.roundRect(cx+aw[4]/2-3.4*mm,y-rowh+1*mm,6.8*mm,rowh-2*mm,1,fill=1,stroke=0)
                    c.setFillColor(colors.HexColor('#BDBDBD')); c.setFont("Helvetica",7.2); c.drawCentredString(cx+aw[4]/2,y-4*mm,"—")
                else:
                    fc,_=RACI[code]; c.setFillColor(colors.HexColor('#'+fc)); c.roundRect(cx+aw[4]/2-3.4*mm,y-rowh+1*mm,6.8*mm,rowh-2*mm,1,fill=1,stroke=0)
                    c.setFillColor(colors.white if code!='I' else colors.black); c.setFont("Helvetica-Bold",7.2); c.drawCentredString(cx+aw[4]/2,y-4*mm,code)
                cx+=aw[4]
            y-=rowh
        y-=4*mm
    # --- página de cambios de tareas y verificación de coherencia ---
    c.showPage(); banner(); y=H-20*mm
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",12); c.drawString(x0,y,f"Ajustes de tareas  Rev4 → {REV}"); y-=7*mm
    IMPC={"ALTA":'#C62828',"MEDIA":'#EF6C00',"BAJA":'#2E7D32',"—":'#90A4AE'}
    for (imp,tipo,det) in CAMBIOS_TAREAS:
        c.setFillColor(colors.HexColor(IMPC[imp])); c.roundRect(x0,y-4.4*mm,13*mm,5*mm,1,fill=1,stroke=0)
        c.setFillColor(colors.white); c.setFont("Helvetica-Bold",6.8); c.drawCentredString(x0+6.5*mm,y-3*mm,imp)
        c.setFillColor(colors.black); c.setFont("Helvetica-Bold",8); c.drawString(x0+15*mm,y-3*mm,tipo)
        c.setFont("Helvetica",7.8)
        for j,ln in enumerate(textwrap.wrap(det,190)):
            c.drawString(x0+38*mm,y-3*mm-j*3.6*mm,ln)
        y-=max(6.5*mm,3.6*mm*len(textwrap.wrap(det,190))+3*mm)
    y-=5*mm
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",12); c.drawString(x0,y,"Verificación de coherencia de rangos"); y-=6*mm
    filas,inc=validar(); tot=sum(f[1] for f in filas); bad=sum(f[2] for f in filas)
    c.setFillColor(colors.HexColor('#2E7D32' if not bad else '#C62828')); c.setFont("Helvetica-Bold",9)
    c.drawString(x0,y,f"{tot} reglas de precedencia verificadas sobre {len(filas)} provisiones · {bad} incoherencias"); y-=6*mm
    c.setFillColor(colors.black); c.setFont("Helvetica",7.8)
    for (nm,ver,fal,na) in filas:
        c.setFillColor(colors.HexColor('#2E7D32' if not fal else '#C62828')); c.circle(x0+1.6*mm,y-1.4*mm,1.5*mm,fill=1,stroke=0)
        c.setFillColor(colors.black); c.drawString(x0+6*mm,y-2.4*mm,f"{nm}  —  {ver} reglas verificadas, {fal} incoherencias, {na} actividades no aplicables"); y-=5*mm
    y-=3*mm; c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9.5); c.drawString(x0,y,"Reglas aplicadas"); y-=5*mm
    c.setFillColor(colors.black); c.setFont("Helvetica",7.4)
    for (pi,si,modo,desc) in PRECEDENCIAS:
        et='inicio y fin' if modo=='SF' else ('inicio' if modo=='S' else 'fin')
        c.drawString(x0,y,f"• {ACTIVITIES[pi][1]}  →  {ACTIVITIES[si][1]}   [{et}] — {desc}"); y-=4.2*mm
    for (nm,a,b,desc,det) in inc:
        c.setFillColor(colors.HexColor('#C62828')); c.drawString(x0,y,f"⚠ {nm}: {a} → {b} — {det}"); y-=4.2*mm
    c.showPage(); c.save(); return path

# ---------- HTML ----------
def build_html():
    contracts=[{"key":k,"name":n,"ven":v,"estado":e,"color":"#"+c,"adj":adj} for (k,n,v,e,c,adj) in PROV]
    acts=[{"fase":f,"act":a,"codes":cd} for (f,a,cd) in ACTIVITIES]
    payload=dict(contracts=contracts,acts=acts,adj={"fase":ADJ_ROW[0],"act":ADJ_ROW[1],"codes":ADJ_ROW[2]},
                 sectors=ACT_SHORT,sectorColors=SECTOR_COLORS,
                 raci={k:{"c":"#"+RACI[k][0],"d":RACI[k][1]} for k in RACI},weight=WEIGHT,
                 ranges=RANGES,adjRange=ADJ_RANGE,
                 cambios=[{"imp":i,"tipo":t,"det":d} for (i,t,d) in CAMBIOS_TAREAS],
                 check=[{"nm":n,"ver":v,"fal":f,"na":a} for (n,v,f,a) in validar()[0]],
                 notas=NOTAS)
    html=HTML_TMPL.replace("__REV__",REV).replace("__FECHA__",FECHA).replace("__DATA__",json.dumps(payload,ensure_ascii=False))
    path=os.path.join(GEN,f"Mapa de calor Interfases - {REV} - {FECHA}.html"); open(path,"w",encoding="utf-8").write(html); return path

HTML_TMPL=r"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mapa de calor de interfases · CPF-2</title>
<style>
:root{--bg:#f4f6f8;--card:#fff;--ink:#1f2933;--muted:#607080;--line:#e2e8f0;--navy:#1F3864;--shadow:0 1px 3px rgba(0,0,0,.08)}
@media(prefers-color-scheme:dark){:root{--bg:#0f141a;--card:#1a222c;--ink:#e6edf3;--muted:#93a1b0;--line:#2a343f;--shadow:0 1px 3px rgba(0,0,0,.5)}}
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--ink);font-size:13.5px}
header{background:var(--navy);color:#fff;padding:15px 22px}h1{margin:0;font-size:19px}.sub{opacity:.85;font-size:12px;margin-top:3px}
.wrap{max-width:1520px;margin:0 auto;padding:16px 22px 60px}
.toolbar{display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin:12px 0}
select{padding:7px 11px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font-size:13px;font-weight:600}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;box-shadow:var(--shadow);margin:14px 0;overflow-x:auto}
.card h2{margin:0 0 4px;font-size:15px}.card .d{color:var(--muted);font-size:12px;margin-bottom:12px}
table{border-collapse:collapse;font-size:11.5px}th,td{border:1px solid var(--line);padding:0}
.hz th{background:rgba(127,127,127,.06);color:var(--muted);font-size:10.5px;font-weight:700;padding:5px 6px;text-align:center;vertical-align:bottom}
.rowlab{text-align:left!important;padding:5px 8px!important;font-weight:600;white-space:normal;line-height:1.2;max-width:260px}
.cellcode{width:38px;height:30px;text-align:center;font-weight:800;color:#fff;font-size:11px}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin:6px 0 2px;font-size:12px}.lg{display:flex;align-items:center;gap:6px}.sw{width:14px;height:14px;border-radius:3px;display:inline-block}
.tcell{width:34px;height:28px;text-align:center;font-size:9.5px;font-weight:700}
.bar{display:flex;align-items:center;gap:8px;margin:5px 0;font-size:12.5px}.bar .nm{width:180px;flex-shrink:0}.bar .bw{flex:1;background:var(--line);border-radius:5px;height:18px;overflow:hidden;max-width:560px}.bar .bf{display:block;height:100%;border-radius:5px;min-width:3px}.bar .vv{width:44px;text-align:right;font-weight:800}
.scalebar{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--muted)}
.grad{width:170px;height:12px;border-radius:3px;background:linear-gradient(90deg,#ffffcc,#fed976,#feb24c,#fd8d3c,#f03b20,#bd0026)}
</style></head><body>
<header><h1>Mapa de calor de interfases — Involucramiento y carga temporal</h1>
<div class="sub">CPF-2 La Calera II · __REV__ · __FECHA__ · Rangos según el cronograma ejecutivo Rev5 (ABB Salas y Ductos 21-09-26 · ABB PMS · Inauco PG-002 Rev.0 · HPH · HARG +49 d)</div></header>
<div class="wrap">
 <div class="toolbar"><b>Contrato:</b><select id="csel"></select><span class="legend" id="raci-legend"></span></div>
 <div class="card"><h2>A · Involucramiento por tarea</h2><div class="d">Cómo participa cada sector en cada actividad, con el rango temporal de cada una. E/A/R/P = involucramiento; I (informado) se muestra tenue; «—» = actividad no aplicable a esta provisión.</div><div id="heatA"></div></div>
 <div class="card"><h2>B · Índice de involucramiento por sector</h2><div class="d">Carga relativa de cada sector en el contrato (peso E=4 · P=3 · R=2 · A=2 · I=0). Color por sector.</div><div id="barsB"></div></div>
 <div class="card"><h2>C · Carga temporal (sector × mes)</h2><div class="d">Cuándo se afecta cada sector (intensidad = involucramiento activo ese mes). <span class="scalebar">menos <span class="grad"></span> más — picos en rojo</span></div><div id="heatC"></div></div>
 <div class="card"><h2>D · Ajustes de tareas Rev4 → __REV__</h2><div class="d">Revisión del set de actividades del resumen: qué se agregó, se desdobló, se fusionó o se renombró.</div><div id="cambios"></div></div>
 <div class="card"><h2>E · Verificación de coherencia de rangos</h2><div class="d">Cada actividad debe respetar el rango de sus antecesoras. Las actividades no aplicables a una provisión se omiten de la verificación.</div><div id="check"></div></div>
 <div class="card"><h2>Notas y criterios</h2><div id="notas"></div></div>
</div>
<script>
const D=__DATA__;const RA=D.raci,W=D.weight,SEC=D.sectors,SC=D.sectorColors;
const csel=document.getElementById('csel');
csel.innerHTML=D.contracts.map(c=>`<option value="${c.key}">${c.name} — ${c.estado}</option>`).join('')+'<option value="__ALL__">TODOS (agregado)</option>';
document.getElementById('raci-legend').innerHTML=Object.entries(RA).map(([k,v])=>`<span class="lg"><span class="sw" style="background:${v.c}"></span>${k} · ${v.d}</span>`).join('');
const MON=[];{let y=2026,m=5;for(let i=0;i<22;i++){MON.push(y+'-'+String(m).padStart(2,'0'));m++;if(m>12){m=1;y++;}}}
const MLAB=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const mlabel=iso=>{const[y,m]=iso.split('-');return MLAB[+m-1]+"'"+y.slice(2);};
function actsOf(ck){const c=D.contracts.find(x=>x.key===ck);return (c&&c.adj)?[D.adj].concat(D.acts):D.acts;}
function rangeOf(ck,ai,isAdj){ // ai index within acts list (0=adj if adj)
 const c=D.contracts.find(x=>x.key===ck);
 if(c&&c.adj){ if(ai===0) return (D.adjRange[ck]||["",""]); return (D.ranges[ck]||[])[ai-1]||["",""]; }
 return (D.ranges[ck]||[])[ai]||["",""];
}
function mrange(r){if(!r||!r[0]||!r[1])return[];const s=MON.indexOf(r[0]),e=MON.indexOf(r[1]);if(s<0||e<0)return[];const o=[];for(let i=s;i<=e;i++)o.push(i);return o;}
function renderA(ck){
 const acts=actsOf(ck), c0=D.contracts.find(x=>x.key===ck);
 let h='<table class="hz"><thead><tr><th class="rowlab">Actividad</th><th>Inicio</th><th>Fin</th>'+SEC.map(s=>`<th>${s.replace(/ /g,'<br>')}</th>`).join('')+'</tr></thead><tbody>';
 acts.forEach((a,ai)=>{const r=rangeOf(ck,ai,c0&&c0.adj), na=!r||!r[0];
  h+=`<tr><td class="rowlab" title="${a.fase}">${a.act}</td>`+
     `<td style="text-align:center;padding:4px;color:var(--muted);font-size:10.5px">${na?'—':mlabel(r[0])}</td>`+
     `<td style="text-align:center;padding:4px;color:var(--muted);font-size:10.5px">${na?'—':mlabel(r[1])}</td>`+
     a.codes.map(code=>{if(na)return '<td class="cellcode" style="background:#f2f2f2;color:#bdbdbd;font-weight:400">—</td>';
       const v=RA[code],f=code==='I';return `<td class="cellcode" style="background:${f?'#eceff1':v.c};color:${f?'#90a4ae':'#fff'}">${code}</td>`;}).join('')+'</tr>';});
 document.getElementById('heatA').innerHTML=h+'</tbody></table>';
}
function sectorIndex(ck){
 const keys = ck==='__ALL__'? D.contracts.map(c=>c.key):[ck];
 const tot=SEC.map(()=>0);
 keys.forEach(k=>{const c=D.contracts.find(x=>x.key===k);
  actsOf(k).forEach((a,ai)=>{const r=rangeOf(k,ai,c&&c.adj); if(!r||!r[0])return;   // N/A no suma
   a.codes.forEach((code,j)=>{tot[j]+=(W[code]||0);});});});
 return tot;}
function renderB(ck){const tot=sectorIndex(ck),mx=Math.max(...tot,1);
 document.getElementById('barsB').innerHTML=SEC.map((s,j)=>`<div class="bar"><span class="nm" style="color:${SC[j]}">${s}</span><span class="bw"><span class="bf" style="width:${100*tot[j]/mx}%;background:${SC[j]}"></span></span><span class="vv">${tot[j]}</span></div>`).join('');
}
function tempLoad(ck){const keys= ck==='__ALL__'? D.contracts.map(c=>c.key):[ck];const load=SEC.map(()=>MON.map(()=>0));
 keys.forEach(k=>{const c=D.contracts.find(x=>x.key===k);const acts=actsOf(k);acts.forEach((a,ai)=>{const months=mrange(rangeOf(k,ai,c&&c.adj));a.codes.forEach((code,j)=>{const w=W[code]||0;if(w)months.forEach(mi=>load[j][mi]+=w);});});});return load;}
function heat(t){const st=[[255,255,204],[254,217,118],[254,178,76],[253,141,60],[240,59,32],[189,0,38]];const x=Math.max(0,Math.min(1,t))*(st.length-1);const i=Math.floor(x),f=x-i;const a=st[i],b=st[Math.min(i+1,st.length-1)];return `rgb(${a.map((v,k)=>Math.round(v+(b[k]-v)*f)).join(',')})`;}
function renderC(ck){const load=tempLoad(ck),mx=Math.max(1,...load.flat());
 let h='<table class="hz"><thead><tr><th class="rowlab">Sector \\ Mes</th>'+MON.map(m=>`<th>${mlabel(m)}</th>`).join('')+'</tr></thead><tbody>';
 SEC.forEach((s,j)=>{h+=`<tr><td class="rowlab" style="color:${SC[j]}">${s}</td>`+load[j].map(v=>{const t=v/mx;return `<td class="tcell" title="${s}·carga ${v}" style="background:${v?heat(t):'transparent'};color:${t>0.55?'#fff':'#7a2e0e'}">${v||''}</td>`;}).join('')+'</tr>';});
 document.getElementById('heatC').innerHTML=h+'</tbody></table>';
}
const IMPC={"ALTA":"#C62828","MEDIA":"#EF6C00","BAJA":"#2E7D32","\u2014":"#90A4AE"};
document.getElementById('cambios').innerHTML='<table class="hz" style="width:100%"><thead><tr><th style="width:70px">Impacto</th><th style="width:100px">Tipo</th><th class="rowlab" style="max-width:none">Detalle</th></tr></thead><tbody>'+
 D.cambios.map(c=>`<tr><td class="cellcode" style="background:${IMPC[c.imp]||'#90A4AE'};width:70px">${c.imp}</td><td class="rowlab" style="max-width:none"><b>${c.tipo}</b></td><td class="rowlab" style="max-width:none;font-weight:400">${c.det}</td></tr>`).join('')+'</tbody></table>';
const totR=D.check.reduce((s,c)=>s+c.ver,0), totF=D.check.reduce((s,c)=>s+c.fal,0);
document.getElementById('check').innerHTML=
 `<div style="font-weight:800;color:${totF?'#C62828':'#2E7D32'};margin-bottom:8px">${totR} reglas de precedencia verificadas sobre ${D.check.length} provisiones · ${totF} incoherencias</div>`+
 '<table class="hz" style="width:100%"><thead><tr><th class="rowlab" style="max-width:none">Provisión</th><th>Reglas verificadas</th><th>Incoherencias</th><th>Actividades N/A</th></tr></thead><tbody>'+
 D.check.map(c=>`<tr><td class="rowlab" style="max-width:none">${c.nm}</td><td style="text-align:center;padding:5px">${c.ver}</td><td class="cellcode" style="background:${c.fal?'#C62828':'#2E7D32'}">${c.fal}</td><td style="text-align:center;padding:5px;color:var(--muted)">${c.na}</td></tr>`).join('')+'</tbody></table>';
document.getElementById('notas').innerHTML='<ul style="font-size:12px;line-height:1.7;padding-left:18px;margin:0">'+D.notas.map(n=>`<li>${n}</li>`).join('')+'</ul>';
function render(){const ck=csel.value;renderA(ck==='__ALL__'?D.contracts[0].key:ck);renderB(ck);renderC(ck);}
csel.onchange=render;render();
</script></body></html>"""

p1=build_xlsx(); p2=build_pdf(); p3=build_html(); print("OK\n ",p1,"\n ",p2,"\n ",p3)
