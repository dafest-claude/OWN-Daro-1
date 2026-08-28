#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
MATRIZ DE INTERFASES — Provisiones críticas CPF-2 (Rev3, 2026-08-21).
Cambios respecto de Rev2:
  - Activación de suministros pasa a Informado (I) en la gestión de OC/contrato.
  - Se agrega un tablero HTML de MAPA DE CALOR:
      A) Involucramiento por tarea (tarea × sector) por contrato — solo E/A/R/P (I = sin peso).
      B) Índice de involucramiento por sector (carga relativa) por contrato.
      C) Carga temporal (sector × mes) — cuándo se afecta cada recurso, por contrato y agregado.
Genera Excel (matriz), PDF (matriz A3) y HTML (mapa de calor + carga temporal).
Nota: las fechas provienen del cronograma integrado Rev2; se recalcularán con los cronogramas
por contrato cuando estén disponibles.
"""
import os, json, textwrap
GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
REV="Rev3"; FECHA="2026-08-21"

ACT_SHORT=["Proveedor","Activación suministros","AESA Proy/Contrato","AESA Ingeniería",
           "AESA Construcciones","AESA Precom/Com","AESA Calidad (QA/QC)","Cliente PPSA"]
RACI={'E':('1F3864','Ejecuta / responsable primario'),'A':('C62828','Aprueba / libera'),
      'R':('EF6C00','Revisa / comenta'),'P':('2E7D32','Participa / Asiste'),'I':('90A4AE','Informado')}
WEIGHT={'E':4,'A':2,'R':2,'P':3,'I':0}

# codes = [Proveedor, Activación, Proyecto, Ingeniería, Construcciones, Precom&Com, Calidad, Cliente]
ACTIVITIES=[
 ("Contrato","Gestión de la OC / contrato y coordinación",           ['P','I','E','I','I','I','I','I']),  # Activación -> I
 ("Ingeniería","Emisión de documentación (IB / ID / IC)",            ['E','I','I','R','I','I','I','I']),
 ("Ingeniería","Revisión y aprobación de documentación",             ['P','I','I','A','R','I','R','I']),
 ("Ingeniería","Freezing / hold points · cierre de ingeniería",      ['E','I','P','A','R','I','R','I']),
 ("Procura / Fabricación","Procura y fabricación (seguimiento)",     ['E','E','R','I','I','I','R','I']),
 ("Procura / Fabricación","Certificación de hitos de pago / avance", ['P','P','E','I','I','I','I','I']),
 ("Pruebas de fábrica (FAT)","Procedimientos de FAT/pruebas (emisión y aprobación)",['E','I','I','R','I','P','A','I']),
 ("Pruebas de fábrica (FAT)","Ejecución de FAT / pruebas en fábrica (asistencia)",  ['E','P','P','P','I','P','A','P']),
 ("Pruebas de fábrica (FAT)","iFAT / pruebas de integración (donde aplica)",        ['E','I','P','P','P','P','R','P']),
 ("Despacho / Sitio","Despacho y entrega en sitio",                  ['E','E','R','I','P','I','A','I']),
 ("Despacho / Sitio","Montaje / integración en sitio",               ['P','I','R','R','E','I','R','I']),
 ("Precom. / Puesta en marcha","SAT / pruebas en sitio",             ['P','I','P','P','P','E','R','P']),
 ("Precom. / Puesta en marcha","Precomisionado",                     ['P','I','P','P','P','E','R','P']),
 ("Precom. / Puesta en marcha","Comisionado / Puesta en marcha (PEM)",['P','I','P','P','I','E','R','P']),
 ("Precom. / Puesta en marcha","Capacitación (operación / mantenimiento)",['E','I','P','I','I','P','I','P']),
 ("Cierre","Documentación final / DataBook / As-built",              ['E','I','I','A','P','P','A','I']),
]
ADJ_ROW=("Adjudicación","Proceso de adjudicación / cierre de contrato",['P','I','E','R','I','I','I','I'])

PROV=[
 ("SALAS","ABB — Salas Eléctricas SE#3 y SE#4","ABB","Adjudicado — OC vigente + planificación","C62828",False),
 ("PMS","ABB — PMS (Power Management System)","ABB","Adjudicado — OC vigente + planificación","EF6C00",False),
 ("INAUCO","Inauco — PCS / SCADA / Comunicaciones","Inauco","Adjudicado — Obra 129-008","1565C0",False),
 ("HIMA_MSH","HIMA — Tableros de marshalling (ESD / F&G / PSS)","HIMA","Adjudicado","6A1B9A",False),
 ("HIMA_CFG","HIMA — Adecuación y configuración SIS","HIMA","EN ADJUDICACIÓN","00838F",True),
]
DATES={
 "ABB — Salas Eléctricas SE#3 y SE#4":{"Freezing / hold points · cierre de ingeniería":"FP Constr. S#3: 31-AGO-2026","Procura y fabricación (seguimiento)":"Fabricación: 01-SEP → 02-NOV-2026","Ejecución de FAT / pruebas en fábrica (asistencia)":"FAT Brasil 03-23-NOV-26 · FAT Salas 08-MAR→19-ABR-27","iFAT / pruebas de integración (donde aplica)":"iFAT: 25-JUN → 14-JUL-2027","Despacho y entrega en sitio":"Despacho 02-ABR→14-MAY · Entrega S#3: 14-MAY-2027","Montaje / integración en sitio":"Montaje shelters (Mendoza): 18-ENE→26-MAR-2027","Comisionado / Puesta en marcha (PEM)":"Comis. 16-JUL-27 → RFSU 31-ENE-2028"},
 "ABB — PMS (Power Management System)":{"Freezing / hold points · cierre de ingeniería":"FP Ing. Básica: 21-AGO-2026","Procura y fabricación (seguimiento)":"Procura → 05-NOV · Armado+SW → 17-DIC-2026","Ejecución de FAT / pruebas en fábrica (asistencia)":"FAT Maqueta 22-SEP-26 · FAT PMS 25-29-ENE-27","iFAT / pruebas de integración (donde aplica)":"FAT Integral Shelter 31-MAR · iFAT 25-JUN→14-JUL-27","Despacho y entrega en sitio":"Despacho + Montaje shelter: 04-FEB→17-MAR-2027","Documentación final / DataBook / As-built":"DataBook PMS: 31-MAR-2027","Comisionado / Puesta en marcha (PEM)":"SAT PMS sitio OCT→DIC-27 · RFSU 31-ENE-2028","Capacitación (operación / mantenimiento)":"Capacitación PMS: FEB-MAR-2028"},
 "Inauco — PCS / SCADA / Comunicaciones":{"Gestión de la OC / contrato y coordinación":"Recepción OC: 28-JUL-2026","Procura y fabricación (seguimiento)":"Construcción tableros: 26-AGO-26 → 16-FEB-27","Ejecución de FAT / pruebas en fábrica (asistencia)":"FAT PCS 29-MAR→22-ABR · FAT SIS 29-MAR→26-ABR-27","iFAT / pruebas de integración (donde aplica)":"Pruebas PCS↔PMS 28-ABR→04-MAY · iFAT JUL-27","SAT / pruebas en sitio":"SAT/Comis campo: 18-MAY → 05-OCT-2027","Comisionado / Puesta en marcha (PEM)":"Comisionado → 31-ENE-2028","Documentación final / DataBook / As-built":"Prueba MCE + CAO: → 07-DIC-2027"},
 "HIMA — Tableros de marshalling (ESD / F&G / PSS)":{"Procura y fabricación (seguimiento)":"Fabricación tablero: 03-JUL → 25-DIC-2026","Despacho y entrega en sitio":"Recepción en Inauco: 25-DIC-2026","Montaje / integración en sitio":"Recableado interno (45 d): 25-DIC → 08-FEB-2027","Ejecución de FAT / pruebas en fábrica (asistencia)":"FAT SIS (conjunto): 29-MAR → 26-ABR-2027","Comisionado / Puesta en marcha (PEM)":"Soporte campo → 05-OCT-2027"},
 "HIMA — Adecuación y configuración SIS":{"Proceso de adjudicación / cierre de contrato":"A definir (en adjudicación)","Ejecución de FAT / pruebas en fábrica (asistencia)":"FAT SIS (conjunto): 29-MAR → 26-ABR-2027","Comisionado / Puesta en marcha (PEM)":"Asistencia PEM: 2027 (a confirmar)"},
}
# rangos por fase (mes) por contrato -> para la carga temporal
PHASE_RANGE={
 "SALAS":{"Contrato":["2026-05","2027-05"],"Ingeniería":["2026-05","2026-08"],"Procura / Fabricación":["2026-09","2026-11"],"Pruebas de fábrica (FAT)":["2026-11","2027-04"],"Despacho / Sitio":["2027-01","2027-05"],"Precom. / Puesta en marcha":["2027-05","2028-01"],"Cierre":["2027-05","2027-07"]},
 "PMS":{"Contrato":["2026-05","2027-04"],"Ingeniería":["2026-06","2026-10"],"Procura / Fabricación":["2026-05","2026-12"],"Pruebas de fábrica (FAT)":["2026-09","2027-07"],"Despacho / Sitio":["2027-02","2027-03"],"Precom. / Puesta en marcha":["2027-05","2028-02"],"Cierre":["2027-03","2027-04"]},
 "INAUCO":{"Contrato":["2026-07","2027-10"],"Ingeniería":["2026-08","2026-11"],"Procura / Fabricación":["2026-08","2027-03"],"Pruebas de fábrica (FAT)":["2027-03","2027-05"],"Despacho / Sitio":["2027-04","2027-05"],"Precom. / Puesta en marcha":["2027-05","2027-10"],"Cierre":["2027-10","2027-12"]},
 "HIMA_MSH":{"Contrato":["2026-07","2027-05"],"Ingeniería":["2026-07","2026-09"],"Procura / Fabricación":["2026-07","2026-12"],"Despacho / Sitio":["2026-12","2027-02"],"Pruebas de fábrica (FAT)":["2027-03","2027-04"],"Precom. / Puesta en marcha":["2027-05","2027-10"],"Cierre":["2027-05","2027-06"]},
 "HIMA_CFG":{"Adjudicación":["2026-09","2026-12"],"Contrato":["2027-01","2027-05"],"Ingeniería":["2027-01","2027-03"],"Procura / Fabricación":["2027-01","2027-03"],"Pruebas de fábrica (FAT)":["2027-03","2027-04"],"Despacho / Sitio":["2027-03","2027-04"],"Precom. / Puesta en marcha":["2027-05","2027-10"],"Cierre":["2027-10","2027-11"]},
}
NOTAS=[
 "Alcance: desde la adjudicación/OC en adelante. HIMA — Adecuación y configuración SIS está EN ADJUDICACIÓN; el resto adjudicado.",
 "Activación de suministros: gestión de la OC/provisión (expediting) — E en procura/fabricación y despacho; Informado en la gestión del contrato.",
 "AESA — Precom&Comisionado y AESA — Calidad (QA/QC) pertenecen al mismo Depto. de Calidad Integral, separados por función.",
 "AESA — Calidad (QA/QC): revisa/aprueba/libera según estadios de fabricación y despacho; aprueba/revisa la documentación de ensayos (protocolos, DataBook).",
 "AESA — Precom&Comisionado ejecuta precom/PEM y asiste a FAT/pruebas.",
 "Hitos/fechas: cronograma integrado Rev2. Se recalcularán con los cronogramas por contrato cuando estén disponibles.",
]
def dfor(prov,act): return DATES.get(prov,{}).get(act,"")

# ---- EXCEL + PDF: reutiliza el enfoque de Rev2 (con la corrección de Activación) ----
def build_xlsx():
    import openpyxl
    from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
    thin=Side(style="thin",color="D9D9D9"); B=Border(thin,thin,thin,thin); NAVY="1F3864"; HEAD="37474F"
    def cell(ws,r,c,v,fill=None,color="000000",bold=False,size=9,center=False,wrap=False):
        cc=ws.cell(r,c,v); cc.font=Font(bold=bold,color=color,size=size)
        if fill: cc.fill=PatternFill("solid",fgColor=fill)
        cc.alignment=Alignment(horizontal="center" if center else "left",vertical="center",wrap_text=wrap); cc.border=B; return cc
    wb=openpyxl.Workbook(); ws=wb.active; ws.title="Matriz de Interfases"; NC=11
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=NC); cell(ws,1,1,f"MATRIZ DE INTERFASES / RESPONSABILIDADES — PROVISIONES CRÍTICAS CPF-2 · {REV} · {FECHA}",NAVY,"FFFFFF",True,12,True); ws.row_dimensions[1].height=22
    r=4; cell(ws,r,1,"Leyenda:",None,"000000",True,9)
    for i,(k,(col,desc)) in enumerate(RACI.items()): cell(ws,r,2+i,f"{k} = {desc}",col,"FFFFFF" if k!='I' else "000000",True,8,True,True)
    r=6
    def header(r):
        cell(ws,r,1,"Fase",HEAD,"FFFFFF",True,8.5,True); cell(ws,r,2,"Actividad / interfase",HEAD,"FFFFFF",True,8.5,True); cell(ws,r,3,"Hito / Fecha (cron. Rev2)",HEAD,"FFFFFF",True,8,True,True)
        for j,a in enumerate(ACT_SHORT): cell(ws,r,4+j,a,HEAD,"FFFFFF",True,7.6,True,True)
        ws.row_dimensions[r].height=30
    for (key,name,ven,estado,col,adj) in PROV:
        ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=NC); cell(ws,r,1,f"▶  {name}   ·   {ven}   ·   {estado}",col,"FFFFFF",True,10); ws.row_dimensions[r].height=18; r+=1
        header(r); r+=1
        for fase,act,codes in (([ADJ_ROW]+ACTIVITIES) if adj else ACTIVITIES):
            cell(ws,r,1,fase,None,"333333",False,8); cell(ws,r,2,act,None,"000000",False,8.5,False,True); cell(ws,r,3,dfor(name,act),None,"1F3864",False,8,False,True)
            for j,code in enumerate(codes):
                fc,_=RACI[code]; cell(ws,r,4+j,code,fc,"FFFFFF" if code!='I' else "000000",True,9,True)
            r+=1
        r+=1
    for c,w in zip("ABCDEFGHIJK",[18,42,30,11,12,12,12,12,12,12,11]): ws.column_dimensions[c].width=w
    ws.freeze_panes="A7"
    ns=wb.create_sheet("Notas");
    for i,n in enumerate(NOTAS): cell(ns,2+i,1,f"• {n}",None,"000000",False,9,False,True); ns.row_dimensions[2+i].height=28
    ns.column_dimensions['A'].width=150
    path=os.path.join(GEN,f"Matriz de Interfases - Provisiones criticas - {REV} - {FECHA}.xlsx"); wb.save(path); return path

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
    aw=[19*mm,60*mm,54*mm]+[ (x1-x0-133*mm)/8 ]*8
    FSHORT={"Procura / Fabricación":"Procura/Fabric.","Pruebas de fábrica (FAT)":"Pruebas FAT","Precom. / Puesta en marcha":"Precom./PEM"}
    ASHORT={"Emisión de documentación (IB / ID / IC)":"Emisión de doc. (IB/ID/IC)","Revisión y aprobación de documentación":"Revisión y aprobación de doc.","Freezing / hold points · cierre de ingeniería":"Freezing/hold points · cierre ing.","Certificación de hitos de pago / avance":"Certif. hitos de pago / avance","Procedimientos de FAT/pruebas (emisión y aprobación)":"Procedimientos FAT (emisión/aprob.)","Ejecución de FAT / pruebas en fábrica (asistencia)":"Ejecución FAT / pruebas (asistencia)","iFAT / pruebas de integración (donde aplica)":"iFAT / pruebas integración (si aplica)","Comisionado / Puesta en marcha (PEM)":"Comisionado / PEM","Capacitación (operación / mantenimiento)":"Capacitación (op./mant.)","Documentación final / DataBook / As-built":"Doc. final / DataBook / As-built","Gestión de la OC / contrato y coordinación":"Gestión de OC / contrato"}
    def mhead(y):
        cx=x0; c.setFillColor(colors.HexColor('#37474F')); c.rect(x0,y-9*mm,sum(aw),9*mm,fill=1,stroke=0); c.setFillColor(colors.white); c.setFont("Helvetica-Bold",6.8)
        for l,w in zip(["Fase","Actividad / interfase","Hito / Fecha (cronograma)"]+ACT_SHORT,aw):
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
        for ri,(fase,act,codes) in enumerate(([ADJ_ROW]+ACTIVITIES) if adj else ACTIVITIES):
            y,newp=ensure(y,6*mm)
            if newp: y=mhead(y)
            rowh=6.2*mm; cx=x0
            c.setFillColor(colors.HexColor('#F7F9FB') if ri%2 else colors.white); c.rect(x0,y-rowh,sum(aw),rowh,fill=1,stroke=0)
            c.setFillColor(colors.HexColor('#444444')); c.setFont("Helvetica",6.6); c.drawString(cx+1.2*mm,y-4*mm,FSHORT.get(fase,fase)[:22]); cx+=aw[0]
            c.setFillColor(colors.black); c.setFont("Helvetica",6.8); c.drawString(cx+1.2*mm,y-4*mm,ASHORT.get(act,act)[:44]); cx+=aw[1]
            c.setFillColor(colors.HexColor('#1F3864')); c.setFont("Helvetica",6.0)
            for j,ln in enumerate(textwrap.wrap(dfor(name,act),int(aw[2]/1.32))[:2]): c.drawString(cx+1.2*mm,y-3.2*mm-j*2.8*mm,ln)
            cx+=aw[2]
            for code in codes:
                fc,_=RACI[code]; c.setFillColor(colors.HexColor('#'+fc)); c.roundRect(cx+aw[3]/2-3.4*mm,y-rowh+1*mm,6.8*mm,rowh-2*mm,1,fill=1,stroke=0)
                c.setFillColor(colors.white if code!='I' else colors.black); c.setFont("Helvetica-Bold",7.2); c.drawCentredString(cx+aw[3]/2,y-4*mm,code); cx+=aw[3]
            y-=rowh
        y-=4*mm
    c.showPage(); c.save(); return path

# ---- HTML mapa de calor + carga temporal ----
def build_html():
    contracts=[{"key":k,"name":n,"ven":v,"estado":e,"color":"#"+c,"adj":adj} for (k,n,v,e,c,adj) in PROV]
    acts=[{"fase":f,"act":a,"codes":cd} for (f,a,cd) in ACTIVITIES]
    adjrow={"fase":ADJ_ROW[0],"act":ADJ_ROW[1],"codes":ADJ_ROW[2]}
    payload=dict(contracts=contracts,acts=acts,adj=adjrow,sectors=ACT_SHORT,
                 raci={k:{"c":"#"+v[0],"d":d} for k,(v,d) in [(k,(RACI[k],RACI[k][1])) for k in RACI]},
                 weight=WEIGHT,phase=PHASE_RANGE)
    # fix raci payload (color+desc)
    payload["raci"]={k:{"c":"#"+RACI[k][0],"d":RACI[k][1]} for k in RACI}
    html=HTML_TMPL.replace("__REV__",REV).replace("__FECHA__",FECHA).replace("__DATA__",json.dumps(payload,ensure_ascii=False))
    path=os.path.join(GEN,f"Mapa de calor Interfases - {REV} - {FECHA}.html"); open(path,"w",encoding="utf-8").write(html); return path

HTML_TMPL=r"""<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>Mapa de calor de interfases · CPF-2</title>
<style>
:root{--bg:#f4f6f8;--card:#fff;--ink:#1f2933;--muted:#607080;--line:#e2e8f0;--navy:#1F3864;--shadow:0 1px 3px rgba(0,0,0,.08)}
@media(prefers-color-scheme:dark){:root{--bg:#0f141a;--card:#1a222c;--ink:#e6edf3;--muted:#93a1b0;--line:#2a343f;--shadow:0 1px 3px rgba(0,0,0,.5)}}
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--ink);font-size:13.5px}
header{background:var(--navy);color:#fff;padding:15px 22px}h1{margin:0;font-size:19px}.sub{opacity:.85;font-size:12px;margin-top:3px}
.wrap{max-width:1500px;margin:0 auto;padding:16px 22px 60px}
.toolbar{display:flex;gap:12px;flex-wrap:wrap;align-items:center;margin:12px 0}
select{padding:7px 11px;border:1px solid var(--line);border-radius:8px;background:var(--card);color:var(--ink);font-size:13px;font-weight:600}
.card{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;box-shadow:var(--shadow);margin:14px 0;overflow-x:auto}
.card h2{margin:0 0 4px;font-size:15px}.card .d{color:var(--muted);font-size:12px;margin-bottom:12px}
table{border-collapse:collapse;font-size:11.5px}th,td{border:1px solid var(--line);padding:0}
.hz th{background:rgba(127,127,127,.06);color:var(--muted);font-size:10.5px;font-weight:700;padding:5px 6px;text-align:center;vertical-align:bottom}
.rowlab{text-align:left!important;padding:5px 8px!important;font-weight:600;white-space:normal;line-height:1.2;max-width:260px}
.cellcode{width:38px;height:30px;text-align:center;font-weight:800;color:#fff;font-size:11px}
.legend{display:flex;gap:14px;flex-wrap:wrap;margin:6px 0 2px;font-size:12px}.lg{display:flex;align-items:center;gap:6px}.sw{width:14px;height:14px;border-radius:3px;display:inline-block}
.tcell{width:34px;height:26px;text-align:center;font-size:9.5px;color:#0b1f3a}
.bar{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:12.5px}.bar .nm{width:180px;flex-shrink:0}.bar .bw{flex:1;background:var(--line);border-radius:5px;height:16px;overflow:hidden;max-width:520px}.bar .bf{display:block;height:100%;border-radius:5px;background:var(--navy);min-width:2px}.bar .vv{width:44px;text-align:right;font-weight:700}
.note{font-size:11px;color:var(--muted);margin-top:10px}
.scalebar{display:flex;align-items:center;gap:6px;font-size:11px;color:var(--muted)}
.grad{width:160px;height:12px;border-radius:3px;background:linear-gradient(90deg,#f7fbff,#c6dbef,#6baed6,#2171b5,#08306b)}
</style></head><body>
<header><h1>Mapa de calor de interfases — Involucramiento y carga temporal</h1>
<div class="sub">CPF-2 La Calera II · Provisiones críticas · __REV__ · __FECHA__ · (fechas del cronograma integrado Rev2; se actualizarán con los cronogramas por contrato)</div></header>
<div class="wrap">
 <div class="toolbar"><b>Contrato:</b><select id="csel"></select>
   <span class="legend" id="raci-legend"></span></div>

 <div class="card"><h2>A · Involucramiento por tarea</h2>
   <div class="d">Cómo participa cada sector en cada actividad del contrato. Solo E/A/R/P representan involucramiento; I (informado) se muestra tenue.</div>
   <div id="heatA"></div></div>

 <div class="card"><h2>B · Índice de involucramiento por sector</h2>
   <div class="d">Carga relativa de cada sector en el contrato (peso: E=4 · P=3 · R=2 · A=2 · I=0).</div>
   <div id="barsB"></div></div>

 <div class="card"><h2>C · Carga temporal (sector × mes)</h2>
   <div class="d">Cuándo se afecta cada sector a lo largo del tiempo (intensidad = involucramiento activo ese mes). <span class="scalebar">menos <span class="grad"></span> más</span></div>
   <div id="heatC"></div></div>
</div>
<script>
const D=__DATA__;
const RA=D.raci, W=D.weight, SEC=D.sectors;
const csel=document.getElementById('csel');
csel.innerHTML=D.contracts.map(c=>`<option value="${c.key}">${c.name} — ${c.estado}</option>`).join('')+'<option value="__ALL__">TODOS los contratos (agregado)</option>';
document.getElementById('raci-legend').innerHTML=Object.entries(RA).map(([k,v])=>`<span class="lg"><span class="sw" style="background:${v.c}"></span>${k} · ${v.d}</span>`).join('');
// months 2026-05 .. 2028-02
const MON=[];{let y=2026,m=5;for(let i=0;i<22;i++){MON.push(y+'-'+String(m).padStart(2,'0'));m++;if(m>12){m=1;y++;}}}
const MLAB=['Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic'];
const mlabel=iso=>{const[y,m]=iso.split('-');return MLAB[+m-1]+"'"+y.slice(2);};
function actsOf(ck){const c=D.contracts.find(x=>x.key===ck);const base=(c&&c.adj)?[D.adj].concat(D.acts):D.acts;return base;}
function mrange(ck,fase){const p=D.phase[ck]||{};const r=p[fase];if(!r)return[];const s=MON.indexOf(r[0]),e=MON.indexOf(r[1]);if(s<0||e<0)return[];const o=[];for(let i=s;i<=e;i++)o.push(i);return o;}

function renderA(ck){
 const acts=actsOf(ck);
 let h='<table class="hz"><thead><tr><th class="rowlab">Actividad</th>'+SEC.map(s=>`<th>${s.replace(/ /g,'<br>')}</th>`).join('')+'</tr></thead><tbody>';
 acts.forEach(a=>{
   h+=`<tr><td class="rowlab" title="${a.fase}">${a.act}</td>`+a.codes.map(code=>{
     const v=RA[code];const faint=code==='I';
     return `<td class="cellcode" style="background:${faint?'#eceff1':v.c};color:${faint?'#90a4ae':'#fff'}">${code}</td>`;}).join('')+'</tr>';
 });
 h+='</tbody></table>';document.getElementById('heatA').innerHTML=h;
}
function sectorIndex(ck){
 const acts=(ck==='__ALL__')? D.contracts.flatMap(c=>actsOf(c.key)) : actsOf(ck);
 const tot=SEC.map((_,j)=>acts.reduce((s,a)=>s+(W[a.codes[j]]||0),0));
 return tot;
}
function renderB(ck){
 const tot=sectorIndex(ck);const mx=Math.max(...tot,1);
 document.getElementById('barsB').innerHTML=SEC.map((s,j)=>`<div class="bar"><span class="nm">${s}</span><span class="bw"><span class="bf" style="width:${100*tot[j]/mx}%"></span></span><span class="vv">${tot[j]}</span></div>`).join('');
}
function tempLoad(ck){
 // load[sector][month]
 const keys= ck==='__ALL__'? D.contracts.map(c=>c.key):[ck];
 const load=SEC.map(()=>MON.map(()=>0));
 keys.forEach(k=>{actsOf(k).forEach(a=>{const months=mrange(k,a.fase);a.codes.forEach((code,j)=>{const w=W[code]||0;if(w)months.forEach(mi=>load[j][mi]+=w);});});});
 return load;
}
function col(t){ // 0..1 -> blue scale
 const stops=[[247,251,255],[198,219,239],[107,174,214],[33,113,181],[8,48,107]];
 const x=Math.max(0,Math.min(1,t))*(stops.length-1);const i=Math.floor(x),f=x-i;const a=stops[i],b=stops[Math.min(i+1,stops.length-1)];
 return `rgb(${a.map((v,k)=>Math.round(v+(b[k]-v)*f)).join(',')})`;
}
function renderC(ck){
 const load=tempLoad(ck);const mx=Math.max(1,...load.flat());
 let h='<table class="hz"><thead><tr><th class="rowlab">Sector \\ Mes</th>'+MON.map(m=>`<th>${mlabel(m)}</th>`).join('')+'</tr></thead><tbody>';
 SEC.forEach((s,j)=>{h+=`<tr><td class="rowlab">${s}</td>`+load[j].map(v=>{const t=v/mx;return `<td class="tcell" title="${s} · carga ${v}" style="background:${v?col(t):'transparent'};color:${t>0.5?'#fff':'#0b1f3a'}">${v||''}</td>`;}).join('')+'</tr>';});
 h+='</tbody></table>';document.getElementById('heatC').innerHTML=h;
}
function render(){const ck=csel.value;renderA(ck==='__ALL__'?D.contracts[0].key:ck);renderB(ck);renderC(ck);
 if(ck==='__ALL__'){document.getElementById('heatA').insertAdjacentHTML('afterbegin','<div class="note">La vista A muestra el 1er contrato como referencia (el involucramiento por tarea es común); B y C están agregados.</div>');}}
csel.onchange=render;render();
</script></body></html>"""

p1=build_xlsx(); p2=build_pdf(); p3=build_html()
print("OK\n ",p1,"\n ",p2,"\n ",p3)
