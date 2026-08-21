#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tie-in CONSOLIDADO (Electricidad + Instrumentación + IT/OT) — Rev1 (2026-08-21).
Novedades:
  - Listado eléctrico ACAL-00102-LG-E-0004 Rev.1: se REDEFINIÓ la condición de paro
    (PARO / SIN PARO / 'PARO o con TCT' / a definir).
  - Se incorporan los tie-in de Instrumentación (34) e IT/OT (12), con Categoría T1/T2/T3
    (T1 = cualquier momento; T2 = parada parcial; T3 = parada total) y Alcance (N/RE/RC/RL/D).
Genera: informe ejecutivo (PDF), dashboard (XLSX) y cuadro de conteo/seguimiento (HTML),
todos filtrables (especialidad, ¿paro?, nivel/tipo, alcance, área, vendor) e incluyen un
ítem de Control de cambios (Rev0 -> Rev1).
"""
import os, json
import xlrd
from collections import Counter, defaultdict

ENT="/home/user/OWN-Daro-1/Constructibilidad/Datos entrada"
GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
REV="Rev1"; FECHA="2026-08-21"
EL_SRC=os.path.join(ENT,"ACAL-00102-LG-E-0004-1.xls")
IN_SRC=os.path.join(ENT,"Listado de TIE-INs INS-IT_OT.xls")

# ---------------- extracción ----------------
def extract_el():
    sh=xlrd.open_workbook(EL_SRC).sheet_by_name("TIE IN electricos")
    # header row 5; columnas por nombre
    hdr=5; colmap={}
    for ci in range(sh.ncols):
        v=str(sh.cell_value(hdr,ci)).strip().upper()
        if v and v not in colmap: colmap[v]=ci
    idx={k:colmap.get(k) for k in ['ITEM','AREA','UBICACION','EQUIPO','TAREA','FUNCIÓN','ETAPA','CABLE','RESPONSABLE','OBSERVACIONES']}
    out=[]; carry={}
    for ri in range(hdr+1,sh.nrows):
        gc=lambda k: str(sh.cell_value(ri,idx[k])).replace('\n',' ').strip() if idx.get(k) is not None else ''
        rec={k:gc(k) for k in idx}
        if not rec['ITEM']: continue
        for k in ['AREA','UBICACION','EQUIPO','ETAPA','RESPONSABLE','TAREA']:
            if rec[k]: carry[k]=rec[k]
            else: rec[k]=carry.get(k,'')
        out.append(rec)
    return out

def extract_ins(sheet):
    sh=xlrd.open_workbook(IN_SRC).sheet_by_name(sheet)
    cols={'Area':0,'Item':2,'NoTieIn':4,'Tag':8,'Ubic':12,'Sistema':15,'Alcance':18,'Categoria':20,'Desc':22,'Resp':31,'Vendor':34,'Obs':43}
    out=[]; carry={}
    for ri in range(6,sh.nrows):
        gc=lambda c: str(sh.cell_value(ri,c)).replace('\n',' ').strip()
        rec={k:gc(c) for k,c in cols.items()}
        if not rec['Item']: continue
        for k in ['Area','Ubic','Sistema','Alcance','Categoria','Resp','Vendor']:
            if rec[k]: carry[k]=rec[k]
            else: rec[k]=carry.get(k,'')
        out.append(rec)
    return out

# ---------------- clasificación unificada ----------------
def paro_el(e):
    e=e.strip().upper()
    if e=='PARO': return 'Paro total','PARO'
    if 'TCT' in e: return 'Con/sin paro (TCT)','PARO o con TCT'
    if e=='SIN PARO': return 'Sin paro','SIN PARO'
    return 'A definir', (e or '-')
def paro_cat(c):
    c=c.strip().upper()
    if c=='T3': return 'Paro total','T3 · parada total'
    if c=='T2': return 'Paro parcial','T2 · parada parcial'
    if c=='T1': return 'Sin paro','T1 · cualquier momento'
    return 'A definir', (c or '-')
def nivel_el(func,tarea,equipo):
    fu=func.lower(); ta=tarea.lower(); eq=equipo.lower()
    if 'puesta a tierra' in fu or 'pat' in ta or 'malla de pat' in eq: return 'PAT'
    if 'potencia mt' in fu: return 'Potencia MT'
    if fu.startswith('potencia bt') or fu=='potencia bt' or 'iluminación' in fu or 'reposición de reserva' in fu or 'resistencia calefactora' in fu: return 'Potencia BT'
    return 'Control/Señales EL'
def nivel_in(sistema):
    s=sistema.upper()
    if any(k in s for k in ['SSS','ESD','PSS','USS','FGS','F&G']): return 'Instrum. Seguridad (SIS)'
    return 'Instrum. Control (PCS)'
ALC={'N':'Instalación nueva','D':'Desmantelamiento','RC':'Recalibración','RE':'Reemplazo','RL':'Relocalización'}
def norm_area(a):
    a=a.replace('.0','').strip(); return a or '—'

rows=[]
# EL
for r in extract_el():
    p,pd=paro_el(r['ETAPA'])
    rows.append(dict(esp='Electricidad', id=f"EL-{r['ITEM'].replace('.0','')}",
        area=norm_area(r['AREA']), ubic=r['UBICACION'], sistema=r['EQUIPO'],
        desc=(r['TAREA'] or r['FUNCIÓN']), func=r['FUNCIÓN'], cable=r['CABLE'],
        nivel=nivel_el(r['FUNCIÓN'],r['TAREA'],r['EQUIPO']), alcance='—',
        paro=p, parodet=pd, vendor='—', obs=r['OBSERVACIONES']))
# IN
for r in extract_ins("Tie-In Instrumentos"):
    p,pd=paro_cat(r['Categoria'])
    rows.append(dict(esp='Instrumentación', id=r['NoTieIn'] or f"IN-{r['Item'].replace('.0','')}",
        area=norm_area(r['Area']), ubic=r['Ubic'], sistema=r['Sistema'],
        desc=r['Desc'], func=r['Tag'], cable='',
        nivel=nivel_in(r['Sistema']), alcance=ALC.get(r['Alcance'].upper(),r['Alcance'] or '—'),
        paro=p, parodet=pd, vendor=(r['Vendor'] if r['Vendor'] not in('','-') else '—'), obs=r['Obs']))
# IT-OT
for r in extract_ins("Tie-In IT-OT"):
    p,pd=paro_cat(r['Categoria'])
    rows.append(dict(esp='IT-OT', id=r['NoTieIn'] or f"OT-{r['Item'].replace('.0','')}",
        area=norm_area(r['Area']), ubic=r['Ubic'], sistema=r['Sistema'],
        desc=r['Desc'], func=r['Tag'], cable='',
        nivel='IT/OT Red/Comunic.', alcance=ALC.get(r['Alcance'].upper(),r['Alcance'] or '—'),
        paro=p, parodet=pd, vendor=(r['Vendor'] if r['Vendor'] not in('','-') else '—'), obs=r['Obs']))

TOTAL=len(rows)
ESP=['Electricidad','Instrumentación','IT-OT']
PARO=['Paro total','Paro parcial','Con/sin paro (TCT)','Sin paro','A definir']
NIVEL=['Potencia MT','Potencia BT','Control/Señales EL','PAT','Instrum. Control (PCS)','Instrum. Seguridad (SIS)','IT/OT Red/Comunic.']
cnt=lambda key: {k:sum(1 for r in rows if r[key]==k) for k in {r[key] for r in rows}}
c_esp={e:sum(1 for r in rows if r['esp']==e) for e in ESP}
c_paro={p:sum(1 for r in rows if r['paro']==p) for p in PARO}
c_niv={n:sum(1 for r in rows if r['nivel']==n) for n in NIVEL}
c_area=dict(sorted(Counter(r['area'] for r in rows).items()))
REQ_PARO=c_paro['Paro total']+c_paro['Paro parcial']
mx_ep={e:{p:sum(1 for r in rows if r['esp']==e and r['paro']==p) for p in PARO} for e in ESP}
pct=lambda n:f"{100*n/TOTAL:.0f}%"

CHANGELOG=[
 ("Rev0","2026-08-18","Versión inicial. Solo listado ELÉCTRICO (ACAL-00102-LG-E-0004 Rev.0): 71 tie-in. Condición de paro con etapas PARO / PREPARO-POSPARO / POSPARO (50 en PARO)."),
 ("Rev1","2026-08-21","(1) Se adopta el listado eléctrico Rev.1 con la condición de paro REDEFINIDA: PARO=24, SIN PARO=43, 'PARO o con TCT'=1, a definir=3. (2) Se INCORPORAN Instrumentación (34) e IT/OT (12) con Categoría T1/T2/T3 y Alcance. (3) Consolidado 117 tie-in, filtrable por especialidad, ¿paro?, nivel/tipo, alcance, área y vendor."),
]

# ---------------- PDF ----------------
def build_pdf():
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    path=os.path.join(GEN,f"Tie-in Consolidado - Informe Ejecutivo - {REV} - {FECHA}.pdf")
    W,H=A4; c=canvas.Canvas(path,pagesize=A4)
    NAVY=colors.HexColor('#1F3864'); RED=colors.HexColor('#C62828'); ORA=colors.HexColor('#EF6C00'); GRN=colors.HexColor('#2E7D32'); GREY=colors.HexColor('#555555')
    x0=15*mm; x1=W-15*mm
    c.setFillColor(NAVY); c.rect(0,H-16*mm,W,16*mm,fill=1,stroke=0)
    c.setFillColor(colors.white); c.setFont("Helvetica-Bold",13.5); c.drawString(x0,H-11*mm,"TIE-IN CONSOLIDADO (EL · IN · IT/OT) — INFORME EJECUTIVO")
    c.setFont("Helvetica",8.5); c.drawRightString(x1,H-11*mm,f"{REV} · {FECHA}")
    y=H-22*mm; c.setFillColor(GREY); c.setFont("Helvetica",8)
    c.drawString(x0,y,"CPF-2 La Calera II · Fuentes: ACAL-00102-LG-E-0004 Rev.1 (EL) + Listado TIE-INs INS-IT/OT · Alcance AESA")
    y-=8*mm
    kpis=[("Total tie-in",str(TOTAL),NAVY),("Requiere paro (total+parcial)",f"{REQ_PARO}  ({pct(REQ_PARO)})",RED),
          ("Sin paro",str(c_paro['Sin paro']),GRN),("Especialidades","EL · IN · IT/OT",colors.HexColor('#1565C0'))]
    bw=(x1-x0-3*3*mm)/4
    for i,(lab,val,col) in enumerate(kpis):
        bx=x0+i*(bw+3*mm); c.setFillColor(colors.HexColor('#F2F4F7')); c.roundRect(bx,y-15*mm,bw,15*mm,3,fill=1,stroke=0)
        c.setFillColor(col); c.rect(bx,y-15*mm,2*mm,15*mm,fill=1,stroke=0)
        c.setFillColor(GREY); c.setFont("Helvetica",6.8); c.drawString(bx+3.5*mm,y-5*mm,lab)
        c.setFillColor(col); c.setFont("Helvetica-Bold",13); c.drawString(bx+3.5*mm,y-12.5*mm,val)
    y-=21*mm
    def tabla(titulo,headers,data,widths,y,hlf=None):
        c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9.5); c.drawString(x0,y,titulo); y-=4.6*mm
        tw=sum(widths); c.setFillColor(colors.HexColor('#37474F')); c.rect(x0,y-4.6*mm,tw,4.6*mm,fill=1,stroke=0)
        cx=x0; c.setFillColor(colors.white); c.setFont("Helvetica-Bold",7.2)
        for h,w in zip(headers,widths): c.drawString(cx+1.3*mm,y-3.3*mm,h); cx+=w
        y-=4.6*mm
        for ri,row in enumerate(data):
            hl=hlf and hlf(row)
            c.setFillColor(colors.HexColor('#FBE9E7') if hl else (colors.HexColor('#F7F9FB') if ri%2 else colors.white))
            c.rect(x0,y-4.6*mm,tw,4.6*mm,fill=1,stroke=0); cx=x0; c.setFillColor(colors.black)
            for j,(v,w) in enumerate(zip(row,widths)):
                c.setFont("Helvetica-Bold" if (j==0 or hl) else "Helvetica",7.2)
                if j==0: c.drawString(cx+1.3*mm,y-3.3*mm,str(v))
                else: c.drawRightString(cx+w-1.3*mm,y-3.3*mm,str(v))
                cx+=w
            y-=4.6*mm
        c.setStrokeColor(colors.HexColor('#CCCCCC')); c.setLineWidth(0.4); c.rect(x0,y,tw,(len(data)+1)*4.6*mm,fill=0,stroke=1)
        return y-5*mm
    W3=[70*mm,26*mm,24*mm]
    y=tabla("1 · Por especialidad",["Especialidad","Cantidad","%"],[[e,c_esp[e],pct(c_esp[e])] for e in ESP],W3,y)
    y=tabla("2 · ¿Requiere paro? (condición redefinida / categoría T1-T3)",["Condición","Cantidad","%"],
            [[p,c_paro[p],pct(c_paro[p])] for p in PARO],W3,y,hlf=lambda r:r[0] in ('Paro total','Paro parcial'))
    W5=[46*mm,26*mm,26*mm,26*mm,22*mm]
    y=tabla("3 · ¿Requiere paro?  ·  cruce por especialidad",["Especialidad","Paro total","Paro parcial","Sin paro","Otros"],
            [[e,mx_ep[e]['Paro total'],mx_ep[e]['Paro parcial'],mx_ep[e]['Sin paro'],mx_ep[e]['Con/sin paro (TCT)']+mx_ep[e]['A definir']] for e in ESP],W5,y)
    W3b=[70*mm,26*mm,24*mm]
    y=tabla("4 · Por nivel / tipo de señal",["Nivel / tipo","Cantidad","%"],[[n,c_niv[n],pct(c_niv[n])] for n in NIVEL if c_niv[n]],W3b,y)
    # control de cambios
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9.5); c.drawString(x0,y,"5 · Control de cambios / Actualización"); y-=5*mm
    for rev,fch,txt in CHANGELOG:
        c.setFillColor(RED if rev==REV else GREY); c.setFont("Helvetica-Bold",7.6); c.drawString(x0,y,f"{rev} · {fch}")
        c.setFillColor(colors.black); c.setFont("Helvetica",7.2)
        import textwrap
        for i,ln in enumerate(textwrap.wrap(txt,118)):
            c.drawString(x0+26*mm,y,ln); y-=3.8*mm
        y-=1.5*mm
    # conclusiones
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9.5); c.drawString(x0,y,"Conclusiones ejecutivas"); y-=4.6*mm
    concl=[
     f"• {REQ_PARO} de {TOTAL} tie-in ({pct(REQ_PARO)}) requieren paro (total o parcial); {c_paro['Sin paro']} son SIN paro.",
     f"• La redefinición del listado eléctrico bajó los tie-in de paro de 50 a {mx_ep['Electricidad']['Paro total']} (el resto pasó a SIN PARO / TCT).",
     f"• IT/OT es el frente más condicionado por paro total ({mx_ep['IT-OT']['Paro total']} de {c_esp['IT-OT']}); Instrumentación combina T1/T2/T3.",
     f"• Paro total por especialidad: EL={mx_ep['Electricidad']['Paro total']} · IN={mx_ep['Instrumentación']['Paro total']} · IT/OT={mx_ep['IT-OT']['Paro total']}.",
    ]
    c.setFont("Helvetica",7.6); c.setFillColor(colors.black)
    for ln in concl: c.drawString(x0,y,ln[:150]); y-=4.2*mm
    c.setFont("Helvetica-Oblique",6.6); c.setFillColor(GREY)
    c.drawString(x0,12*mm,"Categorías IN/IT-OT: T1=cualquier momento (sin paro) · T2=parada parcial · T3=parada total. EL: PARO / SIN PARO / 'PARO o con TCT' (trabajo con tensión) / a definir.")
    c.drawString(x0,9*mm,"Valores de referencia, sujetos a la revisión vigente de cada listado.")
    c.showPage(); c.save(); return path

# ---------------- XLSX ----------------
def build_xlsx():
    import openpyxl
    from openpyxl.styles import Font,PatternFill,Alignment,Border,Side
    path=os.path.join(GEN,f"Dashboard Tie-in Consolidado - {REV} - {FECHA}.xlsx")
    wb=openpyxl.Workbook(); thin=Side(style="thin",color="D9D9D9"); B=Border(thin,thin,thin,thin)
    NAVY="1F3864"; HEAD="37474F"; RED="C62828"
    ws=wb.active; ws.title="Dashboard"
    ws["A1"]=f"DASHBOARD TIE-IN CONSOLIDADO (EL · IN · IT/OT) · {REV} · {FECHA}"; ws.merge_cells("A1:F1")
    ws["A1"].font=Font(bold=True,size=13,color="FFFFFF"); ws["A1"].fill=PatternFill("solid",fgColor=NAVY); ws.row_dimensions[1].height=22
    ws["A2"]="Fuentes: ACAL-00102-LG-E-0004 Rev.1 (EL) + Listado TIE-INs INS-IT/OT · CPF-2 La Calera II"; ws.merge_cells("A2:F2"); ws["A2"].font=Font(italic=True,size=9,color="555555")
    kp=[("Total",TOTAL,NAVY),("Requiere paro",REQ_PARO,RED),("Sin paro",c_paro['Sin paro'],"2E7D32"),("Especialidades",3,"1565C0")]
    for i,(lab,val,col) in enumerate(kp):
        L=chr(65+i); ws[f"{L}4"]=lab; ws[f"{L}4"].font=Font(size=9,color="FFFFFF",bold=True); ws[f"{L}4"].fill=PatternFill("solid",fgColor=col); ws[f"{L}4"].alignment=Alignment(horizontal="center")
        ws[f"{L}5"]=val; ws[f"{L}5"].font=Font(size=20,bold=True,color=col); ws[f"{L}5"].alignment=Alignment(horizontal="center")
    r=[7]
    def block(title,headers,data,hl=None):
        ws.cell(r[0],1,title).font=Font(bold=True,size=11,color=NAVY); r[0]+=1
        for j,h in enumerate(headers):
            cc=ws.cell(r[0],1+j,h); cc.font=Font(bold=True,color="FFFFFF",size=9); cc.fill=PatternFill("solid",fgColor=HEAD); cc.alignment=Alignment(horizontal="center" if j else "left"); cc.border=B
        r[0]+=1
        for row in data:
            for j,v in enumerate(row):
                cc=ws.cell(r[0],1+j,v); cc.border=B; cc.alignment=Alignment(horizontal="center" if j else "left")
                if j==0: cc.font=Font(bold=True,size=9)
                if hl and hl(row): cc.fill=PatternFill("solid",fgColor="FBE9E7"); cc.font=Font(bold=True,size=9,color=RED if j else "000000")
            r[0]+=1
        r[0]+=1
    block("Por especialidad",["Especialidad","Cantidad","%"],[[e,c_esp[e],pct(c_esp[e])] for e in ESP])
    block("¿Requiere paro?",["Condición","Cantidad","%"],[[p,c_paro[p],pct(c_paro[p])] for p in PARO],hl=lambda row:row[0] in('Paro total','Paro parcial'))
    block("¿Requiere paro? · por especialidad",["Especialidad","Paro total","Paro parcial","Sin paro","TCT/def."],
          [[e,mx_ep[e]['Paro total'],mx_ep[e]['Paro parcial'],mx_ep[e]['Sin paro'],mx_ep[e]['Con/sin paro (TCT)']+mx_ep[e]['A definir']] for e in ESP])
    block("Por nivel / tipo de señal",["Nivel / tipo","Cantidad"],[[n,c_niv[n]] for n in NIVEL if c_niv[n]])
    block("Por área",["Área","Cantidad"],[[a,n] for a,n in c_area.items()])
    block("Control de cambios",["Rev","Fecha","Detalle"],[[rev,fch,txt] for rev,fch,txt in CHANGELOG])
    for col,w in zip("ABCDEF",[40,14,16,12,12,10]): ws.column_dimensions[col].width=w
    # Listado
    ls=wb.create_sheet("Listado")
    HED=["Especialidad","ID / N° Tie-In","Área","Ubicación","Sistema/Equipo","Descripción / tarea","Nivel / tipo","Alcance","¿Requiere paro?","Detalle","Vendor","Cable","Estado","Fecha ejec.","Observaciones"]
    ls.append(HED)
    for j,_ in enumerate(HED,1):
        cc=ls.cell(1,j); cc.font=Font(bold=True,color="FFFFFF",size=9); cc.fill=PatternFill("solid",fgColor=HEAD); cc.alignment=Alignment(horizontal="center",wrap_text=True); cc.border=B
    ESPF={'Electricidad':'C62828','Instrumentación':'1565C0','IT-OT':'6A1B9A'}
    PARF={'Paro total':'C62828','Paro parcial':'EF6C00','Con/sin paro (TCT)':'8D6E63','Sin paro':'2E7D32','A definir':'90A4AE'}
    for d in rows:
        ls.append([d['esp'],d['id'],d['area'],d['ubic'],d['sistema'],d['desc'],d['nivel'],d['alcance'],
                   d['paro'],d['parodet'],d['vendor'],d['cable'],'','',d['obs']])
        rr=ls.max_row
        ls.cell(rr,1).fill=PatternFill("solid",fgColor=ESPF[d['esp']]); ls.cell(rr,1).font=Font(color="FFFFFF",bold=True,size=8)
        ls.cell(rr,9).fill=PatternFill("solid",fgColor=PARF[d['paro']]); ls.cell(rr,9).font=Font(color="FFFFFF",bold=True,size=8)
        for j in range(1,len(HED)+1): ls.cell(rr,j).border=B
    for col,w in zip("ABCDEFGHIJKLMNO",[14,15,7,16,18,30,20,15,17,16,10,15,10,11,24]): ls.column_dimensions[col].width=w
    ls.freeze_panes="A2"; ls.auto_filter.ref=f"A1:O{ls.max_row}"
    wb.save(path); return path

# ---------------- HTML ----------------
def build_html():
    path=os.path.join(GEN,f"Dashboard Tie-in Consolidado - {REV} - {FECHA}.html")
    data=[{k:d[k] for k in ['esp','id','area','ubic','sistema','desc','func','cable','nivel','alcance','paro','parodet','vendor','obs']} for d in rows]
    chlog=[{"rev":rev,"fch":fch,"txt":txt} for rev,fch,txt in CHANGELOG]
    html=(HTML_TMPL.replace("__REV__",REV).replace("__FECHA__",FECHA)
          .replace("__TOTAL__",str(TOTAL)).replace("__REQ__",str(REQ_PARO)).replace("__PCTREQ__",pct(REQ_PARO))
          .replace("__SINPARO__",str(c_paro['Sin paro']))
          .replace("__DATA__",json.dumps(data,ensure_ascii=False))
          .replace("__CHLOG__",json.dumps(chlog,ensure_ascii=False)))
    open(path,"w",encoding="utf-8").write(html); return path

HTML_TMPL=r"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Tie-in Consolidado · Cuadro de conteo</title>
<style>
:root{--bg:#f4f6f8;--card:#fff;--ink:#1f2933;--muted:#607080;--line:#e2e8f0;--red:#C62828;--ora:#EF6C00;--blue:#1565C0;--green:#2E7D32;--navy:#1F3864;--pur:#6A1B9A;--brown:#8D6E63;--shadow:0 1px 3px rgba(0,0,0,.08)}
@media(prefers-color-scheme:dark){:root{--bg:#0f141a;--card:#1a222c;--ink:#e6edf3;--muted:#93a1b0;--line:#2a343f;--shadow:0 1px 3px rgba(0,0,0,.5)}}
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--ink);font-size:14px}
header{background:var(--navy);color:#fff;padding:15px 22px}h1{margin:0;font-size:19px}.sub{opacity:.85;font-size:12px;margin-top:3px}
.wrap{max-width:1360px;margin:0 auto;padding:16px 22px 60px}
.kpis{display:flex;gap:12px;flex-wrap:wrap;margin:14px 0}.kpi{flex:1;min-width:150px;background:var(--card);border:1px solid var(--line);border-left:4px solid var(--navy);border-radius:10px;padding:12px 14px;box-shadow:var(--shadow)}
.kpi.red{border-left-color:var(--red)}.kpi.green{border-left-color:var(--green)}.kpi.blue{border-left-color:var(--blue)}
.kpi .l{font-size:11px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}.kpi .v{font-size:23px;font-weight:800;margin-top:2px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:12px;margin:6px 0}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:13px;box-shadow:var(--shadow)}.card h3{margin:0 0 9px;font-size:12.5px}
.brow{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:12px}.brow .nm{width:132px;flex-shrink:0}.brow .barwrap{flex:1;background:var(--line);border-radius:5px;height:15px;overflow:hidden}.brow .bar{display:block;height:100%;border-radius:5px;min-width:2px}.brow .n{width:40px;text-align:right;font-weight:700}
details{margin:12px 0;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:6px 14px;box-shadow:var(--shadow)}
summary{cursor:pointer;font-weight:700;font-size:13px;padding:6px 0}
.chg{font-size:12px;margin:6px 0;padding-left:10px;border-left:3px solid var(--line)}.chg b{color:var(--red)}
.filters{display:flex;gap:10px;flex-wrap:wrap;align-items:end;margin:14px 0 6px}
.fg{display:flex;flex-direction:column;gap:3px}.fg label{font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.03em}
select,input[type=text]{padding:6px 9px;border:1px solid var(--line);border-radius:7px;background:var(--card);color:var(--ink);font-size:12.5px;min-width:130px}
.count{font-size:12px;color:var(--muted);margin:6px 0}.btn{padding:6px 12px;border:1px solid var(--line);border-radius:7px;background:var(--card);color:var(--ink);cursor:pointer;font-size:12px;font-weight:600}
table{width:100%;border-collapse:collapse;background:var(--card);border-radius:10px;overflow:hidden;box-shadow:var(--shadow);font-size:11.5px}
th,td{padding:6px 8px;border-bottom:1px solid var(--line);text-align:left}th{background:rgba(127,127,127,.08);position:sticky;top:0;font-size:10px;text-transform:uppercase;color:var(--muted)}
.tag{color:#fff;border-radius:4px;padding:1px 6px;font-size:10px;font-weight:700;white-space:nowrap}
.pt{background:var(--red)}.pp{background:var(--ora)}.pn{background:var(--green)}.pd{background:#90A4AE}.px{background:var(--brown)}
</style></head><body>
<header><h1>Tie-in Consolidado (Electricidad · Instrumentación · IT/OT) — Cuadro de conteo y seguimiento</h1>
<div class="sub">CPF-2 La Calera II · Fuentes: ACAL-00102-LG-E-0004 Rev.1 + Listado TIE-INs INS-IT/OT · __REV__ · __FECHA__</div></header>
<div class="wrap">
 <div class="kpis">
  <div class="kpi"><div class="l">Total tie-in</div><div class="v">__TOTAL__</div></div>
  <div class="kpi red"><div class="l">Requiere paro (total+parcial)</div><div class="v">__REQ__ <span style="font-size:14px">(__PCTREQ__)</span></div></div>
  <div class="kpi green"><div class="l">Sin paro</div><div class="v">__SINPARO__</div></div>
  <div class="kpi blue"><div class="l">Especialidades</div><div class="v">EL·IN·IT/OT</div></div>
 </div>
 <div class="grid">
  <div class="card"><h3>Por especialidad</h3><div id="c-esp"></div></div>
  <div class="card"><h3>¿Requiere paro?</h3><div id="c-paro"></div></div>
  <div class="card"><h3>Por nivel / tipo de señal</h3><div id="c-niv"></div></div>
 </div>
 <details><summary>📝 Control de cambios / Actualización</summary><div id="chlog"></div></details>
 <div class="filters" id="filters"></div>
 <div class="count" id="count"></div>
 <div style="max-height:560px;overflow:auto;border-radius:10px">
 <table id="tbl"><thead><tr>
  <th>Esp.</th><th>ID / N° Tie-In</th><th>Área</th><th>Ubicación</th><th>Sistema/Equipo</th><th>Descripción / tarea</th><th>Nivel / tipo</th><th>Alcance</th><th>¿Paro?</th><th>Vendor</th>
 </tr></thead><tbody></tbody></table></div>
 <div style="font-size:11px;color:var(--muted);margin-top:10px">Condición de paro — EL: PARO / SIN PARO / 'PARO o con TCT' (trabajo con tensión) / a definir. IN·IT/OT: T1=cualquier momento (sin paro) · T2=parada parcial · T3=parada total. Valores de referencia, sujetos a la revisión vigente de cada listado.</div>
</div>
<script>
const DATA=__DATA__, CHLOG=__CHLOG__;
const ESPC={'Electricidad':'var(--red)','Instrumentación':'var(--blue)','IT-OT':'var(--pur)'};
const PAROC={'Paro total':'var(--red)','Paro parcial':'var(--ora)','Con/sin paro (TCT)':'var(--brown)','Sin paro':'var(--green)','A definir':'#90A4AE'};
const PAROCLS={'Paro total':'pt','Paro parcial':'pp','Con/sin paro (TCT)':'px','Sin paro':'pn','A definir':'pd'};
const NIVC={'Potencia MT':'var(--pur)','Potencia BT':'var(--red)','Control/Señales EL':'var(--ora)','PAT':'var(--brown)','Instrum. Control (PCS)':'var(--blue)','Instrum. Seguridad (SIS)':'#00838F','IT/OT Red/Comunic.':'var(--green)'};
function bars(el,pairs,colf){const mx=Math.max(...pairs.map(p=>p[1]),1);el.innerHTML=pairs.filter(p=>p[1]).map(([k,n])=>`<div class="brow"><span class="nm" title="${k}">${k}</span><span class="barwrap"><span class="bar" style="width:${100*n/mx}%;background:${colf(k)}"></span></span><span class="n">${n}</span></div>`).join('');}
const cnt=f=>{const m={};DATA.forEach(d=>m[f(d)]=(m[f(d)]||0)+1);return m;};
const esp=cnt(d=>d.esp),paro=cnt(d=>d.paro),niv=cnt(d=>d.nivel);
bars(document.getElementById('c-esp'),['Electricidad','Instrumentación','IT-OT'].map(k=>[k,esp[k]||0]),k=>ESPC[k]);
bars(document.getElementById('c-paro'),['Paro total','Paro parcial','Con/sin paro (TCT)','Sin paro','A definir'].map(k=>[k,paro[k]||0]),k=>PAROC[k]);
bars(document.getElementById('c-niv'),['Potencia MT','Potencia BT','Control/Señales EL','PAT','Instrum. Control (PCS)','Instrum. Seguridad (SIS)','IT/OT Red/Comunic.'].map(k=>[k,niv[k]||0]),k=>NIVC[k]);
document.getElementById('chlog').innerHTML=CHLOG.map(c=>`<div class="chg"><b>${c.rev} · ${c.fch}</b> — ${c.txt}</div>`).join('');
// filtros
const DIMS=[['esp','Especialidad'],['paro','¿Requiere paro?'],['nivel','Nivel / tipo'],['alcance','Alcance'],['area','Área'],['vendor','Vendor']];
const state={};
const uniq=k=>[...new Set(DATA.map(d=>d[k]))].sort();
const fdiv=document.getElementById('filters');
DIMS.forEach(([k,lab])=>{const g=document.createElement('div');g.className='fg';
 g.innerHTML=`<label>${lab}</label><select data-k="${k}"><option value="">Todos</option>${uniq(k).map(v=>`<option>${v}</option>`).join('')}</select>`;fdiv.appendChild(g);});
const sg=document.createElement('div');sg.className='fg';sg.innerHTML=`<label>Buscar</label><input type="text" id="q" placeholder="tag, sistema, descripción…">`;fdiv.appendChild(sg);
const bg=document.createElement('div');bg.className='fg';bg.innerHTML=`<label>&nbsp;</label><button class="btn" id="reset">Limpiar</button>`;fdiv.appendChild(bg);
fdiv.querySelectorAll('select').forEach(s=>s.onchange=()=>{state[s.dataset.k]=s.value;render();});
document.getElementById('q').addEventListener('input',e=>{state.q=e.target.value.toLowerCase().trim();render();});
document.getElementById('reset').onclick=()=>{fdiv.querySelectorAll('select').forEach(s=>s.value='');document.getElementById('q').value='';Object.keys(state).forEach(k=>delete state[k]);render();};
function render(){
 const tb=document.querySelector('#tbl tbody');
 const rows=DATA.filter(d=>{
   for(const [k] of DIMS){if(state[k]&&d[k]!==state[k])return false;}
   if(state.q){const s=(d.id+' '+d.sistema+' '+d.desc+' '+d.func+' '+d.ubic+' '+d.cable).toLowerCase();if(!s.includes(state.q))return false;}
   return true;});
 document.getElementById('count').textContent=`${rows.length} de ${DATA.length} tie-in`;
 tb.innerHTML=rows.map(d=>`<tr>
  <td><span class="tag" style="background:${ESPC[d.esp]}">${d.esp==='Instrumentación'?'IN':(d.esp==='IT-OT'?'IT/OT':'EL')}</span></td>
  <td><b>${d.id}</b></td><td>${d.area}</td><td>${d.ubic}</td><td>${d.sistema}</td><td>${d.desc}</td>
  <td>${d.nivel}</td><td>${d.alcance}</td>
  <td><span class="tag ${PAROCLS[d.paro]}" title="${d.parodet}">${d.paro}</span></td><td>${d.vendor}</td></tr>`).join('');
}
render();
</script></body></html>"""

p1=build_pdf(); p2=build_xlsx(); p3=build_html()
print("TOTAL",TOTAL,"| por esp",c_esp,"| paro",c_paro,"| req_paro",REQ_PARO)
print("OK\n ",p1,"\n ",p2,"\n ",p3)
