#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Tie-in CONSOLIDADO (Electricidad + Instrumentación + IT/OT) — Rev2 (2026-09-15).

Fuentes NUEVAS (revisiones preliminares cargadas el 15-09-2026):
  - Eléctricos ...... ACAL-00102-LG-E-0004-1 1.xls      (act. preliminar de la Rev.1)
  - Instrum./IT-OT .. Listado de TIE-INs INS-IT_OT - Rev1.xls
Fuentes ANTERIORES (base de la Rev1 de este análisis):
  - ACAL-00102-LG-E-0004-1.xls  ·  Listado de TIE-INs INS-IT_OT.xls

El diff Rev1->Rev2 se calcula automáticamente comparando ambas revisiones celda a celda,
y se leen del archivo las marcas del emisor:
  · RESALTADO EN AMARILLO  -> ítem revisado / con definición abierta
  · TACHADO (strike-through) -> ítem que el emisor saca del listado de tie-in
Genera: informe ejecutivo (PDF y HTML), dashboard (XLSX) y cuadro de conteo/seguimiento (HTML).
No borra revisiones anteriores.
"""
import os, json, textwrap
import xlrd
from collections import Counter, OrderedDict

ENT="/home/user/OWN-Daro-1/Constructibilidad/Datos entrada"
GEN="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
REV="Rev2"; FECHA="2026-09-15"; REVANT="Rev1"; FECHAANT="2026-08-21"

EL_NEW=os.path.join(ENT,"ACAL-00102-LG-E-0004-1 1.xls")
EL_OLD=os.path.join(ENT,"ACAL-00102-LG-E-0004-1.xls")
IN_NEW=os.path.join(ENT,"Listado de TIE-INs INS-IT_OT - Rev1.xls")
IN_OLD=os.path.join(ENT,"Listado de TIE-INs INS-IT_OT.xls")

# ---------------- lectura con formato (amarillo / tachado) ----------------
YELLOW=13   # índice de color (255,255,0) en la paleta BIFF
def fmt(b,sh,r,c):
    try: xf=b.xf_list[sh.cell_xf_index(r,c)]
    except Exception: return (None,False)
    return (xf.background.pattern_colour_index, bool(b.font_list[xf.font_index].struck_out))

def marks(b,sh,r,cols):
    ye=False; st=False
    for c in cols:
        if c>=sh.ncols: continue
        fl,s=fmt(b,sh,r,c); v=str(sh.cell_value(r,c)).strip()
        if fl==YELLOW and v: ye=True
        if s and v: st=True
    return ye,st

# ---------------- extracción ELÉCTRICOS ----------------
EL_F=['ITEM','AREA','UBICACION','EQUIPO','TAREA','FUNCIÓN','ETAPA','CABLE','RESPONSABLE','N° PLANO','OBSERVACIONES','MATERIALES']
def extract_el(path):
    b=xlrd.open_workbook(path,formatting_info=True); sh=b.sheet_by_name("TIE IN electricos")
    hdr=5; colmap={}
    for ci in range(sh.ncols):
        v=str(sh.cell_value(hdr,ci)).strip().upper()
        if v and v not in colmap: colmap[v]=ci
    idx={k:colmap.get(k) for k in EL_F}
    keycols=[c for c in idx.values() if c is not None]
    out=[]; carry={}
    for ri in range(hdr+1,sh.nrows):
        gc=lambda k: str(sh.cell_value(ri,idx[k])).replace('\n',' ').strip() if idx.get(k) is not None else ''
        rec={k:gc(k) for k in idx}
        if not rec['ITEM']: continue
        # ¿la condición de paro es propia de la fila o heredada de la celda combinada de arriba?
        rec['_etapa_own']=bool(rec['ETAPA'])
        for k in ['AREA','UBICACION','EQUIPO','ETAPA','RESPONSABLE','TAREA','N° PLANO']:
            if rec[k]: carry[k]=rec[k]
            else: rec[k]=carry.get(k,'')
        rec['ITEM']=rec['ITEM'].replace('.0',''); rec['AREA']=rec['AREA'].replace('.0','')
        rec['_ye'],rec['_st']=marks(b,sh,ri,keycols)
        out.append(rec)
    return out

# ---------------- extracción INSTRUM. / IT-OT ----------------
IN_C=OrderedDict([('Area',0),('Item',2),('NoTieIn',4),('Tag',8),('Ubic',12),('Sistema',15),
                  ('Alcance',18),('Categoria',20),('Desc',22),('Resp',31),('Vendor',34),
                  ('Doc',37),('Obs',43),('ObsInsp',53)])
def extract_in(path,sheet,hdr):
    b=xlrd.open_workbook(path,formatting_info=True); sh=b.sheet_by_name(sheet)
    keycols=[c for c in IN_C.values() if c<sh.ncols]
    out=[]; carry={}
    for ri in range(hdr+1,sh.nrows):
        gc=lambda c: str(sh.cell_value(ri,c)).replace('\n',' ').strip() if c<sh.ncols else ''
        rec={k:gc(c) for k,c in IN_C.items()}
        if not rec['Item']: continue
        # 'Area' NO se arrastra: en esta revisión sólo T-O-005 la tiene vacía y
        # arrastrarla del ítem anterior le asignaría un área equivocada.
        for k in ['Ubic','Sistema','Alcance','Categoria','Resp','Vendor','Doc']:
            if rec[k]: carry[k]=rec[k]
            else: rec[k]=carry.get(k,'')
        rec['Item']=rec['Item'].replace('.0',''); rec['Area']=rec['Area'].replace('.0','')
        rec['_ye'],rec['_st']=marks(b,sh,ri,keycols)
        out.append(rec)
    return out

EL_N=extract_el(EL_NEW); EL_O=extract_el(EL_OLD)
IN_N=extract_in(IN_NEW,"Tie-In Instrumentos",0);  IN_O=extract_in(IN_OLD,"Tie-In Instrumentos",5)
OT_N=extract_in(IN_NEW,"Tie-In IT-OT",0);         OT_O=extract_in(IN_OLD,"Tie-In IT-OT",5)

# ---------------- clasificación ----------------
def paro_el(e):
    """(bucket, detalle, condicionamiento)"""
    E=e.strip().upper()
    if not E or E=='-': return 'A definir','-','A definir'
    base='Paro total' if E.startswith('PARO') else ('Sin paro' if E.startswith('SIN PARO') else None)
    if 'TCT' in E: return 'Con/sin paro (TCT)',e,'Con trabajo con tensión'
    if base is None: return 'A definir',e,'A definir'
    cond='Firme'
    U=E
    if 'A DEFINIR CON VENDOR' in U or 'CONSENSUAR CON EL VENDOR' in U: cond='A confirmar con el vendor'
    elif 'RELEVAMIENTO' in U: cond='A confirmar por relevamiento'
    elif '(T2)' in U: cond='Sujeto a condición operativa'
    return base,e,cond

def paro_cat(c):
    C=c.strip().upper()
    base=None
    if C.startswith('T3'): base='Paro total'
    elif C.startswith('T2'): base='Paro parcial'
    elif C.startswith('T1'): base='Sin paro'
    if base is None: return 'A definir',(c or '-'),'A definir',False
    cond='Firme'; pre=False
    if 'PREPARO' in C: cond='Trabajo de PRE-PARO'; pre=True
    elif 'PIPING' in C: cond='A confirmar con piping'
    det={'Paro total':'T3 · parada total','Paro parcial':'T2 · parada parcial','Sin paro':'T1 · cualquier momento'}[base]
    if c.strip() not in ('T1','T2','T3'): det+=f"  ({c.strip()})"
    return base,det,cond,pre

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
norm_area=lambda a:(a.replace('.0','').strip() or '—')

# ---------------- diff Rev1 -> Rev2 ----------------
CAMBIOS=[]   # filas del control de cambios detallado
def reg(esp,ident,campo,ant,nue,tipo):
    CAMBIOS.append(dict(esp=esp,id=ident,campo=campo,ant=ant,nue=nue,tipo=tipo))

rows=[]
EL_BLOQUE={}   # condición nueva -> ítems que la heredan por celda combinada
# ---- ELÉCTRICOS ----
EL_OM={r['ITEM']:r for r in EL_O}
for r in EL_N:
    p,pd_,cond=paro_el(r['ETAPA'])
    o=EL_OM.get(r['ITEM']); cambios=[]
    if o is None: estado='Nuevo'
    else:
        for k,lab in [('ETAPA','Condición de paro'),('TAREA','Tarea'),('FUNCIÓN','Función'),
                      ('CABLE','Cable'),('RESPONSABLE','Responsable'),('OBSERVACIONES','Observaciones'),
                      ('UBICACION','Ubicación'),('EQUIPO','Equipo')]:
            if r[k]!=o[k]:
                cambios.append(f"{lab}: «{o[k] or '—'}» → «{r[k] or '—'}»")
                # la condición de paro viene de una celda combinada: se registra una sola vez
                # en el ítem que la origina y luego se anota el alcance del bloque.
                if k=='ETAPA' and not r['_etapa_own']:
                    EL_BLOQUE.setdefault(r['ETAPA'],[]).append(r['ITEM'])
                else:
                    reg('Electricidad',f"EL-{r['ITEM']}",lab,o[k],r[k],'Modificado')
        estado='Modificado' if cambios else 'Sin cambios'
    if r['_st']: estado='Tachado — baja a confirmar'
    rows.append(dict(esp='Electricidad', id=f"EL-{r['ITEM']}", area=norm_area(r['AREA']),
        ubic=r['UBICACION'], sistema=r['EQUIPO'], desc=(r['TAREA'] or r['FUNCIÓN']), func=r['FUNCIÓN'],
        cable=r['CABLE'], nivel=nivel_el(r['FUNCIÓN'],r['TAREA'],r['EQUIPO']), alcance='—',
        paro=p, parodet=pd_, cond=cond, preparo=False, vendor='—', resp=r['RESPONSABLE'],
        doc=r['N° PLANO'], obs=r['OBSERVACIONES'], obsinsp='',
        estado=estado, tachado=bool(r['_st']), resaltado=bool(r['_ye']), cambios=cambios))

# los ítems eléctricos que heredan la condición por celda combinada se anotan
# en el cambio del ítem que la origina, en lugar de repetir la misma línea N veces
def _rango(items):
    return f"{items[0]}–{items[-1]}" if len(items)>2 else ", ".join(items)
for cond_nueva,its in EL_BLOQUE.items():
    for ch in CAMBIOS:
        if ch['esp']=='Electricidad' and ch['campo'].startswith('Condición de paro') and ch['nue']==cond_nueva:
            ch['campo']=f"Condición de paro (celda combinada: alcanza también a los ítems {_rango(its)})"
            break

# ---- INSTRUMENTACIÓN e IT/OT ----
def proc(new,old,esp,nivelf):
    om={x['NoTieIn']:x for x in old}
    om_item={(x['Area'],x['Item']):x for x in old}
    for r in new:
        p,pd_,cond,pre=paro_cat(r['Categoria'])
        o=om.get(r['NoTieIn']); cambios=[]; estado='Sin cambios'
        renum=None
        if o is None:
            o=om_item.get((r['Area'],r['Item']))
            # renumeración: mismo ítem/desc con otro N° de tie-in
            if o is None:
                cand=[x for x in old if x['Item']==r['Item'] and x['NoTieIn'] not in {y['NoTieIn'] for y in new}]
                if cand: o=cand[0]
            if o is not None:
                renum=o['NoTieIn']; estado='Renumerado'
                cambios.append(f"N° de tie-in: «{o['NoTieIn']}» → «{r['NoTieIn']}»")
                reg(esp,r['NoTieIn'],'N° de tie-in',o['NoTieIn'],r['NoTieIn'],'Renumerado')
        if o is None:
            estado='Nuevo'; reg(esp,r['NoTieIn'],'Alta','—',r['Desc'][:90],'Nuevo')
        else:
            for k,lab in [('Categoria','Categoría (paro)'),('Area','Área'),('Alcance','Alcance'),('Tag','Tag'),
                          ('Ubic','Ubicación'),('Sistema','Sistema'),('Desc','Descripción'),
                          ('Resp','Responsable'),('Vendor','Asistencia de vendor'),
                          ('Doc','Documento de ref.'),('Obs','Observaciones')]:
                if r[k]!=o[k]:
                    cambios.append(f"{lab}: «{o[k] or '—'}» → «{r[k] or '—'}»")
                    reg(esp,r['NoTieIn'],lab,o[k],r[k],'Modificado')
            if cambios and estado=='Sin cambios': estado='Modificado'
        if r['ObsInsp'] and not (o and o.get('ObsInsp')):
            cambios.append(f"Observación de inspección (columna nueva): «{r['ObsInsp']}»")
        if r['_st']:
            estado='Reclasificado a PRE-PARO (tachado)' if pre else 'Tachado — baja a confirmar'
        rows.append(dict(esp=esp, id=r['NoTieIn'], area=norm_area(r['Area']), ubic=r['Ubic'],
            sistema=r['Sistema'], desc=r['Desc'], func=r['Tag'], cable='',
            nivel=nivelf(r['Sistema']), alcance=ALC.get(r['Alcance'].upper(),r['Alcance'] or '—'),
            paro=p, parodet=pd_, cond=cond, preparo=pre,
            vendor=(r['Vendor'] if r['Vendor'] not in ('','-') else '—'), resp=r['Resp'],
            doc=r['Doc'], obs=r['Obs'], obsinsp=r['ObsInsp'],
            estado=estado, tachado=bool(r['_st']), resaltado=bool(r['_ye']), cambios=cambios))
    # bajas
    nn={x['NoTieIn'] for x in new}
    renumerados={c['ant'] for c in CAMBIOS if c['tipo']=='Renumerado'}
    for y in old:
        if y['NoTieIn'] not in nn and y['NoTieIn'] not in renumerados:
            reg(esp,y['NoTieIn'],'Baja',y['Desc'][:90],'—','Eliminado')

proc(IN_N,IN_O,'Instrumentación',nivel_in)
proc(OT_N,OT_O,'IT-OT',lambda s:'IT/OT Red/Comunic.')

# ---------------- métricas ----------------
TOTAL=len(rows)
ESP=['Electricidad','Instrumentación','IT-OT']
PARO=['Paro total','Paro parcial','Con/sin paro (TCT)','Sin paro','A definir']
NIVEL=['Potencia MT','Potencia BT','Control/Señales EL','PAT','Instrum. Control (PCS)','Instrum. Seguridad (SIS)','IT/OT Red/Comunic.']
ESTADOS=['Sin cambios','Modificado','Nuevo','Renumerado','Reclasificado a PRE-PARO (tachado)','Tachado — baja a confirmar']
COND=['Firme','A confirmar con el vendor','A confirmar por relevamiento','A confirmar con piping',
      'Sujeto a condición operativa','Trabajo de PRE-PARO','Con trabajo con tensión','A definir']
c_esp={e:sum(1 for r in rows if r['esp']==e) for e in ESP}
c_paro={p:sum(1 for r in rows if r['paro']==p) for p in PARO}
c_niv={n:sum(1 for r in rows if r['nivel']==n) for n in NIVEL}
c_est={e:sum(1 for r in rows if r['estado']==e) for e in ESTADOS}
c_cond={k:sum(1 for r in rows if r['cond']==k) for k in COND}
c_area=dict(sorted(Counter(r['area'] for r in rows).items()))
REQ_PARO=c_paro['Paro total']+c_paro['Paro parcial']
TACH=sum(1 for r in rows if r['tachado']); PRE=sum(1 for r in rows if r['preparo'])
EFECT=TOTAL-TACH
MODIF=sum(1 for r in rows if r['estado']!='Sin cambios')
ABIERTOS=sum(1 for r in rows if r['cond'].startswith('A confirmar') or r['cond']=='A definir')
VEND_PARO=sum(1 for r in rows if r["paro"]=="Paro total" and r["cond"]=="A confirmar con el vendor")
mx_ep={e:{p:sum(1 for r in rows if r['esp']==e and r['paro']==p) for p in PARO} for e in ESP}
pct=lambda n:f"{100*n/TOTAL:.0f}%"
dlt=lambda d:("0" if d==0 else f"{d:+d}")

# valores Rev1 (para el comparativo) — recalculados con la misma taxonomía
R1={'total':117,'paro_total':39,'paro_parcial':10,'tct':1,'sin_paro':64,'def':3,'req':49}

CHANGELOG=[
 ("Rev0","2026-08-18","Versión inicial. Solo listado ELÉCTRICO (ACAL-00102-LG-E-0004 Rev.0): 71 tie-in. Condición de paro con etapas PARO / PREPARO-POSPARO / POSPARO (50 en PARO)."),
 ("Rev1","2026-08-21","(1) Listado eléctrico Rev.1 con la condición de paro REDEFINIDA: PARO=24, SIN PARO=43, 'PARO o con TCT'=1, a definir=3. (2) Se INCORPORAN Instrumentación (34) e IT/OT (12) con Categoría T1/T2/T3 y Alcance. (3) Consolidado 117 tie-in, filtrable por especialidad, ¿paro?, nivel/tipo, alcance, área y vendor."),
 ("Rev2",FECHA,
  f"Se adoptan las revisiones preliminares del 15-09-2026 de ambos listados. (1) INSTRUMENTACIÓN: 12 ítems recategorizados a «T1 (preparo)» y TACHADOS por el emisor — trabajos previos al paro que dejan de contarse como tie-in propiamente dicho; 4 ítems bajan de T2 a T1 (it. 2, 28, 29 y 30); 2 quedan «T1 (ver categorización piping)»; 2 cambios de asistencia de vendor a HIMA/INAUCO; se incorpora la columna «Observaciones de inspección» con el relevamiento de campo de 22 ítems. "
  f"(2) IT/OT: T-O-940005 se renumera a T-O-005; T-O-900008 queda TACHADO (baja a confirmar); 4 ítems resaltados con definición abierta (patchera/switch de provisión INAUCO). "
  f"(3) ELÉCTRICOS: 5 condiciones de paro se precisan con su condicionamiento (2 «SIN PARO a definir con vendor», 1 «SIN PARO con relevamiento», 1 «PARO a consensuar con el vendor», 1 «SIN PARO (T2)» sujeto a condición operativa). "
  f"(4) Se agregan las dimensiones ESTADO EN LA REVISIÓN y CONDICIONAMIENTO, y el KPI de tie-in efectivos. Consolidado {TOTAL} ítems listados / {EFECT} efectivos; requieren paro {REQ_PARO} (Rev1: {R1['req']})."),
]

# ---------------- PDF ----------------
def build_pdf():
    from reportlab.lib.pagesizes import A4
    from reportlab.lib import colors
    from reportlab.lib.units import mm
    from reportlab.pdfgen import canvas
    path=os.path.join(GEN,f"Tie-in Consolidado - Informe Ejecutivo - {REV} - {FECHA}.pdf")
    W,H=A4; c=canvas.Canvas(path,pagesize=A4)
    NAVY=colors.HexColor('#1F3864'); RED=colors.HexColor('#C62828'); ORA=colors.HexColor('#EF6C00')
    GRN=colors.HexColor('#2E7D32'); GREY=colors.HexColor('#555555'); AMB=colors.HexColor('#B8860B')
    x0=15*mm; x1=W-15*mm
    def header():
        c.setFillColor(NAVY); c.rect(0,H-16*mm,W,16*mm,fill=1,stroke=0)
        c.setFillColor(colors.white); c.setFont("Helvetica-Bold",13.5)
        c.drawString(x0,H-11*mm,"TIE-IN CONSOLIDADO (EL · IN · IT/OT) — INFORME EJECUTIVO")
        c.setFont("Helvetica",8.5); c.drawRightString(x1,H-11*mm,f"{REV} · {FECHA}")
    header()
    y=H-22*mm; c.setFillColor(GREY); c.setFont("Helvetica",8)
    c.drawString(x0,y,"CPF-2 La Calera II · Fuentes: ACAL-00102-LG-E-0004 (EL) + Listado TIE-INs INS-IT/OT Rev.1 — revisiones preliminares 15-09-2026")
    y-=8*mm
    kpis=[("Tie-in listados",str(TOTAL),NAVY),("Efectivos (sin tachados)",str(EFECT),colors.HexColor('#1565C0')),
          ("Requieren paro",f"{REQ_PARO} ({pct(REQ_PARO)})",RED),("Modificados en Rev2",str(MODIF),AMB)]
    bw=(x1-x0-3*3*mm)/4
    for i,(lab,val,col) in enumerate(kpis):
        bx=x0+i*(bw+3*mm); c.setFillColor(colors.HexColor('#F2F4F7')); c.roundRect(bx,y-15*mm,bw,15*mm,3,fill=1,stroke=0)
        c.setFillColor(col); c.rect(bx,y-15*mm,2*mm,15*mm,fill=1,stroke=0)
        c.setFillColor(GREY); c.setFont("Helvetica",6.6); c.drawString(bx+3.5*mm,y-5*mm,lab)
        c.setFillColor(col); c.setFont("Helvetica-Bold",12.5); c.drawString(bx+3.5*mm,y-12.5*mm,val)
    y-=21*mm
    def ensure(yy,need):
        if yy-need<18*mm:
            c.showPage(); header(); return H-24*mm
        return yy
    def tabla(titulo,headers,data,widths,y,hlf=None,lead=4.6):
        y=ensure(y,(len(data)+3)*lead*mm if len(data)<14 else 60*mm)
        c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9.5); c.drawString(x0,y,titulo); y-=lead*mm
        tw=sum(widths); c.setFillColor(colors.HexColor('#37474F')); c.rect(x0,y-lead*mm,tw,lead*mm,fill=1,stroke=0)
        cx=x0; c.setFillColor(colors.white); c.setFont("Helvetica-Bold",7.2)
        for h,w in zip(headers,widths): c.drawString(cx+1.3*mm,y-3.3*mm,h); cx+=w
        y-=lead*mm; y0=y
        for ri,row in enumerate(data):
            hl=hlf and hlf(row)
            c.setFillColor(colors.HexColor('#FBE9E7') if hl else (colors.HexColor('#F7F9FB') if ri%2 else colors.white))
            c.rect(x0,y-lead*mm,tw,lead*mm,fill=1,stroke=0); cx=x0; c.setFillColor(colors.black)
            for j,(v,w) in enumerate(zip(row,widths)):
                c.setFont("Helvetica-Bold" if (j==0 or hl) else "Helvetica",7.2)
                if j==0: c.drawString(cx+1.3*mm,y-3.3*mm,str(v))
                else: c.drawRightString(cx+w-1.3*mm,y-3.3*mm,str(v))
                cx+=w
            y-=lead*mm
        c.setStrokeColor(colors.HexColor('#CCCCCC')); c.setLineWidth(0.4)
        c.rect(x0,y,tw,(y0-y)+lead*mm,fill=0,stroke=1)
        return y-5*mm
    W3=[70*mm,26*mm,24*mm]
    y=tabla("1 · Por especialidad",["Especialidad","Cantidad","%"],[[e,c_esp[e],pct(c_esp[e])] for e in ESP],W3,y)
    y=tabla("2 · ¿Requiere paro?  ·  comparativo Rev1 → Rev2",["Condición",f"{REVANT}",f"{REV}","Δ"],
            [["Paro total",R1['paro_total'],c_paro['Paro total'],dlt(c_paro['Paro total']-R1['paro_total'])],
             ["Paro parcial",R1['paro_parcial'],c_paro['Paro parcial'],dlt(c_paro['Paro parcial']-R1['paro_parcial'])],
             ["Con/sin paro (TCT)",R1['tct'],c_paro['Con/sin paro (TCT)'],dlt(c_paro['Con/sin paro (TCT)']-R1['tct'])],
             ["Sin paro",R1['sin_paro'],c_paro['Sin paro'],dlt(c_paro['Sin paro']-R1['sin_paro'])],
             ["A definir",R1['def'],c_paro['A definir'],dlt(c_paro['A definir']-R1['def'])],
             ["TOTAL que requiere paro",R1['req'],REQ_PARO,dlt(REQ_PARO-R1['req'])]],
            [58*mm,24*mm,24*mm,18*mm],y,hlf=lambda r:r[0].startswith('TOTAL'))
    W5=[46*mm,26*mm,26*mm,26*mm,22*mm]
    y=tabla("3 · ¿Requiere paro?  ·  cruce por especialidad",["Especialidad","Paro total","Paro parcial","Sin paro","Otros"],
            [[e,mx_ep[e]['Paro total'],mx_ep[e]['Paro parcial'],mx_ep[e]['Sin paro'],mx_ep[e]['Con/sin paro (TCT)']+mx_ep[e]['A definir']] for e in ESP],W5,y)
    y=tabla("4 · Estado de cada ítem en esta revisión",["Estado","Cantidad","%"],
            [[e,c_est[e],pct(c_est[e])] for e in ESTADOS if c_est[e]],[70*mm,26*mm,24*mm],y,
            hlf=lambda r:'achado' in r[0])
    y=tabla("5 · Condicionamiento de la condición de paro",["Condicionamiento","Cantidad","%"],
            [[k,c_cond[k],pct(c_cond[k])] for k in COND if c_cond[k]],[70*mm,26*mm,24*mm],y,
            hlf=lambda r:r[0].startswith('A confirmar') or r[0]=='A definir')
    y=tabla("6 · Por nivel / tipo de señal",["Nivel / tipo","Cantidad","%"],
            [[n,c_niv[n],pct(c_niv[n])] for n in NIVEL if c_niv[n]],W3,y)
    # control de cambios
    y=ensure(y,50*mm)
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9.5); c.drawString(x0,y,"7 · Control de cambios / Actualización"); y-=5*mm
    for rev,fch,txt in CHANGELOG:
        y=ensure(y,20*mm)
        c.setFillColor(RED if rev==REV else GREY); c.setFont("Helvetica-Bold",7.6); c.drawString(x0,y,f"{rev} · {fch}")
        c.setFillColor(colors.black); c.setFont("Helvetica",7.0)
        for ln in textwrap.wrap(txt,120):
            y=ensure(y,8*mm); c.drawString(x0+26*mm,y,ln); y-=3.6*mm
        y-=1.5*mm
    # detalle de cambios Rev1 -> Rev2
    y=ensure(y,40*mm)
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9.5)
    c.drawString(x0,y,f"8 · Detalle de cambios {REVANT} → {REV}  ({len(CAMBIOS)} modificaciones)"); y-=5*mm
    c.setFont("Helvetica",6.6)
    for ch in CAMBIOS:
        y=ensure(y,10*mm)
        c.setFillColor(colors.HexColor('#37474F')); c.setFont("Helvetica-Bold",6.8)
        c.drawString(x0,y,f"{ch['id']}"); c.setFont("Helvetica",6.6); c.setFillColor(colors.black)
        txt=f"{ch['campo']}:  {(ch['ant'] or '—')[:62]}   →   {(ch['nue'] or '—')[:62]}"
        c.drawString(x0+24*mm,y,txt[:135]); y-=3.4*mm
    # conclusiones
    y=ensure(y,45*mm); y-=3*mm
    c.setFillColor(NAVY); c.setFont("Helvetica-Bold",9.5); c.drawString(x0,y,"Conclusiones ejecutivas"); y-=4.6*mm
    concl=[
     f"• Se listan {TOTAL} tie-in; {TACH} quedan TACHADOS por el emisor, por lo que los tie-in EFECTIVOS son {EFECT}.",
     f"• Requieren paro (total + parcial) {REQ_PARO} de {TOTAL} ({pct(REQ_PARO)}); en la {REVANT} eran {R1['req']}: bajan {R1['req']-REQ_PARO} por la recategorización de T2 a T1 en Instrumentación.",
     f"• {PRE} ítems de Instrumentación pasan a «T1 (preparo)»: son trabajos PREVIOS al paro, no tie-in de paro; deben planificarse antes de la ventana de parada.",
     f"• {ABIERTOS} ítems tienen la condición de paro CONDICIONADA (a confirmar con vendor, por relevamiento o con piping) — son el frente de definición más urgente.",
     f"• ATENCIÓN: {VEND_PARO} de los {c_paro['Paro total']} tie-in de PARO TOTAL son eléctricos del bloque de SALA DE GENERACIÓN y quedan «a consensuar con el vendor» — es decir, la mayor parte de la carga de la parada eléctrica todavía no está cerrada.",
     f"• Paro total por especialidad: EL={mx_ep['Electricidad']['Paro total']} · IN={mx_ep['Instrumentación']['Paro total']} · IT/OT={mx_ep['IT-OT']['Paro total']}.",
     f"• IT/OT sigue siendo el frente más condicionado: {mx_ep['IT-OT']['Paro total']} de {c_esp['IT-OT']} requieren parada total.",
    ]
    c.setFont("Helvetica",7.4); c.setFillColor(colors.black)
    for ln in concl:
        y=ensure(y,8*mm)
        for i,seg in enumerate(textwrap.wrap(ln,132)):
            c.drawString(x0+(0 if i==0 else 2.5*mm),y,seg); y-=3.8*mm
    c.setFont("Helvetica-Oblique",6.4); c.setFillColor(GREY)
    c.drawString(x0,12*mm,"Categorías IN/IT-OT: T1=cualquier momento (sin paro) · T2=parada parcial · T3=parada total · «(preparo)»=trabajo previo al paro. EL: PARO / SIN PARO / 'PARO o con TCT' / a definir.")
    c.drawString(x0,9*mm,"Marcas del emisor leídas del archivo: resaltado amarillo = ítem revisado con definición abierta · tachado = ítem que sale del listado de tie-in. Revisiones preliminares, sujetas a emisión formal.")
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
    ws["A2"]="Fuentes: ACAL-00102-LG-E-0004 (EL) + Listado TIE-INs INS-IT/OT Rev.1 — revisiones preliminares 15-09-2026 · CPF-2 La Calera II"
    ws.merge_cells("A2:F2"); ws["A2"].font=Font(italic=True,size=9,color="555555")
    kp=[("Listados",TOTAL,NAVY),("Efectivos",EFECT,"1565C0"),("Requiere paro",REQ_PARO,RED),
        ("Tachados",TACH,"8D6E63"),("Pre-paro",PRE,"00838F"),("Modificados Rev2",MODIF,"B8860B")]
    for i,(lab,val,col) in enumerate(kp):
        L=chr(65+i)
        ws[f"{L}4"]=lab; ws[f"{L}4"].font=Font(size=9,color="FFFFFF",bold=True)
        ws[f"{L}4"].fill=PatternFill("solid",fgColor=col); ws[f"{L}4"].alignment=Alignment(horizontal="center")
        ws[f"{L}5"]=val; ws[f"{L}5"].font=Font(size=18,bold=True,color=col); ws[f"{L}5"].alignment=Alignment(horizontal="center")
    r=[7]
    def block(title,headers,data,hl=None):
        ws.cell(r[0],1,title).font=Font(bold=True,size=11,color=NAVY); r[0]+=1
        for j,h in enumerate(headers):
            cc=ws.cell(r[0],1+j,h); cc.font=Font(bold=True,color="FFFFFF",size=9)
            cc.fill=PatternFill("solid",fgColor=HEAD); cc.alignment=Alignment(horizontal="center" if j else "left",wrap_text=True); cc.border=B
        r[0]+=1
        for row in data:
            for j,v in enumerate(row):
                cc=ws.cell(r[0],1+j,v); cc.border=B; cc.alignment=Alignment(horizontal="center" if j else "left",wrap_text=(j==2 and title.startswith("Control")))
                if j==0: cc.font=Font(bold=True,size=9)
                if hl and hl(row): cc.fill=PatternFill("solid",fgColor="FBE9E7"); cc.font=Font(bold=True,size=9,color=RED if j else "000000")
            r[0]+=1
        r[0]+=1
    block("Por especialidad",["Especialidad","Cantidad","%"],[[e,c_esp[e],pct(c_esp[e])] for e in ESP])
    block(f"¿Requiere paro? · comparativo {REVANT} → {REV}",["Condición",REVANT,REV,"Δ"],
          [["Paro total",R1['paro_total'],c_paro['Paro total'],c_paro['Paro total']-R1['paro_total']],
           ["Paro parcial",R1['paro_parcial'],c_paro['Paro parcial'],c_paro['Paro parcial']-R1['paro_parcial']],
           ["Con/sin paro (TCT)",R1['tct'],c_paro['Con/sin paro (TCT)'],c_paro['Con/sin paro (TCT)']-R1['tct']],
           ["Sin paro",R1['sin_paro'],c_paro['Sin paro'],c_paro['Sin paro']-R1['sin_paro']],
           ["A definir",R1['def'],c_paro['A definir'],c_paro['A definir']-R1['def']],
           ["TOTAL requiere paro",R1['req'],REQ_PARO,REQ_PARO-R1['req']]],
          hl=lambda row:str(row[0]).startswith('TOTAL'))
    block("¿Requiere paro? · por especialidad",["Especialidad","Paro total","Paro parcial","Sin paro","TCT/def."],
          [[e,mx_ep[e]['Paro total'],mx_ep[e]['Paro parcial'],mx_ep[e]['Sin paro'],mx_ep[e]['Con/sin paro (TCT)']+mx_ep[e]['A definir']] for e in ESP])
    block("Estado en esta revisión",["Estado","Cantidad","%"],
          [[e,c_est[e],pct(c_est[e])] for e in ESTADOS if c_est[e]],hl=lambda row:'achado' in str(row[0]))
    block("Condicionamiento de la condición de paro",["Condicionamiento","Cantidad","%"],
          [[k,c_cond[k],pct(c_cond[k])] for k in COND if c_cond[k]],
          hl=lambda row:str(row[0]).startswith('A confirmar') or row[0]=='A definir')
    block("Por nivel / tipo de señal",["Nivel / tipo","Cantidad"],[[n,c_niv[n]] for n in NIVEL if c_niv[n]])
    block("Por área",["Área","Cantidad"],[[a,n] for a,n in c_area.items()])
    block("Control de cambios",["Rev","Fecha","Detalle"],[[rev,fch,txt] for rev,fch,txt in CHANGELOG])
    for col,w in zip("ABCDEF",[44,14,16,12,12,12]): ws.column_dimensions[col].width=w

    # ---- Listado ----
    ls=wb.create_sheet("Listado")
    HED=["Especialidad","ID / N° Tie-In","Área","Ubicación","Sistema/Equipo","Descripción / tarea",
         "Nivel / tipo","Alcance","¿Requiere paro?","Detalle categoría","Condicionamiento",
         "Estado en Rev2","Vendor","Responsable","Doc. ref.","Cable","Observaciones",
         "Observaciones de inspección","Cambios Rev1→Rev2","Fecha ejec.","Estado real"]
    ls.append(HED)
    for j,_ in enumerate(HED,1):
        cc=ls.cell(1,j); cc.font=Font(bold=True,color="FFFFFF",size=9)
        cc.fill=PatternFill("solid",fgColor=HEAD); cc.alignment=Alignment(horizontal="center",wrap_text=True); cc.border=B
    ESPF={'Electricidad':'C62828','Instrumentación':'1565C0','IT-OT':'6A1B9A'}
    PARF={'Paro total':'C62828','Paro parcial':'EF6C00','Con/sin paro (TCT)':'8D6E63','Sin paro':'2E7D32','A definir':'90A4AE'}
    ESTF={'Sin cambios':None,'Modificado':'FFF3E0','Nuevo':'E8F5E9','Renumerado':'E3F2FD',
          'Reclasificado a PRE-PARO (tachado)':'FFF9C4','Tachado — baja a confirmar':'FFE0B2'}
    for d in rows:
        ls.append([d['esp'],d['id'],d['area'],d['ubic'],d['sistema'],d['desc'],d['nivel'],d['alcance'],
                   d['paro'],d['parodet'],d['cond'],d['estado'],d['vendor'],d['resp'],d['doc'],d['cable'],
                   d['obs'],d['obsinsp'],"  |  ".join(d['cambios']),'',''])
        rr=ls.max_row
        ls.cell(rr,1).fill=PatternFill("solid",fgColor=ESPF[d['esp']]); ls.cell(rr,1).font=Font(color="FFFFFF",bold=True,size=8)
        ls.cell(rr,9).fill=PatternFill("solid",fgColor=PARF[d['paro']]); ls.cell(rr,9).font=Font(color="FFFFFF",bold=True,size=8)
        f=ESTF.get(d['estado'])
        if f: ls.cell(rr,12).fill=PatternFill("solid",fgColor=f); ls.cell(rr,12).font=Font(bold=True,size=8)
        if d['tachado']:
            for j in (2,6): ls.cell(rr,j).font=Font(strike=True,bold=(j==2),size=8.5)
        if d['cond'].startswith('A confirmar') or d['cond']=='A definir':
            ls.cell(rr,11).fill=PatternFill("solid",fgColor="FFF9C4"); ls.cell(rr,11).font=Font(bold=True,size=8,color="B8860B")
        for j in range(1,len(HED)+1): ls.cell(rr,j).border=B
    for col,w in zip("ABCDEFGHIJKLMNOPQRSTU",[14,15,7,16,18,34,20,15,16,20,22,26,12,11,18,14,30,26,40,11,12]):
        ls.column_dimensions[col].width=w
    ls.freeze_panes="A2"; ls.auto_filter.ref=f"A1:U{ls.max_row}"

    # ---- Cambios ----
    cs=wb.create_sheet("Cambios Rev1-Rev2")
    cs["A1"]=f"DETALLE DE CAMBIOS {REVANT} ({FECHAANT}) → {REV} ({FECHA})"; cs.merge_cells("A1:E1")
    cs["A1"].font=Font(bold=True,size=12,color="FFFFFF"); cs["A1"].fill=PatternFill("solid",fgColor=NAVY)
    cs["A2"]="Diferencias detectadas automáticamente entre las revisiones de los listados de origen."
    cs.merge_cells("A2:E2"); cs["A2"].font=Font(italic=True,size=9,color="555555")
    HC=["Especialidad","ID / N° Tie-In","Campo","Valor anterior","Valor nuevo","Tipo"]
    for j,h in enumerate(HC,1):
        cc=cs.cell(4,j,h); cc.font=Font(bold=True,color="FFFFFF",size=9)
        cc.fill=PatternFill("solid",fgColor=HEAD); cc.alignment=Alignment(horizontal="center",wrap_text=True); cc.border=B
    TIPF={'Modificado':'FFF3E0','Nuevo':'E8F5E9','Eliminado':'FFCDD2','Renumerado':'E3F2FD'}
    rr=5
    for ch in CAMBIOS:
        for j,v in enumerate([ch['esp'],ch['id'],ch['campo'],ch['ant'],ch['nue'],ch['tipo']],1):
            cc=cs.cell(rr,j,v); cc.border=B; cc.alignment=Alignment(wrap_text=True,vertical="top")
            cc.font=Font(size=8.5,bold=(j==2))
        cs.cell(rr,6).fill=PatternFill("solid",fgColor=TIPF.get(ch['tipo'],"FFFFFF"))
        cs.cell(rr,6).font=Font(bold=True,size=8.5)
        rr+=1
    for col,w in zip("ABCDEF",[15,16,22,58,58,14]): cs.column_dimensions[col].width=w
    cs.freeze_panes="A5"; cs.auto_filter.ref=f"A4:F{cs.max_row}"

    # ---- Marcas del emisor ----
    ms=wb.create_sheet("Marcas del emisor")
    ms["A1"]="ÍTEMS MARCADOS EN EL LISTADO DE ORIGEN (resaltado amarillo / tachado)"; ms.merge_cells("A1:F1")
    ms["A1"].font=Font(bold=True,size=12,color="FFFFFF"); ms["A1"].fill=PatternFill("solid",fgColor=NAVY)
    ms["A2"]="Amarillo = ítem revisado con definición abierta · Tachado = ítem que el emisor saca del listado de tie-in."
    ms.merge_cells("A2:F2"); ms["A2"].font=Font(italic=True,size=9,color="555555")
    HM=["Especialidad","ID / N° Tie-In","Resaltado","Tachado","Estado en Rev2","Descripción / tarea"]
    for j,h in enumerate(HM,1):
        cc=ms.cell(4,j,h); cc.font=Font(bold=True,color="FFFFFF",size=9)
        cc.fill=PatternFill("solid",fgColor=HEAD); cc.alignment=Alignment(horizontal="center"); cc.border=B
    rr=5
    for d in rows:
        if not (d['tachado'] or d['resaltado']): continue
        for j,v in enumerate([d['esp'],d['id'],"AMARILLO" if d['resaltado'] else "","TACHADO" if d['tachado'] else "",d['estado'],d['desc']],1):
            cc=ms.cell(rr,j,v); cc.border=B; cc.font=Font(size=8.5,bold=(j==2)); cc.alignment=Alignment(wrap_text=(j==6),vertical="top")
        if d['resaltado']: ms.cell(rr,3).fill=PatternFill("solid",fgColor="FFFF00")
        if d['tachado']:
            ms.cell(rr,4).fill=PatternFill("solid",fgColor="FFE0B2")
            ms.cell(rr,2).font=Font(strike=True,bold=True,size=8.5)
        rr+=1
    for col,w in zip("ABCDEF",[15,16,12,12,34,70]): ms.column_dimensions[col].width=w
    ms.freeze_panes="A5"; ms.auto_filter.ref=f"A4:F{ms.max_row}"
    wb.save(path); return path

# ---------------- HTML dashboard ----------------
def build_html():
    path=os.path.join(GEN,f"Dashboard Tie-in Consolidado - {REV} - {FECHA}.html")
    keys=['esp','id','area','ubic','sistema','desc','func','cable','nivel','alcance','paro','parodet',
          'cond','preparo','vendor','resp','doc','obs','obsinsp','estado','tachado','resaltado','cambios']
    data=[{k:d[k] for k in keys} for d in rows]
    html=(HTML_TMPL
        .replace("__REV__",REV).replace("__FECHA__",FECHA).replace("__REVANT__",REVANT)
        .replace("__TOTAL__",str(TOTAL)).replace("__EFECT__",str(EFECT))
        .replace("__REQ__",str(REQ_PARO)).replace("__PCTREQ__",pct(REQ_PARO))
        .replace("__REQANT__",str(R1['req'])).replace("__DREQ__",dlt(REQ_PARO-R1['req']))
        .replace("__TACH__",str(TACH)).replace("__PRE__",str(PRE))
        .replace("__MODIF__",str(MODIF)).replace("__ABIERTOS__",str(ABIERTOS))
        .replace("__DATA__",json.dumps(data,ensure_ascii=False))
        .replace("__CAMBIOS__",json.dumps(CAMBIOS,ensure_ascii=False))
        .replace("__CHLOG__",json.dumps([{"rev":r,"fch":f,"txt":t} for r,f,t in CHANGELOG],ensure_ascii=False))
        .replace("__R1__",json.dumps(R1)))
    open(path,"w",encoding="utf-8").write(html); return path

HTML_TMPL=r"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1"><title>Tie-in Consolidado · Dashboard __REV__</title>
<style>
:root{--bg:#f4f6f8;--card:#fff;--ink:#1f2933;--muted:#607080;--line:#e2e8f0;--red:#C62828;--ora:#EF6C00;--blue:#1565C0;--green:#2E7D32;--navy:#1F3864;--pur:#6A1B9A;--brown:#8D6E63;--teal:#00838F;--amber:#B8860B;--yel:#FFF9C4;--shadow:0 1px 3px rgba(0,0,0,.08)}
@media(prefers-color-scheme:dark){:root{--bg:#0f141a;--card:#1a222c;--ink:#e6edf3;--muted:#93a1b0;--line:#2a343f;--yel:#4a441c;--shadow:0 1px 3px rgba(0,0,0,.5)}}
*{box-sizing:border-box}body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--ink);font-size:14px}
header{background:var(--navy);color:#fff;padding:15px 22px}h1{margin:0;font-size:19px}.sub{opacity:.85;font-size:12px;margin-top:3px}
.wrap{max-width:1400px;margin:0 auto;padding:16px 22px 60px}
.kpis{display:flex;gap:12px;flex-wrap:wrap;margin:14px 0}
.kpi{flex:1;min-width:140px;background:var(--card);border:1px solid var(--line);border-left:4px solid var(--navy);border-radius:10px;padding:12px 14px;box-shadow:var(--shadow)}
.kpi.red{border-left-color:var(--red)}.kpi.green{border-left-color:var(--green)}.kpi.blue{border-left-color:var(--blue)}
.kpi.brown{border-left-color:var(--brown)}.kpi.teal{border-left-color:var(--teal)}.kpi.amber{border-left-color:var(--amber)}
.kpi .l{font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}
.kpi .v{font-size:23px;font-weight:800;margin-top:2px}.kpi .d{font-size:11px;color:var(--muted);margin-top:2px}
.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(280px,1fr));gap:12px;margin:6px 0}
.card{background:var(--card);border:1px solid var(--line);border-radius:10px;padding:13px;box-shadow:var(--shadow)}.card h3{margin:0 0 9px;font-size:12.5px}
.brow{display:flex;align-items:center;gap:8px;margin:4px 0;font-size:12px}
.brow .nm{width:148px;flex-shrink:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
.brow .barwrap{flex:1;background:var(--line);border-radius:5px;height:15px;overflow:hidden}
.brow .bar{display:block;height:100%;border-radius:5px;min-width:2px}.brow .n{width:40px;text-align:right;font-weight:700}
details{margin:12px 0;background:var(--card);border:1px solid var(--line);border-radius:10px;padding:6px 14px;box-shadow:var(--shadow)}
summary{cursor:pointer;font-weight:700;font-size:13px;padding:6px 0}
.chg{font-size:12px;margin:6px 0;padding-left:10px;border-left:3px solid var(--line)}.chg b{color:var(--red)}
.filters{display:flex;gap:10px;flex-wrap:wrap;align-items:end;margin:14px 0 6px}
.fg{display:flex;flex-direction:column;gap:3px}.fg label{font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.03em}
select,input[type=text]{padding:6px 9px;border:1px solid var(--line);border-radius:7px;background:var(--card);color:var(--ink);font-size:12.5px;min-width:130px}
.count{font-size:12px;color:var(--muted);margin:6px 0}
.btn{padding:6px 12px;border:1px solid var(--line);border-radius:7px;background:var(--card);color:var(--ink);cursor:pointer;font-size:12px;font-weight:600}
.btn.on{background:var(--navy);color:#fff;border-color:var(--navy)}
table{width:100%;border-collapse:collapse;background:var(--card);border-radius:10px;overflow:hidden;box-shadow:var(--shadow);font-size:11.5px}
th,td{padding:6px 8px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}
th{background:rgba(127,127,127,.08);position:sticky;top:0;font-size:10px;text-transform:uppercase;color:var(--muted);z-index:2}
tr.tach td:nth-child(2),tr.tach td:nth-child(6){text-decoration:line-through;opacity:.75}
tr.tach td:nth-child(6) .mini{text-decoration:none;opacity:1}
tr.tach{background:rgba(141,110,99,.10)}tr.res{box-shadow:inset 3px 0 0 #FFD54F}
.tag{color:#fff;border-radius:4px;padding:1px 6px;font-size:10px;font-weight:700;white-space:nowrap;display:inline-block}
.pt{background:var(--red)}.pp{background:var(--ora)}.pn{background:var(--green)}.pd{background:#90A4AE}.px{background:var(--brown)}
.mini{font-size:10px;color:var(--muted);display:block;margin-top:2px}
.cnd{background:var(--yel);color:var(--amber);border-radius:4px;padding:1px 5px;font-size:10px;font-weight:700}
.chip{display:inline-block;background:rgba(127,127,127,.12);border-radius:4px;padding:1px 6px;font-size:10px;margin:1px 2px 1px 0}
.legend{font-size:11px;color:var(--muted);margin-top:10px;line-height:1.6}
.ctab{width:100%;border-collapse:collapse;font-size:11.5px}.ctab td,.ctab th{padding:5px 7px;border-bottom:1px solid var(--line)}
.ctab .old{color:var(--muted);text-decoration:line-through}.ctab .new{color:var(--green);font-weight:600}
</style></head><body>
<header><h1>Tie-in Consolidado (Electricidad · Instrumentación · IT/OT) — Dashboard de conteo y seguimiento</h1>
<div class="sub">CPF-2 La Calera II · Fuentes: ACAL-00102-LG-E-0004 (EL) + Listado TIE-INs INS-IT/OT Rev.1 — revisiones preliminares 15-09-2026 · <b>__REV__ · __FECHA__</b></div></header>
<div class="wrap">
 <div class="kpis">
  <div class="kpi"><div class="l">Tie-in listados</div><div class="v">__TOTAL__</div></div>
  <div class="kpi blue"><div class="l">Efectivos (sin tachados)</div><div class="v">__EFECT__</div><div class="d">__TACH__ tachados por el emisor</div></div>
  <div class="kpi red"><div class="l">Requieren paro</div><div class="v">__REQ__ <span style="font-size:14px">(__PCTREQ__)</span></div><div class="d">__REVANT__: __REQANT__ · <b>__DREQ__</b></div></div>
  <div class="kpi teal"><div class="l">Trabajos de pre-paro</div><div class="v">__PRE__</div><div class="d">antes de la ventana de parada</div></div>
  <div class="kpi amber"><div class="l">Condición a confirmar</div><div class="v">__ABIERTOS__</div><div class="d">vendor / relevamiento / piping</div></div>
  <div class="kpi green"><div class="l">Modificados en __REV__</div><div class="v">__MODIF__</div></div>
 </div>
 <div class="grid">
  <div class="card"><h3>Por especialidad</h3><div id="c-esp"></div></div>
  <div class="card"><h3>¿Requiere paro?</h3><div id="c-paro"></div></div>
  <div class="card"><h3>Estado en esta revisión</h3><div id="c-est"></div></div>
  <div class="card"><h3>Condicionamiento de la condición de paro</h3><div id="c-cond"></div></div>
  <div class="card"><h3>Por nivel / tipo de señal</h3><div id="c-niv"></div></div>
  <div class="card"><h3>Por área</h3><div id="c-area"></div></div>
 </div>
 <details open><summary>🔄 Comparativo __REVANT__ → __REV__ · condición de paro</summary><div id="comp"></div></details>
 <details><summary>📝 Control de cambios / Actualización (por revisión)</summary><div id="chlog"></div></details>
 <details><summary>📋 Detalle de cambios __REVANT__ → __REV__ (ítem por ítem)</summary><div id="cambios"></div></details>
 <div class="filters" id="filters"></div>
 <div class="count" id="count"></div>
 <div style="max-height:600px;overflow:auto;border-radius:10px">
 <table id="tbl"><thead><tr>
  <th>Esp.</th><th>ID / N° Tie-In</th><th>Área</th><th>Ubicación</th><th>Sistema/Equipo</th>
  <th>Descripción / tarea</th><th>Nivel / tipo</th><th>Alcance</th><th>¿Paro?</th>
  <th>Estado __REV__</th><th>Vendor</th>
 </tr></thead><tbody></tbody></table></div>
 <div class="legend">
  <b>Condición de paro</b> — EL: PARO / SIN PARO / «PARO o con TCT» (trabajo con tensión) / a definir. IN·IT-OT: T1 = cualquier momento (sin paro) · T2 = parada parcial · T3 = parada total · «(preparo)» = trabajo previo al paro.<br>
  <b>Marcas del emisor leídas del archivo</b> — <span style="box-shadow:inset 3px 0 0 #FFD54F;padding-left:8px">barra amarilla</span> = ítem resaltado (definición abierta) · <span style="text-decoration:line-through">tachado</span> = ítem que el emisor saca del listado de tie-in.<br>
  Revisiones preliminares, sujetas a emisión formal de ambos listados.
 </div>
</div>
<script>
const DATA=__DATA__, CHLOG=__CHLOG__, CAMBIOS=__CAMBIOS__, R1=__R1__;
const ESPC={'Electricidad':'var(--red)','Instrumentación':'var(--blue)','IT-OT':'var(--pur)'};
const PAROC={'Paro total':'var(--red)','Paro parcial':'var(--ora)','Con/sin paro (TCT)':'var(--brown)','Sin paro':'var(--green)','A definir':'#90A4AE'};
const PAROCLS={'Paro total':'pt','Paro parcial':'pp','Con/sin paro (TCT)':'px','Sin paro':'pn','A definir':'pd'};
const NIVC={'Potencia MT':'var(--pur)','Potencia BT':'var(--red)','Control/Señales EL':'var(--ora)','PAT':'var(--brown)','Instrum. Control (PCS)':'var(--blue)','Instrum. Seguridad (SIS)':'var(--teal)','IT/OT Red/Comunic.':'var(--green)'};
const ESTC={'Sin cambios':'#90A4AE','Modificado':'var(--ora)','Nuevo':'var(--green)','Renumerado':'var(--blue)','Reclasificado a PRE-PARO (tachado)':'var(--teal)','Tachado — baja a confirmar':'var(--brown)'};
const CONDC={'Firme':'var(--green)','A confirmar con el vendor':'var(--amber)','A confirmar por relevamiento':'var(--amber)','A confirmar con piping':'var(--amber)','Sujeto a condición operativa':'var(--ora)','Trabajo de PRE-PARO':'var(--teal)','Con trabajo con tensión':'var(--brown)','A definir':'#90A4AE'};
const ORD={esp:['Electricidad','Instrumentación','IT-OT'],
 paro:['Paro total','Paro parcial','Con/sin paro (TCT)','Sin paro','A definir'],
 nivel:['Potencia MT','Potencia BT','Control/Señales EL','PAT','Instrum. Control (PCS)','Instrum. Seguridad (SIS)','IT/OT Red/Comunic.'],
 estado:['Sin cambios','Modificado','Nuevo','Renumerado','Reclasificado a PRE-PARO (tachado)','Tachado — baja a confirmar'],
 cond:['Firme','A confirmar con el vendor','A confirmar por relevamiento','A confirmar con piping','Sujeto a condición operativa','Trabajo de PRE-PARO','Con trabajo con tensión','A definir']};
function bars(el,pairs,colf){const mx=Math.max(...pairs.map(p=>p[1]),1);
 el.innerHTML=pairs.filter(p=>p[1]).map(([k,n])=>`<div class="brow"><span class="nm" title="${k}">${k}</span><span class="barwrap"><span class="bar" style="width:${100*n/mx}%;background:${colf(k)}"></span></span><span class="n">${n}</span></div>`).join('');}
const cnt=f=>{const m={};DATA.forEach(d=>m[f(d)]=(m[f(d)]||0)+1);return m;};
const esp=cnt(d=>d.esp),paro=cnt(d=>d.paro),niv=cnt(d=>d.nivel),est=cnt(d=>d.estado),cond=cnt(d=>d.cond),area=cnt(d=>d.area);
bars(document.getElementById('c-esp'),ORD.esp.map(k=>[k,esp[k]||0]),k=>ESPC[k]);
bars(document.getElementById('c-paro'),ORD.paro.map(k=>[k,paro[k]||0]),k=>PAROC[k]);
bars(document.getElementById('c-est'),ORD.estado.map(k=>[k,est[k]||0]),k=>ESTC[k]);
bars(document.getElementById('c-cond'),ORD.cond.map(k=>[k,cond[k]||0]),k=>CONDC[k]);
bars(document.getElementById('c-niv'),ORD.nivel.map(k=>[k,niv[k]||0]),k=>NIVC[k]);
bars(document.getElementById('c-area'),Object.keys(area).sort().map(k=>[k,area[k]]),()=>'var(--navy)');
// comparativo
const COMP=[['Paro total',R1.paro_total,paro['Paro total']||0],['Paro parcial',R1.paro_parcial,paro['Paro parcial']||0],
 ['Con/sin paro (TCT)',R1.tct,paro['Con/sin paro (TCT)']||0],['Sin paro',R1.sin_paro,paro['Sin paro']||0],
 ['A definir',R1.def,paro['A definir']||0]];
const reqN=(paro['Paro total']||0)+(paro['Paro parcial']||0);
document.getElementById('comp').innerHTML=`<table class="ctab"><tr><th>Condición</th><th>__REVANT__</th><th>__REV__</th><th>Δ</th></tr>`+
 COMP.map(([k,a,b])=>{const d=b-a;
   // en 'Sin paro' un aumento es favorable; en el resto, lo favorable es que baje
   const bueno = (k==='Sin paro') ? d>0 : d<0;
   const col = d===0?'var(--muted)':(bueno?'var(--green)':'var(--red)');
   return `<tr><td>${k}</td><td>${a}</td><td><b>${b}</b></td><td style="color:${col}">${d>0?'+':''}${d}</td></tr>`;}).join('')+
 `<tr style="background:rgba(198,40,40,.08)"><td><b>TOTAL que requiere paro</b></td><td>${R1.req}</td><td><b>${reqN}</b></td><td style="color:${reqN-R1.req<0?'var(--green)':'var(--red)'}"><b>${reqN-R1.req>0?'+':''}${reqN-R1.req}</b></td></tr></table>`;
document.getElementById('chlog').innerHTML=CHLOG.map(c=>`<div class="chg"><b>${c.rev} · ${c.fch}</b> — ${c.txt}</div>`).join('');
document.getElementById('cambios').innerHTML=`<div style="max-height:380px;overflow:auto"><table class="ctab">
 <tr><th>Esp.</th><th>ID</th><th>Campo</th><th>Anterior</th><th>Nuevo</th><th>Tipo</th></tr>`+
 CAMBIOS.map(c=>`<tr><td>${c.esp==='Instrumentación'?'IN':(c.esp==='IT-OT'?'IT/OT':'EL')}</td><td><b>${c.id}</b></td><td>${c.campo}</td>
  <td class="old">${(c.ant||'—').slice(0,160)}</td><td class="new">${(c.nue||'—').slice(0,160)}</td><td>${c.tipo}</td></tr>`).join('')+
 `</table></div><div style="font-size:11px;color:var(--muted);margin:8px 0">${CAMBIOS.length} modificaciones detectadas entre revisiones.</div>`;
// filtros
const DIMS=[['esp','Especialidad'],['paro','¿Requiere paro?'],['estado','Estado en __REV__'],['cond','Condicionamiento'],
            ['nivel','Nivel / tipo'],['alcance','Alcance'],['area','Área'],['vendor','Vendor']];
const state={};
const uniq=k=>{const u=[...new Set(DATA.map(d=>d[k]))];return ORD[k]?ORD[k].filter(v=>u.includes(v)):u.sort();};
const fdiv=document.getElementById('filters');
DIMS.forEach(([k,lab])=>{const g=document.createElement('div');g.className='fg';
 g.innerHTML=`<label>${lab}</label><select data-k="${k}"><option value="">Todos</option>${uniq(k).map(v=>`<option>${v}</option>`).join('')}</select>`;fdiv.appendChild(g);});
const sg=document.createElement('div');sg.className='fg';sg.innerHTML=`<label>Buscar</label><input type="text" id="q" placeholder="tag, sistema, descripción…">`;fdiv.appendChild(sg);
const tg=document.createElement('div');tg.className='fg';tg.innerHTML=`<label>Vistas rápidas</label><span>
 <button class="btn" id="b-req">Solo requieren paro</button>
 <button class="btn" id="b-cam">Solo con cambios</button>
 <button class="btn" id="b-ef">Solo efectivos</button>
 <button class="btn" id="reset">Limpiar</button></span>`;fdiv.appendChild(tg);
fdiv.querySelectorAll('select').forEach(s=>s.onchange=()=>{state[s.dataset.k]=s.value;render();});
document.getElementById('q').addEventListener('input',e=>{state.q=e.target.value.toLowerCase().trim();render();});
const tog=(id,key)=>{const b=document.getElementById(id);b.onclick=()=>{state[key]=!state[key];b.classList.toggle('on',state[key]);render();};};
tog('b-req','onlyreq');tog('b-cam','onlycam');tog('b-ef','onlyef');
document.getElementById('reset').onclick=()=>{fdiv.querySelectorAll('select').forEach(s=>s.value='');
 document.getElementById('q').value='';Object.keys(state).forEach(k=>delete state[k]);
 ['b-req','b-cam','b-ef'].forEach(i=>document.getElementById(i).classList.remove('on'));render();};
function render(){
 const tb=document.querySelector('#tbl tbody');
 const rows=DATA.filter(d=>{
   for(const [k] of DIMS){if(state[k]&&d[k]!==state[k])return false;}
   if(state.onlyreq&&!(d.paro==='Paro total'||d.paro==='Paro parcial'))return false;
   if(state.onlycam&&d.estado==='Sin cambios')return false;
   if(state.onlyef&&d.tachado)return false;
   if(state.q){const s=(d.id+' '+d.sistema+' '+d.desc+' '+d.func+' '+d.ubic+' '+d.cable+' '+d.obs+' '+d.obsinsp).toLowerCase();if(!s.includes(state.q))return false;}
   return true;});
 const req=rows.filter(d=>d.paro==='Paro total'||d.paro==='Paro parcial').length;
 document.getElementById('count').textContent=`${rows.length} de ${DATA.length} tie-in  ·  ${req} requieren paro  ·  ${rows.filter(d=>d.tachado).length} tachados  ·  ${rows.filter(d=>d.estado!=='Sin cambios').length} con cambios`;
 tb.innerHTML=rows.map(d=>`<tr class="${d.tachado?'tach':''} ${d.resaltado?'res':''}">
  <td><span class="tag" style="background:${ESPC[d.esp]}">${d.esp==='Instrumentación'?'IN':(d.esp==='IT-OT'?'IT/OT':'EL')}</span></td>
  <td><b>${d.id}</b></td><td>${d.area}</td><td>${d.ubic}</td><td>${d.sistema}</td>
  <td>${d.desc}${d.obsinsp?`<span class="mini">🔎 ${d.obsinsp}</span>`:''}${d.cambios.length?`<span class="mini" style="color:var(--amber)">✎ ${d.cambios.join(' · ')}</span>`:''}</td>
  <td>${d.nivel}</td><td>${d.alcance}</td>
  <td><span class="tag ${PAROCLS[d.paro]}" title="${d.parodet}">${d.paro}</span>${d.cond!=='Firme'?`<span class="mini"><span class="cnd">${d.cond}</span></span>`:''}</td>
  <td><span class="chip" style="color:${ESTC[d.estado]};font-weight:700">${d.estado}</span></td>
  <td>${d.vendor}</td></tr>`).join('');
}
render();
</script></body></html>"""

# ---------------- HTML informe ejecutivo ----------------
def build_html_informe():
    path=os.path.join(GEN,f"Tie-in Consolidado - Informe Ejecutivo - {REV} - {FECHA}.html")
    def tabla(headers,data,hl=None):
        h="".join(f"<th>{x}</th>" for x in headers)
        b=""
        for row in data:
            cls=' class="hl"' if (hl and hl(row)) else ''
            b+=f"<tr{cls}>"+"".join(f"<td{' class=num' if j else ''}>{v}</td>" for j,v in enumerate(row))+"</tr>"
        return f"<table class='rep'><thead><tr>{h}</tr></thead><tbody>{b}</tbody></table>"
    t1=tabla(["Especialidad","Cantidad","%"],[[e,c_esp[e],pct(c_esp[e])] for e in ESP])
    t2=tabla(["Condición",REVANT,REV,"Δ"],
        [["Paro total",R1['paro_total'],c_paro['Paro total'],dlt(c_paro['Paro total']-R1['paro_total'])],
         ["Paro parcial",R1['paro_parcial'],c_paro['Paro parcial'],dlt(c_paro['Paro parcial']-R1['paro_parcial'])],
         ["Con/sin paro (TCT)",R1['tct'],c_paro['Con/sin paro (TCT)'],dlt(c_paro['Con/sin paro (TCT)']-R1['tct'])],
         ["Sin paro",R1['sin_paro'],c_paro['Sin paro'],dlt(c_paro['Sin paro']-R1['sin_paro'])],
         ["A definir",R1['def'],c_paro['A definir'],dlt(c_paro['A definir']-R1['def'])],
         ["<b>TOTAL que requiere paro</b>",R1['req'],f"<b>{REQ_PARO}</b>",f"<b>{dlt(REQ_PARO-R1['req'])}</b>"]],
        hl=lambda r:'TOTAL' in str(r[0]))
    t3=tabla(["Especialidad","Paro total","Paro parcial","Sin paro","Otros"],
        [[e,mx_ep[e]['Paro total'],mx_ep[e]['Paro parcial'],mx_ep[e]['Sin paro'],mx_ep[e]['Con/sin paro (TCT)']+mx_ep[e]['A definir']] for e in ESP])
    t4=tabla(["Estado en esta revisión","Cantidad","%"],[[e,c_est[e],pct(c_est[e])] for e in ESTADOS if c_est[e]],
             hl=lambda r:'achado' in str(r[0]))
    t5=tabla(["Condicionamiento","Cantidad","%"],[[k,c_cond[k],pct(c_cond[k])] for k in COND if c_cond[k]],
             hl=lambda r:str(r[0]).startswith('A confirmar') or r[0]=='A definir')
    t6=tabla(["Nivel / tipo de señal","Cantidad","%"],[[n,c_niv[n],pct(c_niv[n])] for n in NIVEL if c_niv[n]])
    tc=tabla(["Esp.","ID","Campo","Valor anterior","Valor nuevo","Tipo"],
        [[c['esp'][:2].upper() if c['esp']=='IT-OT' else ('IN' if c['esp']=='Instrumentación' else 'EL'),
          f"<b>{c['id']}</b>",c['campo'],
          f"<span class='old'>{(c['ant'] or '—')[:150]}</span>",
          f"<span class='new'>{(c['nue'] or '—')[:150]}</span>",c['tipo']] for c in CAMBIOS])
    chl="".join(f"<div class='chg'><b>{r} · {f}</b> — {t}</div>" for r,f,t in CHANGELOG)
    concl="".join(f"<li>{x}</li>" for x in [
      f"Se listan <b>{TOTAL}</b> tie-in; <b>{TACH}</b> quedan TACHADOS por el emisor, por lo que los tie-in <b>efectivos son {EFECT}</b>.",
      f"Requieren paro (total + parcial) <b>{REQ_PARO}</b> de {TOTAL} ({pct(REQ_PARO)}); en la {REVANT} eran {R1['req']} — bajan <b>{R1['req']-REQ_PARO}</b> por la recategorización de T2 a T1 en Instrumentación.",
      f"<b>{PRE}</b> ítems de Instrumentación pasan a «T1 (preparo)»: son trabajos <b>previos al paro</b>, no tie-in de paro. Deben planificarse antes de la ventana de parada y liberan carga de la parada.",
      f"<b>{ABIERTOS}</b> ítems tienen la condición de paro <b>condicionada</b> (a confirmar con el vendor, por relevamiento o con piping): es el frente de definición más urgente.",
      f"<b>Atención:</b> {VEND_PARO} de los {c_paro['Paro total']} tie-in de <b>paro total</b> son eléctricos del bloque de SALA DE GENERACIÓN y quedan «a consensuar con el vendor»: la mayor parte de la carga de la parada eléctrica <b>todavía no está cerrada</b>.",
      f"Paro total por especialidad: EL={mx_ep['Electricidad']['Paro total']} · IN={mx_ep['Instrumentación']['Paro total']} · IT/OT={mx_ep['IT-OT']['Paro total']}.",
      f"IT/OT sigue siendo el frente más condicionado: {mx_ep['IT-OT']['Paro total']} de {c_esp['IT-OT']} requieren parada total, con definiciones abiertas de provisión de patchera y switches (INAUCO).",
      "Ambos listados son <b>revisiones preliminares</b>: las marcas de resaltado y tachado del emisor se mantienen visibles en el dashboard para su confirmación en la emisión formal.",
    ])
    html=f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>Tie-in Consolidado · Informe Ejecutivo {REV}</title>
<style>
:root{{--bg:#f4f6f8;--card:#fff;--ink:#1f2933;--muted:#607080;--line:#e2e8f0;--red:#C62828;--ora:#EF6C00;--blue:#1565C0;--green:#2E7D32;--navy:#1F3864;--amber:#B8860B;--shadow:0 1px 3px rgba(0,0,0,.08)}}
@media(prefers-color-scheme:dark){{:root{{--bg:#0f141a;--card:#1a222c;--ink:#e6edf3;--muted:#93a1b0;--line:#2a343f;--shadow:0 1px 3px rgba(0,0,0,.5)}}}}
*{{box-sizing:border-box}}body{{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;background:var(--bg);color:var(--ink);font-size:14px}}
header{{background:var(--navy);color:#fff;padding:16px 24px}}h1{{margin:0;font-size:19px}}.sub{{opacity:.85;font-size:12px;margin-top:4px}}
.wrap{{max-width:1080px;margin:0 auto;padding:18px 24px 60px}}
h2{{font-size:14px;margin:26px 0 8px;color:var(--navy);border-bottom:2px solid var(--line);padding-bottom:5px}}
@media(prefers-color-scheme:dark){{h2{{color:#8fb4ee}}}}
.kpis{{display:flex;gap:12px;flex-wrap:wrap;margin:16px 0}}
.kpi{{flex:1;min-width:150px;background:var(--card);border:1px solid var(--line);border-left:4px solid var(--navy);border-radius:10px;padding:12px 14px;box-shadow:var(--shadow)}}
.kpi.red{{border-left-color:var(--red)}}.kpi.blue{{border-left-color:var(--blue)}}.kpi.amber{{border-left-color:var(--amber)}}
.kpi .l{{font-size:10.5px;color:var(--muted);text-transform:uppercase;letter-spacing:.04em}}
.kpi .v{{font-size:24px;font-weight:800;margin-top:2px}}.kpi .d{{font-size:11px;color:var(--muted)}}
table.rep{{width:100%;border-collapse:collapse;background:var(--card);border-radius:8px;overflow:hidden;box-shadow:var(--shadow);font-size:12px}}
table.rep th,table.rep td{{padding:6px 9px;border-bottom:1px solid var(--line);text-align:left}}
table.rep th{{background:#37474F;color:#fff;font-size:10.5px;text-transform:uppercase}}
table.rep td.num{{text-align:right}}table.rep tr.hl{{background:rgba(198,40,40,.09);font-weight:700}}
.old{{color:var(--muted);text-decoration:line-through}}.new{{color:var(--green);font-weight:600}}
.chg{{font-size:12px;margin:8px 0;padding-left:11px;border-left:3px solid var(--line)}}.chg b{{color:var(--red)}}
ul.concl{{font-size:12.5px;line-height:1.75;padding-left:18px}}
.note{{font-size:11px;color:var(--muted);margin-top:22px;line-height:1.6;border-top:1px solid var(--line);padding-top:10px}}
details{{margin:10px 0;background:var(--card);border:1px solid var(--line);border-radius:8px;padding:6px 14px;box-shadow:var(--shadow)}}
summary{{cursor:pointer;font-weight:700;font-size:12.5px;padding:6px 0}}
</style></head><body>
<header><h1>Tie-in Consolidado (EL · IN · IT/OT) — Informe Ejecutivo</h1>
<div class="sub">CPF-2 La Calera II · Fuentes: ACAL-00102-LG-E-0004 (EL) + Listado TIE-INs INS-IT/OT Rev.1 — revisiones preliminares 15-09-2026 · <b>{REV} · {FECHA}</b></div></header>
<div class="wrap">
 <div class="kpis">
  <div class="kpi"><div class="l">Tie-in listados</div><div class="v">{TOTAL}</div></div>
  <div class="kpi blue"><div class="l">Efectivos</div><div class="v">{EFECT}</div><div class="d">{TACH} tachados por el emisor</div></div>
  <div class="kpi red"><div class="l">Requieren paro</div><div class="v">{REQ_PARO} <span style="font-size:14px">({pct(REQ_PARO)})</span></div><div class="d">{REVANT}: {R1['req']} · {REQ_PARO-R1['req']:+d}</div></div>
  <div class="kpi amber"><div class="l">Condición a confirmar</div><div class="v">{ABIERTOS}</div><div class="d">vendor / relevamiento / piping</div></div>
 </div>
 <h2>1 · Por especialidad</h2>{t1}
 <h2>2 · ¿Requiere paro? · comparativo {REVANT} → {REV}</h2>{t2}
 <h2>3 · ¿Requiere paro? · cruce por especialidad</h2>{t3}
 <h2>4 · Estado de cada ítem en esta revisión</h2>{t4}
 <h2>5 · Condicionamiento de la condición de paro</h2>{t5}
 <h2>6 · Por nivel / tipo de señal</h2>{t6}
 <h2>7 · Control de cambios / Actualización</h2>{chl}
 <h2>8 · Detalle de cambios {REVANT} → {REV} ({len(CAMBIOS)} modificaciones)</h2>
 <details open><summary>Ver el detalle ítem por ítem</summary><div style="max-height:460px;overflow:auto">{tc}</div></details>
 <h2>Conclusiones ejecutivas</h2><ul class="concl">{concl}</ul>
 <div class="note">
  Categorías IN/IT-OT: T1 = cualquier momento (sin paro) · T2 = parada parcial · T3 = parada total · «(preparo)» = trabajo previo al paro.
  EL: PARO / SIN PARO / «PARO o con TCT» (trabajo con tensión) / a definir.<br>
  Marcas del emisor leídas del archivo de origen: resaltado amarillo = ítem revisado con definición abierta · tachado = ítem que el emisor saca del listado de tie-in.
  Ambos listados son revisiones preliminares, sujetas a emisión formal.
 </div>
</div></body></html>"""
    open(path,"w",encoding="utf-8").write(html); return path

p1=build_pdf(); p2=build_xlsx(); p3=build_html(); p4=build_html_informe()
print(f"TOTAL={TOTAL}  efectivos={EFECT}  tachados={TACH}  preparo={PRE}  modificados={MODIF}  cond.abiertas={ABIERTOS}")
print(" por esp:",c_esp)
print(" paro   :",c_paro,"| requiere paro:",REQ_PARO,f"(Rev1 {R1['req']})")
print(" estado :",{k:v for k,v in c_est.items() if v})
print(" cond   :",{k:v for k,v in c_cond.items() if v})
print(" cambios detectados:",len(CAMBIOS))
print("OK\n ",p1,"\n ",p2,"\n ",p3,"\n ",p4)
