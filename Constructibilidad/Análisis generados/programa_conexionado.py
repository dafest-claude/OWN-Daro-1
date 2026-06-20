# -*- coding: utf-8 -*-
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter
from datetime import date, timedelta
import os

BASE="/home/user/OWN-Daro-1/Constructibilidad/Análisis generados"
OUT=os.path.join(BASE,"Programa_Conexionado_EL_IN_CPF2_Rev3.xlsx")
def d(y,m,dd): return date(y,m,dd)
def meom(dt):  # ultimo dia del mes
    nx=date(dt.year+(dt.month//12),(dt.month%12)+1,1); return nx-timedelta(days=1)

wb=openpyxl.Workbook(); ws=wb.active; ws.title="Cronograma"
NAVY="1F3B63"; ORANGE="ED7D31"; BLUE="2E75B6"; DKOR="C55A11"; GREEN="548235"; REDc="C00000"; GREY="808080"
thin=Side(style="thin",color="D0D0D0"); border=Border(left=thin,right=thin,top=thin,bottom=thin)

# Titulo
ws["A1"]="C5551 - LA CALERA 2  ·  Determinación de plazos para conexionado  ·  Electricidad e Instrumentación"
ws["A1"].font=Font(bold=True,size=13,color=NAVY)
ws["A2"]="Reconstrucción del programa (imagen original) · RFSU 31-ENE-2028 · Rev.3 · 2026-06-14 · Gantt SEMANAL · barras pintadas según Inicio/Fin · REFERENCIA original en amarillo (ver hoja 'Ref_Imagen_Original')"
ws["A2"].font=Font(size=9,italic=True,color="C00000")

# Semanas (columnas) + encabezado de meses agrupado
weeks=[]
wk=d(2026,6,1)
while wk<=d(2028,2,29):
    weeks.append(wk); wk=wk+timedelta(days=7)
MONROW=3; HDRROW=4; COL0=11  # K
heads=["Activity ID","Activity Name","Dur","Start","Finish","TF","Cant","hhs","Q/sm","hhs/Q"]
for j,h in enumerate(heads,1):
    cc=ws.cell(HDRROW,j,h); cc.font=Font(bold=True,color="FFFFFF",size=9); cc.fill=PatternFill("solid",fgColor=NAVY)
    cc.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); cc.border=border
# fila de semanas (grilla) + agrupado por mes (merge)
navyfill=PatternFill("solid",fgColor=NAVY)
seg_start=0
for k,w in enumerate(weeks):
    cc=ws.cell(HDRROW,COL0+k); cc.fill=navyfill; cc.border=Border(left=Side(style="thin",color="3A5A85"),bottom=thin)
    # cambio de mes -> cerrar segmento anterior y rotular
    if k>0 and (w.year,w.month)!=(weeks[k-1].year,weeks[k-1].month):
        c1=get_column_letter(COL0+seg_start); c2=get_column_letter(COL0+k-1)
        ws.merge_cells("%s%d:%s%d"%(c1,MONROW,c2,MONROW))
        mc=ws.cell(MONROW,COL0+seg_start,weeks[seg_start]); mc.number_format="mmm-yy"
        mc.font=Font(bold=True,color="FFFFFF",size=8); mc.fill=navyfill; mc.alignment=Alignment(horizontal="center")
        seg_start=k
# ultimo segmento
c1=get_column_letter(COL0+seg_start); c2=get_column_letter(COL0+len(weeks)-1)
ws.merge_cells("%s%d:%s%d"%(c1,MONROW,c2,MONROW))
mc=ws.cell(MONROW,COL0+seg_start,weeks[seg_start]); mc.number_format="mmm-yy"
mc.font=Font(bold=True,color="FFFFFF",size=8); mc.fill=navyfill; mc.alignment=Alignment(horizontal="center")

# (id,name,dur,start,finish,tf,cant,hhs,qsm,hhsq,section,indent)
R=[
 ("","REFERENCIA — CRONOGRAMA ORIGINAL (amarillo · aprox · ver hoja imagen)","","","","","","","","","REF",0),
 ("","PHs (Permisos / Hold Points)","",d(2027,8,1),d(2027,9,30),"","","","","","REF",1),
 ("","Llegan canalizaciones Eléctricas","",d(2026,10,1),d(2027,3,31),"","","","","","REF",1),
 ("","Llegan canalizaciones Instrumentación","",d(2026,7,1),d(2026,10,31),"","","","","","REF",1),
 ("","Llegan cables Instrumentación","",d(2026,9,1),d(2026,10,31),"","","","","","REF",1),
 ("","Llegan cables Eléctricos","",d(2026,9,1),d(2027,5,31),"","","","","","REF",1),
 ("","Llegada de Instrumentos","",d(2026,8,1),d(2027,5,31),"","","","","","REF",1),
 ("","Fundaciones","",d(2026,7,1),d(2027,4,30),"","","","","","REF",1),
 ("","EE.MM. (equipos electromecánicos)","",d(2026,7,1),d(2027,8,31),"","","","","","REF",1),
 ("","Canalizaciones Eléctricas","",d(2026,10,1),d(2026,12,31),"","","","","","REF",1),
 ("","Canalizaciones Instrumentación","",d(2026,11,1),d(2027,1,31),"","","","","","REF",1),
 ("","Montaje de Instrumentos (original)","",d(2027,1,1),d(2027,9,30),"","","","","","REF",1),
 # Montaje instrumentos
 ("","Montaje de Instrumentos",222.9,d(2027,2,10),d(2027,10,28),"","","","","","INST",0),
 ("","Válvulas","",d(2027,2,10),d(2027,10,28),"",362,7682,"",21.22,"INST",1),
 ("","Instrumentos","",d(2027,2,10),d(2027,10,28),"",838,11643,"",13.80,"INST",1),
 # Cables IN
 ("","Tendido y conex. cables Instrumentación",226.3,d(2027,4,3),d(2027,12,23),"","","","","","CABIN",0),
 ("","Tendido","",d(2027,4,3),d(2027,12,23),"",58713,29967,1816,0.51,"CABIN",1),
 ("","Conexionado","",d(2027,4,3),d(2027,12,23),"",12684,6890,1308,0.54,"CABIN",1),
 # Cables EL
 ("","Tendido y conex. cables Eléctricos",156.9,d(2027,5,30),d(2027,11,29),"","","","","","CABEL",0),
 ("","Tendido","",d(2027,5,30),d(2027,11,29),"",93555,39862,4175,0.43,"CABEL",1),
 ("","Conexionado","",d(2027,5,30),d(2027,11,29),"",7137,5315,1062,0.74,"CABEL",1),
 # Montajes salas/equipos
 ("CS023090","Montaje estructura de soporte / Sala INS",35,d(2027,2,17),d(2027,3,29),59,"","","","","SALAS",0),
 ("CE014050","Montaje de shelter - Sala Eléctrica # 3",38,d(2027,5,28),d(2027,7,10),"",8,"","","","SALAS",0),
 ("CE018050","Montaje de shelter - Sala Eléctrica # 4 (AJUST)",48,d(2027,5,14),d(2027,6,22),"",8,"","","","SALAS",0),
 ("CDE1180","Montaje de tableros y equipos eléctricos",54,d(2027,7,12),d(2027,9,11),"",106,"","","","SALAS",0),
 ("CE014060","Montaje de Transf/Gen Emerg - Salas Eléctricas",20,d(2027,7,23),d(2027,8,14),"",80,"","","","SALAS",0),
 # Precom / comisionado
 ("CDW1000","Precomisionado",220,d(2027,5,3),d(2028,1,13),"",0,"","","","PRECOM",0),
 ("","Precom sistema eléctrico (bomba-motor-PMS)","",d(2027,5,3),d(2028,1,13),"","","","","","PRECOM",1),
 ("","Precom SCADA / PCS / PSS / ESD / F&G","",d(2027,8,1),d(2028,1,13),"","","","","","PRECOM",1),
 ("CDW1010","Comisionado",200,d(2027,7,16),d(2028,1,31),"",0,"","","","PRECOM",0),
 ("","Propuesta 1 — Energización","",d(2027,10,15),d(2027,10,31),"","","","","","PRECOM",1),
 ("","Propuesta 2 — Energización","",d(2027,11,15),d(2027,11,30),"","","","","","PRECOM",1),
]
SECCOL={"REF":"FFC000","INST":ORANGE,"CABIN":BLUE,"CABEL":DKOR,"SALAS":GREEN,"PRECOM":REDc}
r0=HDRROW+1; sec_rows={}
for i,(aid,name,dur,st,fi,tf,cant,hhs,qsm,hhsq,sec,ind) in enumerate(R):
    r=r0+i; sec_rows.setdefault(sec,[]).append(r)
    vals=[aid,("   "*ind)+name,dur,st,fi,tf,cant,hhs,qsm,hhsq]
    for j,v in enumerate(vals,1):
        cc=ws.cell(r,j,v if v!="" else None); cc.border=border; cc.font=Font(size=9,bold=(ind==0 and name and aid=="" and sec!="HITOS"))
        if j in (4,5) and isinstance(v,date): cc.number_format="dd-mmm-yy"; cc.alignment=Alignment(horizontal="center")
        if j in (3,6,7,8,9,10): cc.alignment=Alignment(horizontal="center")
    # color etiqueta de seccion (header rows)
    if ind==0 and (aid=="" or sec in("SALAS","PRECOM")):
        ws.cell(r,2).font=Font(size=9,bold=True,color=SECCOL[sec])
    fill=PatternFill("solid",fgColor=SECCOL[sec])
    for k,w in enumerate(weeks):
        cc=ws.cell(r,COL0+k); cc.border=border
        # pintar la barra si la semana [w, w+6] se solapa con [start, finish]
        if isinstance(st,date) and isinstance(fi,date) and st<=w+timedelta(days=6) and fi>=w:
            cc.fill=fill
# anchos
for col,wd in zip("ABCDEFGHIJ",[10,40,6,10,10,5,8,9,7,7]): ws.column_dimensions[col].width=wd
for k in range(len(weeks)): ws.column_dimensions[get_column_letter(COL0+k)].width=2.0
ws.freeze_panes=ws.cell(r0,COL0)
ws.row_dimensions[1].height=18

# ===== Hoja 2: Datos / Productividad =====
ws2=wb.create_sheet("Datos_Productividad")
ws2["A1"]="DATOS DE PRODUCTIVIDAD Y PUNTAS — transcripción de la imagen (VERIFICAR)"; ws2["A1"].font=Font(bold=True,size=12,color=NAVY)
ws2["A2"]="Avance referencia:  ELEC 75%   ·   INS 60%"; ws2["A2"].font=Font(bold=True,color="C00000")
def block(ws,r0,title,headers,rows):
    ws.cell(r0,1,title).font=Font(bold=True,color="FFFFFF"); 
    for c in range(1,len(headers)+1): ws.cell(r0,c).fill=PatternFill("solid",fgColor=NAVY)
    for j,h in enumerate(headers,1):
        cc=ws.cell(r0+1,j,h); cc.font=Font(bold=True,size=8); cc.fill=PatternFill("solid",fgColor="DCE6F2"); cc.alignment=Alignment(wrap_text=True,horizontal="center")
    for i,row in enumerate(rows):
        for j,v in enumerate(row,1): ws.cell(r0+2+i,j,v)
    return r0+2+len(rows)+1
n=block(ws2,4,"ORIGINAL",["Frente","Hhs/día","N°Pers Base 10/d","Rendim. (Q/día)","standard AESA (Q/día)","Improductividad","VMDS"],
 [["Frente 1",132.43,13.24,259.46,0.32,1.59,0],["Frente 2",9.13,0.91,186.84,1.20,0.45,None],
  ["Frente 3",254.13,25.41,596.43,0.35,1.22,0],["Frente 4",10.17,1.02,151.67,0.97,0.77,None]])
n=block(ws2,n,"REFERENCIAS (prod / hhs / cuadrillas)",
 ["Frente","P80 B (Q/hhs)","NEC B (Q/hhs)","P80 C (Q/hhs)","MEGA (Q/hhs)","LAS (Q/sem)","Duración aj. (días)","hhs VMDS","hhs NEC B","hhs Combinadas","N°Pers VMDS","N°Pers NECB","N°Pers COMB","Q/sem"],
 [["Frente 1",0.37,0.54,0.47,1000,9600,176.00,21489,31940,27830,12,18,16,2002],
  ["Frente 2",None,7.71,None,330,1377,52.80,None,None,97756,None,None,185,1441],
  ["Frente 3",0.26,0.82,0.63,1219,6800,168.00,24324,77089,59127,14,46,35,3341],
  ["Frente 4",None,6.47,None,210,1219,50.40,None,None,46162,None,None,92,850]])
n=block(ws2,n,"LUEGO DE INSTALADAS LAS SALAS",
 ["Item","Alcance d/montadas las salas (puntas)","Duración entre salas y energización (días)","Alcance entre salas (puntas)","Alcance entre salas (hhs)","N°Pers Base 10/d (ORIG)","N°Pers Base 10/d (NECB)","Q/sem"],
 [["Caso 1",7610,132,4134,58653,3,44,346],["Caso 2",5353,96,3986,34622,4,36,335],["Caso 3 (50% shelter 3)",2676,60,1993,17311,3,29,268]])
ws2.cell(n+1,1,"Notas: 'Considerando un 50% de puntas para el shelter 3'.  'Plazo remanente entre el completamiento del montaje de shelter #3 hasta la propuesta #1 de energización'.").font=Font(italic=True,size=8)
for col,w in zip("ABCDEFGHIJKLMN",[20,12,12,12,11,11,11,11,11,11,9,9,9,9]): ws2.column_dimensions[col].width=w

# ===== Hoja 3: Notas =====
ws3=wb.create_sheet("Notas_Hipotesis")
notas=["NOTAS E HIPÓTESIS (de la imagen original)","",
 "• RFSU objetivo: 31-ENE-2028. Cronograma MUY ajustado.",
 "• 'Se deben adelantar estas actividades y gestionar PHs → Re-instalar Instrumentos' (montaje de instrumentos).",
 "• Tendido cables: anotaciones 'Inicio c/20%' y 'Inicio c/50%' (arranque de tendido con avance parcial de canalizaciones).",
 "• Dos propuestas de ENERGIZACIÓN: Propuesta 1 (~OCT-27) y Propuesta 2 (~NOV-27).",
 "• Precomisionado: sistema eléctrico (bomba-motor-PMS) + SCADA/PCS/PSS/ESD/F&G.",
 "• Montaje de shelters: Sala INS, Sala Eléctrica #3, Sala Eléctrica #4 (AJUST).",
 "• 'Hipótesis para comprobar' (nota del autor).","",
 "PENDIENTE DE VERIFICAR contra la imagen:",
 "  - Valores de hhs / Q/sm / hhs/Q de cada frente.",
 "  - Fechas exactas de los hitos previos (canalizaciones, cables, instrumentos, fundaciones, EE.MM.).",
 "  - Tablas de productividad (ORIGINAL / REFERENCIAS / LUEGO DE INSTALADAS LAS SALAS).",
 "  - Asignación de los 4 'frentes' a las actividades."]
for i,t in enumerate(notas,1):
    cc=ws3.cell(i,1,t); cc.font=Font(bold=(i==1 or t.endswith(":")),size=11 if i==1 else 9)
ws3.column_dimensions["A"].width=120

# ===== Hoja 4: imagen real del cronograma ORIGINAL (referencia fiel) =====
ws4=wb.create_sheet("Ref_Imagen_Original")
ws4["A1"]="CRONOGRAMA ORIGINAL — recorte fiel de la imagen (referencia para análisis futuro)"
ws4["A1"].font=Font(bold=True,size=12,color=NAVY)
ws4["A2"]="Las barras amarillas de la hoja 'Cronograma' aproximan estas barras. Eje: JUN-26 → ENE-28."
ws4["A2"].font=Font(italic=True,size=9)
# generar el recorte desde la imagen fuente (autónomo)
refpng=os.path.join(BASE,"_ref_original.png")
SRC=os.path.join(os.path.dirname(BASE),"Datos entrada","Imagen PNG.png")
try:
    from PIL import Image as PILImage
    src=PILImage.open(SRC); w,h=src.size
    src.crop((int(w*0.295),0,int(w*0.66),int(h*0.42))).save(refpng)
except Exception as e:
    print("aviso: no se pudo generar recorte:",e)
if os.path.exists(refpng):
    from openpyxl.drawing.image import Image as XLImage
    img=XLImage(refpng); img.anchor="A4"; ws4.add_image(img)

wb.save(OUT)
try: os.remove(refpng)
except OSError: pass
print("OK ->",OUT)
