const pptxgen = require("pptxgenjs");

const NAVY="1E2761", INK="1F2933", MUT="667788", LINE="E4E8EE", CARD="F5F7FA";
const RED="C62828", AMB="B8860B", GRN="2E7D32";

// Provisiones ordenadas por fecha de entrega (Cronograma Ejecutivo Rev11, 27-10-2026)
const P=[
 {c:"33691E", n:"CAT Miron — Transformadores", e:"11-DIC-2026", av:null,
  nota:"13 unidades · ingeniería 100%, fabr. 0%", nc:MUT},
 {c:"6A1B9A", n:"HPH — Marshalling (Alemania)",        e:"15-DIC-2026", av:100,
  nota:"Fabricación completada · DAP Ezeiza", nc:GRN},
 {c:"EF6C00", n:"ABB — PMS",                            e:"17-FEB-2027", av:53,
  nota:"Adelantado 9 puntos sobre el plan", nc:GRN},
 {c:"AD1457", n:"ABB — Sala eléctrica SE#4 (MT)",       e:"25-MAR-2027", av:36,
  nota:"Contractual 08-04-27 · 14 d de margen", nc:MUT},
 {c:"C62828", n:"ABB — Sala eléctrica SE#3 (BT)",       e:"20-ABR-2027", av:36,
  nota:"Contractual 14-05-27 · 21 d de margen", nc:MUT},
 {c:"1565C0", n:"Inauco — PCS / SCADA",                 e:"04-MAY-2027", av:6,
  nota:"Rev 0 para construcción · sin corrimientos", nc:MUT},
 {c:"00838F", n:"HARG — HIMA Argentina (SIS)",          e:"10-MAY-2027", av:0,
  nota:"ORDEN DE COMPRA PENDIENTE", nc:RED},
 {c:"795548", n:"ABB — Ductos de barras",               e:"08-JUL-2027", av:5,
  nota:"PLAZO SIN CONFIRMAR", nc:RED},
];

const KPI=[
 {l:"Cierre del camino crítico", v:"10-MAY-2027", d:"OC HARG → HPH → HARG → Inauco", c:NAVY},
 {l:"Última entrega del paquete", v:"08-JUL-2027", d:"Ductos de barras — sin confirmar", c:RED},
 {l:"Entregas dentro de 2026",    v:"2 de 8",      d:"Transformadores y marshalling", c:GRN},
 {l:"Puntos de atención abiertos",v:"14",          d:"11 en severidad alta", c:AMB},
];

const ACC=[
 {t:"Colocar la OC de HIMA Argentina", d:"Sin colocar. Gobierna el cierre del 10-05-27: cada semana de demora lo corre una semana.", u:"PA-04"},
 {t:"Emitir la carta de crédito de HPH", d:"Plazo pedido por el proveedor: 30-10-2026. Retiene el alcance ampliado.", u:"PA-14"},
 {t:"Aprobar la ingeniería de las salas", d:"Freezing point vencido el 25-09-26. Los equipos ya están aprobados; falta el shelter.", u:"PA-08"},
 {t:"Responder el reclamo de los ductos", d:"ABB pide adecuar la fecha de las posiciones 120/130/140 y evalúa impacto económico.", u:"PA-07"},
 {t:"Entregar los mapas del PMS", d:"Mapas Modbus y Profinet. La configuración de comunicaciones arranca el 02-11-26.", u:"PA-11"},
];

const pres=new pptxgen();
pres.layout="LAYOUT_WIDE";                 // 13.3 x 7.5
pres.author="AESA"; pres.title="Provisiones críticas CPF-2 — resumen gerencial";
pres.theme={ headFontFace:"Cambria", bodyFontFace:"Calibri" };

const s=pres.addSlide();
s.background={ color:"FFFFFF" };

// ---------- encabezado ----------
s.addText("Provisiones críticas CPF-2 · estado al 27-10-2026",
 { x:0.5, y:0.26, w:10.8, h:0.48, fontSize:27, bold:true, color:NAVY,
   fontFace:"Cambria", isTextBox:true, margin:0, objectName:"titulo" });
s.addText("Cronograma ejecutivo de provisiones críticas · ocho provisiones · seguimiento para PPSA",
 { x:0.5, y:0.80, w:10.8, h:0.28, fontSize:12.5, color:MUT, isTextBox:true, margin:0, objectName:"subtitulo" });
s.addShape(pres.ShapeType.roundRect,
 { x:11.45, y:0.32, w:1.35, h:0.52, rectRadius:0.1, fill:{ color:NAVY }, line:{ color:NAVY }, objectName:"badgeRev" });
s.addText("Rev 11", { x:11.45, y:0.32, w:1.35, h:0.52, fontSize:14, bold:true, color:"FFFFFF",
   align:"center", valign:"middle", isTextBox:true, margin:0, objectName:"badgeRevTxt" });

// ---------- KPIs ----------
const kw=2.8875, kx=[0.5, 3.6375, 6.775, 9.9125];
KPI.forEach((k,i)=>{
  s.addShape(pres.ShapeType.roundRect,
   { x:kx[i], y:1.26, w:kw, h:0.98, rectRadius:0.07, fill:{ color:CARD }, line:{ color:LINE, width:1 },
     objectName:"kpiBox"+i });
  s.addText(k.l.toUpperCase(), { x:kx[i]+0.18, y:1.34, w:kw-0.36, h:0.2, fontSize:8.5, bold:true,
     color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"kpiL"+i });
  s.addText(k.v, { x:kx[i]+0.18, y:1.54, w:kw-0.36, h:0.34, fontSize:21, bold:true, color:k.c,
     fontFace:"Cambria", isTextBox:true, margin:0, objectName:"kpiV"+i });
  s.addText(k.d, { x:kx[i]+0.18, y:1.92, w:kw-0.36, h:0.26, fontSize:9, color:MUT,
     isTextBox:true, margin:0, objectName:"kpiD"+i });
});

// ---------- panel izquierdo: provisiones ----------
const LX=0.5, LW=7.8;
s.addText("Entregas comprometidas por provisión", { x:LX, y:2.46, w:LW, h:0.26, fontSize:13, bold:true,
   color:NAVY, fontFace:"Cambria", isTextBox:true, margin:0, objectName:"hdrProv" });

const CN=LX+0.32, CE=LX+3.28, CB=LX+4.72, CP=LX+5.92, CS=LX+6.5;
s.addText("PROVISIÓN", { x:CN, y:2.78, w:2.9, h:0.18, fontSize:8, bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"thP" });
s.addText("ENTREGA",   { x:CE, y:2.78, w:1.4, h:0.18, fontSize:8, bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"thE" });
s.addText("AVANCE",    { x:CB, y:2.78, w:1.2, h:0.18, fontSize:8, bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"thA" });

const Y0=3.02, RH=0.47;
P.forEach((p,i)=>{
  const y=Y0+i*RH;
  if(i>0) s.addShape(pres.ShapeType.line, { x:LX, y:y-0.03, w:LW, h:0, line:{ color:LINE, width:0.75 }, objectName:"sep"+i });
  s.addShape(pres.ShapeType.ellipse, { x:LX, y:y+0.10, w:0.15, h:0.15, fill:{ color:p.c }, line:{ color:p.c }, objectName:"dot"+i });
  s.addText(p.n, { x:CN, y:y+0.02, w:2.95, h:0.22, fontSize:11.5, bold:true, color:INK, isTextBox:true, margin:0, objectName:"nm"+i });
  s.addText(p.nota, { x:CN, y:y+0.23, w:3.3, h:0.2, fontSize:9, color:p.nc, bold:p.nc===RED, isTextBox:true, margin:0, objectName:"nt"+i });
  s.addText(p.e, { x:CE, y:y+0.10, w:1.4, h:0.26, fontSize:13, bold:true, color:p.c, fontFace:"Cambria", isTextBox:true, margin:0, objectName:"en"+i });
  if(p.av===null){
    s.addText("sin dato global", { x:CB, y:y+0.13, w:1.5, h:0.2, fontSize:9, italic:true, color:MUT, isTextBox:true, margin:0, objectName:"av"+i });
  }else{
    s.addShape(pres.ShapeType.roundRect, { x:CB, y:y+0.15, w:1.1, h:0.13, rectRadius:0.06, fill:{ color:LINE }, line:{ color:LINE }, objectName:"bg"+i });
    if(p.av>0) s.addShape(pres.ShapeType.roundRect, { x:CB, y:y+0.15, w:Math.max(1.1*p.av/100,0.07), h:0.13, rectRadius:0.06, fill:{ color:p.c }, line:{ color:p.c }, objectName:"bf"+i });
    s.addText(p.av+"%", { x:CP, y:y+0.09, w:0.55, h:0.24, fontSize:11, bold:true, color:(p.av<=6?RED:INK), isTextBox:true, margin:0, objectName:"pc"+i });
  }
});

// ---------- panel derecho: foco gerencial ----------
const RX=8.62, RW=4.18;
s.addShape(pres.ShapeType.roundRect, { x:RX, y:2.46, w:RW, h:4.32, rectRadius:0.08,
   fill:{ color:CARD }, line:{ color:LINE, width:1 }, objectName:"panelDer" });
s.addText("Foco gerencial", { x:RX+0.22, y:2.60, w:RW-0.44, h:0.26, fontSize:13, bold:true, color:NAVY,
   fontFace:"Cambria", isTextBox:true, margin:0, objectName:"hdrFoco" });
s.addText([
  { text:"En 7 de las 8 provisiones ", options:{ bold:false } },
  { text:"el pendiente crítico es un insumo de AESA", options:{ bold:true } },
  { text:", no del proveedor. Seis sostienen su fecha de entrega.", options:{ bold:false } },
], { x:RX+0.22, y:2.88, w:RW-0.44, h:0.58, fontSize:11, color:INK, isTextBox:true, margin:0,
     lineSpacingMultiple:1.12, objectName:"claim" });

s.addText("DECISIONES Y ACCIONES INMEDIATAS", { x:RX+0.22, y:3.52, w:RW-0.44, h:0.2, fontSize:8.5,
   bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"hdrAcc" });

const AY=3.78, AH=0.58;
ACC.forEach((a,i)=>{
  const y=AY+i*AH;
  s.addShape(pres.ShapeType.ellipse, { x:RX+0.22, y:y+0.03, w:0.22, h:0.22, fill:{ color:NAVY }, line:{ color:NAVY }, objectName:"num"+i });
  s.addText(String(i+1), { x:RX+0.22, y:y+0.03, w:0.22, h:0.22, fontSize:9.5, bold:true, color:"FFFFFF",
     align:"center", valign:"middle", isTextBox:true, margin:0, objectName:"numT"+i });
  s.addText([
    { text:a.t, options:{ bold:true, color:INK } },
    { text:"   "+a.u, options:{ bold:true, color:RED, fontSize:8.5 } },
  ], { x:RX+0.52, y:y, w:RW-0.74, h:0.2, fontSize:10, isTextBox:true, margin:0, objectName:"accT"+i });
  s.addText(a.d, { x:RX+0.52, y:y+0.195, w:RW-0.74, h:0.37, fontSize:8.8, color:MUT,
     isTextBox:true, margin:0, lineSpacingMultiple:1.06, objectName:"accD"+i });
});

// ---------- pie ----------
s.addText("Fuentes: ABB Salas y Ductos 5155-00-0009-VE-CO-PG-001 Rev 0 «Para construcción» (14-10-26) e informe de avance Nº 5 · ABB PMS SEP-2026 Rev 0 e informe Nº 4 · "+
          "CAT Miron 5155-00-0022-VE-CO-PG-001 Rev A · Inauco 5155-00-0029-VE-CA-PG-001 Rev 0 «Para construcción» (23-10-26) · HPH Supplier Progress Report Nº 1 (15-10-26) · HARG PSCH_CPF2_Rev3.",
 { x:0.5, y:6.92, w:12.3, h:0.44, fontSize:7.8, color:MUT, isTextBox:true, margin:0,
   lineSpacingMultiple:1.1, objectName:"fuentes" });

s.addNotes("Resumen de una página del Cronograma Ejecutivo de Provisiones Críticas Rev11 (27-10-2026). "+
 "Camino crítico de sistemas: la OC de HIMA Argentina gobierna el cierre del 10-05-2027. "+
 "La última entrega del paquete son los ductos de barras el 08-07-2027, con el plazo aún sin confirmar por ABB. "+
 "Seis de las ocho provisiones sostienen su fecha; las dos excepciones (ductos y HARG) dependen de definiciones de AESA.");

pres.writeFile({ fileName: process.argv[2] }).then(f=>console.log("OK",f));
