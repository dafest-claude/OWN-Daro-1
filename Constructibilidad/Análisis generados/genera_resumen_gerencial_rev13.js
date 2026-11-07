const pptxgen = require("pptxgenjs");

const NAVY="1E2761", INK="1F2933", MUT="667788", LINE="E4E8EE", CARD="F5F7FA";
const RED="C62828", AMB="B8860B", GRN="2E7D32";

// Provisiones ordenadas por fecha de entrega (Cronograma Ejecutivo Rev13, 07-11-2026)
const P=[
 {c:"33691E", n:"CAT Miron — Transformadores", e:"11-DIC-2026", av:null,
  nota:"Inspección OK · 1.ª entrega el 13-11-2026", nc:GRN, dv:"Sin desvío", dc:GRN},
 {c:"6A1B9A", n:"HPH — Marshalling (Alemania)",        e:"15-DIC-2026", av:100,
  nota:"Fabricación completada · carta de crédito vencida", nc:AMB, dv:"Sin desvío", dc:GRN},
 {c:"EF6C00", n:"ABB — PMS",                            e:"17-FEB-2027", av:54,
  nota:"EL ARMADO NO ARRANCÓ EN OCTUBRE", nc:RED, dv:"En riesgo", dc:RED},
 {c:"AD1457", n:"ABB — Sala eléctrica SE#4 (MT)",       e:"25-MAR-2027", av:49,
  nota:"Fecha contractual 08-04-2027", nc:MUT, dv:"14 d antes", dc:GRN},
 {c:"C62828", n:"ABB — Sala eléctrica SE#3 (BT)",       e:"23-ABR-2027", av:34,
  nota:"Fecha contractual 14-05-2027", nc:MUT, dv:"21 d antes", dc:GRN},
 {c:"1565C0", n:"Inauco — PCS / SCADA",                 e:"04-MAY-2027", av:6,
  nota:"Sin corrimientos · integración PCS↔PMS 30-11", nc:MUT, dv:"Sin desvío", dc:GRN},
 {c:"00838F", n:"HARG — HIMA Argentina (SIS)",          e:"24-MAY-2027", av:0,
  nota:"OC POSTERGADA A LA SEMANA DEL 16-11", nc:RED, dv:"+63 días", dc:RED},
 {c:"795548", n:"ABB — Ductos de barras",               e:"08-JUL-2027", av:15,
  nota:"PLAZO SIN CONFIRMAR", nc:RED, dv:"+28 días", dc:RED},
];

const KPI=[
 {l:"Cierre del camino crítico", v:"24-MAY-2027", d:"+14 días · OC HARG → HPH → HARG → Inauco", c:RED},
 {l:"Última entrega del paquete", v:"08-JUL-2027", d:"Ductos de barras — sin confirmar", c:RED},
 {l:"Provisiones con desvío",     v:"3 de 8",      d:"HARG +63 d · Ductos +28 d · PMS en riesgo", c:RED},
 {l:"Puntos de atención abiertos",v:"19",          d:"13 en severidad alta · 4 nuevos", c:AMB},
];

const ACC=[
 {t:"Colocar la OC de HIMA Argentina", u:"PA-04",
  d:"Postergada por segunda vez a la semana del 16-11. Cada semana suma una semana al cierre."},
 {t:"Emitir la carta de crédito de HPH", u:"PA-14",
  d:"VENCIÓ el 30-10-2026. Retiene el material de la orden ampliada, ya reservado en el ERP."},
 {t:"Cerrar la ingeniería de las salas", u:"PA-08",
  d:"ABB pide no abrir comentarios nuevos. El shelter de SE#3 se fabrica desde el 23-11."},
 {t:"Resolver la agenda del 30-11", u:"PA-19",
  d:"FAT de TGMT y CCM en Sorocaba contra la integración PCS↔PMS en Buenos Aires."},
 {t:"Entregar los mapas del PMS", u:"PA-11",
  d:"Segundo mes pendientes. La especialidad Configuración sigue en 0%."},
 {t:"Reclamar el armado del PMS", u:"PA-18",
  d:"Sin actividades en octubre y ABB sostiene el 17-12: pedir cronograma recalculado."},
];

const pres=new pptxgen();
pres.layout="LAYOUT_WIDE";                 // 13.3 x 7.5
pres.author="AESA"; pres.title="Provisiones críticas CPF-2 — resumen gerencial (Rev 13)";
pres.theme={ headFontFace:"Cambria", bodyFontFace:"Calibri" };

const s=pres.addSlide();
s.background={ color:"FFFFFF" };

// ---------- encabezado ----------
s.addText("Provisiones críticas CPF-2 · estado al 07-11-2026",
 { x:0.5, y:0.28, w:10.9, h:0.46, fontSize:26, bold:true, color:NAVY,
   fontFace:"Cambria", isTextBox:true, margin:0, objectName:"titulo" });
s.addText("Cronograma ejecutivo de provisiones críticas · ocho provisiones · seguimiento para PPSA",
 { x:0.5, y:0.80, w:10.9, h:0.28, fontSize:12.5, color:MUT, isTextBox:true, margin:0, objectName:"subtitulo" });
s.addShape(pres.ShapeType.roundRect,
 { x:11.65, y:0.32, w:1.15, h:0.52, rectRadius:0.1, fill:{ color:NAVY }, line:{ color:NAVY }, objectName:"badgeRev" });
s.addText("Rev 13", { x:11.65, y:0.32, w:1.15, h:0.52, fontSize:13, bold:true, color:"FFFFFF",
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

const CN=LX+0.32, CE=LX+3.28, CB=LX+4.72, CP=LX+5.92, CD=LX+6.60;
s.addText("PROVISIÓN", { x:CN, y:2.78, w:2.9, h:0.18, fontSize:8, bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"thP" });
s.addText("ENTREGA",   { x:CE, y:2.78, w:1.4, h:0.18, fontSize:8, bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"thE" });
s.addText("AVANCE",    { x:CB, y:2.78, w:1.2, h:0.18, fontSize:8, bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"thA" });
s.addText("DESVÍO",    { x:CD, y:2.78, w:1.2, h:0.18, fontSize:8, bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"thD" });

const Y0=2.98, RH=0.45;
P.forEach((p,i)=>{
  const y=Y0+i*RH;
  if(i>0) s.addShape(pres.ShapeType.line, { x:LX, y:y-0.03, w:LW, h:0, line:{ color:LINE, width:0.75 }, objectName:"sep"+i });
  s.addShape(pres.ShapeType.ellipse, { x:LX, y:y+0.10, w:0.15, h:0.15, fill:{ color:p.c }, line:{ color:p.c }, objectName:"dot"+i });
  s.addText(p.n, { x:CN, y:y+0.02, w:2.95, h:0.22, fontSize:11.5, bold:true, color:INK, isTextBox:true, margin:0, objectName:"nm"+i });
  s.addText(p.nota, { x:CN, y:y+0.23, w:2.88, h:0.2, fontSize:9, color:p.nc, bold:p.nc===RED, isTextBox:true, margin:0, objectName:"nt"+i });
  s.addText(p.e, { x:CE, y:y+0.10, w:1.4, h:0.26, fontSize:13, bold:true, color:p.c, fontFace:"Cambria", isTextBox:true, margin:0, objectName:"en"+i });
  if(p.av===null){
    s.addText("sin dato global", { x:CB, y:y+0.13, w:1.5, h:0.2, fontSize:9, italic:true, color:MUT, isTextBox:true, margin:0, objectName:"av"+i });
  }else{
    s.addShape(pres.ShapeType.roundRect, { x:CB, y:y+0.15, w:1.1, h:0.13, rectRadius:0.06, fill:{ color:LINE }, line:{ color:LINE }, objectName:"bg"+i });
    if(p.av>0) s.addShape(pres.ShapeType.roundRect, { x:CB, y:y+0.15, w:Math.max(1.1*p.av/100,0.07), h:0.13, rectRadius:0.06, fill:{ color:p.c }, line:{ color:p.c }, objectName:"bf"+i });
    s.addText(p.av+"%", { x:CP, y:y+0.09, w:0.55, h:0.24, fontSize:11, bold:true, color:(p.av<=6?RED:INK), isTextBox:true, margin:0, objectName:"pc"+i });
  }
  s.addText(p.dv, { x:CD, y:y+0.09, w:1.18, h:0.24, fontSize:10, bold:(p.dc===RED), color:p.dc,
     isTextBox:true, margin:0, objectName:"dv"+i });
});
s.addText("Desvío medido contra la fecha CONTRACTUAL en las salas eléctricas y contra el plan original de cada proveedor en el resto. El PMS figura «en riesgo»: ABB sostiene la entrega pero el armado de tableros no arrancó.",
 { x:LX, y:6.62, w:LW, h:0.2, fontSize:8, italic:true, color:MUT, isTextBox:true, margin:0, objectName:"notaDesvio" });

// ---------- panel derecho: foco gerencial ----------
const RX=8.62, RW=4.18;
s.addShape(pres.ShapeType.roundRect, { x:RX, y:2.46, w:RW, h:4.38, rectRadius:0.08,
   fill:{ color:CARD }, line:{ color:LINE, width:1 }, objectName:"panelDer" });
s.addText("Foco gerencial", { x:RX+0.22, y:2.60, w:RW-0.44, h:0.26, fontSize:13, bold:true, color:NAVY,
   fontFace:"Cambria", isTextBox:true, margin:0, objectName:"hdrFoco" });
s.addText([
  { text:"La OC de ", options:{ bold:false } },
  { text:"HIMA Argentina", options:{ bold:true } },
  { text:" se posterga por segunda vez y mueve el cierre del paquete del 10-05 al ", options:{ bold:false } },
  { text:"24-05-2027", options:{ bold:true } },
  { text:". ABB aprobó la ingeniería de todos los equipos de las salas, pero el ", options:{ bold:false } },
  { text:"armado del PMS no arrancó", options:{ bold:true } },
  { text:" en octubre.", options:{ bold:false } },
], { x:RX+0.22, y:2.86, w:RW-0.44, h:0.84, fontSize:10, color:INK, isTextBox:true, margin:0,
     valign:"top", lineSpacingMultiple:1.08, objectName:"claim" });

s.addText("DECISIONES Y ACCIONES INMEDIATAS", { x:RX+0.22, y:3.76, w:RW-0.44, h:0.2, fontSize:8.5,
   bold:true, color:MUT, charSpacing:0.6, isTextBox:true, margin:0, objectName:"hdrAcc" });

const AY=3.98, AH=0.455;
ACC.forEach((a,i)=>{
  const y=AY+i*AH;
  s.addShape(pres.ShapeType.ellipse, { x:RX+0.22, y:y+0.03, w:0.22, h:0.22, fill:{ color:NAVY }, line:{ color:NAVY }, objectName:"num"+i });
  s.addText(String(i+1), { x:RX+0.22, y:y+0.03, w:0.22, h:0.22, fontSize:9.5, bold:true, color:"FFFFFF",
     align:"center", valign:"middle", isTextBox:true, margin:0, objectName:"numT"+i });
  s.addText([
    { text:a.t, options:{ bold:true, color:INK } },
    { text:"   "+a.u, options:{ bold:true, color:RED, fontSize:8.5 } },
  ], { x:RX+0.52, y:y, w:RW-0.74, h:0.2, fontSize:10, isTextBox:true, margin:0, objectName:"accT"+i });
  s.addText(a.d, { x:RX+0.52, y:y+0.185, w:RW-0.74, h:0.32, fontSize:8.4, color:MUT,
     isTextBox:true, margin:0, lineSpacingMultiple:1.06, objectName:"accD"+i });
});

// ---------- pie ----------
s.addText("Fuentes: ABB PMS informe de avance Nº 5 (E-2611027, 30-10-26) · ABB Salas Eléctricas informe de avance Nº 6 (R-2631036_A, 04-11-26) · ABB Salas y Ductos 5155-00-0009-VE-CO-PG-001 Rev 0 «Para construcción» (14-10-26) · "+
          "CAT Miron acta de inspección IN-FC-CAT-19082-261021 (21-10-26) · Inauco 5155-00-0029-VE-CA-PG-001 Rev 0 (23-10-26) · HPH Supplier Progress Report Nº 1 (15-10-26) · HARG PSCH_CPF2_Rev3 con corrimiento +63 d.",
 { x:0.5, y:6.92, w:12.3, h:0.44, fontSize:7.8, color:MUT, isTextBox:true, margin:0,
   lineSpacingMultiple:1.1, objectName:"fuentes" });

s.addNotes("Resumen de una página del Cronograma Ejecutivo de Provisiones Críticas Rev13 (07-11-2026). "+
 "Lo que cambió: (1) AESA informa que la orden de compra de HIMA Argentina no sale esta semana sino la siguiente; se adopta el 16-11-2026 y el corrimiento pasa de 49 a 63 días corridos, con lo que el cierre del camino crítico se mueve del 10-05-2027 al 24-05-2027. Es el primer corrimiento del proyecto que cambia la fecha final del paquete. "+
 "(2) Informe ABB PMS Nº5 del 30-10-2026: avance 54% contra 54% previsto, pero el armado de tableros NO arrancó en octubre (sin actividades en Fabricación) y pasa a noviembre en MEHCCO S.A.; ABB sostiene el fin de armado del 17-12 sin sustento aritmético. El acopio quedó verificado al 100% el 22-10 con inspector de AESA y certificado por USD 220.055,72. "+
 "(3) Informe ABB Salas Nº6 del 04-11-2026: hito del período, aprobada la ingeniería de detalle de TODOS los equipos para fabricación; la del shelter sigue con aprobación parcial y ABB pide no abrir comentarios nuevos. Fechas objetivo ratificadas: SE#3 23-04-27 y SE#4 25-03-27, 21 y 14 días mejor que la contractual. Acopio de tableros auxiliares al 100%. "+
 "(4) Nueva previsión de FAT: UPS 18-20/11 en Santa Fe, TGMT 30/11-04/12 y CCM 30/11-12/12 en Sorocaba (invitaciones enviadas) y auxiliares 14-16/12 en AMBA, con la prueba de integración PCS↔PMS la misma semana en Buenos Aires: tres eventos con witness el 30 de noviembre en dos países. "+
 "Puntos de atención: 19 abiertos, 13 de severidad alta, cuatro nuevos (PA-18 armado del PMS, PA-19 agenda de witness, PA-20 módulo adicional fuera de alcance, PA-21 procedimientos de FAT y SAT). "+
 "En 7 de las 8 provisiones el pendiente crítico sigue siendo un insumo de AESA.");

pres.writeFile({ fileName: process.argv[2] }).then(f=>console.log("OK",f));
