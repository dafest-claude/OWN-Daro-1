#!/usr/bin/env python3
"""
Generador de APP HTML cronológica – Suministros La Calera II CPF2 (IN + EL)
===========================================================================
Lee TODOS los planes de suministros de ../info_suministros/ y produce un único
archivo HTML autocontenido (sin dependencias externas) que consolida la
información de Tracker + Dashboard + Infográfico y permite:
  · elegir un corte/fecha y ver KPIs, embudo de compra y críticos
  · comparar dos cortes (deltas de pipeline y cambios en críticos)
  · ver la tendencia de las métricas a lo largo de todos los cortes

Uso: cada semana se sube el nuevo plan a info_suministros/ y se re-ejecuta este
script; el corte nuevo se incorpora automáticamente.

Nombre de salida: App_Suministros_IN_EL_LaCalera_II_<REV>_<fecha>.html
"""
import os, re, glob, json, html, warnings
from datetime import date, datetime
import openpyxl

warnings.filterwarnings('ignore')

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PLANS_DIR  = os.path.join(SCRIPT_DIR, '..', 'info_suministros')

# ── Revisión de la APP (nombre + fecha) ──────────────────────────────────────
APP_REV   = 'Rev2'
GEN_DATE  = date(2026, 7, 24)          # fecha de generación de esta revisión
OUT_HTML  = os.path.join(SCRIPT_DIR,
            f'App_Suministros_IN_EL_LaCalera_II_{APP_REV}_{GEN_DATE.strftime("%d%m%y")}.html')

# 'first' (sin fecha en el nombre) se ordena como base inicial
FIRST_DATE = date(2026, 5, 14)


def to_d(v):
    if isinstance(v, datetime): return v.date()
    if isinstance(v, date):     return v
    return None

def dmy(d):
    return d.strftime('%d/%m/%y') if isinstance(d, date) else ''

def iso(d):
    return d.strftime('%Y-%m-%d') if isinstance(d, date) else None

def parse_tag_date(tag):
    """ddmmyy -> date"""
    try:
        return date(2000 + int(tag[4:6]), int(tag[2:4]), int(tag[0:2]))
    except Exception:
        return None


def read_pipeline(wb):
    ws = wb['Cuadro resumen']
    nombre_map = {'electricidad': 'EL', 'instrumentos': 'IN'}
    out = {}
    for row in ws.iter_rows(min_row=8, max_row=15, values_only=True):
        nm = str(row[1]).strip().lower() if row[1] else ''
        code = nombre_map.get(nm)
        if not code:
            continue
        out[code] = {
            'cant_ri':  row[2] or 0, 'emitidas': row[3] or 0,
            'solped':   row[4] or 0, 'ofertas':  row[5] or 0,
            'at':       row[6] or 0, 'oc':       row[7] or 0,
        }
    return out


def read_criticos(wb, cut_date):
    ws = wb['Suministros críticos']
    items = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or not row[0]:
            continue
        esp = str(row[2]) if row[2] else ''
        if esp not in ('IN', 'EL'):
            continue
        ri      = to_d(row[4]);  solped = to_d(row[5]);  recof = to_d(row[6])
        at_c    = to_d(row[7]);  necoc  = to_d(row[8]);  ocreal = to_d(row[9])
        oc_n    = str(row[11]).strip() if row[11] else ''
        prov    = str(row[12]).strip() if row[12] else ''
        kom     = to_d(row[13])
        status  = re.sub(r'\s+', ' ', str(row[19]).strip()) if row[19] else ''
        oc_eff  = ocreal or kom
        has_oc  = bool(oc_n) and oc_n not in ('-', '') and oc_eff is not None

        # estado + días
        if has_oc:
            estado = 'OC'
            dias = (oc_eff - ri).days if (ri and oc_eff) else None
        else:
            if ri and cut_date and ri <= cut_date:
                dias = (cut_date - ri).days
            else:
                dias = None
            if at_c:       estado = 'AT CERR'     # AT cerrado, aguarda OC
            elif recof:    estado = 'EN AT'        # ofertas recibidas, en análisis técnico
            elif solped:   estado = 'EN SOLPED'
            elif ri:       estado = 'RI EMITIDA'
            else:          estado = 'PENDIENTE'

        desc = re.sub(r'^RI\s*-\s*', '', str(row[3]).strip()) if row[3] else ''
        items.append({
            'esp': esp, 'crit': str(row[1]).strip() if row[1] else '', 'desc': desc,
            'ri': dmy(ri), 'solped': dmy(solped), 'recof': dmy(recof),
            'at': dmy(at_c), 'necoc': dmy(necoc),
            'oc_eff': dmy(oc_eff), 'oc_n': oc_n if has_oc else '',
            'prov': prov, 'dias': dias, 'estado': estado,
            'status': status[:600],
        })
    # orden: OC primero por estado luego por esp
    return items


def build_cut(fpath):
    b = os.path.basename(fpath)
    m = re.search(r'\((\d{6})\)', b)
    tag = m.group(1) if m else 'first'
    cut_date = parse_tag_date(tag) if tag != 'first' else FIRST_DATE
    label = dmy(cut_date) if tag != 'first' else 'Inicial'
    wb = openpyxl.load_workbook(fpath, data_only=True)
    pipe = read_pipeline(wb)
    crit = read_criticos(wb, cut_date)
    wb.close()

    el = pipe.get('EL', {}); inn = pipe.get('IN', {})
    ri_total   = (el.get('cant_ri', 0) + inn.get('cant_ri', 0))
    oc         = (el.get('oc', 0) + inn.get('oc', 0))
    en_gestion = (el.get('emitidas', 0) + inn.get('emitidas', 0))
    kpis = {
        'ri_total': ri_total, 'oc': oc, 'en_gestion': en_gestion,
        'sin_emitir': ri_total - en_gestion,
        'solped': el.get('solped', 0) + inn.get('solped', 0),
        'ofertas': el.get('ofertas', 0) + inn.get('ofertas', 0),
        'at': el.get('at', 0) + inn.get('at', 0),
    }
    return {
        'tag': tag, 'date': iso(cut_date), 'label': label, 'file': b,
        'is_first': tag == 'first',
        'pipeline': {'EL': el, 'IN': inn}, 'kpis': kpis, 'criticos': crit,
    }


def load_all():
    files = glob.glob(os.path.join(PLANS_DIR, '*.xlsx'))
    cuts = [build_cut(f) for f in files]
    # ordenar cronológicamente (first primero)
    cuts.sort(key=lambda c: (0, '') if c['is_first'] else (1, c['date']))
    return cuts


# ═══════════════════════════════════════════════════════════════════════════
# PLANTILLA HTML
# ═══════════════════════════════════════════════════════════════════════════
def render_html(cuts):
    data_json = json.dumps(cuts, ensure_ascii=False)
    meta = {
        'rev': APP_REV, 'gen': GEN_DATE.strftime('%d/%m/%Y'),
        'n': len(cuts),
        'primero': cuts[0]['label'], 'ultimo': cuts[-1]['label'],
    }
    meta_json = json.dumps(meta, ensure_ascii=False)

    return TEMPLATE.replace('/*__DATA__*/', data_json).replace('/*__META__*/', meta_json)


TEMPLATE = r"""<!doctype html>
<html lang="es">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Suministros La Calera II CPF2 — Tablero cronológico IN + EL</title>
<style>
  :root{
    --azul:#1F3864; --azulm:#2F5496; --azulcl:#D6E4F0;
    --el:#2E6B2E; --elcl:#D7E9D7; --in:#9C4500; --incl:#F5E0D0;
    --gris:#404040; --griscl:#F2F2F2; --borde:#C9D2DE;
    --ok:#2E7D32; --okcl:#C8E6C9; --warn:#B8860B; --warncl:#FFECB3;
    --crit:#C00000; --critcl:#FFCDD2; --morado:#7030A0;
    --bg:#eef1f5; --card:#ffffff; --tx:#20272e; --mut:#66707b;
  }
  @media (prefers-color-scheme: dark){
    :root{ --bg:#12161c; --card:#1b2129; --tx:#e6ebf1; --mut:#9aa6b2;
           --borde:#2c3542; --griscl:#232b34; --azulcl:#1e2a3d; }
  }
  *{box-sizing:border-box}
  body{margin:0;font-family:-apple-system,Segoe UI,Roboto,Helvetica,Arial,sans-serif;
       background:var(--bg);color:var(--tx);font-size:14px;line-height:1.4}
  header{background:linear-gradient(100deg,var(--azul),var(--azulm));color:#fff;
         padding:14px 22px;box-shadow:0 2px 10px rgba(0,0,0,.18)}
  header h1{margin:0;font-size:19px;font-weight:800;letter-spacing:.2px}
  header .sub{font-size:12px;color:#d6e4f0;margin-top:3px}
  header .rev{display:inline-block;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.3);
        border-radius:6px;padding:2px 9px;font-weight:700;font-size:12px;margin-right:8px}
  .wrap{max-width:1180px;margin:0 auto;padding:16px}
  .tabs{display:flex;gap:6px;margin:14px 0 6px;flex-wrap:wrap}
  .tab{padding:8px 16px;border-radius:8px 8px 0 0;background:transparent;border:1px solid transparent;
       cursor:pointer;font-weight:700;color:var(--mut);font-size:13px}
  .tab.active{background:var(--card);color:var(--azul);border-color:var(--borde);border-bottom-color:var(--card)}
  @media (prefers-color-scheme: dark){ .tab.active{color:var(--azulcl)} }
  .panel{background:var(--card);border:1px solid var(--borde);border-radius:0 10px 10px 10px;
         padding:18px;box-shadow:0 1px 4px rgba(0,0,0,.06)}
  .row{display:flex;gap:14px;flex-wrap:wrap;align-items:center}
  label.fld{font-size:12px;color:var(--mut);font-weight:700;margin-right:6px}
  select{font-size:14px;padding:7px 10px;border-radius:8px;border:1px solid var(--borde);
          background:var(--card);color:var(--tx);font-weight:600}
  .kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:16px 0}
  .kpi{border-radius:12px;padding:14px 12px;color:#fff;text-align:center}
  .kpi .n{font-size:34px;font-weight:800;line-height:1}
  .kpi .l{font-size:11px;margin-top:6px;opacity:.95;font-weight:600}
  .kpi .d{font-size:12px;font-weight:800;margin-top:4px}
  .k1{background:var(--azul)} .k2{background:var(--el)} .k3{background:var(--in)} .k4{background:var(--morado)}
  h3.sec{font-size:14px;color:var(--azul);margin:20px 0 8px;border-bottom:2px solid var(--azulcl);padding-bottom:4px}
  @media (prefers-color-scheme: dark){ h3.sec{color:var(--azulcl)} }
  .funnel{display:grid;grid-template-columns:120px 1fr;gap:6px 10px;align-items:center;margin:8px 0}
  .fstage{font-size:12px;font-weight:700;color:var(--mut);text-align:right}
  .fbars{display:flex;flex-direction:column;gap:3px}
  .fbar{height:16px;border-radius:4px;display:flex;align-items:center;min-width:2px;transition:width .3s}
  .fbar span{font-size:11px;font-weight:800;color:#fff;padding:0 6px;white-space:nowrap}
  .legend{display:flex;gap:16px;font-size:12px;margin:4px 0 0}
  .legend b{display:inline-block;width:11px;height:11px;border-radius:3px;margin-right:5px;vertical-align:-1px}
  table{border-collapse:collapse;width:100%;font-size:12.5px;margin-top:6px}
  th{background:var(--gris);color:#fff;padding:7px 8px;text-align:left;font-size:11px;position:sticky;top:0}
  td{padding:6px 8px;border-bottom:1px solid var(--borde);vertical-align:middle}
  tr.el td:first-child{border-left:4px solid var(--el)}
  tr.in td:first-child{border-left:4px solid var(--in)}
  .badge{display:inline-block;padding:2px 8px;border-radius:20px;font-size:10.5px;font-weight:800;color:#fff;white-space:nowrap}
  .b-oc{background:var(--ok)} .b-at{background:var(--warn)} .b-recof{background:var(--azulm)}
  .b-solped{background:#5B7FB8} .b-ri{background:#8894a3} .b-pend{background:#aeb6bf}
  .dias{font-weight:800}
  .d-ok{color:var(--ok)} .d-warn{color:var(--warn)} .d-crit{color:var(--crit)}
  .tblwrap{overflow-x:auto;border:1px solid var(--borde);border-radius:8px}
  .find{background:var(--griscl);border-left:4px solid var(--azulm);border-radius:6px;padding:9px 12px;margin:8px 0;font-size:12.5px}
  .find.warn{border-left-color:var(--warn)} .find.crit{border-left-color:var(--crit)} .find.ok{border-left-color:var(--ok)}
  .cmp{display:grid;grid-template-columns:1fr 1fr;gap:14px}
  .delta{font-weight:800;font-size:12px}
  .up{color:var(--ok)} .down{color:var(--azulm)} .eq{color:var(--mut)}
  .mini{font-size:11px;color:var(--mut)}
  .ctrls{display:flex;gap:10px;flex-wrap:wrap;align-items:center;margin-bottom:8px}
  .pill{padding:5px 11px;border-radius:20px;border:1px solid var(--borde);background:transparent;
        cursor:pointer;font-weight:700;font-size:12px;color:var(--mut)}
  .pill.active{background:var(--azul);color:#fff;border-color:var(--azul)}
  svg{max-width:100%}
  .foot{font-size:11px;color:var(--mut);margin-top:14px;text-align:center}
  .status-note{font-size:11px;color:var(--mut);font-style:italic}
  .diffbox{border:1px solid var(--borde);border-radius:8px;padding:8px 10px;margin:6px 0;font-size:12px}
  .diffbox b{color:var(--azul)}
  @media (prefers-color-scheme: dark){ .diffbox b{color:var(--azulcl)} }
  mark{background:var(--warncl);color:inherit;padding:0 3px;border-radius:3px;font-weight:700}
  .cmt{border:1px solid var(--borde);border-radius:9px;padding:10px 12px;margin:8px 0;background:var(--card)}
  .cmt.changed{border-left:4px solid var(--azulm);background:var(--azulcl)}
  .cmt .hd{display:flex;gap:8px;align-items:center;flex-wrap:wrap;margin-bottom:5px}
  .cmt .hd b{font-size:13px}
  .cmt .txt{font-size:12.5px;line-height:1.5}
  .cmt .prev{font-size:11px;color:var(--mut);margin-top:6px;border-top:1px dashed var(--borde);padding-top:5px}
  .mov{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px;margin:6px 0}
  .movcard{border:1px solid var(--borde);border-radius:9px;padding:10px 12px;background:var(--griscl)}
  .movcard .big{font-size:22px;font-weight:800;color:var(--azul)}
  @media (prefers-color-scheme: dark){ .movcard .big{color:var(--azulcl)} }
  .movcard .lbl{font-size:11px;color:var(--mut);font-weight:700}
  .flag{display:inline-block;background:var(--critcl);color:var(--crit);border-radius:6px;padding:2px 8px;font-size:11px;font-weight:800;margin:3px 4px 3px 0}
  @media(max-width:720px){ .kpis{grid-template-columns:repeat(2,1fr)} .cmp{grid-template-columns:1fr} }
</style>
</head>
<body>
<header>
  <h1>Suministros La Calera II CPF2 — Tablero cronológico · Instrumentación (IN) + Electricidad (EL)</h1>
  <div class="sub"><span class="rev" id="revlbl"></span><span id="metalbl"></span></div>
</header>
<div class="wrap">
  <div class="tabs">
    <div class="tab active" data-tab="vista">Vista por corte</div>
    <div class="tab" data-tab="cambios">Cambios &amp; comentarios</div>
    <div class="tab" data-tab="comparar">Comparar cortes</div>
    <div class="tab" data-tab="tendencia">Tendencia</div>
    <div class="tab" data-tab="criticos">Críticos (historial)</div>
  </div>

  <!-- ── VISTA POR CORTE ── -->
  <div class="panel" id="p-vista">
    <div class="row">
      <div><label class="fld">Fecha / corte del plan</label>
        <select id="selVista"></select></div>
      <div class="mini" id="vistaFile"></div>
    </div>
    <div class="kpis" id="vistaKpis"></div>
    <h3 class="sec">Embudo de gestión de compra por especialidad</h3>
    <div class="legend">
      <span><b style="background:var(--el)"></b>ELECTRICIDAD</span>
      <span><b style="background:var(--in)"></b>INSTRUMENTACIÓN</span>
    </div>
    <div id="vistaFunnel"></div>
    <h3 class="sec">Ítems críticos — cadena RI → OC</h3>
    <div class="tblwrap"><table id="vistaCrit"></table></div>
    <div class="status-note">RI→OC: días desde la RI a la OC (adjudicados) o días transcurridos al corte (en gestión). Pasá el mouse por el estado para ver el comentario del plan.</div>
    <h3 class="sec">Lectura del corte</h3>
    <div id="vistaFind"></div>
  </div>

  <!-- ── CAMBIOS & COMENTARIOS ── -->
  <div class="panel" id="p-cambios" style="display:none">
    <div class="row">
      <div><label class="fld">Corte del plan</label><select id="selCambios"></select></div>
      <div class="mini" id="cambiosSub"></div>
    </div>
    <h3 class="sec">Movimientos del período (vs corte anterior)</h3>
    <div id="cambiosMov"></div>
    <h3 class="sec">Comentarios relevantes — ítems críticos</h3>
    <div id="cambiosComents"></div>
    <div class="status-note">Se resaltan montos, fechas límite y palabras clave (aprobación, vencida, urgente, riesgo, NecOC, OCA…). El sello <span class="badge b-recof" style="padding:1px 6px">ACTUALIZADO</span> marca los comentarios que cambiaron respecto del corte anterior.</div>
  </div>

  <!-- ── COMPARAR ── -->
  <div class="panel" id="p-comparar" style="display:none">
    <div class="row">
      <div><label class="fld">Corte A (base)</label><select id="selA"></select></div>
      <div><label class="fld">Corte B (comparado)</label><select id="selB"></select></div>
    </div>
    <div class="kpis" id="cmpKpis"></div>
    <h3 class="sec">Pipeline por etapa — variación A → B</h3>
    <div class="tblwrap"><table id="cmpPipe"></table></div>
    <h3 class="sec">Cambios en ítems críticos (estado / OC / fechas)</h3>
    <div id="cmpCrit"></div>
  </div>

  <!-- ── TENDENCIA ── -->
  <div class="panel" id="p-tendencia" style="display:none">
    <div class="ctrls">
      <span class="mini">Métrica:</span>
      <button class="pill active" data-metric="oc">Con OC</button>
      <button class="pill" data-metric="at">En AT</button>
      <button class="pill" data-metric="ofertas">En petición</button>
      <button class="pill" data-metric="solped">En SOLPED</button>
      <button class="pill" data-metric="emitidas">Emitidas</button>
    </div>
    <div class="ctrls">
      <span class="mini">Series:</span>
      <button class="pill active" data-serie="TOTAL">IN+EL</button>
      <button class="pill" data-serie="IN">Instrumentación</button>
      <button class="pill" data-serie="EL">Electricidad</button>
    </div>
    <div id="trendChart"></div>
    <div class="tblwrap" style="margin-top:12px"><table id="trendTbl"></table></div>
  </div>

  <!-- ── CRÍTICOS HISTORIAL ── -->
  <div class="panel" id="p-criticos" style="display:none">
    <div class="row">
      <div><label class="fld">Ítem crítico</label><select id="selCrit"></select></div>
    </div>
    <div class="tblwrap" style="margin-top:10px"><table id="critHist"></table></div>
    <div class="status-note">Evolución del ítem a lo largo de todos los cortes cargados.</div>
  </div>

  <div class="foot" id="foot"></div>
</div>

<script>
const DATA = /*__DATA__*/;
const META = /*__META__*/;

// ── helpers ──────────────────────────────────────────────────────────────
const $ = s => document.querySelector(s);
const el = (t,c,h)=>{const e=document.createElement(t); if(c)e.className=c; if(h!=null)e.innerHTML=h; return e;};
const ESTADO = {
  'OC':{c:'b-oc',t:'OC COLOCADA'}, 'AT CERR':{c:'b-at',t:'AT CERRADO'},
  'EN AT':{c:'b-recof',t:'EN AT'}, 'REC.OF':{c:'b-recof',t:'REC. OFERTAS'},
  'EN SOLPED':{c:'b-solped',t:'EN SOLPED'},
  'RI EMITIDA':{c:'b-ri',t:'RI EMITIDA'}, 'PENDIENTE':{c:'b-pend',t:'PENDIENTE'}
};
const diasClass = (d)=> d==null?'':(d<=90?'d-ok':(d<=140?'d-warn':'d-crit'));
const STAGES = [['solped','En SOLPED'],['ofertas','Petición Ofertas'],['at','Análisis Técnico'],['oc','Con OC']];

// índices "reales" (con fecha) para selects por defecto
const lastIdx = DATA.length-1;

// header
$('#revlbl').textContent = 'App '+META.rev+' · Generado '+META.gen;
$('#metalbl').textContent = META.n+' cortes de plan ('+META.primero+' → '+META.ultimo+')';
$('#foot').textContent = 'Fuente: Planes de Suministros La Calera II · hoja Cuadro resumen + Suministros críticos · '
   + META.n + ' cortes · App '+META.rev+' generada '+META.gen;

// ── tabs ─────────────────────────────────────────────────────────────────
document.querySelectorAll('.tab').forEach(t=>t.onclick=()=>{
  document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));
  t.classList.add('active');
  ['vista','cambios','comparar','tendencia','criticos'].forEach(n=>{
    $('#p-'+n).style.display = (n===t.dataset.tab)?'block':'none';
  });
});

// ── selects ──────────────────────────────────────────────────────────────
function fillSelect(sel, selectedIdx){
  DATA.forEach((c,i)=>{
    const o=el('option',null,(c.is_first?'Inicial (base)':c.label)+'  ·  '+c.tag);
    o.value=i; sel.appendChild(o);
  });
  sel.value = selectedIdx;
}
fillSelect($('#selVista'), lastIdx);
fillSelect($('#selA'), Math.max(0,lastIdx-3));
fillSelect($('#selB'), lastIdx);

// ── VISTA ────────────────────────────────────────────────────────────────
function kpiCard(n,l,cls,delta){
  const d = delta==null?'':`<div class="d">${delta>0?'▲ +'+delta:(delta<0?'▼ '+delta:'=')}</div>`;
  return `<div class="kpi ${cls}"><div class="n">${n}</div><div class="l">${l}</div>${d}</div>`;
}
function renderVista(){
  const i=+$('#selVista').value, c=DATA[i], prev = i>0?DATA[i-1]:null;
  $('#vistaFile').textContent = 'Archivo: '+c.file;
  const k=c.kpis, pk=prev?prev.kpis:null;
  const dd=(a,b)=> b==null?null:(a-b);
  $('#vistaKpis').innerHTML =
     kpiCard(k.ri_total,'RIs TOTALES IN+EL','k1',dd(k.ri_total,pk&&pk.ri_total))
   + kpiCard(k.oc,'OCs COLOCADAS','k2',dd(k.oc,pk&&pk.oc))
   + kpiCard(k.en_gestion,'EN GESTIÓN DE COMPRA','k3',dd(k.en_gestion,pk&&pk.en_gestion))
   + kpiCard(k.sin_emitir,'RIs SIN EMITIR','k4',dd(k.sin_emitir,pk&&pk.sin_emitir));

  // funnel
  let maxv=1;
  STAGES.forEach(([kk])=>['EL','IN'].forEach(e=>maxv=Math.max(maxv,(c.pipeline[e]||{})[kk]||0)));
  let f='<div class="funnel">';
  STAGES.forEach(([kk,lbl])=>{
    const ve=(c.pipeline.EL||{})[kk]||0, vi=(c.pipeline.IN||{})[kk]||0;
    const w=v=>Math.max(2,Math.round(v/maxv*100));
    f+=`<div class="fstage">${lbl}</div><div class="fbars">
      <div class="fbar" style="width:${w(ve)}%;background:var(--el)"><span>${ve}</span></div>
      <div class="fbar" style="width:${w(vi)}%;background:var(--in)"><span>${vi}</span></div>
    </div>`;
  });
  f+='</div>';
  $('#vistaFunnel').innerHTML=f;

  // criticos table
  renderCritTable($('#vistaCrit'), c.criticos);
  // findings
  $('#vistaFind').innerHTML = findingsFor(i);
}
function renderCritTable(tbl, items){
  const order={'OC':0,'AT CERR':1,'EN AT':2,'REC.OF':3,'EN SOLPED':4,'RI EMITIDA':5,'PENDIENTE':6};
  const its=[...items].sort((a,b)=>(order[a.estado]-order[b.estado])||a.esp.localeCompare(b.esp));
  let h='<tr><th>Esp</th><th>Suministro</th><th>RI</th><th>SOLPED</th><th>Rec.Of</th><th>AT cierre</th>'
      +'<th>OC efect.</th><th>N° OC</th><th>RI→OC</th><th>Estado</th></tr>';
  its.forEach(it=>{
    const es=ESTADO[it.estado]||ESTADO['PENDIENTE'];
    const d=it.dias==null?'—':(it.dias+(it.estado==='OC'?' d':' d*'));
    h+=`<tr class="${it.esp.toLowerCase()}" title="${escapeHtml(it.status)}">
      <td><b>${it.esp}</b></td><td>${escapeHtml(it.desc)}</td>
      <td>${it.ri||'—'}</td><td>${it.solped||'—'}</td><td>${it.recof||'—'}</td>
      <td>${it.at||'—'}</td><td>${it.oc_eff||'—'}</td><td>${it.oc_n||'—'}</td>
      <td class="dias ${diasClass(it.dias)}">${d}</td>
      <td><span class="badge ${es.c}">${es.t}</span></td></tr>`;
  });
  tbl.innerHTML=h;
}
function escapeHtml(s){return (s||'').replace(/[&<>"]/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[m]));}

// ── hallazgos automáticos ────────────────────────────────────────────────
function findingsFor(i){
  const c=DATA[i], prev=i>0?DATA[i-1]:null;
  let out=[];
  // OCs nuevas
  if(prev){
    const doc=c.kpis.oc-prev.kpis.oc;
    if(doc>0) out.push(['ok',`+${doc} OC(s) colocada(s) respecto del corte anterior (${prev.label}). Total OCs IN+EL: ${c.kpis.oc}.`]);
    const dsol=c.kpis.solped-prev.kpis.solped;
    if(dsol<0) out.push(['',`SOLPED en proceso: ${prev.kpis.solped} → ${c.kpis.solped} (avanzaron ${Math.abs(dsol)}).`]);
    const dat=c.kpis.at-prev.kpis.at;
    if(dat>0) out.push(['',`+${dat} ítem(s) ingresaron a Análisis Técnico (AT ${prev.kpis.at} → ${c.kpis.at}).`]);
  }
  // criticos: OC recien emitidas, y estancados
  c.criticos.forEach(it=>{
    if(it.estado==='OC' && it.dias!=null)
      { if(prev){ const p=prev.criticos.find(x=>x.desc===it.desc); if(p && p.estado!=='OC') out.push(['ok',`RESUELTO: ${it.desc} — OC ${it.oc_n} (${it.dias} d RI→OC).`]); } }
    if((it.estado==='EN AT'||it.estado==='AT CERR') && it.dias!=null && it.dias>=80)
      out.push(['crit',`RIESGO: ${it.desc} — ${it.dias} d sin OC (${ESTADO[it.estado].t}). ${it.necoc?('NecOC '+it.necoc):''}`]);
  });
  if(!out.length) out.push(['','Sin variaciones relevantes respecto del corte anterior.']);
  return out.map(([cl,t])=>`<div class="find ${cl}">${escapeHtml(t)}</div>`).join('');
}

// ── COMPARAR ─────────────────────────────────────────────────────────────
function deltaSpan(d){ if(d==null)return''; if(d>0)return `<span class="delta up">▲ +${d}</span>`; if(d<0)return `<span class="delta down">▼ ${d}</span>`; return `<span class="delta eq">=</span>`; }
function renderCompare(){
  const a=DATA[+$('#selA').value], b=DATA[+$('#selB').value];
  const ka=a.kpis, kb=b.kpis;
  $('#cmpKpis').innerHTML =
     kpiCard(kb.oc,'OCs COLOCADAS','k2',kb.oc-ka.oc)
   + kpiCard(kb.en_gestion,'EN GESTIÓN','k3',kb.en_gestion-ka.en_gestion)
   + kpiCard(kb.at,'EN AT','k1',kb.at-ka.at)
   + kpiCard(kb.sin_emitir,'SIN EMITIR','k4',kb.sin_emitir-ka.sin_emitir);
  // pipeline table
  let h=`<tr><th>Especialidad · Etapa</th><th>${a.label}</th><th>${b.label}</th><th>Δ</th></tr>`;
  ['EL','IN'].forEach(e=>{
    h+=`<tr class="${e.toLowerCase()}"><td colspan="4"><b>${e==='EL'?'ELECTRICIDAD':'INSTRUMENTACIÓN'}</b></td></tr>`;
    [['solped','En SOLPED'],['ofertas','En Petición Ofertas'],['at','En AT'],['oc','Con OC'],['emitidas','Emitidas']].forEach(([k,l])=>{
      const va=(a.pipeline[e]||{})[k]||0, vb=(b.pipeline[e]||{})[k]||0;
      h+=`<tr><td style="padding-left:16px">${l}</td><td>${va}</td><td>${vb}</td><td>${deltaSpan(vb-va)}</td></tr>`;
    });
  });
  $('#cmpPipe').innerHTML=h;
  // criticos diff
  let d='';
  b.criticos.forEach(it=>{
    const p=a.criticos.find(x=>x.desc===it.desc);
    if(!p){ d+=`<div class="diffbox"><b>${escapeHtml(it.desc)}</b> — nuevo en ${b.label}</div>`; return; }
    let ch=[];
    if(p.estado!==it.estado) ch.push(`estado: ${ESTADO[p.estado].t} → <b>${ESTADO[it.estado].t}</b>`);
    if(p.oc_n!==it.oc_n && it.oc_n) ch.push(`OC: — → <b>${it.oc_n}</b>`);
    if(p.at!==it.at && it.at) ch.push(`AT cierre: ${p.at||'—'} → ${it.at}`);
    if(p.recof!==it.recof && it.recof) ch.push(`Rec.Of: ${p.recof||'—'} → ${it.recof}`);
    if(p.solped!==it.solped && it.solped) ch.push(`SOLPED: ${p.solped||'—'} → ${it.solped}`);
    if(p.status!==it.status) ch.push(`<span class="mini">comentario actualizado</span>`);
    if(ch.length) d+=`<div class="diffbox"><b>${it.esp} · ${escapeHtml(it.desc)}</b><br>${ch.join(' · ')}</div>`;
  });
  if(!d) d='<div class="mini">Sin cambios en ítems críticos entre ambos cortes.</div>';
  $('#cmpCrit').innerHTML=d;
}

// ── TENDENCIA (SVG) ──────────────────────────────────────────────────────
let trendMetric='oc', trendSerie='TOTAL';
function serieVal(c,serie,metric){
  if(serie==='TOTAL') return ((c.pipeline.EL||{})[metric]||0)+((c.pipeline.IN||{})[metric]||0);
  return (c.pipeline[serie]||{})[metric]||0;
}
function renderTrend(){
  const W=1040,H=300,pad=44, n=DATA.length;
  const vals=DATA.map(c=>serieVal(c,trendSerie,trendMetric));
  const maxv=Math.max(2,...vals);
  const x=i=> pad + i*( (W-pad*1.4-pad)/(Math.max(1,n-1)) );
  const y=v=> H-pad - v/maxv*(H-pad*1.8);
  const col= trendSerie==='EL'?'var(--el)':(trendSerie==='IN'?'var(--in)':'var(--azulm)');
  let g=`<svg viewBox="0 0 ${W} ${H}" width="100%">`;
  // grid + y labels
  for(let t=0;t<=4;t++){ const gy=pad+ t*(H-pad*1.8)/4; const gv=Math.round(maxv*(1-t/4));
    g+=`<line x1="${pad}" y1="${gy}" x2="${W-pad*0.4}" y2="${gy}" stroke="var(--borde)" stroke-width="1"/>`;
    g+=`<text x="${pad-8}" y="${gy+4}" font-size="11" fill="var(--mut)" text-anchor="end">${gv}</text>`;
  }
  // line
  let pts=vals.map((v,i)=>`${x(i)},${y(v)}`).join(' ');
  g+=`<polyline points="${pts}" fill="none" stroke="${col}" stroke-width="3"/>`;
  vals.forEach((v,i)=>{
    g+=`<circle cx="${x(i)}" cy="${y(v)}" r="4" fill="${col}"/>`;
    g+=`<text x="${x(i)}" y="${y(v)-9}" font-size="11" font-weight="700" fill="var(--tx)" text-anchor="middle">${v}</text>`;
    const lab=DATA[i].is_first?'ini':DATA[i].label.slice(0,5);
    g+=`<text x="${x(i)}" y="${H-pad+16}" font-size="10" fill="var(--mut)" text-anchor="middle" transform="rotate(35 ${x(i)} ${H-pad+16})">${lab}</text>`;
  });
  g+=`</svg>`;
  $('#trendChart').innerHTML=g;
  // table
  let h='<tr><th>Corte</th>'+DATA.map(c=>`<th>${c.is_first?'ini':c.label.slice(0,5)}</th>`).join('')+'</tr>';
  [['oc','Con OC'],['at','En AT'],['ofertas','En petición'],['solped','En SOLPED'],['emitidas','Emitidas']].forEach(([k,l])=>{
    h+=`<tr><td><b>${l}</b> (${trendSerie==='TOTAL'?'IN+EL':trendSerie})</td>`
      + DATA.map(c=>`<td>${serieVal(c,trendSerie,k)}</td>`).join('')+'</tr>';
  });
  $('#trendTbl').innerHTML=h;
}

// ── CRÍTICOS HISTORIAL ───────────────────────────────────────────────────
function fillCritSelect(){
  const names=[...new Set(DATA[lastIdx].criticos.map(c=>c.desc))];
  names.forEach(nm=>{const o=el('option',null,nm);o.value=nm;$('#selCrit').appendChild(o);});
}
function renderCritHist(){
  const nm=$('#selCrit').value;
  let h='<tr><th>Corte</th><th>Estado</th><th>RI</th><th>SOLPED</th><th>Rec.Of</th><th>AT</th><th>N° OC</th><th>RI→OC</th><th>Comentario</th></tr>';
  DATA.forEach(c=>{
    const it=c.criticos.find(x=>x.desc===nm); if(!it) return;
    const es=ESTADO[it.estado]||ESTADO['PENDIENTE'];
    const d=it.dias==null?'—':(it.dias+(it.estado==='OC'?' d':' d*'));
    h+=`<tr class="${it.esp.toLowerCase()}"><td><b>${c.is_first?'Inicial':c.label}</b></td>
      <td><span class="badge ${es.c}">${es.t}</span></td>
      <td>${it.ri||'—'}</td><td>${it.solped||'—'}</td><td>${it.recof||'—'}</td><td>${it.at||'—'}</td>
      <td>${it.oc_n||'—'}</td><td class="dias ${diasClass(it.dias)}">${d}</td>
      <td class="mini">${escapeHtml(it.status)}</td></tr>`;
  });
  $('#critHist').innerHTML=h;
}

// ── CAMBIOS & COMENTARIOS ────────────────────────────────────────────────
// resalta montos, fechas límite y palabras clave dentro del comentario
function highlight(txt){
  let s = escapeHtml(txt);
  const pats = [
    /USD\s?[\d.,]+\s?k?/gi, /\bantes del \d{1,2}\/\d{1,2}(?:\/\d{2,4})?\b/gi,
    /\bNecOC[^.,;]*/gi,
    /\b(aprobaci[oó]n|aprobar|vencid[ao]|vence|urgente|riesgo|cr[ií]tic[ao]s?|\bOCA\b|impacto|escalar|demora|estancad[ao])\b/gi
  ];
  pats.forEach(p=> s = s.replace(p, m=>`<mark>${m}</mark>`));
  return s;
}
// detecta ítems con contenido "relevante" (decisiones, deadlines, riesgos)
function esRelevante(txt){
  return /(USD|aprobaci|aprobar|antes del|vencid|urgente|riesgo|cr[ií]tic|OCA|NecOC|impacto|escalar|estancad)/i.test(txt||'');
}
function fillCambiosSelect(){
  DATA.forEach((c,i)=>{
    const o=el('option',null,(c.is_first?'Inicial (base)':c.label)+'  ·  '+c.tag);
    o.value=i; $('#selCambios').appendChild(o);
  });
  $('#selCambios').value=lastIdx;
}
function renderCambios(){
  const i=+$('#selCambios').value, c=DATA[i], prev=i>0?DATA[i-1]:null;
  $('#cambiosSub').textContent = prev
     ? ('Comparando '+ (c.is_first?'Inicial':c.label) +' vs corte anterior '+ (prev.is_first?'Inicial':prev.label))
     : 'Primer corte cargado (sin corte anterior para comparar).';

  // ── Movimientos del período ──
  const k=c.kpis, pk=prev?prev.kpis:null;
  let mov='<div class="mov">';
  const card=(n,l)=>`<div class="movcard"><div class="big">${n}</div><div class="lbl">${l}</div></div>`;
  if(pk){
    const doc=k.oc-pk.oc, dat=k.at-pk.at, dsol=k.solped-pk.solped, dof=k.ofertas-pk.ofertas;
    mov+=card((doc>=0?'+':'')+doc, 'OCs nuevas (total '+k.oc+')');
    mov+=card((dat>=0?'+':'')+dat, 'Δ en AT (total '+k.at+')');
    mov+=card((dof>=0?'+':'')+dof, 'Δ en petición (total '+k.ofertas+')');
    mov+=card((dsol>=0?'+':'')+dsol, 'Δ SOLPED (total '+k.solped+')');
  } else { mov+='<div class="mini">Sin corte anterior.</div>'; }
  mov+='</div>';
  // banderas de atención automáticas
  let flags=[];
  c.criticos.forEach(it=>{
    if(esRelevante(it.status) && (!prev || (prev.criticos.find(x=>x.desc===it.desc)||{}).status!==it.status))
      flags.push(it.desc);
  });
  if(flags.length) mov+='<div style="margin-top:8px"><span class="flag">Atención</span> '+
     'Comentarios relevantes nuevos en: '+flags.map(escapeHtml).join(' · ')+'</div>';
  $('#cambiosMov').innerHTML=mov;

  // ── Comentarios relevantes ──
  // orden: cambiados primero, luego relevantes, luego resto
  const its=[...c.criticos].map(it=>{
    const p=prev?prev.criticos.find(x=>x.desc===it.desc):null;
    const changed = p ? (p.status!==it.status || p.oc_n!==it.oc_n || p.estado!==it.estado) : false;
    return {it, p, changed, rel:esRelevante(it.status)};
  }).sort((a,b)=> (b.changed-a.changed) || (b.rel-a.rel));
  let h='';
  its.forEach(({it,p,changed})=>{
    const es=ESTADO[it.estado]||ESTADO['PENDIENTE'];
    const d = it.dias==null?'':` · <b>${it.dias}${it.estado==='OC'?' d RI→OC':' d*'}</b>`;
    h+=`<div class="cmt ${changed?'changed':''}">
      <div class="hd"><span class="badge ${es.c}">${es.t}</span>
        <b>${it.esp} · ${escapeHtml(it.desc)}</b>
        ${it.oc_n?`<span class="mini">OC ${it.oc_n}</span>`:''}
        ${changed?'<span class="badge b-recof">ACTUALIZADO</span>':''}
        <span class="mini">${d}</span></div>
      <div class="txt">${highlight(it.status||'—')}</div>
      ${(changed&&p&&p.status&&p.status!==it.status)?`<div class="prev"><b>Antes (${prev.is_first?'Inicial':prev.label}):</b> ${escapeHtml(p.status)}</div>`:''}
    </div>`;
  });
  $('#cambiosComents').innerHTML=h;
}

// ── eventos ──────────────────────────────────────────────────────────────
$('#selVista').onchange=renderVista;
$('#selCambios').onchange=renderCambios;
$('#selA').onchange=renderCompare; $('#selB').onchange=renderCompare;
$('#selCrit').onchange=renderCritHist;
document.querySelectorAll('[data-metric]').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('[data-metric]').forEach(x=>x.classList.remove('active'));
  b.classList.add('active'); trendMetric=b.dataset.metric; renderTrend();
});
document.querySelectorAll('[data-serie]').forEach(b=>b.onclick=()=>{
  document.querySelectorAll('[data-serie]').forEach(x=>x.classList.remove('active'));
  b.classList.add('active'); trendSerie=b.dataset.serie; renderTrend();
});

// init
fillCambiosSelect();
renderVista(); renderCambios(); renderCompare(); renderTrend(); fillCritSelect(); renderCritHist();
</script>
</body>
</html>
"""


def main():
    print('Leyendo planes de', PLANS_DIR)
    cuts = load_all()
    for c in cuts:
        print(f"  {c['label']:10} ({c['tag']})  OC={c['kpis']['oc']}  AT={c['kpis']['at']}  crit={len(c['criticos'])}")
    htmlout = render_html(cuts)
    with open(OUT_HTML, 'w', encoding='utf-8') as f:
        f.write(htmlout)
    print(f"\nOK -> {os.path.basename(OUT_HTML)} ({os.path.getsize(OUT_HTML)//1024} KB) · {len(cuts)} cortes")


if __name__ == '__main__':
    main()
