/* Maestro de Contratos - panel lateral (Office.js)
   Lee la hoja "Maestro Contratos" del libro abierto y construye el modelo en memoria. */

let DATA = [];
let byCode = {};
let BASE = 'n';        // 'n' neto / 'b' bruto
let current = null;
let activeIdx = -1;

const SHEET_NAME = "Maestro Contratos";

// indices de medidas dentro de la fila normalizada
const M = {contN:4,contB:5,factN:6,factB:7,pfactN:8,pfactB:9,cobN:10,cobB:11,pcfN:12,pcfB:13,pccN:14,pccB:15};

// ---- cabeceras esperadas (clave -> texto de cabecera normalizado) ----
const HEADERS = {
  proy:   "proyecto ge",
  desc:   "descripcion proyecto ge",
  cli:    "nombre cliente",
  cliCod: "codigo cliente",
  pais:   "pais cliente",
  emp:    "nombre empresa",
  pm:     "project manager",
  div:    "divisa contrato",
  origen: "origen contrato",
  estado: "estado",
  contN:  "contrato neto",  contB:"contrato bruto",
  factN:  "facturado neto", factB:"facturado bruto",
  pfactN: "pendiente facturar neto", pfactB:"pendiente facturar bruto",
  cobN:   "cobrado neto", cobB:"cobrado bruto",
  pcfN:   "pdte cobro s/facturado neto", pcfB:"pdte cobro s/facturado bruto",
  pccN:   "pdte cobro s/contrato neto",  pccB:"pdte cobro s/contrato bruto",
};

function norm(s){
  return (s===null||s===undefined?'':String(s))
    .toLowerCase().trim()
    .normalize('NFD').replace(/[̀-ͯ]/g,'')   // quita acentos
    .replace(/\s+/g,' ');
}
function n2(v){ const x = typeof v==='number'? v : parseFloat(v); return isFinite(x)? x : 0; }
function sclean(v){ return v===null||v===undefined? '' : String(v).trim(); }

// ---------- carga desde el libro ----------
function setStatus(msg, cls){
  const el = document.getElementById('status');
  el.className = 'statusline' + (cls? ' '+cls : '');
  el.innerHTML = msg;
}

async function loadData(){
  document.getElementById('reload').disabled = true;
  setStatus('<span class="spinner"></span>Leyendo hoja "'+SHEET_NAME+'"&hellip;');
  try{
    await Excel.run(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load('items/name');
      await ctx.sync();
      const names = sheets.items.map(s => s.name);
      let target = names.find(nm => norm(nm) === norm(SHEET_NAME)) || names.find(nm => norm(nm).indexOf('maestro')>=0);
      if(!target){
        throw new Error('No encuentro la hoja "'+SHEET_NAME+'". Hojas: '+names.join(', '));
      }
      const sheet = sheets.getItem(target);
      const used = sheet.getUsedRange(true);
      used.load('values');
      await ctx.sync();
      buildModel(used.values, target);
    });
  }catch(e){
    const msg = (e && e.message) ? e.message : String(e);
    setStatus('Error: '+msg, 'err');
  }finally{
    document.getElementById('reload').disabled = false;
  }
}

function buildModel(values, sheetName){
  if(!values || values.length < 2){ setStatus('La hoja no tiene datos.', 'err'); return; }
  const head = values[0].map(norm);
  const col = {};
  for(const key in HEADERS){
    col[key] = head.indexOf(HEADERS[key]);
  }
  if(col.proy < 0){
    setStatus('No encuentro la columna "Proyecto GE" en la hoja "'+sheetName+'".', 'err');
    return;
  }
  const get = (row, key) => { const i = col[key]; return i>=0 ? row[i] : null; };

  const projects = {};
  let nfilas = 0;
  for(let r = 1; r < values.length; r++){
    const row = values[r];
    const code = sclean(get(row,'proy'));
    if(!code) continue;
    nfilas++;
    let p = projects[code];
    if(!p){
      p = projects[code] = {
        c: code, d: sclean(get(row,'desc')),
        cli: sclean(get(row,'cli')), cliCod: sclean(get(row,'cliCod')),
        pais: sclean(get(row,'pais')),
        emp: new Set(), pm: new Set(), div: new Set(), rows: []
      };
    }
    const emp = sclean(get(row,'emp')), pm = sclean(get(row,'pm')), div = sclean(get(row,'div'));
    if(emp) p.emp.add(emp);
    if(pm)  p.pm.add(pm);
    if(div) p.div.add(div);
    p.rows.push([
      emp, sclean(get(row,'origen')), sclean(get(row,'estado')), div,
      n2(get(row,'contN')),  n2(get(row,'contB')),
      n2(get(row,'factN')),  n2(get(row,'factB')),
      n2(get(row,'pfactN')), n2(get(row,'pfactB')),
      n2(get(row,'cobN')),   n2(get(row,'cobB')),
      n2(get(row,'pcfN')),   n2(get(row,'pcfB')),
      n2(get(row,'pccN')),   n2(get(row,'pccB')),
    ]);
  }

  DATA = Object.keys(projects).sort().map(c => {
    const p = projects[c];
    return { c:p.c, d:p.d, cli:p.cli, cliCod:p.cliCod, pais:p.pais,
             emp:[...p.emp].sort(), pm:[...p.pm].sort(), div:[...p.div].sort(),
             rows:p.rows,
             _idx:(p.c+' '+p.d+' '+p.cli+' '+p.cliCod).toLowerCase() };
  });
  byCode = {};
  DATA.forEach(p => byCode[p.c] = p);

  document.getElementById('search').disabled = false;
  document.getElementById('empty-count').innerHTML = 'Hay <b>'+DATA.length+'</b> proyectos disponibles.';
  setStatus('Listo &middot; '+DATA.length+' proyectos &middot; '+nfilas+' lineas &middot; hoja "'+sheetName+'"', 'ok');

  // si habia un proyecto abierto, refrescarlo
  if(current && byCode[current.c]){ current = byCode[current.c]; render(); }
}

// ---------- formato ----------
function fmt(v, div){
  const n = new Intl.NumberFormat('es-ES',{maximumFractionDigits:0}).format(Math.round(v||0));
  return n + (div ? ' '+div : '');
}
function negCls(v){ return v<0 ? 'num-neg' : ''; }
function esc(t){ return (t||'').replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function stateBadge(st){
  if(st==='Terminado') return '<span class="badge b-term">Terminado</span>';
  if(st==='Abierto') return '<span class="badge b-abierto">Abierto</span>';
  if(st==='Completado pdte cobro') return '<span class="badge b-pdte">Compl. pdte cobro</span>';
  return '<span class="badge b-none">&mdash;</span>';
}

// ---------- buscador ----------
const $search = document.getElementById('search');
const $results = document.getElementById('results');

function search(q){
  q = q.trim().toLowerCase();
  if(!q){ $results.classList.remove('show'); return; }
  const terms = q.split(/\s+/);
  const hits = [];
  for(const p of DATA){
    if(terms.every(t => p._idx.indexOf(t)>=0)){ hits.push(p); if(hits.length>=60) break; }
  }
  renderResults(hits);
}
function renderResults(hits){
  activeIdx = -1;
  if(!hits.length){ $results.innerHTML = '<div class="empty">Sin resultados</div>'; $results.classList.add('show'); return; }
  $results.innerHTML = hits.map((p,i)=>
    '<div class="item" data-code="'+p.c+'" data-i="'+i+'">'+
      '<div><span class="code">'+p.c+'</span> &nbsp;'+(esc(p.d)||'(sin descripcion)')+'</div>'+
      '<div class="sub">'+(esc(p.cli)||'(sin cliente)')+' &middot; '+p.emp.length+' empresa(s) &middot; '+(p.div.join(', ')||'-')+'</div>'+
    '</div>').join('');
  $results.classList.add('show');
  [...$results.querySelectorAll('.item')].forEach(el=>{
    el.addEventListener('mousedown', e=>{ e.preventDefault(); selectCode(el.dataset.code); });
  });
}
$search.addEventListener('input', e=> search(e.target.value));
$search.addEventListener('focus', e=> { if(e.target.value) search(e.target.value); });
$search.addEventListener('keydown', e=>{
  const items = [...$results.querySelectorAll('.item')];
  if(e.key==='ArrowDown'){ e.preventDefault(); activeIdx=Math.min(activeIdx+1,items.length-1); }
  else if(e.key==='ArrowUp'){ e.preventDefault(); activeIdx=Math.max(activeIdx-1,0); }
  else if(e.key==='Enter'){ e.preventDefault(); if(activeIdx>=0&&items[activeIdx]) selectCode(items[activeIdx].dataset.code); return; }
  else if(e.key==='Escape'){ $results.classList.remove('show'); return; }
  items.forEach((el,i)=> el.classList.toggle('active', i===activeIdx));
  if(items[activeIdx]) items[activeIdx].scrollIntoView({block:'nearest'});
});
document.addEventListener('click', e=>{ if(!e.target.closest('.searchbar')) $results.classList.remove('show'); });

function selectCode(code){
  current = byCode[code];
  if(!current) return;
  $search.value = code;
  $results.classList.remove('show');
  render();
}
function setBase(b){
  BASE = b;
  document.getElementById('btn-neto').classList.toggle('on', b==='n');
  document.getElementById('btn-bruto').classList.toggle('on', b==='b');
  if(current) render();
}

// ---------- render ----------
function mi(name){ return BASE==='n' ? M[name+'N'] : M[name+'B']; }
function sum(rs, fn){ let t=0; for(const r of rs) t+=fn(r); return t; }
function kpi(label, val, cls){ return '<div class="kpi '+(cls||'')+'"><div class="l">'+label+'</div><div class="v">'+val+'</div></div>'; }
function kpiPct(label, ratio, pct){
  const txt = ratio===null ? '&mdash;' : new Intl.NumberFormat('es-ES',{maximumFractionDigits:1}).format(ratio*100)+'%';
  return '<div class="kpi"><div class="l">'+label+'</div><div class="v">'+txt+'</div><div class="bar"><i style="width:'+pct+'%"></i></div></div>';
}

function render(){
  const p = current;
  document.getElementById('empty-state').style.display = 'none';
  const cont=(r)=>r[mi('cont')], fact=(r)=>r[mi('fact')], pfact=(r)=>r[mi('pfact')],
        cob=(r)=>r[mi('cob')], pcf=(r)=>r[mi('pcf')], pcc=(r)=>r[mi('pcc')];

  const divs = {};
  for(const r of p.rows){ const d=r[3]||'-'; (divs[d]=divs[d]||[]).push(r); }

  let html = '<div class="card head">'+
    '<div class="pcode">'+p.c+'</div>'+
    '<div class="pdesc">'+(esc(p.d)||'(sin descripcion)')+'</div>'+
    '<div class="facts">'+
      fact_('Cliente', (esc(p.cli)||'&mdash;')+(p.cliCod?' <span style="color:var(--muted)">('+p.cliCod+')</span>':''))+
      fact_('Pais cliente', p.pais||'&mdash;')+
      fact_('Empresa(s)', p.emp.map(esc).join('<br>')||'&mdash;')+
      fact_('Project Manager', p.pm.map(esc).join('<br>')||'&mdash;')+
      fact_('Divisa(s)', p.div.join(', ')||'&mdash;')+
      fact_('Lineas (origen x divisa)', String(p.rows.length))+
    '</div></div>';

  for(const d of Object.keys(divs).sort()){
    const rs = divs[d];
    const tC=sum(rs,cont), tF=sum(rs,fact), tPf=sum(rs,pfact), tCob=sum(rs,cob), tPcf=sum(rs,pcf), tPcc=sum(rs,pcc);
    const eje = tC!==0 ? tF/tC : null;
    const ejePct = eje!==null ? Math.max(0,Math.min(100, eje*100)) : 0;
    html += '<div class="card divblock">'+
      '<h3>&#128181; Resumen en '+d+' <span class="hint" style="font-weight:400">('+rs.length+' linea/s)</span></h3>'+
      '<div class="kpis">'+
        kpi('Contrato', fmt(tC,d))+
        kpi('Facturado', fmt(tF,d))+
        kpiPct('% Ejecucion', eje, ejePct)+
        kpi('Pendiente facturar', fmt(tPf,d), tPf<0?'neg':'')+
        kpi('Cobrado', fmt(tCob,d))+
        kpi('Pdte cobro s/fact', fmt(tPcf,d), tPcf>0?'neg':'')+
        kpi('Pdte cobro s/contrato', fmt(tPcc,d), tPcc>0?'neg':'')+
      '</div></div>';
  }

  html += '<div class="card"><h3 style="margin:0 0 6px;font-size:12px;color:var(--accent2)">Detalle por origen</h3>'+
    '<div class="tablewrap"><table><thead><tr>'+
      '<th class="l">Empresa</th><th class="l">Origen</th><th class="l">Estado</th><th class="l">Div</th>'+
      '<th>Contrato</th><th>Facturado</th><th>% Ej.</th><th>Pdte fact</th>'+
      '<th>Cobrado</th><th>Pdte s/fact</th><th>Pdte s/contr</th>'+
    '</tr></thead><tbody>';
  const ordered = [...p.rows].sort((a,b)=> (a[3]+a[1]).localeCompare(b[3]+b[1]));
  for(const r of ordered){
    const c=cont(r), f=fact(r), e=c!==0?f/c:null;
    const ept = e===null?'&mdash;':new Intl.NumberFormat('es-ES',{maximumFractionDigits:1}).format(e*100)+'%';
    html += '<tr>'+
      '<td class="l">'+esc(r[0])+'</td><td class="l">'+esc(r[1])+'</td><td class="l">'+stateBadge(r[2])+'</td><td class="l">'+r[3]+'</td>'+
      '<td class="'+negCls(c)+'">'+fmt(c)+'</td>'+
      '<td class="'+negCls(f)+'">'+fmt(f)+'</td>'+
      '<td>'+ept+'</td>'+
      '<td class="'+negCls(pfact(r))+'">'+fmt(pfact(r))+'</td>'+
      '<td class="'+negCls(cob(r))+'">'+fmt(cob(r))+'</td>'+
      '<td class="'+negCls(pcf(r))+'">'+fmt(pcf(r))+'</td>'+
      '<td class="'+negCls(pcc(r))+'">'+fmt(pcc(r))+'</td>'+
    '</tr>';
  }
  html += '</tbody></table></div></div>';

  document.getElementById('detail').innerHTML = html;
  window.scrollTo({top:0,behavior:'smooth'});
}
function fact_(l,v){ return '<div class="fact"><div class="l">'+l+'</div><div class="v">'+v+'</div></div>'; }

// ---------- arranque ----------
Office.onReady((info) => {
  if(info.host === Office.HostType.Excel){
    loadData();
  }else{
    setStatus('Este complemento es para Excel.', 'err');
  }
});
