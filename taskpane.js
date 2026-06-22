/* Maestro de Contratos - panel lateral (Office.js)
   Lee la hoja "Maestro Contratos" del libro abierto y construye el modelo en memoria.
   Soporta multiseleccion de proyectos (por codigo, texto o "anadir todos" una raiz). */

let DATA = [];
let byCode = {};
let BASE = 'n';
let selected = [];
let lastHits = [];
let activeIdx = -1;

const SHEET_NAME = "Maestro Contratos";
const M = {contN:4,contB:5,factN:6,factB:7,pfactN:8,pfactB:9,cobN:10,cobB:11,pcfN:12,pcfB:13,pccN:14,pccB:15};

const HEADERS = {
  proy:"proyecto ge", desc:"descripcion proyecto ge", cli:"nombre cliente", cliCod:"codigo cliente",
  pais:"pais cliente", emp:"nombre empresa", pm:"project manager", div:"divisa contrato",
  origen:"origen contrato", estado:"estado",
  contN:"contrato neto", contB:"contrato bruto", factN:"facturado neto", factB:"facturado bruto",
  pfactN:"pendiente facturar neto", pfactB:"pendiente facturar bruto",
  cobN:"cobrado neto", cobB:"cobrado bruto",
  pcfN:"pdte cobro s/facturado neto", pcfB:"pdte cobro s/facturado bruto",
  pccN:"pdte cobro s/contrato neto", pccB:"pdte cobro s/contrato bruto",
};

function norm(s){
  return (s===null||s===undefined?'':String(s)).toLowerCase().trim()
    .normalize('NFD').replace(/[̀-ͯ]/g,'').replace(/\s+/g,' ');
}
function n2(v){ const x = typeof v==='number'? v : parseFloat(v); return isFinite(x)? x : 0; }
function sclean(v){ return v===null||v===undefined? '' : String(v).trim(); }
function uniq(a){ return [...new Set(a)]; }
function esc(t){ return (t||'').replace(/[&<>"]/g, c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }

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
      if(!target) throw new Error('No encuentro la hoja "'+SHEET_NAME+'". Hojas: '+names.join(', '));
      const sheet = sheets.getItem(target);
      const used = sheet.getUsedRange(true);
      used.load('values');
      await ctx.sync();
      buildModel(used.values, target);
    });
  }catch(e){
    setStatus('Error: '+((e && e.message) ? e.message : String(e)), 'err');
  }finally{
    document.getElementById('reload').disabled = false;
  }
}

function buildModel(values, sheetName){
  if(!values || values.length < 2){ setStatus('La hoja no tiene datos.', 'err'); return; }
  const head = values[0].map(norm);
  const col = {};
  for(const key in HEADERS) col[key] = head.indexOf(HEADERS[key]);
  if(col.proy < 0){ setStatus('No encuentro la columna "Proyecto GE".', 'err'); return; }
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
      p = projects[code] = { c:code, d:sclean(get(row,'desc')), cli:sclean(get(row,'cli')),
        cliCod:sclean(get(row,'cliCod')), pais:sclean(get(row,'pais')),
        emp:new Set(), pm:new Set(), div:new Set(), rows:[] };
    }
    const emp=sclean(get(row,'emp')), pm=sclean(get(row,'pm')), div=sclean(get(row,'div'));
    if(emp) p.emp.add(emp); if(pm) p.pm.add(pm); if(div) p.div.add(div);
    p.rows.push([ emp, sclean(get(row,'origen')), sclean(get(row,'estado')), div,
      n2(get(row,'contN')), n2(get(row,'contB')), n2(get(row,'factN')), n2(get(row,'factB')),
      n2(get(row,'pfactN')), n2(get(row,'pfactB')), n2(get(row,'cobN')), n2(get(row,'cobB')),
      n2(get(row,'pcfN')), n2(get(row,'pcfB')), n2(get(row,'pccN')), n2(get(row,'pccB')) ]);
  }
  DATA = Object.keys(projects).sort().map(c => {
    const p = projects[c];
    return { c:p.c, d:p.d, cli:p.cli, cliCod:p.cliCod, pais:p.pais,
      emp:[...p.emp].sort(), pm:[...p.pm].sort(), div:[...p.div].sort(), rows:p.rows,
      _idx:(p.c+' '+p.d+' '+p.cli+' '+p.cliCod).toLowerCase() };
  });
  byCode = {};
  DATA.forEach(p => byCode[p.c] = p);
  // refrescar seleccion existente con los nuevos objetos
  selected = selected.map(p => byCode[p.c]).filter(Boolean);

  document.getElementById('search').disabled = false;
  document.getElementById('empty-count').innerHTML = 'Hay <b>'+DATA.length+'</b> proyectos disponibles.';
  setStatus('Listo &middot; '+DATA.length+' proyectos &middot; '+nfilas+' lineas &middot; hoja "'+sheetName+'"', 'ok');
  render(); renderChips();
}

// ---------- formato ----------
function fmt(v, div){
  return new Intl.NumberFormat('es-ES',{maximumFractionDigits:0}).format(Math.round(v||0)) + (div?' '+div:'');
}
function negCls(v){ return v<0 ? 'num-neg' : ''; }
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
  if(!q){ $results.classList.remove('show'); lastHits=[]; return; }
  const terms = q.split(/\s+/);
  const hits = [];
  for(const p of DATA){ if(terms.every(t => p._idx.indexOf(t)>=0)){ hits.push(p); if(hits.length>=200) break; } }
  lastHits = hits;
  renderResults(hits);
}
function renderResults(hits){
  activeIdx = -1;
  if(!hits.length){ $results.innerHTML = '<div class="empty">Sin resultados</div>'; $results.classList.add('show'); return; }
  const head = '<div class="addall" onmousedown="addAll(event)">&#10133; Anadir todos los '+hits.length+' resultados</div>';
  const body = hits.slice(0,200).map((p,i)=>{
    const sel = selected.some(x=>x.c===p.c) ? ' sel' : '';
    return '<div class="item'+sel+'" data-code="'+p.c+'" data-i="'+i+'">'+
      '<div><span class="code">'+p.c+'</span> &nbsp;'+(esc(p.d)||'(sin descripcion)')+'</div>'+
      '<div class="sub">'+(esc(p.cli)||'(sin cliente)')+' &middot; '+p.emp.length+' empresa(s) &middot; '+(p.div.join(', ')||'-')+'</div>'+
    '</div>';
  }).join('');
  $results.innerHTML = head + body;
  $results.classList.add('show');
  [...$results.querySelectorAll('.item')].forEach(el=>{
    el.addEventListener('mousedown', e=>{ e.preventDefault(); addCode(el.dataset.code); });
  });
}
function addAll(e){ if(e) e.preventDefault(); lastHits.forEach(p=>addCode(p.c, true)); render(); renderChips(); renderResults(lastHits); }

$search.addEventListener('input', e=> search(e.target.value));
$search.addEventListener('focus', e=> { if(e.target.value) search(e.target.value); });
$search.addEventListener('keydown', e=>{
  const items = [...$results.querySelectorAll('.item')];
  if(e.key==='ArrowDown'){ e.preventDefault(); activeIdx=Math.min(activeIdx+1,items.length-1); }
  else if(e.key==='ArrowUp'){ e.preventDefault(); activeIdx=Math.max(activeIdx-1,0); }
  else if(e.key==='Enter'){ e.preventDefault(); if(activeIdx>=0&&items[activeIdx]) addCode(items[activeIdx].dataset.code); return; }
  else if(e.key==='Escape'){ $results.classList.remove('show'); return; }
  items.forEach((el,i)=> el.classList.toggle('active', i===activeIdx));
  if(items[activeIdx]) items[activeIdx].scrollIntoView({block:'nearest'});
});
document.addEventListener('click', e=>{ if(!e.target.closest('.searchbar')) $results.classList.remove('show'); });

function addCode(code, bulk){
  const p = byCode[code]; if(!p) return;
  if(!selected.some(x=>x.c===code)) selected.push(p);
  if(!bulk){ render(); renderChips(); renderResults(lastHits); }
}
function removeCode(code){ selected = selected.filter(x=>x.c!==code); render(); renderChips(); if(lastHits.length) renderResults(lastHits); }
function clearAll(){ selected=[]; render(); renderChips(); if(lastHits.length) renderResults(lastHits); }

function renderChips(){
  const $c = document.getElementById('chips');
  if(!selected.length){ $c.innerHTML=''; return; }
  $c.innerHTML = selected.map(p=>
    '<span class="chip"><span class="cc">'+p.c+'</span> <span class="x" onclick="removeCode(\''+p.c+'\')">&times;</span></span>').join('')
    + '<span class="clr" onclick="clearAll()">Quitar todos ('+selected.length+')</span>';
}

function setBase(b){
  BASE = b;
  document.getElementById('btn-neto').classList.toggle('on', b==='n');
  document.getElementById('btn-bruto').classList.toggle('on', b==='b');
  render();
}

// ---------- render ----------
function mi(name){ return BASE==='n' ? M[name+'N'] : M[name+'B']; }
function sum(rs, idx){ let t=0; for(const r of rs) t+=r[idx]; return t; }
function kpi(label, val, cls){ return '<div class="kpi '+(cls||'')+'"><div class="l">'+label+'</div><div class="v">'+val+'</div></div>'; }
function kpiPct(label, ratio, pct){
  const txt = ratio===null ? '&mdash;' : new Intl.NumberFormat('es-ES',{maximumFractionDigits:1}).format(ratio*100)+'%';
  return '<div class="kpi"><div class="l">'+label+'</div><div class="v">'+txt+'</div><div class="bar"><i style="width:'+pct+'%"></i></div></div>';
}
function commonRoot(codes){
  if(codes.length<2) return '';
  let pre = codes[0];
  for(const c of codes){ while(!c.startsWith(pre)) pre = pre.slice(0,-1); if(!pre) break; }
  return pre.length>=2 ? pre : '';
}
function fact_(l,v){ return '<div class="fact"><div class="l">'+l+'</div><div class="v">'+v+'</div></div>'; }

function render(){
  const $d = document.getElementById('detail');
  const es = document.getElementById('empty-state');
  if(!selected.length){ if(es) es.style.display='block'; $d.innerHTML=''; return; }
  if(es) es.style.display='none';

  const ci=mi('cont'), fii=mi('fact'), pfi=mi('pfact'), coi=mi('cob'), pcfi=mi('pcf'), pcci=mi('pcc');
  const entries = [];
  selected.forEach(p => p.rows.forEach(r => entries.push({p, r})));
  const multi = selected.length > 1;
  const codes = selected.map(p=>p.c);
  const root = commonRoot(codes);

  let title, desc='';
  if(multi){ title = selected.length+' proyectos'+(root?' &middot; raiz <span style="color:var(--accent)">'+root+'</span>':''); }
  else { title = selected[0].c; desc = esc(selected[0].d)||'(sin descripcion)'; }

  const projList = selected.map(p=>'<span style="color:var(--accent2)">'+p.c+'</span> '+esc(p.d)).join('<br>');
  const clientes = uniq(selected.map(p=>p.cli).filter(Boolean)).map(esc).join('<br>')||'&mdash;';
  const paises = uniq(selected.map(p=>p.pais).filter(Boolean)).join(', ')||'&mdash;';
  const empresas = uniq(entries.map(e=>e.r[0]).filter(Boolean)).map(esc).join('<br>')||'&mdash;';
  const pms = uniq([].concat(...selected.map(p=>p.pm))).map(esc).join('<br>')||'&mdash;';
  const divisas = uniq(entries.map(e=>e.r[3]).filter(Boolean)).join(', ')||'&mdash;';

  let html = '<div class="card head"><div class="pcode">'+title+'</div>'+
    (desc?'<div class="pdesc">'+desc+'</div>':'')+'<div class="facts">'+
    (multi?fact_('Proyectos ('+selected.length+')', projList):'')+
    fact_('Cliente(s)', clientes)+fact_('Pais(es)', paises)+fact_('Empresa(s)', empresas)+
    fact_('Project Manager(s)', pms)+fact_('Divisa(s)', divisas)+fact_('Lineas', String(entries.length))+
    '</div></div>';

  const divs = {};
  for(const e of entries){ const d=e.r[3]||'-'; (divs[d]=divs[d]||[]).push(e.r); }
  for(const d of Object.keys(divs).sort()){
    const rs = divs[d];
    const tC=sum(rs,ci), tF=sum(rs,fii), tPf=sum(rs,pfi), tCob=sum(rs,coi), tPcf=sum(rs,pcfi), tPcc=sum(rs,pcci);
    const eje = tC!==0 ? tF/tC : null;
    const ejePct = eje!==null ? Math.max(0,Math.min(100, eje*100)) : 0;
    html += '<div class="card divblock"><h3>&#128181; Resumen en '+d+' <span class="hint" style="font-weight:400">('+rs.length+' linea/s)</span></h3>'+
      '<div class="kpis">'+kpi('Contrato', fmt(tC,d))+kpi('Facturado', fmt(tF,d))+kpiPct('% Ejecucion', eje, ejePct)+
      kpi('Pendiente facturar', fmt(tPf,d), tPf<0?'neg':'')+kpi('Cobrado', fmt(tCob,d))+
      kpi('Pdte cobro s/fact', fmt(tPcf,d), tPcf>0?'neg':'')+kpi('Pdte cobro s/contrato', fmt(tPcc,d), tPcc>0?'neg':'')+
      '</div></div>';
  }

  html += '<div class="card"><h3 style="margin:0 0 6px;font-size:12px;color:var(--accent2)">Detalle por origen</h3>'+
    '<div class="tablewrap"><table><thead><tr>'+
    '<th class="l">Proy.</th><th class="l">Empresa</th><th class="l">Origen</th><th class="l">Estado</th><th class="l">Div</th>'+
    '<th>Contrato</th><th>Facturado</th><th>% Ej.</th><th>Pdte fact</th><th>Cobrado</th><th>Pdte s/fact</th><th>Pdte s/contr</th>'+
    '</tr></thead><tbody>';
  const ord = entries.slice().sort((a,b)=> (a.p.c+a.r[3]+a.r[1]).localeCompare(b.p.c+b.r[3]+b.r[1]));
  for(const e of ord){
    const r=e.r, c=r[ci], f=r[fii], ej=c!==0?f/c:null;
    const ejt = ej===null?'&mdash;':new Intl.NumberFormat('es-ES',{maximumFractionDigits:1}).format(ej*100)+'%';
    html += '<tr><td class="l" style="color:var(--accent2)">'+e.p.c+'</td>'+
      '<td class="l">'+esc(r[0])+'</td><td class="l">'+esc(r[1])+'</td><td class="l">'+stateBadge(r[2])+'</td><td class="l">'+r[3]+'</td>'+
      '<td class="'+negCls(c)+'">'+fmt(c)+'</td><td class="'+negCls(f)+'">'+fmt(f)+'</td><td>'+ejt+'</td>'+
      '<td class="'+negCls(r[pfi])+'">'+fmt(r[pfi])+'</td><td class="'+negCls(r[coi])+'">'+fmt(r[coi])+'</td>'+
      '<td class="'+negCls(r[pcfi])+'">'+fmt(r[pcfi])+'</td><td class="'+negCls(r[pcci])+'">'+fmt(r[pcci])+'</td></tr>';
  }
  html += '</tbody></table></div></div>';
  $d.innerHTML = html;
}

Office.onReady((info) => {
  if(info.host === Office.HostType.Excel){ loadData(); }
  else { setStatus('Este complemento es para Excel.', 'err'); }
});
