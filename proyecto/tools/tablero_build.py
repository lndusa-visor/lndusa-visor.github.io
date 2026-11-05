#!/usr/bin/env python3
"""Arma proyecto/tablero.html desde el HTML del tablero con PR SAP.
Uso: python3 tools/tablero_build.py <tablero_con_PR_SAP.html> <salida.html>
El original no se toca. La copia no lleva proyectos adentro: los lee de ../tablero-datos/.
Cada cambio exige encontrar su texto exacto una sola vez; si el tablero cambia, se detiene."""
import re, sys
src, out = sys.argv[1], sys.argv[2]
EMBED = '--embed' in sys.argv   # el HTML del compañero con sus proyectos adentro + «Nuevo proyecto» en un panel; los proyectos nuevos se leen de GitHub
s = open(src, encoding='utf8').read()
def R(a, b):
    global s
    if s.count(a) != 1: sys.exit(f'No encontré una sola vez: {a[:70]}')
    s = s.replace(a, b)

LOADER = r"""const TB = (function(){ const base = '../tablero-datos/', get = u => { const r = new XMLHttpRequest(); r.open('GET', u+'?v='+Date.now(), false); r.send(); if(r.status!==200) throw new Error(u+' · '+r.status); return JSON.parse(r.responseText); };
  const stop = msg => { document.documentElement.innerHTML = `<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Tablero</title><style>body{margin:0;padding:24px 16px;background:#f3f4f2;color:#1d2326;font:16px/1.45 system-ui,sans-serif}main{max-width:560px;margin:10vh auto}a{display:inline-block;margin-top:12px;padding:10px 14px;background:#1f5fa8;color:#fff;text-decoration:none;border-radius:5px}@media (prefers-color-scheme:dark){body{background:#15191b;color:#e8ecee}}</style></head><body><main><h1>Tablero de avance</h1><p>${msg}</p><a href="index.html">Nuevo proyecto</a></main></body>`; throw new Error('Tablero sin datos'); };
  let L = null; try{ L = get(base+'index.json'); }catch(e){ stop('Todavía no hay proyectos publicados. Crea uno en «Nuevo proyecto» y presiona «Publicar en el tablero».'); }
  const list = L.proyectos||[], q = new URLSearchParams(location.search).get('g'); let sv = null; try{ sv = localStorage.getItem('wtb-g'); }catch(e){}
  const cur = list.find(p=>p.code===q) || list.find(p=>p.code===sv) || list[0]; if(!cur) stop('Todavía no hay proyectos publicados. Crea uno en «Nuevo proyecto» y presiona «Publicar en el tablero».');
  try{ localStorage.setItem('wtb-g', cur.code); }catch(e){}
  let meta = null; try{ meta = get(base+encodeURIComponent(cur.folder)+'/proyecto.json'); }catch(e){ stop('No pude leer el proyecto '+cur.code+'.'); }
  return {list, cur, title:meta.title, projs:meta.projs, keys:meta.keys, scene:k=>{ try{ return get(base+encodeURIComponent(cur.folder)+'/'+encodeURIComponent(k)+'.json'); }catch(e){ stop('No pude leer el 3D de '+cur.code+'.'); } }}; })();
const PROJS = TB.projs;
const PKEYS = TB.keys;"""

TB_EMBED = r"""const TB = (function(){ const base = 'https://lndusa-visor.github.io/tablero-datos/', q = new URLSearchParams(location.search).get('g'), local = {code:'US.068.4', title:'US.068.4 - PDC Piperacks · WTS 55 m & WWTP 30 m'};
  PROJS.w55.sheet = 'BOM WTS-55M - CONNECTION'; PROJS.w30.sheet = 'BOM WWTP-30M - CONNECTION';
  const get = u => { const r = new XMLHttpRequest(); r.open('GET', u+'?v='+Date.now(), false); r.send(); if(r.status!==200) throw new Error(u+' · '+r.status); return JSON.parse(r.responseText); };
  if(q && q!==local.code){ try{ const f = q.replace(/[^A-Za-z0-9._-]+/g,'_'), m = get(base+encodeURIComponent(f)+'/proyecto.json'); Object.assign(PROJS, m.projs); return {list:[local,{code:q}], cur:{code:q}, title:m.title, keys:m.keys, base, scene:k=>get(base+encodeURIComponent(f)+'/'+encodeURIComponent(k)+'.json')}; }catch(e){ console.warn('proyecto',e); } }
  return {list:[local], cur:local, title:local.title, keys:['w55','w30'], base, scene:k=>SCENES[k]}; })();
const PKEYS = TB.keys;"""
PANEL = r"""<div id="np-box" style="display:none;position:fixed;inset:0;z-index:9999;background:var(--bg,#fff);grid-template-rows:auto minmax(0,1fr)"><div style="display:flex;align-items:center;justify-content:space-between;gap:12px;padding:8px 14px;border-bottom:2px solid var(--brand,#1f5fa8)"><b id="np-title" style="font-size:15px"></b><button id="np-close" type="button" class="hload">✕</button></div><iframe id="np-frame" title="Nuevo proyecto" style="width:100%;height:100%;border:0"></iframe></div>
<script id="nuevo-proyecto">
/* ===== Agregado: «Nuevo proyecto» dentro del tablero. No cambia las pestañas: abre la carga de BOM, IFC y pernos en un panel y agrega los proyectos publicados al selector. ===== */
(function(){ const L=(es,en)=>LANG==='es'?es:en, $=id=>document.getElementById(id), fill=()=>{ const gp=$('grpsel'), gb=$('grpbox'), nw=$('h-newt'), gl=$('grplb'); if(nw) nw.textContent=L('Nuevo proyecto','New project'); if(gl) gl.textContent=L('Contrato','Contract'); if(!gp) return; gp.innerHTML=TB.list.map(p=>`<option value="${String(p.code).replace(/"/g,'&quot;')}">${p.code}</option>`).join(''); gp.value=TB.cur.code; gb.hidden=TB.list.length<2; };
  const lists=()=>TB.solo?Promise.resolve():fetch(TB.base+'index.json?v='+Date.now(),{cache:'no-store'}).then(r=>r.ok?r.json():null).then(j=>{ if(!j) return; const seen=new Set(TB.list.map(p=>p.code)); (j.proyectos||[]).forEach(p=>{ if(!seen.has(p.code)){ TB.list.push({code:p.code}); seen.add(p.code); } }); fill(); }).catch(()=>{});
  $('grpsel').onchange=e=>{ location.href=location.pathname+'?g='+encodeURIComponent(e.target.value); };
  const box=$('np-box'); $('h-new').onclick=()=>{ const f=$('np-frame'); if(!f.src) f.src='https://lndusa-visor.github.io/proyecto/?v='+Date.now(); box.style.display='grid'; }; $('np-close').onclick=()=>{ box.style.display='none'; lists(); };
  $('np-title').textContent='Nuevo proyecto'; fill(); ['lang-en','lang-es'].forEach(i=>{ const b=$(i); if(b) b.addEventListener('click',()=>setTimeout(fill,0)); }); lists(); })();
</script>"""
# fondo blanco en el 3D, en tema claro y oscuro; el texto de ayuda pasa a gris oscuro para que se lea
R('.hint{position:absolute;left:10px;bottom:8px;font:11.5px var(--f-mono);color:#93a1ae;pointer-events:none}',
  '.hint{position:absolute;left:10px;bottom:8px;font:11.5px var(--f-mono);color:#4f5b66;pointer-events:none}\n:root,:root[data-theme]{--view:#ffffff!important}')
VIRGEN = '--virgen' in sys.argv   # sin proyectos adentro: mismo archivo para cualquier proyecto; si el nombre lleva el código (US.071.1), abre ese
TB_VIRGEN = r"""const TB = (function(){ const base = 'https://lndusa-visor.github.io/tablero-datos/', NP = 'https://lndusa-visor.github.io/proyecto/', get = u => { const r = new XMLHttpRequest(); r.open('GET', u+'?v='+Date.now(), false); r.send(); if(r.status!==200) throw new Error(u+' · '+r.status); return JSON.parse(r.responseText); }, E = t => String(t==null?'':t).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
  const home = (msg, list) => { document.documentElement.innerHTML = `<head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Tablero de avance</title><style>:root{--bg:#f3f4f2;--ink:#1d2326;--card:#fff;--line:#d5dbdf;--acc:#1f5fa8}@media (prefers-color-scheme:dark){:root{--bg:#15191b;--ink:#e8ecee;--card:#1f2528;--line:#33393d;--acc:#6ea8e6}}body{margin:0;padding:32px 16px;background:var(--bg);color:var(--ink);font:16px/1.45 system-ui,sans-serif}main{max-width:620px;margin:0 auto}h1{margin:0 0 6px;font-size:26px}p{margin:0 0 18px}.b{display:inline-block;margin:0 0 18px;padding:11px 16px;border:0;border-radius:5px;background:var(--acc);color:#fff;font:600 15px system-ui,sans-serif;cursor:pointer}a.p{display:block;margin:0 0 8px;padding:12px 14px;background:var(--card);border:1px solid var(--line);border-radius:5px;color:inherit;text-decoration:none}a.p b{display:block}a.p span{font-size:14px;opacity:.75}#n{display:none;position:fixed;inset:0;background:var(--bg);grid-template-rows:auto minmax(0,1fr)}#n div{display:flex;justify-content:space-between;align-items:center;padding:8px 14px;border-bottom:2px solid var(--acc)}#n iframe{width:100%;height:100%;border:0}#n button{padding:6px 12px;cursor:pointer}</style></head><body><main><h1>Tablero de avance</h1><p>${E(msg)}</p><button class="b" id="o" type="button">＋ Nuevo proyecto</button>${list.map(p=>`<a class="p" href="?g=${encodeURIComponent(p.code)}"><b>${E(p.code)}</b><span>${E(p.name||p.title||'')}</span></a>`).join('')}</main><div id="n"><div><b>Nuevo proyecto</b><button id="c" type="button">✕</button></div><iframe id="f" title="Nuevo proyecto"></iframe></div></body>`;
    const $ = i => document.getElementById(i); $('o').onclick = () => { const f=$('f'); if(!f.src) f.src=NP+'?v='+Date.now(); $('n').style.display='grid'; }; $('c').onclick = () => location.reload(); throw new Error('Elegir proyecto'); };
  // cada HTML abre un solo proyecto: el que lleva en su nombre («US.071.1 Tablero.html»); no lista los demás
  const name = decodeURIComponent(location.pathname.split('/').pop()||''), q = new URLSearchParams(location.search).get('g'), m = name.match(/^(.+?)\s+Tablero/i), u = name.match(/US[._ -]?\d{3}[._ -]\d+/i), code = q || (m && m[1].trim()) || (u && u[0].toUpperCase().replace(/[_ -]/g,'.'));
  if(!code) home('Crea un proyecto. Al terminar se descarga su propio HTML: guárdalo en la carpeta del proyecto y ábrelo desde ahí.', []);
  const folder = code.replace(/[^A-Za-z0-9._-]+/g,'_'), cur = {code, folder}; let meta = null;
  try{ meta = get(base+encodeURIComponent(folder)+'/proyecto.json'); }catch(e){ home('Todavía no encuentro el proyecto '+code+'. Si lo acabas de crear, espera 2 minutos y recarga. Revisa también la conexión a internet.', []); }
  return {solo:true, list:[cur], cur, base, title:meta.title, projs:meta.projs, keys:meta.keys, scene:k=>{ try{ return get(base+encodeURIComponent(folder)+'/'+encodeURIComponent(k)+'.json'); }catch(e){ home('No pude leer el 3D de '+code+'.', []); } }}; })();
const PROJS = TB.projs;
const PKEYS = TB.keys;"""
if EMBED or VIRGEN:
    lines = s.split('\n')
    if VIRGEN:
        ix = [i for i, l in enumerate(lines) if l.startswith('const SCENES = {')]
        if len(ix) != 1 or not lines[ix[0]+1].startswith('const PROJS = {') or not lines[ix[0]+2].startswith('const PKEYS ='): sys.exit('No encontré SCENES / PROJS / PKEYS')
        lines[ix[0]:ix[0]+3] = [TB_VIRGEN]
    else:
        ix = [i for i, l in enumerate(lines) if l.startswith('const PKEYS =')]
        if len(ix) != 1: sys.exit('No encontré PKEYS')
        lines[ix[0]] = TB_EMBED
    s = '\n'.join(lines)
    R("let PROJ = (location.hash.match(/p=(w55|w30|all)/)||[])[1] || lsGet('wtp-proj') || 'w55'; if(!['w55','w30','all'].includes(PROJ)) PROJ = 'w55';",
      "let PROJ = decodeURIComponent((location.hash.match(/p=([^&]+)/)||[])[1]||'') || lsGet('wtp-proj') || PKEYS[0]; if(![...PKEYS,'all'].includes(PROJ)) PROJ = PKEYS[0];")
    R("const PSC = ALLV ? (PKEYS.includes(lsGet('wtp-last')) ? lsGet('wtp-last') : 'w55') : PROJ;", "const PSC = ALLV ? (PKEYS.includes(lsGet('wtp-last')) ? lsGet('wtp-last') : PKEYS[0]) : PROJ;")
    R("const SCENE = SCENES[PSC], PM = PROJS[PSC];", "const SCENE = TB.scene(PSC), PM = PROJS[PSC];")
    R("I18N[l].title = 'US.068.4 - PDC Piperacks · WTS 55 m & WWTP 30 m'; Object.assign", "I18N[l].title = TB.title; Object.assign")
    R("Object.assign(I18N.en, {sumH1:'US.068.4 - PDC Piperacks · WTS 55 m & WWTP 30 m', pSel:'Project', pAll:'Summary · 2 projects', sumTitle:'Summary of the 2 projects',", "Object.assign(I18N.en, {sumH1:TB.title, pSel:'Project', pAll:`Summary · ${PKEYS.length} projects`, sumTitle:`Summary of the ${PKEYS.length} projects`,")
    R("Object.assign(I18N.es, {sumH1:'US.068.4 - PDC Piperacks · WTS 55 m & WWTP 30 m', pSel:'Proyecto', pAll:'Resumen · 2 proyectos', sumTitle:'Resumen de los 2 proyectos',", "Object.assign(I18N.es, {sumH1:TB.title, pSel:'Proyecto', pAll:`Resumen · ${PKEYS.length} proyectos`, sumTitle:`Resumen de los ${PKEYS.length} proyectos`,")
    R("const proj = wb.SheetNames.includes('BOM WTS-55M - CONNECTION') ? 'w55' : wb.SheetNames.includes('BOM WWTP-30M - CONNECTION') ? 'w30' : null;", "const proj = (m => m.length===1 ? m[0] : null)(PKEYS.filter(k=>wb.SheetNames.includes(PROJS[k].sheet)));   // misma hoja en varias unidades (un BOM dividido): manda la unidad abierta")
    R("const markKey = s => { const m = String(s||'').toUpperCase().match(/(?:WTS|WWTP)-[A-Z0-9]+/);", "const markKey = s => { const m = String(s||'').toUpperCase().match(/[A-Z0-9]+[-_][A-Z0-9]+/);")
    R('<div class="hrow"><label class="vsel"><span data-i="pSel"></span><select id="projsel"></select></label>',
      '<div class="hrow"><label class="vsel" id="grpbox" hidden><span id="grplb"></span><select id="grpsel"></select></label><label class="vsel"><span data-i="pSel"></span><select id="projsel"></select></label><button class="hload" id="h-new" type="button"><span>＋</span> <span id="h-newt"></span></button>')
    k = s.rindex('</script>') + len('</script>')
    s = s[:k] + '\n' + PANEL + s[k:]
    if VIRGEN: R('<title>PDC Piperacks Weld Map</title>', '<title>Tablero de avance</title>')
    open(out, 'w', encoding='utf8').write(s)
    print('ok', 'virgen' if VIRGEN else 'embed', len(s)); sys.exit(0)

lines = s.split('\n')
ix = [i for i, l in enumerate(lines) if l.startswith('const SCENES = {')]
if len(ix) != 1 or not lines[ix[0]+1].startswith('const PROJS = {') or not lines[ix[0]+2].startswith('const PKEYS ='): sys.exit('No encontré SCENES / PROJS / PKEYS')
lines[ix[0]:ix[0]+3] = [LOADER]
s = '\n'.join(lines)

R("let PROJ = (location.hash.match(/p=(w55|w30|all)/)||[])[1] || lsGet('wtp-proj') || 'w55'; if(!['w55','w30','all'].includes(PROJ)) PROJ = 'w55';",
  "let PROJ = decodeURIComponent((location.hash.match(/p=([^&]+)/)||[])[1]||'') || lsGet('wtp-proj') || PKEYS[0]; if(![...PKEYS,'all'].includes(PROJ)) PROJ = PKEYS[0];")
R("const PSC = ALLV ? (PKEYS.includes(lsGet('wtp-last')) ? lsGet('wtp-last') : 'w55') : PROJ;", "const PSC = ALLV ? (PKEYS.includes(lsGet('wtp-last')) ? lsGet('wtp-last') : PKEYS[0]) : PROJ;")
R("const SCENE = SCENES[PSC], PM = PROJS[PSC];", "const SCENE = TB.scene(PSC), PM = PROJS[PSC];")
R("I18N[l].title = 'US.068.4 - PDC Piperacks · WTS 55 m & WWTP 30 m'; Object.assign", "I18N[l].title = TB.title; Object.assign")
R("Object.assign(I18N.en, {sumH1:'US.068.4 - PDC Piperacks · WTS 55 m & WWTP 30 m', pSel:'Project', pAll:'Summary · 2 projects', sumTitle:'Summary of the 2 projects',",
  "Object.assign(I18N.en, {sumH1:TB.title, pSel:'Unit', pAll:`Summary · ${PKEYS.length} unit${PKEYS.length>1?'s':''}`, sumTitle:`Summary of ${PKEYS.length>1?'the '+PKEYS.length+' units':'the unit'}`,")
R("Object.assign(I18N.es, {sumH1:'US.068.4 - PDC Piperacks · WTS 55 m & WWTP 30 m', pSel:'Proyecto', pAll:'Resumen · 2 proyectos', sumTitle:'Resumen de los 2 proyectos',",
  "Object.assign(I18N.es, {sumH1:TB.title, pSel:'Unidad', pAll:`Resumen · ${PKEYS.length} unidad${PKEYS.length>1?'es':''}`, sumTitle:`Resumen de ${PKEYS.length>1?'las '+PKEYS.length+' unidades':'la unidad'}`,")
R("const proj = wb.SheetNames.includes('BOM WTS-55M - CONNECTION') ? 'w55' : wb.SheetNames.includes('BOM WWTP-30M - CONNECTION') ? 'w30' : null;",
  "const proj = (m => m.length===1 ? m[0] : null)(PKEYS.filter(k=>wb.SheetNames.includes(PROJS[k].sheet)));   // misma hoja en varias unidades (un BOM dividido): manda la unidad abierta")
R("const markKey = s => { const m = String(s||'').toUpperCase().match(/(?:WTS|WWTP)-[A-Z0-9]+/);", "const markKey = s => { const m = String(s||'').toUpperCase().match(/[A-Z0-9]+[-_][A-Z0-9]+/);")
R('<div class="hrow"><label class="vsel"><span data-i="pSel"></span><select id="projsel"></select></label>',
  '<div class="hrow"><label class="vsel" id="grpbox" hidden><span id="grplb"></span><select id="grpsel"></select></label><label class="vsel"><span data-i="pSel"></span><select id="projsel"></select></label><a class="hload" id="h-new" href="index.html" target="_blank" rel="noopener" style="text-decoration:none"><span>＋</span> <span id="h-newt"></span></a>')
R("""  const ps = document.getElementById('projsel'); if(ps){""",
  """  { const gb = document.getElementById('grpbox'), gl = document.getElementById('grplb'), gp = document.getElementById('grpsel'), nw = document.getElementById('h-newt'); if(nw) nw.textContent = LANG==='es' ? 'Nuevo proyecto' : 'New project'; if(gl) gl.textContent = LANG==='es' ? 'Proyecto' : 'Project';
    if(gp){ gp.innerHTML = TB.list.map(p=>`<option value="${p.code.replace(/"/g,'&quot;')}">${p.code}</option>`).join(''); gp.value = TB.cur.code; gb.hidden = TB.list.length<2; gp.onchange = ()=>{ location.href = location.pathname+'?g='+encodeURIComponent(gp.value); }; } }
  const ps = document.getElementById('projsel'); if(ps){""")
R('<title>PDC Piperacks Weld Map</title>', '<title>Tablero de avance</title>')
open(out, 'w', encoding='utf8').write(s)
print('ok', len(s))
