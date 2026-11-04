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
  const lists=()=>fetch(TB.base+'index.json?v='+Date.now(),{cache:'no-store'}).then(r=>r.ok?r.json():null).then(j=>{ if(!j) return; const seen=new Set(TB.list.map(p=>p.code)); (j.proyectos||[]).forEach(p=>{ if(!seen.has(p.code)){ TB.list.push({code:p.code}); seen.add(p.code); } }); fill(); }).catch(()=>{});
  $('grpsel').onchange=e=>{ location.href=location.pathname+'?g='+encodeURIComponent(e.target.value); };
  const box=$('np-box'); $('h-new').onclick=()=>{ const f=$('np-frame'); if(!f.src) f.src='https://lndusa-visor.github.io/proyecto/?v='+Date.now(); box.style.display='grid'; }; $('np-close').onclick=()=>{ box.style.display='none'; lists(); };
  $('np-title').textContent='Nuevo proyecto'; fill(); ['lang-en','lang-es'].forEach(i=>{ const b=$(i); if(b) b.addEventListener('click',()=>setTimeout(fill,0)); }); lists(); })();
</script>"""
if EMBED:
    lines = s.split('\n')
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
    R("const proj = wb.SheetNames.includes('BOM WTS-55M - CONNECTION') ? 'w55' : wb.SheetNames.includes('BOM WWTP-30M - CONNECTION') ? 'w30' : null;", "const proj = PKEYS.find(k=>wb.SheetNames.includes(PROJS[k].sheet)) || null;")
    R("const markKey = s => { const m = String(s||'').toUpperCase().match(/(?:WTS|WWTP)-[A-Z0-9]+/);", "const markKey = s => { const m = String(s||'').toUpperCase().match(/[A-Z0-9]+[-_][A-Z0-9]+/);")
    R('<div class="hrow"><label class="vsel"><span data-i="pSel"></span><select id="projsel"></select></label>',
      '<div class="hrow"><label class="vsel" id="grpbox" hidden><span id="grplb"></span><select id="grpsel"></select></label><label class="vsel"><span data-i="pSel"></span><select id="projsel"></select></label><button class="hload" id="h-new" type="button"><span>＋</span> <span id="h-newt"></span></button>')
    k = s.rindex('</script>') + len('</script>')
    s = s[:k] + '\n' + PANEL + s[k:]
    open(out, 'w', encoding='utf8').write(s)
    print('ok embed', len(s)); sys.exit(0)

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
  "const proj = PKEYS.find(k=>wb.SheetNames.includes(PROJS[k].sheet)) || null;")
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
