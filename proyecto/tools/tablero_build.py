#!/usr/bin/env python3
"""Arma proyecto/tablero.html desde el HTML del tablero con PR SAP.
Uso: python3 tools/tablero_build.py <tablero_con_PR_SAP.html> <salida.html>
El original no se toca. La copia no lleva proyectos adentro: los lee de ../tablero-datos/.
Cada cambio exige encontrar su texto exacto una sola vez; si el tablero cambia, se detiene."""
import re, sys
src, out = sys.argv[1], sys.argv[2]
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
