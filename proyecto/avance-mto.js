/* Excel de avance — hojas de material: TAKEOFF BASE, MTO SUMMARY, PROFILE CUTTING PLAN, PLATE NESTING.
   Salen de las filas de partes de la hoja BOM … CONNECTION; se usa después de avance-gen.js (AV.sheetXml).
   Reglas comprobadas celda por celda con el Excel de 55 m:
   - Barras: largo descendente (a igual largo, orden del BOM), cada pieza a la barra más justa donde quepa; cada corte gasta kerf; recorte al final de barra.
     Las barras con el mismo acomodo se listan juntas, primero los acomodos que más se repiten.
   - Placas: MaxRects por área descendente, Best Short Side Fit, primera lámina donde quepa, sin reordenar láminas. */
(function(G){ 'use strict';
  const AV=G.AV=G.AV||{};
  const nat=(a,b)=>String(a).localeCompare(String(b),undefined,{numeric:true});
  // redondeo con empate al par (como el redondeo del MTO original): 39.8125 → 39.812
  function rd(x,n){ const f=10**n, y=x*f, fl=Math.floor(y), d=y-fl; let k; if(Math.abs(d-0.5)<1e-9) k=fl%2===0?fl:fl+1; else k=Math.round(y); return +(k/f).toFixed(n); }
  // 6 cifras significativas: 186.0625 → 186.062 · 59.5625 → 59.5625 · 6 → 6
  function g6(x){ const e=x===0?0:Math.floor(Math.log10(Math.abs(x))); return String(rd(x,Math.max(0,5-e))); }
  function frac(s){ s=String(s||'').trim(); const m=s.match(/^(?:(\d+)[ -])?(\d+)\/(\d+)$/); if(m) return (+(m[1]||0))+m[2]/m[3]; const n=parseFloat(s); return isNaN(n)?null:n; }
  function ftIn(s){ const m=String(s||'').match(/(\d+)'\s*-?\s*(\d+(?:\s+\d+\/\d+)?|\d+\/\d+)?/); if(!m) return null; return +m[1]*12+(m[2]?frac(m[2]):0); }
  function kind(desc){ const d=String(desc).toUpperCase(); if(/^PL/.test(d)) return 'PLATE'; if(/^L\d/.test(d)) return 'ANGLE'; return 'PROFILE'; }
  function plateDim(desc){ const m=String(desc).match(/^PL\s*([\d\/ ]+)"?\s*[Xx]\s*([\d\/ ]+)"?/i); return m?{t:frac(m[1].trim()),w:frac(m[2].trim()),tok:`PL${m[1].trim()}"`}:null; }
  // lb/ft nominal de la barra: W8X13 → 13; ángulos de tabla AISC; si no, del peso del BOM
  const ANGLES={'L2X2X1/4':3.19,'L2X2X3/8':4.7,'L2-1/2X2-1/2X1/4':4.1,'L3X3X1/4':4.9,'L3X3X3/8':7.2,'L3X3X1/2':9.4,'L4X4X1/4':6.6,'L4X4X3/8':9.8,'L4X4X1/2':12.8,'L5X5X3/8':12.3,'L5X5X1/2':16.2,'L6X6X3/8':14.9,'L6X6X1/2':19.6};
  function lbft(desc,fallback){ const d=String(desc).toUpperCase().replace(/\s/g,''); const m=d.match(/^(?:W|HP|M|S|C|MC)\d+(?:\.\d+)?X(\d+(?:\.\d+)?)$/); if(m) return +m[1]; if(ANGLES[d]) return ANGLES[d]; return fallback; }

  /* partes del BOM: lee las filas P y H de la hoja (columna AP), generada o de un Excel existente */
  AV.mtoParts=function(S){ const out=[]; let asm=null;
    let asmRow=null; S.rows.forEach(R=>{ const c={}; R.cells.forEach(x=>c[x.c]=x); const t=c.AP&&c.AP.v, v=k=>c[k]?c[k].v:null; if(t==='A'){ asm=v('B'); asmRow=R.r; }
      if(t==='P'||t==='H') out.push({asm,asmRow,row:R.r,mark:v('B'),qty:v('D'),desc:v('E'),grade:v('F'),lenTxt:v('G'),len:ftIn(v('G')),wt:v('H'),twt:c.I?c.I.v:null}); });
    return out; };

  function bestFit(pieces,S,kerf,trim){ const bars=[]; pieces.forEach(p=>{ const need=p.len+kerf; let b=null; for(const x of bars) if(x.free>=need-1e-9&&(!b||x.free<b.free)) b=x; if(!b){ b={free:S-trim,items:[],used:0}; bars.push(b); } b.items.push(p); b.free-=need; b.used+=need; }); return bars; }
  function barText(b,asmOf){ const o=[]; b.items.forEach(p=>{ const x=o.find(i=>i.m===p.mark&&i.len===p.len); if(x) x.n++; else o.push({m:p.mark,len:p.len,n:1}); });
    const str=(x,y)=>x<y?-1:x>y?1:0; o.sort((x,y)=>y.len-x.len||str(x.m,y.m)); return o.map(i=>`${i.m} ${g6(i.len)}"${i.n>1?` ×${i.n}`:''} (${asmOf[i.m].join(' / ')})`).join(' · '); }
  function groupBars(bars){ const count={}; bars.forEach(b=>count[b.txt]=(count[b.txt]||0)+1); const seen=new Map(); bars.forEach(b=>{ if(!seen.has(b.txt)) seen.set(b.txt,[]); seen.get(b.txt).push(b); });
    return [...seen.values()].map((l,i)=>({l,i})).sort((a,b)=>b.l.length-a.l.length||a.i-b.i).flatMap(x=>x.l); }

  /* placas: MaxRects, misma geometría que el acomodo de láminas */
  function prune(fr,R){ const e=1e-9, out=[]; for(const f of fr){ if(R.x>=f.x+f.w-e||R.x+R.w<=f.x+e||R.y>=f.y+f.h-e||R.y+R.h<=f.y+e){ out.push(f); continue; }
      if(R.x>f.x+e) out.push({x:f.x,y:f.y,w:R.x-f.x,h:f.h}); if(R.x+R.w<f.x+f.w-e) out.push({x:R.x+R.w,y:f.y,w:f.x+f.w-R.x-R.w,h:f.h}); if(R.y>f.y+e) out.push({x:f.x,y:f.y,w:f.w,h:R.y-f.y}); if(R.y+R.h<f.y+f.h-e) out.push({x:f.x,y:R.y+R.h,w:f.w,h:f.y+f.h-R.y-R.h}); }
    return out.filter((a,i)=>!out.some((b,j)=>j!==i&&a.x>=b.x-e&&a.y>=b.y-e&&a.x+a.w<=b.x+b.w+e&&a.y+a.h<=b.y+b.h+e&&(j<i||a.w!==b.w||a.h!==b.h||a.x!==b.x||a.y!==b.y))); }
  function bssf(free,w,h){ let best=null; for(const f of free) for(const [ww,hh] of [[w,h],[h,w]]){ if(ww>f.w+1e-9||hh>f.h+1e-9) continue; const s1=Math.min(f.w-ww,f.h-hh), s2=Math.max(f.w-ww,f.h-hh); if(!best||s1<best.s1-1e-9||(Math.abs(s1-best.s1)<1e-9&&s2<best.s2-1e-9)) best={x:f.x,y:f.y,w:ww,h:hh,s1,s2}; } return best; }
  function nest(pieces,W,H,gap){ const bins=[], over=[]; [...pieces].sort((x,y)=>y.a*y.b-x.a*x.b).forEach(p=>{ const w=p.a+gap,h=p.b+gap; let bb=null,f=null;
      for(const b of bins){ f=bssf(b.free,w,h); if(f){ bb=b; break; } }
      if(!bb){ const nb={free:[{x:0,y:0,w:W+gap,h:H+gap}],items:[],area:0}; f=bssf(nb.free,w,h); if(!f){ over.push(p); return; } bins.push(nb); bb=nb; }
      bb.free=prune(bb.free,f); bb.items.push(p); bb.area+=p.a*p.b; });
    return {bins,over}; }

  const cells=(spec,vals)=>spec.split(' ').map(t=>{ const m=t.match(/^([A-Z]+)(\d+)$/); return Object.assign({c:m[1],s:+m[2]},(vals&&vals[m[1]])||{}); });
  const txt=v=>({v}), num=v=>({v}), fx=(f,v)=>({f,v});
  const RQ=s=>`'${s}'`;
  const TK='TAKEOFF BASE', PL='PROFILE CUTTING PLAN', NS='PLATE NESTING';
  const noteRow=(r,s,v,ht,ch)=>({r,ht:ht||'15',ch:!!ch,cells:[{c:'A',s,v}]});
  const sumV=(A,k)=>A.reduce((t,x)=>t+(x[k]||0),0);

  /* opciones: parts, stock {W8X13:40,…}, sheet [ancho,largo] ft, kerf, trim, gap, edge (pulgadas),
     proj, bomDoc, boltDoc, bomSheet, bomTotalRow, mtoFile (opcional), spares, bolts [{desc,grade,field,shop}], anchors {desc,n} */
  AV.mtoSheets=function(o){ const kerf=o.kerf??0.125, trim=o.trim??0.5, gap=o.gap??0.25, edge=o.edge??0.5, sheet=o.sheet||[4,8], rev=o.rev||'Rev 0';
    const P=o.parts.map(p=>{ const cat=kind(p.desc), d=cat==='PLATE'?plateDim(p.desc):null; return Object.assign({},p,{cat,d,lenFt:rd(p.len/12,4)}); });
    const n=P.length, last=4+n, tot=last+1;

    /* ---- TAKEOFF BASE ---- */
    const tk=[noteRow(1,60,'TAKEOFF BASE','18',true),noteRow(2,61,'One row per BOM part line, with numeric quantities. Feeds the VENDOR REQUISITION and the MTO SUMMARY.')];
    const HT=['ASSEMBLY','PART MARK','TOTAL QTY.','DESCRIPTION','CATEGORY','MATERIAL','UNIT LENGTH (ft)','WIDTH (in)','THICKNESS (in)','TOTAL LENGTH (ft)','TOTAL AREA (ft²)','UNIT WEIGHT (lb)','TOTAL WEIGHT (lb)'], LET='ABCDEFGHIJKLM'.split('');
    tk.push({r:4,ht:'31.5',ch:true,cells:LET.map((c,i)=>({c,s:5,v:HT[i]}))});
    P.forEach((p,i)=>{ const r=5+i, w=p.d?p.d.w:null, area=p.cat==='PLATE'?p.qty*p.lenFt*w/12:0; p.tl=p.qty*p.lenFt; p.ta=area; p.tw=p.twt!=null?p.twt:p.qty*p.wt;
      const v={A:txt(p.asm),B:txt(p.mark),C:num(p.qty),D:txt(p.desc),E:txt(p.cat),F:txt(p.grade),G:num(p.lenFt),J:fx(`C${r}*G${r}`,p.tl),K:fx(`IF(E${r}="PLATE",C${r}*G${r}*H${r}/12,0)`,area),L:num(p.wt),M:fx(`C${r}*L${r}`,p.tw)};
      if(p.d){ v.H=num(w); v.I=num(p.d.t); }
      tk.push({r,ht:'15',ch:false,cells:cells('A10 B10 C133 D10 E10 F10 G154 H154 I154 J14 K14 L31 M31',v)}); });
    const tq=sumV(P,'qty'), tl=sumV(P,'tl'), ta=sumV(P,'ta'), tw=sumV(P,'tw');
    tk.push({r:tot,ht:'15',ch:false,cells:cells('A99 B69 C114 D69 E69 F69 G69 H69 I69 J51 K51 L69 M51',{A:txt('TOTAL'),C:fx(`SUM(C5:C${last})`,tq),J:fx(`SUM(J5:J${last})`,tl),K:fx(`SUM(K5:K${last})`,ta),M:fx(`SUM(M5:M${last})`,tw)})});

    /* ---- agrupar perfiles y placas ---- */
    const catRank={PROFILE:0,ANGLE:1}, profMap=new Map(), plateMap=new Map();
    P.forEach(p=>{ if(p.cat==='PLATE'){ const k=p.d.t+'|'+p.grade; if(!plateMap.has(k)) plateMap.set(k,{tok:p.d.tok,t:p.d.t,grade:p.grade,rows:[]}); plateMap.get(k).rows.push(p); }
      else { const k=p.desc+'|'+p.grade; if(!profMap.has(k)) profMap.set(k,{desc:p.desc,grade:p.grade,cat:p.cat,rows:[]}); profMap.get(k).rows.push(p); } });
    const profs=[...profMap.values()].sort((a,b)=>catRank[a.cat]-catRank[b.cat]||nat(a.desc,b.desc));
    const plates=[...plateMap.values()].sort((a,b)=>a.t-b.t||nat(a.grade,b.grade));

    /* ---- PROFILE CUTTING PLAN ---- */
    const kerfTxt=kerf===0.125?'1/8':String(kerf), trimTxt=trim===0.5?'1/2':String(trim);
    const tail=o.mtoFile?` Same plan as file ${o.mtoFile}.`:'';
    const pl=[noteRow(1,60,'PROFILE CUTTING PLAN (MTO)','18',true),noteRow(2,61,`Bar-by-bar cutting plan per profile and grade (first-fit decreasing, saw kerf ${kerfTxt}", ${trimTxt}" end trim). Stock: ${profs.map(g=>`${g.desc} ${o.stock[g.desc]} ft`).join(', ')}.${tail}`)];
    const PH=['PROFILE','GRADE','BAR No.','STOCK (ft)','No. CUTS','CUTS (part mark · length in · assembly)','USED (ft)','DROP (ft)','YIELD','BAR WEIGHT (lb)'];
    pl.push({r:4,ht:'31.5',ch:true,cells:'ABCDEFGHIJ'.split('').map((c,i)=>({c,s:5,v:PH[i]}))});
    let r=5; const over=[];
    profs.forEach((g,gi)=>{ const S=o.stock[g.desc], lf=lbft(g.desc,rd(g.rows[0].wt/(g.rows[0].len/12),1)); g.S=S; g.lf=lf; const asmOf={}, pieces=[];
      g.rows.forEach(p=>{ (asmOf[p.mark]=asmOf[p.mark]||[]); if(!asmOf[p.mark].includes(p.asm)) asmOf[p.mark].push(p.asm); });
      const merged=new Map(); g.rows.forEach(p=>{ const k=p.mark+'|'+p.len; if(!merged.has(k)) merged.set(k,{mark:p.mark,len:p.len,qty:0}); merged.get(k).qty+=p.qty; });
      [...merged.values()].forEach(m=>{ for(let q=0;q<m.qty;q++) pieces.push({mark:m.mark,len:m.len}); });
      const sorted=pieces.map((p,i)=>({p,i})).sort((a,b)=>b.p.len-a.p.len||a.i-b.i).map(x=>x.p);
      sorted.filter(p=>p.len+kerf>S*12-trim+1e-9).forEach(p=>over.push(p));
      const bars=bestFit(sorted.filter(p=>p.len+kerf<=S*12-trim+1e-9),S*12,kerf,trim); bars.forEach(b=>{ b.txt=barText(b,asmOf); }); const list=groupBars(bars); const f=r;
      list.forEach((b,i)=>{ const G=rd(b.used/12,3); b.G=G; pl.push({r,ht:b.txt.length>100?'22.35':'15',ch:false,cells:cells('A130 B10 C29 D133 E29 F160 G161 H161 I74 J75',{A:txt(g.desc),B:txt(g.grade),C:num(i+1),D:num(S),E:num(b.items.length),F:txt(b.txt),G:num(G),H:fx(`D${r}-G${r}`,S-G),I:fx(`IFERROR(G${r}/D${r},0)`,G/S),J:fx(`D${r}*${lf}`,S*lf)})}); r++; });
      const l=r-1, gs=list.reduce((t,b)=>t+b.G,0), cut=list.reduce((t,b)=>t+b.items.length,0), st=list.length*S;
      g.bars=list.length; g.subRow=r;
      pl.push({r,ht:'15',ch:false,cells:cells('A99 B69 C54 D69 E66 F69 G52 H52 I136 J55',{A:txt(`Subtotal ${g.desc} ${g.grade}`),C:fx(`COUNT(C${f}:C${l})`,list.length),E:fx(`SUM(E${f}:E${l})`,cut),G:fx(`SUM(G${f}:G${l})`,gs),H:fx(`SUM(H${f}:H${l})`,st-gs),I:fx(`IFERROR(G${r}/SUMPRODUCT(D${f}:D${l}),0)`,gs/st),J:fx(`SUM(J${f}:J${l})`,st*lf)})}); r++;
      if(gi<profs.length-1) r++; });

    /* ---- PLATE NESTING ---- */
    const sw=Math.min(...sheet), sl=Math.max(...sheet), sTxt=`${sw}' x ${sl}'`, sA=sw*sl;
    const ns=[noteRow(1,60,'PLATE NESTING (MTO)','18',true),noteRow(2,61,`Sheet-by-sheet nesting per thickness and grade (MaxRects, parts may rotate, 1/4" cutting allowance, 1/2" edge margin). Sheet ${sTxt} (best yield for this job).${tail}`)];
    const NH=['THICKNESS','GRADE','SHEET No.','SHEET SIZE','SHEET AREA (ft²)','No. PIECES','PIECES (part mark ×qty)','NET AREA (ft²)','YIELD','SHEET WEIGHT (lb)'];
    ns.push({r:4,ht:'31.5',ch:true,cells:'ABCDEFGHIJ'.split('').map((c,i)=>({c,s:5,v:NH[i]}))});
    r=5; const overP=[];
    plates.forEach((g,gi)=>{ const merged=new Map(), pcs=[]; g.rows.forEach(p=>{ const k=p.mark+'|'+p.len+'|'+p.d.w; if(!merged.has(k)) merged.set(k,{mark:p.mark,len:p.len,w:p.d.w,qty:0}); merged.get(k).qty+=p.qty; });
      [...merged.values()].forEach(m=>{ const a=Math.max(m.len,m.w), b=Math.min(m.len,m.w); for(let q=0;q<m.qty;q++) pcs.push({mark:m.mark,a,b}); });
      const N=nest(pcs,sl*12-2*edge,sw*12-2*edge,gap); overP.push(...N.over); const f=r, wt=sA*40.84*g.t; let net=0, count=0;
      N.bins.forEach((b,i)=>{ const c={}; b.items.forEach(p=>c[p.mark]=(c[p.mark]||0)+1); const t=Object.keys(c).sort().map(k=>`${k} ×${c[k]}`).join(' · '), H=rd(b.area/144,2); net+=H; count+=b.items.length;
        ns.push({r,ht:t.length>80?'22.35':'15',ch:false,cells:cells('A130 B10 C29 D29 E133 F29 G160 H161 I74 J75',{A:txt(g.tok),B:txt(g.grade),C:num(i+1),D:txt(sTxt),E:num(sA),F:num(b.items.length),G:txt(t),H:num(H),I:fx(`IFERROR(H${r}/E${r},0)`,H/sA),J:fx(`E${r}*40.84*${g.t}`,wt)})}); r++; });
      const l=r-1, nb=N.bins.length; g.sheets=nb; g.subRow=r;
      ns.push({r,ht:'15',ch:false,cells:cells('A99 B69 C54 D69 E55 F66 G69 H52 I136 J55',{A:txt(`Subtotal ${g.tok} ${g.grade}`),C:fx(`COUNT(C${f}:C${l})`,nb),E:fx(`SUM(E${f}:E${l})`,nb*sA),F:fx(`SUM(F${f}:F${l})`,count),H:fx(`SUM(H${f}:H${l})`,net),I:fx(`IFERROR(H${r}/E${r},0)`,net/(nb*sA)),J:fx(`SUM(J${f}:J${l})`,nb*wt)})}); r++;
      if(gi<plates.length-1) r++; });

    /* ---- MTO SUMMARY ---- */
    const tb=c=>`${RQ(TK)}!$${c}$5:$${c}$${last}`;
    const bomDoc=o.bomDoc||'', boltDoc=o.boltDoc||'', spares=o.spares??0.05;
    const mt=[noteRow(1,60,'MATERIAL TAKE-OFF (MTO) SUMMARY','18',true),noteRow(2,61,`${o.proj} · from ${bomDoc} ${rev} (TAKEOFF BASE), ${PL}, ${NS} and Bolt Summary ${boltDoc} ${rev}. Net = BOM; gross = stock to buy.`)];
    const MH=['CATEGORY','DESCRIPTION','GRADE','PIECES','NET LENGTH (ft) / AREA (ft²)','NET WEIGHT (lb)','STOCK','No. BARS / SHEETS','GROSS LENGTH (ft) / AREA (ft²)','GROSS WEIGHT (lb)','YIELD','DROP / SCRAP (lb)','UNIT'];
    mt.push({r:4,ht:'31.5',ch:true,cells:LET.map((c,i)=>({c,s:5,v:MH[i]})).concat({c:'N',s:5,v:'LONGEST PIECE (PROJECT)'})});
    const longP=g=>(o.longest||{})[`${g.desc}|${g.grade}`], longL=g=>(o.longest||{})[`PL${g.t}|${g.grade}`], lg=x=>x?txt(x.txt):null;   // dato, no fórmula: columna nueva a la derecha, no mueve ninguna
    r=5; const MS='A63 B10 C10 D75 E155 F100 G29 H108 I38 J75 K76 L75 M29 N29', M=[];
    profs.forEach(g=>{ const q=sumV(g.rows,'qty'), nl=sumV(g.rows,'tl'), nw=sumV(g.rows,'tw'), gl=g.bars*g.S, gw=gl*g.lf; M.push({q,nw,gw});
      const sf=c=>`SUMIFS(${tb(c)},${tb('D')},B${r},${tb('F')},C${r})`;
      mt.push({r,ht:'15',ch:false,cells:cells(MS,{A:txt(g.cat),B:txt(g.desc),C:txt(g.grade),D:fx(sf('C'),q),E:fx(sf('J'),nl),F:fx(sf('M'),nw),G:txt(`${g.S} ft`),H:fx(`${RQ(PL)}!C${g.subRow}`,g.bars),I:fx(`H${r}*${g.S}`,gl),J:fx(`I${r}*${g.lf}`,gw),K:fx(`IFERROR(E${r}/I${r},0)`,nl/gl),L:fx(`J${r}-F${r}`,gw-nw),M:txt('ft'),N:lg(longP(g))})}); r++; });
    plates.forEach(g=>{ const q=sumV(g.rows,'qty'), na=sumV(g.rows,'ta'), nw=sumV(g.rows,'tw'), ga=g.sheets*sA, gw=ga*40.84*g.t; M.push({q,nw,gw});
      const sf=(c,pl)=>`SUMIFS(${tb(c)},${tb('I')},${g.t},${tb('F')},C${r}${pl?`,${tb('E')},"PLATE"`:''})`;
      mt.push({r,ht:'15',ch:false,cells:cells(MS,{A:txt('PLATE'),B:txt(g.tok),C:txt(g.grade),D:fx(sf('C',1),q),E:fx(sf('K'),na),F:fx(sf('M',1),nw),G:txt(sTxt),H:fx(`${RQ(NS)}!C${g.subRow}`,g.sheets),I:fx(`H${r}*${sA}`,ga),J:fx(`I${r}*40.84*${g.t}`,gw),K:fx(`IFERROR(E${r}/I${r},0)`,na/ga),L:fx(`J${r}-F${r}`,gw-nw),M:txt('ft²'),N:lg(longL(g))})}); r++; });
    const a=5, b=r-1, tq2=sumV(M,'q'), tn=sumV(M,'nw'), tg=sumV(M,'gw'), T=r;
    mt.push({r,ht:'15',ch:false,cells:cells('A99 B69 C69 D55 E69 F55 G69 H69 I69 J55 K53 L55 M69',{A:txt('TOTAL STEEL'),D:fx(`SUM(D${a}:D${b})`,tq2),F:fx(`SUM(F${a}:F${b})`,tn),J:fx(`SUM(J${a}:J${b})`,tg),K:fx(`IFERROR(F${r}/J${r},0)`,tn/tg),L:fx(`SUM(L${a}:L${b})`,tg-tn)})}); r++;
    mt.push({r,ht:'15',ch:false,cells:cells('A156 F100 G33',{A:txt('Check vs BOM total weight'),F:fx(`${RQ(o.bomSheet)}!I${o.bomTotalRow}`,o.bomTotal??tn),G:fx(`IF(ABS(F${r}-F${T})<1,"OK","CHECK")`,'OK')})}); r+=2;
    mt.push({r,ht:'19.5',ch:true,cells:LET.map((c,i)=>({c,s:5,v:{A:'BOLTING',B:'DESCRIPTION',C:'GRADE',D:'FIELD',E:'SHOP',F:'TOTAL',G:'NUTS',H:'WASHERS',I:'SPARES %',J:'TO BUY (bolts)',M:'UNIT'}[c]??null}))}); r++;
    const bf=r; let F=0,Sh=0,Buy=0; (o.bolts||[]).forEach(x=>{ const t=x.field+x.shop, buy=Math.ceil(t*(1+spares)-1e-9); F+=x.field; Sh+=x.shop; Buy+=buy;
      mt.push({r,ht:'15',ch:false,cells:cells('A63 B10 C10 D157 E157 F123 G75 H75 I158 J123 M29',{A:txt('BOLT'),B:txt(x.desc),C:txt(x.grade),D:num(x.field),E:num(x.shop),F:fx(`D${r}+E${r}`,t),G:fx(`F${r}`,t),H:fx(`F${r}`,t),I:fx(`'VENDOR REQUISITION'!$C$7`,spares),J:fx(`ROUNDUP(F${r}*(1+I${r}),0)`,buy),M:txt('ea')})}); r++; });
    const bl=r-1, bt=F+Sh, sumF=c=>`SUM(${c}${bf}:${c}${bl})`;
    mt.push({r,ht:'15',ch:false,cells:cells('A99 B69 C69 D55 E55 F55 G55 H55 I69 J55 K69 L69 M69',{A:txt('TOTAL BOLTS'),D:fx(sumF('D'),F),E:fx(sumF('E'),Sh),F:fx(sumF('F'),bt),G:fx(sumF('G'),bt),H:fx(sumF('H'),bt),J:fx(sumF('J'),Buy)})}); r++;
    const dia=o.boltDia||'3/4"', nf=x=>Number(x).toLocaleString('en-US');
    mt.push(noteRow(r,159,`Check: Bolt Summary nuts ${dia} = ${nf(o.nuts??bt)} · washers ${dia} HD = ${nf(o.washers??bt)}`)); r+=2;
    mt.push(noteRow(r,85,`Profiles and angles: net length from BOM part lengths; gross from the ${PL} (FFD, kerf ${kerfTxt}").`)); r++;
    mt.push(noteRow(r,85,`Plates: net area = BOM width x length (rectangle); gross from ${NS} (MaxRects, 1/4" allowance). Plate weight factor 40.84 lb/ft² per inch.`)); r++;
    mt.push(noteRow(r,85,`Bolting per Bolt Summary ${boltDoc} ${rev} (0% added).`+(o.anchors?` ${o.anchors.desc} adhesive anchors (${o.anchors.n}) are listed in VENDOR REQUISITION for scope confirmation.`:'')));

    return {takeoff:{rows:tk,lastCol:'M',filter:`$A$4:$M$${last}`,subst:[[/<autoFilter ref="[^"]*">/,`<autoFilter ref="A4:M${last}">`]]},
      mto:{rows:mt,lastCol:'M'},plan:{rows:pl,lastCol:'J'},nesting:{rows:ns,lastCol:'J'},over,overPlates:overP,last,tot,profs,plates,takeoffParts:P}; };

  /* ---- VENDOR REQUISITION y PURCHASE ORDERS: mismo diseño y fórmulas; cambian los datos del proyecto.
     PURCHASE ORDERS sale sin órdenes: se llena a mano. Los campos manuales de la requisición quedan vacíos. ---- */
  const LONGNAME=d=>{ const u=String(d).toUpperCase(); return /^W\d/.test(u)?'Wide flange beam':/^HSS/.test(u)?'Hollow structural section':/^(C|MC)\d/.test(u)?'Channel':/^L\d/.test(u)?'Equal-leg angle':/^PIPE/.test(u)?'Pipe':'Profile'; };
  AV.vendorSheets=function(o,R){ const VR='VENDOR REQUISITION', PO='PURCHASE ORDERS', rev=o.rev||'Rev 0', last=R.last, tb=c=>`'TAKEOFF BASE'!$${c}$5:$${c}$${last}`, po=c=>`'${PO}'!$${c}$8:$${c}$207`;
    const nf=x=>Number(x).toLocaleString('en-US'), pct=x=>Math.round(x*100), rows=[], merges=[], mats={}, sheet=o.sheet||[4,8], sw=Math.min(...sheet), sl=Math.max(...sheet), sTxt=`${sw}' x ${sl}'`, sA=sw*sl, spares=o.spares??0.05;
    rows.push(noteRow(1,77,'MATERIAL REQUISITION TO VENDOR','18',true),noteRow(2,61,`${o.proj} – LINDSAYCA, Inc. – Ref. BOM ${o.bomDocFull||o.bomDoc} (${rev}) / Bolt Summary ${o.boltDocFull||o.boltDoc} (${rev})${o.ptpDocFull?` / Point-to-Point Bolt List ${o.ptpDocFull} (${rev})`:''}`));
    const H4=(r,a,g,sI)=>({r,ht:'15',ch:false,cells:cells(`A70 B70 C102 D102 E102 G70 H70 I${sI} J${sI} K${sI}`,{A:txt(a),G:txt(g)})});
    rows.push(H4(4,'Requisition No.','Date',103),H4(5,'Requested by','Vendor',102),H4(6,'Delivery location','Required on-site date',103));
    rows.push({r:7,ht:'15',ch:false,cells:cells('A70 B70 C104 D59 N59',{A:txt('% bolt spares'),C:num(spares),D:txt('Bolt Summary states 0% added; adjust to purchasing policy.'),N:txt('Columns N–U: purchasing. SAP CODE and PR No. are typed here; vendors and POs are typed in PURCHASE ORDERS (one row per PO line, several vendors per item allowed).')})});
    const VH=['ITEM','CATEGORY','DESCRIPTION','MATERIAL / SPEC','UNIT','NET QTY.','STOCK SIZE','No. BARS / SHEETS','QTY. TO BUY','APPROX. WEIGHT (lb)','REMARKS','DATE RECEIVED IN SHOP','MATERIAL KEY (3D control)','SAP CODE','PR No.','VENDOR (first PO)','No. OF POs','QTY. ORDERED','PENDING TO ORDER','QTY. RECEIVED','PURCHASE STATUS'], VL='ABCDEFGHIJKLMNOPQRSTU'.split('');
    rows.push({r:9,ht:'31.5',ch:true,cells:VL.map((c,i)=>({c,s:5,v:VH[i]}))});
    ['A4:B4','C4:E4','G4:H4','I4:K4','A5:B5','C5:E5','G5:H5','I5:K5','A6:B6','C6:E6','G6:H6','I6:K6','A7:B7'].forEach(m=>merges.push(m));
    let r=10, item=0; const SEC='A105 B105 C105 D105 E105 F105 G105 H105 I105 J105 K105 L105 M105', SUB='A69 B69 C99 D69 E99 F51 G69 H69 I51 J114 K69 L69 M69';
    const sec=t=>{ rows.push({r,ht:'15',ch:false,cells:cells(SEC,{A:txt(t)})}); merges.push(`A${r}:M${r}`); r++; };
    // columnas de compra, iguales en todos los ítems
    const buyCols=(x,q)=>({P:fx(`IFERROR(INDEX(${po('G')},MATCH(A${x},${po('A')},0)),"")`,''),Q:fx(`COUNTIF(${po('A')},A${x})`,0),R:fx(`SUMIFS(${po('K')},${po('A')},A${x})`,0),S:fx(`MAX(0,I${x}-R${x})`,q),T:fx(`SUMIFS(${po('M')},${po('A')},A${x})`,0),
      U:fx(`IF(I${x}=0,"NOT REQUIRED",IF(AND(T${x}>=I${x},T${x}>0),"RECEIVED",IF(T${x}>0,"PARTIALLY RECEIVED",IF(R${x}>=I${x},"ORDERED",IF(R${x}>0,"PARTIALLY ORDERED",IF(O${x}<>"","PR ISSUED","TO REQUEST"))))))`,q===0?'NOT REQUIRED':'TO REQUEST')});
    const subRow=(lab,unit,f,l,net,buy,wt,all)=>{ const x=r; rows.push({r,ht:'15',ch:false,cells:cells(all?'A69 B69 C99 D69 E99 F114 G69 H69 I114 J69 K69 L69 M69':SUB,Object.assign({C:txt(lab),E:txt(unit),F:fx(`SUM(F${f}:F${l})`,net),I:fx(`SUM(I${f}:I${l})`,buy)},all?{}:{J:fx(`SUM(J${f}:J${l})`,wt)}))}); r+=2; return x; };
    const MATS='A29 B10 C10 D10 E10 F106 G107 H108 I109 J110 K111 L37 M112 N34 O21 P113 Q64 R75 S75 T75 U33', tots=[]; let letter=0; const nextL=()=>String.fromCharCode(65+letter++);
    const profGroup=(list,title,lab)=>{ if(!list.length) return; sec(`${nextL()}. ${title} (linear feet)`); const f=r; let n=0,b=0,w=0;
      list.forEach(g=>{ const net=sumV(g.rows,'tl'), buy=g.bars*g.S, wt=buy*g.lf; item++; n+=net;b+=buy;w+=wt; mats[`${g.desc}|${g.grade}`]=r;
        rows.push({r,ht:'15',ch:false,cells:cells(MATS,Object.assign({A:num(item),B:txt(g.cat),C:txt(`${LONGNAME(g.desc)} ${g.desc}`),D:txt(g.grade),E:txt('ft'),F:fx(`SUMIFS(${tb('J')},${tb('D')},"${g.desc}",${tb('F')},D${r})`,net),G:num(g.S),H:fx(`'PROFILE CUTTING PLAN'!C${g.subRow}`,g.bars),I:fx(`G${r}*H${r}`,buy),J:fx(`I${r}*${g.lf}`,wt),K:txt(`${g.bars} bars of ${g.S} ft · yield ${pct(net/buy)}%`),M:txt(`${g.desc}|${g.grade}`)},buyCols(r,buy)))}); r++; });
      tots.push({x:subRow(lab,'ft',f,r-1,n,b,w),w}); };
    const W=R.profs.filter(g=>g.cat==='PROFILE'), allW=W.every(g=>/^W\d/i.test(g.desc));
    profGroup(W,allW?'STRUCTURAL PROFILES – W':'STRUCTURAL PROFILES','Subtotal profiles'); profGroup(R.profs.filter(g=>g.cat==='ANGLE'),'ANGLES','Subtotal angles');
    if(R.plates.length){ sec(`${nextL()}. PLATES (square feet) – ${sTxt} sheet = ${sA} ft²`); const f=r; let n=0,b=0,w=0;
      R.plates.forEach(g=>{ const net=sumV(g.rows,'ta'), buy=g.sheets*sA, wt=buy*40.84*g.t, pcs=sumV(g.rows,'qty'), key=`PL${Number.isInteger(g.t)?g.t.toFixed(1):g.t}|${g.grade}`; item++; n+=net;b+=buy;w+=wt; mats[key]=r;
        rows.push({r,ht:'15',ch:false,cells:cells(MATS.replace('G107','G115'),Object.assign({A:num(item),B:txt('PLATE'),C:txt(`Plate t=${g.tok.replace(/^PL/,'')}`),D:txt(g.grade),E:txt('ft²'),F:fx(`SUMIFS(${tb('K')},${tb('I')},${g.t},${tb('F')},D${r})`,net),G:txt(sTxt),H:fx(`'PLATE NESTING'!C${g.subRow}`,g.sheets),I:fx(`H${r}*${sA}`,buy),J:fx(`I${r}*40.84*${g.t}`,wt),K:txt(`${pcs} pieces · yield ${pct(net/buy)}%`),M:txt(key)},buyCols(r,buy)))}); r++; });
      tots.push({x:subRow('Subtotal plates','ft²',f,r-1,n,b,w),w}); }
    sec(`${nextL()}. GRATING (square feet)`); { const f=r; item++; rows.push({r,ht:'25.5',ch:true,cells:cells('A29 B10 C10 D10 E10 F116 G10 H10 I10 J10 K111 L117 M10 N34 O21 P113 Q64 R75 S75 T75 U33',Object.assign({A:num(item),B:txt('GRATING'),C:txt('Welded steel grating (type TBD)'),E:txt('ft²'),F:num(0),I:fx(`F${r}`,0),J:num(0),K:txt(`No grating in BOM ${rev}. Complete if the drawing is issued.`)},buyCols(r,0)))}); r++; tots.push({x:subRow('Subtotal grating','ft²',f,r-1,0,0,0),w:0}); }
    const bolts=o.bolts||[], dia=o.boltDia||'3/4"', totB=sumV(bolts,'field')+sumV(bolts,'shop'), bspec=(bolts[0]&&bolts[0].grade)||'A325N';
    sec(`${nextL()}. BOLTS (units) – ${bspec}, each bolt with 1 heavy hex nut and 1 hardened washer (Bolt Summary ${o.boltDoc}, field + shop)`);
    { const f=r, BR='A29 B10 C10 D10 E10 F118 G10 H10 I119 J10 K10 L37 M10 N34 O21 P113 Q64 R75 S75 T75 U33', up=x=>Math.ceil(x*(1+spares)-1e-9); let net=0,buy=0; const bf=r;
      bolts.forEach(x=>{ const t=x.field+x.shop, q=up(t); item++; net+=t; buy+=q; rows.push({r,ht:'15',ch:false,cells:cells(BR,Object.assign({A:num(item),B:txt('BOLT'),C:txt(`Heavy hex structural bolt ${String(x.desc).replace(/\s*heavy hex$/i,'')}`),D:txt(x.grade),E:txt('ea'),F:num(t),I:fx(`ROUNDUP(F${r}*(1+$C$7),0)`,q),K:txt(`Field ${nf(x.field)} + Shop ${nf(x.shop)} (Bolt Summary)`)},buyCols(r,q)))}); r++; });
      const bl=r-1, link=Array.from({length:bl-bf+1},(_,i)=>`F${bf+i}`).join('+');
      [['NUT',`Heavy hex nut Ø${dia}`,o.nutSpec||'A563 DH',o.nuts??totB],['WASHER',`Hardened washer Ø${dia}`,o.washerSpec||'F436',o.washers??totB]].forEach(([c,d,sp,sum])=>{ const q=up(totB); item++; net+=totB; buy+=q; rows.push({r,ht:'15',ch:false,cells:cells(BR.replace('F118','F110'),Object.assign({A:num(item),B:txt(c),C:txt(d),D:txt(sp),E:txt('ea'),F:fx(link,totB),I:fx(`ROUNDUP(F${r}*(1+$C$7),0)`,q),K:txt(`1 per bolt (Bolt Summary: ${nf(sum)})`)},buyCols(r,q)))}); r++; });
      if(o.anchors){ const a=o.anchors, n=a.n; item++; net+=n; buy+=n; const model=a.model!=null?` · 3D model: ${a.model}${a.basePlates?` (${a.basePlates} base plates × ${a.perPlate})`:''}`:'';
        rows.push({r,ht:'38.25',ch:true,cells:cells(BR.replace('K10','K111'),Object.assign({A:num(item),B:txt('ANCHOR'),C:txt(`Adhesive anchor ${a.desc}`),D:txt(a.spec||'HILTI HIT (rod + adhesive TBD)'),E:txt('ea'),F:num(n),I:fx(`F${r}`,n),K:txt(`Bolt Summary ${o.boltDoc}: ${n}${model}. Confirm adhesive, rod grade and supply scope (civil / site).`)},buyCols(r,n)))}); r++; }
      subRow('Subtotal bolting (bolts + nuts + washers + anchors)','ea',f,r-1,net,buy,0,true); }
    const buyW=tots.reduce((t,x)=>t+x.w,0), netW=sumV(R.takeoffParts,'tw'), kg=Math.round(buyW*0.45359237).toLocaleString('en-US');
    rows.push({r,ht:'15',ch:false,cells:cells('C99 J114 K99',{C:txt('TOTAL STEEL WEIGHT TO BUY (lb)'),J:fx(tots.map(t=>`J${t.x}`).join('+'),buyW),K:fx(`TEXT(J${r}*0.45359237,"#,##0")&" kg"`,`${kg} kg`)})}); const rb=r; r++;
    rows.push({r,ht:'15',ch:false,cells:cells('C99 J114 K69',{C:txt('Check: BOM net weight (lb)'),J:fx(`'TAKEOFF BASE'!M${R.tot}`,netW),K:fx(`TEXT(J${rb}/J${r}-1,"0.0%")&" purchase allowance over net"`,`${((buyW/netW-1)*100).toFixed(1)}% purchase allowance over net`)})}); const end=r; r+=2;
    rows.push(noteRow(r,120,'Notes:')); r++;
    [`1. Net quantities calculated from the TAKEOFF BASE sheet (${o.bomDoc} ${rev}). Quantity to buy = bars/sheets from the PROFILE CUTTING PLAN and PLATE NESTING sheets x stock size.`,'2. Profiles in linear feet (ft); plates in square feet (ft²); bolting in units (ea).','3. Yellow cells: data to fill in or editable parameters. Blue: input data; green: link to another sheet.','4. Mill test reports (MTR) required for all steel. Coating not stated in the BOM (COATING column empty). Bolting finish per drawing (confirm galvanized).','5. DATE RECEIVED IN SHOP: when each material arrives, enter the date. The 3D CONTROL sheet uses it to move pieces from "NOT YET ARRIVED" to "MATERIAL IN SHOP".'].forEach(t=>{ rows.push(noteRow(r,121,t)); r++; });

    /* PURCHASE ORDERS */
    const vr=c=>`'${VR}'!$${c}$10:$${c}$${end}`, m=x=>`MATCH($A${x},${vr('A')},0)`, pr=[noteRow(1,60,'PURCHASE ORDERS – VENDOR / PR / PO REGISTER','18',true),noteRow(2,61,'One row per purchase order line. Type the VENDOR REQUISITION ITEM: description, unit and SAP code are filled in. Several rows for the same item = several vendors / POs. Shown in the Weld Map dashboard (button Material).')];
    pr.push({r:4,ht:'15',ch:false,cells:cells('A70 B122 C72 D122 E72 F122 G123 I124',{A:txt('Totals'),B:txt('PO lines'),C:fx('COUNTA(I8:I207)',0),D:txt('Vendors'),E:fx('SUMPRODUCT((G8:G207<>"")/COUNTIF(G8:G207,G8:G207&""))',0),F:txt('Approx. lb ordered'),G:fx('SUM(Q8:Q207)',0),I:txt('Example (not real data): ITEM 4 · VENDOR Steel Supplier A · PR 1000123 · PO 4500012345 · 10/20/2026 · 360 ft; second row ITEM 4 · VENDOR Steel Supplier B · PO 4500012399 · 360 ft → 2 vendors for item 4.')})});
    const PH=['ITEM (VENDOR REQ.)','CATEGORY','DESCRIPTION','MATERIAL / SPEC','UNIT','SAP CODE','VENDOR','PR No.','PO No.','PO DATE','QTY. ORDERED','PROMISED DATE','QTY. RECEIVED','DATE RECEIVED','REMARKS','QTY. TO BUY (item)','APPROX. lb ORDERED'];
    pr.push({r:7,ht:'33.75',ch:true,cells:'ABCDEFGHIJKLMNOPQ'.split('').map((c,i)=>({c,s:5,v:PH[i]}))});
    for(let x=8;x<=207;x++){ const look=c=>fx(`IF($A${x}="","",IFERROR(INDEX(${vr(c)},${m(x)})&"",""))`,'');
      pr.push({r:x,ht:'15',ch:false,cells:cells('A34 B113 C113 D113 E125 F125 G21 H125 I21 J37 K126 L37 M126 N37 O127 P100 Q75',{B:look('B'),C:look('C'),D:look('D'),E:look('E'),F:look('N'),H:look('O'),P:fx(`IF($A${x}="","",IFERROR(INDEX(${vr('I')},${m(x)}),""))`,''),Q:fx(`IF(OR($A${x}="",K${x}=""),"",IFERROR(K${x}/INDEX(${vr('I')},${m(x)})*INDEX(${vr('J')},${m(x)}),""))`,'')})}); }
    const itemRows=rows.filter(R=>R.cells.some(c=>c.c==='U')&&R.r>9).map(R=>R.r), dv=AV.ranges('L',itemRows).join(' ');
    return {vendor:{rows,lastCol:'U',merges,subst:[[/(<dataValidation [^>]*?sqref=")[^"]*(")/,`$1${dv}$2`]]},po:{rows:pr,lastCol:'Q'},mats,itemCount:item}; };

  /* pernos para las hojas de material, desde lo que ya leyó la página del Bolt Summary */
  AV.boltInput=function(B,model){ const out={bolts:[],nuts:0,washers:0,anchors:null,boltDia:'3/4"'}; if(!B||!B.rows) return out; const m=new Map();
    const lenTxt=l=>{ const w=Math.floor(l+1e-9), r=Math.round((l-w)*16); if(!r) return `${w}"`; let n=r,d=16; while(n%2===0){ n/=2; d/=2; } return `${w?w+' ':''}${n}/${d}"`; };
    B.rows.forEach(b=>{ if(b.sec==='Nuts'){ out.nuts+=b.qty; return; } if(b.sec==='Washers'){ out.washers+=b.qty; return; } if(b.sec!=='Field'&&b.sec!=='Shop') return;
      const dia=String(b.dia||'').replace(/"/g,''); if(/HILTI|ANCHOR/i.test(b.grade)){ const d=`${String(b.grade).replace(/\s+$/,'')} Ø${dia}" x ${lenTxt(b.len)}`; out.anchors={desc:d,n:(out.anchors?out.anchors.n:0)+b.qty}; return; }
      const k=dia+'|'+b.len+'|'+b.grade; if(!m.has(k)) m.set(k,{dia,len:b.len,grade:b.grade,field:0,shop:0}); const x=m.get(k); if(b.sec==='Field') x.field+=b.qty; else x.shop+=b.qty; });
    out.bolts=[...m.values()].sort((a,b)=>a.len-b.len).map(x=>({desc:`Ø${x.dia}" x ${lenTxt(x.len)} heavy hex`,grade:x.grade,field:x.field,shop:x.shop})); if(m.size) out.boltDia=[...m.values()][0].dia+'"';
    if(out.anchors&&model) Object.assign(out.anchors,model); const t=(B.meta&&B.meta.title||'').match(/[A-Z0-9-]*BOM-\d+/); out.boltDocFull=t?t[0]:''; out.boltDoc=(out.boltDocFull.match(/BOM-\d+/)||[''])[0]; return out; };

  AV.shopBoltText=B=>{ const x=(B&&B.bolts||[]).find(y=>y.shop>0); return x?`${x.grade} ${String(x.desc).replace(/ heavy hex$/,'')}`:null; };
  // anclajes del modelo 3D: se suman de la columna Q de la hoja BOM («Anchored … ×4» por pieza)
  AV.anchorModel=function(S){ let pl=0,tot=0,per=null,mixed=false; S.rows.forEach(R=>{ const c={}; R.cells.forEach(x=>c[x.c]=x); const t=c.AP&&c.AP.v; if(t!=='A') return; const m=String(c.Q&&c.Q.v||'').match(/anchored [^×]*×(\d+)/i); if(!m) return; const q=+c.D.v, n=+m[1]; pl+=q; tot+=q*n; if(per==null) per=n; else if(per!==n) mixed=true; });
    return tot?{model:tot,basePlates:mixed?null:pl,perPlate:mixed?null:per}:null; };

  /* ---- CUTTING TRACKING: una línea por parte del BOM, ligada a la hoja BOM; la tabla diaria siempre tiene 120 días (filas 11–130) ---- */
  const serial=d=>{ const t=d instanceof Date?d:new Date(String(d).slice(0,10)+'T00:00:00Z'); return Math.round((Date.UTC(t.getUTCFullYear(),t.getUTCMonth(),t.getUTCDate())-Date.UTC(1899,11,30))/864e5); };
  AV.cuttingSheet=function(o,parts){ const n=parts.length, f=11, l=10+n, BS=`'${o.bomSheet}'!`, DAYS=130, st=serial(o.startDate), rows=[];
    const rng=c=>`$${c}$${f}:$${c}$${l}`, QD=[['I','J'],['K','L'],['M','N'],['O','P'],['Q','R'],['S','T']], isP=p=>kind(p.desc)==='PLATE';
    const tot=parts.reduce((a,p)=>{ const t=isP(p)?1:0; a.q[t]+=p.qty; a.l[t]++; a.w+=(p.twt!=null?p.twt:p.qty*p.wt); return a; },{q:[0,0],l:[0,0],w:0});
    rows.push(noteRow(1,60,'CUTTING TRACKING – PROFILES AND PLATES','18',true),noteRow(2,61,'One line per BOM part (read from the BOM sheet). Each day pieces are cut, enter the QTY. and DATE in the next free entry (up to 6 per line): the ✔ is set automatically when the sum reaches the required quantity.'));
    const NOTE='H62 I62 J62 K62 L62 M62 N62 O62 P62', H4={H:txt('% CUT = parts cut ÷ parts required (same as X9). Not all parts are cut on the same day: each line has 6 cut entries (QTY. and DATE). TOTAL QTY. CUT = sum of the entries; LAST CUT = most recent date. In red: quantity without date or total above the required quantity.')};
    rows.push({r:4,ht:'30',ch:true,cells:cells(`A7 B7 C7 D7 E7 F7 ${NOTE}`,Object.assign({A:txt('TYPE'),B:txt('LINES'),C:txt('LINES CUT ✔'),D:txt('PARTS REQUIRED'),E:txt('PARTS CUT'),F:txt('% CUT (parts)')},H4))});
    ['PROFILE','PLATE'].forEach((t,i)=>{ const r=5+i; rows.push({r,ht:'15',ch:false,cells:cells(`A63 B64 C64 D64 E64 F65 ${NOTE}`,{A:txt(t),B:fx(`COUNTIF(${rng('D')},A${r})`,tot.l[i]),C:fx(`COUNTIFS(${rng('D')},A${r},${rng('W')},"☑")`,0),D:fx(`SUMIFS(${rng('H')},${rng('D')},A${r})`,tot.q[i]),E:fx(`SUMIFS(${rng('U')},${rng('D')},A${r})`,0),F:fx(`IFERROR(E${r}/D${r},0)`,0)})}); });
    rows.push({r:7,ht:'15',ch:false,cells:cells(`A66 B54 C54 D54 E54 F53 ${NOTE}`,{A:txt('TOTAL'),B:fx('SUM(B5:B6)',n),C:fx('SUM(C5:C6)',0),D:fx('SUM(D5:D6)',tot.q[0]+tot.q[1]),E:fx('SUM(E5:E6)',0),F:fx('IFERROR(E7/D7,0)',0)})});
    rows.push({r:8,ht:'15',ch:false,cells:cells('AB67',{AB:txt('CUTS PER DAY')})});
    const colDates=QD.map(q=>rng(q[1])).join(','), S9=c=>fx(`SUM(${c}11:${c}666)`,c==='H'?tot.q[0]+tot.q[1]:c==='Y'?tot.w:0);
    rows.push({r:9,ht:'15',ch:false,cells:cells('G68 H55 I55 J69 K55 L69 M55 N69 O55 P69 Q55 R69 S55 T69 U55 V69 W54 X53 Y55 Z55 AB70 AC71',{G:txt('TOTAL'),H:S9('H'),I:S9('I'),K:S9('K'),M:S9('M'),O:S9('O'),Q:S9('Q'),S:S9('S'),U:S9('U'),W:fx('COUNTIF(W11:W666,"☑")',0),X:fx('IFERROR(U9/H9,0)',0),Y:S9('Y'),Z:S9('Z'),AB:txt('First day'),AC:fx(`IF(COUNT(${colDates})=0,'DAILY REPORT'!$C$7,MIN(${colDates}))`,st)})});
    const HH=['ITEM','ASSEMBLY','PART','TYPE','DESCRIPTION','GRADE','LENGTH','QTY. REQUIRED','QTY. CUT 1','CUT DATE 1','QTY. CUT 2','CUT DATE 2','QTY. CUT 3','CUT DATE 3','QTY. CUT 4','CUT DATE 4','QTY. CUT 5','CUT DATE 5','QTY. CUT 6','CUT DATE 6','TOTAL QTY. CUT','LAST CUT','CUT ✔','% CUT','WEIGHT (lb)','WEIGHT CUT (lb)'], HC='ABCDEFGHIJKLMNOPQRSTUVWXYZ'.split('');
    rows.push({r:10,ht:'31.5',ch:true,cells:HC.map((c,i)=>({c,s:7,v:HH[i]})).concat([['AB','DATE'],['AC','PARTS CUT (day)'],['AD','PARTS CUT (cum.)'],['AE','% CUT (parts)'],['AF','WEIGHT CUT DAY (lb)']].map(([c,v])=>({c,s:7,v})))});
    const PS='A64 B33 C29 D29 E10 F10 G10 H64 I34 J37 K34 L37 M34 N37 O34 P37 Q34 R37 S34 T37 U72 V73 W35 X74 Y38 Z38', DS='AB73 AC75 AD75 AE76 AF75';
    for(let i=0;i<Math.max(n,DAYS-10);i++){ const r=f+i, p=parts[i], v={};
      if(p){ const q=p.row, W=p.twt!=null?p.twt:p.qty*p.wt; Object.assign(v,{A:num(i+1),B:fx(`${BS}B${p.asmRow}`,p.asm),C:fx(`${BS}B${q}`,p.mark),D:txt(isP(p)?'PLATE':'PROFILE'),E:fx(`${BS}E${q}`,p.desc),F:fx(`${BS}F${q}`,p.grade),G:fx(`${BS}G${q}`,p.lenTxt),H:fx(`${BS}D${q}`,p.qty),
          U:fx(`N(I${r})+N(K${r})+N(M${r})+N(O${r})+N(Q${r})+N(S${r})`,0),V:fx(`IF(COUNT(J${r},L${r},N${r},P${r},R${r},T${r})=0,"",MAX(J${r},L${r},N${r},P${r},R${r},T${r}))`,''),W:fx(`IF(AND(U${r}>0,U${r}>=H${r}),"☑","☐")`,'☐'),X:fx(`IFERROR(MIN(1,U${r}/H${r}),0)`,0),Y:fx(`${BS}I${q}`,W),Z:fx(`Y${r}*X${r}`,0)}); }
      if(r<=DAYS) Object.assign(v,{AB:fx(r===f?'$AC$9':`AB${r-1}+1`,st+i),AC:{f:`IF(AB${r}<=TODAY(),${QD.map(([a,b])=>`SUMIFS(${rng(a)},${rng(b)},AB${r})`).join('+')},"")`},AD:{f:r===f?`IF(AB${r}<=TODAY(),AC${r},"")`:`IF(AB${r}<=TODAY(),N(AD${r-1})+AC${r},"")`},AE:{f:`IF(AB${r}<=TODAY(),IFERROR(AD${r}/$H$9,0),"")`},
        AF:{f:`IF(AB${r}<=TODAY(),${QD.map(([a,b])=>`SUMPRODUCT((${rng(b)}=AB${r})*${rng(a)}*${rng('Y')}/${rng('H')})`).join('+')},"")`}});
      rows.push({r,ht:'15',ch:false,cells:cells([p?PS:'',r<=DAYS?DS:''].filter(Boolean).join(' '),v)}); }
    return {rows,lastCol:'AF',filter:`$A$10:$Z$${l}`,merges:['H4:P7'],subst:[[/<autoFilter ref="[^"]*">/,`<autoFilter ref="A10:Z${l}">`],[/sqref="[^"]*"/g,m=>m.replace(/([A-Z]+11:[A-Z]+)\d+/g,`$1${l}`)]]}; };

  /* ---- WELDING PROGRESS: una fila por pieza compuesta (nodo); tabla diaria fija de 120 días (filas 21–140) ---- */
  const elemName=t=>String(t||'').toUpperCase().replace(/_/g,' ').replace(/^VERTICAL /,'VERT. ').replace(/^HORIZONTAL /,'HORIZ. ');
  AV.weldingSheet=function(o,M,bom){ const B={}; bom.rows.forEach(R=>{ const c={}; R.cells.forEach(x=>c[x.c]=x); const t=c.AP&&c.AP.v; if(t==='A') B[c.B.v]={r:R.r,E:c.E.v,P:[]}; else if(t==='P'){ const m=Object.values(B).pop(); m.P.push({r:R.r,mark:c.B.v,aq:c.AQ?c.AQ.v:0,at:c.AT?c.AT.v:null}); } });
    const BS=`'${o.bomSheet}'!`, st=serial(o.startDate), N=[]; let node=0;
    M.filter(x=>x.comp).forEach(x=>{ const b=B[x.mark], f=b.P[0].r, l=b.P[b.P.length-1].r, G=rd(x.wp.reduce((t,p)=>t+(+p.weld||0),0),1), fs=[...new Set(b.P.map(p=>p.at).filter(Boolean))].sort((a,c)=>{ const v=s=>{ const m=String(s).match(/(\d+)\/(\d+)/); return m?m[1]/m[2]:0; }; return v(a)-v(c); });
      const seq=x.seq||[], ord=[...x.wp].sort((p,q)=>seq.indexOf(p.k)-seq.indexOf(q.k)), E=ord.map(p=>`${p.mark} ×${p.qty/x.qty} ${String(p.role).replace(/ stub$/,'')}`).join('; '), F=x.wp.reduce((t,p)=>t+p.n,0), I=rd(x.twt/x.qty-(x.loose||0)/x.qty,1);
      for(let i=0;i<x.qty;i++) N.push({mark:x.mark,node:++node,el:elemName(b.E),E,F,G,H:fs.join(' / '),I,A:b.r,f,l,q:x.qty,i}); });
    const n=N.length, a=21, z=20+n, T=z+1, rg=c=>`$${c}$${a}:$${c}$${z}`, sum=k=>N.reduce((t,x)=>t+x[k],0), W=sum('G'), rows=[], ALL=rg('D');
    rows.push(noteRow(1,60,'WELDING PROGRESS TRACKING','18',true),noteRow(2,61,'One row per composite piece with shop welds (NODE). Weld count and length derived from the Tekla model (IFC contact faces); all welds are shop fillet welds.'));
    const els=[...new Set(N.map(x=>x.el))], fills=[...new Set(N.map(x=>x.H))].sort((p,q)=>p.length-q.length), E0={BEAM:0}; 
    const elRow=r=>{ const e=els[r-6]; if(!e) return {}; const g=N.filter(x=>x.el===e), wi=g.reduce((t,x)=>t+x.G,0); return {L:txt(e),M:fx(`COUNTIF(${ALL},L${r})`,g.length),N:fx(`COUNTIFS(${ALL},L${r},${rg('O')},"RELEASED")`,0),O:fx(`SUMIFS(${rg('G')},${ALL},L${r})`,wi),P:fx(`IF(O${r}=0,0,SUMPRODUCT((${ALL}=L${r})*${rg('P')}*${rg('G')})/O${r})`,0),Q:fx(`SUMIFS(${rg('S')},${ALL},L${r})`,0),R:fx(`SUMIFS(${rg('T')},${ALL},L${r})`,wi)}; };
    const S=(r,sp,v)=>rows.push({r,ht:'15',ch:false,cells:cells(sp,v)});
    S(4,'A70 B70 C70 D70 E70 G70 H70 I70 J70 L70 M70 N70 O70 P70',{A:txt('WEIGHT OF EACH STAGE IN % PROGRESS'),G:txt('SUMMARY'),L:txt('BY ELEMENT')});
    S(5,'A10 B10 C10 D10 E128 G10 H10 I10 J72 L129 M129 N129 O129 P129 Q129 R129',{A:txt('Fit-up'),E:num(0.2),G:txt('Composite pieces'),J:fx(`COUNTA(${rg('B')})`,n),L:txt('ELEMENT'),M:txt('PIECES'),N:txt('RELEASED'),O:txt('WELD IN.'),P:txt('PROGRESS'),Q:txt('IN. WELDED'),R:txt('IN. PENDING')});
    S(6,'A10 B10 C10 D10 E128 G10 H10 I10 J72 L10 M64 N64 O75 P76 Q75 R75',Object.assign({A:txt('Welding'),E:num(0.6),G:txt('Fitted (100 %)'),J:fx(`COUNTIF(${rg('V')},">=0.999")`,0)},elRow(6)));
    S(7,'A10 B10 C10 D10 E128 G10 H10 I10 J72 L10 M64 N64 O75 P76 Q75 R75',Object.assign({A:txt('VT inspection'),E:num(0.2),G:txt('Welded (100 %)'),J:fx(`COUNTIF(${rg('R')},">=0.999")`,0)},elRow(7)));
    S(8,'A130 B130 C130 D130 E131 G10 H10 I10 J72 L10 M64 N64 O75 P76 Q75 R75',Object.assign({A:txt('Total'),E:fx('SUM(E5:E7)',1),G:txt('Released (VT accepted)'),J:fx(`COUNTIF(${rg('O')},"RELEASED")`,0)},elRow(8)));
    S(9,'G10 H10 I10 J72 L10 M64 N64 O75 P76 Q75 R75',Object.assign({G:txt('To repair'),J:fx(`COUNTIF(${rg('O')},"REPAIR*")`,0)},elRow(9)));
    S(10,'G10 H10 I10 J123 L10 M64 N64 O75 P76 Q75 R75',Object.assign({G:txt('Total weld inches (weld map)'),J:fx(`SUM(${rg('G')})`,W)},elRow(10)));
    S(11,'G10 H10 I10 J123',{G:txt('Inches welded'),J:fx(`SUM(${rg('S')})`,0)});
    S(12,'G10 H10 I10 J123 L70 M70 N70 O70 P70 Q70 R70',{G:txt('Inches pending'),J:fx(`SUM(${rg('T')})`,W),L:txt('BY FILLET SIZE')});
    S(13,'G130 H130 I130 J53 L129 M129 N129 O129 P129',{G:txt('% INCHES WELDED'),J:fx(`IF(SUM(${rg('G')})=0,0,SUM(${rg('S')})/SUM(${rg('G')}))`,0),L:txt('FILLET'),M:txt('PIECES'),N:txt('WELD IN.'),O:txt('IN. WELDED'),P:txt('IN. PENDING')});
    const fRow=r=>{ const h=fills[r-14]; if(!h) return {}; const g=N.filter(x=>x.H===h), wi=g.reduce((t,x)=>t+x.G,0); return {L:txt(h),M:fx(`COUNTIF(${rg('H')},L${r})`,g.length),N:fx(`SUMIFS(${rg('G')},${rg('H')},L${r})`,wi),O:fx(`SUMIFS(${rg('S')},${rg('H')},L${r})`,0),P:fx(`SUMIFS(${rg('T')},${rg('H')},L${r})`,wi)}; };
    S(14,'G130 H130 I130 J53 L29 M64 N75 O75 P75',Object.assign({G:txt('PROGRESS (weighted by inches)'),J:fx(`IF(SUM(${rg('G')})=0,0,SUMPRODUCT(${rg('P')},${rg('G')})/SUM(${rg('G')}))`,0)},fRow(14)));
    S(15,'G10 H10 I10 J65 L29 M64 N75 O75 P75',Object.assign({G:txt('Progress by pieces'),J:fx(`IF(COUNTA(${rg('B')})=0,0,AVERAGE(${rg('P')}))`,0)},fRow(15)));
    rows.push(noteRow(17,85,'FIT-UP DATE, WELD DATE, PIECE DIE STAMP, % FIT / % WELDED, INCHES and STATUS come automatically from the daily records of the BOM sheet (green). Only VT INSPECTION (list) and VT DATE are typed here (yellow).'));
    S(18,'A124 W67',{A:txt('Example (not real data): Fit-up date 11/16/2026 · Weld date 11/17/2026 · Die stamp C1-1 · VT ACCEPTED · VT date 11/18/2026  →  STATUS = RELEASED, PROGRESS = 100%.'),W:txt('WELD INCHES PER DAY (BOM: welded qty × line inches, on the welder date)')});
    const HT=['ITEM','MARK','NODE','ELEMENT','PARTS TO WELD (mark · qty · function)','No. WELDS','WELD LENGTH (in)','FILLET','WEIGHT (lb)','FIT-UP DATE (BOM)','WELD DATE (BOM, 100 %)','PIECE DIE STAMP (BOM)','VT INSPECTION','VT DATE','STATUS','% PROGRESS','REMARKS','% WELDED (BOM, by weight)','INCHES WELDED','INCHES PENDING','LAST WELD DATE (BOM)','% FIT (BOM)'];
    rows.push({r:20,ht:'31.5',ch:true,cells:'ABCDEFGHIJKLMNOPQRSTUV'.split('').map((c,i)=>({c,s:5,v:HT[i]})).concat(['DATE','IN. WELDED (day)','IN. WELDED (cum.)','% OF TOTAL IN.'].map((v,i)=>({c:'WXYZ'[i],s:7,v})))});
    const PS='A29 B130 C29 D10 E132 F133 G31 H29 I31 J134 K134 L125 M21 N37 O29 P74 Q127 R74 S135 T38 U73 V74', TS='A69 B99 C69 D69 E69 F55 G55 H69 I55 J69 K69 L69 M69 N69 O69 P136 Q69 R136 S55 T55 U69', DS='W73 X75 Y75 Z76', DAYS=140, lastB=bom.last, br=c=>`${BS}$${c}$4:$${c}$${lastB}`;
    for(let r=a;r<=Math.max(DAYS,T);r++){ const v={}; const x=N[r-a];
      if(x){ const Y=`${BS}Y${x.f}:Y${x.l}`, U=`${BS}U${x.f}:U${x.l}`, up=(c)=>`MAX(0,MIN(1,${BS}${c}${x.A}*${x.q}-${x.i}))`;
        Object.assign(v,{A:num(x.node),B:txt(x.mark),C:num(x.node),D:txt(x.el),E:txt(x.E),F:num(x.F),G:num(x.G),H:txt(x.H),I:num(x.I),J:fx(`IF(V${r}>0,IF(COUNT(${U})=0,"",MIN(${U})),"")`,''),K:fx(`IF(R${r}>=0.999,IF(COUNT(${Y})=0,IF(COUNT(${U})=0,"",MAX(${U})),MAX(${Y})),"")`,''),L:fx(`IF(${BS}Z${x.A}<>"",${BS}Z${x.A},"")`,''),
          O:fx(`IF(OR(M${r}="ACCEPTED",M${r}="ACCEPTED AFTER REPAIR"),"RELEASED",IF(M${r}="REJECTED","REPAIR",IF(R${r}>=0.999,"WELDED",IF(R${r}>0,"IN WELDING",IF(V${r}>=0.999,"FITTED",IF(V${r}>0,"IN FIT-UP","PENDING"))))))`,'PENDING'),P:fx(`$E$5*V${r}+$E$6*R${r}+IF(O${r}="RELEASED",$E$7,0)`,0),R:fx(up('AG'),0),
          S:fx(`G${r}*IF(O${r}="RELEASED",1,IFERROR(MAX(0,MIN(1,SUM(${BS}AR${x.f}:AR${x.l})/SUM(${BS}AQ${x.f}:AQ${x.l})*${x.q}-${x.i})),0))`,0),T:fx(`G${r}-S${r}`,x.G),U:fx(`IF(COUNT(${Y})=0,"",MAX(${Y}))`,''),V:fx(up('AF'),0)}); }
      if(r===T) Object.assign(v,{B:txt('TOTAL'),F:fx(`SUM(F${a}:F${z})`,sum('F')),G:fx(`SUM(G${a}:G${z})`,W),I:fx(`SUM(I${a}:I${z})`,sum('I')),P:fx('J14',0),R:fx(`IF(G${T}=0,0,S${T}/G${T})`,0),S:fx(`SUM(S${a}:S${z})`,0),T:fx(`SUM(T${a}:T${z})`,W)});
      if(r<=DAYS) Object.assign(v,{W:fx(r===a?"'DAILY REPORT'!$C$7":`W${r-1}+1`,st+r-a),X:{f:`IF(W${r}<=TODAY(),SUMIFS(${br('AR')},${br('Y')},W${r},${br('AP')},"<>A"),"")`},Y:{f:r===a?`IF(W${r}<=TODAY(),X${r},"")`:`IF(W${r}<=TODAY(),N(Y${r-1})+X${r},"")`},Z:{f:`IF(W${r}<=TODAY(),IFERROR(Y${r}/$G$${T},0),"")`}});
      rows.push({r,ht:'15',ch:false,cells:cells([x?PS:r===T?TS:'',r<=DAYS?DS:''].filter(Boolean).join(' '),v)}); }
    return {rows,lastCol:'Z',filter:`$A$20:$V$${z}`,merges:['A4:E4','A5:D5','A6:D6','A7:D7','A8:D8','G10:I10','G11:I11','G12:I12','G13:I13','G14:I14','G15:I15','G4:J4','G5:I5','G6:I6','G7:I7','G8:I8','G9:I9','L12:R12','L4:P4'],subst:[[/<autoFilter ref="[^"]*">/,`<autoFilter ref="A20:V${z}">`],[/sqref="[^"]*"/g,m=>m.replace(/([A-Z]+21:[A-Z]+)\d+/g,`$1${z}`)]],nodes:N}; };

  /* ---- PAINTING PROGRESS: una fila por pieza de montaje; pesos y áreas por pieza del BOM original (sin las partes sueltas) ---- */
  AV.paintingSheet=function(o,M,bom,W){ const B={}; bom.rows.forEach(R=>{ const c={}; R.cells.forEach(x=>c[x.c]=x); if(c.AP&&c.AP.v==='A') B[c.B.v]={E:c.E.v}; });
    const nodes={}, P=[]; W.nodes.forEach(n=>{ (nodes[n.mark]=nodes[n.mark]||[]).push(n.node); });
    M.forEach(x=>{ const el=elemName(B[x.mark].E), comp=x.comp, a=+((x.sa-(x.looseSa||0))/x.qty).toFixed(4), w=+((x.twt-(x.looseT||0))/x.qty).toFixed(4);  // sin redondear: la suma debe dar el total del BOM
      for(let i=0;i<x.qty;i++) P.push({mark:x.mark,c:`${i+1}/${x.qty}`,node:comp?nodes[x.mark][i]:null,el,comp,a,w});
      (x.lp||[]).forEach(l=>{ for(let i=0;i<l.qty;i++) P.push({mark:l.mark,parent:x.mark,c:`${i+1}/${l.qty}`,node:null,el:elemName(l.type),comp:false,a:+(l.sa/l.qty).toFixed(4),w:+(l.twt/l.qty).toFixed(4)}); }); });
    const n=P.length, a=25, z=24+n, T=z+1, rg=c=>`$${c}$${a}:$${c}$${z}`, rows=[], S=(r,ht,sp,v,ch)=>rows.push({r,ht,ch:!!ch,cells:cells(sp,v)}), sum=k=>P.reduce((t,x)=>t+x[k],0), TA=sum('a'), wr=W.nodes.length+20, nComp=P.filter(x=>x.comp).length;
    rows.push(noteRow(1,60,'PAINTING PROGRESS TRACKING – COATING SYSTEM CS1','18',true),noteRow(2,61,`One row per erection piece (${n}). Paint area per BOM (Surface Area). Composite pieces are not painted until their welding is RELEASED.`));
    const els=[...new Set(P.map(x=>x.el))].sort(), elv=r=>{ const e=els[r-6]; if(!e) return {}; const g=P.filter(x=>x.el===e), ar=g.reduce((t,x)=>t+x.a,0); return {M:txt(e),N:fx(`COUNTIF(${rg('E')},M${r})`,g.length),O:fx(`SUMIFS(${rg('G')},${rg('E')},M${r})`,ar),P:fx(`IF(O${r}=0,0,SUMIFS(${rg('G')},${rg('E')},M${r},${rg('J')},">0")/O${r})`,0)}; };
    S(4,'15','A70 B70 C70 D70 E70 F70 H70 I70 J70 K70 M70 N70 O70 P70',{A:txt('COATING SYSTEM CS1 (spec. tables 9.1 / 13.1) – reference'),H:txt('SUMMARY'),M:txt('BY ELEMENT')});
    S(5,'51','A137 B137 C137 D137 E137 F137 H10 I10 J10 K72 M129 N129 O129 P129',{A:txt('COAT'),B:txt('GENERIC TYPE'),C:txt('SHERWIN WILLIAMS'),D:txt('PPG (alternate)'),E:txt('DFT min. (µm)'),F:txt('DFT min. (mils)'),H:txt('Pieces to paint'),K:fx(`COUNTA(${rg('B')})`,n),M:txt('ELEMENT'),N:txt('PIECES'),O:txt('AREA (ft²)'),P:txt('PAINTED (area)')},true);
    const C='A130 B10 C10 D132 E64 F138 H10 I10 J10', E=' M10 N64 O75 P76', co=(r,k,v)=>S(r,'15',C+` K${k}`+E,Object.assign(v,elv(r)));
    co(6,123,{A:txt('Surface preparation'),B:txt('Abrasive blast cleaning per spec.'),C:txt('—'),D:txt('—'),H:txt('Total area (ft²)'),K:fx(`SUM(${rg('G')})`,TA)});
    co(7,72,{A:txt('Primer coat'),B:txt('Inorganic zinc silicate'),C:txt('Zinc Clad II Plus'),D:txt('Dimetcote 9'),E:num(75),F:num(3),H:txt('Waiting for weld release'),K:fx(`COUNTIF(${rg('N')},"WAITING*")`,nComp)});
    co(8,72,{A:txt('Mid coat'),B:txt('Hi solids epoxy'),C:txt('Macropoxy 646-100'),D:txt('Amerlock 2 VOC'),E:num(125),F:num(4.9),H:txt('Pieces painted'),K:fx(`COUNT(${rg('J')})`,0)});
    co(9,123,{A:txt('Topcoat'),B:txt('Polysiloxane'),C:txt('Sher-Loxane 800 / Hi Solids'),D:txt('PSX 700'),E:num(125),F:num(4.9),H:txt('Area painted (ft²)'),K:fx(`SUMIFS(${rg('G')},${rg('J')},">0")`,0)});
    co(10,72,{A:txt('Total system'),B:txt('Total DFT ≥ 325 µm (checked at final inspection)'),C:txt('—'),D:txt('—'),E:num(325),F:num(12.8),H:txt('READY TO SHIP (inspection accepted)'),K:fx(`COUNT(${rg('M')})`,0)});
    S(11,'15','H10 I10 J10 K72'+E,Object.assign({H:txt('IN TRANSIT (with packing list)'),K:fx(`COUNTIF(${rg('N')},"IN TRANSIT")`,0)},elv(11)));
    S(12,'15','A85 H130 I130 J130 K53'+E,Object.assign({A:txt('PAINTING DATE = date the piece was painted (complete system). B&P weight 25 % of the piece; final inspection accepted = READY TO SHIP (10 %).'),H:txt('PROGRESS (area painted)'),K:fx(`IF(SUM(${rg('G')})=0,0,SUMPRODUCT(${rg('O')},${rg('G')})/SUM(${rg('G')}))`,0)},elv(12)));
    for(let r=13;r<=16;r++) S(r,'15',E.trim(),elv(r));
    rows.push(noteRow(21,85,'How to fill: PAINTING DATE when the piece is painted, then FINAL INSPECTION (list) and its date. STATUS and % PROGRESS are calculated; WELD RELEASED is read from WELDING PROGRESS. READY TO SHIP = final inspection accepted. When shipped, enter PACKING LIST No. and SHIP DATE → IN TRANSIT.'),noteRow(22,124,'Fill example (not real data): Painting date 12/10/2026 · Final inspection READY TO SHIP · 12/11/2026 → STATUS = READY TO SHIP · Packing list PL-010 · Ship date 12/14/2026 → IN TRANSIT.'));
    const HH=['ITEM','MARK','PIECE','NODE','ELEMENT','CONDITION','AREA (ft²)','WEIGHT (lb)','WELD RELEASED','PAINTING DATE','FINAL INSPECTION','INSP. DATE','READY TO SHIP DATE','STATUS','% PROGRESS (B&P)','REMARKS','PACKING LIST No.','SHIP DATE'];
    rows.push({r:24,ht:'45',ch:true,cells:'ABCDEFGHIJKLMNOPQR'.split('').map((c,i)=>({c,s:5,v:HH[i]}))});
    P.forEach((x,i)=>{ const r=a+i, st=`'WELDING PROGRESS'!`; S(r,'15','A29 B130 C29 D29 E10 F29 G32 H31 I125 J37 K21 L37 M139 N29 O74 P127 Q21 R37',Object.assign({A:num(i+1),B:txt(x.mark),C:txt(x.c),E:txt(x.el),F:txt(x.comp?'COMPOSITE':'LOOSE'),G:num(x.a),H:num(x.w),I:x.comp?fx(`IF(INDEX(${st}$O$21:$O$${wr},MATCH(D${r},${st}$C$21:$C$${wr},0))="RELEASED","YES","NO")`,'NO'):txt('N/A'),
      M:fx(`IF(K${r}="READY TO SHIP",IF(L${r}<>"",L${r},IF(J${r}<>"",J${r},"")),"")`,''),N:fx(`IF(AND(Q${r}<>"",K${r}="READY TO SHIP"),"IN TRANSIT",IF(K${r}="READY TO SHIP","READY TO SHIP",IF(K${r}="REJECTED","REPAIR",IF(J${r}<>"","PAINTING",IF(I${r}="NO","WAITING FOR WELD","PENDING")))))`,x.comp?'WAITING FOR WELD':'PENDING'),O:fx(`IF(OR(J${r}<>"",K${r}="READY TO SHIP"),1,0)`,0)},x.node?{D:num(x.node)}:{})); });
    S(T,'15','A69 B99 C69 D69 E69 F69 G55 H55 I69 J69 K69 L69 M69 N69 O136 P69 Q69 R69',{B:txt('TOTAL'),G:fx(`SUM(G${a}:G${z})`,TA),H:fx(`SUM(H${a}:H${z})`,sum('w')),O:fx('K12',0)});
    return {rows,lastCol:'R',filter:`$A$24:$R$${z}`,merges:['A4:F4','H10:J10','H11:J11','H12:J12','H4:K4','H5:J5','H6:J6','H7:J7','H8:J8','H9:J9','M4:P4'],subst:[[/<autoFilter ref="[^"]*">/,`<autoFilter ref="A24:R${z}">`],[/sqref="[^"]*"/g,m=>m.replace(/([A-Z]+25:[A-Z]+)\d+/g,`$1${z}`)]],pieces:P}; };

  // resumen del modelo que usan WELDING y PAINTING (sin referencias circulares)
  AV.modelSummary=M=>M.map(x=>({mark:x.a.mark,qty:x.a.qty,comp:x.comp,pwt:x.a.row&&x.a.row.pwt,twt:x.a.row&&x.a.row.twt,sa:x.a.row&&x.a.row.saAll,seq:x.weldSeq,
    loose:x.parts.filter(p=>p.loose).reduce((t,p)=>t+(p.r.pwt||0)*p.r.qty,0),looseT:x.parts.filter(p=>p.loose).reduce((t,p)=>t+(p.r.twt||0),0),looseSa:x.parts.filter(p=>p.loose).reduce((t,p)=>t+(p.r.saAll||0),0),
    lp:x.parts.filter(p=>p.loose).map(p=>({mark:p.r.mark,qty:p.r.qty,pwt:p.r.pwt,twt:p.r.twt,sa:p.r.saAll,type:p.ptype||(/^L\s*\d/i.test(String(p.r.D||p.r.desc||''))?'CLIP_ANGLE':p.role)})),
    wp:x.parts.filter(p=>p.welded).map(p=>({k:p.k,mark:p.r.mark,qty:p.r.qty,role:p.role,n:p.weldN,weld:p.weld}))}));

  /* elevación de cada pieza: centro de su caja en Z, medido desde lo más bajo del modelo (ft).
     Orden de piezas iguales: por X, Y, Z; las partes sueltas siguen el orden de su ensamble padre. */
  AV.pieceElevations=function(G,M){ const I=G.instances||[], box=new Map(); let base=Infinity;
    I.forEach((x,i)=>{ const F=AV._faces(G,x.mark,x.W); let mn=[1e12,1e12,1e12],mx=[-1e12,-1e12,-1e12]; F.forEach(f=>f.pts.forEach(p=>{ for(let k=0;k<3;k++){ if(p[k]<mn[k]) mn[k]=p[k]; if(p[k]>mx[k]) mx[k]=p[k]; } })); if(mn[0]>1e11) return; box.set(i,{x,c:[0,1,2].map(k=>(mn[k]+mx[k])/2)}); if(mn[2]<base) base=mn[2]; });
    const key=b=>[Math.round(b.c[0]),Math.round(b.c[1]),Math.round(b.c[2])], cmp=(a,b)=>{ const p=key(a),q=key(b); return p[0]-q[0]||q[1]-p[1]||p[2]-q[2]; }, H=b=>rd((b.c[2]-base)/304.8,2), byMark=new Map();
    box.forEach(b=>{ const m=b.x.mark; if(!byMark.has(m)) byMark.set(m,[]); byMark.get(m).push(b); }); byMark.forEach(l=>l.sort(cmp));
    const out={}, get=m=>byMark.get(m)||[]; M.forEach(x=>{ const mains=get(x.mark); out[x.mark]=mains.map(H); (x.lp||[]).forEach(l=>{ const list=[]; mains.forEach(h=>{ get(l.mark).filter(b=>b.x.asmHost===h.x.asmHost).forEach(b=>list.push(H(b))); }); out[`${x.mark}>${l.mark}`]=list; }); });
    return out; };

  /* ---- 3D CONTROL: una fila por pieza física, mismo orden que PAINTING PROGRESS ---- */
  AV.controlSheet=function(o,M,bom,R,V,W,P,E){ const BS=`'${o.bomSheet}'!`, CT=`'CUTTING TRACKING'!`, VR=`'VENDOR REQUISITION'!`, WP=`'WELDING PROGRESS'!`, PP=`'PAINTING PROGRESS'!`, WK=`'WEEKLY CUT-OFFS'!`, DR=`'DAILY REPORT'!`;
    const B={}, cut={}; o.parts.forEach((p,i)=>{ cut[p.row]=11+i; }); let cur=null; bom.rows.forEach(r=>{ const c={}; r.cells.forEach(x=>c[x.c]=x); const t=c.AP&&c.AP.v; if(t==='A'){ cur=B[c.B.v]={r:r.r,P:[],H:{}}; } else if(t==='P') cur.P.push({r:r.r,mark:c.B.v,desc:c.E.v,grade:c.F.v}); else if(t==='H') cur.H[c.B.v]={r:r.r,mark:c.B.v,desc:c.E.v,grade:c.F.v}; });
    const mat=l=>{ const d=plateDim(l.desc); return d?{show:`${d.tok}|${l.grade}`,key:`PL${Number.isInteger(d.t)?d.t.toFixed(1):d.t}|${l.grade}`}:{show:`${l.desc}|${l.grade}`,key:`${l.desc}|${l.grade}`}; };
    const mats=lines=>{ const m=lines.map(mat), first=m[0], rest=[...new Map(m.slice(1).filter(x=>x.show!==first.show).map(x=>[x.show,x])).values()].sort((a,b)=>a.show<b.show?-1:a.show>b.show?1:0); return [first,...rest]; };
    const n=P.pieces.length, a=20, z=19+n, rows=[], rg=c=>`$${c}$${a}:$${c}$${z}`, S=(r,ht,sp,v,ch)=>rows.push({r,ht,ch:!!ch,cells:cells(sp,v)}), TW=P.pieces.reduce((t,x)=>t+x.w,0);
    rows.push(noteRow(1,60,'FABRICATION CONTROL PER PIECE – STATUS FOR THE 3D MODEL','18',true),noteRow(2,61,'One row per physical piece. 3D ID = piece number in the 3D viewer (Weld Map HTML). Status is calculated at the CUT-OFF DATE from the dates in the other sheets.'));
    S(4,'15','A70 B70 C71 D71 E59',{A:txt('CUT DATE'),C:{f:`${DR}$C$8`},E:txt('Default: DAILY REPORT date (today). Type another date to see the status on that day.')});
    S(6,'15','A129 B129 C129 D129 E129 F129 G129 H129 I129 J129',{A:txt('STATUS AT CUT-OFF'),C:txt('PIECES'),D:txt('WEIGHT (lb)'),E:txt('% WEIGHT'),F:txt('MEANING / SOURCE OF THE DATE')});
    [['NOT YET ARRIVED','Material not yet arrived (VENDOR REQUISITION / manual date)'],['MATERIAL IN SHOP','Material received in shop, not cut yet'],['MATERIAL CUT','Composite piece: all its parts cut, fit-up not started (loose pieces go straight to IN FABRICATION when cut)'],['IN FABRICATION','Loose: cut, fit-up not recorded yet · Composite: fit-up started, welding not 100 % yet'],['LOOSE COMPLETED','Loose piece with fit-up recorded in the BOM (FIT QTY + Fitter Date)'],['COMPOSITE COMPLETED','Composite piece welded 100 % in the BOM (and its parts fitted)'],['PAINTED','Painted: PAINTING DATE recorded (PAINTING PROGRESS), final inspection pending'],['READY TO SHIP','Final paint inspection accepted = ready for shipment'],['IN TRANSIT','Shipped with packing list (informative)']].forEach(([t,m],i)=>{ const r=7+i, s=140+i; S(r,'15',`A${s} B${s} C64 D75 E76 F132 G132 H132 I132 J132`,{A:txt(t),C:fx(`COUNTIF(${rg('V')},A${r})`,i?0:n),D:fx(`SUMIFS(${rg('G')},${rg('V')},A${r})`,i?0:TW),E:fx(`IF(SUM(${rg('G')})=0,0,D${r}/SUM(${rg('G')}))`,i?0:1),F:txt(m)}); });
    S(16,'15','A99 B99 C54 D55 E53',{A:txt('TOTAL'),C:fx('SUM(C7:C15)',n),D:fx('SUM(D7:D15)',TW),E:fx('SUM(E7:E15)',1)});
    S(17,'15','A99 B99 C99 D99 E53 F85',{A:txt('WEIGHTED PHYSICAL PROGRESS'),E:{f:`IFERROR(SUMPRODUCT(${rg('G')},${rg('AK')})/SUM(${rg('G')}),0)`},F:txt('Σ (weight × % PROGRESS of each piece) ÷ total weight. Milestone weights in WEEKLY CUT-OFFS (engineering 5 · procurement 10 · manufacturing 50 · B&P 25 · RTS 10).')});
    rows.push(noteRow(18,85,'How to fill: in this sheet only the manual dates (yellow) when needed. Material: VENDOR REQUISITION; cut: CUTTING TRACKING; fit-up/welding: WELDING PROGRESS and BOM; paint, READY TO SHIP and packing list: PAINTING PROGRESS. Each date also takes the earliest later date (a piece in painting was already cut, fitted and welded).'));
    const HH=['3D ID','MARK','PIECE','NODE','ELEMENT','CONDITION','WEIGHT (lb)','ELEV. (ft)','REQUIRED MATERIALS','MATERIAL DATE (auto)','MATERIAL DATE (manual)','CUT DATE (auto)','CUT DATE (manual)','FABRICATION START (manual)','D. ARRIVAL','D. CUT','D. FABRICATION','D. WELDING','D. PAINTING','D. READY TO SHIP','D. IN TRANSIT','STATUS AT CUT-OFF','REMARKS','% FIT (piece)','% WELD (piece)','% FABRICATION (piece)','D. FIT-UP (BOM)','D. WELD (BOM)','PIECE DIE STAMP','PACKING LIST No.','WELD INCHES','INCHES WELDED','D. COMPLETED','% CUT (piece)','D. FIT-UP COMPLETE','AREA (ft²)','% PROGRESS (piece)'];
    rows.push({r:19,ht:'39.75',ch:true,cells:HH.map((v,i)=>({c:AV.colName(i+1),s:5,v}))});
    const PS='A72 B130 C29 D29 E10 F29 G31 H149 I150 J134 K37 L134 M37 N37 O73 P73 Q73 R73 S73 T73 U73 V29 W127 X74 Y74 Z131 AA73 AB73 AC29 AD29 AE38 AF38 AG73 AH74 AI73 AJ151 AK65', ms=[7,8,9,10,11,12,13], miss=[];
    const seen={}, wrow={}; W.nodes.forEach((x,i)=>{ wrow[x.node]=21+i; });
    P.pieces.forEach((x,i)=>{ const r=a+i, p=25+i, comp=x.comp, [ci,cq]=x.c.split('/').map(Number), own=x.parent?`${x.parent}>${x.mark}`:x.mark; seen[own]=(seen[own]||0)+1; const e=E[own]&&E[own][seen[own]-1];
      const bx=B[x.parent||x.mark], h=x.parent?bx.H[x.mark]:null, lines=h?[h]:bx.P, top=h?h.r:bx.r, f=lines[0].r, l=lines[lines.length-1].r, U=`${BS}U${f}:U${l}`, Y=`${BS}Y${f}:Y${l}`, mm=mats(h?[h]:bx.P), mv=mm.map(m=>{ const q=V.mats[m.key]; if(!q) miss.push(m.key); return `${VR}$L$${q}`; }).join(','), cr=lines.map(q=>cut[q.r]);
      const cond=cr.map(c=>`${CT}$U$${c}>=${CT}$H$${c}*${ci}/${cq}`).join(','), maxV=cr.map(c=>`${CT}$V$${c}`).join(','), pct=cr.map(c=>`MIN(1,N(${CT}$U$${c})/(${CT}$H$${c}*${ci}/${cq}))`).join('+'), wr=comp?wrow[x.node]:null;
      const w=(c,d)=>`IF(F${r}="LOOSE",${WK}$D$${c},${WK}$C$${c})`, up=(c)=>`IF(ISNUMBER(${c}${r})*(N(${c}${r})<=$C$4),1,`;
      const v={A:num(i+1),B:txt(x.mark),C:txt(x.c),E:txt(x.el),F:txt(comp?'COMPOSITE':'LOOSE'),G:num(x.w),I:txt(mm.map(m=>m.show).join(', ')),J:fx(`IF(COUNT(${mv})<${mm.length},"",MAX(${mv}))`,''),L:fx(`IF(AND(${cond}),MAX(${maxV}),"")`,''),
        O:fx(`IF(COUNT(J${r},K${r},P${r})=0,"",IF(ISNUMBER(K${r}),MIN(K${r},P${r}),MIN(J${r},P${r})))`,''),P:fx(`IF(COUNT(L${r},M${r},Q${r})=0,"",MIN(L${r},M${r},Q${r}))`,''),
        S:fx(`IF(COUNT(${PP}J${p},T${r})=0,"",MIN(${PP}J${p},T${r}))`,''),T:fx(`IF(COUNT(${PP}M${p},U${r})=0,"",MIN(${PP}M${p},U${r}))`,''),U:fx(`IF(AND(${PP}Q${p}<>"",${PP}M${p}<>""),IF(${PP}R${p}<>"",${PP}R${p},${PP}M${p}),"")`,''),
        V:fx(`IF(AND(ISNUMBER(U${r}),U${r}<=$C$4),"IN TRANSIT",IF(AND(ISNUMBER(T${r}),T${r}<=$C$4),"READY TO SHIP",IF(AND(ISNUMBER(S${r}),S${r}<=$C$4),"PAINTED",IF(AND(ISNUMBER(AG${r}),AG${r}<=$C$4),IF(F${r}="LOOSE","LOOSE COMPLETED","COMPOSITE COMPLETED"),IF(AND(NOT(F${r}="LOOSE"),ISNUMBER(Q${r}),Q${r}<=$C$4),"IN FABRICATION",IF(AND(ISNUMBER(P${r}),P${r}<=$C$4),IF(F${r}="LOOSE","IN FABRICATION","MATERIAL CUT"),IF(AND(ISNUMBER(O${r}),O${r}<=$C$4),"MATERIAL IN SHOP","NOT YET ARRIVED")))))))`,'NOT YET ARRIVED'),
        X:fx(`MAX(0,MIN(1,${BS}AF${top}*${cq}-${ci-1}))`,0),Y:fx(`MAX(0,MIN(1,${BS}AG${top}*${cq}-${ci-1}))`,0),Z:fx(`${DR}$C$5*X${r}+${DR}$C$6*Y${r}`,0),AA:fx(`IF(X${r}>0,IF(COUNT(${U})=0,"",MIN(${U})),"")`,''),AD:fx(`IF(${PP}Q${p}<>"",${PP}Q${p},"")`,''),
        AH:fx(`IF(P${r}<>"",1,(${pct})/${cr.length})`,0),AJ:fx(`${PP}G${p}`,x.a),
        AK:{f:`${w(7)}+${w(8)}*ISNUMBER(O${r})*(N(O${r})<=$C$4)+${w(9)}*${up('P')}($C$4>=TODAY())*N(AH${r}))+${w(10)}*${up('AI')}($C$4>=TODAY())*N(X${r}))+${w(11)}*${up('AG')}($C$4>=TODAY())*N(Y${r}))+${w(12)}*ISNUMBER(S${r})*(N(S${r})<=$C$4)+${w(13)}*ISNUMBER(T${r})*(N(T${r})<=$C$4)`}};
      if(e!=null) v.H=num(e); if(x.node) v.D=num(x.node);
      if(comp) Object.assign(v,{Q:fx(`IF(COUNT(N${r},${WP}J${wr},AA${r},AI${r},AG${r})=0,"",MIN(N${r},${WP}J${wr},AA${r},AI${r},AG${r}))`,''),R:fx(`IF(COUNT(${WP}K${wr},AB${r},AG${r})=0,"",MIN(${WP}K${wr},AB${r},AG${r}))`,''),AB:fx(`IF(Y${r}>0,IF(COUNT(${Y})=0,AA${r},MIN(${Y})),"")`,''),AC:fx(`IF(${WP}L${wr}<>"",${WP}L${wr},IF(${BS}Z${top}<>"",${BS}Z${top},""))`,''),AE:fx(`${WP}G${wr}`,W.nodes[wr-21].G),AF:fx(`${WP}S${wr}`,0),AG:fx(`IF(COUNT(${WP}K${wr},S${r})=0,"",MIN(${WP}K${wr},S${r}))`,''),AI:fx(`IF(AND(X${r}>=0.999,COUNT(${U})>0),IF(ISNUMBER(AG${r}),MIN(AG${r},MAX(${U})),MAX(${U})),IF(ISNUMBER(AG${r}),AG${r},""))`,'')});
      else Object.assign(v,{Q:fx(`IF(COUNT(N${r},AG${r})=0,"",MIN(N${r},AG${r}))`,''),R:txt('N/A'),AB:txt('N/A'),AC:fx(`IF(${BS}Z${top}<>"",${BS}Z${top},"")`,''),AE:num(0),AF:num(0),AG:fx(`IF(COUNT(AA${r},S${r})=0,"",MIN(AA${r},S${r}))`,''),AI:fx(`IF(ISNUMBER(AG${r}),AG${r},"")`,'')});
      S(r,'15',PS,v); });
    return {rows,lastCol:'AK',filter:`$A$19:$AK$${z}`,subst:[[/<autoFilter ref="[^"]*">/,`<autoFilter ref="A19:AK${z}">`],[/sqref="[^"]*"/g,m=>m.replace(/([A-Z]+20:[A-Z]+)\d+/g,`$1${z}`)]],miss}; };

  /* ---- DAILY REPORT, WEEKLY NOTES, WEEKLY CUT-OFFS y CONNECTION SUMMARY: se conservan de la plantilla tal cual.
     Solo se reapuntan las filas finales de otras hojas y los textos que nombran datos del proyecto. ---- */
  const xmlEsc=t=>String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
  function retarget(x,T,N){ const q=s=>s.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'), bo=`(?:'|&apos;)${q(T.bomSheet)}(?:'|&apos;)!`, sh=(name)=>new RegExp(`((?:'|&apos;)${q(name)}(?:'|&apos;)!(?:\\$?[A-Z]+\\$?\\d+:)?\\$?[A-Z]+\\$?)(\\d+)\\b`,'g');
    return x.replace(/<f([^>]*)>([^<]*)<\/f>/g,(m,attr,f)=>{ let t=f; t=t.replace(new RegExp(`(${bo}\\$?[A-Z]+\\$?\\d+:\\$?[A-Z]+\\$?)${T.bomLast}\\b`,'g'),`$1${N.bomLast}`).replace(new RegExp(`(${bo}\\$?[A-Z]+\\$?)${T.bomTot}\\b`,'g'),`$1${N.bomTot}`);
      t=t.split(`&apos;${T.bomSheet}&apos;`).join(`&apos;${N.bomSheet}&apos;`).split(`'${T.bomSheet}'`).join(`'${N.bomSheet}'`);
      t=t.replace(sh('3D CONTROL'),(a,b,c)=>b+(+c===T.ctlLast?N.ctlLast:c)).replace(sh('WELDING PROGRESS'),(a,b,c)=>b+(+c===T.wTot?N.wTot:c)); return `<f${attr}>${t}</f>`; }).replace(/(<c r="[A-Z]+\d+"[^>]*>)(<f[^>]*>[^<]*<\/f>)<v>[^<]*<\/v>/g,(m,c,f)=>c.replace(/ t="(?:n|str|b|e)"/,'')+f); }  // sin valor guardado: Excel los calcula al abrir
  function setCell(x,ref,val){ const re=new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`); return x.replace(re,(m,at)=>{ const s=(at.match(/ s="\d+"/)||[''])[0]; return typeof val==='number'?`<c r="${ref}"${s}><v>${val}</v></c>`:`<c r="${ref}"${s} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(val)}</t></is></c>`; }); }
  const cellText=(x,ref,sst)=>{ const m=x.match(new RegExp(`<c r="${ref}"[^>]*?t="s"[^>]*><v>(\\d+)</v>`)); return m?sst[+m[1]]:''; };
  AV.finishSheets=function(tpl,sst,T,o){ const N={bomSheet:o.bomSheet,bomLast:o.bomLast,bomTot:o.bomTotalRow,ctlLast:o.ctlLast,wTot:o.wTot}, out={daily:retarget(tpl.daily,T,N),cuts:retarget(tpl.cuts,T,N)};
    const L=o.loose||[], lEn=L.map(l=>`${l.mark} ×${l.qty} under ${l.asm}`).join(' · '), A=o.anchors||{};
    // proyecto nuevo: WEEKLY NOTES sale sin notas (se escriben a mano); solo se vacían los datos B:K de las filas escritas, nunca una fórmula
    let n=retarget(tpl.notes,T,N); const clearCell=(x,ref)=>x.replace(new RegExp(`<c r="${ref}"([^>]*?)(?:/>|>[\\s\\S]*?</c>)`),(m,at)=>`<c r="${ref}"${(at.match(/ s="\d+"/)||[''])[0]}/>`);
    for(let r=6;r<=12;r++) 'BCDEFGHIJK'.split('').forEach(c=>{ n=clearCell(n,c+r); }); out.notes=n;
    const d0=new Date(`${o.startDate}T12:00:00`); out.cuts=setCell(out.cuts,'B4',serial(o.startDate)+((5-d0.getDay()+7)%7));   // primer corte: primer viernes desde la fecha de inicio
    let c=retarget(tpl.summary,T,N); const ct=r=>cellText(c,r,sst), A16=ct('A16').replace(/\([^()]*\bunder\b[^()]*?, shop bolted/,`(${lEn}, shop bolted`), bd=o.boltDoc;
    [['A1',`SUMMARY – CONNECTION TYPE (${o.proj})`],['E12',`Bolt Summary ${bd} – SHOP BOLT`],['E13',`Bolt Summary ${bd} – FIELD BOLT`],['A14',`Adhesive anchors ${A.desc}`],['C14',A.model],['E14',`IFC ${A.model} (${A.basePlates} base plates × ${A.perPlate}) · Bolt Summary ${A.n}`],['A16',A16]].forEach(([r,v])=>{ c=setCell(c,r,v); }); out.summary=c; return out; };
  AV.looseList=M=>{ const o=[]; M.forEach(x=>(x.lp||[]).forEach(l=>o.push({mark:l.mark,qty:l.qty,asm:x.mark}))); return o; };

  /* ---- Excel de avance completo de una unidad: plantilla + hojas generadas. Después va «Agregar al Excel de avance» (troqueles y pestañas). ---- */
  const TPL={bomSheet:'BOM TPL - CONNECTION',bomLast:204,bomTot:206,ctlLast:579,wTot:148};
  // el Bolt Summary de la unidad: el que nombra su proyecto; si no, el del mismo orden de carga
  AV.boltForUnit=(list,meta,i)=>{ const key=(String(meta.model||meta.project||'').match(/(\d+)\s*m\b/i)||[])[1]; const hit=key&&list.find(b=>new RegExp(key+'\\s*m\\b','i').test(String(b.meta&&b.meta.job||''))); return hit||list[i]||list[0]; };
  AV.generate=async function(tplBuf,bomTpl,o){ const un=t=>t.replace(/&quot;/g,'"').replace(/&apos;/g,"'").replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');
    const Z=await JSZip.loadAsync(tplBuf), rd=n=>Z.file(`xl/worksheets/sheet${n}.xml`).async('string'), wr=(n,x)=>Z.file(`xl/worksheets/sheet${n}.xml`,x);
    const M0=AV.model(o.unit), MS=AV.modelSummary(M0), G=asmGeo.get(normMark(M0[0].a.mark)).G, ELEV=AV.pieceElevations(G,MS), ui=o.unit&&P.units?P.units.findIndex(u=>u.label===o.unit):-1, uM=ui>=0&&P.units[ui].meta||P.meta, bl=bolts&&bolts.list&&ui>=0?AV.boltForUnit(bolts.list,uM,ui):bolts, Nm=AV.names(uM,M0.map(x=>x.a.mark),bl&&bl.meta), B=AV.boltInput(bl), shopBolt=AV.shopBoltText(B);
    const S=AV.bomSheet(M0,{tpl:bomTpl,title:Nm.title,boltDoc:Nm.boltDoc,shopBolt}), parts=AV.mtoParts(S), am=AV.anchorModel(S); if(B.anchors&&am) Object.assign(B.anchors,am);
    const totI=(S.rows.find(R=>R.r===S.TOT).cells.find(c=>c.c==='I')||{}).v, bomDoc=(Nm.doc.match(/BOM-\d+/)||[''])[0];
    const longest={}; P.rows.filter(x=>!x.isAsm&&x.len).forEach(x=>{ const g=String(x.grade||'').trim(), k=x.kind==='plate'?`PL${x.t}|${g}`:`${x.D}|${g}`; if(!longest[k]||x.len>longest[k].len) longest[k]={len:x.len,txt:String(x.lenText||'').trim()}; });   // de todo el proyecto, no solo de esta unidad
    const OPT=Object.assign({},B,{parts,longest,stock:o.stock,sheet:o.sheet||[4,8],proj:Nm.proj,bomDoc,bomDocFull:Nm.doc,bomSheet:Nm.sheet,bomTotalRow:S.TOT,bomTotal:totI,spares:o.spares??0.05,startDate:o.startDate,mtoFile:o.mtoFile,ptpDocFull:o.ptpDocFull,kerf:o.kerf,trim:o.trim,gap:o.gap,edge:o.edge});
    const R=AV.mtoSheets(OPT), V=AV.vendorSheets(OPT,R), C=AV.cuttingSheet(OPT,parts), bomI={rows:S.rows,last:S.last}, Wd=AV.weldingSheet(OPT,MS,bomI), Pt=AV.paintingSheet(OPT,MS,bomI,Wd), Ct=AV.controlSheet(OPT,MS,bomI,R,V,Wd,Pt,ELEV);
    const sstx=await Z.file('xl/sharedStrings.xml').async('string'), sst=[...sstx.matchAll(/<si>([\s\S]*?)<\/si>/g)].map(m=>un([...m[1].matchAll(/<t[^>]*>([^<]*)<\/t>/g)].map(x=>x[1]).join('')));
    const GR=[...(G.assemblies||new Map()).values()].filter(v=>/GROUT/i.test(v.name||'')).map(v=>({mark:v.mark,inst:(G.instances||[]).filter(i=>i.asmMark===v.mark).length})), FS=AV.finishSheets({daily:await rd(3),notes:await rd(4),summary:await rd(5),cuts:await rd(11)},sst,TPL,Object.assign({},OPT,{bomLast:S.last,ctlLast:19+Pt.pieces.length,wTot:21+Wd.nodes.length,loose:AV.looseList(MS),grout:GR.reduce((t,v)=>t+v.inst,0),groutMark:GR[0]?String(GR[0].mark).replace(/\(.*$/,''):'',shopBolt}));
    const put=async(n,Sh)=>wr(n,AV.sheetXml(await rd(n),Sh)); await put(1,S); await put(2,C); await put(6,V.vendor); await put(7,V.po); await put(8,Wd); await put(9,Pt); await put(10,Ct); await put(12,R.takeoff); await put(13,R.mto); await put(14,R.plan); await put(15,R.nesting);
    wr(3,FS.daily.replace(/(<c r="C7"[^>]*>)<v>[^<]*<\/v>/,`$1<v>${serial(o.startDate)}</v>`)); wr(4,FS.notes); wr(5,FS.summary); wr(11,FS.cuts);
    let wb=await Z.file('xl/workbook.xml').async('string'); wb=wb.split(TPL.bomSheet).join(Nm.sheet.replace(/&/g,'&amp;')).replace(/\$A\$3:\$AT\$\d+/,`$A$3:$AT$${S.last}`).replace(/\$A\$10:\$Z\$\d+/,C.filter).replace(/\$A\$24:\$R\$\d+/,Pt.filter).replace(/\$A\$4:\$M\$\d+/,R.takeoff.filter).replace(/\$A\$20:\$V\$\d+/,Wd.filter).replace(/\$A\$19:\$AK\$\d+/,Ct.filter); Z.file('xl/workbook.xml',wb);
    const blob=await Z.generateAsync({type:'blob',compression:'DEFLATE',mimeType:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
    return {blob,name:`BOM_Connection_Type_${(P.ifcSplit&&o.unit?`${Nm.proj} ${o.unit}`:Nm.proj).replace(/\s+/g,'_')}_EN.xlsx`,over:R.over,overPlates:R.overPlates,miss:Ct.miss}; };
  // Excel pide reparar lo que LibreOffice deja en las gráficas: noMultiLvlLbl en un eje de fechas e id="0" en el dibujo. No toca hojas ni fórmulas.
  AV.fixCharts=async function(Z){ let n=0; for(const name of Object.keys(Z.files)){ const ch=/^xl\/charts\/chart\d+\.xml$/.test(name), dr=/^xl\/drawings\/drawing\d+\.xml$/.test(name); if(!ch&&!dr) continue; const x=await Z.file(name).async('string'); let y=x;
      if(ch) y=x.replace(/<c:dateAx>[\s\S]*?<\/c:dateAx>/g,m=>m.replace(/<c:noMultiLvlLbl[^>]*\/>/g,''));
      else if(x.includes('<xdr:cNvPr id="0"')){ let k=Math.max(1,...[...x.matchAll(/<xdr:cNvPr id="(\d+)"/g)].map(m=>+m[1])); y=x.replace(/<xdr:cNvPr id="0"/g,()=>`<xdr:cNvPr id="${++k}"`); }
      if(y!==x){ Z.file(name,y); n++; } } return n; };
  AV._rd=rd; AV._g6=g6; AV._ftIn=ftIn;
})(typeof window!=='undefined'?window:globalThis);
