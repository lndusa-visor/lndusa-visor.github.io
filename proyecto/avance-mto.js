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
    S.rows.forEach(R=>{ const c={}; R.cells.forEach(x=>c[x.c]=x); const t=c.AP&&c.AP.v, v=k=>c[k]?c[k].v:null; if(t==='A') asm=v('B');
      if(t==='P'||t==='H') out.push({asm,mark:v('B'),qty:v('D'),desc:v('E'),grade:v('F'),len:ftIn(v('G')),wt:v('H')}); });
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
    P.forEach((p,i)=>{ const r=5+i, w=p.d?p.d.w:null, area=p.cat==='PLATE'?p.qty*p.lenFt*w/12:0; p.tl=p.qty*p.lenFt; p.ta=area; p.tw=p.qty*p.wt;
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
    mt.push({r:4,ht:'31.5',ch:true,cells:LET.map((c,i)=>({c,s:5,v:MH[i]}))});
    r=5; const MS='A63 B10 C10 D75 E155 F100 G29 H108 I38 J75 K76 L75 M29', M=[];
    profs.forEach(g=>{ const q=sumV(g.rows,'qty'), nl=sumV(g.rows,'tl'), nw=sumV(g.rows,'tw'), gl=g.bars*g.S, gw=gl*g.lf; M.push({q,nw,gw});
      const sf=c=>`SUMIFS(${tb(c)},${tb('D')},B${r},${tb('F')},C${r})`;
      mt.push({r,ht:'15',ch:false,cells:cells(MS,{A:txt(g.cat),B:txt(g.desc),C:txt(g.grade),D:fx(sf('C'),q),E:fx(sf('J'),nl),F:fx(sf('M'),nw),G:txt(`${g.S} ft`),H:fx(`${RQ(PL)}!C${g.subRow}`,g.bars),I:fx(`H${r}*${g.S}`,gl),J:fx(`I${r}*${g.lf}`,gw),K:fx(`IFERROR(E${r}/I${r},0)`,nl/gl),L:fx(`J${r}-F${r}`,gw-nw),M:txt('ft')})}); r++; });
    plates.forEach(g=>{ const q=sumV(g.rows,'qty'), na=sumV(g.rows,'ta'), nw=sumV(g.rows,'tw'), ga=g.sheets*sA, gw=ga*40.84*g.t; M.push({q,nw,gw});
      const sf=(c,pl)=>`SUMIFS(${tb(c)},${tb('I')},${g.t},${tb('F')},C${r}${pl?`,${tb('E')},"PLATE"`:''})`;
      mt.push({r,ht:'15',ch:false,cells:cells(MS,{A:txt('PLATE'),B:txt(g.tok),C:txt(g.grade),D:fx(sf('C',1),q),E:fx(sf('K'),na),F:fx(sf('M',1),nw),G:txt(sTxt),H:fx(`${RQ(NS)}!C${g.subRow}`,g.sheets),I:fx(`H${r}*${sA}`,ga),J:fx(`I${r}*40.84*${g.t}`,gw),K:fx(`IFERROR(E${r}/I${r},0)`,na/ga),L:fx(`J${r}-F${r}`,gw-nw),M:txt('ft²')})}); r++; });
    const a=5, b=r-1, tq2=sumV(M,'q'), tn=sumV(M,'nw'), tg=sumV(M,'gw'), T=r;
    mt.push({r,ht:'15',ch:false,cells:cells('A99 B69 C69 D55 E69 F55 G69 H69 I69 J55 K53 L55 M69',{A:txt('TOTAL STEEL'),D:fx(`SUM(D${a}:D${b})`,tq2),F:fx(`SUM(F${a}:F${b})`,tn),J:fx(`SUM(J${a}:J${b})`,tg),K:fx(`IFERROR(F${r}/J${r},0)`,tn/tg),L:fx(`SUM(L${a}:L${b})`,tg-tn)})}); r++;
    mt.push({r,ht:'15',ch:false,cells:cells('A156 F100 G33',{A:txt('Check vs BOM total weight'),F:fx(`${RQ(o.bomSheet)}!I${o.bomTotalRow}`,tn),G:fx(`IF(ABS(F${r}-F${T})<1,"OK","CHECK")`,'OK')})}); r+=2;
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
  AV._rd=rd; AV._g6=g6; AV._ftIn=ftIn;
})(typeof window!=='undefined'?window:globalThis);
