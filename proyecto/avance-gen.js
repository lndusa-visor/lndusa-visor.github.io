/* Generador del Excel de avance (formato BOM_Connection_Type) a partir del BOM + IFC + Bolt Summary.
   Se carga en la página «Nuevo proyecto» y usa lo que ella ya leyó: P (BOM), asmGeo / ifc (IFC), bolts (Bolt Summary).
   Hojas: una por una, con la misma estructura y fórmulas que el Excel de referencia de los Piperacks 30 m / 55 m. */
(function(G){
  const AV={};
  const MM=25.4; let BTOL=30; AV.setTol=t=>BTOL=t;
  const sub=(a,b)=>[a[0]-b[0],a[1]-b[1],a[2]-b[2]], add=(a,b)=>[a[0]+b[0],a[1]+b[1],a[2]+b[2]], dot=(a,b)=>a[0]*b[0]+a[1]*b[1]+a[2]*b[2];
  const cross=(a,b)=>[a[1]*b[2]-a[2]*b[1],a[2]*b[0]-a[0]*b[2],a[0]*b[1]-a[1]*b[0]], len=a=>Math.hypot(a[0],a[1],a[2]), nrm=a=>{ const l=len(a)||1; return [a[0]/l,a[1]/l,a[2]/l]; };
  // W = [ejeX, ejeY, ejeZ, origen] (filas) → punto en el sistema de la pieza
  const tf=(W,v)=>[W[0][0]*v[0]+W[1][0]*v[1]+W[2][0]*v[2]+W[3][0],W[0][1]*v[0]+W[1][1]*v[1]+W[2][1]*v[2]+W[3][1],W[0][2]*v[0]+W[1][2]*v[1]+W[2][2]*v[2]+W[3][2]];
  const inv=M=>{ const [X,Y,Z,O]=M; return [[X[0],Y[0],Z[0]],[X[1],Y[1],Z[1]],[X[2],Y[2],Z[2]],[-dot(X,O),-dot(Y,O),-dot(Z,O)]]; };
  function facesOf(geo,mark,W){ const p=geo.get(mark); if(!p||!p.faces) return []; return p.faces.map(f=>f.map(v=>tf(W,v))).filter(f=>f.length>=3).map(f=>{ let n=[0,0,0]; for(let i=0;i<f.length;i++){ const a=f[i],b=f[(i+1)%f.length]; n=add(n,cross(a,b)); } return {pts:f,n:nrm(n),d:dot(nrm(n),f[0])}; }); }
  // recorte de un polígono por otro convexo (Sutherland–Hodgman) en 2D
  function clip(subject,clipper){ let out=subject; const area=P=>P.reduce((s,p,i)=>{ const q=P[(i+1)%P.length]; return s+p[0]*q[1]-q[0]*p[1]; },0); const C=area(clipper)<0?clipper.slice().reverse():clipper;
    for(let i=0;i<C.length&&out.length;i++){ const a=C[i], b=C[(i+1)%C.length], inside=p=>(b[0]-a[0])*(p[1]-a[1])-(b[1]-a[1])*(p[0]-a[0])>=-1e-6; const inp=out; out=[];
      for(let j=0;j<inp.length;j++){ const P=inp[j], Q=inp[(j+1)%inp.length], ip=inside(P), iq=inside(Q);
        if(ip&&iq) out.push(Q); else if(ip&&!iq) out.push(X(P,Q,a,b)); else if(!ip&&iq){ out.push(X(P,Q,a,b)); out.push(Q); } } }
    return out;
    function X(P,Q,a,b){ const A1=Q[1]-P[1],B1=P[0]-Q[0],C1=A1*P[0]+B1*P[1],A2=b[1]-a[1],B2=a[0]-b[0],C2=A2*a[0]+B2*a[1],D=A1*B2-A2*B1; if(Math.abs(D)<1e-12) return Q; return [(B2*C1-B1*C2)/D,(A1*C2-A2*C1)/D]; } }
  const convex=P=>{ let s=0; for(let i=0;i<P.length;i++){ const a=P[i],b=P[(i+1)%P.length],c=P[(i+2)%P.length]; const z=(b[0]-a[0])*(c[1]-b[1])-(b[1]-a[1])*(c[0]-b[0]); if(Math.abs(z)<1e-9) continue; if(!s) s=Math.sign(z); else if(Math.sign(z)!==s) return false; } return true; };
  // pulgadas de soldadura entre dos partes: bordes de las zonas de contacto (caras paralelas opuestas que se tocan); se ignoran los bordes cortos (espesores)
  let cwN=0; // cordones contados (No. WELDS)
  function contactWeld(FA,FB,tolMM){ tolMM=tolMM||0.8; let L=0;
    for(const a of FA) for(const b of FB){ if(Math.abs(dot(a.n,b.n))<0.999) continue; if(Math.abs(dot(a.n,b.pts[0])-a.d)>tolMM) continue;
      const u=nrm(sub(a.pts[1],a.pts[0])), v=cross(a.n,u), to2=p=>[dot(p,u),dot(p,v)];
      const pa=a.pts.map(to2), pb=b.pts.map(to2); let poly=null;
      if(convex(pa)) poly=clip(pb,pa); else if(convex(pb)) poly=clip(pa,pb); else continue;
      if(!poly||poly.length<3) continue;
      const E=[]; for(let i=0;i<poly.length;i++){ const p=poly[i], q=poly[(i+1)%poly.length]; E.push(Math.hypot(q[0]-p[0],q[1]-p[1])); }
      if(AV.dbg) AV.dbg.push(E.map(e=>+(e/MM).toFixed(3)));
      // franja delgada (canto de placa contra una cara): solo los dos lados largos; perfil o traslape: todo el contorno salvo puntas de ≤ 1/4"
      const R=E.filter(e=>e>0.5); if(R.length<3) continue; const thin=R.length===4&&Math.min(...R)<=25.4&&Math.max(...R)>Math.min(...R)+1;
      if(thin){ const mn=Math.min(...R); for(const e of R) if(e>mn+0.5){ L+=e; cwN++; } } else for(const e of R) if(e>6.6){ L+=e; cwN++; } }
    return L/MM; }
  // pernos que cruzan una parte (segmento del eje contra la caja de la parte en sus coordenadas)
  function crosses(B,p,W,tol){ tol=tol==null?1:tol; const Wi=inv(W), pa=tf(Wi,B.a), pb=tf(Wi,B.b), dv=sub(pb,pa); let t0=0,t1=1;
    for(let j=0;j<3;j++){ if(Math.abs(dv[j])<1e-9){ if(pa[j]<p.mn[j]-tol||pa[j]>p.mx[j]+tol) return false; continue; } let u=(p.mn[j]-tol-pa[j])/dv[j], w=(p.mx[j]+tol-pa[j])/dv[j]; if(u>w) [u,w]=[w,u]; t0=Math.max(t0,u); t1=Math.min(t1,w); } return t0<=t1; }
  const ROLE={END_PLATE:'end plate',SHEAR_PLATE:'shear plate',GUSSET_PLATE:'gusset',STIFFENER_PLATE:'stiffener',BASE_PLATE:'base plate',CAP_PLATE:'cap plate',CLIP_ANGLE:'clip angle',PLATE:'shear plate'};
  const frac=x=>{ const w=Math.floor(x+1e-9), r=Math.round((x-w)*16); if(r===16) return (w+1)+''; if(!r) return w+''; let n=r,d=16; while(n%2===0){ n/=2; d/=2; } return (w?w+' ':'')+n+'/'+d; };

  /* ---------- modelo del proyecto (una unidad = un IFC + un BOM) ---------- */
  AV.model=function(unit){
    const rows=P.rows.filter(r=>!unit||r.unit===unit), asms=P.asms.filter(a=>!unit||(a.row&&a.row.unit===unit)||(!a.row&&!unit));
    const names=(typeof ifc!=='undefined'&&ifc&&ifc.plateType)?ifc.plateType:new Map();
    const boltSpec=(()=>{ const B=typeof bolts!=='undefined'&&bolts&&bolts.rows?bolts.rows:[]; const F=B.filter(b=>/field|campo|site/i.test(b.sec||b.cat||b.type||'')); const cnt=new Map(); F.forEach(b=>{ const k=b.grade+'|'+b.dia; cnt.set(k,(cnt.get(k)||0)+b.qty); });
      const ks=[...cnt].sort((x,y)=>y[1]-x[1]).map(x=>x[0].split('|')); const f=ks[0]||['A325N','3/4'], an=ks.find(k=>k[1]!==f[1]);
      const q=d=>{ d=String(d||'').replace(/"/g,''); return d+'"'; }; return {grade:f[0],dia:q(f[1]),anchor:an?{grade:an[0],dia:q(an[1])}:null}; })();
    let node=0; const out=[];
    asms.forEach(a=>{ const g=asmGeo.get(normMark(a.mark)); const Gm=g?g.G:null, A=g?g.A:null;
      const W=typeof ifcWeld==='function'?ifcWeld(a.mark):null; const welded=new Set((W&&W.w||[]).map(normMark)), bolted=new Set((W&&W.b||[]).map(normMark));
      const isLoosePart=r=>r!==a.main&&bolted.has(normMark(r.mark))&&(r.kind==='plate'||/^L\s*\d/i.test(String(r.D||r.desc||'')));
      // en la referencia, todo lo que no es la principal ni una placa/ángulo apernado en taller va como parte soldada (aunque lleve pernos, p. ej. un muñón W)
      const looseMarks=new Set(a.parts.filter(isLoosePart).map(r=>normMark(r.mark)));
      a.parts.forEach(r=>{ const k=normMark(r.mark); if(r!==a.main&&!looseMarks.has(k)&&(welded.size||bolted.size)) welded.add(k); });
      const comp=(a.auto||'')==='Compuesta';
      // soldadura: contacto de cada parte soldada con la pieza principal (en el primer ensamble del IFC)
      const weldIn=new Map(), weldN=new Map(), weldSeq=[];
      if(Gm&&A){ const main=A.members.find(m=>normMark(m.mark)===normMark(W&&W.main||a.mark))||A.members[0]; const FM=facesOf(Gm,main.mark,main.W);
        // cada parte soldada: contacto con la principal; el contacto entre dos partes soldadas se le cuenta a la que va después en el ensamble del IFC
        const sec=A.members.filter(m=>m!==main&&welded.has(normMark(m.mark))), FS=sec.map(m=>facesOf(Gm,m.mark,m.W));
        sec.forEach((m,j)=>{ if(!weldSeq.includes(normMark(m.mark))) weldSeq.push(normMark(m.mark)); const n0=cwN; let L=contactWeld(FS[j],FM); for(let i=0;i<j;i++) L+=contactWeld(FS[j],FS[i]); const k=normMark(m.mark); weldIn.set(k,(weldIn.get(k)||0)+L); weldN.set(k,(weldN.get(k)||0)+cwN-n0); }); }
      // pernos de campo por instancia y piezas a las que se une
      const inst=(Gm&&Gm.instances||[]).filter(x=>normMark(x.asmMark||'')===normMark(a.mark));
      const hosts=[...new Set(inst.map(x=>x.asmHost))];
      // la referencia toma la instancia más cercana al origen del modelo (primer pórtico): menor X, luego mayor Y, luego menor Z
      const org=h=>{ const x=inst.find(i=>i.asmHost===h); return x?x.W[3]:[0,0,0]; }; hosts.sort((a,b)=>{ const A=org(a),B=org(b); for(let k=0;k<3;k++){ const d=(k===1?-1:1)*(A[k]-B[k]); if(Math.abs(d)>1) return d; } return 0; });
      if(AV.useA&&A&&A.members[0]&&A.members[0].asmHost!=null){ const h0=A.members[0].asmHost; const i=hosts.indexOf(h0); if(i>0){ hosts.splice(i,1); hosts.unshift(h0); } } const BS=(Gm&&Gm.bolts||[]).filter(b=>!/work|shop|taller/i.test(b.loc));
      // como en la referencia: pernos de campo del diámetro del Bolt Summary (sin anclajes) de la primera instancia y las piezas que une
      let nb=0, na=0; const to=new Set(); const dmm=(()=>{ const m=String(boltSpec.dia||'3/4').match(/(\d+)\/(\d+)|(\d+(?:\.\d+)?)/); return m?(m[1]?(+m[1]/+m[2]):+m[3])*MM:19.05; })();
      if(hosts.length&&BS.length){ const byHost=new Map(); (Gm.instances||[]).forEach(x=>{ if(!byHost.has(x.asmHost)) byHost.set(x.asmHost,[]); byHost.get(x.asmHost).push(x); });
        const h=hosts[0], mine=byHost.get(h)||[]; BS.forEach(B=>{ if(!mine.some(x=>{ const p=Gm.get(x.mark); return p&&crosses(B,p,x.W,BTOL); })) return; if(Math.abs((B.d||dmm)-dmm)>1.5){ na++; return; } nb++;
          byHost.forEach((xs,h2)=>{ if(h2===h) return; const am=xs[0].asmMark; if(!am||/\(\?\)|^DND/i.test(am)) return; if(xs.some(x=>{ const p=Gm.get(x.mark); return p&&crosses(B,p,x.W); })) to.add(am); }); }); }
      const perInst=nb, anchors=na;
      // piezas sueltas (clip angles…): pernos de taller y de campo de una pieza, y a qué ensambles va en campo
      const looseInfo=new Map();
      if(hosts.length&&Gm){ const all=(Gm.bolts||[]), mine=(Gm.instances||[]).filter(x=>x.asmHost===hosts[0]);
        looseMarks.forEach(k=>{ const x=mine.find(i=>normMark(i.mark)===k); if(!x) return; const p=Gm.get(x.mark); if(!p) return; let shop=0, field=0; const t2=new Set();
          all.forEach(B=>{ if(!crosses(B,p,x.W,BTOL)) return; if(/work|shop|taller/i.test(B.loc)) { shop++; return; } if(Math.abs((B.d||dmm)-dmm)>1.5) return; field++;
            (Gm.instances||[]).forEach(o=>{ if(o.asmHost===hosts[0]||!o.asmMark||/\(\?\)|^DND/i.test(o.asmMark)) return; const q=Gm.get(o.mark); if(q&&crosses(B,q,o.W)) t2.add(o.asmMark); }); });
          looseInfo.set(k,{shop,field,to:[...t2].sort(natCmp)}); }); }
      // orden de las partes soldadas en el ensamble del IFC (para el texto de la fila del ensamble)
      const ifcOrder=A?A.members.map(m=>normMark(m.mark)):[];
      const parts=a.parts.map(r=>{ const k=normMark(r.mark); const role=ROLE[(names.get&&names.get(r.mark.toUpperCase()))||(names.get&&names.get(k))||'']||(r.kind==='plate'?'plate':/^W/i.test(r.D||'')?'welded beam stub':'part');
        const t=r.t||0; const fil=role==='base plate'&&t>=1?'5/16"':'1/4"';
        return {r,k,main:r===a.main,welded:welded.has(k)&&r!==a.main,loose:looseMarks.has(k),role,fillet:fil,weld:weldIn.get(k)||0,weldN:weldN.get(k)||0,li:looseInfo.get(k)||null}; });
      const nodes=comp?(a.qty>1?`${node+1}–${node+a.qty}`:String(node+1)):'—'; if(comp) node+=a.qty;
      out.push({a,comp,parts,nodes,perInst,anchors,to:[...to].sort(natCmp),spec:boltSpec,ifcOrder,weldSeq}); });
    return out; };

  /* ---------- utilidades de hoja (OOXML) ---------- */
  const colNum=c=>{ let n=0; for(const ch of c) n=n*26+ch.charCodeAt(0)-64; return n; };
  const colName=n=>{ let s=''; while(n>0){ const m=(n-1)%26; s=String.fromCharCode(65+m)+s; n=Math.floor((n-1)/26); } return s; };
  const xesc=t=>String(t).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
  AV.colNum=colNum; AV.colName=colName;
  // celda: {v} número | {v:'texto'} | {f, v} fórmula (v = valor guardado; texto si es string)
  function cellXml(ref,s,c){ const sa=s!=null?` s="${s}"`:''; if(!c||(c.v==null&&c.f==null)) return `<c r="${ref}"${sa}/>`;
    if(c.f!=null){ const isS=typeof c.v==='string'; return `<c r="${ref}"${sa}${isS?' t="str"':''}><f>${xesc(c.f)}</f>${c.v!=null?`<v>${isS?xesc(c.v):c.v}</v>`:''}</c>`; }
    if(typeof c.v==='number') return `<c r="${ref}"${sa}><v>${c.v}</v></c>`;
    return `<c r="${ref}"${sa} t="inlineStr"><is><t xml:space="preserve">${xesc(c.v)}</t></is></c>`; }
  function rowXml(R){ const cells=R.cells.slice().sort((a,b)=>colNum(a.c)-colNum(b.c)).map(c=>cellXml(c.c+R.r,c.s,c)).join('');
    return `<row r="${R.r}"${R.ht?` ht="${R.ht}" customHeight="${R.ch?'true':'false'}"`:''}>${cells}</row>`; }
  // rangos compactos: [5,9,10,11] → "5 9:11" para una columna
  const ranges=(col,rows)=>{ const o=[]; let a=null,b=null; rows.slice().sort((x,y)=>x-y).forEach(r=>{ if(a==null){ a=b=r; } else if(r===b+1) b=r; else { o.push(a===b?col+a:`${col}${a}:${col}${b}`); a=b=r; } }); if(a!=null) o.push(a===b?col+a:`${col}${a}:${col}${b}`); return o; };
  AV.ranges=ranges;
  // sustituye sheetData y los rangos que dependen del número de filas en el XML de la hoja de la plantilla
  AV.sheetXml=function(tpl,S){ let x=tpl.replace(/<sheetData>[\s\S]*<\/sheetData>|<sheetData\/>/,'<sheetData>'+S.rows.map(rowXml).join('')+'</sheetData>');
    const lastCol=S.lastCol||'AT', maxR=Math.max(...S.rows.map(r=>r.r));
    x=x.replace(/<dimension ref="[^"]*"\/>/,`<dimension ref="A1:${lastCol}${maxR}"/>`);
    if(S.merges){ x=x.replace(/<mergeCells[\s\S]*?<\/mergeCells>/,`<mergeCells count="${S.merges.length}">`+S.merges.map(m=>`<mergeCell ref="${m}"/>`).join('')+'</mergeCells>'); }
    if(S.subst) S.subst.forEach(([re,rep])=>{ x=x.replace(re,rep); });
    return x; };

  /* ---------- hoja BOM … CONNECTION ---------- */
  AV.bomSheet=function(M,o){ const T=o.tpl, rows=[]; let r=4;
    const stOf=k=>{ const m={}; T[k].c.forEach(c=>m[c[0]]=c[1]); return m; };
    const add=(k,vals,ht)=>{ const st=stOf(k), cols=new Set(T[k].c.map(c=>c[0]).concat(Object.keys(vals))); const cells=[...cols].map(c=>Object.assign({c,s:st[c]!=null?st[c]:(T[k].c[0]||[])[1]},vals[c]||{}));
      rows.push({r,ht:ht||T[k].ht,ch:ht?true:T[k].ch,cells}); return r++; };
    // encabezados (filas 1–3) de la plantilla; el título con el proyecto
    ['h1','h2','h3'].forEach((k,i)=>{ const cells=T[k].c.map(c=>({c:c[0],s:c[1],v:c[2]!=null?(k==='h1'&&c[0]==='A'?o.title:c[2]):null})); rows.push({r:i+1,ht:T[k].ht,ch:T[k].ch,cells}); });
    const R1=[],V1=[], groups=[]; let nAsm=0, nPieces=0; const looseNotes=[];
    const W=n=>Math.round(n*10)/10, round2=n=>Math.round(n*100)/100;
    const spec=(M[0]&&M[0].spec)||{grade:'A325N',dia:'3/4"'}; const bTxt=(n,to,an,na)=>n?`Bolted ${spec.grade} Ø${spec.dia} ×${n}`+(to.length?` · to ${to.join(', ')}`:'')+(na&&an?` · anchored ${an.grade} Ø${an.dia} ×${na}`:''):(na&&an?`Anchored ${an.grade} Ø${an.dia} ×${na}`:'—');
    M.forEach(x=>{ const a=x.a, main=x.parts.find(p=>p.main)||x.parts[0]; if(!main) return; nAsm++; nPieces+=a.qty;
      const wp=x.parts.filter(p=>p.welded).sort((p,q)=>natCmp(p.r.mark,q.r.mark)), hp=x.parts.filter(p=>p.loose).sort((p,q)=>natCmp(p.r.mark,q.r.mark));
      const looseQty=hp.reduce((t,p)=>t+p.r.qty,0); nPieces+=looseQty; hp.forEach(p=>looseNotes.push(`${p.r.mark} ×${p.r.qty} under ${a.mark}`));
      const rA=r, p0=r+1, nP=1+wp.length, plP=rA+nP, pl=plP+hp.length; const L=x.comp?'COMPOSITE':'LOOSE';
      const roles=[]; x.ifcOrder.forEach(k=>{ const p=wp.find(q=>q.k===k); if(!p) return; const t=p.role==='welded beam stub'?'welded beam':p.role; if(!roles.includes(t)) roles.push(t); });
      wp.forEach(p=>{ const t=p.role==='welded beam stub'?'welded beam':p.role; if(!roles.includes(t)) roles.push(t); });
      const lq=looseQty?` · ${looseQty} non-welded parts counted as LOOSE pieces`:'';
      const O=x.comp?`Shop welded (fillet): ${roles.join(', ')}${lq}`:`No welding${lq}`;
      const sumI=[main,...wp,...hp].reduce((t,p)=>t+W(p.r.qty*p.r.pwt),0), sumK=[main,...wp,...hp].reduce((t,p)=>t+round2(p.r.qty*p.r.sa1),0), sumIP=[main,...wp].reduce((t,p)=>t+W(p.r.qty*p.r.pwt),0);
      const sumAQ=wp.reduce((t,p)=>t+W(p.weld),0);
      add(x.comp?'AC':'AL',{A:{v:a.sno},B:{v:a.mark},D:{v:a.qty},E:{v:a.desc||''},G:{v:(a.row&&a.row.lenText)||''},
        H:{f:`ROUND(SUM(I${p0}:I${pl})/D${rA},1)`,v:W(sumI/a.qty)},I:{f:`SUM(I${p0}:I${pl})`,v:W(sumI)},J:{f:`ROUND(K${rA}/D${rA},2)`,v:round2(sumK/a.qty)},K:{f:`SUM(K${p0}:K${pl})`,v:round2(sumK)},
        L:{v:L},M:{v:x.comp?'YES':'NO'},N:{v:x.nodes},O:{v:O},P:{v:'—'},Q:{v:bTxt(x.perInst,x.to,spec.anchor,x.anchors)},
        S:{f:`IF(AF${rA}>=0.9999,"☑","☐")`,v:'☐'},W:x.comp?{f:`IF(AND(AF${rA}>=0.9999,AG${rA}>=0.9999),"☑","☐")`,v:'☐'}:{v:'N/A'},
        AC:{f:`SUM(I${p0}:I${plP})`,v:W(sumIP)},AD:{f:`SUM(AA${p0}:AA${plP})`,v:0},AE:{f:`SUM(AB${p0}:AB${plP})`,v:0},AF:{f:`IFERROR(AD${rA}/AC${rA},0)`,v:0},AG:{f:`IFERROR(AE${rA}/AC${rA},0)`,v:0},
        AH:{f:`'DAILY REPORT'!$C$5*AF${rA}+'DAILY REPORT'!$C$6*AG${rA}`,v:0},AI:{f:`ROUNDDOWN(D${rA}*MIN(AF${rA},AG${rA})+0.0001,0)`,v:0},AJ:{f:`IFERROR(AD${rA}/$I$@TOT@,0)`,v:0},AK:{f:`IFERROR(AE${rA}/$I$@TOT@,0)`,v:0},
        AL:{f:`'DAILY REPORT'!$C$5*AJ${rA}`,v:0},AM:{f:`'DAILY REPORT'!$C$6*AK${rA}`,v:0},AN:{f:`AL${rA}+AM${rA}`,v:0},AO:{v:L},AP:{v:'A'},
        AQ:{f:`SUM(AQ${p0}:AQ${pl})`,v:W(sumAQ)},AR:{f:`SUM(AR${p0}:AR${pl})`,v:0},AS:{f:`AQ${rA}-AR${rA}`,v:W(sumAQ)}}, x.comp?null:'51');
      const base=p=>{ const q=p.r, rr=r; return {A:{v:q.sno},B:{v:q.mark},D:{v:q.qty},E:{v:q.desc||q.D||''},F:{v:q.grade||''},G:{v:q.lenText||''},H:{v:q.pwt},I:{f:`ROUND(D${rr}*H${rr},1)`,v:W(q.qty*q.pwt)},J:{v:q.sa1},K:{f:`ROUND(D${rr}*J${rr},2)`,v:round2(q.qty*q.sa1)},
        S:{f:`IF(AND(ISNUMBER(R${rr}),R${rr}>=D${rr}),"☑","☐")`,v:'☐'},AA:{f:`I${rr}*IFERROR(MIN(1,R${rr}/D${rr}),0)`,v:0},AO:{v:L},AP:{v:'P'}}; };
      const na=()=>({V:{v:'N/A'},W:{v:'N/A'},X:{v:'N/A'},Y:{v:'N/A'},AB:{f:`AA${r}`,v:0}});
      const weldF=()=>({W:{f:`IF(AND(ISNUMBER(V${r}),V${r}>=D${r}),"☑","☐")`,v:'☐'},AB:{f:`I${r}*IFERROR(MIN(1,V${r}/D${r}),0)`,v:0}});
      // principal
      if(x.comp){ R1.push(r); V1.push(r); add('PM',Object.assign(base(main),{L:{v:'Main member'},M:{v:'MAIN'},O:{v:'Receives welded parts'},P:{v:'—'}},weldF())); }
      else { R1.push(r); add('PL',Object.assign(base(main),{L:{v:'Main member'},M:{v:'NO'},O:{v:hp.length?'No welding · receives ship-loose parts (shop bolted)':'No welding'},P:{v:'—'}},na())); }
      wp.forEach(p=>{ R1.push(r); V1.push(r); const w=W(p.weld); add('PW',Object.assign(base(p),{L:{v:'Welded part'},M:{v:'YES'},O:{v:`Shop fillet weld ${p.fillet} – ${p.role}`},P:{v:a.mark},
        AQ:{v:w},AR:{f:`IF(N(AQ${r})=0,"",AQ${r}*IFERROR(MIN(1,N(V${r})/D${r}),0))`,v:w?0:''},AS:{f:`IF(N(AQ${r})=0,"",AQ${r}-AR${r})`,v:w?w:''},AT:{v:p.fillet}},weldF())); });
      hp.forEach(p=>{ R1.push(r); const li=p.li||{shop:0,field:0,to:[]}; const q=p.r;
        add('H',Object.assign(base(p),{L:{v:'LOOSE'},M:{v:'NO'},N:{v:'—'},O:{v:`LOOSE piece – not welded, fabricated and shipped as an individual piece · shop bolted to ${a.mark} (${spec.grade} Ø${spec.dia} ×${li.shop})`},P:{v:'—'},
          Q:{v:bTxt(li.field,li.to)},AC:{f:`SUM(I${r}:I${r})`,v:W(q.qty*q.pwt)},AD:{f:`SUM(AA${r}:AA${r})`,v:0},AE:{f:`SUM(AB${r}:AB${r})`,v:0},AF:{f:`IFERROR(AD${r}/AC${r},0)`,v:0},AG:{f:`IFERROR(AE${r}/AC${r},0)`,v:0},
          AH:{f:`'DAILY REPORT'!$C$5*AF${r}+'DAILY REPORT'!$C$6*AG${r}`,v:0},AI:{f:`ROUNDDOWN(D${r}*MIN(AF${r},AG${r})+0.0001,0)`,v:0},AJ:{f:`IFERROR(AD${r}/$I$@TOT@,0)`,v:0},AK:{f:`IFERROR(AE${r}/$I$@TOT@,0)`,v:0},
          AL:{f:`'DAILY REPORT'!$C$5*AJ${r}`,v:0},AM:{f:`'DAILY REPORT'!$C$6*AK${r}`,v:0},AN:{f:`AL${r}+AM${r}`,v:0},AO:{v:'LOOSE'},AP:{v:'H'}},na())); });
      groups.push({rA,p0,plP,pl,comp:x.comp,loose:hp.length});
      add('blank',{}); add('blank',{}); });
    const last=r-1, TOT=r+1, rng=c=>`${c}4:${c}${last}`, sif=(c,t)=>`SUMIFS(${rng(c)},${rng('AP')},"${t}")`;
    const totI=rows.filter(R=>R.cells.some(c=>c.c==='AP'&&c.v==='A')).reduce((t,R)=>t+(R.cells.find(c=>c.c==='I').v||0),0);
    const totK=rows.filter(R=>R.cells.some(c=>c.c==='AP'&&c.v==='A')).reduce((t,R)=>t+(R.cells.find(c=>c.c==='K').v||0),0);
    const totAQ=rows.filter(R=>R.cells.some(c=>c.c==='AP'&&c.v==='A')).reduce((t,R)=>t+(R.cells.find(c=>c.c==='AQ').v||0),0);
    const totIH=totI; // AC de A + H = peso de P (sin H) + H
    const acA=rows.filter(R=>R.cells.some(c=>c.c==='AP'&&(c.v==='A'||c.v==='H'))).reduce((t,R)=>t+(R.cells.find(c=>c.c==='AC').v||0),0);
    r=TOT;
    const ah=c=>({f:`${sif(c,'A')}+${sif(c,'H')}`,v:0});
    add('T1',{A:{v:`Total Weight for ${nAsm} Assemblies / ${nPieces} Pieces (lb)`},I:{f:sif('I','A'),v:W(totI)},AA:{f:`SUM(${rng('AA')})`,v:0},AB:{f:`SUM(${rng('AB')})`,v:0},
      AC:{f:`${sif('AC','A')}+${sif('AC','H')}`,v:W(acA)},AD:ah('AD'),AE:ah('AE'),AF:{f:`IFERROR(AD${TOT}/AC${TOT},0)`,v:0},AG:{f:`IFERROR(AE${TOT}/AC${TOT},0)`,v:0},
      AH:{f:`'DAILY REPORT'!$C$5*AF${TOT}+'DAILY REPORT'!$C$6*AG${TOT}`,v:0},AI:ah('AI'),AJ:ah('AJ'),AK:ah('AK'),AL:ah('AL'),AM:ah('AM'),AN:ah('AN'),
      AQ:{f:sif('AQ','A'),v:W(totAQ)},AR:{f:sif('AR','A'),v:0},AS:{f:sif('AS','A'),v:W(totAQ)},AT:{f:`IFERROR(AR${TOT}/AQ${TOT},0)`,v:0}});
    add('T2',{A:{v:`Total Surface Area for ${nAsm} Assemblies (sq-ft)`},K:{f:sif('K','A'),v:round2(totK)}});
    r++;
    const cnt=(lab,t)=>{ const sel=rows.filter(R=>R.cells.some(c=>c.c==='AP'&&(c.v==='A'||c.v==='H'))&&R.cells.some(c=>c.c==='L'&&c.v===t));
      add('T3',{A:{v:lab},I:{f:`SUMIFS(${rng('D')},${rng('AP')},"A",${rng('L')},"${t}")+SUMIFS(${rng('D')},${rng('AP')},"H",${rng('L')},"${t}")`,v:sel.reduce((s,R)=>s+R.cells.find(c=>c.c==='D').v,0)},
        J:{v:'marks:'},K:{f:`COUNTIFS(${rng('AP')},"A",${rng('L')},"${t}")+COUNTIFS(${rng('AP')},"H",${rng('L')},"${t}")`,v:sel.length}}); };
    cnt('Loose pieces (no welding, shipped to field as individual pieces)','LOOSE'); cnt('Composite pieces (shop welded)','COMPOSITE');
    r+=2;
    const sb=o.shopBolt||`${spec.grade} Ø${spec.dia}`;
    const notes=[`Improvement over the original BOM: columns L to Q (connection type read from the Tekla IFC model + Bolt Summary${o.boltDoc?' '+o.boltDoc:''} and 3D bolts).`,
      'Structure and quantities of the original engineering BOM (MTO) are kept: same assemblies, same parts, same weights.',
      'PIECE: COMPOSITE = main member with parts WELDED onto it. Every part that is NOT welded to the main member is a LOOSE piece on its own line (column L = LOOSE, bold, row type H): it keeps its place under its assembly (original BOM structure), is fabricated and shipped as an individual piece and has its own 3D ID.'+(looseNotes.length?` Here: ${looseNotes.join(' · ')} (clip angles, shop bolted with ${sb}).`:''),
      'WELDED (YES/NO): on the assembly row, whether the piece needs shop welding. On each part: YES = welded; WELDED TO = the part it is welded to.',
      'Welds were not exported in the IFC (Welds:Off). Weld lines, count and length were derived from the contact faces between parts; fillet size = 1/4" (5/16" on the 1" base plates). Column AQ = weld inches of each welded part line (all its quantity) from the weld map. Confirm with shop drawings.',
      'DAILY TRACKING (columns R to AN): every day enter on each line the quantity fitted and welded, the name and the date; on the piece row, the PIECE DIE STAMP. The ✔ are set automatically.',
      'Part weights and areas (columns H to K) are the values of the engineering BOM (Tekla report); assembly rows add up their parts.',
      'Shop-bolted and loose lines: no welding (N/A); their welded weight equals the fitted weight (bolted = fitted).',
      '% PROGRESS PIECE = 60 % fit + 40 % weld (weights editable in DAILY REPORT). % BATCH = contribution of the piece to the project total.'];
    notes.forEach(t=>add('N',{A:{v:t}}));
    rows.forEach(R=>R.cells.forEach(c=>{ if(c.f) c.f=c.f.replace(/@TOT@/g,TOT); }));
    // fusiones, validaciones, formato condicional, filtro
    const merges=['A1:K1','L1:Q1','R1:AN1','AQ1:AT1','A2:A3','B2:C3','D2:D3','E2:E3'].concat(['L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','AA','AB','AC','AD','AE','AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR','AS','AT'].map(c=>`${c}2:${c}3`));
    for(let i=4;i<=last;i++) merges.push(`B${i}:C${i}`);
    const dvq=[], dvd=[]; groups.forEach(g=>{ dvq.push(g.comp?`R${g.p0}:R${g.pl}`.replace(/R(\d+):R\1$/,'R$1'):`R${g.p0}:R${g.pl}`.replace(/R(\d+):R\1$/,'R$1')); if(g.comp) dvq.push(`V${g.p0}:V${g.plP}`.replace(/V(\d+):V\1$/,'V$1'));
      dvd.push(`U${g.p0}:U${g.pl}`.replace(/U(\d+):U\1$/,'U$1')); if(g.comp) dvd.push(`Y${g.p0}:Y${g.plP}`.replace(/Y(\d+):Y\1$/,'Y$1')); });
    const L4=last;
    const subst=[[/<autoFilter ref="[^"]*"/,`<autoFilter ref="A3:AT${L4}"`],
      [/sqref="R4:R\d+"/,`sqref="R4:R${L4}"`],[/sqref="V4:V\d+"/,`sqref="V4:V${L4}"`],[/sqref="S4:S\d+ W4:W\d+"/,`sqref="S4:S${L4} W4:W${L4}"`],[/sqref="AH4:AH\d+"/g,`sqref="AH4:AH${L4}"`],[/sqref="AF4:AG\d+"/g,`sqref="AF4:AG${L4}"`],
      [/<xm:sqref>AH4:AH\d+<\/xm:sqref>/,`<xm:sqref>AH4:AH${L4}</xm:sqref>`],[/<xm:sqref>AF4:AG\d+<\/xm:sqref>/,`<xm:sqref>AF4:AG${L4}</xm:sqref>`],
      [/(<dataValidation [^>]*?type="whole"[^>]*?)sqref="[^"]*"|(<dataValidation [^>]*?)sqref="[^"]*"([^>]*?type="whole")/,(m,a,b,c)=>a?`${a}sqref="${dvq.join(' ')}"`:`${b}sqref="${dvq.join(' ')}"${c}`],
      [/(<dataValidation [^>]*?)sqref="[^"]*"([^>]*?type="date")/,(m,a,c)=>`${a}sqref="${dvd.join(' ')}"${c}`]];
    return {rows,merges,subst,last,TOT,filter:`$A$3:$AT$${L4}`,lastCol:'AT'}; };

  // nombres del proyecto para títulos y hojas: «BOP_WWTP PDC Piperacks_30 m» → «WWTP PDC Piperacks 30m»
  AV.names=function(meta,marks,boltMeta){ const proj=String(meta&&(meta.project||meta.model)||'').replace(/^BOP_/i,'').replace(/_/g,' ').replace(/(\d+)\s+m\b/gi,'$1m').replace(/\s+/g,' ').trim();
    const doc=String(meta&&meta.bom||'').replace(/^_+/,''); const pre=(String((marks||[])[0]||'').match(/^([A-Za-z0-9]+)-/)||[])[1]||''; const len=(proj.match(/(\d+)m\b/i)||[])[1];
    return {proj,doc,title:`ASSEMBLY LIST WITH PARTS (BOM) – ${proj}${doc?` (${doc})`:''}`,sheet:`BOM ${pre}${len?'-'+len+'M':''} - CONNECTION`.replace(/\s+/g,' '),boltDoc:((boltMeta&&boltMeta.title||'').match(/BOM-\d+/)||[''])[0]}; };
  AV._faces=facesOf; AV._cw=contactWeld;
  G.AV=AV;
})(typeof window!=='undefined'?window:globalThis);
