(function(G){
const ECM=[null,[26,10,[[1,16]]],[44,16,[[1,28]]],[70,26,[[1,44]]],[100,18,[[2,32]]],[134,24,[[2,43]]],[172,16,[[4,27]]],[196,18,[[4,31]]],[242,22,[[2,38],[2,39]]],[292,22,[[3,36],[2,37]]],[346,26,[[4,43],[1,44]]]];
const ALIGN=[null,[],[6,18],[6,22],[6,26],[6,30],[6,34],[6,22,38],[6,24,42],[6,26,46],[6,28,50]];
const EXP=new Array(512), LOG=new Array(256);
(function(){ let x=1; for(let i=0;i<255;i++){ EXP[i]=x; LOG[x]=i; x<<=1; if(x&256) x^=0x11D; } for(let i=255;i<512;i++) EXP[i]=EXP[i-255]; })();
const mul=(a,b)=>a&&b?EXP[LOG[a]+LOG[b]]:0;
function rsGen(n){ let g=[1]; for(let i=0;i<n;i++){ const ng=new Array(g.length+1).fill(0); for(let j=0;j<g.length;j++){ ng[j]^=g[j]; ng[j+1]^=mul(g[j],EXP[i]); } g=ng; } return g; }
function rsEnc(data,n){ const g=rsGen(n), res=new Array(n).fill(0); for(const d of data){ const f=d^res[0]; res.shift(); res.push(0); for(let i=0;i<n;i++) res[i]^=mul(g[i+1],f); } return res; }
function utf8(s){ return Array.from(new TextEncoder().encode(s)); }
function qrMatrix(text){
  const bytes=utf8(text); let ver=0;
  for(let v=1;v<=10;v++){ const [tot,ec,bl]=ECM[v]; const dataCw=bl.reduce((a,[n,k])=>a+n*k,0); const cc=v<10?8:16; if(4+cc+8*bytes.length<=dataCw*8){ ver=v; break; } }
  if(!ver) throw new Error('Texto demasiado largo para el QR');
  const [tot,ecLen,bl]=ECM[ver], dataCw=bl.reduce((a,[n,k])=>a+n*k,0), cc=ver<10?8:16;
  const bits=[]; const put=(v,n)=>{ for(let i=n-1;i>=0;i--) bits.push((v>>>i)&1); };
  put(4,4); put(bytes.length,cc); bytes.forEach(b=>put(b,8));
  put(0,Math.min(4,dataCw*8-bits.length)); while(bits.length%8) bits.push(0);
  const cw=[]; for(let i=0;i<bits.length;i+=8){ let b=0; for(let j=0;j<8;j++) b=(b<<1)|bits[i+j]; cw.push(b); }
  for(let p=0xEC;cw.length<dataCw;p^=0xEC^0x11) cw.push(p);
  // bloques + RS, intercalado
  const blocks=[]; let k=0; bl.forEach(([n,len])=>{ for(let i=0;i<n;i++){ const d=cw.slice(k,k+len); k+=len; blocks.push({d,e:rsEnc(d,ecLen)}); } });
  const out=[]; const maxD=Math.max(...blocks.map(b=>b.d.length));
  for(let i=0;i<maxD;i++) blocks.forEach(b=>{ if(i<b.d.length) out.push(b.d[i]); });
  for(let i=0;i<ecLen;i++) blocks.forEach(b=>out.push(b.e[i]));
  const size=ver*4+17, M=[], F=[]; for(let r=0;r<size;r++){ M.push(new Array(size).fill(0)); F.push(new Array(size).fill(false)); }
  const set=(r,c,v)=>{ M[r][c]=v?1:0; F[r][c]=true; };
  const finder=(r0,c0)=>{ for(let r=-1;r<=7;r++) for(let c=-1;c<=7;c++){ const rr=r0+r, cc2=c0+c; if(rr<0||cc2<0||rr>=size||cc2>=size) continue;
    const on=(r>=0&&r<=6&&(c===0||c===6))||(c>=0&&c<=6&&(r===0||r===6))||(r>=2&&r<=4&&c>=2&&c<=4); set(rr,cc2,on); } };
  finder(0,0); finder(0,size-7); finder(size-7,0);
  for(let i=8;i<size-8;i++){ set(6,i,i%2===0); set(i,6,i%2===0); }
  const al=ALIGN[ver]; for(const r of al) for(const c of al){ if((r===6&&c===6)||(r===6&&c===size-7)||(r===size-7&&c===6)) continue;
    for(let dr=-2;dr<=2;dr++) for(let dc=-2;dc<=2;dc++) set(r+dr,c+dc,Math.max(Math.abs(dr),Math.abs(dc))!==1); }
  // reservar formato y versión
  for(let i=0;i<9;i++){ if(!F[8][i]) set(8,i,0); if(!F[i][8]) set(i,8,0); }
  for(let i=0;i<8;i++){ set(8,size-1-i,0); set(size-1-i,8,0); }
  set(size-8,8,1); // módulo oscuro
  if(ver>=7){ let rem=ver; for(let i=0;i<12;i++) rem=(rem<<1)^((rem>>>11)*0x1F25); const vb=(ver<<12)|rem;
    for(let i=0;i<18;i++){ const b=(vb>>>i)&1, a=size-11+(i%3), c=Math.floor(i/3); set(c,a,b); set(a,c,b); } }
  // datos en zigzag
  const dbits=[]; out.forEach(b=>{ for(let i=7;i>=0;i--) dbits.push((b>>>i)&1); });
  let bi=0;
  for(let right=size-1;right>=1;right-=2){ if(right===6) right=5;
    for(let vert=0;vert<size;vert++) for(let j=0;j<2;j++){ const c=right-j, up=((right+1)&2)===0, r=up?size-1-vert:vert;
      if(!F[r][c]){ M[r][c]=bi<dbits.length?dbits[bi]:0; bi++; } } }
  const MASKS=[(r,c)=>(r+c)%2===0,(r,c)=>r%2===0,(r,c)=>c%3===0,(r,c)=>(r+c)%3===0,(r,c)=>(Math.floor(r/2)+Math.floor(c/3))%2===0,
    (r,c)=>(r*c)%2+(r*c)%3===0,(r,c)=>((r*c)%2+(r*c)%3)%2===0,(r,c)=>((r+c)%2+(r*c)%3)%2===0];
  const drawFormat=(Q,mask)=>{ const data=(0<<3)|mask; let rem=data; for(let i=0;i<10;i++) rem=(rem<<1)^((rem>>>9)*0x537); const fb=((data<<10)|rem)^0x5412;
    const b=i=>(fb>>>i)&1;
    for(let i=0;i<=5;i++) Q[i][8]=b(i); Q[7][8]=b(6); Q[8][8]=b(7); Q[8][7]=b(8); for(let i=9;i<15;i++) Q[8][14-i]=b(i);
    for(let i=0;i<8;i++) Q[8][size-1-i]=b(i); for(let i=8;i<15;i++) Q[size-15+i][8]=b(i); Q[size-8][8]=1; };
  const penalty=Q=>{ let p=0;
    for(let pass=0;pass<2;pass++) for(let i=0;i<size;i++){ let run=1; for(let j=1;j<size;j++){ const a=pass?Q[j][i]:Q[i][j], b=pass?Q[j-1][i]:Q[i][j-1];
      if(a===b){ run++; if(run===5) p+=3; else if(run>5) p++; } else run=1; } }
    for(let r=0;r<size-1;r++) for(let c=0;c<size-1;c++){ const v=Q[r][c]; if(v===Q[r][c+1]&&v===Q[r+1][c]&&v===Q[r+1][c+1]) p+=3; }
    const pat=[1,0,1,1,1,0,1,0,0,0,0], pat2=[0,0,0,0,1,0,1,1,1,0,1];
    for(let i=0;i<size;i++) for(let j=0;j+11<=size;j++){ let h1=true,h2=true,v1=true,v2=true;
      for(let k=0;k<11;k++){ if(Q[i][j+k]!==pat[k]) h1=false; if(Q[i][j+k]!==pat2[k]) h2=false; if(Q[j+k][i]!==pat[k]) v1=false; if(Q[j+k][i]!==pat2[k]) v2=false; }
      p+=40*(h1+h2+v1+v2); }
    let dark=0; Q.forEach(r=>r.forEach(v=>dark+=v)); p+=Math.floor(Math.abs(dark*20/(size*size)-10))*10; return p; };
  let best=null, bestP=Infinity;
  for(let m=0;m<8;m++){ const Q=M.map(r=>r.slice()); for(let r=0;r<size;r++) for(let c=0;c<size;c++) if(!F[r][c]&&MASKS[m](r,c)) Q[r][c]^=1;
    drawFormat(Q,m); const p=penalty(Q); if(p<bestP){ bestP=p; best=Q; } }
  return {size, version:ver, get:(r,c)=>best[r][c]===1};
}
function qrSVG(text,px,qz){
  qz=qz==null?4:qz; const q=qrMatrix(text), n=q.size+2*qz; let d='';
  for(let r=0;r<q.size;r++) for(let c=0;c<q.size;c++) if(q.get(r,c)) d+=`M${c+qz} ${r+qz}h1v1h-1z`;
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${px}" height="${px}" viewBox="0 0 ${n} ${n}" shape-rendering="crispEdges"><rect width="${n}" height="${n}" fill="#fff"/><path d="${d}" fill="#000"/></svg>`;
}
G.qrMatrix=qrMatrix; G.qrSVG=qrSVG;
})(typeof window!=='undefined'?window:globalThis);
