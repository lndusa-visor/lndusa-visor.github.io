#!/usr/bin/env python3
"""Arma proyecto/avance-plantilla.xlsx desde un Excel de avance terminado: quita datos del proyecto,
las pestañas agregadas y los resultados guardados. Conserva diseño, estilos, fórmulas fijas, gráficas y reglas.
Uso: avance_plantilla.py entrada.xlsx salida.xlsx"""
import re,sys,zipfile
src,dst=sys.argv[1:3]
z=zipfile.ZipFile(src); OLD='BOM WTS-55M - CONNECTION'; NEW='BOM TPL - CONNECTION'
GEN={1,2,6,7,8,9,10,12,13,14,15}; KEEP={3,4,5,11}; DROP={16,17,18,19}
wb=z.read('xl/workbook.xml').decode(); rels=z.read('xl/_rels/workbook.xml.rels').decode(); ct=z.read('[Content_Types].xml').decode()
files={}
drop=set(n for n in z.namelist() if n.startswith('[trash]') or n.startswith('xl/media/'))
for i in DROP:
    drop.add(f'xl/worksheets/sheet{i}.xml'); r=f'xl/worksheets/_rels/sheet{i}.xml.rels'; drop.add(r)
    for d in re.findall(r'drawings/(drawing\d+\.xml)',z.read(r).decode() if r in z.namelist() else ''):
        drop.add('xl/drawings/'+d); drop.add(f'xl/drawings/_rels/{d}.rels')
for i in DROP:
    rid=re.search(rf'Id="(rId\d+)"[^>]*Target="worksheets/sheet{i}\.xml"|Target="worksheets/sheet{i}\.xml"[^>]*Id="(rId\d+)"',rels); rid=rid[1] or rid[2]
    wb=re.sub(rf'<sheet [^>]*r:id="{rid}"[^>]*/>','',wb); rels=re.sub(rf'<Relationship [^>]*worksheets/sheet{i}\.xml"[^>]*/>','',rels)
for d in drop: ct=re.sub(rf'<Override [^>]*PartName="/{re.escape(d)}"[^>]*/>','',ct)
wb=wb.replace(OLD,NEW)
ct=re.sub(r'<Default Extension="(png|jpe?g)"[^>]*/>','',ct)
sst=z.read('xl/sharedStrings.xml').decode(); head=sst[:sst.index('<si>')]; items=re.findall(r'<si>.*?</si>',sst,re.S); used=set()
def strip(x):
    x=re.sub(r'(<c r="[A-Z]+\d+"[^>]*>)(<f[^>]*>[^<]*</f>)<v>[^<]*</v>',lambda m:re.sub(r' t="(?:n|str|b|e)"','',m[1])+m[2],x); return x.replace(OLD,NEW)
for i in range(1,16):
    n=f'xl/worksheets/sheet{i}.xml'; x=z.read(n).decode()
    if i in GEN: x=re.sub(r'<sheetData>.*</sheetData>','<sheetData/>',x,flags=re.S)
    else:
        x=strip(x); used|={int(k) for k in re.findall(r't="s"[^>]*><v>(\d+)</v>',x)}
    files[n]=x
def clean(t):
    for a,b in ((r'WTS-a1 ×308 under WTS-B3','X-a1 ×0 under X-B1'),(r'WTS-a1 ×308 bajo WTS-B3','X-a1 ×0 bajo X-B1'),(r'WTS-B3','X-B1'),(r'WTS PDC Piperacks 55m','PROJECT'),(r'BOM-0123','BOM-0000'),(r'BOM-0122','BOM-0000'),(r'\b97 HILTI HIT Ø1&quot; x 10&quot;','0 ANCHOR'),(r'\b97 anclas adhesivas HILTI HIT Ø1&quot; x 10&quot;','0 anclas adhesivas ANCHOR'),(r'HILTI HIT Ø1&quot; x 10&quot;','ANCHOR'),(r'has 96 \(24 base plates × 4\)','has 0 (0 base plates × 0)'),(r'tiene 96 \(24 placas base × 4\)','tiene 0 (0 placas base × 0)'),(r'IFC 96 \(24 base plates × 4\) · Bolt Summary 97','IFC 0'),(r'^(<si><t[^>]*>)24 ','\\g<1>0 '),(r'\(DND0\)','(X)')):
        t=re.sub(a,b,t)
    return t
out=[clean(s) if i in used else '<si><t></t></si>' for i,s in enumerate(items)]
files['xl/sharedStrings.xml']=head+''.join(out)+'</sst>'
leak=[re.sub('<[^>]+>','',s)[:110] for s in out if re.search(r'WTS|WWTP|Piperack|US\.?0|BOM-0[1-9]|HILTI|Guyana|GUYANA',s)]
print('textos con datos que quedan:',leak)
files['xl/workbook.xml']=wb; files['xl/_rels/workbook.xml.rels']=rels; files['[Content_Types].xml']=ct
o=zipfile.ZipFile(dst,'w',zipfile.ZIP_DEFLATED)
for n in z.namelist():
    if n in drop or n.endswith('/'): continue
    b=z.read(n)
    if n in files: b=files[n].encode()
    elif n.startswith('xl/charts/chart') and n.endswith('.xml'):
        b=re.sub(r'<c:(num|str)Cache>.*?</c:\1Cache>','',b.decode(),flags=re.S).encode()
    o.writestr(n,b)
o.close(); print('plantilla',len(o.namelist()) if False else 'ok')
