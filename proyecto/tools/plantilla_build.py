#!/usr/bin/env python3
"""Arma proyecto/plantilla.html: el tablero vacío con «Nuevo proyecto» adentro, para guardarlo en SharePoint.
Uso (desde proyecto/): python3 tools/plantilla_build.py
Mete en un solo archivo la página index.html, avance-*.js, las plantillas de Excel, del tablero y del visor.
Con eso, GitHub solo se usa para publicar y abrir los visores de los QR.
Hay que correrlo cada vez que cambie cualquiera de esos archivos; tablero-virgen.html se edita directo."""
import base64, gzip, json, re, sys
b64 = lambda b: base64.b64encode(b).decode()
rd = lambda p: open(p, encoding='utf8').read()
page = rd('index.html')
for js in ('avance-gen.js', 'avance-mto.js', 'avance-scene.js'):
    tag = f'<script src="{js}"></script>'
    code = rd(js)
    if page.count(tag) != 1 or '</script' in code.lower(): sys.exit('No pude incrustar ' + js)
    page = page.replace(tag, '<script>\n' + code + '\n</script>', 1)
emb = {'xlsx': b64(open('avance-plantilla.xlsx', 'rb').read()),
       'bom': b64(open('avance-tpl-bom.json', 'rb').read()),
       'tablero': b64(open('tablero-virgen.html', 'rb').read()),
       'visor': b64(open('../visor/index.html', 'rb').read())}
head = '<script>window.EMB=' + json.dumps(emb) + '</script>\n'
first = page.index('<script')
page = page[:first] + head + page[first:]
tablero = rd('tablero-virgen.html')
if tablero.count('<!--NP-PACK-->') != 1: sys.exit('Falta la marca NP-PACK')
pack = b64(gzip.compress(page.encode('utf8'), 9, mtime=0))
out = tablero.replace('<!--NP-PACK-->', '<script type="text/plain" id="np-pack">' + pack + '</script>')
open('plantilla.html', 'w', encoding='utf8').write(out)
print('plantilla.html', round(len(out.encode()) / 1024), 'KB')
