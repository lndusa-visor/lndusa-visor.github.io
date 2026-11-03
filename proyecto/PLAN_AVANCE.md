# Generador del Excel de avance (BOM_Connection_Type) — plan y estado

Meta: desde BOM (Tekla, .xls) + IFC + Bolt Summary de cada unidad, generar el Excel de avance
idéntico en estructura y fórmulas al de los Piperacks 30 m / 55 m (archivos de referencia del compañero),
más lo nuevo: Stamp BOM con letra, LISTA DE TROQUELADOS, PR SAP y PLAN DE CORTE.
Después: el tablero (Weld Map) arma su 3D (SCENES) desde el IFC.

Referencias (sesión actual): /tmp/claude-0/wm/b30.xlsx, b55.xlsx (Excel del compañero),
/tmp/claude-0/np/*.xls (BOM), *Bolt_Summary*.xlsx, /tmp/claude-0/ifc/u2/*.ifc, /tmp/claude-0/wm/wm.html (tablero).

## Hecho
- proyecto/index.html: lista de troquelados (1 hoja, QR por pieza), PR SAP única, publicar visor, «Agregar al Excel de avance».
- Tablero del compañero + PR SAP en Material (archivo aparte, no publicado en GitHub).

## Hojas del Excel (fórmulas por hoja: total / formas distintas)
BOM CONNECTION 1255/89 · CUTTING TRACKING 1756/242 · DAILY REPORT 1043/37 · WEEKLY NOTES 40/1 ·
CONNECTION SUMMARY 13/10 · VENDOR REQUISITION 139/45 · PURCHASE ORDERS 1603/11 · WELDING PROGRESS 1265/469 ·
PAINTING PROGRESS 1051/18 · 3D CONTROL 6668/1931 · WEEKLY CUT-OFFS 512/20 · TAKEOFF BASE 286/7 ·
MTO SUMMARY 84/50 · PROFILE CUTTING PLAN 177/30 · PLATE NESTING 36/19

## BOM CONNECTION (estructura observada)
- Por ensamble: fila A (resumen, fórmulas SUM de sus partes), filas P (main member primero, luego partes soldadas),
  filas H (partes sueltas apernadas, p. ej. clip angles) al final del grupo; 2 filas vacías entre grupos.
- L: LOOSE/COMPOSITE (A), Main member / Welded part (P), LOOSE (H). M: YES/NO/MAIN. N: nodos (compuestas numeradas 1..n por instancia, «4–27»).
- O: texto de soldadura con tipo de placa (end plate, shear plate, gusset, welded beam stub). P: soldada a. Q: pernos de campo «Bolted A325N Ø3/4" ×N · to …».
- AQ: pulgadas de soldadura por pieza (geometría de contacto del IFC), AT: filete (1/4", 5/16" en placas base de 1").
- Fila total: «Total Weight for N Assemblies / M Pieces (lb)» con SUMIFS por AP.

## Estado del generador (proyecto/avance-gen.js)
- AV.model: nodos, piezas soldadas/sueltas, pulgadas de soldadura por contacto (reglas: franja delgada = 2 lados largos; perfil = contorno sin puntas ≤ 1/4";
  contacto entre dos partes soldadas → a la que va después en el ensamble del IFC), pernos de campo (instancia de menor X / mayor Y, tolerancia 30 mm, anclajes aparte).
- AV.bomSheet + AV.sheetXml: hoja BOM … CONNECTION sobre la plantilla (estilos en avance-tpl-bom.json).
- Comparación celda por celda: 30 m → 3 diferencias, 55 m → 2 diferencias (de ~10 000 celdas):
  · end plate de B9 (30 m) / B6 (55 m): 61 vs 63 in (la referencia cuenta las puntas del ala aquí pero no en B2 con la misma sección);
  · gusset p10 (30 m) / p7 (55 m) sobre ángulo: 32.8 vs 24.8 in (la referencia no cuenta los extremos de 4");
  · nota 181 del 30 m: la referencia nombra el Point-to-Point viejo (BOM-0111); el generador nombra el Bolt Summary cargado.

## Siguiente
1. Generador JS (Node + navegador) hoja por hoja, comparando celda por celda contra b30/b55.
2. Orden: BOM CONNECTION → TAKEOFF/MTO/CUTTING PLAN/NESTING → VENDOR REQ/PO → CUTTING TRACKING → WELDING → PAINTING → 3D CONTROL → DAILY/WEEKLY.
3. Tablero: SCENES (partes, soldaduras, pernos, ptp, grid, info) desde el IFC.
