# Generador del Excel de avance (BOM_Connection_Type) — plan y estado

Meta: desde BOM (Tekla, .xls) + IFC + Bolt Summary de cada unidad, generar el Excel de avance
idéntico en estructura y fórmulas al de los Piperacks 30 m / 55 m (archivos de referencia del compañero),
más lo nuevo: Stamp BOM con letra, LISTA DE TROQUELADOS, PR SAP y PLAN DE CORTE.
Después: el tablero (Weld Map) arma su 3D (SCENES) desde el IFC.

Referencias: se adjuntan en cada sesión (no están en GitHub, y no deben subirse: GitHub Pages es público).
Referencia completa: BOM_Connection_Type_WTS_PDC_Piperacks_55m_EN_con_troquel.xlsx (19 hojas).

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

## TAKEOFF BASE, MTO SUMMARY, PROFILE CUTTING PLAN, PLATE NESTING (proyecto/avance-mto.js)
- AV.mtoParts lee las filas P/H de la hoja BOM; AV.mtoSheets arma las 4 hojas. Estilos fijos en el código (mismos índices de styles.xml).
- Comparación contra el Excel de 55 m: 0 diferencias en valor, fórmula, resultado guardado, estilo y alto de fila (3216 celdas). Recalculado en LibreOffice: 0 diferencias, 0 errores.
- Probado de punta a punta en 30 m y 55 m: página real con BOM + IFC + Bolt Summary → AV.model → AV.bomSheet → AV.mtoParts → hojas. 0 diferencias en las 6 hojas.
- Corte: largo descendente (empates en orden del BOM), cada pieza a la barra más justa; kerf por corte, recorte al final; barras iguales juntas, primero las más repetidas. Dentro de la barra: largo descendente y marca como texto. Usado = redondeo a 3 decimales con empate al par; largos con 6 cifras significativas.
- Placas (misma marca con dos largos = piezas distintas): área descendente, Best Short Side Fit, primera lámina donde quepa, lámina útil = lámina − bordes + separación, sin reordenar láminas. Ojo: difiere del nest() de index.html (mismo número de láminas en 55 m, otro acomodo).
- Entran como opciones, no se deducen: largo de barra por perfil (stock) y tamaño de lámina (los elige el HTML); % de repuesto (VENDOR REQUISITION C7). Pernos y anclajes: AV.boltInput lee lo que la página ya leyó del Bolt Summary; AV.anchorModel y AV.shopBoltText completan notas.

## VENDOR REQUISITION y PURCHASE ORDERS (AV.vendorSheets en avance-mto.js)
- Mismo diseño y fórmulas; solo cambian los datos del proyecto. PURCHASE ORDERS sale sin órdenes. Comparadas contra 55 m.
- AV.vendorSheets devuelve mats (clave de material → fila), que necesitará 3D CONTROL. Probadas también de punta a punta.

## CUTTING TRACKING (AV.cuttingSheet en avance-mto.js)
- Una línea por parte, ligada a la hoja BOM; tabla diaria fija de 120 días (filas 11–130, la usa la gráfica). 0 diferencias en 30 m y 55 m desde los archivos de origen. Fecha inicial = DAILY REPORT C7.
- Los números de estilo cambian entre los Excel de 30 y 55: la plantilla final usará un solo styles.xml; la prueba compara la apariencia real de cada celda.

## WELDING PROGRESS (AV.weldingSheet en avance-mto.js)
- Una fila por nodo; tabla diaria fija de 120 días (filas 21–140). AV.model ahora da cordones (weldN) y orden de soldadura (weldSeq). G = suma exacta de pulgadas; F = cordones.
- Comparada desde archivos de origen. Diferencias: F y G en las mismas piezas ya conocidas (B6 / B7, B9: 16 vs 24 y 61 vs 63; VB2: 4 vs 2 y 32.8 vs 24.8), y los totales que dependen de ellas.
- Columna I en 8 columnas de 30 m: 0.1 lb (889.9 vs 889.8, 877.5 vs 877.4). La referencia no coincide con el BOM; fuente aún no encontrada.
- K19, L19 y el valor de la columna L los completa el paso de troqueles.

## PAINTING PROGRESS (AV.paintingSheet en avance-mto.js)
- Una fila por pieza de montaje; piezas sueltas en su propia fila. Fórmulas iguales a la referencia en 30 m y 55 m.

## Regla: el BOM manda y las fórmulas no se tocan
- 9798 fórmulas en 55 m y 7677 en 30 m idénticas a la referencia, salvo los troqueles de la columna Z, que agrega su paso.
- Peso y área unitarios (H, J de BOM CONNECTION; L de TAKEOFF; G, H de PAINTING) llevan los decimales del total del BOM; se ven redondeados por el formato de celda.
- Suma de partes = 57378.7 lb (55 m) / 33897.3 lb (30 m). Suma de ensambles del BOM = 57378.9 / 33898.3. Esa diferencia ya viene dentro del BOM original.

## Decisión sobre soldadura
- El usuario decidió dejar la regla como está: no se ajusta por B6/B7/B9 ni VB2. Esas diferencias contra la referencia quedan aceptadas.

## Control de pruebas
- Los troqueles de la columna Z de BOM CONNECTION no los pone el generador: los agrega el paso «Agregar al Excel de avance». Falta encadenarlo.

## Siguiente
1. Generador JS (Node + navegador) hoja por hoja, comparando celda por celda contra b30/b55.
2. Orden: BOM CONNECTION (hecha) → TAKEOFF/MTO/CUTTING PLAN/NESTING (hechas) → VENDOR REQ/PO (hechas) → CUTTING TRACKING (hecha) → WELDING (hecha) → PAINTING (hecha) → 3D CONTROL → DAILY/WEEKLY.
3. Tablero: SCENES (partes, soldaduras, pernos, ptp, grid, info) desde el IFC.
