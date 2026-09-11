---
name: comparar-apus
description: Compara los APUs (análisis de precios unitarios) y el presupuesto de la oferta de un oferente contra los archivos de referencia del contratante, rubro por rubro, y entrega un Excel con las diferencias graduadas por gravedad. Úsala siempre que el usuario mencione APUs, análisis de precios unitarios, rubros, oferentes, licitación, presupuesto de obra o desagregación/VAE, y quiera verificar, validar, revisar o comparar una oferta contra un archivo base — aunque no diga la palabra "comparar" y aunque los archivos vengan en PDF, en Excel o mezclados. También aplica cuando pide revisar si el oferente "copió bien" los rubros, cambió cantidades, unidades, descripciones o indirectos.
---

# Comparación de APUs: referencia vs. oferente

En una licitación de obra, el contratante publica los análisis de precios
unitarios de cada rubro y el oferente debe reproducirlos usando sus propios
precios. Puede cambiar tarifas, jornales, precios unitarios y rendimientos —
esa es su oferta. Lo que **no** puede cambiar es la composición del rubro:
qué equipos, mano de obra y materiales entran, en qué cantidad y en qué unidad.

Tu trabajo es encontrar dónde se apartó, sin ahogar al revisor en ruido. Un
reporte donde los 137 rubros salen en rojo no sirve para nada: la mitad del
valor está en separar el error real de la diferencia de escritura, y la otra
mitad en separar lo que perjudica al contrato de lo que no.

## El criterio de gravedad

Esto es lo que hay que entender antes que nada, porque decide todo lo demás.

**La oferta no tiene que ser idéntica: tiene que cubrir el rubro.** De ahí sale
una asimetría que a primera vista sorprende y que es el corazón de la revisión:

- Que **falte** un ítem, o que su cantidad venga **por debajo** de la
  referencia, es **grave**. El oferente se comprometió a ejecutar el rubro con
  menos de lo que el contratante estimó necesario.
- Que **sobre** un ítem, o que su cantidad venga **por encima**, es una
  observación **menor**. Está destinando más recursos a ejecutar el rubro; no
  perjudica al contrato.
- Que los ítems estén en **otro orden** es informativo y nada más.

La baja por redondeo no es excepción: `1,005 → 1,00` es menos cantidad y se
observa igual, porque el efecto sobre la obra es el mismo venga de donde venga.
La cantidad que queda **en cero** se señala aparte: es la más grave de todas,
porque el insumo desaparece del costo del rubro sin dejar rastro.

## Qué se compara y qué no

**Sí:** nombre y unidad del rubro, detalle/especificación, la descripción,
cantidad y unidad de cada ítem de EQUIPOS / MANO DE OBRA / MATERIALES /
TRANSPORTE, el porcentaje de indirectos, y —si los dos archivos traen
presupuesto— la cantidad de cada rubro y el precio unitario.

**No:** tarifas, jornales, precios unitarios de insumos, rendimientos y costos.
Son del oferente. Compararlos llena el reporte de ruido y esconde lo que
importa. **El monto de la oferta tampoco se observa**: ya puntúa por sí mismo
en la calificación económica, así que un "+10% de sobrecosto" como titular es
un hallazgo que no le sirve a nadie.

Con una excepción que sí es observable: que el precio unitario del presupuesto
de la oferta **no cuadre con el que sale de su propio APU**. Ahí uno de los dos
documentos no es el que el oferente va a ejecutar. `apu.comparar_presupuesto()`
lo comprueba solo.

### Indirectos

El oferente puede bajar sus indirectos hasta **5 puntos porcentuales** respecto
a los del presupuesto referencial, pero no subirlos. Dentro de esa tolerancia
es una observación menor; fuera —o hacia arriba— es grave.

Los indirectos son **una** decisión que se repite en todos los rubros, así que
se reportan agrupados en una línea que dice cuánto bajó y en qué rubros, y no
entran en el estado de cada rubro. Reportarlo 137 veces taparía todo lo demás.

### Mano de obra y categorías

Renombrar la mano de obra manteniendo el código de estructura ocupacional
(`ESTRUC. OCUPAC. E2 PEON` → `Estr. Oc. E2 PEON`) es el mismo obrero escrito
distinto: activa `nomenclatura_mo_menor=True` cuando veas ese patrón. Pero
**cambiar la categoría** (C3 → C1) no está permitido y sigue siendo grave: es
otro salario y otro perfil.

## Flujo

### 1. Reconocer los dos archivos

```bash
python3 -c "
import openpyxl
wb = openpyxl.load_workbook('archivo.xlsx', read_only=True)
print(len(wb.sheetnames)); print(wb.sheetnames[:25]); print(wb.sheetnames[-8:])"
```

Para PDF: `pdfinfo archivo.pdf | grep Pages`. No hace falta OCR: son Excel
impresos y traen texto real.

**Si el mismo dato viene en Excel y en PDF, usa el Excel.** El oferente entrega
lo que quiere y a menudo manda las dos cosas; la carpeta del contratante casi
siempre tiene el presupuesto en `.xlsx` además del PDF del portal. El Excel
trae el valor de cada celda y el PDF trae un dibujo de esa celda: no hay pelea
posible. Y **la referencia también se lee mal**: en una revisión real, tres de
las siete observaciones rojas venían de haber leído el presupuesto referencial
en PDF (`e=15cm` por `e=5cm`, `70` por `9`, `250` por `225`), no de la oferta.

Cuenta las hojas con cuidado. Un libro puede traer **dos juegos de APUs**: uno
antiguo (`Rubro1`…`Rubro190`) y el vigente (`1`…`195`), más hojas de apoyo. Si
el total no cuadra con el número de rubros, busca el segundo juego antes de
seguir. Elegir el equivocado invalida todo el análisis y no da ninguna señal de
error — solo produce cientos de diferencias falsas.

### 2. Establecer y verificar la correspondencia

Nunca asumas que la hoja `Rubro N` corresponde al rubro `N`. **Verifícalo
comparando los nombres de los 78, 137 o los que sean**:

```bash
python3 -c "
import sys; sys.path.insert(0, '<skill>/scripts')
import lectores, apu
fa, ref = lectores.leer_libro('referencia.xlsx', {i: 'Rubro%d' % i for i in range(1, N+1)})
fb, ofe = lectores.leer_libro('oferta.xlsx',     {i: str(i)         for i in range(1, N+1)})
mal = [(i, ref[i]['nombre'][:40], ofe[i]['nombre'][:40])
       for i in ref if apu.nd(ref[i]['nombre']) != apu.nd(ofe[i]['nombre'])]
print('formatos:', fa, fb, '| desalineados:', len(mal)); print(mal[:10])"
```

Si salen desalineados, no fuerces la comparación: busca la hoja de PRESUPUESTO
de la referencia, que trae el orden real con su código, y enlaza por código en
vez de por posición. Avísale al usuario, porque es un hallazgo por sí mismo.

### 3. Leer los archivos

Excel y PDF se autodetectan, y las dos combinaciones se mezclan sin problema
(referencia en Excel con oferta en PDF es lo más común).

```python
import sys; sys.path.insert(0, '<skill>/scripts')
import lectores, lector_pdf

formato, ref = lectores.leer_libro('referencia.xlsx', {i: str(i) for i in range(1, N+1)})
ofe = lector_pdf.leer_pdf('oferta.pdf')      # detecta preset, decimal y modo

import presupuesto
pres_ref, hoja_r = presupuesto.leer('referencia.xlsx')    # ({}, None) si no lo trae
pres_ofe, hoja_o = presupuesto.leer('oferta.pdf')         # el PDF tambien vale
```

`presupuesto.leer()` acepta `.xlsx` y `.pdf`; con un PDF se apoya en
`tabla_pdf`, que lee la tabla por celdas y comprueba cada fila con su propia
aritmética. La lectura por celdas necesita `pdfplumber`
(`pip install pdfplumber --break-system-packages` si no está); sin él,
`tabla_pdf` cae solo al modo por espacios, que es el que produce los hallazgos
falsos, así que instálalo antes de leer un presupuesto en PDF. `tabla_pdf.inconsistentes(pres)` devuelve los rubros en los que
cantidad x precio unitario no da el precio total: son los que hay que mirar a
mano antes de creerse nada de esa tabla.

Hay un atajo para todo esto: `scripts/probar.py referencia oferta salida.xlsx`
hace la lectura, valida la correspondencia y el enlace del presupuesto, genera
el reporte e imprime lo que hay que mirar. Es la forma rapida de empezar con un
par de archivos nuevos.

`references/formatos.md` describe los tres formatos de Excel y los tres de PDF
ya resueltos, y cómo agregar el tuyo si la autodetección falla.

**Verifica la lectura antes de comparar los cientos de rubros.** Compara el
total de ítems de los dos archivos: si difieren mucho, el problema es tu lector,
no la oferta. Después imprime dos o tres rubros lado a lado contra las celdas
crudas del original. Una descripción vacía o una cantidad en `None` es la señal
típica de un mapa de columnas equivocado.

### 4. Comparar y generar el reporte

```python
import apu
apu.reporte(ref, ofe, '/ruta/Comparacion APUs X vs Oferta Y.xlsx',
            titulo_ref='referencia.xlsx  -  hojas 1 a 137',
            titulo_ofe='oferta.pdf  -  137 páginas, una por rubro',
            correspondencia='Hoja N <-> página N. Verificada en las 137.',
            pres_ref=pres_ref, pres_ofe=pres_ofe,
            pdf=True)                      # activa TEXTO CORTADO
```

Sale con cinco hojas: **RESUMEN** (semáforo por rubro), **DIFERENCIAS** (todo),
**PRESUPUESTO** (si se leyeron los dos), **SOLO ERRORES** (lo que hay que
corregir) y **NOTAS** (el criterio aplicado, para que el reporte se defienda
solo).

**Las hojas están enlazadas por fórmulas, y eso es deliberado.** El revisor y su
equipo repasan DIFERENCIAS y reclasifican a mano lo que para ellos no es un
error de fondo — `CEMENTO` frente a `CEMENTO PORTLAND TIPO 1`, por ejemplo. Para
que esa decisión valga algo, la hoja trae `TIPO DETECTADO` (gris, lo que
encontró el análisis, queda de respaldo), `TIPO REVISADO` (editable, con lista
desplegable) y una columna de `OBSERVACION`. Al cambiar el tipo revisado se
recalculan los conteos del RESUMEN, el estado del rubro, los colores —son
formato condicional, no relleno fijo— y la hoja SOLO ERRORES, sin volver a
generar nada.

El enlace entre hojas se hace con `COUNTIFS` e `INDEX/MATCH`, no con fórmulas
de matriz dinámica: `FILTER` habría sido más corto, pero openpyxl no sabe
escribirlas y Excel las borra al abrir el archivo, avisando de que "reparó" el
libro. Así funciona en cualquier versión. `apu.reporte_estatico()` queda para
cuando se quiera un archivo sin ninguna fórmula.

Otro detalle que cuesta media hora descubrir: los colores son formato
condicional, y en un formato diferencial Excel pinta el fondo con `bgColor`. El
`fgColor` que sirve en una celda normal ahí se ignora y la regla se aplica sin
que se vea nada.

El RESUMEN incorpora además, cuando hay presupuestos, la cantidad de obra de
cada uno y el precio unitario del APU de la oferta contra el de su propio
presupuesto. Son dos verificaciones que el APU por sí solo no permite.

Tres decisiones se toman solas y no hace falta preguntarlas: si comparar el
detalle (se descarta cuando la oferta lo trae en blanco en casi todos los
rubros), y las dos banderas de nomenclatura de mano de obra y unidades
equivalentes, que se activan cuando el patrón aparece en la mayoría de los casos
y no en dos sueltos. Todas quedan anotadas en NOTAS. Pásalas a mano solo para
apartarte del criterio.

### 5. Verificar y reportar

Antes de entregar, abre las celdas crudas de los dos o tres hallazgos más
gruesos y confirma que existen de verdad. Es barato y evita acusar a un oferente
por un fallo de tu lector. Cuando encuentres algo llamativo —una cantidad en 0,
un material sustituido— revísalo contra el subtotal del rubro: si el subtotal
también cambió, el hallazgo es real.

En el chat da primero el cuadro de estados, después **los hallazgos graves
agrupados por patrón** y al final, en una línea, las observaciones menores. Al
revisor le sirve más "faltan 10 ítems de mano de obra repartidos en 8 rubros"
que diez líneas sueltas — y le sirve saber que los 92 casos de cantidad mayor
no le quitan el sueño.

## Cómo se clasifica cada diferencia

| Tipo | Cuándo | Color |
|---|---|---|
| `DIFERENCIA` | Falta un ítem, cantidad menor o en cero, unidad cambiada, insumo sustituido, categoría salarial cambiada, indirectos fuera de tolerancia, cantidad de rubro menor en el presupuesto, precio que no cuadra con el APU | rojo |
| `MENOR` | Ítem adicional, cantidad mayor, mismo contenido escrito distinto, indirectos dentro de tolerancia | amarillo |
| `TEXTO CORTADO` | El PDF truncó el texto al imprimir | azul |
| `ORDEN` | Mismos ítems en distinto orden | azul |

Dos banderas de `apu.comparar()` se detectan solas, porque responden a una
decisión de forma que el oferente tomó una vez y repitió en todo el documento:

- `nomenclatura_mo_menor` — renombró toda la mano de obra manteniendo el código
  ocupacional.
- `unidades_equivalentes_menor` — abrevia las unidades de otra forma (`galón` →
  `Gln`, `l` → `Ltr`). Un cambio real de unidad (`Kg` → `u`) se sigue
  reportando como grave.

`detectar_patrones()` las activa cuando el patrón explica la mayoría de las
diferencias de su tipo, nunca por dos casos sueltos: dos ítems con el mismo
código ocupacional pueden ser casualidad, treinta son una decisión.

## Lo que cuesta caro si se pasa por alto

**Los ítems se emparejan por descripción, nunca por posición.** Si el oferente
intercala un ítem propio, comparar por posición corre todo lo que sigue y
produce una cascada de "falta" y "sobra" falsos. `apu.emparejar()` hace cuatro
pasadas: descripción idéntica, parecida, palabras contenidas una en otra
(`CEMENTO` / `CEMENTO PORTLAND TIPO 1`) y código ocupacional.

**Un total de ítems que no cuadra es tu bug, no su error.** Antes de acusar al
oferente de omitir 30 materiales, revisa el lector.

**El estado del rubro lo marca el peor hallazgo.** Un rubro con diez
diferencias de formato y una cantidad de menos está para revisar, no en verde.

**El presupuesto es donde se paga.** Un oferente puede copiar los 78 APUs sin
tocar una coma y aun así bajar la cantidad de un rubro en el presupuesto. Si
los dos archivos lo traen, compáralo siempre.

**La celda impresa esconde decimales.** Excel imprime `0,0125` como `0,01` si la
celda tiene formato de dos decimales, y en el PDF el resto ya no existe. Leído
tal cual, eso sale como una cantidad menor que la referencia —grave— cuando el
oferente no bajó nada. Se resuelve despejando: la fila es lineal en la cantidad
(`cantidad x tarifa x rendimiento = costo`), así que el valor completo es el
costo dividido para los demás factores, y solo se acepta si redondeado a los
decimales impresos da exactamente la cifra impresa. `leer_pdf()` lo hace solo;
`lector_pdf.cantidades_recuperadas()` devuelve la lista para anotarla en NOTAS.
Si el oferente escribió `0,01` de verdad, el costo cuadra con `0,01` y no se
toca nada.

**El PDF no miente, pero se lee mal.** Estas plantillas imprimen la descripción
larga en tres renglones —la cabeza arriba, las cifras en el del medio, la cola
abajo—, y entre dos filas conviven la cola de una y la cabeza de la siguiente.
Leído por columnas de espacios, el rubro se queda con el trozo del medio y el
vecino hereda lo que sobra; la firma electrónica, que se imprime *encima* de la
tabla, mete el apellido del firmante dentro de un rubro. Los lectores ya traen
esto resuelto —celdas con pdfplumber, cruce con la lectura por espacios, filtro
de firmas—, pero el criterio que hay detrás es el que importa cuando aparezca
una plantilla nueva: **una diferencia que sólo afecta al texto, en un documento
que viene de un PDF, es del lector hasta que se demuestre lo contrario.** Ábrela
contra el original antes de reportarla. La forma barata de demostrarlo es la
aritmética de la fila y el subtotal del rubro: lo que no cuadra está mal leído.

## Si el formato no encaja

Copia de `scripts/lectores.py` la función más parecida y ajusta el mapa de
columnas; son unas quince líneas. Para PDF, agrega un preset a
`lector_pdf.ETIQUETAS` con su `pre`, `coma_decimal` y `modo` —y `detalle_fiable:
False` si la plantilla imprime la celda DETALLE encima de la de UNIDAD, para que
la comparación no invente diferencias de especificación.
`references/formatos.md` tiene el detalle de cada formato resuelto y sus
trampas.
