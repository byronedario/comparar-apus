# Formatos de APU ya resueltos

Léelo cuando el autodetector falle o cuando el archivo no encaje en ninguno de
los tres formatos de Excel. Todos los APUs tienen la misma anatomía —cabecera,
cuatro secciones, subtotales— y lo único que cambia es dónde está cada cosa.

## Contenido

- [Anatomía común](#anatomía-común)
- [Excel A — hojas "RubroN" con cabecera en una sola celda](#excel-a)
- [Excel B — hoja numerada clásica](#excel-b)
- [Excel C — plantilla con desagregación / VAE](#excel-c)
- [PDF — impreso desde Excel](#pdf)
- [Cómo agregar un formato nuevo](#cómo-agregar-un-formato-nuevo)

## Anatomía común

Un rubro por hoja o por página, con esta estructura:

```
cabecera:  código, nombre del rubro, unidad, especificación/detalle
EQUIPOS         descripción | cantidad | tarifa | costo hora | rendimiento | costo
MANO DE OBRA    descripción | cantidad | jornal | costo hora | rendimiento | costo
MATERIALES      descripción | unidad | cantidad | precio unitario | costo
TRANSPORTE      descripción | unidad | distancia | cantidad | tarifa | costo
                subtotales, costo directo, indirectos, costo total
```

Se comparan descripción, unidad y cantidad. Todo lo demás es del oferente.

Las secciones se localizan siempre igual: el nombre de la sección en la columna
A (a veces con dos puntos, `EQUIPOS:`), y el bloque de ítems cierra en una fila
que dice `SUBTOTAL` o `PARCIAL`. Lo que cambia entre formatos es cuántas filas
de encabezado hay entre el título de la sección y el primer ítem.

## Excel A

`lectores.formato_a` — hojas llamadas `Rubro1`, `Rubro2`, …

La cabecera va **dentro de una sola celda** (A3) con saltos de línea:

```
A3 = "RUBRO No : 5044 \n DESCRIPCIÓN : CERRAM.PROVISIONAL... \n ESPECIFICACIÓN : Tabla dura..."
```

La unidad está en otra columna de la misma fila 3, como `"UNIDAD:     m"`.

- Una sola fila de encabezado por sección; los ítems empiezan en `r0 + 2`.
- El bloque cierra en `PARCIAL M/N/O/P`, **en una columna que varía** (E o F
  según el archivo), por eso se busca la palabra en toda la fila.
- EQUIPOS y MANO DE OBRA: descripción en B, cantidad en C.
- MATERIALES y TRANSPORTE: descripción en B, unidad en C, cantidad en D.

Visto en: `5.2_PRESUPUESTO-APUS-VAE Serd.xlsm`, `APUS diego ocaña SECRETARIA.xlsx`.

⚠️ Estos libros suelen traer **también** un juego de hojas numeradas `1`…`N` con
los mismos rubros. Las hojas `RubroN` pueden ser una versión antigua y estar en
otro orden. Verifica cuál corresponde con la oferta antes de decidir.

## Excel B

`lectores.formato_b` — hojas llamadas `1`, `2`, …

Etiqueta y valor en celdas contiguas:

```
fila 7:  A="RUBRO:"   B=nombre
fila 6:  F="CÓDIGO:"  G=valor
fila 8:  F="UNIDAD:"  G=valor
fila 9:  B=especificación (texto suelto, sin etiqueta)
```

- Dos filas de encabezado por sección; los ítems empiezan en `r0 + 3`.
- Cierra en `SUBTOTAL M/N/O/P`.
- EQUIPOS y MANO DE OBRA: descripción en B, cantidad en C.
- MATERIALES y TRANSPORTE: descripción en B, unidad en D, cantidad en E
  (la columna C queda vacía).

Visto en: `5.2 APUS_CAUPICHU principal.xlsx`, hojas numeradas del archivo Serd.

## Excel C

`lectores.formato_c` — plantilla con desagregación tecnológica / VAE.

Igual que la B pero **la descripción del ítem arranca en la columna A**, y a la
derecha hay columnas que no son del APU: `PESO RELATIVO ELEMENTO`,
`CPC Elemento`, `NP/ND/EP`, `VAE (%)`. Esas se ignoran.

```
A="NOMBRE DEL RUBRO:"  B=nombre        (a veces solo "RUBRO:")
A="CÓDIGO DEL RUBRO:"  B=código
A="DETALLE:"           B=especificación
E="UNIDAD: "           F=valor
```

- EQUIPOS y MANO DE OBRA: descripción en A, cantidad en B.
- MATERIALES: descripción en A, unidad en C, cantidad en D.
- TRANSPORTE: descripción en A, unidad en B, **distancia en C**, cantidad en D.

Visto en: `5-12-30 Apus con VAE Oferente.xlsx`,
`3-14-APUS CON DESEGREGACION Oferente.xlsx`,
`10-16-22 LICO-MDMQ-2026-4909 DIEGO FINAL.xlsx`.

## PDF

`lector_pdf.leer_pdf(ruta, etiquetas, coma_decimal, modo)`. Un rubro por página,
impreso desde Excel, con texto real: `pdftotext -layout` conserva las columnas
como espacios. Antes de programar nada, mira el texto crudo:

```bash
pdftotext -layout "oferta.pdf" - | head -80
```

Cuatro cosas dan más guerra de lo que parece:

**1. La descripción larga se parte en varios renglones.** El renglón con las
cifras puede ser el primero, el del medio o el último del grupo, y a veces un
renglón suelto arrastra una sola cifra (el costo de la celda multilínea). Por eso
un ítem se reconoce por traer **dos o más cifras**, y los renglones con una o
ninguna se acumulan como texto de la descripción.

**2. La cabecera está centrada verticalmente.** Un valor de tres líneas deja la
etiqueta (`RUBRO:`, `DETALLE`) en la del medio, no arriba. Por eso cada fragmento
se asigna a la etiqueta más cercana. Y una misma fila puede llevar el final de
una celda a la izquierda y la etiqueta de otra columna a la derecha, así que hay
que cortar en el primer hueco ancho y repartir los dos lados.

**3. `modo` decide de dónde se leen las columnas del ítem:**

- `'derecha'` — la fila termina en las cifras del APU. Se barre desde la derecha.
- `'izquierda'` — la fila sigue con las columnas de VAE. Barrer desde la derecha
  se atasca en el token `EP`/`NP`, así que la cantidad se toma como la primera
  cifra después de la descripción.

**4. `coma_decimal`** — `0,8300` con punto de miles (`1.342,86`) frente a
`0.83`. Equivocarlo multiplica o divide cantidades por mil sin avisar.

Presets en `lector_pdf.ETIQUETAS`:

| preset | numeración | etiquetas | modo | coma decimal |
|---|---|---|---|---|
| `rubro_n` | `RUBRO N° 1` | `RUBRO:`, `UNIDAD`, `DETALLE` | `derecha` | no |
| `hoja_n_de` | `HOJA: 1 DE 137` | `RUBRO:`, `UNIDAD:`, `ESPECIFICACIÓN:` | `izquierda` | sí |
| `uem_codigo` | `RUBRO No : 5044` | `DESCRIPCIÓN:`, `UNIDAD:`, `ESPECIFICACIÓN:` | `izquierda` | no |

Cada preset lleva además `ignorar`: las etiquetas vecinas cuyo texto no debe
colarse en el nombre (`RENDIMIENTO:`, `OFERENTE:`, `CÓDIGO:`). Se comparan como
palabra completa, porque `PROYECTOR MODULAR` empieza por `PROYECTO`.

**`leer_pdf(ruta)` sin más argumentos autodetecta** el preset, el decimal y el
modo: prueba cada combinación sobre las primeras páginas y se queda con la que
saca más ítems con descripción y cantidad, penalizando las cantidades absurdas
que delatan un decimal mal elegido. Pásalos a mano solo si eso falla.

### PDF de la UEM (`uem_codigo`)

El más enredado de los tres, y por eso tiene su propio preprocesador
(`_pre_uem`). Cuatro cosas hay que deshacer antes de parsear:

1. **Columna CODIGO a la izquierda de cada ítem** (`03`, `0701`, `M1`, `EV114`).
   Si se deja, en modo `izquierda` el código se toma como la cantidad.
2. **El código del rubro es alfanumérico** (`5044`, pero también `P524`, `V302`).
   `parse_pagina` espera un entero, así que se sustituye por el número de página
   y el código real se guarda en `codigo`.
3. **La cabecera de las columnas de VAE sale partida en fragmentos sueltos**
   (`PESO` / `RELATI` / `VO`, `ELEME` / `NTO` / `(%)`), repartidos antes y
   después del nombre de la sección. Si se cuelan, se pegan a la descripción del
   primer ítem del bloque.
4. **El nombre de la sección arrastra `- 0 0`** de esas mismas columnas, y con
   cifras dentro deja de reconocerse como cabecera de sección.

Además la sección cierra en `PARCIAL M/N/O/P`, no en `SUBTOTAL`, y `UNIDAD:`
comparte renglón con `RUBRO No :`, que es donde `parse_pagina` empieza a buscar
las etiquetas — por eso el preprocesador parte ese renglón en dos.

### El pie del APU

`lector_pdf.pie_pdf()` y `lectores.pie()` sacan costo directo, porcentaje de
indirectos y precio total, que es lo que necesita `apu.margen()`. Dos trampas:

- El importe es la **primera** cifra de la fila, no la última: varias plantillas
  cierran con una columna de porcentaje (el `100` del costo directo) y tomar la
  última devuelve ese 100 en vez del dinero.
- El porcentaje aparece de cuatro formas: dentro del rótulo
  (`COSTO INDIRECTO ( 20.0000 % )`), pegado a él (`COSTOS INDIRECTOS 7.8%`),
  **en el renglón siguiente** cuando la celda está centrada verticalmente (por
  eso se mira hasta dos renglones más abajo), y **sin el signo `%`**
  (`COSTO INDIRECTO    17,00    0,2732`). En ese último caso cuál de las dos
  cifras es el porcentaje no se adivina: se comprueba cuál, aplicada al costo
  directo, da la otra.

  Que el porcentaje se lea importa más de lo que parece. Cuando no está,
  `apu.margen()` lo deduce del precio, que viene redondeado a dos decimales, y
  un rubro barato sale desviado: con costo directo `0,0574` y precio `0,07` el
  margen calculado es 21,95 % cuando el declarado es 17 %. Una sola decisión del
  oferente se convierte así en un hallazgo falso.

## PDF: lo que rompe la lectura y cómo se resuelve

Los tres formatos de PDF de arriba se leen con `pdftotext -layout`, que conserva
las columnas como espacios. Eso alcanza para las páginas de APU, pero **no para
la tabla de rubros** (presupuesto, tabla de cantidades y precios, oferta
económica), que va por `scripts/tabla_pdf.py`. Estas son las trampas, todas
vistas en archivos reales del mismo proceso:

**La descripción se imprime en tres renglones.** Cabeza arriba, cifras en el del
medio, cola abajo. Entre dos filas de cifras conviven la cola de una y la cabeza
de la siguiente, y por columnas de texto no hay forma de saber cuál es cuál:

```
   MALLA NYLON ALTO IMPACTO COLOR VERDE          <- cabeza del 75
75 ROMBO 5cm, RESISTENTE A LA INTEMPERIE CON m2 3800 2,52 ...
   PROTECCIÓN UV.                                <- cola del 75
   DESALOJO EQUIPO PESADO TIERRA/ESCOMBROS       <- cabeza del 76
76                                           m3  300 6,68 ...
```

Sin tratar, el 75 pierde la cabeza y el 76 hereda la cola. La salida es leer por
**celdas** con pdfplumber, que devuelve los tres renglones ya unidos: estas
tablas están rayadas y las líneas son datos reales del PDF, no una suposición.

**La celda se desborda.** Cuando el texto no cabe, pdfplumber lo corta a media
palabra y le entrega el sobrante a la columna vecina: la unidad deja de parecer
una unidad (`A INTEmM2PERIE` es `A LA INTEMPERIE` y `m2` impresos uno encima del
otro). Por eso `tabla_pdf.leer()` lee también por espacios y cruza las dos
lecturas fila por fila.

**El árbitro es la aritmética de la fila.** Cantidad x precio unitario = precio
total. Una fila mal leída casi nunca cuadra y una bien leída casi siempre, así
que no hay que elegir lector a ciegas: se elige por fila. Lo que sigue sin
cuadrar lo devuelve `tabla_pdf.inconsistentes()` para mirarlo a mano.

**La firma electrónica se imprime ENCIMA de la tabla.** No al final: sus
renglones aterrizan dentro de las filas. En un caso real el rubro 77 salió como
`SOBREACARREO DE ESCOMBROS, TIERRA DE NARVAEZ EXCAVACIÓN Y MAT. PETREOS.`, con
el apellido del firmante metido en medio; en otro, el bloque de firmas de la
última página aportó tres ítems de TRANSPORTE al rubro 78. `lector_pdf.FIRMA` y
`tabla_pdf.FIRMA` los quitan, junto con el nombre suelto que dejan debajo.

**El mismo documento repite los rubros.** Después del presupuesto suelen venir
la tabla del VAE y la desagregación tecnológica, con los mismos 78 rubros y
otras columnas. Manda la **primera** aparición; si no, la segunda tabla pisa la
buena.

**Un PDF de APUs no es un presupuesto.** Cada sección de un APU también tiene
descripción, cantidad y un total, así que el mismo lector devuelve una tabla de
rubros inventada con los insumos. `tabla_pdf` exige que la cabecera nombre el
NÚMERO de ítem, que la numeración vaya del 1 al N casi sin huecos y que la
mayoría de las filas cuadren; si no, devuelve `{}`.

**El separador decimal no se adivina, se cuenta.** Un grupo de miles tiene
exactamente tres dígitos, así que una coma o un punto seguidos de **cuatro**
dígitos son el decimal sin discusión, y las plantillas de APU imprimen las
cantidades con cuatro decimales. Probar los dos y puntuar el resultado no avisa
cuando se equivoca: con el separador cambiado, `0,0940` se lee `940` y el rubro
entero sale mal sin que nada falle.

**DETALLE impreso encima de UNIDAD.** En el formato `ushay_vae` las dos celdas
se superponen y el texto sale entrelazado. El preset lo declara con
`detalle_fiable: False` (la especificación se vacía y la comparación la descarta
sola) y la unidad del rubro se rellena desde la tabla de cantidades del mismo
documento.

## Cómo agregar un formato nuevo

1. Vuelca una hoja o página cruda y ubica: fila de la cabecera, filas de
   encabezado por sección, y qué columna es descripción, unidad y cantidad.

   ```python
   import openpyxl, apu
   ws = openpyxl.load_workbook('archivo.xlsx', read_only=True, data_only=True)['1']
   g = apu.grid(ws, 70, 10)
   for r in range(1, 71):
       print(r, [('' if c is None else str(c)[:28]) for c in g[r][1:11]])
   ```

2. Copia la función de `lectores.py` más parecida, renómbrala y ajusta el mapa
   de columnas. Regístrala en el diccionario `FORMATOS`.

3. Compruébala en un rubro contra las celdas crudas, y compara el total de ítems
   contra el otro archivo antes de correr todos los rubros.
