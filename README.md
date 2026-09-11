# comparar-apus

Skill para Claude que compara los **análisis de precios unitarios (APUs)** y el
presupuesto de la oferta de un oferente contra los archivos de referencia del
contratante, rubro por rubro, y entrega un Excel con las diferencias graduadas
por gravedad.

Nació revisando ofertas reales de licitaciones de obra pública en Ecuador
(LICO del Municipio de Quito), y el criterio que aplica es el de un revisor de
comisión técnica. Ese criterio está documentado abajo y **se puede cambiar
pidiéndoselo a Claude en lenguaje normal** — no hace falta tocar código.

---

## Qué hace

Le das dos archivos —el de referencia del contratante y el del oferente, en
Excel, en PDF o mezclados— y produce un Excel con cinco hojas:

| Hoja | Qué trae |
|---|---|
| **RESUMEN** | Semáforo por rubro: correcto, observación menor, orden/texto cortado, a revisar |
| **DIFERENCIAS** | Todos los hallazgos, con el valor de la referencia y el de la oferta lado a lado |
| **PRESUPUESTO** | Cantidades de rubro y cuadre de precios, si los dos archivos traen presupuesto |
| **SOLO ERRORES** | Lo que hay que corregir, sin ruido |
| **NOTAS** | El criterio aplicado, para que el reporte se defienda solo |

Lee sola tres plantillas de Excel y tres de PDF, y detecta por su cuenta el
formato, el separador decimal y la disposición de columnas. Los PDFs son Excel
impresos con texto real: no necesita OCR.

---

## Instalación

### En la app de Claude (Free, Pro, Max)

1. Descarga [`comparar-apus.skill`](https://github.com/byronedario/comparar-apus/raw/main/comparar-apus.skill)
   (botón derecho → guardar enlace como).
2. En Claude: **Personalizar → Habilidades → Agregar**, y sube el archivo.
3. Listo. Se activa sola cuando menciones APUs, rubros, oferentes o una oferta
   que quieras revisar; o puedes forzarla escribiendo `/comparar-apus`.

Si tu cliente de correo o tu navegador no deja pasar la extensión `.skill`,
renómbrala a `.zip`: por dentro es lo mismo.

### En Claude Code y otras herramientas

```bash
git clone https://github.com/byronedario/comparar-apus.git
cp -r comparar-apus/comparar-apus ~/.claude/skills/
```

El formato (`SKILL.md` + `scripts/` + `references/`) es el estándar abierto
Agent Skills, así que la carpeta también funciona en las herramientas de línea
de comandos que lo adoptaron.

### Uso

Adjunta los dos archivos y escribe algo como:

> Compara los APUs de esta oferta contra el archivo base "APUS proyecto X.xlsx".

No necesita más. Decide sola si vale la pena comparar el campo
detalle/especificación, y avisa en las notas cuando lo descarta.

---

## El criterio de evaluación

Esto es lo que hay que entender antes de usarla, y lo que conviene revisar si
trabajas con otro pliego o en otro país.

### La regla de fondo: la asimetría

**La oferta no tiene que ser idéntica a la referencia: tiene que cubrir el
rubro.** El oferente pone sus propios precios —esa es su oferta— pero no puede
cambiar la composición del rubro. De ahí sale una asimetría que decide todo lo
demás:

| Situación | Gravedad | Por qué |
|---|---|---|
| Falta un ítem | **Grave** | El rubro queda sin un componente que el contratante consideró necesario |
| Cantidad **menor** a la de la referencia | **Grave** | Se comprometió a ejecutar con menos de lo estimado |
| Cantidad **en cero** | **Grave**, y se señala aparte | El insumo desaparece del costo sin dejar rastro en el APU |
| Ítem **adicional** | Menor | Destina más recursos al rubro; no perjudica al contrato |
| Cantidad **mayor** | Menor | Igual que el anterior |
| Ítems en **otro orden** | Informativo | No es error: las cantidades se comparan contra el ítem correcto |
| Mismo contenido escrito distinto | Menor | Mayúsculas, tildes, `#` por `No.` |

La baja por redondeo no es excepción: `1,005 → 1,00` es menos cantidad y se
observa igual, porque el efecto sobre la obra es el mismo venga de donde venga.

### Indirectos

El oferente puede bajar sus indirectos **hasta 5 puntos porcentuales** respecto
a los del presupuesto referencial, pero no subirlos. Dentro de esa tolerancia
es observación menor; fuera de ella —o hacia arriba— es grave.

Como es **una** decisión que se repite en todos los rubros, se reporta agrupada
en una línea que dice cuánto bajó y en qué rubros, y no arrastra a cada rubro
al estado "a revisar". Reportarlo 137 veces taparía todo lo demás.

### Mano de obra

Renombrar la mano de obra manteniendo el código de estructura ocupacional
(`ESTRUC. OCUPAC. E2 PEON` → `Estr. Oc. E2 PEON`) es el mismo obrero escrito
distinto, y baja a observación menor. **Cambiar la categoría** (C3 → C1) no
está permitido y sigue siendo grave: es otro salario y otro perfil.

### Lo que NO se compara, y por qué

Tarifas, jornales, precios unitarios de insumos, rendimientos y costos son del
oferente. Compararlos llena el reporte de ruido y esconde lo que importa.

**El monto de la oferta tampoco se observa**: ya puntúa por sí mismo en la
calificación económica, así que un titular tipo "+10% de sobrecosto" es un
hallazgo que no le sirve a nadie.

La única excepción del lado de los precios: que el precio unitario del
**presupuesto** de la oferta no cuadre con el que sale de su **propio APU**.
Ahí uno de los dos documentos no es el que el oferente va a ejecutar, y eso sí
es observable.

### El presupuesto

Un oferente puede copiar los APUs sin tocar una coma y aun así bajar la
cantidad de un rubro en el presupuesto, que es donde de verdad se paga. Si los
dos archivos traen hoja de presupuesto, se comparan las cantidades de cada
rubro además de los APUs.

---

## Cómo cambiar el criterio sin tocar código

Todo lo de arriba está escrito en `SKILL.md` y en `scripts/apu.py` con su
razonamiento explicado, así que puedes pedirle a Claude que lo ajuste en
lenguaje normal. Abre un chat, adjunta la carpeta de la skill (o pídele que la
lea de tu perfil) y dile qué quieres cambiar. Ejemplos que funcionan:

**Cambiar la tolerancia de indirectos**

> En mi pliego la tolerancia de indirectos es de 3 puntos, no de 5. Ajusta la
> skill comparar-apus y vuelve a empaquetarla.

**Que los indirectos no se observen**

> Quita la comparación de indirectos de la skill comparar-apus: en mi entidad
> el porcentaje es libre.

**Cambiar la asimetría**

> En mi caso cualquier cantidad distinta a la de la referencia es observable,
> tanto por encima como por debajo. Cambia el criterio de la skill.

**Tratar el redondeo aparte otra vez**

> Que las cantidades que solo bajaron por redondeo a dos decimales salgan como
> observación informativa y no como error grave.

**Comparar también precios o rendimientos**

> Agrega a la skill una hoja que compare los rendimientos de la oferta contra
> los de la referencia, marcada como informativa.

**Agregar el formato de tu institución**

> Este PDF de oferta no lo lee bien. Lee references/formatos.md, agrega un
> preset nuevo para este formato y verifica la lectura contra el original.

Si prefieres tocarlo tú, los dos puntos más probables son:

- `scripts/apu.py` → `TOLERANCIA_INDIRECTOS = 5.0`
- `scripts/apu.py` → `tipo_cantidad()`, que es donde vive la asimetría

Después de cualquier cambio hay que volver a subir la skill a Claude para que
tome la versión nueva.

---

## Formatos que ya lee

**Excel** — tres plantillas, autodetectadas: hojas `RubroN` con la cabecera
embebida en una celda; hoja numerada clásica; y la plantilla con desagregación
tecnológica / VAE.

**PDF** — tres formatos, autodetectados junto con el separador decimal y la
disposición de columnas: numeración `RUBRO N° 1`, numeración `HOJA: 1 DE 137`,
y el formato de la UEM con código de rubro alfanumérico y columna de código por
ítem.

`references/formatos.md` describe cada uno, sus trampas y cómo agregar el tuyo.
Si un archivo no encaja, la skill sabe cómo extenderse: son unas quince líneas.

---

## Estructura

```
comparar-apus/
├── SKILL.md                 el criterio y el flujo de trabajo
├── scripts/
│   ├── apu.py               motor: normalización, emparejamiento, clasificación, reporte
│   ├── lectores.py          lectura de Excel y de la hoja de presupuesto
│   └── lector_pdf.py        lectura de PDF con autodetección de formato
├── references/
│   └── formatos.md          anatomía de cada plantilla resuelta
└── evals/
    └── evals.json           casos de prueba con los resultados esperados
```

Los scripts son Python normal (`openpyxl` y `pdftotext`), sin nada atado a
Claude: se pueden correr a mano o desde otra herramienta.

---

## Advertencia

Esta skill **ayuda a revisar**, no decide. Verifica siempre los hallazgos
gruesos contra el archivo original antes de trasladarlos a un informe: una
observación mal fundamentada en un proceso de contratación pública tiene
consecuencias para el oferente y para quien la firma.

El criterio que trae por defecto es el de un pliego concreto. Antes de usarla
en otra entidad, revisa la sección del criterio y ajústala.

## Licencia

MIT.
