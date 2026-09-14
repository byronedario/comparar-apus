"""Lectores de hojas de Excel con APUs. Tres disposiciones ya vistas + autodeteccion.

Todas devuelven la estructura que espera apu.py:
    {'codigo', 'nombre', 'unidad', 'detalle', 'secs': {SECCION: [items]}}

Si tu archivo no encaja en ninguna, copia la funcion mas parecida y ajusta el
mapa de columnas. Lo unico que cambia entre formatos es: donde esta la cabecera,
cuantas filas de encabezado tiene cada seccion, y que columna es descripcion,
unidad y cantidad.
"""
import re
from apu import SECS, clean, num, grid, fila_vacia


def _cabecera_secciones(A, max_fil):
    """Fila donde arranca cada seccion (tolera 'EQUIPOS' y 'EQUIPOS:')."""
    ini = {}
    for r in range(1, max_fil + 1):
        t = clean(A(r, 1)).upper().rstrip(':').strip()
        if t in SECS and t not in ini:
            ini[t] = r
    return ini


def pie(A, max_fil=80, max_col=9):
    """Costo directo, porcentaje de indirectos y precio total del rubro.

    Cada plantilla lo escribe distinto: el porcentaje puede ir dentro del propio
    rotulo ("COSTO INDIRECTO ( 20.0000 % )"), en una celda aparte como 20, o como
    la fraccion 0.2. Se normaliza siempre a puntos porcentuales.

    El importe es la PRIMERA cifra de la fila: varias plantillas cierran con una
    columna de porcentaje (el 100 del costo directo, el 0 del precio total) y
    tomar la ultima devuelve ese 100 en vez del dinero.
    """
    out = {'directo': None, 'indirecto_pct': None, 'precio': None}
    for r in range(1, max_fil + 1):
        fila = [clean(A(r, c)) for c in range(1, max_col + 1)]
        txt = ' '.join(fila).upper()
        cifras = [num(x) for x in fila if num(x) is not None]
        if not txt.strip():
            continue
        if re.search(r'COSTO\s+DIRECTO|DIRECTO\s*\(M', txt) and cifras:
            out['directo'] = cifras[0]
        elif re.search(r'INDIRECTO', txt) and 'OTROS' not in txt:
            m = re.search(r'\(\s*([\d.,]+)\s*%', txt)
            if m:
                out['indirecto_pct'] = num(m.group(1))
            elif len(cifras) >= 2:
                p = cifras[0]
                out['indirecto_pct'] = p * 100 if p is not None and p < 1 else p
            if cifras:
                out['_ind_valor'] = cifras[-1]
        elif re.search(r'PRECIO\s+UNITARIO\s+TOTAL|COSTO\s+TOTAL|VALOR\s+OFERTADO', txt) \
                and cifras:
            out['precio'] = cifras[0]
    out.pop('_ind_valor', None)
    return out


def _fin_bloque(A, r, marcas, max_col=9):
    """True si la fila r cierra el bloque de items."""
    fila = [clean(A(r, c)).upper() for c in range(1, max_col + 1)]
    if fila[0].rstrip(':').strip() in SECS:
        return True
    return any(x.startswith(m) for x in fila for m in marcas)


# ---------------------------------------------------------------------------
# Formato A: hojas "RubroN" con la cabecera embebida en la celda A3
#   A3 = "RUBRO No : 5044 \n DESCRIPCION : ... \n ESPECIFICACION : ..."
#   la unidad esta en otra columna de la fila 3, como "UNIDAD: m"
#   una sola fila de encabezado por seccion; el bloque cierra en "PARCIAL x"
# ---------------------------------------------------------------------------
def formato_a(ws):
    g = grid(ws, 80, 9)
    A = lambda r, c: g[r][c]

    head = clean(A(3, 1))
    m = re.search(r'DESCRIPCI[ÓO]N\s*:\s*(.*?)(?:\s*ESPECIFICACI[ÓO]N\s*:|$)', head, re.I)
    nombre = clean(m.group(1)) if m else ''
    m = re.search(r'ESPECIFICACI[ÓO]N\s*:\s*(.*)$', head, re.I)
    detalle = clean(m.group(1)) if m else ''
    m = re.search(r'RUBRO\s*No\.?\s*:?\s*([A-Za-z0-9\-\.]+)', head, re.I)
    codigo = m.group(1) if m else ''
    unidad = ''
    for c in range(2, 10):
        t = clean(A(3, c))
        if re.search(r'UNIDAD', t, re.I):
            unidad = clean(re.sub(r'.*UNIDAD\s*:?', '', t, flags=re.I))

    secs = {}
    for s, r0 in _cabecera_secciones(A, 80).items():
        items, r = [], r0 + 2
        while r <= 80 and not _fin_bloque(A, r, ('PARCIAL',)):
            if s in ('EQUIPOS', 'MANO DE OBRA'):
                d, u, q = A(r, 2), None, A(r, 3)
            else:
                d, u, q = A(r, 2), A(r, 3), A(r, 4)
            if not fila_vacia(d, q):
                items.append({'desc': clean(d), 'uni': clean(u), 'cant': num(q)})
            r += 1
        secs[s] = items
    d = {'codigo': codigo, 'nombre': nombre, 'unidad': unidad,
         'detalle': detalle, 'secs': secs}
    d.update(pie(A))
    return d


# ---------------------------------------------------------------------------
# Formato B: hoja numerada clasica, etiqueta y valor en celdas contiguas
#   "RUBRO:" | nombre        "CODIGO:" | valor        "UNIDAD:" | valor
#   dos filas de encabezado por seccion; cierra en "SUBTOTAL x"
#   MATERIALES: descripcion en B, unidad en D, cantidad en E
# ---------------------------------------------------------------------------
def formato_b(ws):
    g = grid(ws, 80, 9)
    A = lambda r, c: g[r][c]

    nombre = codigo = unidad = detalle = ''
    for r in range(1, 13):
        for c in range(1, 9):
            t = clean(A(r, c))
            if re.fullmatch(r'RUBRO\s*:?', t, re.I) and not nombre:
                nombre = clean(A(r, c + 1)) or clean(A(r, c + 2))
            elif re.fullmatch(r'C[ÓO]DIGO\s*:?', t, re.I) and not codigo:
                codigo = clean(A(r, c + 1))
            elif re.fullmatch(r'UNIDAD\s*:?', t, re.I) and not unidad:
                unidad = clean(A(r, c + 1))

    ini = _cabecera_secciones(A, 80)
    r0 = min(ini.values()) if ini else 12
    for r in range(1, r0):                       # especificacion: texto suelto en B
        t = clean(A(r, 2))
        if t and t != nombre and t not in ('0', '-', '.') \
                and not re.match(r'"?CONSTRUC|PROYECTO|RUBRO|UNIDAD|C[ÓO]DIGO', t, re.I):
            detalle = t
            break

    secs = {}
    for s, r0 in ini.items():
        items, r = [], r0 + 3
        while r <= 80 and not _fin_bloque(A, r, ('SUBTOTAL', 'TOTAL COSTO', 'ESTOS PRECIOS')):
            if s in ('EQUIPOS', 'MANO DE OBRA'):
                d, u, q = A(r, 2), None, A(r, 3)
            else:
                d, u, q = A(r, 2), A(r, 4), A(r, 5)
            if not fila_vacia(d, q):
                items.append({'desc': clean(d), 'uni': clean(u), 'cant': num(q)})
            r += 1
        secs[s] = items
    d = {'codigo': codigo, 'nombre': nombre, 'unidad': unidad,
         'detalle': detalle, 'secs': secs}
    d.update(pie(A))
    return d


# ---------------------------------------------------------------------------
# Formato C: plantilla con desegregacion / columnas de VAE
#   "NOMBRE DEL RUBRO:" | nombre    "DETALLE:" | ...    "UNIDAD:" | valor
#   la descripcion del item arranca en la columna A
#   dos filas de encabezado; a la derecha van Peso Relativo, CPC, NP/ND/EP, VAE
# ---------------------------------------------------------------------------
def formato_c(ws):
    g = grid(ws, 95, 9)
    A = lambda r, c: g[r][c]

    nombre = codigo = unidad = detalle = ''
    for r in range(1, 13):
        for c in range(1, 8):
            t = clean(A(r, c))
            if re.fullmatch(r'(NOMBRE DEL )?RUBRO\s*:?', t, re.I) and not nombre:
                nombre = clean(A(r, c + 1)) or clean(A(r, c + 2))
            elif re.fullmatch(r'C[ÓO]DIGO( DEL RUBRO)?\s*:?', t, re.I) and not codigo:
                codigo = clean(A(r, c + 1))
            elif re.fullmatch(r'UNIDAD\s*:?', t, re.I) and not unidad:
                unidad = clean(A(r, c + 1)) or clean(A(r, c + 2))
            elif re.fullmatch(r'DETALLE\s*:?', t, re.I) and not detalle:
                detalle = clean(A(r, c + 1)) or clean(A(r, c + 2))

    secs = {}
    for s, r0 in _cabecera_secciones(A, 95).items():
        items, r = [], r0 + 3
        while r <= 95 and not _fin_bloque(A, r, ('SUBTOTAL', 'TOTAL COSTO', 'ESTOS PRECIOS')):
            if s in ('EQUIPOS', 'MANO DE OBRA'):
                d, u, q = A(r, 1), None, A(r, 2)
            elif s == 'MATERIALES':
                d, u, q = A(r, 1), A(r, 3), A(r, 4)
            else:                                # TRANSPORTE: desc | uni | dist | cant
                d, u, q = A(r, 1), A(r, 2), A(r, 4)
            if not fila_vacia(d, q):
                items.append({'desc': clean(d), 'uni': clean(u), 'cant': num(q)})
            r += 1
        secs[s] = items
    d = {'codigo': codigo, 'nombre': nombre, 'unidad': unidad,
         'detalle': detalle, 'secs': secs}
    d.update(pie(A))
    return d


# ---------------------------------------------------------------------------
# Formato generico: no fija ninguna columna, las busca por su NOMBRE
#
# Los tres formatos de arriba llevan el mapa de columnas escrito a mano, asi que
# una plantilla que mueva la tabla una columna a la derecha -o que ponga UNIDAD
# entre la descripcion y la cantidad solo en MATERIALES- devuelve cero items sin
# que nada falle. Este lector mira la fila de encabezados de cada seccion y se
# queda con la columna que se llama DESCRIPCION, UNIDAD y CANTIDAD. Pruebalo
# antes de escribir un formato nuevo: casi siempre sobra con esto.
# ---------------------------------------------------------------------------
_COLS = {'desc': r'DESCRIPCI[ÓO]N|CUADRILLA|RUBRO|DETALLE',
         'uni': r'UNIDAD|UND\b|^U\.?$',
         'cant': r'CANTIDAD|CANT\b'}
_FIN = ('SUBTOTAL', 'PARCIAL', 'TOTAL COSTO', 'COSTO TOTAL', 'ESTOS PRECIOS',
        'ESTE PRECIO', 'INDIRECTOS', 'UTILIDAD', 'VALOR OFERTADO')


# El nombre de la seccion y el de su columna pueden compartir celda, separados
# por un salto de linea ("EQUIPO\nDESCRIPCION", "CANTIDAD\nA"): se mira solo el
# primer renglon. Y hay plantillas que escriben la seccion en singular.
SINONIMOS_SEC = {'EQUIPO': 'EQUIPOS', 'EQUIPOS Y HERRAMIENTAS': 'EQUIPOS',
                 'MANO DE OBRA': 'MANO DE OBRA', 'MATERIAL': 'MATERIALES',
                 'MATERIALES': 'MATERIALES', 'TRANSPORTE': 'TRANSPORTE',
                 'EQUIPOS': 'EQUIPOS'}


def _primera_linea(v):
    return clean(str(v).split('\n')[0]) if v is not None else ''


def _seccion_en(fila):
    """(seccion, columna) si esta fila anuncia una seccion; si no, (None, None)."""
    for c, v in enumerate(fila[1:], 1):
        t = _primera_linea(v).upper().rstrip(':').strip()
        if t in SINONIMOS_SEC:
            return SINONIMOS_SEC[t], c
    return None, None


def _mapa_columnas(fila):
    """{'desc': col, 'uni': col, 'cant': col} leyendo los nombres de la fila.

    Se miran TODOS los renglones de cada celda: la celda que anuncia la seccion
    suele traer debajo, en la misma celda, el nombre de su columna
    ("EQUIPO\nDESCRIPCION").
    """
    m = {}
    for c, v in enumerate(fila[1:], 1):
        if v is None:
            continue
        for linea in str(v).split('\n'):
            t = clean(linea).upper()
            if not t:
                continue
            for campo, pat in _COLS.items():
                if campo not in m and re.match(pat, t):
                    m[campo] = c
                    break
    return m


def formato_generico(ws, max_fil=120, max_col=16, g=None):
    g = g if g is not None else grid(ws, max_fil, max_col)
    max_fil = min(max_fil, len(g) - 1)
    A = lambda r, c: g[r][c]
    filas = [[A(r, c) for c in range(0, max_col + 1)] for r in range(0, max_fil + 1)]

    campos_cab = {'nombre': '', 'codigo': '', 'unidad': '', 'detalle': ''}
    ROTULOS = [('nombre', r'(?:NOMBRE DEL |DESCRIPCI[ÓO]N DEL )?RUBRO'),
               ('codigo', r'C[ÓO]DIGO'),
               ('unidad', r'UNIDAD'),
               ('detalle', r'DETALLE|ESPECIFICACI[ÓO]N')]
    for r in range(1, min(20, max_fil) + 1):
        for c in range(1, max_col):
            celda = clean(A(r, c))
            if not celda:
                continue
            # el rotulo puede traer su valor en la MISMA celda ("RUBRO: REPLANTEO
            # MANUAL...") o en la de al lado; se mira primero dentro
            m = re.match(r'(?i)^\s*([A-ZÁÉÍÓÚÑ .]{3,30}?)\s*:\s*(.*)$', celda)
            t, dentro = (m.group(1).strip() + ':', clean(m.group(2))) if m else (celda, '')
            val = dentro or next((clean(A(r, k)) for k in range(c + 1, max_col + 1)
                                  if clean(A(r, k))), '')
            # si lo siguiente a la derecha es otro rotulo ("DETALLE:" seguido de
            # "R(H/U):"), esta celda esta vacia y no hay que robarle el valor al
            # vecino
            if val.endswith(':'):
                val = ''
            for campo, pat in ROTULOS:
                if not campos_cab[campo] and re.fullmatch(r'(?:%s)\s*:?' % pat, t, re.I):
                    campos_cab[campo] = val
                    break

    secs, r = {}, 1
    while r <= max_fil:
        sec, _ = _seccion_en(filas[r])
        if not sec:
            r += 1
            continue
        # la fila de encabezados va justo debajo; puede haber una segunda linea
        mapa, rr = {}, r
        while rr <= min(r + 4, max_fil):
            m = _mapa_columnas(filas[rr])
            if 'desc' in m and 'cant' in m:
                mapa, r = m, rr
                break
            rr += 1
        if not mapa:
            r += 1
            continue
        items, rr = [], r + 1
        while rr <= max_fil:
            texto = ' '.join(clean(x) for x in filas[rr][1:]).upper().strip()
            if texto.startswith(_FIN) or _seccion_en(filas[rr])[0]:
                break
            d = clean(A(rr, mapa['desc']))
            q = A(rr, mapa['cant'])
            u = A(rr, mapa['uni']) if 'uni' in mapa else None
            # el codigo de estructura ocupacional puede ir en su propia columna
            # ("PEON" | "EO E2"): pegado a la descripcion, el emparejamiento por
            # codigo vuelve a funcionar
            for c in range(mapa['desc'] + 1, mapa['cant']):
                t = clean(A(rr, c))
                if re.fullmatch(r'(?i)(E\.?O\.?|ESTR?\.? ?OC\.?)\s*[A-E][1-3]', t):
                    d = clean(d + ' ' + t)
                    break
            # un item sin cantidad no es un item: son los renglones de plantilla
            # ("Herramienta Menor 0% de M.O.") y las filas de relleno
            if num(q) is not None and not fila_vacia(d, q):
                items.append({'desc': d, 'uni': clean(u), 'cant': num(q)})
            rr += 1
        secs[sec] = items
        r = rr

    # Un nombre de rubro nunca es solo un numero: hay plantillas donde "RUBRO:"
    # lleva el numero y el nombre vive en "DETALLE:". Si es el caso, se cambian.
    if re.fullmatch(r'\d{1,3}', campos_cab['nombre'] or '') and campos_cab['detalle']:
        campos_cab['nombre'], campos_cab['detalle'] = campos_cab['detalle'], ''

    d = dict(campos_cab, secs=secs)
    d.update(pie(A, max_fil, max_col))
    return d


# ---------------------------------------------------------------------------
# Un solo libro con TODOS los rubros apilados en una hoja
#
# `leer_libro` da por hecho una hoja por rubro. Hay plantillas que ponen los 78
# APUs uno debajo de otro en la misma hoja, separados por su propia cabecera
# ("RUBRO : 3", "HOJA 3 DE 78"). Se corta la hoja en bloques y cada bloque se
# lee con el formato generico, que no depende de en que fila empiece nada.
# ---------------------------------------------------------------------------
_CABECERA_RUBRO = re.compile(r'(?i)^\s*RUBRO\s*(?:No\.?|N[°º])?\s*:\s*(\d{1,3})\b')
_HOJA_N = re.compile(r'(?i)\bHOJA\s*:?\s*(\d{1,3})\s*(?:DE|/)\s*\d{1,3}\b')


def cortes_por_rubro(g, max_fil, max_col):
    """[(n, fila_inicio)] de cada rubro apilado en la hoja.

    Se busca "RUBRO : n" y, si no aparece, "HOJA n DE 78", que es la otra marca
    con la que estas plantillas separan un analisis del siguiente.
    """
    cortes = []
    for r in range(1, max_fil + 1):
        for c in range(1, max_col + 1):
            t = clean(g[r][c])
            if not t:
                continue
            m = _CABECERA_RUBRO.match(t) or _HOJA_N.search(t)
            if m:
                n = int(m.group(1))
                if not cortes or cortes[-1][0] != n:
                    cortes.append((n, r))
                break
    return cortes


def leer_hoja_unica(ruta, hoja=None, max_col=16):
    """{n: apu} de un libro que apila todos los rubros en una sola hoja."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    hoja = hoja or wb.sheetnames[0]
    ws = wb[hoja]
    max_fil = ws.max_row or 0
    g = grid(ws, max_fil, max_col)
    cortes = cortes_por_rubro(g, max_fil, max_col)
    out = {}
    for i, (n, r0) in enumerate(cortes):
        r1 = cortes[i + 1][1] - 1 if i + 1 < len(cortes) else max_fil
        # el bloque se re-indexa desde 1 para que formato_generico lo lea como
        # si fuera una hoja suelta
        sub = [[None] * (max_col + 1)] + [g[r][:] for r in range(r0, r1 + 1)]
        out[n] = formato_generico(None, len(sub) - 1, max_col, g=sub)
    return out


FORMATOS = {'A': formato_a, 'B': formato_b, 'C': formato_c,
            'G': formato_generico}


def detectar(ws):
    """Prueba los tres formatos y devuelve (clave, apu) del que mejor lee.

    Gana el que saque mas items con nombre y unidad. Es fiable porque un mapa de
    columnas equivocado devuelve celdas vacias o basura, no items validos. Aun
    asi, revisa el resultado contra la hoja original antes de correr las 100+.
    """
    mejor, clave, punt = None, None, -1
    for k, fn in FORMATOS.items():
        try:
            d = fn(ws)
        except Exception:
            continue
        items = [x for v in d['secs'].values() for x in v]
        # un mapa de columnas equivocado deja items sin cantidad o sin unidad
        # en materiales, asi que se puntua la calidad, no solo el conteo
        p = len(items)
        p += sum(1 for x in items if x['cant'] is not None)
        p += sum(1 for x in d['secs'].get('MATERIALES', []) if x['uni'])
        p += (5 if d['nombre'] else 0) + (5 if d['unidad'] else 0)
        if p > punt:
            mejor, clave, punt = d, k, p
    return clave, mejor


def leer_libro(ruta, hojas, formato=None):
    """Lee {n: apu} de un libro. hojas es {n: 'nombre de hoja'}.

    formato None autodetecta con la primera hoja y aplica ese a todas, que es lo
    correcto: dentro de un mismo archivo la plantilla no cambia.
    """
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    orden = sorted(hojas)
    if formato is None:
        formato, _ = detectar(wb[hojas[orden[0]]])
    fn = FORMATOS[formato]
    return formato, {n: fn(wb[hojas[n]]) for n in orden}



def leer_presupuesto(ruta, hoja=None):
    """{n: item} del presupuesto. Delega en presupuesto.py, que mapea las
    columnas por su nombre; aqui solo se conserva el nombre historico."""
    import presupuesto
    return presupuesto.leer(ruta, hoja)[0]
