"""Lector de la tabla de rubros (presupuesto / cantidades y precios) en PDF.

Este modulo existe por un error que costo caro: leer un presupuesto en PDF
partiendo el texto por huecos de espacios produce hallazgos que no existen.

En estas plantillas, una descripcion larga se imprime en TRES renglones -la
cabeza arriba, las cifras en el del medio, la cola abajo- y entre dos filas de
cifras pueden convivir la cola de una y la cabeza de la siguiente. Partiendo por
columnas de texto no hay forma de saber cual es cual, asi que el rubro se queda
con el trozo del medio y el vecino hereda lo que sobra. En un caso real:

    75  ROMBO 5cm, RESISTENTE A LA INTEMPERIE CON     <- perdio la cabeza
    76  PROTECCION UV. DESALOJO EQUIPO PESADO ...     <- heredo la cola del 75
    77  SOBREACARREO DE ESCOMBROS, TIERRA DE NARVAEZ  <- se colo una firma

Ninguno de los tres era un hallazgo: los tres eran el lector.

La solucion es no adivinar: estas tablas estan rayadas, y pdfplumber devuelve el
contenido de cada CELDA, con sus tres renglones ya unidos y sin nada de la celda
vecina. Se usa eso siempre que se pueda y solo se cae al modo por espacios si el
PDF no trae lineas de tabla.
"""
import re

from apu import clean, sin_tildes

# Las firmas electronicas se imprimen ENCIMA de la tabla, no al final: sus
# renglones aterrizan dentro de las celdas y se leen como parte del rubro.
FIRMA = re.compile(
    r'firmado\s+electr[oó]nicamente|validar\s+[uú]nicamente|firmaec|'
    r'ir\s+a\s+la\s+firma|al\s+menos\s+una\s+firma', re.I)

# Como se llama cada columna en las plantillas vistas. El texto llega ya en
# mayusculas y sin tildes, asi que los patrones se escriben tal cual: pasarlos
# por sin_tildes() los rompe, porque .upper() convierte \w en \W.
COLUMNAS = {
    'n':      (r'^ITEM$', r'^ITEM ?N', r'^N[°º]?$', r'^NO\.?$', r'^NUMER',
               r'^RUBRO ?N'),
    'codigo': (r'^CODIGO', r'^COD'),
    'desc':   (r'^DESCRIPCION', r'^RUBRO', r'^DETALLE'),
    'uni':    (r'^UNIDAD', r'^UND$', r'^U\.?$'),
    'cant':   (r'^CANTIDAD', r'^CANT'),
    # ojo con el solapamiento: "PRECIO UNITARIO" y "PRECIO TOTAL" empiezan igual,
    # asi que el patron del unitario exige la palabra completa
    'punit':  (r'^PREC\w* ?UNITARIO', r'^P\.? ?UNITARIO', r'^V\.? ?UNITARIO'),
    'total':  (r'^PRECIO ?GLOBAL', r'^SUBTOTAL', r'^TOTAL', r'^VALOR ?TOTAL',
               r'^PREC\w* ?TOTAL'),
}


def _cual(texto):
    t = re.sub(r'\s+', ' ', sin_tildes(clean(texto))).strip()
    if not t or 'CPC' in t:
        return None
    for campo, patrones in COLUMNAS.items():
        for p in patrones:
            if re.match(p, t):
                return campo
    return None


def _num(t):
    """1.234,56 y 1,234.56 a float. El separador de miles es el que se repite."""
    t = clean(t).replace(' ', '')
    if not t or not re.search(r'\d', t):
        return None
    t = re.sub(r'[^\d.,-]', '', t)
    if ',' in t and '.' in t:
        t = t.replace('.', '').replace(',', '.') if t.rfind(',') > t.rfind('.') \
            else t.replace(',', '')
    elif ',' in t:
        # 2,800 son dos mil ochocientos; 2,80 son dos con ocho
        ent, _, dec = t.rpartition(',')
        t = t.replace(',', '') if len(dec) == 3 and ent else t.replace(',', '.')
    try:
        return float(t)
    except ValueError:
        return None


def _limpiar(celda):
    """Texto de una celda, sin los renglones de firma que cayeron encima."""
    if not celda:
        return ''
    vivas = [l for l in str(celda).split('\n') if not FIRMA.search(l)]
    # el nombre del firmante queda suelto bajo su rotulo: si la firma ocupaba
    # el renglon anterior, lo que sigue en MAYUSCULAS sin minusculas es suyo
    if len(vivas) < len(str(celda).split('\n')):
        vivas = [l for l in vivas
                 if not (l.strip() and l.strip() == l.strip().upper()
                         and len(l.strip().split()) <= 3 and not re.search(r'\d', l))]
    return clean(' '.join(vivas))


def _cuadra(it, tol=0.02):
    """La fila se comprueba sola: cantidad x precio unitario == precio total.

    Es el unico juez fiable que tiene un lector de PDF. Una fila mal leida casi
    nunca cuadra, y una bien leida casi siempre.
    """
    c, p, t = it.get('cant'), it.get('punit'), it.get('total')
    if c is None or p is None or t is None:
        return None
    return abs(c * p - t) <= max(tol, abs(t) * 0.001)


def _unidad_plausible(u):
    """Una unidad es un token corto: m, m2, m3, u, Kg, Gln, m3-km, saco.

    Cuando la descripcion no cabe en su celda, pdfplumber le entrega el sobrante
    a la columna de al lado y la unidad deja de parecerlo ('A INTEmM2PERIE' es
    'A LA INTEMPERIE' y 'm2' impresos uno encima del otro). Es la senal de que
    esa fila hay que leerla por espacios, aunque sus cifras cuadren.
    """
    return bool(re.fullmatch(r'[A-Za-z][A-Za-z0-9/.\-]{0,5}', clean(u)))


def leer_plumber(ruta, min_items=5):
    """Lectura por celdas. Exacta mientras el texto quepa en su celda."""
    try:
        import pdfplumber
    except ImportError:
        return {}

    out, cab = {}, None
    with pdfplumber.open(ruta) as pdf:
        for pagina in pdf.pages:
            for tabla in pagina.extract_tables() or []:
                for fila in tabla:
                    celdas = [_limpiar(c) for c in fila]
                    mapa = {}
                    for j, c in enumerate(celdas):
                        campo = _cual(c)
                        if campo and campo not in mapa:
                            mapa[campo] = j
                    # La cabecera de la tabla de rubros nombra el NUMERO de
                    # item, ademas de la descripcion y la cantidad. Exigir el
                    # numero es lo que impide confundirla con la cabecera de una
                    # seccion de APU ("CODIGO DESCRIPCION CANTIDAD TARIFA ...
                    # TOTAL COSTO"), que si no produce una tabla de rubros
                    # inventada a partir de los insumos.
                    if 'n' in mapa and 'desc' in mapa and 'cant' in mapa:
                        cab = mapa
                        continue
                    if not cab:
                        continue
                    item = {k: (celdas[j] if j < len(celdas) else '')
                            for k, j in cab.items()}
                    n = _num(item.get('n'))
                    cant = _num(item.get('cant'))
                    # las filas de capitulo no llevan cantidad; las de detalle
                    # bajo un rubro no llevan numero
                    if n is None or cant is None or not item.get('desc'):
                        continue
                    # Manda la primera aparicion: despues del presupuesto, estos
                    # documentos suelen repetir los 78 rubros en la tabla del VAE
                    # o de la desagregacion, con otras columnas, y esa segunda
                    # tabla pisaba la buena.
                    if int(n) in out:
                        continue
                    out[int(n)] = {
                        'codigo': item.get('codigo', ''),
                        'desc': item.get('desc', ''),
                        'uni': item.get('uni', ''),
                        'cant': cant,
                        'punit': _num(item.get('punit')),
                        'total': _num(item.get('total')),
                        'hoja': 'PDF', 'fila': None,
                    }
    return out if len(out) >= min_items else {}


# ---------------------------------------------------------------------------
# Lectura por espacios, para el PDF sin lineas de tabla y para reparar la fila
# cuyo texto se desbordo de su celda (pdfplumber la corta a media palabra y le
# entrega el sobrante a la columna vecina, que asi deja de ser un numero).
# ---------------------------------------------------------------------------

def _toks(linea):
    return [t for t in re.split(r'\s{2,}', linea.strip()) if t]


def _es_num(t):
    return bool(re.fullmatch(r'-?[\d.,]*\d[\d.,]*%?', t))


def _fila(toks):
    """(n, codigo, desc, uni, cant, punit, total) de una fila de la tabla.

    La unidad es el ancla: es el ULTIMO token no numerico de la fila, y detras
    vienen siempre cantidad, precio unitario y precio total en ese orden. Las
    columnas que algunas plantillas agregan despues (peso relativo, VAE) no
    estorban, porque se cuentan desde la unidad hacia la derecha y no desde el
    final.
    """
    ult = max((i for i, t in enumerate(toks) if not _es_num(t)), default=-1)
    if ult < 0 or len(toks) < ult + 4:
        return None
    uni, cifras = toks[ult], toks[ult + 1:ult + 4]
    izq = toks[:ult]
    n = codigo = None
    if izq and re.fullmatch(r'\d{1,3}', izq[0]):
        n, izq = int(izq[0]), izq[1:]
    if izq and re.fullmatch(r'[A-Z]?\d{3,5}', izq[0]):
        codigo, izq = izq[0], izq[1:]
    return (n, codigo or '', ' '.join(izq), uni,
            _num(cifras[0]), _num(cifras[1]), _num(cifras[2]))


def leer_layout(ruta, min_items=5):
    """Lectura por espacios con pdftotext -layout.

    La descripcion larga se imprime en tres renglones: cabeza arriba, cifras en
    el del medio, cola abajo. Entre dos filas de cifras conviven la cola de una
    y la cabeza de la siguiente; el reparto es "el primer renglon suelto es la
    cola de la de arriba, el resto es la cabeza de la de abajo", que es como lo
    compone Excel al imprimir.
    """
    import subprocess
    txt = subprocess.run(['pdftotext', '-layout', ruta, '-'],
                         capture_output=True, text=True).stdout
    lineas = [l for l in txt.replace('\f', '\n').split('\n') if not FIRMA.search(l)]

    filas, sueltas, pend_n = [], [], None
    for l in lineas:
        t = _toks(l)
        if not t:
            continue
        if len(t) == 1 and re.fullmatch(r'\d{1,3}', t[0]):
            pend_n = int(t[0])          # el numero solo, en su propio renglon
            continue
        f = _fila(t)
        if f and (f[0] is not None or pend_n is not None):
            n = f[0] if f[0] is not None else pend_n
            # Con el numero de rubro en su propio renglon, TODO lo suelto que
            # hay encima es la cabeza de esta fila. Con el numero en la misma
            # linea de las cifras no hay esa marca, y entonces vale la forma en
            # que Excel compone la celda: el primer renglon suelto es la cola de
            # la fila de arriba y el resto es la cabeza de esta.
            if filas and sueltas and pend_n is None:
                filas[-1][1].append(sueltas[0])
                sueltas = sueltas[1:]
            filas.append([n, list(sueltas) + ([f[2]] if f[2] else []), f])
            sueltas, pend_n = [], None
        elif not any(_es_num(x) for x in t):
            sueltas.append(' '.join(t))
    if filas and sueltas:
        filas[-1][1].append(sueltas[0])

    out = {}
    for n, partes, f in filas:
        if n is None or n in out:      # manda la primera aparicion, ver arriba
            continue
        out[n] = {'codigo': f[1], 'desc': clean(' '.join(partes)), 'uni': f[3],
                  'cant': f[4], 'punit': f[5], 'total': f[6],
                  'hoja': 'PDF', 'fila': None}
    return out if len(out) >= min_items else {}


def _es_tabla(items, min_items=5):
    """Forma de tabla de rubros: numerada del 1 al N y casi sin huecos.

    Es una puerta de FORMA, no de calidad: se aplica a cada lector por separado
    para descartar lo que no era una tabla de rubros -el mismo lector sobre un
    PDF de APUs devuelve los insumos de cada rubro como si fueran rubros-, pero
    sin exigirle todavia que este bien leida, porque de eso se encarga el cruce
    entre los dos lectores.
    """
    return bool(items) and len(items) >= min_items and len(items) >= 0.6 * max(items)


def _bien_leida(items):
    """Calidad de la lectura YA cruzada: unidades que lo parecen y filas que
    cuadran. Si no llega, es mejor devolver nada que un presupuesto inventado."""
    if not items:
        return False
    unis = sum(1 for f in items.values() if _unidad_plausible(f.get('uni')))
    if unis < 0.7 * len(items):
        return False
    juicios = [_cuadra(f) for f in items.values()]
    ciertos = [j for j in juicios if j is not None]
    return not ciertos or sum(ciertos) >= 0.7 * len(ciertos)


def leer(ruta, min_items=5):
    """{n: {codigo, desc, uni, cant, punit, total}} de un presupuesto en PDF.

    Lee de las dos maneras y se queda, fila por fila, con la que cuadre la
    aritmetica (cantidad x precio unitario = precio total) y traiga una unidad
    que parezca una unidad. Cuando las dos valen gana la de celdas, que es la
    que no se inventa descripciones. Devuelve {} si el PDF no trae una tabla de
    rubros reconocible, para que quien llame pueda seguir sin el presupuesto en
    vez de reventar o de comparar contra algo inventado.
    """
    a = leer_plumber(ruta, min_items)
    b = leer_layout(ruta, min_items)
    if not _es_tabla(a, min_items):
        a = {}
    if not _es_tabla(b, min_items):
        b = {}
    if not a:
        a, b = b, {}
    for n, fila_b in b.items():
        fila_a = a.get(n)
        if fila_a is None:
            a[n] = fila_b
            continue
        rota = (_cuadra(fila_a) is False
                or not _unidad_plausible(fila_a.get('uni')))
        sana = _cuadra(fila_b) is not False and _unidad_plausible(fila_b.get('uni'))
        if rota and sana:
            a[n] = fila_b
    # lo que sigue sin parecer una unidad no lo es: se deja vacio, porque una
    # unidad inventada se compara y sale como un cambio de unidad que no existe
    if not _bien_leida(a):
        return {}
    for fila in a.values():
        if not _unidad_plausible(fila.get('uni')):
            fila['uni'] = ''
    return a


def inconsistentes(items):
    """Numeros de rubro cuya aritmetica no cuadra: lo que hay que mirar a mano."""
    return sorted(n for n, it in items.items() if _cuadra(it) is False)
