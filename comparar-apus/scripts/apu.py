"""Motor de comparacion de APUs: normalizacion, emparejamiento, clasificacion y reporte.

Este modulo NO sabe leer archivos. Tu escribes dos funciones que devuelven APUs
en la estructura de abajo, y este modulo hace el resto.

Estructura de un APU:
    {'codigo': '5044',            # opcional
     'nombre': 'CERRAM.PROVISIONAL TELA EMBAL. hl=2.1m',
     'unidad': 'm',
     'detalle': 'Tabla dura, pingos 0.50m...',   # '' si no aplica
     'secs': {'EQUIPOS':      [{'desc': ..., 'uni': ..., 'cant': 2.0}, ...],
              'MANO DE OBRA': [...],
              'MATERIALES':   [...],
              'TRANSPORTE':   [...]}}

Uso tipico:
    import apu
    ref = {i: leer_referencia(i) for i in range(1, N + 1)}
    ofe = {i: leer_oferta(i)     for i in range(1, N + 1)}
    apu.reporte(ref, ofe, '/ruta/Comparacion ... .xlsx',
                titulo_ref='archivo referencia.xlsx',
                titulo_ofe='archivo oferta.pdf',
                correspondencia='Hoja N <-> pagina N. Verificada en las 137.')
"""
import difflib
import re
import unicodedata
from decimal import Decimal, ROUND_HALF_UP

SECS = ['EQUIPOS', 'MANO DE OBRA', 'MATERIALES', 'TRANSPORTE']


# ===========================================================================
# 1. Normalizacion y lectura de celdas
# ===========================================================================

def clean(v):
    """Texto de una celda: sin saltos de linea, sin espacios repetidos."""
    if v is None:
        return ''
    s = str(v).replace('_x000D_', ' ')
    s = s.replace('\r', ' ').replace('\n', ' ').replace('\xa0', ' ')
    return re.sub(r'\s+', ' ', s).strip()


def sin_tildes(s):
    s = unicodedata.normalize('NFD', (s or '').upper())
    return ''.join(c for c in s if unicodedata.category(c) != 'Mn')


def nd(s):
    """Clave de comparacion: mayusculas, sin tildes, sin puntuacion.

    Absorbe las variantes de escritura que no cambian el contenido:
    'Kg'/'KG', '#12'/'No. 12', '100%'/'100 por ciento', guion bajo por guion.
    """
    s = sin_tildes(clean(s))
    s = s.replace('%', 'PORCIENTO').replace('POR CIENTO', 'PORCIENTO')
    s = s.replace('#', 'NO').replace('N°', 'NO')
    return re.sub(r'[^A-Z0-9]', '', s)


def num(v):
    """Numero de una celda. Acepta coma decimal y punto de miles: '1.342,86'."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip()
    if not t or t in ('-', '–'):
        return None
    if ',' in t:                       # formato europeo: el punto es de miles
        t = t.replace('.', '').replace(',', '.')
    try:
        return float(t)
    except ValueError:
        return None


CABECERAS = ('DESCRIPCION', 'DESCRIPCIÓN', 'CODIGO', 'CÓDIGO', 'UNIDAD',
             'CANTIDAD', 'TARIFA', 'A', 'B', 'R')


def fila_vacia(desc, cant):
    """True para los renglones de relleno y para los encabezados repetidos.

    Muchas plantillas dejan filas en blanco con ceros, o repiten la fila de
    titulos dentro del bloque. Si se cuelan aparecen como items fantasma.

    Un item siempre tiene descripcion: si viene vacia, o la fila es relleno o
    estas leyendo la columna equivocada. Ojo con lo segundo -- una descripcion
    vacia con cantidad distinta de cero es la senal tipica de un mapa de
    columnas mal elegido.
    """
    d = clean(desc)
    if d in ('', '0', '-', '.'):
        return True
    return d.upper() in CABECERAS and num(cant) in (None, 0.0)


def grid(ws, max_fil, max_col):
    """Hoja de openpyxl como matriz 1-indexada: g[fila][columna].

    Funciona con read_only=True, donde ws.cell() no esta disponible.
    """
    g = [[None] * (max_col + 1) for _ in range(max_fil + 1)]
    for i, fila in enumerate(ws.iter_rows(min_row=1, max_row=max_fil,
                                          min_col=1, max_col=max_col,
                                          values_only=True), 1):
        for j, v in enumerate(fila, 1):
            if j <= max_col:
                g[i][j] = v
    return g


def fnum(x):
    """Numero a texto sin ceros de cola, para mostrar en el reporte."""
    if x is None:
        return ''
    return ('%.10f' % x).rstrip('0').rstrip('.')


# ===========================================================================
# 2. Emparejamiento de items
# ===========================================================================
# El oferente reordena los items, los renombra o intercala algunos propios.
# Comparar por posicion produce falsos "falta" y "sobra" en cascada: basta un
# item extra al principio para que todo lo de abajo quede corrido. Por eso el
# emparejamiento va por descripcion, en pasadas de menor a mayor tolerancia.

UMBRAL = 0.62

VACIAS = {'DE', 'DEL', 'LA', 'EL', 'LOS', 'LAS', 'Y', 'O', 'A', 'EN', 'PARA',
          'CON', 'POR', 'TIPO', 'INC', 'INCLUYE'}


def palabras(s):
    return {w for w in re.split(r'[^A-Z0-9]+', sin_tildes(s)) if w and w not in VACIAS}


def cod_ocupacional(s):
    """Codigo de estructura ocupacional (E2, D2, B3, C1...) de la mano de obra.

    Identifica al obrero aunque el oferente reescriba la categoria:
    'ESTRUC. OCUPAC. E2 PEON' y 'Estr. Oc. E2 PEON' son el mismo E2.
    """
    m = re.search(r'\b(?:ESTRUC?\.?\s*)?(?:OCUPAC?\.?|OC\.?)\s*([A-E][1-3])\b',
                  (s or '').upper())
    return m.group(1) if m else None


def emparejar(a, b, sec=None):
    """Empareja items de referencia (a) y oferta (b).

    Devuelve [(item_ref, item_ofe, pos_ref, pos_ofe)] en el orden de la
    referencia, con los sobrantes de la oferta al final. Cualquiera de los dos
    items puede ser None: eso es un faltante o un adicional de verdad.
    """
    libres = list(range(len(b)))
    par = [None] * len(a)

    # 1) descripcion identica una vez normalizada
    for i, x in enumerate(a):
        for j in libres:
            if nd(b[j]['desc']) == nd(x['desc']):
                par[i] = j
                libres.remove(j)
                break

    # 2) descripcion parecida (recortes, abreviaturas, erratas)
    for i, x in enumerate(a):
        if par[i] is not None or not libres:
            continue
        mejor, punt = None, UMBRAL
        for j in libres:
            r = difflib.SequenceMatcher(None, nd(x['desc']), nd(b[j]['desc'])).ratio()
            if r > punt:
                mejor, punt = j, r
        if mejor is not None:
            par[i] = mejor
            libres.remove(mejor)

    # 3) las palabras de uno contenidas en las del otro
    #    CEMENTO / CEMENTO PORTLAND TIPO 1 ; MODULO DE ANDAMIO / ANDAMIO (MODULO)
    for i, x in enumerate(a):
        if par[i] is not None or not libres:
            continue
        px = palabras(x['desc'])
        if not px:
            continue
        mejor, punt = None, 0
        for j in libres:
            py = palabras(b[j]['desc'])
            com = px & py
            if com and (com == px or com == py):
                r = len(com) / max(len(px), len(py))
                if r > punt:
                    mejor, punt = j, r
        if mejor is not None:
            par[i] = mejor
            libres.remove(mejor)

    # 4) mano de obra: mismo codigo ocupacional aunque cambie el nombre
    if sec == 'MANO DE OBRA':
        for i, x in enumerate(a):
            if par[i] is not None or not libres:
                continue
            cx = cod_ocupacional(x['desc'])
            if not cx:
                continue
            for j in libres:
                if cod_ocupacional(b[j]['desc']) == cx:
                    par[i] = j
                    libres.remove(j)
                    break

    filas = [(a[i], b[par[i]] if par[i] is not None else None, i, par[i])
             for i in range(len(a))]
    filas += [(None, b[j], None, j) for j in sorted(libres)]
    return filas


def desordenado(filas):
    """True si los items emparejados no aparecen en el mismo orden relativo."""
    pos = [j for _, _, i, j in filas if i is not None and j is not None]
    return pos != sorted(pos)


# ===========================================================================
# 3. Clasificacion de diferencias
# ===========================================================================
# Un reporte donde todo es rojo no sirve para revisar una oferta. La gracia
# esta en separar lo que hay que corregir de lo que solo esta escrito distinto.

EQUIV_UNIDAD = [
    {'GALON', 'GAL', 'GLN', 'GALONES'},
    {'L', 'LT', 'LTR', 'LITRO', 'LITROS'},
    {'U', 'UN', 'UNIDAD', 'C/U'},
    {'KG', 'KGS', 'KILO', 'KILOS', 'KILOGRAMO'},
    {'M', 'ML', 'METRO'},
    {'M2', 'MT2'},
    {'M3', 'MT3'},
    {'SACO', 'SACOS', 'SC'},
    {'JGO', 'JUEGO'},
    {'ROLLO', 'ROLL'},
    {'GLB', 'GLOBAL'},
]


def misma_unidad(a, b):
    x, y = nd(a), nd(b)
    return x == y or any(x in g and y in g for g in EQUIV_UNIDAD)


def red2(x):
    """Redondeo a 2 decimales al estilo Excel (medio hacia arriba)."""
    return float(Decimal(repr(x)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))


def tipo_texto(ref, ofe, pdf=False):
    """Clasifica una diferencia de texto.

    None            -> identicos
    'MENOR'         -> mismo contenido, distinta escritura
    'TEXTO CORTADO' -> la oferta es un prefijo de la referencia (solo en PDF,
                       donde la celda impresa no alcanzo y el texto se trunco)
    'DIFERENCIA'    -> contenido distinto
    """
    if clean(ref) == clean(ofe):
        return None
    x, y = nd(ref), nd(ofe)
    if x == y:
        return 'MENOR'
    # La celda impresa se queda sin sitio y corta la descripcion. Vale como
    # truncamiento si lo que falta es la cola: o bien lo leido ya es largo, o
    # bien lo que falta son dos o tres caracteres ("CAMION CISTERNA 13 TON.
    # 10.000 L" por "... 10.000 LT.", que es exactamente lo que imprime el PDF).
    if pdf and x and y and x.startswith(y) and (len(y) > 25 or len(x) - len(y) <= 3):
        return 'TEXTO CORTADO'
    return 'DIFERENCIA'


def tipo_cantidad(ref, ofe):
    """Clasifica una diferencia de cantidad, y devuelve (campo, tipo).

    La direccion lo es todo, porque el pliego no pide que la oferta sea igual:
    pide que cubra el rubro. Si el oferente pone MAS cantidad esta destinando
    mas recursos a ejecutarlo, y eso no perjudica al contratante. Si pone MENOS,
    el rubro queda sin cubrir con lo que la referencia considero necesario, y
    eso hay que observarlo aunque la baja sea de decimales o venga de redondear
    la plantilla a dos decimales: el efecto sobre la obra es el mismo.

    La cantidad que queda en cero se separa porque es la mas grave de todas:
    el insumo desaparece del costo del rubro sin dejar rastro en el APU.
    """
    if ref is None or ofe is None:
        return ('CANTIDAD', None if ref == ofe else 'DIFERENCIA')
    if abs(ref - ofe) <= 1e-9:
        return ('CANTIDAD', None)
    if ofe > ref:
        return ('CANTIDAD MAYOR QUE LA REFERENCIA', 'MENOR')
    if ofe == 0:
        return ('CANTIDAD EN CERO', 'DIFERENCIA')
    return ('CANTIDAD MENOR QUE LA REFERENCIA', 'DIFERENCIA')


# Tolerancia legal: el oferente puede bajar sus indirectos hasta 5 puntos
# porcentuales respecto a los del presupuesto referencial, pero no subirlos.
TOLERANCIA_INDIRECTOS = 5.0


def margen(a):
    """Porcentaje de indirectos del APU, en puntos.

    Se prefiere el porcentaje DECLARADO, que es el que se compara contra el del
    presupuesto referencial. El margen efectivo (precio / costo directo) sirve de
    respaldo cuando la plantilla no lo escribe, pero no como primera opcion: el
    precio impreso viene redondeado a dos decimales y eso mueve el resultado unas
    centesimas en cada rubro, lo que convierte una sola decision del oferente en
    treinta hallazgos distintos.
    """
    p = num(a.get('indirecto_pct'))
    if p is not None:
        return p
    d, t = num(a.get('directo')), num(a.get('precio'))
    if d and t:
        return (t / d - 1) * 100
    return None


def tipo_indirectos(ref, ofe):
    """(texto, tipo) al comparar los indirectos de un rubro."""
    a, b = margen(ref), margen(ofe)
    if a is None or b is None:
        return None, None
    d = a - b                                    # cuanto BAJO la oferta
    if abs(d) <= 0.05:
        return None, None
    if d < 0:
        return ('subio %.2f puntos' % -d, 'DIFERENCIA')
    if d <= TOLERANCIA_INDIRECTOS + 1e-9:
        return ('bajo %.2f puntos, dentro de los %g permitidos'
                % (d, TOLERANCIA_INDIRECTOS), 'MENOR')
    return ('bajo %.2f puntos, mas de los %g permitidos'
            % (d, TOLERANCIA_INDIRECTOS), 'DIFERENCIA')


# ===========================================================================
# 4. Comparacion completa
# ===========================================================================

def detalle_comparable(ref, ofe, umbral=0.8):
    """False si la oferta dejo el detalle en blanco casi siempre.

    Muchas plantillas de oferta no imprimen la especificacion. Compararla
    entonces produce una diferencia por rubro que no dice nada y tapa el resto,
    asi que se detecta y se anota en las notas en vez de reportarlo.
    """
    con_ref = [n for n in ref if clean(ref[n].get('detalle'))]
    if not con_ref:
        return False
    vacios = sum(1 for n in con_ref if n in ofe and not clean(ofe[n].get('detalle')))
    return vacios / len(con_ref) < umbral


def detectar_patrones(ref, ofe):
    """(nomenclatura_mo_menor, unidades_equivalentes_menor) mirando la oferta.

    Las dos banderas responden a una decision de forma que el oferente tomo una
    vez y repitio en todo el documento: reescribir la nomenclatura de la mano de
    obra, o abreviar las unidades a su manera. Se detectan igual que las veria
    una persona -- si la mayoria de las diferencias de ese tipo son del patron y
    no casos sueltos -- y asi no hay que adivinarlas rubro por rubro.

    Un caso aislado NO activa la bandera: dos items con el mismo codigo
    ocupacional pueden ser un renombrado, pero treinta ya son una decision.
    """
    mo_total = mo_patron = un_total = un_patron = 0
    for n in ref:
        O = ofe.get(n)
        if O is None:
            continue
        for sec in SECS:
            for ia, ib, _, _ in emparejar(ref[n]['secs'].get(sec, []),
                                          O['secs'].get(sec, []), sec):
                if ia is None or ib is None:
                    continue
                if sec == 'MANO DE OBRA' and nd(ia['desc']) != nd(ib['desc']):
                    mo_total += 1
                    ca, cb = cod_ocupacional(ia['desc']), cod_ocupacional(ib['desc'])
                    mo_patron += bool(ca and ca == cb)
                if sec in ('MATERIALES', 'TRANSPORTE') and nd(ia['uni']) != nd(ib['uni']):
                    un_total += 1
                    un_patron += bool(misma_unidad(ia['uni'], ib['uni']))
    sistematico = lambda p, t: t >= 5 and p >= 0.6 * t
    return sistematico(mo_patron, mo_total), sistematico(un_patron, un_total)


def comparar(ref, ofe, pdf=False, comparar_detalle=None, comparar_indirectos=True,
             nomenclatura_mo_menor=None, unidades_equivalentes_menor=None):
    """Compara dos diccionarios {n: apu} y devuelve (filas, resumen).

    pdf                          la oferta viene de un PDF: activa TEXTO CORTADO.
    comparar_detalle             None decide solo: se compara salvo que la oferta
                                 traiga el campo vacio en casi todos los rubros.
    comparar_indirectos          compara el margen de cada rubro contra la
                                 tolerancia de 5 puntos. Ponlo en False solo si
                                 el lector no pudo sacar el pie del APU.
    nomenclatura_mo_menor        el oferente renombro sistematicamente la mano de
                                 obra manteniendo el codigo ocupacional. None lo
                                 detecta solo.
    unidades_equivalentes_menor  abrevia las unidades de otra forma (galon ->
                                 Gln). None lo detecta solo. Ojo: un cambio real
                                 de unidad (Kg -> u) se sigue reportando grave.
    """
    filas, resumen = [], []
    if comparar_detalle is None:
        comparar_detalle = detalle_comparable(ref, ofe)
    if nomenclatura_mo_menor is None or unidades_equivalentes_menor is None:
        mo, un = detectar_patrones(ref, ofe)
        if nomenclatura_mo_menor is None:
            nomenclatura_mo_menor = mo
        if unidades_equivalentes_menor is None:
            unidades_equivalentes_menor = un

    for n in sorted(ref):
        S = ref[n]
        O = ofe.get(n)

        if O is None:
            filas.append([n, S.get('codigo', ''), S['nombre'], 'DOCUMENTO', '',
                          'RUBRO AUSENTE', 'existe en la referencia',
                          'no esta en la oferta', 'DIFERENCIA'])
            resumen.append([n, S.get('codigo', ''), S['nombre'], '',
                            S['unidad'], '', 1, 0, 0, 0, 'REVISAR', None])
            continue

        def add(sec, item, campo, vs, vo, k):
            if k:
                filas.append([n, S.get('codigo', ''), S['nombre'], sec, item,
                              campo, vs, vo, k])

        add('CABECERA', '', 'NOMBRE DEL RUBRO', S['nombre'], O['nombre'],
            tipo_texto(S['nombre'], O['nombre'], pdf))
        add('CABECERA', '', 'UNIDAD DEL RUBRO', S['unidad'], O['unidad'],
            tipo_texto(S['unidad'], O['unidad']))
        if comparar_detalle:
            add('CABECERA', '', 'DETALLE / ESPECIFICACION',
                S.get('detalle', ''), O.get('detalle', ''),
                tipo_texto(S.get('detalle', ''), O.get('detalle', ''), pdf))

        for sec in SECS:
            a = S['secs'].get(sec, [])
            b = O['secs'].get(sec, [])
            pares = emparejar(a, b, sec)

            if len(a) != len(b):
                add(sec, '', 'N. DE ITEMS', str(len(a)), str(len(b)), 'MENOR')
            if desordenado(pares):
                add(sec, '', 'ORDEN DE LOS ITEMS',
                    ' | '.join(x['desc'] for x in a),
                    ' | '.join(x['desc'] for x in b), 'ORDEN')

            for ia, ib, pi, _ in pares:
                lbl = '#%s %s' % (pi + 1 if pi is not None else '-',
                                  ((ia or ib)['desc'] or '')[:50])
                if ia is None:
                    # un item de mas es mas recurso destinado al rubro: se anota,
                    # pero no es lo que hay que corregir
                    add(sec, lbl, 'ITEM ADICIONAL EN LA OFERTA', '', ib['desc'],
                        'MENOR')
                    continue
                if ib is None:
                    add(sec, lbl, 'ITEM FALTA EN LA OFERTA', ia['desc'], '',
                        'DIFERENCIA')
                    continue

                k = tipo_texto(ia['desc'], ib['desc'], pdf)
                if (k == 'DIFERENCIA' and nomenclatura_mo_menor
                        and sec == 'MANO DE OBRA'
                        and cod_ocupacional(ia['desc'])
                        and cod_ocupacional(ia['desc']) == cod_ocupacional(ib['desc'])):
                    k = 'MENOR'
                add(sec, lbl, 'DESCRIPCION', ia['desc'], ib['desc'], k)

                if sec in ('MATERIALES', 'TRANSPORTE'):
                    ku = tipo_texto(ia['uni'], ib['uni'])
                    if (ku == 'DIFERENCIA' and unidades_equivalentes_menor
                            and misma_unidad(ia['uni'], ib['uni'])):
                        ku = 'MENOR'
                    add(sec, lbl, 'UNIDAD', ia['uni'], ib['uni'], ku)

                campo_c, kc = tipo_cantidad(ia['cant'], ib['cant'])
                add(sec, lbl, campo_c, fnum(ia['cant']), fnum(ib['cant']), kc)

        mias = [r for r in filas if r[0] == n]
        dif = sum(1 for r in mias if r[-1] == 'DIFERENCIA')
        men = sum(1 for r in mias if r[-1] == 'MENOR')
        orn = sum(1 for r in mias if r[-1] == 'ORDEN')
        cor = sum(1 for r in mias if r[-1] in ('TEXTO CORTADO', 'REDONDEO'))
        estado = ('REVISAR' if dif else 'ORDEN / TEXTO CORTADO' if (orn or cor)
                  else 'OBSERVACION MENOR' if men else 'OK')
        # el ultimo campo es el precio unitario que sale del APU de la oferta;
        # el reporte lo contrasta con el de su propio presupuesto
        resumen.append([n, S.get('codigo', ''), S['nombre'], O['nombre'],
                        S['unidad'], O['unidad'], dif, cor, men, orn, estado,
                        num(O.get('precio'))])

    if comparar_indirectos:
        filas += _filas_indirectos(ref, ofe)
    return filas, resumen


def _filas_indirectos(ref, ofe):
    """Los indirectos son UNA decision del oferente, no 78.

    Casi siempre aplica el mismo margen a todos los rubros, asi que reportarlo
    rubro por rubro llenaria el informe con la misma linea repetida y taparia lo
    demas. Se agrupa por cuanto bajo, y se dice en que rubros. Por eso tampoco
    entra en el estado de cada rubro: es una observacion sobre la oferta.
    """
    import collections
    grupos = collections.defaultdict(list)
    for n in sorted(ref):
        if n not in ofe:
            continue
        a, b = margen(ref[n]), margen(ofe[n])
        txt, k = tipo_indirectos(ref[n], ofe[n])
        if k:
            grupos[(round(a - b, 1), txt, k)].append((n, a, b))

    filas = []
    for (_, txt, k), rs in sorted(grupos.items()):
        ns = [str(n) for n, _, _ in rs]
        lista = ', '.join(ns[:12]) + ('' if len(ns) <= 12 else ' y %d mas' % (len(ns) - 12))
        filas.append(['', '', 'OBSERVACION SOBRE TODA LA OFERTA', 'INDIRECTOS',
                      '%d rubro(s): %s' % (len(rs), lista), 'INDIRECTOS (%s)' % txt,
                      fnum(rs[0][1]) + ' %', fnum(rs[0][2]) + ' %', k])
    return filas


# ===========================================================================
# 4b. Presupuesto
# ===========================================================================
# El APU dice como se compone un rubro; el presupuesto dice cuanto de ese rubro
# se va a ejecutar. Un oferente puede copiar los APUs sin tocar una coma y aun
# asi bajar la cantidad de un rubro en el presupuesto, que es donde de verdad se
# paga. Por eso se comparan los dos, y por eso se comprueba ademas que el precio
# unitario del presupuesto de la oferta sea el que sale de su propio APU: si no
# cuadran, uno de los dos documentos no es el que el oferente va a ejecutar.

def comparar_presupuesto(pa, pb, ofe=None, tol=0.005, pdf=False):
    """Compara dos presupuestos {n: {desc, uni, cant, punit, total}}.

    ofe: los APUs de la oferta, para cruzar el precio unitario. Opcional.
    pdf: alguno de los dos presupuestos se leyo de un PDF. Lo unico que cambia
         es que una descripcion truncada al imprimir deja de ser DIFERENCIA y
         pasa a TEXTO CORTADO; sin esto, el recorte del PDF llena la hoja
         PRESUPUESTO de rojo y tapa los hallazgos de verdad.
    """
    filas = []

    def add(n, desc, campo, vs, vo, k):
        if k:
            filas.append([n, desc, campo, vs, vo, k])

    for n in sorted(pa):
        A, B = pa[n], pb.get(n)
        if B is None:
            add(n, A['desc'], 'RUBRO AUSENTE EN EL PRESUPUESTO DE LA OFERTA',
                A['desc'], '', 'DIFERENCIA')
            continue

        add(n, A['desc'], 'DESCRIPCION DEL RUBRO', A['desc'], B['desc'],
            tipo_texto(A['desc'], B['desc'], pdf))
        # la unidad vacia en un PDF no es un cambio de unidad: es una celda que
        # el lector no pudo separar de la descripcion que se desbordo encima
        if pdf and not clean(B['uni']):
            add(n, A['desc'], 'UNIDAD DEL RUBRO', A['uni'], '',
                'TEXTO CORTADO' if clean(A['uni']) else None)
        else:
            add(n, A['desc'], 'UNIDAD DEL RUBRO', A['uni'], B['uni'],
                None if misma_unidad(A['uni'], B['uni']) else 'DIFERENCIA')

        campo, k = tipo_cantidad(A['cant'], B['cant'])
        add(n, A['desc'], campo.replace('CANTIDAD', 'CANTIDAD DEL RUBRO'),
            fnum(A['cant']), fnum(B['cant']), k)

        # coherencia interna de la oferta: presupuesto contra su propio APU
        if ofe and n in ofe:
            p_apu, p_pre = num(ofe[n].get('precio')), num(B.get('punit'))
            if p_apu and p_pre and abs(p_apu - p_pre) > max(tol, abs(p_pre) * 1e-4):
                add(n, A['desc'], 'PRECIO DEL PRESUPUESTO CONTRA SU PROPIO APU',
                    'APU de la oferta: %s' % fnum(p_apu),
                    'presupuesto de la oferta: %s' % fnum(p_pre), 'DIFERENCIA')

    for n in sorted(pb):
        if n not in pa:
            add(n, pb[n]['desc'], 'RUBRO ADICIONAL EN LA OFERTA', '',
                pb[n]['desc'], 'MENOR')
    return filas


# ===========================================================================
# 5. Reporte en Excel
# ===========================================================================

def reporte(ref, ofe, salida, titulo_ref='', titulo_ofe='', correspondencia='',
            notas_extra=(), pres_ref=None, pres_ofe=None, **kw):
    """Compara y escribe el Excel con las hojas enlazadas por formulas.

    El equipo reclasifica en la hoja DIFERENCIAS y el RESUMEN, los colores y
    SOLO ERRORES se recalculan solos. Devuelve (filas, resumen) para armar el
    resumen del chat. Para Excel anterior a 2021 usa reporte_estatico().
    """
    import reporte as reporte_mod
    filas, resumen = comparar(ref, ofe, **kw)
    fpres = (comparar_presupuesto(pres_ref, pres_ofe, ofe, pdf=kw.get('pdf', False))
             if pres_ref and pres_ofe else [])
    reporte_mod.generar(salida, filas, resumen, fpres, pres_ref, pres_ofe,
                        titulo_ref, titulo_ofe, correspondencia, notas_extra,
                        TOLERANCIA_INDIRECTOS)
    return filas, resumen


def reporte_estatico(ref, ofe, salida, titulo_ref='', titulo_ofe='', correspondencia='',
                     notas_extra=(), pres_ref=None, pres_ofe=None, **kw):
    """Compara y escribe el Excel con los conteos fijos. kw se pasan tal cual a comparar().

    pres_ref / pres_ofe: los dos presupuestos, si se pudieron leer. Agregan la
    hoja PRESUPUESTO y sus hallazgos entran en SOLO ERRORES como los demas.
    """
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter

    filas, resumen = comparar(ref, ofe, **kw)
    fpres = (comparar_presupuesto(pres_ref, pres_ofe, ofe, pdf=kw.get('pdf', False))
             if pres_ref and pres_ofe else [])

    wb = openpyxl.Workbook()
    TH = Font(bold=True, color='FFFFFF', size=10)
    FH = PatternFill('solid', fgColor='1F3864')
    COLOR = {
        'DIFERENCIA': PatternFill('solid', fgColor='FFC7CE'),
        'REVISAR': PatternFill('solid', fgColor='FFC7CE'),
        'MENOR': PatternFill('solid', fgColor='FFF2CC'),
        'OBSERVACION MENOR': PatternFill('solid', fgColor='FFF2CC'),
        'ORDEN': PatternFill('solid', fgColor='DDEBF7'),
        'ORDEN DISTINTO': PatternFill('solid', fgColor='DDEBF7'),
        'REDONDEO': PatternFill('solid', fgColor='DDEBF7'),
        'TEXTO CORTADO': PatternFill('solid', fgColor='DDEBF7'),
        'ORDEN / TEXTO CORTADO': PatternFill('solid', fgColor='DDEBF7'),
        'OK': PatternFill('solid', fgColor='C6EFCE'),
    }
    borde = Border(*[Side(style='thin', color='BFBFBF')] * 4)

    def hoja(ws, cabeceras, datos, anchos, col_color=None):
        ws.append(cabeceras)
        for c in range(1, len(cabeceras) + 1):
            cel = ws.cell(1, c)
            cel.font, cel.fill = TH, FH
            cel.alignment = Alignment(horizontal='center', vertical='center',
                                      wrap_text=True)
        for d in datos:
            ws.append(d)
        for i, a in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(i)].width = a
        ws.freeze_panes = 'A2'
        ws.auto_filter.ref = ws.dimensions
        ws.row_dimensions[1].height = 30
        if col_color:
            for r in range(2, ws.max_row + 1):
                f = COLOR.get(ws.cell(r, col_color).value)
                if f:
                    ws.cell(r, col_color).fill = f
        for fila in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=len(cabeceras)):
            for c in fila:
                c.border = borde
                if c.row > 1:
                    c.alignment = Alignment(vertical='top', wrap_text=True)
                    c.font = Font(size=10)

    ws = wb.active
    ws.title = 'RESUMEN'
    hoja(ws, ['RUBRO N.', 'CODIGO', 'RUBRO EN LA REFERENCIA', 'RUBRO EN LA OFERTA',
              'UNIDAD REFERENCIA', 'UNIDAD OFERTA', 'GRAVES',
              'ORDEN / CORTE', 'MENORES', 'DIF. ORDEN', 'ESTADO'],
         resumen, [10, 10, 48, 48, 15, 14, 12, 13, 13, 11, 22], 11)

    HD = ['RUBRO N.', 'CODIGO', 'RUBRO', 'SECCION', 'ITEM', 'CAMPO',
          'VALOR EN LA REFERENCIA', 'VALOR EN LA OFERTA', 'TIPO']
    AN = [10, 10, 42, 15, 42, 26, 46, 46, 16]
    hoja(wb.create_sheet('DIFERENCIAS'), HD, filas, AN, 9)

    if fpres:
        hoja(wb.create_sheet('PRESUPUESTO'),
             ['RUBRO N.', 'RUBRO', 'CAMPO', 'VALOR EN LA REFERENCIA',
              'VALOR EN LA OFERTA', 'TIPO'], fpres, [10, 52, 42, 40, 40, 16], 6)

    graves = [r for r in filas if r[-1] == 'DIFERENCIA']
    graves += [[r[0], '', r[1], 'PRESUPUESTO', '', r[2], r[3], r[4], r[5]]
               for r in fpres if r[-1] == 'DIFERENCIA']
    hoja(wb.create_sheet('SOLO ERRORES'), HD, graves, AN, 9)

    cnt = lambda v: str(sum(1 for r in resumen if r[10] == v))
    notas = [
        ['CRITERIO DE COMPARACION', ''],
        ['Referencia', titulo_ref],
        ['Documento analizado', titulo_ofe],
        ['Correspondencia', correspondencia],
        ['Campos comparados',
         'Nombre del rubro, unidad, indirectos y, en cada seccion (EQUIPOS / MANO '
         'DE OBRA / MATERIALES / TRANSPORTE): descripcion, cantidad y unidad de '
         'cada item. Si se leyeron los dos presupuestos, tambien la cantidad de '
         'cada rubro y el precio unitario contra el APU de la propia oferta.'],
        ['Campos NO comparados',
         'Tarifas, jornales, precios unitarios de insumos, rendimientos y costos: '
         'son propios de cada oferente. El monto de la oferta tampoco se observa, '
         'porque ya puntua por si mismo en la calificacion economica.'],
        ['Emparejamiento de items',
         'Los items se emparejan por descripcion, no por posicion. Un item listado '
         'en otro orden NO se reporta como faltante. "ITEM FALTA" y "ITEM ADICIONAL" '
         'solo aparecen cuando el item no esta en ninguna posicion de la seccion.'],
        ['', ''],
        ['COMO SE GRADUA CADA HALLAZGO', ''],
        ['El criterio',
         'La oferta no tiene que ser identica: tiene que cubrir el rubro. Lo que se '
         'observa es que FALTE algo o que venga de MENOS. Que sobre, no.'],
        ['DIFERENCIA (grave)',
         'Item que falta, cantidad menor a la de la referencia (aunque la baja sea '
         'de decimales o venga de redondear), cantidad que queda en cero, unidad '
         'cambiada, insumo sustituido, cambio de categoria salarial, indirectos '
         'fuera de la tolerancia y, en el presupuesto, cantidad de rubro menor o '
         'precio que no cuadra con el APU de la propia oferta.'],
        ['MENOR',
         'Item adicional, cantidad mayor a la de la referencia, mismo contenido '
         'escrito distinto e indirectos dentro de la tolerancia. Se anota, pero no '
         'es lo que hay que corregir: el oferente destina mas recursos al rubro.'],
        ['ORDEN', 'Los items estan en distinto orden. No es error por si mismo: '
                  'las cantidades y unidades igual se comparan contra el item correcto.'],
        ['TEXTO CORTADO', 'El texto coincide pero aparece truncado en el PDF porque '
                          'la celda impresa no alcanzo. Verificar en el original.'],
        ['Indirectos',
         'El oferente puede bajar sus indirectos hasta %g puntos porcentuales '
         'respecto a los de la referencia, pero no subirlos. Dentro de esa '
         'tolerancia el hallazgo es MENOR; fuera, grave.' % TOLERANCIA_INDIRECTOS],
        ['', ''],
        ['RESULTADO', ''],
        ['Rubros con hallazgos graves', cnt('REVISAR')],
        ['Rubros solo con observaciones menores', cnt('OBSERVACION MENOR')],
        ['Rubros con orden distinto o texto cortado', cnt('ORDEN / TEXTO CORTADO')],
        ['Rubros correctos', cnt('OK')],
        ['Total de rubros revisados', str(len(resumen))],
        ['Hallazgos en el presupuesto',
         str(sum(1 for r in fpres if r[-1] == 'DIFERENCIA')) + ' graves de '
         + str(len(fpres)) if fpres else 'no se compararon los presupuestos'],
    ]
    # notas_extra son pares [concepto, detalle]: un texto suelto se recorreria
    # letra por letra y llenaria la hoja de columnas
    notas += [list(x) if isinstance(x, (list, tuple)) else ['', str(x)]
              for x in notas_extra]
    hoja(wb.create_sheet('NOTAS'), ['CONCEPTO', 'DETALLE'], notas, [40, 118])

    wb.save(salida)
    return filas, resumen


def resumir(filas, resumen):
    """Conteos para reportar en el chat."""
    import collections
    return {
        'estados': collections.Counter(r[10] for r in resumen),
        'campos': collections.Counter(r[5] for r in filas if r[-1] == 'DIFERENCIA'),
        'errores': sum(1 for r in filas if r[-1] == 'DIFERENCIA'),
        'menores': collections.Counter(r[5] for r in filas if r[-1] == 'MENOR'),
    }
