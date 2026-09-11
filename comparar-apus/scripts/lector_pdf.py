"""Lector de APUs en PDF (un rubro por pagina, impreso desde Excel).

Se apoya en `pdftotext -layout`, que conserva las columnas como espacios. No
necesita OCR: estos PDFs traen texto real. Comprueba siempre con
`pdftotext -layout archivo.pdf - | head -60` antes de programar nada.

Dos cosas de este formato dan mas guerra de lo que parece:

1. Una descripcion larga se parte en varios renglones. El renglon con las cifras
   puede ser el primero, el del medio o el ultimo del grupo, y a veces un
   renglon suelto arrastra una sola cifra (el costo). Por eso un item se
   reconoce por traer DOS O MAS cifras al final, y los renglones con una o
   ninguna se acumulan como texto de la descripcion.

2. En la cabecera, Excel centra verticalmente el texto de la celda: un valor de
   tres lineas deja la etiqueta ("RUBRO:", "DETALLE") en la del medio. Por eso
   cada linea suelta se asigna a la etiqueta mas cercana, no a la de arriba.
"""
import re
import subprocess

from apu import SECS, clean, fila_vacia

# 1.234,56  |  0,8300  |  12  |  5,00%  |  1234.56
NUMTOK = re.compile(r'^-?\d{1,3}(?:[.,]\d{3})*(?:[.,]\d+)?%?$|^-?\d+(?:[.,]\d+)?%?$')


def paginas(pdf):
    txt = subprocess.run(['pdftotext', '-layout', pdf, '-'],
                         capture_output=True, text=True).stdout
    return txt.split('\f')


def _f(t, coma_decimal):
    t = (t or '').strip()
    if not t or t.endswith('%') or t in ('-', '–'):
        return None
    t = t.replace('.', '').replace(',', '.') if coma_decimal else t.replace(',', '')
    try:
        return float(t)
    except ValueError:
        return None


def _cifras_finales(linea):
    """(texto de la izquierda, [cifras pegadas a la derecha])."""
    toks = re.split(r'\s{2,}', linea.strip())
    i = len(toks)
    nums = []
    while i > 0:
        t = toks[i - 1].strip()
        if re.fullmatch(r'[-–]', t) or NUMTOK.match(t):
            nums.insert(0, t)
            i -= 1
        else:
            break
    return ' '.join(toks[:i]).strip(), nums


def parse_pagina(txt, etiquetas, coma_decimal=True, modo='derecha'):
    """Un APU a partir del texto de una pagina.

    etiquetas: {'num': regex con un grupo para el numero de rubro,
                'nombre': regex de la etiqueta, 'unidad': ..., 'detalle': ...}
               Usa None en las que no existan en ese formato.

    modo: de donde se leen las columnas del item.
      'derecha'   la fila termina en las cifras del APU. Se identifica un item
                  por traer dos o mas cifras al final.
      'izquierda' la fila sigue con columnas ajenas al APU (Peso Relativo, CPC
                  Elemento, NP/ND/EP, VAE %). Barrer desde la derecha se
                  atasca en el token 'EP'/'NP', asi que la cantidad se toma
                  como la primera cifra a partir de la descripcion.
    """
    lineas = txt.split('\n')

    # --- cabecera: solo lo anterior a la primera seccion --------------------
    fin = len(lineas)
    for i, ln in enumerate(lineas):
        if re.sub(r'\s{2,}.*$', '', ln.strip()).upper() in SECS:
            fin = i
            break
    head = lineas[:fin]

    num, ini = None, 0
    if etiquetas.get('num'):
        for i, ln in enumerate(head):
            m = re.search(etiquetas['num'], ln)
            if m:
                num, ini = int(m.group(1)), i + 1
                break

    # El valor de una celda vive en su columna: lo que venga despues de un hueco
    # ancho pertenece a otra columna de la misma fila (UNIDAD:, RENDIMIENTO:...)
    corte = lambda s: clean(re.sub(r'\s{2,}.*$', '', s.strip()))
    ignorar = etiquetas.get('ignorar')

    campos = {k: [] for k in ('nombre', 'unidad', 'detalle')}
    pos, sueltos = {}, []
    for i in range(ini, len(head)):
        s = head[i].strip()
        for k in ('nombre', 'unidad', 'detalle'):
            pat = etiquetas.get(k)
            if not pat or k in pos:
                continue
            m = re.search(pat, s)
            if not m:
                continue
            # una misma fila puede llevar el final de una celda a la izquierda y
            # la etiqueta de otra columna a la derecha
            izq = corte(s[:m.start()])
            if izq:
                sueltos.append((i, izq))
            resto = corte(s[m.end():])
            # la etiqueta puede salir vacia arriba y repetirse con el valor mas
            # abajo; nos quedamos con la aparicion que trae contenido
            if resto:
                pos[k] = i
                campos[k] = [(i, resto)]
            elif k not in pos:
                pos[k] = i
            break
    # una celda alta reparte su texto arriba y abajo de la etiqueta, porque Excel
    # lo centra verticalmente: cada fragmento va a la etiqueta mas cercana
    sueltos += [(i, corte(head[i])) for i in range(ini, len(head))
                if head[i].strip() and i not in pos.values()]
    # la unidad es un token corto que nunca envuelve, asi que no recibe fragmentos
    destinos = [k for k in pos if k != 'unidad']
    for i, s in sueltos:
        if not s or not destinos or (ignorar and re.search(r'\b(?:%s)' % ignorar, s)):
            continue
        cerca = min(destinos, key=lambda k: abs(pos[k] - i))
        if abs(pos[cerca] - i) <= 2:
            campos[cerca].append((i, s))

    unir = lambda k: clean(' '.join(t for _, t in sorted(campos[k])))
    nombre, detalle = unir('nombre'), unir('detalle')
    # la unidad es un token corto: nunca envuelve a las lineas vecinas
    unidad = clean(' '.join(t for i, t in campos['unidad']))

    # --- secciones ----------------------------------------------------------
    secs, cur, pend, sufijo = {}, None, [], False
    for ln in lineas:
        s = ln.strip()
        if not s:
            continue
        cab = re.sub(r'\s{2,}.*$', '', s).strip().upper()
        # el ancho de la pagina puede arrastrar titulos de columna junto al
        # nombre de la seccion, pero nunca cifras
        if cab in SECS and not re.search(r'\d', s):
            cur, pend, sufijo = cab, [], False
            secs[cur] = []
            continue
        if cur is None:
            continue
        if re.match(r'SUBTOTAL|TOTAL COSTO|COSTOS? INDIRECTO|OTROS INDIRECTOS|'
                    r'UTILIDAD|VALOR OFERTADO|ESTOS PRECIOS|Fecha', s, re.I):
            if re.match(r'TOTAL COSTO|COSTO TOTAL|UTILIDAD|VALOR OFERTADO|ESTOS PRECIOS',
                        s, re.I):
                cur = None
            continue
        if re.match(r'DESCRIPCI[ÓO]N|Peso Relativo|CPC|VAE|NP/EP|elemento\b', s, re.I):
            continue

        if modo == 'izquierda':
            toks = [t.strip() for t in re.split(r'\s{2,}', s) if t.strip()]
            k = next((i for i, t in enumerate(toks) if NUMTOK.match(t)), None)
            if k is None:                        # descripcion partida en dos renglones
                pend.append(s.strip())
                continue
            if cur in ('EQUIPOS', 'MANO DE OBRA'):
                desc, uni = ' '.join(toks[:k]), ''
            else:
                if k == 0:
                    continue                     # renglon sin descripcion: relleno
                desc, uni = ' '.join(toks[:k - 1]), toks[k - 1]
            if pend:
                desc = clean(' '.join(pend) + ' ' + desc)
                pend = []
            cant = _f(toks[k], coma_decimal)
            desc = clean(desc)
            if not fila_vacia(desc, cant):
                secs[cur].append({'desc': desc, 'uni': clean(uni), 'cant': cant})
            continue

        izq, nums = _cifras_finales(s)
        reales = [x for x in nums if _f(x, coma_decimal) is not None]
        if len(reales) < 2:                      # renglon de descripcion partida
            if not izq:
                continue
            if sufijo and secs.get(cur):
                secs[cur][-1]['desc'] = clean(secs[cur][-1]['desc'] + ' ' + izq)
                sufijo = False
            else:
                pend.append(izq)
            continue
        sufijo = (not izq) and bool(pend)
        if not izq and not pend:
            continue

        pref = clean(' '.join(pend))
        pend = []
        if pref:
            izq = clean(pref + ' ' + izq)

        if cur in ('EQUIPOS', 'MANO DE OBRA'):
            desc, uni = izq, ''
        else:
            toks = re.split(r'\s{2,}', s.strip())
            j = len(toks) - len(nums)
            if j - 1 >= 0:
                uni, desc = toks[j - 1].strip(), ' '.join(toks[:j - 1]).strip()
                if NUMTOK.match(uni):
                    desc, uni = izq, ''
                elif pref:
                    desc = clean(pref + ' ' + desc)
            else:
                desc, uni = izq, ''
        cant = _f(nums[0], coma_decimal)
        desc = clean(desc)
        if not fila_vacia(desc, cant):
            secs[cur].append({'desc': desc, 'uni': clean(uni), 'cant': cant})

    return {'n': num, 'codigo': '', 'nombre': nombre, 'unidad': unidad,
            'detalle': detalle, 'secs': secs}



# ---------------------------------------------------------------------------
# Formato UEM: la columna CODIGO va a la izquierda de cada item, el codigo del
# rubro es alfanumerico (5044, P524, V302), la seccion cierra en "PARCIAL M/N/O/P"
# y la cabecera de las columnas de VAE se imprime partida en fragmentos sueltos
# ("PESO / RELATI / VO") repartidos alrededor del nombre de la seccion.
# ---------------------------------------------------------------------------
COD_RUBRO = re.compile(r'(RUBRO\s*No\.?\s*:?\s*)([A-Za-z0-9\-\.]+)')

FRAG = {'PESO', 'RELATI', 'VO', 'ELEME', 'NTO', 'ELEMENTO', 'PRE', 'CPC', 'VAE',
        'NP/EP', 'NP/ND', '/ND', '(%)', 'VAE (%)', 'CPC ELEMENTO', 'TOTAL COSTO',
        'DESCRIPCION', 'CANTIDAD', 'TARIFA', 'UNIDAD', 'JORNAL/HR', 'RENDIMIENTO'}

CODITEM = re.compile(r'^\s*[A-Za-z]{0,3}\d{1,5}[A-Za-z]?\s{2,}')


def _pre_uem(pagina):
    out = []
    for ln in pagina.split('\n'):
        s = ln.rstrip()
        if not s.strip():
            continue
        toks = [t.strip() for t in re.split(r'\s{2,}', s.strip()) if t.strip()]
        if toks and all(t.upper() in FRAG for t in toks):
            continue
        # el nombre de la seccion arrastra "- 0 0" de las columnas de VAE, y con
        # cifras dentro deja de reconocerse como cabecera de seccion
        cab = toks[0].upper() if toks else ''
        if cab in SECS:
            out.append(cab)
            continue
        if re.match(r'\s*(PARCIAL|CODIGO|COSTO\s+DIRECTO|COSTO\s+INDIRECTO|'
                    r'PRECIO\s+UNITARIO|%)', s, re.I):
            continue
        # "RUBRO No : 5044      UNIDAD: m" -> dos renglones: parse_pagina empieza
        # a buscar etiquetas DESPUES del renglon del numero
        m = re.search(r'\s{2,}(UNIDAD:.*)$', s)
        if m and COD_RUBRO.search(s):
            out.append(s[:m.start()])
            out.append(m.group(1))
            continue
        out.append(CODITEM.sub('   ', s))
    return '\n'.join(out)


def pie_pdf(pagina, coma_decimal):
    """Costo directo, indirectos y precio total impresos al pie de la pagina."""
    out = {'directo': None, 'indirecto_pct': None, 'precio': None}
    lineas = pagina.split('\n')
    for i, ln in enumerate(lineas):
        s = ln.strip()
        if not s:
            continue
        cifras = [x for x in re.split(r'\s{2,}|\s(?=[\d-])', s) if NUMTOK.match(x.strip())]
        val = _f(cifras[0], coma_decimal) if cifras else None
        if re.search(r'COSTO\s+DIRECTO', s, re.I) and val is not None:
            out['directo'] = val
        elif re.search(r'INDIRECTO', s, re.I) and 'OTROS' not in s.upper():
            # "( 20.0000 % )" o "COSTO INDIRECTO 7.8%": el porcentaje es la cifra
            # que lleva el signo pegado, no el importe que va al final. Y a veces
            # el rotulo ("COSTOS INDIRECTOS %:") queda en un renglon y su valor
            # cae en el siguiente, porque la celda esta centrada verticalmente
            m = re.search(r'([\d.,]+)\s*%', s)
            if not m:
                for j in range(i + 1, min(i + 3, len(lineas))):
                    m = re.search(r'([\d.,]+)\s*%', lineas[j])
                    if m:
                        break
            if m:
                out['indirecto_pct'] = _f(m.group(1), coma_decimal)
        elif re.search(r'PRECIO\s+UNITARIO\s+TOTAL|COSTO\s+TOTAL|VALOR\s+OFERTADO',
                       s, re.I) and val is not None:
            out['precio'] = val
    return out

# Etiquetas de los formatos ya vistos. Ajusta o agrega el tuyo.
ETIQUETAS = {
    # "RUBRO N° 1" arriba, luego "RUBRO:", "DETALLE", "UNIDAD" con el valor
    # separado por espacios
    'rubro_n': {'num': r'RUBRO\s*N[°º]\s*(\d+)',
                'nombre': r'\bRUBRO\s*:',
                'unidad': r'\bUNIDAD\b',
                'detalle': r'\bDETALLE\b',
                'ignorar': r'OFERENTE\b|PROYECTO\b|ANALISIS DE PRECIOS'},
    # "HOJA: 1 DE 137" arriba, "RUBRO:", "UNIDAD:", "ESPECIFICACIÓN:"
    'hoja_n_de': {'num': r'HOJA:\s*(\d+)\s*DE\s*\d+',
                  'nombre': r'\bRUBRO:\s*',
                  'unidad': r'\bUNIDAD:\s*',
                  'detalle': r'\bESPECIFICACI[ÓO]N:\s*',
                  'ignorar': r'RENDIMIENTO:|OFERENTE:|OBJETO\b|PROCESO:|C[ÓO]DIGO:|'
                             r'AN[ÁA]LISIS DE PRECIOS|HOJA:'},
    # "RUBRO No : <codigo>" con UNIDAD en la misma fila y columna CODIGO por item
    'uem_codigo': {'num': r'RUBRO\s*No\.?\s*:?\s*(\d+)',
                   'nombre': r'DESCRIPCI[ÓO]N\s*:',
                   'unidad': r'\bUNIDAD:\s*',
                   'detalle': r'ESPECIFICACI[ÓO]N\s*:',
                   'ignorar': r'UNIDAD EJECUTORA|ANALISIS PRECIO|DETERMINACI[ÓO]N|'
                              r'RECREACI[ÓO]N|CODIGO\b',
                   'pre': _pre_uem, 'coma_decimal': False, 'modo': 'izquierda'},
}

# Ajustes por defecto de cada preset, para que leer_pdf pueda autodetectar
DEFECTOS = {'rubro_n':   {'coma_decimal': False, 'modo': 'derecha'},
            'hoja_n_de': {'coma_decimal': True,  'modo': 'izquierda'},
            'uem_codigo': {'coma_decimal': False, 'modo': 'izquierda'}}


def _una(pagina, et, coma_decimal, modo, i):
    """Lee una pagina ya sabiendo el preset. Devuelve el APU o None."""
    pre = et.get('pre')
    txt = pagina
    cod = ''
    if pre is _pre_uem:
        # el codigo del rubro puede ser alfanumerico (P524) y parse_pagina espera
        # un entero: se sustituye por el numero de pagina y se guarda aparte
        m = COD_RUBRO.search(txt)
        cod = m.group(2) if m else ''
        txt = COD_RUBRO.sub(lambda mm: mm.group(1) + str(i), txt, count=1)
    if pre:
        txt = pre(txt)
    d = parse_pagina(txt, et, coma_decimal, modo)
    if not d['n']:
        return None
    if cod:
        d['n'], d['codigo'] = i, cod
    d.update(pie_pdf(pagina, coma_decimal))
    return d


def calidad(apus):
    """Puntua una lectura. Un preset equivocado deja items sin cantidad, o
    descripciones vacias, o no encuentra ni la mitad de las paginas."""
    items = [x for d in apus.values() for v in d['secs'].values() for x in v]
    return (len(apus) * 3 + len(items)
            + sum(1 for x in items if x['cant'] is not None)
            + sum(1 for x in items if x['desc'])
            + sum(3 for d in apus.values() if d['nombre'])
            + sum(3 for d in apus.values() if d['unidad']))


def detectar_pdf(ruta):
    """(preset, coma_decimal, modo) probando los presets sobre las primeras
    paginas. El decimal se prueba en los dos sentidos porque equivocarlo
    multiplica o divide las cantidades por mil sin dar ninguna senal."""
    pgs = [p for p in paginas(ruta) if p.strip()][:6]
    mejor, punt = None, -1
    for nombre, et in ETIQUETAS.items():
        base = DEFECTOS.get(nombre, {})
        for coma in {base.get('coma_decimal', True), not base.get('coma_decimal', True)}:
            modo = et.get('modo', base.get('modo', 'derecha'))
            try:
                leidas = {}
                for i, p in enumerate(pgs, 1):
                    d = _una(p, et, coma, modo, i)
                    if d:
                        leidas[d['n']] = d
            except Exception:
                continue
            if not leidas:
                continue
            p_ = calidad(leidas)
            # una coma decimal mal elegida inventa cantidades enormes
            cants = [x['cant'] for d in leidas.values() for v in d['secs'].values()
                     for x in v if x['cant'] is not None]
            if cants and max(cants) > 100000:
                p_ -= 50
            if p_ > punt:
                mejor, punt = (nombre, coma, modo), p_
    return mejor


def leer_pdf(ruta, etiquetas=None, coma_decimal=None, modo=None):
    """{n: apu} de un PDF con un rubro por pagina.

    Sin argumentos autodetecta el formato, el decimal y el modo. Pasalos a mano
    solo si la deteccion falla; `references/formatos.md` explica cada uno.
    """
    if etiquetas is None:
        det = detectar_pdf(ruta)
        if not det:
            raise ValueError('no se reconocio el formato del PDF: mira '
                             'references/formatos.md y pasa el preset a mano')
        etiquetas, coma_auto, modo_auto = det
        coma_decimal = coma_auto if coma_decimal is None else coma_decimal
        modo = modo_auto if modo is None else modo
    et = ETIQUETAS[etiquetas] if isinstance(etiquetas, str) else etiquetas
    if coma_decimal is None:
        coma_decimal = et.get('coma_decimal', True)
    if modo is None:
        modo = et.get('modo', 'derecha')

    out = {}
    for i, p in enumerate([x for x in paginas(ruta) if x.strip()], 1):
        d = _una(p, et, coma_decimal, modo, i)
        if d:
            out[d['n']] = d
    return out
