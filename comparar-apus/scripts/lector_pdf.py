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


# La firma electronica se imprime ENCIMA de la tabla, no al final: sus renglones
# caen dentro de las filas y se leen como items o como parte de una descripcion.
# En un caso real, "SOBREACARREO DE ESCOMBROS, TIERRA DE" salio con un NARVAEZ
# en medio, y la ultima pagina de un pliego aporto tres items de TRANSPORTE que
# eran el bloque de firmas.
FIRMA = re.compile(
    r'firmado\s+electr[oó]nicamente|validar\s+[uú]nicamente|firmaec|'
    r'ir\s+a\s+la\s+firma|al\s+menos\s+una\s+firma', re.I)


def sin_firmas(txt):
    """Quita los renglones de firma y el nombre suelto que dejan debajo."""
    fuera, lineas = [], txt.split('\n')
    for i, l in enumerate(lineas):
        if FIRMA.search(l):
            fuera.append(i)
            # el nombre del firmante cae en los renglones de abajo, en
            # mayusculas y sin cifras
            for j in range(i + 1, min(i + 4, len(lineas))):
                s = lineas[j].strip()
                if not s or re.search(r'\d', s):
                    break
                nombre = (s == s.upper() and len(s.split()) <= 5) or \
                    re.match(r'(?i)^(ing|arq|ab|dr|lic|sr|sra|econ)\.', s)
                if not nombre:
                    break
                fuera.append(j)
    return '\n'.join(l for i, l in enumerate(lineas) if i not in set(fuera))


def paginas(pdf):
    txt = subprocess.run(['pdftotext', '-layout', pdf, '-'],
                         capture_output=True, text=True).stdout
    return [sin_firmas(p) for p in txt.split('\f')]


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


def _decimales(token):
    """Cuantos decimales trae IMPRESA una cifra. Es el dato que dice hasta donde
    se puede confiar en ella: Excel imprime 0,0125 como "0,01" si la celda tiene
    formato de dos decimales, y el PDF ya no guarda el resto."""
    m = re.search(r'[.,](\d+)\s*%?$', (token or '').strip())
    return len(m.group(1)) if m else 0


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

    num, ini, fila_num = None, 0, None
    if etiquetas.get('num'):
        for i, ln in enumerate(head):
            m = re.search(etiquetas['num'], ln)
            if m:
                num, fila_num, ini = int(m.group(1)), i, i + 1
                break

    # El valor de una celda vive en su columna: lo que venga despues de un hueco
    # ancho pertenece a otra columna de la misma fila (UNIDAD:, RENDIMIENTO:...)
    corte = lambda s: clean(re.sub(r'\s{2,}.*$', '', s.strip()))
    ignorar = etiquetas.get('ignorar')

    campos = {k: [] for k in ('nombre', 'unidad', 'detalle')}
    pos, sueltos = {}, []
    # La fila del numero de rubro tambien puede llevar una etiqueta a su derecha
    # ("RUBRO No : 0002    UNIDAD: m2", "Rubro: 4    Unidad: m2"): se le miran
    # las etiquetas, pero no se recoge su texto suelto, que es el propio rotulo
    # del numero.
    filas = ([fila_num] if fila_num is not None else []) + list(range(ini, len(head)))
    for i in filas:
        s = head[i].strip()
        # todas las etiquetas de la linea, no solo la primera: cuando DETALLE y
        # UNIDAD comparten renglon, quedarse con una deja a la otra convertida
        # en texto suelto, que acaba pegado al nombre del rubro
        marcas = []
        for k in ('nombre', 'unidad', 'detalle'):
            pat = etiquetas.get(k)
            if not pat or k in pos:
                continue
            m = re.search(pat, s)
            if m:
                marcas.append((m.start(), m.end(), k))
        marcas.sort()
        if not marcas:
            continue
        izq = corte(s[:marcas[0][0]])
        if izq and i != fila_num:
            sueltos.append((i, izq))
        for j, (_, fin, k) in enumerate(marcas):
            hasta = marcas[j + 1][0] if j + 1 < len(marcas) else len(s)
            resto = corte(s[fin:hasta])
            # la etiqueta puede salir vacia arriba y repetirse con el valor mas
            # abajo; nos quedamos con la aparicion que trae contenido
            if resto:
                pos[k] = i
                campos[k] = [(i, resto)]
            elif k not in pos:
                pos[k] = i
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
            cifras = [_f(t, coma_decimal) for t in toks[k:]]
            decs = [_decimales(t) for t in toks[k:]]
            cant = cifras[0] if cifras else None
            dec = decs[0] if decs else 0
            desc = clean(desc)
            if not fila_vacia(desc, cant):
                secs[cur].append({'desc': desc, 'uni': clean(uni), 'cant': cant,
                                  'cifras': cifras, 'dec': dec, 'decs': decs})
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
        cifras = [_f(x, coma_decimal) for x in nums]
        decs = [_decimales(x) for x in nums]
        cant = cifras[0] if cifras else None
        dec = decs[0] if decs else 0
        desc = clean(desc)
        if not fila_vacia(desc, cant):
            secs[cur].append({'desc': desc, 'uni': clean(uni), 'cant': cant,
                              'cifras': cifras, 'dec': dec, 'decs': decs})

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


def _pre_rubro_detalle(pagina):
    """Normaliza una pagina del formato `rubro_detalle`.

    A la derecha del APU hay un bloque de VAE (Peso Relativo, CPC, NP/EP/ND) y,
    bajo el nombre de cada columna, un renglon con la FORMULA ("PE= D / CD",
    "PMO= D / CD"). Ese renglon no trae cifras del APU, asi que el lector lo
    tomaba como texto suelto y se lo pegaba al primer insumo de cada seccion:
    "PE= D / CD CPC X Y(%) VE=PE * Y(%) HERRAMIENTA MANUAL". Eran 148 hallazgos
    de descripcion, todos falsos.
    """
    out = []
    for l in pagina.split('\n'):
        # cabeceras de columna y renglones de formula: nombran las columnas, no
        # son insumos. Se miran sobre la linea entera y ANTES de recortar nada,
        # porque recortar parte la formula ("PM= D / CD" se queda en "PM= D") y
        # entonces ya no se reconoce.
        if re.match(r'\s*(DESCRIPCI[ÓO]N|CUADRILLA TIPO)\b', l) or \
           re.search(r'\bP[EMOT]*\s*=\s*D\s*/\s*CD\b|VE\s*=\s*PE\b', l):
            continue
        # el rotulo de la seccion comparte renglon con la cabecera del bloque de
        # VAE; el resto de las columnas no estorba, porque la cantidad se lee
        # desde la descripcion hacia la derecha y no desde el final
        i = l.find('Peso Relativo')
        out.append(l[:i].rstrip() if i > 0 else l.rstrip())
    return '\n'.join(out)


def _pre_ushay(pagina):
    """Normaliza una pagina del formato USHAY con columnas de VAE.

    Tres arreglos, los tres por como pdftotext ve esta plantilla:

    1. Recorta las columnas de la derecha (PESO RELATIVO, CPC, NP/ND/EP, VAE).
       Traen cifras que no son del APU, y con ellas un renglon que solo lleva
       texto parece una fila de item.
    2. Junta la descripcion dentro de su columna. Excel la estira con espacios
       ("PLANCHA   ACERO   ASTM  A  36"), y al partir por huecos anchos ese 36
       se leia como la cantidad.
    3. Pega a su fila los renglones de texto que envolvieron. Si la fila de
       cifras no trae texto propio, el renglon que la sigue es la cola de su
       descripcion; si lo trae, el renglon que la sigue es su cola SALVO que
       dos lineas mas abajo venga una fila de solo cifras, porque entonces es
       la cabeza de la descripcion del item siguiente.
    """
    out, cut, col, dentro = [], None, None, []
    for l in pagina.split('\n'):
        s = l.strip()
        if 'PESO RELATIVO' in l or re.search(r'C=AxB|PRT ?= ?Td/Q|VAEt', l):
            continue
        if s.startswith('DESCRIPCION'):
            i = l.find('ELEMENTO')
            cut = i if i > 0 else cut
            m = re.search(r'\b(CANTIDAD|UNIDAD|DISTANCIA)\b', l[len('DESCRIPCION'):])
            col = m.start() + len('DESCRIPCION') if m else col
            continue
        if re.match(r'^(EQUIPOS|MANO DE OBRA|MATERIALES|TRANSPORTE)\b', s):
            out.append(l)
            dentro.append(len(out) - 1)
            continue
        l = l[:cut] if cut else l
        if col and len(l) > 3:
            izq, der = re.sub(r'\s+', ' ', l[:col]).strip(), l[col:].strip()
            l = (izq + '   ' + der) if izq and der else (izq or ' ' * 20 + der)
        if col:
            dentro.append(len(out))
        out.append(l)
    return '\n'.join(_pegar_envueltas(out, set(dentro)))


_FUERA = re.compile(r'^(SUBTOTAL|PARCIAL|TOTAL|COSTO|ESTOS PRECIOS|OTROS|EQUIPOS|'
                    r'MANO DE OBRA|MATERIALES|TRANSPORTE)')


def _pegar_envueltas(lineas, dentro):
    """Une a su fila de cifras los renglones de texto que quedaron sueltos.

    Solo dentro de las secciones de items: en la cabecera, "CODIGO DEL RUBRO: 4"
    seguido de "NOMBRE DEL RUBRO: ..." cumple el mismo patron y no se toca.
    """
    parte = lambda l: [t for t in re.split(r'\s{2,}', l.strip()) if t]
    es_num = lambda t: bool(NUMTOK.match(t))
    for i, l in enumerate(lineas[:-1]):
        if i not in dentro or (i + 1) not in dentro:
            continue
        t = parte(l)
        if not t:
            continue
        sig = lineas[i + 1].strip()
        if not sig or _FUERA.match(sig.upper()):
            continue
        t2 = parte(sig)
        if any(es_num(x) for x in t2):
            continue
        if all(es_num(x) for x in t) and len(t) >= 2:
            lineas[i] = ' '.join(t2) + '   ' + l.strip()
        elif any(es_num(x) for x in t) and not es_num(t[0]):
            sig2 = parte(lineas[i + 2]) if i + 2 < len(lineas) else []
            if sig2 and len(sig2) >= 2 and all(es_num(x) for x in sig2):
                continue
            lineas[i] = t[0] + ' ' + ' '.join(t2) + '   ' + '   '.join(t[1:])
        else:
            continue
        lineas[i + 1] = ''
    return lineas


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
            elif re.search(r'INDIRECTOS?\s*%', s, re.I) and cifras:
                # "INDIRECTOS %    20.000    3.72": el signo va pegado al rotulo
                # y no a la cifra, asi que el porcentaje es el primer numero de
                # la fila
                v = _f(cifras[0], coma_decimal)
                if v is not None and 0 < v <= 100:
                    out['indirecto_pct'] = v
            elif out['directo']:
                # Hay plantillas que imprimen el porcentaje sin el signo:
                # "COSTO INDIRECTO    17,00    0,2732". Cual de las dos cifras
                # es el porcentaje no se adivina, se comprueba: la que aplicada
                # al costo directo da la otra.
                vals = [v for v in (_f(x, coma_decimal) for x in cifras)
                        if v is not None]
                for a in vals:
                    if 0 < a <= 100 and any(
                            abs(out['directo'] * a / 100 - b) <= max(0.01, abs(b) * 0.02)
                            for b in vals if b is not a):
                        out['indirecto_pct'] = a
                        break
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
    # "HOJA 4,00 DE 78" arriba, "Rubro:" con el NUMERO (no el nombre), "Unidad:"
    # y "Detalle:", que es donde vive el nombre del rubro
    'rubro_detalle': {'num': r'\bRubro:\s*(\d+)',
                      'nombre': r'\bDetalle:\s*',
                      'unidad': r'\bUnidad:\s*',
                      'detalle': None,
                      'ignorar': r'HOJA\b|CODIGO:|OFERENTE|PROYECTO|'
                                 r'AN[ÁA]LISIS|ANALISIS',
                      'pre': _pre_rubro_detalle,
                      'coma_decimal': True, 'modo': 'izquierda'},
    # "Hoja 4 de 78" arriba, "RUBRO:" con el nombre y "UNIDAD:" en el mismo
    # renglon, "DETALLE:" con "R(H/U):" a su derecha
    'hoja_n_de_rubro': {'num': r'Hoja\s+(\d+)\s+de\s+\d+',
                        'nombre': r'\bRUBRO:\s*',
                        'unidad': r'\bUNIDAD:\s*',
                        'detalle': r'\bDETALLE:\s*',
                        'ignorar': r'ANALISIS DE PRECIOS|R\(H/U\)|CONSTRUCCION|'
                                   r'INFRAESTRUCTURA|LIGA DEPORTIVA|Hoja\s+\d+\s+de',
                        'coma_decimal': False, 'modo': 'izquierda'},
    # USHAY sin numero de rubro impreso: un rubro por pagina, "RUBRO:" con el
    # nombre y "UNIDAD:" en el mismo renglon
    'ushay_rubro': {'num': None, 'num_por_pagina': True,
                    'nombre': r'\bRUBRO:\s*',
                    'unidad': r'\bUNIDAD:\s*',
                    'detalle': r'\bDETALLE:\s*',
                    'ignorar': r'PROYECTO:|OFERENTE:|ANALISIS DE PRECIOS|'
                               r'DETERMINACI|CONSTRUCCION Y MEJORAMIENTO',
                    'pre': _pre_ushay,
                    'coma_decimal': False, 'modo': 'izquierda'},
    # USHAY con las columnas del VAE a la derecha: "CODIGO DEL RUBRO: 4",
    # "NOMBRE DEL RUBRO:", "DETALLE:" impreso encima de "UNIDAD:"
    'ushay_vae': {'num': r'C[ÓO]DIGO DEL RUBRO:\s*(\d+)',
                  'nombre': r'NOMBRE DEL RUBRO:\s*',
                  'unidad': r'\bUNIDAD:\s*',
                  'detalle': r'\bDETALLE:\s*',
                  # Excel imprime la celda DETALLE ENCIMA de la de UNIDAD: el
                  # texto sale entrelazado y no se puede comparar. Se vacia para
                  # que la comparacion lo descarte sola en vez de inventar
                  # diferencias de especificacion que no existen.
                  'detalle_fiable': False,
                  'ignorar': r'NOMBRE DEL PROYECTO|NOMBRE DEL OFERENTE|'
                             r'ANALISIS DE PRECIOS|DETERMINACI|Hoja\s+\d+\s+de',
                  'pre': _pre_ushay,
                  'coma_decimal': True, 'modo': 'izquierda'},
}

# Ajustes por defecto de cada preset, para que leer_pdf pueda autodetectar
DEFECTOS = {'rubro_n':   {'coma_decimal': False, 'modo': 'derecha'},
            'hoja_n_de': {'coma_decimal': True,  'modo': 'izquierda'},
            'uem_codigo': {'coma_decimal': False, 'modo': 'izquierda'},
            'rubro_detalle': {'coma_decimal': True, 'modo': 'izquierda'},
            'hoja_n_de_rubro': {'coma_decimal': False, 'modo': 'izquierda'},
            'ushay_vae': {'coma_decimal': True, 'modo': 'izquierda'},
            'ushay_rubro': {'coma_decimal': False, 'modo': 'izquierda'}}


# Debajo de esta linea ya no hay items: esta el pie del APU y, mas abajo, el
# bloque de firmas. Leer mas alla agrega al rubro cosas que no son insumos.
# \s casa tambien el salto de linea: con \s+ el patron une la cabecera "TOTAL
# COSTO" de una seccion con el "DIRECTO" del pie treinta renglones mas abajo y
# corta la pagina por la mitad. Aqui los separadores son espacios de la misma
# linea, nada mas.
PIE = re.compile(r'(?im)^.*(TOTAL[ \t]+COSTO[ \t]+DIRECTO|COSTO[ \t]+DIRECTO[ \t]*:|'
                 r'COSTO[ \t]+TOTAL[ \t]+DEL[ \t]+RUBRO|PRECIO[ \t]+UNITARIO[ \t]+TOTAL[ \t]*:|'
                 r'VALOR[ \t]+OFERTADO).*$')


def _corta_en_pie(txt):
    m = PIE.search(txt)
    return txt[:m.start()] if m else txt


def _una(pagina, et, coma_decimal, modo, i):
    """Lee una pagina ya sabiendo el preset. Devuelve el APU o None."""
    pre = et.get('pre')
    txt = _corta_en_pie(pagina)
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
    if not d['n'] and et.get('num_por_pagina'):
        # hay plantillas que no imprimen el numero del rubro en ninguna parte:
        # va un rubro por pagina y en el mismo orden del presupuesto, asi que el
        # numero de pagina ES el numero de rubro. Se acepta solo si la pagina
        # trajo nombre, para no numerar caratulas ni anexos.
        if d.get('nombre'):
            d['n'] = i
    if not d['n']:
        return None
    if cod:
        d['n'], d['codigo'] = i, cod
    d.update(pie_pdf(pagina, coma_decimal))
    return d


def decimal_del_pdf(texto):
    """True si el separador decimal es la coma. None si no se puede decidir.

    El truco es mirar solo las cifras que NO pueden ser un separador de miles:
    un grupo de miles tiene exactamente tres digitos, asi que una coma o un
    punto seguidos de CUATRO digitos son, sin discusion, el decimal. Las
    plantillas de APU imprimen las cantidades con cuatro decimales, de modo que
    en la practica siempre hay de donde decidir.

    Adivinarlo por prueba y error, en cambio, no avisa cuando se equivoca: con
    el separador cambiado "0,0940" se lee 940 y el rubro entero sale mal sin
    que nada falle.
    """
    coma = len(re.findall(r'\d,\d{4,}', texto))
    punto = len(re.findall(r'\d\.\d{4,}', texto))
    if coma == punto:
        return None
    return coma > punto


def calidad(apus):
    """Puntua una lectura. Un preset equivocado deja items sin cantidad, o
    descripciones vacias, o no encuentra ni la mitad de las paginas."""
    items = [x for d in apus.values() for v in d['secs'].values() for x in v]
    return (len(apus) * 3 + len(items)
            + sum(1 for x in items if x['cant'] is not None)
            + sum(1 for x in items if x['desc'])
            + sum(3 for d in apus.values() if d['nombre'])
            + sum(3 for d in apus.values() if d['unidad']))


def detectar_pdf(ruta, _solo=None):
    """(preset, coma_decimal, modo) probando los presets sobre las primeras
    paginas. El decimal se prueba en los dos sentidos porque equivocarlo
    multiplica o divide las cantidades por mil sin dar ninguna senal."""
    pgs = [p for p in paginas(ruta) if p.strip()][:6]
    mejor, punt = None, -1
    cierto = decimal_del_pdf('\n'.join(pgs))
    # Numerar por pagina es el ultimo recurso: si algun preset encuentra el
    # numero del rubro IMPRESO, ese gana. Si no, un PDF con dos paginas de
    # presupuesto delante de los APUs se lee corrido, con cada rubro numerado
    # dos puestos mas alla. Por eso se prueban primero los presets que leen el
    # numero, y solo si ninguno sirve se prueban los que cuentan paginas.
    candidatos = ([(k, v) for k, v in (_solo or ETIQUETAS).items()
                   if _solo or not v.get('num_por_pagina')])
    for nombre, et in candidatos:
        base = DEFECTOS.get(nombre, {})
        opciones = ({cierto} if cierto is not None
                    else {base.get('coma_decimal', True), not base.get('coma_decimal', True)})
        for coma in opciones:
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
    if mejor is None and candidatos:
        por_pagina = {k: v for k, v in ETIQUETAS.items() if v.get('num_por_pagina')}
        return detectar_pdf(ruta, _solo=por_pagina) if por_pagina else None
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
            if et.get('detalle_fiable') is False:
                d['detalle'] = ''
            out[d['n']] = d
    _recuperar_cantidades(out)
    return _completar_con_tabla(out, ruta)


def _cantidad_oculta(cant, dec, cifras, decs=None, tol=0.02):
    """Despeja del propio renglon los decimales que la celda no imprimio.

    Excel imprime 0,0125 como "0,01" si la celda tiene formato de dos decimales,
    pero el COSTO de la fila se calcula con el valor completo. Como la fila es
    lineal en la cantidad (cantidad x tarifa x rendimiento = costo, o cantidad x
    precio unitario = costo), el valor que falta se despeja: es el costo
    dividido para el resto de los factores.

    Solo se acepta el valor despejado si, redondeado a los decimales que SI se
    imprimieron, da exactamente la cifra impresa. Asi esto nunca cambia una
    cantidad: solo le devuelve los decimales que el formato de celda escondio.
    Si el oferente escribio 0,01 de verdad, el costo cuadra con 0,01 y no pasa
    nada.
    """
    if cant is None or not cant or dec is None or dec > 2 or len(cifras) < 3:
        return None
    for j in range(2, min(len(cifras), 6)):
        medios = cifras[1:j]
        if any(m is None or m <= 0 for m in medios) or cifras[j] is None:
            return None if j == 2 else None
        # una columna intermedia puede ser un calculo previo de la misma fila
        # ("COSTO HORA" = cantidad x tarifa); multiplicarla contaria dos veces
        util = [m for i, m in enumerate(medios)
                if not (0 < i < len(medios) - 1 and abs(m - cant * medios[i - 1]) <= tol * max(1.0, abs(m)))]
        c = 1.0
        for m in util:
            c *= m
        if c <= 0:
            continue
        total = cifras[j]
        if abs(cant * c - total) <= max(5e-5, abs(total) * 0.02):
            return None                      # la cifra impresa ya cuadra
        q = total / c
        if q <= 0:
            continue
        if round(q, dec) != round(cant, dec) or abs(q - cant) <= 1e-9:
            continue
        # el costo tambien viene redondeado, asi que el despeje sale con ruido
        # (0,01250939). Se devuelve el numero mas corto que sigue reproduciendo
        # el costo impreso dentro de su propia precision: 0,0125.
        dt = (decs[j] if decs and j < len(decs) and decs[j] else 4)
        margen = 0.5 * 10 ** (-dt) + 1e-9
        # Y no se despeja nada si el costo viene demasiado redondeado para
        # sostenerlo. Hay plantillas que imprimen TODA la fila con dos decimales:
        # ahi la aritmetica no distingue 0,01 de 0,009, y aceptar el despeje
        # convertiria una cantidad correcta en una "cantidad menor" inventada.
        # Se exige que la diferencia que se reclama pese al menos cinco veces el
        # redondeo del propio costo.
        if abs(q - cant) * c < 5 * margen:
            continue
        # Y solo se acepta si ese numero es CORTO. Una cantidad de obra se
        # escribe con cuatro decimales como mucho; si hacen falta seis para
        # reproducir el costo, lo que falla es la identificacion de las columnas
        # y no el formato de la celda, y mas vale no tocar nada: inventar
        # 0,09508 donde el oferente puso 0,1 seria el error contrario.
        for k in range(dec + 1, 5):
            qk = round(q, k)
            if qk > 0 and abs(qk * c - total) <= margen:
                return qk
        return None
    return None


def _recuperar_cantidades(apus):
    """Aplica `_cantidad_oculta` a todos los items. Devuelve cuantos recupero."""
    n = 0
    for d in apus.values():
        for filas in d['secs'].values():
            for it in filas:
                q = _cantidad_oculta(it.get('cant'), it.get('dec'),
                                     it.get('cifras') or [], it.get('decs'))
                if q is not None:
                    it['cant_impresa'], it['cant'] = it['cant'], q
                    n += 1
    return n


def cantidades_recuperadas(apus):
    """[(rubro, seccion, descripcion, impresa, recuperada)] de lo que se despejo."""
    out = []
    for n in sorted(apus):
        for sec, filas in apus[n]['secs'].items():
            for it in filas:
                if 'cant_impresa' in it:
                    out.append((n, sec, it['desc'], it['cant_impresa'], it['cant']))
    return out


def _unidad_plausible(u):
    return bool(re.fullmatch(r'[A-Za-z][A-Za-z0-9/.\-]{0,5}', (u or '').strip()))


def _completar_con_tabla(apus, ruta):
    """Rellena la unidad del rubro con la tabla de cantidades del mismo PDF.

    En algunas plantillas la celda DETALLE se imprime ENCIMA de la celda UNIDAD
    y las dos salen entrelazadas: la unidad del rubro es ilegible por mucho que
    se afine el lector de la cabecera. Pero el mismo documento suele traer, en
    sus primeras paginas, la tabla de rubros con la unidad en su propia columna.
    Se toma de ahi en vez de inventarla o de dejarla vacia, que en la
    comparacion saldria como un cambio de unidad que no existe.
    """
    faltan = [n for n, d in apus.items() if not _unidad_plausible(d.get('unidad'))]
    if not faltan:
        return apus
    try:
        import tabla_pdf
        tabla = tabla_pdf.leer(ruta)
    except Exception:
        return apus
    for n in faltan:
        fila = tabla.get(n)
        if fila and _unidad_plausible(fila.get('uni')):
            apus[n]['unidad'] = fila['uni']
    return apus
