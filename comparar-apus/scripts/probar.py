"""Prueba de humo de la skill: genera el reporte y revisa lo que no se puede
comprobar leyendo el codigo.

Uso:
    python3 probar.py <referencia> <oferta> [salida.xlsx]

Los dos archivos pueden ser .xlsx o .pdf, en cualquier combinacion.

Imprime, en este orden:
  1. formato detectado y total de items de cada archivo
  2. si la correspondencia hoja a hoja cuadra por nombre de rubro
  3. que hoja de presupuesto encontro en cada libro y cuantos items leyo
  4. si el numero de item del presupuesto corresponde al numero de hoja del APU
  5. el precio unitario del APU frente al del presupuesto, en una muestra
  6. el cuadro de estados

Los puntos 3, 4 y 5 son los que hay que mirar con lupa la primera vez que se usa
un archivo nuevo: si el presupuesto se enlaza mal, las columnas de cantidad y
valor del RESUMEN comparan rubros distintos y no avisan de nada.
"""
import re
import sys

import apu
import lectores
import lector_pdf
import presupuesto


def hojas_de(ruta):
    """{n: nombre_de_hoja} de las hojas que son APUs, sea 'N' o 'RubroN'."""
    import openpyxl
    wb = openpyxl.load_workbook(ruta, read_only=True)
    num, rub = {}, {}
    for h in wb.sheetnames:
        t = h.strip()
        if t.isdigit():
            num[int(t)] = h
        m = re.fullmatch(r'(?i)rubro\s*(\d+)', t)
        if m:
            rub[int(m.group(1))] = h
    # si hay los dos juegos, gana el mas numeroso; si empatan, el numerado
    return num if len(num) >= len(rub) else rub


def leer(ruta):
    """(etiqueta_de_formato, {n: apu}) de un xlsx o de un pdf."""
    if re.search(r'\.pdf$', ruta, re.I):
        det = lector_pdf.detectar_pdf(ruta)
        return 'PDF %s' % (det[0] if det else '?'), lector_pdf.leer_pdf(ruta)
    hojas = hojas_de(ruta)
    f, d = lectores.leer_libro(ruta, hojas)
    return 'Excel %s' % f, d


def main(ref_path, ofe_path, salida=None):
    salida = salida or 'Comparacion APUs.xlsx'

    fr, ref = leer(ref_path)
    fo, ofe = leer(ofe_path)
    comunes = sorted(set(ref) & set(ofe))
    ir = sum(len(v) for d in ref.values() for v in d['secs'].values())
    io = sum(len(v) for d in ofe.values() for v in d['secs'].values())
    print('formato -> referencia: %s, oferta: %s' % (fr, fo))
    print('rubros -> referencia: %d, oferta: %d, comunes: %d | items: %d vs %d'
          % (len(ref), len(ofe), len(comunes), ir, io))
    if len(ref) != len(ofe):
        print('  !! distinta cantidad de rubros: revisa si hay dos juegos de APUs')
    if ir and abs(ir - io) > max(10, 0.05 * ir):
        print('  !! los totales se alejan mucho: sospecha del lector, no de la oferta')

    mal = [n for n in comunes if apu.nd(ref[n]['nombre']) != apu.nd(ofe[n]['nombre'])]
    print('correspondencia por nombre -> desalineados: %d' % len(mal))
    for n in mal[:5]:
        print('   %s | %s || %s' % (n, ref[n]['nombre'][:44], ofe[n]['nombre'][:44]))

    pref, hp_ref = presupuesto.leer(ref_path)
    pofe, hp_ofe = presupuesto.leer(ofe_path)
    print('presupuesto -> referencia: %r (%d items) | oferta: %r (%d items)'
          % (hp_ref, len(pref), hp_ofe, len(pofe)))

    # el enlace del presupuesto se valida contra el nombre del rubro del APU
    for etiq, pres in (('referencia', pref), ('oferta', pofe)):
        if not pres:
            continue
        coinc = sum(1 for n in comunes
                    if n in pres and apu.nd(pres[n]['desc']) == apu.nd(ref[n]['nombre']))
        print('   enlace %s: %d de %d items casan por descripcion'
              % (etiq, coinc, len(comunes)))
        if coinc < 0.8 * len(comunes):
            print('     !! el numero de item NO corresponde al numero de hoja;'
                  ' hay que enlazar por codigo de rubro antes de usar estas columnas')

    print('precio unitario (APU vs presupuesto), primeros 8 rubros:')
    for n in comunes[:8]:
        va = apu.num(ofe[n].get('precio'))
        vp = (pofe.get(n) or {}).get('punit')
        ok = ('' if va is None or vp is None
              else (' OK' if abs(va - vp) < 0.005 else '  <-- NO CUADRA'))
        print('   %-4s APU=%-10s presupuesto=%-10s%s' % (n, va, vp, ok))

    print('indirectos -> referencia: %s | oferta: %s'
          % (apu.margen(ref[comunes[0]]), apu.margen(ofe[comunes[0]])))

    # con la oferta impresa en PDF se activa TEXTO CORTADO: sin eso, la
    # especificacion que la celda impresa no alcanzo a mostrar se reporta como
    # si el oferente la hubiera cambiado
    es_pdf = bool(re.search(r'\.pdf$', ofe_path, re.I))
    filas, res = apu.reporte(ref, ofe, salida, pres_ref=pref, pres_ofe=pofe,
                             pdf=es_pdf,
                             titulo_ref=ref_path, titulo_ofe=ofe_path,
                             correspondencia='Rubro a rubro, verificado por nombre '
                                             'en los %d rubros.' % len(comunes))
    r = apu.resumir(filas, res)
    print('estados ->', dict(r['estados']))
    print('graves ->', r['errores'], '| por campo:', dict(r['campos']))
    print('escrito:', salida)


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(1)
    main(*sys.argv[1:4])
