"""Regenera el bloque `const DATA = {...}` de index.html a partir de:
  - STOCK MENUDENCIAS DD-MM.xlsx  (stock de planta, una fila por caja)
  - Ventas menudencias.xlsx       (hojas 12-Saldos por ventas, Puertos, 11-Shipments)

La configuracion fija (lista de codigos, peso por res, share, pools, consumo,
puntual, faena proyectada y semana de cada venta) se toma del DATA que ya esta
publicado en index.html, asi que se conserva de un dia al otro.

Uso:
  python actualizar.py STOCK.xlsx VENTAS.xlsx [--corte AAAA-MM-DD] [--dry-run]
"""
import argparse, collections, datetime as dt, json, re, sys
from pathlib import Path
import openpyxl

HTML = Path(__file__).with_name('index.html')
RX = re.compile(r'const DATA = (\{.*?\});\n')


def leer_data_actual(html):
    return json.loads(RX.search(html).group(1))


def fecha(v):
    if v is None or v == '':
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, (int, float)):  # 20260928
        s = str(int(v)); return dt.date(int(s[:4]), int(s[4:6]), int(s[6:]))
    return dt.datetime.strptime(str(v).strip(), '%d/%m/%Y').date()


def rows(ws):
    return list(ws.iter_rows(values_only=True))


# ---------------------------------------------------------------- stock
def leer_stock(path):
    ws = openpyxl.load_workbook(path, read_only=True, data_only=True).active
    r = rows(ws); h = r[0]
    return [dict(zip(h, x)) for x in r[1:] if x[h.index('CODIGO')]]


def armar_stock(prev, cajas, corte):
    por_cod = collections.defaultdict(list)
    for c in cajas:
        por_cod[c['CODIGO']].append(c)
    codigos = []
    for base in prev['codigos']:
        c = dict(base)
        filas = por_cod.get(c['c'], [])
        cam = collections.defaultdict(float)
        for f in filas:
            cam[f['camara']] += f['kg'] or 0
        c['stock'] = round(sum(f['kg'] or 0 for f in filas), 2)
        c['cajas'] = int(sum(f['CAJAS'] or 0 for f in filas))
        c['ret'] = round(sum(f['kg'] or 0 for f in filas if f['Retenido'] != 'APTO'), 2)
        c['cam'] = {k: round(v, 2) for k, v in sorted(cam.items()) if round(v, 2)}
        codigos.append(c)
    # partidas: por codigo + fecha produccion + vencimiento
    g = collections.defaultdict(lambda: [0.0, 0])
    for c in cajas:
        if c['CODIGO'] not in {x['c'] for x in prev['codigos']}:
            continue
        k = (c['CODIGO'], fecha(c['FECHA PROD']), fecha(c['Fec. VENC']))
        g[k][0] += c['kg'] or 0; g[k][1] += c['CAJAS'] or 0
    partidas = [dict(c=k[0], fp=k[1].isoformat(), fv=k[2].isoformat(), kg=round(v[0], 2), cj=int(v[1]),
                     ant=(corte - k[1]).days, dv=(k[2] - corte).days)
                for k, v in sorted(g.items())]
    sin_config = sorted(set(por_cod) - {x['c'] for x in prev['codigos']})
    return codigos, partidas, sin_config


# ---------------------------------------------------------------- ventas
def leer_saldos(ws):
    out = []; eng = gen = dest = code = None
    for r in rows(ws)[5:]:
        if r[0] is not None: eng = r[0]
        if r[1] is not None: gen = r[1]
        if r[2] is not None: dest = r[2]
        if r[3] is not None and not str(r[3]).startswith('Total'): code = r[3]
        if r[4] is None: continue
        out.append(dict(code=code, venta=str(r[4]), vend=r[5] or 0, carg=r[6] or 0,
                        vtarto=r[7] or 0, pend=r[8] or 0))
    return out


def leer_puertos(ws):
    out = {}
    for r in rows(ws)[5:]:
        if r[6] is None: continue
        out[str(r[6])] = dict(cliente=r[2], puerto=r[7] or r[3], pais=r[4])
    return out


def leer_ships(ws):
    out = []; cons = ven = puerto = nro = sem = None
    for r in rows(ws)[3:]:
        if r[0]: cons = r[0]
        if r[1]: ven = str(r[1]); nro = None
        if r[2]: puerto = r[2]
        if r[3]: nro = str(r[3])
        if r[7]: sem = r[7]
        if not r[4]: continue
        out.append(dict(venta=ven, nro=nro, consignee=cons, puerto=puerto, semana=r[7] or sem, c=r[4], kg=r[8] or 0))
    return out


def armar_ventas(prev, saldos, puertos, ships_rows):
    cods = {c['c']: c for c in prev['codigos'] if c['c'] and not c.get('virtual')}
    prev_v = {v['nro']: v for v in prev['ventas']}
    por_v = collections.defaultdict(list)
    for s in saldos:
        if s['code'] in cods:
            por_v[s['venta']].append(s)
    # ventas viejas con saldos residuales no se siguen: solo entran las ya publicadas
    # y las nuevas con numero >= a la venta mas baja publicada
    piso = min(int(n) for n in prev_v if n.isdigit())
    ventas, faltan_semana = [], []
    for nro, its in sorted(por_v.items(), key=lambda x: int(x[0]) if x[0].isdigit() else 0):
        if nro not in prev_v and not (nro.isdigit() and int(nro) >= piso):
            continue
        items = {s['code']: max(0.0, s['pend']) for s in its if s['pend'] > 0.005}
        # compensacion: lo cargado de mas en un codigo cubre lo pendiente de otro del mismo generico
        sobra = {s['code']: -s['vtarto'] for s in its if s['vtarto'] < -0.005}
        comp = []
        for c in list(items):
            for o in list(sobra):
                if o != c and cods[o]['g'] == cods[c]['g'] and sobra[o] > 0:
                    kg = min(items[c], sobra[o]); sobra[o] -= kg; items[c] -= kg
                    comp.append(dict(c=c, kg=round(kg, 2), por=o))
            if items[c] <= 0.005: del items[c]
        if not items:
            continue
        p = puertos.get(nro, {}); pv = prev_v.get(nro, {})
        cli = p.get('cliente') or pv.get('cliente')
        if cli in (None, '(en blanco)'): cli = 'A CONFIRMAR'
        v = dict(nro=nro, cliente=cli, puerto=p.get('puerto') or pv.get('puerto'),
                 week=pv.get('week'), madre=pv.get('madre'),
                 items={k: round(x, 2) for k, x in items.items()},
                 contratado={s['code']: round(s['vend'], 2) for s in its if s['vend']},
                 cargado={s['code']: round(s['carg'], 2) for s in its if s['carg']})
        if comp: v['comp'] = comp
        v['wsrc'] = pv.get('wsrc', 'reporte'); v['pais'] = p.get('pais') or pv.get('pais')
        if not v['week']: faltan_semana.append(nro)
        ventas.append(v)
    abiertas = {v['nro'] for v in ventas}
    sh = collections.OrderedDict()
    for r in ships_rows:
        if r['venta'] in abiertas and r['c'] in cods and r['nro']:
            k = (r['venta'], r['nro'])
            if k not in sh:
                sh[k] = dict(venta=r['venta'], nro=r['nro'], semana=r['semana'], consignee=r['consignee'],
                             puerto=r['puerto'], items={})
            sh[k]['items'][r['c']] = sh[k]['items'].get(r['c'], 0) + float(r['kg'])
    return ventas, list(sh.values()), faltan_semana


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('stock'); ap.add_argument('ventas')
    ap.add_argument('--corte', help='fecha de corte AAAA-MM-DD (por defecto, la del nombre del archivo de stock)')
    ap.add_argument('--semanas', help='JSON con semana por venta para ventas nuevas, ej {"10380":"W50"}')
    ap.add_argument('--dry-run', action='store_true')
    ap.add_argument('--out-json')
    a = ap.parse_args()

    html = HTML.read_text(encoding='utf-8')
    prev = leer_data_actual(html)
    if a.corte:
        corte = dt.date.fromisoformat(a.corte)
    else:
        m = re.search(r'(\d{1,2})-(\d{1,2})', Path(a.stock).name)
        corte = dt.date(dt.date.today().year, int(m.group(2)), int(m.group(1)))

    codigos, partidas, sin_config = armar_stock(prev, leer_stock(a.stock), corte)
    wb = openpyxl.load_workbook(a.ventas, read_only=True, data_only=True)
    ventas, ships, faltan = armar_ventas(prev, leer_saldos(wb['12-Saldos por ventas']),
                                         leer_puertos(wb['Puertos']), leer_ships(wb['11-Shipments']))
    if a.semanas:
        sem = json.loads(a.semanas)
        for v in ventas:
            if v['nro'] in sem: v['week'] = sem[v['nro']]
        faltan = [n for n in faltan if n not in sem]

    sc = corte.isocalendar()[1]
    semanas = [s for s in prev['semanas'] if s['w'] >= sc] or prev['semanas']
    data = dict(corte=corte.isoformat(), semanaCorte=sc, semanas=semanas, codigos=codigos,
                ventas=ventas, ships=ships, partidas=partidas)

    print(f"corte {data['corte']} W{sc} | {sum(c['stock'] for c in codigos)/1000:.1f} tn, "
          f"{sum(c['cajas'] for c in codigos)} cajas | {len(ventas)} ventas, {len(ships)} shipments, {len(partidas)} partidas")
    if sin_config: print('Codigos en stock sin configurar (no entran):', ', '.join(sin_config))
    if faltan: print('VENTAS NUEVAS SIN SEMANA (pasar --semanas):', ', '.join(faltan))
    if a.out_json:
        Path(a.out_json).write_text(json.dumps(data, ensure_ascii=False, indent=1), encoding='utf-8')
    if a.dry_run:
        return
    nuevo = RX.sub(lambda m: 'const DATA = ' + json.dumps(data, ensure_ascii=False, separators=(',', ':')) + ';\n', html, count=1)
    HTML.write_text(nuevo, encoding='utf-8', newline='\n')
    print('index.html actualizado')


if __name__ == '__main__':
    main()
