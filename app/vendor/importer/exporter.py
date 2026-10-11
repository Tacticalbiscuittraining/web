#!/usr/bin/env python3
"""Rellena las plantillas de Robert (tools/plantillas) con un plan de la app.
Es lo contrario de importer.py: lo que sale se puede volver a importar.

Uso:  python3 tools/exporter.py datos.json -o salida.docx|.xlsx
datos.json = {"kind": "diet"|"training", "name", "goal", "review", "weight", "data": {"diet"|"training": …}}
En la app se ejecuta en el navegador (Pyodide), igual que el importador.
"""
import copy, datetime as dt, io, json, os, re, sys, zipfile
from lxml import etree

HERE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(HERE, 'plantillas')
DAYS = ['LUNES', 'MARTES', 'MIÉRCOLES', 'JUEVES', 'VIERNES', 'SÁBADO', 'DOMINGO']

def fnum(x, dec=1):
    """68.4 → '68,4' · 150.0 → '150'"""
    if x in (None, ''): return ''
    try: x = float(x)
    except (TypeError, ValueError): return str(x)
    s = ('%.' + str(dec) + 'f') % x
    s = s.rstrip('0').rstrip('.') if '.' in s else s
    return s.replace('.', ',')

def thousands(x):
    return format(int(round(float(x))), ',').replace(',', '.') if x not in (None, '') else ''

# ================================================================ Word
W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
def q(t): return '{%s}%s' % (W, t)

def cell_rpr(tc, row):
    """Formato para el texto de una celda: el de su primer texto, o el de la etiqueta de la fila sin negrita."""
    r = tc.find('.//' + q('r'))
    if r is not None and r.find(q('rPr')) is not None: return copy.deepcopy(r.find(q('rPr')))
    for r in row.iter(q('r')):
        rp = r.find(q('rPr'))
        if rp is None: continue
        rp = copy.deepcopy(rp)
        for t in ('b', 'bCs', 'color'):
            for e in rp.findall(q(t)): rp.remove(e)
        c = etree.SubElement(rp, q('color')); c.set(q('val'), '1F2D33')
        return rp
    return None

def set_cell(tc, text, row=None):
    """Pone el texto en la celda conservando su formato (una línea = un párrafo)."""
    rpr = cell_rpr(tc, row if row is not None else tc.getparent())
    ps = tc.findall(q('p'))
    p0 = ps[0]
    for p in ps[1:]: tc.remove(p)
    for ch in list(p0):
        if ch.tag != q('pPr'): p0.remove(ch)
    lines = str(text).split('\n') if text not in (None, '') else []
    for i, ln in enumerate(lines or ['']):
        p = p0 if i == 0 else copy.deepcopy(p0)
        if i:
            for ch in list(p):
                if ch.tag != q('pPr'): p.remove(ch)
            tc.append(p)
        if not ln: continue
        r = etree.SubElement(p, q('r'))
        if rpr is not None: r.append(copy.deepcopy(rpr))
        t = etree.SubElement(r, q('t')); t.text = ln
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

def replace_blank(tc, value):
    """'________ kg' → '68,4 kg' (solo cambia el hueco)."""
    if value in (None, ''): return
    for t in tc.iter(q('t')):
        if t.text and '____' in t.text:
            t.text = re.sub(r'_{4,}', value, t.text, count=1)
            rp = t.getparent().find(q('rPr'))
            if rp is not None:
                for e in rp.findall(q('color')): e.set(q('val'), '1F2D33')
            return

def rows(tbl): return tbl.findall(q('tr'))
def cells(tr): return tr.findall(q('tc'))
def ctext(tc): return '\n'.join(''.join(x.text or '' for x in p.iter(q('t'))) for p in tc.iter(q('p'))).strip()

def add_row(tbl, model, before=None):
    nr = copy.deepcopy(model)
    for tc in cells(nr): set_cell(tc, '', nr)
    if before is not None: before.addprevious(nr)
    else: tbl.append(nr)
    return nr

MEALS = ['Desayuno', 'Media mañana', 'Comida', 'Pre-entreno', 'Merienda', 'Post-entreno', 'Cena']
def norm_meal(s):
    s = (s or '').strip().lower()
    for m in MEALS:
        if s.startswith(m.lower()): return m
    return None

def mtot(m):
    P = sum(float(i.get('p') or 0) for i in m.get('items', []))
    C = sum(float(i.get('c') or 0) for i in m.get('items', []))
    F = sum(float(i.get('f') or 0) for i in m.get('items', []))
    return P, C, F, 4 * P + 4 * C + 9 * F

def g_of(i):
    g = i.get('g')
    return fnum(g, 0) if g not in (None, '') else ''

def meal_compact(m):
    its = [('%s g %s' % (g_of(i), i.get('food', ''))).strip() for i in m.get('items', []) if i.get('food')]
    return ' + '.join(its)

def meal_menu(m):
    """Celda del menú semanal: plato y alimentos (como en tus planes: 1.ª línea el plato, luego cantidades)."""
    its = [('%s g %s' % (g_of(i), i.get('food', ''))).strip() for i in m.get('items', []) if i.get('food')]
    dish = (m.get('dish') or '').strip()
    if dish: return '\n'.join([dish] + its)
    return ' + '.join(its) or '—'

def fill_plan(J):
    import docx
    doc = docx.Document(os.path.join(TPL, 'PLAN_NUTRICIONAL.docx'))
    d = J['data'].get('diet') or {}
    T = d.get('types') or {}
    body = doc.element.body
    tbls = list(body.iter(q('tbl')))
    head = lambda t: ctext(cells(rows(t)[0])[0]).upper() if rows(t) and cells(rows(t)[0]) else ''
    trainDays = d.get('trainDays') or [0] * 7
    nT = sum(1 for x in trainDays if x)
    wt = J.get('weight') or 0
    def tot(key):
        ms = (T.get(key) or {}).get('meals') or []
        S = [0, 0, 0, 0]
        for m in ms:
            for j, v in enumerate(mtot(m)): S[j] += v
        return S
    tt, tr = tot('train'), tot('rest')
    kt = (T.get('train') or {}).get('kcal') or tt[3]
    kr = (T.get('rest') or {}).get('kcal') or tr[3]
    avg = (kt * nT + kr * (7 - nT)) / 7 if (kt or kr) else 0
    ini = d.get('inicio') or ''
    m = re.match(r'(\d{4})-(\d\d)-(\d\d)', ini)
    if m: ini = '%s/%s/%s' % (m.group(3), m.group(2), m.group(1))
    kv = {'Nombre completo': J.get('name', ''), 'Objetivo del plan': J.get('goal', ''), 'Fecha de inicio': ini,
          'Duración': d.get('dur', ''), 'Próxima revisión': J.get('review', '') if J.get('review') not in ('—', '-') else '',
          'Edad': fnum(d.get('edad')), 'Altura (cm)': fnum(d.get('altura')), 'Días entreno / sem.': d.get('diasTxt') or (str(nT) if nT else ''),
          'Nivel de actividad': d.get('nivel', '')}
    menu_done = False
    for t in tbls:
        R = rows(t); h = head(t)
        if not R: continue
        # tablas «etiqueta | valor»
        if len(R) == 1 and h in ('NOMBRE COMPLETO', 'OBJETIVO DEL PLAN', 'FECHA DE INICIO', 'EDAD', 'NIVEL DE ACTIVIDAD'):
            cs = cells(R[0])
            for i in range(0, len(cs) - 1, 2):
                lab = ctext(cs[i])
                if lab in kv and kv[lab] not in (None, ''): set_cell(cs[i + 1], kv[lab], R[0])
        elif h == 'PESO ACTUAL':
            cs = cells(R[1])
            for c, v in zip(cs, [fnum(wt), fnum(d.get('goalW')), thousands(d.get('bmr')) if d.get('bmr') else '', thousands(d.get('tdee')) if d.get('tdee') else '']):
                replace_blank(c, v)
        elif h == 'KCAL DÍA ENTRENO':
            cs = cells(R[1]); bal = round(avg - float(d['tdee'])) if d.get('tdee') and avg else None
            vals = [thousands(kt) if kt else '', thousands(kr) if kr else '',
                    ('' if bal is None else ('+' if bal > 0 else '−' if bal < 0 else '') + thousands(abs(bal))), thousands(avg) if avg else '']
            for c, v in zip(cs, vals): replace_blank(c, v)
        elif h == 'MACRONUTRIENTE':
            two = lambda a, b: '%s / %s' % (a, b)
            for r in R[1:]:
                cs = cells(r); lab = ctext(cs[0]).lower()
                if lab.startswith('prote'): j, f = 0, 4
                elif lab.startswith('hidratos'): j, f = 1, 4
                elif lab.startswith('grasas'): j, f = 2, 9
                elif lab.startswith('fibra'): set_cell(cs[1], d.get('fibra') or '25 – 30', r); continue
                elif lab.startswith('total'):
                    if tt[3] or tr[3]: set_cell(cs[3], two(round(tt[3]), round(tr[3])), r)
                    continue
                else: continue
                a, b = tt[j], tr[j]
                if not (a or b): continue
                set_cell(cs[1], two(fnum(a, 0), fnum(b, 0)), r)
                if wt: set_cell(cs[2], two(fnum(a / wt), fnum(b / wt)), r)
                set_cell(cs[3], two(round(a * f), round(b * f)), r)
                set_cell(cs[4], two(('%d %%' % round(a * f / tt[3] * 100)) if tt[3] else '', ('%d %%' % round(b * f / tr[3] * 100)) if tr[3] else ''), r)
        elif 'AGUA' in h:
            replace_blank(cells(R[0])[1], d.get('waterTxt') or fnum(d.get('water')))
        elif h == 'SUPLEMENTO':
            sp = [s for s in (d.get('supps') or []) if s.get('name')]
            while len(R) - 1 < len(sp): add_row(t, R[-1]); R = rows(t)
            for r, s in zip(R[1:], sp):
                for c, v in zip(cells(r), [s.get('name'), s.get('dose'), s.get('when'), s.get('obs')]): set_cell(c, v or '', r)
        elif h == 'TOMA':
            key = 'train' if not getattr(fill_plan, '_t', None) else 'rest'
            fill_plan._t = True
            ms = (T.get(key) or {}).get('meals') or []
            byrow = {norm_meal(ctext(cells(r)[0])): r for r in R[1:] if not ctext(cells(r)[0]).upper().startswith('TOTAL')}
            total_row = next((r for r in R if ctext(cells(r)[0]).upper().startswith('TOTAL')), None)
            S = [0, 0, 0, 0]
            for mm in ms:
                r = byrow.get(norm_meal(mm.get('name')))
                if r is None or ctext(cells(r)[2]):
                    r = add_row(t, R[1], before=total_row); set_cell(cells(r)[0], mm.get('name', ''), r)
                P, C, F, K = mtot(mm)
                for j, v in enumerate((P, C, F, K)): S[j] += v
                cs = cells(r)
                set_cell(cs[1], mm.get('time') or '', r)
                set_cell(cs[2], meal_compact(mm) or (mm.get('dish') or ''), r)
                if mm.get('items'):
                    set_cell(cs[3], fnum(P), r); set_cell(cs[4], fnum(C), r); set_cell(cs[5], fnum(F), r); set_cell(cs[6], str(round(K)), r)
            if total_row is not None and ms:
                cs = cells(total_row)
                for j, v in enumerate(S[:3]): set_cell(cs[3 + j], fnum(v), total_row)
                set_cell(cs[6], str(round(S[3])), total_row)
        elif h in ('DÍA', 'DIA') and len(cells(R[0])) > 2 and ctext(cells(R[0])[1]).upper().startswith('DESAYUNO') and not menu_done:
            menu_done = True
            cols = [norm_meal(ctext(c)) for c in cells(R[0])[1:]]
            for i, r in enumerate(R[1:8]):
                cs = cells(r)
                key = 'd%d' % i if ('d%d' % i) in T else ('train' if trainDays[i] else 'rest')
                if trainDays[i]:
                    for x in cs[0].iter(q('t')):
                        if x.text == '☐': x.text = '☒'
                    for x in cs[0].iter('{http://schemas.microsoft.com/office/word/2010/wordml}checked'):
                        x.set('{http://schemas.microsoft.com/office/word/2010/wordml}val', '1')
                ms = (T.get(key) or {}).get('meals') or []
                txt = {}
                last = None
                for mm in ms:
                    k = norm_meal(mm.get('name'))
                    if k in cols: txt[k] = meal_menu(mm); last = k
                    elif last: txt[last] += '\n' + (mm.get('name') or '') + ': ' + meal_menu(mm)
                for k, c in zip(cols, cs[1:]):
                    set_cell(c, txt.get(k, '—'), r)
        elif h.startswith('GRUPO') and d.get('exch'):
            for r, g in zip(R[1:], d['exch']):
                cs = cells(r)
                set_cell(cs[0], g[0], r); set_cell(cs[1], '\n'.join('•  ' + x for x in g[2]), r); set_cell(cs[2], g[1], r); set_cell(cs[3], g[3], r)
        elif h.startswith('⚠') and d.get('avoid'):
            set_cell(cells(R[0])[1], d['avoid'], R[0])
        elif h == 'PAUTA' and d.get('pautas'):
            P = d['pautas']
            while len(R) - 1 < len(P): add_row(t, R[-1]); R = rows(t)
            for r, p in zip(R[1:], P):
                set_cell(cells(r)[0], p[0], r); set_cell(cells(r)[1], p[1], r)
        elif h.startswith('OBSERVACIONES'):
            L = [x for x in (d.get('notes') or '').split('\n') if x.strip()]
            while len(R) - 1 < len(L): add_row(t, R[-1]); R = rows(t)
            for r, ln in zip(R[1:], L): set_cell(cells(r)[0], ln, r)
    fill_plan._t = None
    out = io.BytesIO(); doc.save(out); return out.getvalue()

# ================================================================ Excel (se toca solo el XML de las celdas: el diseño queda intacto)
S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
def qs(t): return '{%s}%s' % (S, t)
def col_n(c):
    n = 0
    for ch in c: n = n * 26 + ord(ch) - 64
    return n
def split_ref(ref):
    m = re.match(r'([A-Z]+)(\d+)$', ref); return m.group(1), int(m.group(2))

class Sheet:
    def __init__(self, xml):
        self.root = etree.fromstring(xml)
        self.data = self.root.find(qs('sheetData'))
    def cell(self, ref):
        col, rn = split_ref(ref)
        row = None
        for r in self.data.findall(qs('row')):
            n = int(r.get('r'))
            if n == rn: row = r; break
            if n > rn:
                row = etree.Element(qs('row')); row.set('r', str(rn)); r.addprevious(row); break
        if row is None:
            row = etree.SubElement(self.data, qs('row')); row.set('r', str(rn))
        for c in row.findall(qs('c')):
            cr = c.get('r'); cc, _ = split_ref(cr)
            if cr == ref: return c
            if col_n(cc) > col_n(col):
                nc = etree.Element(qs('c')); nc.set('r', ref); c.addprevious(nc); return nc
        nc = etree.SubElement(row, qs('c')); nc.set('r', ref); return nc
    def set(self, ref, v):
        c = self.cell(ref)
        for ch in list(c): c.remove(ch)
        for a in ('t',):
            if a in c.attrib: del c.attrib[a]
        if v is None or v == '': return
        if isinstance(v, (dt.date, dt.datetime)):
            v = (dt.datetime(v.year, v.month, v.day) - dt.datetime(1899, 12, 30)).days
        if isinstance(v, bool): v = int(v)
        if isinstance(v, (int, float)):
            etree.SubElement(c, qs('v')).text = repr(v) if isinstance(v, float) else str(v)
        else:
            c.set('t', 'inlineStr')
            is_ = etree.SubElement(c, qs('is')); t = etree.SubElement(is_, qs('t')); t.text = str(v)
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    def bytes(self):
        return etree.tostring(self.root, xml_declaration=True, encoding='UTF-8', standalone=True)

def numish(v):
    """'8' → 8 · '60' → 60 · '8-10' → '8-10'"""
    s = str(v if v is not None else '').strip()
    if re.fullmatch(r'\d+', s): return int(s)
    if re.fullmatch(r'\d+[.,]\d+', s): return float(s.replace(',', '.'))
    return s

WD = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo']
def fill_training(J):
    T = J['data'].get('training') or {}
    src = os.path.join(TPL, 'ENTRENAMIENTO_CLIENTE.xlsx')
    zin = zipfile.ZipFile(src)
    wb = etree.fromstring(zin.read('xl/workbook.xml'))
    rels = etree.fromstring(zin.read('xl/_rels/workbook.xml.rels'))
    rid = {r.get('Id'): 'xl/' + r.get('Target').lstrip('/').replace('xl/', '', 1) for r in rels}
    R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    path = {s.get('name'): rid[s.get(R)] for s in wb.find(qs('sheets'))}
    sheets = {}
    def sh(name):
        if name not in sheets: sheets[name] = Sheet(zin.read(path[name]))
        return sheets[name]
    meso = T.get('meso') or {}
    days = [d for d in (T.get('days') or []) if d.get('exs')]
    res = sh('Resumen')
    res.set('C6', J.get('name', '')); res.set('C7', J.get('goal', ''))
    res.set('C8', meso.get('level') or '')
    if re.match(r'\d{4}-\d\d-\d\d$', meso.get('start') or ''): res.set('C9', dt.date.fromisoformat(meso['start']))
    res.set('C10', len(days) or '')
    res.set('C11', int(meso.get('n') or 1))
    for i, w in enumerate((meso.get('weeks') or [])[:4]):
        r = 15 + i
        res.set('D%d' % r, w.get('phase') or ''); res.set('E%d' % r, numish(w.get('rir')))
        res.set('F%d' % r, w.get('vol') or ''); res.set('G%d' % r, w.get('int') or ''); res.set('H%d' % r, w.get('notes') or '')
    day_sheets = [s.get('name') for s in wb.find(qs('sheets')) if s.get('name').startswith('Día')]
    left = []
    for di, d in enumerate(days):
        if di >= len(day_sheets): left.append(d.get('name', '')); continue
        ws = sh(day_sheets[di]); rr = 22 + di
        wd = d.get('wd')
        res.set('C%d' % rr, WD[wd] if isinstance(wd, int) and 0 <= wd < 7 else '')
        res.set('D%d' % rr, d.get('focus') or '')
        res.set('G%d' % rr, numish(d.get('dur') or ''))
        res.set('H%d' % rr, d.get('place') or '')
        exs = d['exs'][:8]
        if len(d['exs']) > 8: left.append('%s: %s' % (d.get('name', ''), ', '.join(e.get('name', '') for e in d['exs'][8:])))
        for i, e in enumerate(exs):
            r = 10 + i
            ws.set('D%d' % r, e.get('name', '')); ws.set('E%d' % r, e.get('group') or '')
            ws.set('F%d' % r, e.get('tempo') or ''); ws.set('G%d' % r, e.get('rest') or ''); ws.set('H%d' % r, e.get('tech') or '')
            wk = e.get('wk') or {}
            for wi, top in enumerate((28, 41, 54, 67)):
                rw = top + i; P = wk.get(str(wi)) or {}
                if wi == 0:
                    ws.set('E%d' % rw, numish(P.get('sets') or e.get('sets'))); ws.set('F%d' % rw, numish(e.get('reps')))
                elif P.get('sets'):
                    ws.set('E%d' % rw, numish(P['sets']))
                kg = P.get('kg') if P.get('kg') not in (None, '') else e.get('kg')
                if kg not in (None, ''):
                    ws.set('G%d' % rw, 'RIR' if kg == 'Según RIR' else 'BW' if kg == 'Peso corporal' else numish(kg))
        for k, top in (('warm', 20), ('cool', 89)):
            for j, x in enumerate([x for x in re.split(r'\s*·\s*', d.get(k) or '') if x][:3]): ws.set('B%d' % (top + j), x)
        for j, x in enumerate([x for x in (d.get('note') or '').split('\n') if x.strip()][:3]): ws.set('B%d' % (99 + j), x)
    # Excel/LibreOffice recalculan las fórmulas al abrir
    cp = wb.find(qs('calcPr'))
    if cp is None: cp = etree.SubElement(wb, qs('calcPr'))
    cp.set('fullCalcOnLoad', '1')
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as z:
        for it in zin.infolist():
            if it.filename == 'xl/calcChain.xml': continue
            data = zin.read(it.filename)
            if it.filename == 'xl/workbook.xml': data = etree.tostring(wb, xml_declaration=True, encoding='UTF-8', standalone=True)
            else:
                for nm, s in sheets.items():
                    if path[nm] == it.filename: data = s.bytes()
            if it.filename == '[Content_Types].xml': data = data.replace(b'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>', b'')
            if it.filename == 'xl/_rels/workbook.xml.rels': data = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b'', data)
            z.writestr(it, data)
    return out.getvalue(), left

def export(J):
    """Devuelve (bytes, extensión, avisos)."""
    if J.get('kind') == 'diet': return fill_plan(J), 'docx', []
    b, left = fill_training(J)
    warn = ['No caben en la plantilla (4 días y 8 ejercicios por día): ' + '; '.join(left)] if left else []
    return b, 'xlsx', warn

if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(); ap.add_argument('datos'); ap.add_argument('-o', required=True)
    a = ap.parse_args()
    b, ext, w = export(json.load(open(a.datos, encoding='utf-8')))
    open(a.o, 'wb').write(b)
    for x in w: print('Aviso:', x)
