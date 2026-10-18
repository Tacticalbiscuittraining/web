#!/usr/bin/env python3
"""Rellena las plantillas de Robert (tools/plantillas) con un plan de la app.
Es lo contrario de importer.py: lo que sale se puede volver a importar.

Uso:  python3 tools/exporter.py datos.json -o salida.docx|.xlsx
datos.json = {"kind": "diet"|"training"|"run", "name", "goal", "review", "weight", "data": {"diet"|"training": …}}
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

def place_meals(names, slots):
    """Asigna cada toma a una fila/columna de la plantilla. Las que no tienen hueco propio
    (p. ej. «Almuerzo») van al primer hueco libre después de la toma anterior y renombran esa fila.
    Devuelve {nombre: índice de slot} y {índice: nombre nuevo}."""
    pos, rename, used, last = {}, {}, set(), -1
    for nm in names:
        if nm in pos: last = max(last, pos[nm]); continue
        k = norm_meal(nm)
        i = slots.index(k) if k in slots else None
        if i is not None and i in used and rename.get(i): i = None
        if i is None:
            i = next((j for j in range(last + 1, len(slots)) if j not in used and not any(norm_meal(n) == slots[j] for n in names)), None)
            if i is not None: rename[i] = nm
        if i is None: continue
        pos[nm] = i; used.add(i); last = i
    return pos, rename

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
            body_rows = [r for r in R[1:] if not ctext(cells(r)[0]).upper().startswith('TOTAL')]
            total_row = next((r for r in R if ctext(cells(r)[0]).upper().startswith('TOTAL')), None)
            pos, rename = place_meals([mm.get('name') or '' for mm in ms], [norm_meal(ctext(cells(r)[0])) for r in body_rows])
            for i, nm in rename.items(): set_cell(cells(body_rows[i])[0], nm, body_rows[i])
            S = [0, 0, 0, 0]
            for mm in ms:
                i = pos.get(mm.get('name') or '')
                r = body_rows[i] if i is not None else None
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
            # tomas con otro nombre (p. ej. «Almuerzo»): ocupan un hueco libre y renombran la columna
            names = []
            for i in range(7):
                key = 'd%d' % i if ('d%d' % i) in T else ('train' if trainDays[i] else 'rest')
                for mm in (T.get(key) or {}).get('meals') or []:
                    if (mm.get('name') or '') not in names: names.append(mm.get('name') or '')
            pos, rename = place_meals(names, cols)
            hdr = cells(R[0])[1:]
            for j, nm in rename.items(): set_cell(hdr[j], nm.upper(), R[0])
            cols = [rename.get(j) or c for j, c in enumerate(cols)]
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
                    j = pos.get(mm.get('name') or '')
                    k = cols[j] if j is not None else None
                    if k and k not in txt: txt[k] = meal_menu(mm); last = k
                    elif k: txt[k] += '\n' + (mm.get('name') or '') + ': ' + meal_menu(mm)
                    elif last: txt[last] += '\n' + (mm.get('name') or '') + ': ' + meal_menu(mm)
                    else: txt[cols[0]] = (mm.get('name') or '') + ': ' + meal_menu(mm) + ('\n' + txt[cols[0]] if cols[0] in txt else '')
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
    def cache(self, ref, v):
        """Guarda el resultado de la fórmula de la celda (la fórmula se queda)."""
        c = self.find_cell(ref)
        if c is None or c.find(qs('f')) is None: return
        for ch in list(c):
            if ch.tag != qs('f'): c.remove(ch)
        if isinstance(v, (int, float)): c.attrib.pop('t', None); etree.SubElement(c, qs('v')).text = str(v)
        else: c.set('t', 'str'); etree.SubElement(c, qs('v')).text = v
    def find_cell(self, ref):
        _, rn = split_ref(ref)
        for r in self.data.findall(qs('row')):
            if int(r.get('r')) == rn:
                for c in r.findall(qs('c')):
                    if c.get('r') == ref: return c
                return None
        return None
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

# Las fórmulas de la plantilla no llevan resultado guardado: Excel las recalcula al abrir, pero la vista previa
# del iPhone (y otros visores) las muestra vacías. Se guarda el resultado de las sencillas: =IF(X="","",Y)
IF_EMPTY = re.compile(r'^IF\((?:([^!()]+)!)?\$?([A-Z]+)\$?(\d+)="","",(?:([^!()]+)!)?\$?([A-Z]+)\$?(\d+)\)$')
def cache_formulas(get, names, shared):
    done = {}
    def val(sheet, ref):
        if (sheet, ref) in done: return done[(sheet, ref)]
        c = get(sheet).find_cell(ref)
        if c is None: return ''
        if c.find(qs('f')) is not None: return None  # fórmula aún sin resolver
        v, t = c.find(qs('v')), c.get('t')
        if t == 'inlineStr': return ''.join(x.text or '' for x in c.iter(qs('t')))
        if v is None or v.text is None: return ''
        if t == 's': return shared[int(v.text)]
        return v.text if t in ('str', 'e') else float(v.text)
    todo = []
    for n in names:
        for c in get(n).data.iter(qs('c')):
            f = c.find(qs('f'))
            m = IF_EMPTY.match((f.text or '').strip()) if f is not None else None
            if m: todo.append((n, c, m))
    sh_of = lambda n, g: g.strip("'") if g else n
    for _ in range(6):  # cadenas: semana 4 ← semana 3 ← semana 2 ← semana 1
        left = []
        for n, c, m in todo:
            x = val(sh_of(n, m.group(1)), m.group(2) + m.group(3))
            y = '' if x == '' else val(sh_of(n, m.group(4)), m.group(5) + m.group(6))
            if x is None or y is None: left.append((n, c, m)); continue
            done[(n, c.get('r'))] = y
            for ch in list(c):
                if ch.tag != qs('f'): c.remove(ch)
            if isinstance(y, float):
                c.attrib.pop('t', None); etree.SubElement(c, qs('v')).text = str(int(y)) if y.is_integer() else repr(y)
            else:
                c.set('t', 'str'); etree.SubElement(c, qs('v')).text = y
        if not left: break
        todo = left

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
    days = [d for d in (T.get('days') or []) if d.get('exs') and not d.get('of')]
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
    def fill_day(ws, d):
        exs = d['exs'][:8]
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
            ow = T.get('onlyWeek')
            if isinstance(ow, int) and 0 < ow < 4:
                rw = (28, 41, 54, 67)[ow] + i
                ws.set('E%d' % rw, numish(e.get('sets'))); ws.set('F%d' % rw, numish(e.get('reps')))
            D = e.get('done') or {}
            if D:
                rw = (28, 41, 54, 67)[ow if isinstance(ow, int) and 0 <= ow < 4 else 0] + i
                for col, key in (('I', 'sets'), ('J', 'reps'), ('K', 'kg'), ('L', 'rir')): ws.set('%s%d' % (col, rw), numish(D.get(key)))
        for k, top in (('warm', 20), ('cool', 89)):
            for j, x in enumerate([x for x in re.split(r'\s*·\s*', d.get(k) or '') if x][:3]): ws.set('B%d' % (top + j), x)
        for j, x in enumerate([x for x in (d.get('note') or '').split('\n') if x.strip()][:3]): ws.set('B%d' % (99 + j), x)
    # hojas nuevas copiadas de una de la plantilla (mismo diseño), justo detrás de «after»
    extra, ct_add, rel_add = {}, [], []
    nsheet = max(int(re.search(r'(\d+)\.xml$', v).group(1)) for v in path.values()) + 1
    ndraw = 1 + max([int(m) for m in re.findall(r'xl/drawings/drawing(\d+)\.xml$', '\n'.join(zin.namelist()), re.M)] or [0])
    rel_ids = [int(r.get('Id')[3:]) for r in rels if r.get('Id', '').startswith('rId')]
    def clone(base, name, after):
        nonlocal nsheet, ndraw
        part = 'xl/worksheets/sheet%d.xml' % nsheet
        sheets[name] = Sheet(zin.read(path[base])); path[name] = part
        brel = path[base].replace('worksheets/', 'worksheets/_rels/') + '.rels'
        if brel in zin.namelist():
            rx = zin.read(brel)
            dm = re.search(rb'Target="../drawings/(drawing\d+)\.xml"', rx)
            if dm:
                dsrc = 'xl/drawings/%s.xml' % dm.group(1).decode()
                extra['xl/drawings/drawing%d.xml' % ndraw] = zin.read(dsrc)
                drel = dsrc.replace('drawings/', 'drawings/_rels/') + '.rels'
                if drel in zin.namelist(): extra['xl/drawings/_rels/drawing%d.xml.rels' % ndraw] = zin.read(drel)
                ct_add.append('<Override PartName="/xl/drawings/drawing%d.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>' % ndraw)
                rx = rx.replace(dm.group(0), b'Target="../drawings/drawing%d.xml"' % ndraw); ndraw += 1
            extra[part.replace('worksheets/', 'worksheets/_rels/') + '.rels'] = rx
        ct_add.append('<Override PartName="/%s" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' % part)
        rid_new = 'rId%d' % (max(rel_ids) + 1); rel_ids.append(max(rel_ids) + 1)
        rel_add.append('<Relationship Id="%s" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet%d.xml"/>' % (rid_new, nsheet))
        S_ = wb.find(qs('sheets')); be = [s for s in S_ if s.get('name') == after][0]
        while be.getnext() is not None and be.getnext().get('name', '').startswith(after + ' ('): be = be.getnext()
        pos = list(S_).index(be) + 1
        ne = etree.Element(qs('sheet')); ne.set('name', name); ne.set('sheetId', str(max(int(s.get('sheetId')) for s in S_) + 1)); ne.set('state', 'visible'); ne.set(R, rid_new)
        be.addnext(ne)
        dn = wb.find(qs('definedNames'))
        for x in (dn if dn is not None else []):
            if x.get('localSheetId') is not None and int(x.get('localSheetId')) >= pos: x.set('localSheetId', str(int(x.get('localSheetId')) + 1))
        nsheet += 1
        ws = sheets[name]; ws.set('B2', 'PLAN DE ENTRENAMIENTO · %s' % name.upper())
        return ws
    # más de 8 ejercicios: hoja «Día A (cont.)» con los siguientes, numerados a partir del 9
    def more(name, d, base):
        rest, k = d['exs'][8:], 1
        while rest:
            cn = name + ' (cont.)' if k == 1 else '%s (cont. %d)' % (name, k)
            ws = clone(base, cn, name if k == 1 else prev); prev = cn
            for i in range(8):
                for top in (10, 28, 41, 54, 67): ws.set('B%d' % (top + i), 8 * k + i + 1)
            if name not in day_sheets:
                for ref in ('F5', 'K5'): ws.set(ref, cellval(sheets[name], ref))
            dd = dict(d, exs=rest[:8], warm='', cool='', note=''); fill_day(ws, dd); all_days.append((cn, dd))
            rest, k = rest[8:], k + 1
    def cellval(ws, ref):
        c = ws.find_cell(ref)
        if c is None or c.find(qs('f')) is not None: return ''
        return ''.join(x.text or '' for x in c.iter(qs('t')))
    all_days, slot = [], {}
    LET = 'ABCDEFGHIJKLMNOP'
    for di, d in enumerate(days):
        wd = d.get('wd'); wdn = WD[wd] if isinstance(wd, int) and 0 <= wd < 7 else ''
        if di < len(day_sheets):
            name = day_sheets[di]; ws = sh(name); rr = 22 + di
            res.set('C%d' % rr, wdn); res.set('D%d' % rr, d.get('focus') or '')
            res.set('G%d' % rr, numish(d.get('dur') or '')); res.set('H%d' % rr, d.get('place') or '')
        else:  # más días que la plantilla: hoja nueva igual que la del último día
            name = 'Día ' + LET[di] if di < len(LET) else 'Día %d' % (di + 1)
            ws = clone(day_sheets[-1], name, all_days[-1][0] if all_days else day_sheets[-1])
            ws.set('F5', d.get('focus') or ''); ws.set('K5', wdn)
            ws.set('B3', 'Tactical Biscuit Training · %s min · %s' % (d.get('dur') or '60', d.get('place') or 'Gimnasio'))
        slot[d['name']] = name
        fill_day(ws, d); all_days.append((name, d))
        if len(d['exs']) > 8: more(name, d, day_sheets[min(di, len(day_sheets) - 1)])
    # variantes (p. ej. «Día B · Casa»): hoja propia «Día B (Casa)» copiada de la del día base, justo detrás
    for v in [d for d in (T.get('days') or []) if d.get('exs') and d.get('of')]:
        base = slot.get(v['of'])
        if not base: left.append('%s (variante de %s)' % (v.get('name', ''), v['of'])); continue
        label = v.get('place') or v['name'].split('·')[-1].strip()
        name = '%s (%s)' % (base, label)
        ws = clone(base if base in day_sheets else day_sheets[-1], name, base)
        if base not in day_sheets:
            for ref in ('F5', 'K5'): ws.set(ref, cellval(sheets[base], ref))
        fill_day(ws, v); all_days.append((name, v))
        if len(v['exs']) > 8: left.append('%s: %s' % (v.get('name', ''), ', '.join(e.get('name', '') for e in v['exs'][8:])))
    # resultados que no son =IF(X="","",Y): fechas de las semanas, cabeceras y enlaces a vídeo
    st = dt.date.fromisoformat(meso['start']) if re.match(r'\d{4}-\d\d-\d\d$', meso.get('start') or '') else None
    rng = lambda i: '%s – %s' % ((st + dt.timedelta(7 * i)).strftime('%d/%m'), (st + dt.timedelta(7 * i + 6)).strftime('%d/%m')) if st else ''
    if st: res.cache('G10', (st + dt.timedelta(27) - dt.date(1899, 12, 30)).days)
    W = (meso.get('weeks') or [])[:4]
    for i in range(4): res.cache('C%d' % (15 + i), rng(i))
    if 'Check-in semanal' in path:
        for r in range(6, 10): sh('Check-in semanal').cache('K%d' % r, '0 / %d' % (len(days) or 4))
    used = dict(all_days)
    for name in day_sheets + [n for n, _ in all_days if n not in day_sheets]:
        ws = sh(name); d = used.get(name) or {'exs': []}
        for i, top in enumerate((24, 37, 50, 63)):
            w = W[i] if i < len(W) else {}
            ws.cache('B%d' % top, 'SEMANA %d  ·  %s  ·  RIR objetivo %s' % (i + 1, str(w.get('phase') or '').upper(), w.get('rir') if w.get('rir') is not None else '') + ('  ·  ' + rng(i) if st else ''))
        for i in range(len(d['exs'][:8])): ws.cache('M%d' % (10 + i), '▶ Ver vídeo')
    shared = []
    if 'xl/sharedStrings.xml' in zin.namelist():
        ss = etree.fromstring(zin.read('xl/sharedStrings.xml'))
        shared = [''.join(t.text or '' for t in si.iter(qs('t'))) for si in ss.findall(qs('si'))]
    cache_formulas(sh, [n for n in path if n == 'Resumen' or n.startswith('Día')], shared)
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
            if it.filename == '[Content_Types].xml':
                data = data.replace(b'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/>', b'')
                data = data.replace(b'</Types>', ''.join(ct_add).encode() + b'</Types>')
            if it.filename == 'xl/_rels/workbook.xml.rels':
                data = re.sub(rb'<Relationship [^>]*calcChain[^>]*/>', b'', data)
                data = data.replace(b'</Relationships>', ''.join(rel_add).encode() + b'</Relationships>')
            z.writestr(it, data)
        for nm, s in sheets.items():
            if path[nm] not in zin.namelist(): z.writestr(path[nm], s.bytes())
        for k, b in extra.items(): z.writestr(k, b)
    return out.getvalue(), left

# ---------------------------------------------------------------- plan de carrera (Excel propio; importer.parse_run lo vuelve a leer)
RUN_TN = {'suave': 'Suave', 'larga': 'Salida larga', 'calidad': 'Calidad', 'test': 'Test', 'carrera': 'Día de carrera', 'recup': 'Recuperación', 'fuerza': 'Fuerza', 'descanso': 'Descanso'}
RUN_TC = {'suave': 'D5EEEB', 'larga': 'F4ECCB', 'calidad': 'F9DDD5', 'test': 'E4DCEF', 'carrera': 'F2D0CE', 'recup': 'DCE9E4', 'fuerza': 'E1E5E8', 'descanso': 'EEF0F2'}

def fill_run(J):
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    R = (J.get('data') or {}).get('run') or {}
    race, WK = R.get('race') or {}, R.get('weeks') or []
    dark, gold = PatternFill('solid', fgColor='1F2A33'), PatternFill('solid', fgColor='C9A227')
    thin = Side(style='thin', color='C8CDD2'); box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical='top')
    def iso(s):
        try: return dt.date.fromisoformat(str(s)[:10])
        except Exception: return None
    st = iso(R.get('start'))
    wb = openpyxl.Workbook()
    def title(ws, t, sub, width):
        ws.append([None, t]); ws.cell(1, 2).font = Font(bold=True, size=16, color='FFFFFF')
        for c in range(1, width + 1): ws.cell(1, c).fill = dark
        ws.append([None, sub]); ws.cell(2, 2).font = Font(italic=True, color='5B6670'); ws.append([])
    def head(ws, cols):
        ws.append(cols); r = ws.max_row
        for i, _ in enumerate(cols, 1):
            c = ws.cell(r, i); c.font = Font(bold=True, color='FFFFFF'); c.fill = dark; c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center'); c.border = box
    def body(ws, vals, fills=None):
        ws.append(vals); r = ws.max_row
        for i, _ in enumerate(vals, 1):
            c = ws.cell(r, i); c.alignment = wrap; c.border = box
            if fills and fills[i - 1]: c.fill = PatternFill('solid', fgColor=fills[i - 1])
    def section(ws, t, width):
        ws.append([]); ws.append([t]); r = ws.max_row
        for c in range(1, width + 1): ws.cell(r, c).fill = gold
        ws.cell(r, 1).font = Font(bold=True, color='1F2A33')
    # ---- hoja 1: planificación semana a semana
    ws = wb.active; ws.title = 'Plan carrera'
    title(ws, 'PLAN DE CARRERA · ' + (race.get('name') or J.get('name') or ''), 'Tactical Biscuit Training · ' + (J.get('name') or ''), 12)
    head(ws, ['SEMANA', 'FECHAS', 'FASE'] + DAYS + ['VOLUMEN', 'CLAVE DE LA SEMANA'])
    for w, W in enumerate(WK):
        dts = W.get('dates') or (('%s – %s' % ((st + dt.timedelta(7 * w)).strftime('%d/%m'), (st + dt.timedelta(7 * w + 6)).strftime('%d/%m'))) if st else '')
        ds = W.get('days') or []
        cells, fills = [], []
        for i in range(7):
            S = ds[i] if i < len(ds) else None
            cells.append((S or {}).get('txt') or '—'); fills.append(RUN_TC.get((S or {}).get('t')) if S and S.get('txt') else None)
        vol = W.get('vol') or ''
        body(ws, ['Semana %d' % (w + 1), dts, W.get('phase') or ''] + cells + [('≈ ' + vol) if vol else '', W.get('key') or ''], [None, None, None] + fills + [None, None])
    for x in R.get('legend') or []: ws.append(['• ' + x])
    for col, wd in zip('ABCDEFGHIJKL', [11, 14, 16] + [26] * 7 + [12, 32]): ws.column_dimensions[col].width = wd
    ws.freeze_panes = 'D5'
    # ---- hoja 2: resumen (carrera, zonas, reglas)
    ws = wb.create_sheet('Resumen')
    title(ws, 'RESUMEN DEL PLAN', 'Tactical Biscuit Training', 4)
    car = race.get('name') or ''
    if race.get('place'): car += ' · ' + race['place']
    rd = iso(race.get('date'))
    for k, v in [('Cliente', J.get('name') or ''), ('Carrera', car), ('Fecha', dt.datetime(rd.year, rd.month, rd.day) if rd else ''), ('Distancia / desnivel', race.get('dist') or ''),
                 ('Modalidad', race.get('mod') or ''), ('Tiempo estimado', race.get('est') or ''), ('Salida', race.get('salida') or ''), ('Obligatorio', race.get('oblig') or ''),
                 ('Inicio del plan', dt.datetime(st.year, st.month, st.day) if st else ''), ('Semanas', len(WK))]:
        if v in ('', None): continue
        ws.append([k, v]); r = ws.max_row; ws.cell(r, 1).font = Font(bold=True)
        if isinstance(v, dt.datetime): ws.cell(r, 2).number_format = 'DD/MM/YYYY'
        ws.cell(r, 2).alignment = Alignment(horizontal='left')
    if R.get('zones'):
        section(ws, 'TUS ZONAS DE PULSO', 4); head(ws, ['ZONA', 'PULSO (ppm)', 'SENSACIÓN', 'PARA QUÉ LA USAMOS'])
        for z in R['zones']: body(ws, [str(x) for x in (list(z) + ['', '', '', ''])[:4]])
        if R.get('zonesNote'): ws.append([R['zonesNote']])
    if R.get('rules'):
        section(ws, 'REGLAS DEL PLAN', 4)
        for i, x in enumerate(R['rules'], 1): body(ws, [str(i), x])
    for col, wd in zip('ABCD', [22, 40, 30, 40]): ws.column_dimensions[col].width = wd
    # ---- hoja 3: estrategia de carrera
    S = R.get('strat') or {}
    if S.get('segs'):
        ws = wb.create_sheet('Estrategia')
        title(ws, 'ESTRATEGIA DE CARRERA', race.get('name') or '', 8)
        head(ws, ['KM', 'TRAMO', 'DESNIVEL', 'PULSO', 'CÓMO', 'CÓMO HACERLO', 'TIEMPO', 'ACUMULADO'])
        for x in S['segs']: body(ws, [str(v) for v in (list(x) + [''] * 8)[:8]])
        if S.get('total'): ws.append([S['total']])
        for key, t, cols, n in [('food', 'COMIDA Y BEBIDA DURANTE LA CARRERA', ['MOMENTO', 'QUÉ TOMAR', 'NOTAS'], 'foodNote'), ('controls', 'CONTROLES Y TU PASO ESTIMADO', ['KM', 'CONTROL', 'HORARIO', 'TU PASO', 'QUÉ HACER'], 'controlsNote')]:
            if not S.get(key): continue
            ws.append([]); head(ws, cols)
            for x in S[key]: body(ws, [str(v) for v in (list(x) + [''] * len(cols))[:len(cols)]])
            if S.get(n): ws.append([S[n]])
        if S.get('prev'):
            section(ws, 'DÍAS PREVIOS Y MATERIAL', 8)
            for x in S['prev']: body(ws, [str(v) for v in (list(x) + ['', ''])[:2]])
        for col, wd in zip('ABCDEFGH', [12, 26, 14, 14, 18, 40, 12, 12]): ws.column_dimensions[col].width = wd
    out = io.BytesIO(); wb.save(out); return out.getvalue()

# ---------------------------------------------------------------- composición corporal (Excel propio; importer.parse_comp lo vuelve a leer)
def fill_comp(J):
    """J['comp'] = {name, altura, nac, sexo, notas, fechas: [iso…], rows: [[etiqueta, 'h'|'in'|'calc', [valores…]], …]}"""
    import openpyxl
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    C = J.get('comp') or {}
    F = C.get('fechas') or []
    n = max(1, len(F))
    dark, gold, calc = PatternFill('solid', fgColor='1F2A33'), PatternFill('solid', fgColor='C9A227'), PatternFill('solid', fgColor='EEF0F2')
    thin = Side(style='thin', color='C8CDD2'); box = Border(left=thin, right=thin, top=thin, bottom=thin)
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = 'Composición corporal'
    ws.append(['Composición corporal · ' + (J.get('name') or '')]); ws.cell(1, 1).font = Font(bold=True, size=16, color='FFFFFF')
    for c in range(1, n + 2): ws.cell(1, c).fill = dark
    ws.append(['Tactical Biscuit Training']); ws.cell(2, 1).font = Font(italic=True, color='5B6670'); ws.append([])
    def iso(s):
        try: return dt.datetime.fromisoformat(str(s)[:10])
        except Exception: return None
    ws.append(['Fecha'] + [iso(x) for x in F]); r = ws.max_row
    for c in range(1, n + 2):
        x = ws.cell(r, c); x.font = Font(bold=True, color='FFFFFF'); x.fill = dark; x.border = box; x.alignment = Alignment(horizontal='center')
        if c > 1: x.number_format = 'DD/MM/YYYY'
    for lab, kind, vals in C.get('rows') or []:
        if kind == 'h':
            ws.append([str(lab).upper()]); r = ws.max_row
            for c in range(1, n + 2): ws.cell(r, c).fill = gold
            ws.cell(r, 1).font = Font(bold=True, color='1F2A33'); continue
        ws.append([lab] + [v if v not in ('', None) else None for v in vals]); r = ws.max_row
        for c in range(1, n + 2):
            x = ws.cell(r, c); x.border = box
            if kind == 'calc': x.fill = calc; x.font = Font(italic=True, color='5B6670')
            if c > 1: x.alignment = Alignment(horizontal='center')
    ws.column_dimensions['A'].width = 38
    for c in range(2, n + 2): ws.column_dimensions[openpyxl.utils.get_column_letter(c)].width = 13
    ws.freeze_panes = 'B5'
    d = wb.create_sheet('Datos')
    for k, v in [('Cliente', J.get('name') or ''), ('Fecha de nacimiento', C.get('nac') or ''), ('Sexo', C.get('sexo') or ''), ('Altura (cm)', C.get('altura') or ''), ('Notas', C.get('notas') or '')]:
        d.append([k, v]); d.cell(d.max_row, 1).font = Font(bold=True)
    d.column_dimensions['A'].width = 22; d.column_dimensions['B'].width = 60
    out = io.BytesIO(); wb.save(out); return out.getvalue()

def export(J):
    """Devuelve (bytes, extensión, avisos)."""
    if J.get('kind') == 'diet': return fill_plan(J), 'docx', []
    if J.get('kind') == 'run': return fill_run(J), 'xlsx', []
    if J.get('kind') == 'comp': return fill_comp(J), 'xlsx', []
    b, left = fill_training(J)
    warn = ['No caben en la plantilla (más de 8 ejercicios en una variante): ' + '; '.join(left)] if left else []
    return b, 'xlsx', warn

if __name__ == '__main__':
    import argparse
    ap = argparse.ArgumentParser(); ap.add_argument('datos'); ap.add_argument('-o', required=True)
    a = ap.parse_args()
    b, ext, w = export(json.load(open(a.datos, encoding='utf-8')))
    open(a.o, 'wb').write(b)
    for x in w: print('Aviso:', x)
