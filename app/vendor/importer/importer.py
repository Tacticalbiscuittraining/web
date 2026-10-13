#!/usr/bin/env python3
"""Convierte las plantillas de Robert (Word / Excel) en un archivo de importación
para Team Biscuit (botón «Importar archivo de plan»).

Uso:
  python3 tools/importer.py --plan Plan.docx --entreno Entreno.xlsx \
      --comp COMPOSICION.xlsx --ficha Ficha.docx -o cliente.json

Todos los archivos son opcionales; se importa lo que se pase.
Al terminar imprime un informe con lo que no ha podido interpretar.
"""
import argparse, datetime as dt, json, re, secrets, sys
import docx, openpyxl
from docx.oxml.ns import qn

REPORT = []
def warn(msg): REPORT.append(msg)
def nid(): return 'i' + secrets.token_hex(5)
def r1(x): return round(x * 10) / 10
def num(s):
    if s is None: return None
    if isinstance(s, (int, float)): return float(s)
    m = re.search(r'-?\d+(?:[.,]\d+)?', str(s).replace('.', '').replace(',', '.') if re.search(r'\d\.\d{3}', str(s)) else str(s).replace(',', '.'))
    return float(m.group(0)) if m else None

# ---------------------------------------------------------------- alimentos
# Tabla oficial de alimentos (src/foods.json, la misma que usa la app):
# USDA FoodData Central + tablas españolas (Moreiras/MAPA) + etiquetas de supermercado.
# Valores por 100 g de porción comestible; HC = hidratos disponibles (sin fibra).
import os, unicodedata
FOODDB = json.load(open(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src', 'foods.json'), encoding='utf-8'))
def _n(t):
    t = unicodedata.normalize('NFD', t.lower())
    return ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn')
GENERIC = {'aove': 'AOVE', 'fruta': 'Fruta', 'verdura': 'Verdura', 'verduras': 'Verdura', 'pescado blanco': 'Pescado blanco', 'legumbre': 'Legumbre'}
FOOD = []  # (regex, nombre que se muestra, clave de la tabla, [P, HC, G])
for x in FOODDB:
    for al in [re.sub(r'\s*\(.*?\)', '', x['n'])] + x.get('al', []):
        disp = GENERIC.get(_n(al), x['n'])
        FOOD.append((r'\b' + re.escape(_n(al)) + r'\b', disp, x['n'], [x['p'], x['c'], x['f']]))
FOOD.sort(key=lambda r: -len(r[0]))
EGG_G = 55  # huevo M sin cáscara

def parse_compact(cell):
    """'3 huevos + 70g avena + 150g plátano' · '2 latas atún (160g), sin pan' · 'verdura (sin AOVE)'"""
    items, notes, names = [], [], []
    for tok in re.split(r'\s*\+\s*|,\s*', cell):
        tok = tok.strip().rstrip('*').strip()
        if not tok: continue
        par = re.search(r'\((?!\d)([^)]*)\)', tok)
        if par:
            notes.append(par.group(1).strip()); tok = (tok[:par.start()] + tok[par.end():]).strip()
        m = re.match(r'^(\d+)\s+huevos?\b', tok, re.I)
        if m:
            n = int(m.group(1)); g = n * EGG_G; per = next([x['p'], x['c'], x['f']] for x in FOODDB if x['n'] == 'Huevos')
            items.append({'id': nid(), 'food': 'Huevos M (%d)' % n, 'g': g, 'p': r1(per[0]*g/100), 'c': r1(per[1]*g/100), 'f': r1(per[2]*g/100), 'ref': 'Huevos', 'alt': '', '_cnt': 1})
            names.append('huevos'); continue
        m = re.match(r'^(\d+(?:[.,]\d+)?)\s*(g|ml)\s*(.+)$', tok, re.I)
        if m:
            it = parse_item(m.group(1) + ' ' + m.group(2) + ' ' + m.group(3))
        elif re.match(r'^\d+\s+.+\(\d', tok):
            it = parse_item(tok)
        elif re.match(r'^verdura', tok, re.I):
            it = parse_item('200 g ' + tok); it['_cnt'] = 1
        else:
            notes.append(tok); continue
        if it and 'note' not in it:
            items.append(it)
            nm = re.sub(r'^(\d+(?:[.,]\d+)?\s*(g|ml)\s*|\d+\s+latas?\s+)', '', tok, flags=re.I)
            nm = re.sub(r'\s*\(.*$', '', nm).strip()
            if not re.search(r'aove|aceite', nm, re.I): names.append(nm.lower())
        elif it: notes.append(it['note'])
    dish = (', '.join(names[:-1]) + ' y ' + names[-1]) if len(names) > 1 else (names[0] if names else cell)
    return dish[:1].upper() + dish[1:], items, notes

def split_top(s):
    """Separa por «+» y «,» fuera de paréntesis."""
    out, depth, cur = [], 0, ''
    for ch in s:
        if ch == '(': depth += 1
        elif ch == ')': depth = max(0, depth - 1)
        if ch in '+,' and depth == 0:
            out.append(cur); cur = ''
        else: cur += ch
    out.append(cur)
    return [x.strip() for x in out if x.strip()]

def food_sub(txt):
    """Como food_of, pero tolera palabras pegadas («patatacocida», «yjamon»)."""
    t = _n(txt).replace(' ', '')
    for rx, disp, ref, m in FOOD:
        key = rx[2:-2].replace('\\ ', '').replace(' ', '')
        if key and key in t: return disp, ref, m
    return None

RAW_OF = {'Lentejas cocidas': 'Lentejas (crudas)', 'Arroz cocido': 'Arroz (crudo)', 'Pasta cocida': 'Pasta (cruda)', 'Patata cocida': 'Patata (cruda)'}
GEN_PROT = r'prote[ií]na|carne|pescado o carne|pescado'
def parse_free(cell):
    """Texto libre del menú: 'Avena (110g) con canela,4 huevos enteros,2 platanos (180g) + cafe'"""
    s = re.sub(r'(\d)\s*(g|ml)(?=[A-Za-zÁÉÍÓÚáéíóúñÑ])', r'\1\2 ', cell)
    toks = []
    for tok in split_top(s):
        inner = re.search(r'\(([^()]*\d\s*(?:g|ml)[^()]*\+[^()]*)\)', tok)
        if inner: toks += split_top(inner.group(1))
        else: toks += [x for x in re.split(r'(?<=\))\s*(?:\by\b|\bcon\b)\s+(?=[^()]*\([^()]*\d)', tok) if x.strip()]
    items, notes, names = [], [], []
    for tok in toks:
        tok = tok.strip(' .*:')
        if not tok: continue
        low = _n(tok)
        m = re.search(r'(\d+(?:[.,]\d+)?)\s*(g|ml)\b', low)
        eg = re.search(r'(\d+)\s*huevos?|huevos?[^()]*\((\d+)\s*ud', low)
        name = re.sub(r'\([^)]*\)?', ' ', tok)
        name = re.sub(r'\d+(?:[.,]\d+)?\s*(g|ml)\b|\d+\s*(latas?|uds?|piezas?)|\b(crudo|cruda|escurrido|enteros?|aprox\.?)\b|~|%|^\s*\d+(/\d+)?\s*', ' ', name, flags=re.I)
        name = re.sub(r'.*:', '', name)
        name = re.sub(r'\s+', ' ', name).strip(' ,.-')
        if eg and not m:
            n = int(eg.group(1) or eg.group(2)); g = n * EGG_G; unit = 'g'
            f = ('Huevos (%d)' % n, 'Huevos', next([x['p'], x['c'], x['f']] for x in FOODDB if x['n'] == 'Huevos'))
        elif m:
            g = float(m.group(1).replace(',', '.')); unit = m.group(2)
            if re.search(r'\d\s*g\s*pan', low): name = 'pan integral' if 'integral' in low else 'pan'
            f = food_of(name) or food_sub(name) or food_of(tok)
            if not f and re.search(GEN_PROT, low):
                x = next(x for x in FOODDB if x['n'] == 'Pechuga de pollo'); f = ('Proteína (carne o pescado)', x['n'], [x['p'], x['c'], x['f']])
        else:
            notes.append(tok); continue
        if not f:
            warn('Alimento sin datos nutricionales: «%s» (se importa sin macros)' % tok)
            items.append({'id': nid(), 'food': name[:1].upper() + name[1:], 'g': g, 'p': 0, 'c': 0, 'f': 0, 'ref': '', 'alt': '', '_src': name, **({'u': 'ml'} if unit == 'ml' else {})})
        else:
            if f[1] == 'Pechuga de pollo a la plancha (cocinada)' and not re.search(r'cocin|cocid', low):
                x = next(x for x in FOODDB if x['n'] == 'Pechuga de pollo'); f = (x['n'], x['n'], [x['p'], x['c'], x['f']])
            if re.search(r'\bcrud', low) and f[1] in RAW_OF:
                x = next((x for x in FOODDB if x['n'] == RAW_OF[f[1]]), None)
                if x: f = (x['n'], x['n'], [x['p'], x['c'], x['f']])
            disp, ref, per = f
            it = {'id': nid(), 'food': disp, 'g': g, 'p': r1(per[0] * g / 100), 'c': r1(per[1] * g / 100), 'f': r1(per[2] * g / 100), 'ref': ref, 'alt': '', '_src': name or disp}
            if unit == 'ml': it['u'] = 'ml'
            items.append(it)
        nm2 = re.sub(r'\s*\(.*\)$', '', items[-1]['food']).lower() if items[-1].get('ref') else name.lower()
        if nm2 and not re.search(r'aove|aceite', nm2, re.I): names.append(nm2)
    dish = (', '.join(names[:-1]) + ' y ' + names[-1]) if len(names) > 1 else (names[0] if names else cell)
    return dish[:1].upper() + dish[1:], items, notes

def food_of(txt):
    t = _n(txt)
    t = t + ' ' + re.sub(r'\b(\w{4,}?)(es|s)\b', r'\1', t)  # también en singular («platanos» → «platano»)
    best = None  # el alimento que aparece antes en el texto («arroz/legumbre» → arroz); a igualdad, el nombre más largo
    for rx, disp, ref, m in FOOD:
        mm = re.search(rx, t)
        if mm and (best is None or mm.start() < best[0]): best = (mm.start(), (disp, ref, m))
    return best[1] if best else None

def parse_item(line):
    """'160 g yogur griego 0%' · '2 tortitas arroz (26 g)' · '10 ml AOVE' · '1 huevo duro (55 g)'"""
    line = line.strip().strip('·•-').strip()
    if not line or line in ('—', '-'):
        return None
    unit = 'g'
    m = re.match(r'^(\d+(?:[.,]\d+)?)\s*(g|ml)\s+(.+)$', line, re.I)
    if m:
        g, unit, name = float(m.group(1).replace(',', '.')), m.group(2).lower(), m.group(3)
    else:
        m = re.match(r'^(\d+)\s+(.+?)\s*\((\d+(?:[.,]\d+)?)\s*(g|ml)\)\s*(.*)$', line, re.I)
        if m:
            g, unit, name = float(m.group(3).replace(',', '.')), m.group(4).lower(), m.group(2) + (' ' + m.group(5) if m.group(5) else '')
            name = name.strip() + ' (' + m.group(1) + ')'
        else:
            return {'note': line}
    f = food_of(name) or food_sub(name)
    if not f:
        warn('Alimento sin datos nutricionales: «%s» (se importa sin macros)' % name)
        return {'id': nid(), 'food': name[:1].upper() + name[1:], 'g': g, 'p': 0, 'c': 0, 'f': 0, 'ref': '', 'alt': '', **({'u': 'ml'} if unit == 'ml' else {})}
    disp, ref, per = f
    extra = re.search(r'\((\d+)\)$', name)
    if extra: disp = re.sub(r'\s*\(.*\)$', '', disp) + ' (' + extra.group(1) + ')'
    if re.search(r'duro', name, re.I): disp = 'Huevo duro'
    if re.search(r'plancha', name, re.I) and ref == 'Huevos': disp = 'Huevo a la plancha'
    if re.search(r'tostad', name, re.I): disp += ' tostado'
    it = {'id': nid(), 'food': disp, 'g': g, 'p': r1(per[0] * g / 100), 'c': r1(per[1] * g / 100), 'f': r1(per[2] * g / 100), 'ref': ref, 'alt': ''}
    it['_src'] = name
    if unit == 'ml': it['u'] = 'ml'
    return it

CHECKED = set()
def check(items, target, label):
    """No toca nada: compara la toma con la tabla oficial y avisa si tu plantilla no cuadra."""
    key = label + '|' + '|'.join('%s%s' % (i['food'], i['g']) for i in items)
    if key in CHECKED: return
    CHECKED.add(key)
    calc = [sum(i[k] for i in items) for k in 'pcf']
    lim = [5, 8, 4]
    bad = [(nm, tv, cv) for nm, tv, cv, l in zip(('P', 'HC', 'G'), target, calc, lim) if tv is not None and abs(tv - cv) > l]
    if bad:
        warn('%s: tu plantilla dice %s; con la tabla oficial sale %s' % (label,
             ' · '.join('%s %g' % (nm, tv) for nm, tv, cv in bad), ' · '.join('%s %g' % (nm, round(cv, 1)) for nm, tv, cv in bad)))

# ---------------------------------------------------------------- Word helpers
def cell_text(tc):
    return '\n'.join(''.join(x.text or '' for x in p.iter(qn('w:t'))) for p in tc.iter(qn('w:p'))).strip()

def rows_of(tbl):
    out = []
    for tr in tbl.iter(qn('w:tr')):
        cells, prev = [], None
        for tc in tr.iter(qn('w:tc')):
            cells.append(cell_text(tc))
        out.append(cells)
    return out

def tables(doc):
    return [rows_of(t) for t in doc.element.body.iter(qn('w:tbl'))]

MEALS = ['Desayuno', 'Media mañana', 'Comida', 'Pre-entreno', 'Merienda', 'Post-entreno', 'Cena']
def norm_meal(s):
    s = s.strip().lower()
    for m in MEALS:
        if s.startswith(m.lower()): return m
    return s[:1].upper() + s[1:]

def kv_tables(T):
    kv = {}
    for t in T:
        for r in t:
            for i in range(0, len(r) - 1):
                if r[i] and r[i] not in kv: kv[r[i]] = r[i + 1]
    return kv

# ---------------------------------------------------------------- plan nutricional
def parse_plan(path):
    doc = docx.Document(path)
    T = tables(doc)
    kv = kv_tables(T)
    diet = {'supps': [], 'notes': ''}
    head = lambda t: (t[0][0] if t and t[0] else '').strip().upper()
    out = {}
    out['name'] = kv.get('Nombre completo', '').strip()
    out['goal'] = kv.get('Objetivo del plan', '').strip()
    ini = kv.get('Fecha de inicio', '')
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', ini)
    if m: diet['inicio'] = '%s-%02d-%02d' % (m.group(3), int(m.group(2)), int(m.group(1)))
    diet['dur'] = (kv.get('Duración') or kv.get('Duración del plan') or '').strip()
    out['phase'] = re.split(r'\s+--\s+|\s+—\s+', diet['dur'])[0].strip()
    rv = kv.get('Próxima revisión', '').strip()
    out['review'] = rv if rv and rv not in ('-', '—') else '—'
    diet['edad'] = num(kv.get('Edad')) or ''
    diet['altura'] = num(kv.get('Altura (cm)')) or ''
    diet['diasTxt'] = kv.get('Días entreno / sem.', '').strip()
    diet['nivel'] = kv.get('Nivel de actividad', '').strip()
    week = None; struct = {}; exch = None
    # formato «Etiqueta / valor» en una sola celda (diseño nuevo)
    for t in T:
        for r in t:
            for c in r:
                m2 = re.match(r'\s*(.+?)\s*/\s*(.+)$', c.replace('\n', ' / ', 1))
                if not m2: continue
                k, v = m2.group(1).strip().lower(), m2.group(2).strip()
                if k.startswith('peso actual'): out['_peso'] = num(v)
                elif k == 'tmb': diet['bmr'] = num(v.replace('.', '')) or 0
                elif k == 'tdee': diet['tdee'] = num(v.replace('.', '')) or 0
                elif k.startswith('calorías en entreno') or k.startswith('calorias en entreno'): out['_kt'] = num(v.replace('.', ''))
                elif k.startswith('calorías en descanso') or k.startswith('calorias en descanso'): out['_kr'] = num(v.replace('.', ''))
                elif k.startswith('calorías diarias') or k.startswith('calorias diarias'): diet['kcalAvg'] = num(v.replace('.', '')) or 0
                elif 'agua diaria' in k:
                    mm = re.match(r'\s*([\d,]+\s*[–-]\s*[\d,]+|[\d,]+)', v)
                    if mm: diet['waterTxt'] = mm.group(1).strip(); diet['water'] = num(re.split(r'[–-]', mm.group(1))[-1]) or 2
                elif 'suplementación' in k or 'suplementacion' in k:
                    if v and not v.lower().startswith('sin '): diet['supps'].append({'id': nid(), 'name': v.split(' ')[0], 'dose': v, 'when': '', 'obs': ''})
    prev_sec = ''
    for t in T:
        h = head(t)
        if len(t) == 1 and len(t[0]) == 1 and re.match(r'^\d\d\s', h): prev_sec = h; continue
        if h == 'PESO ACTUAL':
            v = t[1]; diet['goalW'] = num(v[1]) or 0; diet['bmr'] = num(v[2]) or 0; diet['tdee'] = num(v[3]) or 0
            out['_peso'] = num(v[0])
        elif h == 'KCAL DÍA ENTRENO':
            v = t[1]; out['_kt'], out['_kr'] = num(v[0]), num(v[1]); diet['kcalAvg'] = num(v[3]) or 0
        elif h == 'MACRONUTRIENTE':
            for r in t:
                if r[0].lower().startswith('fibra'): diet['fibra'] = r[1]
        elif 'AGUA' in h:
            w = t[0][1]
            mm = re.match(r'\s*([\d,]+\s*[–-]\s*[\d,]+|[\d,]+)', w)
            if mm:
                diet['waterTxt'] = mm.group(1).strip()
                diet['water'] = num(mm.group(1).split('–')[-1].split('-')[-1]) or 2
        elif h == 'SUPLEMENTO':
            for r in t[1:]:
                if r[0] and not re.match(r'(sin |ningun|—|-$)', r[0].strip().lower()):
                    diet['supps'].append({'id': nid(), 'name': r[0], 'dose': r[1], 'when': r[2], 'obs': r[3]})
        elif h == 'TOMA':
            key = 'train' if 'train' not in struct else 'rest'
            meals = {}
            for r in t[1:]:
                if r[0].upper().startswith('TOTAL'): continue
                if len(r) <= 4:
                    if num(r[1]) is not None: meals[norm_meal(r[0])] = {'time': '', 'text': '', 'tot': [num(r[1]), num(r[2]), num(r[3])]}
                elif r[2]:
                    meals[norm_meal(r[0])] = {'time': r[1], 'text': r[2], 'tot': [num(r[3]), num(r[4]), num(r[5])]}
            struct[key] = meals
        elif h in ('DÍA', 'DIA', '') and len(t[0]) > 2 and re.match(r'(DESAYUNO|ALMUERZO|MEDIA MA|COMIDA(?!S))', t[0][1].strip().upper()) and not week:
            cols = [norm_meal(x) for x in t[0][1:]]
            week = []; marks = []
            for r in t[1:8]:
                marks.append(1 if '☒' in r[0] else 0 if '☐' in r[0] else None)
                day = []
                for mname, cell in zip(cols, r[1:]):
                    lines = [x.strip() for x in cell.split('\n') if x.strip()]
                    if not lines or lines[0] in ('—', '-'): continue
                    dish, items, notes = lines[0], [], []
                    if len(lines) == 1 and ('+' in lines[0] or ',' in lines[0] or re.match(r'^\d', lines[0])):
                        r0 = len(REPORT); dish, items, notes = parse_compact(lines[0]); w1 = REPORT[r0:]; del REPORT[r0:]
                        d2, i2, n2 = parse_free(lines[0]); w2 = REPORT[r0:]; del REPORT[r0:]
                        q = lambda L: (len([x for x in L if x.get('ref')]), -len([x for x in L if not x.get('ref')]))
                        if q(i2) > q(items): dish, items, notes, w1 = d2, i2, n2, w2
                        REPORT.extend(w1)
                    for ln in lines[1:]:
                        it = parse_item(ln)
                        if it is None: continue
                        if 'note' in it: notes.append(it['note'])
                        else: items.append(it)
                    day.append({'id': nid(), 'name': mname, 'dish': dish, 'time': '', 'note': ' · '.join(notes), 'items': items})
                week.append(day)
        elif h.startswith('GRUPO'):
            exch = []
            for r in t[1:]:
                its = [re.sub(r'^[•·\s]+', '', x).strip() for x in r[1].split('\n') if x.strip()]
                exch.append([r[0], r[2] if len(r) > 2 else '', its, r[3] if len(r) > 3 else ''])
        elif h.startswith('⚠'):
            diet['avoid'] = t[0][1]
        elif h == 'PAUTA':
            diet['pautas'] = [[r[0], r[1]] for r in t[1:] if r[0]]
        elif prev_sec and re.search(r'NOTAS|OBSERVACIONES|RECOMENDACIONES', prev_sec) and len(t[0]) == 1 and not re.match(r'^\d\d\s', h):
            diet['notes'] = '\n'.join(re.sub(r'^[•·\s]+', '', r[0]).strip() for r in t if r[0].strip())
        elif h.startswith('OBSERVACIONES'):
            diet['notes'] = '\n'.join(re.sub(r'^[•·\s]+', '', r[0]).strip() for r in t[1:] if r[0].strip())
    if exch: diet['exch'] = exch
    # calibración de los días de ejemplo con los totales de la estructura del día
    def same(day, sm):
        return day and all(any(m['name'] == k and sm[k]['text'].replace(' ', '').lower().find(str(int(m['items'][0]['g'])) if m['items'] else '') >= 0 for m in day) for k in sm)
    types = {}
    if week:
        for i, day in enumerate(week):
            types['d%d' % i] = {'meals': day}
        for key in ('train', 'rest'):
            sm = struct.get(key)
            if not sm: continue
            # el día cuyo menú coincide con la tabla de estructura (p. ej. lunes = entreno, domingo = descanso)
            def mok(m):
                if m['name'] not in sm: return False
                tx = sm[m['name']]['text'].lower()
                if len(m['items']) != len(re.split(r'\s*\+\s*', tx)): return False
                return all((it.get('_cnt') or re.search(r'\b%s\s*(g|ml)\b' % int(it['g']), tx)) and
                           re.split(r'[\s(/]', it.get('_src', it['food']).lower())[0] in tx for it in m['items'])
            if not any(v['text'] for v in sm.values()):
                # estructura solo con macros por toma (sin menú): no hay día modelo que buscar
                continue
            cand = [i for i, d in enumerate(week) if set(m['name'] for m in d) == set(sm) and all(mok(m) for m in d)]
            if not cand:
                warn('No encuentro el día del menú que corresponde a la estructura de %s' % ('entreno' if key == 'train' else 'descanso'))
                continue
            for d in week:
                for m in d:
                    if m['items'] and not m.get('_cal') and mok(m): check(m['items'], sm[m['name']]['tot'], '%s (%s)' % (m['name'], 'entreno' if key == 'train' else 'descanso')); m['_cal'] = 1
            src = week[cand[0]]
            types[key] = {'meals': [dict({k: v for k, v in m.items() if k != '_cal'}, id=nid(), items=[dict(it, id=nid()) for it in m['items']]) for m in src]}
        if 'train' not in types and 'rest' not in types and len(week) == 7:
            # sin día modelo: el lunes hace de día de entreno y el domingo de descanso (los objetivos de kcal salen del plan)
            for key, di in (('train', 0), ('rest', 6)):
                if week[di]: types[key] = {'meals': [dict(m, id=nid(), items=[dict(it, id=nid()) for it in m['items']]) for m in week[di]]}
            tot = struct.get('train') or struct.get('rest')
            if tot:
                T0 = [sum((v['tot'][j] or 0) for v in tot.values()) for j in range(3)]
                for i, d in enumerate(week):
                    S0 = [sum(it[k] for m in d for it in m['items']) for k in 'pcf']
                    bad = [(nm, a, b) for nm, a, b in zip(('P', 'HC', 'G'), T0, S0) if a and abs(a - b) > max(10, a * .12)]
                    if bad: warn('%s: con la tabla oficial sale %s (tu plan: %s)' % (['Lunes','Martes','Miércoles','Jueves','Viernes','Sábado','Domingo'][i], ' · '.join('%s %g' % (nm, round(b)) for nm, a, b in bad), ' · '.join('%s %g' % (nm, a) for nm, a, b in bad)))
        for k, ty in types.items():
            K = sum(4 * i['p'] + 4 * i['c'] + 9 * i['f'] for m in ty['meals'] for i in m['items'])
            ty['kcal'] = round(K); ty['p'] = round(sum(i['p'] for m in ty['meals'] for i in m['items']))
        if 'train' in types and out.get('_kt'): types['train']['kcal'] = int(out['_kt'])
        if 'rest' in types and out.get('_kr'): types['rest']['kcal'] = int(out['_kr'])
        # días tipo entreno/descanso según la estructura de cada día
        tr = types.get('train', {}).get('meals', [])
        diet['trainDays'] = [marks[i] if marks[i] is not None else (1 if len(week[i]) == len(tr) else 0) for i in range(7)]
        for ty in types.values():
            for m in ty['meals']:
                for it in m['items']: it.pop('_cnt', None); it.pop('_src', None)
        for d in list(week) + [t['meals'] for t in types.values()]:
            for m in d:
                m.pop('_cal', None)
                for it in m['items']: it.pop('_cnt', None); it.pop('_src', None)
    diet['types'] = types
    out['diet'] = diet
    return out

# ---------------------------------------------------------------- entrenamiento
def norm_rest(x):
    s = str(x or '').strip()
    m = re.fullmatch(r"(\d+)\s*[-–]\s*(\d+)\s*'", s)
    if m: return '%d-%d s' % (int(m.group(1)) * 60, int(m.group(2)) * 60)
    m = re.fullmatch(r"(\d+(?:[.,]\d+)?)\s*'", s)
    if m: return '%d s' % round(float(m.group(1).replace(',', '.')) * 60)
    m = re.fullmatch(r"(\d+)\s*(''|\"|s|seg)", s)
    if m: return '%s s' % m.group(1)
    return s
def parse_sets(series, reps, kg):
    n = int(num(series) or 0)
    rs = [x for x in re.split(r'\s*[-–]\s*', str(reps)) if x.strip()] if reps not in (None, '') else []
    ks = [x for x in re.split(r'\s*[-–]\s*', str(kg)) if x.strip()] if kg not in (None, '') else []
    if isinstance(kg, (int, float)): ks = [kg]
    if isinstance(reps, (int, float)): rs = [reps]
    n = max(n, len(rs), len(ks))
    if not n: return [], ''
    note = []
    def rv(x):
        s = str(x).lower()
        if 'min' in s: note.append(str(x)); return int((num(s) or 1) * 60)
        if 'seg' in s or re.search(r'\d\s*s\b', s): note.append(str(x)); return int(num(s) or 0)
        return int(num(s) or 0)
    def kv(x):
        s = str(x)
        if re.search(r'[a-z+]', s, re.I) and not re.fullmatch(r'\s*[\d.,]+\s*(kg)?\s*', s, re.I): note.append(s)
        return num(s) or 0
    R = [rv(rs[i] if i < len(rs) else rs[-1]) if rs else 0 for i in range(n)]
    K = [kv(ks[i] if i < len(ks) else ks[-1]) if ks else 0 for i in range(n)]
    return [[K[i], R[i]] for i in range(n)], ' · '.join(dict.fromkeys(note))

def is_training_v1(wb):
    return any(re.match(r'D[ií]a\s', n) for n in wb.sheetnames) and any(str(ws['D14'].value or '').upper().startswith('EJERCICIO') for ws in wb.worksheets)

def parse_training_v1(wb):
    """Formato antiguo: una hoja por día («Dia A - Fuerza Base») con una sola semana."""
    WD = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
    days, level, week = [], '', None
    for ws in wb.worksheets:
        if str(ws['D14'].value or '').upper()[:9] != 'EJERCICIO': continue
        v = lambda a: ws[a].value
        title = ws.title.strip()
        m = re.match(r'(D[ií]a\s+\w+)\s*[-–]?\s*(.*)$', title)
        dname = re.sub(r'^Dia', 'Día', m.group(1)) if m else title
        ses = str(v('E7') or '')
        wm = re.search(r'\(([^)]+)\)', ses)
        wdn = (wm.group(1) if wm else '').strip().lower().replace('miercoles', 'miércoles').replace('sabado', 'sábado')
        rest = norm_rest(str(v('F9') or '').split('(')[0].strip().replace(' seg', ' s'))
        level = level or str(v('I9') or '').strip(' *')
        exs = []
        for r in range(15, 23):
            name = v('D%d' % r)
            if not name: continue
            n = lambda x: (str(int(x)) if isinstance(x, float) and x.is_integer() else str(x)) if x not in (None, '') else ''
            e = {'id': nid(), 'name': str(name).strip(), 'sets': n(v('E%d' % r)), 'reps': n(v('F%d' % r)), 'kg': n(v('G%d' % r)), 'rest': rest, 'tempo': '', 'tech': ''}
            done = {k: n(v(c + str(r))) for k, c in (('sets', 'H'), ('reps', 'I'), ('kg', 'J'), ('rir', 'K'))}
            if any(done.values()): e['done'] = done  # lo que hizo el cliente (solo para el historial)
            exs.append(e)
        if not exs: continue
        wk = re.search(r'\d+', str(v('G7') or ''))
        if wk: week = int(wk.group(0)) - 1
        blk = lambda a, b: [str(v('B%d' % r)).strip() for r in range(a, b) if v('B%d' % r)]
        note = blk(40, 43)
        cli = blk(35, 38)
        if cli: note += ['Notas del cliente: ' + ' · '.join(cli)]
        days.append({'id': nid(), 'name': dname, 'focus': (m.group(2).strip() if m and m.group(2) else str(v('B9') or '')), 'wd': WD.index(wdn) if wdn in WD else -1,
                     'dur': '60', 'place': 'Gimnasio', 'warm': ' · '.join(blk(25, 28)), 'cool': ' · '.join(blk(30, 33)), 'note': '\n'.join(note), 'exs': exs})
    T = {'meso': {'n': 1, 'level': level, 'start': '', 'weeks': []}, 'auto': True, 'days': days}
    if week is not None and 0 <= week < 4: T['onlyWeek'] = week
    return {'training': T, 'log': [], 'chk': []}

def parse_training(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    if 'Resumen' not in wb.sheetnames and is_training_v1(wb): return parse_training_v1(wb)
    res = wb['Resumen']
    cells = {c.coordinate: c.value for r in res.iter_rows() for c in r if c.value not in (None, '')}
    start = cells.get('C9')
    meso = {'n': int(cells.get('C11') or 1), 'level': cells.get('C8') or '', 'start': start.strftime('%Y-%m-%d') if isinstance(start, dt.datetime) else '', 'weeks': []}
    for row in range(15, 19):
        notes = cells.get('H%d' % row) or ''
        if str(notes).startswith('✔'): notes = ''
        meso['weeks'].append({'phase': cells.get('D%d' % row) or '', 'rir': cells.get('E%d' % row) if cells.get('E%d' % row) is not None else '', 'vol': cells.get('F%d' % row) or '', 'int': cells.get('G%d' % row) or '', 'notes': notes})
    days, logs = [], []
    for ws in wb.worksheets:
        if not ws.title.startswith('Día'): continue
        v = lambda a: ws[a].value
        exs = []
        for r in range(10, 18):
            name = v('D%d' % r)
            if not name: continue
            exs.append({'id': nid(), 'name': str(name).strip(), 'sets': '', 'reps': '', 'kg': '', 'rest': norm_rest(v('G%d' % r)), '_n': v('B%d' % r), 'tempo': str(v('F%d' % r) or ''), 'tech': str(v('H%d' % r) or ''), 'wk': {}})
        if not exs: continue
        byname = {e['name']: e for e in exs}
        warm = [str(v('B%d' % r)) for r in range(20, 23) if v('B%d' % r)]
        cool = [str(v('B%d' % r)) for r in range(89, 92) if v('B%d' % r)]
        note = [str(v('B%d' % r)) for r in range(99, 102) if v('B%d' % r)]
        dname = ws.title.strip()
        for wi, top in enumerate((28, 41, 54, 67)):
            date = v('D%d' % (top - 3))
            for r in range(top, top + 8):
                en = v('C%d' % r)
                if en and str(en).strip() in byname: e = byname[str(en).strip()]
                else:
                    e = next((x for x in exs if x['_n'] == v('B%d' % r)), None)
                    if not e: continue
                series, reps, kgo, rir = v('E%d' % r), v('F%d' % r), v('G%d' % r), v('H%d' % r)
                if wi == 0:
                    e['sets'], e['reps'] = str(series or ''), str(reps or '')
                P = {}
                if series and str(series) != e['sets']: P['sets'] = str(series)
                if kgo not in (None, ''):
                    if isinstance(kgo, str) and re.search(r'\d\s*s\b', kgo):
                        P['why'] = 'Esta semana: %s series de %s.' % (series, kgo)
                    elif isinstance(kgo, str) and kgo.strip().upper() == 'RIR': P['kg'] = 'Según RIR'
                    elif isinstance(kgo, str) and kgo.strip().upper() in ('BW', 'PC'): P['kg'] = 'Peso corporal'
                    else: P['kg'] = kgo if isinstance(kgo, str) else (int(kgo) if float(kgo).is_integer() else kgo)
                if P: e['wk'][str(wi)] = P
                # lo realizado
                ds, dr, dk, dri, dn = v('I%d' % r), v('J%d' % r), v('K%d' % r), v('L%d' % r), v('M%d' % r)
                if dr not in (None, '') or dk not in (None, ''):
                    sets, extra = parse_sets(ds, dr, dk)
                    if not sets: continue
                    if not isinstance(date, dt.datetime) and not meso['start']:
                        warn('%s · %s: hay series hechas pero sin fecha de inicio; no las he pasado a sus registros' % (dname, e['name'])); continue
                    ts = int((date if isinstance(date, dt.datetime) else dt.datetime.strptime(meso['start'], '%Y-%m-%d') + dt.timedelta(days=7 * wi)).replace(hour=19).timestamp() * 1000)
                    logs.append({'id': nid(), 'ts': ts, 'day': dname + ' · S%d' % (wi + 1), 'ex': e['name'], 'sets': sets,
                                 'rir': int(dri) if isinstance(dri, (int, float)) else None,
                                 'note': ' · '.join(x for x in [str(dn).strip() if dn else '', extra] if x), 'w': wi, 'exId': e['id'], 'fresh': False})
        for e in exs:
            e.pop('_n', None)
            if not e['wk']: del e['wk']
        rr = next((r for r in range(22, 30) if str(cells.get('B%d' % r, '')).strip() == dname), 22 + len(days))
        wdn = str(cells.get('C%d' % rr, '') or '').strip().lower()
        WD = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo']
        dur = str(cells.get('G%d' % rr, '') or '')
        days.append({'id': nid(), 'name': dname, 'focus': str(v('F5') or cells.get('D%d' % rr, '') or ''), 'wd': -1,
                     'dur': (re.search(r'\d+\s*[-–]\s*\d+', dur).group(0).replace(' ', '') if re.search(r'\d+\s*[-–]\s*\d+', dur) else (re.findall(r'\d+', dur) or ['60'])[-1]),
                     'place': str(cells.get('H%d' % rr, '') or 'Gimnasio'), 'warm': ' · '.join(warm), 'cool': ' · '.join(cool), 'note': '\n'.join(note), 'exs': exs})
    chk = []
    if 'Check-in' in wb.sheetnames or any(s.lower().startswith('check') for s in wb.sheetnames):
        ws = wb[[s for s in wb.sheetnames if s.lower().startswith('check')][0]]
        for r in range(6, 10):
            d, peso = ws['C%d' % r].value, ws['D%d' % r].value
            if isinstance(d, dt.datetime) and peso:
                parts = []
                if ws['E%d' % r].value: parts.append('Pasos/día: %s' % format(int(num(ws['E%d' % r].value)), ',').replace(',', '.'))
                if ws['F%d' % r].value: parts.append('Sueño: %s h' % str(ws['F%d' % r].value).replace('.', ','))
                for col, lab in (('G', 'Energía'), ('H', 'Estrés'), ('I', 'Hambre')):
                    if ws['%s%d' % (col, r)].value not in (None, ''): parts.append('%s: %s/10' % (lab, ws['%s%d' % (col, r)].value))
                if ws['L%d' % r].value: parts.append(str(ws['L%d' % r].value))
                chk.append({'id': nid(), 'ts': int(d.replace(hour=9).timestamp() * 1000), 'week': r - 5, 'peso': float(peso), 'animo': 0, 'comm': ' · '.join(parts), 'fresh': False})
    return {'training': {'meso': meso, 'auto': True, 'days': days}, 'log': sorted(logs, key=lambda l: -l['ts']), 'chk': chk}

# ---------------------------------------------------------------- composición corporal
CC_MAP = [
    (r'^peso', 'peso'), (r'porcentaje grasa|^masa grasa \(%\)|^grasa corporal \(%\)', 'grasa'), (r'porcentaje m[uú]sculo \(%\)|^m[uú]sculo \(%\)|^masa muscular \(%\)', 'musc'),
    (r'^masa [oó]sea \(%\)', 'osea'),
    (r'(porcentaje |poercentaje )?m[uú]sculo esquel[eé]tico \(%\)', 'esq'), (r'porcentaje [oó]seo', 'osea'),
    (r'porcentaje prote[ií]na|^prote[ií]na \(%\)', 'prot'), (r'agua corporal \(%\)|porcentaje agua', 'agua'),
    (r'grasa subcut', 'subc'), (r'grasa visceral', 'visc'), (r'metab[oó]lica basal', 'tmb'), (r'edad metab', 'edadm'),
    (r'^cuello', 'cuello'), (r'^hombros', 'hombros'), (r'b[ií]ceps derecho|brazo derecho', 'bicD'), (r'b[ií]ceps izquierdo|brazo izquierdo', 'bicI'),
    (r'^pecho', 'pecho'), (r'^cintura \(ombligo\)', 'abdomen'), (r'^cintura', 'cintura'), (r'^abdomen', 'abdomen'), (r'^cadera', 'cadera'),
    (r'^mus(c)?lo derecho', 'musD'), (r'^muslo izquierdo', 'musI'), (r'^gemelo derecho', 'gemD'), (r'^gemelo izquierdo', 'gemI'),
]
def cc_key(label):
    l = str(label).strip().lower()
    for rx, k in CC_MAP:
        if re.search(rx, l): return k
    return None

def parse_comp(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    grid = {(c.row, c.column): c.value for r in ws.iter_rows() for c in r if c.value not in (None, '')}
    revs, skipped, alt = [], set(), None
    # bloques: una columna de etiquetas y a su derecha columnas de sesiones con fecha
    label_cols = sorted({col for (row, col), v in grid.items() if isinstance(v, str) and re.match(r'^fecha$', v.strip(), re.I)})
    for lc in label_cols:
        frow = [row for (row, col), v in grid.items() if col == lc and isinstance(v, str) and v.strip().lower() == 'fecha'][0]
        for col in range(lc + 1, lc + 14):
            d = grid.get((frow, col))
            if not isinstance(d, dt.datetime): continue
            v = {}
            for (row, c2), lab in grid.items():
                if c2 != lc or not isinstance(lab, str) or row == frow or lab.strip().isupper(): continue
                val = grid.get((row, col))
                if val in (None, ''): continue
                if re.match(r'^\s*altura', lab, re.I): alt = val; continue
                k = cc_key(lab)
                if k in ('grasa', 'musc', 'esq', 'osea', 'prot', 'agua', 'subc') and isinstance(val, (int, float)) and 0 < val <= 1:
                    val = round(val * 100, 1)  # celda con formato de porcentaje (0,253 = 25,3 %)
                if k and k not in v: v[k] = str(val).replace('.', ',') if isinstance(val, float) else str(val)
                elif not k: skipped.add(lab.strip())
            revs.append({'fecha': d.strftime('%Y-%m-%d'), 'v': v})
    revs.sort(key=lambda r: r['fecha'])
    for s in sorted(skipped):
        if not re.search(r'imc|masa .*\(kg\)|peso corporal sin grasa|masa grasa \(kg\)|nombre|objetivo|clasificaci|sexo|nacimiento', s, re.I):
            warn('Composición: no he importado «%s» (no hay casilla equivalente)' % s)
    w = [num(r['v']['peso'].replace(',', '.')) for r in revs if r['v'].get('peso')]
    return {'comp': {'nac': '', 'sexo': '', 'altura': alt or '', 'notas': '', 'revs': revs}, 'weights': w}

# ---------------------------------------------------------------- ficha inicial
def parse_ficha(path):
    doc = docx.Document(path)
    I = {'parq': {}, 'dia': {}, 'zonas': [], 'suplem': [], 'steps': []}
    PH = 'haga clic o pulse aquí'
    def clean(s):
        s = '\n'.join(x for x in s.split('\n') if PH not in x.lower()).strip()
        return s
    def checked(cells):
        return [re.sub(r'^\d*\s*☒\s*', '', c).strip() for c in cells if '☒' in c]
    rows = [r for t in tables(doc) for r in t]
    def find(label):
        for r in rows:
            if r and r[0].strip().lower().startswith(label.lower()):
                return r
        return None
    def val(label, idx=1):
        r = find(label)
        return clean(r[idx]) if r and len(r) > idx else ''
    def pick(v):
        m = re.search(r'☒\s*([^☐☒]+)', v or '')
        return m.group(1).strip() if m else ('' if '☐' in (v or '') else v)
    def pair(label):
        for r in rows:
            for i, c in enumerate(r[:-1]):
                if c.strip().lower().startswith(label.lower()):
                    return clean(r[i + 1])
        return ''
    def opt(label, mapping=None):
        r = find(label)
        if not r: return ''
        ch = checked(r[1:])
        if not ch: return ''
        if mapping: ch = [mapping.get(x, x) for x in ch]
        return ch
    def block(label):
        for t in tables(doc):
            if t and t[0] and t[0][0].strip().lower().startswith(label.lower()):
                txt = '\n'.join(x for x in t[0][0].split('\n')[1:])
                rest = '\n'.join(r[0] for r in t[1:])
                return clean((txt + '\n' + rest).strip())
        return ''
    I['nombre'] = val('Nombre completo')
    nac = pair('Fecha de nacimiento')
    m = re.match(r'(\d{1,2})/(\d{1,2})/(\d{4})', nac or '')
    if m: I['nac'] = '%s-%02d-%02d' % (m.group(3), int(m.group(2)), int(m.group(1)))
    I['sexo'] = next((re.sub(r'^\d*\s*☒\s*', '', c).strip() for r in rows for i, x in enumerate(r) if x.strip().lower().startswith('sexo') for c in r[i + 1:i + 4] if '☒' in c), '') or pick(pair('Sexo')); I['tel'] = val('Teléfono'); em = pair('Email')
    if '@' in em: I['email'] = em
    I['ocupacion'] = re.sub(r'\s*/\s*', ' / ', val('Ocupación'))
    t = opt('Tipo de trabajo', {'Activo (de pie / movimiento)': 'Activo (de pie / en movimiento)'})
    if t: I['trabajo'] = t[0]
    r = find('Nivel de estrés')
    if r:
        for c in r[1:]:
            mm = re.match(r'^(\d+)\s*☒', c.strip())
            if mm: I['estres'] = mm.group(1)
    I['sueno'] = val('Horas de sueño'); I['descansa'] = pair('¿Descansas bien?')
    h = opt('Hábitos')
    if h: I['habitos'] = h[0]
    al = val('Alcohol')
    I['lesiones'] = block('Lesiones actuales'); I['patologias'] = block('Patologías')
    med = val('Medicación habitual'); I['medicacion'] = '' if 'nombre del medicamento' in med.lower() else med
    sup = val('Suplementación actual'); I['suplActual'] = '' if 'proteína, creatina' in sup.lower() else sup
    a = opt('¿Tienes autorización médica')
    if a: I['autMedica'] = a[0]
    I['obsMed'] = block('Observaciones médicas')
    n = opt('Nivel de experiencia')
    if n: I['nivel'] = n[0]
    I['deportes'] = block('Deportes o actividades')
    an = val('Años de entrenamiento')
    if an and not an.lower().startswith('aproximad'): I['anosPesas'] = an
    rec = val('Récords personales'); I['records'] = '' if rec.lower().startswith('sentadilla, press') else rec
    I['pasos'] = val('Pasos diarios').replace('/', '-')
    s = opt('¿Has tenido entrenador')
    if s: I['asesoria'] = s[0]
    I['asesoriaQue'] = block('Si tuviste asesoría')
    lg = opt('¿Dónde vas a entrenar?', {'Al aire libre': 'Aire libre'})
    if lg: I['lugar'] = lg[0] if len(lg) == 1 else 'Mixto'
    I['material'] = block('Material disponible')
    I['diasSem'] = val('Días de entrenamiento'); I['minSes'] = pair('Minutos por sesión')
    mo = opt('Mejor momento')
    if mo: I['momento'] = mo[0]
    I['limHorario'] = block('Limitaciones de horario')
    I['comidasDia'] = val('Número de comidas'); I['comidasQuiere'] = pair('¿Cuántas te gustaría')
    co = opt('¿Quién cocina', {'Mi pareja / familia': 'Pareja / familia'})
    if co: I['cocina'] = co[0] if len(co) == 1 else 'Mixto'
    tc = val('Tiempo disponible'); I['tCocinar'] = '' if tc.lower().startswith('minutos aprox') else tc
    I['noGusta'] = block('Alimentos que NO'); I['alergias'] = block('Alergias o intolerancias')
    ta = opt('Tipo de alimentación')
    if ta: I['tipoAlim'] = ta[0]
    I['tentacion'] = block('¿Cuál es tu mayor')
    ax = opt('¿Comes por ansiedad')
    if ax: I['ansiedad'] = ax[0]
    I['agua'] = val('Vasos de agua'); I['alcohol'] = pair('¿Consumes alcohol?') or al
    sp = opt('Suplementación nutricional', {'Vitaminas': 'Vitaminas / minerales'})
    if sp: I['suplem'] = sp
    I['obj3'] = block('¿Cuál es tu objetivo principal'); I['objLargo'] = block('¿Cuál es tu objetivo a largo')
    I['costado'] = block('¿Qué es lo que más'); I['porQueAhora'] = block('¿Por qué quieres'); I['dispuesto'] = block('¿Qué estás dispuesto')
    for lab, k in (('Motivación actual', 'motivacion'), ('Compromiso', 'compromiso')):
        r = find(lab)
        if r:
            for c in r[1:]:
                mm = re.match(r'^(\d+)\s*☒', c.strip())
                if mm: I[k] = mm.group(1)
    return {k: v for k, v in I.items() if v not in ('', None)}


# ---------------------------------------------------------------- carrera
WDN = ['LUNES', 'MARTES', 'MIÉRCOLES', 'JUEVES', 'VIERNES', 'SÁBADO', 'DOMINGO']
def run_type(txt, hint=''):
    t = (txt or '').lower(); h = (hint or '').lower()
    if not t.strip() or t.strip() in ('—', '-'): return None
    if 'test' in t: return 'test'
    if t.startswith('fuerza'): return 'fuerza'
    if re.match(r'^descanso(\s|$|\.|,)', t) and not re.search(r'\d+\s*min', t): return 'descanso'
    if re.match(r'^descanso', t): return 'recup'
    if 'larga' in h: return 'larga'
    if re.search(r'z3|z4|ritmo|cuestas|subidas|series', t) and not re.search(r'^rodaje', t): return 'calidad'
    if re.search(r'^(sin correr|caminar|\d+\s*min caminando|\d+-?\d*\s*min caminando)|bici|movilidad|paseo|caminando suave', t): return 'recup'
    return 'suave'

def dmy_to_iso(s, year):
    m = re.search(r'(\d{1,2})/(\d{1,2})(?:/(\d{4}))?', str(s or ''))
    if not m: return None
    return '%s-%02d-%02d' % (m.group(3) or year, int(m.group(2)), int(m.group(1)))

def sheet_rows(ws):
    return [[c.value for c in r] for r in ws.iter_rows()]

def parse_run(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    R = {'race': {}, 'zones': [], 'rules': [], 'legend': [], 'weeks': [], 'logs': []}
    year = dt.date.today().year
    # ---- planificación semanal (cualquier hoja con SEMANA + días de la semana en la cabecera)
    for ws in wb.worksheets:
        rows = sheet_rows(ws)
        for ri, row in enumerate(rows):
            heads = [str(x).strip().upper() if x else '' for x in row]
            cols = {}
            for ci, h in enumerate(heads):
                for wi, wn in enumerate(WDN):
                    if h.startswith(wn): cols[ci] = (wi, h)
            if 'SEMANA' not in heads or len(cols) < 2: continue
            ci_f = next((i for i, h in enumerate(heads) if h.startswith('FECHA')), None)
            ci_p = next((i for i, h in enumerate(heads) if h.startswith('FASE')), None)
            ci_v = next((i for i, h in enumerate(heads) if h.startswith('VOLUMEN')), None)
            ci_k = next((i for i, h in enumerate(heads) if h.startswith('CLAVE')), None)
            ci_s = heads.index('SEMANA')
            R['title'] = next((str(r[1]).strip() for r in rows[:3] if len(r) > 1 and r[1]), '')
            for r in rows[ri + 1:]:
                lab = str(r[ci_s] or '').strip()
                if lab and lab.isupper() and len(lab) < 40 and not re.match(r'^(SEMANA\s*)?S?\d+$', lab): break
                if not re.match(r'^(semana\s*)?s?\d+$', lab, re.I):
                    if lab and len(lab) > 25: R['legend'].append(lab.lstrip('• ').strip())
                    if not lab and any(r): continue
                    if not any(r): continue
                    continue
                days = [None] * 7
                for ci, (wi, h) in cols.items():
                    txt = str(r[ci] or '').strip()
                    tp = run_type(txt, h)
                    if tp: days[wi] = {'t': tp, 'txt': txt}
                W = {'dates': str(r[ci_f] or '') if ci_f is not None else '', 'phase': str(r[ci_p] or '') if ci_p is not None else '',
                     'vol': str(r[ci_v] or '').replace('≈', '').strip() if ci_v is not None else '', 'key': str(r[ci_k] or '') if ci_k is not None else '', 'days': days}
                R['weeks'].append(W)
            if R['weeks']:
                R['start'] = dmy_to_iso(R['weeks'][0]['dates'], year)
                break
        if R['weeks']: break
    # ---- resumen: carrera, zonas, reglas, análisis
    for ws in wb.worksheets:
        rows = sheet_rows(ws)
        kv = {}
        for r in rows:
            for i in range(len(r) - 1):
                if isinstance(r[i], str) and r[i + 1] not in (None, ''): kv.setdefault(r[i].strip(), r[i + 1])
        if 'Carrera' in kv or 'Fecha' in kv and 'Distancia / desnivel' in kv:
            rc = R['race']
            rc['name'] = str(kv.get('Carrera', '')).split('·')[0].strip()
            rc['place'] = str(kv.get('Carrera', '')).split('·')[1].strip() if '·' in str(kv.get('Carrera', '')) else ''
            d = kv.get('Fecha'); rc['date'] = d.strftime('%Y-%m-%d') if isinstance(d, dt.datetime) else (dmy_to_iso(d, year) or '')
            rc['dist'] = str(kv.get('Distancia / desnivel', '')); rc['mod'] = str(kv.get('Modalidad', ''))
            rc['est'] = str(kv.get('Tiempo estimado', '')); rc['salida'] = str(kv.get('Salida', '')); rc['oblig'] = str(kv.get('Obligatorio', ''))
            ini = kv.get('Inicio del plan')
            if ini: R['start'] = ini.strftime('%Y-%m-%d') if isinstance(ini, dt.datetime) else (dmy_to_iso(ini, year) or R.get('start'))
        for ri, r in enumerate(rows):
            vals = [str(x).strip() if x is not None else '' for x in r]
            nz = [v for v in vals if v]
            if nz[:2] == ['ZONA', 'PULSO (ppm)']:
                for r2 in rows[ri + 1:]:
                    v2 = [str(x).strip() for x in r2 if x not in (None, '')]
                    if not v2 or not re.match(r'^Z\d', v2[0]):
                        if v2 and len(v2[0]) > 40: R['zonesNote'] = v2[0]
                        break
                    R['zones'].append((v2 + ['', '', '', ''])[:4])
            if nz[:3] == ['FECHA', 'SESIÓN', 'DISTANCIA']:
                R['hist'] = []
                for r2 in rows[ri + 1:]:
                    v2 = [str(x).strip() for x in r2 if x not in (None, '')]
                    if not v2: break
                    if v2[0].lower().startswith('conclusi'): R['concl'] = v2[0]; break
                    R['hist'].append((v2 + [''] * 6)[:6])
            if nz and re.search(r'REGLAS', nz[0]):
                for r2 in rows[ri + 1:]:
                    v2 = [str(x).strip() for x in r2 if x not in (None, '')]
                    if not v2: break
                    if re.match(r'^\d+$', v2[0]) and len(v2) > 1: R['rules'].append(v2[1])
                    elif v2[0].startswith('•'): R['rules'].append(v2[0].lstrip('• ').strip())
                    elif v2[0].isupper(): break
                    else: R['rules'].append(v2[0])
            if nz and nz[0].upper().startswith('SEMANA TIPO'): R['legend'].append(nz[0])
    # ---- estrategia de carrera
    for ws in wb.worksheets:
        rows = sheet_rows(ws)
        S = {}
        for ri, r in enumerate(rows):
            vals = [str(x).strip() if x is not None else '' for x in r]
            nz = [v for v in vals[:10] if v]
            if nz[:2] == ['KM', 'TRAMO']:
                S['segs'] = []
                for r2 in rows[ri + 1:]:
                    v2 = [str(x).strip() for x in r2[:10] if x not in (None, '')]
                    if not v2: break
                    if not re.match(r'^[\d,]+\s*[–-]\s*[\d,]+$', v2[0]): S['total'] = v2[0]; break
                    S['segs'].append((v2 + [''] * 8)[:8])
            if nz[:2] == ['MOMENTO', 'QUÉ TOMAR']:
                S['food'] = []
                for r2 in rows[ri + 1:]:
                    v2 = [str(x).strip() for x in r2[:10] if x not in (None, '')]
                    if not v2: break
                    if len(v2) == 1: S['foodNote'] = v2[0]; break
                    S['food'].append((v2 + ['', ''])[:3])
            if nz[:2] == ['KM', 'CONTROL']:
                S['controls'] = []
                for r2 in rows[ri + 1:]:
                    v2 = [str(x).strip() for x in r2[:10] if x not in (None, '')]
                    if not v2: break
                    if len(v2) == 1: S['controlsNote'] = v2[0]; break
                    S['controls'].append((v2 + [''] * 5)[:5])
            if nz and 'DÍAS PREVIOS' in nz[0]:
                S['prev'] = []
                for r2 in rows[ri + 1:]:
                    v2 = [str(x).strip() for x in r2[:10] if x not in (None, '')]
                    if len(v2) < 2: break
                    S['prev'].append(v2[:2])
        if S.get('segs'): R['strat'] = S; break
    R['legend'] = [x for x in dict.fromkeys(R['legend']) if x not in R['rules']]
    # el día de la carrera
    rd = R['race'].get('date')
    if rd and R.get('start'):
        k = (dt.date.fromisoformat(rd) - dt.date.fromisoformat(R['start'])).days
        if 0 <= k < 7 * len(R['weeks']):
            s = R['weeks'][k // 7]['days'][k % 7]
            if s: s['t'] = 'carrera'
    return R if R['weeks'] else None

# ---------------------------------------------------------------- principal
def build(plan=None, entreno=None, carrera=None, comp=None, ficha=None, extra=None):
    """Devuelve (archivo de importación, avisos). Lo usan la línea de comandos y la app (en el navegador)."""
    del REPORT[:]; CHECKED.clear()
    data, name = {}, ''
    if plan:
        p = parse_plan(plan); name = p.pop('name', '') or name
        for k in list(p):
            if k.startswith('_'): p.pop(k)
        data.update(p)
    if entreno: data.update(parse_training(entreno))
    for src in [x for x in (carrera, entreno) if x]:
        rr = parse_run(src)
        if rr: data['run'] = rr; break
    if comp:
        cc = parse_comp(comp); data['comp'] = cc['comp']; data['weights'] = cc['weights']
    if data.get('chk'):
        data['weights'] = (data.get('weights') or []) + [c['peso'] for c in sorted(data['chk'], key=lambda c: c['ts'])]
    if ficha:
        I = parse_ficha(ficha); data['int'] = I; name = name or I.get('nombre', '')
        if data.get('comp'):
            data['comp']['nac'] = I.get('nac', ''); data['comp']['sexo'] = I.get('sexo', '')
    if extra:
        for k, v in extra.items():
            if isinstance(v, dict) and isinstance(data.get(k), dict): data[k].update(v)
            else: data[k] = v
    if name: data['name'] = name
    out = {'tbImport': 1, 'version': 1, 'created': dt.datetime.now().isoformat(timespec='seconds'), 'name': name, 'data': data}
    return out, list(dict.fromkeys(REPORT))

def classify(path):
    """Adivina qué es cada archivo: plan, ficha, entreno, carrera o composición."""
    low = path.lower()
    if low.endswith('.docx'):
        T = tables(docx.Document(path))
        heads = {(t[0][0] if t and t[0] else '').strip().upper() for t in T}
        txt = ' '.join(c for t in T for r in t for c in r).lower()
        meal_tbl = any(t and len(t[0]) > 2 and 'DESAYUNO' in ' '.join(t[0]).upper() for t in T)
        if ('fecha de nacimiento' in txt or 'par-q' in txt) and len(re.findall(r'\d\s*(?:g|ml)\b', txt)) < 15: return 'ficha'
        if 'TOMA' in heads or meal_tbl or 'MACRONUTRIENTE' in heads: return 'plan'
        if 'nivel de estrés' in txt: return 'ficha'
        return 'plan'
    if low.endswith('.xlsx') or low.endswith('.xlsm'):
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        names = wb.sheetnames
        if 'Resumen' in names and any(n.startswith('Día') for n in names): return 'entreno'
        wb.close()
        if is_training_v1(openpyxl.load_workbook(path, data_only=True)): return 'entreno'
        try:
            if parse_run(path): return 'carrera'
        except Exception: pass
        return 'comp'
    return None

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('--plan'); ap.add_argument('--entreno'); ap.add_argument('--carrera', help='Excel con el plan de carrera (o el mismo del entreno si lleva la hoja «Plan carrera»)'); ap.add_argument('--comp'); ap.add_argument('--ficha')
    ap.add_argument('--extra', help='JSON con campos que añadir o corregir (p. ej. restricciones del entrenador)')
    ap.add_argument('-o', '--out', required=True)
    a = ap.parse_args()
    out, rep = build(a.plan, a.entreno, a.carrera, a.comp, a.ficha, json.load(open(a.extra, encoding='utf-8')) if a.extra else None)
    json.dump(out, open(a.out, 'w', encoding='utf-8'), ensure_ascii=False, indent=1)
    print('OK →', a.out)
    for r in rep: print(' ·', r)

if __name__ == '__main__':
    main()
