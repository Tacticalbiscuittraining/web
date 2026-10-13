#!/usr/bin/env python3
"""Perfil completo de un cliente en Word (.docx), para guardarlo antes de eliminarlo.

La app prepara el contenido (bloques) y aquí solo se maqueta.
Uso:  python3 tools/perfil.py bloques.json -o perfil.docx
bloques.json = {"title", "sub", "blocks": [...], "imgs": [rutas]}
  ["h", nivel, texto] · ["p", texto] · ["kv", [[etiqueta, valor], …]]
  ["t", [cabecera…], [[fila…], …]] · ["img", índice, pie] · ["pb"]
En la app se ejecuta en el navegador (Pyodide), igual que el importador.
"""
import io, json, sys
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt, RGBColor

GREEN = RGBColor(0x2A, 0x6F, 0x5E)

def shade(cell, hex_):
    tcPr = cell._tc.get_or_add_tcPr()
    s = OxmlElement('w:shd'); s.set(qn('w:val'), 'clear'); s.set(qn('w:color'), 'auto'); s.set(qn('w:fill'), hex_)
    tcPr.append(s)

def txt(v):
    if v is None: return ''
    if isinstance(v, bool): return 'Sí' if v else 'No'
    if isinstance(v, float): return ('%g' % v).replace('.', ',')
    return str(v)

def put(cell, v, bold=False, size=9):
    cell.text = ''
    p = cell.paragraphs[0]
    lines = txt(v).split('\n')
    for i, ln in enumerate(lines):
        r = p.add_run(ln); r.bold = bold; r.font.size = Pt(size)
        if i < len(lines) - 1: r.add_break()

def table(doc, head, rows, widths=None):
    n = max([len(head)] + [len(r) for r in rows]) if (head or rows) else 0
    if not n: return
    t = doc.add_table(rows=0, cols=n)
    t.style = 'Table Grid'; t.alignment = WD_TABLE_ALIGNMENT.CENTER
    if head:
        cs = t.add_row().cells
        for i in range(n):
            put(cs[i], head[i] if i < len(head) else '', True); shade(cs[i], 'E3F2EF')
        trPr = t.rows[0]._tr.get_or_add_trPr(); h = OxmlElement('w:tblHeader'); h.set(qn('w:val'), 'true'); trPr.append(h)
    for r in rows:
        cs = t.add_row().cells
        for i in range(n): put(cs[i], r[i] if i < len(r) else '')
    if widths:
        for row in t.rows:
            for i, w in enumerate(widths[:n]): row.cells[i].width = Cm(w)
    doc.add_paragraph()

def build(d, imgs=None):
    doc = Document()
    sec = doc.sections[0]
    sec.left_margin = sec.right_margin = Cm(1.8); sec.top_margin = sec.bottom_margin = Cm(1.6)
    st = doc.styles['Normal']; st.font.name = 'Calibri'; st.font.size = Pt(10)
    for lv in (1, 2, 3):
        doc.styles['Heading %d' % lv].font.color.rgb = GREEN
    h = doc.add_heading(d.get('title', 'Perfil del cliente'), 0)
    if d.get('sub'): doc.add_paragraph(d['sub']).runs[0].italic = True
    imgs = imgs or []
    for b in d.get('blocks', []):
        k = b[0]
        if k == 'h': doc.add_heading(txt(b[2]), min(3, max(1, int(b[1]))))
        elif k == 'p':
            for ln in txt(b[1]).split('\n\n'): doc.add_paragraph(ln)
        elif k == 'kv':
            rows = [r for r in b[1] if txt(r[1]).strip()]
            if not rows: continue
            t = doc.add_table(rows=0, cols=2); t.style = 'Table Grid'
            for lab, val in rows:
                cs = t.add_row().cells; put(cs[0], lab, True); shade(cs[0], 'F3F6F5'); put(cs[1], val)
                cs[0].width = Cm(6); cs[1].width = Cm(11.4)
            doc.add_paragraph()
        elif k == 't': table(doc, b[1], b[2])
        elif k == 'img':
            i = int(b[1])
            if i < len(imgs) and imgs[i]:
                try:
                    doc.add_picture(io.BytesIO(imgs[i]) if isinstance(imgs[i], (bytes, bytearray)) else imgs[i], width=Cm(8))
                    if len(b) > 2 and b[2]: doc.add_paragraph(txt(b[2])).runs[0].font.size = Pt(9)
                except Exception:
                    doc.add_paragraph('(No se ha podido incluir la foto %s)' % txt(b[2] if len(b) > 2 else ''))
        elif k == 'pb': doc.add_page_break()
    out = io.BytesIO(); doc.save(out)
    return out.getvalue()

if __name__ == '__main__':
    a = sys.argv[1:]
    d = json.load(open(a[0], encoding='utf-8'))
    o = a[a.index('-o') + 1] if '-o' in a else 'perfil.docx'
    open(o, 'wb').write(build(d, [open(p, 'rb').read() for p in d.get('imgs', [])]))
    print(o)
