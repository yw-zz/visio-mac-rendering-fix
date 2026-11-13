"""Turn an SVG picture from an Office file into a one-page vector PDF with correct fonts.

Office stores SVG pictures with a PNG fallback. Word for Mac's PDF export rasterises them
(about 200 ppi), so tables and diagrams saved as SVG come out blurry. This module renders the
SVG back to vectors:

  * shapes (rect, line, path, ...) are rendered by MuPDF's SVG engine;
  * text is drawn by this module with the real installed fonts (MuPDF's SVG engine only has
    the 14 standard PDF fonts, so e.g. Arial Narrow would come out as a wider substitute).
    Text stays real, selectable text with the font embedded.

Supported text features: x / y, text-anchor, font-family lists (per-character fallback),
font-weight, font-style, font-size, fill, tspan with dx / dy / x / y and font changes,
transform on text and its ancestors. Whitespace follows SVG's default xml:space handling.
"""
import math
import re
import xml.etree.ElementTree as ET

try:
    import pymupdf as fitz
except ImportError:            # pragma: no cover
    import fitz

SVG_NS = 'http://www.w3.org/2000/svg'
IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
FALLBACK_FAMILIES = ('Arial', 'Helvetica', 'Arial Unicode MS', 'Times New Roman')


def _tag(el):
    return el.tag.split('}')[-1]

def _mul(a, b):
    """SVG matrices (a b c d e f): apply b first, then a."""
    return (a[0] * b[0] + a[2] * b[1], a[1] * b[0] + a[3] * b[1],
            a[0] * b[2] + a[2] * b[3], a[1] * b[2] + a[3] * b[3],
            a[0] * b[4] + a[2] * b[5] + a[4], a[1] * b[4] + a[3] * b[5] + a[5])

def _apply(m, x, y):
    return (m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5])

def _parse_transform(s):
    m = IDENT
    for name, args in re.findall(r'(\w+)\s*\(([^)]*)\)', s or ''):
        v = [float(t) for t in re.split(r'[\s,]+', args.strip()) if t]
        if name == 'matrix' and len(v) == 6:
            t = tuple(v)
        elif name == 'translate':
            t = (1, 0, 0, 1, v[0], v[1] if len(v) > 1 else 0)
        elif name == 'scale':
            t = (v[0], 0, 0, v[1] if len(v) > 1 else v[0], 0, 0)
        elif name == 'rotate':
            r = math.radians(v[0]); c, s_ = math.cos(r), math.sin(r)
            t = (c, s_, -s_, c, 0, 0)
            if len(v) == 3:
                t = _mul(_mul((1, 0, 0, 1, v[1], v[2]), t), (1, 0, 0, 1, -v[1], -v[2]))
        else:
            continue
        m = _mul(m, t)
    return m

def _num(s, default=0.0):
    if s is None:
        return default
    m = re.match(r'\s*([-+]?[\d.]+(?:e[-+]?\d+)?)', s)
    return float(m.group(1)) if m else default

def _style(el, inherited):
    st = dict(inherited)
    for k in ('font-family', 'font-size', 'font-weight', 'font-style', 'fill', 'text-anchor'):
        if el.get(k) is not None:
            st[k] = el.get(k)
    for k, v in re.findall(r'([\w-]+)\s*:\s*([^;]+)', el.get('style', '')):
        if k in ('font-family', 'font-size', 'font-weight', 'font-style', 'fill', 'text-anchor'):
            st[k] = v.strip()
    return st

def _color(s):
    s = (s or '#000000').strip()
    if s == 'none':
        return None
    named = {'black': '#000000', 'white': '#ffffff', 'red': '#ff0000', 'blue': '#0000ff', 'green': '#008000'}
    s = named.get(s.lower(), s)
    if re.fullmatch(r'#[0-9a-fA-F]{3}', s):
        s = '#' + ''.join(c * 2 for c in s[1:])
    if re.fullmatch(r'#[0-9a-fA-F]{6}', s):
        return tuple(int(s[i:i + 2], 16) / 255 for i in (1, 3, 5))
    m = re.fullmatch(r'rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)', s)
    if m:
        return tuple(int(g) / 255 for g in m.groups())
    return (0, 0, 0)


class _Fonts:
    """Pick a font file for each character, following the SVG font-family list."""
    def __init__(self, index):
        self.index, self.cache = index, {}

    def pick(self, families, bold, italic, ch):
        for fam in list(families) + list(FALLBACK_FAMILIES):
            for b, i in ((bold, italic), (bold, False), (False, False)):
                key = (fam.lower(), b, i)
                if key not in self.cache:
                    self.cache[key] = self.index.get(fam, b, i)
                f = self.cache[key]
                if f and ord(ch) in f[2]:
                    return key, f
        return None, None


def _runs(text_el, fonts):
    """Lay out a <text> element: list of (x, y, string, size, font_key, font, color) in its own
    user space (before the text's transform), with text-anchor applied per text chunk."""
    st0 = _style(text_el, {'font-size': '16', 'font-family': 'Arial', 'fill': '#000000'})
    items = []                                   # (dx_before, dy, x_abs, y_abs, text, style)

    def walk(el, st, first):
        st = _style(el, st)
        items.append(('start', el, st))
        if el.text:
            items.append(('text', el.text, st))
        for ch in el:
            if _tag(ch) == 'tspan':
                walk(ch, st, False)
            if ch.tail:
                items.append(('text', ch.tail, st))
    walk(text_el, st0, True)

    # SVG default whitespace handling over the whole element
    pieces, prev_space = [], True
    for kind, *rest in items:
        if kind == 'text':
            t = re.sub(r'[\r\n]', '', rest[0]).replace('\t', ' ')
            t = re.sub(r' +', ' ', t)
            if prev_space:
                t = t.lstrip(' ')
            if t:
                prev_space = t.endswith(' ')
            pieces.append(('text', t, rest[1]))
        else:
            pieces.append((kind, *rest))
    for i in range(len(pieces) - 1, -1, -1):     # strip trailing space of the element
        if pieces[i][0] == 'text' and pieces[i][1]:
            pieces[i] = ('text', pieces[i][1].rstrip(' '), pieces[i][2])
            break

    x = y = 0.0
    chunks, cur = [], []
    for kind, a, b in pieces:
        if kind == 'start':
            el, st = a, b
            nx, ny = el.get('x'), el.get('y')
            if nx is not None or ny is not None:
                if cur:
                    chunks.append(cur); cur = []
                if nx is not None:
                    x = _num(nx.split()[0])
                if ny is not None:
                    y = _num(ny.split()[0])
            x += _num((el.get('dx') or '0').split()[0]) if el.get('dx') else 0
            y += _num((el.get('dy') or '0').split()[0]) if el.get('dy') else 0
            continue
        txt, st = a, b
        size = _num(st.get('font-size'), 16)
        fams = [f.strip().strip('\'"') for f in st.get('font-family', 'Arial').split(',') if f.strip()]
        w = st.get('font-weight', '400')
        bold = w == 'bold' or w == 'bolder' or (w.isdigit() and int(w) >= 600)
        italic = st.get('font-style', 'normal') in ('italic', 'oblique')
        color = _color(st.get('fill'))
        run_txt, run_key, run_font, run_x = '', None, None, x
        for ch in txt:
            key, f = fonts.pick(fams, bold, italic, ch)
            if f is None:
                key, f = fonts.pick(fams, bold, italic, '?')
                ch = '?'
            if key != run_key and run_txt:
                cur.append([run_x, y, run_txt, size, run_key, run_font, color, st.get('text-anchor', 'start')])
                run_txt, run_x = '', x
            if not run_txt:
                run_x = x
            run_key, run_font = key, f
            tt, _, cmap = f
            x += tt['hmtx'][cmap[ord(ch)]][0] * size / tt['head'].unitsPerEm
            run_txt += ch
        if run_txt:
            cur.append([run_x, y, run_txt, size, run_key, run_font, color, st.get('text-anchor', 'start')])
    if cur:
        chunks.append(cur)
    out = []
    for ch in chunks:                            # text-anchor applies per chunk
        anchor = ch[0][7]
        x0 = min(r[0] for r in ch)
        x1 = max(r[0] + sum(r[5][0]['hmtx'][r[5][2][ord(c)]][0] for c in r[2]) * r[3] / r[5][0]['head'].unitsPerEm
                 for r in ch)
        shift = -(x1 - x0) / 2 if anchor == 'middle' else -(x1 - x0) if anchor == 'end' else 0
        for r in ch:
            out.append((r[0] + shift, r[1], r[2], r[3], r[4], r[5], r[6]))
    return out


def _one_char_per_glyph(data):
    """Return (font bytes, {codepoint: canonical codepoint}) where every glyph is reachable from
    exactly one Unicode character (the lowest). PDF text extraction maps glyphs back to
    characters, so e.g. a space glyph also mapped from U+00A0 would copy as a non-breaking space."""
    import io
    from fontTools.ttLib import TTFont
    f = TTFont(io.BytesIO(data))
    best, canon = {}, {}
    full = f.getBestCmap() or {}
    for cp, g in sorted(full.items()):
        best.setdefault(g, cp)
    for cp, g in full.items():
        canon[cp] = best[g]
    for t in f['cmap'].tables:
        if t.isUnicode():
            t.cmap = {cp: g for cp, g in t.cmap.items() if best.get(g) == cp}
    buf = io.BytesIO()
    f.save(buf)
    return buf.getvalue(), canon


def svg_to_pdf(svg_bytes, font_index):
    """Return a one-page PyMuPDF document drawing the SVG as vectors with real fonts."""
    ET.register_namespace('', SVG_NS)
    ET.register_namespace('xlink', 'http://www.w3.org/1999/xlink')
    root = ET.fromstring(svg_bytes)
    vb = [float(v) for v in re.split(r'[\s,]+', root.get('viewBox', '').strip()) if v] or None

    # collect text with accumulated transforms, then remove it from the tree
    texts = []
    def visit(el, m):
        m = _mul(m, _parse_transform(el.get('transform')))
        for ch in list(el):
            if _tag(ch) == 'text':
                texts.append((ch, _mul(m, _parse_transform(ch.get('transform')))))
                el.remove(ch)
            else:
                visit(ch, m)
    visit(root, IDENT)

    shapes = fitz.open(stream=ET.tostring(root), filetype='svg')
    page_rect = shapes[0].rect
    out = fitz.open('pdf', shapes.convert_to_pdf())
    page = out[0]
    if vb and len(vb) == 4:
        sx, sy = page_rect.width / vb[2], page_rect.height / vb[3]
        base = (sx, 0, 0, sy, -vb[0] * sx, -vb[1] * sy)
    else:
        base = IDENT

    fonts = _Fonts(font_index)
    aliases = {}

    canon = {}

    def alias_for(key):
        if key not in aliases:
            aliases[key] = 'VF%d' % len(aliases)
            data, canon[key] = _one_char_per_glyph(font_index.font_bytes(*key))
            page.insert_font(fontname=aliases[key], fontbuffer=data)
        return aliases[key]
    for el, m in texts:
        full = _mul(base, m)
        k = math.sqrt(abs(full[0] * full[3] - full[1] * full[2])) or 1.0
        # morph matrices act in PDF (y-up) space; SVG matrices are y-down: conjugate by diag(1, -1)
        rot = fitz.Matrix(full[0] / k, -full[1] / k, -full[2] / k, full[3] / k, 0, 0)
        for x, y, s, size, key, font, color in _runs(el, fonts):
            if color is None:
                continue
            px, py = _apply(full, x, y)
            p = fitz.Point(px, py)
            upright = abs(rot.a - 1) < 1e-6 and abs(rot.d - 1) < 1e-6 and abs(rot.b) < 1e-6 and abs(rot.c) < 1e-6
            name = alias_for(key)
            s = ''.join(chr(canon[key].get(ord(c), ord(c))) for c in s)
            page.insert_text(p, s, fontsize=size * k, fontname=name, color=color,
                             morph=None if upright else (p, rot))
    return out
