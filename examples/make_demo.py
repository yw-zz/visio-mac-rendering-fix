"""Build a Visio-style figure as a dual EMF+ file, and demo.docx / demo.pptx that embed it.

The figure is a small sensor signal chain with a measured oscilloscope capture, written the way
Visio writes figures, so it reproduces the problems this project fixes:

  * labels in Arial Narrow, drawn as GDI text inside EMF+ GetDC sections
      -> Office for Mac shows them in italic on screen;
  * a translucent "signal path" highlight: an EMF+ alpha fill, plus a GDI fallback that fakes
    transparency with the PATINVERT / MASKPEN stipple trick
      -> Office for Mac shows a solid block on screen, hiding what is behind it;
  * a 960x540 scope capture stored losslessly (EMF+ PNG object + GDI DIB fallback)
      -> Word's PDF export resamples it to ~200 ppi (visio-pdf-hires restores it).

Usage: python3 make_demo.py [outdir]
Needs Pillow; python-docx and python-pptx for the Office files.
"""
import io
import math
import os
import random
import re
import struct
import sys
import zipfile

W, H = 1000, 300          # figure size in device pixels (96 dpi)
FONT = 'Arial Narrow'


# ------------------------------------------------------------------ record helpers
def rec(t, payload=b''):
    return struct.pack('<II', t, 8 + len(payload)) + payload

def plus(t, flags, data=b''):
    return struct.pack('<HHII', t, flags, 12 + len(data), len(data)) + data

def plus_comment(*records):
    body = b'EMF+' + b''.join(records)
    return rec(70, struct.pack('<I', len(body)) + body)

def argb(a, r, g, b):
    return (a << 24) | (r << 16) | (g << 8) | b

def cref(r, g, b):
    return r | (g << 8) | (b << 16)

NULL_PEN, BLACK_PEN, WHITE_BRUSH, DEFAULT_FONT = 0x80000008, 0x80000007, 0x80000000, 0x8000000D


# ------------------------------------------------------------------ shapes (EMF+ and GDI fallback)
def _solid(r, g, b):
    """Select a brush and a same-colour pen (Office for Mac ignores NULL_PEN on RECTANGLE)."""
    return [rec(39, struct.pack('<IIII', 1, 0, cref(r, g, b), 0)), rec(37, struct.pack('<I', 1)),
            rec(38, struct.pack('<IIiiI', 4, 0, 1, 0, cref(r, g, b))), rec(37, struct.pack('<I', 4))]

def _unsolid():
    return [rec(37, struct.pack('<I', WHITE_BRUSH)), rec(37, struct.pack('<I', BLACK_PEN)),
            rec(40, struct.pack('<I', 1)), rec(40, struct.pack('<I', 4))]

def rects(color, rs):
    """Solid rectangles: EMF+ FillRects + GDI RECTANGLE fallback."""
    r, g, b = color
    out = [plus_comment(plus(0x400A, 0x8000, struct.pack('<II', argb(255, r, g, b), len(rs))
                             + b''.join(struct.pack('<4f', *q) for q in rs))),
           *_solid(r, g, b)]
    for x, y, w, h in rs:
        out.append(rec(43, struct.pack('<4i', round(x), round(y), round(x + w), round(y + h))))
    return out + _unsolid()

def box(x, y, w, h, t=2, color=(0, 0, 0)):
    return rects(color, [(x, y, w, t), (x, y + h - t, w, t), (x, y, t, h), (x + w - t, y, t, h)])

def polygon(color, pts):
    """Filled polygon: EMF+ FillPolygon + GDI POLYGON16."""
    r, g, b = color
    out = [plus_comment(plus(0x400C, 0x8000, struct.pack('<II', argb(255, r, g, b), len(pts))
                             + b''.join(struct.pack('<ff', *p) for p in pts))),
           *_solid(r, g, b)]
    xs, ys = [round(p[0]) for p in pts], [round(p[1]) for p in pts]
    out.append(rec(86, struct.pack('<4iI', min(xs), min(ys), max(xs), max(ys), len(pts))
                   + b''.join(struct.pack('<hh', x, y) for x, y in zip(xs, ys))))
    return out + _unsolid()

def arrow(x0, y, x1, color=(0, 0, 0)):
    return rects(color, [(x0, y - 1, x1 - x0 - 9, 2)]) + polygon(color, [(x1 - 10, y - 5), (x1, y), (x1 - 10, y + 5)])

def translucent(x, y, w, h, color, alpha):
    """Translucent rectangle: EMF+ alpha fill + Visio's GDI stipple fallback (XOR, MASKPEN, XOR)."""
    r, g, b = color
    plus_fill = plus_comment(plus(0x400A, 0x8000, struct.pack('<II', argb(alpha, r, g, b), 1)
                                  + struct.pack('<4f', x, y, w, h)))
    hdr = struct.pack('<IiiHHIIiiII', 40, 8, 8, 1, 1, 0, 32, 0, 0, 2, 0) + b'\x00\x00\x00\x00\xff\xff\xff\x00'
    bits = b''.join(bytes([0xAA if i % 2 == 0 else 0x55, 0, 0, 0]) for i in range(8))
    pat = struct.pack('<IIIIII', 2, 0, 32, len(hdr), 32 + len(hdr), len(bits)) + hdr + bits
    X, Y, X2, Y2 = round(x), round(y), round(x + w), round(y + h)
    blt = lambda: rec(76, struct.pack('<4i', X, Y, X2, Y2) + struct.pack('<iiiiI', X, Y, X2 - X, Y2 - Y, 0x5A0049)
                      + struct.pack('<ii6fIIIIII', 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0))
    path = [rec(59), rec(27, struct.pack('<ii', X, Y)), rec(54, struct.pack('<ii', X2, Y)),
            rec(54, struct.pack('<ii', X2, Y2)), rec(54, struct.pack('<ii', X, Y2)), rec(61), rec(60)]
    return [plus_fill,
            rec(39, struct.pack('<IIII', 1, 0, cref(r, g, b), 0)), rec(37, struct.pack('<I', 1)), blt(),
            rec(20, struct.pack('<I', 9)), rec(24, struct.pack('<I', 0xFFFFFF)),
            rec(94, pat), rec(37, struct.pack('<I', 2)), rec(37, struct.pack('<I', NULL_PEN)),
            *path, rec(62, struct.pack('<4i', 0, 0, -1, -1)),
            rec(37, struct.pack('<I', 1)), rec(20, struct.pack('<I', 13)), blt(),
            rec(37, struct.pack('<I', WHITE_BRUSH)), rec(40, struct.pack('<I', 2)), rec(40, struct.pack('<I', 1))]


# ------------------------------------------------------------------ text (GDI inside GetDC, like GDI+)
_metrics = None
def _advances(s, size, bold):
    """Per-character advances in pixels, from the installed Arial Narrow if available."""
    global _metrics
    if _metrics is None:
        _metrics = {}
        try:
            sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))
            from visio_mac_rendering_fix.fonts import FontIndex
            idx = FontIndex()
            for bd in (False, True):
                f = idx.get(FONT, bd, False)
                if f:
                    _metrics[bd] = (f[0]['hmtx'], f[2], f[0]['head'].unitsPerEm)
        except Exception:
            pass
    m = _metrics.get(bold)
    if not m:
        return [round(size * (0.23 if c == ' ' else 0.41)) for c in s]
    hmtx, cmap, upm = m
    return [round(hmtx[cmap.get(ord(c), cmap[ord('?')])][0] * size / upm) for c in s]

def text(x, y, s, size=20, bold=False, color=(0, 0, 0), align='left'):
    """Baseline at y. align: left / center (x is the centre)."""
    adv = _advances(s, size, bold)
    if align == 'center':
        x -= sum(adv) / 2
    lf = struct.pack('<5i8B', -size, 0, 0, 0, 700 if bold else 400, 0, 0, 0, 0, 0, 0, 5, 34)
    lf += FONT.encode('utf-16le').ljust(64, b'\x00')
    elw = lf + b'\x00' * (320 - len(lf))
    n = len(s)
    txt = s.encode('utf-16le') + (b'\x00\x00' if n % 2 else b'')
    off_str = 8 + 68
    off_dx = off_str + len(txt)
    emrtext = struct.pack('<iiIIIiiiiI', 0, 0, n, off_str, 0, 0, 0, -1, -1, off_dx)
    body = struct.pack('<iiiiIff', 0, 0, -1, -1, 2, 0.0, 0.0) + emrtext + txt + struct.pack('<%di' % n, *adv)
    return [plus_comment(plus(0x4004, 0)),                                   # GetDC
            rec(33), rec(36, struct.pack('<6fI', 1, 0, 0, 1, x, y, 2)),
            rec(82, struct.pack('<I', 3) + elw), rec(37, struct.pack('<I', 3)),
            rec(22, struct.pack('<I', 24)), rec(18, struct.pack('<I', 1)),
            rec(24, struct.pack('<I', cref(*color))), rec(84, body),
            rec(37, struct.pack('<I', DEFAULT_FONT)), rec(40, struct.pack('<I', 3)),
            rec(34, struct.pack('<i', -1))]


# ------------------------------------------------------------------ embedded bitmap
TRACES = (('Input', 75, (110, 110, 110)), ('Clock', 205, (40, 170, 190)),
          ('Data', 320, (220, 40, 90)), ('RF out', 450, (120, 170, 40)))

def scope_capture(w=960, h=540):
    """A synthetic oscilloscope screenshot: thin coloured traces on a light grid."""
    from PIL import Image, ImageDraw
    im = Image.new('RGB', (w, h), (255, 255, 255))
    d = ImageDraw.Draw(im)
    for i in range(1, 10):
        d.line([(i * w / 10, 0), (i * w / 10, h)], fill=(215, 215, 215))
    for i in range(1, 8):
        d.line([(0, i * h / 8), (w, i * h / 8)], fill=(215, 215, 215))
    rnd = random.Random(1)
    bits = [rnd.randint(0, 1) for _ in range(12)]
    bit = lambda t: bits[min(11, int(t * 12))]
    shapes = (lambda t: math.sin(2 * math.pi * 6 * t),                          # analog input
              lambda t: 1 if math.sin(2 * math.pi * 24 * t) > 0 else -1,       # clock
              lambda t: 1 if bit(t) else -1,                                   # data
              lambda t: (0.25 + 0.75 * bit(t)) * math.sin(2 * math.pi * 110 * t))  # RF output
    for (_, y0, color), fn, amp in zip(TRACES, shapes, (40, 28, 28, 55)):
        pts = [(x, y0 - amp * fn(x / w) + rnd.uniform(-1.2, 1.2)) for x in range(w)]
        d.line(pts, fill=color, width=1)
    return im

def image(x, y, w, h, im, obj_id=1, attr_id=2):
    """EMF+ PNG image object (chunked like GDI+) + DrawImagePoints, and a GDI STRETCHDIBITS fallback."""
    png = io.BytesIO(); im.save(png, 'PNG'); png = png.getvalue()
    bitmap = struct.pack('<IIiiiII', 0xDBC01002, 1, im.width, im.height, 0, 0x0026200A, 1) + png
    out = [plus_comment(plus(0x4008, (8 << 8) | attr_id,                      # ImageAttributes object
                             struct.pack('<6I', 0xDBC01002, 0, 0, 0, 0, 0)))]
    chunk = 65000                                                            # multiple of 4
    total = len(bitmap)                                                      # TotalObjectSize: true length
    bitmap += b'\x00' * (-total % 4)                                         # EMF+ records are 4-byte aligned
    for i in range(0, len(bitmap), chunk):                                   # continued Object records
        out.append(plus_comment(plus(0x4008, 0x8000 | (5 << 8) | obj_id,
                                     struct.pack('<I', total) + bitmap[i:i + chunk])))
    out.append(plus_comment(plus(0x401B, obj_id, struct.pack('<II4fI', attr_id, 2, -0.5, -0.5, im.width, im.height, 3)
                                 + struct.pack('<6f', x, y, x + w, y, x, y + h))))
    # GDI fallback: 24-bit bottom-up DIB
    stride = (im.width * 3 + 3) & ~3
    rows = im.convert('RGB').transpose(1).tobytes('raw', 'BGR')             # FLIP_TOP_BOTTOM: DIB is bottom-up
    bits = b''.join(rows[r * im.width * 3:(r + 1) * im.width * 3].ljust(stride, b'\x00') for r in range(im.height))
    bmi = struct.pack('<IiiHHIIiiII', 40, im.width, im.height, 1, 24, 0, len(bits), 0, 0, 0, 0)
    X, Y, X2, Y2 = round(x), round(y), round(x + w), round(y + h)
    fixed = struct.pack('<4i6iIIIIIIii', X, Y, X2, Y2, X, Y, 0, 0, im.width, im.height,
                        80, len(bmi), 80 + len(bmi), len(bits), 0, 0x00CC0020, X2 - X, Y2 - Y)
    out.append(rec(81, fixed + bmi + bits))
    return out


# ------------------------------------------------------------------ the figure
def figure_records():
    r = []
    blue, orange, grey = (40, 120, 220), (230, 120, 30), (90, 90, 90)
    r += text(20, 28, 'Sensor signal chain', 22, bold=True)
    # signal-path highlight behind the chain
    r += translucent(14, 60, 568, 132, blue, 70)
    r += text(400, 184, 'Signal path', 18, color=blue)
    blocks = [('Sensor', 'Bridge, 2 kΩ'), ('Amplifier', 'Gain 40 dB'), ('ADC', '12-bit, 1 MS/s'), ('Radio', '2.4 GHz')]
    bx, by, bw, bh, gap = 30, 80, 112, 64, 34
    for i, (name, spec) in enumerate(blocks):
        x = bx + i * (bw + gap)
        r += rects((255, 255, 255), [(x, by, bw, bh)])
        r += box(x, by, bw, bh)
        r += text(x + bw / 2, by + 30, name, 20, bold=True, align='center')
        r += text(x + bw / 2, by + 52, spec, 16, color=grey, align='center')
        if i < len(blocks) - 1:
            r += arrow(x + bw + 2, by + bh / 2, x + bw + gap - 2)
    # supply rail
    r += rects(orange, [(bx + bw / 2 - 1, 235, 3 * (bw + gap) + 3, 3)])
    for i in range(4):
        cx = bx + i * (bw + gap) + bw / 2
        r += rects(orange, [(cx - 1, by + bh + 2, 3, 236 - by - bh - 2)])
    r += text(bx + bw / 2 + 6, 262, 'Supply rail 1.8 V', 18, color=orange)
    # scope capture
    sx, sy, sw, sh = 668, 46, 316, 178
    r += text(sx + sw / 2, 32, 'Measured waveforms', 22, bold=True, align='center')
    r += image(sx, sy, sw, sh, scope_capture())
    r += box(sx - 2, sy - 2, sw + 4, sh + 4)
    for label, y_px, col in TRACES:
        r += text(sx - 8 - sum(_advances(label, 16, False)), sy + sh * y_px / 540 + 6, label, 16, color=col)
    r += text(sx + sw / 2, sy + sh + 24, '1 us/div', 16, color=grey, align='center')
    return r


def make_emf():
    body = [plus_comment(plus(0x4001, 1, struct.pack('<IIII', 0xDBC01002, 1, 96, 96)),     # Header (dual)
                         plus(0x4030, 2, struct.pack('<f', 1.0)))]                          # page unit: pixel
    body += figure_records()
    body.append(plus_comment(plus(0x4002, 0)))
    body.append(rec(14, struct.pack('<III', 0, 16, 20)))
    content = b''.join(body)
    frame = (0, 0, round(W * 2540 / 96), round(H * 2540 / 96))
    header = struct.pack('<II4i4iIIIIHHIIIiiii', 1, 108, 0, 0, W - 1, H - 1, *frame,
                         0x464D4520, 0x10000, 108 + len(content), 0, 5, 0, 0, 0, 0, 1920, 1080, 508, 286)
    header += struct.pack('<III', 0, 0, 0) + struct.pack('<ii', 508000, 285750)
    data = bytearray(header[:108] + content)
    off = n = 0
    while off < len(data):
        t, sz = struct.unpack_from('<II', data, off)
        n += 1
        if t == 14:
            break
        off += sz
    struct.pack_into('<I', data, 52, n)
    return bytes(data)


# ------------------------------------------------------------------ Office files
def _placeholder_png():
    from PIL import Image
    buf = io.BytesIO(); Image.new('RGB', (W, H), 'white').save(buf, 'PNG'); return buf.getvalue()

def _swap_media_to_emf(path, emf):
    tmp = path + '.tmp'
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            d, name = zi.read(it.filename), it.filename
            if '/media/' in name and name.endswith('.png'):
                name, d = name[:-4] + '.emf', emf
            if name.endswith('.rels'):
                d = d.replace(b'.png"', b'.emf"')
            if name == '[Content_Types].xml' and b'Extension="emf"' not in d:
                d = re.sub(rb'(<Types[^>]*>)', rb'\1<Default Extension="emf" ContentType="image/x-emf"/>', d, count=1)
            zi_out = zipfile.ZipInfo(name, it.date_time)
            zi_out.compress_type = zipfile.ZIP_DEFLATED
            zo.writestr(zi_out, d)
    os.replace(tmp, path)

def main(outdir='.'):
    os.makedirs(outdir, exist_ok=True)
    emf = make_emf()
    with open(os.path.join(outdir, 'demo.emf'), 'wb') as f:
        f.write(emf)
    try:
        from docx import Document
        from docx.shared import Inches
        from pptx import Presentation
        from pptx.util import Inches as PI
    except ImportError:
        print('python-docx / python-pptx not installed: wrote demo.emf only')
        return
    d = Document()
    d.add_paragraph('A Visio-style figure (dual EMF+). Open this file in Word for Mac, then compare '
                    'demo_macfix.docx; export both to PDF and compare with visio-pdf-hires.')
    d.add_picture(io.BytesIO(_placeholder_png()), width=Inches(6.5))
    p = os.path.join(outdir, 'demo.docx'); d.save(p); _swap_media_to_emf(p, emf)
    prs = Presentation()
    prs.slide_width, prs.slide_height = PI(13.333), PI(7.5)
    s = prs.slides.add_slide(prs.slide_layouts[6])
    s.shapes.add_picture(io.BytesIO(_placeholder_png()), PI(0.4), PI(1.2), width=PI(12.5))
    p = os.path.join(outdir, 'demo.pptx'); prs.save(p); _swap_media_to_emf(p, emf)
    print('wrote demo.emf, demo.docx, demo.pptx to', os.path.abspath(outdir))


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else '.')
