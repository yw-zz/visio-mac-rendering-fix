"""Build a small Visio-style dual EMF+ figure (and .docx / .pptx files that embed it).

The figure reproduces the two Office-for-Mac rendering bugs this tool fixes:
  * labels in Arial Narrow, drawn as GDI text inside EMF+ GetDC sections (Mac shows italic);
  * a translucent shape: EMF+ alpha fill, plus a GDI fallback that fakes transparency with the
    PATINVERT / MASKPEN stipple trick (Mac shows a solid block).

Usage: python3 make_demo.py [outdir]     (needs python-docx and python-pptx for the Office files)
"""
import os
import struct
import sys
import zipfile
import re

W, H = 800, 300          # device pixels at 96 dpi


def rec(t, payload=b''):
    return struct.pack('<II', t, 8 + len(payload)) + payload

def plus(t, flags, data=b''):
    return struct.pack('<HHII', t, flags, 12 + len(data), len(data)) + data

def plus_comment(*records):
    body = b'EMF+' + b''.join(records)
    return rec(70, struct.pack('<I', len(body)) + body)

def argb(a, r, g, b):
    return (a << 24) | (r << 16) | (g << 8) | b

def colorref(r, g, b):
    return r | (g << 8) | (b << 16)

def fill_rects_plus(color, rects):
    return plus(0x400A, 0x8000, struct.pack('<II', color, len(rects))
                + b''.join(struct.pack('<4f', *r) for r in rects))

def gdi_solid_rects(handle, color, rects):
    out = [rec(39, struct.pack('<IIII', handle, 0, color, 0)), rec(37, struct.pack('<I', handle)),
           rec(37, struct.pack('<I', 0x80000008))]                          # NULL_PEN
    for x, y, w, h in rects:
        out.append(rec(43, struct.pack('<4i', int(x), int(y), int(x + w) + 1, int(y + h) + 1)))
    out += [rec(37, struct.pack('<I', 0x80000000)), rec(40, struct.pack('<I', handle))]
    return out

def _advances(s, size, bold):
    """Per-character advance widths (pixels) like GDI writes them; real Arial Narrow metrics if found."""
    try:
        sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'src'))
        from visio_mac_rendering_fix.fonts import FontIndex
        f = FontIndex().get('Arial Narrow', bold, False)
    except Exception:
        f = None
    if f is None:
        return [int(round(size * (0.23 if c == ' ' else 0.41))) for c in s]
    tt, _, cmap = f
    k = size / tt['head'].unitsPerEm
    return [int(round(tt['hmtx'][cmap[ord(c)]][0] * k)) for c in s]


def text_run(handle, x, y, s, size=40, bold=False, color=(0, 0, 0)):
    """GDI text inside a GetDC section, the way GDI+ / Visio writes it."""
    logfont = struct.pack('<5i8B', -size, 0, 0, 0, 700 if bold else 400, 0, 0, 0, 0, 0, 0, 5, 34)
    logfont += 'Arial Narrow'.encode('utf-16le').ljust(64, b'\x00')
    elw = logfont + b'\x00' * (320 - len(logfont))        # LogFontExDv-sized blob (fields unused)
    adv = _advances(s, size, bold)
    n = len(s)
    off_str = 8 + 68
    txt = s.encode('utf-16le') + (b'\x00\x00' if n % 2 else b'')
    off_dx = off_str + len(txt)
    emrtext = struct.pack('<iiIIIiiiiI', 0, 0, n, off_str, 0, 0, 0, -1, -1, off_dx)
    body = struct.pack('<iiiiIff', 0, 0, -1, -1, 2, 0.0, 0.0) + emrtext + txt + struct.pack('<%di' % n, *adv)
    return [plus_comment(plus(0x4004, 0)),                                      # GetDC
            rec(33),                                                             # SAVEDC
            rec(36, struct.pack('<6fI', 1, 0, 0, 1, x, y, 2)),                   # MODIFYWORLDTRANSFORM
            rec(82, struct.pack('<I', handle) + elw),                            # EXTCREATEFONTINDIRECTW
            rec(37, struct.pack('<I', handle)),
            rec(22, struct.pack('<I', 24)),                                      # TA_BASELINE
            rec(18, struct.pack('<I', 1)),                                       # TRANSPARENT
            rec(24, struct.pack('<I', colorref(*color))),
            rec(84, body),                                                       # EXTTEXTOUTW
            rec(37, struct.pack('<I', 0x8000000D)), rec(40, struct.pack('<I', handle)),
            rec(34, struct.pack('<i', -1))]                                      # RESTOREDC

def stipple_rect(x, y, w, h, color):
    """GDI fallback for a translucent rectangle: XOR / MASKPEN checkerboard / XOR."""
    r, g, b = color
    hdr = struct.pack('<IiiHHIIiiII', 40, 8, 8, 1, 1, 0, 32, 0, 0, 2, 0) + b'\x00\x00\x00\x00\xff\xff\xff\x00'
    bits = b''.join(bytes([0xAA if i % 2 == 0 else 0x55, 0, 0, 0]) for i in range(8))
    pat = struct.pack('<IIIIII', 2, 0, 32, len(hdr), 32 + len(hdr), len(bits)) + hdr + bits
    blt = lambda: rec(76, struct.pack('<4i', int(x), int(y), int(x + w), int(y + h)) +
                      struct.pack('<iiiiI', int(x), int(y), int(w), int(h), 0x5A0049) +
                      struct.pack('<ii6fIIIIII', 0, 0, 1, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0))
    path = [rec(59), rec(27, struct.pack('<ii', int(x), int(y))),
            rec(54, struct.pack('<ii', int(x + w), int(y))), rec(54, struct.pack('<ii', int(x + w), int(y + h))),
            rec(54, struct.pack('<ii', int(x), int(y + h))), rec(61), rec(60)]
    return [rec(39, struct.pack('<IIII', 1, 0, colorref(r, g, b), 0)), rec(37, struct.pack('<I', 1)),
            blt(),
            rec(20, struct.pack('<I', 9)),                       # R2_MASKPEN
            rec(24, struct.pack('<I', 0xFFFFFF)),
            rec(94, pat), rec(37, struct.pack('<I', 2)),
            rec(37, struct.pack('<I', 0x80000008)),
            *path, rec(62, struct.pack('<4i', 0, 0, -1, -1)),
            rec(37, struct.pack('<I', 1)),
            rec(20, struct.pack('<I', 13)),
            blt(),
            rec(37, struct.pack('<I', 0x80000000)), rec(40, struct.pack('<I', 2)), rec(40, struct.pack('<I', 1))]


def make_emf():
    body = []
    body.append(plus_comment(
        plus(0x4001, 1, struct.pack('<IIII', 0xDBC01002, 1, 96, 96)),           # Header, dual
        plus(0x4030, 2, struct.pack('<f', 1.0))))                                 # SetPageTransform pixel
    # frame and a black wire (EMF+ fill + GDI fallback)
    frame = [(20, 20, 760, 3), (20, 277, 760, 3), (20, 20, 3, 260), (777, 20, 3, 260), (60, 205, 680, 4)]
    body.append(plus_comment(fill_rects_plus(argb(255, 0, 0, 0), frame)))
    body += gdi_solid_rects(1, 0, frame)
    # translucent blue band crossing the wire
    band = (420, 150, 300, 110)
    body.append(plus_comment(fill_rects_plus(argb(110, 40, 120, 220), [band])))
    body += stipple_rect(*band, (40, 120, 220))
    # labels
    body += text_run(3, 60, 90, 'Arial Narrow label', 44)
    body += text_run(3, 60, 160, 'Bold label 6.78MHz', 44, bold=True, color=(200, 30, 30))
    body.append(plus_comment(plus(0x4002, 0)))                                    # EMF+ EndOfFile
    body.append(rec(14, struct.pack('<III', 0, 16, 20)))                    # EMR_EOF
    content = b''.join(body)
    hdr_len = 108
    frame_01mm = (0, 0, int(W * 2540 / 96), int(H * 2540 / 96))
    header = struct.pack('<II4i4iIIIIHHIIIii ii', 1, hdr_len, 0, 0, W - 1, H - 1, *frame_01mm,
                         0x464D4520, 0x10000, hdr_len + len(content), 0, 4, 0, 0, 0, 0,
                         1920, 1080, 508, 286)
    header += struct.pack('<III', 0, 0, 0) + struct.pack('<ii', 508000, 285750)
    header = header[:hdr_len]
    data = bytearray(header + content)
    # record count
    off, count = 0, 0
    while off < len(data):
        t, sz = struct.unpack_from('<II', data, off)
        count += 1
        if t == 14:
            break
        off += sz
    struct.pack_into('<I', data, 52, count)
    return bytes(data)


def _swap_media_to_emf(path, placeholder_ext):
    """Replace the placeholder picture in an Office file with demo.emf (correct name + content type)."""
    tmp = path + '.tmp'
    emf = make_emf()
    with zipfile.ZipFile(path) as zi, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zo:
        for it in zi.infolist():
            d = zi.read(it.filename)
            name = it.filename
            if '/media/' in name and name.endswith(placeholder_ext):
                name, d = name[:-len(placeholder_ext)] + '.emf', emf
            if name.endswith('.rels'):
                d = d.replace(placeholder_ext.encode() + b'"', b'.emf"')
            if name == '[Content_Types].xml' and b'Extension="emf"' not in d:
                d = re.sub(rb'(<Types[^>]*>)', rb'\1<Default Extension="emf" ContentType="image/x-emf"/>', d, count=1)
            zi_out = zipfile.ZipInfo(name, it.date_time)
            zi_out.compress_type = zipfile.ZIP_DEFLATED
            zo.writestr(zi_out, d)
    os.replace(tmp, path)


def main(outdir='.'):
    os.makedirs(outdir, exist_ok=True)
    emf = make_emf()
    open(os.path.join(outdir, 'demo.emf'), 'wb').write(emf)
    try:
        import io
        from docx import Document
        from docx.shared import Inches
        from pptx import Presentation
        from pptx.util import Inches as PI
        png = _placeholder_png()
        d = Document()
        d.add_paragraph('Visio-style figure (dual EMF+):')
        d.add_picture(io.BytesIO(png), width=Inches(6))
        p = os.path.join(outdir, 'demo.docx'); d.save(p); _swap_media_to_emf(p, '.png')
        prs = Presentation()
        s = prs.slides.add_slide(prs.slide_layouts[6])
        s.shapes.add_picture(io.BytesIO(png), PI(0.5), PI(1.5), width=PI(9))
        p = os.path.join(outdir, 'demo.pptx'); prs.save(p); _swap_media_to_emf(p, '.png')
    except ImportError:
        print('python-docx / python-pptx not installed: wrote demo.emf only')
    print('wrote demo files to', os.path.abspath(outdir))


def _placeholder_png():
    """A W x H white PNG, so Office sizes the picture to the EMF's aspect ratio."""
    import zlib
    raw = b''.join(b'\x00' + b'\xff' * (W * 3) for _ in range(H))
    chunk = lambda t, d: struct.pack('>I', len(d)) + t + d + struct.pack('>I', zlib.crc32(t + d) & 0xFFFFFFFF)
    return (b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', W, H, 8, 2, 0, 0, 0))
            + chunk(b'IDAT', zlib.compress(raw)) + chunk(b'IEND', b''))


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else '.')
