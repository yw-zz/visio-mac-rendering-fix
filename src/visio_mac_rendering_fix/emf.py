"""Rewrite a dual EMF+ metafile so Office for Mac's on-screen renderer draws it correctly.

Background
  Visio (and other GDI+ apps) produce "dual" EMF files: an EMF+ layer plus a plain GDI fallback
  layer. Office for Windows, and Office for Mac's PDF export, draw the EMF+ layer and run GDI
  records only inside EMF+ "GetDC" sections, which is where Visio puts its text. Office for Mac's
  on-screen view plays only the GDI layer, where
    * text resolves through Office's bundled fonts; e.g. its "Arial Narrow" has only the Italic
      faces, so labels draw italic;
    * transparency is faked with PATINVERT / MASKPEN stipple raster ops that the Mac ignores,
      so translucent shapes become solid blocks.

What fix_emf() does
  1. Each text run whose font is installed becomes glyph outlines, drawn twice:
       - an anti-aliased EMF+ FillPath in device pixels (Windows and PDF export draw this), and
       - a GDI path fill placed outside the GetDC section (only the Mac on-screen view draws it).
     The GetDC section is then reopened so the remaining GDI state records still run on Windows.
  2. Stipple transparency in the GDI fallback becomes a plain fill pre-blended 50% with white.
     Other stipple variants are dropped. EMF+ renderers never draw these records.
"""
import struct
from fontTools.pens.basePen import BasePen

SUB = 16            # GDI path coords are written x16 under a 1/16 world scale (sub-unit precision)
PLUS_OBJ_ID = 63    # EMF+ object slot used for glyph paths (GDI+ writers fill slots from 0)
PATINVERT, PATAND = 0x5A0049, 0xA000C9
IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)

# EMF record types used here
EMR_HEADER, EMR_EOF = 1, 14
EMR_SETWINDOWORGEX, EMR_SETVIEWPORTORGEX = 10, 12
EMR_SETPOLYFILLMODE, EMR_SETROP2, EMR_SETTEXTALIGN, EMR_SETTEXTCOLOR = 19, 20, 22, 24
EMR_MOVETOEX, EMR_SAVEDC, EMR_RESTOREDC = 27, 33, 34
EMR_SETWORLDTRANSFORM, EMR_MODIFYWORLDTRANSFORM = 35, 36
EMR_SELECTOBJECT, EMR_CREATEPEN, EMR_CREATEBRUSHINDIRECT, EMR_DELETEOBJECT = 37, 38, 39, 40
EMR_LINETO, EMR_BEGINPATH, EMR_ENDPATH, EMR_CLOSEFIGURE, EMR_FILLPATH = 54, 59, 60, 61, 62
EMR_POLYBEZIERTO, EMR_GDICOMMENT, EMR_BITBLT = 5, 70, 76
EMR_EXTCREATEFONTINDIRECTW, EMR_EXTTEXTOUTW = 82, 84
EMR_CREATEMONOBRUSH, EMR_CREATEDIBPATTERNBRUSHPT, EMR_EXTCREATEPEN = 93, 94, 95
R2_MASKPEN, R2_COPYPEN = 9, 13
WHITE_BRUSH, NULL_PEN = 0x80000000, 0x80000008


def rec(t, payload=b''):
    return struct.pack('<II', t, 8 + len(payload)) + payload

def plus_rec(t, flags, data=b''):
    return struct.pack('<HHII', t, flags, 12 + len(data), len(data)) + data

def plus_comment(records):
    body = b'EMF+' + b''.join(records)
    return rec(EMR_GDICOMMENT, struct.pack('<I', len(body)) + body)

def parse(b):
    """List of (type, offset, size) for every record up to EOF."""
    out = []
    off = 0
    while off + 8 <= len(b):
        t, sz = struct.unpack_from('<II', b, off)
        out.append((t, off, sz))
        if sz < 8 or t == EMR_EOF:
            break
        off += sz
    return out

def is_dual_emf(b):
    if len(b) < 108 or struct.unpack_from('<I', b, 0)[0] != EMR_HEADER or b[40:44] != b' EMF':
        return False
    return b'EMF+' in b[:8192]

def _xmul(a, b):
    """Point goes through a, then b."""
    a11, a12, a21, a22, ax, ay = a
    b11, b12, b21, b22, bx, by = b
    return (a11 * b11 + a12 * b21, a11 * b12 + a12 * b22,
            a21 * b11 + a22 * b21, a21 * b12 + a22 * b22,
            ax * b11 + ay * b21 + bx, ax * b12 + ay * b22 + by)

def _xapply(m, x, y):
    return (x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5])

def _ends_with_getdc(b, off):
    dsz = struct.unpack_from('<I', b, off + 8)[0]
    p, end, last = off + 16, off + 12 + dsz, None
    while p + 12 <= end:
        pt, _, psz, _ = struct.unpack_from('<HHII', b, p)
        last = pt
        if psz < 12:
            break
        p += psz
    return last == 0x4004

def _blend_white(c, a=0.5):
    f = lambda v: int(round(v * a + 255 * (1 - a)))
    return f(c & 0xFF) | (f((c >> 8) & 0xFF) << 8) | (f((c >> 16) & 0xFF) << 16)


class _Collect(BasePen):
    def __init__(self, gs, xf):
        super().__init__(gs)
        self.xf, self.ops = xf, []
    def _moveTo(self, p): self.ops.append(('M', [self.xf(p)]))
    def _lineTo(self, p): self.ops.append(('L', [self.xf(p)]))
    def _curveToOne(self, p1, p2, p3): self.ops.append(('C', [self.xf(p) for p in (p1, p2, p3)]))
    def _closePath(self): self.ops.append(('Z', []))


def _glyph_ops(b, off, font, align, fonts, glyph_index):
    """Outline ops in the record's logical coordinates, or None to leave the record as text."""
    if font is None:
        return None
    face, h, w, esc, weight, italic = font
    if esc != 0 or w != 0 or (align & 1):         # escapement, stretched font, TA_UPDATECP
        return None
    f = fonts.get(face, weight >= 600, bool(italic))
    if f is None:
        return None
    tt, gs, cmap = f
    gm = struct.unpack_from('<I', b, off + 24)[0]
    nch, offstr, opts = struct.unpack_from('<III', b, off + 44)
    if gm != 2 or opts & ~(0x2000 | 0x10):         # GM_ADVANCED; allow ETO_PDY, ETO_GLYPH_INDEX
        return None
    if (opts & 0x10) and not glyph_index:
        return None
    refx, refy = struct.unpack_from('<ii', b, off + 36)
    offdx = struct.unpack_from('<I', b, off + 72)[0]
    if opts & 0x10:
        order = tt.getGlyphOrder()
        gids = struct.unpack_from('<%dH' % nch, b, off + offstr)
        if any(g >= len(order) for g in gids):
            return None
        glyphs = [order[g] for g in gids]
    else:
        s = b[off + offstr:off + offstr + 2 * nch].decode('utf-16le', 'replace')
        if any(ord(c) not in cmap for c in s):
            return None
        glyphs = [cmap[ord(c)] for c in s]
    upm, os2 = tt['head'].unitsPerEm, tt['OS/2']
    if h < 0:
        em = -h
    elif h > 0:
        em = h * upm / (os2.usWinAscent + os2.usWinDescent)
    else:
        return None
    k = em / upm
    step = 2 if (opts & 0x2000) else 1
    if offdx:
        dx = struct.unpack_from('<%di' % (nch * step), b, off + offdx)
        adv = [dx[i * step] for i in range(nch)]
        advy = [dx[i * step + 1] for i in range(nch)] if step == 2 else [0] * nch
    else:
        adv = [tt['hmtx'][g][0] * k for g in glyphs]
        advy = [0] * nch
    x0, y0 = float(refx), float(refy)
    ha, va = align & 6, align & 24
    if ha == 6:
        x0 -= sum(adv) / 2
    elif ha == 2:
        x0 -= sum(adv)
    if va == 0:
        y0 += os2.usWinAscent * k
    elif va == 8:
        y0 -= os2.usWinDescent * k
    ops, px, py = [], x0, y0
    for g, a, ay in zip(glyphs, adv, advy):
        pen = _Collect(gs, lambda p, px=px, py=py: (px + p[0] * k, py - p[1] * k))
        gs[g].draw(pen)
        ops += pen.ops
        px += a
        py += ay
    return ops


def _gdi_path(ops):
    ip = lambda p: (int(round(p[0] * SUB)), int(round(p[1] * SUB)))
    out = []
    for kind, pts in ops:
        if kind == 'M':
            out.append(rec(EMR_MOVETOEX, struct.pack('<ii', *ip(pts[0]))))
        elif kind == 'L':
            out.append(rec(EMR_LINETO, struct.pack('<ii', *ip(pts[0]))))
        elif kind == 'C':
            q = [ip(p) for p in pts]
            xs, ys = [p[0] for p in q], [p[1] for p in q]
            out.append(rec(EMR_POLYBEZIERTO, struct.pack('<iiiiI', min(xs), min(ys), max(xs), max(ys), 3)
                           + b''.join(struct.pack('<ii', *p) for p in q)))
        elif kind == 'Z':
            out.append(rec(EMR_CLOSEFIGURE))
    return out


def _plus_paths(ops, dev):
    """EmfPlusPath object payloads (float device coords), split to stay far below 64 KB."""
    paths, pts, types = [], [], []
    def flush():
        if pts:
            data = struct.pack('<III', 0xDBC01002, len(pts), 0)
            data += b''.join(struct.pack('<ff', *p) for p in pts) + bytes(types)
            paths.append(data + b'\x00' * (-len(data) % 4))
        pts.clear()
        types.clear()
    for kind, p in ops:
        if kind == 'M':
            if len(pts) > 4000:
                flush()
            pts.append(dev(p[0])); types.append(0)
        elif kind == 'L':
            pts.append(dev(p[0])); types.append(1)
        elif kind == 'C':
            for q in p:
                pts.append(dev(q)); types.append(3)
        elif kind == 'Z' and types:
            types[-1] |= 0x80
    flush()
    return paths


def _text_records(ops, color, H, xf, worg, vorg, save_idx):
    def dev(p):
        x, y = _xapply(xf, p[0], p[1])
        return (x - worg[0] + vorg[0], y - worg[1] + vorg[1])
    argb = 0xFF000000 | ((color & 0xFF) << 16) | (color & 0xFF00) | ((color >> 16) & 0xFF)
    plus = [plus_rec(0x4025, 0, struct.pack('<I', save_idx)),   # Save
            plus_rec(0x402B, 0),                                  # ResetWorldTransform
            plus_rec(0x401E, 0x09)]                               # SetAntiAliasMode: on
    for pdata in _plus_paths(ops, dev):
        plus.append(plus_rec(0x4008, (3 << 8) | PLUS_OBJ_ID, pdata))                    # Object: Path
        plus.append(plus_rec(0x4014, 0x8000 | PLUS_OBJ_ID, struct.pack('<I', argb)))    # FillPath
    plus.append(plus_rec(0x4026, 0, struct.pack('<I', save_idx)))                       # Restore
    gdi = [rec(EMR_SAVEDC),
           rec(EMR_MODIFYWORLDTRANSFORM, struct.pack('<6fI', 1.0 / SUB, 0, 0, 1.0 / SUB, 0, 0, 2)),
           rec(EMR_CREATEBRUSHINDIRECT, struct.pack('<IIII', H, 0, color & 0xFFFFFF, 0)),
           rec(EMR_SELECTOBJECT, struct.pack('<I', H)),
           rec(EMR_SELECTOBJECT, struct.pack('<I', NULL_PEN)),
           rec(EMR_SETPOLYFILLMODE, struct.pack('<I', 2)),
           rec(EMR_BEGINPATH), *_gdi_path(ops), rec(EMR_ENDPATH),
           rec(EMR_FILLPATH, struct.pack('<iiii', 0, 0, -1, -1)),
           rec(EMR_RESTOREDC, struct.pack('<i', -1)),
           rec(EMR_DELETEOBJECT, struct.pack('<I', H))]
    # EMF+ comment first (ends the GetDC section, so EMF+ renderers skip the GDI copy),
    # then the GDI copy, then a GetDC to resume the original section.
    return plus_comment(plus) + b''.join(gdi) + plus_comment([plus_rec(0x4004, 0)])


def fix_emf(data, fonts, text=True, transparency=True, glyph_index=True, stats=None):
    """Return the rewritten EMF bytes. `fonts` is a FontIndex. Non-dual EMFs are returned as is."""
    b = bytes(data)
    if not is_dual_emf(b):
        return b
    st = stats if stats is not None else {}
    nhandles = struct.unpack_from('<H', b, 56)[0]
    H = nhandles                                  # one extra GDI object slot for our brush
    objs = {}
    cur_font = cur_brush = None
    text_color = align = 0
    in_dc = False
    xf, worg, vorg, stack = IDENT, (0, 0), (0, 0), []   # GDI state as EMF+ renderers see it
    stipple, stipple_brush_live, save_idx = None, False, 0x7EAD0000
    recs = parse(b)
    out = []
    i = 0
    while i < len(recs):
        t, off, sz = recs[i]
        i += 1
        r = b[off:off + sz]
        if t == EMR_GDICOMMENT and b[off + 12:off + 16] == b'EMF+':
            in_dc = _ends_with_getdc(b, off)
        elif in_dc and t == EMR_SAVEDC:
            stack.append((xf, worg, vorg))
        elif in_dc and t == EMR_RESTOREDC:
            n = struct.unpack_from('<i', b, off + 8)[0]
            for _ in range(-n if n < 0 else 0):
                if stack:
                    xf, worg, vorg = stack.pop()
        elif in_dc and t == EMR_SETWORLDTRANSFORM:
            xf = struct.unpack_from('<6f', b, off + 8)
        elif in_dc and t == EMR_MODIFYWORLDTRANSFORM:
            m = struct.unpack_from('<6f', b, off + 8)
            mode = struct.unpack_from('<I', b, off + 32)[0]
            xf = IDENT if mode == 1 else _xmul(m, xf) if mode == 2 else _xmul(xf, m) if mode == 3 else m
        elif in_dc and t == EMR_SETWINDOWORGEX:
            worg = struct.unpack_from('<ii', b, off + 8)
        elif in_dc and t == EMR_SETVIEWPORTORGEX:
            vorg = struct.unpack_from('<ii', b, off + 8)
        elif t == EMR_EXTCREATEFONTINDIRECTW:
            ih, h, w, esc, _, wt = struct.unpack_from('<Iiiiii', b, off + 8)
            face = b[off + 40:off + 104].decode('utf-16le', 'replace').split('\x00')[0]
            objs[ih] = ('font', (face, h, w, esc, wt, b[off + 32]))
        elif t == EMR_CREATEBRUSHINDIRECT:
            ih, style, color = struct.unpack_from('<III', b, off + 8)
            objs[ih] = ('brush', color)
        elif t in (EMR_CREATEPEN, EMR_EXTCREATEPEN, EMR_CREATEMONOBRUSH, EMR_CREATEDIBPATTERNBRUSHPT):
            kind = 'pattern' if t in (EMR_CREATEMONOBRUSH, EMR_CREATEDIBPATTERNBRUSHPT) else 'other'
            objs[struct.unpack_from('<I', b, off + 8)[0]] = (kind, None)
        elif t == EMR_DELETEOBJECT:
            objs.pop(struct.unpack_from('<I', b, off + 8)[0], None)
        elif t == EMR_SELECTOBJECT:
            ih = struct.unpack_from('<I', b, off + 8)[0]
            kind, val = objs.get(ih, (None, None))
            if kind == 'font':
                cur_font = val
            elif kind == 'brush':
                cur_brush = val
            if stipple is not None and kind == 'pattern':
                if not stipple_brush_live:
                    out.append(rec(EMR_CREATEBRUSHINDIRECT, struct.pack('<IIII', H, 0, stipple, 0)))
                    stipple_brush_live = True
                out.append(rec(EMR_SELECTOBJECT, struct.pack('<I', H)))
                continue
        elif t == EMR_SETROP2 and stipple is not None:
            out.append(rec(EMR_SETROP2, struct.pack('<I', R2_COPYPEN)))
            continue
        elif t == EMR_SETTEXTCOLOR:
            text_color = struct.unpack_from('<I', b, off + 8)[0]
        elif t == EMR_SETTEXTALIGN:
            align = struct.unpack_from('<I', b, off + 8)[0]
        elif t == EMR_BITBLT and transparency and not in_dc and sz >= 44:
            rop = struct.unpack_from('<I', b, off + 40)[0]
            if rop == PATINVERT and stipple is not None:          # closing XOR of the stipple trick
                stipple = None
                if stipple_brush_live:
                    out.append(rec(EMR_SELECTOBJECT, struct.pack('<I', WHITE_BRUSH)))
                    out.append(rec(EMR_DELETEOBJECT, struct.pack('<I', H)))
                    stipple_brush_live = False
                st['transparency_filled'] = st.get('transparency_filled', 0) + 1
                continue
            if rop == PATINVERT:                                   # opening XOR: path-mask variant?
                kinds, j = [], i
                while j < len(recs) and j < i + 400:
                    tj, oj, _ = recs[j]
                    if tj == EMR_BITBLT:
                        kinds.append(struct.unpack_from('<I', b, oj + 40)[0])
                        break
                    if tj == EMR_SETROP2:
                        kinds.append(('rop2', struct.unpack_from('<I', b, oj + 8)[0]))
                    j += 1
                if kinds and kinds[-1] == PATINVERT and ('rop2', R2_MASKPEN) in kinds and cur_brush is not None:
                    stipple = _blend_white(cur_brush)
                    continue
            if rop in (PATINVERT, PATAND):
                st['transparency_dropped'] = st.get('transparency_dropped', 0) + 1
                continue
        elif t == EMR_EXTTEXTOUTW and text and in_dc:
            ops = _glyph_ops(b, off, cur_font, align, fonts, glyph_index)
            if ops is not None:
                if ops:
                    out.append(_text_records(ops, text_color, H, xf, worg, vorg, save_idx))
                    save_idx += 1
                st['text_converted'] = st.get('text_converted', 0) + 1
                continue
            st['text_kept'] = st.get('text_kept', 0) + 1
            if cur_font:
                st.setdefault('fonts_not_converted', set()).add(cur_font[0])
        out.append(r)
    if not (st.get('text_converted') or st.get('transparency_filled') or st.get('transparency_dropped')):
        return b
    body = bytearray(b''.join(out))
    struct.pack_into('<I', body, 48, len(body))           # nBytes
    struct.pack_into('<H', body, 56, nhandles + 1)        # nHandles
    struct.pack_into('<I', body, 52, len(parse(bytes(body))))  # nRecords
    return bytes(body)
