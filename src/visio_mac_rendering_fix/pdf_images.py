"""Put full-resolution images back into a PDF exported from Word.

Word's PDF export resamples bitmaps inside EMF figures (e.g. Visio figures with scope
screenshots or photos) to about 200 ppi and often JPEG-compresses them. The original bitmaps
are still stored, losslessly, in each figure's EMF+ layer inside the .docx/.pptx. This tool
matches every image in the PDF to those originals and swaps the original pixels in. Placement
on the page is untouched, so the layout is identical; only sharpness changes.

  1. Images Word kept whole are found by comparing them with each original as drawn.
  2. Images Word cropped to their visible part are found by mapping their rectangle on the
     page back into the figure, using the images matched in step 1 to calibrate the mapping.
Every replacement must look like the image it replaces (correlation >= --min-similarity).

Command: visio-pdf-hires exported.pdf source.docx [-o out.pdf]
Needs the optional extra: pip install 'visio-mac-rendering-fix[pdf]' (PyMuPDF, Pillow, numpy).
"""
import argparse
import io
import math
import os
import struct
import sys
import zipfile

try:
    import numpy as np
    from PIL import Image
    try:
        import pymupdf as fitz  # PyMuPDF >= 1.24
    except ImportError:
        import fitz
    HAVE_DEPS = True
except ImportError:  # pragma: no cover
    HAVE_DEPS = False

MISSING_DEPS = ("visio-pdf-hires needs PyMuPDF, Pillow and numpy:\n"
                "  pip install 'visio-mac-rendering-fix[pdf]'")

IDENT = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)


# ------------------------------------------------------------------ EMF+ parsing
def _mul(a, b):
    """Apply a, then b (GDI+ matrix order m11, m12, m21, m22, dx, dy)."""
    return (a[0] * b[0] + a[1] * b[2], a[0] * b[1] + a[1] * b[3],
            a[2] * b[0] + a[3] * b[2], a[2] * b[1] + a[3] * b[3],
            a[4] * b[0] + a[5] * b[2] + b[4], a[4] * b[1] + a[5] * b[3] + b[5])

def _apply(m, x, y):
    return (x * m[0] + y * m[2] + m[4], x * m[1] + y * m[3] + m[5])

def _decode_image(data):
    """PIL image from an EmfPlusImage object holding a compressed bitmap (PNG, JPEG, ...)."""
    if len(data) < 28:
        return None
    _, itype = struct.unpack_from('<II', data, 0)
    if itype != 1:
        return None
    btype = struct.unpack_from('<I', data, 24)[0]
    if btype != 1:
        return None
    try:
        im = Image.open(io.BytesIO(data[28:]))
        im.load()
        return im
    except Exception:
        return None

def emfplus_draws(emf):
    """List of draws: dict(img, src=(sx, sy, sw, sh), pts=(P0, P1, P2)) with the destination
    parallelogram in EMF+ device coordinates (P0 = top-left, P1 = top-right, P2 = bottom-left)."""
    out, off, pending, objs, xf, stack = [], 0, {}, {}, IDENT, {}
    while off + 8 <= len(emf):
        t, sz = struct.unpack_from('<II', emf, off)
        if sz < 8:
            break
        if t == 70 and emf[off + 12:off + 16] == b'EMF+':
            p, end = off + 16, off + 12 + struct.unpack_from('<I', emf, off + 8)[0]
            while p + 12 <= end:
                ty, fl, ps, ds = struct.unpack_from('<HHII', emf, p)
                if ps < 12:
                    break
                d0 = p + 12
                if ty == 0x4008:                                        # Object
                    oid, otype = fl & 0xFF, (fl >> 8) & 0x7F
                    data = None
                    if fl & 0x8000:                                     # continued object
                        total = struct.unpack_from('<I', emf, d0)[0]
                        buf = pending.setdefault(oid, [total, bytearray()])
                        buf[1] += emf[d0 + 4:d0 + ds]
                        if len(buf[1]) >= total:
                            data = bytes(buf[1][:total]); del pending[oid]
                    else:
                        data = (bytes(pending.pop(oid)[1]) if oid in pending else b'') + emf[d0:d0 + ds]
                    if data is not None:
                        objs[oid] = _decode_image(data) if otype == 5 else None
                elif ty == 0x402A:
                    xf = struct.unpack_from('<6f', emf, d0)
                elif ty == 0x402B:
                    xf = IDENT
                elif ty in (0x402C, 0x402D, 0x402E, 0x402F):
                    if ty == 0x402C:
                        m = struct.unpack_from('<6f', emf, d0)
                    elif ty == 0x402D:
                        m = (1, 0, 0, 1) + struct.unpack_from('<2f', emf, d0)
                    elif ty == 0x402E:
                        a, b = struct.unpack_from('<2f', emf, d0); m = (a, 0, 0, b, 0, 0)
                    else:
                        r = math.radians(struct.unpack_from('<f', emf, d0)[0])
                        m = (math.cos(r), math.sin(r), -math.sin(r), math.cos(r), 0, 0)
                    xf = _mul(xf, m) if fl & 0x2000 else _mul(m, xf)
                elif ty in (0x4025, 0x4028):                            # Save, BeginContainerNoParams
                    stack[struct.unpack_from('<I', emf, d0)[0]] = xf
                elif ty == 0x4027:                                      # BeginContainer
                    dx, dy, dw, dh, sx, sy, sw, sh = struct.unpack_from('<8f', emf, d0)
                    stack[struct.unpack_from('<I', emf, d0 + 32)[0]] = xf
                    if sw and sh:
                        xf = _mul((dw / sw, 0, 0, dh / sh, dx - sx * dw / sw, dy - sy * dh / sh), xf)
                elif ty in (0x4026, 0x4029):                            # Restore, EndContainer
                    xf = stack.get(struct.unpack_from('<I', emf, d0)[0], xf)
                elif ty in (0x401A, 0x401B):                            # DrawImage, DrawImagePoints
                    img = objs.get(fl & 0xFF)
                    if img is not None:
                        unit = struct.unpack_from('<I', emf, d0 + 4)[0]
                        src = struct.unpack_from('<4f', emf, d0 + 8)
                        if unit != 2 or not src[2] or not src[3]:
                            src = (0, 0, img.width, img.height)
                        if ty == 0x401B:
                            if fl & 0x4000:
                                pts = [struct.unpack_from('<hh', emf, d0 + 28 + 4 * i) for i in range(3)]
                            else:
                                pts = [struct.unpack_from('<ff', emf, d0 + 28 + 8 * i) for i in range(3)]
                        else:
                            fmt = '<4h' if fl & 0x4000 else '<4f'
                            x, y, w, h = struct.unpack_from(fmt, emf, d0 + 24)
                            pts = [(x, y), (x + w, y), (x, y + h)]
                        out.append(dict(img=img, src=src, pts=tuple(_apply(xf, *q) for q in pts)))
                p += ps
        if t == 14:
            break
        off += sz
    return out


def _dest_bbox(dr):
    P0, P1, P2 = dr['pts']
    xs = [P0[0], P1[0], P2[0], P1[0] + P2[0] - P0[0]]
    ys = [P0[1], P1[1], P2[1], P1[1] + P2[1] - P0[1]]
    return (min(xs), min(ys), max(xs), max(ys))

def _crop_for_device_rect(dr, R):
    """Source-pixel box showing device rectangle R of draw dr (None if degenerate)."""
    (P0, P1, P2), (sx, sy, sw, sh) = dr['pts'], dr['src']
    ux, uy = (P1[0] - P0[0], P1[1] - P0[1]), (P2[0] - P0[0], P2[1] - P0[1])
    det = ux[0] * uy[1] - ux[1] * uy[0]
    if abs(det) < 1e-9:
        return None
    us, vs = [], []
    for X, Y in ((R[0], R[1]), (R[2], R[1]), (R[0], R[3]), (R[2], R[3])):
        a = ((X - P0[0]) * uy[1] - (Y - P0[1]) * uy[0]) / det
        b = ((Y - P0[1]) * ux[0] - (X - P0[0]) * ux[1]) / det
        us.append(sx + a * sw); vs.append(sy + b * sh)
    img = dr['img']
    box = (max(0, round(min(us))), max(0, round(min(vs))),
           min(img.width, round(max(us))), min(img.height, round(max(vs))))
    return box if box[2] - box[0] > 4 and box[3] - box[1] > 4 else None


# ------------------------------------------------------------------ image helpers
def on_white(im):
    if im.mode in ('RGBA', 'LA', 'P', 'PA'):
        im = im.convert('RGBA')
        bg = Image.new('RGB', im.size, (255, 255, 255))
        bg.paste(im, mask=im.split()[-1])
        return bg
    return im.convert('RGB')

def similarity(pdf_img, cand):
    """Correlation after low-pass filtering both, so JPEG noise and resampling don't matter."""
    from PIL import ImageFilter
    w, h = max(8, pdf_img.width // 2), max(8, pdf_img.height // 2)
    blur = ImageFilter.GaussianBlur(1)
    a = np.asarray(pdf_img.convert('L').filter(blur).resize((w, h), Image.BOX), dtype=float).ravel()
    o = np.asarray(cand.convert('L').resize(pdf_img.size, Image.LANCZOS).filter(blur)
                   .resize((w, h), Image.BOX), dtype=float).ravel()
    if a.std() < 1 or o.std() < 1:
        return 0.0
    return float(np.corrcoef(a, o)[0, 1])

def is_photo(im):
    """Photographs compress far better as JPEG; screenshots and line art stay lossless."""
    small = im.resize((min(im.width, 256), min(im.height, 256)))
    return len(set(small.getdata())) > 6000

def write_image(doc, xref, src, smask=0):
    """Overwrite image XObject `xref` in place with `src`. If the PDF image had a soft mask
    (`smask` xref), the mask is rewritten from src's alpha; otherwise src is flattened on white.
    Returns the encoding used."""
    import zlib
    for key in ('Mask', 'DecodeParms', 'Decode', 'ImageMask'):
        doc.xref_set_key(xref, key, 'null')
    rgba = src.convert('RGBA')
    if smask:
        alpha = rgba.split()[-1]
        doc.update_stream(smask, zlib.compress(alpha.tobytes(), 9), compress=False)
        for key, val in (('Width', str(alpha.width)), ('Height', str(alpha.height)), ('ColorSpace', '/DeviceGray'),
                         ('BitsPerComponent', '8'), ('Filter', '/FlateDecode'), ('DecodeParms', 'null'),
                         ('Decode', 'null'), ('Matte', 'null')):    # new colours are not premultiplied
            doc.xref_set_key(smask, key, val)
        im = rgba.convert('RGB')
    else:
        doc.xref_set_key(xref, 'SMask', 'null')
        im = on_white(rgba)
    if is_photo(im):
        buf = io.BytesIO(); im.save(buf, 'JPEG', quality=95, subsampling=0)
        data, filt, kind = buf.getvalue(), '/DCTDecode', 'JPEG q95'
    else:
        data, filt, kind = zlib.compress(im.tobytes(), 9), '/FlateDecode', 'lossless'
    doc.update_stream(xref, data, compress=False)
    for key, val in (('Width', str(im.width)), ('Height', str(im.height)), ('ColorSpace', '/DeviceRGB'),
                     ('BitsPerComponent', '8'), ('Filter', filt)):
        doc.xref_set_key(xref, key, val)
    return kind


# ------------------------------------------------------------------ main
def restore_pdf_images(pdf, source, out, min_similarity=0.97, log=print):
    """Write `out`: `pdf` with its images replaced by the full-resolution originals found in
    `source` (the Office file it was exported from). Returns a list of replacement records."""
    if not HAVE_DEPS:
        raise RuntimeError(MISSING_DEPS)
    if os.path.abspath(out) == os.path.abspath(pdf):
        raise ValueError('output must differ from input')
    figures = {}                                  # emf part name -> list of draws
    with zipfile.ZipFile(source) as z:
        for n in z.namelist():
            if n.lower().endswith('.emf'):
                draws = emfplus_draws(z.read(n))
                for dr in draws:
                    dr['rgba'] = dr['img'].convert('RGBA').crop(tuple(map(round, (
                        dr['src'][0], dr['src'][1], dr['src'][0] + dr['src'][2], dr['src'][1] + dr['src'][3]))))
                    dr['full'] = on_white(dr['rgba'])
                if draws:
                    figures[n] = draws
    log(f'{sum(len(v) for v in figures.values())} original images in {len(figures)} figures')

    doc = fitz.open(pdf)
    todo = []                                     # (page, xref, rect, pdf image)
    for pno, page in enumerate(doc):
        for item in page.get_images(full=True):
            xref = item[0]
            rects = page.get_image_rects(xref)
            if not rects:
                continue
            info = doc.extract_image(xref)
            im = Image.open(io.BytesIO(info['image'])); im.load()
            smask = item[1]
            if smask:                                    # composite with its soft mask, as displayed
                try:
                    m = Image.open(io.BytesIO(doc.extract_image(smask)['image'])).convert('L').resize(im.size)
                    im = im.convert('RGB'); im.putalpha(m)
                except Exception:
                    pass
            todo.append((pno, xref, rects[0], on_white(im), info['ext'], smask))

    done, anchors, records = {}, {}, []                        # anchors: (figure, page) -> [(device bbox, page rect)]
    def replace(pno, xref, rect, pim, ext, smask, new, score, label):
        if xref in done or new.width * new.height <= pim.width * pim.height * 1.2:
            return False
        kind = write_image(doc, xref, new, smask)
        records.append(dict(page=pno + 1, xref=xref, old=pim.size, new=new.size, encoding=kind,
                            source=label, score=score))
        log(f'  page {pno + 1}: {pim.width}x{pim.height} {ext} (~{pim.width / (rect.width / 72):.0f} ppi) -> '
              f'{new.width}x{new.height} {kind} (~{new.width / (rect.width / 72):.0f} ppi)  [{label}, match {score:.3f}]')
        done[xref] = True
        return True

    # pass 1: images Word kept whole
    for pno, xref, rect, pim, ext, smask in todo:
        best = max(((similarity(pim, dr['full']), fig, dr) for fig, ds in figures.items() for dr in ds),
                   key=lambda x: x[0], default=None)
        if best and best[0] >= min_similarity:
            score, fig, dr = best
            anchors.setdefault((fig, pno), []).append((_dest_bbox(dr), rect))
            replace(pno, xref, rect, pim, ext, smask, dr['rgba'], score, fig.split('/')[-1])

    # pass 2: map each remaining image's page rectangle back into the figure (handles crops)
    for pno, xref, rect, pim, ext, smask in todo:
        if xref in done:
            continue
        best = None
        for (fig, apno), pairs in anchors.items():
            if apno != pno:
                continue
            # page = a * device + b, fitted per axis from the anchor rectangles
            dx = [(d[0], r.x0) for d, r in pairs] + [(d[2], r.x1) for d, r in pairs]
            dy = [(d[1], r.y0) for d, r in pairs] + [(d[3], r.y1) for d, r in pairs]
            (ax, bx), (ay, by) = (np.polyfit([u for u, _ in dx], [v for _, v in dx], 1),
                                  np.polyfit([u for u, _ in dy], [v for _, v in dy], 1))
            R = ((rect.x0 - bx) / ax, (rect.y0 - by) / ay, (rect.x1 - bx) / ax, (rect.y1 - by) / ay)
            R = (min(R[0], R[2]), min(R[1], R[3]), max(R[0], R[2]), max(R[1], R[3]))
            for dr in figures[fig]:
                D = _dest_bbox(dr)
                if R[0] >= D[2] or R[2] <= D[0] or R[1] >= D[3] or R[3] <= D[1]:
                    continue
                box = _crop_for_device_rect(dr, R)
                if box is None:
                    continue
                cand = dr['img'].convert('RGBA').crop(box)
                sc = similarity(pim, on_white(cand))
                ix = max(0, min(R[2], D[2]) - max(R[0], D[0])) * max(0, min(R[3], D[3]) - max(R[1], D[1]))
                iou = ix / ((R[2] - R[0]) * (R[3] - R[1]) + (D[2] - D[0]) * (D[3] - D[1]) - ix)
                if best is None or sc > best[0]:
                    best = (sc, fig, cand, box, iou)
        # position on the page is independent evidence: an image sitting exactly where an original
        # is drawn needs less pixel agreement (thin traces lose a lot when downsampled)
        need = min_similarity if not best or best[4] < 0.97 else min(min_similarity, 0.90)
        if best and best[0] >= need:
            label = best[1].split('/')[-1] + ('' if best[4] >= 0.97 else f' crop {best[3]}')
            replace(pno, xref, rect, pim, ext, smask, best[2], best[0], label + f', position match {best[4]:.2f}')
        elif best and best[0] > 0.5:
            log(f'  page {pno + 1}: {pim.width}x{pim.height} closest original only {best[0]:.3f}, left as is')

    doc.save(out, garbage=3, deflate=True)
    log(f'replaced {len(done)} image(s); wrote {out}')
    return records


def main(argv=None):
    ap = argparse.ArgumentParser(prog='visio-pdf-hires', description=__doc__.split('\n\n')[0])
    ap.add_argument('pdf', help='PDF exported from the Office file')
    ap.add_argument('source', help='the .docx / .pptx / .xlsx it was exported from')
    ap.add_argument('-o', '--output', help='default: <pdf name>_hires.pdf')
    ap.add_argument('--min-similarity', type=float, default=0.97,
                    help='how closely a replacement must resemble the image it replaces (default 0.97)')
    ap.add_argument('-q', '--quiet', action='store_true')
    a = ap.parse_args(argv)
    if not HAVE_DEPS:
        print(MISSING_DEPS, file=sys.stderr)
        return 2
    out = a.output or os.path.splitext(a.pdf)[0] + '_hires.pdf'
    try:
        restore_pdf_images(a.pdf, a.source, out, a.min_similarity,
                           log=(lambda *_: None) if a.quiet else print)
    except (ValueError, OSError, zipfile.BadZipFile) as e:
        print(f'error: {e}', file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
