"""Apply fix_emf() to every EMF inside an Office Open XML package (.docx/.pptx/.xlsx and friends)."""
import os
import zipfile

import hashlib

from .emf import fix_emf, is_dual_emf, is_fixed_emf, unfix_emf, strip_markers
from .fonts import FontIndex

SUPPORTED = ('.docx', '.docm', '.dotx', '.dotm', '.pptx', '.pptm', '.potx', '.potm',
             '.ppsx', '.ppsm', '.xlsx', '.xlsm', '.xltx', '.xltm')


def default_output(path):
    root, ext = os.path.splitext(path)
    return root + '_macfix' + ext


def fix_package(src, dst, fonts=None, **options):
    """Copy `src` to `dst`, rewriting embedded dual EMF+ images. Returns {part name: stats}."""
    if os.path.abspath(src) == os.path.abspath(dst):
        raise ValueError('output must differ from input')
    fonts = fonts or FontIndex()
    report = {}
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename.lower().endswith('.emf') and is_dual_emf(data):
                stats = {}
                data = fix_emf(data, fonts, stats=stats, **options)
                report[item.filename] = stats
            zo.writestr(item, data)
    return report


def default_undo_output(path):
    root, ext = os.path.splitext(path)
    return root + '_original' + ext


def _looks_converted(data):
    """A dual EMF whose text was turned into outlines by an old version (no undo data):
    our glyph fills use EMF+ object slot 63 and the font records remain without any text."""
    import struct
    from .emf import parse, EMR_GDICOMMENT, EMR_EXTTEXTOUTW
    b = bytes(data)
    if not is_dual_emf(b) or is_fixed_emf(b):
        return False
    text = ours = 0
    for t, off, sz in parse(b):
        if t == EMR_EXTTEXTOUTW:
            text += 1
        elif t == EMR_GDICOMMENT and b[off + 12:off + 16] == b'EMF+':
            p, end = off + 16, off + 12 + struct.unpack_from('<I', b, off + 8)[0]
            while p + 12 <= end:
                pt, pf, ps, _ = struct.unpack_from('<HHII', b, p)
                if pt == 0x4014 and (pf & 0xFF) == 63:
                    ours += 1
                if ps < 12:
                    break
                p += ps
    return ours > 0 and text == 0


def legacy_originals(originals, fonts=None):
    """{sha1 of what versions <= 0.4 would have produced: original EMF bytes} for every dual
    EMF in the given Office files, to recover figures converted without undo data."""
    fonts = fonts or FontIndex()
    table = {}
    for path in originals or []:
        with zipfile.ZipFile(path) as z:
            for n in z.namelist():
                if not n.lower().endswith('.emf'):
                    continue
                orig = unfix_emf(z.read(n))
                if not is_dual_emf(orig) or _looks_converted(orig):
                    continue
                legacy = strip_markers(fix_emf(orig, fonts))
                table.setdefault(hashlib.sha1(legacy).hexdigest(), orig)
    return table


def unfix_package(src, dst, originals=None, fonts=None):
    """Copy `src` to `dst` with every converted figure restored to its original EMF.
    Returns {part name: 'undone' | 'recovered' | 'converted, no original found'}."""
    if os.path.abspath(src) == os.path.abspath(dst):
        raise ValueError('output must differ from input')
    table = None
    report = {}
    with zipfile.ZipFile(src) as zi, zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zo:
        for item in zi.infolist():
            data = zi.read(item.filename)
            if item.filename.lower().endswith('.emf'):
                if is_fixed_emf(data):
                    data = unfix_emf(data)
                    report[item.filename] = 'undone'
                elif _looks_converted(data):
                    if table is None:
                        table = legacy_originals(originals, fonts) if originals else {}
                    orig = table.get(hashlib.sha1(data).hexdigest())
                    if orig is not None:
                        data = orig
                        report[item.filename] = 'recovered'
                    else:
                        report[item.filename] = 'converted, no original found'
            zo.writestr(item, data)
    return report
