"""Apply fix_emf() to every EMF inside an Office Open XML package (.docx/.pptx/.xlsx and friends)."""
import os
import zipfile

from .emf import fix_emf, is_dual_emf
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
