"""Locate installed font files by family name and style."""
import logging
import os
import sys
from fontTools.ttLib import TTFont, TTCollection

# system fonts often have harmless quirks (odd 'created' dates, padding) that fontTools logs
logging.getLogger('fontTools').setLevel(logging.ERROR)

def default_font_dirs():
    home = os.path.expanduser('~')
    if sys.platform == 'darwin':
        return ['/System/Library/Fonts/Supplemental', '/System/Library/Fonts', '/Library/Fonts',
                os.path.join(home, 'Library/Fonts'),
                # fonts bundled with Office for Mac (used only if a face is missing elsewhere)
                '/Applications/Microsoft Word.app/Contents/Resources/DFonts',
                '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts']
    if os.name == 'nt':
        windir = os.environ.get('WINDIR', r'C:\Windows')
        local = os.environ.get('LOCALAPPDATA', '')
        return [os.path.join(windir, 'Fonts'), os.path.join(local, 'Microsoft', 'Windows', 'Fonts')]
    return ['/usr/share/fonts', '/usr/local/share/fonts', os.path.join(home, '.fonts'),
            os.path.join(home, '.local/share/fonts')]

def _name(font, ids):
    tbl = font['name']
    for nid in ids:
        rec = tbl.getName(nid, 3, 1, 0x409) or tbl.getName(nid, 1, 0, 0)
        if rec is not None:
            return str(rec)
    return None

class FontIndex:
    """Maps (family, bold, italic) -> (path, font number). Built lazily on first lookup."""

    def __init__(self, dirs=None):
        self.dirs = list(dirs) if dirs else default_font_dirs()
        self._index = None
        self._loaded = {}

    def _build(self):
        idx = {}
        for d in self.dirs:
            if not os.path.isdir(d):
                continue
            for root, _, files in os.walk(d):
                for fn in files:
                    ext = fn.lower().rsplit('.', 1)[-1]
                    if ext not in ('ttf', 'otf', 'ttc', 'otc'):
                        continue
                    path = os.path.join(root, fn)
                    try:
                        count = len(TTCollection(path, lazy=True).fonts) if ext in ('ttc', 'otc') else 1
                    except Exception:
                        continue
                    for num in range(count):
                        try:
                            f = TTFont(path, fontNumber=num, lazy=True)
                            if 'OS/2' not in f or 'name' not in f:
                                continue
                            sel = f['OS/2'].fsSelection
                            bold = bool(sel & 0x20) or f['OS/2'].usWeightClass >= 600
                            italic = bool(sel & 0x01)
                            for fam in {_name(f, [1]), _name(f, [16])}:
                                if fam:
                                    idx.setdefault((fam.lower(), bold, italic), (path, num))
                        except Exception:
                            continue
        self._index = idx

    def font_bytes(self, family, bold, italic):
        """The font file's bytes (a single font, extracted from a collection if needed)."""
        if self._index is None:
            self._build()
        path, num = self._index[(family.lower(), bold, italic)]
        if not path.lower().endswith(('.ttc', '.otc')):
            with open(path, 'rb') as f:
                return f.read()
        import io
        buf = io.BytesIO()
        TTFont(path, fontNumber=num).save(buf)
        return buf.getvalue()

    def get(self, family, bold, italic):
        """Return (TTFont, glyphset, cmap) for an exact style match, or None."""
        if self._index is None:
            self._build()
        key = (family.lower(), bold, italic)
        hit = self._index.get(key)
        if hit is None:
            return None
        if hit not in self._loaded:
            f = TTFont(hit[0], fontNumber=hit[1])
            self._loaded[hit] = (f, f.getGlyphSet(), f.getBestCmap())
        return self._loaded[hit]
