import io
import os
import struct
import sys
import tempfile
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'src'))
sys.path.insert(0, os.path.join(HERE, '..', 'examples'))

from visio_mac_rendering_fix import FontIndex, fix_emf, fix_package, is_dual_emf   # noqa: E402
from visio_mac_rendering_fix.emf import parse, _ends_with_getdc                   # noqa: E402
import make_demo                                                                  # noqa: E402

FONTS = FontIndex()
HAVE_FONT = FONTS.get('Arial Narrow', False, False) is not None


def plus_types(b):
    """All EMF+ record types, in order."""
    out = []
    for t, off, sz in parse(b):
        if t == 70 and b[off + 12:off + 16] == b'EMF+':
            p, end = off + 16, off + 12 + struct.unpack_from('<I', b, off + 8)[0]
            while p < end:
                pt, _, psz, _ = struct.unpack_from('<HHII', b, p)
                out.append(pt)
                p += psz
    return out


def gdi_records(b):
    """(type, offset, inside_getdc) for every non-EMF+ record."""
    out, in_dc = [], False
    for t, off, sz in parse(b):
        if t == 70 and b[off + 12:off + 16] == b'EMF+':
            in_dc = _ends_with_getdc(b, off)
        else:
            out.append((t, off, in_dc))
    return out


class StructureTests(unittest.TestCase):
    def check_valid(self, b):
        recs = parse(b)
        self.assertEqual(recs[0][0], 1)
        self.assertEqual(recs[-1][0], 14)
        self.assertEqual(sum(sz for _, _, sz in recs), len(b))
        self.assertEqual(struct.unpack_from('<I', b, 48)[0], len(b))
        self.assertEqual(struct.unpack_from('<I', b, 52)[0], len(recs))
        for t, off, sz in recs:
            if t == 70 and b[off + 12:off + 16] == b'EMF+':
                p, end = off + 16, off + 12 + struct.unpack_from('<I', b, off + 8)[0]
                while p < end:
                    _, _, psz, dsz = struct.unpack_from('<HHII', b, p)
                    self.assertEqual(psz, dsz + 12)
                    self.assertEqual(psz % 4, 0)
                    p += psz
                self.assertEqual(p, end)

    def test_demo_is_dual(self):
        b = make_demo.make_emf()
        self.assertTrue(is_dual_emf(b))
        self.check_valid(b)

    def test_plain_emf_untouched(self):
        b = make_demo.make_emf().replace(b'EMF+', b'XXXX')   # plain EMF: no EMF+ layer
        self.assertFalse(is_dual_emf(b))
        self.assertEqual(fix_emf(b, FONTS), b)

    def test_transparency_only(self):
        src = make_demo.make_emf()
        out = fix_emf(src, FONTS, text=False, stats=(st := {}))
        self.check_valid(out)
        self.assertEqual(st.get('transparency_filled'), 1)
        # no raster-op blits left; EMF+ layer and every GetDC-section record untouched
        rop = lambda b: [struct.unpack_from('<I', b, o + 40)[0] for t, o, _ in gdi_records(b) if t == 76]
        self.assertTrue(rop(src))
        self.assertFalse(rop(out))
        self.assertEqual(plus_types(src), plus_types(out))
        dc = lambda b: [b[o:o + struct.unpack_from('<I', b, o + 4)[0]] for t, o, d in gdi_records(b) if d]
        self.assertEqual(dc(src), dc(out))

    @unittest.skipUnless(HAVE_FONT, 'Arial Narrow not installed')
    def test_text_converted(self):
        src = make_demo.make_emf()
        out = fix_emf(src, FONTS, stats=(st := {}))
        self.check_valid(out)
        runs = sum(1 for t, _, _ in gdi_records(src) if t == 84)
        self.assertGreater(runs, 10)
        self.assertEqual(st.get('text_converted'), runs)
        # no text records left; each run became an EMF+ path fill plus a GDI copy outside GetDC
        recs = gdi_records(out)
        self.assertFalse([1 for t, _, _ in recs if t == 84])
        self.assertEqual(plus_types(out).count(0x4014), runs)
        fills_outside = [1 for t, _, d in recs if t == 62 and not d]
        self.assertGreaterEqual(len(fills_outside), runs)
        self.assertEqual(struct.unpack_from('<H', out, 56)[0], struct.unpack_from('<H', src, 56)[0] + 1)

    @unittest.skipUnless(HAVE_FONT, 'Arial Narrow not installed')
    def test_package_roundtrip(self):
        emf = make_demo.make_emf()
        with tempfile.TemporaryDirectory() as d:
            src, dst = os.path.join(d, 'a.docx'), os.path.join(d, 'b.docx')
            with zipfile.ZipFile(src, 'w') as z:
                z.writestr('[Content_Types].xml', '<Types/>')
                z.writestr('word/document.xml', '<doc/>')
                z.writestr('word/media/image1.emf', emf)
            report = fix_package(src, dst, fonts=FONTS)
            self.assertIn('word/media/image1.emf', report)
            with zipfile.ZipFile(src) as a, zipfile.ZipFile(dst) as b:
                self.assertEqual(a.namelist(), b.namelist())
                self.assertEqual(a.read('word/document.xml'), b.read('word/document.xml'))
                self.assertNotEqual(a.read('word/media/image1.emf'), b.read('word/media/image1.emf'))
            with self.assertRaises(ValueError):
                fix_package(src, src, fonts=FONTS)


if __name__ == '__main__':
    unittest.main()
