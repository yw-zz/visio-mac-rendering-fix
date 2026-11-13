import os
import sys
import tempfile
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'src'))
sys.path.insert(0, os.path.join(HERE, '..', 'examples'))

from visio_mac_rendering_fix import FontIndex, fix_emf, unfix_emf, is_fixed_emf, fix_package   # noqa: E402
from visio_mac_rendering_fix.emf import strip_markers                                         # noqa: E402
from visio_mac_rendering_fix.package import unfix_package                                     # noqa: E402
import make_demo                                                                               # noqa: E402

FONTS = FontIndex()
HAVE_FONT = FONTS.get('Arial Narrow', False, False) is not None


@unittest.skipUnless(HAVE_FONT, 'Arial Narrow not installed')
class UndoTests(unittest.TestCase):
    def setUp(self):
        self.orig = make_demo.make_emf()
        self.fixed = fix_emf(self.orig, FONTS)

    def test_undo_is_exact(self):
        self.assertTrue(is_fixed_emf(self.fixed))
        self.assertFalse(is_fixed_emf(self.orig))
        self.assertEqual(unfix_emf(self.fixed), self.orig)
        self.assertEqual(unfix_emf(self.orig), self.orig)          # no undo data: unchanged

    def test_fixing_twice(self):
        self.assertEqual(fix_emf(self.fixed, FONTS), self.fixed)    # idempotent
        self.assertEqual(unfix_emf(fix_emf(self.fixed, FONTS)), self.orig)

    def _docx(self, d, name, emf):
        p = os.path.join(d, name)
        with zipfile.ZipFile(p, 'w') as z:
            z.writestr('word/document.xml', '<doc/>')
            z.writestr('word/media/image1.emf', emf)
        return p

    def test_package_undo(self):
        with tempfile.TemporaryDirectory() as d:
            src = self._docx(d, 'a.docx', self.orig)
            fixed, back = os.path.join(d, 'b.docx'), os.path.join(d, 'c.docx')
            fix_package(src, fixed, fonts=FONTS)
            self.assertEqual(unfix_package(fixed, back), {'word/media/image1.emf': 'undone'})
            with zipfile.ZipFile(src) as a, zipfile.ZipFile(back) as b:
                self.assertEqual(a.read('word/media/image1.emf'), b.read('word/media/image1.emf'))

    def test_recover_old_style_conversion(self):
        """Files converted by <= 0.4 have no undo data; an older copy of the document recovers them."""
        legacy = strip_markers(self.fixed)
        self.assertFalse(is_fixed_emf(legacy))
        with tempfile.TemporaryDirectory() as d:
            converted = self._docx(d, 'v6.docx', legacy)
            older = self._docx(d, 'v5.docx', self.orig)
            out = os.path.join(d, 'out.docx')
            self.assertEqual(unfix_package(converted, out), {'word/media/image1.emf': 'converted, no original found'})
            self.assertEqual(unfix_package(converted, out, originals=[older]),
                             {'word/media/image1.emf': 'recovered'})
            with zipfile.ZipFile(out) as z:
                self.assertEqual(z.read('word/media/image1.emf'), self.orig)


if __name__ == '__main__':
    unittest.main()
