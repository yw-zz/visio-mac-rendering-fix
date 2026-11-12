import io
import os
import sys
import tempfile
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'src'))
sys.path.insert(0, os.path.join(HERE, '..', 'examples'))

from visio_mac_rendering_fix import pdf_images   # noqa: E402

try:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    from PIL import Image
    import make_demo
except ImportError:              # optional [pdf] extra not installed
    fitz = None


@unittest.skipIf(fitz is None or not pdf_images.HAVE_DEPS, 'PyMuPDF / Pillow / numpy not installed')
class RestoreTests(unittest.TestCase):
    def test_finds_original_image(self):
        draws = pdf_images.emfplus_draws(make_demo.make_emf())
        self.assertEqual(len(draws), 1)
        self.assertEqual(draws[0]['img'].size, (960, 540))

    def test_restores_downsampled_image(self):
        """A PDF holding a 200-ppi JPEG copy of the figure's scope capture gets the original back."""
        emf = make_demo.make_emf()
        full = make_demo.scope_capture()
        small = full.resize((393, 221), Image.LANCZOS)
        buf = io.BytesIO(); small.save(buf, 'JPEG', quality=80)
        with tempfile.TemporaryDirectory() as d:
            src, pdf, out = (os.path.join(d, n) for n in ('a.docx', 'a.pdf', 'a_hires.pdf'))
            with zipfile.ZipFile(src, 'w') as z:
                z.writestr('word/media/image1.emf', emf)
            doc = fitz.open()
            page = doc.new_page()
            page.insert_image(fitz.Rect(300, 100, 448, 183), stream=buf.getvalue())
            page.insert_text((72, 72), 'Fig. 1')
            placed = page.get_image_rects(page.get_images(full=True)[0][0])[0]
            doc.save(pdf)
            records = pdf_images.restore_pdf_images(pdf, src, out, log=lambda *_: None)
            self.assertEqual(len(records), 1)
            self.assertEqual(records[0]['new'], (960, 540))
            res = fitz.open(out)
            xref = res[0].get_images(full=True)[0][0]
            info = res.extract_image(xref)
            self.assertEqual((info['width'], info['height']), (960, 540))
            self.assertEqual(res[0].get_image_rects(xref)[0], placed)          # placement unchanged
            self.assertIn('Fig. 1', res[0].get_text())
            with self.assertRaises(ValueError):
                pdf_images.restore_pdf_images(pdf, src, pdf)

    def test_leaves_unrelated_images(self):
        emf = make_demo.make_emf()
        other = Image.effect_noise((300, 200), 60).convert('RGB')
        buf = io.BytesIO(); other.save(buf, 'PNG')
        with tempfile.TemporaryDirectory() as d:
            src, pdf, out = (os.path.join(d, n) for n in ('a.docx', 'a.pdf', 'b.pdf'))
            with zipfile.ZipFile(src, 'w') as z:
                z.writestr('word/media/image1.emf', emf)
            doc = fitz.open(); doc.new_page().insert_image(fitz.Rect(50, 50, 200, 150), stream=buf.getvalue())
            doc.save(pdf)
            self.assertEqual(pdf_images.restore_pdf_images(pdf, src, out, log=lambda *_: None), [])


if __name__ == '__main__':
    unittest.main()
