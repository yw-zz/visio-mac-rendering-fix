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


@unittest.skipIf(fitz is None or not pdf_images.HAVE_DEPS, 'PyMuPDF / Pillow / numpy not installed')
class CommandLineTests(unittest.TestCase):
    def _files(self, d):
        emf = make_demo.make_emf()
        small = make_demo.scope_capture().resize((393, 221), Image.LANCZOS)
        buf = io.BytesIO(); small.save(buf, 'PNG')
        src, pdf = os.path.join(d, 'a.docx'), os.path.join(d, 'a.pdf')
        with zipfile.ZipFile(src, 'w') as z:
            z.writestr('word/media/image1.emf', emf)
        doc = fitz.open(); doc.new_page().insert_image(fitz.Rect(300, 100, 448, 183), stream=buf.getvalue())
        doc.save(pdf)
        return src, pdf

    def test_either_order(self):
        with tempfile.TemporaryDirectory() as d:
            src, pdf = self._files(d)
            self.assertEqual(pdf_images.main([src, pdf, '-q']), 0)
            self.assertTrue(os.path.exists(os.path.join(d, 'a_hires.pdf')))
            out2 = os.path.join(d, 'b.pdf')
            self.assertEqual(pdf_images.main([pdf, src, '-o', out2, '-q']), 0)
            self.assertTrue(os.path.exists(out2))

    def test_bad_arguments(self):
        with tempfile.TemporaryDirectory() as d:
            src, pdf = self._files(d)
            for args in ([pdf], [pdf, pdf], [src, src], [src, pdf, 'x.txt']):
                with self.assertRaises(SystemExit):
                    pdf_images.main(args + ['-q'])


@unittest.skipIf(fitz is None or not pdf_images.HAVE_DEPS, 'PyMuPDF / Pillow / numpy not installed')
class EmbeddedOriginalTests(unittest.TestCase):
    def test_src_box_pixel_centre_convention(self):
        dr = dict(img=Image.new('RGB', (707, 442)), src=(-0.5, -0.5, 707, 442))
        self.assertEqual(pdf_images._src_box(dr), (0, 0, 707, 442))

    def test_uses_larger_original_from_embedded_visio_drawing(self):
        """Visio shrinks pictures in a figure's preview; the full-size original (stored flipped
        relative to the preview here) is inside the embedded .vsdx and should be used."""
        full = make_demo.scope_capture()                                  # 960 x 540
        preview = full.resize((480, 270), Image.LANCZOS).transpose(Image.Transpose.FLIP_TOP_BOTTOM)
        orig_scope = make_demo.scope_capture
        make_demo.scope_capture = lambda: preview
        try:
            emf = make_demo.make_emf()
        finally:
            make_demo.scope_capture = orig_scope
        vsdx = io.BytesIO()
        with zipfile.ZipFile(vsdx, 'w') as v:
            buf = io.BytesIO(); full.save(buf, 'PNG')
            v.writestr('visio/media/image1.png', buf.getvalue())
        blurry = io.BytesIO(); preview.resize((240, 135), Image.LANCZOS).save(blurry, 'PNG')
        with tempfile.TemporaryDirectory() as d:
            src, pdf, out = (os.path.join(d, n) for n in ('a.docx', 'a.pdf', 'a_hires.pdf'))
            with zipfile.ZipFile(src, 'w') as z:
                z.writestr('word/media/image1.emf', emf)
                z.writestr('word/embeddings/Microsoft_Visio_Drawing.vsdx', vsdx.getvalue())
            doc = fitz.open()
            doc.new_page().insert_image(fitz.Rect(100, 100, 420, 280), stream=blurry.getvalue())
            doc.save(pdf)
            recs = pdf_images.restore_pdf_images(pdf, src, out, log=lambda *_: None, max_ppi=0)
            self.assertEqual(len(recs), 1)
            self.assertEqual(recs[0]['new'], (960, 540))
            self.assertIn('vsdx', recs[0]['source'])
            capped = os.path.join(d, 'capped.pdf')
            recs = pdf_images.restore_pdf_images(pdf, src, capped, log=lambda *_: None, max_ppi=150)
            self.assertEqual(recs[0]['new'][0], round(960 * 150 / (960 / (320 / 72))))   # 320 pt wide at 150 ppi


@unittest.skipIf(fitz is None or not pdf_images.HAVE_DEPS, 'PyMuPDF / Pillow / numpy not installed')
class PlainPictureTests(unittest.TestCase):
    def test_restores_plain_png(self):
        full = make_demo.scope_capture()
        buf = io.BytesIO(); full.save(buf, 'PNG')
        small = io.BytesIO(); full.resize((320, 180), Image.LANCZOS).convert('RGB').save(small, 'JPEG', quality=80)
        with tempfile.TemporaryDirectory() as d:
            src, pdf, out = (os.path.join(d, n) for n in ('a.docx', 'a.pdf', 'b.pdf'))
            with zipfile.ZipFile(src, 'w') as z:
                z.writestr('word/media/fig.png', buf.getvalue())
            doc = fitz.open(); doc.new_page().insert_image(fitz.Rect(100, 100, 420, 280), stream=small.getvalue())
            doc.save(pdf)
            recs = pdf_images.restore_pdf_images(pdf, src, out, log=lambda *_: None)
            self.assertEqual([r['new'] for r in recs], [(960, 540)])
            self.assertEqual(recs[0]['encoding'], 'lossless')


@unittest.skipUnless(os.environ.get('VPH_TEST_WORD') == '1',
                     'drives Microsoft Word for Mac; set VPH_TEST_WORD=1 to run')
class WordExportTests(unittest.TestCase):
    def test_export_demo(self):
        from visio_mac_rendering_fix import word_export
        with tempfile.TemporaryDirectory() as d:
            make_demo.main(d)
            pdf = os.path.join(d, 'demo.pdf')
            pages = word_export.export_docx_to_pdf(os.path.join(d, 'demo.docx'), pdf)
            self.assertEqual(pages, 1)
            self.assertEqual(fitz.open(pdf).page_count, 1)


if __name__ == '__main__':
    unittest.main()
