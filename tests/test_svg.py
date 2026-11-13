import io
import os
import sys
import tempfile
import unittest
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, '..', 'src'))

from visio_mac_rendering_fix import pdf_images            # noqa: E402
from visio_mac_rendering_fix.fonts import FontIndex       # noqa: E402

try:
    try:
        import pymupdf as fitz
    except ImportError:
        import fitz
    from visio_mac_rendering_fix.svg_vector import svg_to_pdf
except ImportError:
    fitz = None

FONTS = FontIndex()
HAVE_FONT = FONTS.get('Arial Narrow', False, False) is not None

SVG = b'''<svg width="300" height="160" viewBox="0 0 150 80" xmlns="http://www.w3.org/2000/svg">
<rect x="0" y="0" width="150" height="80" fill="#FFFFFF"/>
<rect x="30" y="2" width="60" height="20" fill="#E6E6E6"/>
<line x1="0" y1="24" x2="150" y2="24" stroke="#000000" stroke-width="0.5"/>
<text font-family="Arial Narrow,Arial" font-weight="700" font-size="8" text-anchor="middle" x="60" y="15">Instant-CSK</text>
<text font-family="Arial Narrow,Arial" font-size="7" x="5" y="40">D<tspan font-size="5" dy="1.5">norm</tspan><tspan dy="-1.5"> = d / 2</tspan></text>
<text font-family="Arial Narrow,Arial" font-weight="700" font-size="7" text-anchor="middle"
      transform="matrix(0 -1 1 0 -40 70)" x="10" y="60">Uplink</text>
</svg>'''

DOCUMENT = '''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:drawing>
<pic:pic><pic:blipFill><a:blip r:embed="rId1"><a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">
<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId2"/></a:ext></a:extLst>
</a:blip><a:stretch><a:fillRect/></a:stretch></pic:blipFill></pic:pic></w:drawing></w:r></w:p></w:body></w:document>'''

RELS = '''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.svg"/>
</Relationships>'''


@unittest.skipIf(fitz is None or not pdf_images.HAVE_DEPS, 'PyMuPDF / Pillow / numpy not installed')
@unittest.skipUnless(HAVE_FONT, 'Arial Narrow not installed')
class SvgTests(unittest.TestCase):
    def test_svg_to_pdf_uses_real_font(self):
        doc = svg_to_pdf(SVG, FONTS)
        page = doc[0]
        self.assertEqual((round(page.rect.width), round(page.rect.height)), (300, 160))
        fonts = {f[3] for f in page.get_fonts()}
        self.assertTrue(any('Narrow' in f for f in fonts), fonts)
        text = page.get_text()
        for s in ('Instant-CSK', 'Uplink', 'norm'):
            self.assertIn(s, text)
        # rotated label reads upwards and stays inside the picture
        words = {w[4]: w for w in page.get_text('words')}
        up = words['Uplink']
        self.assertGreater(up[3] - up[1], up[2] - up[0])            # taller than wide
        self.assertTrue(page.rect.contains(fitz.Rect(up[:4])))
        # centred label is centred on x = 60 user units = 120 pt
        cx = (words['Instant-CSK'][0] + words['Instant-CSK'][2]) / 2
        self.assertAlmostEqual(cx, 120, delta=2)

    def test_restores_rasterised_svg_as_vector(self):
        fallback = svg_to_pdf(SVG, FONTS)[0].get_pixmap(dpi=144)          # what Office stores as PNG
        png = fallback.tobytes('png')
        blurry = svg_to_pdf(SVG, FONTS)[0].get_pixmap(dpi=60).tobytes('png')  # what Word exports
        with tempfile.TemporaryDirectory() as d:
            src, pdf, out = (os.path.join(d, n) for n in ('a.docx', 'a.pdf', 'a_hires.pdf'))
            with zipfile.ZipFile(src, 'w') as z:
                z.writestr('word/document.xml', DOCUMENT)
                z.writestr('word/_rels/document.xml.rels', RELS)
                z.writestr('word/media/image1.png', png)
                z.writestr('word/media/image2.svg', SVG)
            doc = fitz.open()
            page = doc.new_page()
            rect = fitz.Rect(100, 100, 400, 260)
            page.insert_image(rect, stream=blurry)
            doc.save(pdf)
            recs = pdf_images.restore_pdf_images(pdf, src, out, log=lambda *_: None)
            self.assertEqual([r['new'] for r in recs], ['vector'])
            res = fitz.open(out)[0]
            self.assertIn('Instant-CSK', res.get_text())
            self.assertTrue(all(fitz.Rect(w[:4]).intersects(rect) for w in res.get_text('words')))
            imgs = [res.parent.extract_image(i[0]) for i in res.get_images(full=True)]
            self.assertTrue(all(i['width'] * i['height'] <= 1 for i in imgs))   # raster removed


if __name__ == '__main__':
    unittest.main()
