import argparse
import os
import sys

from . import __version__
from .fonts import FontIndex
from .package import SUPPORTED, default_output, fix_package


def main(argv=None):
    p = argparse.ArgumentParser(
        prog='visio-mac-rendering-fix',
        description='Fix Visio (EMF+) figures that render wrong in Office for Mac '
                    '(italic labels, solid blocks instead of transparency). '
                    'Writes a new file; the input is never modified.')
    p.add_argument('input', nargs='+', help='.docx / .pptx / .xlsx files (and macro/template variants)')
    p.add_argument('-o', '--output', help='output path (only with a single input; '
                   'default: <name>_macfix.<ext> next to the input)')
    p.add_argument('--font-dir', action='append', default=None, metavar='DIR',
                   help='folder to search for font files (repeatable; replaces the OS defaults)')
    p.add_argument('--no-text', action='store_true', help='leave text runs as text')
    p.add_argument('--no-transparency', action='store_true', help='leave stipple transparency as is')
    p.add_argument('--no-glyph-index', action='store_true',
                   help='skip text stored as glyph indices (use if such labels come out as wrong characters)')
    p.add_argument('-q', '--quiet', action='store_true')
    p.add_argument('--version', action='version', version='%(prog)s ' + __version__)
    a = p.parse_args(argv)

    if a.output and len(a.input) > 1:
        p.error('-o/--output works with a single input file')
    fonts = FontIndex(a.font_dir)
    status = 0
    for src in a.input:
        if not src.lower().endswith(SUPPORTED):
            print(f'skip {src}: not an Office Open XML file ({", ".join(SUPPORTED)})', file=sys.stderr)
            status = 1
            continue
        dst = a.output or default_output(src)
        try:
            report = fix_package(src, dst, fonts=fonts, text=not a.no_text,
                                 transparency=not a.no_transparency, glyph_index=not a.no_glyph_index)
        except Exception as e:                       # keep going with other files
            print(f'error {src}: {e}', file=sys.stderr)
            status = 1
            continue
        if not a.quiet:
            missing = set()
            for part, st in sorted(report.items()):
                missing |= st.pop('fonts_not_converted', set())
                print(f'  {part}: ' + ', '.join(f'{k}={v}' for k, v in st.items()) if st else f'  {part}: unchanged')
            if not report:
                print(f'{src}: no dual EMF+ images found')
            if missing:
                print('  fonts not found (left as text): ' + ', '.join(sorted(missing)))
            print(f'wrote {dst}')
    return status


if __name__ == '__main__':
    sys.exit(main())
