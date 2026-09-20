"""Build TrueType (.ttf) versions of the ITQ brand font Titillium from its OTFs.

Why: Word on Windows silently exports Calibri for any CFF/OpenType (.otf) font
(verified 2026-09-18 with an unrelated system-wide OTF as control), but handles
TrueType fine. Titillium ships as CFF OTF; Titillium Web (a different typeface) is
TTF, which is why only the Web cut ever survived a Word PDF export here. These
builds keep the real Titillium outlines and every name record, so a document that
names "Titillium" embeds Titillium. install_titillium_ttf.ps1 registers them.

Method follows fontTools' reference Snippets/otf2ttf.py: cubic->quadratic with
Cu2QuPen (max error 1 unit/1000 upem, reversed contour direction), a fresh v1.0
maxp (a CFF-style v0.5 maxp loads in GDI but Word's PDF embedder rejects it and
falls back), left side bearings from xMin, post 2.0 (glyph names), plus a gasp
table. Licence: Titillium is SIL OFL 1.1 (fsType 8 = editable embedding); these
builds are for ITQ's internal use.

    python build_titillium_ttf.py                     # all Titillium-*.otf found in the
                                                      # default locations -> fonts/ttf/
    python build_titillium_ttf.py --src DIR --out DIR # explicit
    python build_titillium_ttf.py --one in.otf out.ttf [--rename FAMILY]
"""
import argparse
import glob
import logging
import os
import sys

# The source Titillium 'fraction' glyph has a 65507-unit advance (a source defect,
# carried over unchanged); fontTools warns about it on every build. Keep stderr to
# real errors so PowerShell callers do not mistake the warning for a failure.
logging.getLogger('fontTools').setLevel(logging.ERROR)

from fontTools.pens.cu2quPen import Cu2QuPen  # noqa: E402
from fontTools.pens.ttGlyphPen import TTGlyphPen
from fontTools.ttLib import TTFont, newTable

HERE = os.path.dirname(os.path.abspath(__file__))
# The official OTFs shipped with this script come first: <dir>/fonts/otf next to the
# script (repo, itq-word-plugin) or one level up (portable skill keeps scripts/ apart).
# Installed fonts and the itq-design skill are fallbacks for a bare checkout.
LOCAL_OTF = [os.path.join(HERE, 'fonts', 'otf'), os.path.join(os.path.dirname(HERE), 'fonts', 'otf')]
DEFAULT_SRC = LOCAL_OTF + [
    os.path.join(os.environ.get('LOCALAPPDATA', ''), 'Microsoft', 'Windows', 'Fonts'),
    os.path.join(os.environ.get('WINDIR', r'C:\Windows'), 'Fonts'),
    os.path.join(os.path.expanduser('~'), '.claude', 'skills', 'itq-design', 'fonts'),
]


def default_out():
    """fonts/ttf beside whichever fonts/otf is shipped with this script."""
    for d in LOCAL_OTF:
        if os.path.isdir(d):
            return os.path.join(os.path.dirname(d), 'ttf')
    return os.path.join(HERE, 'fonts', 'ttf')
MAX_ERR = 1.0  # font units at 1000 upem; visually lossless


def otf_to_ttf(font, rename_family=None):
    assert font.sfntVersion == 'OTTO' and 'CFF ' in font, 'not a CFF OpenType font'
    order = font.getGlyphOrder()
    gs = font.getGlyphSet()
    upem = font['head'].unitsPerEm
    font['loca'] = newTable('loca')
    glyf = font['glyf'] = newTable('glyf')
    glyf.glyphOrder = order
    glyf.glyphs = {}
    for name in order:
        tp = TTGlyphPen(gs)
        gs[name].draw(Cu2QuPen(tp, MAX_ERR * upem / 1000.0, reverse_direction=True))
        glyf.glyphs[name] = tp.glyph()
    del font['CFF ']
    if 'VORG' in font:
        del font['VORG']
    glyf.compile(font)
    hmtx = font['hmtx']
    for name, g in glyf.glyphs.items():
        if hasattr(g, 'xMin'):
            hmtx[name] = (hmtx[name][0], g.xMin)
    maxp = font['maxp'] = newTable('maxp')
    maxp.tableVersion = 0x00010000
    maxp.maxZones = 1
    maxp.maxTwilightPoints = 0
    maxp.maxStorage = 0
    maxp.maxFunctionDefs = 0
    maxp.maxInstructionDefs = 0
    maxp.maxStackElements = 0
    maxp.maxSizeOfInstructions = 0
    maxp.maxComponentElements = max(len(getattr(g, 'components', [])) for g in glyf.glyphs.values())
    maxp.compile(font)
    post = font['post']
    post.formatType = 2.0
    post.extraNames = []
    post.mapping = {}
    post.glyphOrder = order
    try:
        post.compile(font)
    except OverflowError:
        post.formatType = 3
    gasp = font['gasp'] = newTable('gasp')
    gasp.version = 1
    gasp.gaspRange = {65535: 15}
    font.sfntVersion = '\x00\x01\x00\x00'
    if rename_family:
        old = font['name'].getDebugName(1)
        for rec in font['name'].names:
            try:
                v = rec.toUnicode()
            except Exception:
                continue
            if old in v:
                rec.string = v.replace(old, rename_family)
    return font


def convert_one(src, dst, rename_family=None):
    f = otf_to_ttf(TTFont(src), rename_family)
    os.makedirs(os.path.dirname(os.path.abspath(dst)) or '.', exist_ok=True)
    f.save(dst)
    n = f['name']
    return n.getDebugName(1), n.getDebugName(2), n.getDebugName(6)


def find_sources(dirs):
    """Titillium-*.otf in the given dirs; first hit per basename wins."""
    seen = {}
    for d in dirs:
        if not d or not os.path.isdir(d):
            continue
        for p in sorted(glob.glob(os.path.join(d, 'Titillium-*.otf'))):
            seen.setdefault(os.path.basename(p), p)
    return seen


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--src', action='append', metavar='DIR',
                    help='directory holding Titillium-*.otf (repeatable; default: per-user fonts, '
                         'C:\\Windows\\Fonts, the itq-design skill fonts/, ./fonts/otf)')
    ap.add_argument('--out', default=None, metavar='DIR', help='output directory (default: fonts/ttf beside the shipped fonts/otf)')
    ap.add_argument('--one', nargs=2, metavar=('IN.otf', 'OUT.ttf'), help='convert a single file instead')
    ap.add_argument('--rename', metavar='FAMILY', help='with --one: rename the family (probes only)')
    args = ap.parse_args(argv)

    if args.one:
        fam, sub, ps = convert_one(args.one[0], args.one[1], args.rename)
        print('%s  family="%s" style="%s" ps="%s"' % (args.one[1], fam, sub, ps))
        return 0

    sources = find_sources(args.src or DEFAULT_SRC)
    if not sources:
        print('build_titillium_ttf: no Titillium-*.otf found in: %s\n'
              'Ship fonts/otf next to this script (see NOTICE.md) or pass --src DIR.'
              % ', '.join(args.src or DEFAULT_SRC), file=sys.stderr)
        return 1
    out = args.out or default_out()
    for base, src in sources.items():
        dst = os.path.join(out, base[:-4] + '.ttf')
        fam, sub, ps = convert_one(src, dst)
        print('%-32s family="%s" style="%s" ps="%s"  <- %s' % (os.path.basename(dst), fam, sub, ps, src))
    print('%d face(s) written to %s' % (len(sources), os.path.abspath(out)))
    return 0


if __name__ == '__main__':
    sys.exit(main())
