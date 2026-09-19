"""Normalise font family references to the ITQ brand font "Titillium" in a .docx.

Rewrites "Titillium Web" -> "Titillium" wherever Word reads a family name
(document, styles, fontTable, numbering, headers/footers, notes, theme, and the
glossary copies of those). The brand font is Titillium (marketing, 2026-09-18);
"Titillium Web" is Google's derivative cut with different widths and letterforms
and must not appear in ITQ documents.

History: between 2026-09-01 and 2026-09-18 the template was retargeted the other
way, because Word on Windows silently exports Calibri for any CFF/OpenType (.otf)
font. That is now solved at the machine level: build_titillium_ttf.py converts the
licensed Titillium OTFs to TrueType and install_titillium_ttf.ps1 registers them,
so a document that names "Titillium" exports a PDF that embeds Titillium.

Usage:  python retarget_titillium.py path/to/doc.docx
Rewrites the file in place (atomic). Idempotent; "Titillium Bd"/"Titillium Lt"
are never touched.
"""
import os
import re
import sys
import tempfile
import zipfile

# w:ascii= w:hAnsi= w:cs= w:eastAsia= (runs/styles)   w:name= (fontTable)   typeface= (theme)
_RE = re.compile(r'((?:w:ascii|w:hAnsi|w:cs|w:eastAsia|w:name|typeface)=")Titillium Web(")')
FONT_PARTS = re.compile(
    r'^word/(?:glossary/)?(document|styles|fontTable|numbering|header\d*|footer\d*|footnotes|endnotes)\.xml$'
    r'|^word/(?:glossary/)?theme/theme\d+\.xml$')


def retarget(path):
    """Rewrite `path` in place; return the number of references changed."""
    # Build the replacement next to the target so the final swap is an atomic
    # same-filesystem os.replace, never a copy that could leave the original
    # half-overwritten if interrupted.
    target_dir = os.path.dirname(os.path.abspath(path)) or '.'
    fd, tmp = tempfile.mkstemp(suffix='.docx', prefix='.retarget-', dir=target_dir)
    os.close(fd)
    changed = 0
    try:
        with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if FONT_PARTS.match(item.filename):
                    new, n = _RE.subn(r'\1Titillium\2', data.decode('utf-8'))
                    changed += n
                    data = new.encode('utf-8')
                zout.writestr(item, data)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            os.remove(tmp)
    return changed


if __name__ == '__main__':
    if len(sys.argv) != 2:
        sys.exit(__doc__)
    n = retarget(sys.argv[1])
    print('normalised %d font references to "Titillium" in %s' % (n, os.path.basename(sys.argv[1])))
