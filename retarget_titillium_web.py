"""Retarget font family "Titillium" -> "Titillium Web" inside a rendered .docx.

Why: on some Windows machines Word silently lays out and exports Calibri for the
CFF-based "Titillium" OTFs even though they are installed (see itq-word-render
skill, "Fonts are NOT embedded"). The TTF family "Titillium Web" renders and
embeds correctly and is explicitly allowed by the ITQ brand rules.

Usage:  python retarget_titillium_web.py "path\\to\\doc.docx"
Rewrites the file in place; run finalize.py afterwards to re-bake TOC + PDF.
Idempotent: "Titillium Web", "Titillium Bd", "Titillium Lt" are never double-hit.
"""
import re
import sys
import zipfile
import tempfile
import os

# w:ascii="Titillium"  w:hAnsi="Titillium"  w:cs="Titillium"  w:eastAsia="Titillium"
# w:name="Titillium" (fontTable)  typeface="Titillium" (theme)
_RE = re.compile(r'((?:w:ascii|w:hAnsi|w:cs|w:eastAsia|w:name|typeface)=")Titillium(")')
_PARTS = re.compile(r'^word/(document|styles|fontTable|numbering|header\d*|footer\d*|footnotes|endnotes)\.xml$|^word/theme/theme\d+\.xml$')


def retarget(path):
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
                if _PARTS.match(item.filename):
                    text = data.decode('utf-8')
                    new, n = _RE.subn(r'\1Titillium Web\2', text)
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
    print(f'retargeted {n} font references to "Titillium Web" in {os.path.basename(sys.argv[1])}')
