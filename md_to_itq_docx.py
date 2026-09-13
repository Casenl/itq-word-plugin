"""
Convert a markdown file into an ITQ-branded Word document.

Works with:
  - Notion's enhanced-markdown export (as produced by the Notion MCP server)
  - Standard markdown (GitHub READMEs, hand-written notes, etc.)

Handles:
  # / ## / ### headings  (# -> Heading 1, ## -> Heading 2; --heading-offset
                          shifts this. Manual numbers in the heading text are
                          kept and the style's auto-numbering is suppressed so
                          they don't double up.)
  **bold**, *italic*, `inline code`
  - bullets          (ITQ Bullets via numId 15)
  1. numbered items  (ITQ Numbering via numId 18)
  <callout icon="..." color="..."> ... </callout>         Notion-enhanced only
  <table [header-row="true"]> <tr><td>..</td>..</tr> ... </table>  Notion-enhanced
  | GitHub-style | pipe tables |      first row = header; backslash-pipe escapes a pipe
  ``` fenced code blocks ```          (mermaid rendered via kroki.io)
  > blockquote
  --- horizontal rule
  ![Caption](path.png)                 a paragraph that is ONLY an image: centred
                                       picture (max 6.2in wide, natural width if
                                       smaller at 96 dpi), alt text as caption.
                                       png/jpg/jpeg; path relative to the .md
                                       file (fallback: cwd). Missing files warn
                                       and leave a "[Figure missing: ...]" note.

Not yet supported: inline [links](url); images INSIDE a text paragraph are
dropped with a warning (file:line) rather than rendered.

After python-docx saves, a post-processing pass patches the zip to:
  - populate the cover-page SDT content controls (Client/Project/Subtitle/
    Title/Version) that are bound to customXml/item2.xml
  - update docProps/core.xml title and creator
  - clear <w:showingPlcHdr/> and overwrite the cached SDT text so the cover
    renders correctly without needing Word to refresh the binding
  - strip emoji characters from every visible text run
"""
import io
import os
import re
import sys
import shutil
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from copy import deepcopy

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.image.image import Image as DocxImage
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm, Emu, Inches


# Renderer version. Bump on every behaviour change so that a document rendered
# by a stale bundled copy is diagnosable from the file itself: it is written
# into docProps/custom.xml as the ITQRendererVersion custom property, which
# Word preserves across save and PDF export.
__version__ = '1.5.1'


def warn(msg):
    """Non-fatal diagnostics go to stderr so the render still completes and
    stdout keeps carrying only the 'wrote ...' line."""
    print('warning: ' + msg, file=sys.stderr)


# ---------- Emoji stripping ---------------------------------------------------

# Broad emoji / pictograph ranges. We also strip variation selectors (FE0F) and
# zero-width joiners so residue doesn't remain after emoji characters vanish.
_EMOJI_RE = re.compile(
    '['
    '\U0001F300-\U0001FAFF'   # Supplemental Symbols, Pictographs, etc.
    '\U0001F900-\U0001F9FF'   # Supplemental Symbols and Pictographs (legacy)
    '\U0001F600-\U0001F64F'   # Emoticons
    '\U0001F680-\U0001F6FF'   # Transport and Map
    '\U0001F700-\U0001F77F'
    '\U0001F780-\U0001F7FF'
    '\U0001F800-\U0001F8FF'
    '\U00002600-\U000027BF'   # Misc symbols + Dingbats
    '\U0001F1E6-\U0001F1FF'   # Regional indicator (flags)
    '\U0000FE00-\U0000FE0F'   # Variation selectors
    '\U0000200D'              # ZWJ
    ']+',
    flags=re.UNICODE,
)


def strip_emoji(text, trim=False):
    """Remove emoji characters from text.

    By default preserves leading/trailing whitespace — critical when called
    on individual <w:t> runs, where a leading space carries the break between
    a bold run and the following plain run. Set trim=True for contexts that
    operate on a whole logical field value (headings, SDT values, etc.).
    """
    if not text:
        return text
    cleaned = _EMOJI_RE.sub('', text)
    # Collapse runs of whitespace only if emoji removal introduced them.
    if text != cleaned:
        cleaned = re.sub(r'[ \t]{2,}', ' ', cleaned)
    if trim:
        cleaned = cleaned.strip()
    return cleaned


# ---------- Block parser ------------------------------------------------------

class Block:
    def __init__(self, kind, **kw):
        self.kind = kind
        self.__dict__.update(kw)


def _split_table(raw):
    rows = []
    row_re = re.compile(r'<tr>(.*?)</tr>', re.DOTALL)
    cell_re = re.compile(r'<td>(.*?)</td>', re.DOTALL)
    for m in row_re.finditer(raw):
        cells = [c.strip() for c in cell_re.findall(m.group(1))]
        rows.append(cells)
    return rows


# GitHub-flavoured pipe tables. The delimiter row (---, :---, ---:, :---:)
# marks the line above as the header. A leading/trailing pipe is optional;
# an escaped pipe (\|) stays inside its cell and is unescaped by unescape_md
# at render time, so it needs no special handling here.
_PIPE_SEP_RE = re.compile(r'^\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$')
_PIPE_SPLIT_RE = re.compile(r'(?<!\\)\|')


def _split_pipe_row(line):
    s = line.strip()
    if s.startswith('|'):
        s = s[1:]
    if s.endswith('|') and not s.endswith('\\|'):
        s = s[:-1]
    return [c.strip() for c in _PIPE_SPLIT_RE.split(s)]


def _has_pipe(line):
    return bool(_PIPE_SPLIT_RE.search(line))


_BLOCK_START_RE = re.compile(r'^(#{1,6}\s|[-*]\s|\d+\.\s|>|```|---\s*$|<callout|<table)')


def _is_block_start(stripped):
    """True when a line opens another block-level construct, which ends a
    pipe table even if the line itself happens to contain a pipe."""
    return bool(_BLOCK_START_RE.match(stripped))


def _is_pipe_table_start(lines, i):
    # Header row must contain an unescaped pipe, and the delimiter row must
    # contain BOTH a dash run and a pipe, so a bare '---' rule after a prose
    # line that happens to contain '|' is never mistaken for a table.
    if not _has_pipe(lines[i]) or i + 1 >= len(lines):
        return False
    nxt = lines[i + 1].strip()
    if not ('|' in nxt and '-' in nxt and _PIPE_SEP_RE.match(nxt)):
        return False
    # GFM: the delimiter row must have exactly as many cells as the header,
    # otherwise the lines are ordinary paragraphs.
    return len(_split_pipe_row(nxt)) == len(_split_pipe_row(lines[i]))


# Images. A paragraph consisting of nothing but `![alt](target)` becomes an
# image block; an image embedded in running text is dropped (with a warning)
# because inline pictures have no sensible layout in the house template. A
# backslash-escaped `\![` is Notion's way of saying "literal text" and is left
# for unescape_md; image syntax quoted inside an inline-code span is code.
_CODE_SPAN_RE = re.compile(r'`[^`\n]+`')
# `<path with spaces>` or `path "optional title"`; CommonMark also allows a
# title in single quotes or parentheses: `path (title)`.
_IMAGE_TARGET_RE = re.compile(r'^(?:<([^>]*)>|(\S+))(?:\s+("[^"]*"|\'[^\']*\'|\([^)]*\)))?$')


def _scan_balanced(text, start, open_ch, close_ch):
    """Index of the bracket closing the one at `start`, honouring nesting
    and backslash-escaped brackets (e.g. alt text `![a \\] b](x.png)`
    keeps its escaped `]` from ending the label early); -1 when
    unterminated."""
    depth = 0
    k = start
    n = len(text)
    while k < n:
        c = text[k]
        if c == '\\' and k + 1 < n:
            k += 2
            continue
        if c == open_ch:
            depth += 1
        elif c == close_ch:
            depth -= 1
            if depth == 0:
                return k
        k += 1
    return -1


def _scan_image_dest_close(text, start):
    """Index of the ')' closing the destination/title group opened by the
    '(' at `start` (as `_scan_balanced` does for '(' / ')'), but also
    honouring the inline-link syntax that can legally appear inside that
    group: a backslash escape (`\\)`), an angle-bracketed destination
    (`<a).b>`), and a quoted title (`"a)b"` / `'a)b'`) can all contain a
    literal ')' that must NOT be mistaken for the group's own close --
    unlike plain nested parens (`run(1).png`), which still need counting.
    -1 when unterminated."""
    depth = 0
    i = start
    n = len(text)
    while i < n:
        c = text[i]
        if c == '\\' and i + 1 < n:
            i += 2
            continue
        if c == '(':
            depth += 1
        elif c == ')':
            depth -= 1
            if depth == 0:
                return i
        elif c == '<':
            end = text.find('>', i + 1)
            if end >= 0:
                i = end + 1
                continue
        elif c in ('"', "'") and text[i - 1] in (' ', '\t'):
            # Markdown only permits a title after whitespace following the
            # destination, so a quote/apostrophe anywhere else (e.g. inside
            # an unquoted destination like "user's.png") is just a literal
            # character, not the start of a quoted span to jump over.
            end = text.find(c, i + 1)
            if end >= 0:
                i = end + 1
                continue
        i += 1
    return -1


def find_image_tokens(text):
    r"""Every `![alt](dest)` in `text` as (start, end, alt, dest).

    Alt text may contain balanced square brackets ("See [R-57]"), the
    destination balanced parentheses ("run(1).png"). Escaped `\![` and
    anything inside an inline-code span are skipped."""
    tokens = []
    code_spans = [(m.start(), m.end()) for m in _CODE_SPAN_RE.finditer(text)]
    i = 0
    while True:
        i = text.find('![', i)
        if i < 0:
            break
        if (i > 0 and text[i - 1] == '\\') or any(a <= i < b for a, b in code_spans):
            i += 2
            continue
        j = _scan_balanced(text, i + 1, '[', ']')
        if j < 0 or j + 1 >= len(text) or text[j + 1] != '(':
            i += 2
            continue
        k = _scan_image_dest_close(text, j + 1)
        if k < 0:
            i += 2
            continue
        tokens.append((i, k + 1, text[i + 2:j], text[j + 2:k]))
        i = k + 1
    return tokens


def _whole_line_image(stripped):
    """(alt, dest) when the line is exactly one image token, else None."""
    if not stripped.startswith('!['):
        return None
    toks = find_image_tokens(stripped)
    if len(toks) != 1:
        return None
    start, end, alt, dest = toks[0]
    if start != 0 or end != len(stripped) or not dest.strip():
        return None
    return alt, dest


_MD_ESCAPE_RE = re.compile(r'\\([!"#$%&\'()*+,\-./:;<=>?@\[\\\]^_`{|}~])')


def _unescape_md(text):
    """Drop the backslash from CommonMark escapes (backslash + ASCII
    punctuation). The scanner keeps the escape so `plot\\).png` is not cut
    at the `)`; the filesystem lookup needs the literal `plot).png`."""
    return _MD_ESCAPE_RE.sub(r'\1', text)


def _parse_image_target(raw):
    raw = raw.strip()
    m = _IMAGE_TARGET_RE.match(raw)
    if not m:
        return _unescape_md(raw)  # e.g. an unbracketed path with spaces and no title
    return _unescape_md(m.group(1) if m.group(1) is not None else m.group(2))


def _drop_inline_images(text, lineno, source_name, alt_fallback=False):
    """Remove `![..](..)` occurrences from running text, warning per image.

    With `alt_fallback`, a text that consists of nothing but images keeps
    the images' alt text instead of becoming empty: a heading or list item
    such as `# ![Logo](logo.png)` then reads "Logo" rather than producing an
    empty heading (and an empty TOC entry). Callers still drop the block
    when the result is empty."""
    if '![' not in text:
        return text
    toks = find_image_tokens(text)
    if not toks:
        return text
    out = []
    alts = []
    pos = 0
    for start, end, alt, _dest in toks:
        warn('%s:%d: inline image dropped (only whole-paragraph images render): %s'
             % (source_name, lineno, text[start:end]))
        out.append(text[pos:start])
        alts.append(alt.strip())
        pos = end
    out.append(text[pos:])
    cleaned = re.sub(r'[ \t]{2,}', ' ', ''.join(out)).strip()
    if not cleaned and alt_fallback:
        cleaned = ' '.join(a for a in alts if a)
        if cleaned:
            warn('%s:%d: kept the alt text as the block text: %s' % (source_name, lineno, cleaned))
    return cleaned


def parse_blocks(text, source_name='<markdown>'):
    lines = text.splitlines()
    i, n = 0, len(lines)
    blocks = []

    while i < n:
        line = lines[i]
        stripped = line.strip()

        if not stripped or stripped == '<empty-block/>':
            i += 1
            continue

        if stripped == '---':
            blocks.append(Block('hr')); i += 1; continue

        if stripped.startswith('```'):
            lang = stripped[3:].strip()
            i += 1
            buf = []
            while i < n and not lines[i].strip().startswith('```'):
                buf.append(lines[i]); i += 1
            i += 1
            blocks.append(Block('code', lang=lang, text='\n'.join(buf)))
            continue

        if stripped.startswith('<callout'):
            icon_m = re.search(r'icon="([^"]*)"', stripped)
            color_m = re.search(r'color="([^"]*)"', stripped)
            icon = icon_m.group(1) if icon_m else ''
            color = color_m.group(1) if color_m else ''
            i += 1
            buf = []
            while i < n and lines[i].strip() != '</callout>':
                buf.append(lines[i].lstrip('\t ')); i += 1
            i += 1
            blocks.append(Block('callout', icon=icon, color=color,
                                text='\n'.join(buf).strip()))
            continue

        if stripped.startswith('<table'):
            header_row = 'header-row="true"' in stripped
            i += 1
            buf = []
            while i < n and lines[i].strip() != '</table>':
                buf.append(lines[i]); i += 1
            i += 1
            rows = _split_table('\n'.join(buf))
            if rows:
                blocks.append(Block('table', header=header_row, rows=rows))
            continue

        if _is_pipe_table_start(lines, i):
            header = _split_pipe_row(lines[i])
            width = len(header)
            rows = [header]
            i += 2  # skip the header row and the one syntactic delimiter row
            while i < n:
                s = lines[i].strip()
                if not s or not _has_pipe(s) or _is_block_start(s):
                    break
                cells = _split_pipe_row(s)
                # GFM: table width is fixed by the header; excess cells are
                # ignored, missing cells are padded.
                cells = (cells + [''] * width)[:width]
                rows.append(cells)
                i += 1
            blocks.append(Block('table', header=True, rows=rows))
            continue

        m = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if m:
            text = _drop_inline_images(m.group(2).strip(), i + 1, source_name,
                                       alt_fallback=True)
            if text:
                blocks.append(Block('heading', level=len(m.group(1)), text=text))
            else:
                warn('%s:%d: heading emptied by image removal, skipped' % (source_name, i + 1))
            i += 1
            continue

        if stripped.startswith('>'):
            buf = []
            while i < n and lines[i].strip().startswith('>'):
                raw = lines[i].strip().lstrip('>').lstrip()
                cleaned = _drop_inline_images(raw, i + 1, source_name, alt_fallback=True)
                # Drop only lines that image removal emptied; an intentional
                # blank quoted line (`>`) keeps the paragraph break it marks.
                if cleaned or not raw:
                    buf.append(cleaned)
                i += 1
            quote = '\n'.join(buf).strip()
            if quote:
                blocks.append(Block('blockquote', text=quote))
            continue

        m = re.match(r'^(\d+)\.\s+(.*)$', stripped)
        if m:
            items = []
            while i < n:
                mm = re.match(r'^(\d+)\.\s+(.*)$', lines[i].strip())
                if not mm: break
                items.append(_drop_inline_images(mm.group(2), i + 1, source_name,
                                                 alt_fallback=True))
                i += 1
            items = [it for it in items if it]
            if items:
                blocks.append(Block('olist', items=items))
            continue

        if stripped.startswith('- ') or stripped.startswith('* '):
            items = []
            while i < n:
                s = lines[i].strip()
                if s.startswith('- ') or s.startswith('* '):
                    items.append(_drop_inline_images(s[2:], i + 1, source_name,
                                                     alt_fallback=True)); i += 1
                else:
                    break
            items = [it for it in items if it]
            if items:
                blocks.append(Block('ulist', items=items))
            continue

        img = _whole_line_image(stripped)
        if img:
            blocks.append(Block('image', alt=img[0].strip(),
                                target=_parse_image_target(img[1]),
                                line=i + 1))
            i += 1
            continue

        # Paragraph: in Notion-flavoured markdown each non-empty line is its
        # own logical paragraph (no soft-wrapping), so emit one block per line.
        para = _drop_inline_images(stripped, i + 1, source_name)
        if para:
            blocks.append(Block('paragraph', text=para))
        i += 1

    return blocks


# ---------- Inline formatting -------------------------------------------------

_INLINE_RE = re.compile(r'(\*\*[^*]+\*\*|\*[^*\n]+\*|`[^`\n]+`)')

# Notion export escapes non-formatting characters with a backslash (e.g. \$,
# \~, \<, \[) so other markdown processors don't mis-parse them. None of
# these collide with our inline formatting markers (*, _, `), so we can
# safely unescape them before running the inline regex.
_ESCAPE_RE = re.compile(r'\\([$~<>|#!&{}()\[\]+\-.:\\])')


def unescape_md(text):
    return _ESCAPE_RE.sub(r'\1', text)


def render_runs(paragraph, text, base_bold=False):
    text = strip_emoji(text, trim=True)
    text = unescape_md(text)
    pos = 0
    for m in _INLINE_RE.finditer(text):
        if m.start() > pos:
            r = paragraph.add_run(text[pos:m.start()])
            if base_bold: r.bold = True
        token = m.group(0)
        if token.startswith('**'):
            r = paragraph.add_run(token[2:-2]); r.bold = True
        elif token.startswith('`'):
            r = paragraph.add_run(token[1:-1])
            r.font.name = 'Consolas'
        else:
            r = paragraph.add_run(token[1:-1])
            r.italic = True
            if base_bold: r.bold = True
        pos = m.end()
    if pos < len(text):
        r = paragraph.add_run(text[pos:])
        if base_bold: r.bold = True


# ---------- Cell / shading helpers -------------------------------------------

CALLOUT_COLORS = {
    'red_bg':    'F8D7DA',
    'blue_bg':   'D7E7F8',
    'yellow_bg': 'FFF3CD',
    'green_bg':  'D1E7DD',
    'orange_bg': 'FFE5CC',
    'gray_bg':   'E7E7E7',
    'grey_bg':   'E7E7E7',
}


def shade_cell(cell, hex_fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_fill)
    tcPr.append(shd)


def set_cell_border(cell, color='BFBFBF', sz='4'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = OxmlElement('w:tcBorders')
    for side in ('top', 'left', 'bottom', 'right'):
        b = OxmlElement(f'w:{side}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), sz)
        b.set(qn('w:color'), color)
        tcBorders.append(b)
    tcPr.append(tcBorders)


# ---------- Kroki mermaid rendering ------------------------------------------

KROKI_URL = 'https://kroki.io/mermaid/png'


def render_mermaid(source, out_path):
    """POST mermaid source to kroki.io and save the returned PNG."""
    req = urllib.request.Request(
        KROKI_URL,
        data=source.encode('utf-8'),
        headers={
            'Content-Type': 'text/plain',
            # kroki.io is behind Cloudflare, which 403-blocks non-browser
            # User-Agents (error 1010). Present a browser UA.
            'User-Agent': ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                           'AppleWebKit/537.36 (KHTML, like Gecko) '
                           'Chrome/124.0.0.0 Safari/537.36'),
            'Accept': 'image/png',
        },
    )
    with urllib.request.urlopen(req, timeout=45) as resp:
        data = resp.read()
    if not data.startswith(b'\x89PNG'):
        raise RuntimeError(f'kroki returned non-PNG response ({len(data)} bytes)')
    with open(out_path, 'wb') as f:
        f.write(data)
    return out_path


# ---------- Renderer ----------------------------------------------------------

IMAGE_MAX_WIDTH_IN = 6.2      # matches the mermaid figure width
IMAGE_SCREEN_DPI = 96.0       # "natural size" of a screenshot / exported PNG
IMAGE_FORMATS = ('.png', '.jpg', '.jpeg')
ITQ_ROYAL_BLUE = RGBColor(0x04, 0x05, 0x5A)


def resolve_image_path(target, source_dir=None):
    """Turn a markdown image target into a filesystem path.

    Relative targets resolve against the markdown file's directory first and
    the current working directory second. A URL-encoded form (Notion exports
    write `Screen%20Shot.png`) is tried as a fallback. Returns the first
    candidate that exists, else the primary candidate (for the message).
    """
    variants = [target]
    unquoted = urllib.parse.unquote(target)
    if unquoted != target:
        variants.append(unquoted)
    if os.path.isabs(target):
        bases = [None]
    else:
        bases = ([source_dir] if source_dir else []) + [os.getcwd()]
    candidates = []
    for base in bases:
        for v in variants:
            candidates.append(v if base is None else os.path.join(base, v))
    for c in candidates:
        if os.path.isfile(c):
            return c
    return candidates[0]


class ItqRenderer:
    def __init__(self, template_path, asset_dir, source_dir=None,
                 source_name='<markdown>'):
        self.d = Document(template_path)
        self.asset_dir = asset_dir
        self.source_dir = source_dir
        self.source_name = source_name
        os.makedirs(asset_dir, exist_ok=True)
        self._mermaid_seq = 0
        self._discover_list_numids()

    # --- list numbering detection --------------------------------------------

    def _discover_list_numids(self):
        """Find the numIds whose abstractNum carries the ITQBullets /
        ITQNumbering styleLink and actual <w:lvl> definitions."""
        self.bullet_numid = None
        self.num_numid = None
        numbering = self.d.part.numbering_part.element
        w = numbering.nsmap['w']

        # Step 1: abstractNumId -> styleLink name (only for abstractNums that
        # actually define levels, not numStyleLink redirects).
        style_to_absid = {}
        for abs_num in numbering.findall(qn('w:abstractNum')):
            abs_id = abs_num.get(qn('w:abstractNumId'))
            if abs_num.find(qn('w:lvl')) is None:
                continue
            style_link = abs_num.find(qn('w:styleLink'))
            if style_link is None:
                continue
            name = style_link.get(qn('w:val'))
            if name:
                # Normalise — the template uses "ITQBullets" (no space)
                style_to_absid[name.replace(' ', '')] = abs_id

        # Step 2: find the first numId that targets each abstract num
        bullet_abs = style_to_absid.get('ITQBullets')
        num_abs = style_to_absid.get('ITQNumbering')

        for num in numbering.findall(qn('w:num')):
            nid = num.get(qn('w:numId'))
            abs_ref = num.find(qn('w:abstractNumId'))
            if abs_ref is None: continue
            target = abs_ref.get(qn('w:val'))
            if target == bullet_abs and self.bullet_numid is None:
                self.bullet_numid = nid
            if target == num_abs and self.num_numid is None:
                self.num_numid = nid

    def _apply_list(self, para, numid):
        pPr = para._p.get_or_add_pPr()
        numPr = OxmlElement('w:numPr')
        ilvl = OxmlElement('w:ilvl'); ilvl.set(qn('w:val'), '0'); numPr.append(ilvl)
        num = OxmlElement('w:numId'); num.set(qn('w:val'), str(numid)); numPr.append(num)
        pPr.append(numPr)

    # --- body helpers --------------------------------------------------------

    def _suppress_numbering(self, para):
        """Override the Heading style's inherited list numbering on this
        paragraph (numId 0 = 'no numbering'). Lets headings keep the manual
        numbers authored in the markdown without the template's auto-numbering
        doubling them up (e.g. '3 1. About ITQ'). outlineLvl is unaffected, so
        the TOC still nests correctly."""
        pPr = para._p.get_or_add_pPr()
        numPr = OxmlElement('w:numPr')
        numId = OxmlElement('w:numId'); numId.set(qn('w:val'), '0')
        numPr.append(numId)
        pPr.append(numPr)

    def add_heading(self, level, text):
        """`level` is the target Word heading level (1-9) directly."""
        mapped = min(9, max(1, level))
        p = self.d.add_paragraph(style=f'Heading {mapped}')
        render_runs(p, text)
        self._suppress_numbering(p)

    def add_paragraph(self, text):
        p = self.d.add_paragraph()
        # The template's Normal style sets space-after to 0, so consecutive
        # body paragraphs would render with no separation. Give each a small
        # space-after so authored paragraph breaks read as breaks.
        p.paragraph_format.space_after = Pt(8)
        render_runs(p, text)

    def add_blockquote(self, text):
        style = 'Quote' if 'Quote' in [s.name for s in self.d.styles] else 'Normal'
        p = self.d.add_paragraph(style=style)
        render_runs(p, text)
        for r in p.runs:
            r.italic = True

    def add_bullet(self, text):
        p = self.d.add_paragraph(style='List Paragraph')
        render_runs(p, text)
        if self.bullet_numid is not None:
            self._apply_list(p, self.bullet_numid)

    def add_numbered(self, text):
        p = self.d.add_paragraph(style='List Paragraph')
        render_runs(p, text)
        if self.num_numid is not None:
            self._apply_list(p, self.num_numid)

    def add_hr(self):
        p = self.d.add_paragraph()
        pPr = p._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single')
        bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1')
        bottom.set(qn('w:color'), '9A9A9A')
        pBdr.append(bottom)
        pPr.append(pBdr)

    def add_code(self, text, lang=''):
        lang_norm = (lang or '').strip().lower()

        # Render mermaid diagrams as images via kroki
        if lang_norm == 'mermaid':
            self._mermaid_seq += 1
            out = os.path.join(self.asset_dir, f'mermaid-{self._mermaid_seq}.png')
            try:
                render_mermaid(text, out)
                self.d.add_picture(out, width=Inches(6.2))
                return
            except Exception as e:
                # Fall through to code block rendering if kroki fails
                note = self.d.add_paragraph()
                nr = note.add_run(f'[mermaid render failed: {e}]')
                nr.italic = True
                nr.font.color.rgb = RGBColor(0xC0, 0x00, 0x00)

        header = self.d.add_paragraph()
        hr = header.add_run(f'[{lang or "code"} block]')
        hr.italic = True; hr.font.size = Pt(9)
        hr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
        for line in text.splitlines():
            p = self.d.add_paragraph()
            pPr = p._p.get_or_add_pPr()
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), 'F4F4F4')
            pPr.append(shd)
            r = p.add_run(line if line else ' ')
            r.font.name = 'Consolas'; r.font.size = Pt(9)

    # --- images --------------------------------------------------------------

    def _add_figure_note(self, text):
        p = self.d.add_paragraph()
        p.paragraph_format.space_after = Pt(8)
        r = p.add_run(text)
        r.italic = True
        return p

    def _add_caption(self, alt):
        # The template ships a 'Caption' style (9pt, ITQ accent colour, spacing
        # before/after); reuse it so captions match anything the author adds
        # by hand in Word. Fall back to an explicit small italic Royal Blue.
        has_style = 'Caption' in [s.name for s in self.d.styles]
        p = self.d.add_paragraph(style='Caption' if has_style else None)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        render_runs(p, alt)
        if not has_style:
            p.paragraph_format.space_after = Pt(8)
            for r in p.runs:
                r.italic = True
                r.font.size = Pt(9)
                r.font.color.rgb = ITQ_ROYAL_BLUE
        return p

    def add_image(self, alt, target, line=None):
        """Centred picture, at most IMAGE_MAX_WIDTH_IN wide (natural width at
        96 dpi if smaller), aspect ratio kept, alt text as caption underneath.
        Never raises: every failure warns and leaves an italic note in place."""
        where = '%s:%s' % (self.source_name, line if line is not None else '?')
        # Remote images are never fetched. That includes protocol-relative
        # (`//host/x.png`) and UNC (`\\host\share\x.png`) targets, which
        # os.path.isfile() would otherwise probe over SMB on Windows.
        if (re.match(r'^[a-z][a-z0-9+.-]*://', target, re.IGNORECASE)
                or target.startswith(('//', '\\\\'))):
            warn('%s: remote images are not downloaded, skipped: %s' % (where, target))
            self._add_figure_note('[Figure missing: %s]' % target)
            return
        path = resolve_image_path(target, self.source_dir)
        if not os.path.isfile(path):
            warn('%s: image file not found, skipped: %s (looked for %s)'
                 % (where, target, path))
            self._add_figure_note('[Figure missing: %s]' % target)
            return
        ext = os.path.splitext(path)[1].lower()
        if ext not in IMAGE_FORMATS:
            warn('%s: unsupported image format %s, skipped: %s (supported: %s)'
                 % (where, ext or '(none)', target, ', '.join(IMAGE_FORMATS)))
            self._add_figure_note('[Figure missing: %s]' % target)
            return
        p = None
        try:
            px_width = DocxImage.from_file(path).px_width
            natural = Inches(px_width / IMAGE_SCREEN_DPI)
            width = min(Inches(IMAGE_MAX_WIDTH_IN), natural)
            p = self.d.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.keep_with_next = True
            p.add_run().add_picture(path, width=width)
        except Exception as e:  # corrupt file, unreadable header, ...
            # Remove only the paragraph THIS call created (add_picture can
            # fail after it exists), never whatever preceded the figure.
            if p is not None:
                p._p.getparent().remove(p._p)
            warn('%s: could not embed image %s: %s' % (where, target, e))
            self._add_figure_note('[Figure missing: %s]' % target)
            return
        if alt:
            self._add_caption(alt)
        else:
            p.paragraph_format.keep_with_next = False
            p.paragraph_format.space_after = Pt(8)

    def add_callout(self, icon, color, text):
        t = self.d.add_table(rows=1, cols=1)
        t.autofit = True
        cell = t.rows[0].cells[0]
        shade_cell(cell, CALLOUT_COLORS.get(color, 'F0F0F0'))
        set_cell_border(cell)
        p = cell.paragraphs[0]
        first = True
        for para_text in re.split(r'\n\s*\n', text):
            if not first:
                p = cell.add_paragraph()
            first = False
            render_runs(p, para_text.replace('\n', ' ').strip())
        self.d.add_paragraph()

    def add_table(self, rows, header=True):
        if not rows:
            return
        ncols = max(len(r) for r in rows)
        t = self.d.add_table(rows=len(rows), cols=ncols)
        t.style = self.d.styles['ITQ Tabel']
        for i, row in enumerate(rows):
            for j in range(ncols):
                cell = t.rows[i].cells[j]
                p = cell.paragraphs[0]
                for r in list(p.runs): r.text = ''
                val = row[j] if j < len(row) else ''
                render_runs(p, val, base_bold=(header and i == 0))
        self.d.add_paragraph()

    # --- template hygiene ----------------------------------------------------

    def clear_template_demo(self):
        """Remove the 'Introduction / Subheading 1 / Subheading 2' demo
        section and the 4-column demo table that the template ships with."""
        if len(self.d.tables) >= 2:
            demo = self.d.tables[1]._tbl
            demo.getparent().remove(demo)
        for idx in range(32, 22, -1):
            if idx < len(self.d.paragraphs):
                p = self.d.paragraphs[idx]
                p._p.getparent().remove(p._p)

    def fill_version_history(self, rows):
        vh = self.d.tables[0]
        for i, row in enumerate(rows, start=1):
            if i >= len(vh.rows): break
            for j, val in enumerate(row):
                p = vh.rows[i].cells[j].paragraphs[0]
                for r in list(p.runs): r.text = ''
                p.add_run(val)
        while len(vh.rows) > 1 + len(rows):
            last = vh.rows[-1]._tr
            last.getparent().remove(last)

    def save(self, path):
        self.d.save(path)


# ---------- Post-save zip patcher --------------------------------------------
#
# We deliberately do NOT parse document.xml through ElementTree. document.xml
# declares ~20 namespaces (w, r, m, v, wp, w14, w15, w16se, wpg, wpi, wpc,
# wpp, sl, ...) and ET re-serialises any it doesn't recognise as ns0:, ns1:,
# etc., which corrupts SDT data-binding xpaths and triggers Word's "found
# unreadable content" recovery prompt. Targeted byte-level edits keep every
# other byte of document.xml exactly as python-docx emitted it.

def _xml_escape(value):
    return (value
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;'))


def _patch_custom_xml(raw_bytes, fields):
    """Replace <Tag>..</Tag> or <Tag/> in customXml/item2.xml."""
    xml = raw_bytes.decode('utf-8')
    for tag, value in fields.items():
        escaped = _xml_escape(value)
        pat_open_close = re.compile(
            r'<' + re.escape(tag) + r'(\s[^>]*)?>.*?</' + re.escape(tag) + r'>',
            re.DOTALL)
        pat_self_close = re.compile(
            r'<' + re.escape(tag) + r'(\s[^>]*)?/>')
        new_open_close = f'<{tag}>{escaped}</{tag}>'
        if pat_open_close.search(xml):
            xml = pat_open_close.sub(new_open_close, xml, count=1)
        elif pat_self_close.search(xml):
            xml = pat_self_close.sub(new_open_close, xml, count=1)
    return xml.encode('utf-8')


def _patch_core_xml(raw_bytes, title, creator, subject=None):
    """Replace <dc:title>..</dc:title> etc. in docProps/core.xml."""
    xml = raw_bytes.decode('utf-8')

    def replace(xml, tag, value):
        escaped = _xml_escape(value)
        # Open/close form
        pat = re.compile(r'<' + tag + r'(\s[^>]*)?>.*?</' + tag + r'>', re.DOTALL)
        if pat.search(xml):
            return pat.sub(f'<{tag}>{escaped}</{tag}>', xml, count=1)
        # Self-closing form
        pat2 = re.compile(r'<' + tag + r'(\s[^>]*)?/>')
        if pat2.search(xml):
            return pat2.sub(f'<{tag}>{escaped}</{tag}>', xml, count=1)
        return xml

    xml = replace(xml, 'dc:title', title)
    xml = replace(xml, 'dc:creator', creator)
    if subject is not None:
        xml = replace(xml, 'dc:subject', subject)
    return xml.encode('utf-8')


_CUSTOM_STAMP_RE = re.compile(
    r'<property\b[^>]*\bname="ITQRendererVersion".*?</property>', re.DOTALL)
_CUSTOM_PID_RE = re.compile(r'\bpid="(\d+)"')


def _patch_docprops_custom(raw_bytes):
    """Record which renderer build produced this file as a custom document
    property.

    Word preserves custom properties across save and PDF export, unlike
    cp:lastModifiedBy, which it overwrites with the current user -- so this is
    the only stamp that survives the finalize pass and reaches shipped output.
    Lets a stale bundled copy be identified from the artifact alone.
    """
    xml = raw_bytes.decode('utf-8')
    xml = _CUSTOM_STAMP_RE.sub('', xml)
    pids = [int(p) for p in _CUSTOM_PID_RE.findall(xml)]
    pid = max(pids) + 1 if pids else 2
    prop = (
        '<property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" '
        'pid="%d" name="ITQRendererVersion"><vt:lpwstr>%s</vt:lpwstr></property>'
        % (pid, _xml_escape('itq-word-render ' + __version__))
    )
    if '</Properties>' in xml:
        xml = xml.replace('</Properties>', prop + '</Properties>', 1)
    return xml.encode('utf-8')


_SDT_RE = re.compile(r'<w:sdt>(.*?)</w:sdt>', re.DOTALL)
_ALIAS_RE = re.compile(r'<w:alias[^>]*w:val="([^"]*)"')
_TAG_RE = re.compile(r'<w:tag[^>]*w:val="([^"]*)"')
_XPATH_RE = re.compile(r'<w:dataBinding[^>]*w:xpath="([^"]*)"')
_SHOWINGPLC_RE = re.compile(r'<w:showingPlcHdr\s*/>')
_SDT_CONTENT_OPEN_RE = re.compile(r'<w:sdtContent>')
_WT_RE = re.compile(r'<w:t(\s[^>]*)?>([^<]*)</w:t>', re.DOTALL)


def _identify_sdt(inner, keys):
    """Return the matcher key for this SDT, or None."""
    # A <w:tag> is the most intentional handle (set deliberately in the
    # template), so it wins over alias/data-binding.
    tag = _TAG_RE.search(inner)
    if tag and tag.group(1) in keys:
        return tag.group(1)
    alias = _ALIAS_RE.search(inner)
    if alias and alias.group(1) in keys:
        return alias.group(1)
    xp = _XPATH_RE.search(inner)
    if xp:
        path = xp.group(1)
        for key in keys:
            if path.endswith(f':{key}[1]') or path.endswith(f':{key}'):
                return key
    return None


_PLACEHOLDER_RSTYLE_RE = re.compile(r'<w:rStyle[^>]*w:val="PlaceholderText"[^>]*/>')


def _rewrite_sdt(inner, value):
    """Given the contents between <w:sdt> and </w:sdt>, rewrite the first
    <w:t> inside sdtContent to the new value and drop showingPlcHdr."""
    value = strip_emoji(value, trim=True)
    escaped = _xml_escape(value)
    # Drop the showingPlcHdr flag wherever it appears inside sdtPr
    inner = _SHOWINGPLC_RE.sub('', inner)
    # Locate sdtContent block
    m = _SDT_CONTENT_OPEN_RE.search(inner)
    if not m:
        return inner
    content_start = m.end()
    content = inner[content_start:]
    # Strip Word's grey placeholder run style, otherwise the injected value
    # ships looking like un-filled template text.
    content = _PLACEHOLDER_RSTYLE_RE.sub('', content)
    # Replace the FIRST w:t's text, blank all subsequent ones
    seen = {'hit': False}

    def repl(match):
        attrs = match.group(1) or ''
        if not seen['hit']:
            seen['hit'] = True
            # Ensure xml:space="preserve" so leading/trailing spaces survive
            if 'xml:space' not in attrs:
                attrs = ' xml:space="preserve"' + attrs
            return f'<w:t{attrs}>{escaped}</w:t>'
        return f'<w:t{attrs}></w:t>'

    content = _WT_RE.sub(repl, content)
    content = _inherit_sdt_run_props(inner[:content_start], content)
    return inner[:content_start] + content


_SDTPR_RPR_RE = re.compile(r'<w:sdtPr>\s*(<w:rPr>.*?</w:rPr>)', re.DOTALL)
_RUN_OPEN_RE = re.compile(r'<w:r(?:\s[^>]*)?>')


def _inherit_sdt_run_props(sdt_head, content):
    """Give the run that carries the injected value the control's own run
    properties when it has none of its own.

    Word applies <w:sdtPr><w:rPr> to placeholder text only. Once a real run
    sits in sdtContent without an <w:rPr>, it inherits the paragraph style
    instead, which on the dark cover page turns the Client value Royal Blue
    on Royal Blue (invisible) while Author/Date, whose template runs carry an
    explicit white <w:rPr>, render fine. Copying the sdtPr rPr onto a bare run
    restores what the template author intended."""
    props = _SDTPR_RPR_RE.search(sdt_head)
    if not props:
        return content
    t_pos = content.find('<w:t')
    if t_pos < 0:
        return content
    run_open = None
    for m in _RUN_OPEN_RE.finditer(content, 0, t_pos):
        run_open = m
    if run_open is None:
        return content
    between = content[run_open.end():t_pos]
    # An rPr left empty after the PlaceholderText style was stripped counts
    # as absent: Word drops it on save and the run falls back to the paragraph.
    empty = _EMPTY_RPR_RE.search(between)
    if empty:
        start = run_open.end() + empty.start()
        end = run_open.end() + empty.end()
        return content[:start] + props.group(1) + content[end:]
    if '<w:rPr>' in between or '<w:rPr ' in between:
        return content
    return content[:run_open.end()] + props.group(1) + content[run_open.end():]


_EMPTY_RPR_RE = re.compile(r'<w:rPr>\s*</w:rPr>|<w:rPr\s*/>')


def _patch_document_xml(raw_bytes, field_values):
    """Rewrite SDT content and strip emojis from all <w:t> text nodes via
    targeted regex substitution. Leaves every unrelated byte untouched."""
    xml = raw_bytes.decode('utf-8')
    keys = set(field_values.keys())

    # 1) Rewrite identified SDTs (non-greedy, handles nested SDTs correctly
    # because _SDT_RE is non-greedy and applied repeatedly with a callback).
    def rewrite(match):
        inner = match.group(1)
        key = _identify_sdt(inner, keys)
        if key is None:
            return match.group(0)
        return '<w:sdt>' + _rewrite_sdt(inner, field_values[key]) + '</w:sdt>'

    xml = _SDT_RE.sub(rewrite, xml)

    # 2) Emoji sweep across all <w:t> text nodes (no whitespace trimming).
    def scrub(match):
        attrs = match.group(1) or ''
        inner_text = match.group(2)
        if not inner_text:
            return match.group(0)
        cleaned = _EMOJI_RE.sub('', inner_text)
        if cleaned == inner_text:
            return match.group(0)
        # Preserve xml:space if the cleaned text now has leading/trailing ws
        if (cleaned != cleaned.strip()) and 'xml:space' not in attrs:
            attrs = ' xml:space="preserve"' + attrs
        return f'<w:t{attrs}>{cleaned}</w:t>'

    xml = _WT_RE.sub(scrub, xml)
    return xml.encode('utf-8')


_UPDATEFIELDS_RE = re.compile(r'<w:updateFields\b[^>]*/>')
_SETTINGS_OPEN_RE = re.compile(r'(<w:settings\b[^>]*>)')


def _patch_settings_xml(raw_bytes):
    """Ensure <w:updateFields w:val="true"/> is present so Word recomputes the
    TOC (and other fields) the first time the document is opened — otherwise a
    template TOC field renders stale ('Error! Bookmark not defined.')."""
    xml = raw_bytes.decode('utf-8')
    if _UPDATEFIELDS_RE.search(xml):
        return raw_bytes  # already set
    # w:updateFields is valid as the first child of w:settings; Word accepts it
    # immediately after the opening tag regardless of strict schema ordering.
    new_xml, n = _SETTINGS_OPEN_RE.subn(
        r'\1<w:updateFields w:val="true"/>', xml, count=1)
    return new_xml.encode('utf-8') if n else raw_bytes


_PARA_PPR_RE = re.compile(r'<w:pPr>.*?</w:pPr>', re.DOTALL)
_RPR_BLOCK_RE = re.compile(r'<w:rPr>.*?</w:rPr>', re.DOTALL)

# Marks the data-binding xpath of the cover/title-page "Version" row, which
# every metadata block in the template carries (the dark cover text box and
# the light title page both bind ProjectName/ClientName/Version the same
# way). Anchoring on this rather than a fixed paragraph index lets a new
# Classification row be added next to Version wherever that block appears,
# without needing a dedicated content control for it in the template.
_VERSION_XPATH_MARK = 'ns0:Version[1]'


def _insert_cover_classification_rows(raw_bytes, classification):
    """Add a 'Classification' row directly after every cover/title-page
    'Version' row, reusing that row's own paragraph and run formatting so
    the new row automatically matches white-on-dark or Royal-Blue/black-on-
    light, whichever the surrounding block already uses. No-op (byte-
    identical passthrough) when classification is empty."""
    if not classification:
        return raw_bytes
    xml = raw_bytes.decode('utf-8')
    escaped = _xml_escape(strip_emoji(classification, trim=True))
    inserts = []  # (insert_pos, new_paragraph_xml)
    search_from = 0
    while True:
        idx = xml.find(_VERSION_XPATH_MARK, search_from)
        if idx < 0:
            break
        start = max(xml.rfind('<w:p>', 0, idx), xml.rfind('<w:p ', 0, idx))
        end = xml.find('</w:p>', idx)
        if start < 0 or end < 0:
            search_from = idx + len(_VERSION_XPATH_MARK)
            continue
        end_full = end + len('</w:p>')
        para = xml[start:end_full]
        ppr_m = _PARA_PPR_RE.search(para)
        rprs = _RPR_BLOCK_RE.findall(para[ppr_m.end():]) if ppr_m else []
        if not ppr_m or len(rprs) < 2:
            # Unexpected structure (template changed) — skip rather than
            # emit a malformed paragraph.
            search_from = end_full
            continue
        bold_rpr, plain_rpr = rprs[0], rprs[1]  # label style, value style
        new_para = (
            '<w:p>' + ppr_m.group(0)
            + '<w:r>' + bold_rpr + '<w:t xml:space="preserve">Classification</w:t></w:r>'
            + '<w:r>' + plain_rpr + '<w:tab/></w:r>'
            + '<w:r>' + plain_rpr + '<w:t xml:space="preserve">' + escaped + '</w:t></w:r>'
            + '</w:p>'
        )
        inserts.append((end_full, new_para))
        search_from = end_full
    for pos, new_para in reversed(inserts):
        xml = xml[:pos] + new_para + xml[pos:]
    return xml.encode('utf-8')


def _append_footer_classification(raw_bytes, classification):
    """Append ' | CLASSIFICATION' to the end of the footer's page-number
    line, reusing the run formatting already used there (Titillium Web,
    Royal Blue, 8pt) so it reads as part of the same footer text. No-op
    (byte-identical passthrough) when classification is empty."""
    if not classification:
        return raw_bytes
    xml = raw_bytes.decode('utf-8')
    ftr_end = xml.rfind('</w:ftr>')
    if ftr_end < 0:
        return raw_bytes
    p_end = xml.rfind('</w:p>', 0, ftr_end)
    if p_end < 0:
        return raw_bytes
    rpr_end = xml.rfind('</w:rPr>', 0, p_end)
    rpr_start = xml.rfind('<w:rPr>', 0, rpr_end) if rpr_end >= 0 else -1
    rpr = xml[rpr_start:rpr_end + len('</w:rPr>')] if rpr_start >= 0 else ''
    label = strip_emoji(classification, trim=True).upper()
    run = '<w:r>' + rpr + '<w:t xml:space="preserve"> | ' + _xml_escape(label) + '</w:t></w:r>'
    return (xml[:p_end] + run + xml[p_end:]).encode('utf-8')


def postprocess_docx(path, meta):
    """Patch the saved docx in place to populate SDTs and strip emojis."""
    sdt_values = {
        'ClientName':  meta['client'],
        'ProjectName': meta['project'],
        'Ondertitel':  meta.get('subtitle', ''),
        'Version':     meta['version'],
        # core-property-bound SDTs
        'title':       meta['title_cover'],
        'Titel':       meta['title_cover'],   # alias
        'creator':     meta['author'],
        'Auteur':      meta['author'],        # alias
    }
    # Management-summary control: filled only when --summary is supplied.
    # Targeted by the template tag 'ManagementSummary'.
    if meta.get('summary'):
        sdt_values['ManagementSummary'] = meta['summary']

    classification = meta.get('classification') or ''

    tmp = path + '.tmp'
    with zipfile.ZipFile(path, 'r') as zin, zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            name = item.filename
            if name == 'customXml/item2.xml':
                data = _patch_custom_xml(data, {
                    'ClientName':  strip_emoji(meta['client'], trim=True),
                    'ProjectName': strip_emoji(meta['project'], trim=True),
                    'Ondertitel':  strip_emoji(meta.get('subtitle', ''), trim=True),
                    'Version':     strip_emoji(meta['version'], trim=True),
                })
            elif name == 'docProps/core.xml':
                data = _patch_core_xml(
                    data,
                    title=strip_emoji(meta['title_cover'], trim=True),
                    creator=strip_emoji(meta['author'], trim=True),
                )
            elif name == 'word/document.xml':
                data = _patch_document_xml(data, sdt_values)
                data = _insert_cover_classification_rows(data, classification)
            elif name in ('word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'):
                data = _patch_document_xml(data, sdt_values)
                data = _append_footer_classification(data, classification)
            elif name in ('word/header1.xml', 'word/header2.xml', 'word/header3.xml'):
                data = _patch_document_xml(data, sdt_values)
            elif name == 'word/settings.xml':
                data = _patch_settings_xml(data)
            elif name == 'docProps/custom.xml':
                data = _patch_docprops_custom(data)
            zout.writestr(item, data)
    shutil.move(tmp, path)


# ---------- Orchestration -----------------------------------------------------

def render(template, source_md, out_path, meta):
    asset_dir = os.path.join(os.path.dirname(out_path), 'assets')
    # meta['source_path'] (the .md file) anchors relative image paths and
    # names the file in warnings. Callers that pass markdown text only get
    # cwd-relative resolution.
    source_path = meta.get('source_path')
    source_dir = os.path.dirname(os.path.abspath(source_path)) if source_path else None
    source_name = source_path if source_path else '<markdown>'
    r = ItqRenderer(template, asset_dir, source_dir=source_dir,
                    source_name=source_name)
    r.fill_version_history(meta['version_history'])
    r.clear_template_demo()

    # Markdown heading level -> Word heading level. Default offset 0 means
    # standard markdown (# -> Heading 1, ## -> Heading 2, ### -> Heading 3).
    # Pass offset -1 for Notion-export markdown, where the page title is a
    # single leading `#` and real sections start at `##` (reproduces the old
    # behaviour: ## -> Heading 1).
    offset = meta.get('heading_offset', 0)

    # Document title as first Heading 1
    r.add_heading(1, meta['title_cover'])

    blocks = parse_blocks(source_md, source_name=source_name)
    for b in blocks:
        if b.kind == 'heading':
            r.add_heading(max(1, b.level + offset), b.text)
        elif b.kind == 'paragraph':
            r.add_paragraph(b.text)
        elif b.kind == 'image':
            r.add_image(b.alt, b.target, b.line)
        elif b.kind == 'ulist':
            for item in b.items: r.add_bullet(item)
        elif b.kind == 'olist':
            for item in b.items: r.add_numbered(item)
        elif b.kind == 'callout':
            r.add_callout(b.icon, b.color, b.text)
        elif b.kind == 'table':
            r.add_table(b.rows, header=b.header)
        elif b.kind == 'code':
            r.add_code(b.text, b.lang)
        elif b.kind == 'blockquote':
            r.add_blockquote(b.text)
        elif b.kind == 'hr':
            r.add_hr()

    r.save(out_path)
    # SDT / core.xml / custom XML population + final emoji sweep
    postprocess_docx(out_path, meta)


def _default_template():
    """Locate the ITQ template shipped alongside this script."""
    here = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(here, 'templates', 'ITQ - Word document Template.docx')


def _build_cli():
    import argparse
    p = argparse.ArgumentParser(
        description='Convert a markdown file into an ITQ-branded Word document.',
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )
    p.add_argument('--source', required=True,
                   help='Path to the markdown file (.md)')
    p.add_argument('--out', required=True,
                   help='Path for the generated .docx')
    p.add_argument('--template', default=_default_template(),
                   help='Path to the ITQ Word template (.docx)')
    p.add_argument('--title', default='Document title',
                   help='Document title shown on the cover and in headers/footers')
    p.add_argument('--subtitle', default='',
                   help='Document subtitle shown on the cover')
    p.add_argument('--client', default='ITQ — Internal',
                   help='Client / recipient — appears in cover metadata block')
    p.add_argument('--project', default='',
                   help='Project name — appears in cover metadata block')
    p.add_argument('--author', default='',
                   help='Author — appears in cover metadata block and core properties')
    p.add_argument('--version', default='1.0',
                   help='Version label — appears in cover metadata block')
    p.add_argument('--classification', default='',
                   help='Document classification (e.g. INTERNAL, CONFIDENTIAL). '
                        'Adds a Classification row next to Client/Project/Version '
                        'on the cover and title page, and a short uppercase label '
                        'in the footer of every body page. Omit for no change to '
                        'the output (default behaviour).')
    p.add_argument('--heading-offset', type=int, default=0,
                   help='Shift markdown heading levels onto Word heading levels. '
                        '0 (default) = standard markdown (# -> Heading 1). '
                        '-1 = Notion-export markdown (## -> Heading 1).')
    p.add_argument('--summary', default=None,
                   help='Management-summary text, OR a path to a .md/.txt file '
                        'whose contents become the summary. Fills the template '
                        "control tagged 'ManagementSummary'. Omit to leave blank.")
    p.add_argument('--version-history', action='append', default=None,
                   metavar='VER|DATE|AUTHOR|CHANGES',
                   help='Version history row, pipe-separated. Repeat for '
                        'multiple rows. Omit to use a single default row.')
    return p


def _parse_history(rows):
    if not rows:
        from datetime import date
        return [('0.9', date.today().strftime('%-d %B %Y') if sys.platform != 'win32'
                 else date.today().strftime('%#d %B %Y'),
                 'Draft', 'Initial draft.')]
    parsed = []
    for row in rows:
        parts = [p.strip() for p in row.split('|')]
        while len(parts) < 4:
            parts.append('')
        parsed.append(tuple(parts[:4]))
    return parsed


if __name__ == '__main__':
    args = _build_cli().parse_args()

    with open(args.source, 'r', encoding='utf-8') as f:
        md = f.read()

    # --summary accepts inline text or a path to a .md/.txt file.
    summary = args.summary
    if summary and os.path.isfile(summary):
        with open(summary, 'r', encoding='utf-8') as f:
            summary = f.read().strip()

    meta = {
        'title_cover':     args.title,
        'subtitle':        args.subtitle,
        'client':          args.client,
        'project':         args.project,
        'author':          args.author,
        'version':         args.version,
        'classification':  args.classification,
        'summary':         summary,
        'version_history': _parse_history(args.version_history),
        'heading_offset':  args.heading_offset,
        'source_path':     args.source,
    }

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    render(args.template, md, args.out, meta)
    print(f'wrote {args.out}  [itq-word-render {__version__}]')
