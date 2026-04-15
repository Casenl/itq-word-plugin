"""
Convert a markdown file into an ITQ-branded Word document.

Works with:
  - Notion's enhanced-markdown export (as produced by the Notion MCP server)
  - Standard markdown (GitHub READMEs, hand-written notes, etc.)

Handles:
  # / ## / ### headings  (## -> Heading 1)
  **bold**, *italic*, `inline code`
  - bullets          (ITQ Bullets via numId 15)
  1. numbered items  (ITQ Numbering via numId 18)
  <callout icon="..." color="..."> ... </callout>         Notion-enhanced only
  <table [header-row="true"]> <tr><td>..</td>..</tr> ... </table>  Notion-enhanced
  ``` fenced code blocks ```          (mermaid rendered via kroki.io)
  > blockquote
  --- horizontal rule

Not yet supported: GitHub-style pipe tables, inline [links](url), ![images].

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
import urllib.request
import xml.etree.ElementTree as ET
import zipfile
from copy import deepcopy

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Cm, Emu, Inches


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


def parse_blocks(text):
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

        m = re.match(r'^(#{1,6})\s+(.*)$', stripped)
        if m:
            blocks.append(Block('heading', level=len(m.group(1)),
                                text=m.group(2).strip()))
            i += 1
            continue

        if stripped.startswith('>'):
            buf = []
            while i < n and lines[i].strip().startswith('>'):
                buf.append(lines[i].strip().lstrip('>').lstrip())
                i += 1
            blocks.append(Block('blockquote', text='\n'.join(buf).strip()))
            continue

        m = re.match(r'^(\d+)\.\s+(.*)$', stripped)
        if m:
            items = []
            while i < n:
                mm = re.match(r'^(\d+)\.\s+(.*)$', lines[i].strip())
                if not mm: break
                items.append(mm.group(2))
                i += 1
            blocks.append(Block('olist', items=items))
            continue

        if stripped.startswith('- ') or stripped.startswith('* '):
            items = []
            while i < n:
                s = lines[i].strip()
                if s.startswith('- ') or s.startswith('* '):
                    items.append(s[2:]); i += 1
                else:
                    break
            blocks.append(Block('ulist', items=items))
            continue

        # Paragraph: in Notion-flavoured markdown each non-empty line is its
        # own logical paragraph (no soft-wrapping), so emit one block per line.
        blocks.append(Block('paragraph', text=stripped))
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
            'User-Agent': 'ITQ-doc-builder/1.0',
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

class ItqRenderer:
    def __init__(self, template_path, asset_dir):
        self.d = Document(template_path)
        self.asset_dir = asset_dir
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

    def add_heading(self, level, text):
        mapped = max(1, level - 1)
        if mapped > 9: mapped = 9
        p = self.d.add_paragraph(style=f'Heading {mapped}')
        render_runs(p, text)

    def add_paragraph(self, text):
        p = self.d.add_paragraph()
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
                caption = self.d.add_paragraph()
                cr = caption.add_run('Decision flowchart (rendered via kroki.io)')
                cr.italic = True
                cr.font.size = Pt(9)
                cr.font.color.rgb = RGBColor(0x80, 0x80, 0x80)
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


_SDT_RE = re.compile(r'<w:sdt>(.*?)</w:sdt>', re.DOTALL)
_ALIAS_RE = re.compile(r'<w:alias[^>]*w:val="([^"]*)"')
_XPATH_RE = re.compile(r'<w:dataBinding[^>]*w:xpath="([^"]*)"')
_SHOWINGPLC_RE = re.compile(r'<w:showingPlcHdr\s*/>')
_SDT_CONTENT_OPEN_RE = re.compile(r'<w:sdtContent>')
_WT_RE = re.compile(r'<w:t(\s[^>]*)?>([^<]*)</w:t>', re.DOTALL)


def _identify_sdt(inner, keys):
    """Return the matcher key for this SDT, or None."""
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
    return inner[:content_start] + content


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
            elif name in ('word/document.xml', 'word/footer1.xml',
                          'word/footer2.xml', 'word/footer3.xml',
                          'word/header1.xml', 'word/header2.xml',
                          'word/header3.xml'):
                data = _patch_document_xml(data, sdt_values)
            zout.writestr(item, data)
    shutil.move(tmp, path)


# ---------- Orchestration -----------------------------------------------------

def render(template, source_md, out_path, meta):
    asset_dir = os.path.join(os.path.dirname(out_path), 'assets')
    r = ItqRenderer(template, asset_dir)
    r.fill_version_history(meta['version_history'])
    r.clear_template_demo()

    # Document title as first Heading 1
    r.add_heading(2, meta['title_cover'])

    blocks = parse_blocks(source_md)
    for b in blocks:
        if b.kind == 'heading':
            r.add_heading(b.level, b.text)
        elif b.kind == 'paragraph':
            r.add_paragraph(b.text)
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

    meta = {
        'title_cover':     args.title,
        'subtitle':        args.subtitle,
        'client':          args.client,
        'project':         args.project,
        'author':          args.author,
        'version':         args.version,
        'version_history': _parse_history(args.version_history),
    }

    os.makedirs(os.path.dirname(os.path.abspath(args.out)) or '.', exist_ok=True)
    render(args.template, md, args.out, meta)
    print(f'wrote {args.out}')
