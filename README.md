# itq-word-plugin

Convert any markdown file — a Notion page export, a GitHub README, hand-written notes — into an **ITQ-branded Word document** that uses the official ITQ template. Cover page, fonts, colours, headings, tables, bullets, version history, and rendered Mermaid diagrams all come out styled correctly.

The output is a `.docx` that opens cleanly in Word with all the ITQ styles applied — no manual reformatting.

---

## What you get

| Input | Output |
|---|---|
| A Notion page (internal guide, proposition, runbook, whitepaper, …) | A Word `.docx` in ITQ styling: cover page with Client / Project / Author / Version filled in, ITQ headings and lists, branded tables, colour-coded callout boxes, rendered Mermaid diagrams, and auto-generated version history. |

---

## Requirements

- **Python 3.9+** (tested on 3.11 and 3.13)
- **python-docx** — installed via `requirements.txt`
- **Internet access** — only needed when your markdown contains a Mermaid diagram (we render it through the free public [kroki.io](https://kroki.io) service). Everything else is offline.
- **The ITQ Word template** — shipped in `templates/`, no extra download needed.

---

## Install

### Windows

```powershell
# 1. Clone (or copy) this folder somewhere convenient
cd C:\path\to\your\workspace
git clone <repo-url> itq-word-plugin
cd itq-word-plugin

# 2. Create a virtual environment (optional but recommended)
py -m venv .venv
.\.venv\Scripts\Activate.ps1

# 3. Install dependencies
pip install -r requirements.txt

# 4. Smoke-test the CLI
py md_to_itq_docx.py --help
```

If PowerShell blocks the activation script, run PowerShell as admin once and
execute `Set-ExecutionPolicy -Scope CurrentUser RemoteSigned`.

### macOS

```bash
# 1. Clone
cd ~/code
git clone <repo-url> itq-word-plugin
cd itq-word-plugin

# 2. Virtual environment
python3 -m venv .venv
source .venv/bin/activate

# 3. Dependencies
pip install -r requirements.txt

# 4. Smoke-test the CLI
python md_to_itq_docx.py --help
```

On a fresh Mac you may need the Xcode command line tools (`xcode-select --install`) for pip to build wheels, but python-docx ships pre-built so this should not be required.

---

## Usage

### 1. Get the markdown out of Notion

The converter expects **Notion's "enhanced markdown" format** — the one produced by the Notion MCP server inside Claude Code. It's a superset of standard markdown that adds `<callout>` boxes and `<table><tr><td>` blocks.

The easiest way to get it:

1. Open Claude Code with the Notion MCP server configured (Anthropic's official Notion integration).
2. Paste the Notion page URL and ask Claude to **fetch** the page.
3. Ask Claude to save the fetched markdown to a `.md` file in this project, e.g. `examples/my-guide.md`.

If you don't have access to the MCP tool, Notion's **Export → Markdown & CSV** option produces *standard* GitHub-flavoured markdown. That also works for the common things (headings, bold, lists, blockquotes) but **pipe-style tables (`| col1 | col2 |`) are not yet supported** — they are passed through as plain text. Native markdown tables will be added in a future release. If your document doesn't have tables, the plain Notion export works fine.

### 2. Run the converter

```bash
python md_to_itq_docx.py \
    --source   examples/my-guide.md \
    --out      output/MyGuide.docx \
    --title    "My Guide" \
    --subtitle "A clear subtitle" \
    --client   "ITQ — Internal" \
    --project  "My Project" \
    --author   "Your Name — Department" \
    --version  "1.0"
```

`--source` and `--out` are required; every other flag has a sensible default. Run with `--help` for the full list.

### 3. Version history rows (optional)

Each `--version-history` flag adds one row to the Version history table on page 2. Use pipe separators `VERSION|DATE|AUTHOR|CHANGES`:

```bash
python md_to_itq_docx.py --source my.md --out out.docx \
    --version-history "0.9|15 April 2026|CTO Office|Initial draft" \
    --version-history "1.0|20 April 2026|CTO Office|Board-approved"
```

Omit the flag to get a single default row.

### 4. All CLI options

```
--source PATH         (required) Markdown file (.md)
--out PATH            (required) Output path for the generated .docx
--template PATH       ITQ Word template (default: templates/ITQ - Word document Template.docx)
--title TEXT          Document title (cover + header/footer)
--subtitle TEXT       Cover subtitle
--client TEXT         Client / recipient
--project TEXT        Project name
--author TEXT         Author
--version TEXT        Version label
--version-history ... Pipe-separated row; repeat for multiple rows
```

---

## What's inside the repo

```
itq-word-plugin/
├── README.md                          This file
├── requirements.txt                   python-docx
├── md_to_itq_docx.py                  The converter (single file)
├── finalize.py                        TOC bake + PDF export (needs Word)
├── .gitignore
└── templates/
    └── ITQ - Word document Template.docx
```

> **This folder is generated.** Both scripts and the template are copied
> byte-for-byte from the canonical repo, `Casenl/markdown-to-itq-word`. Fix bugs
> there and regenerate; edits made here are overwritten on the next build.

---

## Finishing the document: table of contents and PDF

The converter has no layout engine, so the table of contents ships as an
un-evaluated Word field — page numbers are blank until Word resolves it. Two ways
to finish the job:

**Just open it in Word.** The converter sets `updateFields`, so Word rebuilds the
table of contents the first time a human opens the `.docx`. For a document you are
handing to a person, that is enough — you need do nothing.

**Or bake it and export a PDF**, which is required if you are shipping a PDF rather
than a `.docx`:

```bash
python finalize.py --docx doc.docx --pdf doc.pdf
```

This drives Word directly (COM on Windows, AppleScript on macOS). Exit codes:
`0` done · `2` no Word on this machine, so the `.docx` was left to self-heal ·
`1` Word was there but the automation failed. Omit `--pdf` to bake the TOC only.

⚠️ **Fonts are not embedded in the template.** The `.docx` is safe to send anywhere
(the recipient's Word supplies Titillium, or falls back to Calibri), but a PDF is
only as branded as the machine that exported it. Export on a machine without
Titillium installed and the PDF silently comes out in Calibri, with no warning.
Install the ITQ fonts before exporting a customer-facing PDF, and check the result.

**Which build made this file?** Every rendered `.docx` records its renderer version
as the custom document property `ITQRendererVersion` (File → Info → Properties, or
read `docProps/custom.xml`). If a document renders wrongly and that version is
behind the canonical repo, this copy is stale — pull before debugging further.

---

## Supported markdown features

| Feature | Notion-enhanced | Standard markdown | Notes |
|---|---|---|---|
| `#` `##` `###` headings | ✅ | ✅ | `##` maps to Heading 1 (the Notion page title is the doc title). |
| `**bold**`, `*italic*`, `` `code` `` | ✅ | ✅ | |
| `-` / `*` bullets | ✅ | ✅ | Rendered with the **ITQ Bullets** list style. |
| `1.` numbered lists | ✅ | ✅ | Rendered with the **ITQ Numbering** list style. |
| `> blockquote` | ✅ | ✅ | |
| `---` horizontal rule | ✅ | ✅ | |
| ` ```mermaid ``` ` code fences | ✅ | ✅ | **Rendered to PNG via kroki.io** and embedded as an inline image. |
| Other `` ``` `` code fences | ✅ | ✅ | Rendered as grey monospace blocks. |
| `<callout icon="..." color="...">` | ✅ | ❌ | Rendered as coloured callout boxes (red_bg, blue_bg, yellow_bg, green_bg, orange_bg, gray_bg). |
| `<table><tr><td>…</td></tr></table>` | ✅ | ❌ | Rendered with the **ITQ Tabel** style. Header row is detected via `header-row="true"`. |
| Pipe-style tables (`\| a \| b \|`) | ❌ | ❌ | **Not yet supported.** Raw text is passed through. |
| Inline links `[text](url)` | ❌ | ❌ | Rendered as raw text. |
| Images `![](...)` | ❌ | ❌ | Not yet implemented (except Mermaid via kroki). |
| Emojis 🎯🚀 | — | — | **Stripped** from every heading, paragraph, table cell and cover field. ITQ brand guidelines don't use emoji in formal documents. |
| Notion backslash escapes (`\$`, `\~`, `\<`) | ✅ | — | Unescaped so `\$17K/yr` becomes `$17K/yr`. |

---

## How it works (for the curious)

The converter is a single Python file. It runs in four phases:

1. **Parse** — walks the source markdown line-by-line into a flat list of block objects (`heading`, `paragraph`, `ulist`, `olist`, `callout`, `table`, `code`, `blockquote`, `hr`). The parser is non-greedy and preserves per-line paragraph breaks, which matters because Notion export does not soft-wrap.

2. **Render into python-docx** — opens the ITQ Word template and appends content using the template's own styles (`Heading 1`, `ITQ Tabel`, `ITQ Bullets`, `ITQ Numbering`, `List Paragraph`). For Mermaid code fences we POST the source to `https://kroki.io/mermaid/png`, save the returned PNG under `assets/`, and embed it via `document.add_picture()`.

3. **Save** — python-docx writes the `.docx` zip.

4. **Post-save patch** — re-opens the zip and runs **targeted regex edits** on three parts:
   - `customXml/item2.xml` — sets the `<ClientName>`, `<ProjectName>`, `<Ondertitel>`, `<Version>` custom properties that the cover-page SDT content controls are bound to.
   - `docProps/core.xml` — sets `<dc:title>` and `<dc:creator>`.
   - `word/document.xml` and `word/footer*.xml` — for every SDT bound to those properties, drops the `<w:showingPlcHdr/>` flag and overwrites the cached text so Word displays the new value immediately instead of the `[Client]` / `[Title]` placeholder. Also runs a final emoji-strip pass on every `<w:t>` text node.

We deliberately do **not** parse/serialize these XML parts through `ElementTree`. ET strips namespace declarations it doesn't know about, renames prefixes (`dcterms:` → `ns2:`, `w:` → `ns0:`), changes attribute quoting, and collapses empty elements — all of which cause Word's "found unreadable content" recovery prompt on open. Byte-preserving regex edits keep every untouched byte of the original XML intact.

### Template cover page and SDTs

The ITQ Word template ships with **Structured Document Tag** content controls on the cover page and in the metadata block. They look like `[Client]`, `[Title]`, `[Subtitle]` in the template but are bound to:

- **Core properties** (`docProps/core.xml`) — title, creator
- **Custom XML** (`customXml/item2.xml`, an `<Extra>` element) — ClientName, ProjectName, Ondertitel, Version

When you set these values, *every* SDT bound to them updates automatically (cover page, metadata block, header, footer — all three places at once). That's why we patch the source data and remove the "show placeholder" flag, rather than trying to find and edit every SDT instance by hand.

### Kroki and ITQ branding

Mermaid diagrams are rendered by posting the raw Mermaid source to `https://kroki.io/mermaid/png` with a `User-Agent` header (Kroki rejects requests without one). Kroki returns a PNG that honours any `classDef` / `style` directives in the Mermaid source, so if your diagram includes the ITQ colour palette:

```mermaid
classDef defaultRoute fill:#04055A,color:#fff,stroke:#04055A,stroke-width:2px
classDef altRoute fill:#FF4F41,color:#fff,stroke:#FF4F41,stroke-width:2px
```

… the rendered image comes back in ITQ Royal Blue (`#04055A`) and ITQ Red (`#FF4F41`). No extra configuration needed.

Kroki is a free public service; if your organisation blocks it, you can self-host it with one Docker command — see [docs.kroki.io](https://docs.kroki.io/kroki/setup/install/) — and change the `KROKI_URL` constant near the top of `md_to_itq_docx.py`.

---

## Troubleshooting

| Symptom | Cause / fix |
|---|---|
| `HTTPError: 403 Forbidden` when rendering Mermaid | Kroki rejects requests without a User-Agent header. Check the script's `render_mermaid()` function still sets `User-Agent`. |
| `urllib.error.URLError` on Mermaid render | You're offline or Kroki is blocked. The Mermaid block falls back to a monospace code block in the document — your diagram isn't visible but the document still builds. |
| Cover page shows `[Client]` / `[Title]` instead of your values | Old `md_to_itq_docx.py` with the ElementTree post-processor. Make sure you're on the current version — the post-processor should be regex-based. |
| Word says "Word found unreadable content. Do you want to recover?" | Same cause as above. Upgrade to the current version. |
| Text is crammed together around bold words | Old version of `strip_emoji` — it was trimming whitespace from every `<w:t>` run including leading spaces. Upgrade. |
| `ValueError: assigned style is type LIST (4), need type PARAGRAPH (1)` | You're trying to pass `ITQ Bullets` as a paragraph style. It's a *list* style — the converter already handles this by applying the numId directly. Don't rename `ITQBullets` / `ITQNumbering` in the template. |
| `PermissionError: [Errno 13]` on save | The output `.docx` is open in Word. Close it and rerun, or use a different `--out` path. |
| Tables render as plain text | Your markdown uses GitHub-style pipe tables. Not supported yet — either convert to `<table>` / `<tr>` / `<td>` blocks, or ask for pipe-table support to be added. |
| Table of contents has no page numbers | Expected until Word resolves the field. Open the `.docx` in Word, or run `finalize.py`. |
| Management summary looks grey and italic, like unfilled template text | Stale copy: Word's `PlaceholderText` style was not stripped from the filled control. Check `ITQRendererVersion` in the document properties and pull the current version. |
| Paragraphs run together, headings numbered twice, or version history says "0.9 Draft" | Stale copy predating the 2026-06 fixes. Check `ITQRendererVersion` and pull. |
| Exported PDF is in Calibri, not Titillium | The exporting machine does not have the ITQ fonts installed. Fonts are not embedded in the template; install Titillium and re-export. |
| `finalize.py` exits 2 | Microsoft Word is not installed here. Not an error: the `.docx` is valid and rebuilds its own TOC when opened in Word. Do not ship a PDF from this machine. |

---

## Extending the converter

The script is deliberately a single file with no framework. Common extensions:

- **Add a new callout colour** → append to `CALLOUT_COLORS` in the script.
- **Render to a different template** → pass `--template path/to/other.docx`. As long as the other template has the same style names (`Heading 1-9`, `ITQ Tabel`, `ITQ Bullets`, `ITQ Numbering`, `List Paragraph`), it will work. New templates with different style names need a small `_discover_list_numids()` tweak.
- **Support pipe tables** → add a new block kind in `parse_blocks()` that matches `^\|.*\|$` lines and splits them on `|`. Feed rows to the existing `add_table()` helper.
- **Add image support** → handle `![alt](url)` in the inline regex; download to `assets/` and insert via `add_picture()`.
- **Change Kroki endpoint** → edit `KROKI_URL` at the top of the script.

---

## License

MIT. The ITQ Word template shipped under `templates/` is ITQ brand
material — if you fork this repo for use outside ITQ, swap in your own
template and matching style names.
