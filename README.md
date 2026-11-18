# ToWord — PDF → editable Word, with Khmer support

Upload a PDF (or image / DOC / ODT / RTF / DOCX), convert it to Word, edit it in the
browser, then download it as `.docx` or PDF. Khmer (ភាសាខ្មែរ) is handled throughout.

## Run

```bash
./run.sh            # then open http://localhost:8000
```

Requires (already present on this Mac): Python 3.12, LibreOffice (`soffice`), Tesseract.

## How it works

| Step | Tool |
|---|---|
| Digital PDF → DOCX, keeping layout/tables/images | `pdf2docx` + Khmer text repair |
| Khmer PDF with a broken text layer (typical Word + Khmer OS export) → DOCX | **Hybrid**: `pdf2docx` layout, each Khmer phrase re-read by OCR from its exact position (`app/hybrid.py`) |
| Scanned PDF / images / broken Khmer text layer → DOCX | Tesseract OCR (`khm+eng`, model in `tessdata/`) |
| DOC / ODT / RTF → DOCX | LibreOffice |
| DOCX → editor HTML | `app/docx2html.py` (keeps Khmer fonts, sizes, alignment, tables, images) |
| Editor | TinyMCE 7 (GPL, served locally from `static/vendor/`) |
| Edited HTML → DOCX / PDF | LibreOffice, then Khmer font fixing (`app/khmer.py`) |

**Automatic mode** checks the PDF's text layer first:
- scanned or unreadable legacy font codes → OCR;
- contains Khmer → OCR a sample of Khmer phrases and compare with the text layer.
  Agreement below 0.9 means the text layer is wrong (Word often stores wrong
  codes for Khmer OS glyphs, e.g. `នាមត្រកូេ និងនាមខលួន` for `នាមត្រកូល និងនាមខ្លួន`)
  → hybrid mode; otherwise normal layout mode.

### Side boxes (`app/boxes.py`)
Bordered boxes that sit beside other text (for example the "បិទរូបថតថ្មី ៤ x ៦" photo box
on forms) confuse pdf2docx's reading order. They are removed from the PDF before
pdf2docx runs and re-inserted as floating tables at their exact page position.
In the editor they are floated boxes. Edited downloads keep them next to their paragraph.
Boxes that are part of a table grid or contain images are left to pdf2docx.

### Khmer specifics
- PDFs exported from Chrome, Word and LibreOffice often extract Khmer with duplicated
  vowels or signs (`ព្រះះ`, `សាាសនា`, `ទៅទៅ`). `khmer.repair_khmer` fixes these.
- Every Khmer run in the output `.docx` gets the chosen Khmer font in the
  complex-script slot (`w:cs`) and the Latin slots, plus `km-KH` language, so Word,
  LibreOffice and Google Docs all display it correctly. Fonts the document already
  uses (for example Khmer Muol titles) are kept.
- `fonts/` holds open-licensed (OFL) Google Khmer fonts: Battambang, Siemreap, Moul,
  Hanuman, Koulen, Kantumruy Pro and Noto Khmer. PDF export uses them in place of the
  Khmer OS fonts they are modelled on, so PDFs render even without Khmer OS installed.

## Layout
```
app/main.py        FastAPI server + job API
app/converters.py  conversion pipeline (layout, OCR, LibreOffice)
app/docx2html.py   DOCX → HTML for the editor
app/khmer.py       Khmer detection, text repair, DOCX font fixing
static/            front-end (index.html, app.js, styles.css, TinyMCE)
data/jobs/         uploaded files and drafts (deleted after 24 h)
```

## Limits
- Hybrid/OCR text can still contain small OCR mistakes (a missing vowel, a wrong
  consonant), mostly in short words and decorative Muol fonts. Proofread names and numbers.
- OCR output is plain flowing text with estimated sizes and alignment. Tables in
  scanned documents come out as text.
- Editing in the browser is HTML-based. Complex layouts (text boxes, floating
  shapes) may shift. "Download → Word (.docx) — original conversion" gives the
  exact conversion result without browser edits.
