"""Khmer-specific helpers: script detection, extraction-quality checks, and
DOCX font fixing so Khmer text renders correctly in Word / LibreOffice / Google Docs."""
from __future__ import annotations

import re

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

KHMER_RE = re.compile(r"[ក-៿᧠-᧿]")
# Khmer consonants + independent vowels (things a dependent sign may legally follow)
_BASE = "ក-ឳ"
# Dependent vowels, signs and COENG (must follow a base or another sign)
_DEP = "ា-៓៝"
_ORPHAN_SIGN_RE = re.compile(rf"(?<![{_BASE}{_DEP}])[{_DEP}]")

KHMER_FONTS = [
    "Khmer OS Battambang",
    "Khmer OS Siemreap",
    "Khmer OS",
    "Khmer OS Muol Light",
    "Noto Sans Khmer",
    "Noto Serif Khmer",
    "Battambang",
    "Kantumruy Pro",
    "Hanuman",
    "Moul",
    "Khmer MN",
]


def has_khmer(text: str) -> bool:
    return bool(KHMER_RE.search(text or ""))


def ratio_khmer(text: str) -> float:
    letters = [c for c in text if not c.isspace()]
    if not letters:
        return 0.0
    return sum(1 for c in letters if KHMER_RE.match(c)) / len(letters)


def text_quality(text: str, pages: int) -> dict:
    """Heuristics to decide whether direct text extraction is usable or OCR is needed.

    - Scanned PDFs have (almost) no text layer.
    - Khmer PDFs produced with broken ToUnicode maps / legacy fonts (Limon, ABC)
      give garbage: private-use chars, U+FFFD, or dependent vowels with no base
      consonant (pre-base vowels like "េ" stored in visual order)."""
    stripped = re.sub(r"\s+", "", text or "")
    n = len(stripped)
    pua = sum(1 for c in stripped if 0xE000 <= ord(c) <= 0xF8FF)
    bad = stripped.count("�")
    khmer_chars = len(KHMER_RE.findall(stripped))
    defects_raw = count_khmer_defects(text)
    # What is still wrong after our automatic repair:
    repaired = repair_khmer(text or "")
    defects_left = count_khmer_defects(repaired)
    per_page = n / max(pages, 1)
    reasons = []
    if per_page < 25:
        reasons.append("little or no text layer (likely scanned)")
    if n and (pua + bad) / n > 0.05:
        reasons.append("unreadable glyph codes (legacy/embedded font)")
    if khmer_chars > 30 and defects_left / khmer_chars > 0.02:
        reasons.append("Khmer text layer is corrupted")
    return {
        "chars": n,
        "khmer_chars": khmer_chars,
        "khmer_defects": defects_raw,
        "khmer_defects_after_repair": defects_left,
        "needs_ocr": bool(reasons),
        "reasons": reasons,
    }


# --- Repair of text extracted from PDFs --------------------------------------
# Many PDF producers (Chrome, Word, LibreOffice) map one shaped Khmer glyph to
# several characters, so extraction yields duplicated signs: "ព្រះះ", "សាាសនា",
# "ឆ្នាំាំ", "ទៅទៅ". Khmer never legitimately repeats a dependent sign, and
# literal syllable reduplication is written with "ៗ", so these are safe to fold.
_DUP_SIGNS_RE = re.compile(rf"([{_DEP}]{{1,3}})\1+")
_PRE_BASE = "េ-ៅ"  # េ ែ ៃ ោ ៅ (vowels drawn before / around the base)
_DUP_SYLLABLE_RE = re.compile(
    rf"([{_BASE}](?:្[{_BASE}])*[{_DEP}]*[{_PRE_BASE}][{_DEP}]*)\1"
)
_DUP_COENG_RE = re.compile(r"(្[ក-អ])\1+")
_DUP_ACROSS_BREAK_RE = re.compile(rf"([{_DEP}]{{1,3}})(\s+)\1(?![{_DEP}])")


def count_khmer_defects(text: str) -> int:
    text = text or ""
    return (len(_DUP_SIGNS_RE.findall(text)) + len(_DUP_SYLLABLE_RE.findall(text))
            + len(_DUP_COENG_RE.findall(text)) + len(_ORPHAN_SIGN_RE.findall(text)))


def repair_khmer(text: str) -> str:
    if not text or not has_khmer(text):
        return text
    prev = None
    while prev != text:
        prev = text
        text = _DUP_COENG_RE.sub(r"\1", text)
        text = _DUP_SIGNS_RE.sub(r"\1", text)
        text = _DUP_SYLLABLE_RE.sub(r"\1", text)
        # duplicate carried over a line break: "ព្រះ\nះរាជា" -> "ព្រះ\nរាជា"
        text = _DUP_ACROSS_BREAK_RE.sub(r"\1\2", text)
    return text


def repair_docx_text(docx_path: str) -> int:
    """Apply repair_khmer to every paragraph of a .docx. Works across runs of
    the same paragraph so duplicates split between runs are fixed too.
    Returns the number of paragraphs changed."""
    doc = Document(docx_path)
    changed = 0
    prev_tail = ""
    for p in doc.element.body.iter(qn("w:p")):
        ts = [t for t in p.iter(qn("w:t"))]
        if not ts:
            continue
        full = "".join(t.text or "" for t in ts)
        tail, prev_tail = prev_tail, full.rstrip()
        if not has_khmer(full):
            continue
        fixed = repair_khmer(full)
        # orphan sign duplicated from the end of the previous paragraph
        m = re.match(rf"[{_DEP}]{{1,3}}", fixed)
        if m and tail.endswith(m.group(0)):
            fixed = fixed[m.end():]
        if fixed == full:
            continue
        changed += 1
        # Re-distribute fixed text over runs: walk original text, keeping each
        # character that survives the repair in its original run.
        keep = _alignment_keep_mask(full, fixed)
        pos = 0
        for t in ts:
            s = t.text or ""
            t.text = "".join(c for i, c in enumerate(s) if keep[pos + i])
            pos += len(s)
            if t.text != t.text.strip():
                t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    doc.save(docx_path)
    return changed


def _alignment_keep_mask(original: str, fixed: str) -> list[bool]:
    """fixed is original with some characters deleted; mark which survive."""
    keep = [False] * len(original)
    j = 0
    for i, c in enumerate(original):
        if j < len(fixed) and fixed[j] == c:
            # Greedy match; deletions only ever remove repeated material so a
            # greedy left-to-right alignment is correct enough.
            remaining_orig = len(original) - i
            remaining_fixed = len(fixed) - j
            if remaining_orig >= remaining_fixed:
                keep[i] = True
                j += 1
    return keep


# Schema order of <w:rPr> children (Word rejects out-of-order elements).
_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
              "dstrike", "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid",
              "vanish", "webHidden", "color", "spacing", "w", "kern", "position", "sz", "szCs",
              "highlight", "u", "effect", "bdr", "shd", "fitText", "vertAlign", "rtl", "cs",
              "em", "lang", "eastAsianLayout", "specVanish", "oMath"]


def _get_or_add(rpr, name: str):
    el = rpr.find(qn("w:" + name))
    if el is not None:
        return el
    el = OxmlElement("w:" + name)
    idx = _RPR_ORDER.index(name)
    for i, child in enumerate(list(rpr)):
        tag = child.tag.split("}")[-1]
        if tag in _RPR_ORDER and _RPR_ORDER.index(tag) > idx:
            rpr.insert(i, el)
            return el
    rpr.append(el)
    return el


def is_khmer_font(name: str | None) -> bool:
    n = (name or "").lower()
    return any(k in n for k in ("khmer", "battambang", "siemreap", "moul", "muol", "hanuman",
                                "kantumruy", "bokor", "koulen", "dangrek", "fasthand", "metal",
                                "freehand", "suwannaphum", "angkor", "preahvihear", "taprom"))


def _set_run_font(rpr, font: str, khmer_run: bool):
    rfonts = _get_or_add(rpr, "rFonts")
    # Respect a Khmer font the document already uses (e.g. Khmer Muol titles).
    existing = rfonts.get(qn("w:cs")) or rfonts.get(qn("w:ascii"))
    if khmer_run and is_khmer_font(existing):
        font = existing
    # Word picks the "complex script" (cs) font for Khmer characters.
    rfonts.set(qn("w:cs"), font)
    for attr in ("w:csTheme",):
        if rfonts.get(qn(attr)) is not None:
            del rfonts.attrib[qn(attr)]
    if khmer_run:
        # Also set Latin slots so apps that ignore w:cs (Google Docs, some
        # LibreOffice versions) still use a Khmer-capable font.
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia"):
            rfonts.set(qn(attr), font)
        for attr in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme"):
            if rfonts.get(qn(attr)) is not None:
                del rfonts.attrib[qn(attr)]
        _get_or_add(rpr, "lang").set(qn("w:bidi"), "km-KH")
        # Mirror size / bold between Latin and complex-script slots so every
        # app (Word uses szCs for Khmer, Google Docs uses sz) agrees.
        for a, b in (("sz", "szCs"), ("b", "bCs"), ("i", "iCs")):
            ea, eb = rpr.find(qn("w:" + a)), rpr.find(qn("w:" + b))
            if ea is not None and eb is None:
                el = _get_or_add(rpr, b)
                for k, v in ea.attrib.items():
                    el.set(k, v)
            elif eb is not None and ea is None:
                el = _get_or_add(rpr, a)
                for k, v in eb.attrib.items():
                    el.set(k, v)


def _fix_part(element, font: str):
    for r in element.iter(qn("w:r")):
        text = "".join(t.text or "" for t in r.iter(qn("w:t")))
        if not text:
            continue
        khmer_run = has_khmer(text)
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            if not khmer_run:
                continue
            rpr = OxmlElement("w:rPr")
            r.insert(0, rpr)
        _set_run_font(rpr, font, khmer_run)


def apply_khmer_font(docx_path: str, font: str = "Khmer OS Battambang") -> None:
    """Rewrite a .docx so every Khmer run uses `font`, and set document defaults."""
    doc = Document(docx_path)

    # Document defaults (affects new text the user types in Word).
    styles_el = doc.styles.element
    doc_defaults = styles_el.find(qn("w:docDefaults"))
    if doc_defaults is not None:
        rpr_default = doc_defaults.find(qn("w:rPrDefault"))
        if rpr_default is None:
            rpr_default = OxmlElement("w:rPrDefault")
            doc_defaults.insert(0, rpr_default)
        rpr = rpr_default.find(qn("w:rPr"))
        if rpr is None:
            rpr = OxmlElement("w:rPr")
            rpr_default.append(rpr)
        _set_run_font(rpr, font, khmer_run=False)
        _get_or_add(rpr, "lang").set(qn("w:bidi"), "km-KH")

    _fix_part(doc.element.body, font)
    for section in doc.sections:
        for part in (section.header, section.footer, section.first_page_header,
                     section.first_page_footer, section.even_page_header, section.even_page_footer):
            try:
                if part is not None and not part.is_linked_to_previous:
                    _fix_part(part._element, font)
            except Exception:
                pass
    doc.save(docx_path)
