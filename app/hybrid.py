"""Hybrid conversion: keep pdf2docx's layout, but replace Khmer text with OCR.

Many Khmer PDFs (especially Microsoft Word exports using Khmer OS fonts) have a
text layer whose Unicode mapping is wrong: glyphs are drawn correctly but the
extracted characters are wrong letters ("នាមត្រកូេ និងនាមខលួន" instead of
"នាមត្រកូល និងនាមខ្លួន"). The positions of the glyphs are still exact, so we:

1. let pdf2docx extract the raw text (with per-character boxes),
2. split every text line into Khmer "phrases" (separated by dot leaders, gaps),
3. render each phrase's exact box at high DPI and OCR it as a single line,
4. put the OCR text back in place, so pdf2docx builds the Word layout
   (tables, positions, fonts, sizes) around correct Unicode text.
"""
from __future__ import annotations

import difflib
import os
import re
import subprocess
import tempfile
import threading
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pymupdf

from . import khmer

_LEADERS = set(".…_-–—·•‥⋯|")
_state = threading.local()


# --------------------------------------------------------------------------- #
# OCR helpers
# --------------------------------------------------------------------------- #
def _run_tesseract_list(images: list[Path], lang: str, psm: str, tessdata: Path,
                        tesseract: str) -> dict[Path, str]:
    """OCR many single-line images with one tesseract process per chunk
    (the Khmer model is loaded once per process instead of once per image)."""
    results: dict[Path, str] = {}
    if not images:
        return results
    env = dict(os.environ, TESSDATA_PREFIX=str(tessdata), OMP_THREAD_LIMIT="1")
    workers = max(1, min(6, os.cpu_count() or 2, len(images)))
    chunks = [images[i::workers] for i in range(workers)]

    def run(chunk: list[Path]):
        if not chunk:
            return
        with tempfile.NamedTemporaryFile("w", suffix=".txt", delete=False) as lf:
            lf.write("\n".join(str(p) for p in chunk))
            listfile = lf.name
        try:
            res = subprocess.run([tesseract, listfile, "stdout", "-l", lang, "--psm", psm],
                                 capture_output=True, env=env, timeout=900)
            pages = res.stdout.decode("utf-8", errors="ignore").split("\f")
            for p, txt in zip(chunk, pages):
                results[p] = " ".join(txt.split())
        finally:
            os.unlink(listfile)

    with ThreadPoolExecutor(max_workers=workers) as ex:
        list(ex.map(run, chunks))
    return results


def _ocr_segments(items: list[tuple[Path, str]], base_lang: str, tessdata: Path,
                  tesseract: str) -> list[str]:
    """items: (image, original_text). Pure-Khmer phrases use the Khmer-only
    model (avoids Latin hallucinations); psm 13 (raw line) is the most reliable
    for tight crops, psm 7 is the fallback when it returns nothing."""
    by_lang: dict[str, list[Path]] = {}
    for img, orig in items:
        lang = base_lang
        if "khm" in base_lang and not re.search(r"[A-Za-z]", orig):
            lang = "khm"
        by_lang.setdefault(lang, []).append(img)
    out: dict[Path, str] = {}
    for lang, imgs in by_lang.items():
        out.update(_run_tesseract_list(imgs, lang, "13", tessdata, tesseract))
        empty = [i for i in imgs if not out.get(i)]
        if empty:
            out.update({k: v for k, v in
                        _run_tesseract_list(empty, lang, "7", tessdata, tesseract).items() if v})
    return [out.get(img, "") for img, _ in items]


_EDGE_PUNCT = set(".,:;'‘’\"“”`|/\\_-=~·•»«()[]{}<>!?")


def clean_ocr(ocr: str, orig: str) -> str:
    """Remove punctuation the OCR picked up from neighbouring dot leaders or
    table lines, unless the original phrase has it at the same edge."""
    t = ocr.strip()
    while t and t[0] in _EDGE_PUNCT and not orig.startswith(t[0]):
        t = t[1:].lstrip()
    while t and t[-1] in _EDGE_PUNCT and not orig.endswith(t[-1]):
        t = t[:-1].rstrip()
    return t


# Phrases found in almost every Cambodian official document. Decorative Muol
# fonts are hard to OCR, so close matches are snapped to the exact spelling.
COMMON_PHRASES = [
    "ព្រះរាជាណាចក្រកម្ពុជា", "ជាតិ សាសនា ព្រះមហាក្សត្រ", "ព្រះមហាក្សត្រ",
    "រាជរដ្ឋាភិបាលកម្ពុជា", "ក្រសួងអប់រំ យុវជន និងកីឡា", "ក្រសួងសុខាភិបាល",
    "មករា", "កុម្ភៈ", "មីនា", "មេសា", "ឧសភា", "មិថុនា", "កក្កដា", "សីហា", "កញ្ញា",
    "តុលា", "វិច្ឆិកា",
]


def snap_common(text: str, *candidates: str) -> str:
    best, score = None, 0.0
    for phrase in COMMON_PHRASES:
        target = phrase.replace(" ", "")
        for cand in candidates:
            if cand:
                r = difflib.SequenceMatcher(None, cand.replace(" ", ""), target).ratio()
                if r > score:
                    best, score = phrase, r
    if best and score >= 0.85 and abs(len(best) - len(text)) <= max(2, len(best) // 5):
        return best
    return text


# --------------------------------------------------------------------------- #
# Phrase segmentation and replacement
# --------------------------------------------------------------------------- #
def _line_segments(line: dict):
    """Split a text line (across all its spans) into phrases.

    Yields lists of (span_index, char_index). Phrases break at dot leaders,
    runs of 2+ spaces and wide horizontal gaps. They may cross span
    boundaries (Word often splits one word, e.g. ឈ្មោះ, over two spans)."""
    flat = []
    for si, span in enumerate(line.get("spans", [])):
        for ci, ch in enumerate(span.get("chars", [])):
            flat.append((si, ci, ch))
    seg: list[tuple[int, int, dict]] = []
    out: list[list[tuple[int, int]]] = []

    def flush():
        while seg and seg[-1][2]["c"].isspace():
            seg.pop()
        if seg:
            out.append([(a, b) for a, b, _ in seg])
        seg.clear()

    prev = None
    for k, (si, ci, ch) in enumerate(flat):
        c = ch["c"]
        if c in _LEADERS:
            flush()
            prev = None
            continue
        if c.isspace():
            if not seg:
                continue
            nxt = flat[k + 1][2]["c"] if k + 1 < len(flat) else ""
            if nxt.isspace() or nxt in _LEADERS:
                flush()
                prev = None
            else:
                seg.append((si, ci, ch))
            continue
        if prev is not None:
            h = max(ch["bbox"][3] - ch["bbox"][1], 1)
            if ch["bbox"][0] - prev["bbox"][2] > 0.9 * h:
                flush()
        seg.append((si, ci, ch))
        prev = ch
    flush()
    return out


def _new_chars(old: list[dict], text: str) -> list[dict]:
    """Chars for `text`, spread evenly over the box of the replaced chars
    (pdf2docx only needs plausible geometry)."""
    x0 = min(c["bbox"][0] for c in old)
    y0 = min(c["bbox"][1] for c in old)
    x1 = max(c["bbox"][2] for c in old)
    y1 = max(c["bbox"][3] for c in old)
    oy = old[0]["origin"][1]
    w = (x1 - x0) / max(len(text), 1)
    return [{"c": ch, "bbox": (x0 + k * w, y0, x0 + (k + 1) * w, y1), "origin": (x0 + k * w, oy)}
            for k, ch in enumerate(text)]


def _pad_and_save(pix: pymupdf.Pixmap, path: Path, pad: int = 12):
    big = pymupdf.Pixmap(pymupdf.csGRAY, pymupdf.IRect(0, 0, pix.width + 2 * pad, pix.height + 2 * pad), False)
    big.set_rect(big.irect, (255,))
    pix.set_origin(pad, pad)
    big.copy(pix, pix.irect)
    big.save(path)


def _collect_phrases(blocks: list[dict]) -> list[dict]:
    """All phrases on a page, merged across PyMuPDF line boundaries when two
    phrases sit on the same baseline and touch (Word often splits one word,
    e.g. ឈ្មោះ, into separate text runs that PyMuPDF reports as two lines)."""
    phrases = []
    for bi, block in enumerate(blocks):
        for li, line in enumerate(block.get("lines", [])):
            d = tuple(round(v, 3) for v in line.get("dir", (1, 0)))
            if d != (1.0, 0.0):
                continue  # rotated text: leave as is
            spans = line.get("spans", [])
            for seg in _line_segments(line):
                chs = [spans[si]["chars"][ci] for si, ci in seg]
                phrases.append({
                    "keys": [(bi, li, si, ci) for si, ci in seg],
                    "chars": chs,
                    "fonts": {spans[si].get("font", "") for si, _ in seg},
                    "x0": min(c["bbox"][0] for c in chs), "x1": max(c["bbox"][2] for c in chs),
                    "y0": min(c["bbox"][1] for c in chs), "y1": max(c["bbox"][3] for c in chs),
                })
    def joinable(q, p) -> bool:
        """p continues q to the right on the same row."""
        h = max(q["y1"] - q["y0"], p["y1"] - p["y0"], 1)
        qc, pc = (q["y0"] + q["y1"]) / 2, (p["y0"] + p["y1"]) / 2
        overlap = min(q["y1"], p["y1"]) - max(q["y0"], p["y0"])
        if abs(qc - pc) > 0.35 * h and overlap < 0.6 * min(q["y1"] - q["y0"], p["y1"] - p["y0"]):
            return False
        gap = p["x0"] - q["x1"]
        left, right = q["chars"][-1]["c"], p["chars"][0]["c"]
        if left == "\u17D2":
            limit = 1.5 * h   # a word never ends with a coeng: always joined
        elif khmer.has_khmer(left) and khmer.has_khmer(right):
            limit = 0.5 * h
        else:
            limit = 0.35 * h
        return -0.5 * h < gap < limit

    phrases.sort(key=lambda p: p["x0"])
    merged: list[dict] = []
    for p in phrases:
        target = None
        for q in merged:
            if joinable(q, p) and (target is None or q["x1"] > target["x1"]):
                target = q
        if target is None:
            merged.append(p)
            continue
        q = target
        q["keys"] += p["keys"]
        q["chars"] += p["chars"]
        q["fonts"] |= p["fonts"]
        q["x0"], q["x1"] = min(q["x0"], p["x0"]), max(q["x1"], p["x1"])
        q["y0"], q["y1"] = min(q["y0"], p["y0"]), max(q["y1"], p["y1"])
    return merged


def fix_blocks(page: pymupdf.Page, blocks: list[dict], lang: str, tessdata: Path,
               tesseract: str, dpi: int = 300, stats: dict | None = None) -> list[dict]:
    jobs = []   # (phrase, original)
    items = []  # (image, original)
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        for ph in _collect_phrases(blocks):
            orig = "".join(c["c"] for c in ph["chars"])
            kfont = any(khmer.is_khmer_font(f) for f in ph["fonts"])
            if not (khmer.has_khmer(orig) or (kfont and re.search(r"[A-Za-z]", orig))):
                continue
            x0, y0, x1, y1 = ph["x0"], ph["y0"], ph["x1"], ph["y1"]
            h = y1 - y0
            if h <= 0 or x1 <= x0:
                continue
            clip = pymupdf.Rect(x0 - 1.0, y0 - 0.12 * h, x1 + 1.0, y1 + 0.12 * h) & page.rect
            if clip.is_empty:
                continue
            pix = page.get_pixmap(clip=clip, dpi=dpi if h >= 10 else int(dpi * 1.5),
                                  colorspace=pymupdf.csGRAY)
            f = tmp / f"s{len(items):05d}.png"
            _pad_and_save(pix, f)
            items.append((f, orig))
            jobs.append((ph, orig))
        texts = _ocr_segments(items, lang, tessdata, tesseract)

    sims: list[tuple[float, int]] = []
    insert_at: dict[tuple, list[dict]] = {}
    drop: set[tuple] = set()
    changed = 0
    for (ph, orig), txt in zip(jobs, texts):
        txt = clean_ocr(txt, orig)
        if not txt:
            continue
        if khmer.has_khmer(orig) and not khmer.has_khmer(txt):
            continue  # OCR noise for a Khmer phrase: keep the original
        txt = snap_common(txt, txt, khmer.repair_khmer(orig))
        sims.append((difflib.SequenceMatcher(None, khmer.repair_khmer(orig), txt).ratio(), len(orig)))
        # new text goes where the phrase starts; its other chars are dropped
        insert_at[ph["keys"][0]] = _new_chars(ph["chars"], txt)
        drop.update(ph["keys"])
        changed += 1

    touched = {(bi, li, si) for bi, li, si, _ in drop}
    for bi, li, si in touched:
        span = blocks[bi]["lines"][li]["spans"][si]
        new = []
        for ci, ch in enumerate(span["chars"]):
            key = (bi, li, si, ci)
            if key in insert_at:
                new.extend(insert_at[key])
            elif key not in drop:
                new.append(ch)
        span["chars"] = new
    if stats is not None:
        stats["segments"] = stats.get("segments", 0) + len(jobs)
        stats["replaced"] = stats.get("replaced", 0) + changed
        stats.setdefault("similarities", []).extend(sims)
    return blocks


# --------------------------------------------------------------------------- #
# pdf2docx integration
# --------------------------------------------------------------------------- #
_patched = False
_patch_lock = threading.Lock()


def _install_patch():
    global _patched
    with _patch_lock:
        if _patched:
            return
        from pdf2docx.page.RawPageFitz import RawPageFitz

        original = RawPageFitz._preprocess_text

        def patched(self, **settings):
            blocks = original(self, **settings)
            cfg = getattr(_state, "cfg", None)  # only active inside pdf_to_docx_hybrid
            if cfg and self.page_engine is not None:
                try:
                    blocks = fix_blocks(self.page_engine, blocks, cfg["lang"], cfg["tessdata"],
                                        cfg["tesseract"], stats=cfg["stats"])
                except Exception:
                    import traceback
                    traceback.print_exc()
                if cfg.get("progress"):
                    cfg["done"] += 1
                    cfg["progress"](cfg["done"], cfg["total"])
            return blocks

        RawPageFitz._preprocess_text = patched
        _patched = True


def pdf_to_docx_hybrid(pdf: Path, out: Path, lang: str, tessdata: Path, tesseract: str,
                       progress=None) -> dict:
    from pdf2docx import Converter

    _install_patch()
    stats: dict = {}
    with pymupdf.open(pdf) as d:
        total = d.page_count
    _state.cfg = {"lang": lang, "tessdata": tessdata, "tesseract": tesseract,
                  "stats": stats, "progress": progress, "done": 0, "total": total}
    try:
        cv = Converter(str(pdf))
        try:
            cv.convert(str(out), multi_processing=False, parse_stream_table=False)
        finally:
            cv.close()
    finally:
        _state.cfg = None
    return stats


def probe_text_layer(pdf: Path, lang: str, tessdata: Path, tesseract: str,
                     max_lines: int = 40) -> float | None:
    """OCR a sample of Khmer phrases and compare them with the (repaired) text
    layer. Returns the length-weighted similarity 0..1, or None if the PDF has
    too little Khmer text to judge."""
    with pymupdf.open(pdf) as doc:
        for page in doc:
            raw = page.get_text("rawdict", flags=pymupdf.TEXT_MEDIABOX_CLIP)
            picked, count = [], 0
            for b in raw.get("blocks", []):
                if b.get("type") != 0 or count >= max_lines:
                    continue
                lines = []
                for line in b.get("lines", []):
                    if count >= max_lines:
                        break
                    if khmer.has_khmer("".join(c["c"] for s in line["spans"] for c in s["chars"])):
                        lines.append(line)
                        count += 1
                if lines:
                    picked.append({**b, "lines": lines})
            if count < 3:
                continue
            stats: dict = {}
            fix_blocks(page, picked, lang, tessdata, tesseract, stats=stats)
            sims = [(r, n) for r, n in stats.get("similarities", []) if n >= 4]
            total = sum(n for _, n in sims)
            if total >= 40:
                return sum(r * n for r, n in sims) / total
    return None
