"""Side boxes: bordered rectangles that sit *beside* other text.

Typical example: the "បិទរូបថតថ្មី ៤ x ៦" photo box at the top right of
Cambodian forms, next to the national motto. pdf2docx turns such a box into a
table in the text flow, which breaks the reading order (text that is beside the
box ends up below it).

Fix:
1. find side boxes on each page,
2. remove them (border + inner text) from the copy of the PDF that pdf2docx
   reads, so the main text flows in the right order,
3. re-insert each box into the .docx as a floating table anchored to its
   exact page position, so Word places it where it was and wraps text around it.
"""
from __future__ import annotations

from pathlib import Path

import pymupdf
from docx import Document
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls, qn
from xml.sax.saxutils import escape

from . import khmer


# --------------------------------------------------------------------------- #
# Detection
# --------------------------------------------------------------------------- #
def _stroked_rects(page: pymupdf.Page) -> list[tuple[pymupdf.Rect, dict]]:
    out = []
    for dr in page.get_drawings():
        items = dr.get("items", [])
        if dr.get("color") is None:
            continue  # fill-only shape, not a bordered box
        if len(items) == 1 and items[0][0] == "re":
            out.append((pymupdf.Rect(items[0][1]), dr))
        elif len(items) == 4 and all(it[0] == "l" for it in items):
            r = dr["rect"]
            # four axis-aligned lines forming a closed rectangle
            if all(abs(it[1].x - it[2].x) < 0.5 or abs(it[1].y - it[2].y) < 0.5 for it in items):
                out.append((pymupdf.Rect(r), dr))
    return out


def find_side_boxes(page: pymupdf.Page) -> list[dict]:
    """Bordered rectangles that are isolated (not part of a table grid), not
    too wide, and have page text beside them on the same rows."""
    rects = _stroked_rects(page)
    all_paths = [pymupdf.Rect(d["rect"]) for d in page.get_drawings() if d.get("color") is not None]
    words = page.get_text("words")
    images = [pymupdf.Rect(i["bbox"]) for i in page.get_image_info()]
    pw = page.rect.width
    boxes = []
    for r, dr in rects:
        if r.width < 30 or r.height < 20 or r.width > 0.5 * pw:
            continue
        grown = r + (-3, -3, 3, 3)
        # part of a table / other line art? (anything touching the box that is
        # not the box itself)
        touching = [p for p in all_paths if p.intersects(grown) and not (abs(p.x0 - r.x0) < 1 and abs(p.y0 - r.y0) < 1 and abs(p.x1 - r.x1) < 1 and abs(p.y1 - r.y1) < 1)]
        touching = [p for p in touching if not r.contains(p)]  # inner decorations are fine
        if touching:
            continue
        if any(r.intersects(im) for im in images):
            continue  # boxes with pictures/QR codes: leave to pdf2docx
        # text beside the box: a word outside it whose vertical span overlaps it
        beside = 0
        for w in words:
            wr = pymupdf.Rect(w[:4])
            if r.contains(wr) or wr.intersects(r):
                continue
            ov = min(wr.y1, r.y1) - max(wr.y0, r.y0)
            if ov > 0.5 * wr.height:
                beside += 1
        if beside < 2:
            continue  # the box has its own row: pdf2docx handles that fine
        boxes.append({
            "rect": r,
            "stroke": dr.get("color") or (0, 0, 0),
            "width": dr.get("width") or 0.5,
        })
    return boxes


# --------------------------------------------------------------------------- #
# Extraction of the box content
# --------------------------------------------------------------------------- #
def _box_lines(page: pymupdf.Page, rect: pymupdf.Rect, hybrid_cfg: dict | None) -> list[dict]:
    raw = page.get_text("rawdict", clip=rect, flags=pymupdf.TEXT_MEDIABOX_CLIP)
    blocks = [b for b in raw.get("blocks", []) if b.get("type") == 0]
    if hybrid_cfg:
        from . import hybrid
        blocks = hybrid.fix_blocks(page, blocks, hybrid_cfg["lang"], hybrid_cfg["tessdata"],
                                   hybrid_cfg["tesseract"])
    lines = []
    for b in blocks:
        for ln in b.get("lines", []):
            spans = ln.get("spans", [])
            text = "".join(c["c"] for s in spans for c in s["chars"]).strip()
            if not text:
                continue
            text = khmer.repair_khmer(text)
            big = max(spans, key=lambda s: len(s["chars"]))
            lb = pymupdf.Rect(ln["bbox"])
            centre = (lb.x0 + lb.x1) / 2
            if abs(centre - (rect.x0 + rect.x1) / 2) < rect.width * 0.12:
                align = "center"
            elif lb.x0 - rect.x0 > rect.x1 - lb.x1 + rect.width * 0.2:
                align = "right"
            else:
                align = "left"
            lines.append({
                "text": " ".join(text.split()),
                "size": round(big.get("size", 11) * 2) / 2,
                "bold": bool(big.get("flags", 0) & 16),
                "font": big.get("font", ""),
                "align": align,
                "top": lb.y0,
                "bottom": lb.y1,
            })
    # merge fragments of the same visual line
    lines.sort(key=lambda l: (round(l["top"]), ))
    merged: list[dict] = []
    for l in lines:
        if merged and abs(merged[-1]["top"] - l["top"]) < 0.5 * (l["bottom"] - l["top"]):
            merged[-1]["text"] += " " + l["text"]
            merged[-1]["bottom"] = max(merged[-1]["bottom"], l["bottom"])
        else:
            merged.append(l)
    return merged


def _anchor_text(page: pymupdf.Page, rect: pymupdf.Rect) -> str:
    """Text of the first line outside the box that starts at/after the box top:
    the box is anchored to the paragraph holding it (for the editor's flow)."""
    best = None
    for b in page.get_text("dict").get("blocks", []):
        for ln in b.get("lines", []):
            lb = pymupdf.Rect(ln["bbox"])
            if lb.intersects(rect) and rect.contains(lb):
                continue
            text = "".join(s["text"] for s in ln["spans"]).strip()
            if not text or lb.y1 < rect.y0 - 2:
                continue
            if best is None or (lb.y0, lb.x0) < best[0]:
                best = ((lb.y0, lb.x0), text)
    return khmer.repair_khmer(best[1]) if best else ""


def extract_side_boxes(pdf: Path, cleaned_pdf: Path, hybrid_cfg: dict | None = None) -> dict[int, list[dict]]:
    """Write `cleaned_pdf` without the side boxes; return {page_index: [box]}."""
    found: dict[int, list[dict]] = {}
    doc = pymupdf.open(pdf)
    try:
        for pno, page in enumerate(doc):
            boxes = find_side_boxes(page)
            if not boxes:
                continue
            for box in boxes:
                box["lines"] = _box_lines(page, box["rect"], hybrid_cfg)
                box["anchor"] = _anchor_text(page, box["rect"])
            for box in boxes:
                g = box["width"] + 0.8
                page.add_redact_annot(box["rect"] + (-g, -g, g, g))
            page.apply_redactions(images=pymupdf.PDF_REDACT_IMAGE_NONE,
                                  graphics=pymupdf.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                                  text=pymupdf.PDF_REDACT_TEXT_REMOVE)
            found[pno] = [{**b, "rect": tuple(b["rect"])} for b in boxes]
        doc.save(cleaned_pdf)
    finally:
        doc.close()
    return found


# --------------------------------------------------------------------------- #
# Re-insertion into the .docx as floating tables
# --------------------------------------------------------------------------- #
def _tw(pt: float) -> int:
    return int(round(pt * 20))


def _hex(color) -> str:
    return "".join(f"{int(max(0, min(1, c)) * 255):02X}" for c in (color or (0, 0, 0))[:3])


def _box_table_xml(box: dict, vert: str = "page") -> str:
    """vert="page": exact page position (direct conversion).
    vert="text": top aligned with the anchor paragraph (after editing, the
    text may have moved, so the box follows its paragraph)."""
    x0, y0, x1, y1 = box["rect"]
    w, h = x1 - x0, y1 - y0
    sz = max(2, int(round(box["width"] * 8)))  # eighths of a point
    col = _hex(box["stroke"])
    border = f'w:val="single" w:sz="{sz}" w:space="0" w:color="{col}"'
    paras = []
    prev_bottom = y0
    for i, ln in enumerate(box["lines"]):
        before = max(0.0, ln["top"] - prev_bottom - (2 if i else 0))
        prev_bottom = ln["bottom"]
        font = ln["font"]
        is_kh = khmer.has_khmer(ln["text"])
        fattr = ""
        if font and "+" in font:
            font = font.split("+", 1)[1]
        if is_kh:
            fattr = ""  # Khmer font is applied later by khmer.apply_khmer_font
        elif font:
            fattr = f'<w:rFonts w:ascii="{escape(font)}" w:hAnsi="{escape(font)}"/>'
        b = "<w:b/><w:bCs/>" if ln["bold"] else ""
        half = int(round(ln["size"] * 2))
        paras.append(
            f'<w:p><w:pPr><w:spacing w:before="{_tw(before)}" w:after="0" w:line="240" w:lineRule="auto"/>'
            f'<w:jc w:val="{ln["align"]}"/></w:pPr>'
            f'<w:r><w:rPr>{fattr}{b}<w:sz w:val="{half}"/><w:szCs w:val="{half}"/></w:rPr>'
            f'<w:t xml:space="preserve">{escape(ln["text"])}</w:t></w:r></w:p>'
        )
    if not paras:
        paras.append("<w:p/>")
    return (
        f'<w:tbl {nsdecls("w")}><w:tblPr>'
        f'<w:tblpPr w:leftFromText="113" w:rightFromText="113" w:topFromText="0" w:bottomFromText="0" '
        f'w:vertAnchor="{vert}" w:horzAnchor="page" w:tblpX="{_tw(x0)}" '
        f'w:tblpY="{_tw(y0) if vert == "page" else 0}"/>'
        f'<w:tblOverlap w:val="never"/>'
        f'<w:tblW w:w="{_tw(w)}" w:type="dxa"/>'
        f'<w:tblBorders><w:top {border}/><w:left {border}/><w:bottom {border}/><w:right {border}/></w:tblBorders>'
        f'<w:tblLayout w:type="fixed"/>'
        f'<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="57" w:type="dxa"/>'
        f'<w:bottom w:w="0" w:type="dxa"/><w:right w:w="57" w:type="dxa"/></w:tblCellMar>'
        f'<w:tblLook w:val="0000"/></w:tblPr>'
        f'<w:tblGrid><w:gridCol w:w="{_tw(w)}"/></w:tblGrid>'
        f'<w:tr><w:trPr><w:trHeight w:val="{_tw(h)}" w:hRule="exact"/></w:trPr>'
        f'<w:tc><w:tcPr><w:tcW w:w="{_tw(w)}" w:type="dxa"/></w:tcPr>{"".join(paras)}</w:tc></w:tr></w:tbl>'
    )


def _para_text(el) -> str:
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _similar(a: str, b: str) -> float:
    import difflib
    a, b = "".join(a.split()), "".join(b.split())
    if not a or not b:
        return 0.0
    n = min(len(a), len(b), 40)
    return difflib.SequenceMatcher(None, a[:n], b[:n]).ratio()


def insert_side_boxes(docx: Path, boxes: dict[int, list[dict]]) -> int:
    """Insert each box as a floating table, anchored before the paragraph that
    holds the text next to it (Word positions it absolutely on the page; the
    anchor only matters for the editor's text flow)."""
    if not boxes:
        return 0
    doc = Document(docx)
    body = doc.element.body
    children = list(body)
    starts = [0]
    for i, el in enumerate(children):
        if el.tag == qn("w:p") and el.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
            starts.append(i + 1)
    starts.append(len(children))
    inserted = 0
    for pno, page_boxes in boxes.items():
        if pno + 1 >= len(starts):
            continue
        lo, hi = starts[pno], starts[pno + 1]
        paras = [el for el in children[lo:hi] if el.tag == qn("w:p")]
        if not paras:
            continue
        min_idx = 0  # boxes are processed top to bottom; anchors keep that order
        for box in sorted(page_boxes, key=lambda b: (b["rect"][1], b["rect"][0])):
            anchor, anchor_idx = None, min_idx
            if box.get("anchor"):
                scored = [(_similar(_para_text(p), box["anchor"]), i, p)
                          for i, p in enumerate(paras) if i >= min_idx and _para_text(p).strip()]
                if scored:
                    score, i, p = max(scored, key=lambda x: (x[0], -x[1]))
                    if score >= 0.45:
                        anchor, anchor_idx = p, i
            if anchor is None:
                anchor = paras[min(min_idx, len(paras) - 1)]
            min_idx = anchor_idx + 1
            anchor.addprevious(parse_xml(_box_table_xml(box)))
            inserted += 1
    doc.save(docx)
    return inserted


# --------------------------------------------------------------------------- #
# Editor round trip: HTML boxes <-> floating tables
# --------------------------------------------------------------------------- #
MARKER = "TWBOX{}TWBOX"


def pull_boxes_from_html(html: str) -> tuple[str, list[dict]]:
    """Replace editor box tables (class tw-box) with marker paragraphs and
    return their content, so LibreOffice does not turn them into frames."""
    from bs4 import BeautifulSoup
    import re

    soup = BeautifulSoup(html, "html.parser")
    found = []
    for tbl in soup.select("table.tw-box"):
        try:
            rect = tuple(float(tbl.get(f"data-{k}")) for k in ("x0", "y0", "x1", "y1"))
        except (TypeError, ValueError):
            continue
        lines = []
        for p in tbl.find_all(["p", "h1", "h2", "h3", "h4", "h5", "h6"]):
            text = " ".join(p.get_text(" ").split())
            if not text:
                continue
            style = (p.get("style") or "")
            m = re.search(r"text-align:\s*(\w+)", style)
            align = {"center": "center", "right": "right", "justify": "both"}.get(m.group(1) if m else "", "left")
            size = 11.0
            span = p.find(style=re.compile("font-size"))
            if span:
                sm = re.search(r"font-size:\s*([\d.]+)pt", span["style"])
                if sm:
                    size = float(sm.group(1))
            lines.append({"text": text, "align": align, "size": size,
                          "bold": p.find(["strong", "b"]) is not None, "font": "",
                          "top": 0, "bottom": 0})
        # keep the original vertical padding of the first line
        if lines:
            try:
                lines[0]["top"] = rect[1] + float(tbl.get("data-pad") or 0)
                lines[0]["bottom"] = lines[0]["top"]
                for l in lines[1:]:
                    l["top"] = l["bottom"] = lines[0]["top"]
            except ValueError:
                pass
        idx = len(found)
        found.append({"rect": rect, "lines": lines, "stroke": (0, 0, 0),
                      "width": float(tbl.get("data-bw") or 0.5)})
        marker = soup.new_tag("p")
        marker.string = MARKER.format(idx)
        tbl.replace_with(marker)
    return str(soup), found


def push_boxes_into_docx(docx: Path, found: list[dict]) -> None:
    """Replace marker paragraphs (see pull_boxes_from_html) by floating tables."""
    if not found:
        return
    doc = Document(docx)
    for p in list(doc.element.body.iter(qn("w:p"))):
        text = _para_text(p).strip()
        if text.startswith("TWBOX") and text.endswith("TWBOX"):
            try:
                idx = int(text[5:-5])
            except ValueError:
                continue
            if 0 <= idx < len(found):
                p.addprevious(parse_xml(_box_table_xml(found[idx], vert="text")))
            p.getparent().remove(p)
    doc.save(docx)
