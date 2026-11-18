"""Conversion pipeline: PDF / image / office file -> DOCX, DOCX <-> HTML (editor)."""
from __future__ import annotations

import base64
import os
import re
import shutil
import statistics
import subprocess
import tempfile
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path

import pymupdf
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.shared import Pt, Cm

from . import khmer

ROOT = Path(__file__).resolve().parent.parent
TESSDATA = ROOT / "tessdata"
FONTS_DIR = ROOT / "fonts"
SOFFICE = shutil.which("soffice") or shutil.which("libreoffice")
TESSERACT = shutil.which("tesseract")

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp", ".webp"}
OFFICE_EXT = {".doc", ".odt", ".rtf", ".txt", ".html", ".htm"}


class ConversionError(Exception):
    pass


# --------------------------------------------------------------------------- #
# LibreOffice helper
# --------------------------------------------------------------------------- #
def soffice_convert(src: Path, out_dir: Path, target: str, infilter: str | None = None) -> Path:
    if not SOFFICE:
        raise ConversionError("LibreOffice (soffice) is not installed.")
    profile = Path(tempfile.mkdtemp(prefix="lo-profile-"))
    # Bundled open-licensed Khmer fonts are made visible to LibreOffice through
    # its private profile, so PDFs render Khmer even if they are not installed.
    (profile / "user").mkdir()
    if FONTS_DIR.is_dir():
        (profile / "user" / "fonts").symlink_to(FONTS_DIR)
    cmd = [SOFFICE, f"-env:UserInstallation={profile.as_uri()}", "--headless"]
    if infilter:
        cmd.append(f"--infilter={infilter}")
    cmd += ["--convert-to", target, "--outdir", str(out_dir), str(src)]
    try:
        subprocess.run(cmd, check=True, capture_output=True, timeout=180)
    except subprocess.CalledProcessError as e:
        raise ConversionError(f"LibreOffice failed: {e.stderr.decode(errors='ignore')[:300]}")
    finally:
        shutil.rmtree(profile, ignore_errors=True)
    ext = target.split(":")[0]
    out = out_dir / (src.stem + "." + ext)
    if not out.exists():
        raise ConversionError("LibreOffice produced no output.")
    return out


# --------------------------------------------------------------------------- #
# PDF analysis
# --------------------------------------------------------------------------- #
def analyse_pdf(pdf: Path) -> dict:
    with pymupdf.open(pdf) as doc:
        if doc.needs_pass:
            raise ConversionError("This PDF is password-protected.")
        text = "".join(page.get_text() for page in doc)
        q = khmer.text_quality(text, doc.page_count)
        q["pages"] = doc.page_count
        return q


# --------------------------------------------------------------------------- #
# Layout-preserving conversion (text layer)
# --------------------------------------------------------------------------- #
def pdf_to_docx_layout(pdf: Path, out: Path) -> None:
    from pdf2docx import Converter

    with tempfile.TemporaryDirectory() as td:
        clean, side_boxes = _without_side_boxes(pdf, Path(td), None)
        cv = Converter(str(clean))
        try:
            # Borderless "stream" tables split Khmer lines into many cells; keep
            # only real (ruled) tables.
            cv.convert(str(out), multi_processing=False, parse_stream_table=False)
        finally:
            cv.close()
    _cleanup_pdf2docx(out)
    khmer.repair_docx_text(str(out))
    _put_side_boxes_back(out, side_boxes)


def _without_side_boxes(pdf: Path, tmp: Path, hybrid_cfg: dict | None):
    """Copy of the PDF without side boxes (see boxes.py) + the removed boxes."""
    from . import boxes
    try:
        clean = tmp / "clean.pdf"
        found = boxes.extract_side_boxes(pdf, clean, hybrid_cfg)
        if found:
            return clean, found
    except Exception:
        import traceback
        traceback.print_exc()
    return pdf, {}


def _put_side_boxes_back(docx: Path, side_boxes: dict) -> None:
    from . import boxes
    if side_boxes:
        boxes.insert_side_boxes(docx, side_boxes)


def pdf_to_docx_hybrid(pdf: Path, out: Path, lang: str = "khm+eng", progress=None) -> dict:
    """Layout from pdf2docx + Khmer text re-read by OCR phrase by phrase."""
    from . import hybrid
    if not TESSERACT:
        raise ConversionError("Tesseract OCR is not installed.")
    cfg = {"lang": lang, "tessdata": TESSDATA, "tesseract": TESSERACT}
    with tempfile.TemporaryDirectory() as td:
        clean, side_boxes = _without_side_boxes(pdf, Path(td), cfg)
        stats = hybrid.pdf_to_docx_hybrid(clean, out, lang, TESSDATA, TESSERACT, progress=progress)
    _cleanup_pdf2docx(out)
    khmer.repair_docx_text(str(out))
    _put_side_boxes_back(out, side_boxes)
    stats["side_boxes"] = sum(len(v) for v in side_boxes.values())
    return stats


def khmer_text_layer_score(pdf: Path, lang: str = "khm+eng") -> float | None:
    """1.0 = Khmer text layer agrees with OCR, low = broken text layer."""
    from . import hybrid
    if not TESSERACT:
        return None
    try:
        return hybrid.probe_text_layer(pdf, lang, TESSDATA, TESSERACT)
    except Exception:
        return None


def _cleanup_pdf2docx(docx: Path) -> None:
    """Fix common pdf2docx artefacts: empty font names, and words emitted as
    separate runs with no space between them ("ThisisanEnglish...")."""
    from docx.oxml.ns import qn
    import re

    doc = Document(docx)
    word_end = re.compile(r"[A-Za-z0-9,.;:!?)\]'\"]$")
    word_start = re.compile(r"^[A-Za-z0-9(\['\"]")
    for p in doc.element.body.iter(qn("w:p")):
        prev_t = None
        for r in p.iter(qn("w:r")):
            rfonts = r.find(qn("w:rPr") + "/" + qn("w:rFonts")) if r.find(qn("w:rPr")) is not None else None
            if rfonts is not None:
                for k in list(rfonts.attrib):
                    if rfonts.attrib[k] == "":
                        del rfonts.attrib[k]
            t = r.find(qn("w:t"))
            if t is None or not t.text:
                if r.find(qn("w:br")) is not None or r.find(qn("w:tab")) is not None:
                    prev_t = None
                continue
            if prev_t is not None and word_end.search(prev_t.text or "") and word_start.search(t.text):
                prev_t.text = prev_t.text + " "
                prev_t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            prev_t = t
    doc.save(docx)


# --------------------------------------------------------------------------- #
# OCR conversion (scanned PDFs, images, broken Khmer text layers)
# --------------------------------------------------------------------------- #
def _ocr_image(png: Path, lang: str) -> list[dict]:
    if not TESSERACT:
        raise ConversionError("Tesseract OCR is not installed.")
    env = dict(os.environ, TESSDATA_PREFIX=str(TESSDATA), OMP_THREAD_LIMIT="1")
    res = subprocess.run(
        [TESSERACT, str(png), "stdout", "-l", lang, "--psm", "3",
         "-c", "preserve_interword_spaces=1", "tsv"],
        capture_output=True, env=env, timeout=600,
    )
    if res.returncode != 0:
        raise ConversionError("OCR failed: " + res.stderr.decode(errors="ignore")[:300])
    rows = res.stdout.decode("utf-8", errors="ignore").splitlines()[1:]
    paras: dict[tuple, dict] = {}
    page_w = page_h = 1
    for row in rows:
        c = row.split("\t")
        if len(c) < 12:
            continue
        level = int(c[0])
        left, top, w, h = map(int, c[6:10])
        if level == 1:
            page_w, page_h = w, h
            continue
        if level != 5:
            continue
        word = c[11]
        if not word.strip():
            continue
        key = (int(c[2]), int(c[3]))
        p = paras.setdefault(key, {"lines": {}, "box": [left, top, left + w, top + h]})
        p["lines"].setdefault(int(c[4]), []).append((left, top, w, h, word))
        b = p["box"]
        b[0], b[1] = min(b[0], left), min(b[1], top)
        b[2], b[3] = max(b[2], left + w), max(b[3], top + h)

    out = []
    for key in sorted(paras, key=lambda k: (paras[k]["box"][1], paras[k]["box"][0])):
        p = paras[key]
        lines, heights = [], []
        for ln in sorted(p["lines"]):
            words = sorted(p["lines"][ln], key=lambda x: x[0])
            lx0 = words[0][0]
            lx1 = max(w[0] + w[2] for w in words)
            lines.append((" ".join(w[4] for w in words), lx0, lx1))
            heights += [w[3] for w in words]
        out.append({
            "lines": lines, "box": p["box"],
            "line_h": _percentile(heights, 0.9) if heights else 0,
            "page_w": page_w, "page_h": page_h,
        })
    return out


def _percentile(values, q):
    v = sorted(values)
    return v[min(len(v) - 1, int(len(v) * q))]


def _join_lines(lines, content_l, content_r) -> list[str]:
    """Join wrapped lines into one string; keep short lines (headings, address
    blocks, centred titles) as separate lines (hard breaks)."""
    width = max(content_r - content_l, 1)
    segs: list[str] = []
    prev_full = False
    for text, x0, x1 in lines:
        if segs and prev_full:
            last = segs[-1]
            if khmer.has_khmer(last[-1:]) and khmer.has_khmer(text[:1]):
                segs[-1] = last + text  # Khmer wraps without spaces
            else:
                segs[-1] = last + " " + text
        else:
            segs.append(text)
        prev_full = (x1 - x0) > width * 0.8
    return segs


def _render_pages(pdf: Path, tmp: Path, dpi: int = 300) -> list[Path]:
    pngs = []
    with pymupdf.open(pdf) as doc:
        for i, page in enumerate(doc):
            pix = page.get_pixmap(dpi=dpi, colorspace=pymupdf.csGRAY)
            f = tmp / f"page-{i:04d}.png"
            pix.save(f)
            pngs.append(f)
    return pngs


def ocr_to_docx(src: Path, out: Path, lang: str = "khm+eng", dpi: int = 300,
                progress=None) -> None:
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        if src.suffix.lower() in IMAGE_EXT:
            pngs = [src]
            dpi = 300
        else:
            pngs = _render_pages(src, tmp, dpi)
        done = 0
        results: list[list[dict]] = [None] * len(pngs)  # type: ignore
        with ThreadPoolExecutor(max_workers=max(1, min(4, os.cpu_count() or 2))) as ex:
            futs = {ex.submit(_ocr_image, p, lang): i for i, p in enumerate(pngs)}
            for fut in futs:
                results[futs[fut]] = fut.result()
                done += 1
                if progress:
                    progress(done, len(pngs))

    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(2)
    sec.top_margin = sec.bottom_margin = Cm(2)
    normal = doc.styles["Normal"]
    normal.font.size = Pt(12)

    for pi, paras in enumerate(results):
        if pi > 0:
            doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
        if not paras:
            continue
        lefts = [p["box"][0] for p in paras]
        rights = [p["box"][2] for p in paras]
        content_l, content_r = min(lefts), max(rights)
        content_c = (content_l + content_r) / 2
        width = max(content_r - content_l, 1)

        def est_size(p):
            text_all = " ".join(t for t, _, _ in p["lines"])
            pt = p["line_h"] / dpi * 72
            factor = 1.17 if khmer.ratio_khmer(text_all) > 0.5 else 0.9
            return pt / factor

        sizes = [est_size(p) for p in paras]
        body = statistics.median(sizes)
        for p, raw in zip(paras, sizes):
            l, t, r, b = p["box"]
            segs = _join_lines(p["lines"], content_l, content_r)
            text_all = " ".join(segs)
            para = doc.add_paragraph()
            # Font size from the ink height of words (calibrated: Khmer glyph
            # boxes with stacked vowels are ~1.17 em, Latin ~0.9 em).
            # small fragments give noisy estimates: snap near-body sizes to body
            size = body if raw < body * 1.15 else raw
            size = max(8, min(40, round(size)))
            for i, seg in enumerate(segs):
                run = para.add_run(seg)
                run.font.size = Pt(size)
                if i < len(segs) - 1:
                    run.add_break()
            centre = (l + r) / 2
            if abs(centre - content_c) < width * 0.05 and (l - content_l) > width * 0.1:
                para.alignment = WD_ALIGN_PARAGRAPH.CENTER
            elif (l - content_l) > width * 0.45 and abs(r - content_r) < width * 0.05:
                para.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            elif len(segs) == 1 and len(p["lines"]) > 1:
                para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    doc.save(out)


# --------------------------------------------------------------------------- #
# Other inputs
# --------------------------------------------------------------------------- #
def office_to_docx(src: Path, out: Path) -> None:
    res = soffice_convert(src, out.parent, 'docx:MS Word 2007 XML')
    if res != out:
        shutil.move(res, out)


# --------------------------------------------------------------------------- #
# DOCX <-> HTML for the in-browser editor
# --------------------------------------------------------------------------- #
def docx_to_html(docx: Path) -> str:
    """DOCX -> editor HTML, keeping Khmer fonts (see docx2html.py)."""
    from .docx2html import docx_to_html as convert
    try:
        return convert(docx)
    except Exception:
        import traceback
        traceback.print_exc()
        return _docx_to_html_mammoth(docx)


def _docx_to_html_mammoth(docx: Path) -> str:
    import mammoth
    with open(docx, "rb") as f:
        return mammoth.convert_to_html(f).value


def _inline_images(html: str, base: Path) -> str:
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    mime = {".png": "image/png", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
            ".gif": "image/gif", ".svg": "image/svg+xml", ".bmp": "image/bmp"}
    for img in soup.find_all("img"):
        src = img.get("src", "")
        if src.startswith("data:") or src.startswith("http"):
            continue
        f = base / src
        if f.exists():
            m = mime.get(f.suffix.lower(), "application/octet-stream")
            img["src"] = f"data:{m};base64," + base64.b64encode(f.read_bytes()).decode()
    return str(soup)


_DROP_PROPS = {"writing-mode", "direction", "line-height", "font-family", "position",
               "left", "top", "right", "bottom"}
_DEFAULT_VALUES = {("font-style", "normal"), ("font-weight", "normal"), ("color", "#000000"),
                   ("text-align", "left"), ("text-decoration", "none")}


def _clean_decls(decls: str) -> dict[str, str]:
    out = {}
    for d in decls.split(";"):
        if ":" not in d:
            continue
        k, v = d.split(":", 1)
        k, v = k.strip().lower(), v.replace("!important", "").replace("! important", "").strip()
        if not k or not v or k in _DROP_PROPS or (k, v.lower()) in _DEFAULT_VALUES:
            continue
        # Zero margins/indents are noise.
        if k.startswith(("margin", "padding", "text-indent")) and re.fullmatch(r"-?0(\.0+)?(in|cm|pt|px)?", v):
            continue
        out[k] = v
    return out


def _is_khmer_font(family: str) -> bool:
    return khmer.is_khmer_font(family)


def _extract_body(html: str) -> str:
    """Inline LibreOffice's class-based CSS into style attributes so the HTML is
    self-contained and editable (no !important rules locking alignment, no
    tight line-heights that clip Khmer vowels and subscripts)."""
    from bs4 import BeautifulSoup

    soup = BeautifulSoup(html, "html.parser")
    rules: dict[str, dict[str, str]] = {}
    for st in soup.find_all("style"):
        for sel, body in re.findall(r"([^{}]+)\{([^{}]*)\}", st.get_text()):
            for s in sel.split(","):
                s = s.strip()
                m = re.fullmatch(r"(?:[a-z0-9]+)?\.([\w-]+)", s)
                if m:
                    r = rules.setdefault(m.group(1), {})
                    r.update(_clean_decls(body))
                    fam = re.search(r"font-family\s*:\s*([^;]+)", body)
                    if fam and _is_khmer_font(fam.group(1)):
                        r["font-family"] = fam.group(1).replace("!important", "").strip()
        st.decompose()

    # Keep font-family only when it is a Khmer font or when it differs from the
    # default (pdf2docx output names the real fonts; LO puts them on spans).
    for el in soup.find_all(True):
        classes = el.get("class") or []
        style: dict[str, str] = {}
        for c in classes:
            style.update(rules.get(c, {}))
        if el.get("style"):
            for k, v in _clean_decls(el["style"]).items():
                style[k] = v
            fam = re.search(r"font-family\s*:\s*([^;]+)", el["style"])
            if fam:
                style["font-family"] = fam.group(1).strip()
        if classes:
            del el["class"]
        if el.has_attr("dir"):
            del el["dir"]
        if style:
            el["style"] = "; ".join(f"{k}: {v}" for k, v in style.items())
        elif el.has_attr("style"):
            del el["style"]
    for a in soup.find_all("a"):
        if a.get("name") and not a.get_text(strip=True) and not a.find(True):
            a.decompose()  # bookmark anchors
    body = soup.body
    return (body.decode_contents() if body else str(soup)).strip()


def html_to_docx(html: str, out: Path, font: str | None, setup: dict | None = None) -> None:
    from . import boxes
    html = html.replace("<!-- pagebreak -->", '<p style="page-break-before: always; margin: 0"></p>')
    # Side boxes are re-created as exact floating tables after LibreOffice
    # (LibreOffice would turn a floated table into a loosely placed frame).
    html, side_boxes = boxes.pull_boxes_from_html(html)
    doc_html = (
        "<!DOCTYPE html><html><head><meta charset='utf-8'>"
        "<style>body{font-family:'%s';}</style></head><body>%s</body></html>"
        % (font or "Khmer OS Battambang", html)
    )
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td)
        src = tmp / "edited.html"
        src.write_text(doc_html, "utf-8")
        if SOFFICE:
            res = soffice_convert(src, tmp, "docx:MS Word 2007 XML", infilter="HTML (StarWriter)")
            shutil.copy(res, out)
        else:  # degraded fallback
            from htmldocx import HtmlToDocx
            d = Document()
            HtmlToDocx().add_html_to_document(html, d)
            d.save(out)
    boxes.push_boxes_into_docx(out, side_boxes)
    if font:
        khmer.apply_khmer_font(str(out), font)
    apply_page_setup(out, setup)


# Khmer OS fonts are the de-facto standard in Cambodian Word documents but are
# often not installed on the server. The Google Fonts below are the same
# designs (by the same designer) and are bundled in ./fonts.
PDF_FONT_SUBSTITUTES = {
    "khmer os battambang": "Battambang",
    "khmer os siemreap": "Siemreap",
    "khmer os": "Battambang",
    "khmer os system": "Battambang",
    "khmer os content": "Hanuman",
    "khmer os muol light": "Moul",
    "khmer os muol": "Moul",
    "khmer muol": "Moul",
    "khmer os muol pali": "Moul",
    "khmer os bokor": "Moul",
    "khmer os metal chrieng": "Koulen",
}
_installed_fonts: set[str] | None = None


def installed_font_families() -> set[str]:
    global _installed_fonts
    if _installed_fonts is None:
        fams: set[str] = set()
        fc = shutil.which("fc-list")
        if fc:
            try:
                res = subprocess.run([fc, ":", "family"], capture_output=True, timeout=30)
                for line in res.stdout.decode(errors="ignore").splitlines():
                    for f in line.split(","):
                        fams.add(f.strip().replace("\\", "").lower())
            except Exception:
                pass
        fams |= {"battambang", "siemreap", "moul", "hanuman", "koulen", "kantumruy pro",
                 "noto sans khmer", "noto serif khmer"}
        _installed_fonts = fams
    return _installed_fonts


def substitute_missing_fonts(docx: Path) -> None:
    from docx.oxml.ns import qn

    have = installed_font_families()
    doc = Document(docx)
    attrs = [qn("w:ascii"), qn("w:hAnsi"), qn("w:cs"), qn("w:eastAsia")]
    for root in (doc.element.body, doc.styles.element):
        for rf in root.iter(qn("w:rFonts")):
            for a in attrs:
                v = rf.get(a)
                if not v:
                    continue
                # Always use the modern builds of Khmer OS designs: old Khmer OS
                # .ttf files often lack proper script tables, and LibreOffice
                # then falls back to another font even when they are installed.
                sub = PDF_FONT_SUBSTITUTES.get(v.lower())
                if sub is None and khmer.is_khmer_font(v) and v.lower() not in have:
                    sub = "Battambang"
                if sub:
                    rf.set(a, sub)
    doc.save(docx)


def page_setup(docx: Path) -> dict:
    """Page size and margins (in pt) of the first section."""
    sec = Document(docx).sections[0]
    def pt(v, default):
        return round(v.pt, 1) if v is not None else default
    return {
        "width": pt(sec.page_width, 595.3), "height": pt(sec.page_height, 841.9),
        "left": pt(sec.left_margin, 56.7), "right": pt(sec.right_margin, 56.7),
        "top": pt(sec.top_margin, 56.7), "bottom": pt(sec.bottom_margin, 56.7),
    }


def apply_page_setup(docx: Path, setup: dict | None) -> None:
    if not setup:
        return
    from docx.shared import Pt as _Pt
    doc = Document(docx)
    for sec in doc.sections:
        sec.page_width, sec.page_height = _Pt(setup["width"]), _Pt(setup["height"])
        sec.left_margin, sec.right_margin = _Pt(setup["left"]), _Pt(setup["right"])
        sec.top_margin, sec.bottom_margin = _Pt(setup["top"]), _Pt(setup["bottom"])
    doc.save(docx)


def docx_to_pdf(docx: Path, out: Path) -> None:
    with tempfile.TemporaryDirectory() as td:
        tmp = Path(td) / out.with_suffix(".docx").name
        shutil.copy(docx, tmp)
        substitute_missing_fonts(tmp)
        res = soffice_convert(tmp, out.parent, "pdf")
    if res != out:
        shutil.move(res, out)
