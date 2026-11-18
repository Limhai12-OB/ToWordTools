"""ToWord — PDF to editable Word converter with Khmer support.

Run:  .venv/bin/uvicorn app.main:app --port 8000
"""
from __future__ import annotations

import json
import shutil
import threading
import time
import traceback
import uuid
from pathlib import Path
from urllib.parse import quote

from fastapi import FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel

from . import converters as conv
from . import khmer

ROOT = Path(__file__).resolve().parent.parent
DATA = ROOT / "data" / "jobs"
DATA.mkdir(parents=True, exist_ok=True)
STATIC = ROOT / "static"
MAX_UPLOAD = 60 * 1024 * 1024
JOB_TTL = 24 * 3600

ALLOWED = {".pdf", ".docx"} | conv.IMAGE_EXT | conv.OFFICE_EXT

app = FastAPI(title="ToWord")


# --------------------------------------------------------------------------- #
# Job bookkeeping (file-based so jobs survive a server restart)
# --------------------------------------------------------------------------- #
def job_dir(job_id: str) -> Path:
    if not job_id.isalnum():
        raise HTTPException(400, "bad job id")
    d = DATA / job_id
    if not d.exists():
        raise HTTPException(404, "Job not found (it may have expired).")
    return d


def write_state(d: Path, **kw) -> dict:
    f = d / "state.json"
    state = json.loads(f.read_text()) if f.exists() else {}
    state.update(kw)
    tmp = d / "state.json.tmp"
    tmp.write_text(json.dumps(state, ensure_ascii=False))
    tmp.replace(f)
    return state


def read_state(d: Path) -> dict:
    return json.loads((d / "state.json").read_text())


def cleanup_old_jobs():
    now = time.time()
    for d in DATA.iterdir():
        if d.is_dir() and now - d.stat().st_mtime > JOB_TTL:
            shutil.rmtree(d, ignore_errors=True)


cleanup_old_jobs()


# --------------------------------------------------------------------------- #
# Conversion worker
# --------------------------------------------------------------------------- #
def run_conversion(d: Path, src: Path, mode: str, lang: str, font: str):
    out = d / "converted.docx"
    try:
        ext = src.suffix.lower()
        used = mode
        quality = None
        if ext == ".pdf":
            write_state(d, status="working", message="Analysing PDF…", progress=5)
            quality = conv.analyse_pdf(src)
            write_state(d, quality=quality, pages=quality["pages"])
            if mode == "auto":
                if quality["needs_ocr"]:
                    used = "ocr"
                elif quality["khmer_chars"] > 0:
                    write_state(d, message="Checking Khmer text…", progress=10)
                    score = conv.khmer_text_layer_score(src, lang if "khm" in lang else "khm+eng")
                    quality["khmer_layer_score"] = score
                    write_state(d, quality=quality)
                    # Below 0.9 the text layer disagrees with what is printed
                    # (typical for Word + Khmer OS PDFs): re-read Khmer by OCR.
                    used = "hybrid" if score is not None and score < 0.9 else "layout"
                else:
                    used = "layout"
            if used == "hybrid":
                write_state(d, message="Converting layout and re-reading Khmer text…", progress=15,
                            mode_used="hybrid")

                def hprog(done, total):
                    write_state(d, progress=15 + int(70 * done / total),
                                message=f"Fixing Khmer text, page {done} of {total}…")

                stats = conv.pdf_to_docx_hybrid(src, out, lang if "khm" in lang else "khm+eng", progress=hprog)
                write_state(d, hybrid={k: v for k, v in stats.items() if k != "similarities"})
            elif used == "layout":
                write_state(d, message="Converting with layout…", progress=20, mode_used="layout")
                conv.pdf_to_docx_layout(src, out)
            else:
                write_state(d, message="Running Khmer/English OCR…", progress=10, mode_used="ocr")

                def prog(done, total):
                    write_state(d, progress=10 + int(75 * done / total),
                                message=f"OCR page {done} of {total}…")

                conv.ocr_to_docx(src, out, lang=lang, progress=prog)
        elif ext in conv.IMAGE_EXT:
            used = "ocr"
            write_state(d, status="working", message="Running OCR on image…", progress=20, mode_used="ocr")
            conv.ocr_to_docx(src, out, lang=lang)
        elif ext == ".docx":
            used = "docx"
            shutil.copy(src, out)
        else:
            used = "office"
            write_state(d, status="working", message="Converting document…", progress=30, mode_used="office")
            conv.office_to_docx(src, out)

        write_state(d, message="Applying Khmer fonts…", progress=88)
        if font:
            khmer.apply_khmer_font(str(out), font)
        write_state(d, message="Preparing editor…", progress=93)
        try:
            write_state(d, page=conv.page_setup(out))
        except Exception:
            pass
        html = conv.docx_to_html(out)
        (d / "converted.html").write_text(html, "utf-8")
        write_state(d, status="done", progress=100, message="Done", mode_used=used)
    except conv.ConversionError as e:
        write_state(d, status="error", error=str(e))
    except Exception as e:  # pragma: no cover
        traceback.print_exc()
        write_state(d, status="error", error=f"Unexpected error: {e}")


# --------------------------------------------------------------------------- #
# API
# --------------------------------------------------------------------------- #
@app.get("/api/config")
def config():
    return {
        "fonts": khmer.KHMER_FONTS,
        "ocr": bool(conv.TESSERACT),
        "office": bool(conv.SOFFICE),
        "accept": sorted(ALLOWED),
    }


@app.post("/api/upload")
async def upload(
    file: UploadFile = File(...),
    mode: str = Form("auto"),
    lang: str = Form("khm+eng"),
    font: str = Form("Khmer OS Battambang"),
):
    name = Path(file.filename or "document").name
    ext = Path(name).suffix.lower()
    if ext not in ALLOWED:
        raise HTTPException(400, f"Unsupported file type {ext or '(none)'}.")
    if mode not in {"auto", "layout", "hybrid", "ocr"}:
        raise HTTPException(400, "bad mode")
    if lang not in {"khm+eng", "khm", "eng"}:
        raise HTTPException(400, "bad language")
    if font and font not in khmer.KHMER_FONTS:
        raise HTTPException(400, "unknown font")

    job_id = uuid.uuid4().hex
    d = DATA / job_id
    d.mkdir(parents=True)
    src = d / ("source" + ext)
    size = 0
    with open(src, "wb") as f:
        while chunk := await file.read(1 << 20):
            size += len(chunk)
            if size > MAX_UPLOAD:
                f.close()
                shutil.rmtree(d, ignore_errors=True)
                raise HTTPException(413, "File is larger than 60 MB.")
            f.write(chunk)

    write_state(d, id=job_id, filename=name, ext=ext, status="queued", progress=0,
                message="Queued", mode=mode, lang=lang, font=font, created=time.time())
    threading.Thread(target=run_conversion, args=(d, src, mode, lang, font), daemon=True).start()
    return {"id": job_id}


@app.get("/api/jobs/{job_id}")
def job_status(job_id: str):
    d = job_dir(job_id)
    st = read_state(d)
    st["has_draft"] = (d / "draft.html").exists()
    return st


@app.get("/api/jobs/{job_id}/html")
def job_html(job_id: str, draft: bool = True):
    d = job_dir(job_id)
    f = d / "draft.html" if draft and (d / "draft.html").exists() else d / "converted.html"
    if not f.exists():
        raise HTTPException(409, "Conversion not finished.")
    return Response(f.read_text("utf-8"), media_type="text/html; charset=utf-8")


@app.get("/api/jobs/{job_id}/source")
def job_source(job_id: str):
    d = job_dir(job_id)
    st = read_state(d)
    src = d / ("source" + st["ext"])
    media = "application/pdf" if st["ext"] == ".pdf" else None
    return FileResponse(src, media_type=media)


def _download_name(job_dir_: Path, ext: str, suffix: str = "") -> str:
    stem = Path(read_state(job_dir_)["filename"]).stem
    return f"{stem}{suffix}{ext}"


def _attachment(path: Path, filename: str, media: str) -> FileResponse:
    ascii_name = filename.encode("ascii", "ignore").decode() or "document"
    headers = {"Content-Disposition":
               f"attachment; filename=\"{ascii_name}\"; filename*=UTF-8''{quote(filename)}"}
    return FileResponse(path, media_type=media, headers=headers)


DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


@app.get("/api/jobs/{job_id}/converted.docx")
def converted_docx(job_id: str):
    """The direct conversion result (best layout fidelity, before editing)."""
    d = job_dir(job_id)
    f = d / "converted.docx"
    if not f.exists():
        raise HTTPException(409, "Conversion not finished.")
    return _attachment(f, _download_name(d, ".docx"), DOCX_MIME)


class ReconvertBody(BaseModel):
    mode: str
    lang: str = "khm+eng"


@app.post("/api/jobs/{job_id}/reconvert")
def reconvert(job_id: str, body: ReconvertBody):
    """Convert the same source file again with a different mode (new job)."""
    d = job_dir(job_id)
    st = read_state(d)
    if body.mode not in {"layout", "hybrid", "ocr"} or body.lang not in {"khm+eng", "khm", "eng"}:
        raise HTTPException(400, "bad mode")
    new_id = uuid.uuid4().hex
    nd = DATA / new_id
    nd.mkdir(parents=True)
    src = nd / ("source" + st["ext"])
    shutil.copy(d / ("source" + st["ext"]), src)
    write_state(nd, id=new_id, filename=st["filename"], ext=st["ext"], status="queued", progress=0,
                message="Queued", mode=body.mode, lang=body.lang, font=st.get("font"),
                created=time.time())
    threading.Thread(target=run_conversion, args=(nd, src, body.mode, body.lang, st.get("font")),
                     daemon=True).start()
    return {"id": new_id}


class SaveBody(BaseModel):
    html: str


@app.post("/api/jobs/{job_id}/draft")
def save_draft(job_id: str, body: SaveBody):
    d = job_dir(job_id)
    (d / "draft.html").write_text(body.html, "utf-8")
    return {"saved": time.time()}


@app.delete("/api/jobs/{job_id}/draft")
def discard_draft(job_id: str):
    d = job_dir(job_id)
    (d / "draft.html").unlink(missing_ok=True)
    return {"ok": True}


class ExportBody(BaseModel):
    html: str
    font: str | None = None
    format: str = "docx"


@app.post("/api/jobs/{job_id}/export")
def export(job_id: str, body: ExportBody):
    d = job_dir(job_id)
    if body.font and body.font not in khmer.KHMER_FONTS:
        raise HTTPException(400, "unknown font")
    if body.format not in {"docx", "pdf"}:
        raise HTTPException(400, "bad format")
    (d / "draft.html").write_text(body.html, "utf-8")
    out_dir = d / "export"
    shutil.rmtree(out_dir, ignore_errors=True)
    out_dir.mkdir()
    docx = out_dir / "edited.docx"
    try:
        conv.html_to_docx(body.html, docx, body.font, read_state(d).get("page"))
        if body.format == "pdf":
            pdf = out_dir / "edited.pdf"
            conv.docx_to_pdf(docx, pdf)
            return _attachment(pdf, _download_name(d, ".pdf", "-edited"), "application/pdf")
    except conv.ConversionError as e:
        return JSONResponse({"detail": str(e)}, status_code=500)
    return _attachment(docx, _download_name(d, ".docx", "-edited"), DOCX_MIME)


app.mount("/", StaticFiles(directory=STATIC, html=True), name="static")
