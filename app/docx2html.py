"""DOCX -> HTML for the browser editor.

Written specifically to keep what matters for Khmer documents: the complex-script
(w:cs) font of each run, font sizes, alignment, indentation, spacing, tables
(including merged cells and borders), images, line/page breaks and lists.
"""
from __future__ import annotations

import base64
import html
import re
from docx import Document
from docx.oxml.ns import qn

from . import khmer

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_EMBED = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed"
R_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
A_BLIP = "{http://schemas.openxmlformats.org/drawingml/2006/main}blip"
WP_EXTENT = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}extent"
V_IMAGEDATA = "{urn:schemas-microsoft-com:vml}imagedata"

HIGHLIGHT = {
    "yellow": "#ffff00", "green": "#00ff00", "cyan": "#00ffff", "magenta": "#ff00ff",
    "blue": "#0000ff", "red": "#ff0000", "darkBlue": "#000080", "darkCyan": "#008080",
    "darkGreen": "#008000", "darkMagenta": "#800080", "darkRed": "#800000",
    "darkYellow": "#808000", "darkGray": "#808080", "lightGray": "#c0c0c0", "black": "#000000",
}
ALIGN = {"center": "center", "right": "right", "end": "right", "both": "justify",
         "distribute": "justify", "left": None, "start": None}


def _val(el, tag, attr="w:val"):
    if el is None:
        return None
    c = el.find(qn(tag))
    if c is None:
        return None
    return c.get(qn(attr))


def _on(el, tag):
    """Boolean toggle property (w:b, w:i...). Returns True/False/None."""
    if el is None:
        return None
    c = el.find(qn(tag))
    if c is None:
        return None
    v = c.get(qn("w:val"))
    return v not in ("0", "false", "off")


def _twip(v) -> float | None:
    try:
        return int(v) / 20.0  # points
    except (TypeError, ValueError):
        return None


class Converter:
    def __init__(self, path):
        self.doc = Document(path)
        self.part = self.doc.part
        self.styles = {}
        st_el = self.doc.styles.element
        for s in st_el.findall(qn("w:style")):
            self.styles[s.get(qn("w:styleId"))] = s
        dd = st_el.find(qn("w:docDefaults"))
        self.default_rpr = dd.find(qn("w:rPrDefault") + "/" + qn("w:rPr")) if dd is not None else None
        self.default_ppr = dd.find(qn("w:pPrDefault") + "/" + qn("w:pPr")) if dd is not None else None
        self.default_pstyle = None
        for sid, s in self.styles.items():
            if s.get(qn("w:type")) == "paragraph" and s.get(qn("w:default")) in ("1", "true"):
                self.default_pstyle = sid
        self.numbering = self._load_numbering()

    # ------------------------------------------------------------------ styles
    def _style_chain(self, sid):
        seen, chain = set(), []
        while sid and sid in self.styles and sid not in seen:
            seen.add(sid)
            s = self.styles[sid]
            chain.append(s)
            sid = _val(s, "w:basedOn")
        return chain  # most specific first

    def _rpr_chain(self, p_style, r_style, rpr):
        chain = []
        if rpr is not None:
            chain.append(rpr)
        for s in self._style_chain(r_style):
            chain.append(s.find(qn("w:rPr")))
        for s in self._style_chain(p_style):
            chain.append(s.find(qn("w:rPr")))
        chain.append(self.default_rpr)
        return [c for c in chain if c is not None]

    def _ppr_chain(self, p_style, ppr):
        chain = [ppr] if ppr is not None else []
        for s in self._style_chain(p_style):
            chain.append(s.find(qn("w:pPr")))
        chain.append(self.default_ppr)
        return [c for c in chain if c is not None]

    @staticmethod
    def _first(chain, fn):
        for el in chain:
            v = fn(el)
            if v is not None:
                return v
        return None

    # --------------------------------------------------------------- numbering
    def _load_numbering(self):
        try:
            numbering = self.part.numbering_part.element
        except Exception:
            return {}
        abstract = {}
        for a in numbering.findall(qn("w:abstractNum")):
            lv = {}
            for l in a.findall(qn("w:lvl")):
                lv[l.get(qn("w:ilvl"))] = _val(l, "w:numFmt")
            abstract[a.get(qn("w:abstractNumId"))] = lv
        nums = {}
        for n in numbering.findall(qn("w:num")):
            nums[n.get(qn("w:numId"))] = abstract.get(_val(n, "w:abstractNumId"), {})
        return nums

    # ------------------------------------------------------------------- body
    def convert(self) -> str:
        out: list[str] = []
        self._blocks(self.doc.element.body, out)
        self._close_lists(out)
        return "\n".join(out)

    def _blocks(self, parent, out):
        self._list_stack = getattr(self, "_list_stack", [])
        for child in parent:
            tag = child.tag
            if tag == qn("w:p"):
                self._paragraph(child, out)
            elif tag == qn("w:tbl"):
                self._close_lists(out)
                self._table(child, out)
            elif tag == qn("w:sdt"):
                content = child.find(qn("w:sdtContent"))
                if content is not None:
                    self._blocks(content, out)

    def _close_lists(self, out):
        while getattr(self, "_list_stack", []):
            out.append(f"</{self._list_stack.pop()}>")

    def _paragraph(self, p, out):
        ppr = p.find(qn("w:pPr"))
        p_style = _val(ppr, "w:pStyle") or self.default_pstyle
        pchain = self._ppr_chain(p_style, ppr)
        css = {}

        jc = self._first(pchain, lambda e: _val(e, "w:jc"))
        if jc and ALIGN.get(jc):
            css["text-align"] = ALIGN[jc]
        ind = self._first(pchain, lambda e: e.find(qn("w:ind")))
        if ind is not None:
            left = _twip(ind.get(qn("w:left")) or ind.get(qn("w:start")))
            right = _twip(ind.get(qn("w:right")) or ind.get(qn("w:end")))
            first = _twip(ind.get(qn("w:firstLine")))
            hanging = _twip(ind.get(qn("w:hanging")))
            if left:
                css["margin-left"] = f"{left:g}pt"
            if right:
                css["margin-right"] = f"{right:g}pt"
            if first:
                css["text-indent"] = f"{first:g}pt"
            elif hanging:
                css["text-indent"] = f"-{hanging:g}pt"
        sp = self._first(pchain, lambda e: e.find(qn("w:spacing")))
        before = after = 0.0
        if sp is not None:
            before = _twip(sp.get(qn("w:before"))) or 0
            after = _twip(sp.get(qn("w:after"))) or 0
        css["margin-top"] = f"{min(before, 48):g}pt"
        css["margin-bottom"] = f"{min(after, 48):g}pt"

        # Heading level from style name / outline level
        tag = "p"
        sname = ""
        if p_style in self.styles:
            n = self.styles[p_style].find(qn("w:name"))
            sname = (n.get(qn("w:val")) if n is not None else "") or ""
        low = sname.lower()
        if low.startswith("heading ") and low[8:].isdigit() and 1 <= int(low[8:]) <= 6:
            tag = f"h{low[8:]}"
        elif low == "title":
            tag = "h1"

        # Lists
        num_pr = self._first(pchain, lambda e: e.find(qn("w:numPr")))
        list_tag = None
        if num_pr is not None:
            num_id = _val(num_pr, "w:numId")
            ilvl = _val(num_pr, "w:ilvl") or "0"
            if num_id and num_id != "0":
                fmt = self.numbering.get(num_id, {}).get(ilvl, "bullet")
                list_tag = "ul" if fmt in (None, "bullet", "none") else "ol"
                depth = int(ilvl) + 1
                while len(self._list_stack) > depth:
                    out.append(f"</{self._list_stack.pop()}>")
                while len(self._list_stack) < depth:
                    self._list_stack.append(list_tag)
                    out.append(f"<{list_tag}>")
                css.pop("margin-left", None)
                css.pop("text-indent", None)
        if list_tag is None:
            self._close_lists(out)

        inner, page_break = self._runs(p, p_style)
        if page_break:
            self._close_lists(out)
            out.append("<!-- pagebreak -->")
            if not inner.strip():
                return
        if not inner.strip():
            inner = "&nbsp;"
        style = "; ".join(f"{k}: {v}" for k, v in css.items())
        el = f'<{tag} style="{style}">{inner}</{tag}>'
        out.append(f"<li>{el}</li>" if list_tag else el)

    def _runs(self, p, p_style):
        parts: list[str] = []
        page_break = False
        for child in p:
            if child.tag == qn("w:r"):
                s, pb = self._run(child, p_style)
                parts.append(s)
                page_break = page_break or pb
            elif child.tag == qn("w:hyperlink"):
                rid = child.get(R_ID)
                href = ""
                if rid and rid in self.part.rels:
                    href = self.part.rels[rid].target_ref
                inner = "".join(self._run(r, p_style)[0] for r in child.findall(qn("w:r")))
                parts.append(f'<a href="{html.escape(href)}">{inner}</a>' if href else inner)
            elif child.tag in (qn("w:ins"), qn("w:smartTag"), qn("w:fldSimple")):
                for r in child.iter(qn("w:r")):
                    parts.append(self._run(r, p_style)[0])
            elif child.tag == qn("w:sdt"):
                for r in child.iter(qn("w:r")):
                    parts.append(self._run(r, p_style)[0])
        # The page break marker goes before the paragraph content when the
        # paragraph only contains the break (typical in OCR / pdf2docx output).
        return "".join(parts), page_break

    def _run(self, r, p_style):
        rpr = r.find(qn("w:rPr"))
        r_style = _val(rpr, "w:rStyle")
        chain = self._rpr_chain(p_style, r_style, rpr)
        content: list[str] = []
        text_acc = ""
        page_break = False
        for c in r:
            if c.tag == qn("w:t"):
                content.append(html.escape(c.text or ""))
                text_acc += c.text or ""
            elif c.tag == qn("w:tab"):
                content.append('<span class="tab">&emsp;&emsp;</span>')
            elif c.tag == qn("w:br"):
                if c.get(qn("w:type")) == "page":
                    page_break = True
                else:
                    content.append("<br>")
            elif c.tag == qn("w:cr"):
                content.append("<br>")
            elif c.tag in (qn("w:drawing"), qn("w:pict"), qn("w:object")):
                content.append(self._image(c))
        if not content:
            return "", page_break
        body = "".join(content)
        if not text_acc:
            return body, page_break

        css = {}
        is_kh = khmer.has_khmer(text_acc)

        def font(e):
            rf = e.find(qn("w:rFonts"))
            if rf is None:
                return None
            if is_kh:
                return rf.get(qn("w:cs")) or rf.get(qn("w:ascii"))
            return rf.get(qn("w:ascii")) or rf.get(qn("w:hAnsi"))

        f = self._first(chain, font)
        if f:
            css["font-family"] = f"'{f}'"

        def size(e):
            v = _val(e, "w:szCs") if is_kh else None
            return v or _val(e, "w:sz")

        sz = self._first(chain, size)
        if sz:
            try:
                css["font-size"] = f"{int(sz) / 2:g}pt"
            except ValueError:
                pass
        color = self._first(chain, lambda e: _val(e, "w:color"))
        color = (color or "").lstrip("#")
        if color and color.lower() not in ("auto", "000000"):
            css["color"] = f"#{color}"
        hl = self._first(chain, lambda e: _val(e, "w:highlight"))
        if hl and hl in HIGHLIGHT:
            css["background-color"] = HIGHLIGHT[hl]
        shd = self._first(chain, lambda e: _val(e, "w:shd", "w:fill"))
        if shd and shd not in ("auto", "FFFFFF") and "background-color" not in css:
            css["background-color"] = f"#{shd}"
        va = self._first(chain, lambda e: _val(e, "w:vertAlign"))

        b = self._first(chain, lambda e: _on(e, "w:bCs" if is_kh else "w:b")) if is_kh else None
        if b is None:
            b = self._first(chain, lambda e: _on(e, "w:b"))
        i = self._first(chain, lambda e: _on(e, "w:i"))
        u = self._first(chain, lambda e: _val(e, "w:u"))
        strike = self._first(chain, lambda e: _on(e, "w:strike"))

        if b:
            body = f"<strong>{body}</strong>"
        if i:
            body = f"<em>{body}</em>"
        if u and u != "none":
            body = f"<u>{body}</u>"
        if strike:
            body = f"<s>{body}</s>"
        if va == "superscript":
            body = f"<sup>{body}</sup>"
        elif va == "subscript":
            body = f"<sub>{body}</sub>"
        if css:
            style = "; ".join(f"{k}: {v}" for k, v in css.items())
            body = f'<span style="{style}">{body}</span>'
        return body, page_break

    def _image(self, el) -> str:
        blip = next(el.iter(A_BLIP), None)
        rid = blip.get(R_EMBED) if blip is not None else None
        if rid is None:
            img = next(el.iter(V_IMAGEDATA), None)
            rid = img.get(R_ID) if img is not None else None
        if not rid or rid not in self.part.rels:
            return ""
        try:
            part = self.part.rels[rid].target_part
            data = part.blob
            mime = part.content_type
        except Exception:
            return ""
        if mime in ("image/x-emf", "image/x-wmf", "image/emf", "image/wmf"):
            return ""  # browsers cannot display these
        ext = next(el.iter(WP_EXTENT), None)
        size = ""
        if ext is not None:
            try:
                w = int(ext.get("cx")) / 9525  # EMU -> px
                h = int(ext.get("cy")) / 9525
                size = f' width="{round(w)}" height="{round(h)}"'
            except (TypeError, ValueError):
                pass
        b64 = base64.b64encode(data).decode()
        return f'<img src="data:{mime};base64,{b64}"{size}>'

    # ------------------------------------------------------------------ tables
    def _floating_box_open(self, tbl, tbl_pr, tblp) -> str:
        """A page-positioned floating table (side box, e.g. the photo box of a
        form) -> floated HTML table that remembers its exact page position."""
        def num(v, default=0.0):
            try:
                return float(v)
            except (TypeError, ValueError):
                return default
        x0 = num(tblp.get(qn("w:tblpX"))) / 20
        y0 = num(tblp.get(qn("w:tblpY"))) / 20
        w = num(_val(tbl_pr, "w:tblW", "w:w")) / 20 or 100
        tr_h = tbl.find(qn("w:tr") + "/" + qn("w:trPr") + "/" + qn("w:trHeight"))
        h = num(tr_h.get(qn("w:val")) if tr_h is not None else None) / 20 or 60
        top = tbl_pr.find(qn("w:tblBorders") + "/" + qn("w:top"))
        bw = num(top.get(qn("w:sz")) if top is not None else None, 4) / 8
        first_p = next(tbl.iter(qn("w:p")), None)
        pad = 0.0
        if first_p is not None:
            sp = first_p.find(qn("w:pPr") + "/" + qn("w:spacing"))
            if sp is not None:
                pad = num(sp.get(qn("w:before"))) / 20
        try:
            page_w = self.doc.sections[0].page_width.pt
        except Exception:
            page_w = 595
        side = "right" if x0 + w / 2 > page_w / 2 else "left"
        margin = "0 0 6pt 10pt" if side == "right" else "0 10pt 6pt 0"
        style = (f"float: {side}; width: {w:g}pt; height: {h:g}pt; margin: {margin}; "
                 f"border-collapse: collapse")
        return (f'<table class="tw-box" data-x0="{x0:g}" data-y0="{y0:g}" data-x1="{x0 + w:g}" '
                f'data-y1="{y0 + h:g}" data-pad="{pad:g}" data-bw="{bw:g}" style="{style}"><tbody>')

    def _table(self, tbl, out):
        tbl_pr = tbl.find(qn("w:tblPr"))
        borders_default = self._borders(tbl_pr.find(qn("w:tblBorders")) if tbl_pr is not None else None)
        t_style = _val(tbl_pr, "w:tblStyle")
        if not borders_default and t_style in self.styles:
            st_pr = self.styles[t_style].find(qn("w:tblPr"))
            if st_pr is not None:
                borders_default = self._borders(st_pr.find(qn("w:tblBorders")))
        jc = _val(tbl_pr, "w:jc")
        tstyle = "border-collapse: collapse; width: 100%"
        if jc == "center":
            tstyle += "; margin-left: auto; margin-right: auto"
        tblp = tbl_pr.find(qn("w:tblpPr")) if tbl_pr is not None else None
        if tblp is not None and tblp.get(qn("w:horzAnchor")) == "page":
            out.append(self._floating_box_open(tbl, tbl_pr, tblp))
        else:
            out.append(f'<table style="{tstyle}"><tbody>')

        rows = tbl.findall(qn("w:tr"))
        # Compute rowspans for vMerge.
        grid: list[list[dict]] = []
        for tr in rows:
            cells = []
            col = 0
            for tc in tr.findall(qn("w:tc")):
                tcpr = tc.find(qn("w:tcPr"))
                span = int(_val(tcpr, "w:gridSpan") or 1)
                vm_el = tcpr.find(qn("w:vMerge")) if tcpr is not None else None
                vmerge = None
                if vm_el is not None:
                    vmerge = vm_el.get(qn("w:val")) or "continue"
                cells.append({"tc": tc, "col": col, "span": span, "vmerge": vmerge, "rowspan": 1})
                col += span
            grid.append(cells)
        for ri, cells in enumerate(grid):
            for c in cells:
                if c["vmerge"] == "restart":
                    rs = 1
                    for nxt in grid[ri + 1:]:
                        m = next((x for x in nxt if x["col"] == c["col"]), None)
                        if m and m["vmerge"] == "continue":
                            rs += 1
                        else:
                            break
                    c["rowspan"] = rs

        for cells in grid:
            out.append("<tr>")
            for c in cells:
                if c["vmerge"] == "continue":
                    continue
                tc = c["tc"]
                tcpr = tc.find(qn("w:tcPr"))
                css = {}
                cb = self._borders(tcpr.find(qn("w:tcBorders")) if tcpr is not None else None)
                b = {**borders_default, **cb}
                for side in ("top", "bottom", "left", "right"):
                    if side in b:
                        css[f"border-{side}"] = b[side]
                    elif side.replace("left", "start").replace("right", "end") in b:
                        css[f"border-{side}"] = b[side.replace("left", "start").replace("right", "end")]
                if "insideH" in b:
                    css.setdefault("border-top", b["insideH"])
                    css.setdefault("border-bottom", b["insideH"])
                if "insideV" in b:
                    css.setdefault("border-left", b["insideV"])
                    css.setdefault("border-right", b["insideV"])
                fill = _val(tcpr, "w:shd", "w:fill")
                if fill and fill not in ("auto", "FFFFFF"):
                    css["background-color"] = f"#{fill}"
                va = _val(tcpr, "w:vAlign")
                if va in ("center", "bottom"):
                    css["vertical-align"] = "middle" if va == "center" else "bottom"
                w = tcpr.find(qn("w:tcW")) if tcpr is not None else None
                if w is not None and w.get(qn("w:type")) == "dxa":
                    pts = _twip(w.get(qn("w:w")))
                    if pts:
                        css["width"] = f"{pts:g}pt"
                attrs = ""
                if c["span"] > 1:
                    attrs += f' colspan="{c["span"]}"'
                if c["rowspan"] > 1:
                    attrs += f' rowspan="{c["rowspan"]}"'
                style = "; ".join(f"{k}: {v}" for k, v in css.items())
                inner: list[str] = []
                saved = self._list_stack
                self._list_stack = []
                self._blocks(tc, inner)
                self._close_lists(inner)
                self._list_stack = saved
                out.append(f'<td{attrs} style="{style}">{"".join(inner)}</td>')
            out.append("</tr>")
        out.append("</tbody></table>")

    @staticmethod
    def _borders(el) -> dict:
        res = {}
        if el is None:
            return res
        for b in el:
            side = b.tag.split("}")[1]
            val = b.get(qn("w:val"))
            if val in (None, "nil", "none"):
                res[side] = "none"
                continue
            try:
                sz = float(b.get(qn("w:sz")) or 4) / 8.0
            except ValueError:
                sz = 0.5
            color = b.get(qn("w:color"))
            color = "#000000" if not color or color == "auto" else "#" + color.lstrip("#")
            style = "double" if val == "double" else "dashed" if "dash" in val else "dotted" if "dot" in val else "solid"
            res[side] = f"{max(sz, 0.5):g}pt {style} {color}"
        return res


_ADJ_SPAN_RE = re.compile(r'<span style="([^"]*)">((?:(?!<span|</span>).)*)</span><span style="\1">', re.S)
_ADJ_TAG_RE = re.compile(r"</(strong|em|u|s)><\1>")


def _merge_runs(h: str) -> str:
    prev = None
    while prev != h:
        prev = h
        h = _ADJ_TAG_RE.sub("", h)
        h = _ADJ_SPAN_RE.sub(r'<span style="\1">\2', h)
    return h


def docx_to_html(path) -> str:
    return _merge_runs(Converter(path).convert())
