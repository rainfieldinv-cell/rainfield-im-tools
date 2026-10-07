# -*- coding: utf-8 -*-
"""PDF 원본 → (표지·연락처 끝쪽·증권사 머리말/꼬리말 지운) 워드.

PDF 는 쪽이 정해져 있어 워드보다 오히려 쉽다.
  1. 1쪽(증권사 표지)은 뺀다.
  2. 마지막 쪽에 메일·전화번호가 2개 넘게 있으면 증권사 연락처 쪽으로 보고 뺀다.
  3. 쪽마다 같은 자리에 되풀이되는 위·아래 띠(증권사 머리말·로고·쪽번호)를 지운다.
  4. pdf2docx 로 워드로 바꾼다(2026-10-06 실측 비교에서 워드로 열기보다 표가 훨씬 원본에 가까웠다).
"""
import io
import re
import tempfile
import os
from collections import Counter

import fitz

EMAIL_RE = re.compile(r"[\w.\-]+@[\w\-]+\.[\w.]+")
PHONE_RE = re.compile(r"0\d{1,2}[-)\s.]\d{3,4}[-\s.]\d{4}")

TOP_BAND = 0.16       # 쪽 높이의 위 16% 를 머리말 띠로 본다(12% 로는 머리말 밑줄이 빠져 빈 표로 남았다 — 등촌역)
BOTTOM_BAND = 0.12    # 아래 12% 를 꼬리말 띠로 본다
REPEAT = 0.5          # 절반 넘는 쪽에 같은 자리로 나오면 '되풀이' 로 본다


def text_per_page(pdf_bytes):
    """쪽당 글자 수(글자를 그림으로 바꾼 PDF 를 미리 알아보려고)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = sum(len(p.get_text().strip()) for p in doc)
    return n / max(doc.page_count, 1)


def last_page_is_contact(doc):
    t = doc[-1].get_text()
    return len(set(EMAIL_RE.findall(t))) + len(set(PHONE_RE.findall(t))) >= 2


def _band_items(page):
    """위·아래 띠 안의 글줄·그림·선을 (종류, 반올림한 자리, 내용) 로."""
    h = page.rect.height
    top, bot = h * TOP_BAND, h * (1 - BOTTOM_BAND)
    out = []
    for b in page.get_text("dict")["blocks"]:
        r = fitz.Rect(b["bbox"])
        if r.y1 <= top or r.y0 >= bot:
            if b["type"] == 0:
                txt = "".join(s["text"] for l in b["lines"] for s in l["spans"]).strip()
                key = re.sub(r"\d+", "#", txt)            # 쪽번호는 숫자만 다르다
                out.append(("t", (round(r.x0 / 6), round(r.y0 / 6)), key, r))
            else:
                out.append(("i", (round(r.x0 / 6), round(r.y0 / 6), round(r.width / 6)), "", r))
    for d in page.get_drawings():
        r = d["rect"]
        if r.y1 <= top or r.y0 >= bot:
            out.append(("d", (round(r.x0 / 6), round(r.y0 / 6), round(r.width / 6)), "", r))
    return out


def strip_repeating_bands(doc):
    """여러 쪽에 같은 자리로 되풀이되는 위·아래 띠 요소를 하얗게 지운다. 지운 쪽 수를 돌려준다."""
    per_page = [_band_items(p) for p in doc]
    cnt = Counter()
    for items in per_page:
        cnt.update({(k, pos, key) for k, pos, key, _ in items})
    need = max(2, int(doc.page_count * REPEAT))
    rep = {sig for sig, c in cnt.items() if c >= need}
    touched = 0
    for page, items in zip(doc, per_page):
        rects = [r for k, pos, key, r in items if (k, pos, key) in rep]
        if not rects:
            continue
        for r in rects:
            # fill=False — 흰 네모로 덮으면 pdf2docx 가 그 네모를 빈 표로 만든다(실제: 등촌역 쪽마다 빈 표)
            page.add_redact_annot(r + (-1, -1, 1, 1), fill=False)
        try:
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE,
                                  graphics=fitz.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED)
        except TypeError:                                   # 옛 판에는 graphics 인자가 없다
            page.apply_redactions(images=fitz.PDF_REDACT_IMAGE_REMOVE)
        touched += 1
    return touched


def pdf_to_docx(pdf_bytes, drop_cover=True, drop_last=None):
    """PDF bytes → (docx bytes, 보고)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    rep = {"원본 쪽수": doc.page_count}
    if drop_last is None:
        drop_last = doc.page_count > 2 and last_page_is_contact(doc)
    if drop_last and doc.page_count > 2:
        doc.delete_page(doc.page_count - 1)
    rep["원본 끝쪽 뺌"] = bool(drop_last)
    if drop_cover and doc.page_count > 1:
        doc.delete_page(0)
    rep["원본 표지 뺌"] = bool(drop_cover)
    rep["머리말·꼬리말 지운 쪽"] = strip_repeating_bands(doc)

    from pdf2docx import Converter
    with tempfile.TemporaryDirectory() as td:
        src = os.path.join(td, "in.pdf")
        dst = os.path.join(td, "out.docx")
        doc.save(src)
        cv = Converter(src)
        try:
            cv.convert(dst)
        finally:
            cv.close()
        with open(dst, "rb") as f:
            out = f.read()
    out = _make_room_for_company_header(out)
    return out, rep


# 회사 양식(형식_1)의 쪽 여백 — 머리말·꼬리말이 들어갈 자리
_TOP_CM, _BOTTOM_CM, _HEADER_CM, _FOOTER_CM = 3.0, 2.5, 1.7, 1.2


def _make_room_for_company_header(docx_bytes):
    """pdf2docx 는 위 여백 0cm·아래 0.2cm 로 잡고 쪽 안 위치를 '문단 앞 간격' 으로 맞춘다.
    여기에 회사 머리말·꼬리말을 얹으면 본문이 밀려 **쪽마다 넘쳐 쪽수가 두 배** 가 됐다(2026-10-07 실측).
    → 여백을 회사 양식대로 넓히고, 늘어난 만큼 각 쪽 첫 문단의 앞 간격을 줄여 본문은 제자리에 둔다.
      (원본 증권사 머리말·꼬리말 띠는 이미 지웠으므로 그 자리가 비어 있다.)
    """
    import docx
    from docx.shared import Cm, Pt
    from docx.oxml.ns import qn
    from docx.enum.section import WD_SECTION
    d = docx.Document(io.BytesIO(docx_bytes))
    body = d.element.body
    blocks = [e for e in body if e.tag in (qn("w:p"), qn("w:tbl"))]
    # 구역(=원본 한 쪽)마다 블록 묶기: 구역은 '문단 안 sectPr' 로 끝난다
    groups, cur = [], []
    for el in blocks:
        cur.append(el)
        if el.tag == qn("w:p") and el.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
            groups.append(cur)
            cur = []
    groups.append(cur)

    def empty(el):
        return el.tag == qn("w:p") and not "".join(el.itertext()).strip() \
            and el.find(".//" + qn("w:drawing")) is None and el.find(qn("w:pPr") + "/" + qn("w:sectPr")) is None

    prev_size = None
    for s, grp in zip(d.sections, groups):
        grow = Cm(_TOP_CM).pt - s.top_margin.pt
        s.top_margin, s.bottom_margin = Cm(_TOP_CM), Cm(max(_BOTTOM_CM, s.bottom_margin.cm))
        s.header_distance, s.footer_distance = Cm(_HEADER_CM), Cm(_FOOTER_CM)
        # ★pdf2docx 는 쪽 위 빈 공간을 '빈 문단(고정 줄 높이)' 으로 만든다 → 그 높이부터 차례로 깎는다
        for el in grp:
            if grow <= 0:
                break
            pf = docx.text.paragraph.Paragraph(el, None).paragraph_format if el.tag == qn("w:p") else None
            if pf is None:
                break
            before = pf.space_before.pt if pf.space_before is not None else 0
            cut = min(before, grow)
            pf.space_before = Pt(before - cut)
            grow -= cut
            if not empty(el):
                break
            line = pf.line_spacing.pt if hasattr(pf.line_spacing, "pt") else 12.0
            after = pf.space_after.pt if pf.space_after is not None else 0
            if line + after <= grow:
                el.getparent().remove(el)
                grow -= line + after
            else:
                pf.line_spacing = Pt(max(1.0, line - grow))
                grow = 0
        # ★원본 한 쪽 = 한 구역(새 쪽). 글꼴이 바뀌어 몇 줄만 넘쳐도 '거의 빈 쪽' 이 생겼다
        #   → 앞 구역과 쪽 크기가 같으면 이어서 흐르게 한다(쪽 크기가 바뀌는 곳만 새 쪽).
        size = (round(s.page_width.cm, 1), round(s.page_height.cm, 1))
        if prev_size == size:
            s.start_type = WD_SECTION.CONTINUOUS
        prev_size = size
    out = io.BytesIO()
    d.save(out)
    return out.getvalue()


def suggest_title(pdf_bytes):
    """1쪽(표지)에서 가장 큰 글씨 줄들 → 표지 제목 후보(최대 2줄)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    lines = []
    for b in doc[0].get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = re.sub(r"\s+", " ", "".join(s["text"] for s in l["spans"])).strip()
            if not t or len(t) > 60 or re.search(r"memorandum|confidential|private|securities|division", t, re.I):
                continue
            if re.fullmatch(r"[\d.\-/\s]+", t):
                continue
            lines.append((max(s["size"] for s in l["spans"]), t))
    if not lines:
        return []
    top = max(s for s, _ in lines)
    out = []
    for s, t in lines:
        if s >= top * 0.8 and t not in out:
            out.append(t)
    return out[:2]


def cover_photo_candidates(pdf_bytes, limit=6):
    """PDF 안 그림 중 표지 사진 후보(색이 다양한 큰 실사 사진 먼저)."""
    from PIL import Image
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    seen, out = set(), []
    for page in doc:
        for img in page.get_images(full=True):
            xref = img[0]
            if xref in seen:
                continue
            seen.add(xref)
            try:
                pix = fitz.Pixmap(doc, xref)
                if pix.n - pix.alpha >= 4:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                if pix.width < 300 or pix.height < 150:
                    continue
                blob = pix.tobytes("png")
                im = Image.open(io.BytesIO(blob)).convert("RGB")
                colors = len(set(im.resize((80, 80)).getdata()))
                bright = sum(im.convert("L").resize((60, 60)).getdata()) / 3600
            except Exception:
                continue
            ratio = pix.width / pix.height
            out.append(((colors > 1500, bright < 210, 1.2 <= ratio <= 2.6, pix.width * pix.height), blob))
    out.sort(key=lambda t: t[0], reverse=True)
    return [b for _, b in out[:limit]]
