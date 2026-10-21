# -*- coding: utf-8 -*-
"""PDF 원본 → 회사 양식 규격에 맞춘 워드 본문.

2026-10-08 사용자 확정 규칙
  · 결과는 **회사 양식 규격(A4·여백) 고정**, 머리말·꼬리말 위치·크기도 양식 고정.
  · **원본 1쪽 = 결과 1쪽.** 원본 쪽 내용이 양식 안 칸보다 크면 **비율 그대로 줄여** 넣는다(작으면 그대로).
  · 1쪽(증권사 표지)은 빼고, 마지막 쪽이 증권사 연락처면 뺀다.
  · 쪽마다 같은 자리에 되풀이되는 위·아래 띠(증권사 머리말·로고·쪽번호)는 지운다.

순서
  1. 뺄 쪽·띠 지우기(fitz)  2. 쪽마다 남은 내용의 범위를 잰다
  3. pdf2docx 로 워드(원본 한 쪽 = 한 구역)  4. 구역마다 양식 규격으로 바꾸고 내용 전체를 같은 비율로 줄인다
  5. pdf2docx 가 실패해 건너뛴 쪽은 **원본 쪽을 고화질 그림으로** 넣는다(내용이 빠지면 안 된다)
"""
import io
import logging
import os
import re
import tempfile
from collections import Counter

import fitz

EMAIL_RE = re.compile(r"[\w.\-]+@[\w\-]+\.[\w.]+")
PHONE_RE = re.compile(r"0\d{1,2}[-)\s.]\d{3,4}[-\s.]\d{4}")

TOP_BAND = 0.16       # 쪽 높이의 위 16% 를 머리말 띠로 본다(12% 로는 머리말 밑줄이 빠져 빈 표로 남았다 — 등촌역)
BOTTOM_BAND = 0.12    # 아래 12% 를 꼬리말 띠로 본다
REPEAT = 0.5          # 절반 넘는 쪽에 같은 자리로 나오면 '되풀이' 로 본다

# 회사 양식(형식_1) 규격 — cm
PAGE_W, PAGE_H = 21.0, 29.7
# ★양식(최종) 파일의 값 그대로(twips → cm): 위 1701·아래 1440·왼 1440·오른 1440·머리말 975·꼬리말 697.
#   예전엔 2.5/1.7 로 어림해 표지·개요 쪽(양식 구역)과 본문 쪽의 머리말 자리·폭이 조금씩 달랐다(2026-10-14 사용자 지적).
_TW = 2.54 / 1440
MIN_SZ = 18          # 본문 글자 크기(반 pt 단위) = 9pt — 2026-10-14 사용자 확정 '본문 내용은 9'
# ★규격 = 내곡동 원본 규격(2026-10-14 사용자 확정): 원본 실측 본문 폭 516pt(42.5~558.5), 본문 아래 끝 784pt,
#   머리말 글 37pt·머리말 선 80pt, 쪽번호 794pt. 좌우 여백은 같게(사용자 '가로 여백 같게').
LEFT_CM = RIGHT_CM = (21.0 - 516 / 72 * 2.54) / 2          # ≈1.40cm
TOP_CM, BOTTOM_CM = 2.85, 2.05
HEADER_CM, FOOTER_CM = 1.30, 1.10
BOX_W_PT = (PAGE_W - LEFT_CM - RIGHT_CM) / 2.54 * 72        # 본문 칸 가로(pt)
BOX_H_PT = (PAGE_H - TOP_CM - BOTTOM_CM) / 2.54 * 72        # 본문 칸 세로(pt)
HEIGHT_SAFETY = 0.89  # 글꼴이 Pretendard 로 바뀌면 줄이 조금 높아진다 → 세로는 8% 여유(5% 로는 등촌역 쪽 몇 개가 넘쳤다)

PAGE_NO_RE = re.compile(r"[-–\s]*(?:page\s*)?\d{1,3}(?:\s*(?:/|of)\s*\d{1,3})?[-–\s]*", re.I)
HIGHLIGHT_RE = re.compile(r"executive\s*summary|investment\s*highlight|투자\s*하이라이트|투자\s*포인트|"
                          r"핵심\s*요약|하이라이트|highlights?|overview", re.I)   # 신촌: 'OVERVIEW'


# ──────────────────────────────────────────────
# 원본 살펴보기
# ──────────────────────────────────────────────
def text_per_page(pdf_bytes):
    """쪽당 글자 수(글자를 그림으로 바꾼 PDF 를 미리 알아보려고)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = sum(len(p.get_text().strip()) for p in doc)
    return n / max(doc.page_count, 1)


_END_RE = re.compile(r"^[\s\-–—]*(끝|end|the\s*end|감사합니다|thank\s*you)[\s\-–—.!]*$", re.I)


def last_page_is_contact(doc):
    """원본 마지막 쪽이 연락처 쪽이거나 '-끝-'·'감사합니다' 만 있는 끝 쪽이면 True(그 쪽은 뺀다 — 결과 끝은 회사 연락처 쪽).
    ★고산3지구 30쪽 '-끝-' 이 남았다(2026-10-15 사용자 17번)."""
    pg = doc[-1]
    t = pg.get_text()
    if len(set(EMAIL_RE.findall(t))) + len(set(PHONE_RE.findall(t))) >= 2:
        return True
    lines = [l.strip() for l in t.splitlines() if l.strip()]
    lines = [l for l in lines if not re.fullmatch(r"[\d\s/\-–—.|pPage]+", l)]   # 쪽번호 줄
    if pg.get_images() and any(fitz.Rect(i["bbox"]).get_area() > pg.rect.get_area() * 0.2 for i in pg.get_image_info()):
        return False                                               # 큰 그림이 있는 쪽은 본문
    return bool(lines) and len(lines) <= 3 and any(_END_RE.match(l) for l in lines)


SECTION_HEAD_RE = re.compile(r"^\s*(\d{1,2}\s*\.|[Ⅰ-Ⅻ]\s*\.?|[IVX]{1,4}\s*\.)\s*\S")


def highlight_band(page):
    """쪽 위쪽에 하이라이트 제목이 있고 그 아래에 다른 대단원 제목(예: '1. 주요금융조건')이 이어지면
    → 하이라이트 부분의 (위 y, 아래 y). 하이라이트만 있는 쪽이거나 하이라이트가 없으면 None.
    ★등촌역 3쪽: 위 1/4 이 Executive Summary, 나머지가 '1. 주요금융조건' 표 — 쪽째 빼면 금융조건이 사라진다(2026-10-13)."""
    h = page.rect.height
    lines = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if t:
                lines.append((l["bbox"][1], max(s["size"] for s in l["spans"]), t))
    lines.sort()
    title = next(((y, sz) for y, sz, t in lines if y < h * 0.35 and HIGHLIGHT_RE.search(t) and len(t) < 40), None)
    if title is None:
        return None
    ty, tsz = title
    nxt = next((y for y, sz, t in lines if y > ty + 5 and y < h * 0.85 and sz >= tsz - 1 and len(t) < 40
                and SECTION_HEAD_RE.match(t) and not HIGHLIGHT_RE.search(t)), None)
    return (ty, nxt) if nxt is not None else None


def find_highlight_pages(pdf_bytes):
    """하이라이트(Executive Summary 등) 쪽 번호(1부터) 목록. 표지 다음 6쪽 안에서, 쪽 위쪽 글로 찾는다.
    하이라이트와 본문이 한 쪽에 섞인 쪽은 넣지 않는다(그 쪽은 convert 가 하이라이트 부분만 지운다)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    hits = []
    toc = set(find_toc_pages(doc))                           # 목차 쪽은 하이라이트 아님(목차에 'Highlight' 글이 있다)
    for i in range(1, min(doc.page_count, 7)):
        if i in toc:
            continue
        top = doc[i].get_text("text", clip=fitz.Rect(0, 0, doc[i].rect.width, doc[i].rect.height * 0.35))
        if HIGHLIGHT_RE.search(top):
            hits.append(i + 1)
        elif hits:
            break
    return hits


BOND_NEED_PT = 270    # 사모사채 개요(제목 + 9줄 표)에 드는 높이 어림(약 9.5cm)
FIN_HEAD_PT = 75      # 금융 개요 대단원 제목 + '1. 금융 구조도' 소제목 + 사이 띄움
FIT_FLOOR = 0.30      # 쪽 줄이기 최저 비율 — 9pt 고정 글만으로도 안 들어가는 빽빽한 표 쪽(내곡 분양가 테이블)만 여기까지 내려간다


_TOC_RE = re.compile(r"(목\s*차|차\s*례|contents|table\s*of\s*contents|index)", re.I)


def find_toc_pages(doc):
    """앞쪽(표지 포함 6쪽 안)에서 목차·차례 쪽(0부터). 쪽 위쪽 절반에 '목차/차례/Contents' 만 적힌 줄이 있고,
    쪽 글이 짧은 줄 위주(목차 항목)인 쪽."""
    out = []
    for i in range(min(6, doc.page_count)):
        pg = doc[i]
        lines = []
        for b in pg.get_text("dict")["blocks"]:
            for l in b.get("lines", []):
                t = re.sub(r"\s+", " ", "".join(s["text"] for s in l["spans"])).strip()
                if t:
                    lines.append((l["bbox"][1], t))
        hit = any(y < pg.rect.height * 0.5 and _TOC_RE.fullmatch(t.strip(" .:·-|")) for y, t in lines)
        if not hit:
            continue
        # 목차 항목 = 점선(....)으로 쪽번호까지 이은 줄. 그런 줄이 있거나 쪽 글 줄이 적으면 목차 쪽
        leaders = sum(1 for _, t in lines if re.search(r"(\.\s?){5,}|…{2,}|·{5,}", t))
        if leaders >= 2 or len(lines) <= 30:
            out.append(i)
    return out


def bond_place(rep, highlight_pages, diagram_png=None):
    """사모사채 개요 자리 — 하이라이트 마지막 쪽 뒤(2026-10-08 사용자 확정).

    그 쪽에 개요가 들어갈 여백이 남으면 (k, True) = 이어서, 아니면 (k, False) = 새 쪽.
    ★개요 + 금융 개요 > 1. 금융 구조도(그림)는 한 쪽에 같이 있어야 한다(2026-10-21) → 구조도가 있으면 그 높이(본문 폭 기준)까지 넣어 잰다.
    하이라이트를 못 찾았으면 None = 표지 바로 다음 쪽.
    """
    if not highlight_pages:
        return None
    need = BOND_NEED_PT
    if diagram_png:
        try:
            pix = fitz.Pixmap(diagram_png)
            need += FIN_HEAD_PT + BOX_W_PT * pix.height / max(pix.width, 1) + 12
        except Exception:
            need += FIN_HEAD_PT + 300
    last = max(highlight_pages)
    for k, p in enumerate(rep.get("쪽", [])):
        if p["원본쪽"] == last:
            return (k, BOX_H_PT - p["높이"] >= need)
    return None


def page_png(pdf_bytes, pno, zoom=0.6):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    return doc[pno].get_pixmap(matrix=fitz.Matrix(zoom, zoom)).tobytes("png")


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


def all_photos(pdf_bytes):
    """원본에 든 그림 **전부**(로고처럼 여러 쪽에 되풀이되는 것·아주 작은 것만 뺀다).

    돌려주는 것: [{'blob','w','h','page','score'}] — 표지에 쓸 만한 순(색이 다양한 큰 실사 사진 먼저).
    blob 은 원본에 든 해상도 그대로(가능하면 원본 파일 그대로) 꺼낸다.
    """
    from PIL import Image
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    pages_of = Counter()
    for page in doc:
        for x in {img[0] for img in page.get_images(full=True)}:
            pages_of[x] += 1
    seen, out = set(), []
    for pno, page in enumerate(doc):
        for img in page.get_images(full=True):
            xref, smask = img[0], img[1]
            if xref in seen or pages_of[xref] >= 3:
                continue
            seen.add(xref)
            try:
                if smask:
                    pix = fitz.Pixmap(doc, xref)
                    if pix.n - pix.alpha >= 4:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    pix = fitz.Pixmap(pix, fitz.Pixmap(doc, smask))
                    blob = pix.tobytes("png")
                else:
                    ext = doc.extract_image(xref)
                    blob = ext["image"]
                    if ext.get("ext") not in ("jpeg", "jpg", "png"):
                        pix = fitz.Pixmap(doc, xref)
                        if pix.n - pix.alpha >= 4:
                            pix = fitz.Pixmap(fitz.csRGB, pix)
                        blob = pix.tobytes("png")
                im = Image.open(io.BytesIO(blob))
                if im.mode == "CMYK":
                    im = im.convert("RGB")
                    b = io.BytesIO()
                    im.save(b, "JPEG", quality=95)
                    blob = b.getvalue()
                w, h = im.size
                if w < 120 or h < 80:
                    continue
                rgb = im.convert("RGB")
                colors = len(set(rgb.resize((80, 80)).getdata()))
                bright = sum(rgb.convert("L").resize((60, 60)).getdata()) / 3600
            except Exception:
                continue
            ratio = w / h
            score = (colors > 1500, bright < 210, 1.2 <= ratio <= 2.6, w * h)
            out.append({"blob": blob, "w": w, "h": h, "page": pno + 1, "score": score})
    out.sort(key=lambda d: d["score"], reverse=True)
    return out


# ──────────────────────────────────────────────
# 띠 지우기 · 내용 범위 재기
# ──────────────────────────────────────────────
EDGE_TOP, EDGE_BOTTOM = 0.055, 0.05   # 쪽 맨 가장자리 여백 — 본문이 안 들어오는 자리


def _is_page_no(txt, pno):
    """쪽번호인가. 숫자만 덜렁 있으면 그 쪽의 실제 번호와 거의 같을 때만(표 숫자 '266' 을 지운 일이 있다)."""
    if not PAGE_NO_RE.fullmatch(txt):
        return False
    if not re.search(r"/|of|page|페이지|^[-–]|[-–]$", txt.strip(), re.I):
        # 숫자만 덜렁 있는 건 여기서 정하지 않는다 → '쪽마다 1씩 커지는지' 로 판단(strip_repeating_bands 의 num).
        # ★예전엔 '쪽번호 ±3' 이면 바로 지웠다 → 앞쪽 제목 네모 번호가 지워졌다(남포동 2쪽 '1', 7쪽 '4').
        return False
    nums = re.findall(r"\d+", txt)
    return len(nums) >= 1 and pno - 3 <= int(nums[0]) <= pno + 3


def _band_items(page):
    """위·아래 띠 안의 글줄·그림·선을 (종류, 반올림한 자리, 내용, 범위) 로."""
    h = page.rect.height
    pno = page.number + 1
    top, bot = h * TOP_BAND, h * (1 - BOTTOM_BAND)
    out = []
    for b in page.get_text("dict")["blocks"]:
        # ★글은 '줄' 단위로 본다 — 덩어리(block) 단위면 머리말이 아래 본문과 한 덩어리로 묶여 띠 밖으로
        #   잡혀 하나도 안 지워졌다(고산3지구 머리말 249줄 그대로)
        if b["type"] == 0:
            units = [(fitz.Rect(l["bbox"]), "".join(s["text"] for s in l["spans"]).strip()) for l in b["lines"]]
        else:
            # ★그림은 '내용' 까지 같아야 되풀이로 본다 — 성수동은 제목 숫자('2.', '8.')가 그림 조각인데,
            #   자리·크기만 보니 숫자가 달라도 같은 그림으로 쳐서 제목 숫자가 지워졌다
            import hashlib
            digest = hashlib.md5(b.get("image", b"")[:4096]).hexdigest()[:10]
            units = [(fitz.Rect(b["bbox"]), "\x00" + digest)]
        for r, txt in units:
            if not (r.y1 <= top or r.y0 >= bot):
                continue
            if txt is not None and txt.startswith("\x00"):
                out.append(("i", (round(r.x0 / 6), round(r.y0 / 6), round(r.width / 6)), txt, r))
                continue
            if txt is not None:
                if not txt:
                    continue
                # 가로 쪽은 높이가 낮아 5% 가 33pt 밖에 안 된다(성수동 'Ⅰ. Project 개요' 가 남았다) → 최소 40pt
                # 아래쪽 가장자리는 안 쓴다 — 쪽 바닥까지 내려온 표 숫자('33,389')를 지운 일이 있다(등촌역).
                #   아래 꼬리말은 '되풀이' 와 쪽번호 규칙으로 잡힌다.
                edge = r.y1 <= max(h * EDGE_TOP, 40)
                if edge or _is_page_no(txt.replace("페이지", "page"), pno):
                    # 쪽번호('13 / 14', '- 3 -', 'Page 3')는 자릿수에 따라 자리가 조금씩 달라 '되풀이' 로
                    # 안 잡혔다(등촌역 13·14쪽에 남음) → 띠 안에 있으면 무조건 지운다.
                    # 쪽 맨 가장자리 여백의 글(고산 2쪽 오른쪽 위 머리말)도 한 번만 나와도 지운다.
                    out.append(("pn", (), "", r))
                    continue
                key = re.sub(r"\d+", "#", txt)            # 쪽번호는 숫자만 다르다
                if re.fullmatch(r"\s*\d{1,3}\s*", txt) and r.y0 >= bot and int(txt) == pno:
                    out.append(("pn", (), "", r))         # 아래 띠에 그 쪽 실제 번호와 똑같은 숫자(안산 30쪽 '30')
                    continue
                if re.fullmatch(r"\s*\d{1,3}\s*", txt):
                    # 숫자만 있는 줄 — 쪽번호를 본문부터 1로 다시 세는 문서가 있어(고산3지구 5쪽의 '1')
                    # '높이' 가 여러 쪽에서 같으면 쪽번호로 본다(아래 strip_repeating_bands 에서 센다)
                    # 자리는 '위 띠/아래 띠' 만 본다(안산 30쪽만 쪽번호 높이가 조금 달라 남았다)
                    out.append(("num", ("위" if r.y1 <= top else "아래",), txt.strip(), r))
                    continue
                if len(key.strip()) < 4:
                    continue                               # 'Ⅲ.' '개요' 같은 짧은 제목 조각은 머리말로 안 본다(성수동)
                out.append(("t", (round(r.x0 / 6), round(r.y0 / 6)), key, r))
            else:
                out.append(("i", (round(r.x0 / 6), round(r.y0 / 6), round(r.width / 6)), "", r))
    for d in page.get_drawings():
        r = d["rect"]
        if r.y1 <= top or r.y0 >= bot:
            out.append(("d", (round(r.x0 / 6), round(r.y0 / 6), round(r.width / 6)), "", r))
    return out


def _safe_text_targets(doc, pno, rects):
    """글 줄 지우기를 사본 쪽에서 먼저 해 보고, 다른 본문 글까지 사라지게 하는 줄은 빼고 돌려준다.

    ★PDF 지우기는 글자 '모양 상자' 가 겹치면 지운다. 글꼴 정보가 엉터리로 큰 PDF 에서는 꼬리말 한 줄을
      지우는데 그 위 표 숫자까지 사라졌다(등촌역 8쪽: '오이코스자산운용 대체투자본부' 를 지우니 '33,389' 도).
    반환: (지워도 되는 범위들, PDF 에선 못 지운 줄들의 글 — 워드로 바꾼 뒤 지운다)
    """
    page = doc[pno]
    targets = {tuple(r) for r in rects}
    keep_txt = Counter(t for r, t in _lines(page) if tuple(r) not in targets)
    ok, skipped = [], []
    for r in rects:
        tmp = fitz.open()
        tmp.insert_pdf(doc, from_page=pno, to_page=pno)
        tp = tmp[0]
        tp.add_redact_annot(r, fill=False)
        _apply(tp, images=False, graphics=None, text=True)
        after = Counter(t for rr, t in _lines(tp))
        lost = keep_txt - after
        if lost:
            skipped.append(next((t for rr, t in _lines(page) if tuple(rr) == tuple(r)), ""))
        else:
            ok.append(r)
    return ok, [t for t in skipped if t]


def strip_repeating_bands(doc, skipped=None):
    """여러 쪽에 같은 자리로 되풀이되는 위·아래 띠 요소를 지운다. 지운 쪽 수를 돌려준다.

    skipped(list) 를 주면 PDF 에서 안전하게 못 지운 머리말·꼬리말 글을 거기 모은다(워드로 바꾼 뒤 지운다).
    """
    per_page = [_band_items(p) for p in doc]
    cnt = Counter()
    for items in per_page:
        cnt.update({(k, pos, key) for k, pos, key, _ in items})
    need = max(2, int(doc.page_count * REPEAT))
    rep = {sig for sig, c in cnt.items() if c >= need}
    # ★같은 머리말 글이라도 쪽마다 자리가 조금 다를 수 있다(고산3지구 2쪽 DISCLAIMER 의 오른쪽 위 머리말).
    #   → 글은 '자리' 를 빼고 '내용' 만으로도 되풀이를 센다.
    txt_cnt = Counter()
    for items in per_page:
        txt_cnt.update({key for k, pos, key, _ in items if k == "t" and len(key) >= 8})
    rep_txt = {key for key, c in txt_cnt.items() if c >= need}
    # 숫자만 있는 줄은 '쪽마다 1씩 커지는'(숫자 − 쪽번호가 같은) 것만 쪽번호로 본다.
    # ★예전엔 같은 높이에 숫자만 여러 쪽 나오면 지웠다 → 제목 네모 안 번호(고산 8쪽 '2 사업 개요' 의 2, 21쪽 5,
    #   25쪽 8, 26쪽 9)가 지워졌다(2026-10-12 사용자 지적). 표·차트 숫자도 이 띠에 많이 걸린다.
    num_cnt = Counter()
    for page, items in zip(doc, per_page):
        num_cnt.update({(pos, int(key) - page.number) for k, pos, key, _ in items if k == "num"})
    num_ok = {sig for sig, c in num_cnt.items() if c >= max(3, need // 2)}
    touched = 0
    for page, items in zip(doc, per_page):
        hit = [(k, r) for k, pos, key, r in items
               if k == "pn" or (k != "num" and (k, pos, key) in rep) or (k == "t" and key in rep_txt)
               or (k == "num" and (pos, int(key) - page.number) in num_ok)]
        if not hit:
            continue
        # ★두 번에 나눠 지운다. 한 번에 지우면 선·그림 자리에 걸친 **글자까지** 지워진다
        #   (성수동: 쪽마다 같은 자리의 제목 밑줄을 지우다 그 위 제목 '2. 개발계획' 이 같이 사라졌다).
        #   ① 글 줄 → 글만  ② 선·그림 → 그 도형만(글은 그대로)
        # fill=False — 흰 네모로 덮으면 pdf2docx 가 그 네모를 빈 표로 만든다(실제: 등촌역 쪽마다 빈 표)
        texts = [r for k, r in hit if k in ("t", "pn", "num")]
        # ★선·그림은 그 쪽 본문 글보다 위(위 띠) / 아래(아래 띠)에 있을 때만 머리말·꼬리말로 본다.
        #   제목 아래 밑줄처럼 본문 사이에 끼어 있으면 쪽마다 같은 자리여도 본문 장식이다(성수동 제목 밑줄).
        hit_txt = set(map(tuple, texts))
        body = [r for r, t in _lines(page) if tuple(r) not in hit_txt]
        h = page.rect.height
        top_body = min((r.y0 for r in body if r.y0 < h * 0.5), default=h)
        bot_body = max((r.y1 for r in body if r.y1 > h * 0.5), default=0)
        shapes = [r for k, r in hit if k not in ("t", "pn", "num")
                  and ((r.y1 <= h * 0.5 and r.y1 <= top_body + 2) or (r.y0 > h * 0.5 and r.y0 >= bot_body - 2))]
        if texts:
            texts, miss = _safe_text_targets(doc, page.number, texts)
            if skipped is not None:
                skipped.extend(miss)
            for r in texts:
                page.add_redact_annot(r, fill=False)
            if texts:
                _apply(page, images=False, graphics=None, text=True)
        # ★선은 '걸치기만 해도' 지운다 — 원본 머리말 밑줄이 쪽 바깥(x=-1)까지 그어져 있어 '완전히 든 것만' 으로는
        #   안 지워졌다(등촌역). 본문보다 위·아래에 있는 것만 골랐으므로 본문 선은 건드리지 않는다.
        #   그림은 하나씩 따로 지운다 — 딱 붙은 두 조각을 한 번에 지우니 하나가 남았다(등촌역 머리말 장식).
        for r in shapes:
            page.add_redact_annot(r + (-1, -1, 1, 1), fill=False)
            _apply(page, images=True, graphics="touched", text=False)
        touched += 1
    return touched


ROLE_RE = re.compile(r"담당|팀장|부장|차장|과장|대리|주임|사원|이사|상무|전무|본부|부서|팀\b|연락처|문의|"
                     r"tel|e-?mail|mobile|fax|contact|휴대|전화|이메일", re.I)


def _lines(page):
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            t = "".join(s["text"] for s in l["spans"]).strip()
            if t:
                out.append((fitz.Rect(l["bbox"]), t))
    return out


def strip_contacts(page):
    """원본 증권사 담당자 연락처 묶음(이름·직급·전화·메일)을 지운다 — 어느 쪽에 있든(2026-10-08 사용자 확정).

    메일·전화번호가 2개 이상인 줄들을 씨앗으로, 위아래로 바짝 붙은(30pt 안) 이름·직급·'담당 부서' 줄까지
    넓힌 범위의 글·선·배경·그림을 지운다. 떨어져 있는 본문 문단은 건드리지 않는다.
    (실제: 고산3지구 2쪽 DISCLAIMER 아래 '금융조달 담당 부서 – 키움증권 …' 4명 연락처)
    반환: 지웠으면 True
    """
    lines = _lines(page)
    hits = [(r, t) for r, t in lines if EMAIL_RE.search(t) or PHONE_RE.search(t)]
    found = set()
    for _, t in hits:
        found |= set(EMAIL_RE.findall(t)) | set(PHONE_RE.findall(t))
    if len(found) < 2:
        return False
    box = fitz.Rect()
    for r, _ in hits:
        box |= r
    grew = True
    while grew:
        grew = False
        for r, t in lines:
            if r in box or box.contains(r):
                continue
            near = (0 <= box.y0 - r.y1 <= 30) or (0 <= r.y0 - box.y1 <= 30) or r.intersects(box)
            if near and (ROLE_RE.search(t) or len(t) <= 14):
                box |= r
                grew = True
    box = box + (-8, -6, 8, 6)
    # 연락처 묶음에 걸친 얇은 띠·선(제목 줄 회색 배경, 칸 선)도 함께 — 큰 테두리는 건드리지 않는다
    for d in page.get_drawings():
        r = d["rect"]
        if r.intersects(box) and r.height < 40 and r.width < page.rect.width * 0.95:
            box |= r
    page.add_redact_annot(box, fill=False)
    _apply(page, images=True, graphics="covered", text=True)
    return True


_CHART_DEBUG = False


def find_charts(page):
    """선·막대로 그린 차트 자리(fitz.Rect 목록). ★pdf2docx 는 차트의 막대·꺾은선을 못 옮겨 축 숫자만 흩어졌다
    (내곡동 '(3) 월별 분양 추이', 2026-10-14 사용자 21번 '못 옮기는 차트는 원본 그 부분을 그림으로').
    그림 묶음 중 비스듬한 선·곡선이 있거나, 기울어진 글(축 글)이 있거나, 바닥이 같은 막대가 여럿인 것."""
    W, H = page.rect.width, page.rect.height
    items = []
    for dr in page.get_drawings():
        r = dr["rect"]
        if r.get_area() > W * H * 0.4 or r.y1 < H * TOP_BAND or r.y0 > H * (1 - BOTTOM_BAND):
            continue
        diag = curve = 0
        for it in dr["items"]:
            if it[0] == "l":
                a, b = it[1], it[2]
                if abs(a.x - b.x) > 1 and abs(a.y - b.y) > 1:
                    diag += 1
            elif it[0] == "c":
                curve += 1
        items.append((fitz.Rect(r), diag, curve, dr.get("fill") is not None))
    if len(items) < 15:
        return []
    n = len(items)
    par = list(range(n))

    def f(a):
        while par[a] != a:
            par[a] = par[par[a]]
            a = par[a]
        return a
    infl = [it[0] + (-6, -6, 6, 6) for it in items]
    order = sorted(range(n), key=lambda i: infl[i].x0)
    for ai, a in enumerate(order):
        for b in order[ai + 1:]:
            if infl[b].x0 > infl[a].x1:
                break
            if infl[a].intersects(infl[b]):
                par[f(a)] = f(b)
    groups = {}
    for i in range(n):
        groups.setdefault(f(i), []).append(items[i])
    rot = [fitz.Rect(l["bbox"]) for b in page.get_text("dict")["blocks"] for l in b.get("lines", [])
           if l.get("dir") and abs(l["dir"][1]) > 0.2]
    out = []
    for g in groups.values():
        if len(g) < 15:
            continue
        box = fitz.Rect(g[0][0])
        for it in g[1:]:
            box |= it[0]
        if box.width < 80 or box.height < 50:
            continue
        diag = sum(it[1] for it in g)
        curve = sum(it[2] for it in g)
        nrot = sum(1 for r in rot if (box + (-20, -20, 20, 20)).intersects(r))
        words = page.get_text("words", clip=box)
        bars = [it[0] for it in g if it[3] and it[0].height > 3 and it[0].width < box.width * 0.2
                and not any(fitz.Rect(w[:4]).intersects(it[0] + (-0.5, -0.5, 0.5, 0.5)) for w in words)]
        base = {}
        for r in bars:
            base.setdefault(round(r.y1), []).append(r.height)
        bar_ok = any(len(v) >= 5 and max(v) - min(v) > 5 for v in base.values())
        try:
            tabs = [fitz.Rect(t.bbox) for t in page.find_tables().tables]
        except Exception:
            tabs = []
        in_tab = max(((box & t).get_area() / box.get_area() for t in tabs if t.intersects(box)), default=0)
        if _CHART_DEBUG:
            print("  chart?", [round(x) for x in box], "diag", diag, "curve", curve, "rot", nrot, "bar", bar_ok,
                  "tab", round(in_tab, 2), "n", len(g))
        # 표(선으로 칸이 나뉜 것)와 겹치면 차트로 안 본다 — 단 기울어진 축 글·막대가 있으면 차트
        chars = "".join(w[4] for w in words)
        chars = [c for c in chars if not c.isspace()]
        digit = sum(1 for c in chars if c.isdigit() or c in ".,%()-~/년월") / max(1, len(chars))
        if _CHART_DEBUG:
            print("    digit", round(digit, 2), "chars", len(chars))
        # 차트 글은 대부분 숫자(축 눈금·%·날짜) — 표·그림 설명(한글 많음)과 가른다
        if in_tab < 0.5 and digit >= 0.55 and (nrot >= 3 or bar_ok or diag >= 4 or curve >= 24):
            # 축·범례 글(바로 붙은 짧은 글)까지 넣는다
            for b in page.get_text("dict")["blocks"]:
                for l in b.get("lines", []):
                    lr = fitz.Rect(l["bbox"])
                    t = "".join(s["text"] for s in l["spans"]).strip()
                    if t and len(t) <= 20 and (box + (-25, -18, 25, 18)).intersects(lr):
                        box |= lr
            out.append(box + (-2, -2, 2, 2))
    return out


def charts_to_images(page):
    """차트 자리를 원본 그대로 그림으로 바꿔 넣는다(PDF 안에서). 바꾼 수."""
    boxes = find_charts(page)
    for box in boxes:
        pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=box)
        png = pix.tobytes("png")
        page.add_redact_annot(box, fill=False)
        _apply(page, images=True, graphics="covered", text=True)
        page.insert_image(box, stream=png)
    return len(boxes)


def strip_band_logos(page):
    """위·아래 머리말 띠 안의 작은 그림(로고)을 지운다 — 한 쪽에만 있어도(되풀이 여부 상관없이)."""
    h, w = page.rect.height, page.rect.width
    n = 0
    infos = page.get_image_info()
    for info in infos:
        r = fitz.Rect(info["bbox"])
        in_band = r.y1 <= h * TOP_BAND or r.y0 >= h * (1 - BOTTOM_BAND)
        # ★더 큰 그림 위에 얹힌 작은 그림(평면도 위 '1F' 표시 등)은 로고가 아니다. 지우면 PDF 지우기가 겹친 큰 그림까지
        #   통째로 지워 내곡동 52쪽 'B Type 평면도' 가 사라졌다(2026-10-15).
        on_big = any(fitz.Rect(o["bbox"]).intersects(r) and fitz.Rect(o["bbox"]).get_area() > r.get_area() * 4
                     for o in infos if o is not info)
        if in_band and not on_big and r.width < w * 0.35 and r.height < h * 0.10:
            page.add_redact_annot(r + (-1, -1, 1, 1), fill=False)
            n += 1
    if n:
        _apply(page, images=True, graphics=None, text=False)
    return n


def _apply(page, images=True, graphics="covered", text=True):
    """지우기 실행. graphics: None(선 안 건드림) / 'covered'(범위에 완전히 든 선만) / 'touched'(걸친 선도)."""
    kw = {"images": fitz.PDF_REDACT_IMAGE_REMOVE if images else fitz.PDF_REDACT_IMAGE_NONE,
          "graphics": {"covered": fitz.PDF_REDACT_LINE_ART_REMOVE_IF_COVERED,
                       "touched": fitz.PDF_REDACT_LINE_ART_REMOVE_IF_TOUCHED}.get(graphics, fitz.PDF_REDACT_LINE_ART_NONE),
          "text": fitz.PDF_REDACT_TEXT_REMOVE if text else fitz.PDF_REDACT_TEXT_NONE}
    try:
        page.apply_redactions(**kw)
    except TypeError:                                       # 옛 판에는 graphics·text 인자가 없다
        page.apply_redactions(images=kw["images"])


# ──────────────────────────────────────────────
# 글씨를 그림으로 그린 PDF 바로잡기
# ──────────────────────────────────────────────
# ★성수동(2026-10-08): 파워포인트 글자 효과로 만든 PDF 는 보이는 글씨가 '2×2 색 그림 + 글자 모양 마스크' 이고
#   진짜 글씨는 그 위에 투명도 0 으로 깔려 있다(쪽마다 그림 260~480개). pdf2docx 가 이 그림을 하나하나 풀다가
#   한 쪽에 몇 분씩 걸려 웹에서 10분 넘게 안 끝났고, 끝나도 워드에 글씨가 두 번(그림+글) 나온다.
#   → 그 그림을 빼고 진짜 글씨를 그림 색으로 보이게 한다(굵게 보이던 테두리 칠도 같은 색). 2쪽 60초+ → 0.5초.
TINY_PX = 4                                        # 이 이하 크기의 그림만 '색 칠용' 으로 본다
HEAVY_IMAGES = 150                                 # 한 쪽 그림 조각이 이만큼 넘으면 쪽 통째 그림으로(샘플 17개 중 성수동 15쪽만)
_NUM = r"-?(?:\d+\.?\d*|\.\d+)"


def _alpha0_states(doc, page):
    """투명도 0(ca 0 / CA 0) 그리기 상태 이름들 — 안 보이게 깐 글씨에 쓰인다."""
    out = set()
    try:
        val = doc.xref_get_key(page.xref, "Resources/ExtGState")[1]
    except Exception:
        return out
    for name, xr in re.findall(r"/([\w.\-]+)\s+(\d+)\s+0\s+R", val or ""):
        if re.search(r"/(ca|CA)\s+0(?:\.0+)?(?![\d.])", doc.xref_object(int(xr))):
            out.add(name)
    return out


def _image_kind(doc, x, cache):
    """'gone' = 마스크가 전부 0(완전히 안 보이는 그림), (r,g,b) = 마스크 있는 아주 작은 색 그림, None = 그대로."""
    xref, smask, w, h = x[0], x[1], x[2], x[3]
    if xref not in cache:
        k = None
        try:
            if smask and max(fitz.Pixmap(doc, smask).samples or b"\0") == 0:
                k = "gone"
            elif smask and w <= TINY_PX and h <= TINY_PX:
                pix = fitz.Pixmap(doc, xref)
                if not pix.colorspace or pix.colorspace.n != 3:
                    pix = fitz.Pixmap(fitz.csRGB, pix)
                s, n, m = pix.samples, pix.n, pix.width * pix.height
                k = tuple(sum(s[i + c] for i in range(0, len(s), n)) / m / 255 for c in range(3))
        except Exception:
            k = None
        cache[xref] = k
    return cache[xref]


def fix_image_text(doc):
    """그림으로 그린 글씨 → 진짜 글씨로. 바꾼 개수를 돌려준다.
    진짜 글씨가 바로 뒤에 투명하게 깔려 있을 때만 바꾸고, 아니면 그림을 그대로 둔다."""
    cache, done = {}, 0
    for page in doc:
        names = {}
        for x in page.get_images(full=True):
            if x[9] == 0:                              # 쪽 내용에 바로 붙은 그림만(폼 안 그림은 안 건드림)
                k = _image_kind(doc, x, cache)
                if k is not None:
                    names[x[7]] = k
        if not names:
            continue
        hidden = _alpha0_states(doc, page)
        page.clean_contents()
        cx = page.get_contents()[0]
        s = doc.xref_stream(cx).decode("latin-1")
        pat = re.compile(r"/(" + "|".join(re.escape(n) for n in names) + r")\s*Do\b")
        out, pos, n = [], 0, 0
        for m in pat.finditer(s):
            if m.start() < pos:
                continue
            k = names[m.group(1)]
            if k == "gone":
                out.append(s[pos:m.start()] + " ")
                pos, n = m.end(), n + 1
                continue
            bt = s.find("BT", m.end())
            et = s.find("ET", bt) if bt >= 0 else -1
            seg = s[bt:et] if et > 0 else ""
            if (not seg or bt - m.end() > 600 or pat.search(s, m.end(), bt)
                    or not any(re.search(r"/" + re.escape(h) + r"\s+gs", seg) for h in hidden)):
                continue
            for h in hidden:
                seg = re.sub(r"/" + re.escape(h) + r"\s+gs", " ", seg)
            seg = re.sub(r"\b[37]\s+Tr\b", "0 Tr", seg)                # 3·7 = 안 그리는 글씨
            seg = re.sub(r"(?:%s\s+){1,4}(?:g|rg|k|sc|scn|G|RG|K|SC|SCN)\b" % _NUM, " ", seg)
            seg = seg.replace("BT", "BT %.4f %.4f %.4f rg %.4f %.4f %.4f RG " % (k + k), 1)
            out.append(s[pos:m.start()] + " " + s[m.end():bt] + seg)
            pos, n = et, n + 1
        out.append(s[pos:])
        if n:
            doc.update_stream(cx, "".join(out).encode("latin-1"))
            done += n
    return done


def content_box(page):
    """띠를 지운 뒤 쪽에 남은 내용(글·그림·선)의 범위."""
    r = fitz.Rect()
    for b in page.get_text("dict")["blocks"]:
        if b["type"] == 0 and not "".join(s["text"] for l in b["lines"] for s in l["spans"]).strip():
            continue
        r |= fitz.Rect(b["bbox"])
    for d in page.get_drawings():
        dr = d["rect"]
        if dr.width > 1 or dr.height > 1:
            r |= dr
    for info in page.get_image_info():
        r |= fitz.Rect(info["bbox"])
    return r & page.rect if not r.is_empty else fitz.Rect(0, 0, 1, 1)


# ──────────────────────────────────────────────
# 워드로 바꾸고 규격 맞추기
# ──────────────────────────────────────────────
class _CatchIgnored(logging.Handler):
    """pdf2docx 기록에서 '버린 쪽' 을 모으고, 진행 상황을 progress(0~1, 글) 로 알린다."""
    def __init__(self, progress=None):
        super().__init__()
        self.pages = []
        self.progress = progress
        self.stage = 0

    def emit(self, record):
        msg = record.getMessage()
        m = re.search(r"Ignore page (\d+)", msg)
        if m:
            self.pages.append(int(m.group(1)))
        if not self.progress:
            return
        if "[2/4]" in msg:
            self.progress(0.03, "원본 짜임새 살펴보는 중…")
        elif "[3/4]" in msg:
            self.stage = 3
        elif "[4/4]" in msg:
            self.stage = 4
        m = re.match(r"\((\d+)/(\d+)\) Page", msg)
        if m and self.stage in (3, 4):
            i, n = int(m.group(1)), int(m.group(2))
            if self.stage == 3:
                self.progress(0.05 + 0.65 * i / n, f"원본 읽는 중… {i} / {n}쪽")
            else:
                self.progress(0.70 + 0.25 * i / n, f"워드에 옮기는 중… {i} / {n}쪽")


# ★pdf2docx 속도(2026-10-12): '원본 살펴보기' 단계의 묶기(Collection.group)가 모든 짝을 다 비교해서(n²)
#   고산3지구 30쪽에서만 60초(웹은 몇 분) — 진행 막대가 첫 단계에 멈춘 것처럼 보였다.
#   묶기 조건 자체는 그대로 두고, 위치를 정렬해 '절대 닿을 수 없는 먼 짝' 만 건너뛴다 → 결과는 똑같고 빠르다.
#   빠르게 하는 건 아래 세 가지(줄 묶기·칸 묶기·같은 색 도형 합치기)뿐이고, 모르는 조건은 원래 방식 그대로.
_GROUP_MARGIN = 2.0          # 건너뛰기 여유(pt) — 조건의 허용치(1e-3, TINY_DIST)보다 넉넉하게


def _group_axis(fun, items):
    """이 묶기 조건이 어느 축(0=가로, 1=세로)에서 겹쳐야만 참이 되는지. 모르면 None(원래 방식)."""
    name = getattr(fun, "__qualname__", "")
    if name.startswith("Shapes._merge_shapes"):
        return 0                                           # 겹치거나 맞닿아야(TINY_DIST) 함 → 가로로도 겹침
    if name.startswith(("Collection.group_by_rows", "Collection.group_by_columns")):
        cells = dict(zip(fun.__code__.co_freevars, (c.cell_contents for c in (fun.__closure__ or ()))))
        if cells.get("text_direction") and any(getattr(e, "is_vertical_text", False) for e in items):
            return None                                    # 세로쓰기 섞이면 원래 방식
        return 1 if "rows" in name else 0                  # 줄 = 세로 범위 겹침, 칸 = 가로 범위 겹침
    return None


def _fast_group(self, fun, _orig):
    items = self._instances
    axis = _group_axis(fun, items) if len(items) > 60 else None
    if axis is None:
        return _orig(self, fun)
    lo = [e.bbox[axis] for e in items]
    hi = [e.bbox[axis + 2] for e in items]
    if any(h < l for l, h in zip(lo, hi)):
        return _orig(self, fun)
    from pdf2docx.common.algorithm import graph_bfs
    order = sorted(range(len(items)), key=lambda i: lo[i])
    edges = []
    for p, i in enumerate(order):
        lim = hi[i] + _GROUP_MARGIN
        for j in order[p + 1:]:
            if lo[j] > lim:
                break                                      # 이 뒤는 더 멀다 — 조건이 참일 수 없음
            a, b = (i, j) if i < j else (j, i)
            if fun(items[a], items[b]):                    # 원래와 같은 순서로 같은 조건 검사
                edges.append((a, b))
    index_groups = [set() for _ in items]
    for a, b in sorted(edges):                             # 원래와 같은 순서로 넣어 묶음 순서까지 같게
        index_groups[a].add(b)
        index_groups[b].add(a)
    return [self.__class__([items[i] for i in g]) for g in graph_bfs(index_groups)]


def convert(pdf_bytes, drop_cover=True, drop_last=None, progress=None, drop_pages=None):
    """PDF bytes → (docx bytes, 보고 dict).

    보고['쪽'] = 결과 구역 순서대로 원본 쪽 번호(1부터)·줄인 비율·쓰인 높이(pt)·그림으로 넣었는지.
    progress(비율 0~1, 글) 를 주면 진행 상황을 알린다.
    """
    if progress:
        progress(0.01, "원본 정리하는 중… (머리말·꼬리말·연락처 지우기)")
    import docx
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = doc.page_count
    rep = {"원본 쪽수": n}
    if drop_last is None:
        drop_last = n > 2 and last_page_is_contact(doc)
    keep = list(range(n))
    if drop_last and n > 2:
        keep.remove(n - 1)
    if drop_cover and n > 1:
        keep.remove(0)
    # 원본 하이라이트 쪽은 회사 Investment Highlight 쪽으로 바뀐다 → 뺀다(2026-10-12 사용자 확정). 1부터 센 쪽 번호.
    rep["원본 하이라이트 뺀 쪽"] = []
    for pg in sorted(set(drop_pages or [])):
        if 0 <= pg - 1 < n and (pg - 1) in keep and len(keep) > 1:
            keep.remove(pg - 1)
            rep["원본 하이라이트 뺀 쪽"].append(pg)
    # 원본 목차·차례 쪽은 뺀다(2026-10-14 사용자 확정 — 대단원 번호가 바뀌어 원본 목차가 맞지 않는다)
    rep["원본 목차 뺀 쪽"] = []
    for i in find_toc_pages(doc):
        if i in keep and len(keep) > 1:
            keep.remove(i)
            rep["원본 목차 뺀 쪽"].append(i + 1)
    rep["원본 끝쪽 뺌"], rep["원본 표지 뺌"] = bool(drop_last and n > 2), bool(drop_cover and n > 1)
    rep["그림 글씨 고침"] = fix_image_text(doc)
    # 하이라이트와 본문이 섞인 쪽(표지 다음 6쪽 안) → 하이라이트 부분만 지운다(회사 하이라이트 쪽으로 대체)
    rep["원본 하이라이트 일부 지운 쪽"] = []
    for i in []:                       # 2026-10-13 끔: 회사 하이라이트 쪽을 안 쓰므로 원본 하이라이트를 지우지 않는다
        band = highlight_band(doc[i])
        if band is None:
            continue
        y0, y1 = band
        page = doc[i]
        page.add_redact_annot(fitz.Rect(0, y0 - 4, page.rect.width, y1 - 4), fill=False)
        _apply(page, images=True, graphics="covered", text=True)
        rep["원본 하이라이트 일부 지운 쪽"].append(i + 1)
    leftover = []
    rep["머리말·꼬리말 지운 쪽"] = strip_repeating_bands(doc, leftover)
    # 원본 담당자 연락처(어느 쪽이든)·머리말 띠의 로고(한 쪽에만 있어도) — 2026-10-08 사용자 확정
    rep["연락처 지운 쪽"] = [i + 1 for i in keep if strip_contacts(doc[i])]
    rep["로고 지운 쪽"] = [i + 1 for i in keep if strip_band_logos(doc[i])]
    rep["차트 그림으로 넣은 쪽"] = [i + 1 for i in keep if charts_to_images(doc[i])]
    boxes = [content_box(doc[i]) for i in range(n)]
    page_lines = {i: page_text_lines(doc[i]) for i in keep}
    sizes = [(doc[i].rect.width, doc[i].rect.height) for i in range(n)]
    # 그림 조각이 아주 많은 쪽(성수동 15쪽 지도 = 조각 237개) 은 pdf2docx 가 한 쪽에 몇 분씩 걸린다 → 처음부터 쪽 그림으로
    heavy = {i for i in keep if len(doc[i].get_image_info()) >= HEAVY_IMAGES}
    rep["그림 조각 많아 그림으로"] = [i + 1 for i in sorted(heavy)]
    conv = [i for i in keep if i not in heavy]
    doc.select(conv)

    from pdf2docx import Converter
    import docx.table as _dt
    catcher = _CatchIgnored(progress)
    root = logging.getLogger()
    old_level = root.level
    root.addHandler(catcher)
    root.setLevel(min(old_level or logging.INFO, logging.INFO))     # 쪽 진행 기록(INFO)까지 받도록
    # ★pdf2docx 는 표 칸 합치기가 하나라도 겹치면 **그 쪽을 통째로 버린다**('Ignore page …').
    #   안산 8쪽('정보통신업 사업자 관련 설명' 표)·33쪽이 원본 그대로도 이렇게 빠졌다(2026-10-08).
    #   → 바꾸는 동안만, 합치기가 안 되는 칸 하나는 합치지 않고 넘어가게 해 표 나머지를 살린다.
    _orig_merge = _dt._Cell.merge

    def _safe_merge(self, other):
        try:
            return _orig_merge(self, other)
        except Exception:
            rep.setdefault("못 합친 칸", 0)
            rep["못 합친 칸"] += 1
            return self
    _dt._Cell.merge = _safe_merge
    from pdf2docx.common.Collection import Collection as _Coll
    _orig_group = _Coll.group
    _Coll.group = lambda self, fun: _fast_group(self, fun, _orig_group)
    # ★칸 바탕색: pdf2docx 는 '칠한 네모 하나가 칸을 거의(90%) 덮고 길이까지 딱 맞아야' 바탕색으로 본다.
    #   원본 칠이 칸보다 조금 작거나 두 조각이면 색이 빠져 흰 칸이 됐다(내곡동 '구 분|내용' 머리글 → 흰 글씨가 안 보임,
    #   구분 칸 연한 파랑도 빠짐, 2026-10-12). → 바꾸는 동안만: 칸을 가장 많이 덮는 칠이 60% 넘으면 그 색.
    from pdf2docx.table.TableStructure import CellStructure as _CS
    from pdf2docx.common.Element import Element as _El
    _orig_shading = _CS.parse_shading

    def _loose_shading(self, fills):
        _orig_shading(self, fills)
        if self.shading is not None:
            return
        x0, y0, x1, y1 = self.merged_bbox
        cell = fitz.Rect(x0, y0, x1, y1)
        area = cell.get_area()
        if not area:
            return
        best, best_r = None, 0.0
        for shape in fills:
            r = (fitz.Rect(shape.bbox) & cell).get_area() / area
            if r > best_r:
                best, best_r = shape, r
        if False and best is not None and best_r >= 0.6:   # 2026-10-13 끔: 망가진 표에 검은 덩어리를 퍼뜨림(방배동). 빠진 머리글 색은 fill_white_text_cells 가 채운다
            self.shading = best
    _CS.parse_shading = _loose_shading
    try:
        with tempfile.TemporaryDirectory() as td:
            src, dst = os.path.join(td, "in.pdf"), os.path.join(td, "out.docx")
            doc.save(src)
            cv = Converter(src)
            try:
                cv.convert(dst)
            finally:
                cv.close()
            with open(dst, "rb") as f:
                raw = f.read()
        try:
            from pdf2docx.font.Fonts import Fonts
            font_lh = {f.name: f.line_height for f in Fonts.extract(doc) if f.line_height}
        except Exception:
            font_lh = {}
        try:
            width_ratio = font_width_ratios(doc)
        except Exception:
            width_ratio = {"*": 1.0}
    finally:
        root.removeHandler(catcher)
        root.setLevel(old_level)
        _dt._Cell.merge = _orig_merge
        _Coll.group = _orig_group
        _CS.parse_shading = _orig_shading
    if progress:
        progress(0.96, "회사 양식 크기에 맞추는 중…")
    ignored = {conv[p - 1] for p in catcher.pages if 0 < p <= len(conv)} | heavy

    d = docx.Document(io.BytesIO(raw))
    # 망가진 '표 안의 표' 는 다른 손질 전에 원본 그림으로(손질 뒤엔 채운 칸이 늘어 멀쩡한 표까지 걸린다)
    _orig = fitz.open(stream=pdf_bytes, filetype="pdf")
    _done0 = [k for k in keep if k not in ignored]
    # ★표는 원본 선·칠·글 자리로 다시 짠다(칸 합치기·안쪽 표까지 원본 그대로, 2026-10-15 사용자 22번 '가장 중요')
    import table_rebuild
    rep["원본 선으로 다시 짠 표"] = []
    for (_secs, _grps), k in zip(_pages(d), _done0):
        try:
            got = table_rebuild.rebuild_page(doc[conv.index(k)], _grps, d.part)
        except Exception as e:                                # 못 짜면 pdf2docx 표 그대로
            logging.getLogger(__name__).warning("표 다시 짜기 실패 %s쪽: %s", k + 1, e)
            got = []
        if got:
            rep["원본 선으로 다시 짠 표"].append((k + 1, len(got)))
    rep["가짜 굵게 살림"] = 0
    for (_secs, _grps), k in zip(_pages(d), _done0):
        try:
            rep["가짜 굵게 살림"] += apply_bold_segments(_grps, fake_bold_segments(doc[conv.index(k)]))
        except Exception as e:
            logging.getLogger(__name__).warning("굵게 살리기 실패 %s쪽: %s", k + 1, e)
    rep["망가진 안쪽 표 그림으로"] = []
    for (_secs, _grps), k in zip(_pages(d), _done0):
        if clean_broken_nested_tables(d, _grps):   # 2026-10-13 끔: 그림(색 안 바뀜)·다시 짜기(그림·글 빠짐) 둘 다 사용자 기준 미달
            rep["망가진 안쪽 표 그림으로"].append(k + 1)
    rep["워드에서 지운 머리말 줄"] = drop_paragraphs(d, set(leftover))
    # fix_wrap_tabs(줄바꿈 탭→띄어쓰기)는 안 쓴다 — 목차·글머리 줄을 이어 붙여 원본을 망가뜨렸다(내곡동, 2026-10-12).
    # 대신 글자 폭을 원본과 같게 맞춰(match_char_width) 탭·줄바꿈이 원본 자리에 그대로 떨어지게 한다.
    rep["기호 글꼴 고침"] = fix_symbol_runs(d)
    rep["줄 간격 글꼴 맞춤"] = match_line_spacing(d, font_lh)
    rep["글자 폭 원본 맞춤"] = match_char_width(d, width_ratio)
    rep["글머리 앞 줄바꿈"] = break_before_bullets(d)
    rep["글머리 표 → 문단"] = bullet_tables_to_paragraphs(d)
    rep["글머리 줄 문단 나눔"] = split_bullet_lines(d)
    rep["표 밖 줄바꿈·탭 원본 줄대로"] = split_body_breaks(d)
    rep["글머리 내어쓰기"] = hang_bullets(d, width_ratio)
    rep["표에서 온 글머리 위치 맞춤"] = align_table_bullets(d)
    rep["글머리 다음 줄 맞춤"] = align_continuations(d)
    rep["줄바꿈 탭·쪼개진 줄·줄 간격 정리"] = normalize_text_flow(d)
    rep["칸 좌우 여백 같게"] = even_cell_margins(d)
    rep["선만 있는 표 지움"] = drop_grid_tables(d)
    rep["칸 안 작은 표 풂"] = unwrap_cell_boxes(d)
    rep["쪼개진 표 합침"] = merge_split_tables(d)
    rep["칸 안 쪼개진 표 합침"] = merge_split_tables_in_cells(d)
    rep["흰 글씨 칸 바탕"] = fill_white_text_cells(d)
    rep["표 줄 높이 풀어줌"] = loosen_row_heights(d)
    rep["흰 바탕 투명하게"] = clear_white_overlays(d)
    done = [k for k in keep if k not in ignored]
    page_info = []
    for (secs, grps), k in zip(_pages(d), done):
        s, used = _fit_page(secs, grps, boxes[k], page_lines.get(k))
        page_info.append({"원본쪽": k + 1, "비율": s, "높이": used, "그림": False})
    # 실패해 건너뛴 쪽 → 원본 쪽을 그림으로(제자리에)
    for k in sorted(ignored):
        idx = sum(1 for p in done if p < k) + sum(1 for p in ignored if p < k)
        img = fitz.open(stream=pdf_bytes, filetype="pdf")[k].get_pixmap(matrix=fitz.Matrix(3, 3), clip=boxes[k])
        used = _insert_picture_page(d, idx, img.tobytes("png"), boxes[k])
        page_info.insert(idx, {"원본쪽": k + 1, "비율": None, "높이": used, "그림": True})
    rep["쪽"] = page_info
    rep["그림으로 넣은 쪽"] = [k + 1 for k in sorted(ignored)]
    out = io.BytesIO()
    d.save(out)
    return out.getvalue(), rep


def _section_groups(d):
    """구역(=원본 한 쪽)마다 블록 묶음. 구역은 '문단 안 sectPr' 로 끝난다(마지막은 본문 sectPr)."""
    from docx.oxml.ns import qn
    groups, cur = [], []
    for el in d.element.body:
        if el.tag not in (qn("w:p"), qn("w:tbl")):
            continue
        cur.append(el)
        if el.tag == qn("w:p") and el.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
            groups.append(cur)
            cur = []
    groups.append(cur)
    return groups


def fake_bold_segments(page):
    """'채우기 + 외곽선' 으로 두껍게 그린 글(가짜 굵게) 조각들(빈칸 뺀 글). ★고산3지구 원본 굵은 글씨는 굵은 글꼴이 아니라
    같은 글자를 외곽선으로 한 번 더 그린 것 → pdf2docx 가 못 알아봐 굵게가 빠졌다(2026-10-15)."""
    boxes = []
    try:
        for t in page.get_texttrace():
            if t.get("type") in (1, 2) and (t.get("linewidth") or 0) > 0:
                for c in t["chars"]:
                    boxes.append(fitz.Rect(c[3]))
    except Exception:
        return []
    if not boxes:
        return []
    segs = []
    for b in page.get_text("rawdict")["blocks"]:
        for l in b.get("lines", []):
            cur = ""
            for sp in l["spans"]:
                for c in sp["chars"]:
                    ch = c["c"]
                    if not ch.strip():
                        continue
                    r = fitz.Rect(c["bbox"])
                    cen = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
                    if any(bx.contains(cen) for bx in boxes):
                        cur += ch
                    else:
                        if len(cur) >= 4:
                            segs.append(cur)
                        cur = ""
            if len(cur) >= 4:
                segs.append(cur)
    return segs


def apply_bold_segments(grps, segs):
    """docx 쪽(구역 묶음) 안 문단에서 segs 글을 찾아 그 글자만 굵게(글 조각을 나눠서). 굵게 한 곳 수."""
    import copy
    from docx.oxml.ns import qn
    if not segs:
        return 0
    segs = sorted(set(segs), key=len, reverse=True)
    n = 0
    for el in grps_iter(grps):
        for p in list(el.iter(qn("w:p"))):
            runs = [r for r in p if r.tag == qn("w:r") and r.find(qn("w:t")) is not None]
            if not runs:
                continue
            chars = []                                         # (run 번호, 글자)
            for i, r in enumerate(runs):
                for ch in "".join(t.text or "" for t in r.findall(qn("w:t"))):
                    chars.append((i, ch))
            compact = [(k, ch) for k, (i, ch) in enumerate(chars) if ch.strip()]
            cs = "".join(ch for _, ch in compact)
            mask = [False] * len(chars)
            for sg in segs:
                start = 0
                while True:
                    j = cs.find(sg, start)
                    if j < 0:
                        break
                    a, b = compact[j][0], compact[j + len(sg) - 1][0]
                    for k in range(a, b + 1):
                        mask[k] = True
                    start = j + len(sg)
            if not any(mask):
                continue
            k = 0
            for i, r in enumerate(runs):
                txt = "".join(t.text or "" for t in r.findall(qn("w:t")))
                m = mask[k:k + len(txt)]
                k += len(txt)
                if not txt or not any(m):
                    continue
                pieces, cur, cb = [], "", m[0]
                for ch, mb in zip(txt, m):
                    if mb != cb and ch.strip():
                        pieces.append((cur, cb))
                        cur, cb = "", mb
                    cur += ch
                pieces.append((cur, cb))
                rpr = r.find(qn("w:rPr"))
                for text, bold in pieces:
                    nr = r.makeelement(qn("w:r"), {})
                    np_ = copy.deepcopy(rpr) if rpr is not None else nr.makeelement(qn("w:rPr"), {})
                    nr.append(np_)
                    if bold:
                        for e in np_.findall(qn("w:b")):
                            np_.remove(e)
                        be = np_.makeelement(qn("w:b"), {})
                        rf = np_.find(qn("w:rFonts"))
                        (rf.addnext(be) if rf is not None else np_.insert(0, be))
                        n += 1
                    t = nr.makeelement(qn("w:t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
                    t.text = text
                    nr.append(t)
                    r.addprevious(nr)
                r.getparent().remove(r)
    return n


def drop_paragraphs(d, texts):
    """PDF 에서 안전하게 못 지운 머리말·꼬리말 줄을 워드에서 지운다(그 글과 똑같은 문단만, 표 밖만)."""
    from docx.oxml.ns import qn
    if not texts:
        return 0
    want = {re.sub(r"\s+", "", t) for t in texts}
    n = 0
    # 머리말·꼬리말은 늘 그 쪽(구역)의 맨 처음·맨 끝에 있다 → 구역마다 글 있는 문단 앞뒤 3개만 본다
    #   (숫자만 있는 쪽번호 '30' 이 본문 어딘가의 '30' 문단까지 지우지 않게)
    for grp in _section_groups(d):
        paras = [e for e in grp if e.tag == qn("w:p") and "".join(x.text or "" for x in e.iter(qn("w:t"))).strip()]
        for p in paras[:3] + paras[-3:]:
            if p.getparent() is None:
                continue
            t = re.sub(r"\s+", "", "".join(x.text or "" for x in p.iter(qn("w:t"))))
            if t and t in want and p.find(qn("w:pPr") + "/" + qn("w:sectPr")) is None:
                p.getparent().remove(p)
                n += 1
    return n


def fix_wrap_tabs(d):
    """pdf2docx 가 '줄이 바뀌는 자리' 에 넣은 탭을 띄어쓰기로 바꾸고, 둘째 줄부터 들여쓰기로 대신한다.

    ★원본은 글머리 문단의 둘째 줄을 글머리 뒤에 맞추려고 [첫 줄 글][탭][둘째 줄 글] 로 만들어져 있다.
      글꼴이 Pretendard 로 바뀌어 첫 줄이 짧아지면 탭이 같은 줄 오른쪽으로 튀어
      '…법인 요건을          갖출 예정' 처럼 벌어졌다(안산 8쪽, 2026-10-08).
    조건: 문단에 직접 정한 탭 자리가 1~3개이고 모두 왼쪽 들여쓰기 근처(2cm 안)에 있으며, 탭 앞에 글이 있을 때.
      연달아 있는 탭(안산 '[부가통신사업의 정의]' 는 2개)은 띄어쓰기 하나로. 둘째 줄부터는 가장 먼 탭 자리에.
    """
    from docx.oxml.ns import qn
    n = 0
    for p in d.element.body.iter(qn("w:p")):
        ppr = p.find(qn("w:pPr"))
        tabs = ppr.find(qn("w:tabs")) if ppr is not None else None
        stops = tabs.findall(qn("w:tab")) if tabs is not None else []
        if not 1 <= len(stops) <= 3:
            continue
        ind = ppr.find(qn("w:ind"))
        left = int(ind.get(qn("w:left"), "0")) if ind is not None and (ind.get(qn("w:left")) or "0").lstrip("-").isdigit() else 0
        poss = [int(t.get(qn("w:pos"), "0")) for t in stops]
        pos = max(poss)
        if pos <= left or pos > left + 1134:
            continue
        seen_text, changed, prev_tab = False, False, False
        for r in p.iter(qn("w:r")):
            for c in list(r):
                if c.tag == qn("w:t") and (c.text or "").strip() and not (c.text or "").strip() in "▪•·∙-–※✓➢◦○●■□▶▷":
                    seen_text, prev_tab = True, False
                elif c.tag == qn("w:br"):
                    prev_tab = False
                elif c.tag == qn("w:tab") and seen_text:
                    if not prev_tab:                       # 연달아 있는 탭은 띄어쓰기 하나로
                        t = r.makeelement(qn("w:t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
                        t.text = " "
                        c.addprevious(t)
                    r.remove(c)
                    changed, prev_tab = True, True
        if changed:
            if ind is None:
                ind = ppr.makeelement(qn("w:ind"), {})
                after = next((ppr.find(qn(t)) for t in ("w:jc", "w:rPr", "w:sectPr")
                              if ppr.find(qn(t)) is not None), None)       # ind 는 jc·rPr·sectPr 보다 앞
                if after is not None:
                    after.addprevious(ind)
                else:
                    ppr.append(ind)
            ind.attrib.pop(qn("w:firstLine"), None)
            ind.set(qn("w:left"), str(pos))
            ind.set(qn("w:hanging"), str(pos - left))
            n += 1
    return n


_SYMBOL_FONT_RE = re.compile(r"wingding|webding|symbol|marlett", re.I)


def match_line_spacing(d, font_lh):
    """'배수' 줄 간격을 Pretendard 기준으로 고쳐 원본 줄 간격(pt)을 그대로 지킨다. 고친 문단 수를 돌려준다.
    ★pdf2docx 는 줄 간격을 '원본 줄 높이 ÷ (원본 글꼴 한 줄 높이 × 글자 크기)' 배수로 적는다(글꼴 파일에서 읽은 값,
    맑은고딕 ≈1.33·나눔고딕 ≈1.15). 글꼴을 Pretendard(워드에서 1.5625)로 바꾸면 모든 줄이 그만큼 높아져
    표 칸 글씨가 잘리고 쪽이 넘쳤다(2026-10-12). → 배수 × (원본 글꼴 줄 높이 ÷ 1.5625).
    한 문단의 기준은 pdf2docx 와 같게 '가장 큰 글자' 의 글꼴."""
    from docx.oxml.ns import qn
    n = 0
    for p in d.element.body.iter(qn("w:p")):
        sp = p.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
        if sp is None or (sp.get(qn("w:lineRule")) or "auto") != "auto" or not (sp.get(qn("w:line")) or "").isdigit():
            continue
        best = None
        for r in p.iter(qn("w:r")):
            if not "".join(t.text or "" for t in r.iter(qn("w:t"))).strip():
                continue
            rpr = r.find(qn("w:rPr"))
            sz = rpr.find(qn("w:sz")) if rpr is not None else None
            size = int(sz.get(qn("w:val"))) if sz is not None and (sz.get(qn("w:val")) or "").isdigit() else 20
            rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
            name = (rf.get(qn("w:ascii")) or rf.get(qn("w:eastAsia"))) if rf is not None else None
            if best is None or size > best[0]:
                best = (size, name)
        f = font_lh.get(best[1]) if best else None
        if not f or f >= LINE_EM:
            # 원본 글꼴 줄이 Pretendard 보다 높으면(맑은고딕 1.729) 바꿔도 줄이 낮아질 뿐 → 그대로 둔다.
            # 원본에 맞춰 늘리면 여유가 사라져 쪽이 넘쳤다(돈암 14→17쪽, 2026-10-12).
            continue
        sp.set(qn("w:line"), str(max(20, int(round(int(sp.get(qn("w:line"))) * f / LINE_EM)))))
        n += 1
    return n


def _font_key(name):
    return re.sub(r"[^a-z0-9가-힣]", "", (name or "").lower().split("+")[-1])


def font_width_ratios(doc):
    """원본 PDF 글꼴마다 '원본 글자 폭 합 ÷ 같은 글자의 Pretendard 폭 합'. {(글꼴키, 굵게?): 비율, '*': 전체}."""
    W = _widths()
    acc = {}
    for p in doc:
        for b in p.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                if l.get("dir") != (1.0, 0.0):
                    continue
                for s in l["spans"]:
                    bold = "bold" in s["font"].lower() or bool(s.get("flags", 0) & 16)
                    tbl = W["bold" if bold else "light"]
                    key = (_font_key(re.sub(r"(?i)[-,]?(bold|regular|light|medium)$", "", s["font"])), bold)
                    for c in s["chars"]:
                        ow = c["bbox"][2] - c["bbox"][0]
                        if not c["c"].strip() or ow <= 0 or ow > s["size"] * 2:
                            continue
                        a = acc.setdefault(key, [0.0, 0.0])
                        a[0] += ow
                        a[1] += s["size"] * _char_w(c["c"], tbl)
    out = {k: v[0] / v[1] for k, v in acc.items() if v[1] > 200}
    tot = [sum(v[0] for v in acc.values()), sum(v[1] for v in acc.values())]
    out["*"] = tot[0] / tot[1] if tot[1] else 1.0
    out["글자"] = _orig_char_em(doc)
    return out


def _orig_char_em(doc):
    """원본에서 글꼴·굵기·글자마다 실제 폭(글자 크기 1 기준, 중앙값). {(글꼴키, 굵게?, 글자): 폭}
    ★평균 비율로 모든 글자를 똑같이 벌리면 목차 점선처럼 원래 폭이 비슷한 글자까지 늘어나 줄이 넘쳤다(내곡동)."""
    from statistics import median
    acc = {}
    for p in doc:
        for b in p.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                if l.get("dir") != (1.0, 0.0):
                    continue
                for s in l["spans"]:
                    if s["size"] <= 0:
                        continue
                    bold = "bold" in s["font"].lower() or bool(s.get("flags", 0) & 16)
                    fk = _font_key(re.sub(r"(?i)[-,]?(bold|regular|light|medium)$", "", s["font"]))
                    cs = s["chars"]
                    for i, c in enumerate(cs):
                        # 폭 = 다음 글자 시작까지(글자 사이 간격 포함). 줄 마지막 글자는 상자 폭.
                        ow = (cs[i + 1]["origin"][0] - c["origin"][0]) if i + 1 < len(cs) else c["bbox"][2] - c["bbox"][0]
                        if ow <= 0 or ow > s["size"] * 2:
                            continue
                        acc.setdefault((fk, bold, c["c"]), []).append(ow / s["size"])
    out = {k: median(v) for k, v in acc.items()}
    out.update(_bullet_gap_em(doc))
    return out


BULLETS = "•▪■□◦○●➢✓※-–·∙▶▷◆◇"
WIDTH_FIT = 0.98


def _bullet_gap_em(doc):
    """줄 맨 앞 글머리 기호에서 '그 줄 첫 글자' 까지 거리(글자 크기 기준, 중앙값). {('글머리', 기호): 거리}
    ★원본은 '•' 뒤에 다른 글꼴(Arial)의 넓은 빈칸을 두어 띄웠는데 pdf2docx 가 빈칸만 든 조각을 버려
    '•서초구' 처럼 붙었다(내곡동 금융조건, 2026-10-12)."""
    from statistics import median
    acc = {}
    for p in doc:
        for b in p.get_text("rawdict")["blocks"]:
            for l in b.get("lines", []):
                if l.get("dir") != (1.0, 0.0):
                    continue
                cs = [(c, s["size"]) for s in l["spans"] for c in s["chars"]]
                if len(cs) < 2 or cs[0][0]["c"] not in BULLETS:
                    continue
                nxt = next((c for c, _ in cs[1:] if c["c"].strip()), None)
                if nxt is None or cs[0][1] <= 0:
                    continue
                gap = (nxt["origin"][0] - cs[0][0]["origin"][0]) / cs[0][1]
                if 0 < gap < 4:
                    acc.setdefault(("글머리", cs[0][0]["c"]), []).append(gap)
    return {k: median(v) for k, v in acc.items()}


def _split_bullet_runs(d):
    """'•준공 및…' 처럼 글머리 기호와 글이 한 조각이면 기호만 따로 떼어 낸다(기호 뒤 띄움을 따로 주려고)."""
    import copy
    from docx.oxml.ns import qn
    for r in list(d.element.body.iter(qn("w:r"))):
        ts = r.findall(qn("w:t"))
        if len(ts) != 1 or len([c for c in r if c.tag != qn("w:rPr")]) != 1:
            continue
        s = ts[0].text or ""
        if len(s) < 2 or s[0] not in BULLETS or s[1] in BULLETS:
            continue
        rest = s[1:].lstrip(" ")
        if not rest:
            continue
        nr = copy.deepcopy(r)
        ts[0].text = s[0]
        nt = nr.find(qn("w:t"))
        nt.text = rest
        nt.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        r.addnext(nr)


LINE_BULLETS = "•▪■□◦○●➢✓※▶▷◆◇"


def break_before_bullets(d):
    """문단 가운데 끼어 있는 글머리 기호(원본에선 늘 새 줄 맨 앞) 앞에 줄바꿈을 넣는다. 넣은 수.
    ★pdf2docx 는 원본 여러 줄을 한 문단에 넣고 '폭이 꽉 차서 넘어가는' 것에 기대 줄을 나눈다 → 글자 폭이 조금만
    달라도 '…제출 ▷자금 차입…' 처럼 다음 항목이 같은 줄로 붙었다(등촌역 인출선행조건, 2026-10-12)."""
    from docx.oxml.ns import qn
    n = 0
    for p in d.element.body.iter(qn("w:p")):
        seen = False
        for r in list(p.iter(qn("w:r"))):
            for t in r.findall(qn("w:t")):
                s = t.text or ""
                if seen and s and s[0] in LINE_BULLETS:           # ★빈 글('' in …)은 늘 참이라 빈 조각마다 줄바꿈이 들어갔다
                    prev = r.getprevious()
                    if not (prev is not None and prev.tag == qn("w:r") and prev.find(qn("w:br")) is not None):
                        br_r = r.makeelement(qn("w:r"), {})
                        if r.find(qn("w:rPr")) is not None:
                            import copy
                            br_r.append(copy.deepcopy(r.find(qn("w:rPr"))))
                        br_r.append(br_r.makeelement(qn("w:br"), {}))
                        r.addprevious(br_r)
                        # 줄바꿈 바로 앞의 탭(원래 줄 넘김용)은 지운다 — 남으면 빈 칸이 생긴다
                        pv = br_r.getprevious()
                        if pv is not None and pv.tag == qn("w:r"):
                            for tb in pv.findall(qn("w:tab")):
                                pv.remove(tb)
                        n += 1
                if s.strip():
                    seen = True
    return n


def drop_grid_tables(d):
    """글이 하나도 없는 '선만 있는 표'(pdf2docx 가 안쪽 표의 선만 따로 떼어 만든 것)를 지우고, 같은 쪽의
    칸 수가 같고 테두리 없는 표에 테두리를 준다. ★등촌역 3쪽: 'Tr.A' 안쪽 표가 테두리 없이 나오고 그 선이
    쪽 아래 빈 표로 따로 생겼다(2026-10-12)."""
    from docx.oxml.ns import qn
    n = 0
    for tbl in list(d.element.body.iter(qn("w:tbl"))):
        if "".join(t.text or "" for t in tbl.iter(qn("w:t"))).strip() or tbl.find(".//" + qn("w:drawing")) is not None:
            continue
        rows = tbl.findall(qn("w:tr"))
        if len(rows) < 1:
            continue
        ncol = max(len(r.findall(qn("w:tc"))) for r in rows)
        # 같은 쪽(구역)에서 칸 수가 같은 테두리 없는 표 찾기 — 앞쪽으로 거슬러 올라가며 구역 끝에서 멈춘다
        target, el = None, tbl.getprevious()
        while el is not None and target is None:
            if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
                break
            for t in ([el] if el.tag == qn("w:tbl") else []) + list(el.iter(qn("w:tbl"))):
                if t is tbl:
                    continue
                trs = t.findall(qn("w:tr"))
                if trs and max(len(r.findall(qn("w:tc"))) for r in trs) == ncol and t.find(".//" + qn("w:tcBorders")) is None \
                        and "".join(x.text or "" for x in t.iter(qn("w:t"))).strip():
                    target = t
                    break
            el = el.getprevious()
        if target is not None:
            tblpr = target.find(qn("w:tblPr"))
            b = tblpr.makeelement(qn("w:tblBorders"), {})
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
                e = b.makeelement(qn("w:" + side), {qn("w:val"): "single", qn("w:sz"): "4", qn("w:color"): "000000"})
                b.append(e)
            old = tblpr.find(qn("w:tblBorders"))
            if old is not None:
                tblpr.remove(old)
            after = next((tblpr.find(qn(t)) for t in ("w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
                          if tblpr.find(qn(t)) is not None), None)
            after.addprevious(b) if after is not None else tblpr.append(b)
        parent = tbl.getparent()
        if parent.tag == qn("w:tc") and len(parent.findall(qn("w:p"))) == 0:
            parent.append(parent.makeelement(qn("w:p"), {}))     # 칸엔 문단이 하나는 있어야 한다
        parent.remove(tbl)
        n += 1
    return n


BULLET_EM = {"•": 0.41, "▪": 0.86, "■": 0.86, "➢": 0.96, "✓": 0.75, "◆": 0.86, "○": 0.86, "●": 0.86, "▶": 0.86,
             "※": 1.0, "-": 0.4, "·": 0.3, "ㆍ": 0.5, "◦": 0.6, "□": 0.86, "▷": 0.86, "◇": 0.86}   # Segoe UI Symbol 실측 폭(em)
HANG_BULLETS = "•▪■□◦○●➢✓※▶▷◆◇◼◻-·∙ㆍ"


def _par_text(p):
    from docx.oxml.ns import qn
    return "".join(t.text or "" for t in p.iter(qn("w:t")))


def _ind_vals(p):
    """(left, firstLine, hanging) twips."""
    from docx.oxml.ns import qn
    ind = p.find(f"{qn('w:pPr')}/{qn('w:ind')}")

    def g(a):
        v = ind.get(qn("w:" + a)) if ind is not None else None
        return int(float(v)) if v and re.fullmatch(r"-?[\d.]+", v) else 0
    return (g("left") or g("start")), g("firstLine"), g("hanging")


def _set_ind(p, left, hanging=0, right=None):
    from docx.oxml.ns import qn
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = p.makeelement(qn("w:pPr"), {})
        p.insert(0, ppr)
    ind = ppr.find(qn("w:ind"))
    if ind is None:
        ind = ppr.makeelement(qn("w:ind"), {})
        after = next((ppr.find(qn(t)) for t in ("w:jc", "w:textAlignment", "w:outlineLvl", "w:rPr", "w:sectPr")
                      if ppr.find(qn(t)) is not None), None)
        after.addprevious(ind) if after is not None else ppr.append(ind)
    for a in ("firstLine", "hanging", "start", "end"):
        ind.attrib.pop(qn("w:" + a), None)
    ind.set(qn("w:left"), str(max(0, int(left))))
    if hanging:
        ind.set(qn("w:hanging"), str(int(hanging)))
    if right is not None:
        ind.set(qn("w:right"), str(int(right)))


def _has_visible_border(tbl):
    from docx.oxml.ns import qn
    for b in tbl.iter(qn("w:tcBorders"), qn("w:tblBorders")):
        for e in b:
            if e.get(qn("w:val")) not in (None, "nil", "none"):
                return True
    return False


_FROM_BULLET_TABLE = []      # 표에서 문단으로 바꾼 글머리 문단(같은 글머리 이웃과 위치 맞춤용)


def _join_runs(p, runs):
    """p 끝에 runs 를 이어 붙인다 — 앞 글 끝 빈칸·뒤 글 앞 빈칸을 정리해 사이에 빈칸 딱 하나."""
    from docx.oxml.ns import qn
    ts = [t for t in p.iter(qn("w:t"))]
    if ts:
        ts[-1].text = (ts[-1].text or "").rstrip() + " "
        ts[-1].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    first = True
    for r in runs:
        if first:
            for t in r.findall(qn("w:t")):
                t.text = (t.text or "").lstrip()
                if t.text:
                    first = False
                    break
        p.append(r)


def align_table_bullets(d):
    """표에서 바꾼 글머리 문단을 바로 이웃한 같은 글머리 문단과 같은 들여쓰기로(원본은 ➢ 들이 한 줄로 서 있다).
    바꾼 문단 수."""
    from docx.oxml.ns import qn
    n = 0

    def bullet_of(p):
        t = _par_text(p).lstrip()
        return t[0] if t and t[0] in HANG_BULLETS else None
    for p in _FROM_BULLET_TABLE:
        if p.getparent() is None:
            continue
        b = bullet_of(p)
        if not b:
            continue
        ref = None
        for sib in list(p.itersiblings()) + list(p.itersiblings(preceding=True)):
            if sib.tag != qn("w:p") or sib in _FROM_BULLET_TABLE:
                continue
            if bullet_of(sib) == b:
                ref = sib
                break
            if bullet_of(sib):                           # 다른 글머리가 먼저 나오면 다른 묶음
                break
        if ref is None:
            continue
        l, fl, hg = _ind_vals(ref)
        _set_ind(p, l, hg, 0)
        rsp = ref.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
        psp = p.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
        if rsp is not None and psp is not None and rsp.get(qn("w:lineRule"), "auto") == "auto":
            # 이웃 글머리 문단들은 한 줄짜리라 줄 사이 = 앞 간격 + 줄 높이. 여러 줄로 합친 이 문단도 줄 사이를 그만큼
            try:
                pitch = int(rsp.get(qn("w:before"), "0")) + int(rsp.get(qn("w:line"), "240"))
                psp.set(qn("w:line"), str(max(240, pitch)))
                psp.set(qn("w:lineRule"), "auto")
            except ValueError:
                pass
        n += 1
    _FROM_BULLET_TABLE.clear()
    return n


def align_continuations(d):
    """글머리 문단 바로 뒤, 글머리 없이 이어지는 줄(pdf2docx 가 따로 뗀 둘째 줄)을 글머리 글 시작 자리에 맞춘다.
    ★내곡동 3쪽 '26 일 PF 리파이낸싱을 완료함.' 이 글 시작보다 오른쪽으로 밀려 있었다(원본은 글 시작과 같은 자리).
    글머리보다 오른쪽 ~ 글 시작+20pt 사이에 있는 줄만(그보다 더 들여 쓴 건 하위 항목이라 그대로). 맞춘 문단 수."""
    from docx.oxml.ns import qn
    n = 0
    for p in list(d.element.body.iter(qn("w:p"))):
        t = _par_text(p).lstrip()
        if not t or t[0] not in HANG_BULLETS:
            continue
        l, fl, hg = _ind_vals(p)
        if not hg:
            continue
        start, text_at = l - hg, l
        q = p.getnext()
        while q is not None and q.tag == qn("w:p"):
            qt = _par_text(q).strip()
            if not qt or qt[0] in HANG_BULLETS or q.find(f".//{qn('w:drawing')}") is not None or \
                    q.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
                break
            ql, qfl, qhg = _ind_vals(q)
            qs = ql + qfl - qhg
            if not (start + 60 < qs <= text_at + 400):
                break
            _set_ind(q, text_at, 0, 0)
            n += 1
            q = q.getnext()
    return n


def bullet_tables_to_paragraphs(d):
    """선 없는 표인데 줄마다 [글머리 하나][글] 두 칸인 것 → 보통 글머리 문단으로. 바로 뒤에 이어지는 줄(글머리 없이
    글 시작 자리에 들여 쓴 문단)은 그 문단에 이어 붙인다.
    ★pdf2docx 가 '➢ 본건은…' 첫 줄을 보이지 않는 표로, 둘째 줄부터를 따로 문단으로 만들어
    글머리가 혼자 한 줄을 차지하고 글이 아래로 밀렸다(내곡동 3쪽, 2026-10-14). 바꾼 표 수."""
    import copy
    from docx.oxml.ns import qn
    n = 0
    for tbl in list(d.element.body.iter(qn("w:tbl"))):
        if tbl.getparent() is None:
            continue
        rows = tbl.findall(qn("w:tr"))
        # 맨 위 글 없는 줄(제목 밑줄 등)은 표로 남기고 그 아래 글머리 줄만 문단으로
        keep = 0
        while keep < len(rows) and not "".join(x.text or "" for x in rows[keep].iter(qn("w:t"))).strip():
            keep += 1
        rows = rows[keep:]
        if not rows or any(_has_visible_border(tr) for tr in rows):
            continue
        ok = True
        for tr in rows:
            tcs = tr.findall(qn("w:tc"))
            if len(tcs) != 2 or tcs[0].find(".//" + qn("w:tbl")) is not None or tcs[1].find(".//" + qn("w:tbl")) is not None:
                ok = False
                break
            t0 = "".join(x.text or "" for x in tcs[0].iter(qn("w:t"))).strip()
            if len(t0) != 1 or t0 not in HANG_BULLETS:
                ok = False
                break
        if not ok:
            continue
        new_ps = []
        for tr in rows:
            tcs = tr.findall(qn("w:tc"))
            bullet_r = next(r for r in tcs[0].iter(qn("w:r")) if "".join(x.text or "" for x in r.iter(qn("w:t"))).strip())
            text_ps = [p for p in tcs[1] if p.tag == qn("w:p")]
            p = copy.deepcopy(text_ps[0]) if text_ps else tbl.makeelement(qn("w:p"), {})
            ppr = p.find(qn("w:pPr"))
            first_run = next((c for c in p if c.tag == qn("w:r")), None)
            br = copy.deepcopy(bullet_r)
            if first_run is not None:
                first_run.addprevious(br)
            else:
                p.append(br)
            for extra in text_ps[1:]:                      # 칸 안 여러 문단은 띄어쓰기 하나로 이어 붙임
                _join_runs(p, [copy.deepcopy(r) for r in extra.findall(qn("w:r"))])
            if ppr is not None:
                jc = ppr.find(qn("w:jc"))
                if jc is not None:
                    jc.set(qn("w:val"), "left")
                ind = ppr.find(qn("w:ind"))
                if ind is not None:
                    for a in list(ind.attrib):
                        ind.attrib.pop(a)
            new_ps.append(p)
        # 표 바로 뒤 이어지는 줄(글머리 없이 들여 쓴 문단)을 마지막 문단에 붙인다
        nxt = tbl.getnext()
        while nxt is not None and nxt.tag == qn("w:p"):
            txt = "".join(x.text or "" for x in nxt.iter(qn("w:t")))
            if not txt.strip() or txt.lstrip()[0] in HANG_BULLETS or nxt.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
                break
            l, fl, hg = _ind_vals(nxt)
            if l + fl - hg < 100:                           # 글 시작 자리로 들여 쓰지 않은 문단 = 새 글
                break
            _join_runs(new_ps[-1], nxt.findall(qn("w:r")))
            after = nxt.getnext()
            nxt.getparent().remove(nxt)
            nxt = after
        _FROM_BULLET_TABLE.extend(new_ps)
        if keep:
            anchor = tbl
            for p in new_ps:
                anchor.addnext(p)
                anchor = p
            for tr in rows:
                tbl.remove(tr)
        else:
            for p in new_ps:
                tbl.addprevious(p)
            tbl.getparent().remove(tbl)
        n += 1
    return n


def split_bullet_lines(d):
    """한 문단에 줄바꿈(w:br)으로 이어 붙은 줄 중 글머리로 시작하는 줄은 따로 문단으로 뗀다.
    ★pdf2docx 는 원본 여러 줄을 한 문단에 줄바꿈으로 넣는다 → 둘째 글머리 줄이 칸 왼쪽에 붙었다(내곡동 6쪽, 2026-10-14)."""
    import copy
    from docx.oxml.ns import qn
    n = 0
    for p in list(d.element.body.iter(qn("w:p"))):
        runs = [r for r in p if r.tag == qn("w:r")]
        cuts = []
        for i, r in enumerate(runs):
            brs = r.findall(qn("w:br"))
            if not brs or any(b.get(qn("w:type")) == "page" for b in brs):
                continue
            nxt = "".join(_par_text(x) for x in runs[i + 1:i + 3]).lstrip()
            if nxt[:1] and nxt[0] in HANG_BULLETS:
                cuts.append(i)
        for i in reversed(cuts):
            r = runs[i]
            newp = p.makeelement(qn("w:p"), {})
            ppr = p.find(qn("w:pPr"))
            if ppr is not None:
                c = copy.deepcopy(ppr)
                for sp in c.findall(qn("w:sectPr")):
                    c.remove(sp)
                newp.append(c)
            for x in runs[i + 1:]:
                newp.append(x)
            p.remove(r)
            p.addnext(newp)
            runs = runs[:i]
            n += 1
    return n


_NUM_START_RE = re.compile(r"^\s*(\(?\d{1,2}[.)]|\(?[가-하][.)]|[①-⑳]|[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩ]|※|\*|\[|<|【|■|□|▶|▷)")


def _run_w_pt(r):
    """글 조각 폭(pt) — Pretendard 글자 폭 표 + 글자 간격(w:spacing). 원본 글자 폭에 맞춘 뒤라 원본 폭과 거의 같다."""
    from docx.oxml.ns import qn
    rpr = r.find(qn("w:rPr"))
    sz = rpr.find(qn("w:sz")) if rpr is not None else None
    size = int(sz.get(qn("w:val"))) / 2 if sz is not None and (sz.get(qn("w:val")) or "").isdigit() else 10.0
    b = rpr.find(qn("w:b")) if rpr is not None else None
    tbl = _widths()["bold" if b is not None and b.get(qn("w:val")) not in ("0", "false") else "light"]
    spc = rpr.find(qn("w:spacing")) if rpr is not None else None
    extra = int(spc.get(qn("w:val"))) / 20 if spc is not None and (spc.get(qn("w:val")) or "").lstrip("-").isdigit() else 0.0
    txt = "".join(t.text or "" for t in r.findall(qn("w:t")))
    return sum(_char_w(ch, tbl) * size + extra for ch in txt)


def _par_size(p):
    from docx.oxml.ns import qn
    v = [int(x.get(qn("w:val"))) / 2 for x in p.iter(qn("w:sz")) if (x.get(qn("w:val")) or "").isdigit()
         and x.getparent().tag == qn("w:rPr") and x.getparent().getparent().find(qn("w:t")) is not None]
    return max(v) if v else 10.0


def _line_h(p):
    """(앞 간격 pt, 한 줄 높이 pt, 줄 규칙, line 값) — match_line_spacing 뒤라 '배수' 는 Pretendard 기준."""
    from docx.oxml.ns import qn
    sp = p.find(f"{qn('w:pPr')}/{qn('w:spacing')}")

    def g(a):
        v = sp.get(qn("w:" + a)) if sp is not None else None
        try:
            return int(float(v)) if v is not None else None
        except ValueError:
            return None
    before = (g("before") or 0) / 20
    line, rule = g("line"), (sp.get(qn("w:lineRule")) if sp is not None else None) or "auto"
    size = _par_size(p)
    if not line:
        return before, size * LINE_EM, rule, line
    if rule == "auto":
        return before, line / 240 * LINE_EM * size, rule, line
    return before, line / 20, rule, line


def _set_pitch(p, before_pt, pitch_pt):
    """줄 간격 = '최소' pitch_pt(원본 줄 사이 거리), 앞 간격 = before_pt. '최소' 라 글자가 커져도 잘리지 않는다."""
    from docx.oxml.ns import qn
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        ppr = p.makeelement(qn("w:pPr"), {})
        p.insert(0, ppr)
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = ppr.makeelement(qn("w:spacing"), {})
        after = next((ppr.find(qn(t)) for t in ("w:ind", "w:jc", "w:rPr", "w:sectPr") if ppr.find(qn(t)) is not None), None)
        after.addprevious(sp) if after is not None else ppr.append(sp)
    sp.set(qn("w:before"), str(max(0, int(round(before_pt * 20)))))
    sp.set(qn("w:line"), str(max(20, int(round(pitch_pt * 20)))))
    sp.set(qn("w:lineRule"), "atLeast")


def _set_jc(p, val):
    from docx.oxml.ns import qn
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return
    jc = ppr.find(qn("w:jc"))
    if jc is None:
        jc = ppr.makeelement(qn("w:jc"), {})
        after = next((ppr.find(qn(t)) for t in ("w:rPr", "w:sectPr") if ppr.find(qn(t)) is not None), None)
        after.addprevious(jc) if after is not None else ppr.append(jc)
    jc.set(qn("w:val"), val)


def normalize_text_flow(d):
    """표 밖 본문 글을 '원본 줄 모양 흉내' 에서 '이어지는 글' 로 — 우리 양식 폭에서 워드가 자연스럽게 줄을 바꾸게.
    ★2026-10-21 사용자(이천 Executive summary): ① 3줄 문단 첫 줄만 시작점이 다름 ② 글머리 줄 간격이 들쭉날쭉
      ③ '…이며, 해당        대출 실행시' 처럼 크게 벌어짐(군데군데 많음).
    원인: pdf2docx 가 원본 줄마다 ⓐ 따로 문단을 만들고(꽉 찬 첫 줄은 '가운데 정렬' 로 잘못 봄) ⓑ 줄 바뀐 자리에 탭을 넣고
      ⓒ 줄 간격을 '앞 간격+줄 높이' 로 흉내 냄 → 양식 폭(516pt)이 원본(503pt)과 다르고 글꼴도 달라 줄 바뀌는 자리가 달라지면 다 어긋났다.
      (원본과 같은 폭으로 해도 글꼴이 바뀌면 줄 바뀌는 자리가 달라져 해결 안 됨 → 글을 이어 붙이는 게 근본.)
    ① 줄 바뀐 자리 탭 지움: 탭 앞 글이 거의 한 줄(줄 폭 70% 이상)이고 뒤에 글이 있을 때만(목차·칸 맞춤 탭은 앞 글이 짧아 안 건드림 —
       예전 '모든 탭 → 띄어쓰기' 는 내곡 목차를 망가뜨렸다).
    ② 쪼개진 줄 합치기: 윗줄이 거의 꽉 찼고(80% 이상) 아랫줄이 글머리·번호로 시작하지 않으며 크기·굵기·간격·시작점이 같을 때.
       원본 줄 끝 띄어쓰기는 pdf2docx 가 남겨 둔다(어절 중간 줄바꿈 '조합|원' 은 띄어쓰기 없음) → 그대로 붙인다. 합친 문단은 양쪽 맞춤.
    ③ 줄 간격 같게: 합친 문단·이어지는 글머리 묶음은 모든 줄이 원본 줄 사이 거리(앞 간격+줄 높이)로 — 첫 줄 자리는 그대로.
    고친 곳 수."""
    from docx.oxml.ns import qn
    body = d.element.body
    n = 0
    # 문단 → 그 구역 본문 폭(pt)
    width_of, pend = {}, []
    for el in body:
        sp = None
        if el.tag == qn("w:p"):
            pend.append(el)
            sp = el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
        elif el.tag == qn("w:sectPr"):
            sp = el
        if sp is not None:
            sz_, mg = sp.find(qn("w:pgSz")), sp.find(qn("w:pgMar"))
            try:
                w = (int(sz_.get(qn("w:w"))) - int(mg.get(qn("w:left"))) - int(mg.get(qn("w:right")))) / 20
            except (TypeError, ValueError, AttributeError):
                w = 480.0
            for q in pend:
                width_of[q] = w
            pend = []
    for q in pend:
        width_of[q] = 480.0

    def plain(p):
        return p is not None and p.tag == qn("w:p") and p.getparent() is body and \
            p.find(".//" + qn("w:drawing")) is None and p.find(".//" + qn("w:pict")) is None and \
            p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is None and \
            not any(b.get(qn("w:type")) == "page" for b in p.iter(qn("w:br")))

    def starts(p):                                          # (첫 줄 시작, 둘째 줄부터 시작) twips
        l, fl, hg = _ind_vals(p)
        return l + fl - hg, l

    def lead_bullet(p):
        """(글머리 글자 또는 None, [글머리][탭] 의 그 탭 run 또는 None)"""
        runs = [r for r in p if r.tag == qn("w:r")]
        for i, r in enumerate(runs):
            t = "".join(x.text or "" for x in r.findall(qn("w:t"))).strip()
            if not t and r.find(qn("w:tab")) is None:
                continue
            if t and t[0] in HANG_BULLETS and len(t) == 1 and i + 1 < len(runs) and runs[i + 1].find(qn("w:tab")) is not None:
                return t, runs[i + 1]
            return (t[0] if t and t[0] in HANG_BULLETS else None), None
        return None, None

    def bold(x):
        for r in x.iter(qn("w:r")):
            if _par_text(r).strip():
                rb = r.find(f"{qn('w:rPr')}/{qn('w:b')}")
                return rb is not None and rb.get(qn("w:val")) not in ("0", "false")
        return False

    # ① 줄 바뀐 자리 탭
    for p in list(body.findall(qn("w:p"))):
        if not plain(p):
            continue
        W = width_of.get(p, 480.0)
        b, btab = lead_bullet(p)
        ppr = p.find(qn("w:pPr"))
        tabs = ppr.find(qn("w:tabs")) if ppr is not None else None
        stops = [int(t.get(qn("w:pos"))) for t in (tabs.findall(qn("w:tab")) if tabs is not None else [])
                 if (t.get(qn("w:pos")) or "").lstrip("-").isdigit()]
        first, cont = starts(p)
        if btab is not None:
            text_at = max(stops) if stops and max(stops) > first else cont
        else:
            text_at = first
        acc, avail = 0.0, W - text_at / 20 - 2
        runs = [r for r in p if r.tag == qn("w:r")]
        seen_after = btab is None
        changed = False
        for i, r in enumerate(runs):
            if not seen_after:
                if r is btab:
                    seen_after = True
                continue
            if r.find(qn("w:br")) is not None:
                acc, avail = 0.0, W - cont / 20 - 2
                continue
            if r.find(qn("w:tab")) is not None and not _par_text(r).strip():
                rest = "".join(_par_text(x) for x in runs[i + 1:]).strip()
                if rest and acc >= 0.7 * avail:
                    for tb in r.findall(qn("w:tab")):
                        r.remove(tb)
                    changed = True
                    avail = W - cont / 20 - 2
                acc = 0.0
                continue
            acc += _run_w_pt(r)
        if changed:
            n += 1
            if btab is not None and stops and max(stops) > first and tabs is not None:
                pos = max(stops)                            # [글머리][탭] + 탭 자리 → 둘째 줄을 글 시작에(내어쓰기)
                _set_ind(p, pos, pos - first)
                ppr.remove(tabs)

    # ② 쪼개진 줄 합치기
    merged = []
    p = body[0] if len(body) else None
    while p is not None:
        q = p.getnext()
        if not (plain(p) and plain(q)):
            p = q
            continue
        t1, t2 = _par_text(p), _par_text(q)
        ok = bool(t1.strip()) and bool(t2.strip()) and t2.lstrip()[:1] not in HANG_BULLETS and \
            not _NUM_START_RE.match(t2) and not re.search(r"([.:。]|[다임음함됨]\.?)\s*$", t1.rstrip()) and \
            not any(b_.get(qn("w:type")) in (None, "textWrapping") for b_ in p.iter(qn("w:br")))
        if ok:
            b1, h1, r1, l1 = _line_h(p)
            b2, h2, r2, l2 = _line_h(q)
            f1, c1 = starts(p)
            f2, _ = starts(q)
            W = width_of.get(p, 480.0)
            _, lbt = lead_bullet(p)
            # 윗줄의 '마지막 줄' 이 얼마나 찼나(이미 합친 문단이면 전체 길이를 줄 폭으로 나눈 나머지)
            used = sum(_run_w_pt(r) for r in p if r.tag == qn("w:r") and r is not lbt)
            first_avail = W - ((c1 if lbt is not None else max(f1, 0)) / 20)
            cont_avail = W - c1 / 20
            if used > first_avail:
                last = (used - first_avail) % max(cont_avail, 1)
                fill = last / max(cont_avail, 1)
            else:
                fill = used / max(first_avail, 1)
            jc = p.find(f"{qn('w:pPr')}/{qn('w:jc')}")
            ok = abs(_par_size(p) - _par_size(q)) < 0.6 and bold(p) == bold(q) and r1 == r2 and \
                abs((l1 or 0) - (l2 or 0)) <= max(12, 0.1 * (l1 or 240)) and b2 <= 14 and \
                abs(f2 - c1) <= 120 and fill >= 0.8 and \
                (jc is None or jc.get(qn("w:val")) in ("left", "both", "center", "start", "distribute"))
        if not ok:
            p = q
            continue
        pitch = b2 + h2                                     # 원본 줄 사이 거리
        lead = True
        for ch in [c for c in q if c.tag != qn("w:pPr")]:
            if lead and ch.tag == qn("w:r") and ch.find(qn("w:tab")) is not None and not _par_text(ch).strip():
                continue                                    # 아랫줄 맨 앞 탭(들여쓰기 흉내)은 버림
            lead = False
            p.append(ch)
        q.getparent().remove(q)
        if p.get("_pitch_b") is None:
            p.set("_pitch_b", str(b1 + h1))                 # 첫 줄 아래 끝(앞 간격+줄 높이)은 그대로 지킨다
            merged.append(p)
        p.set("_pitch", str(max(pitch, float(p.get("_pitch") or 0))))
        _set_jc(p, "both")
        n += 1                                              # p 를 그대로 두고 다음 줄(3줄 이상)도 이어 본다
    for p in merged:
        pitch, top = float(p.get("_pitch")), float(p.get("_pitch_b"))
        _set_pitch(p, top - pitch, pitch)
        for a in ("_pitch", "_pitch_b"):
            p.attrib.pop(a, None)

    # ③ 이어지는 글머리 묶음 줄 간격 같게(같은 글머리·같은 시작점, 사이에 다른 것 없을 때)
    p = body[0] if len(body) else None
    while p is not None:
        b = lead_bullet(p)[0] if plain(p) else None
        if b is None:
            p = p.getnext()
            continue
        grp, q = [p], p.getnext()
        while plain(q) and lead_bullet(q)[0] == b and abs(starts(q)[0] - starts(p)[0]) <= 60:
            grp.append(q)
            q = q.getnext()
        # ★줄 사이 거리는 '원본에서 한 줄짜리' 글머리로만 잰다 — 여러 줄 문단 사이의 '앞 간격' 은 문단 간격이지 줄 간격이 아니다
        #   (돈암 하이라이트: 글머리마다 여러 줄 문단 → 줄 간격이 두 배로 벌어져 한 쪽이 넘쳤다, 2026-10-21)
        def one_line(x):
            _, bt = lead_bullet(x)
            f_, c_ = starts(x)
            ppr_ = x.find(qn("w:pPr"))
            tb_ = ppr_.find(qn("w:tabs")) if ppr_ is not None else None
            st_ = [int(t.get(qn("w:pos"))) for t in (tb_.findall(qn("w:tab")) if tb_ is not None else [])
                   if (t.get(qn("w:pos")) or "").lstrip("-").isdigit()]
            at = max(st_ + [c_]) if bt is not None else f_
            w_ = sum(_run_w_pt(r) for r in x if r.tag == qn("w:r") and r is not bt and r.find(qn("w:tab")) is None)
            return w_ <= width_of.get(x, 480.0) - at / 20 - 2
        singles = [x for x in grp[1:] if one_line(x)]
        if len(grp) >= 2 and singles:
            pitches = sorted(sum(_line_h(x)[:2]) for x in singles)
            pitch = pitches[len(pitches) // 2]
            if pitch >= max(_par_size(x) for x in grp) * 1.2:
                for i, x in enumerate(grp):
                    bx, hx, _, _ = _line_h(x)
                    _set_pitch(x, (bx + hx - pitch) if i == 0 else 0.0, pitch)
                n += 1
        p = q
    return n


def hang_bullets(d, ratios):
    """글머리 문단 = 글머리 + 탭 + 글, '내어쓰기' 로 둘째 줄부터 글 시작에 맞춘다(원본 모양).
    ★예전엔 글머리 뒤를 글자 간격(w:spacing)으로 벌려 첫 줄 간격이 크게 벌어지고(내곡동 3쪽 '➢ 본건은'),
    둘째 줄·'(*)' 줄은 칸 왼쪽에 붙었다(내곡동 6쪽). 글머리~글 거리는 원본에서 잰 값(글자 크기 기준)."""
    import copy
    from docx.oxml.ns import qn
    em = (ratios or {}).get("글자", {})
    n = 0
    for p in d.element.body.iter(qn("w:p")):
        runs = [r for r in p if r.tag == qn("w:r") and r.find(qn("w:t")) is not None]
        first = next((r for r in runs if _par_text(r).strip()), None)
        if first is None:
            continue
        txt = _par_text(first).lstrip()
        b = txt[0]
        if b not in HANG_BULLETS:
            continue
        nx = first.getnext()
        if txt.strip() == b and nx is not None and nx.tag == qn("w:r") and nx.find(qn("w:tab")) is not None:
            for r_ in (first, nx):                             # 이미 [글머리][탭][글] (원본 선으로 다시 짠 표 — 원본 자리 그대로)
                rp = r_.find(qn("w:rPr"))                      # 글머리 글자 벌림이 남으면 탭을 넘어 글이 밀린다
                for sp_ in (rp.findall(qn("w:spacing")) if rp is not None else []):
                    rp.remove(sp_)
            continue
        if b in "-·∙ㆍ" and len(_par_text(p).strip()) < 3:
            continue
        ts = first.findall(qn("w:t"))
        rest = (ts[0].text or "").lstrip()[1:]
        ts[0].text = b
        for t in ts[1:]:
            rest += t.text or ""
            first.remove(t)
        rpr = first.find(qn("w:rPr"))
        if rpr is not None:
            for sp in rpr.findall(qn("w:spacing")):
                rpr.remove(sp)
        tab_r = first.makeelement(qn("w:r"), {})
        if rpr is not None:
            tab_r.append(copy.deepcopy(rpr))
        tab_r.append(tab_r.makeelement(qn("w:tab"), {}))
        first.addnext(tab_r)
        rest = rest.lstrip()
        if rest:
            rr = copy.deepcopy(first)
            for t in rr.findall(qn("w:t")):
                rr.remove(t)
            t = rr.makeelement(qn("w:t"), {"{http://www.w3.org/XML/1998/namespace}space": "preserve"})
            t.text = rest
            rr.append(t)
            tab_r.addnext(rr)
        else:
            for t in p.iter(qn("w:t")):
                if t.getparent() in (first, tab_r):
                    continue
                t.text = (t.text or "").lstrip()
                if t.text:
                    break
        gap_em = em.get(("글머리", b)) or 1.3
        # ★글머리 글자보다 좁으면 탭이 다음 기본 탭(36pt)까지 튀어 글머리 뒤가 크게 벌어졌다(고산 '▪' = 0.86em, 2026-10-15)
        gap_em = max(gap_em, BULLET_EM.get(b, 0.9) + 0.3)
        gap = int(round(min(max(gap_em, 0.8), 2.5) * 9 * 20))
        left, fl, hg = _ind_vals(p)
        _set_ind(p, left + fl - hg + gap, gap)
        ppr = p.find(qn("w:pPr"))
        tabs = ppr.find(qn("w:tabs")) if ppr is not None else None
        if tabs is not None:
            ppr.remove(tabs)
        n += 1
    return n


def even_cell_margins(d):
    """칸 안 왼쪽·오른쪽 여백을 같게(85 twips ≈0.15cm). 가운데·오른쪽 정렬 문단은 들여쓰기 0.
    왼쪽 정렬 문단은 칸 안 가장 왼쪽 문단을 0 으로 맞추고 상대 위치(들여 쓴 항목·내어쓰기)는 그대로.
    ★pdf2docx 는 칸 여백 0 + 문단 들여쓰기로 글 위치를 흉내 내 정렬이 원본과 어긋났다(2026-10-14 사용자 지적)."""
    from docx.oxml.ns import qn
    n = 0
    for tc in d.element.body.iter(qn("w:tc")):
        ps = [p for p in tc if p.tag == qn("w:p")]
        tcpr = tc.find(qn("w:tcPr"))
        if not ps or tcpr is None:
            continue
        mar = tcpr.find(qn("w:tcMar"))
        if mar is None:
            mar = tcpr.makeelement(qn("w:tcMar"), {})
            after = next((tcpr.find(qn(t)) for t in ("w:textDirection", "w:tcFitText", "w:vAlign", "w:hideMark")
                          if tcpr.find(qn(t)) is not None), None)
            after.addprevious(mar) if after is not None else tcpr.append(mar)
        for e in list(mar):
            if e.tag in (qn("w:start"), qn("w:end"), qn("w:left"), qn("w:right")):
                mar.remove(e)
        # ★좁은 칸은 여백도 좁게(칸 폭의 6%, 최대 85) — 좁은 숫자 칸에서 글자가 한 자씩 꺾여 표가 몇 배로 길어졌다(효자 분양사례, 2026-10-15)
        tw_ = tcpr.find(qn("w:tcW"))
        try:
            cw = int(float(tw_.get(qn("w:w")))) if tw_ is not None and tw_.get(qn("w:type")) in (None, "dxa") else 0
        except (TypeError, ValueError):
            cw = 0
        M = min(85, max(10, int(cw * 0.06))) if cw else 85
        top = mar.find(qn("w:top"))
        lft = mar.makeelement(qn("w:left"), {qn("w:w"): str(M), qn("w:type"): "dxa"})
        top.addnext(lft) if top is not None else mar.insert(0, lft)
        bot = mar.find(qn("w:bottom"))
        rgt = mar.makeelement(qn("w:right"), {qn("w:w"): str(M), qn("w:type"): "dxa"})
        bot.addnext(rgt) if bot is not None else mar.append(rgt)

        def jcv(p):
            jc = p.find(f"{qn('w:pPr')}/{qn('w:jc')}")
            return jc.get(qn("w:val")) if jc is not None else "left"
        lefts = []
        for p in ps:
            if jcv(p) in ("center", "right", "end"):
                _set_ind(p, 0, 0, 0)
            elif _par_text(p).strip():
                l, fl, hg = _ind_vals(p)
                lefts.append(l + fl - hg)
        base = min(lefts) if lefts else 0
        for p in ps:
            if jcv(p) in ("center", "right", "end"):
                continue
            l, fl, hg = _ind_vals(p)
            if hg:
                _set_ind(p, l - base, hg, 0)
            else:
                _set_ind(p, max(0, l + fl - base), 0, 0)
        n += 1
    return n


def split_body_breaks(d):
    """표 밖 문단: pdf2docx 가 [제목 줄][줄바꿈][탭탭 본문 줄][탭 다음 줄]… 로 한 문단에 넣은 것을 원본 줄대로 고친다.
    ① 줄바꿈 자리마다 문단을 나눔(제목 '■ 도시계획변경 사전협상제도 개요' 와 본문이 한 문단이었다)
    ② 새 문단 맨 앞 탭은 지움(탭 두 개로 첫 줄이 크게 들여 써졌다), 들여쓰기는 원래 문단 글머리 자리
    ③ 줄 가운데 '긴 글 뒤 탭' = 원본 줄이 바뀐 자리 → 줄바꿈('…2023      년' 처럼 벌어졌다). 고산 18쪽, 2026-10-15."""
    import copy
    from docx.oxml.ns import qn
    n = 0
    body = d.element.body
    for p in list(body.findall(qn("w:p"))):
        if p.find(f".//{qn('w:drawing')}") is not None:
            continue
        brs = [r for r in p.findall(qn("w:r")) if any(b.get(qn("w:type")) in (None, "textWrapping") for b in r.findall(qn("w:br")))]
        # ★[줄바꿈] 바로 뒤가 [탭] 으로 시작하는 문단(고산 18쪽 모양)만 — 다른 문단(내곡 하이라이트 글머리 등)은 건드리면 망가졌다
        def _tab_after(r):
            n_ = r.getnext()
            return n_ is not None and n_.tag == qn("w:r") and n_.find(qn("w:tab")) is not None and                 not "".join(t.text or "" for t in n_.findall(qn("w:t"))).strip()
        if not any(_tab_after(r) for r in brs):
            continue
        if _par_text(p).lstrip()[:1] in HANG_BULLETS and _par_text(p).lstrip()[:1] != "■":
            continue
        tabs_mid = []
        acc = ""
        for r in p.findall(qn("w:r")):
            if r.find(qn("w:br")) is not None:
                acc = ""
            if r.find(qn("w:tab")) is not None and len(acc.strip()) >= 25:
                tabs_mid.append(r)
                acc = ""
            acc += "".join(t.text or "" for t in r.findall(qn("w:t")))
        if not brs and not tabs_mid:
            continue
        for r in tabs_mid:                                     # ③ 줄 바뀐 자리 탭 → 줄바꿈
            for tb in r.findall(qn("w:tab")):
                r.replace(tb, r.makeelement(qn("w:br"), {}))
        if not brs:
            n += 1
            continue
        l, fl, hg = _ind_vals(p)
        base_left = max(0, l + fl - hg)
        ppr = p.find(qn("w:pPr"))
        segs, cur = [], []
        for ch in [c for c in p if c.tag != qn("w:pPr")]:
            if ch in brs:
                segs.append(cur)
                cur = []
                continue
            cur.append(ch)
        segs.append(cur)
        segs = [sg for sg in segs if "".join(t.text or "" for x in sg for t in x.iter(qn("w:t"))).strip()]
        if len(segs) < 2:
            continue
        for ch in [c for c in p if c.tag != qn("w:pPr")]:
            p.remove(ch)
        for x in segs[0]:
            p.append(x)
        prev = p
        for sg in segs[1:]:
            np_ = p.makeelement(qn("w:p"), {})
            if ppr is not None:
                c = copy.deepcopy(ppr)
                for sp in c.findall(qn("w:sectPr")):
                    c.remove(sp)
                np_.append(c)
            while sg and sg[0].tag == qn("w:r") and sg[0].find(qn("w:tab")) is not None and                     not "".join(t.text or "" for t in sg[0].findall(qn("w:t"))):
                sg = sg[1:]                                    # ② 맨 앞 탭 지움
            for x in sg:
                np_.append(x)
            _set_ind(np_, base_left, 0, 0)
            sp_ = np_.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
            if sp_ is not None:
                sp_.set(qn("w:before"), "0")
            if not np_.find(qn("w:pPr")) is None and _par_text(np_).lstrip()[:1] in HANG_BULLETS:
                sp_b = np_.find(f"{qn('w:pPr')}/{qn('w:spacing')}")   # 글머리 소제목('■ 고산2지구…')은 위와 원본만큼 띄움(빈 줄 하나)
                if sp_b is not None:
                    sp_b.set(qn("w:before"), "200")
            prev.addnext(np_)
            prev = np_
            n += 1
        # ★원본 '앞 간격' 은 첫 줄(제목) 위 간격 — 그대로 두면 나눈 뒤 첫 문단 위에 두 줄 넘는 빈칸이 생겼다('4) 대관 협의 역량', 고산 19쪽)
        sp0 = p.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
        if sp0 is not None and int(float(sp0.get(qn("w:before")) or 0)) > 300:
            sp0.set(qn("w:before"), "200")
        if p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:   # 구역 끝 표시는 마지막 문단으로
            sect = p.find(f"{qn('w:pPr')}/{qn('w:sectPr')}")
            prev.find(qn("w:pPr")).append(sect)
    return n


def match_char_width(d, ratios):
    """Pretendard 글자 사이를 벌려(좁혀) 원본과 같은 폭으로 → 줄이 원본 자리에서 바뀐다. 고친 글 조각 수.
    ★Pretendard 는 원본 글꼴보다 약 9% 좁다(샘플 17개 실측 0.85~0.97). 그대로 두면 줄바꿈 자리가 달라져
    목차 점선·쪽번호가 다음 줄로 밀렸다(내곡동, 2026-10-12). 사용자: '글꼴 바꿔도 간격을 늘려 똑같이 하면 되잖아'."""
    from docx.oxml.ns import qn
    W = _widths()
    n = 0
    _split_bullet_runs(d)
    for r in d.element.body.iter(qn("w:r")):
        text = "".join(t.text or "" for t in r.findall(qn("w:t")))
        chars = [c for c in text]
        if not chars:
            continue
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            continue
        if any(_SYMBOL_FONT_RE.search(v or "") for v in (rpr.find(qn("w:rFonts")).attrib.values()
                                                         if rpr.find(qn("w:rFonts")) is not None else [])):
            continue
        b = rpr.find(qn("w:b"))
        bold = b is not None and b.get(qn("w:val")) not in ("0", "false")
        rf = rpr.find(qn("w:rFonts"))
        name = _font_key((rf.get(qn("w:ascii")) or rf.get(qn("w:eastAsia"))) if rf is not None else "")
        ratio = ratios.get((name, bold)) or ratios.get((name, not bold)) or ratios.get("*", 1.0)
        ratio = min(max(ratio, 0.85), 1.25)
        szel = rpr.find(qn("w:sz"))
        size = int(szel.get(qn("w:val"))) / 2 if szel is not None and (szel.get(qn("w:val")) or "").isdigit() else 10.0
        tbl = W["bold" if bold else "light"]
        pret = sum(size * _char_w(c, tbl) for c in chars)
        em = ratios.get("글자", {})
        orig = 0.0
        bullet = text.strip() if len(text.strip()) == 1 and text.strip() in BULLETS else None
        par = next((a for a in r.iterancestors() if a.tag == qn("w:p")), None)
        first = par is not None and next((x for x in par.iter(qn("w:t")) if (x.text or "").strip()), None) is \
            next((x for x in r.iter(qn("w:t")) if (x.text or "").strip()), None)
        if bullet and not first:                              # 줄바꿈·탭 바로 뒤의 글머리도 '줄 맨 앞'
            prev = r.getprevious()
            while prev is not None and prev.tag == qn("w:r") and \
                    not "".join(t.text or "" for t in prev.findall(qn("w:t"))).strip() and \
                    prev.find(qn("w:tab")) is None and prev.find(qn("w:br")) is None:
                prev = prev.getprevious()
            first = prev is not None and prev.tag == qn("w:r") and (
                prev.find(qn("w:tab")) is not None or prev.find(qn("w:br")) is not None or
                (prev[-1].tag in (qn("w:tab"), qn("w:br")) if len(prev) else False))
        if bullet and first and em.get(("글머리", bullet)):
            orig = size * em[("글머리", bullet)]                 # 글머리 기호 + 원래 띄운 거리
            chars = [bullet] + [c for c in text if c != bullet]
        else:
            for c in chars:                                   # 글자마다 원본 실제 폭 — 없으면 그 글꼴 평균 비율로
                o = em.get((name, bold, c)) or em.get((name, not bold, c))
                orig += size * o if o else size * _char_w(c, tbl) * ratio
            # 원본에서 칸·줄 폭을 꽉 채운 줄은 워드에서 아주 조금만 길어도 다음 줄로 넘어간다(내곡동 대출기간)
            # → 원본 폭의 WIDTH_FIT 로 맞춰 살짝 여유를 둔다
            orig *= WIDTH_FIT
        extra = (orig - pret) / len(chars)                    # 글자 하나당 더할 폭(pt)
        extra = max(-0.15 * size, min((2.0 if bullet and first else 0.3) * size, extra))
        tw = int(round(extra * 20))
        if not tw:
            continue
        sp = rpr.find(qn("w:spacing"))
        if sp is None:
            sp = rpr.makeelement(qn("w:spacing"), {})
            # rPr 안 순서(워드 규칙): … color, spacing, w, kern, position, sz, szCs, highlight, u, … — 어기면 '손상'
            after = next((rpr.find(qn(t)) for t in ("w:w", "w:kern", "w:position", "w:sz", "w:szCs", "w:highlight",
                                                     "w:u", "w:effect", "w:bdr", "w:shd", "w:fitText", "w:vertAlign",
                                                     "w:rtl", "w:cs", "w:em", "w:lang", "w:eastAsianLayout")
                          if rpr.find(qn(t)) is not None), None)
            after.addprevious(sp) if after is not None else rpr.append(sp)
        sp.set(qn("w:val"), str(tw + int(sp.get(qn("w:val")) or 0) if (sp.get(qn("w:val")) or "").lstrip("-").isdigit() else tw))
        n += 1
    return n


def fill_white_text_cells(d):
    """글자가 전부 흰색인데 바탕색이 빠진 칸 → 진한 바탕(나중에 머리글 칠하기가 회사 네이비로 바꾼다).
    ★내곡동 '구 분|내용' 머리글: 원본은 둘 다 남색+흰 글씨인데 '내용' 칸 색이 빠져 흰 글씨가 안 보였다(2026-10-12)."""
    from docx.oxml.ns import qn
    n = 0
    for tc in d.element.body.iter(qn("w:tc")):
        runs = [r for r in tc.iter(qn("w:r")) if "".join(t.text or "" for t in r.iter(qn("w:t"))).strip()]
        if not runs or any(tbl is not None for tbl in [tc.find(".//" + qn("w:tbl"))]):
            continue
        cols = [(r.find(f"{qn('w:rPr')}/{qn('w:color')}").get(qn("w:val")) or "").upper()
                if r.find(f"{qn('w:rPr')}/{qn('w:color')}") is not None else "" for r in runs]
        if not all(c == "FFFFFF" for c in cols):
            continue
        tcpr = tc.find(qn("w:tcPr"))
        shd = tcpr.find(qn("w:shd")) if tcpr is not None else None
        if shd is not None and (shd.get(qn("w:fill")) or "").upper() not in ("", "AUTO", "FFFFFF"):
            continue
        if tcpr is None:
            tcpr = tc.makeelement(qn("w:tcPr"), {})
            tc.insert(0, tcpr)
        if shd is None:
            shd = tcpr.makeelement(qn("w:shd"), {})
            # shd 는 tcPr 안에서 noWrap/tcMar/textDirection/tcFitText/vAlign/hideMark 보다 앞
            after = next((tcpr.find(qn(t)) for t in ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
                                                      "w:vAlign", "w:hideMark") if tcpr.find(qn(t)) is not None), None)
            after.addprevious(shd) if after is not None else tcpr.append(shd)
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), _table_dark_fill(tc) or "243F60")
        n += 1
    _fill_vmerge_continues(d)
    _fill_gaps_in_header_rows(d)
    return n


def _fill_gaps_in_header_rows(d):
    """머리글 줄에서 양옆이 같은 진한 색인데 혼자 비어 흰 칸 → 같은 색(이천 대월 '계약금' 아래 칸)."""
    import copy
    from docx.oxml.ns import qn

    def fill_of(tc):
        shd = tc.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
        f = (shd.get(qn("w:fill")) or "").upper() if shd is not None else ""
        return f if f and f != "AUTO" and _lum_hex(f) < 0.5 else None
    for tr in d.element.body.iter(qn("w:tr")):
        tcs = tr.findall(qn("w:tc"))
        for i in range(1, len(tcs) - 1):
            tc = tcs[i]
            if "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip() or tc.find(f"{qn('w:tcPr')}/{qn('w:shd')}") is not None:
                continue
            a, b = fill_of(tcs[i - 1]), fill_of(tcs[i + 1])
            if a and a == b:
                pr = tc.find(qn("w:tcPr"))
                sh = copy.deepcopy(tcs[i - 1].find(f"{qn('w:tcPr')}/{qn('w:shd')}"))
                after = next((pr.find(qn(t)) for t in ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
                                                        "w:vAlign", "w:hideMark") if pr.find(qn(t)) is not None), None)
                after.addprevious(sh) if after is not None else pr.append(sh)


def _lum_hex(h):
    try:
        r, g, b = (int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
        return 0.299 * r + 0.587 * g + 0.114 * b
    except Exception:
        return 1.0


def _table_dark_fill(tc):
    """칸이 든 표에서 가장 많이 쓴 진한 바탕색 — 빠진 머리글 색을 원본 색으로 채우려고(이천 대월: 검정)."""
    from collections import Counter
    from docx.oxml.ns import qn
    tbl = next((a for a in tc.iterancestors() if a.tag == qn("w:tbl")), None)
    if tbl is None:
        return None
    cnt = Counter()
    for c in tbl.iter(qn("w:tc")):
        if next((a for a in c.iterancestors() if a.tag == qn("w:tbl")), None) is not tbl:
            continue
        shd = c.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
        f = (shd.get(qn("w:fill")) or "").upper() if shd is not None else ""
        if f and f not in ("AUTO",) and _lum_hex(f) < 0.5:
            cnt[f] += 1
    return cnt.most_common(1)[0][0] if cnt else None


def _fill_vmerge_continues(d):
    """세로로 합친 칸의 아래 칸(이어짐)에 바탕색이 없으면 맨 위 칸 색을 그대로(합친 칸 한쪽만 흰색이 되던 것)."""
    import copy
    from docx.oxml.ns import qn
    for tbl in d.element.body.iter(qn("w:tbl")):
        above = {}                                         # 눈금 시작 → 위 줄 칸
        for tr in tbl.findall(qn("w:tr")):
            pos, row = 0, {}
            for tc in tr.findall(qn("w:tc")):
                pr = tc.find(qn("w:tcPr"))
                gs = pr.find(qn("w:gridSpan")) if pr is not None else None
                vm = pr.find(qn("w:vMerge")) if pr is not None else None
                if vm is not None and vm.get(qn("w:val")) in (None, "continue") and pos in above:
                    src = above[pos].find(f"{qn('w:tcPr')}/{qn('w:shd')}")
                    if src is not None and pr.find(qn("w:shd")) is None:
                        sh = copy.deepcopy(src)
                        after = next((pr.find(qn(t)) for t in ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
                                                                "w:vAlign", "w:hideMark") if pr.find(qn(t)) is not None), None)
                        after.addprevious(sh) if after is not None else pr.append(sh)
                row[pos] = tc
                pos += int(gs.get(qn("w:val"))) if gs is not None else 1
            above = row


def _hex(rgb):
    return "%02X%02X%02X" % tuple(max(0, min(255, int(round(c * 255)))) for c in rgb[:3])


def rebuild_table_from_pdf(page, region, W_scale=1.0):
    """원본 PDF 의 선(얇은 네모)·칠·글로 그 자리 표를 워드 표(w:tbl 요소)로 다시 짠다. 실패하면 None.
    ★pdf2docx 가 '표 안의 표' 를 칸 속 작은 표 조각들로 망가뜨린 경우(방배동 금융조건, 2026-10-13)에만 쓴다.
    안쪽 표까지 한 장의 표(칸 합치기)로 평평하게 만든다 — 보이는 모양은 원본과 같다."""
    from docx.oxml.ns import qn
    from lxml import etree
    reg = fitz.Rect(region) + (-1.5, -1.5, 1.5, 1.5)
    lines, fills = [], []
    for dr in page.get_drawings():
        for it in dr["items"]:
            if it[0] == "re":
                r = fitz.Rect(it[1])
            elif it[0] == "l":
                r = fitz.Rect(it[1], it[2]).normalize()
                w = max(dr.get("width") or 0.5, 0.3)
                r = r + ((-w / 2, 0, w / 2, 0) if r.width < 0.1 else (0, -w / 2, 0, w / 2))
            else:
                continue
            if not reg.contains(r):
                continue
            col = dr.get("fill") or dr.get("color")
            if col is None:
                continue
            if min(r.width, r.height) <= 1.6 and max(r.width, r.height) >= 2:
                lines.append((r, col))
            elif r.width > 3 and r.height > 3 and dr.get("fill") is not None:
                fills.append((r, dr["fill"]))
    hl = [(r, c) for r, c in lines if r.width >= r.height]
    vl = [(r, c) for r, c in lines if r.height > r.width]
    if len(hl) < 2 or len(vl) < 1:
        return None

    def cluster(vals, tol=1.6):
        out = []
        for v in sorted(vals):
            if out and v - out[-1][-1] <= tol:
                out[-1].append(v)
            else:
                out.append([v])
        return [sum(c) / len(c) for c in out]
    xs = cluster([r.x0 + r.width / 2 for r, _ in vl] + [r.x0 for r, _ in hl] + [r.x1 for r, _ in hl])
    ys = cluster([r.y0 + r.height / 2 for r, _ in hl] + [r.y0 for r, _ in vl] + [r.y1 for r, _ in vl])
    xs = [x for x in xs if reg.x0 <= x <= reg.x1]
    ys = [y for y in ys if reg.y0 <= y <= reg.y1]
    # 너무 가는 칸(선 두께 차이로 생긴 것)은 없앤다
    xs = [x for i, x in enumerate(xs) if i == 0 or x - xs[i - 1] > 2.2]
    ys = [y for i, y in enumerate(ys) if i == 0 or y - ys[i - 1] > 2.2]
    nx, ny = len(xs) - 1, len(ys) - 1
    if nx < 1 or ny < 1 or nx > 40 or ny > 60:
        return None

    def h_line(j, i):            # 줄 i 위쪽 가로선(눈금 y[i])이 칸 j 를 덮나
        y, x0, x1 = ys[i], xs[j], xs[j + 1]
        for r, c in hl:
            if abs(r.y0 + r.height / 2 - y) <= 1.6 and r.x0 <= x0 + 1.5 and r.x1 >= x1 - 1.5:
                return c, r.height
        return None

    def v_line(j, i):            # 칸 j 왼쪽 세로선(눈금 x[j])이 줄 i 를 덮나
        x, y0, y1 = xs[j], ys[i], ys[i + 1]
        for r, c in vl:
            if abs(r.x0 + r.width / 2 - x) <= 1.6 and r.y0 <= y0 + 1.5 and r.y1 >= y1 - 1.5:
                return c, r.width
        return None

    def fill_at(cx, cy):
        best = None
        for r, c in fills:
            if r.contains(fitz.Point(cx, cy)) and (best is None or r.get_area() < best[0].get_area()):
                best = (r, c)
        return best[1] if best else None
    # 칸 묶기: 가로로 세로선이 없으면 잇고, 세로로 가로선이 없고 같은 폭이면 잇는다
    owner = {}
    cells = []
    for i in range(ny):
        j = 0
        while j < nx:
            k = j + 1
            while k < nx and v_line(k, i) is None:
                k += 1
            cells.append({"r0": i, "r1": i + 1, "c0": j, "c1": k})
            j = k
    for c in cells:
        owner[(c["r0"], c["c0"])] = c
    merged = []
    for c in sorted(cells, key=lambda c: (c["r0"], c["c0"])):
        if c.get("gone"):
            continue
        i = c["r1"]
        while i < ny:
            below = owner.get((i, c["c0"]))
            if below is None or below["c1"] != c["c1"] or below.get("gone"):
                break
            if any(h_line(j, i) is not None for j in range(c["c0"], c["c1"])):
                break
            below["gone"] = True
            c["r1"] = i + 1
            i += 1
        merged.append(c)
    # 글 넣기
    words = page.get_text("dict", clip=reg)["blocks"]
    spans = []
    for b in words:
        for l in b.get("lines", []):
            for s in l["spans"]:
                if s["text"].strip():
                    sr = fitz.Rect(s["bbox"])
                    if xs[0] - 1 <= sr.x0 + sr.width / 2 <= xs[-1] + 1 and ys[0] - 1 <= sr.y0 + sr.height / 2 <= ys[-1] + 1:
                        spans.append((sr, s, l))          # 표 테두리 안의 글만(표 밖 각주는 원래 자리에 남긴다)
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    tbl = etree.Element(qn("w:tbl"))
    tp = etree.SubElement(tbl, qn("w:tblPr"))
    etree.SubElement(tp, qn("w:tblW"), {qn("w:w"): "0", qn("w:type"): "auto"})
    etree.SubElement(tp, qn("w:tblLayout"), {qn("w:type"): "fixed"})
    cm = etree.SubElement(tp, qn("w:tblCellMar"))
    for s_ in ("top", "left", "bottom", "right"):
        etree.SubElement(cm, qn("w:" + s_), {qn("w:w"): "0" if s_ in ("top", "bottom") else "40", qn("w:type"): "dxa"})
    grid = etree.SubElement(tbl, qn("w:tblGrid"))
    for j in range(nx):
        etree.SubElement(grid, qn("w:gridCol"), {qn("w:w"): str(int(round((xs[j + 1] - xs[j]) * 20 * W_scale)))})
    by_row = {}
    for c in merged:
        by_row.setdefault(c["r0"], []).append(c)
    cont = {}                     # (줄, 칸 시작) → 위에서 이어지는 칸
    for c in merged:
        for i in range(c["r0"] + 1, c["r1"]):
            cont.setdefault(i, []).append(c)

    def border_el(tcb, side, ln):
        if ln is None:
            etree.SubElement(tcb, qn("w:" + side), {qn("w:val"): "nil"})
        else:
            col, th = ln
            etree.SubElement(tcb, qn("w:" + side), {qn("w:val"): "single", qn("w:sz"): str(max(2, int(round(th * 8)))),
                                                   qn("w:color"): _hex(col)})
    for i in range(ny):
        tr = etree.SubElement(tbl, qn("w:tr"))
        trp = etree.SubElement(tr, qn("w:trPr"))
        etree.SubElement(trp, qn("w:trHeight"), {qn("w:val"): str(int(round((ys[i + 1] - ys[i]) * 20 * W_scale))),
                                                  qn("w:hRule"): "atLeast"})
        row = sorted(by_row.get(i, []) + [dict(c, _cont=True) for c in cont.get(i, [])], key=lambda c: c["c0"])
        for c in row:
            tc = etree.SubElement(tr, qn("w:tc"))
            tcp = etree.SubElement(tc, qn("w:tcPr"))
            span = c["c1"] - c["c0"]
            etree.SubElement(tcp, qn("w:tcW"), {qn("w:w"): str(int(round((xs[c["c1"]] - xs[c["c0"]]) * 20 * W_scale))),
                                                qn("w:type"): "dxa"})
            if span > 1:
                etree.SubElement(tcp, qn("w:gridSpan"), {qn("w:val"): str(span)})
            if c["r1"] - c["r0"] > 1:
                etree.SubElement(tcp, qn("w:vMerge"), {} if c.get("_cont") else {qn("w:val"): "restart"})
            tcb = etree.SubElement(tcp, qn("w:tcBorders"))
            border_el(tcb, "top", h_line(c["c0"], c["r0"]))
            border_el(tcb, "left", v_line(c["c0"], i))
            border_el(tcb, "bottom", h_line(c["c0"], c["r1"]) if c["r1"] <= ny - 1 or True else None)
            border_el(tcb, "right", v_line(c["c1"], i) if c["c1"] < len(xs) - 1 else None)
            rect = fitz.Rect(xs[c["c0"]], ys[c["r0"]], xs[c["c1"]], ys[c["r1"]])
            f = fill_at(rect.x0 + rect.width / 2, rect.y0 + rect.height / 2)
            if f is not None and _hex(f) != "FFFFFF":
                etree.SubElement(tcp, qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): _hex(f)})
            etree.SubElement(tcp, qn("w:vAlign"), {qn("w:val"): "center"})
            mine = [] if c.get("_cont") else [(r, s, l) for r, s, l in spans
                                              if rect.contains(fitz.Point(r.x0 + r.width / 2, r.y0 + r.height / 2))]
            lines_ = {}
            for r, s, l in mine:
                lines_.setdefault(round(l["bbox"][1] / 2), []).append((r, s))
            if not lines_:
                p = etree.SubElement(tc, qn("w:p"))
                pp = etree.SubElement(p, qn("w:pPr"))
                etree.SubElement(pp, qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0", qn("w:line"): "240", qn("w:lineRule"): "auto"})
                continue
            for key in sorted(lines_):
                segs = sorted(lines_[key], key=lambda x: x[0].x0)
                lx0 = min(r.x0 for r, _ in segs)
                lx1 = max(r.x1 for r, _ in segs)
                mid = (lx0 + lx1) / 2 - (rect.x0 + rect.width / 2)
                jc = "center" if abs(mid) < rect.width * 0.12 else ("left" if mid < 0 else "right")
                p = etree.SubElement(tc, qn("w:p"))
                pp = etree.SubElement(p, qn("w:pPr"))
                etree.SubElement(pp, qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0", qn("w:line"): "240", qn("w:lineRule"): "auto"})
                if jc == "left":
                    etree.SubElement(pp, qn("w:ind"), {qn("w:left"): str(max(0, int((lx0 - rect.x0 - 2) * 20 * W_scale)))})
                etree.SubElement(pp, qn("w:jc"), {qn("w:val"): jc})
                prev_x1 = None
                for r, s in segs:
                    run = etree.SubElement(p, qn("w:r"))
                    rp = etree.SubElement(run, qn("w:rPr"))
                    if s["flags"] & 16 or "bold" in s["font"].lower():
                        etree.SubElement(rp, qn("w:b"))
                    if s["flags"] & 4 or s.get("text", "") and False:
                        pass
                    etree.SubElement(rp, qn("w:color"), {qn("w:val"): "%06X" % s["color"]})
                    etree.SubElement(rp, qn("w:sz"), {qn("w:val"): str(max(2, int(round(s["size"] * 2 * W_scale))))})
                    t = etree.SubElement(run, qn("w:t"))
                    txt = s["text"]
                    if prev_x1 is not None and r.x0 - prev_x1 > s["size"] * 0.25 and not txt.startswith(" "):
                        txt = " " + txt
                    t.text = txt.rstrip() if (r, s) == segs[-1] else txt
                    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    prev_x1 = r.x1
    # 안전장치: 원본 글이 하나라도 빠지면(어느 칸에도 안 들어감) 다시 짜지 않는다(돈암동에서 글이 빠졌다)
    want = "".join(sp[1]["text"] for sp in spans).replace(" ", "")
    got = "".join(t.text or "" for t in tbl.iter(qn("w:t"))).replace(" ", "")
    if len(got) < len(want):
        return None
    return tbl


def clean_broken_nested_tables(d, grps):
    """pdf2docx 가 망가뜨린 '표 안의 표'(칸마다 글 한 토막짜리 작은 표가 3개 넘게 박힘)에서 글 없는 진한 칸의 칠을 지운다.
    ★방배동 금융조건: 원본 머리글의 검정 칠이 빈 칸들로 번져 큰 검은 덩어리가 됐다(2026-10-13). 글이 든 칸 색은 그대로. 고친 칸 수."""
    from docx.oxml.ns import qn
    n = 0
    for el in grps_iter(grps):
        for t in list(el.iter(qn("w:tbl"))):
            if not any(a.tag == qn("w:tbl") and not _is_wrapper_tbl(a) for a in t.iterancestors()):
                continue
            tcs = [c for c in t.iter(qn("w:tc")) if next((a for a in c.iterancestors() if a.tag == qn("w:tbl")), None) is t]
            frags = sum(1 for c in tcs for it in c.findall(qn("w:tbl")) if len(list(it.iter(qn("w:tc")))) <= 2)
            if frags < 3:
                continue
            for c in t.iter(qn("w:tc")):
                shd = c.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
                f = (shd.get(qn("w:fill")) or "") if shd is not None else ""
                if f and _lum_hex(f) < 0.5 and not "".join(x.text or "" for x in c.iter(qn("w:t"))).strip():
                    shd.getparent().remove(shd)
                    n += 1
    return n


def unwrap_cell_boxes(d):
    """칸 안에 든 1줄 1칸짜리 작은 표(pdf2docx 가 원본 '글 + 칠한 네모' 를 이렇게 만든다)를 풀어 글만 칸에 남긴다.
    ★돈암동 머리글 '구 분 | 내 용': 바깥 칸은 네이비로 칠했는데 안의 작은 표가 원본 검정이라 네이비 위에 검은 네모가 떴다
    (2026-10-14 사용자 지적 '헤더 글 부분이 도형처럼 겹친다'). 작은 표의 칠은 바깥 칸에 칠이 없을 때만 옮긴다. 푼 수."""
    import copy
    from docx.oxml.ns import qn
    n = 0
    for tc in list(d.element.body.iter(qn("w:tc"))):
        for t in [c for c in tc if c.tag == qn("w:tbl")]:
            rows = t.findall(qn("w:tr"))
            if len(rows) != 1:
                continue
            cells = rows[0].findall(qn("w:tc"))
            if len(cells) != 1 or cells[0].find(".//" + qn("w:tbl")) is not None:
                continue
            inner = cells[0]
            ishd = inner.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
            tcpr = tc.find(qn("w:tcPr"))
            if ishd is not None and tcpr is not None and tcpr.find(qn("w:shd")) is None:
                sh = copy.deepcopy(ishd)
                after = next((tcpr.find(qn(x)) for x in ("w:noWrap", "w:tcMar", "w:textDirection", "w:tcFitText",
                                                          "w:vAlign", "w:hideMark") if tcpr.find(qn(x)) is not None), None)
                after.addprevious(sh) if after is not None else tcpr.append(sh)
            for p in inner.findall(qn("w:p")):
                t.addprevious(p)
            tc.remove(t)
            n += 1
        # 칸 끝에 남은 1pt 빈 문단(작은 표 뒤 자리 맞춤용)은 지운다 — 칸에 문단이 하나 넘게 있을 때만
        ps = tc.findall(qn("w:p"))
        if len(ps) > 1 and not "".join(x.text or "" for x in ps[-1].iter(qn("w:t"))).strip()                 and ps[-1].find(".//" + qn("w:drawing")) is None and tc[-1] is ps[-1]                 and ps[-1].getprevious() is not None and ps[-1].getprevious().tag != qn("w:tbl"):   # 표 바로 뒤 문단은 워드 규칙상 꼭 있어야 한다(지우면 '파일 손상')
            sp = ps[-1].find(f"{qn('w:pPr')}/{qn('w:spacing')}")
            if sp is not None and sp.get(qn("w:lineRule")) == "exact":
                tc.remove(ps[-1])
    return n


def _is_wrapper_tbl(t):
    from docx.oxml.ns import qn
    rows = t.findall(qn("w:tr"))
    tcs = rows[0].findall(qn("w:tc")) if len(rows) == 1 else []
    return len(tcs) == 1 and tcs[0].find(qn("w:tbl")) is not None


def fix_broken_nested_tables(d, grps, page):
    """pdf2docx 가 망가뜨린 '표 안의 표' 를 원본 그 부분 그림으로 바꾼다. 바꾼 수.
    망가진 표 = 글이 든 안쪽 표인데 글 없는 진한 칸이 5개 이상(원본 머리글 칠이 엉뚱한 빈 칸들에 흩어짐).
    ★방배동 2쪽 '금융조건'(표 안의 표 안의 표): 검은/남색 덩어리로 뒤죽박죽 — 어제 판도 같았다(2026-10-13).
    그 표를 품은 맨 바깥 칸의 내용을 통째로, 원본에서 그 칸 글들이 있는 자리를 잘라 낸 그림으로 바꾼다."""
    from docx.oxml.ns import qn
    from docx.shared import Pt
    n = 0
    for el in grps_iter(grps):
        for t in list(el.iter(qn("w:tbl"))):
            if t.getparent() is None or not any(a.tag == qn("w:tbl") and not _is_wrapper_tbl(a) for a in t.iterancestors()):
                continue
            tcs = [c for c in t.iter(qn("w:tc")) if next((a for a in c.iterancestors() if a.tag == qn("w:tbl")), None) is t]
            dark_empty = 0
            for c in tcs:
                shd = c.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
                f = (shd.get(qn("w:fill")) or "") if shd is not None else ""
                if f and _lum_hex(f) < 0.5 and not "".join(x.text or "" for x in c.iter(qn("w:t"))).strip():
                    dark_empty += 1
            # 칸마다 글 한 토막짜리 작은 표(1~2칸)가 3개 넘게 박힌 것 = pdf2docx 가 망가뜨린 모양(방배동 금융조건)
            frags = 0
            for c in tcs:
                for it in c.findall(qn("w:tbl")):
                    if len(list(it.iter(qn("w:tc")))) <= 2:
                        frags += 1
            deep = frags >= 3
            if (dark_empty < 5 and not deep) or not "".join(x.text or "" for x in t.iter(qn("w:t"))).strip():
                continue
            # 맨 바깥(껍데기 아닌) 표의 칸 = 바꿀 자리
            cells = [a for a in t.iterancestors() if a.tag == qn("w:tc")]
            outer = None
            for c in cells:
                tb = next((a for a in c.iterancestors() if a.tag == qn("w:tbl")), None)
                if tb is not None and not any(x.tag == qn("w:tbl") and not _is_wrapper_tbl(x) for x in tb.iterancestors()):
                    outer = c
            if outer is None:
                continue
            words = {x.strip() for x in ("".join(r.text or "" for r in p.iter(qn("w:t"))) for p in outer.iter(qn("w:p")))
                     if len(x.strip()) >= 2}
            box = fitz.Rect()
            for w in words:
                for r in page.search_for(w)[:1]:
                    box |= r
            if box.is_empty or box.width < 30:
                continue
            box = (box + (-8, -8, 8, 8)) & page.rect
            new_tbl = rebuild_table_from_pdf(page, box)
            if new_tbl is None:
                continue
            # 바깥 칸 안에서 망가진 표를 품은 표만 새 표로 바꾼다(그 앞뒤 문단·각주는 그대로)
            holder = None
            for ch in outer:
                if ch.tag == qn("w:tbl") and (ch is t or any(x is t for x in ch.iter(qn("w:tbl")))):
                    holder = ch
                    break
            if holder is None:
                continue
            holder.addprevious(new_tbl)
            outer.remove(holder)
            if outer[-1].tag != qn("w:p"):
                outer.append(outer.makeelement(qn("w:p"), {}))     # 칸은 문단으로 끝나야 한다
            # 칸 줄 높이는 내용에 맞게 '최소'
            tr = next((a for a in outer.iterancestors() if a.tag == qn("w:tr")), None)
            h = tr.find(f"{qn('w:trPr')}/{qn('w:trHeight')}") if tr is not None else None
            if h is not None:
                h.set(qn("w:hRule"), "atLeast")
            n += 1
            break
    return n


def grps_iter(grps):
    for g in grps:
        for el in g:
            yield el


def _row_cells(tbl):
    """표 첫 줄의 (칸 너비 목록, 왼쪽 들여쓰기) — 같은 표가 쪼개졌는지 비교용(twips)."""
    from docx.oxml.ns import qn
    tr = tbl.find(qn("w:tr"))
    if tr is None:
        return None
    ws = []
    for tc in tr.findall(qn("w:tc")):
        w = tc.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
        span = tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
        ws.append((int(w.get(qn("w:w"))) if w is not None and (w.get(qn("w:w")) or "").isdigit() else -1,
                   span.get(qn("w:val")) if span is not None else "1"))
    ind = tbl.find(f"{qn('w:tblPr')}/{qn('w:tblInd')}")
    return ws, int(ind.get(qn("w:w"))) if ind is not None and (ind.get(qn("w:w")) or "").lstrip("-").isdigit() else 0


def merge_split_tables(d):
    """한 표가 줄마다 따로 쪼개진 것(pdf2docx 가 원본의 굵은 선·빨간 네모를 표 경계로 봄)을 다시 한 표로.
    ★방배동 거래사례 표(2026-10-12 사용자 지적 '표가 아니라 선을 이어붙인 느낌'): 16줄이 1줄짜리 표 16개였다.
    조건: 사이에 빈 문단만 있고, 첫 줄 칸 수·칸 너비(±3pt)·왼쪽 위치(±3pt)가 같을 때만."""
    from docx.oxml.ns import qn
    n = 0
    body = d.element.body
    for tbl in list(body.iter(qn("w:tbl"))):
        if tbl.getparent() is not body or tbl.getparent() is None:
            continue
        sig = _row_cells(tbl)
        while sig:
            gap, nxt = [], tbl.getnext()
            while nxt is not None and nxt.tag == qn("w:p") and _empty(nxt) and len(gap) < 2:
                gap.append(nxt)
                nxt = nxt.getnext()
            if nxt is None or nxt.tag != qn("w:tbl"):
                break
            other = _row_cells(nxt)
            if not other or len(other[0]) != len(sig[0]) or abs(other[1] - sig[1]) > 60 or \
                    any(a[1] != b[1] or abs(a[0] - b[0]) > 60 for a, b in zip(sig[0], other[0])):
                break
            for tr in nxt.findall(qn("w:tr")):
                tbl.append(tr)
            for g in gap:
                body.remove(g)
            body.remove(nxt)
            n += 1
    return n


def _tbl_rows_x(tbl):
    """표의 줄마다 [(시작x, 끝x, 칸)] (twips, 표 들여쓰기 포함)."""
    from docx.oxml.ns import qn
    ind = tbl.find(f"{qn('w:tblPr')}/{qn('w:tblInd')}")
    x0 = float(ind.get(qn("w:w"))) if ind is not None and re.fullmatch(r"-?[\d.]+", ind.get(qn("w:w")) or "") else 0.0
    out, prev = [], []
    for tr in tbl.findall(qn("w:tr")):
        x, row = x0, []
        for tc in tr.findall(qn("w:tc")):
            w = tc.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
            wv = float(w.get(qn("w:w"))) if w is not None and re.fullmatch(r"[\d.]+", w.get(qn("w:w")) or "") else 0.0
            vm = tc.find(f"{qn('w:tcPr')}/{qn('w:vMerge')}")
            if vm is not None and vm.get(qn("w:val")) in (None, "continue"):
                # 세로로 이어지는 칸은 pdf2docx 가 폭을 틀리게 적곤 한다(이천 대월 723 ↔ 위 칸 814) → 위 칸 폭을 쓴다
                above = next((b for b in prev if abs(b[0] - x) <= 60), None)
                if above is not None:
                    wv = above[1] - above[0]
            row.append((x, x + wv, tc))
            x += wv
        out.append((tr, row))
        prev = row
    return out


def _vmerge_header_cells(tbl):
    """다시 짠 표의 머리글(칠한 칸) 구역에서, 같은 칸 눈금 위·아래로 이어진 칸 중 글이 하나뿐이면 세로로 합친다.
    (이천 대월: '구분'·'계' 는 3줄, '계약금' 은 2줄을 세로로 차지하는데 줄마다 따로 칸이 됐다)"""
    from docx.oxml.ns import qn

    def headish(tc):
        shd = tc.find(f"{qn('w:tcPr')}/{qn('w:shd')}")
        f = (shd.get(qn("w:fill")) or "").upper() if shd is not None else ""
        if f not in ("", "AUTO", "FFFFFF"):
            return True
        cols = [(c.get(qn("w:val")) or "").upper() for c in tc.iter(qn("w:color"))]
        return bool(cols) and all(c == "FFFFFF" for c in cols) and \
            bool("".join(t.text or "" for t in tc.iter(qn("w:t"))).strip())

    head_rows = set()                                      # 칸 대부분이 칠했거나 흰 글씨인 줄 = 머리글 줄
    for i, tr in enumerate(tbl.findall(qn("w:tr"))):
        tcs = [tc for tc in tr.findall(qn("w:tc")) if "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip()]
        if tcs and sum(headish(tc) for tc in tcs) >= 0.6 * len(tcs):
            head_rows.add(i)

    def fill(tc, r):
        return r in head_rows

    grid = []                                              # 줄마다 {눈금 시작: (칸, 칸 수)}
    for tr in tbl.findall(qn("w:tr")):
        pos, row = 0, {}
        for tc in tr.findall(qn("w:tc")):
            gs = tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}")
            span = int(gs.get(qn("w:val"))) if gs is not None else 1
            row[pos] = (tc, span)
            pos += span
        grid.append(row)
    cols = sorted({k for row in grid for k in row})
    for c in cols:
        r = 0
        while r < len(grid):
            run = []
            while r < len(grid) and c in grid[r] and fill(grid[r][c][0], r) and \
                    (not run or grid[r][c][1] == run[0][1][1]):
                run.append((r, grid[r][c]))
                r += 1
            if len(run) >= 2:
                texty = [x for x in run if "".join(t.text or "" for t in x[1][0].iter(qn("w:t"))).strip()]
                if len(texty) == 1:
                    import copy
                    src_tc = texty[0][1][0]
                    first_tc = run[0][1][0]
                    if src_tc is not first_tc:                 # 글을 맨 위 칸으로 옮긴다
                        for p in first_tc.findall(qn("w:p")):
                            first_tc.remove(p)
                        for p in src_tc.findall(qn("w:p")):
                            first_tc.append(copy.deepcopy(p))
                    for k, (_, (tc, _)) in enumerate(run):
                        pr = tc.find(qn("w:tcPr"))
                        vm = pr.find(qn("w:vMerge"))
                        if vm is None:
                            vm = pr.makeelement(qn("w:vMerge"), {})
                            after = next((pr.find(qn(t)) for t in ("w:tcBorders", "w:shd", "w:noWrap", "w:tcMar",
                                                                    "w:textDirection", "w:tcFitText", "w:vAlign",
                                                                    "w:hideMark") if pr.find(qn(t)) is not None), None)
                            after.addprevious(vm) if after is not None else pr.append(vm)
                        if k == 0:
                            vm.set(qn("w:val"), "restart")
                        else:
                            vm.attrib.pop(qn("w:val"), None)
                            for p in tc.findall(qn("w:p")):
                                for rr in p.findall(qn("w:r")):
                                    p.remove(rr)
            if not run:
                r += 1


def merge_split_tables_in_cells(d):
    """칸 안에서 한 표가 여러 표로 쪼개지고(칸 경계도 서로 조금 다름) 세로로 합친 칸 글이 따로 떨어진 것을 다시 한 표로.
    ★이천 대월 '중도금대출 일정표'(2026-10-12 사용자 지적): [위 머리줄 표] '계약금' 문단 [가운데 3줄 표] [납입주체 표] 로
    쪼개져 '계약금' 이 따로 놀았다. → 모든 표의 칸 경계를 모아 하나의 칸 눈금으로 다시 짜고, 떨어진 짧은 글은
    그 문단 들여쓰기 위치에 해당하는 위 표 마지막 줄의 빈 칸에 넣는다.
    조건: 같은 칸 안, 사이엔 빈 문단이나 짧은 글(8자 이하) 문단만, 바깥 왼쪽·오른쪽 끝이 서로 3pt 안."""
    import copy
    from docx.oxml.ns import qn
    n = 0
    for tc_outer in list(d.element.body.iter(qn("w:tc"))):
        kids = [k for k in tc_outer if k.tag in (qn("w:p"), qn("w:tbl"))]
        i = 0
        while i < len(kids):
            if kids[i].tag != qn("w:tbl"):
                i += 1
                continue
            group, gaps, j = [kids[i]], [], i + 1
            while j < len(kids):
                k = kids[j]
                if k.tag == qn("w:p"):
                    txt = "".join(t.text or "" for t in k.iter(qn("w:t"))).strip()
                    if len(txt) > 8 or k.find(".//" + qn("w:drawing")) is not None:
                        break
                    gaps.append((len(group), k, txt))
                    j += 1
                    continue
                a, b = _tbl_rows_x(group[0]), _tbl_rows_x(k)
                if not a or not b or not a[0][1] or not b[0][1]:
                    break
                if abs(a[0][1][0][0] - b[0][1][0][0]) > 60 or abs(a[0][1][-1][1] - b[0][1][-1][1]) > 60:
                    break
                group.append(k)
                j += 1
            # 끝에 붙은 빈 문단은 묶음에서 뺀다
            while gaps and gaps[-1][0] == len(group):
                gaps.pop()
            # ★표 사이에 '떨어진 글'(계약금 같은 짧은 글)이 있을 때만 합친다 — 그런 게 없으면 원래 잘 나온 표
            #   (이천 대월 사업진행현황: 합쳤더니 '2025.10 실착공' 줄이 깨졌다)
            if len(group) < 2 or not any(txt for _, _, txt in gaps):
                i = j
                continue
            rows = []
            for gi, t in enumerate(group):
                for tr, row in _tbl_rows_x(t):
                    rows.append((gi, tr, row))
            # 떨어진 짧은 글 → 바로 위 표 마지막 줄에서 그 위치의 빈 칸으로
            for gi, p, txt in gaps:
                if not txt or gi == 0:
                    continue
                ind = p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
                px = float(ind.get(qn("w:left"))) if ind is not None and re.fullmatch(r"-?[\d.]+", ind.get(qn("w:left")) or "") else 0.0
                last = [r for r in rows if r[0] == gi - 1][-1][2]
                for x0, x1, tc in last:
                    if x0 - 20 <= px < x1 and not "".join(t.text or "" for t in tc.iter(qn("w:t"))).strip():
                        tp = next(iter(tc.findall(qn("w:p"))), None)
                        if tp is not None:
                            tc.remove(tp)
                        newp = copy.deepcopy(p)
                        ind2 = newp.find(f"{qn('w:pPr')}/{qn('w:ind')}")
                        if ind2 is not None:
                            ind2.set(qn("w:left"), "0")
                        jc = newp.find(f"{qn('w:pPr')}/{qn('w:jc')}")
                        if jc is not None:
                            jc.set(qn("w:val"), "center")
                        for sh in list(newp.iter(qn("w:shd"))) + list(newp.iter(qn("w:highlight"))):
                            sh.getparent().remove(sh)             # 문단·글자 바탕(검은 상자)은 칸 색에 맡긴다
                        tc.append(newp)
                        break
            # 칸 경계 모으기(40twips 안은 같은 경계)
            xs = sorted(x for _, _, row in rows for x0, x1, _ in row for x in (x0, x1))
            edges = []
            for x in xs:
                if not edges or x - edges[-1][-1] > 40:
                    edges.append([x])
                else:
                    edges[-1].append(x)
            E = [sum(e) / len(e) for e in edges]

            def snap(x):
                return min(range(len(E)), key=lambda k: abs(E[k] - x))
            base = group[0]
            grid = base.find(qn("w:tblGrid"))
            for g in list(grid):
                grid.remove(g)
            for k in range(len(E) - 1):
                grid.append(grid.makeelement(qn("w:gridCol"), {qn("w:w"): str(int(round(E[k + 1] - E[k])))}))
            ind = base.find(f"{qn('w:tblPr')}/{qn('w:tblInd')}")
            if ind is not None:
                ind.set(qn("w:w"), str(int(round(E[0]))))
            for gi, tr, row in rows:
                for x0, x1, tc in row:
                    a, b = snap(x0), snap(x1)
                    span = max(1, b - a)
                    pr = tc.find(qn("w:tcPr"))
                    w = pr.find(qn("w:tcW"))
                    if w is not None:
                        w.set(qn("w:w"), str(int(round(E[b] - E[a]))))
                    gs = pr.find(qn("w:gridSpan"))
                    if span > 1:
                        if gs is None:
                            gs = pr.makeelement(qn("w:gridSpan"), {})
                            w.addnext(gs) if w is not None else pr.insert(0, gs)
                        gs.set(qn("w:val"), str(span))
                    elif gs is not None:
                        pr.remove(gs)
                if gi > 0:
                    base.append(tr)
            for t in group[1:]:
                tc_outer.remove(t)
            for _, p, _ in gaps:
                if p.getparent() is tc_outer:
                    tc_outer.remove(p)
            _vmerge_header_cells(base)
            n += 1
            kids = [k for k in tc_outer if k.tag in (qn("w:p"), qn("w:tbl"))]
            i = kids.index(base) + 1
    return n


def clear_white_overlays(d):
    """떠 있는 그림 중 대부분 흰색인 것(pdf2docx 가 원본의 선·동그라미·테두리 같은 벡터 그림을 '흰 바탕 그림' 으로
    구운 것)의 흰 바탕을 투명하게. ★고산 18쪽(2026-10-12): 빨간 동그라미+회색 테두리 그림이 흰 바탕째 지도 위에 얹혀
    지도가 가려졌다. 사진(흰 바탕이 거의 없는 것)·글 줄 안에 든 그림은 안 건드린다."""
    from PIL import Image
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    done, n = set(), 0
    for anc in d.element.body.iter(WP + "anchor"):
        blip = next(anc.iter(A + "blip"), None)
        rid = blip.get(R + "embed") if blip is not None else None
        if not rid or rid in done:
            continue
        done.add(rid)
        part = d.part.related_parts.get(rid)
        try:
            im = Image.open(io.BytesIO(part.blob)).convert("RGBA")
        except Exception:
            continue
        px = im.getdata()
        white = sum(1 for p in px if p[0] >= 245 and p[1] >= 245 and p[2] >= 245 and p[3] > 0)
        if white < 0.8 * len(px):
            continue
        im.putdata([(r, g, b, 0) if r >= 245 and g >= 245 and b >= 245 else (r, g, b, a) for r, g, b, a in px])
        buf = io.BytesIO()
        im.save(buf, "PNG")
        part._blob = buf.getvalue()
        n += 1
    return n


def loosen_row_heights(d):
    """pdf2docx 는 표 줄 높이를 전부 '정확히'(exact) 로 박는다 → 글꼴이 Pretendard 로 바뀌어 줄이 조금 높아지면
    칸 글씨 아래가 잘린다(고산3지구 '대출 조건' 칸·'사업 개요' 네모 상자, 2026-10-12 사용자 지적).
    → 글이 든 줄은 '최소'(atLeast) 로 바꿔 넘치면 칸이 늘어나게. 글 없는 줄(얇은 선·간격용)은 그대로 둔다."""
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    n = 0
    for tr in reversed(list(d.element.body.iter(W + "tr"))):   # 안쪽 표 줄부터 정해야 바깥 줄 어림이 맞다
        h = tr.find(f"{W}trPr/{W}trHeight")
        if h is None or h.get(W + "hRule") != "exact":
            continue
        if not any((t.text or "").strip() for t in tr.iter(W + "t")):
            continue
        h.set(W + "hRule", "atLeast")
        if (h.get(W + "val") or "").isdigit() and _row_growth(tr) <= 0:
            h.set(W + "hRule", "exact")                    # 글이 들어맞는 줄은 그대로(괜히 늘어나 쪽이 넘치지 않게)
            continue
        n += 1
    return n


_WING_BULLETS = {"": "•", "": "•", "": "○", "": "●", "": "■", "": "▪",
                 "": "◆", "": "◆", "": "➢", "": "➢", "": "✓", "": "•"}


def fix_symbol_runs(d):
    """기호 글꼴(Wingdings 등)이 붙어 있는데 글자는 보통 유니코드 기호인 조각 → 'Segoe UI Symbol'.

    ★pdf2docx 가 '▪'(U+25AA)에 Wingdings 를 붙여 와서, 워드가 Wingdings 의 같은 번호 모양인 '✦' 로
      그렸다(안산). Wingdings 전용 글자(U+F000~F0FF)인 조각은 그대로 둔다.
    """
    from docx.oxml.ns import qn
    n = 0
    for r in d.element.body.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        rf = rpr.find(qn("w:rFonts")) if rpr is not None else None
        if rf is None or not any(_SYMBOL_FONT_RE.search(rf.get(qn("w:" + k)) or "") for k in ("ascii", "hAnsi", "eastAsia", "cs")):
            continue
        txt = "".join(t.text or "" for t in r.findall(qn("w:t")))
        # 자주 쓰는 Wingdings 글머리(전용 번호)는 보통 기호로 — 워드에서 거의 안 보이고 내어쓰기도 못 했다(대전 하이라이트 '')
        if txt.strip() and all(c in _WING_BULLETS or c.isspace() for c in txt):
            for t in r.findall(qn("w:t")):
                t.text = "".join(_WING_BULLETS.get(c, c) for c in (t.text or ""))
            txt = "".join(t.text or "" for t in r.findall(qn("w:t")))
        if txt and all(not (0xF000 <= ord(c) <= 0xF0FF) for c in txt):
            for k in ("ascii", "hAnsi", "eastAsia", "cs"):
                rf.set(qn("w:" + k), "Segoe UI Symbol")
            n += 1
    return n


def _pages(d):
    """구역을 '원본 한 쪽' 단위로 묶는다 → [(구역들, 블록 묶음들)].

    ★pdf2docx 는 한 쪽 안에서도 단(칼럼) 나누기에 구역을 여러 개 쓴다(새 쪽 → 2단 이어서 → 다음 단).
      구역 하나 = 한 쪽으로 보고 전부 '새 쪽' 으로 바꿨더니 빈 쪽이 생겼다(등촌역 7쪽이 구역 3개).
      → '새 쪽으로 시작하는 구역' 에서만 쪽을 나눈다.
    """
    from docx.enum.section import WD_SECTION
    secs, grps = list(d.sections), _section_groups(d)
    out = []
    for i, (sec, grp) in enumerate(zip(secs, grps)):
        if i == 0 or sec.start_type in (WD_SECTION.NEW_PAGE, WD_SECTION.ODD_PAGE, WD_SECTION.EVEN_PAGE):
            out.append(([], []))
        out[-1][0].append(sec)
        out[-1][1].append(grp)
    return out


def _set_geometry(sec, new_page=True):
    from docx.shared import Cm
    from docx.enum.section import WD_SECTION, WD_ORIENT
    sec.orientation = WD_ORIENT.PORTRAIT
    sec.page_width, sec.page_height = Cm(PAGE_W), Cm(PAGE_H)
    sec.top_margin, sec.bottom_margin = Cm(TOP_CM), Cm(BOTTOM_CM)
    sec.left_margin, sec.right_margin = Cm(LEFT_CM), Cm(RIGHT_CM)
    sec.header_distance, sec.footer_distance = Cm(HEADER_CM), Cm(FOOTER_CM)
    if new_page:
        sec.start_type = WD_SECTION.NEW_PAGE


def stretch_table_to_width(tbl, target_tw):
    """맨 바깥 표를 본문 폭(target_tw, twips) 그대로 늘린다 — 왼쪽 들여쓰기 0, 칸 폭은 같은 비율로.
    ★2026-10-14 사용자 지적(편집해 보여 줌): 표가 머리말 선보다 좁고 여백이 더 컸다. 머리말 길이 = 표 시작~끝.
    본문 폭의 70% 보다 좁은 표(작은 표·옆에 둔 표)는 원본 크기 그대로 둔다."""
    from docx.oxml.ns import qn
    grid = tbl.find(qn("w:tblGrid"))
    if grid is None:
        return
    cols = grid.findall(qn("w:gridCol"))
    ws = [float(c.get(qn("w:w")) or 0) for c in cols]

    def row_sum(tr):
        tot = 0.0
        for tc in tr.findall(qn("w:tc")):
            w = tc.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
            if w is None or not re.fullmatch(r"[\d.]+", w.get(qn("w:w")) or "") or w.get(qn("w:type")) not in (None, "dxa"):
                return None
            tot += float(w.get(qn("w:w")))
        return tot
    # ★실제 폭은 칸 폭(tcW) 합 — pdf2docx 는 눈금(gridCol)을 똑같이 나눠 적어 둘이 다를 때가 있다(돈암 9070 vs 10046)
    sums = [x for x in (row_sum(tr) for tr in tbl.findall(qn("w:tr"))) if x]
    total = max(sums) if sums else sum(ws)
    if total <= 0 or total < target_tw * 0.7:
        return
    f = target_tw / total
    first = tbl.find(qn("w:tr"))
    fw = [float(tc.find(f"{qn('w:tcPr')}/{qn('w:tcW')}").get(qn("w:w"))) for tc in first.findall(qn("w:tc"))]         if first is not None and row_sum(first) else None
    if fw and len(fw) == len(cols) and all(tc.find(f"{qn('w:tcPr')}/{qn('w:gridSpan')}") is None
                                           for tc in first.findall(qn("w:tc"))):
        for c, w in zip(cols, fw):                           # 눈금을 첫 줄 칸 폭에 맞춘다
            c.set(qn("w:w"), str(int(round(w * f))))
    else:
        gf = target_tw / (sum(ws) or 1)
        for c, w in zip(cols, ws):
            c.set(qn("w:w"), str(int(round(w * gf))))
    for tc in tbl.iter(qn("w:tc")):                          # 안쪽 표 칸까지 같은 비율(안쪽 표도 같이 넓어진다)
        w = tc.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
        if w is not None and re.fullmatch(r"[\d.]+", w.get(qn("w:w")) or "") and w.get(qn("w:type")) in (None, "dxa"):
            w.set(qn("w:w"), str(int(round(float(w.get(qn("w:w"))) * f))))
    for g in tbl.iter(qn("w:gridCol")):
        if g.getparent() is not grid and re.fullmatch(r"[\d.]+", g.get(qn("w:w")) or ""):
            g.set(qn("w:w"), str(int(round(float(g.get(qn("w:w"))) * f))))
    tp = tbl.find(qn("w:tblPr"))
    ind = tp.find(qn("w:tblInd")) if tp is not None else None
    if ind is not None:
        ind.set(qn("w:w"), "0")
        ind.set(qn("w:type"), "dxa")
    tw = tp.find(qn("w:tblW")) if tp is not None else None
    if tw is not None:
        tw.set(qn("w:w"), str(int(target_tw)))
        tw.set(qn("w:type"), "dxa")


def page_text_lines(page):
    """쪽 글 줄 [(가장 큰 글자 크기, 줄 높이, 줄 가운데 y)] — 9pt 규칙으로 커질 몫 어림용."""
    out = []
    for b in page.get_text("dict")["blocks"]:
        for l in b.get("lines", []):
            sz = [sp["size"] for sp in l["spans"] if sp["text"].strip()]
            if sz:
                out.append((max(sz), l["bbox"][3] - l["bbox"][1], (l["bbox"][1] + l["bbox"][3]) / 2))
    return out


def _fit_page(secs, grps, box, lines=None):
    """원본 한 쪽(구역 1개 이상)을 양식 규격으로: 내용을 같은 비율로 줄이고(크면만) 위쪽 빈 자리를 없앤다.

    반환: (비율, 결과에서 쓰인 높이 pt)
    """
    from docx.oxml.ns import qn
    first = secs[0]
    old_left, old_top = first.left_margin.pt, first.top_margin.pt
    old_w = first.page_width.pt - first.left_margin.pt - first.right_margin.pt
    grow = sum(_row_growth(el) for grp in grps for el in grp)   # 표 칸이 Pretendard 로 늘어날 몫(원본 pt)
    # 어림은 표가 빽빽한 쪽에서 크게 넘친다(돈암동 6쪽 573pt → 쪽이 반으로 줄어 표가 쪼그라들고 쪽이 넘쳤다).
    # → 늘어날 몫은 쪽 높이의 GROW_CAP 까지만 반영한다.
    grow = min(grow, box.height * GROW_CAP)
    # ★가로(표 폭·칸 폭·들여쓰기)는 양식 본문 폭 그대로 두고(넓으면만 줄임), 세로(글자 크기·줄 간격·줄 높이·그림)만 줄여 한 쪽에 맞춘다.
    #   예전엔 통째로 줄여 표가 머리말·꼬리말 선보다 좁았다(방배동 86~94%, 2026-10-13 사용자 지적).
    sw = min(1.0, BOX_W_PT / max(old_w, 1))
    s = min(sw, BOX_H_PT * HEIGHT_SAFETY / max(box.height + grow, 1))
    used = (box.height + grow) * s                         # 결과에서 쓸 높이 어림(그림 키울 남는 높이 셈에 씀)
    if lines and s < 0.999:
        # ★줄인 뒤 본문 글자를 9pt 로 다시 키우면(글자 크기 규칙) 그만큼 길어져 쪽이 넘쳤다(내곡동 하이라이트 각주 한 줄, 2026-10-15)
        #   → 커질 몫(줄 높이 × 커지는 비율, 줄바꿈 늘어남까지 1.5제곱)을 넣고 쪽에 들어가는 가장 큰 비율을 찾는다.
        H = box.height + grow

        def need(x):
            # ★같은 높이에 나란히 있는 줄(표 칸 여러 개)은 한 번만 — 다 더했더니 이천 3쪽이 '글 높이 1,897pt' 로 잡혀
            #   쪽마다 0.30(최저)까지 줄여 그림이 원본의 1/3 로 쪼그라들었다(2026-10-21 사용자 '사진을 왜 이렇게 작게').
            band = {}
            for item in lines:
                sz, h = item[0], item[1]
                yc = item[2] if len(item) > 2 else len(band)
                v = sz * x
                # ★_sz_target 은 '원본' 8~10.5pt 글을 줄인 비율과 상관없이 9pt 로 둔다 → 조건도 원본 크기로(내곡 33·44쪽, 2026-10-20)
                if 8 <= sz <= 10.5 and v < MIN_SZ / 2:
                    # 9pt 로 고정되는 줄의 결과 높이 = h·(9/sz)^1.5(줄바꿈 늘어남 포함) — 쪽을 얼마나 줄이든 그대로다.
                    # ★예전 식 h·(9/(sz·x))^1.5 는 x 가 작을수록 커져(×x^-0.5) 많이 줄인 쪽일수록 부풀려 셌다
                    #   → 아래 x·(H+extra) 에 넣으려고 x 로 나눠 둔다
                    e = h * ((MIN_SZ / 2 / sz) ** 1.5 / x - 1)
                    k_ = round(yc / 3)
                    band[k_] = max(band.get(k_, 0.0), e)
            return x * (H + sum(band.values()))
        target = BOX_H_PT * HEIGHT_SAFETY
        lo, hi = FIT_FLOOR, s
        if need(hi) > target:
            if need(lo) > target:
                # 9pt 로 고정되는 글만으로도 안 들어가는 쪽 — 더 줄여도 글은 안 줄고 그림·표만 쪼그라든다 → 가장 작게 드는 비율로
                s = min((need(lo + i * (hi - lo) / 40), lo + i * (hi - lo) / 40) for i in range(41))[1]
            else:
                for _ in range(25):
                    mid = (lo + hi) / 2
                    if need(mid) <= target:
                        lo = mid
                    else:
                        hi = mid
                s = lo
        used = need(s)
    for i, sec in enumerate(secs):
        _set_geometry(sec, new_page=(i == 0))
        cols = sec._sectPr.find(qn("w:cols"))               # 2단 같은 단 너비·간격도 같은 비율로
        if cols is not None and sw < 0.999:
            for el in [cols] + cols.findall(qn("w:col")):
                for a in ("space", "w"):
                    v = el.get(qn("w:" + a))
                    if v and v.isdigit():
                        el.set(qn("w:" + a), str(int(int(v) * sw)))
    for grp in grps:
        for el in grp:
            scale_xml(el, s, old_left, sw)
            _place_anchors(el, s, old_left, box.y0, sw)
    _flatten_lead_anchors(grps[0])
    lead = max(0.0, (box.y0 - old_top) * s)                 # 위 띠를 지운 자리 = 쓸데없는 빈 칸
    trim_leading(grps[0], lead)
    for grp in grps[:-1]:                                    # 쪽 안의 구역 끝(단 나눔) 문단도 납작하게
        if grp:
            _flatten_holder(grp[-1]) if grp[-1].find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None else None
    trim_trailing(grps[-1])
    if s < 0.999:                                            # 글자 최소 8pt 로 커진 칸 글이 잘리지 않게
        for grp in grps:
            for el in grp:
                for tr in el.iter(qn("w:tr")):
                    h = tr.find(f"{qn('w:trPr')}/{qn('w:trHeight')}")
                    if h is not None and h.get(qn("w:hRule")) == "exact" and                             any((t.text or "").strip() for t in tr.iter(qn("w:t"))):
                        h.set(qn("w:hRule"), "atLeast")
    for grp in grps:                                         # 표는 머리말·꼬리말 선 길이(본문 폭) 그대로 꽉 채운다
        for el in grp:
            if el.tag == qn("w:tbl"):
                stretch_table_to_width(el, BOX_W_PT * 20)
    # 남는 높이 = 쪽 높이 − (내용 + 표 칸이 Pretendard 로 늘어날 몫) × s. 예전엔 늘어날 몫(grow)을 빼먹어 남는 높이를
    #   부풀려 셌고, 그림이 커지며 쪽 끝 표 줄이 넘쳤다(내곡동 59→61쪽, 2026-10-20) → 빼고, 어림 오차 몫으로 절반만 쓴다
    #   + 어림 오차 몫으로 쪽 높이 5% 는 남겨 둔다(그림이 6pt 만 커져도 꽉 찬 쪽 끝 각주가 넘쳤다 — 내곡 31쪽)
    #   + 9pt 로 다시 커지는 몫(used 에 포함)도 빼야 한다 — 빼먹어 내곡 29쪽이 183pt 남는다고 셌다
    grow_pictures_to_width(grps, 0.5 * max(0.0, BOX_H_PT * (HEIGHT_SAFETY - 0.05) - used))
    return s, box.height * s


def grow_pictures_to_width(grps, slack_pt):
    """표 밖에 홀로 있는 그림을 표처럼 본문 폭(머리말 줄 길이)까지 키운다 — 비율 그대로(자르지 않음).
    쪽에 남는 높이(slack)만큼만 키우고, 모자라면 들어가는 만큼 최대한 키워 가운데 정렬.
    ★2026-10-20 사용자: 표는 다 본문 폭인데 그림만 좁다 — 쪽 맞춤에서 그림은 세로 비율(s)로 같이 줄고 표만 다시 늘렸다."""
    from docx.oxml.ns import qn
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    target = BOX_W_PT * 12700
    items = []
    for grp in grps:
        for el in grp:
            if el.tag != qn("w:p") or _par_text(el).strip():
                continue
            for inl in el.iter(WP + "inline"):
                ext = inl.find(WP + "extent")
                try:
                    w, h = int(ext.get("cx")), int(ext.get("cy"))
                except (TypeError, ValueError):
                    continue
                if w <= 0 or h <= 0 or w >= target * 0.98 or w >= 5 * h:   # 납작한 띠(쪽 장식)는 그대로
                    continue
                items.append((el, inl, ext, w, h, h * (target / w) - h))
    if not items:
        return 0
    need = sum(x[5] for x in items)
    slack = max(0.0, slack_pt * 12700)
    k = min(1.0, slack / need) if need > 0 else 0.0
    n = 0
    for p, inl, ext, w, h, extra in items:
        f = 1.0 + k * (target / w - 1.0)
        if f <= 1.001:
            continue
        nw, nh = int(w * f), int(h * f)
        ext.set("cx", str(nw)); ext.set("cy", str(nh))
        for x in inl.iter(A + "ext"):
            if x.getparent().tag == A + "xfrm":
                x.set("cx", str(nw)); x.set("cy", str(nh))
        ppr = p.find(qn("w:pPr"))
        if ppr is None:
            ppr = p.makeelement(qn("w:pPr"), {})
            p.insert(0, ppr)
        for e in ppr.findall(qn("w:ind")):
            ppr.remove(e)
        jc = ppr.find(qn("w:jc"))
        if jc is None:
            jc = ppr.makeelement(qn("w:jc"), {})
            ppr.append(jc)
        jc.set(qn("w:val"), "center")
        n += 1
    return n


# ── 표 칸이 늘어날 높이 어림 ──
# 줄 높이를 '최소' 로 풀어 주면(loosen_row_heights) 글이 넘치는 칸이 늘어나 쪽이 넘친다(고산 32→37쪽).
# → 쪽을 줄이기 전에 칸마다 '이 글이 Pretendard 로 몇 줄·몇 pt 필요한가' 를 어림해 늘어날 몫만큼 더 줄인다.
CELL_MIN_PITCH = float(os.environ.get("CF_CELL_MIN_PITCH", "1.35"))   # 칸·상자 안 줄 간격 최소(글자 크기 배)
_BUMP = [True]   # 어림에 위 최소 간격을 넣을지(_row_growth 가 '넣은 것 − 뺀 것' 을 따로 셀 때 끔)
LINE_EM = 1.5625             # 워드가 실제로 그리는 Pretendard 한 줄 높이 = 글자 크기 × 1.5625
                             # (2026-10-12 시험 문서로 잰 값: 8pt→12.5, 9pt→14.06, 10pt→15.6. 글꼴 파일 값 1.19 로 셈하면 모자라 잘렸다)
WIDE_EM, NARROW_EM = 0.95, 0.55
GROW_CAP = 10.0   # 쪽 전체 묶음은 안 씀 — 표마다 TABLE_GROW_MAX 로 묶는다
TABLE_GROW_MAX = 0.25
NEED_SAFETY = 1.15          # 어림 여유 — 쪽을 줄이면 글자 크기·칸 폭이 반올림으로 조금 다르게 줄어 줄이 하나 더 꺾일 수 있다
                             # (고산 6쪽 '인출선행조건': 줄이기 전 81.8/89.2 맞음 → 줄인 뒤 잘림). 최소 높이라 넉넉해도 손해 없음


_WIDTHS = None


def _widths():
    """Pretendard 글자 폭 표(글자 크기 1 기준) — 양식\\pretendard_widths.json (글꼴 파일에서 뽑음, 웹에도 같이 올림).
    ★예전엔 한글 0.95·영숫자 0.55 로 어림 → 한글은 실제 0.864 라 너무 크게, 숫자는 실제 ≈0.59 라 작게 셌다."""
    global _WIDTHS
    if _WIDTHS is None:
        import json
        try:
            with open(os.path.join(os.path.dirname(os.path.abspath(__file__)), "양식", "pretendard_widths.json"),
                      encoding="utf-8") as f:
                _WIDTHS = json.load(f)
        except Exception:
            _WIDTHS = {"light": {"hangul": 0.864, "w": {}}, "bold": {"hangul": 0.864, "w": {}}}
    return _WIDTHS


def _char_w(ch, tbl):
    if "가" <= ch <= "힣":
        return tbl["hangul"]
    w = tbl["w"].get(ch)
    if w is not None:
        return w
    return 1.0 if ord(ch) > 0x2E80 else NARROW_EM


def _para_height(p, width_pt, last=False):
    from docx.oxml.ns import qn
    sz, segs = None, [0.0]                                 # segs = 줄바꿈(w:br)으로 나뉜 토막별 글 길이(pt)
    for r in p.iter(qn("w:r")):
        v = r.find(f"{qn('w:rPr')}/{qn('w:sz')}")
        rs = int(v.get(qn("w:val"))) / 2 if v is not None and (v.get(qn("w:val")) or "").isdigit() else 10.0
        sz = rs if sz is None else max(sz, rs)
        b = r.find(f"{qn('w:rPr')}/{qn('w:b')}")
        tbl = _widths()["bold" if b is not None and b.get(qn("w:val")) not in ("0", "false") else "light"]
        for ch_el in r:
            if ch_el.tag == qn("w:br") and ch_el.get(qn("w:type")) in (None, "textWrapping"):
                segs.append(0.0)
            elif ch_el.tag == qn("w:t"):
                for ch in ch_el.text or "":
                    segs[-1] += rs * _char_w(ch, tbl)
    sp = p.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
    get = (lambda a: int(sp.get(qn("w:" + a))) if sp is not None and (sp.get(qn("w:" + a)) or "").lstrip("-").isdigit() else None)
    sz = sz or 10.0
    lines = sum(max(1, -(-c // max(width_pt, 1))) for c in segs)
    line, rule = get("line"), (sp.get(qn("w:lineRule")) if sp is not None else None) or "auto"
    one = sz * LINE_EM
    if line:
        one = line / 20 if rule == "exact" else max(one, line / 20) if rule == "atLeast" else one * line / 240
        if rule == "auto" and line < 240 and _BUMP[0]:
            # ★엔진(fix_compressed_line_spacing)이 칸·상자 안 1 미만 배수를 고정 간격(최소 1.35 배)으로 올린다
            #   → 어림도 같게(안 그러면 쪽을 덜 줄여 표 마지막 줄이 다음 쪽으로 넘쳤다 — 돈암 10쪽, 2026-10-20). 본문 문단은 그대로.
            if any(a.tag in (qn("w:tc"), qn("w:txbxContent")) for a in p.iterancestors()):
                one = max(one, sz * CELL_MIN_PITCH)
    return lines * one + (get("before") or 0) / 20 + (0 if last else (get("after") or 0) / 20)


def _row_h(tr):
    """(줄 높이 pt, 규칙) — 높이 없으면 (0, None)."""
    from docx.oxml.ns import qn
    h = tr.find(f"{qn('w:trPr')}/{qn('w:trHeight')}")
    if h is None or not (h.get(qn("w:val")) or "").isdigit():
        return 0.0, None
    return int(h.get(qn("w:val"))) / 20, h.get(qn("w:hRule"))


def _row_need(tr, safety=None):
    """이 줄의 글·안쪽 표가 Pretendard 로 차지할 높이(pt) 어림 — 칸마다 문단 + 안쪽 표 줄 높이 합 중 큰 것.
    safety: 어림 여유 배수(기본 NEED_SAFETY — '최소 높이로 풀지' 정할 때). 쪽 줄이기 계산엔 1.0."""
    safety = NEED_SAFETY if safety is None else safety
    from docx.oxml.ns import qn
    need = 0.0
    for tc in tr.findall(qn("w:tc")):
        w = tc.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
        # 칸 안쪽 여백은 그 칸에 적힌 값(pdf2docx 는 대개 0) — 없을 때만 워드 기본 5.4pt×2.
        # ★예전엔 늘 11pt 를 뺐다 → 좁은 칸(30pt)에서 한 글자씩 꺾인다고 어림해 쪽을 너무 줄였다(돈암동 14→15쪽).
        mar = tc.find(f"{qn('w:tcPr')}/{qn('w:tcMar')}")
        pad, vpad = 10.8, 0.0
        if mar is not None:
            vals = [int(m.get(qn("w:w"))) / 20 for m in mar if (m.get(qn("w:w")) or "").isdigit()
                    and m.tag in (qn("w:start"), qn("w:end"), qn("w:left"), qn("w:right"))]
            pad = sum(vals) if vals else pad
            # 칸 위·아래 여백도 줄 높이에 얹힌다(예전엔 안 세 빽빽한 표 마지막 줄이 넘쳤다 — 내곡 43쪽, 2026-10-20)
            vpad = sum(int(m.get(qn("w:w"))) / 20 for m in mar if (m.get(qn("w:w")) or "").isdigit()
                       and m.tag in (qn("w:top"), qn("w:bottom")))
        width = int(w.get(qn("w:w"))) / 20 - pad if w is not None and (w.get(qn("w:w")) or "").isdigit() else 200
        # ★칸 끝의 빈 문단·마지막 문단 뒤 간격은 잘려도 안 보인다 → 세지 않는다.
        #   (돈암동: 이걸 세서 148줄이 '넘친다' 고 풀었더니 실제로 칸이 늘어 쪽이 넘쳤다. 화면엔 잘린 글이 없었다)
        kids = [ch for ch in tc if ch.tag in (qn("w:p"), qn("w:tbl"))]
        while kids and kids[-1].tag == qn("w:p") and not "".join(t.text or "" for t in kids[-1].iter(qn("w:t"))).strip() \
                and kids[-1].find(".//" + qn("w:drawing")) is None:
            kids.pop()
        hh = vpad if kids else 0.0
        for i, ch in enumerate(kids):
            if ch.tag == qn("w:p"):
                hh += _para_height(ch, width, last=(i == len(kids) - 1))
            else:
                hh += sum(_row_actual(r, safety) for r in ch.findall(qn("w:tr")))
        need = max(need, hh)
    return need * safety


_LOOSE = [False]   # True: 글 든 '고정' 줄도 '최소' 로 셈(엔진 loosen_exact_rows 가 결국 다 풀어 준다)


def _row_actual(tr, safety=None):
    h, rule = _row_h(tr)
    if rule == "exact" and not (_LOOSE[0] and any((t.text or "").strip() for t in tr.iter(
            "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t"))):
        return h
    return max(h, _row_need(tr, safety))


def _row_growth(el):
    """el 이 '표 한 줄' 이면: 그 줄이 늘어날 높이(여유 포함, 풀지 말지 정할 때).
    el 이 표·문단 덩어리면: 맨 바깥 표마다 '늘어날 높이' 합 — 표마다 원래 높이의 TABLE_GROW_MAX 까지만.
    ★워드로 잰 돈암동 표 7개(2026-10-12): 실제 늘어남은 0~+23%(520→640)인데 어림은 칸 합치기·안쪽 표 때문에
    +70%(483→819)까지 넘쳤다 → 표 하나당 25% 로 묶음."""
    from docx.oxml.ns import qn
    if el.tag == qn("w:tr"):
        h, rule = _row_h(el)
        return max(0.0, _row_need(el) - h) if rule == "atLeast" else 0.0
    tbls = [el] if el.tag == qn("w:tbl") else [t for t in el.iter(qn("w:tbl"))
                                               if not any(a.tag == qn("w:tbl") for a in t.iterancestors())]
    total = 0.0
    for t in tbls:
        rows = t.findall(qn("w:tr"))
        orig = sum(_row_h(tr)[0] for tr in rows)
        # ★글 든 '고정' 줄은 엔진(loosen_exact_rows)이 '최소' 로 풀어 늘어난다 → 어림도 풀린 것으로
        #   (안 그러면 늘 0 으로 셌고, 빽빽한 표 쪽 끝 각주가 다음 쪽으로 넘쳤다 — 효자 16·22쪽, 2026-10-20)
        _LOOSE[0] = True
        try:
            bumped = [_row_need(tr, 1.0) for tr in rows]
            _BUMP[0] = False
            try:
                flat = [_row_need(tr, 1.0) for tr in rows]
                plain = sum(_row_actual(tr, 1.0) for tr in rows)   # 칸 최소 간격 빼고 어림
            finally:
                _BUMP[0] = True
        finally:
            _LOOSE[0] = False
        # ★줄 간격 1 미만 → 고정 간격(최소 1.35배)으로 늘어나는 몫은 어림이 아니라 정해진 값 → 25% 묶음 밖에서,
        #   줄 높이가 넉넉해 보여도 그대로 더한다(묶음 안에 넣거나 '넉넉하면 0' 으로 셌더니 빽빽한 표 쪽 끝 줄·각주가
        #   다음 쪽으로 넘쳤다 — 돈암 4·9쪽, 효자 16·22쪽, 2026-10-20). 조금 더 줄이는 쪽이 안전.
        bump = sum(max(0.0, x - y) for x, y in zip(bumped, flat))
        total += min(max(0.0, plain - orig), orig * TABLE_GROW_MAX) + bump
    return total



# ── 같은 비율로 줄이기 ──
_TWIPS = {"spacing": ("before", "after"), "ind": ("left", "right", "firstLine", "hanging", "start", "end"),
          "tab": ("pos",), "tblW": ("w",), "tcW": ("w",), "gridCol": ("w",), "tblInd": ("w",),
          "trHeight": ("val",), "tblCellSpacing": ("w",)}
_HALF = {"sz", "szCs", "kern", "position"}
_EMU = {"extent": ("cx", "cy"), "ext": ("cx", "cy"), "off": ("x", "y"),
        "effectExtent": ("l", "t", "r", "b")}


_NOTE_RE = re.compile(r"^\s*(※|\*|＊|주\s*\d|\(\s*단위|\(?\s*source|\(?\s*출처|자료\s*:)", re.I)


def _sz_target(v, s, el):
    """글자 크기(반 pt) 정하기 — 2026-10-14 사용자 확정.
    · 본문(원본 8~10.5pt) → **9pt**
    · 참고·각주·단위·출처 글(※, *, 주1), (단위…), Source…)과 원본 8pt 미만 → 원본 크기 그대로
    · 제목(원본 11pt 이상) → 줄이더라도 11pt 아래로는 안 내림"""
    W_ = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    if v < 16:
        return v
    p = next((a for a in el.iterancestors() if a.tag == W_ + "p"), None)
    txt = "".join(t.text or "" for t in p.iter(W_ + "t")) if p is not None else ""
    if _NOTE_RE.match(txt):
        return v
    if v <= 21:
        return MIN_SZ
    return max(int(round(v * s)), 22)


def scale_xml(root, s, old_left=0.0, sw=None):
    """root 안의 길이를 줄인다: 세로·크기(글자·간격·줄 높이·그림)는 s 배, 가로(표·칸 폭·들여쓰기·탭)는 sw 배(없으면 s)."""
    sw = s if sw is None else sw
    if s >= 0.999 and sw >= 0.999:
        W_ = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
        for el in root.iter(W_ + "sz", W_ + "szCs"):          # 줄이지 않는 쪽도 본문 글자는 9pt 로
            v = el.get(W_ + "val")
            if v and v.isdigit() and el.getparent().tag == W_ + "rPr":
                el.set(W_ + "val", str(_sz_target(int(v), 1.0, el)))
        return
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    def mul(el, attr, ns=W, minimum=None, f=None):
        k = ns + attr if ns else attr
        v = el.get(k)
        if v is None or not v.lstrip("-").isdigit():
            return
        nv = int(round(int(v) * (s if f is None else f)))
        if minimum is not None:
            nv = max(minimum, nv)
        el.set(k, str(nv))

    for el in root.iter():
        tag = el.tag
        if not isinstance(tag, str):
            continue
        if tag.startswith(W):
            name = tag[len(W):]
            if name in _HALF:
                v = el.get(W + "val")
                if name in ("sz", "szCs") and v and v.isdigit():
                    el.set(W + "val", str(_sz_target(int(v), s, el)))
                else:
                    mul(el, "val", minimum=2)
            elif name == "spacing":
                if el.getparent().tag == W + "rPr":
                    mul(el, "val")
                else:
                    mul(el, "before")
                    mul(el, "after")
                    if el.get(W + "lineRule") in ("exact", "atLeast"):
                        mul(el, "line", minimum=20)
            elif name in _TWIPS and name != "spacing":
                if name in ("tblW", "tcW") and el.get(W + "type") not in (None, "dxa"):
                    continue
                for a in _TWIPS[name]:
                    mul(el, a, f=(s if name == "trHeight" else sw))
            elif name == "tblpPr":
                for a in ("tblpX", "tblpY", "leftFromText", "rightFromText", "topFromText", "bottomFromText"):
                    mul(el, a, f=(sw if a in ("tblpX", "leftFromText", "rightFromText") else s))
            elif name in ("top", "left", "bottom", "right", "start", "end") and \
                    el.getparent() is not None and el.getparent().tag in (W + "tcMar", W + "tblCellMar"):
                mul(el, "w", f=(s if name in ("top", "bottom") else sw))
        elif tag.startswith(WP):
            name = tag[len(WP):]
            if name in _EMU:
                for a in _EMU[name]:
                    mul(el, a, ns="")
            elif name == "posOffset" and el.text and el.text.lstrip("-").isdigit():
                if el.getparent().get("relativeFrom") != "page":       # 쪽 기준 위치는 _place_anchors 가 맞춘다
                    f = sw if el.getparent().tag == WP + "positionH" else s
                    el.text = str(int(int(el.text) * f))
        elif tag.startswith(A):
            name = tag[len(A):]
            if name in ("ext", "off", "chExt", "chOff") and el.getparent() is not None \
                    and el.getparent().tag == A + "xfrm":
                for a in ("cx", "cy", "x", "y"):
                    mul(el, a, ns="")


def _place_anchors(root, s, old_left, old_top, sw=None):
    """쪽 기준으로 박힌 떠 있는 그림·도형 위치를 글과 같은 자리로 옮긴다.
    ★고산 18쪽(2026-10-12 사용자 지적): 글은 양식 여백(위 3cm)으로 내리고 원본 위 빈 자리를 깎았는데
    그림은 원본 쪽 위치(위에서 59.6pt) 그대로라 머리말 위로 올라가고 그림 설명 글과 겹쳤다.
    → 원본 y 를 '본문 시작(old_top) 에서 얼마 아래였나 × 비율 + 양식 위 여백' 으로(가로는 왼쪽 여백 기준)."""
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    for pos in root.iter(WP + "positionH", WP + "positionV"):
        if pos.get("relativeFrom") != "page":
            continue
        off = pos.find(WP + "posOffset")
        if off is None or not (off.text or "").lstrip("-").isdigit():
            continue
        v = int(off.text) / 12700
        if pos.tag == WP + "positionH":
            v = (v - old_left) * (s if sw is None else sw) + LEFT_CM / 2.54 * 72
        else:
            v = (v - old_top) * s + TOP_CM / 2.54 * 72
        off.text = str(int(max(0.0, v) * 12700))


def _empty(el):
    from docx.oxml.ns import qn
    return el.tag == qn("w:p") and not "".join(el.itertext()).strip() \
        and el.find(".//" + qn("w:drawing")) is None and el.find(qn("w:pPr") + "/" + qn("w:sectPr")) is None


def _anchor_only(el):
    """글 없이 떠 있는 그림(wp:anchor)만 든 문단인가."""
    from docx.oxml.ns import qn
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    # 글 판단은 w:t 만(itertext 는 그림 위치 숫자 posOffset 까지 '글' 로 센다)
    return el.tag == qn("w:p") and not "".join(t.text or "" for t in el.iter(qn("w:t"))).strip() \
        and el.find(".//" + WP + "anchor") is not None and el.find(".//" + WP + "inline") is None


def _flatten_lead_anchors(grp):
    """쪽 맨 위(첫 글·표·줄 안 그림이 나오기 전)의 '떠 있는 그림만 든 문단' 을 높이 1pt 로.
    그림은 쪽 기준 위치로 따로 놓이므로 이 문단 높이만큼 아래 글(그림 설명)이 밀려 그림과 겹쳤다(고산 18쪽)."""
    import docx
    import docx.enum.text
    from docx.shared import Pt
    from docx.oxml.ns import qn
    for el in grp:
        if el.tag != qn("w:p"):
            break
        if _anchor_only(el):
            pf = docx.text.paragraph.Paragraph(el, None).paragraph_format
            pf.space_before, pf.space_after = Pt(0), Pt(0)
            pf.line_spacing_rule = docx.enum.text.WD_LINE_SPACING.EXACTLY
            pf.line_spacing = Pt(1)
        elif not _empty(el):
            break


def trim_leading(grp, amount):
    """구역 맨 위의 빈 자리(빈 문단 높이·문단 앞 간격)를 amount(pt) 만큼 깎는다."""
    import docx
    import docx.enum.text
    from docx.shared import Pt
    from docx.oxml.ns import qn
    for el in list(grp):
        if el.tag != qn("w:p"):
            break
        pf = docx.text.paragraph.Paragraph(el, None).paragraph_format
        if _anchor_only(el):
            # 떠 있는 그림만 든 문단은 글 흐름에서 자리만 차지한다(그림은 쪽 기준 위치로 따로 놓임).
            # 맨 위에 있으면 그 높이만큼 아래 글(그림 설명)이 밀려 그림과 겹쳤다(고산 18쪽) → 깎을 몫과 상관없이 납작하게.
            continue                                         # _flatten_lead_anchors 가 이미 납작하게 했다
        if amount <= 0.5:
            break
        before = pf.space_before.pt if pf.space_before is not None else 0
        cut = min(before, amount)
        pf.space_before = Pt(before - cut)
        amount -= cut
        if not _empty(el):
            break
        line = pf.line_spacing.pt if hasattr(pf.line_spacing, "pt") else 12.0
        after = pf.space_after.pt if pf.space_after is not None else 0
        if line + after <= amount:
            el.getparent().remove(el)
            grp.remove(el)
            amount -= line + after
        else:
            pf.line_spacing = Pt(max(1.0, line - amount))
            amount = 0


def trim_trailing(grp):
    """구역 끝의 빈 문단(pdf2docx 가 쪽 바닥까지 채운 것)을 지우고, 구역 끝 문단은 1pt 로 납작하게.

    ★둘 다 그대로 두면 줄인 쪽 끝에서 다음 쪽으로 넘쳐 **거의 빈 쪽** 이 생겼다(등촌역 4쪽).
      원본 1쪽 = 결과 1쪽이라 다음 구역은 어차피 새 쪽에서 시작하므로 지워도 모양이 같다.
    """
    from docx.oxml.ns import qn
    if not grp:
        return
    last = grp[-1]
    ends = last.tag == qn("w:p") and last.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None
    for el in reversed(grp[:-1] if ends else list(grp)):
        if _empty(el):
            el.getparent().remove(el)
            grp.remove(el)
        else:
            break
    if ends:
        _flatten_holder(last)


def _flatten_holder(p):
    """구역 끝 문단을 높이 1pt 짜리로(글자 1pt·줄 간격 고정 1pt·앞뒤 0)."""
    from docx.oxml.ns import qn
    from lxml import etree
    for r in list(p.findall(qn("w:r"))):
        if not "".join(r.itertext()).strip():
            p.remove(r)
    ppr = p.find(qn("w:pPr"))
    sp = ppr.find(qn("w:spacing"))
    if sp is None:
        sp = ppr.makeelement(qn("w:spacing"), {})
        sect = ppr.find(qn("w:sectPr"))
        # spacing 은 rPr·sectPr 보다 앞
        anchor = ppr.find(qn("w:jc")) or ppr.find(qn("w:rPr")) or sect
        anchor.addprevious(sp) if anchor is not None else ppr.append(sp)
    # ★줄 간격을 '고정 1pt'(exact) 로 하면 바로 앞 문단의 줄 안 그림이 줄 높이 0 으로 눌려 아래 끝만 보였다
    #   (고산 18쪽 두 번째 지도, 2026-10-12 워드로 시험: exact→잘림, auto→정상). 글자 1pt + 자동 줄 = 약 1.6pt.
    for k, v in (("before", "0"), ("after", "0"), ("line", "240"), ("lineRule", "auto")):
        sp.set(qn("w:" + k), v)
    rpr = ppr.find(qn("w:rPr"))
    if rpr is None:
        rpr = ppr.makeelement(qn("w:rPr"), {})
        ppr.find(qn("w:sectPr")).addprevious(rpr)
    for tag in ("w:sz", "w:szCs"):
        e = rpr.find(qn(tag))
        if e is None:
            e = rpr.makeelement(qn(tag), {})
            rpr.append(e)
        e.set(qn("w:val"), "2")


def _insert_picture_page(d, idx, png, box):
    """idx 번째 구역 자리에 원본 쪽 그림 한 장짜리 구역을 넣는다. 쓰인 높이(pt)를 돌려준다."""
    import copy
    from docx.shared import Pt
    from docx.oxml.ns import qn
    s = min(1.0, BOX_W_PT / box.width, BOX_H_PT / box.height)
    w, h = box.width * s, box.height * s
    body = d.element.body
    p = d.add_paragraph()
    p.add_run().add_picture(io.BytesIO(png), width=Pt(w), height=Pt(h))
    holder = d.add_paragraph()
    sp = copy.deepcopy(body.find(qn("w:sectPr")))
    for e in sp.findall(qn("w:headerReference")) + sp.findall(qn("w:footerReference")):
        sp.remove(e)
    holder._p.get_or_add_pPr().append(sp)
    body.remove(p._p)
    body.remove(holder._p)
    groups = [g for g in _section_groups(d) if g]
    if idx < len(groups):
        # idx 번째 구역 앞에 [그림 문단 + 구역 끝] 을 끼운다 → 그림만 든 구역이 하나 생긴다
        groups[idx][0].addprevious(p._p)
        groups[idx][0].addprevious(holder._p)
    else:
        # 맨 끝: 앞 구역을 닫는 '구역 끝' 문단을 먼저 두고, 그림은 마지막 구역(본문 sectPr)에 넣는다
        last = body.find(qn("w:sectPr"))
        holder._p.find(qn("w:pPr")).remove(holder._p.find(qn("w:pPr")).find(qn("w:sectPr")))
        holder._p.get_or_add_pPr().append(copy.deepcopy(last))
        last.addprevious(holder._p)
        last.addprevious(p._p)
    _set_geometry(d.sections[idx])
    return h


# ──────────────────────────────────────────────
# 금융 구조도 기본값(원본에서 읽기) — 화면 2단계 칸을 미리 채운다(사람이 고친다)
# ──────────────────────────────────────────────
_DIAG_LABELS = [
    ("차주", re.compile(r"^차\s*주$")),
    ("신탁", re.compile(r"^(신\s*탁\s*사|신탁기관|신탁회사)$")),
    ("시공사", re.compile(r"^(시\s*공\s*사|시공회사)$")),
    ("금액", re.compile(r"^(대출\s*금액|대출\s*규모|대출금|총\s*대출금액)$")),
    ("LTV", re.compile(r"^LTV$", re.I)),
    ("금리", re.compile(r"^(대출\s*금리|금\s*리|적용\s*금리)$")),
    ("연대보증", re.compile(r"^(연대\s*보증(인)?|보증인)$")),
]


def diagram_defaults(pdf_bytes, max_pages=10):
    """원본 앞쪽 표의 '구분 | 내용' 줄에서 차주·신탁사·대출금액·LTV·금리·연대보증을 읽는다(못 찾으면 빈칸)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    found = {}
    for page in list(doc)[1:max_pages]:
        words = page.get_text("words")
        rows = {}
        for w in words:
            rows.setdefault(round((w[1] + w[3]) / 2 / 3), []).append(w)
        for key in sorted(rows):
            ws = sorted(rows[key], key=lambda w: w[0])
            # 맨 앞 1~3 낱말을 붙여 구분 이름인지 본다
            for k in (1, 2, 3):
                label = " ".join(w[4] for w in ws[:k])
                for name, rx in _DIAG_LABELS:
                    if name in found or not rx.match(label.replace(" ", "") if name != "LTV" else label):
                        continue
                    val = " ".join(w[4] for w in ws[k:]).strip(" :·")
                    heads = len(re.findall(r"대출기간|금리|수수료|LTV|구분|내용|비고", val))
                    if val and len(val) < 60 and heads < 2:          # 표 머리글 줄('대출기간 금리 수수료 LTV')은 값이 아니다
                        found[name] = val
                break
    if "금리" in found:
        found["금리"] = re.sub(r"^연\s*", "", found["금리"])
    return {"차주": found.get("차주", ""), "신탁": found.get("신탁", ""), "연대보증": found.get("연대보증", ""),
            "시공사": found.get("시공사", ""),
            "본건": {"금액": found.get("금액", ""), "LTV": found.get("LTV", ""), "금리": found.get("금리", ""), "담보": ""},
            "브릿지": ""}
