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
TOP_CM, BOTTOM_CM, LEFT_CM, RIGHT_CM = 3.0, 2.5, 2.5, 2.5
HEADER_CM, FOOTER_CM = 1.7, 1.2
BOX_W_PT = (PAGE_W - LEFT_CM - RIGHT_CM) / 2.54 * 72        # 본문 칸 가로(pt)
BOX_H_PT = (PAGE_H - TOP_CM - BOTTOM_CM) / 2.54 * 72        # 본문 칸 세로(pt)
HEIGHT_SAFETY = 0.92  # 글꼴이 Pretendard 로 바뀌면 줄이 조금 높아진다 → 세로는 8% 여유(5% 로는 등촌역 쪽 몇 개가 넘쳤다)

PAGE_NO_RE = re.compile(r"[-–\s]*(?:page\s*)?\d{1,3}(?:\s*(?:/|of)\s*\d{1,3})?[-–\s]*", re.I)
HIGHLIGHT_RE = re.compile(r"executive\s*summary|investment\s*highlight|투자\s*하이라이트|투자\s*포인트|"
                          r"핵심\s*요약|하이라이트|highlights?", re.I)


# ──────────────────────────────────────────────
# 원본 살펴보기
# ──────────────────────────────────────────────
def text_per_page(pdf_bytes):
    """쪽당 글자 수(글자를 그림으로 바꾼 PDF 를 미리 알아보려고)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    n = sum(len(p.get_text().strip()) for p in doc)
    return n / max(doc.page_count, 1)


def last_page_is_contact(doc):
    t = doc[-1].get_text()
    return len(set(EMAIL_RE.findall(t))) + len(set(PHONE_RE.findall(t))) >= 2


def find_highlight_pages(pdf_bytes):
    """하이라이트(Executive Summary 등) 쪽 번호(1부터) 목록. 표지 다음 6쪽 안에서, 쪽 위쪽 글로 찾는다."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    hits = []
    for i in range(1, min(doc.page_count, 7)):
        top = doc[i].get_text("text", clip=fitz.Rect(0, 0, doc[i].rect.width, doc[i].rect.height * 0.35))
        if HIGHLIGHT_RE.search(top):
            hits.append(i + 1)
        elif hits:
            break
    return hits


BOND_NEED_PT = 270    # 사모사채 개요(제목 + 9줄 표)에 드는 높이 어림(약 9.5cm)


def bond_place(rep, highlight_pages):
    """사모사채 개요 자리 — 하이라이트 마지막 쪽 뒤(2026-10-08 사용자 확정).

    그 쪽에 개요가 들어갈 여백이 남으면 (k, True) = 이어서, 아니면 (k, False) = 새 쪽.
    하이라이트를 못 찾았으면 None = 표지 바로 다음 쪽.
    """
    if not highlight_pages:
        return None
    last = max(highlight_pages)
    for k, p in enumerate(rep.get("쪽", [])):
        if p["원본쪽"] == last:
            return (k, BOX_H_PT - p["높이"] >= BOND_NEED_PT)
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


def strip_band_logos(page):
    """위·아래 머리말 띠 안의 작은 그림(로고)을 지운다 — 한 쪽에만 있어도(되풀이 여부 상관없이)."""
    h, w = page.rect.height, page.rect.width
    n = 0
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"])
        in_band = r.y1 <= h * TOP_BAND or r.y0 >= h * (1 - BOTTOM_BAND)
        if in_band and r.width < w * 0.35 and r.height < h * 0.10:
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


def convert(pdf_bytes, drop_cover=True, drop_last=None, progress=None):
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
    rep["원본 끝쪽 뺌"], rep["원본 표지 뺌"] = bool(drop_last and n > 2), bool(drop_cover and n > 1)
    rep["그림 글씨 고침"] = fix_image_text(doc)
    leftover = []
    rep["머리말·꼬리말 지운 쪽"] = strip_repeating_bands(doc, leftover)
    # 원본 담당자 연락처(어느 쪽이든)·머리말 띠의 로고(한 쪽에만 있어도) — 2026-10-08 사용자 확정
    rep["연락처 지운 쪽"] = [i + 1 for i in keep if strip_contacts(doc[i])]
    rep["로고 지운 쪽"] = [i + 1 for i in keep if strip_band_logos(doc[i])]
    boxes = [content_box(doc[i]) for i in range(n)]
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
        if best is not None and best_r >= 0.6:
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
    rep["워드에서 지운 머리말 줄"] = drop_paragraphs(d, set(leftover))
    # fix_wrap_tabs(줄바꿈 탭→띄어쓰기)는 안 쓴다 — 목차·글머리 줄을 이어 붙여 원본을 망가뜨렸다(내곡동, 2026-10-12).
    # 대신 글자 폭을 원본과 같게 맞춰(match_char_width) 탭·줄바꿈이 원본 자리에 그대로 떨어지게 한다.
    rep["기호 글꼴 고침"] = fix_symbol_runs(d)
    rep["줄 간격 글꼴 맞춤"] = match_line_spacing(d, font_lh)
    rep["글자 폭 원본 맞춤"] = match_char_width(d, width_ratio)
    rep["글머리 앞 줄바꿈"] = break_before_bullets(d)
    rep["선만 있는 표 지움"] = drop_grid_tables(d)
    rep["쪼개진 표 합침"] = merge_split_tables(d)
    rep["흰 글씨 칸 바탕"] = fill_white_text_cells(d)
    rep["표 줄 높이 풀어줌"] = loosen_row_heights(d)
    rep["흰 바탕 투명하게"] = clear_white_overlays(d)
    done = [k for k in keep if k not in ignored]
    page_info = []
    for (secs, grps), k in zip(_pages(d), done):
        s, used = _fit_page(secs, grps, boxes[k])
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
                if seen and s[:1] in LINE_BULLETS:
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
        shd.set(qn("w:fill"), "243F60")
        n += 1
    return n


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


def _fit_page(secs, grps, box):
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
    s = min(1.0, BOX_W_PT / max(old_w, 1), BOX_H_PT * HEIGHT_SAFETY / max(box.height + grow, 1))
    for i, sec in enumerate(secs):
        _set_geometry(sec, new_page=(i == 0))
        cols = sec._sectPr.find(qn("w:cols"))               # 2단 같은 단 너비·간격도 같은 비율로
        if cols is not None and s < 0.999:
            for el in [cols] + cols.findall(qn("w:col")):
                for a in ("space", "w"):
                    v = el.get(qn("w:" + a))
                    if v and v.isdigit():
                        el.set(qn("w:" + a), str(int(int(v) * s)))
    for grp in grps:
        for el in grp:
            scale_xml(el, s, old_left)
            _place_anchors(el, s, old_left, box.y0)
    _flatten_lead_anchors(grps[0])
    lead = max(0.0, (box.y0 - old_top) * s)                 # 위 띠를 지운 자리 = 쓸데없는 빈 칸
    trim_leading(grps[0], lead)
    for grp in grps[:-1]:                                    # 쪽 안의 구역 끝(단 나눔) 문단도 납작하게
        if grp:
            _flatten_holder(grp[-1]) if grp[-1].find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None else None
    trim_trailing(grps[-1])
    return s, box.height * s


# ── 표 칸이 늘어날 높이 어림 ──
# 줄 높이를 '최소' 로 풀어 주면(loosen_row_heights) 글이 넘치는 칸이 늘어나 쪽이 넘친다(고산 32→37쪽).
# → 쪽을 줄이기 전에 칸마다 '이 글이 Pretendard 로 몇 줄·몇 pt 필요한가' 를 어림해 늘어날 몫만큼 더 줄인다.
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
        pad = 10.8
        if mar is not None:
            vals = [int(m.get(qn("w:w"))) / 20 for m in mar if (m.get(qn("w:w")) or "").isdigit()
                    and m.tag in (qn("w:start"), qn("w:end"), qn("w:left"), qn("w:right"))]
            pad = sum(vals) if vals else pad
        width = int(w.get(qn("w:w"))) / 20 - pad if w is not None and (w.get(qn("w:w")) or "").isdigit() else 200
        # ★칸 끝의 빈 문단·마지막 문단 뒤 간격은 잘려도 안 보인다 → 세지 않는다.
        #   (돈암동: 이걸 세서 148줄이 '넘친다' 고 풀었더니 실제로 칸이 늘어 쪽이 넘쳤다. 화면엔 잘린 글이 없었다)
        kids = [ch for ch in tc if ch.tag in (qn("w:p"), qn("w:tbl"))]
        while kids and kids[-1].tag == qn("w:p") and not "".join(t.text or "" for t in kids[-1].iter(qn("w:t"))).strip() \
                and kids[-1].find(".//" + qn("w:drawing")) is None:
            kids.pop()
        hh = 0.0
        for i, ch in enumerate(kids):
            if ch.tag == qn("w:p"):
                hh += _para_height(ch, width, last=(i == len(kids) - 1))
            else:
                hh += sum(_row_actual(r, safety) for r in ch.findall(qn("w:tr")))
        need = max(need, hh)
    return need * safety


def _row_actual(tr, safety=None):
    h, rule = _row_h(tr)
    return h if rule == "exact" else max(h, _row_need(tr, safety))


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
        est = sum(_row_actual(tr, 1.0) for tr in rows)
        total += min(max(0.0, est - orig), orig * TABLE_GROW_MAX)
    return total


# ── 같은 비율로 줄이기 ──
_TWIPS = {"spacing": ("before", "after"), "ind": ("left", "right", "firstLine", "hanging", "start", "end"),
          "tab": ("pos",), "tblW": ("w",), "tcW": ("w",), "gridCol": ("w",), "tblInd": ("w",),
          "trHeight": ("val",), "tblCellSpacing": ("w",)}
_HALF = {"sz", "szCs", "kern", "position"}
_EMU = {"extent": ("cx", "cy"), "ext": ("cx", "cy"), "off": ("x", "y"),
        "effectExtent": ("l", "t", "r", "b")}


def scale_xml(root, s, old_left=0.0):
    """root 안의 길이(글자 크기·간격·들여쓰기·표 너비·그림 크기·위치)를 모두 s 배로."""
    if s >= 0.999:
        return
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"

    def mul(el, attr, ns=W, minimum=None):
        k = ns + attr if ns else attr
        v = el.get(k)
        if v is None or not v.lstrip("-").isdigit():
            return
        nv = int(round(int(v) * s))
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
                    mul(el, a)
            elif name == "tblpPr":
                for a in ("tblpX", "tblpY", "leftFromText", "rightFromText", "topFromText", "bottomFromText"):
                    mul(el, a)
            elif name in ("top", "left", "bottom", "right", "start", "end") and \
                    el.getparent() is not None and el.getparent().tag in (W + "tcMar", W + "tblCellMar"):
                mul(el, "w")
        elif tag.startswith(WP):
            name = tag[len(WP):]
            if name in _EMU:
                for a in _EMU[name]:
                    mul(el, a, ns="")
            elif name == "posOffset" and el.text and el.text.lstrip("-").isdigit():
                if el.getparent().get("relativeFrom") != "page":       # 쪽 기준 위치는 _place_anchors 가 맞춘다
                    el.text = str(int(int(el.text) * s))
        elif tag.startswith(A):
            name = tag[len(A):]
            if name in ("ext", "off", "chExt", "chOff") and el.getparent() is not None \
                    and el.getparent().tag == A + "xfrm":
                for a in ("cx", "cy", "x", "y"):
                    mul(el, a, ns="")


def _place_anchors(root, s, old_left, old_top):
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
            v = (v - old_left) * s + LEFT_CM / 2.54 * 72
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
