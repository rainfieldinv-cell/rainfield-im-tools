# -*- coding: utf-8 -*-
"""워드 본문에서 '쪽이 어디서 바뀌는지' 어림한다 — 원본 표지·끝쪽을 찾기 위해.

워드 파일에는 쪽 번호가 적혀 있지 않다. 쪽은 워드가 화면에 그릴 때 정해진다.
웹(리눅스)에는 워드가 없으므로 세 가지를 차례로 쓴다.
  1. 쪽 나눔 / 구역 나눔 / '앞에서 쪽 나누기' 설정 — 확실한 경계
  2. lastRenderedPageBreak — 워드가 저장할 때 남긴 '여기서 쪽이 바뀌었다' 표시(있는 문서만)
  3. 줄 높이 어림 — 빈 줄로 표지를 채워 자연스럽게 넘어가는 문서(실제: 신사동 평화빌딩)
어림은 틀릴 수 있어 화면에서 사람이 고칠 수 있게 한다.
"""
import re

from docx.oxml.ns import qn

FB = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
BLOCK_TAGS = (qn("w:p"), qn("w:tbl"), qn("w:sdt"))


def blocks(body):
    return [e for e in body if e.tag in BLOCK_TAGS]


def text_of(el):
    """진짜 글자만(글상자 대체 내용·위치 숫자 빼고)."""
    return "".join(t.text or "" for t in el.iter(qn("w:t"))
                   if not any(a.tag == FB for a in t.iterancestors()))


def _on(el, tag):
    if el is None:
        return None
    b = el.find(qn(tag))
    if b is None:
        return None
    return b.get(qn("w:val"), "true").lower() not in ("0", "false", "off")


def hard_break_after(el):
    """이 요소 '뒤에서' 쪽이 확실히 바뀌나(쪽 나눔·구역 나눔)."""
    if el.tag != qn("w:p"):
        return False
    ppr = el.find(qn("w:pPr"))
    if ppr is not None:
        sp = ppr.find(qn("w:sectPr"))
        if sp is not None:
            t = sp.find(qn("w:type"))
            if t is None or t.get(qn("w:val")) != "continuous":
                return True
    return any(b.get(qn("w:type")) == "page" for b in el.iter(qn("w:br")))


def starts_page(el):
    """이 요소 '앞에서' 쪽이 바뀌나(앞에서 쪽 나누기 설정 · 워드가 남긴 표시)."""
    first_p = el if el.tag == qn("w:p") else next(el.iter(qn("w:p")), None)
    if first_p is None:
        return False
    ppr = first_p.find(qn("w:pPr"))
    if ppr is not None and _on(ppr, "w:pageBreakBefore"):
        return True
    for r in first_p.iter(qn("w:r")):
        if r.find(qn("w:lastRenderedPageBreak")) is not None:
            return True
        if text_of(r).strip():
            return False
    return False


class _Sizer:
    """문단·표 높이를 pt 로 어림한다."""

    def __init__(self, doc):
        self.doc = doc
        st = doc.styles.element
        self.def_sz = 10.0
        self.def_before = self.def_after = 0.0
        self.def_line = 1.0
        dd = st.find(qn("w:docDefaults"))
        if dd is not None:
            s = dd.find(".//" + qn("w:rPrDefault") + "//" + qn("w:sz"))
            if s is not None:
                self.def_sz = int(s.get(qn("w:val"))) / 2
            sp = dd.find(".//" + qn("w:pPrDefault") + "//" + qn("w:spacing"))
            self._spacing(sp, self)
        self.style = {}
        for s in st.findall(qn("w:style")):
            sid = s.get(qn("w:styleId"))
            info = {}
            sz = s.find(qn("w:rPr") + "/" + qn("w:sz"))
            if sz is not None:
                info["sz"] = int(sz.get(qn("w:val"))) / 2
            sp = s.find(qn("w:pPr") + "/" + qn("w:spacing"))
            if sp is not None:
                info["sp"] = sp
            based = s.find(qn("w:basedOn"))
            info["based"] = based.get(qn("w:val")) if based is not None else None
            self.style[sid] = info
        sec = doc.sections[0]
        self.page_h = sec.page_height.pt - sec.top_margin.pt - sec.bottom_margin.pt
        self.page_w = sec.page_width.pt - sec.left_margin.pt - sec.right_margin.pt

    @staticmethod
    def _spacing(sp, obj):
        if sp is None:
            return
        b, a, l, rule = (sp.get(qn("w:" + k)) for k in ("before", "after", "line", "lineRule"))
        if b and b.lstrip("-").isdigit():
            obj.def_before = int(b) / 20
        if a and a.lstrip("-").isdigit():
            obj.def_after = int(a) / 20
        if l and l.isdigit():
            obj.def_line = int(l) / 240 if rule in (None, "auto") else -int(l) / 20   # 음수 = 고정 pt

    def _style_val(self, sid, key):
        seen = set()
        while sid in self.style and sid not in seen:
            seen.add(sid)
            if key in self.style[sid]:
                return self.style[sid][key]
            sid = self.style[sid]["based"]
        return None

    def para(self, p, width=None):
        width = width or self.page_w
        ppr = p.find(qn("w:pPr"))
        sid = None
        if ppr is not None and ppr.find(qn("w:pStyle")) is not None:
            sid = ppr.find(qn("w:pStyle")).get(qn("w:val"))
        sizes = [int(s.get(qn("w:val"))) / 2 for s in p.iter(qn("w:sz"))
                 if (s.get(qn("w:val")) or "").isdigit() and not any(a.tag == FB for a in s.iterancestors())]
        size = max(sizes) if sizes else (self._style_val(sid, "sz") or self.def_sz)

        class S:
            pass
        sp = S()
        sp.def_before, sp.def_after, sp.def_line = self.def_before, self.def_after, self.def_line
        self._spacing(self._style_val(sid, "sp"), sp)
        if ppr is not None:
            self._spacing(ppr.find(qn("w:spacing")), sp)
        line = size * 1.2 * sp.def_line if sp.def_line > 0 else -sp.def_line
        txt = text_of(p)
        n = max(1, -(-int(sum(1.0 if ord(c) > 0x2E80 else 0.55 for c in txt) * size) // max(int(width), 1)))
        h = line * n + sp.def_before + sp.def_after
        for inl in p.iter(WP + "inline"):                 # 글 속에 든 그림은 높이를 더한다
            ext = inl.find(WP + "extent")
            if ext is not None:
                h += int(ext.get("cy")) / 12700
        return h

    def table(self, tbl):
        total = 0.0
        for tr in tbl.findall(qn("w:tr")):
            trh = tr.find(qn("w:trPr") + "/" + qn("w:trHeight"))
            fixed = int(trh.get(qn("w:val"))) / 20 if trh is not None and (trh.get(qn("w:val")) or "").isdigit() else 0
            cells = tr.findall(qn("w:tc"))
            cw = self.page_w / max(len(cells), 1)
            best = 0.0
            for tc in cells:
                best = max(best, sum(self.para(p, cw) for p in tc.findall(qn("w:p"))))
            total += max(best, fixed)
        return total

    def block(self, el):
        if el.tag == qn("w:tbl"):
            return self.table(el)
        if el.tag == qn("w:p"):
            return self.para(el)
        return sum(self.para(p) for p in el.iter(qn("w:p")))


def _sure_end(bl, i):
    """블록 i 뒤에서 쪽이 확실히 바뀌나."""
    return hard_break_after(bl[i]) or (i + 1 < len(bl) and starts_page(bl[i + 1]))


def cover_end(doc, limit=80):
    """원본 표지가 끝나는 블록 번호(그 블록까지가 표지)와 근거. 못 찾으면 (None, 이유).

    줄 높이 어림은 ±2블록쯤 틀린다(실측: 신사동·엘피스). 그래서 어림한 자리 근처에
      ① 확실한 쪽 표시가 있으면 그것을 쓰고
      ② 없으면 '첫 본문 덩어리(글 100자 이상)' 바로 앞 제목을 2쪽의 시작으로 본다
         — 표지 뒤 첫 쪽은 거의 늘 '제목 + 본문' 으로 시작한다.
    """
    bl = blocks(doc.element.body)[:limit]
    sizer = _Sizer(doc)
    used, est = 0.0, None
    for i, el in enumerate(bl):
        if i > 0 and starts_page(el):
            return i - 1, "표시"
        h = sizer.block(el)
        if i > 0 and used + h > sizer.page_h:
            est = i - 1
            break
        used += h
        if hard_break_after(el):
            return i, "쪽나눔"
    if est is None:
        return None, "못 찾음"
    # ② 를 먼저 본다 — 근처의 '쪽 표시' 는 2쪽이 아니라 3쪽 시작일 때가 있다(실측: 신사동)
    s = next((k for k in range(max(0, est - 4), min(len(bl), est + 6))
              if len(text_of(bl[k]).strip()) >= 100), None)
    if s is not None:
        h = s - 1
        while h > 0 and not text_of(bl[h]).strip():
            h -= 1
        if h >= 1 and abs((h - 1) - est) <= 3:
            return h - 1, "어림+제목"
    for j in range(max(0, est - 1), min(len(bl), est + 4)):          # ①
        if _sure_end(bl, j):
            return j, "어림+표시"
    return est, "어림"


def last_page_start(doc):
    """마지막 쪽이 시작되는 블록 번호(확실한 표시가 있을 때만)."""
    bl = blocks(doc.element.body)
    for j in range(len(bl) - 1, 0, -1):
        if starts_page(bl[j]):
            return j
        if hard_break_after(bl[j - 1]):
            return j
    return None
