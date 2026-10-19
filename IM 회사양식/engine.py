# -*- coding: utf-8 -*-
"""IM 회사양식 — 만들기 엔진.

원본 IM 워드(.docx)를 바탕으로 두고 회사 양식만 입힌다. 본문 내용·글씨 크기·쪽 방향은 건드리지 않는다.

  1. 글꼴        → Pretendard (보통 글씨 'Pretendard Light', 굵은 글씨 'Pretendard ExtraBold')
                    ※ 'Pretendard Light' 에 굵게를 걸면 워드가 가짜로 두껍게 그려서, 굵은 글씨는 일반 굵기 계열로
  2. 표 머리글   → 네이비(17365D) + 흰 글씨. 나머지 칸은 원본 그대로(붕어빵과 같은 규칙)
  3. 원본 표지   → 뺀다 (첫 '쪽 나눔' 앞까지)
     원본 끝쪽   → 증권사 연락처 쪽이면 뺀다 (마지막 '쪽 나눔' 뒤에 메일·전화번호가 있을 때)
  4. 회사 양식   → 맨 앞에 회사 표지 + 'Ⅰ. 사모사채 개요', 모든 쪽 머리말·꼬리말을 회사 것으로,
                    맨 뒤에 회사 연락처 쪽 + '끝'

양식은 사용자가 만든 `양식\\형식_1.docx` (2026-10-07 확정)에서 가져온다. 웹에는 딜 내용을 지운
`양식\\회사양식_틀.docx` 를 올린다(make_clean_template).
"""
import copy
import io
import os
import re
from datetime import date

from lxml import etree
from PIL import Image

import docx
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.oxml.ns import qn
from docx.shared import Pt

import pages

HERE = os.path.dirname(os.path.abspath(__file__))
# 웹·내 컴퓨터 모두 딜 내용을 지운 틀을 쓴다(같은 결과가 나오게). 형식_1 을 고쳤으면 양식\_틀만들기.py 를 돌릴 것.
TEMPLATE_CANDIDATES = [os.path.join(HERE, "양식", "회사양식_틀.docx"),
                       os.path.join(HERE, "양식", "형식_1.docx")]

NAVY = "17365D"           # 표 머리글 (사용자 양식 연락처 표와 같은 색)
LABEL = "F2F2F2"          # 사모사채 개요 구분 칸
FONT_LIGHT = "Pretendard Light"
FONT_BOLD = "Pretendard ExtraBold"   # 2026-10-19 사용자: 두꺼운 글씨 = Pretendard ExtraBold
SYMBOL_FONTS = {"wingdings", "wingdings 2", "wingdings 3", "symbol", "webdings", "marlett",
                "mt extra", "segoe ui symbol", "segoe ui emoji"}

BOND_ROWS = ["사모사채명", "사채유형", "발행인", "기초자산", "발행금액",
             "발행일", "만기일", "금융조건", "이자지급주기"]
COVER_PHOTO_RATIO = 15.92 / 7.85   # 양식 표지 사진 칸 가로:세로

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W,
      "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
      "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
      "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "v": "urn:schemas-microsoft-com:vml"}
R_EMBED = "{%s}embed" % NS["r"]
R_ID = "{%s}id" % NS["r"]

EMAIL_RE = re.compile(r"[\w.\-]+@[\w\-]+\.[\w.]+")
PHONE_RE = re.compile(r"0\d{1,2}[-)\s.]\d{3,4}[-\s.]\d{4}")


def template_path():
    for p in TEMPLATE_CANDIDATES:
        if os.path.exists(p):
            return p
    return None


# ★워드 서식 항목은 **순서가 정해져 있다.** 순서를 어기면 워드가 '파일이 손상되었습니다' 로 안 연다
#   (실제로 겪음 — rFonts 를 rStyle 앞에, shd 를 vAlign 뒤에 넣었을 때).
_ORDER = {
    "rPr": ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike",
            "outline", "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden",
            "color", "spacing", "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect",
            "bdr", "shd", "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
            "specVanish", "oMath", "rPrChange"],
    "tcPr": ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
             "textDirection", "tcFitText", "vAlign", "hideMark", "headers", "cellIns", "cellDel",
             "cellMerge", "tcPrChange"],
    "tblPr": ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
              "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd",
              "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription", "tblPrChange"],
    "pPr": ["pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl", "numPr",
            "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap",
            "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd",
            "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc",
            "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle",
            "rPr", "sectPr", "pPrChange"],
}


def _put(parent, child):
    """child 를 parent 안 '제자리' 에 넣는다(이미 같은 것이 있으면 바꿔 끼운다)."""
    order = _ORDER.get(etree.QName(parent).localname)
    name = etree.QName(child).localname
    old = parent.find(child.tag)
    if old is not None:
        parent.replace(old, child)
        return child
    if order and name in order:
        later = set(order[order.index(name) + 1:])
        for i, sib in enumerate(parent):
            if etree.QName(sib).localname in later:
                parent.insert(i, child)
                return child
    parent.append(child)
    return child


def _get(parent, name):
    """parent 의 w:name 자식(없으면 제자리에 새로 만든다)."""
    el = parent.find(qn("w:" + name))
    if el is None:
        el = _put(parent, parent.makeelement(qn("w:" + name), {}))
    return el


# ──────────────────────────────────────────────
# 1. 글꼴
# ──────────────────────────────────────────────
def _is_symbol(rfonts):
    """기호 글꼴인가 — pdf2docx 는 'Wingdings-Regular' 처럼 이름 뒤를 붙여 와서 이름 일부로 본다."""
    for k in ("ascii", "hAnsi", "eastAsia", "cs"):
        v = (rfonts.get(qn("w:" + k)) or "").strip().lower()
        if v and (v in SYMBOL_FONTS or re.search(r"wingding|webding|symbol|marlett", v)):
            return True
    return False


def _set_rfonts(rpr, family):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = _get(rpr, "rFonts")
    elif _is_symbol(rf):
        return
    for k in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
        rf.attrib.pop(qn("w:" + k), None)
    for k in ("ascii", "hAnsi", "eastAsia", "cs"):
        rf.set(qn("w:" + k), family)


def _on(el, tag):
    """<w:b/> 같은 켜고 끄는 서식이 켜져 있나."""
    if el is None:
        return None
    b = el.find(qn(tag))
    if b is None:
        return None
    return b.get(qn("w:val"), "true").lower() not in ("0", "false", "off")


def _style_bold(styles_el):
    """스타일 id → 굵게 여부(바탕 스타일까지 따라감)."""
    raw = {}
    for s in styles_el.findall(qn("w:style")):
        sid = s.get(qn("w:styleId"))
        based = s.find(qn("w:basedOn"))
        raw[sid] = (_on(s.find(qn("w:rPr")), "w:b"), based.get(qn("w:val")) if based is not None else None)
    out = {}
    for sid in raw:
        seen, cur, val = set(), sid, None
        while cur in raw and cur not in seen and val is None:
            seen.add(cur)
            val, cur = raw[cur][0], raw[cur][1]
        out[sid] = bool(val)
    return out


def _font_weight(rpr, bold):
    """굵은 글씨 = 'Pretendard ExtraBold' 글꼴 자체로(굵게 표시는 끔). ★굵게까지 켜 두면 워드가 ExtraBold 를
    한 번 더 가짜로 두껍게 그려 뭉개진다. 얇은 글씨 = 'Pretendard Light'."""
    _set_rfonts(rpr, FONT_BOLD if bold else FONT_LIGHT)
    if bold:
        for tag in ("w:b", "w:bCs"):
            _put(rpr, rpr.makeelement(qn(tag), {qn("w:val"): "0"}))


def apply_fonts(doc):
    styles_el = doc.styles.element
    bold_of = _style_bold(styles_el)
    # 기본값·스타일
    for rpr in styles_el.iter(qn("w:rPr")):
        st = rpr.getparent()
        while st is not None and st.tag != qn("w:style"):
            st = st.getparent()
        bold = bold_of.get(st.get(qn("w:styleId")), False) if st is not None else False
        _font_weight(rpr, bold)
    dd = styles_el.find(qn("w:docDefaults"))
    if dd is not None:
        rprd = dd.find(qn("w:rPrDefault"))
        if rprd is None:
            rprd = dd.makeelement(qn("w:rPrDefault"), {})
            dd.insert(0, rprd)                     # rPrDefault 는 pPrDefault 앞
        rpr = rprd.find(qn("w:rPr"))
        if rpr is None:
            rpr = etree.SubElement(rprd, qn("w:rPr"))
        _set_rfonts(rpr, FONT_LIGHT)
    # 테마 글꼴(스타일이 테마를 가리키는 문서가 있다)
    for rel in doc.part.rels.values():
        if rel.reltype == RT.THEME:
            part = rel.target_part
            blob = part.blob.decode("utf-8")
            blob = re.sub(r'(<a:(?:latin|ea) typeface=")[^"]*(")', r"\g<1>%s\2" % FONT_LIGHT, blob)
            blob = re.sub(r'(<a:font script="Hang" typeface=")[^"]*(")', r"\g<1>%s\2" % FONT_LIGHT, blob)
            part._blob = blob.encode("utf-8")
    # 본문의 글자 하나하나(글상자 안까지)
    body = doc.element.body
    for r in body.iter(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is None:
            continue
        p = r.getparent()
        while p is not None and p.tag != qn("w:p"):
            p = p.getparent()
        ps = None
        if p is not None:
            ppr = p.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:pStyle")) is not None:
                ps = ppr.find(qn("w:pStyle")).get(qn("w:val"))
        rs = rpr.find(qn("w:rStyle"))
        direct = _on(rpr, "w:b")
        bold = direct if direct is not None else (
            bold_of.get(rs.get(qn("w:val")), False) if rs is not None else bold_of.get(ps, False))
        _font_weight(rpr, bold)
    # 번호·글머리표(기호 글꼴은 그대로)
    try:
        num = doc.part.numbering_part.element
        for rpr in num.iter(qn("w:rPr")):
            _set_rfonts(rpr, FONT_LIGHT)
    except Exception:
        pass


# ──────────────────────────────────────────────
# 2. 표 머리글 → 네이비
# ──────────────────────────────────────────────
def _fill(tc):
    tcpr = tc.find(qn("w:tcPr"))
    shd = tcpr.find(qn("w:shd")) if tcpr is not None else None
    if shd is None:
        return None
    f = (shd.get(qn("w:fill")) or "").upper()
    return None if f in ("", "AUTO", "FFFFFF") else f


def _lum(hexs):
    try:
        r, g, b = (int(hexs[i:i + 2], 16) / 255 for i in (0, 2, 4))
    except Exception:
        return 1.0
    return 0.299 * r + 0.587 * g + 0.114 * b


def _first_child(parent, name):
    """parent 맨 앞에 w:name 자식(pPr·rPr·tcPr 처럼 늘 맨 앞인 것)."""
    el = parent.find(qn("w:" + name))
    if el is None:
        el = parent.makeelement(qn("w:" + name), {})
        parent.insert(0, el)
    return el


def _paint_navy(tc):
    tcpr = _first_child(tc, "tcPr")
    shd = tcpr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): NAVY})
    _put(tcpr, shd)
    # 칸 안 글자·문단에 남은 바탕(원본 검정 글자 바탕)은 지운다 — 네이비 칸 위에 검은 상자가 떴다(방배동 5. 분양사례 머리글)
    for p in tc.iter(qn("w:p")):
        for holder in [p.find(qn("w:pPr"))] + [r.find(qn("w:rPr")) for r in p.iter(qn("w:r"))]:
            if holder is None:
                continue
            for e in list(holder.findall(qn("w:shd"))) + list(holder.findall(qn("w:highlight"))):
                holder.remove(e)
            rp = holder.find(qn("w:rPr")) if holder.tag == qn("w:pPr") else None
            if rp is not None:
                for e in list(rp.findall(qn("w:shd"))) + list(rp.findall(qn("w:highlight"))):
                    rp.remove(e)
    # 글씨는 흰색 — 빈 칸에도 박아 둔다(나중에 타이핑하면 흰 글씨로 나오게)
    for p in tc.findall(qn("w:p")):
        ppr = _first_child(p, "pPr")
        targets = [_get(ppr, "rPr")] + [_first_child(r, "rPr") for r in p.findall(qn("w:r"))]
        for rpr in targets:
            _put(rpr, rpr.makeelement(qn("w:color"), {qn("w:val"): "FFFFFF"}))


def _is_frame(tbl):
    """'큰 틀' 표인가 — 맨 윗줄이 표 폭 전체를 덮는 진한 제목 띠 하나.

    예) 안산 '정보통신업 사업자 관련 설명' : 남색 제목 띠 아래에 작은 표 두 개(회색 머리글)가 들어 있다.
        pdf2docx 는 이걸 표 하나로 합쳐 만들어서, 0번 줄 = 큰 틀 제목, 1번 줄 = 안쪽 표 머리글이 된다.
    """
    rows = tbl.findall(qn("w:tr"))
    if len(rows) < 2:
        return False
    first = rows[0].findall(qn("w:tc"))
    ncols = len(tbl.findall(qn("w:tblGrid") + "/" + qn("w:gridCol"))) or max(len(r.findall(qn("w:tc"))) for r in rows)
    if len(first) == 1:
        span = first[0].find(qn("w:tcPr") + "/" + qn("w:gridSpan"))
        wide = ncols <= 1 or (span is not None and int(span.get(qn("w:val"), "1")) >= ncols - 1)
    else:
        wide = False
    f = _fill(first[0]) if first else None
    return wide and bool(f) and _lum(f) < 0.5


_SEC_NO_RE = re.compile(r"\s*(\d{1,2}|[IVX]{1,4}|[Ⅰ-Ⅻ])\.?\s*")


def _heading_row(tbl):
    """대단원 띠 표의 '진짜 줄' 하나(나머지가 글 없는 얇은 줄 = 밑줄용이면 그것들도 함께 돌려준다). 아니면 (None, [])."""
    rows = tbl.findall(qn("w:tr"))
    real, thin = [], []
    for tr in rows:
        h = tr.find(f"{qn('w:trPr')}/{qn('w:trHeight')}")
        hv = int(float(h.get(qn("w:val")))) if h is not None and re.fullmatch(r"[\d.]+", h.get(qn("w:val")) or "") else 999
        if not "".join(t.text or "" for t in tr.iter(qn("w:t"))).strip() and hv < 120:
            thin.append(tr)                                  # 고산3지구: 제목 줄 아래 1.7pt 짜리 밑줄 줄
        else:
            real.append(tr)
    return (real[0], thin) if len(real) == 1 else (None, [])


def restyle_section_headings(doc):
    """원본 대단원 제목 띠 [색 칠한 번호 칸][제목 + 밑줄] → 회사 모양: 번호 칸 네이비 + 흰 번호, 밑줄은 뺀다.
    번호·제목 글은 그대로(2026-10-12 사용자 확정, 그림판 그림: 노란 '3' 상자+노란 밑줄 → 네이비 '3' 상자).
    조건: 맨 바깥 표, 1줄 2칸, 첫 칸이 칠해져 있고 글이 번호(1·2·Ⅰ·II…)뿐, 둘째 칸 글이 짧다(40자 이하)."""
    n = 0
    body = doc.element.body
    for tbl in list(body.iter(qn("w:tbl"))):
        if any(a.tag == qn("w:tbl") for a in tbl.iterancestors()):
            continue
        row, thin = _heading_row(tbl)
        if row is None:
            continue
        cells = row.findall(qn("w:tc"))
        if len(cells) != 2:
            continue
        num_txt = "".join(t.text or "" for t in cells[0].iter(qn("w:t")))
        title = "".join(t.text or "" for t in cells[1].iter(qn("w:t"))).strip()
        if not _SEC_NO_RE.fullmatch(num_txt or "x") or not title or len(title) > 40:
            continue
        fills = [_fill(tc) for tc in cells[0].iter(qn("w:tc"))]
        if not any(fills):
            continue
        # 번호 칸: 안쪽 작은 표(pdf2docx 가 만든 것)를 걷어내고 번호 하나만 가로·세로 가운데로(2026-10-13 사용자 확정)
        c0 = cells[0]
        run0 = next((r for r in c0.iter(qn("w:r")) if "".join(t.text or "" for t in r.iter(qn("w:t"))).strip()), None)
        rpr0 = copy.deepcopy(run0.find(qn("w:rPr"))) if run0 is not None and run0.find(qn("w:rPr")) is not None else None
        for ch in list(c0):
            if ch.tag != qn("w:tcPr"):
                c0.remove(ch)
        p = c0.makeelement(qn("w:p"), {})
        ppr = etree.SubElement(p, qn("w:pPr"))
        etree.SubElement(ppr, qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0"})
        etree.SubElement(ppr, qn("w:ind"), {qn("w:left"): "0", qn("w:right"): "0", qn("w:firstLine"): "0"})
        etree.SubElement(ppr, qn("w:jc"), {qn("w:val"): "center"})
        r = etree.SubElement(p, qn("w:r"))
        if rpr0 is not None:
            for sp in rpr0.findall(qn("w:spacing")):
                rpr0.remove(sp)
            r.append(rpr0)
        t = etree.SubElement(r, qn("w:t"))
        t.text = num_txt.strip()
        c0.append(p)
        tcpr = c0.find(qn("w:tcPr"))
        mar = tcpr.find(qn("w:tcMar"))
        if mar is not None:
            tcpr.remove(mar)
        _paint_navy(c0)
        _put(tcpr, tcpr.makeelement(qn("w:vAlign"), {qn("w:val"): "center"}))
        b = tcpr.find(qn("w:tcBorders"))
        if b is not None:
            # ★테두리(원본 노란 밑줄 → 네이비)는 칠과 사이에 흰 틈이 생겨 흰 줄로 보였다(2026-10-14) → 테두리 없이 칠만
            for e in list(b):
                b.remove(e)
            for side in ("top", "left", "bottom", "right"):
                b.append(b.makeelement(qn("w:" + side), {qn("w:val"): "nil"}))
        # 정사각: 줄 높이 = 번호 칸 폭(사용자: '네모로')
        w = tcpr.find(qn("w:tcW"))
        if w is not None and re.fullmatch(r"[\d.]+", w.get(qn("w:w")) or ""):
            side = int(float(w.get(qn("w:w"))))
            trpr = row.find(qn("w:trPr"))
            if trpr is None:
                trpr = row.makeelement(qn("w:trPr"), {})
                row.insert(0, trpr)
            h = trpr.find(qn("w:trHeight"))
            if h is None:
                h = trpr.makeelement(qn("w:trHeight"), {})
                trpr.append(h)
            h.set(qn("w:val"), str(side))
            h.set(qn("w:hRule"), "exact")
        for tr in thin:                                       # 밑줄용 얇은 줄도 뺀다
            tbl.remove(tr)
        b = cells[1].find(f"{qn('w:tcPr')}/{qn('w:tcBorders')}")   # 제목 칸 밑줄·옆줄 빼기
        if b is not None:
            for e in list(b):
                b.remove(e)
        n += 1
    return n


_ROMAN = ["", "Ⅰ", "Ⅱ", "Ⅲ", "Ⅳ", "Ⅴ", "Ⅵ", "Ⅶ", "Ⅷ", "Ⅸ", "Ⅹ", "Ⅺ", "Ⅻ"]
_ROMAN_ASCII = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"]
_HEAD_PATTERNS = [                                          # (이름, 정규식: 1=앞 2=번호 3=뒤 4=제목)
    ("roman", re.compile(r"^(\s*)([Ⅰ-Ⅻ]|(?:X{0,1}(?:IX|IV|V?I{0,3})))(\s*[.)]\s*|\s+)(\S.*)$")),
    ("paren", re.compile(r"^(\s*\()(\d{1,2})(\)\s*)(\S.*)$")),
    ("digit", re.compile(r"^(\s*)(\d{1,2})(\s*[.)]\s*)(\S.*)$")),
]


def _num_value(kind, tok):
    if kind == "roman":
        tok = tok.strip()
        if tok in _ROMAN:
            return _ROMAN.index(tok)
        return _ROMAN_ASCII.index(tok) if tok in _ROMAN_ASCII else None
    return int(tok)


def _num_text(kind, tok, v):
    if kind == "roman":
        if v <= 0:
            return None
        return (_ROMAN if tok.strip() in _ROMAN else _ROMAN_ASCII)[v]
    return str(v)


def _bond_heading_from_original(doc):
    """사모사채 개요 제목을 원본 대단원 제목 모양 그대로(2026-10-13 사용자 확정).
    ① 원본에 [번호 상자][제목] 띠가 있으면 → 그 띠를 복사(이미 회사 네이비로 바뀐 것), 번호 = 원본 첫 대단원 번호 − 1
       (원본이 1 부터면 0), 제목 '사모사채 개요'.
    ② 없으면 → 원본 대단원 문단(Ⅰ.·1.·(1) 중 글자가 가장 큰 것)을 복사, 같은 번호 모양으로 번호 − 1.
       로마숫자인데 0 이 되면 번호 없이 제목만.
    못 찾으면 None(번호 없이 '사모사채 개요')."""
    body = doc.element.body
    for tbl in body.iter(qn("w:tbl")):
        if any(a.tag == qn("w:tbl") for a in tbl.iterancestors()):
            continue
        row, _ = _heading_row(tbl)
        if row is None or len(row.findall(qn("w:tc"))) != 2 or len(tbl.findall(qn("w:tr"))) != 1:
            continue
        c0, c1 = row.findall(qn("w:tc"))
        num = "".join(t.text or "" for t in c0.iter(qn("w:t"))).strip().rstrip(".")
        if not _SEC_NO_RE.fullmatch(num or "x") or not _fill(c0):
            continue
        kind = "digit" if num.isdigit() else "roman"
        v = _num_value(kind, num)
        if v is None:
            continue
        new = copy.deepcopy(tbl)
        n0, n1 = new.findall(qn("w:tr"))[0].findall(qn("w:tc"))
        txt = _num_text(kind, num, max(v - 1, 1))
        _set_cell_text(n0, txt if txt is not None else "")
        _set_cell_text(n1, "사모사채 개요")
        return new
    for p in body.findall(qn("w:p")):                        # 방배동식: 번호 글자 바탕만 칠한 문단
        hit = _shaded_number_runs(p)
        if not hit:
            continue
        shaded, num, title_r = hit
        num = num.rstrip(".")
        kind = "digit" if num.isdigit() else "roman"
        v = _num_value(kind, num)
        if v is None:
            continue
        new = copy.deepcopy(p)
        idx = list(p).index(title_r)
        for ch in list(new)[idx + 1:]:                       # 제목 뒤(줄바꿈·다음 소제목)는 버린다
            new.remove(ch)
        nruns = [r for r in new.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
        txt = _num_text(kind, num, max(v - 1, 1))
        # 번호가 든 조각만 바꾸고, 번호 뒤 칠한 빈칸(상자 여백)은 원본 그대로 둔다
        for r in nruns[:len(shaded)]:
            for t in r.findall(qn("w:t")):
                if (t.text or "").strip():
                    t.text = ""
        if nruns:
            first_num = next((r for r in nruns[:len(shaded)]), None)
            first_num.find(qn("w:t")).text = txt if txt is not None else ""
            last = nruns[len(shaded)].find(qn("w:t")) if len(nruns) > len(shaded) else None
            if last is not None:
                lead = (last.text or "")[:len(last.text or "") - len((last.text or "").lstrip())]
                last.text = lead + "사모사채 개요"
                last.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        for sp in new.iter(qn("w:sectPr")):
            sp.getparent().remove(sp)
        return new
    lv = _level(_para_heads(doc))                            # 대단원 무리의 첫 제목 모양
    if not lv:
        return None
    _, kind, p, m = lv[0]
    v = _num_value(kind, m.group(2))
    new = copy.deepcopy(p)
    numtxt = _num_text(kind, m.group(2), max(v - 1, 1))
    _set_heading_text(new, m.end(3), m.group(1) + numtxt + m.group(3), "사모사채 개요")
    for sp in new.iter(qn("w:sectPr")):
        sp.getparent().remove(sp)
    return new


def _shaded_number_runs(p):
    """문단 맨 앞이 '칠한 바탕의 번호' 글 조각이면 (번호 조각들, 번호 글, 제목 조각) — 방배동식 대단원
    [■3■][ 담보자산 개요]. 아니면 None."""
    runs = [r for r in p.findall(qn("w:r")) if r.find(qn("w:t")) is not None]
    shaded, num = [], ""
    for r in runs:
        shd = r.find(f"{qn('w:rPr')}/{qn('w:shd')}")
        f = (shd.get(qn("w:fill")) or "").upper() if shd is not None else ""
        txt = "".join(t.text or "" for t in r.findall(qn("w:t")))
        if f not in ("", "AUTO", "FFFFFF"):
            shaded.append(r)
            num += txt
            continue
        if shaded and txt.strip():
            return (shaded, num.strip(), r) if _SEC_NO_RE.fullmatch(num.strip() or "x") else None
        if txt.strip():
            return None
    return None


def restyle_shaded_number_headings(doc):
    """방배동식 대단원(번호 글자 바탕만 칠한 문단 '■2■금융 조건') → [정사각 네이비 번호 칸][제목] 1줄 2칸 표.
    ★글자 바탕은 글자 폭만큼만 칠해져 세로로 긴 직사각형이 됐다(2026-10-13 사용자: '네모로'). 원본 상자는 21×21pt."""
    n = 0
    body = doc.element.body
    for p in list(body.findall(qn("w:p"))):
        hit = _shaded_number_runs(p)
        if not hit:
            continue
        shaded, num, title_r = hit
        idx = list(p).index(title_r)
        tail = list(p)[idx + 1:]                             # 제목 뒤(줄바꿈 + 소제목 등) → 표 뒤 새 문단으로
        title = "".join(t.text or "" for t in title_r.findall(qn("w:t"))).strip()
        rpr_n = copy.deepcopy(shaded[0].find(qn("w:rPr")))
        rpr_t = copy.deepcopy(title_r.find(qn("w:rPr"))) if title_r.find(qn("w:rPr")) is not None else None
        sz = rpr_n.find(qn("w:sz")) if rpr_n is not None else None
        pt = int(sz.get(qn("w:val"))) / 2 if sz is not None and (sz.get(qn("w:val")) or "").isdigit() else 14
        side = int(round(pt * 1.5 * 20))                    # 정사각 한 변(twips) — 글자 크기의 1.5배(원본 14pt → 21pt)
        tbl = etree.Element(qn("w:tbl"))
        tp = etree.SubElement(tbl, qn("w:tblPr"))
        etree.SubElement(tp, qn("w:tblW"), {qn("w:w"): "0", qn("w:type"): "auto"})
        ind = p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
        etree.SubElement(tp, qn("w:tblInd"), {qn("w:w"): (ind.get(qn("w:left")) or "0") if ind is not None else "0",
                                              qn("w:type"): "dxa"})
        etree.SubElement(tp, qn("w:tblLayout"), {qn("w:type"): "fixed"})
        cm = etree.SubElement(tp, qn("w:tblCellMar"))
        for side_n in ("left", "right"):
            etree.SubElement(cm, qn("w:" + side_n), {qn("w:w"): "0", qn("w:type"): "dxa"})
        grid = etree.SubElement(tbl, qn("w:tblGrid"))
        etree.SubElement(grid, qn("w:gridCol"), {qn("w:w"): str(side)})
        etree.SubElement(grid, qn("w:gridCol"), {qn("w:w"): "8000"})
        tr = etree.SubElement(tbl, qn("w:tr"))
        trp = etree.SubElement(tr, qn("w:trPr"))
        etree.SubElement(trp, qn("w:trHeight"), {qn("w:val"): str(side), qn("w:hRule"): "exact"})
        for k, (txt, rpr, w) in enumerate(((num.strip(), rpr_n, side), (title, rpr_t, 8000))):
            tc = etree.SubElement(tr, qn("w:tc"))
            tcp = etree.SubElement(tc, qn("w:tcPr"))
            etree.SubElement(tcp, qn("w:tcW"), {qn("w:w"): str(w), qn("w:type"): "dxa"})
            if k == 1:
                etree.SubElement(tcp, qn("w:tcMar")).append(etree.Element(qn("w:left"), {qn("w:w"): "120", qn("w:type"): "dxa"}))
            etree.SubElement(tcp, qn("w:vAlign"), {qn("w:val"): "center"})
            cp = etree.SubElement(tc, qn("w:p"))
            cpp = etree.SubElement(cp, qn("w:pPr"))
            etree.SubElement(cpp, qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0", qn("w:line"): "240",
                                                    qn("w:lineRule"): "auto"})
            etree.SubElement(cpp, qn("w:jc"), {qn("w:val"): "center" if k == 0 else "left"})
            r = etree.SubElement(cp, qn("w:r"))
            if rpr is not None:
                rr = copy.deepcopy(rpr)
                for e in list(rr.findall(qn("w:shd"))) + list(rr.findall(qn("w:spacing"))):
                    rr.remove(e)
                r.append(rr)
            t = etree.SubElement(r, qn("w:t"))
            t.text = txt
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            if k == 0:
                _paint_navy(tc)
        p.addprevious(tbl)
        if any("".join(x.text or "" for x in el.iter(qn("w:t"))).strip() for el in tail):
            rest = etree.Element(qn("w:p"))
            ppr0 = p.find(qn("w:pPr"))
            if ppr0 is not None:
                rest.append(copy.deepcopy(ppr0))
            for el in tail:
                if el.tag == qn("w:r") and el.find(qn("w:br")) is not None and not rest.findall(qn("w:r")):
                    continue                                 # 맨 앞 줄바꿈은 새 문단이 대신한다
                rest.append(el)
            p.addprevious(rest)
        else:
            gap = etree.Element(qn("w:p"))                   # 표 뒤 빈 문단(다음 표와 붙지 않게)
            gpp = etree.SubElement(gap, qn("w:pPr"))
            etree.SubElement(gpp, qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0", qn("w:line"): "120",
                                                    qn("w:lineRule"): "exact"})
            p.addprevious(gap)
        p.getparent().remove(p)
        n += 1
    return n


def _set_cell_text(tc, text):
    p = next(iter(tc.findall(qn("w:p"))), None)
    if p is None:
        p = etree.SubElement(tc, qn("w:p"))
    for extra in tc.findall(qn("w:p"))[1:]:
        tc.remove(extra)
    if p.find(".//" + qn("w:t")) is None:
        r = etree.SubElement(p, qn("w:r"))
        etree.SubElement(r, qn("w:t"))
    _set_visible_text(p, text)                               # 첫 글 조각이 흰 글씨면 제목이 안 보였다(이천 '0 사모사채 개요')


def strip_cell_trailing_spaces(doc):
    """표 칸 글 끝의 빈칸을 지운다 — 가운데 맞춤 칸에서 글이 한쪽으로 쏠려 보였다(이천 대월 '차    주 ')."""
    n = 0
    for tc in doc.element.body.iter(qn("w:tc")):
        for p in tc.findall(qn("w:p")):
            ts = [t for t in p.iter(qn("w:t")) if t.text]
            if ts and ts[-1].text != ts[-1].text.rstrip():
                ts[-1].text = ts[-1].text.rstrip()
                n += 1
    return n


FIN_TITLE = "본건 기초자산 금융조건"
_FIN_TITLE_RE = re.compile(r"(주요\s*)?금융\s*(개요|조건|구조)|금융\s*조건|금\s*융\s*조\s*건")


_STRUCT_RE = re.compile(r"^[\s■□▪◆◇●○▶▷❑❏\-\d.)(\[\]<>Ⅰ-ⅫIVX]*\s*(금융|사업|거래|딜|대출)?\s*구\s*조\s*(도)?\s*[)\]>]?\s*$")


def drop_structure_paragraphs(doc):
    """원본 구조도가 표 밖 '■ 금융 구조도' 제목 + 그림 문단으로 있는 것 → 제목과 그림을 지운다(2026-10-15 사용자 4번:
    '원본 구조도는 항상 뺀다, 무시하지 말고 치워'). 고산3지구 '1 금융 개요' 쪽. 지운 묶음 수."""
    n = 0
    body = doc.element.body
    for p in list(body.findall(qn("w:p"))):
        if p.getparent() is None:
            continue
        t = _vis_text(p).strip()
        if not t or len(t) > 20 or not _STRUCT_RE.match(t):
            continue
        # 바로 뒤(빈 문단은 건너뛰고) 그림 문단이 있어야 원본 구조도
        nxt = p.getnext()
        gone = []
        while nxt is not None and nxt.tag == qn("w:p") and not _vis_text(nxt).strip() and                 nxt.find(".//" + qn("w:drawing")) is None and nxt.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is None:
            gone.append(nxt)
            nxt = nxt.getnext()
        pic_tbl = nxt is not None and nxt.tag == qn("w:tbl") and nxt.find(".//" + qn("w:drawing")) is not None             and not _vis_text(nxt).strip()                      # 그림만 든 표(안산 '[금융 구조]')
        if nxt is None or (not pic_tbl and (nxt.tag != qn("w:p") or nxt.find(".//" + qn("w:drawing")) is None)):
            continue
        pics = [nxt]                                           # 바로 아래 그림 하나만(그 다음 그림은 다른 내용일 수 있다 — 안산 자금 사용)
        for el in [p] + gone + pics:
            if el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:   # 구역 끝 문단은 남기고 내용만 비움
                for ch in [c for c in el if c.tag != qn("w:pPr")]:
                    el.remove(ch)
            else:
                el.getparent().remove(el)
        n += 1
    return n


def right_align_units(doc):
    """'[단위: 백만원]'·'(단위 : 억원)' 같은 단위 줄은 원본처럼 오른쪽 맞춤(고산 담보필지목록, 사용자 11번). 바꾼 수."""
    n = 0
    for p in doc.element.body.iter(qn("w:p")):
        t = _vis_text(p).strip()
        if not t or len(t) > 30 or not re.match(r"^[\[\(<]?\s*단\s*위\s*[:：]", t):
            continue
        ppr = _first_child(p, "pPr")
        ind = ppr.find(qn("w:ind"))
        if ind is not None:
            for a in list(ind.attrib):
                ind.attrib.pop(a)
            ind.set(qn("w:left"), "0")
            ind.set(qn("w:right"), "0")
        _put(ppr, ppr.makeelement(qn("w:jc"), {qn("w:val"): "right"}))
        n += 1
    return n


def drop_structure_rows(doc):
    """원본 표의 '구조도' 줄(구분 칸 글에 '구조도', 그 줄에 그림)을 지운다 — 회사 금융 구조도가 사모사채 개요 아래에
    따로 들어가므로(2026-10-13 사용자 확정 '구조도 사진 항목 행을 지워서 이어서'). 지운 수."""
    n = 0
    for tr in list(doc.element.body.iter(qn("w:tr"))):
        tcs = tr.findall(qn("w:tc"))
        if len(tcs) < 2:
            continue
        label = "".join(t.text or "" for t in tcs[0].iter(qn("w:t"))).replace(" ", "")
        if len(label) > 12 or not re.search(r"(금융|사업|거래|대출|딜)?구조(도)?$", label):   # 신촌 '금융구조'(도 없음)
            continue
        if tr.find(".//" + qn("w:drawing")) is None and tr.find(".//" + qn("w:pict")) is None:
            continue
        nxt = tr.getnext()                                   # 세로로 합친 칸이 이 줄에서 시작하면 다음 줄로 넘긴다
        for k, tc in enumerate(tcs):
            vm = tc.find(f"{qn('w:tcPr')}/{qn('w:vMerge')}")
            if vm is not None and vm.get(qn("w:val")) == "restart" and nxt is not None:
                ntcs = nxt.findall(qn("w:tc"))
                if k < len(ntcs):
                    nvm = ntcs[k].find(f"{qn('w:tcPr')}/{qn('w:vMerge')}")
                    if nvm is not None:
                        nvm.set(qn("w:val"), "restart")
        tbl = tr.getparent()
        for anc in tr.iterancestors(qn("w:tr")):             # 감싼 바깥 줄의 높이는 지운 줄 몫이 남아 빈 자리가 생긴다 → 풀기
            h = anc.find(f"{qn('w:trPr')}/{qn('w:trHeight')}")
            if h is not None:
                h.getparent().remove(h)
        tbl.remove(tr)
        if not tbl.findall(qn("w:tr")):
            tbl.getparent().remove(tbl)
        n += 1
    return n


def whiten_borders_to_black(doc):
    """표 테두리 중 흰색(FFFFFF)을 그 표의 다른 테두리 색(가장 많이 쓴 색, 없으면 검정)으로. 다른 색은 그대로. 바꾼 수.
    ★2026-10-14 사용자: '흰색 테두리는 없다 — 그 표의 다른 테두리와 같은 색으로'(10-13 엔 무조건 검정이었다)."""
    from collections import Counter
    n = 0
    for tbl in doc.element.body.iter(qn("w:tbl")):
        own = [b for b in tbl.iter(qn("w:tcBorders"), qn("w:tblBorders"))
               if next(b.iterancestors(qn("w:tbl")), None) is tbl]
        cnt = Counter()
        for b in own:
            for e in b:
                col = (e.get(qn("w:color")) or "").lstrip("#").upper()
                if e.get(qn("w:val")) not in ("nil", "none", None) and col not in ("", "AUTO", "FFFFFF"):
                    cnt[col] += 1
        to = cnt.most_common(1)[0][0] if cnt else "000000"
        for b in own:
            for e in b:
                col = (e.get(qn("w:color")) or "").lstrip("#").upper()
                if col == "FFFFFF" and e.get(qn("w:val")) not in ("nil", "none"):
                    e.set(qn("w:color"), to)
                    n += 1
    return n


def _set_visible_text(p, text):
    """글을 '흰 글씨가 아닌' 첫 글 조각에 넣고 나머지 조각은 비운다(돈암동: 첫 조각이 흰 글씨라 제목이 안 보였다)."""
    runs = [r for r in p.iter(qn("w:r")) if r.find(qn("w:t")) is not None]
    def white(r):
        c = r.find(f"{qn('w:rPr')}/{qn('w:color')}")
        return c is not None and (c.get(qn("w:val")) or "").upper() == "FFFFFF"
    keep = next((r for r in runs if not white(r)), runs[0] if runs else None)
    if keep is None:
        return
    for r in runs:
        for t in r.findall(qn("w:t")):
            t.text = ""
    t = keep.find(qn("w:t"))
    t.text = text
    t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def rename_finance_headings(doc):
    """원본 대단원 중 금융 관련 제목(금융개요·주요 금융조건·금융 조건…) → '본건 기초자산 금융조건'
    (2026-10-13 사용자 확정, 번호·모양은 그대로). 번호 상자 표·번호 문단 둘 다. 바꾼 수."""
    n = 0
    boxes = False
    body = doc.element.body
    for tbl in body.iter(qn("w:tbl")):                        # ① 번호 상자 표
        if any(a.tag == qn("w:tbl") for a in tbl.iterancestors()):
            continue
        rows = tbl.findall(qn("w:tr"))
        if len(rows) != 1 or len(rows[0].findall(qn("w:tc"))) != 2:
            continue
        c0, c1 = rows[0].findall(qn("w:tc"))
        num = "".join(t.text or "" for t in c0.iter(qn("w:t"))).strip()
        title = "".join(t.text or "" for t in c1.iter(qn("w:t"))).strip()
        if _SEC_NO_RE.fullmatch(num or "x") and title:
            boxes = True                                     # 번호 상자식 대단원이 있는 원본 → 번호 문단(소제목)은 안 건드린다
        if _SEC_NO_RE.fullmatch(num or "x") and title and len(title) <= 20 and _FIN_TITLE_RE.search(title):
            p = next(pp for pp in c1.iter(qn("w:p")) if "".join(t.text or "" for t in pp.iter(qn("w:t"))).strip())
            _set_visible_text(p, FIN_TITLE)
            n += 1
    if boxes:
        return n
    found = []                                              # ② 번호 문단(가장 큰 글자 무리 = 대단원)
    for p in body.findall(qn("w:p")):
        txt = "".join(t.text or "" for t in p.iter(qn("w:t")))
        if not txt.strip() or len(txt.strip()) > 30:
            continue
        for kind, rx in _HEAD_PATTERNS:
            m = rx.match(txt)
            if m and m.group(2) and _num_value(kind, m.group(2)):
                szs = [int(x.get(qn("w:val"))) for x in p.iter(qn("w:sz")) if (x.get(qn("w:val")) or "").isdigit()]
                found.append((max(szs) if szs else 20, kind, p, m))
                break
    if found:
        top = max(f[0] for f in found)
        for sz, kind, p, m in found:
            if sz >= top - 1 and _FIN_TITLE_RE.search(m.group(4)):
                _set_par_text(p, m.group(1) + m.group(2) + m.group(3) + FIN_TITLE)
                ind = p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
                if ind is not None:                          # 원본 오른쪽 들여쓰기가 좁아 두 줄로 꺾였다(내곡동)
                    ind.set(qn("w:right"), "0")
                n += 1
    return n


def center_cells_vertically(doc):
    """모든 표 칸의 글을 세로 가운데로(가로는 원본 그대로). 2026-10-12 사용자 확정 '표 글은 무조건 셀 가운데'.
    pdf2docx 는 칸 안 위치를 문단 '앞 간격' 으로 맞춰 두므로, 그 간격(첫 문단 앞)을 없애야 진짜 가운데가 된다."""
    n = 0
    for tc in doc.element.body.iter(qn("w:tc")):
        tcpr = tc.find(qn("w:tcPr"))
        if tcpr is None:
            tcpr = tc.makeelement(qn("w:tcPr"), {})
            tc.insert(0, tcpr)
        va = tcpr.find(qn("w:vAlign"))
        if va is None:
            va = tcpr.makeelement(qn("w:vAlign"), {})
            hm = tcpr.find(qn("w:hideMark"))
            hm.addprevious(va) if hm is not None else tcpr.append(va)   # vAlign 은 tcPr 끝(hideMark 앞)
        va.set(qn("w:val"), "center")
        first = next((ch for ch in tc if ch.tag in (qn("w:p"), qn("w:tbl"))), None)
        if first is not None and first.tag == qn("w:p"):
            sp = first.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
            if sp is not None:
                sp.set(qn("w:before"), "0")
        n += 1
    return n


def _is_wrapper(tbl):
    """1줄 1칸짜리 '껍데기' 표(안에 표가 들어 있는 것) — pdf2docx 가 본문 표 전체를 이렇게 감싸곤 한다(방배동)."""
    rows = tbl.findall(qn("w:tr"))
    if len(rows) != 1:
        return False
    tcs = rows[0].findall(qn("w:tc"))
    return len(tcs) == 1 and tcs[0].find(qn("w:tbl")) is not None


def paint_headers(doc):
    """머리글 칸을 회사 네이비로. 고친 칸 수를 돌려준다.

    규칙(사용자 확정)
      · 보통 표 : 맨 위 1~2줄 머리글은 색이 연해도 네이비 + 흰 글씨(붕어빵과 같다), 진한 칸은 어디 있든 네이비.
      · 표 안의 표('큰 틀') : **큰 틀 제목 띠만** 회사 네이비. 그 안 작은 표들의 색은 **원본 그대로**(2026-10-08).
        원본이 이미 네이비여도 회사 네이비(17365D)로 맞춘다.
    """
    n = 0
    body = doc.element.body
    frames = [t for t in body.iter(qn("w:tbl")) if _is_frame(t)]
    frame_ids = set(map(id, frames))
    for tbl in body.iter(qn("w:tbl")):
        if any(not _is_wrapper(a) for a in tbl.iterancestors(qn("w:tbl"))):
            continue                                          # 표 안의 표는 전부 원본 색 그대로(2026-10-08 확정, 10-13 이천 대월에서 바로잡음)
                                                              # ★단 pdf2docx 가 표 전체를 감싼 '1칸짜리 껍데기 표' 는 표로 안 친다(방배동)
        rows = tbl.findall(qn("w:tr"))
        if not rows:
            continue
        cells = [r.findall(qn("w:tc")) for r in rows]
        if tbl in frames:
            _paint_navy(cells[0][0])
            n += 1
            continue
        if len(rows) < 2 or max(len(c) for c in cells) < 2:
            for row in cells:                                 # 한 줄짜리 제목 띠 등 — 진한 칸만
                for tc in row:
                    f = _fill(tc)
                    if f and _lum(f) < 0.5 and f != NAVY:
                        _paint_navy(tc)
                        n += 1
            continue
        head = set()
        for i, r in enumerate(rows):                      # '머리글 줄 반복' 표시가 된 줄
            trpr = r.find(qn("w:trPr"))
            if trpr is not None and _on(trpr, "w:tblHeader"):
                head.add(i)
        body_fills = [_fill(tc) for row in cells[1:] for tc in row]
        common = max(set(body_fills), key=body_fills.count) if body_fills else None
        for i in (0, 1):                                  # 맨 위 1~2줄이 칠해져 있으면 머리글
            if i >= len(rows) - 1:
                break
            fills = [_fill(tc) for tc in cells[i]]
            colored = [f for f in fills if f]
            if len(colored) >= max(1, 0.6 * len(fills)) and (i == 0 or i - 1 in head) \
                    and max(set(colored), key=colored.count) != common:
                head.add(i)
            else:
                break
        for i, row in enumerate(cells):
            for tc in row:
                f = _fill(tc)
                if i in head or (f and _lum(f) < 0.5):     # 진한 칸은 어디 있든 머리글로 본다
                    if f != NAVY or i in head:
                        _paint_navy(tc)
                        n += 1
    # 남색·진파랑 칸은 표 안의 표까지 전부 정확히 회사 네이비로(2026-10-14 사용자: '비슷한 남색이어도 17365D 로')
    for tc in body.iter(qn("w:tc")):
        f = _fill(tc)
        if f and f != NAVY and _navy_like(f):
            _paint_navy(tc)
            n += 1
    return n


def _navy_like(hexs):
    """남색·진파랑(파랑이 가장 세고 어두운 색)인가."""
    try:
        r, g, b = (int(hexs[i:i + 2], 16) for i in (0, 2, 4))
    except Exception:
        return False
    return b >= r + 30 and b >= g and _lum(hexs) < 0.42


# ──────────────────────────────────────────────
# 3. 원본 표지·끝쪽
# ──────────────────────────────────────────────
def _hard_break(el):
    """이 요소가 '여기서 쪽이 끝난다' 를 품고 있나 → 'run'(쪽나눔) / 'sect'(구역나눔) / None"""
    if el.tag != qn("w:p"):
        return None
    ppr = el.find(qn("w:pPr"))
    if ppr is not None:
        sp = ppr.find(qn("w:sectPr"))
        if sp is not None:
            t = sp.find(qn("w:type"))
            if t is None or t.get(qn("w:val")) != "continuous":
                return "sect"
    for br in el.iter(qn("w:br")):
        if br.get(qn("w:type")) == "page":
            return "run"
    return None


def _starts_page(el):
    ppr = el.find(qn("w:pPr")) if el.tag == qn("w:p") else None
    return ppr is not None and _on(ppr, "w:pageBreakBefore")


def _text(els):
    return "".join("".join(e.itertext()) for e in els)


def _blocks(body):
    return [e for e in body if e.tag in (qn("w:p"), qn("w:tbl"), qn("w:sdt"))]


def find_cover(doc):
    """원본 표지 = (끝 블록 번호, 근거, 표지 글). 못 찾으면 (None, 이유, '')."""
    end, why = pages.cover_end(doc)
    if end is None:
        return None, why, ""
    bl = _blocks(doc.element.body)
    return end, why, " ".join(pages.text_of(e).strip() for e in bl[:end + 1] if pages.text_of(e).strip())


def find_contact_tail(doc):
    """맨 끝 쪽이 증권사 연락처 쪽이면 (시작 블록 번호, 글, 연락처 수). 아니면 (None, 글, 수)."""
    start = pages.last_page_start(doc)
    if start is None:
        return None, "", 0
    bl = _blocks(doc.element.body)
    txt = " ".join(pages.text_of(e) for e in bl[start:])
    hits = len(set(EMAIL_RE.findall(txt))) + len(set(PHONE_RE.findall(txt)))
    if hits >= 2 and len(txt) < 3000:
        return start, txt, hits
    return None, txt, hits


def remove_cover(doc, end):
    """블록 0 ~ end 를 뺀다. 표지 끝이 구역 나눔이면 그 구역 설정도 함께 사라진다(표지 전용 설정)."""
    if end is None:
        return False
    bl = _blocks(doc.element.body)
    for el in bl[:end + 1]:
        el.getparent().remove(el)
    return True


def remove_tail(doc, start):
    body = doc.element.body
    bl = _blocks(body)
    prev = bl[start - 1] if start > 0 else None
    for el in bl[start:]:
        body.remove(el)
    if prev is not None and _hard_break(prev) == "sect":
        # 마지막 구역의 쪽 설정이 사라졌다 → 앞 구역 설정을 문서 끝 설정으로 올린다
        sp = prev.find(qn("w:pPr")).find(qn("w:sectPr"))
        prev.find(qn("w:pPr")).remove(sp)
        old = body.find(qn("w:sectPr"))
        body.replace(old, sp)


def strip_trailing_break(doc):
    """본문 맨 끝의 쪽나눔(뒤에 우리 연락처 쪽이 붙으면 빈 쪽이 생긴다)."""
    bl = _blocks(doc.element.body)
    for el in reversed(bl):
        if el.tag != qn("w:p"):
            return
        brs = [b for b in el.iter(qn("w:br")) if b.get(qn("w:type")) == "page"]
        if brs:
            for b in brs:
                r = b.getparent()
                r.remove(b)
            return
        if "".join(el.itertext()).strip() or el.find(".//" + qn("w:drawing")) is not None:
            return


# ──────────────────────────────────────────────
# 원본에서 표지 제목·사진 후보 뽑기
# ──────────────────────────────────────────────
def suggest_title(doc, end=None):
    """원본 표지에서 가장 큰 글씨 줄들 → 표지 제목 후보(최대 2줄)."""
    if end is None:
        end, _, _ = find_cover(doc)
    bl = _blocks(doc.element.body)[: (end + 1) if end is not None else 10]
    lines = []
    for el in bl:
        for p in el.iter(qn("w:p")):
            if any(a.tag == pages.FB for a in p.iterancestors()):
                continue
            if p.find(".//" + qn("w:txbxContent")) is not None:
                continue                                   # 글상자를 품은 문단은 안쪽 문단에서 센다
            t = re.sub(r"\s+", " ", pages.text_of(p)).strip()
            if not t or len(t) > 60:
                continue
            sizes = [int(s.get(qn("w:val"))) for s in p.iter(qn("w:sz")) if (s.get(qn("w:val")) or "").isdigit()]
            lines.append((max(sizes) if sizes else 20, t))
    lines = [(s, t) for s, t in lines
             if not re.search(r"memorandum|confidential|private|^\d{4}[.\-/]?\d{0,2}\.?$", t, re.I)]
    if not lines:
        return []
    top = max(s for s, _ in lines)
    out = []
    for s, t in lines:
        if s >= top - 4 and t not in out:
            out.append(t)
    return out[:2]


def photo_candidates(doc, limit=6):
    """본문 그림 중 표지 사진으로 쓸 만한 것(큰 실사 사진 먼저)."""
    seen, out = set(), []
    for rel in doc.part.rels.values():
        if rel.reltype != RT.IMAGE or rel.is_external:
            continue
        part = rel.target_part
        if part.partname in seen:
            continue
        seen.add(part.partname)
        try:
            im = Image.open(io.BytesIO(part.blob))
            w, h = im.size
            if w < 300 or h < 150:
                continue
            small = im.convert("L").resize((60, 60))
            bright = sum(small.getdata()) / 3600
            # 실사 사진은 색이 수천 가지, 구조도·표 그림은 몇십 가지 — 어두운 구조도를 사진으로
            # 고르던 것을 막는다(실제: 엘피스 표지에 금융구조도가 골라졌다)
            colors = len(set(im.convert("RGB").resize((80, 80)).getdata()))
        except Exception:
            continue
        ratio = w / h
        score = (colors > 1500, bright < 210, 1.2 <= ratio <= 2.6, w * h)
        out.append((score, part.blob))
    out.sort(key=lambda t: t[0], reverse=True)
    return [b for _, b in out[:limit]]


# ──────────────────────────────────────────────
# 4. 회사 양식 붙이기
# ──────────────────────────────────────────────
def _prepare_photo(blob):
    """표지 사진: **자르지 않는다.** 원본 그대로 두고 화질만 또렷하게(2026-10-08 사용자 확정).

    · 가로 1,800px 보다 작으면 같은 비율로 키운 뒤(인쇄 때 뭉개지지 않게) 윤곽만 살짝 또렷하게 한다.
    · 다시 저장할 때 화질이 떨어지지 않게 PNG(투명 있으면) 또는 JPEG 95 로 저장.
    반환: (bytes, (가로px, 세로px))
    """
    from PIL import ImageFilter
    im = Image.open(io.BytesIO(blob))
    has_alpha = im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info)
    im = im.convert("RGBA" if has_alpha else "RGB")
    if im.width < 1800:
        k = 1800 / im.width
        im = im.resize((1800, max(1, int(im.height * k))), Image.LANCZOS)
    im = im.filter(ImageFilter.UnsharpMask(radius=1.2, percent=70, threshold=2))
    out = io.BytesIO()
    if has_alpha:
        im.save(out, "PNG")
    else:
        im.save(out, "JPEG", quality=95, subsampling=0)
    return out.getvalue(), im.size


def _anchor_of(node):
    while node is not None and etree.QName(node).localname not in ("anchor", "inline"):
        node = node.getparent()
    return node


def _pos_v(anchor):
    po = anchor.find("wp:positionV/wp:posOffset", NS)
    return po


def _fit_cover_photo(blip_node, pw, ph):
    """사진을 원본 비율 그대로 '사진 칸' 안에 들어가게 줄이고 가로 가운데에 둔다.

    사진 칸 = 양식 사진 가로 폭 × (사진 윗자리 ~ Disclaimer 상자 위 0.4cm).
    """
    anc = _anchor_of(blip_node)
    if anc is None:
        return
    ext = anc.find("wp:extent", NS)
    max_w = int(ext.get("cx"))
    top = _pos_v(anc)
    room = int(anc.get("_room_cy") or ext.get("cy"))
    w = max_w
    h = int(w * ph / pw)
    if h > room:
        h = room
        w = int(h * pw / ph)
    ext.set("cx", str(w))
    ext.set("cy", str(h))
    for x in anc.iter("{%s}ext" % NS["a"]):
        if x.getparent() is not None and etree.QName(x.getparent()).localname == "xfrm":
            x.set("cx", str(w))
            x.set("cy", str(h))
    ph_el = anc.find("wp:positionH", NS)
    off = ph_el.find("wp:posOffset", NS) if ph_el is not None else None
    if off is not None:
        off.text = str(int(off.text) + (max_w - w) // 2)
    anc.attrib.pop("_room_cy", None)


def _text_width_pt(text, size_pt, bold=True):
    """Pretendard 로 찍었을 때 글 폭(pt) 어림 — 웹에는 글꼴 파일이 없어 글자 종류별 평균 폭으로 잰다."""
    w = 0.0
    for ch in text:
        if ord(ch) >= 0x1100:                 # 한글·한자
            w += 0.92
        elif ch == " ":
            w += 0.28
        elif ch.isupper() or ch.isdigit():
            w += 0.64
        else:
            w += 0.55
    return w * size_pt * (1.04 if bold else 1.0)


def _layout_cover(els, lines):
    """제목 상자를 글 길이에 맞춰 늘리고, 그만큼 날짜·사진을 아래로 내린다(2026-10-08 사용자 지적).

    양식 제목 상자는 두 줄 높이다. 둘째 줄이 길어 상자 폭을 넘으면 세 줄 이상이 되는데, 상자가 그대로면
    아래 날짜와 겹친다(실제: '안산 OUR IDC PF 대출 / 부산_남포동_생활형숙박시설 신축사업 PF대출 후순위').
    """
    anchors = [a for el in els for a in el.iter("{%s}anchor" % NS["wp"])]
    title = date_a = pic = disc = logo = im = None
    for a in anchors:
        t = "".join(x.text or "" for x in a.iter(qn("w:t")))
        if a.find(".//a:blip", NS) is not None:
            if int(a.find("wp:extent", NS).get("cy")) < 400000:
                logo = a                                       # 회사 로고(작은 그림)
            else:
                pic = a
        elif "Information Memorandum" in t:
            im = a
        elif "Disclaimer" in t:
            disc = a
        elif re.search(r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}", t):
            date_a = a
        elif a.find(".//" + qn("w:txbxContent")) is not None and title is None and \
                any(l and l in t for l in lines if l):
            title = a
    if title is None:
        return
    # ★표지 순서(2026-10-15 사용자 사진): 'Information Memorandum' → 회사 로고 → 제목 → 날짜 → 사진 → Disclaimer
    #   (양식은 로고가 맨 위). 로고를 IM 글 바로 아래로, 제목·날짜·사진은 그만큼 내린다.
    if logo is not None and im is not None and all(_pos_v(x) is not None for x in (logo, im, title)):
        im_y = int(_pos_v(im).text)
        im_h = int(im.find("wp:extent", NS).get("cy"))
        logo_h = int(logo.find("wp:extent", NS).get("cy"))
        new_logo = im_y + im_h + 30000
        _pos_v(logo).text = str(new_logo)
        delta = new_logo + logo_h + 120000 - int(_pos_v(title).text)
        if delta > 0:
            for a in (title, date_a, pic):
                if a is not None and _pos_v(a) is not None:
                    _pos_v(a).text = str(int(_pos_v(a).text) + delta)
    ext = title.find("wp:extent", NS)
    body_pr = title.find(".//{http://schemas.microsoft.com/office/word/2010/wordprocessingShape}bodyPr")
    l_ins = int(body_pr.get("lIns", 91440)) if body_pr is not None else 91440
    r_ins = int(body_pr.get("rIns", 91440)) if body_pr is not None else 91440
    width_pt = (int(ext.get("cx")) - l_ins - r_ins) / 12700
    paras = title.find(".//" + qn("w:txbxContent")).findall(qn("w:p"))
    need = 0
    for p in paras:
        txt = "".join(x.text or "" for x in p.iter(qn("w:t")))
        sizes = [int(s.get(qn("w:val"))) / 2 for s in p.iter(qn("w:sz")) if (s.get(qn("w:val")) or "").isdigit()]
        size = max(sizes) if sizes else 20
        if txt.strip():
            need += max(1, -(-int(_text_width_pt(txt, size)) // max(int(width_pt), 1)))
    have = sum(1 for p in paras if "".join(x.text or "" for x in p.iter(qn("w:t"))).strip()) or 1
    per_line = int(ext.get("cy")) // max(len(paras), 1)
    extra = max(0, need - max(have, len(paras))) * per_line
    if extra:
        ext.set("cy", str(int(ext.get("cy")) + extra))
        for x in title.iter("{%s}ext" % NS["a"]):
            if x.getparent() is not None and etree.QName(x.getparent()).localname == "xfrm":
                x.set("cy", str(int(x.get("cy")) + extra))
        for a in (date_a, pic):
            if a is not None and _pos_v(a) is not None:
                _pos_v(a).text = str(int(_pos_v(a).text) + extra)
    if pic is not None and disc is not None and _pos_v(pic) is not None and _pos_v(disc) is not None:
        room = int(_pos_v(disc).text) - int(_pos_v(pic).text) - 144000      # Disclaimer 위 0.4cm 띄움
        pic.set("_room_cy", str(max(room, 360000)))


def _set_par_text(p, text):
    """문단 서식은 두고 글만 바꾼다(첫 글자 조각에 몰아 넣음)."""
    ts = list(p.iter(qn("w:t")))
    if not ts:
        return
    ts[0].text = text
    ts[0].set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    for t in ts[1:]:
        t.text = ""


def _is_page_break_run(r):
    return r.tag == qn("w:r") and any(b.get(qn("w:type")) == "page" for b in r.findall(qn("w:br")))


def _split_cover_para(p):
    """양식 두 번째 문단 = [표지 글상자들][쪽나눔][회사 하이라이트 쪽 그룹들][쪽나눔] (2026-10-12 '(최종)' 양식).
    → (표지 문단, 하이라이트 문단 또는 None) 사본. 하이라이트 문단은 끝의 쪽나눔을 그대로 둬 다음(사모사채 개요)이 새 쪽."""
    cover, hl = copy.deepcopy(p), copy.deepcopy(p)
    runs_c = [c for c in cover if c.tag == qn("w:r")]
    k = next((i for i, r in enumerate(runs_c) if _is_page_break_run(r)), None)
    if k is None:
        return cover, None
    for r in runs_c[k:]:
        cover.remove(r)                                   # 표지: 첫 쪽나눔부터 뒤는 뺀다
    runs_h = [c for c in hl if c.tag == qn("w:r")]
    for r in runs_h[:k + 1]:
        hl.remove(r)                                      # 하이라이트: 표지 글상자·첫 쪽나눔까지 뺀다
    if not any(r.find(".//" + qn("w:drawing")) is not None or "".join(r.itertext()).strip()
               for r in hl.findall(qn("w:r"))):
        return cover, None
    return cover, hl


def _highlight_elements(tpl, doc):
    """회사 Investment Highlight 쪽(양식 그대로, 그룹째) — 내용은 사용자가 워드에서 고친다(2026-10-12 확정)."""
    tb = _blocks(tpl.element.body)
    hl = _split_cover_para(tb[1])[1]
    if hl is None:
        return []
    tpl_rels = tpl.part.rels
    for node in hl.iter():
        for attr in (R_EMBED, R_ID):
            rid = node.get(attr)
            if rid and rid in tpl_rels and tpl_rels[rid].reltype == RT.IMAGE:
                new_rid, _ = doc.part.get_or_add_image(io.BytesIO(tpl_rels[rid].target_part.blob))
                node.set(attr, new_rid)
    _import_styles(tpl, doc, hl)
    _import_numbering(tpl, doc, hl)
    _pin_template_defaults(doc, hl)
    return [hl]


def _import_numbering(tpl, doc, root):
    """root 가 쓰는 글머리/번호(numId)를 양식의 정의째 이 문서로 옮기고 번호를 새로 매긴다.
    ★안 옮기면 이 문서의 같은 번호 정의를 써서 양식의 '→' 가 '1. 2. 3.' 으로 나왔다(하이라이트 쪽, 2026-10-12)."""
    used = {n.get(qn("w:val")) for n in root.iter(qn("w:numId")) if n.get(qn("w:val")) not in (None, "0")}
    if not used:
        return
    try:
        tnum = tpl.part.numbering_part.element
    except Exception:
        return
    dnum = doc.part.numbering_part.element
    abs_ids = [int(a.get(qn("w:abstractNumId"))) for a in dnum.findall(qn("w:abstractNum"))] or [0]
    num_ids = [int(a.get(qn("w:numId"))) for a in dnum.findall(qn("w:num"))] or [0]
    next_abs, next_num = max(abs_ids) + 1, max(num_ids) + 1
    remap = {}
    for nid in used:
        num = next((n for n in tnum.findall(qn("w:num")) if n.get(qn("w:numId")) == nid), None)
        if num is None:
            continue
        aid = num.find(qn("w:abstractNumId")).get(qn("w:val"))
        absn = next((a for a in tnum.findall(qn("w:abstractNum")) if a.get(qn("w:abstractNumId")) == aid), None)
        if absn is None:
            continue
        new_abs = copy.deepcopy(absn)
        new_abs.set(qn("w:abstractNumId"), str(next_abs))
        for nsid in new_abs.findall(qn("w:nsid")):
            new_abs.remove(nsid)                              # 같은 nsid 면 워드가 기존 목록과 합친다
        first_num = dnum.find(qn("w:num"))
        (first_num.addprevious(new_abs) if first_num is not None else dnum.append(new_abs))  # abstractNum 은 num 보다 앞
        new_num = copy.deepcopy(num)
        new_num.set(qn("w:numId"), str(next_num))
        new_num.find(qn("w:abstractNumId")).set(qn("w:val"), str(next_abs))
        dnum.append(new_num)
        remap[nid] = str(next_num)
        next_abs += 1
        next_num += 1
    for n in root.iter(qn("w:numId")):
        if n.get(qn("w:val")) in remap:
            n.set(qn("w:val"), remap[n.get(qn("w:val"))])


def _is_cover_photo(node):
    """그 그림이 표지 사진 자리(가장 큰 그림, 높이 3cm 넘음)인가."""
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    a = node
    while a is not None and etree.QName(a).localname not in ("anchor", "inline"):
        a = a.getparent()
    ext = a.find(WP + "extent") if a is not None else None
    return ext is not None and int(ext.get("cy", "0")) > 3 * 360000


def _cover_elements(tpl, doc, title_lines, date_text, photo):
    """양식 표지(문단 0·1)를 복사해 글·사진을 바꿔 돌려준다."""
    tb = _blocks(tpl.element.body)
    els = [copy.deepcopy(tb[0]), _split_cover_para(tb[1])[0]]
    lines = (list(title_lines) + ["", ""])[:2]
    for el in els:
        for tx in el.iter(qn("w:txbxContent")):
            t = "".join(tx.itertext())
            ps = tx.findall(qn("w:p"))
            if "제안서" in t and len(ps) >= 2:            # 제목 상자(두 줄)
                _set_par_text(ps[0], lines[0])
                _set_par_text(ps[1], lines[1])
            elif re.search(r"(January|February|March|April|May|June|July|August|September|"
                           r"October|November|December)\s+\d{4}", t):
                _set_par_text(ps[0], date_text)
    # 글상자의 '대체 내용'(옛 워드용 VML)은 뺀다 — 상자 크기·자리를 바꿀 때 두 벌을 맞출 필요가 없게.
    # 지금 워드는 mc:Choice 쪽만 그린다.
    for el in els:
        for fb in list(el.iter("{%s}Fallback" % NS["mc"])):
            fb.getparent().remove(fb)
    _layout_cover(els, lines)
    # 사진: 원본 rId 를 새 문서의 그림으로 바꾸거나, 사진 없음이면 그림을 뺀다
    tpl_rels = tpl.part.rels
    for el in els:
        for node in list(el.iter()):
            for attr in (R_EMBED, R_ID):
                rid = node.get(attr)
                if not rid or rid not in tpl_rels or tpl_rels[rid].reltype != RT.IMAGE:
                    continue
                if not _is_cover_photo(node):
                    # 표지 사진이 아닌 그림(맨 위 로고 띠 '그림 16' 등)은 양식 그림 그대로
                    # ★예전엔 모든 그림을 표지 사진으로 바꿔 로고 자리에 작은 사진이 생겼다(2026-10-14 사용자 지적)
                    nid, _ = doc.part.get_or_add_image(io.BytesIO(tpl_rels[rid].target_part.blob))
                    node.set(attr, nid)
                    continue
                if photo is None:
                    box = node
                    while box is not None and etree.QName(box).localname not in ("drawing", "pict"):
                        box = box.getparent()
                    if box is not None and box.getparent() is not None:
                        box.getparent().remove(box)
                    break
                blob, (pw, ph) = _prepare_photo(photo)
                new_rid, _ = doc.part.get_or_add_image(io.BytesIO(blob))
                node.set(attr, new_rid)
                _fit_cover_photo(node, pw, ph)
    for el in els:
        _import_styles(tpl, doc, el)
        _pin_template_defaults(doc, el)
    return els


def _pin_template_defaults(doc, root):
    """양식에서 가져온 부분에 양식의 '기본값' 을 박는다(글자 10pt, 문단 앞뒤 0, 줄간격 1).

    ★양식은 기본값(글자 크기·문단 간격)을 따로 안 적고 문서 기본값에 기대는데, 원본 문서의 기본값은
      다르다. pdf2docx 문서는 기본 11pt·문단 뒤 10pt 라 표지 'Strictly Confidential' 상자가 넘쳐
      'Strictly' 만 보였다(2026-10-07). 직접 적었거나 회사양식 서식(cf_)이 정한 값은 그대로 둔다.
    """
    st = {s.get(qn("w:styleId")): s for s in doc.styles.element.findall(qn("w:style"))}

    def style_has(sid, path):
        seen = set()
        while sid in st and sid not in seen:
            seen.add(sid)
            if st[sid].find(path) is not None:
                return True
            b = st[sid].find(qn("w:basedOn"))
            sid = b.get(qn("w:val")) if b is not None else None
        return False

    for p in root.iter(qn("w:p")):
        ppr = _first_child(p, "pPr")
        ps = ppr.find(qn("w:pStyle"))
        sid = ps.get(qn("w:val")) if ps is not None else None
        if ppr.find(qn("w:spacing")) is None and not style_has(sid, qn("w:pPr") + "/" + qn("w:spacing")):
            _put(ppr, ppr.makeelement(qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0",
                                                         qn("w:line"): "240", qn("w:lineRule"): "auto"}))
        for r in p.findall(qn("w:r")) + [ppr]:
            rpr = _get(ppr, "rPr") if r is ppr else _first_child(r, "rPr")
            rs = rpr.find(qn("w:rStyle"))
            if rpr.find(qn("w:sz")) is None and not style_has(sid, qn("w:rPr") + "/" + qn("w:sz")) \
                    and not (rs is not None and style_has(rs.get(qn("w:val")), qn("w:rPr") + "/" + qn("w:sz"))):
                _put(rpr, rpr.makeelement(qn("w:sz"), {qn("w:val"): "20"}))


def _import_styles(tpl, doc, root):
    """root 안에서 쓰는 양식 서식을 원본에 'cf_<이름>' 으로 복사해 붙이고, 참조를 그쪽으로 바꾼다.

    ★서식 이름(styleId)이 같아도 문서마다 뜻이 다르다. 안산 원본에서는 양식의 표지 서식 이름이
      '제N장' 번호 서식이라 표지에 '제1장…제6장' 이, 머리말 앞에 '1)' 이 붙었다(2026-10-07).
    """
    tst = tpl.styles.element
    dst = doc.styles.element
    by_id = {s.get(qn("w:styleId")): s for s in tst.findall(qn("w:style"))}
    have = {s.get(qn("w:styleId")) for s in dst.findall(qn("w:style"))}

    def bring(sid):
        if sid not in by_id:
            return None
        nid = "cf_" + sid
        if nid in have:
            return nid
        have.add(nid)
        st = copy.deepcopy(by_id[sid])
        st.set(qn("w:styleId"), nid)
        st.attrib.pop(qn("w:default"), None)
        nm = st.find(qn("w:name"))
        if nm is not None:
            nm.set(qn("w:val"), "회사양식 " + nm.get(qn("w:val")))
        for numpr in list(st.iter(qn("w:numPr"))):         # 번호 매기기는 가져오지 않는다
            numpr.getparent().remove(numpr)
        for tag in ("w:basedOn", "w:next", "w:link"):
            ref = st.find(qn(tag))
            if ref is not None:
                new = bring(ref.get(qn("w:val")))
                if new:
                    ref.set(qn("w:val"), new)
                else:
                    st.remove(ref)
        dst.append(st)
        return nid

    for tag in ("w:pStyle", "w:rStyle", "w:tblStyle"):
        for s in list(root.iter(qn(tag))):
            new = bring(s.get(qn("w:val")))
            if new:
                s.set(qn("w:val"), new)
            else:
                s.getparent().remove(s)


def _run(p, text, size, bold=False, color=None, font=None):
    r = p.add_run(text)
    r.font.size = Pt(size)
    rpr = r._r.get_or_add_rPr()
    if font:
        r.bold = bold
        _set_rfonts(rpr, font)
    else:
        _font_weight(rpr, bold)
    if color:
        from docx.shared import RGBColor
        r.font.color.rgb = RGBColor.from_string(color)
    return r


def _para_fmt(p, align="left", before=0, after=0):
    from docx.enum.text import WD_ALIGN_PARAGRAPH as A
    p.alignment = {"left": A.LEFT, "center": A.CENTER}[align]
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)


def _cell_shade(cell, fill):
    tcpr = cell._tc.get_or_add_tcPr()
    _put(tcpr, tcpr.makeelement(qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): fill}))


def _table_borders(tbl, color="000000", sz=4):
    tblpr = tbl._tbl.tblPr
    b = _put(tblpr, tblpr.makeelement(qn("w:tblBorders"), {}))
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        e = etree.SubElement(b, qn("w:" + side))
        e.set(qn("w:val"), "single")
        e.set(qn("w:sz"), str(sz))
        e.set(qn("w:space"), "0")
        e.set(qn("w:color"), color)


def _fixed_widths(tbl, widths_twips):
    from docx.shared import Twips
    tbl.autofit = False
    for row in tbl.rows:
        for c, w in zip(row.cells, widths_twips):
            c.width = Twips(w)
    grid = tbl._tbl.tblGrid
    for gc, w in zip(grid.findall(qn("w:gridCol")), widths_twips):
        gc.set(qn("w:w"), str(w))


def _row_height(row, twips):
    from docx.shared import Twips
    from docx.enum.table import WD_ROW_HEIGHT_RULE
    row.height = Twips(twips)
    row.height_rule = WD_ROW_HEIGHT_RULE.AT_LEAST


def _vcenter(cell):
    from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT as V
    cell.vertical_alignment = V.CENTER


def _build_bond_section(doc, bond):
    """'Ⅰ. 사모사채 개요' 제목 + 9줄 표. 새로 만든 요소들을 돌려준다(본문 끝에 붙었다가 떼어 냄)."""
    body = doc.element.body
    # ★id() 로 가르면 안 된다 — lxml 은 파이썬 쪽 객체를 버렸다 다시 만들어 번호가 재사용된다
    #   (그래서 연락처 표가 2쪽에, 개요 제목이 맨 끝에 가 있었다). 목록을 붙들고 'is' 로 가른다.
    before = list(body)
    h = doc.add_paragraph()
    _para_fmt(h, "left", 0, 10)
    _run(h, "Ⅰ. 사모사채 개요", 16, bold=True)
    t = doc.add_table(rows=len(BOND_ROWS), cols=2)
    _table_borders(t)
    _fixed_widths(t, [1448, 7619])
    for i, key in enumerate(BOND_ROWS):
        a, b = t.rows[i].cells
        _row_height(t.rows[i], 454)
        _cell_shade(a, LABEL)
        for c in (a, b):
            _vcenter(c)
        pa, pb = a.paragraphs[0], b.paragraphs[0]
        _para_fmt(pa, "center")
        _para_fmt(pb, "left")
        _run(pa, "사채 유형" if key == "사채유형" else key, 10)
        _run(pb, (bond.get(key) or "").strip(), 10)
    return [e for e in body if not any(e is b for b in before) and e.tag != qn("w:sectPr")]


def _build_contact_section(doc, contacts):
    body = doc.element.body
    # ★id() 로 가르면 안 된다 — lxml 은 파이썬 쪽 객체를 버렸다 다시 만들어 번호가 재사용된다
    #   (그래서 연락처 표가 2쪽에, 개요 제목이 맨 끝에 가 있었다). 목록을 붙들고 'is' 로 가른다.
    before = list(body)
    cols = ["본부", "직급", "이름", "E-mail", "연락처"]
    t = doc.add_table(rows=1 + len(contacts), cols=len(cols))
    _table_borders(t)
    _fixed_widths(t, [2302, 1170, 1459, 2182, 1903])
    for j, name in enumerate(cols):
        c = t.rows[0].cells[j]
        _cell_shade(c, NAVY)
        _vcenter(c)
        p = c.paragraphs[0]
        _para_fmt(p, "center")
        _run(p, name, 9, bold=True, color="FFFFFF")
    _row_height(t.rows[0], 397)
    for i, row in enumerate(contacts, start=1):
        _row_height(t.rows[i], 397)
        for j, key in enumerate(cols):
            c = t.rows[i].cells[j]
            _vcenter(c)
            p = c.paragraphs[0]
            _para_fmt(p, "center")
            _run(p, row.get(key, ""), 9)
    p = doc.add_paragraph()
    _para_fmt(p, "left", 8, 0)
    _run(p, "본 자료와 관련하여 문의사항이 있으실 경우 아래의 담당자에게 언제든지 연락하여 주시면 "
            "성심껏 답변하여 드리겠습니다.", 9)
    p = doc.add_paragraph()
    _para_fmt(p, "center", 220, 0)
    _run(p, "끝", 11)
    return [e for e in body if not any(e is b for b in before) and e.tag != qn("w:sectPr")]


def _copy_from_template(tpl, doc, el):
    """양식 요소 사본 — 그림·하이퍼링크 연결과 서식까지 이 문서로 옮긴다."""
    new = copy.deepcopy(el)
    rels = tpl.part.rels
    for node in new.iter():
        for attr in (R_EMBED, R_ID):
            rid = node.get(attr)
            if not rid or rid not in rels:
                continue
            rel = rels[rid]
            if rel.reltype == RT.IMAGE:
                nid, _ = doc.part.get_or_add_image(io.BytesIO(rel.target_part.blob))
                node.set(attr, nid)
            elif rel.reltype == RT.HYPERLINK:
                node.set(attr, doc.part.relate_to(rel.target_ref, RT.HYPERLINK, is_external=True))
    _import_styles(tpl, doc, new)
    _pin_template_defaults(doc, new)
    return new


def _tpl_body(tpl):
    return [el for el in tpl.element.body if el.tag != qn("w:sectPr")]


def read_bond_defaults(tpl_path=None):
    """양식 사모사채 개요 표의 기본 글(화면 2단계 기본값). {구분: 내용}"""
    tpl = docx.Document(tpl_path or template_path())
    for t in tpl.tables:
        rows = t.rows
        if rows and rows[0].cells[0].text.strip() == "사모사채명":
            return {(r.cells[0].text.strip().replace("사채 유형", "사채유형")): r.cells[-1].text.strip() for r in rows}
    return {}


def _sub_heading_from_original(doc):
    """원본 '대단원 아래 소제목' 문단(예: '1. 본건 대출 금융조건') — 대단원 무리를 뺀 번호 문단 중 가장 큰 글자의 번호 모양 무리의 첫 문단.
    번호 상자식 대단원이면 번호 문단 중 가장 큰 무리가 소제목. (문단, 종류, 매치) 또는 None."""
    found = _para_heads(doc)
    majors = [h["el"] for h in _major_headings(doc) if not h["box"]]
    lv = _level(found, exclude=majors)
    if not lv:
        return None
    sz, kind, p, m = lv[0]
    return p, kind, m


def _make_sub_heading(src, n, title):
    """원본 소제목 문단을 복사해 번호 n·제목만 바꾼다. src = (_sub_heading_from_original 결과) 또는 None."""
    if src is None:
        p = etree.Element(qn("w:p"))
        ppr = etree.SubElement(p, qn("w:pPr"))
        etree.SubElement(ppr, qn("w:spacing"), {qn("w:before"): "120", qn("w:after"): "120"})
        r = etree.SubElement(p, qn("w:r"))
        rpr = etree.SubElement(r, qn("w:rPr"))
        etree.SubElement(rpr, qn("w:sz"), {qn("w:val"): "22"})
        _font_weight(rpr, True)
        etree.SubElement(r, qn("w:t")).text = f"{n}. {title}"
        return p
    p0, kind, m = src
    new = copy.deepcopy(p0)
    num = (_ROMAN if m.group(2).strip() in _ROMAN else _ROMAN_ASCII)[n] if kind == "roman" else str(n)
    _set_visible_text(new, m.group(1) + num + m.group(3) + title)
    for sp in new.iter(qn("w:sectPr")):
        sp.getparent().remove(sp)
    sp = new.find(f"{qn('w:pPr')}/{qn('w:spacing')}")
    if sp is not None:
        sp.set(qn("w:before"), "120")
    ind = new.find(f"{qn('w:pPr')}/{qn('w:ind')}")
    if ind is not None:
        ind.set(qn("w:right"), "0")
    return new


def _template_bond_els(tpl, doc, bond, diagram_png=None, sub=None):
    """양식의 [사모사채 개요 표][빈 줄][금융 구조도 제목] 을 복사하고, 표 내용은 입력값(없으면 양식 글), 제목 아래 구조도 그림."""
    els = _tpl_body(tpl)
    ti = next((i for i, el in enumerate(els) if el.tag == qn("w:tbl")
               and "사모사채명" in "".join(t.text or "" for t in el.iter(qn("w:t")))), None)
    if ti is None:
        return None
    out = []
    tbl = _copy_from_template(tpl, doc, els[ti])
    for fl in tbl.findall(f"{qn('w:tblPr')}/{qn('w:tblpPr')}"):   # 떠 있는 표 설정을 빼야 제목 아래에 순서대로 놓인다
        fl.getparent().remove(fl)
    ov = tbl.find(f"{qn('w:tblPr')}/{qn('w:tblOverlap')}")
    if ov is not None:
        ov.getparent().remove(ov)
    for tr in tbl.findall(qn("w:tr")):
        tcs = tr.findall(qn("w:tc"))
        if len(tcs) < 2:
            continue
        key = "".join(t.text or "" for t in tcs[0].iter(qn("w:t"))).strip().replace("사채 유형", "사채유형")
        val = (bond or {}).get(key, "").strip()
        if val:
            p = tcs[-1].find(qn("w:p"))
            if p.find(".//" + qn("w:t")) is None:
                r = etree.SubElement(p, qn("w:r"))
                etree.SubElement(r, qn("w:t"))
            _set_par_text(p, val)
            for extra in tcs[-1].findall(qn("w:p"))[1:]:
                tcs[-1].remove(extra)
    out.append(tbl)
    hi = next((i for i in range(ti + 1, len(els)) if "구조도" in "".join(t.text or "" for t in els[i].iter(qn("w:t")))), None)
    if hi is not None and diagram_png:                       # 구조도를 안 만들었으면 '금융 구조도' 제목도 안 넣는다
        out.append(_make_sub_heading(sub, 2, "금융 구조도"))   # 2026-10-14: 원본 소제목 모양(개요와 같은 글꼴·크기)
        for el in []:
            c = _copy_from_template(tpl, doc, el)
            for br in list(c.iter(qn("w:br"))):          # 제목 뒤 쪽나눔은 뺀다(구조도 그림이 그 아래로 와야 한다)
                if br.get(qn("w:type")) == "page":
                    br.getparent().remove(br)
            out.append(c)
        if diagram_png:
            from docx.shared import Cm
            from docx.shared import Cm
            inline = doc.part.new_pic_inline(io.BytesIO(diagram_png), Cm(16.0), None)   # 양식 본문 폭
            p = etree.Element(qn("w:p"))
            ppr = etree.SubElement(p, qn("w:pPr"))
            etree.SubElement(ppr, qn("w:spacing"), {qn("w:before"): "120", qn("w:after"): "0"})
            etree.SubElement(ppr, qn("w:jc"), {qn("w:val"): "center"})
            r = etree.SubElement(p, qn("w:r"))
            dr = etree.SubElement(r, qn("w:drawing"))
            dr.append(inline)
            out.append(p)
    return out


def _template_contact_els(tpl, doc):
    """양식 연락처 부분('본 자료와 관련하여…' 부터 '끝' 까지) 그대로."""
    els = _tpl_body(tpl)
    si = next((i for i, el in enumerate(els) if "본 자료와 관련하여" in "".join(t.text or "" for t in el.iter(qn("w:t")))), None)
    if si is None:
        return None
    # 앞쪽 빈 문단(구조도 제목 뒤 여백)부터가 아니라 안내문부터
    out = [_copy_from_template(tpl, doc, el) for el in els[si:]]
    WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    for el in out:                                           # 글상자는 글에 맞춰 늘어나게(담당자 셋째 줄이 잘렸다)
        for bp in el.iter("{%s}bodyPr" % WPS):
            for ch in list(bp):
                if etree.QName(ch).localname in ("noAutofit", "normAutofit", "spAutoFit"):
                    bp.remove(ch)
            bp.append(etree.Element("{%s}spAutoFit" % A))
    return out


def read_contacts(tpl):
    """양식 마지막 쪽 연락처 표에서 담당자 목록을 읽는다(빈 칸은 건너뜀)."""
    for t in tpl.tables:
        head = [c.text.strip() for c in t.rows[0].cells]
        if "이름" in head and "연락처" in head:
            out, cols = [], ["본부", "직급", "이름", "E-mail", "연락처"]
            for r in t.rows[1:]:
                vals = []
                seen = set()
                for c in r.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    vals.append(c.text.strip())
                vals = [v for v in vals]
                # 사용자 양식에는 빈 칸이 하나 끼어 있다(6칸) → 머리글 이름으로 맞춘다
                cells = [c.text.strip() for c in r.cells]
                row = {}
                for k in cols:
                    j = head.index(k)
                    row[k] = cells[j]
                if not row["본부"]:
                    row["본부"] = next((v for v in cells[:head.index("직급")] if v), "")
                out.append(row)
            return out
    return []


def _copy_hdrftr(tpl, doc, src_part, dst_hf, header_text=None):
    """양식 머리말/꼬리말 내용을 이 문서의 머리말/꼬리말로 옮긴다(그림 연결까지)."""
    # ★조각만 옮기면 안 된다 — 글상자의 mc:Choice Requires="wps" 가 가리키는 이름표(xmlns:wps)가
    #   맨 위(w:hdr)에 선언돼 있어야 워드가 연다. 그래서 양식 머리말 전체를 선언째 통째로 바꿔 끼운다.
    #   (조각만 옮겼을 때 워드가 '파일이 손상되었습니다' 로 안 열었다 — 2026-10-07)
    part = dst_hf.part
    part._element = copy.deepcopy(src_part.element)
    dst = part._element
    _import_styles(tpl, doc, dst)
    _pin_template_defaults(doc, dst)
    for node in dst.iter():
        for attr in (R_EMBED, R_ID):
            rid = node.get(attr)
            if rid and rid in src_part.rels and src_part.rels[rid].reltype == RT.IMAGE:
                blob = src_part.rels[rid].target_part.blob
                new_rid, _ = dst_hf.part.get_or_add_image(io.BytesIO(blob))
                node.set(attr, new_rid)
    if header_text is not None:
        for p in dst.iter(qn("w:p")):
            if "".join(p.itertext()).strip():
                _set_par_text(p, header_text)
                break
    # 쪽번호 8pt 는 양식 파일 자체에 넣었다(양식\_틀만들기.py) — 여기서 따로 안 바꾼다


PAGE_NO_SZ = 16              # 쪽번호 글자 크기 8pt(반 pt 단위) — 2026-10-12 사용자 확정(양식은 11pt 라 너무 컸다)


def _page_no_size(root):
    """PAGE 필드(쪽번호)가 든 문단의 글자 크기를 PAGE_NO_SZ 로."""
    for p in root.iter(qn("w:p")):
        if not any("PAGE" in (t.text or "") for t in p.iter(qn("w:instrText"))) and \
                not any("PAGE" in (f.get(qn("w:instr")) or "") for f in p.iter(qn("w:fldSimple"))):
            continue
        for r in p.iter(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                rpr = r.makeelement(qn("w:rPr"), {})
                r.insert(0, rpr)
            for tag in ("w:sz", "w:szCs"):
                e = rpr.find(qn(tag))
                if e is None:
                    e = rpr.makeelement(qn(tag), {})
                    rpr.append(e)
                e.set(qn("w:val"), str(PAGE_NO_SZ))


def align_tables_to_body(doc):
    """맨 바깥 표의 왼쪽 선을 머리말 가로줄 시작에, 본문 폭 표는 오른쪽 선을 줄 끝에 딱 맞춘다(2026-10-15 사용자:
    '머리말 줄 시작점에서 똑같이 시작해야지 왜 먼저 시작하냐'). ★워드는 표 선을 '들여쓰기 − 첫 칸 왼쪽 여백' 에 그려
    표가 줄보다 4~5pt 왼쪽에서 시작했다(고산 실측: 줄 39.7pt, 표 34.7pt). 고친 표 수."""
    import pdf_input as _pi
    body_w = int(round(_pi.BOX_W_PT * 20))
    n = 0
    cover_end = next((i for i, el in enumerate(doc.element.body)
                      if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None), -1)
    for i_el, tbl in enumerate(doc.element.body):
        if tbl.tag != qn("w:tbl") or i_el < cover_end:
            continue                                           # 표지 쪽 표는 양식 그대로
        tpr = tbl.find(qn("w:tblPr"))
        grid = tbl.find(qn("w:tblGrid"))
        if tpr is None or grid is None:
            continue
        cols = grid.findall(qn("w:gridCol"))
        try:
            total = sum(int(float(c.get(qn("w:w")) or 0)) for c in cols)
        except ValueError:
            continue
        if total <= 0:
            continue
        ind = tpr.find(qn("w:tblInd"))
        try:
            cur = int(float(ind.get(qn("w:w")) or 0)) if ind is not None else 0
        except ValueError:
            cur = 0
        if tpr.find(qn("w:tblpPr")) is not None:
            continue                                           # 떠 있는 표는 그대로
        first_tc = tbl.find(f"{qn('w:tr')}/{qn('w:tc')}")
        mar = 0
        m = first_tc.find(f"{qn('w:tcPr')}/{qn('w:tcMar')}/{qn('w:left')}") if first_tc is not None else None
        if m is None:
            m = tpr.find(f"{qn('w:tblCellMar')}/{qn('w:left')}")
        if m is not None:
            try:
                mar = int(float(m.get(qn("w:w")) or 0))
            except ValueError:
                mar = 0
        else:
            mar = 108                                          # 칸 여백을 안 정한 표(양식 표) = 워드 기본 0.19cm
        left = cur - mar                                       # 지금 표 선 자리(본문 왼쪽 기준)
        is_bond = "사모사채명" in _vis_text(tbl)[:40]
        full = (total >= body_w * 0.9 and abs(left) < 400) or is_bond   # 사모사채 개요 표도 본문 폭(사용자 2번)
        if full:                                               # 본문 폭 표 → 줄 시작~끝에 딱
            k = body_w / total
            acc = 0
            for i, c in enumerate(cols):
                w = int(round(int(float(c.get(qn("w:w")))) * k)) if i < len(cols) - 1 else body_w - acc
                c.set(qn("w:w"), str(w))
                acc += w
            for tr in tbl.findall(qn("w:tr")):                 # 칸 폭도 같은 비율로
                for tc in tr.findall(qn("w:tc")):
                    tw = tc.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
                    if tw is not None and (tw.get(qn("w:type")) in (None, "dxa")):
                        try:
                            tw.set(qn("w:w"), str(int(round(int(float(tw.get(qn("w:w")))) * k))))
                        except ValueError:
                            pass
            tw_ = tpr.find(qn("w:tblW"))
            if tw_ is not None:
                tw_.set(qn("w:w"), str(body_w))
                tw_.set(qn("w:type"), "dxa")
            new_left = 0
        else:
            new_left = max(0, left)
        if ind is None:
            ind = tpr.makeelement(qn("w:tblInd"), {})
            _put(tpr, ind)
        ind.set(qn("w:w"), str(new_left + mar))
        ind.set(qn("w:type"), "dxa")
        for jc in tpr.findall(qn("w:jc")):
            tpr.remove(jc)
        n += 1
    return n


def apply_company_geometry(doc):
    """모든 구역(표지·개요·본문·연락처) 규격을 같게 + 머리말·꼬리말 가로 선을 본문 폭 그대로.
    ★양식 머리말·꼬리말 선은 고정 폭(16cm) 네모 — 규격을 바꾸면 선 길이가 본문과 안 맞는다(내곡동 6쪽 비뚤어짐)."""
    import pdf_input as _pi
    from docx.shared import Cm
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(_pi.PAGE_W), Cm(_pi.PAGE_H)
        sec.top_margin, sec.bottom_margin = Cm(_pi.TOP_CM), Cm(_pi.BOTTOM_CM)
        sec.left_margin, sec.right_margin = Cm(_pi.LEFT_CM), Cm(_pi.RIGHT_CM)
        sec.header_distance, sec.footer_distance = Cm(_pi.HEADER_CM), Cm(_pi.FOOTER_CM)
        for t in sec._sectPr.findall(qn("w:pgMar")):
            t.set(qn("w:gutter"), "0")
    body_emu = int(round(_pi.BOX_W_PT * 12700))
    WP = "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}"
    A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
    seen = set()
    for sec in doc.sections:
        for hf in (sec.header, sec.footer):
            el = hf._element
            if id(hf.part) in seen:
                continue
            seen.add(id(hf.part))
            for anc in el.iter(WP + "anchor"):
                ext = anc.find(WP + "extent")
                if ext is None:
                    continue
                cx, cy = int(ext.get("cx")), int(ext.get("cy"))
                if cy < 100000 and cx > 3000000:                 # 가로 선(얇고 긴 네모)
                    ext.set("cx", str(body_emu))
                    ph = anc.find(WP + "positionH")
                    off = ph.find(WP + "posOffset") if ph is not None else None
                    if off is not None:
                        off.text = "0"
                    for x in anc.iter(A + "ext"):
                        if x.getparent().tag == A + "xfrm":
                            x.set("cx", str(body_emu))


def _renumber_shapes(doc):
    """그림·도형 번호(wp:docPr id)를 문서 전체에서 겹치지 않게 다시 매긴다.

    ★양식에서 옮겨 온 표지·꼬리말 도형 번호가 원본 그림 번호와 겹치면 워드가
      '파일이 손상되었습니다' 로 안 연다(실제로 겪음 — 꼬리말 도형 3·4번이 원본과 겹침).
    """
    n = 0
    for part in doc.part.package.iter_parts():
        el = getattr(part, "_element", None)
        if el is None:
            continue
        for dp in el.iter("{%s}docPr" % NS["wp"]):
            n += 1
            dp.set("id", str(n))


def _merge_nsmap(doc, tpl):
    """양식 표지 글상자가 쓰는 이름표(wps 등) 선언을 원본 문서 맨 위에 합친다.

    mc:Choice Requires="wps" 는 맨 위에 xmlns:wps 가 있어야 워드가 연다. 워드로 만든 문서는 대개
    있지만, 다른 프로그램(pdf2docx 등)이 만든 문서에는 없을 수 있다.
    """
    root = doc.part._element
    extra = {k: v for k, v in tpl.part.element.nsmap.items() if k and k not in root.nsmap}
    if not extra:
        return
    # python-docx 의 파서로 만들어야 '문서' 요소(body 등)로 동작한다 — 그냥 lxml 로 만들면 기능이 없다
    from docx.oxml.parser import oxml_parser
    new = oxml_parser.makeelement(root.tag, nsmap={**root.nsmap, **extra})
    for k, v in root.attrib.items():
        new.set(k, v)
    for ch in list(root):
        new.append(ch)
    doc.part._element = new
    doc._element = new
    doc._Document__body = None          # python-docx 가 붙들고 있던 본문 객체를 새 뿌리로 다시 잡게


def _sect_from_template(tpl, title_page):
    sp = copy.deepcopy(tpl.element.body.find(qn("w:sectPr")))
    for tag in ("w:headerReference", "w:footerReference"):
        for e in sp.findall(qn(tag)):
            sp.remove(e)
    tp = sp.find(qn("w:titlePg"))
    if not title_page and tp is not None:
        sp.remove(tp)
    for a in [k for k in sp.attrib]:
        del sp.attrib[a]
    return sp


def _section_groups(doc):
    """구역마다 블록 묶음(구역은 '문단 안 sectPr' 로 끝나고, 마지막은 본문 sectPr)."""
    groups, cur = [], []
    for el in doc.element.body:
        if el.tag not in (qn("w:p"), qn("w:tbl")):
            continue
        cur.append(el)
        if el.tag == qn("w:p") and el.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
            groups.append(cur)
            cur = []
    groups.append(cur)
    return groups


def _holder(sect):
    """구역 끝 문단(문단 안 sectPr)."""
    from docx.oxml.parser import OxmlElement            # python-docx 요소로 만들어야 문단 기능이 붙는다
    p = OxmlElement("w:p")
    ppr = OxmlElement("w:pPr")
    p.append(ppr)
    ppr.append(sect)
    return p


def _place_bond(doc, tpl, bond_els, place):
    """사모사채 개요를 넣는다(2026-10-08 사용자 확정).

    place = None          → 표지 바로 다음 쪽(하이라이트를 못 찾았을 때)
    place = (k, same)     → 본문 k 번째 구역(=원본 하이라이트 마지막 쪽) 뒤.
                            same=True 면 그 쪽 남은 여백에 이어서, False 면 새 쪽.
    """
    body = doc.element.body
    groups = [g for g in _section_groups(doc) if g]
    if place is None or not groups:
        first = groups[0][0] if groups else body.find(qn("w:sectPr"))
        for el in bond_els:
            first.addprevious(el)
        first.addprevious(_holder(_sect_from_template(tpl, title_page=False)))
        return
    k, same = place
    k = min(k, len(groups) - 1)
    grp = groups[k]
    last = grp[-1]
    ends_section = last.tag == qn("w:p") and last.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None
    if same:
        # 그 쪽 끝의 빈 문단(pdf2docx 가 쪽 아래를 채운 것)을 걷어내고 바로 뒤에 잇는다
        for el in reversed(grp[:-1] if ends_section else grp):
            if el.tag == qn("w:p") and not "".join(el.itertext()).strip() \
                    and el.find(".//" + qn("w:drawing")) is None:
                body.remove(el)
            else:
                break
        anchor = last if ends_section else None
        for el in bond_els:
            if anchor is not None:
                anchor.addprevious(el)
            else:
                body.find(qn("w:sectPr")).addprevious(el)
        return
    if ends_section:
        for el in reversed(bond_els + [_holder(_sect_from_template(tpl, title_page=False))]):
            last.addnext(el)
    else:
        # 마지막 구역 뒤 — 지금 구역을 닫고(본문 sectPr 복사) 그 뒤에 개요를 둔다
        sp = body.find(qn("w:sectPr"))
        sp.addprevious(_holder(copy.deepcopy(sp)))
        for el in bond_els:
            sp.addprevious(el)


def _vis_text(el):
    return "".join(t.text or "" for t in el.iter(qn("w:t")))


def _set_heading_text(p, head_len, new_head, title):
    """번호 문단 글 바꾸기 — 원본 글 앞 head_len 글자(번호+띄움)가 따로 된 글 조각이면 번호 조각·제목 조각 서식을 각각 살린다.
    ★예전엔 글 전체를 첫 조각(번호 조각)에 넣어 제목 글꼴·굵기가 원본과 달라졌다('본건 사모사채 개요', 2026-10-14)."""
    ts = [t for t in p.iter(qn("w:t"))]
    pos, cut = 0, None
    for k, t in enumerate(ts):
        pos += len(t.text or "")
        if pos == head_len:
            cut = k
            break
        if pos > head_len:
            break
    if cut is None or cut + 1 >= len(ts) or not new_head:
        _set_par_text(p, (new_head or "") + title)
        return
    for t in ts:
        t.text = ""
    ts[0].text = new_head
    ts[cut + 1].text = title
    for t in (ts[0], ts[cut + 1]):
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")


def _head_match(txt):
    """번호 제목 글이면 (종류, 매치), 아니면 None."""
    if not txt.strip() or len(txt.strip()) > 40:
        return None
    for kind, rx in _HEAD_PATTERNS:
        m = rx.match(txt)
        if m and m.group(2) and _num_value(kind, m.group(2)) is not None:
            return kind, m
    return None


def _max_sz(el):
    szs = [int(x.get(qn("w:val"))) for x in el.iter(qn("w:sz")) if (x.get(qn("w:val")) or "").isdigit()]
    return max(szs) if szs else 20


def split_heading_breaks(doc):
    """'Ⅱ.금융조건 ↵ 1.주요 금융 조건' 처럼 줄바꿈으로 한 문단에 붙은 제목들을 문단 하나씩으로 나눈다.
    ★pdf2docx 가 대단원과 소제목을 한 문단에 넣어 대단원을 못 찾았다(내곡동, 2026-10-15). 나눈 문단 수."""
    n = 0
    body = doc.element.body
    for p in list(body.findall(qn("w:p"))):
        if not any(br.get(qn("w:type")) in (None, "textWrapping") for br in p.iter(qn("w:br"))):
            continue
        segs, cur, part = [], [], None
        for ch in p:
            if ch.tag == qn("w:pPr"):
                continue
            if ch.tag != qn("w:r"):
                cur.append(copy.deepcopy(ch))
                continue
            rpr = ch.find(qn("w:rPr"))
            part = None
            for c in ch:
                if c.tag == qn("w:rPr"):
                    continue
                if c.tag == qn("w:br") and c.get(qn("w:type")) in (None, "textWrapping"):
                    segs.append(cur)
                    cur, part = [], None
                    continue
                if part is None:
                    part = ch.makeelement(qn("w:r"), {})
                    if rpr is not None:
                        part.append(copy.deepcopy(rpr))
                    cur.append(part)
                part.append(copy.deepcopy(c))
        segs.append(cur)
        texts = ["".join(t.text or "" for s in seg for t in s.iter(qn("w:t"))) for seg in segs]
        if len(segs) < 2 or _head_match(texts[0]) is None:
            continue
        ppr = p.find(qn("w:pPr"))
        made = []
        for seg, txt in zip(segs[:-1], texts[:-1]):
            if not txt.strip():
                continue
            np_ = p.makeelement(qn("w:p"), {})
            if ppr is not None:
                c = copy.deepcopy(ppr)
                for sp in c.findall(qn("w:sectPr")):
                    c.remove(sp)
                np_.append(c)
            for x in seg:
                np_.append(x)
            p.addprevious(np_)
            made.append(np_)
            n += 1
        for ch in [c for c in p if c.tag != qn("w:pPr")]:
            p.remove(ch)
        for x in segs[-1]:
            p.append(x)
        for q in [p] + made:
            _left_heading(q)
    for f in _para_heads(doc):                               # 제목 번호 뒤 띄어쓰기(원본 'II. 금융조건', pdf2docx 가 빼먹음)
        _space_after_number(f[2], f[3])
    return n


def _left_heading(p):
    """pdf2docx 가 '가운데 맞춤 + 큰 오른쪽 들여쓰기' 로 왼쪽 자리를 흉내 낸 제목 → 왼쪽 맞춤, 들여쓰기 0(원본 자리)."""
    ppr = p.find(qn("w:pPr"))
    if ppr is None:
        return
    jc = ppr.find(qn("w:jc"))
    ind = ppr.find(qn("w:ind"))
    try:
        right = int(float(ind.get(qn("w:right")) or 0)) if ind is not None else 0
    except ValueError:
        right = 0
    if jc is not None and jc.get(qn("w:val")) == "center" and right > 1500:
        jc.set(qn("w:val"), "left")
        ind.set(qn("w:right"), "0")
        ind.set(qn("w:left"), "0")


def _space_after_number(p, m):
    """번호 바로 뒤에 띄어쓰기가 없으면 한 칸('2.자금조달' → '2. 자금조달')."""
    if m.group(3) and m.group(3)[-1].isspace():
        return
    head_len = m.end(3)
    pos = 0
    for t in p.iter(qn("w:t")):
        L = len(t.text or "")
        if pos + L >= head_len:
            k = head_len - pos
            t.text = (t.text or "")[:k] + " " + (t.text or "")[k:]
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            return
        pos += L


def _para_heads(doc):
    """맨 바깥 번호 제목 문단 [(크기, 종류, 문단, 매치)] 문서 순서."""
    out = []
    for p in doc.element.body.findall(qn("w:p")):
        hm = _head_match(_vis_text(p))
        if hm:
            out.append((_max_sz(p), hm[0], p, hm[1]))
    return out


def _level(found, exclude=()):
    """가장 큰 글자 제목의 번호 종류 무리(그 종류 중 가장 큰 글자의 85% 이상). [(크기, 종류, 문단, 매치)]
    ★쪽마다 줄인 비율이 달라 같은 대단원도 13.5~15.5pt 로 제각각 — 글자 크기 하나로 묶으면 안 된다(내곡동)."""
    rest = [f for f in found if f[2] not in exclude]
    if not rest:
        return []
    top = max(rest, key=lambda f: f[0])
    kind = top[1]
    lim = top[0] * 0.85
    return [f for f in rest if f[1] == kind and f[0] >= lim]


def _major_headings(doc):
    """원본 대단원 제목 목록(문서 순서). [{el, box, kind, tok, title, m}]
    번호 상자 표(1줄 2칸, 칠한 번호 칸)가 있으면 그것들, 없으면 가장 큰 글자 번호 문단 무리(2개 이상)."""
    body = doc.element.body
    boxes = []
    for el in body:
        if el.tag != qn("w:tbl"):
            continue
        row, _ = _heading_row(el)
        if row is None or len(el.findall(qn("w:tr"))) != 1:
            continue
        tcs = row.findall(qn("w:tc"))
        if len(tcs) != 2:
            continue
        num = _vis_text(tcs[0]).strip().rstrip(".")
        title = _vis_text(tcs[1]).strip()
        if num and re.match(re.escape(num) + r"\.?\s*[^\d\s.]", title):   # 제목 칸에도 번호가 겹쳐 들어온 것(돈암 '1 Executive summary', '3사업개요')
            title = re.sub("^" + re.escape(num) + r"\.?\s*", "", title)
        if _SEC_NO_RE.fullmatch(num or "x") and title and len(title) <= 40 and _fill(tcs[0]):
            boxes.append({"el": el, "box": True, "kind": "digit" if num.isdigit() else "roman", "tok": num,
                          "title": title, "m": None})
    if boxes:
        return boxes
    found = _para_heads(doc)
    lv = _level(found)
    if lv:
        # ★번호가 차례로 이어지는데 빠진 번호(글자가 작게 줄어든 대단원 — 효자 'II. 금융구조도'·'III. 금융 조건')를 채운다
        kind = lv[0][1]
        have = {_num_value(kind, f[3].group(2)) for f in lv}
        order = list(doc.element.body)
        idx = {id(e): i for i, e in enumerate(order)}
        for f in found:
            if f[1] != kind or f in lv or f[2] not in order:
                continue
            v = _num_value(kind, f[3].group(2))
            if v is None or v in have or f[0] < max(x[0] for x in lv) * 0.6:
                continue
            prev = [x for x in lv if _num_value(kind, x[3].group(2)) < v]
            nxt = [x for x in lv if _num_value(kind, x[3].group(2)) > v]
            if prev and nxt and idx[id(max(prev, key=lambda x: idx[id(x[2])])[2])] < idx[id(f[2])] <                     idx[id(min(nxt, key=lambda x: idx[id(x[2])])[2])]:
                lv.append(f)
                have.add(v)
        lv.sort(key=lambda f: idx.get(id(f[2]), 0))
    return [{"el": p, "box": False, "kind": kind, "tok": m.group(2), "title": m.group(4).strip(), "m": m}
            for sz, kind, p, m in lv]


def _fmt_num(kind, tok, v):
    if kind == "roman":
        return (_ROMAN if tok.strip() in _ROMAN else _ROMAN_ASCII)[min(v, 12)]
    return f"{v:0{len(tok)}d}" if tok.startswith("0") else str(v)


def _set_major(h, v=None, title=None):
    """대단원 제목 번호·이름 바꾸기. v=None 이면 번호 없이 제목만(상자식은 번호 칸을 빼고 제목 칸을 넓힌다)."""
    title = h["title"] if title is None else title
    el = h["el"]
    if h["box"]:
        tr = el.find(qn("w:tr"))
        c0, c1 = tr.findall(qn("w:tc"))
        if v is None:
            w0 = c0.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
            w1 = c1.find(f"{qn('w:tcPr')}/{qn('w:tcW')}")
            try:
                w1.set(qn("w:w"), str(int(float(w1.get(qn("w:w")))) + int(float(w0.get(qn("w:w"))))))
            except (AttributeError, TypeError, ValueError):
                pass
            grid = el.find(qn("w:tblGrid"))
            cols = grid.findall(qn("w:gridCol")) if grid is not None else []
            if len(cols) >= 2:
                cols[1].set(qn("w:w"), str(int(float(cols[0].get(qn("w:w")) or 0)) + int(float(cols[1].get(qn("w:w")) or 0))))
                grid.remove(cols[0])
            tr.remove(c0)
        else:
            _set_cell_text(c0, _fmt_num(h["kind"], h["tok"], v))
        p = next((pp for pp in c1.iter(qn("w:p")) if _vis_text(pp).strip()), None)
        if p is not None:
            _set_visible_text(p, title)
        return
    m = h["m"]
    head = "" if v is None else m.group(1) + _fmt_num(h["kind"], m.group(2), v) + m.group(3)
    _set_heading_text(el, m.end(3), head, title)


def _sub_headings_between(doc, start_el, end_el, src):
    """두 대단원 사이 소제목 문단(원본 소제목과 같은 번호 모양, 글자 크기 85% 이상). [(p, kind, m)]"""
    if src is None:
        return []
    p0, kind0, m0 = src
    sz0 = _max_sz(p0)
    out = []
    el = start_el.getnext()
    while el is not None and el is not end_el:
        if el.tag == qn("w:p"):
            hm = _head_match(_vis_text(el))
            if hm and hm[0] == kind0 and _max_sz(el) >= sz0 * 0.85:
                out.append((el, hm[0], hm[1]))
        el = el.getnext()
    return out


def _diagram_par(doc, diagram_png):
    from docx.shared import Cm
    import pdf_input as _pi
    w = min(Cm(16.0), int(_pi.BOX_W_PT * 12700))
    inline = doc.part.new_pic_inline(io.BytesIO(diagram_png), w, None)
    p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(p, qn("w:pPr"))
    etree.SubElement(ppr, qn("w:spacing"), {qn("w:before"): "120", qn("w:after"): "120"})
    etree.SubElement(ppr, qn("w:jc"), {qn("w:val"): "center"})
    r = etree.SubElement(p, qn("w:r"))
    etree.SubElement(r, qn("w:drawing")).append(inline)
    return p


def restructure_sections(doc, bond_els, bond_head, sub, diagram_png, has_highlight):
    """대단원 순서·번호를 회사 순서로(2026-10-14 사용자 확정).
      [원본 하이라이트 대단원 — 번호 없이 글자만]
      Ⅰ. 사모사채 개요  > 1. 본건 사모사채 개요(+표)              ← bond_head ~ bond_last (이미 넣어 둠)
      Ⅱ. 금융 개요      > 1. 금융 구조도(그림) > 2. 본건 기초자산 금융조건(원본 금융조건) > 3. …(원본 다음 소제목들)
      Ⅲ. … 원본 나머지 대단원 차례대로 다시 번호.
    번호 모양·글꼴·크기는 원본 대단원·소제목 것. 보고 dict."""
    rep = {}
    heads = _major_headings(doc)
    if not heads:                                           # 대단원을 못 찾으면 개요 표 뒤에 '2. 금융 구조도' + 그림
        last = bond_last
        s2 = _make_sub_heading(sub, 2, "금융 구조도")
        last.addnext(s2)
        if diagram_png:
            s2.addnext(_diagram_par(doc, diagram_png))
        rep["대단원"] = "못 찾음"
        return rep
    body = doc.element.body

    def pos(el):
        return body.index(el) if el.getparent() is body else -1
    bond_last = bond_els[-1]
    bpos = pos(bond_head)
    hl = None
    if has_highlight:
        import pdf_input as _pi
        # 하이라이트 제목 = 개요 앞에서 마지막으로 나온 '하이라이트 말(Executive Summary 등)' 짧은 줄(번호 있든 없든)
        anchor = None
        for el in body:
            if pos(el) >= bpos:
                break
            t = _vis_text(el).strip()
            if t and len(t) <= 40 and _pi.HIGHLIGHT_RE.search(t):
                anchor = el
        if anchor is not None:
            apos = pos(anchor)
            between = [h for h in heads if apos < pos(h["el"]) < bpos]
            if between:
                # ★하이라이트와 다음 대단원이 한 쪽에 같이 있다(대전 '1 Executive Summary'+'2 금융개요',
                #   등촌 '■ Executive Summary'+'1. 주요금융조건') → 개요를 그 쪽 뒤가 아니라 다음 대단원 바로 앞에(2026-10-15)
                target = between[0]["el"]
                left = bond_els[-1].getnext()                  # 개요 뒤 구역 나눔(새 쪽)만 남으면 빈 쪽이 된다 → 같이 뺀다
                if left is not None and left.tag == qn("w:p") and not _vis_text(left).strip() and                         left.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None and                         left.find(".//" + qn("w:drawing")) is None:
                    left.getparent().remove(left)
                for el in bond_els:
                    if el.getparent() is not None:
                        target.addprevious(el)
                rep["개요 옮김"] = f"'{between[0]['title']}' 앞으로"
                bpos = pos(bond_head)
            hl = next((h for h in heads if h["el"] is anchor), None)   # 번호 있는 대단원일 때만 번호를 뺀다
    # ★대단원 자체가 '금융구조도' 인 것(효자 'II. 금융구조도' + 그림) → 제목·그림을 통째로 뺀다(사용자 4번)
    for h in list(heads):
        if pos(h["el"]) > bpos and _STRUCT_RE.match(h["title"]) and "구조" in h["title"]:
            nxt_h = next((x["el"] for x in heads if pos(x["el"]) > pos(h["el"])), None)
            el, kill = h["el"].getnext(), [h["el"]]
            ok = True
            while el is not None and el is not nxt_h:
                if el.tag == qn("w:tbl") or _vis_text(el).strip():
                    ok = False                                 # 그림 말고 다른 내용이 있으면 제목만 지운다
                    break
                kill.append(el)
                el = el.getnext()
            for x in (kill if ok else [h["el"]]):
                if x.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
                    for ch in [c for c in x if c.tag != qn("w:pPr")]:
                        x.remove(ch)
                else:
                    x.getparent().remove(x)
            rep["원본 구조도 대단원 뺌"] = h["title"]
            heads = _major_headings(doc)
            break
    # 금융 대단원 = 개요 뒤 첫 금융 제목
    fin = next((h for h in heads if pos(h["el"]) > bpos and _FIN_TITLE_RE.search(h["title"])), None)
    style = next((h for h in heads if h["el"] is bond_head), None) or heads[0]
    fin_els = []
    if fin is None:                                         # 원본에 금융 대단원이 없다 → 개요 바로 뒤에 새로
        new = copy.deepcopy(style["el"])
        for sp in new.iter(qn("w:sectPr")):
            sp.getparent().remove(sp)
        bond_last.addnext(new)
        nh = dict(style, el=new)
        _set_major(nh, 1, "금융 개요")
        heads = _major_headings(doc)
        fin = next(h for h in heads if h["el"] is new)
        rep["금융 대단원"] = "원본에 없어 새로"
    else:
        rep["금융 대단원"] = f"원본 '{fin['title']}' → 금융 개요"
        # ★개요와 금융 대단원 사이에 다른 대단원이 있으면(안산: 사업 개요) 금융 대단원 묶음을 개요 바로 밑으로 옮긴다(사용자 3번)
        if any(pos(bond_last) < pos(h["el"]) < pos(fin["el"]) for h in heads):
            nxt_h = next((h["el"] for h in heads if pos(h["el"]) > pos(fin["el"])), None)
            block, el = [], fin["el"]
            while el is not None and el is not nxt_h and el.tag != qn("w:sectPr"):
                block.append(el)
                el = el.getnext()
            anchor = bond_last
            for el in block:
                if el.tag == qn("w:p") and el.find(f"{qn('w:pPr')}/{qn('w:sectPr')}") is not None:
                    # 구역 끝 문단은 제자리에 두고(쪽 나눔 유지) 내용만 옮긴다
                    moved = copy.deepcopy(el)
                    moved.find(qn("w:pPr")).remove(moved.find(f"{qn('w:pPr')}/{qn('w:sectPr')}"))
                    for ch in [c for c in el if c.tag != qn("w:pPr")]:
                        el.remove(ch)
                    el = moved
                anchor.addnext(el)
                anchor = el
            rep["금융 대단원 옮김"] = len(block)
        # ★개요 표 바로 밑에 이어서 '금융 개요'(새 쪽 X, 2026-10-15 사용자 3번) — 사이의 빈 문단·구역 나눔을 걷어낸다
        between, el, ok = [], bond_last.getnext(), True
        while el is not None and el is not fin["el"]:
            if el.tag != qn("w:p") or _vis_text(el).strip() or el.find(".//" + qn("w:drawing")) is not None:
                ok = False
                break
            between.append(el)
            el = el.getnext()
        if ok and el is fin["el"]:
            for b in between:
                b.getparent().remove(b)
            for br in list(fin["el"].iter(qn("w:br"))):          # 제목 앞 쪽 나눔도
                if br.get(qn("w:type")) == "page":
                    br.getparent().remove(br)
            ppr = fin["el"].find(qn("w:pPr")) if fin["el"].tag == qn("w:p") else None
            if ppr is not None:
                for e in ppr.findall(qn("w:pageBreakBefore")):
                    ppr.remove(e)
            gap = etree.Element(qn("w:p"))                     # 위 표와 조금 띄움
            etree.SubElement(etree.SubElement(gap, qn("w:pPr")), qn("w:spacing"),
                             {qn("w:before"): "240", qn("w:after"): "0"})
            fin["el"].addprevious(gap)
            rep["금융 개요 이어서"] = len(between)
    fin["title"] = "금융 개요"
    # 금융 대단원 안 소제목: 1. 금융 구조도, 2. 본건 기초자산 금융조건, 3.~ 원본 나머지
    nxt = next((h["el"] for h in heads if pos(h["el"]) > pos(fin["el"])), None)
    subs = _sub_headings_between(doc, fin["el"], nxt, sub)
    anchor = fin["el"]
    first_fin_sub = subs[0] if subs and _FIN_TITLE_RE.search(subs[0][2].group(4)) else None
    s1 = _make_sub_heading(sub, 1, "금융 구조도")
    anchor.addnext(s1)
    fin_els.append(s1)
    last = s1
    if diagram_png:
        dp = _diagram_par(doc, diagram_png)
        last.addnext(dp)
        last = dp
    if first_fin_sub is None:                               # 원본 금융조건 소제목이 없으면 새로(원본 표 바로 위)
        s2 = _make_sub_heading(sub, 2, FIN_TITLE)
        last.addnext(s2)
        rest = subs
    else:
        p, kind, m = first_fin_sub
        _set_heading_text(p, m.end(3), m.group(1) + _fmt_num(kind, m.group(2), 2) + m.group(3), FIN_TITLE)
        ind = p.find(f"{qn('w:pPr')}/{qn('w:ind')}")
        if ind is not None:
            ind.set(qn("w:right"), "0")
        rest = subs[1:]
    for k, (p, kind, m) in enumerate(rest):
        _set_heading_text(p, m.end(3), m.group(1) + _fmt_num(kind, m.group(2), k + 3) + m.group(3), m.group(4).strip())
    rep["금융 소제목"] = 2 + len(rest)
    # 대단원 번호 다시 매김
    v = 0
    for h in _major_headings(doc):
        if hl is not None and h["el"] is hl["el"]:
            _set_major(h, None)
            continue
        v += 1
        title = "금융 개요" if h["el"] is fin["el"] else ("사모사채 개요" if h["el"] is bond_head else None)
        _set_major(h, v, title)
    rep["대단원"] = v
    rep["하이라이트 번호 뺌"] = hl["title"] if hl else "없음"
    return rep



def build(src_bytes, *, title_lines, header_text, date_text, photo, bond, bond_place=None, tpl_path=None,
          diagram_png=None):
    """(PDF 에서 만든) 본문 워드 bytes → 회사양식 워드 bytes, 보고(dict).

    본문은 pdf_input.convert 가 이미 회사 규격(A4·여백)으로 맞춰 둔 것이다.
    bond_place : _place_bond 참고.
    """
    tpl_path = tpl_path or template_path()
    tpl = docx.Document(tpl_path)
    doc = docx.Document(io.BytesIO(src_bytes))
    rep = {}

    apply_fonts(doc)
    rep["제목 줄 나눔"] = split_heading_breaks(doc)
    rep["머리글 칸"] = paint_headers(doc)
    rep["세로 가운데 칸"] = center_cells_vertically(doc)
    rep["대단원 제목 회사 모양"] = restyle_section_headings(doc) + restyle_shaded_number_headings(doc)
    rep["칸 끝 빈칸 지움"] = strip_cell_trailing_spaces(doc)

    body = doc.element.body
    # 원본 구역들: 쪽번호를 처음부터 다시 세는 설정 지우기(표지부터 이어 세게)
    for sp in body.iter(qn("w:sectPr")):
        pn = sp.find(qn("w:pgNumType"))
        if pn is not None:
            pn.attrib.pop(qn("w:start"), None)

    # ── 사모사채 개요: 하이라이트 뒤 ──
    head = _bond_heading_from_original(doc)
    bond_els = _build_bond_section(doc, bond)
    for el in bond_els:
        body.remove(el)
    rep["흰 테두리 검정"] = whiten_borders_to_black(doc)
    rep["원본 구조도 줄 지움"] = drop_structure_rows(doc)
    rep["원본 구조도 그림 지움"] = drop_structure_paragraphs(doc)
    rep["단위 오른쪽 맞춤"] = right_align_units(doc)
    sub = _sub_heading_from_original(doc)
    tpl_bond = _template_bond_els(tpl, doc, bond, None, sub)   # 양식(최종) 사모사채 개요 표(구조도는 금융 개요 아래로)
    # ★2026-10-14 사용자 확정: 대단원 'Ⅰ. 사모사채 개요'(원본 대단원 모양) > 소제목 '1. 본건 사모사채 개요'(원본 소제목 모양) > 표
    if tpl_bond:
        bond_els = [_make_sub_heading(sub, 1, "본건 사모사채 개요")] + tpl_bond
        rep["개요 제목"] = "원본 소제목 모양" if sub is not None else "기본 모양(원본 소제목 못 찾음)"
    else:                                                   # 양식에서 개요 표를 못 찾았을 때만 옛 방식
        _set_par_text(bond_els[0], "본건 사모사채 개요")
        rep["개요 제목"] = "양식 표 못 찾음"
    if head is not None:
        bond_els = [head] + bond_els
    if bond_place is not None and bond_place[1]:            # 같은 쪽에 이을 때는 위와 조금 띄운다
        gap = etree.Element(qn("w:p"))
        etree.SubElement(etree.SubElement(gap, qn("w:pPr")), qn("w:spacing"),
                         {qn("w:before"): "240", qn("w:after"): "0"})
        bond_els = [gap] + bond_els
    _place_bond(doc, tpl, bond_els, bond_place)
    bond_head = head if head is not None else bond_els[-len(tpl_bond or [None]) - 1]
    rep.update(restructure_sections(doc, bond_els, bond_head, sub, diagram_png, bond_place is not None))
    rep["개요 자리"] = ("표지 다음 쪽" if bond_place is None else
                      f"하이라이트 뒤 {'같은 쪽에 이어서' if bond_place[1] else '새 쪽'}")

    # ── 맨 뒤: 연락처 쪽(세로 A4 구역) ──
    old_last = body.find(qn("w:sectPr"))
    holder = etree.SubElement(body, qn("w:p"))
    body.remove(holder)
    ppr = etree.SubElement(holder, qn("w:pPr"))
    ppr.append(copy.deepcopy(old_last))
    body.insert(list(body).index(old_last), holder)          # 원본 마지막 구역을 문단 구역나눔으로
    body.replace(old_last, _sect_from_template(tpl, title_page=False))
    tail = _template_contact_els(tpl, doc)
    if tail:
        for el in tail:
            body.insert(list(body).index(body.find(qn("w:sectPr"))), el)
    else:
        contacts = read_contacts(tpl)
        for el in _build_contact_section(doc, contacts):
            body.remove(el)
            body.insert(list(body).index(body.find(qn("w:sectPr"))), el)

    # ── 맨 앞: 회사 표지(혼자 한 쪽) ──
    _merge_nsmap(doc, tpl)
    body = doc.element.body
    cover = _cover_elements(tpl, doc, title_lines, date_text, photo)
    # 양식 표지 두 번째 문단 끝의 '쪽 나눔' 은 뺀다 — 구역 끝 문단이 새 쪽을 연다(빈 쪽 방지)
    for br in list(cover[-1].iter(qn("w:br"))):
        if br.get(qn("w:type")) == "page":
            br.getparent().remove(br)
    # 회사 Investment Highlight 쪽(양식 그대로) — 표지 바로 다음, 그 뒤에 사모사채 개요(새 쪽)
    hl = []                                                 # 2026-10-13 양식에서 회사 하이라이트 쪽 뺌 — 원본 하이라이트를 쓴다
    for i, el in enumerate(cover + [_holder(_sect_from_template(tpl, title_page=True))] + hl):
        body.insert(i, el)

    # ── 머리말·꼬리말: 첫 구역에만 넣고 나머지는 '앞과 같게' ──
    tpl_sec = tpl.sections[0]
    secs = doc.sections
    for k, s in enumerate(secs):
        s.different_first_page_header_footer = (k == 0)
        for hf in (s.header, s.footer, s.even_page_header, s.even_page_footer):
            hf.is_linked_to_previous = True
        if k > 0:
            s.first_page_header.is_linked_to_previous = True
            s.first_page_footer.is_linked_to_previous = True
    s0 = secs[0]
    s0.header.is_linked_to_previous = False
    s0.footer.is_linked_to_previous = False
    _copy_hdrftr(tpl, doc, tpl_sec.header.part, s0.header, header_text)
    _copy_hdrftr(tpl, doc, tpl_sec.footer.part, s0.footer)
    s0.first_page_header.is_linked_to_previous = False      # 표지는 머리말·꼬리말 없음
    s0.first_page_footer.is_linked_to_previous = False
    for hf in (s0.first_page_header, s0.first_page_footer):
        for p in hf.paragraphs[1:]:
            p._p.getparent().remove(p._p)
        if hf.paragraphs:
            for r in list(hf.paragraphs[0].runs):
                r._r.getparent().remove(r._r)
    # 홀짝 머리말 설정 끄기
    st = doc.settings.element
    eo = st.find(qn("w:evenAndOddHeaders"))
    if eo is not None:
        st.remove(eo)

    apply_company_geometry(doc)
    rep["표 머리말 줄에 맞춤"] = align_tables_to_body(doc)
    _renumber_shapes(doc)

    out = io.BytesIO()
    doc.save(out)
    rep["쪽 구역 수"] = len(doc.sections)
    return out.getvalue(), rep


def default_date_text(today=None):
    d = today or date.today()
    return d.strftime("%B %Y")
