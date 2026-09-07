# -*- coding: utf-8 -*-
"""원본에서 읽은 표(pdf_tables)를 PPT 표로 **그대로** 그린다.

지금까지는 원본을 LLM 이 다시 쓴 값을 회사 틀에 박아 넣었다. 그래서
  · 틀에 박혀 있던 라벨이 그대로 남고('기조자산' 같은 오타)
  · 칸마다 서식이 달라 '·' 이 있다 없다 하고, 간격·시작점이 제각각이고
  · 두 줄이 되면 둘째 줄이 첫 줄과 안 맞았다
표를 원본 그대로 새로 그리면 이 문제가 전부 사라진다.

규칙(사용자 지시):
  · 요약하지 않는다 — 원본 행·열·병합·글자 그대로
  · 글씨는 9pt 고정. 넘쳐도 그대로 둔다
  · 불릿(·) 없음. 모든 칸의 들여쓰기·줄맞춤을 같게
"""
from pptx.util import Emu, Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn

FONT_PT = 9.0                       # ★9pt 고정(사용자 지시)
FONT_NAME = "피플폰트 Light"
HEAD_FONT = "피플폰트 Bold"
LINE = RGBColor(0xA5, 0xA5, 0xA5)   # 회사 표 테두리색
HEAD_BG = RGBColor(0xF2, 0xF4, 0xF7)
PAD = Inches(0.04)


def _no_bullet(para):
    """불릿을 확실히 끈다. 틀에서 물려받은 '·' 이 칸마다 달랐던 원인."""
    pPr = para._p.get_or_add_pPr()
    for tag in ("a:buChar", "a:buAutoNum", "a:buBlip"):
        for e in pPr.findall(qn(tag)):
            pPr.remove(e)
    if pPr.find(qn("a:buNone")) is None:
        pPr.append(pPr.makeelement(qn("a:buNone"), {}))
    pPr.set("marL", "0")            # 들여쓰기 0 — 둘째 줄도 첫 줄과 같은 자리에서 시작
    pPr.set("indent", "0")


def _cell_style(cell, text, head=False):
    cell.margin_left = cell.margin_right = PAD
    cell.margin_top = cell.margin_bottom = Inches(0.02)
    cell.vertical_anchor = MSO_ANCHOR.MIDDLE
    tf = cell.text_frame
    tf.word_wrap = True
    tf.text = text or ""
    for para in tf.paragraphs:
        _no_bullet(para)
        para.alignment = PP_ALIGN.CENTER if head else PP_ALIGN.LEFT
        if not para.runs:
            para.add_run()
        for run in para.runs:
            run.font.size = Pt(FONT_PT)
            run.font.name = HEAD_FONT if head else FONT_NAME
            run.font.bold = bool(head)


def _borders(cell):
    """회사 표와 **똑같은** 선 서식.

    회사 틀을 뜯어보니 좌우 세로선은 아예 없고 위아래 가로선만 회색으로 긋는다.
    내가 사방에 테두리를 그리는 바람에 다른 표들과 완전히 달라 보였다.
    """
    tcPr = cell._tc.get_or_add_tcPr()
    tcPr.set("marL", "33513")
    tcPr.set("marR", "33513")
    tcPr.set("marT", "0")
    tcPr.set("marB", "0")
    tcPr.set("anchor", "ctr")
    for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
        for e in tcPr.findall(qn(tag)):
            tcPr.remove(e)
    for tag in ("a:lnL", "a:lnR"):                   # 세로선 없음
        ln = tcPr.makeelement(qn(tag), {"w": "12700", "cmpd": "sng"})
        ln.append(ln.makeelement(qn("a:noFill"), {}))
        tcPr.append(ln)
    for tag in ("a:lnT", "a:lnB"):                   # 가로선만 회색
        ln = tcPr.makeelement(qn(tag), {"w": "6350", "cap": "flat",
                                        "cmpd": "sng", "algn": "ctr"})
        fill = ln.makeelement(qn("a:solidFill"), {})
        clr = fill.makeelement(qn("a:schemeClr"), {"val": "bg1"})
        clr.append(clr.makeelement(qn("a:lumMod"), {"val": "65000"}))
        fill.append(clr)
        ln.append(fill)
        tcPr.append(ln)


def _first_row_is_head(t):
    """첫 줄이 머리글인지 — 숫자가 거의 없고 짧으면 머리글로 본다."""
    first = [c for c in t["cells"] if c["r"] == 0]
    if not first:
        return False
    txt = " ".join(c["text"] for c in first)
    digits = sum(ch.isdigit() for ch in txt)
    return len(txt) > 0 and digits <= max(2, len(txt) * 0.1)


def add_table(slide, t, left, top, width):
    """원본 표 t 를 슬라이드에 그린다. 반환: 만든 표 도형."""
    n_r, n_c = t["rows"], t["cols"]
    row_h = Inches(0.22)
    shape = slide.shapes.add_table(n_r, n_c, left, top, width,
                                   Emu(int(row_h * n_r)))
    shape.name = "원본표"
    tbl = shape.table
    tbl.first_row = False            # 파워포인트 기본 줄무늬·머리글 서식 끄기
    tbl.horz_banding = False

    # 열 너비 — 원본 표의 열 폭 비율을 그대로 따른다
    bbox = t.get("bbox")
    widths = t.get("col_ratio")
    if widths and len(widths) == n_c:
        total = sum(widths) or 1.0
        acc = 0
        for i, col in enumerate(tbl.columns):
            w = int(width * widths[i] / total)
            col.width = w
            acc += w
        tbl.columns[n_c - 1].width += int(width) - acc   # 음수 인덱스 안 됨
    else:
        for col in tbl.columns:
            col.width = int(width / n_c)

    head = _first_row_is_head(t)
    for row in tbl.rows:
        row.height = row_h

    filled = set()
    for c in sorted(t["cells"], key=lambda x: (x["r"], x["c"])):
        r0, c0, rs, cs = c["r"], c["c"], c["rs"], c["cs"]
        if (r0, c0) in filled:
            continue
        cell = tbl.cell(r0, c0)
        if rs > 1 or cs > 1:
            try:
                cell.merge(tbl.cell(min(r0 + rs, n_r) - 1, min(c0 + cs, n_c) - 1))
            except Exception:
                pass
        for rr in range(r0, min(r0 + rs, n_r)):
            for cc in range(c0, min(c0 + cs, n_c)):
                filled.add((rr, cc))
        # 회사 표는 첫 '줄' 이 아니라 **첫 열(라벨)** 을 굵게 쓴다.
        # 배경색은 내가 임의로 넣지 않는다(다른 표들과 달라 보였다).
        _cell_style(cell, c["text"], head=(c0 == 0 or (head and r0 == 0)))

    for r in range(n_r):
        for c in range(n_c):
            try:
                _borders(tbl.cell(r, c))
            except Exception:
                pass
    return shape
