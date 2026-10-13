# -*- coding: utf-8 -*-
"""금융 구조도 그림(PNG) — 기존 IM 도구(modules/finance_diagram.py)의 배치·항목을 그대로 옮겨 그림으로 그린다.

기존 IM 은 파워포인트 도형으로 그렸지만, 이 도구는 워드에 넣고 웹(파워포인트 없음)에서도 돌아야 하므로
같은 자리·같은 상자·같은 화살표를 PIL 로 그린다(2026-10-13 사용자 확정: '그 단계를 똑같이').
색은 회사 네이비(17365D). 글꼴은 양식\\fonts 의 Pretendard.

fin = {"차주", "신탁", "본건": {"금액", "LTV", "금리", "담보"}, "브릿지", "연대보증"} — 빈 값은 기존 IM 과 같은 기본 글.
"""
import io
import math
import os

from PIL import Image, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
FONT_B = os.path.join(HERE, "양식", "fonts", "Pretendard-Bold.otf")

NAVY = (0x17, 0x36, 0x5D)
BLUE = (0x3E, 0x95, 0xBE)
RED = (0xC0, 0x00, 0x00)
PINK = (0xFD, 0xE7, 0xE7)
LGRAY = (0xF2, 0xF2, 0xF2)
WHITE = (0xFF, 0xFF, 0xFF)
DARK = (0x33, 0x33, 0x33)

DPI = 200                       # 1인치 = 200점
TXT = 1.55                      # 글자 배율 — 워드 본문 폭(16cm)에 넣으면 그림이 62% 로 줄어 글이 작아진다(9pt → 약 8.6pt)
X0, Y0 = 0.35, 1.65             # 기존 IM 슬라이드 좌표(인치)에서 잘라 낼 왼쪽 위
WIDTH_IN, HEIGHT_IN = 10.2, 4.75


def _font(pt):
    try:
        return ImageFont.truetype(FONT_B, int(pt * TXT * DPI / 72))
    except Exception:
        return ImageFont.load_default()


def _xy(x, y):
    return int((x - X0) * DPI), int((y - Y0) * DPI)


def _box(d, x, y, w, h, text, fill=WHITE, line=NAVY, txt=DARK, size=9):
    x0, y0 = _xy(x, y)
    x1, y1 = _xy(x + w, y + h)
    d.rounded_rectangle([x0, y0, x1, y1], radius=int(0.08 * DPI), fill=fill, outline=line, width=max(2, int(1.25 * DPI / 72)))
    f = _font(size)
    lines = [ln for ln in str(text).split("\n")]
    lh = int(size * TXT * 1.3 * DPI / 72)
    total = lh * len(lines)
    cy = (y0 + y1) / 2 - total / 2
    for i, ln in enumerate(lines):
        tw = d.textlength(ln, font=f)
        while tw > (x1 - x0) - 10 and len(ln) > 2:          # 상자보다 길면 글자 줄이기 대신 줄 끝 자르지 않고 글자 크기 낮춤
            f = _font(size * 0.9)
            size *= 0.9
            tw = d.textlength(ln, font=f)
        d.text(((x0 + x1) / 2 - tw / 2, cy + i * lh), ln, font=f, fill=txt)


def _arrow(d, x1, y1, x2, y2, label, lx, ly, color=NAVY, dash=False):
    a, b = _xy(x1, y1), _xy(x2, y2)
    w = max(2, int(1.5 * DPI / 72))
    if dash:
        L = math.dist(a, b)
        n = max(1, int(L // 18))
        for k in range(0, n, 2):
            p = (a[0] + (b[0] - a[0]) * k / n, a[1] + (b[1] - a[1]) * k / n)
            q = (a[0] + (b[0] - a[0]) * min(k + 1, n) / n, a[1] + (b[1] - a[1]) * min(k + 1, n) / n)
            d.line([p, q], fill=color, width=w)
    else:
        d.line([a, b], fill=color, width=w)
    ang = math.atan2(b[1] - a[1], b[0] - a[0])
    hl, hw = 0.13 * DPI, 0.07 * DPI
    tip = b
    left = (tip[0] - hl * math.cos(ang) + hw * math.sin(ang), tip[1] - hl * math.sin(ang) - hw * math.cos(ang))
    right = (tip[0] - hl * math.cos(ang) - hw * math.sin(ang), tip[1] - hl * math.sin(ang) + hw * math.cos(ang))
    d.polygon([tip, left, right], fill=color)
    f = _font(9)
    lh = int(9 * TXT * 1.3 * DPI / 72)
    px, py = _xy(lx, ly)
    for i, s in enumerate(str(label).split("\n")):
        d.text((px, py + i * lh), s, font=f, fill=color)


def render_png(fin=None):
    """기존 IM build_finance_diagram_slide 와 같은 배치의 금융 구조도 PNG bytes."""
    fin = fin or {}
    bg = fin.get("본건") or {}
    im = Image.new("RGB", (int(WIDTH_IN * DPI), int(HEIGHT_IN * DPI)), WHITE)
    d = ImageDraw.Draw(im)

    차주 = fin.get("차주") or "차주"
    신탁 = fin.get("신탁") or "부동산담보신탁\n(신탁사)"
    연대 = fin.get("연대보증") or ""
    브릿지 = fin.get("브릿지") or ""
    담보 = bg.get("담보") or ""
    본건줄 = "본건 담보대출  " + (bg.get("금액") or "")
    조건줄 = " · ".join([x for x in [("LTV " + bg["LTV"]) if bg.get("LTV") else "",
                                     ("금리 " + bg["금리"]) if bg.get("금리") else ""] if x])

    _box(d, 0.55, 1.85, 1.9, 0.6, str(신탁), fill=LGRAY)
    _box(d, 0.55, 3.4, 1.9, 0.75, "차주\n" + str(차주), fill=WHITE)
    if 연대:
        _box(d, 0.55, 5.45, 1.9, 0.55, str(연대), fill=LGRAY)
    _box(d, 3.7, 3.3, 2.7, 0.95,
         "\n".join([x for x in [본건줄, ("(" + 조건줄 + ")") if 조건줄 else "", 담보] if x]),
         fill=PINK, line=RED, txt=RED)
    _box(d, 8.0, 2.6, 2.3, 0.7, "신규 유동화\nSPC (TBD)", fill=NAVY, txt=WHITE, size=10.5)
    _box(d, 8.0, 4.45, 2.3, 0.65, "투자자\n(사채권자)", fill=BLUE, line=BLUE, txt=WHITE, size=10.5)
    if 브릿지:
        _box(d, 3.7, 5.5, 3.2, 0.75, "브릿지 대출  " + str(브릿지) + "\n→ 본건 대출 상환재원", fill=LGRAY)

    _arrow(d, 1.5, 3.4, 1.5, 2.45, "담보신탁계약", 1.62, 2.75)
    _arrow(d, 2.45, 3.75, 3.7, 3.75, "담보대출", 2.55, 3.45)
    if 연대:
        _arrow(d, 1.5, 5.45, 1.5, 4.15, "연대보증", 1.62, 4.7)
    _arrow(d, 6.4, 3.6, 8.0, 3.0, "대출채권\n(기초자산)", 6.35, 2.8)
    _arrow(d, 9.15, 3.3, 9.15, 4.45, "사모사채\n발행·인수", 9.3, 3.65)
    if 브릿지:
        _arrow(d, 5.3, 5.5, 5.3, 4.25, "상환", 5.42, 4.75, dash=True)

    # 아래 빈 자리 잘라 내기(연대보증·브릿지가 없으면 아래쪽이 비어 있다)
    bottom = 6.35 if (연대 or 브릿지) else 5.3
    im = im.crop((0, 0, im.width, int((bottom - Y0) * DPI)))
    buf = io.BytesIO()
    im.save(buf, "PNG")
    return buf.getvalue()


if __name__ == "__main__":
    import sys
    open(sys.argv[1] if len(sys.argv) > 1 else "diagram.png", "wb").write(render_png({
        "차주": "주식회사 쿠도스", "신탁": "부동산담보신탁\n(TBD)",
        "본건": {"금액": "[45]억원", "LTV": "38.3%", "금리": "[T.B.D]%", "담보": "연립주택 8세대 1순위 담보신탁"},
        "연대보증": "대표이사·대주주 연대보증"}))
