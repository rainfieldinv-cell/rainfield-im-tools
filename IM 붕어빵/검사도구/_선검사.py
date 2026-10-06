# -*- coding: utf-8 -*-
"""원본 PDF에 그어진 선과, 만든 PPT의 표 테두리를 **좌표로 맞대본다.**

눈으로 찾으면 놓친다. 빠진 선이 어디 몇 개인지 숫자로 나오게 한다.
사용: python _선검사.py [쪽번호 ...]
"""
import os
import sys

# ★이 파일은 '검사도구' 폴더 안에 있다. 프로그램(pdf_to_ppt)과 자료(_원본.pdf)는
#   옆 폴더에 있으므로 찾아갈 길을 열어 준다.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)   # 2026-10-06 한 폴더로 합침: 엔진(pdf_to_ppt)이 바로 윗 폴더에 있다
_자료 = os.path.join(_ROOT, "자료")
_기본 = os.path.join(os.path.dirname(_ROOT), "IM(샘플)", "[IM]-돈암동.pdf")
import os
import sys

import fitz

import pdf_to_ppt as P

HERE = _자료
TOL = 2.5


def _pdf_segs(rects, skip=()):
    """원본에 그어진 선 → 세로 [(x, y0, y1)] · 가로 [(y, x0, x1)]

    밑줄로 옮긴 선은 뺀다 — 표 테두리로 안 넣은 게 맞고, 글자 밑줄 서식으로 들어간다.
    """
    v, h = [], []
    for n, it in enumerate(rects):
        if not it["thin"] or n in skip:
            continue
        r = it["r"]
        if r.width <= P.THIN_PT and r.height > P.THIN_PT:
            v.append(((r.x0 + r.x1) / 2, r.y0, r.y1))
        elif r.height <= P.THIN_PT and r.width > P.THIN_PT:
            h.append(((r.y0 + r.y1) / 2, r.x0, r.x1))
    return v, h


def _emitted(xs, ys, cells, lines):
    """_put_table 과 **같은 판정**으로, 실제로 넣게 될 테두리를 모은다."""
    n_col, n_row = len(xs) - 1, len(ys) - 1
    v, h = [], []
    for (r0, c0, rs, cs) in cells:
        x0, y0 = xs[c0], ys[r0]
        x1, y1 = xs[min(c0 + cs, n_col)], ys[min(r0 + rs, n_row)]
        if lines.vline(x0, y0, y1):
            v.append((x0, y0, y1))
        if lines.vline(x1, y0, y1):
            v.append((x1, y0, y1))
        if lines.hline(y0, x0, x1):
            h.append((y0, x0, x1))
        if lines.hline(y1, x0, x1):
            h.append((y1, x0, x1))
    return v, h


def _uncovered(seg, pool):
    """seg(=(축, a, b)) 중 pool 이 덮지 못한 길이."""
    axis, a, b = seg
    covers = sorted((max(a, p), min(b, q)) for (ax, p, q) in pool
                    if abs(ax - axis) <= TOL and min(b, q) - max(a, p) > 0.5)
    left, cur = 0.0, a
    for (p, q) in covers:
        if p > cur:
            left += p - cur
        cur = max(cur, q)
    left += max(0.0, b - cur)
    return left


def check(pno, pg):
    txt = P._lines_of(pg)
    rects, plan, boxes = P._plan(pg, txt)
    pv, ph = _pdf_segs(rects, P._underlines(rects, txt))
    ev, eh = [], []
    for t in plan:
        xs, ys, cells = t["grid"]
        a, b = _emitted(xs, ys, cells, t["lines"])
        ev += a
        eh += b

    def inside(ax, a, b, vertical):
        for r in boxes:
            if vertical and r.x0 - 2 <= ax <= r.x1 + 2 and a >= r.y0 - 2 and b <= r.y1 + 2:
                return True
            if not vertical and r.y0 - 2 <= ax <= r.y1 + 2 and a >= r.x0 - 2 and b <= r.x1 + 2:
                return True
        return False

    miss = []
    for s in pv:
        if not inside(s[0], s[1], s[2], True):
            continue
        left = _uncovered(s, ev)
        if left > 1.0:
            miss.append(("세로", s, left))
    for s in ph:
        if not inside(s[0], s[1], s[2], False):
            continue
        left = _uncovered(s, eh)
        if left > 1.0:
            miss.append(("가로", s, left))
    total = sum(m[2] for m in miss)
    print(f"{pno:>3}쪽 : 표 {len(boxes)}개 · 원본 선 {len(pv)+len(ph)}개 · "
          f"못 넣은 선 {len(miss)}개 (길이 합 {total:.0f}pt)")
    for kind, (ax, a, b), left in sorted(miss, key=lambda m: -m[2])[:8]:
        pos = f"x={ax:.1f} y={a:.1f}~{b:.1f}" if kind == "세로" \
            else f"y={ax:.1f} x={a:.1f}~{b:.1f}"
        print(f"       {kind} {pos}  못 넣은 길이 {left:.1f}pt")
    return len(miss)


if __name__ == "__main__":
    doc = fitz.open(_기본)
    want = [int(a) for a in sys.argv[1:]] or list(range(1, doc.page_count + 1))
    bad = 0
    for p in want:
        bad += check(p, doc[p - 1])
    print(f"\n합계 : 못 넣은 선 {bad}개")
    doc.close()
