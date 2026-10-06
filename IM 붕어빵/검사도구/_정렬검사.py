# -*- coding: utf-8 -*-
"""표 칸 안 글자가 **원본 자리에 놓이는지** 숫자로 잰다.

내가 고른 정렬(왼쪽/가운데/오른쪽)대로 놓았을 때의 자리와,
원본 PDF에서 실제로 그 글자가 있던 자리의 차이를 잰다.
사용: python _정렬검사.py [쪽번호 ...]
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
MARGIN = 1.5          # _put_table 이 주는 칸 좌우 여백
BAD = 3.0             # 이보다 어긋나면 눈에 보인다


def _pad_of(align, boxes, x0, x1):
    """_put_table 과 같은 방식으로 칸 여백을 구한다."""
    wide = max(b.width for b in boxes)
    room = max(0.0, (x1 - x0) - wide - 1.0)
    if align == "왼쪽":
        return min(max(0.0, min(b.x0 for b in boxes) - x0), room)
    if align == "오른쪽":
        return min(max(0.0, x1 - max(b.x1 for b in boxes)), room)
    return 0.0


def _place_err(align, box, x0, x1, pad=0.0, extra=0.0):
    """그 정렬·여백으로 놓았을 때 원본 자리와 얼마나 어긋나나."""
    if align == "가운데":
        return abs((x0 + x1) / 2 - (box.x0 + box.x1) / 2)
    if align == "왼쪽":
        return abs((x0 + pad + extra) - box.x0)
    return abs((x1 - pad - extra) - box.x1)


_NAME = {P.PP_ALIGN.LEFT: "왼쪽", P.PP_ALIGN.CENTER: "가운데",
         P.PP_ALIGN.RIGHT: "오른쪽"}


def check(pno, pg):
    _rects, plan, _boxes = P._plan(pg, P._lines_of(pg))
    txt = P._lines_of(pg)
    used = {}
    bad = []
    n_cell = 0
    for t in plan:
        xs, ys, cells = t["grid"]
        skip = t["skip"]
        n_col, n_row = len(xs) - 1, len(ys) - 1
        for (r0, c0, rs, cs) in cells:
            x0, y0 = xs[c0], ys[r0]
            x1, y1 = xs[min(c0 + cs, n_col)], ys[min(r0 + rs, n_row)]
            picked = []
            for i, ln in enumerate(txt):
                if ln.get("unit"):
                    continue
                b = ln["bbox"]
                if b.y1 < y0 - 1 or b.y0 > y1 + 1 or b.x1 < x0 - 1 or b.x0 > x1 + 1:
                    continue
                done = used.setdefault(i, set())
                idx = [k for k, c in enumerate(ln["chars"])
                       if k not in done
                       and x0 - 0.5 <= (c["bbox"].x0 + c["bbox"].x1) / 2 <= x1 + 0.5
                       and y0 - 1 <= (c["bbox"].y0 + c["bbox"].y1) / 2 <= y1 + 1
                       and not any(P._inside(c["bbox"], s) for s in skip)]
                if idx and "".join(ln["chars"][k]["c"] for k in idx).strip():
                    picked.append((i, idx))
            if not picked:
                continue
            seg = []
            for (i, idx) in picked:
                chars = [txt[i]["chars"][k] for k in idx]
                used[i].update(idx)
                seg.append({"chars": chars, "bbox": P._bbox_of(chars)})
            seg.sort(key=lambda s: (s["bbox"].y0, s["bbox"].x0))
            rows = []
            for s in seg:
                b = s["bbox"]
                near = rows and (min(rows[-1]["bbox"].y1, b.y1)
                                 - max(rows[-1]["bbox"].y0, b.y0)) > 0
                if near:
                    rows[-1]["items"].append(s)
                    rows[-1]["bbox"] = rows[-1]["bbox"] | b
                else:
                    rows.append({"items": [s], "bbox": b})
            boxes = [r["bbox"] for r in rows]
            align = _NAME[P._cell_align(boxes, x0, x1)]
            n_cell += 1

            def _errs(a):
                p = _pad_of(a, boxes, x0, x1)
                out = []
                for b in boxes:
                    if a == "왼쪽":
                        ex = max(0.0, (b.x0 - x0) - p)
                    elif a == "오른쪽":
                        ex = max(0.0, (x1 - b.x1) - p)
                    else:
                        ex = 0.0
                    out.append(_place_err(a, b, x0, x1, p, ex))
                return out

            errs = _errs(align)
            best = min(("왼쪽", "가운데", "오른쪽"), key=lambda a: max(_errs(a)))
            berr = max(_errs(best))
            if max(errs) > BAD:
                bad.append((max(errs), align, best, berr,
                            "".join(c["c"] for c in rows[0]["items"][0]["chars"])[:20],
                            round(x0, 1), round(y0, 1)))
    bad.sort(reverse=True)
    print(f"{pno:>3}쪽 : 글자 있는 칸 {n_cell}개 · 어긋난 칸 {len(bad)}개")
    for err, align, best, berr, t, x, y in bad[:6]:
        print(f"       {err:>5.1f}pt 어긋남  내가고른={align:<4} "
              f"더나은={best:<4}({berr:.1f}pt)  x={x} y={y}  {t!r}")
    return len(bad), n_cell


if __name__ == "__main__":
    doc = fitz.open(_기본)
    want = [int(a) for a in sys.argv[1:]] or list(range(1, doc.page_count + 1))
    tb = tc = 0
    for p in want:
        a, b = check(p, doc[p - 1])
        tb += a
        tc += b
    print(f"\n합계 : 어긋난 칸 {tb} / {tc}개")
    doc.close()
