# -*- coding: utf-8 -*-
"""원본 IM 표에 **실제로 어떤 색이 칠해져 있는지** 세어 본다.

'진한 부분'을 남색으로 바꾸려면 먼저 진한 색이 뭔지 눈이 아니라 숫자로 알아야 한다.
같이 놓인 글자색까지 세어 준다(머리글은 대개 흰 글씨다).

사용: python _색검사.py
"""
import os
import sys

# ★이 파일은 '검사도구' 폴더 안에 있다. 프로그램(pdf_to_ppt)과 자료(_원본.pdf)는
#   옆 폴더에 있으므로 찾아갈 길을 열어 준다.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)   # 2026-10-06 한 폴더로 합침: 엔진(pdf_to_ppt)이 바로 윗 폴더에 있다
_자료 = os.path.join(_ROOT, "자료")
_기본 = os.path.join(os.path.dirname(_ROOT), "IM(샘플)", "[IM]-돈암동.pdf")
import glob
import os
from collections import Counter

import fitz

import pdf_to_ppt as P

HERE = _자료
SOURCES = [_기본,
           r"C:\Users\jbzle\OneDrive\Desktop\IM.pdf"]
SOURCES += sorted(glob.glob(
    r"C:\Users\jbzle\Desktop\종합\자동화\IM\_예전자료"
    r"\rainfield-im(변환기 원본)\PDF,워드(원본)\*.pdf"))


def lum(hx):
    r, g, b = (int(hx[i:i + 2], 16) / 255 for i in (0, 2, 4))
    return 0.299 * r + 0.587 * g + 0.114 * b


def main():
    tally = Counter()          # 색 -> 칸 수
    withwhite = Counter()      # 색 -> 그 위 글자가 흰색인 칸 수
    for path in SOURCES:
        if not os.path.exists(path):
            continue
        doc = fitz.open(path)
        for i in range(doc.page_count):
            pg = doc[i]
            txt = P._lines_of(pg)
            rects, plan, _boxes = P._plan(pg, txt)
            fills = [it for it in rects if it["filled"]]
            for t in plan:
                xs, ys, cells = t["grid"]
                nc, nr = len(xs) - 1, len(ys) - 1
                for (r0, c0, rs, cs) in cells:
                    box = fitz.Rect(xs[c0], ys[r0],
                                    xs[min(c0 + cs, nc)], ys[min(r0 + rs, nr)])
                    hx = P._cell_fill_hex(fills, box)
                    if not hx:
                        continue
                    tally[hx] += 1
                    cols = [c["hex"] for l in txt for c in l["chars"]
                            if c["c"].strip()
                            and box.contains(fitz.Point(
                                (c["bbox"].x0 + c["bbox"].x1) / 2,
                                (c["bbox"].y0 + c["bbox"].y1) / 2))]
                    if cols and all(lum(h) > 0.85 for h in cols):
                        withwhite[hx] += 1
        doc.close()

    print(f"{'색':<9}{'밝기':>6}{'칸수':>7}{'흰글씨칸':>9}")
    for hx, n in sorted(tally.items(), key=lambda kv: lum(kv[0])):
        print(f"#{hx:<8}{lum(hx):>6.2f}{n:>7}{withwhite.get(hx, 0):>9}")


if __name__ == "__main__":
    main()
