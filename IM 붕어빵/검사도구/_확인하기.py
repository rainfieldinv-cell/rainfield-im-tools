# -*- coding: utf-8 -*-
"""만든 PPT 를 눈으로 확인한다 — 원본 PDF 쪽 그림과 나란히 뽑는다.

사용: python _확인하기.py [쪽번호 ...]      (안 적으면 1,2,3쪽)
결과: _확인/원본_NN.png · _확인/결과_NN.png
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

HERE = _자료
OUT = os.path.join(HERE, "_확인")
PDF = _기본
PPT = os.path.join(HERE, "_결과.pptx")

pages = [int(a) for a in sys.argv[1:]] or [1, 2, 3]
os.makedirs(OUT, exist_ok=True)

# ── 원본 PDF 쪽 그림 ────────────────────────────────
doc = fitz.open(PDF)
W = H = None
for p in pages:
    pg = doc[p - 1]
    pm = pg.get_pixmap(dpi=150)
    pm.save(os.path.join(OUT, f"원본_{p:02d}.png"))
    W, H = pm.width, pm.height
doc.close()
print(f"원본 {len(pages)}장 저장 ({W}x{H})")

# ── 만든 PPT 슬라이드 그림 (파워포인트) ──────────────
import win32com.client

ppt = win32com.client.Dispatch("PowerPoint.Application")
try:
    pres = ppt.Presentations.Open(os.path.abspath(PPT), WithWindow=False)
except Exception:
    pres = ppt.Presentations.Open(os.path.abspath(PPT))
for p in pages:
    if p <= pres.Slides.Count:
        pres.Slides(p).Export(os.path.join(OUT, f"결과_{p:02d}.png"),
                              "PNG", W, H)
pres.Close()
try:
    ppt.Quit()
except Exception:
    pass
print(f"결과 {len(pages)}장 저장 → {OUT}")
