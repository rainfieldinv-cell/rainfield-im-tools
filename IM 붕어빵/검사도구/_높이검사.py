# -*- coding: utf-8 -*-
"""표가 원본보다 세로로 늘어나지 않았는지 잰다(파워포인트가 실제로 그린 높이)."""
import os
import sys

# ★이 파일은 '검사도구' 폴더 안에 있다. 프로그램(pdf_to_ppt)과 자료(_원본.pdf)는
#   옆 폴더에 있으므로 찾아갈 길을 열어 준다.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)   # 2026-10-06 한 폴더로 합침: 엔진(pdf_to_ppt)이 바로 윗 폴더에 있다
_자료 = os.path.join(_ROOT, "자료")
import os
import sys

import win32com.client
from pptx import Presentation

EMU = 914400 / 72.0
path = os.path.abspath(sys.argv[1])
prs = Presentation(path)
app = win32com.client.Dispatch("PowerPoint.Application")
try:
    pres = app.Presentations.Open(path, WithWindow=False)
except Exception:
    pres = app.Presentations.Open(path)
bad = tot = 0
worst = []
for i, sl in enumerate(prs.slides, 1):
    want = [sh.height / EMU for sh in sl.shapes
            if getattr(sh, "has_table", False) and sh.has_table]
    if not want:
        continue
    s = pres.Slides(i)
    real = [s.Shapes(k).Height for k in range(1, s.Shapes.Count + 1)
            if s.Shapes(k).HasTable]
    for a, b in zip(want, real):
        tot += 1
        if abs(a - b) > 0.5:
            bad += 1
            worst.append((b - a, i, a, b))
pres.Close()
app.Quit()
worst.sort(reverse=True)
print(f"{os.path.basename(path)} : 표 {tot}개 중 높이 어긋난 것 {bad}개")
for d, i, a, b in worst[:5]:
    print(f"   {i}쪽 {a:.0f} → {b:.0f} ({d:+.0f}pt)")
