# -*- coding: utf-8 -*-
"""한 문서를 **전 쪽** 변환해서 원본과 나란히 그림으로 뽑는다.

★쪽마다 사정이 다르다. 지적받은 쪽만 고치면 안 본 쪽은 그대로 남는다.
  한 번에 전부 뽑아 놓고 훑을 것.

사용: python _전체보기.py [원본PDF경로] [내보낼폴더]
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
PDF = sys.argv[1] if len(sys.argv) > 1 else _기본
OUT = os.path.normpath(os.path.abspath(
    sys.argv[2] if len(sys.argv) > 2 else os.path.join(HERE, "_확인전체")))
PPT = os.path.join(OUT, "_결과.pptx")

os.makedirs(OUT, exist_ok=True)
info = P.convert(open(PDF, "rb").read(), PPT)
print(f"변환 {info['count']}쪽 → {PPT}")

doc = fitz.open(PDF)
W = H = None
for i in range(doc.page_count):
    pm = doc[i].get_pixmap(dpi=110)
    pm.save(os.path.join(OUT, f"원본_{i + 1:02d}.png"))
    W, H = pm.width, pm.height
doc.close()

import win32com.client

ppt = win32com.client.Dispatch("PowerPoint.Application")
try:
    pres = ppt.Presentations.Open(os.path.abspath(PPT), WithWindow=False)
except Exception:
    pres = ppt.Presentations.Open(os.path.abspath(PPT))
for i in range(pres.Slides.Count):
    dst = os.path.normpath(os.path.join(OUT, f"결과_{i + 1:02d}.png"))
    try:
        pres.Slides(i + 1).Export(dst, "PNG", W, H)
    except Exception as e:
        print(f"  {i + 1}쪽 그림 실패: {e}")
pres.Close()
try:
    ppt.Quit()
except Exception:
    pass
print(f"그림 {info['count']}쌍 저장 → {OUT}")
