# -*- coding: utf-8 -*-
"""아무 IM PDF나 넣어서 **검사 + 원본/결과 그림 대조** 까지 한 번에.

사용: python _다른원본검사.py "<원본.pdf>" [쪽번호 ...]
결과: _확인2\원본_NN.png · _확인2\결과_NN.png
"""
import os
import sys

# ★이 파일은 '검사도구' 폴더 안에 있다. 프로그램(pdf_to_ppt)과 자료(_원본.pdf)는
#   옆 폴더에 있으므로 찾아갈 길을 열어 준다.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, _ROOT)   # 2026-10-06 한 폴더로 합침: 엔진(pdf_to_ppt)이 바로 윗 폴더에 있다
_자료 = os.path.join(_ROOT, "자료")
import os
import sys

import fitz

import pdf_to_ppt as P
import _선검사
import _정렬검사

HERE = _자료
OUT = os.path.join(HERE, "_확인2")
PPT = os.path.join(HERE, "_결과2.pptx")

src = sys.argv[1]
pages = [int(a) for a in sys.argv[2:]]
raw = open(src, "rb").read()
os.makedirs(OUT, exist_ok=True)

info = P.convert(raw, PPT)
print(f"=== {os.path.basename(src)} · {info['count']}쪽 ===")
for s in info["pages"]:
    print(f"  {s['쪽']:>3}쪽 : 표 {s['표']:>2} · 네모 {s['네모']:>4} · "
          f"그림 {s['그림']:>2} · 글상자 {s['글상자']:>4}")

doc = fitz.open(stream=raw, filetype="pdf")
print("\n--- 선 검사 ---")
bad = sum(_선검사.check(i + 1, doc[i]) for i in range(doc.page_count))
print(f"합계 : 못 넣은 선 {bad}개")
print("\n--- 정렬 검사 ---")
tb = tc = 0
for i in range(doc.page_count):
    a, b = _정렬검사.check(i + 1, doc[i])
    tb += a
    tc += b
print(f"합계 : 어긋난 칸 {tb} / {tc}개")

if pages:
    W = H = None
    for p in pages:
        pm = doc[p - 1].get_pixmap(dpi=150)
        pm.save(os.path.join(OUT, f"원본_{p:02d}.png"))
        W, H = pm.width, pm.height
    import win32com.client
    app = win32com.client.Dispatch("PowerPoint.Application")
    try:
        pres = app.Presentations.Open(os.path.abspath(PPT), WithWindow=False)
    except Exception:
        pres = app.Presentations.Open(os.path.abspath(PPT))
    for p in pages:
        if p <= pres.Slides.Count:
            pres.Slides(p).Export(os.path.join(OUT, f"결과_{p:02d}.png"),
                                  "PNG", W, H)
    pres.Close()
    try:
        app.Quit()
    except Exception:
        pass
    print(f"\n그림 저장 → {OUT}")
doc.close()
