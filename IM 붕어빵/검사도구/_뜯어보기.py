# -*- coding: utf-8 -*-
"""원본에서 글자 조각이 어떻게 쪼개져 있는지 뜯어본다(간격·기호·워터마크 확인)."""
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
doc = fitz.open(_기본)
pno = int(sys.argv[1]) - 1 if len(sys.argv) > 1 else 2
page = doc[pno]

print(f"--- {pno+1}쪽 · 주석(annot) ---")
try:
    ann = list(page.annots() or [])
    print(f"주석 {len(ann)}개 :", [(a.type, a.info.get('content', '')[:20]) for a in ann])
except Exception as e:
    print("주석 읽기 실패", e)

print(f"\n--- 글자 조각 (처음 6줄) ---")
d = page.get_text("dict")
cnt = 0
for b in d["blocks"]:
    if b["type"] != 0:
        continue
    for ln in b["lines"]:
        cnt += 1
        if cnt > 6:
            break
        print(f"[줄 {cnt}] dir={ln['dir']} bbox={tuple(round(v,1) for v in ln['bbox'])}")
        prev = None
        for sp in ln["spans"]:
            t = sp["text"]
            codes = " ".join(f"U+{ord(c):04X}" for c in t[:4])
            gap = "" if prev is None else f" 앞칸과간격={sp['bbox'][0]-prev:.2f}pt"
            print(f"    {sp['font']:<22} {sp['size']:.1f}pt "
                  f"x={sp['bbox'][0]:.1f}~{sp['bbox'][2]:.1f}{gap}")
            print(f"      글자={t!r}  {codes}")
            prev = sp["bbox"][2]
    if cnt > 6:
        break

print("\n--- 워터마크 후보 : 회색/연한 글자 ---")
for b in d["blocks"]:
    if b["type"] != 0:
        continue
    for ln in b["lines"]:
        for sp in ln["spans"]:
            if abs(ln["dir"][1]) > 0.01:
                print(f"  기울어진 글자 : {sp['text']!r} dir={ln['dir']} "
                      f"font={sp['font']} color={sp['color']:06X}")
print("\n--- 그림 블록 ---")
for b in d["blocks"]:
    if b["type"] == 1:
        print(f"  bbox={tuple(round(v,1) for v in b['bbox'])} "
              f"크기={len(b.get('image') or b'')}바이트")
doc.close()
