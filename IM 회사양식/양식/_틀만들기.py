# -*- coding: utf-8 -*-
"""사용자 양식(형식_1.docx) → 웹에 올릴 '회사양식_틀.docx' 만들기.

형식_1.docx 에는 실제 딜 내용(신촌지역 사업명·표지 조감도·고산3지구 사모사채 개요 값)이 들어 있어
공개 저장소에 올리면 안 된다. 도구는 이 부분을 어차피 매번 새로 채우므로, 딜 내용만 지운 틀을
따로 만들어 올린다. (연락처는 기존 IM·요약본 템플릿에 이미 공개돼 있어 그대로 둔다.)

★형식_1.docx 를 고쳤으면 이걸 다시 돌려야 웹에도 반영된다:
    python _틀만들기.py
"""
import io
import os
import re
import zipfile

from PIL import Image

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "형식_1.docx")
DST = os.path.join(HERE, "회사양식_틀.docx")


def main():
    zin = zipfile.ZipFile(SRC)
    out = io.BytesIO()
    zout = zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED)
    rels = zin.read("word/_rels/document.xml.rels").decode("utf-8")
    body_imgs = set(re.findall(r'Target="(media/[^"]+)"', rels))     # 본문(표지) 그림만 — 머리말·꼬리말 로고는 둔다
    for item in zin.infolist():
        data = zin.read(item.filename)
        if item.filename.startswith("word/media/") and item.filename[5:] in body_imgs:
            im = Image.new("RGB", (40, 20), (200, 200, 200))         # 표지 조감도 → 회색 빈 그림
            b = io.BytesIO()
            fmt = "PNG" if item.filename.lower().endswith(".png") else "JPEG"
            im.save(b, fmt)
            data = b.getvalue()
        elif item.filename == "word/document.xml":
            x = data.decode("utf-8")
            for old, new in [("신촌지역 정비4구역 4-12, 4-1지구", "사업명"), ("브릿지대출 사모사채 제안서", "사모사채 제안서")]:
                x = x.replace(old, new)
            # 사모사채 개요 표의 내용 칸 글을 지운다(구분 칸은 둔다)
            x = re.sub(r"(<w:tbl>.*?사모사채명.*?</w:tbl>)", _clear_values, x, count=1, flags=re.S)
            data = x.encode("utf-8")
        elif re.match(r"word/header\d+\.xml$", item.filename):
            x = data.decode("utf-8")
            x = re.sub(r"(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)",
                       lambda m: m.group(1) + ("" if m.group(2).strip() and "제안서" not in m.group(2) else
                                               ("사업명 사모사채 제안서" if "제안서" in m.group(2) else m.group(2)))
                       + m.group(3), x)
            data = x.encode("utf-8")
        zout.writestr(item, data)
    zout.close()
    with open(DST, "wb") as f:
        f.write(out.getvalue())
    print("만듦:", DST)


def _clear_values(m):
    tbl = m.group(1)
    rows = re.findall(r"<w:tr[ >].*?</w:tr>", tbl, re.S)
    for r in rows:
        cells = re.findall(r"<w:tc>.*?</w:tc>", r, re.S)
        new_r = r
        for c in cells[1:]:                                  # 첫 칸(구분)만 남기고 글 지우기
            new_r = new_r.replace(c, re.sub(r"(<w:t(?: [^>]*)?>)[^<]*(</w:t>)", r"\1\2", c), 1)
        tbl = tbl.replace(r, new_r, 1)
    return tbl


if __name__ == "__main__":
    main()
