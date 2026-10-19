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
SRC = os.path.join(HERE, "회사양식_틀(최종).docx")          # 2026-10-12 사용자 최종 양식(하이라이트 쪽 포함)
if not os.path.exists(SRC):
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
        if item.filename.startswith("word/media/") and item.filename[5:] in body_imgs and len(data) > 50000:
            # 사진(표지 조감도·하이라이트 지도)만 회색으로 — 작은 그림(체크 아이콘 등)은 양식 그대로
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
            # 사모사채 개요 표의 기본 글은 둔다(2026-10-13 양식: 'TBD(신규 유동화 SPC)…' 같은 공통 문구가 기본값).
            # 실제 딜 이름만 지운다.
            x = _genericize_deal(x)
            data = _clear_highlight(x.encode("utf-8"))
        elif re.match(r"word/header\d+\.xml$", item.filename):
            x = data.decode("utf-8")
            x = re.sub(r"(<w:t(?: [^>]*)?>)([^<]*)(</w:t>)",
                       lambda m: m.group(1) + ("" if m.group(2).strip() and "제안서" not in m.group(2) else
                                               ("사업명 사모사채 제안서" if "제안서" in m.group(2) else m.group(2)))
                       + m.group(3), x)
            data = x.encode("utf-8")
        elif re.match(r"word/footer\d+\.xml$", item.filename):
            # 쪽번호 글자 크기 8pt(sz 16) — 2026-10-12 사용자 확정 '쪽번호 양식을 아예 8로'
            data = _page_no_8(data)
        zout.writestr(item, data)
    zout.close()
    with open(DST, "wb") as f:
        f.write(out.getvalue())
    print("만듦:", DST)


def _clear_highlight(xml):
    """회사 하이라이트 쪽(두 번째 문단의 첫 쪽나눔 뒤 그룹들)의 실제 딜 글을 지운다. 모양·그룹·제목은 그대로.
    '<…>' 소제목 → '<하이라이트 제목>', 그 밖의 글 → '내용을 입력하세요'(문단마다 한 번). 'Overview'·'Deal Highlight' 는 둔다."""
    from lxml import etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    root = etree.fromstring(xml)
    body = root.find(W + "body")
    paras = [el for el in body if el.tag == W + "p"]
    if len(paras) < 2:
        return xml
    runs = [r for r in paras[1] if r.tag == W + "r"]
    k = next((i for i, r in enumerate(runs)
              if any(b.get(W + "type") == "page" for b in r.findall(W + "br"))), None)
    if k is None:
        return xml
    keep = {"Overview", "Deal Highlight"}
    for r in runs[k + 1:]:
        for p in r.iter(W + "p"):
            ts = list(p.iter(W + "t"))
            txt = "".join(t.text or "" for t in ts).strip()
            if not txt or txt in keep:
                continue
            new = "<하이라이트 제목>" if txt.startswith("<") and txt.endswith(">") else "내용을 입력하세요"
            ts[0].text = new
            for t in ts[1:]:
                t.text = ""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _genericize_deal(x):
    """사모사채 개요 '기초자산' 칸의 실제 딜 이름(고산3지구 등)을 'OOO' 로. 글이 여러 조각이라 칸 단위로 고친다."""
    from lxml import etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    root = etree.fromstring(x.encode("utf-8"))
    for tr in root.iter(W + "tr"):
        tcs = tr.findall(W + "tc")
        if len(tcs) >= 2 and "".join(t.text or "" for t in tcs[0].iter(W + "t")).strip() == "기초자산":
            ts = list(tcs[-1].iter(W + "t"))
            full = "".join(t.text or "" for t in ts)
            m = re.search(r"\S*?\d?\s*지구\s*", full)
            if m:
                full = full[:m.start()] + "OOO " + full[m.end():]
            if ts:
                ts[0].text = full
                for t in ts[1:]:
                    t.text = ""
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True).decode("utf-8")


def _page_no_8(xml):
    """꼬리말에서 PAGE 필드가 든 문단의 글자 크기를 전부 8pt 로. (글상자 안 문단이 섞여 있어 정규식 말고 XML 로 읽는다)"""
    from lxml import etree
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    root = etree.fromstring(xml)
    for it in root.iter(W + "instrText"):
        if "PAGE" not in (it.text or ""):
            continue
        p = next(a for a in it.iterancestors() if a.tag == W + "p")
        for r in p.iter(W + "r"):
            rpr = r.find(W + "rPr")
            if rpr is None:
                rpr = etree.Element(W + "rPr")
                r.insert(0, rpr)
            for tag in ("sz", "szCs"):
                e = rpr.find(W + tag)
                if e is None:
                    e = etree.SubElement(rpr, W + tag)
                e.set(W + "val", "16")
        rpr = p.find(f"{W}pPr/{W}rPr")
        if rpr is not None:
            for tag in ("sz", "szCs"):
                e = rpr.find(W + tag)
                if e is not None:
                    e.set(W + "val", "16")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


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
    # 글꼴을 회사 글꼴 두 개로(두꺼운 = Pretendard ExtraBold, 얇은 = Pretendard Light) — 2026-10-19 사용자 확정
    import importlib.util
    _g = os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", "..", "글꼴_바꾸기.py")
    _spec = importlib.util.spec_from_file_location("글꼴_바꾸기", _g)
    _mod = importlib.util.module_from_spec(_spec)
    _spec.loader.exec_module(_mod)
    print("글꼴 맞춤:", _mod.convert(os.path.join(os.path.dirname(os.path.abspath(__file__)), "회사양식_틀.docx")))
