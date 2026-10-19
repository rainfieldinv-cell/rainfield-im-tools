# -*- coding: utf-8 -*-
"""틀 파일(pptx·docx)의 한글 글꼴을 회사 글꼴 두 개로 맞춘다(2026-10-19 사용자 확정).

  두꺼운 글씨 → 'Pretendard ExtraBold'
  얇은 글씨   → 'Pretendard Light'
  (피플폰트는 더 이상 안 쓴다 — IM 도구 4개 전부)

바꾸는 글꼴: 피플폰트·Pretendard 계열(굵기 여러 가지)·Apple SD 산돌고딕·맑은 고딕·테마 글꼴.
굵기는 이름으로 가른다(Bold·Black·Heavy·SemiBold·ExtraBold → 두꺼운, 나머지 → 얇은).
docx 는 이름에 굵기가 없으면(Pretendard Variable 등) 그 글자의 '굵게' 표시로 가른다.
Arial·Wingdings·Segoe UI Symbol 같은 영문·기호 글꼴은 그대로 둔다.

사용: python 글꼴_바꾸기.py 파일.pptx [파일.docx …]   (제자리에 고쳐 쓴다)
"""
import re
import sys
import zipfile
import shutil
import os
import tempfile

BOLD = "Pretendard ExtraBold"
LIGHT = "Pretendard Light"
_KO = re.compile(r"피플폰트|pretendard|applesdgothic|apple sd|맑은 고딕|malgun|굴림|바탕|돋움|궁서|gulim|batang|dotum|나눔|nanum|^HY|한컴|함초롬", re.I)
_HEAVY = re.compile(r"bold|black|heavy|B$|EB$|SB$|H$|B00$|EB00$|H00$", re.I)


def map_name(name, bold=None):
    """글꼴 이름 → 회사 글꼴(바꿀 대상이 아니면 None)."""
    if not name or not _KO.search(name):
        return None
    if _HEAVY.search(name.replace("Bold00", "Bold")):
        return BOLD
    if re.search(r"light|thin|regular|medium|L$|M00$|R00$|M$", name, re.I):
        return LIGHT
    return BOLD if bold else LIGHT                      # 굵기 없는 이름(Pretendard Variable 등)


def _pptx_xml(s):
    def sub_rpr(m):
        tag = m.group(0)
        bold = ' b="1"' in tag.split(">", 1)[0]

        def tf(mm):
            new = map_name(mm.group(2), bold)
            return mm.group(1) + (new or mm.group(2)) + mm.group(3)
        return re.sub(r'(<a:(?:latin|ea|cs) typeface=")([^"]+)(")', tf, tag)
    s = re.sub(r"<a:(?:rPr|defRPr|endParaRPr)\b[^>]*?/>|<a:(?:rPr|defRPr|endParaRPr)\b.*?</a:(?:rPr|defRPr|endParaRPr)>",
               sub_rpr, s, flags=re.S)
    # 테마 글꼴(제목 = 두꺼운, 본문 = 얇은)
    def theme(m):
        body = m.group(0)
        f = BOLD if m.group(1) == "major" else LIGHT

        def tf(mm):
            return mm.group(1) + (f if (map_name(mm.group(2)) or mm.group(2) == "") else mm.group(2)) + mm.group(3)
        body = re.sub(r'(<a:(?:latin|ea) typeface=")([^"]*)(")', tf, body)
        return re.sub(r'(<a:font script="Hang" typeface=")([^"]*)(")', lambda mm: mm.group(1) + f + mm.group(3), body)
    s = re.sub(r"<a:(major|minor)Font>.*?</a:\1Font>", theme, s, flags=re.S)
    return s


def _docx_xml(s):
    def sub_rpr(m):
        tag = m.group(0)
        bold = bool(re.search(r'<w:b(?: w:val="(?:1|true|on)")?\s*/>', tag))

        def at(mm):
            new = map_name(mm.group(2), bold)
            return mm.group(1) + (new or mm.group(2)) + mm.group(3)
        return re.sub(r'(w:(?:ascii|hAnsi|eastAsia|cs)=")([^"]+)(")', at, tag)
    return re.sub(r"<w:rPr\b.*?</w:rPr>|<w:rPr\s*/>", sub_rpr, s, flags=re.S)


def convert(path):
    is_docx = path.lower().endswith((".docx", ".dotx"))
    zin = zipfile.ZipFile(path)
    fd, tmp = tempfile.mkstemp(suffix=os.path.splitext(path)[1])
    os.close(fd)
    changed = 0
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename.endswith(".xml"):
                s = data.decode("utf-8")
                new = _docx_xml(s) if is_docx else _pptx_xml(s)
                if new != s:
                    changed += 1
                    data = new.encode("utf-8")
            zout.writestr(item, data)
    zin.close()
    shutil.move(tmp, path)
    return changed


if __name__ == "__main__":
    for p in sys.argv[1:]:
        print(p, "→ 고친 부분", convert(p))
