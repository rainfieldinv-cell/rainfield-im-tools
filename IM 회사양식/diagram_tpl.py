# -*- coding: utf-8 -*-
"""금융 구조도 — 'IM 요약본' 도구의 구조도 단계를 그대로 쓴다(2026-10-14 사용자 확정: '요약본처럼 똑같이').

요약본의 파일을 복사하지 않고 그 폴더에서 바로 불러온다 → 구조도 틀(layout/금융구조도_모음.pptx)·썸네일·
채우기 규칙을 고치면 두 도구에 같이 반영된다.
흐름(요약본이 요약본에 구조도를 넣을 때와 같음): 틀 고르기 → 칸 채우기 → 한 장짜리 PPT → 로고·제목·액자 빼기
→ 사진(내 PC: 파워포인트 / 웹: LibreOffice) → 구조도 부분만 잘라 내기.
"""
import importlib.util
import os
import sys
import tempfile

_HERE = os.path.dirname(os.path.abspath(__file__))
_SUM = os.path.join(os.path.dirname(_HERE), "IM 요약본")


def _load(name, file):
    key = "_cf_sum_" + name
    if key in sys.modules:
        return sys.modules[key]
    spec = importlib.util.spec_from_file_location(key, os.path.join(_SUM, file))
    mod = importlib.util.module_from_spec(spec)
    sys.modules[key] = mod
    spec.loader.exec_module(mod)
    return mod


def _dg():
    # 요약본 diagram.py 가 안에서 'from engine_bits import …' 로 부른다 → 같은 파일을 그 이름으로도 등록
    sys.modules.setdefault("engine_bits", _eb())
    return _load("diagram", "diagram.py")


def _eb():
    return _load("engine_bits", "engine_bits.py")


def layouts():
    """[{no, title, fields:[{key, kind, mark, text}]}]"""
    return _dg().read_layouts()


def tags():
    import json
    try:
        raw = json.load(open(os.path.join(_SUM, "layout", "구조도_분류.json"), encoding="utf-8"))
        return {int(k): v for k, v in raw.items() if k.isdigit()}
    except Exception:
        return {}


def thumb(no):
    return _dg().thumb_path(no)


def autofill(layout, hint, amounts):
    """요약본 _diag_autofill 과 같은 규칙. hint = {차주, 시공사, 신탁사, 연대보증, 발행, 사채권자}, amounts = [금액…].
    못 찾으면 틀의 글을 그대로 둔다(지어내지 않는다)."""
    out, used = {}, 0
    for f in layout["fields"]:
        txt = (f.get("text") or "").strip()
        if not txt:
            continue
        if f["mark"].endswith("-1"):                 # 상자 두 줄짜리: 윗줄 역할(차주/시공사) → 아랫줄 이름
            head = f["mark"][0]
            role = next((g["text"] for g in layout["fields"] if g["mark"] == head), "")
            for key, val in hint.items():
                if val and key in role:
                    out[f["key"]] = str(val)
                    break
            continue
        low = txt.replace(" ", "").lower()
        if low.startswith("tr.") or "억원" in txt:
            if used < len(amounts):
                if amounts[used]:
                    out[f["key"]] = str(amounts[used])
                used += 1
    return out


def build(no, values, removed=None):
    """채운 구조도 → (pptx bytes, 잘라 낸 사진 png bytes 또는 None, 오류 글 또는 None)."""
    dg, eb = _dg(), _eb()
    p = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False).name
    dg.build_diagram(no, values, p, removed=set(removed or ()))
    with open(p, "rb") as f:
        pptx = f.read()
    clean = dg.strip_chrome(p)
    png, err = eb.pptx_slide_png(clean, 1)
    if png:
        png = dg.crop_to_content(png, dg.content_bbox(clean))
    return pptx, png, err


def pptx_to_png(pptx_bytes):
    """사람이 고쳐 올린 구조도 PPT → 잘라 낸 사진."""
    dg, eb = _dg(), _eb()
    p = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False).name
    with open(p, "wb") as f:
        f.write(pptx_bytes)
    clean = dg.strip_chrome(p)
    png, err = eb.pptx_slide_png(clean, 1)
    if png:
        png = dg.crop_to_content(png, dg.content_bbox(clean))
    return png, err
