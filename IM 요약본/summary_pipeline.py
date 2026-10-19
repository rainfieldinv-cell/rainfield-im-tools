# -*- coding: utf-8 -*-
"""요약본 파이프라인 — 원본 PDF → 5~6장 요약본.
  extract_summary(pdf_bytes) -> data(dict)
  build_summary(data, pdf_path, out_path, pages=1) -> out_path
"""
import io, os, json, sys, re
import datetime as _dt
from copy import deepcopy

# ── 경로 ─────────────────────────────────────────────
#   생성기 코드는 'IM 자동화\생성기\' 에 있고,
#   변환 엔진(extractors.py, modules/)은 옆 저장소 'rainfield-im' 에 있다.
# ★요약본 생성기는 'IM 자동화' 폴더 안에서 **독립적으로** 돈다.
#   전엔 옆 저장소(rainfield-im)의 extractors·modules 를 끌어다 썼는데,
#   그 폴더가 움직이거나 정리될 때마다 깨졌다(경로·API키·캐시 3번 터짐).
#   → 필요한 것(extractors.py, claude_api.py, .env, 캐시)을 이 폴더에 두고 그것만 쓴다.
_HERE = os.path.dirname(os.path.abspath(__file__))        # ...\IM 자동화\_내부\생성기
_WORK = os.path.abspath(os.path.join(_HERE, ".."))        # ...\IM 자동화\_내부
if _HERE not in sys.path:
    sys.path.insert(0, _HERE)

from dotenv import load_dotenv
load_dotenv(os.path.join(_HERE, ".env"))                  # 로컬: 같은 폴더 .env
if not os.environ.get("ANTHROPIC_API_KEY"):
    # 배포(Streamlit Cloud): .env 가 없으므로 secrets 에서 받는다.
    try:
        import streamlit as _st
        _k = _st.secrets.get("ANTHROPIC_API_KEY", "")
        if _k:
            os.environ["ANTHROPIC_API_KEY"] = _k
    except Exception:
        pass

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
try:
    from pptx.enum.text import MSO_ANCHOR
except ImportError:          # python-pptx 버전에 따라 이름이 다르다
    from pptx.enum.text import MSO_VERTICAL_ANCHOR as MSO_ANCHOR
from extractors import extract_from_pdf
import claude_api as _claude_api
from claude_api import call_claude
from engine_bits import (replace_text_keep_runs as _replace_text_keep_runs,
                         extract_page_images, repeated_xrefs, mean_brightness)

from pathlib import Path as _Path
_claude_api.CACHE_DIR = _Path(_HERE) / ".claude_cache"    # 캐시도 이 폴더에 고정
_claude_api.CACHE_DIR.mkdir(parents=True, exist_ok=True)

# 레이아웃 틀 — 로컬(IM 자동화\_내부\레이아웃\)과 배포(생성기\layout\) 둘 다 지원
_LAYOUT_CANDIDATES = [
    os.path.join(_HERE, "layout", "요약본 레이아웃.pptx"),   # 저장소 배포용
    os.path.join(_WORK, "레이아웃", "요약본 레이아웃.pptx"),  # 로컬
]


def _layout_path():
    for p in _LAYOUT_CANDIDATES:
        if os.path.exists(p):
            return p
    raise FileNotFoundError(
        "레이아웃 틀(요약본 레이아웃.pptx)을 찾을 수 없습니다.\n" +
        "\n".join(f"  - {p}" for p in _LAYOUT_CANDIDATES))


SYS = """당신은 부동산 금융 IM(투자설명서)을 읽고 '요약본 제안서'용 핵심 데이터를 뽑는 전문가다.
아래 원문에서 정보를 찾아 **JSON만** 출력한다. 규칙:
- ★**요약하지 마라.** 원문에 적힌 문장을 그대로 옮긴다.
  줄이거나 다듬거나 말을 바꾸지 않는다. 여러 줄을 한 줄로 합치지도 않는다.
  원문이 길면 긴 대로 그대로 둔다(넣을지 뺄지는 사용자가 나중에 고른다).
- 값은 원문 그대로(verbatim). 숫자·단위·괄호([30]억원, [43.98]% 등 대괄호 표기)도 원문대로 유지.
- 원문에 없는 값은 지어내지 말고 null. 확인 불가도 null.
- '본건'에 해당하는 트랜치(Tr)는 is_bongeon=true.

출력 JSON 스키마:
{
 "deal_name": "사업명(표지용)",
 "date_ko": "발행일 또는 제안 시점(예: 2026년 8월)",
 "overview": ["사업 개요 핵심 문장(원문 그대로) 3~4줄 — 요약본 'Overview' 칸에 들어간다"],
 "highlights": [
   {"title":"12자내 짧은 제목", "subtitle":"20자내 한 줄 부제", "bullets":["근거 1(상세)","근거 2(상세)"]}
 ],
 "사모사채개요": {"사모사채명":..,"사채유형":..,"발행인":..,"기초자산":..,"발행금액":..,"발행일":..,"만기일":..,"금융조건":..,"이자지급주기":..},
 "담보대출조건": {"금융구조":..,"차주":..,"tranches":[{"구분":"Tr.A","대출금액":..,"LTV":..,"금리":..,"is_bongeon":false}],"LTV기준_각주":..,"주요채권보전":["..",".."],"주요인출선행조건":["..",".."],"상환재원":..,"대출기간":..},
 "사업일정": {"완료":[{"시기":"'23.02","내용":".."}],"예정":[{"시기":"'26.4Q","내용":".."}]},
 "사업개요": {"사업명":..,"대지위치":..,"대지면적":..,"건축면적":..,"연면적":..,"건축규모":..},
 "투입에쿼티": [{"주체":..,"금액":..,"비고":..}],
 "법인개요": {"대상":..,"회사명":..,"설립일":..,"사업자번호":..,"대표자":..,"자본금":..,"주요사업":..,"주요주주":..},
 "재무제표": {"단위":"백만원","연도":["2025.12"],"행":[{"항목":"자산총계","값":[".."]}]},
 "중요항목": [{"제목":..,"내용":["..",".."]}],
 "이미지_있음": {"금융구조도": false, "조감도": true, "위치도": false}
}
사업일정·재무제표 같은 표는 원본에서 그대로 읽으므로 여기서는 대충 채워도 된다.

★"중요항목" — ※표 안에 있는 내용은 담지 마라. 표는 원본에서 그대로 읽어 쓴다.
  표가 아닌 **줄글 설명**만 여기에 담는다(없으면 빈 목록).
  (예: 분양·임대 현황, 인근 시세, 시공사 개요, 신용보강·안전장치, 자금수지,
   토지 확보 현황, 사업비·수지 분석, 조합원 현황, 임차인 현황, 감정평가 결과 등)
  - "제목"은 원문의 표·문단 제목을 그대로 쓴다(예: '인근 아파트 시세').
  - "내용"의 각 줄은 원문의 표를 옮길 때 "항목 : 값" 형태로 쓰고,
    표가 아닌 설명이면 문장 한 줄로 쓴다.
  - 원문 문장을 줄여 쓰지 말고 그대로 옮긴다. 최대 6개.
JSON 외 다른 말 금지."""


def extract_summary(pdf_bytes: bytes) -> dict:
    data = extract_from_pdf(pdf_bytes)
    full = "\n".join(data.get("pages_text", []))
    res = call_claude(SYS, f"[IM 원문]\n{full}", slide_num=801,
                      pdf_context=full, prompt_version="summary_extract_v6")
    if not res.get("ok"):
        raise RuntimeError(res.get("error") or "추출 실패")
    return res["data"]


# ── 빌드 헬퍼 ────────────────────────────────────────
def _walk(shapes):
    for sh in shapes:
        yield sh
        if sh.shape_type == 6:
            yield from _walk(sh.shapes)


def _set(table, r, c, text):
    if r < len(table.rows) and c < len(table.columns):
        _replace_text_keep_runs(table.cell(r, c).text_frame, str(text if text is not None else ""))


def _lines(table, r, c, lines):
    _replace_text_keep_runs(table.cell(r, c).text_frame, "\n".join(str(x) for x in lines if x))


def _add_rows(table, n):
    for _ in range(n):
        table._tbl.append(deepcopy(table._tbl.tr_lst[-1]))


def _del_rows(table, n):
    for _ in range(n):
        tr = table._tbl.tr_lst[-1]; tr.getparent().remove(tr)


def _replace_contains(slide, needle, new_text):
    for sh in _walk(slide.shapes):
        if sh.has_text_frame and needle in sh.text_frame.text:
            _replace_text_keep_runs(sh.text_frame, new_text)
            return True
    return False


def _find_label(slide, needle):
    """슬라이드에서 특정 문구가 들어간 '라벨 글상자'를 찾는다(표 제목용)."""
    for sh in slide.shapes:
        if sh.has_text_frame and needle in sh.text_frame.text:
            return sh
    return None


def _dup_slide_after(prs, src, after_idx):
    """슬라이드를 통째로 복제해 after_idx 바로 뒤에 넣는다(핵심요약 2페이지용).

    python-pptx 에 복제 기능이 없어 직접 한다:
      ①같은 레이아웃으로 새 슬라이드 → 자동 생성된 자리표시자 제거
      ②원본 도형 XML 을 deepcopy 해서 그대로 붙임
      ③그림·차트 같은 외부 파트 관계(rel)를 함께 복사 — 안 하면 이미지가 깨진다
      ④맨 뒤에 생긴 슬라이드를 원하는 위치로 이동
    """
    import copy as _copy
    _R = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
    dst = prs.slides.add_slide(src.slide_layout)
    for sh in list(dst.shapes):
        sh._element.getparent().remove(sh._element)
    for sh in src.shapes:
        el = _copy.deepcopy(sh._element)
        # ★그림이 가리키는 관계 번호(rId)는 **슬라이드마다 따로**다.
        #   XML 만 그대로 베끼면 새 슬라이드에서 그 번호가 없거나 엉뚱한 걸 가리켜
        #   그림이 안 보인다(2페이지에서 회사 로고가 사라지던 원인).
        #   → 원본이 쓰던 대상을 새 슬라이드에 다시 연결하고 번호를 바꿔 끼운다.
        for node in el.iter():
            for attr in ("embed", "link", "id"):
                key = _R + attr
                rid = node.get(key)
                if not rid:
                    continue
                try:
                    rel = src.part.rels[rid]
                except KeyError:
                    continue
                try:
                    node.set(key, dst.part.relate_to(
                        rel.target_ref if rel.is_external else rel._target,
                        rel.reltype, is_external=rel.is_external))
                except Exception:
                    pass
        dst.shapes._spTree.append(el)
    # 맨 뒤 → after_idx 다음 자리로 이동
    sldIdLst = prs.slides._sldIdLst
    ids = list(sldIdLst)
    sldIdLst.remove(ids[-1])
    sldIdLst.insert(after_idx + 1, ids[-1])
    return dst


def _clone_at_end(slide, proto):
    """도형을 복제해 같은 슬라이드 맨 뒤에 붙인다(표·글상자 전용).
       그림은 관계(rel)까지 따라가야 하므로 이 함수로 복제하지 않는다."""
    slide.shapes._spTree.append(deepcopy(proto._element))
    return slide.shapes[-1]


def extra_blocks_of(data: dict):
    """원본에서 찾은 '그 밖의 내용'을 [(제목, [(왼칸, 오른칸), ...]), ...] 로 만든다.

    고정 항목(사모사채개요·담보대출조건·사업일정·법인개요·재무제표)에 안 들어가는
    것도 요약본에 넣을 수 있어야 한다 → 표 한 벌로 그릴 수 있게 정리해 둔다.
    왼칸이 빈 줄은 두 칸을 합쳐 한 줄로 그린다(표가 아닌 설명 문장).
    """
    out = []

    ov = data.get("사업개요") or {}
    if isinstance(ov, dict):
        pairs = [(k, v) for k, v in ov.items() if v]
        if pairs:
            out.append(("사업개요", pairs))

    eq = data.get("투입에쿼티") or []
    if isinstance(eq, list) and eq:
        pairs = []
        for it in eq:
            if not isinstance(it, dict):
                continue
            rhs = " · ".join(str(x) for x in (it.get("금액"), it.get("비고")) if x)
            if it.get("주체") or rhs:
                pairs.append((it.get("주체") or "", rhs))
        if pairs:
            out.append(("투입 에쿼티", pairs))

    for it in (data.get("중요항목") or []):
        if not isinstance(it, dict):
            continue
        title = str(it.get("제목") or "").strip()
        lines = [str(x).strip() for x in (it.get("내용") or []) if str(x).strip()]
        if not title or not lines:
            continue
        pairs = []
        for ln in lines:
            # "항목 : 값" 이면 두 칸으로, 아니면 한 줄 통으로.
            k, sep, v = ln.partition(":")
            if not sep:
                k, sep, v = ln.partition("：")
            if sep and 0 < len(k.strip()) <= 14 and v.strip():
                pairs.append((k.strip(), v.strip()))
            else:
                pairs.append(("", ln))
        out.append((title, pairs))

    # 제목이 겹치면 뒤엣것에 번호를 붙여 칸 이름이 헷갈리지 않게 한다.
    seen, uniq = {}, []
    for title, pairs in out:
        n = seen.get(title, 0) + 1
        seen[title] = n
        uniq.append((title if n == 1 else f"{title} ({n})", pairs))
    return uniq


def _build_extra_block(slide, proto_tbl, proto_lbl, title, pairs):
    """[라벨 + 2열 표] 한 벌을 새로 만들어 붙인다. 반환 (라벨, 표)."""
    lb = None
    if proto_lbl is not None:
        lb = _clone_at_end(slide, proto_lbl)
        _replace_text_keep_runs(lb.text_frame, title)
    tb = _clone_at_end(slide, proto_tbl)
    t = tb.table
    need, cur = max(1, len(pairs)), len(t.rows)
    if need > cur:
        _add_rows(t, need - cur)
    elif need < cur:
        _del_rows(t, cur - need)
    for i, (a, b) in enumerate(pairs):
        if str(a).strip():
            _set(t, i, 0, a)
            _set(t, i, 1, b)
        else:
            # 왼칸이 없으면 두 칸을 합쳐 한 줄로(설명 문장).
            try:
                t.cell(i, 0).merge(t.cell(i, 1))
            except Exception:
                pass
            _set(t, i, 0, b)
    # 틀에서 물려받은 세로 정렬이 칸마다 달라 글이 바닥에 붙는다 → 가운데로 통일.
    for row in t.rows:
        for cell in row.cells:
            try:
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            except Exception:
                pass
    return lb, tb


def _units(s: str) -> float:
    """글자 폭을 em 단위로 센다. 한글·한자는 전각(1.0), 숫자·영문·기호는 반각(0.55)."""
    u = 0.0
    for ch in str(s):
        if ('가' <= ch <= '힣' or '㄰' <= ch <= '㆏'
                or '一' <= ch <= '鿿' or ch in '·／…‘’“”'):
            u += 1.0
        elif ch == ' ':
            u += 0.30
        else:
            u += 0.55
    return u


def _set_table_width(shape, target_w: int):
    """표를 실제로 target_w 폭에 맞춘다.

    ★graphicFrame 의 width 만 바꾸면 **열 너비(gridCol)는 그대로**라
      파워포인트는 원래 폭으로 그린다 — 표가 칸 밖으로 삐져나오거나 덜 차서
      좌·우 단의 표들이 서로 안 맞아 보이던 원인. 열 너비를 비율대로 함께 늘린다.
    """
    try:
        cols = list(shape.table.columns)
        cur = sum(c.width or 0 for c in cols)
    except Exception:
        shape.width = target_w
        return
    if cur <= 0 or not cols:
        shape.width = target_w
        return
    scaled = [int((c.width or 0) * target_w / cur) for c in cols]
    scaled[-1] += target_w - sum(scaled)      # 반올림 오차는 마지막 열이 흡수
    for c, w in zip(cols, scaled):
        if w > 0:
            c.width = w
    shape.width = target_w


def _fit_table_rows(shape):
    """행 높이를 글 양에 맞춘다.

    ★칸의 글이 길면 파워포인트가 행을 알아서 늘리는데, 우리가 계산한 높이는
      늘기 전 값이라 아래 블록을 덮었다. 열 너비 기준으로 줄 수를 미리 세어 둔다.
    """
    try:
        t = shape.table
        widths = [(c.width or 0) / 914400.0 for c in t.columns]
    except Exception:
        return
    for row in t.rows:
        need = 0.0
        for ci, cell in enumerate(row.cells):
            w = widths[ci] if ci < len(widths) else 1.0
            pt = 9.0
            for p in cell.text_frame.paragraphs:
                got = False
                for run in p.runs:
                    if run.font.size:
                        pt, got = run.font.size.pt, True
                        break
                if got:
                    break
            cpl = max(1.0, (w - 0.16) * 72.0 / pt)      # 좌우 여백 0.08in 씩
            lines = 0
            for ln in (cell.text or "").split("\n"):
                lines += max(1, int(-(-_units(ln) // cpl)))
            need = max(need, max(1, lines) * pt * 1.32 / 72.0 + 0.08)
        # ★'기존 높이와 큰 값'이 아니라 **필요한 만큼**으로 맞춘다.
        #   틀에서 물려받은 큰 행 높이를 그대로 두면 한 줄짜리 칸이 뻥 비어 보인다.
        #   최소 높이는 글자 크기를 따라간다(글자를 줄이면 행도 줄어야 한다).
        row.height = int(Inches(max(0.14, need)))


def _table_h(shape):
    """표의 실제 높이(행 높이 합)."""
    try:
        return sum(r.height for r in shape.table.rows)
    except Exception:
        return shape.height or 0


_MIN_TABLE_PT = 6.0      # 이보다 작으면 읽기 힘들다 — 여기서 멈추고 경고를 낸다


def _scale_table_font(shape, ratio: float, base_pt: float = 9.0):
    """표 안 글자 크기를 ratio 배로 줄인다(칸이 넘칠 때만)."""
    try:
        t = shape.table
    except Exception:
        return
    for row in t.rows:
        for cell in row.cells:
            for para in cell.text_frame.paragraphs:
                for run in para.runs:
                    cur = run.font.size.pt if run.font.size else base_pt
                    run.font.size = Pt(max(_MIN_TABLE_PT, round(cur * ratio, 1)))


def _stack(blocks, x_in, w_in, top_in, bottom_in, gap_in=0.16, label_gap_in=0.05):
    """한 단(column)의 블록들을 위에서부터 차곡차곡 다시 배치한다.

    ★레이아웃 틀은 위치가 '고정'이라, 표 행 수가 원본에 따라 늘면 아래 블록과 겹친다
      (연대보증사 재무제표가 법인개요를 덮던 문제). 내용을 채운 뒤 실제 높이로 재배치한다.
    blocks: [(label_shape|None, table_or_pic_shape|None), ...]  위→아래 순서
    반환: 마지막 y(인치). bottom 을 넘으면 넘친 것이므로 호출측에서 판단.
    """
    X, W = Inches(x_in), Inches(w_in)
    y = Inches(top_in)
    for label, body in blocks:
        if label is not None:
            # ★라벨 글상자는 틀에서 1.60in 로 좁다. 회사명이 붙어 길어지면 두 줄로
            #   접히는데 높이는 그대로라 바로 아래 표를 덮었다 → 단 폭만큼 넓힌다.
            label.left, label.top, label.width = X, y, W
            try:
                label.text_frame.word_wrap = False
            except Exception:
                pass
            y = y + (label.height or Inches(0.25)) + Inches(label_gap_in)
        if body is not None:
            body.left, body.top = X, y
            if body.has_table:
                _set_table_width(body, W)     # 열 너비까지 같이 맞춘다
                _fit_table_rows(body)         # 글 양에 맞춰 행 높이 확보
                if getattr(body, "name", "") == "블록:담보대출조건:본문":
                    _loan_fix(body)           # 3번째 줄 칸 안에 Tr 표를 얹는다
                h = _table_h(body)
            elif body.shape_type == 17 and body.has_text_frame:
                body.width = W                # 글 상자(Overview) — 글 양만큼 높이
                h = _text_box_h(body, W)
            else:
                # 그림·구조도는 폭만 바꾸면 찌그러진다 → 높이도 같은 비율로.
                _ow, _oh = body.width or 0, body.height or 0
                _ratio = (W / _ow) if _ow > 0 else 1.0
                body.width = W
                h = int(_oh * _ratio)
                # ★그룹을 줄여도 안쪽 '표'는 안 줄어든다(OOXML 특성) → 직접 줄인다.
                if getattr(body, "name", "") == "금융구조도":
                    try:
                        import diagram as _dgm
                        _dgm.scale_tables_in_group(body, _ratio)
                    except Exception as _se:
                        print(f"[요약본] 구조도 표 축소 실패: {_se}")
            body.height = h
            y = y + h + Inches(gap_in)
    return y / 914400.0


def _fix_cover_overlap(slide):
    """표지에서 제목이 길어 두 줄이 되면 아래 날짜 줄을 덮는 것을 막는다.

    틀의 제목 글상자는 0.34in(한 줄)뿐인데 딜명이 길면 두 줄이 된다.
    글상자는 커지지 않으므로 바로 아래 날짜 줄과 겹친다.
    → 줄 수만큼 제목 상자를 키우고, 그 아래 글상자를 같은 만큼 내린다.
    (그룹의 세로 배율이 1 이라 안쪽 좌표를 그대로 더하면 된다)
    """
    grp = next((sh for sh in slide.shapes if sh.shape_type == 6), None)
    if grp is None:
        return
    title = date_box = None
    for ch in grp.shapes:
        if not ch.has_text_frame:
            continue
        t = (ch.text_frame.text or "")
        if "제안서" in t or "요약본" in t:
            title = ch
        elif any(m in t for m in ("월", "20")) and len(t) < 30:
            date_box = ch
    if title is None:
        return

    lines = (title.text_frame.text or "").count(chr(10)) + 1
    if lines < 2:
        return
    one = title.height or Inches(0.34)
    extra = int(one * (lines - 1))
    title.height = one + extra
    if date_box is not None and (date_box.top or 0) > (title.top or 0):
        date_box.top = (date_box.top or 0) + extra

    # 그룹도 그만큼 키워야 아래 글상자가 그룹 밖으로 나가지 않는다
    try:
        from pptx.util import Emu as _E
        xf = grp._element.find(".//" + _A + "xfrm")
        ext = xf.find(_A + "ext"); che = xf.find(_A + "chExt")
        ext.set("cy", str(int(ext.get("cy")) + extra))
        che.set("cy", str(int(che.get("cy")) + extra))
    except Exception:
        pass


# ── 글머리기호·줄맞춤 통일 ──────────────────────────
#  틀의 칸들이 marL=0, indent=-342900 로 돼 있었다.
#  이러면 첫 줄 글자는 불릿 뒤에서 시작하는데 **줄바꿈된 둘째 줄은 왼쪽 끝(0)으로
#  돌아간다.** 그래서 칸마다 불릿 간격과 줄맞춤이 제각각으로 보였다.
#  제대로 된 내어쓰기는 marL=한 칸, indent=-한 칸 이다.
#    불릿 → 0 자리,  글자 → 한 칸 자리,  줄바꿈된 줄 → 같은 한 칸 자리
_HANG = 228600          # 0.25인치 — 불릿과 글자 사이 간격(모든 칸 동일)


def _norm_paragraph(para):
    pPr = para._p.get_or_add_pPr()
    has_bu = any(e.tag.endswith("}buChar") or e.tag.endswith("}buAutoNum")
                 for e in pPr)
    ind = pPr.get("indent")
    try:
        neg = int(ind) < 0 if ind is not None else False
    except Exception:
        neg = False
    if has_bu or neg:
        pPr.set("marL", str(_HANG))
        pPr.set("indent", str(-_HANG))
    else:
        pPr.set("marL", "0")
        pPr.set("indent", "0")


def normalize_bullets(slide):
    """이 슬라이드의 모든 표·글상자에서 불릿 간격과 줄맞춤을 같게 만든다."""
    for sh in _walk(slide.shapes):
        if sh.has_table:
            t = sh.table
            for r in range(len(t.rows)):
                for c in range(len(t.columns)):
                    for para in t.cell(r, c).text_frame.paragraphs:
                        _norm_paragraph(para)
        elif sh.has_text_frame:
            for para in sh.text_frame.paragraphs:
                _norm_paragraph(para)


def _border(shape, rgb="000000", pt=0.75):
    """그림에 얇은 검정 테두리(사용자 지시)."""
    from pptx.oxml.ns import qn
    spPr = shape._element.spPr
    for e in spPr.findall(qn("a:ln")):
        spPr.remove(e)
    ln = spPr.makeelement(qn("a:ln"), {"w": str(int(pt * 12700))})
    fill = ln.makeelement(qn("a:solidFill"), {})
    fill.append(fill.makeelement(qn("a:srgbClr"), {"val": rgb}))
    ln.append(fill)
    spPr.append(ln)


HL_SYS = """당신은 부동산 금융 IM 을 읽고 제안서 표지 다음 장에 넣을
'핵심 3가지' 카드를 만드는 전문가다. 아래는 원문의 Executive Summary 부분이다.

규칙:
- 원문에 있는 말만 쓴다. 지어내지 않는다. 숫자·단위·괄호는 원문 그대로.
- 카드는 정확히 3개.
- title: 12자 안쪽의 짧은 제목
- subtitle: 20자 안쪽 한 줄(없으면 "")
- bullets: 근거 2~3줄, 각 줄은 원문 문장을 거의 그대로

출력은 JSON 배열만:
[{"title":"..","subtitle":"..","bullets":["..",".."]}, ...]"""


def highlights_from_pages(pdf_bytes, pages):
    """지정한 **쪽만** 읽어 하이라이트 카드 3개를 만든다.

    전에는 처음에 문서 전체로 뽑아둔 결과를 그대로 다시 썼다. 그래서 사용자가
    Executive Summary 쪽 번호를 고쳐도 아무 것도 달라지지 않았다.
    """
    import fitz
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        txt = []
        for pno in pages or []:
            if 1 <= pno <= doc.page_count:
                txt.append(str(pno) + "쪽" + chr(10) + (doc[pno - 1].get_text() or ""))
    finally:
        doc.close()
    body = (chr(10) + chr(10)).join(txt).strip()
    if not body:
        return []
    res = call_claude(HL_SYS, body, slide_num=802, pdf_context=body,
                      prompt_version="highlights_v1")
    if not res.get("ok"):
        raise RuntimeError(res.get("error") or "하이라이트 추출 실패")
    data = res["data"]
    if isinstance(data, dict):
        data = data.get("highlights") or data.get("cards") or []
    out = []
    for h in (data or [])[:3]:
        if not isinstance(h, dict):
            continue
        out.append({"title": str(h.get("title") or "")[:40],
                    "subtitle": str(h.get("subtitle") or "")[:60],
                    "bullets": [str(b) for b in (h.get("bullets") or []) if b]})
    return out


def build_highlight_preview(data: dict, out_path: str) -> str:
    """하이라이트 슬라이드 1장만 만든 PPTX(미리보기용)."""
    return build_summary(data, None, out_path, pages=2, _highlight_only=True)


# ── 하이라이트 제목 블록([체크아이콘]+[제목]) 가운데 정렬 ───────────────
#   문제: 레이아웃 틀의 제목 그룹 3개가 전부 L4.66" W1.52" 로 **고정**이고
#   그 안 글상자는 폭 1.04" 짜리다. 제목이 그보다 길면 글이 상자를 뚫고
#   오른쪽으로만 자라서, **글자 수가 다르면 카드마다 제목 위치가 어긋났다.**
#   (예: '초역세권 입지' 6자 vs '낮은 인허가 리스크' 9자)
#   해결: 제목 글자폭을 재서 그룹 폭을 실제 내용에 맞추고, 그룹 전체를
#   슬라이드 정중앙에 놓는다. 기존 IM 변환기의 하이라이트와 같은 방식.
_TITLE_PT = 18.0    # 틀의 제목 rPr 에 sz 가 없다 → 본문 기본값 18pt
_ICON_GAP = 0.08    # 체크아이콘과 제목 사이 간격(인치)
_A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"


def _text_width_in(s: str, pt: float) -> float:
    """글자폭 추정(인치)."""
    return _units(s) * pt / 72.0


def _center_title_group(grp, txbox, title: str, slide_w: int) -> bool:
    """[체크아이콘 + 제목] 그룹을 글자폭에 맞춰 다시 재고 슬라이드 가운데로 옮긴다."""
    xfrm = grp._element.find(f"{_A}xfrm")
    if xfrm is None:                       # grpSpPr 안에 있는 경우
        xfrm = grp._element.find(f".//{_A}xfrm")
    if xfrm is None:
        return False
    off, ext = xfrm.find(f"{_A}off"), xfrm.find(f"{_A}ext")
    ch_off, ch_ext = xfrm.find(f"{_A}chOff"), xfrm.find(f"{_A}chExt")
    if None in (off, ext, ch_off, ch_ext):
        return False

    pics = [c for c in grp.shapes if c.shape_type == 13]
    if not pics:
        return False
    icon = pics[0]
    icon_w = icon.width or Inches(0.47)

    base_x = int(ch_off.get("x"))          # 자식 좌표계의 왼쪽 끝
    text_w = int(Inches(_text_width_in(title or "", _TITLE_PT) + 0.06))
    text_w = max(text_w, Inches(0.4))

    # 자식 좌표계 재배치: [아이콘][간격][제목]
    icon.left = base_x
    txbox.left = base_x + icon_w + Inches(_ICON_GAP)
    txbox.width = text_w

    total = icon_w + Inches(_ICON_GAP) + text_w
    ext.set("cx", str(int(total)))          # 확대·축소 없이 1:1 이 되도록
    ch_ext.set("cx", str(int(total)))
    off.set("x", str(int((slide_w - total) / 2)))   # 슬라이드 정중앙
    return True


def _tbl_add_col(table):
    """표 맨 오른쪽 앞(마지막 열 바로 앞)에 열 하나 — 마지막 열(합계) 서식을 복제."""
    from pptx.oxml.ns import qn
    tbl = table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    cols = grid.findall(qn("a:gridCol"))
    src = cols[-2] if len(cols) >= 2 else cols[-1]
    new = deepcopy(src)
    cols[-1].addprevious(new)
    for tr in tbl.tr_lst:
        tcs = tr.findall(qn("a:tc"))
        ref = tcs[-2] if len(tcs) >= 2 else tcs[-1]
        tcs[-1].addprevious(deepcopy(ref))


def _tbl_del_col(table, idx):
    from pptx.oxml.ns import qn
    tbl = table._tbl
    grid = tbl.find(qn("a:tblGrid"))
    grid.remove(grid.findall(qn("a:gridCol"))[idx])
    for tr in tbl.tr_lst:
        tr.remove(tr.findall(qn("a:tc"))[idx])


def _num(v):
    try:
        return float(re.sub(r"[^\d.\-]", "", str(v)))
    except Exception:
        return None


def _set_font_light(tf):
    """글 상자 글자를 얇은 회사 글꼴로(2026-10-19 사용자: 두꺼운 = Pretendard ExtraBold, 얇은 = Pretendard Light)."""
    from pptx.oxml.ns import qn
    for p in tf.paragraphs:
        for r in p.runs:
            rpr = r._r.get_or_add_rPr()
            for tag in ("a:latin", "a:ea", "a:cs"):
                e = rpr.find(qn(tag))
                if e is None:
                    e = rpr.makeelement(qn(tag), {})
                    rpr.append(e)
                e.set("typeface", "Pretendard Light")
            rpr.attrib.pop("b", None)


def _text_box_h(shape, w_emu):
    """글 상자 높이 어림 — 단 폭에서 줄 수를 세고 글자 크기 × 1.5 줄 간격."""
    pt = 9.0
    for p in shape.text_frame.paragraphs:
        for r in p.runs:
            if r.font.size:
                pt = r.font.size.pt
                break
        else:
            continue
        break
    cpl = max(1.0, ((w_emu / 914400.0) - 0.45) * 72.0 / pt)
    lines = 0
    for ln in (shape.text_frame.text or "").split("\n"):
        lines += max(1, int(-(-_units(ln) // cpl)))
    return int(Inches(lines * pt * 1.62 / 72.0 + 0.22))


def _loan_fix(body):
    """'기초자산 담보대출 주요 조건' 표: 3번째 줄(대출금액 및 조건) 칸 안에 Tr 표를 얹고 그만큼 줄을 키운다."""
    tr = next((sh for sh in body._parent if getattr(sh, "name", "") == "블록:담보대출조건:부속"), None)
    if tr is None:
        return
    t = body.table
    c0 = t.columns[0].width
    c1 = sum(c.width for c in list(t.columns)[1:])
    _set_table_width(tr, int(c1 - Inches(0.16)))
    tr.left = body.left + c0 + Inches(0.08)
    tr.top = body.top + t.rows[0].height + t.rows[1].height + Inches(0.06)
    note = (t.cell(2, 1).text or "").strip()
    note_h = Inches(0.2) * max(1, note.count("\n") + 1) if note else 0
    t.rows[2].height = int(_table_h(tr) + Inches(0.12) + note_h)


def build_summary(data: dict, pdf_path: str, out_path: str, pages: int = 2,
                  _highlight_only: bool = False) -> str:
    """요약본 PPT. 2026-10-19 새 레이아웃: [표지][Deal Highlight][내용 1][내용 2]([내용 3])[End].
    pages = 내용 쪽 수(2 = 총 5장, 3 = 총 6장)."""
    prs = Presentation(_layout_path())

    # 슬라이드 1 : 표지
    s1 = prs.slides[0]
    name = (data.get('deal_name') or '').strip()
    _replace_contains(s1, "사업명(요약본)", name + chr(10) + "사모사채 제안서(요약본)")
    _now = _dt.datetime.now()
    _replace_contains(s1, "몇월", f"{_now.year}. {_now.month:02d}")
    _fix_cover_overlap(s1)

    # 슬라이드 2 : Deal Highlight — 카드 3개 = [<제목>][대괄호][본문] 그룹
    s2 = prs.slides[1]
    hl = data.get('highlights') or []
    groups = sorted([sh for sh in s2.shapes if sh.shape_type == 6
                     and not any(c.shape_type == 13 for c in sh.shapes)], key=lambda g: g.top)
    for i, g in enumerate(groups[:3]):
        tbs = [c for c in g.shapes if c.shape_type == 17 and c.has_text_frame]
        tt = next((c for c in tbs if c.text_frame.text.strip().startswith("<")), None)
        body = max([c for c in tbs if c is not tt], key=lambda c: c.width or 0, default=None)
        card = hl[i] if i < len(hl) else {}
        title = (card.get('title') or '').strip().strip("<>").strip()
        sub = (card.get('subtitle') or '').strip() if card.get('use_sub', True) else ''
        bul = []
        for _b in (card.get('bullets') or []):
            _b = str(_b).strip()
            while _b[:1] in ("→", "·", "-", "•", "▶", "ㆍ"):
                _b = _b[1:].strip()
            if _b:
                bul.append(_b)
        if tt is not None:
            _replace_text_keep_runs(tt.text_frame, f"<{title}>" if title else "")
            tt.text_frame.word_wrap = False
        if body is not None:
            _replace_text_keep_runs(body.text_frame, "\n".join(([sub] if sub else []) + bul))
            body.text_frame.word_wrap = True
            _set_font_light(body.text_frame)

    if _highlight_only:
        _ids = list(prs.slides._sldIdLst)
        for _i, _sid in enumerate(_ids):
            if _i != 1:
                prs.slides._sldIdLst.remove(_sid)
        prs.save(out_path)
        return out_path

    # 내용 쪽 : 틀의 3·4번 슬라이드를 '부품 상자' 로 쓴다 → 3번(s3)에 부품을 다 모으고 4번은 뺀다
    s3, s4 = prs.slides[2], prs.slides[3]
    E = 914400

    def _tables(slide, nrows=None, ncols=None):
        return [sh for sh in slide.shapes if sh.has_table
                and (nrows is None or len(sh.table.rows) == nrows)
                and (ncols is None or len(sh.table.columns) == ncols)]

    def _name(sh, nm):
        if sh is not None:
            sh.name = nm
        return sh

    # 4번에서 '사업개요'(라벨 + 2열 표)·'사업일정'(라벨 + 3열 표)을 3번으로 옮겨 온다
    for lbl_txt, pick in (("사업개요", lambda: next(iter(_tables(s4, ncols=2)), None)),
                          ("사업일정", lambda: next(iter(_tables(s4, ncols=3)), None))):
        lb, tb = _find_label(s4, lbl_txt), pick()
        for sh, kind in ((lb, "라벨"), (tb, "본문")):
            if sh is not None:
                el = deepcopy(sh._element)
                s3.shapes._spTree.append(el)
                s3.shapes[-1].name = f"블록:{lbl_txt}:{kind}"

    # 3번의 부품 이름표
    _name(_find_label(s3, "Overview"), "블록:Overview:라벨")
    ov_box = max([sh for sh in s3.shapes if sh.shape_type == 17 and sh.has_text_frame
                  and len(sh.text_frame.text) > 40], key=lambda sh: len(sh.text_frame.text), default=None)
    _name(ov_box, "블록:Overview:본문")
    _name(_find_label(s3, "사모사채 개요"), "블록:사모사채개요:라벨")
    T_BOND = _name(next(iter(_tables(s3, nrows=9, ncols=2)), None), "블록:사모사채개요:본문")
    _name(_find_label(s3, "담보대출"), "블록:담보대출조건:라벨")
    T_LOAN = _name(next(iter(_tables(s3, nrows=5, ncols=2)), None), "블록:담보대출조건:본문")
    T_TR = _name(next(iter(_tables(s3, nrows=4)), None), "블록:담보대출조건:부속")
    # 다른 딜 예시(담보토지 매입 현황 글·표·사진, 구조도 예시 그림)는 지운다
    for sh in list(s3.shapes):
        nm = sh.name or ""
        if nm.startswith("블록:"):
            continue
        txt = sh.text_frame.text if sh.has_text_frame else ""
        if ("담보토지" in txt) or (sh.has_table) or (sh.shape_type == 13 and (sh.width or 0) > Inches(1.8)):
            sh._element.getparent().remove(sh._element)

    loan = data.get('담보대출조건') or {}
    bong = next((t for t in (loan.get('tranches') or []) if t.get('is_bongeon')), {})
    기간 = (loan.get('대출기간') or '').strip()

    # Overview 글
    if ov_box is not None:
        _ov = data.get("overview") or []
        if isinstance(_ov, str):
            _ov = [x for x in _ov.split("\n")]
        _ov = [str(x).replace("**", "").strip().lstrip("→·-•▶ㆍ ").strip() for x in _ov if str(x).strip()]
        _replace_text_keep_runs(ov_box.text_frame, "\n".join(_ov))
        ov_box.text_frame.word_wrap = True
        _set_font_light(ov_box.text_frame)

    # 사모사채 개요(고정 + 파생) — 구분열은 그대로, 값열만
    if T_BOND is not None:
        tb = T_BOND.table
        bo = data.get("사모사채개요") or {}
        vals = {3: bo.get("기초자산") or f"{name} 담보대출채권", 4: bo.get("발행금액") or bong.get('대출금액'),
                5: bo.get("발행일") or f"{data.get('date_ko', '')} (예정)",
                6: bo.get("만기일") or f"발행일로부터 {기간 or '약 3개월'} (대출 만기 연동)",
                7: bo.get("금융조건") or (f"채권 금리 : 연 {bong.get('금리', '')} (고정, 세전)" if bong.get('금리') else None)}
        for r, v in vals.items():
            if v and r < len(tb.rows):
                _set(tb, r, 1, v)

    # 기초자산 담보대출 주요 조건 + Tr 표
    if T_LOAN is not None:
        tl = T_LOAN.table
        _set(tl, 0, 1, loan.get('금융구조'))
        _set(tl, 1, 1, loan.get('차주'))
        _set(tl, 2, 1, (f"※ {loan.get('LTV기준_각주')}" if loan.get('LTV기준_각주') else ""))
        try:
            tl.cell(2, 1).vertical_anchor = MSO_ANCHOR.BOTTOM
        except Exception:
            pass
        _lines(tl, 3, 1, loan.get('주요채권보전') or [])
        _lines(tl, 4, 1, loan.get('주요인출선행조건') or ([f"대출기간 : {기간}"] if 기간 else []))
    if T_TR is not None:
        tt = T_TR.table
        trs = loan.get('tranches') or []
        # ★틀의 Tr.B 열(본건)은 빨간 테두리 — 열 수를 바꾸기 전에 칸 서식(보통/본건)을 떠 둔다
        from pptx.oxml.ns import qn as _qn
        _tcpr = lambda r, c: deepcopy(tt.cell(r, c)._tc.find(_qn("a:tcPr")))
        _plain = [_tcpr(r, 1) for r in range(len(tt.rows))]
        _red = [_tcpr(r, min(2, len(tt.columns) - 2)) for r in range(len(tt.rows))]
        want = max(1, len(trs)) + 2                         # 구분 + Tr들 + 합계
        while len(tt.columns) < want:
            _tbl_add_col(tt)
        while len(tt.columns) > want:
            _tbl_del_col(tt, len(tt.columns) - 2)
        nc = len(tt.columns)
        _set(tt, 0, 0, "구분")
        _set(tt, 1, 0, "대출금액(억원)")
        _set(tt, 2, 0, "대출기간(월)")
        _set(tt, 3, 0, "LTV")
        tot = 0.0
        ok = bool(trs)
        for c in range(1, nc - 1):
            v = trs[c - 1] if c - 1 < len(trs) else {}
            _g = (v.get('구분') or '')
            _set(tt, 0, c, _g + (" (본건)" if v.get('is_bongeon') and "본건" not in _g else ""))
            _set(tt, 1, c, v.get('대출금액') or "")
            _set(tt, 2, c, v.get('대출기간') or 기간 or "")
            _set(tt, 3, c, v.get('LTV') or "-")
            n_ = _num(v.get('대출금액'))
            if n_ is None:
                ok = False
            else:
                tot += n_
        for c in range(1, nc - 1):                          # 본건 열만 빨간 테두리
            v = trs[c - 1] if c - 1 < len(trs) else {}
            for r in range(len(tt.rows)):
                tc = tt.cell(r, c)._tc
                old = tc.find(_qn("a:tcPr"))
                src = (_red if v.get('is_bongeon') else _plain)[r]
                if src is not None:
                    if old is not None:
                        tc.remove(old)
                    tc.append(deepcopy(src))
        _set(tt, 0, nc - 1, "합계")
        _set(tt, 1, nc - 1, (f"{tot:,.0f}" if ok else "-"))
        _set(tt, 2, nc - 1, "-")
        _set(tt, 3, nc - 1, "")

    # 사업개요(2열 표) / 사업일정(3열 표)
    lbl_ov2 = next((sh for sh in s3.shapes if sh.name == "블록:사업개요:라벨"), None)
    T_BIZ = next((sh for sh in s3.shapes if sh.name == "블록:사업개요:본문"), None)
    if T_BIZ is not None:
        biz = data.get("사업개요") or {}
        pairs = [(k, v) for k, v in (biz.items() if isinstance(biz, dict) else []) if v]
        t = T_BIZ.table
        need, cur = max(1, len(pairs)), len(t.rows)
        if need > cur:
            _add_rows(t, need - cur)
        elif need < cur:
            _del_rows(t, cur - need)
        for i, (a, b) in enumerate(pairs):
            _set(t, i, 0, a)
            _set(t, i, 1, b)
    T_SCHED = next((sh for sh in s3.shapes if sh.name == "블록:사업일정:본문"), None)
    if T_SCHED is not None:
        ts = T_SCHED.table
        sched = data.get('사업일정') or {}
        done = sched.get('완료') or []
        plan = sched.get('예정') or []
        allrows = [("완료", d) for d in done] + [("예정", p) for p in plan]
        need, cur = max(1, len(allrows)), len(ts.rows)
        if need > cur:
            _add_rows(ts, need - cur)
        elif need < cur:
            _del_rows(ts, cur - need)
        for i, (grp, it) in enumerate(allrows):
            _set(ts, i, 0, grp if (i == 0 or i == len(done)) else "")
            _set(ts, i, 1, it.get('시기'))
            _set(ts, i, 2, it.get('내용'))

    # 조감도(원본 사진)
    try:
        import fitz
        doc = fitz.open(pdf_path)
        skip = repeated_xrefs(doc)
        imgs = []
        for pi in range(doc.page_count):
            imgs += extract_page_images(doc, pi, skip_xrefs=skip)
        imgs = [im for im in imgs if im.get('data')]
        for im in imgs:
            ar = im['width'] / max(1, im['height'])
            im['_score'] = (1 if mean_brightness(im['data']) < 200 else 0,
                            1 if 1.2 <= ar <= 2.6 else 0, im['width'] * im['height'])
        imgs.sort(key=lambda im: im['_score'], reverse=True)
        if imgs:
            _p = s3.shapes.add_picture(io.BytesIO(imgs[0]['data']), Inches(0.25), Inches(0.39),
                                       width=Inches(4.88))
            _p.name = "블록:조감도:본문"
        doc.close()
    except Exception as _ie:
        print(f"[요약본] 조감도 삽입 실패: {_ie}")

    # 쪽마다 고른 칸 배치
    DEFAULTS = [
        {"left": ["Overview", "사모사채개요", "금융구조도"], "right": ["담보대출조건"]},
        {"left": ["사업개요", "조감도"], "right": ["사업일정"]},
        {"left": ["법인개요"], "right": ["재무제표"]},
    ]
    pages = max(1, int(pages or 2))
    sp = list(data.get("_slots_pages") or [])
    slots_pages = [dict(sp[k]) if k < len(sp) and sp[k] else DEFAULTS[min(k, len(DEFAULTS) - 1)]
                   for k in range(pages)]
    _wanted = {k for _s in slots_pages for _side in ("left", "right") for k in (_s.get(_side) or [])}

    # 그 밖의 내용(법인개요·재무제표·중요항목 등)은 [라벨 + 2열 표] 한 벌씩
    _lbl_proto = next((sh for sh in s3.shapes if sh.name == "블록:사업일정:라벨"), None)
    _tbl_proto = T_BIZ
    extras = list(extra_blocks_of(data))
    corp = data.get('법인개요') or {}
    if corp:
        _cn = (corp.get('회사명') or '').strip()
        extras.append(("법인개요", [(k, corp.get(k)) for k in
                                    ("회사명", "설립일", "사업자번호", "대표자", "자본금", "주요사업", "주요주주")
                                    if corp.get(k)]))
    fin = data.get('재무제표') or {}
    if fin and fin.get('행'):
        yrs = fin.get('연도') or []
        rows = [(f"구분(단위 : {fin.get('단위', '백만원')})", " / ".join(str(y) for y in yrs))]
        for row in fin.get('행') or []:
            rows.append((row.get('항목') or "", " / ".join(str(x) for x in (row.get('값') or []))))
        extras.append(("재무제표", rows))
    for _title, _pairs in extras:
        if _title == "사업개요" or (_wanted and _title not in _wanted) or _tbl_proto is None:
            continue
        lb, tb = _build_extra_block(s3, _tbl_proto, _lbl_proto, _title, _pairs)
        if lb is not None:
            lb.name = f"블록:{_title}:라벨"
        tb.name = f"블록:{_title}:본문"

    # 원본 IM 의 표를 그대로
    _orig = data.get("_orig_tables") or {}
    if _orig:
        try:
            import ppt_tables as _pt
            for _nm, _t in _orig.items():
                if _wanted and _nm not in _wanted:
                    continue
                if _lbl_proto is not None:
                    _lb = _clone_at_end(s3, _lbl_proto)
                    _replace_text_keep_runs(_lb.text_frame, _t.get("title") or _nm)
                    _lb.name = f"블록:{_nm}:라벨"
                _sh = _pt.add_table(s3, _t, Inches(0.25), Inches(0.39), Inches(4.90))
                if _sh is None:
                    _sh = next((sh for sh in reversed(list(s3.shapes)) if sh.has_table), None)
                if _sh is not None:
                    _sh.name = f"블록:{_nm}:본문"
        except Exception as _oe:
            print(f"[요약본] 원본 표 그리기 실패: {_oe}")

    # 금융구조도(3단계에서 만든 것) — 그림으로
    _diag = data.get("_diagram") or {}
    _diag_png = _diag.get("png")
    if _diag.get("pptx_bytes") and not (_diag.get("pptx") and os.path.exists(_diag["pptx"])):
        try:
            import tempfile as _tf
            _tmp = _tf.NamedTemporaryFile(suffix=".pptx", delete=False)
            _tmp.write(_diag["pptx_bytes"])
            _tmp.close()
            _diag["pptx"] = _tmp.name
        except Exception as _te:
            print(f"[요약본] 구조도 임시파일 만들기 실패: {_te}")
    if not _diag_png and _diag.get("pptx") and os.path.exists(_diag["pptx"]):
        try:
            from engine_bits import pptx_slide_png as _p2p
            import diagram as _dgm
            _clean = _dgm.strip_chrome(_diag["pptx"])
            _diag_png, _err = _p2p(_clean, 1)
            if _diag_png:
                _diag_png = _dgm.crop_to_content(_diag_png, _dgm.content_bbox(_clean))
            else:
                print(f"[요약본] 구조도를 그림으로 못 바꿈: {_err}")
        except Exception as _de:
            print(f"[요약본] 구조도 그림 변환 실패: {_de}")
    if _diag_png:
        try:
            _ds = s3.shapes.add_picture(io.BytesIO(_diag_png), Inches(0.25), Inches(0.39), width=Inches(4.90))
            _ds.name = "블록:금융구조도:본문"
            _border(_ds)
        except Exception as _de:
            print(f"[요약본] 금융구조도 넣기 실패: {_de}")
    elif "금융구조도" in _wanted:
        print("[요약본] 경고: 금융구조도를 골랐는데 넣을 구조도가 없습니다 "
              "— 3단계에서 구조도를 골라 항목을 채워 주세요.")

    def _blocks_of(slide):
        m = {}
        for sh in slide.shapes:
            nm = sh.name or ""
            if nm.startswith("블록:"):
                _, key, kind = nm.split(":", 2)
                lb, bd = m.get(key, (None, None))
                if kind == "라벨":
                    lb = sh
                elif kind == "본문":
                    bd = sh
                m[key] = (lb, bd)
        return m

    def _arrange(slide, sl, tag):
        BL = _blocks_of(slide)
        used = {k for side in ("left", "right") for k in sl.get(side, [])}
        for key, (lb, bd) in BL.items():
            if key in used:
                continue
            for sh in (lb, bd):
                if sh is not None:
                    sh._element.getparent().remove(sh._element)
            if key == "담보대출조건":
                for sh in list(slide.shapes):
                    if sh.name == "블록:담보대출조건:부속":
                        sh._element.getparent().remove(sh._element)
        _BOT, _TOP = 7.28, 0.39
        for side, x, w, nm in (("left", 0.25, 4.90, "좌단"), ("right", 5.26, 5.32, "우단")):
            seq = [BL[k] for k in sl.get(side, []) if k in BL]
            seq = [(lb, bd) for lb, bd in seq if bd is not None or lb is not None]
            end = _stack(seq, x_in=x, w_in=w, top_in=_TOP, bottom_in=_BOT)
            if end > _BOT:
                print(f"[요약본] 경고: {tag} {nm} 내용이 {end - _BOT:.2f}in 넘침 — 칸을 줄이거나 쪽을 늘려 주세요")

    # 내용 쪽 복제(채운 뒤 통째로) → 쪽마다 고른 것만 남긴다
    copies = [s3]
    for k in range(1, pages):
        copies.append(_dup_slide_after(prs, s3, 2 + k - 1))
    for k, sl in enumerate(copies):
        _arrange(sl, slots_pages[k], f"내용 {k + 1}/{pages}")
    # 4번 틀 쪽(부품만 빌려 온 것)은 결과에서 뺀다 — 관계까지 끊어야 파일 이름이 안 겹친다
    for _sid in list(prs.slides._sldIdLst):
        if prs.part.related_part(_sid.rId) is s4.part:
            prs.part.drop_rel(_sid.rId)
            prs.slides._sldIdLst.remove(_sid)

    for _k, _sl in enumerate(prs.slides):
        if _k != 1:                                         # Deal Highlight 은 틀의 → 간격 그대로
            normalize_bullets(_sl)
    prs.save(out_path)
    return out_path
