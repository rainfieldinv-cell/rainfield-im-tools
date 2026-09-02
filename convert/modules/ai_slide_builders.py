"""
ai_slide_builders.py
─────────────────────────────────────────────────────────
Claude API 기반 슬라이드 데이터 생성 + PPT 빌더.

이 파일은 modules/page_builders.py 를 수정하지 않고
AI 처리가 필요한 슬라이드만 별도로 처리합니다.

현재 구현된 슬라이드 (그룹 A):
  - 슬라이드 2: Executive Summary
  - 슬라이드 5: 1.1 본건 사모사채 개요
  - 슬라이드 7: 2.1 투자구조도
─────────────────────────────────────────────────────────
"""

import copy

from lxml import etree
from pptx.dml.color import RGBColor
from pptx.oxml.ns import qn as _qn
from pptx.util import Inches, Pt, Cm
from pptx.enum.shapes import MSO_SHAPE, MSO_CONNECTOR
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

from modules.claude_api import call_claude, verify_numbers_in_pdf
from modules.page_builders import (
    clone_slide_layout,
    _replace_text_frame_content as _replace_tf_content,
    _replace_footer_business_name as _replace_footer,
)

# 슬라이드 크기 (cm)
_SLIDE_W = 27.517
_SLIDE_H = 19.05

# Rainfield 다크 네이비 + 포인트 그린
_C_DARK   = RGBColor(0x00, 0x20, 0x60)   # 002060
_C_GREEN  = RGBColor(0x70, 0xAD, 0x47)   # 70AD47
_C_GRAY   = RGBColor(0x7F, 0x7F, 0x7F)   # 7F7F7F
_C_WHITE  = RGBColor(0xFF, 0xFF, 0xFF)
_C_LIGHT  = RGBColor(0xF2, 0xF2, 0xF2)   # F2F2F2

# ════════════════════════════════════════════════════════
# 신도림 팔레트 (전체 PPT 공통 — 앞으로 모든 슬라이드에서 이 색만 사용)
# ════════════════════════════════════════════════════════
PALETTE = {
    "navy_dark": RGBColor(0x08, 0x37, 0x7C),  # 08377C 진남색 (표 헤더 주력 / 차주)
    "blue":      RGBColor(0x00, 0x63, 0xA1),  # 0063A1 파랑 (신탁사·시공사·AMC)
    "steel":     RGBColor(0x3E, 0x95, 0xBE),  # 3E95BE 스틸블루 (강조, 보통 투명 35%)
    "maroon":    RGBColor(0x8C, 0x4A, 0x59),  # 8C4A59 자주 (대주)
    "gray":      RGBColor(0xD9, 0xD9, 0xD9),  # D9D9D9 회색
    "label_gray": RGBColor(0xF2, 0xF2, 0xF2),  # F2F2F2 구분열 (흰색 배경1, 5% 더 어둡게)
    "red":       RGBColor(0xC0, 0x00, 0x00),  # C00000 빨강 강조
    "gray_text": RGBColor(0x80, 0x80, 0x80),  # 808080 푸터/보조 회색
    "white":     RGBColor(0xFF, 0xFF, 0xFF),
    "black":     RGBColor(0x00, 0x00, 0x00),
}

# ════════════════════════════════════════════════════════
# 슬라이드 2: Executive Summary
# ════════════════════════════════════════════════════════

SLIDE_2_SYSTEM_PROMPT = """당신은 부동산 PF 투자제안서 작성 전문가입니다.
주어진 PDF 의 'Executive Summary' 페이지 내용(여러 개의 항목/문장)을 읽고, 그 내용을 정확히
'3개 섹션'으로 요약·재구성합니다.

[엄격한 규칙]
1. ★원문에 실제로 있는 내용만 사용하라. 원문에 없는 항목(예: 분양성, 만기일자 등)은 절대
   만들지 마라. "확인 불가" 같은 문구도 쓰지 마라 — 없으면 그 주제를 아예 다루지 마라.
2. 원문 항목들을 의미가 비슷한 것끼리 묶어 정확히 3개 섹션으로 나눈다.
3. 각 섹션은 title(내용에 맞는 8자 이내 제목)과 body(핵심 요약)로 구성.
   - 'Executive Summary' 라는 단어는 title 에 쓰지 마라.
   - title 예: "거래 개요", "사업 구조", "인허가·시공", "핵심 포인트" 등 — 내용에 맞게.
4. body 는 원문을 그대로 복붙하지 말고 요약하라. 한 섹션에 포인트가 여러 개면 '\\n' 로 줄을 나눠라
   (각 줄이 한 포인트). 숫자/사실은 원문 값 그대로.
5. 보통 1번 섹션=거래 개요(차주·사업지·금액·구조), 2번=핵심 포인트(인허가/시공/토지 등 리스크),
   3번=나머지 핵심(사업 구조·책임준공 등). 단 원문 내용에 맞춰 유연하게.
6. 한국어. 출력은 JSON 만.
7. ★입력 텍스트에는 'Executive Summary'가 아닌 다른 내용(그 페이지의 나머지 절반, 표·상세
   금융조건·주석 등)이 섞여 있을 수 있다. 오직 'Executive Summary(핵심요약/투자포인트)'에
   해당하는 서술만 보고 정리하라. 그 외 상세표·부수 내용은 요약에 넣지 말고 무시하라.

[JSON 출력 스키마]
{
  "deal_title": "본 건 거래 한 줄 요약(사업명+토지담보대출 등)",
  "sections": [
    {"title": "8자 이내 제목", "body": "요약(여러 포인트면 \\n 로 구분)"},
    {"title": "8자 이내 제목", "body": "요약"},
    {"title": "8자 이내 제목", "body": "요약"}
  ]
}"""

SLIDE_2_USER_TEMPLATE = """[PDF 원문 — Executive Summary 페이지]
{pdf_text}

위 텍스트 중 'Executive Summary(핵심요약)'에 해당하는 부분만 보고 '3개 섹션'으로 요약해
JSON 으로 출력하라. (다른 상세표·부수 내용이 섞여 있으면 무시)
원문에 없는 내용은 만들지 말 것(특히 분양성·만기일자). sections 는 정확히 3개."""


def generate_executive_summary(pdf_text: str) -> dict:
    """
    슬라이드 2 (Executive Summary) 데이터를 Claude 로 생성합니다.

    Returns
    -------
    call_claude() 반환값 dict
    {"ok": bool, "data": {...}, "usage": {...}, "cached": bool, ...}
    """
    print("=" * 60)
    print(f"[FORCE-DEBUG-AI] generate_executive_summary 호출됨 - PDF 길이={len(pdf_text)}")
    print("=" * 60)

    result = call_claude(
        system_prompt=SLIDE_2_SYSTEM_PROMPT,
        user_prompt=SLIDE_2_USER_TEMPLATE.format(pdf_text=pdf_text),
        slide_num=2,
        pdf_context=pdf_text,
        prompt_version="exec_v3_es_only",
    )

    if result["ok"]:
        v = verify_numbers_in_pdf(result["data"], pdf_text)
        if not v["ok"]:
            print(f"[경고] 슬라이드 2 환각 의심 숫자: {v['hallucinated_numbers']}")
        else:
            print(f"[슬라이드 2] 숫자 검증 통과 ({v['verified_count']}개)")

    return result


# ════════════════════════════════════════════════════════
# 목차(TOC) 자동 추출 — IM 전체를 읽어 섹션/소제목 구성
# ════════════════════════════════════════════════════════

TOC_SYSTEM_PROMPT = """당신은 부동산·금융 IM(투자설명서)을 회사 표준 제안서 목차로 정리하는 전문가입니다.
주어진 IM 원문 전체를 읽고, 표준 4개 대분류에 맞춰 목차(소제목 포함)를 구성합니다.

[표준 대분류 — 제목은 이대로 유지]
01 사모사채 개요
02 금융개요
03 사업개요
04 Appendix

[규칙]
1. 각 대분류의 '소제목'을 IM 내용에서 찾아 채운다(그 섹션에 실제로 있는 내용만).
2. 소제목은 짧은 명사구(예: "사업개요", "입지분석", "분양사례", "차주 개요", "담보 분석").
   서술형 문장 금지. 각 대분류당 1~6개.
3. '01 사모사채 개요'의 소제목은 항상 ["본건 사모사채 개요"] 하나로 둔다(고정 페이지).
4. '02 금융개요'의 첫 소제목은 항상 "금융 구조도"로 시작한다. 이어서 IM의 금융조건 관련 소제목.
5. 소제목엔 번호(1.1 등)를 붙이지 말 것 — 텍스트만.
6. 출력은 JSON 만. 다른 텍스트 금지.

[JSON 출력 스키마]
{
  "sections": [
    {"title": "사모사채 개요", "subtitles": ["본건 사모사채 개요"]},
    {"title": "금융개요",     "subtitles": ["금융 구조도", "..."]},
    {"title": "사업개요",     "subtitles": ["...", "..."]},
    {"title": "Appendix",     "subtitles": ["..."]}
  ]
}"""

TOC_USER_TEMPLATE = """[IM 원문 전체]
{pdf_text}

위 IM을 읽고 표준 4개 대분류(사모사채 개요/금융개요/사업개요/Appendix)의 소제목을 구성해 JSON으로 출력하라."""


def generate_toc(pdf_text: str) -> dict:
    """IM 원문 전체를 읽어 목차(대분류+소제목)를 Claude 로 구성합니다.
       반환: call_claude() dict. data = {"sections": [{"title","subtitles":[...]}, ...]}."""
    print("=" * 60)
    print(f"[FORCE-DEBUG-AI] generate_toc 호출됨 - PDF 길이={len(pdf_text)}")
    print("=" * 60)
    return call_claude(
        system_prompt=TOC_SYSTEM_PROMPT,
        user_prompt=TOC_USER_TEMPLATE.format(pdf_text=pdf_text),
        slide_num=0,
        pdf_context=pdf_text,
        prompt_version="toc_v1",
    )


# ════════════════════════════════════════════════════════
# 슬라이드 5: 1.1 본건 사모사채 개요
# ════════════════════════════════════════════════════════

SLIDE_5_SYSTEM_PROMPT = """당신은 사모사채 발행 전문가입니다.
주어진 PDF 의 Bridge Loan 정보를 바탕으로, 그 대출의 Tr.B 를 기초자산으로 발행하는 사모사채의 조건을 정리합니다.

[엄격한 규칙]
1. 차주, 사업명, 금액, 만기 등 PDF 에 있는 값은 그대로 사용하라.
2. PDF 에 없는 항목(채권 금리, 인수수수료 등)은 반드시 "[ TBD ]" 로 표시하라.
3. 발행금액은 PDF 의 Tr.B 금액과 동일하게.
4. 발행일은 PDF 의 대출 인출 예정일(또는 약정 예정일)과 동일.
5. 만기일은 발행일 기준 6개월 후.
6. fields 배열은 정확히 9개. 순서와 label 은 스키마와 동일하게.
7. 출력은 JSON 만. 다른 텍스트 금지.

[JSON 출력 스키마]
{
  "intro_paragraph": "본 건 사모사채 소개 1~2문장 (차주명, 사업명, Tr.B 금액, 만기 포함)",
  "fields": [
    {"label": "사모사채명", "value": "TBD(신규 유동화 SPC) 제1회 무기명식 무보증 사모사채"},
    {"label": "사채 유형", "value": "국내 전자등록 또는 실물 발행, 무기명식 무보증 사모사채"},
    {"label": "발행인", "value": "TBD(신규 유동화 SPC)"},
    {"label": "기초자산", "value": "[사업명] 토지담보대출 Tr.B 대출채권 — PDF 값 사용"},
    {"label": "발행금액", "value": "[Tr.B 금액] — PDF 값 사용"},
    {"label": "발행일", "value": "[YYYY년 MM월] (예정) — PDF 값 사용"},
    {"label": "만기일", "value": "[YYYY년 MM월 00일] (발행일로부터 약 6개월) — PDF 값 사용"},
    {"label": "금융조건", "value": "All-in [ TBD ]%, 채권 금리: 연 [ TBD ]%(고정, 세전), 인수수수료: 발행금액의 [ TBD ]%"},
    {"label": "이자지급주기", "value": "3개월 단위 후취 또는 만기 전액 후취"}
  ]
}"""

SLIDE_5_USER_TEMPLATE = """[PDF 원문]
{pdf_text}

위 PDF 에서 Tr.B 금액, 발행일, 만기일, 차주명, 사업명을 찾아 사모사채 개요 JSON 을 출력하라.
PDF 에 없는 항목은 반드시 [ TBD ] 로 표시하고, fields 는 정확히 9개."""


def generate_sasae_overview(pdf_text: str) -> dict:
    """
    슬라이드 5 (1.1 본건 사모사채 개요) 데이터를 Claude 로 생성합니다.

    Returns
    -------
    call_claude() 반환값 dict
    """
    print("=" * 60)
    print(f"[FORCE-DEBUG-AI] generate_sasae_overview 호출됨 - PDF 길이={len(pdf_text)}")
    print("=" * 60)

    result = call_claude(
        system_prompt=SLIDE_5_SYSTEM_PROMPT,
        user_prompt=SLIDE_5_USER_TEMPLATE.format(pdf_text=pdf_text),
        slide_num=5,
        pdf_context=pdf_text,
        prompt_version="v1",
    )

    if result["ok"]:
        v = verify_numbers_in_pdf(result["data"], pdf_text)
        if not v["ok"]:
            print(f"[경고] 슬라이드 5 환각 의심 숫자: {v['hallucinated_numbers']}")
        else:
            print(f"[슬라이드 5] 숫자 검증 통과 ({v['verified_count']}개)")

    return result


# ════════════════════════════════════════════════════════
# 슬라이드 7: 2.1 투자구조도
# ════════════════════════════════════════════════════════

SLIDE_7_SYSTEM_PROMPT = """당신은 부동산 금융 투자구조 분석 전문가입니다.
PDF 에서 투자 참여기관을 추출하여 투자구조도용 데이터를 만듭니다.

[엄격한 규칙]
1. 차주, 시행사, 시공사, 신탁사, 대주, 사채권자는 PDF 에 있는 이름 그대로 사용.
2. 금액(담보대출 총액, Tr.A, Tr.B)은 PDF 원문 그대로 사용. 추정 금지.
3. SPC 명은 "신규 유동화 SPC T.B.D." 로 표시.
4. 사채권자는 "사채권자 T.B.D." 로 표시.
5. relationships 는 PDF 에서 파악 가능한 계약관계만 기재.
6. 출력은 JSON 만.

[JSON 출력 스키마]
{
  "intro_paragraph": "투자구조 소개 1~2문장 (총 대출금액, Tr.A/B 구성 포함)",
  "total_loan_amount": "PDF 원문 값",
  "tranches": [
    {"name": "Tranche A", "amount": "PDF 원문 값"},
    {"name": "Tranche B", "amount": "PDF 원문 값"}
  ],
  "entities": {
    "borrower": "차주명 — PDF 값",
    "constructor": "시공사명 — PDF 값",
    "trustee": "신탁사명 — PDF 값",
    "tr_a_lenders": "Tr.A 대주 (PDF 에 없으면 'T.B.D.')",
    "spc": "신규 유동화 SPC T.B.D.",
    "bondholders": "사채권자 T.B.D."
  },
  "relationships": [
    "차주 - 시공사: 공사도급계약",
    "차주 - 대주: 대출약정",
    "차주 - 신탁사: 담보신탁계약",
    "SPC - 사채권자: 사모사채 발행/인수"
  ]
}"""

SLIDE_7_USER_TEMPLATE = """[PDF 원문]
{pdf_text}

위 PDF 에서 투자 참여기관과 금액을 찾아 투자구조도 JSON 을 출력하라.
PDF 에 없는 기관명은 T.B.D. 로 표시하고, 금액은 절대 추정하지 마라."""


def generate_investment_structure(pdf_text: str) -> dict:
    """
    슬라이드 7 (2.1 투자구조도) 데이터를 Claude 로 생성합니다.

    Returns
    -------
    call_claude() 반환값 dict
    """
    print("=" * 60)
    print(f"[FORCE-DEBUG-AI] generate_investment_structure 호출됨 - PDF 길이={len(pdf_text)}")
    print("=" * 60)

    result = call_claude(
        system_prompt=SLIDE_7_SYSTEM_PROMPT,
        user_prompt=SLIDE_7_USER_TEMPLATE.format(pdf_text=pdf_text),
        slide_num=7,
        pdf_context=pdf_text,
        prompt_version="v1",
    )

    if result["ok"]:
        v = verify_numbers_in_pdf(result["data"], pdf_text)
        if not v["ok"]:
            print(f"[경고] 슬라이드 7 환각 의심 숫자: {v['hallucinated_numbers']}")
        else:
            print(f"[슬라이드 7] 숫자 검증 통과 ({v['verified_count']}개)")

    return result


# ════════════════════════════════════════════════════════
# 내부 헬퍼
# ════════════════════════════════════════════════════════

def _replace_text_keep_runs(tf, new_text: str):
    """기존 단락/런의 XML 구조(폰트·크기·색상)를 보존하면서 텍스트만 교체합니다.

    _replace_text_frame_content 와 달리 모든 <a:p>/<a:r>/<a:rPr> 를 제거하지 않고
    기존 단락을 재사용하므로 템플릿에서 지정한 폰트가 그대로 유지됩니다.
    줄 수가 기존 단락보다 많으면 마지막 단락 스타일을 복사해 추가합니다.
    """
    lines = (new_text or '').split('\n')
    txBody = tf._txBody
    all_paras = txBody.findall(_qn('a:p'))

    if not all_paras:
        _replace_tf_content(tf, new_text)
        return

    last_p = all_paras[-1]

    for i, para in enumerate(all_paras):
        if i < len(lines):
            runs = para.findall(_qn('a:r'))
            if runs:
                t_elem = runs[0].find(_qn('a:t'))
                if t_elem is not None:
                    t_elem.text = lines[i]
                else:
                    t_elem = etree.SubElement(runs[0], _qn('a:t'))
                    t_elem.text = lines[i]
                for r in runs[1:]:
                    para.remove(r)
            else:
                new_r = etree.SubElement(para, _qn('a:r'))
                last_runs = last_p.findall(_qn('a:r'))
                if last_runs:
                    old_rPr = last_runs[0].find(_qn('a:rPr'))
                    if old_rPr is not None:
                        new_r.insert(0, copy.deepcopy(old_rPr))
                t_elem = etree.SubElement(new_r, _qn('a:t'))
                t_elem.text = lines[i]
        else:
            runs = para.findall(_qn('a:r'))
            if runs:
                t_elem = runs[0].find(_qn('a:t'))
                if t_elem is not None:
                    t_elem.text = ''
                for r in runs[1:]:
                    para.remove(r)

    for i in range(len(all_paras), len(lines)):
        new_p = copy.deepcopy(last_p)
        new_runs = new_p.findall(_qn('a:r'))
        if new_runs:
            t_elem = new_runs[0].find(_qn('a:t'))
            if t_elem is not None:
                t_elem.text = lines[i]
            for r in new_runs[1:]:
                new_p.remove(r)
        else:
            new_r = etree.SubElement(new_p, _qn('a:r'))
            t_elem = etree.SubElement(new_r, _qn('a:t'))
            t_elem.text = lines[i]
        txBody.append(new_p)


def _find_shape_by_pos(slide, left_cm: float, top_cm: float, tol_cm: float = 0.35):
    """좌표(cm)로 슬라이드에서 텍스트 프레임 도형을 찾습니다."""
    tol = int(tol_cm * 360000)
    lx  = int(left_cm * 360000)
    tx  = int(top_cm  * 360000)
    best, best_dist = None, float('inf')
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        dl = abs(shape.left - lx)
        dt = abs(shape.top  - tx)
        if dl <= tol and dt <= tol:
            d = dl + dt
            if d < best_dist:
                best_dist, best = d, shape
    return best


def _find_table_by_pos(slide, left_cm: float, top_cm: float, tol_cm: float = 0.35):
    """좌표(cm)로 슬라이드에서 TABLE shape 를 찾습니다."""
    tol = int(tol_cm * 360000)
    lx  = int(left_cm * 360000)
    tx  = int(top_cm  * 360000)
    best, best_dist = None, float('inf')
    for shape in slide.shapes:
        if shape.shape_type != 19:   # 19 = TABLE
            continue
        dl = abs(shape.left - lx)
        dt = abs(shape.top  - tx)
        if dl <= tol and dt <= tol:
            d = dl + dt
            if d < best_dist:
                best_dist, best = d, shape
    return best


def _fill_group_label(group_shape, text: str):
    """GROUP 내부 첫 번째 TEXT_BOX 의 텍스트를 교체합니다 (섹션 헤더 레이블)."""
    try:
        for child in group_shape.shapes:
            if child.has_text_frame:
                _replace_text_keep_runs(child.text_frame, text)
                return
    except Exception:
        pass


def _exec_summary_parts(slide):
    """Investment Highlights(슬라이드 2) 템플릿의 구성요소를 '구조'로 찾습니다.

    템플릿(레이아웃.pptx)이 교체되면 도형 name·좌표가 통째로 바뀌므로
    이름/좌표로 찾으면 조용히 실패해 **이전 딜의 템플릿 문구가 그대로 출고**된다.
    (2026-07-29 템플릿 교체 후 실제로 발생 — 포스코이앤씨 문구 잔존.)
    그래서 이름 대신 아래 구조 규칙으로 찾는다:

      본문 블록 = 육각형(AUTO_SHAPE) 2개 + TEXT_BOX 1개 를 가진 최상위 GROUP
      제목 블록 = PICTURE(체크 아이콘) + TEXT_BOX 를 가진 최상위 GROUP
      부제      = 최상위 TEXT_BOX 중 텍스트가 '내용' 인 것

    각 목록은 top(세로) 순으로 정렬해 섹션 1·2·3 에 대응시킨다.

    Returns
    -------
    dict: {"titles": [GROUP×3], "bodies": [TEXT_BOX×3],
           "hex_out": [AUTO_SHAPE×3], "hex_in": [AUTO_SHAPE×3], "subs": [...]}
           찾지 못한 자리는 None.
    """
    title_grps, body_grps, subs = [], [], []
    for sh in slide.shapes:
        if sh.shape_type == 6:                       # GROUP
            try:
                kids = list(sh.shapes)
            except Exception:
                continue
            tbs  = [c for c in kids if c.shape_type == 17]
            hexs = [c for c in kids if c.shape_type == 1]      # AUTO_SHAPE(육각형)
            pics = [c for c in kids if c.shape_type == 13]
            if len(hexs) >= 2 and tbs:
                body_grps.append((sh.top or 0, sh, tbs[0], hexs))
            elif pics and tbs:
                title_grps.append((sh.top or 0, sh))
        elif sh.shape_type == 17 and sh.has_text_frame:
            if sh.text_frame.text.strip() == "내용":
                subs.append((sh.top or 0, sh))

    title_grps.sort(key=lambda x: x[0])
    body_grps.sort(key=lambda x: x[0])
    subs.sort(key=lambda x: x[0])

    def _pad(seq, n=3):
        return (list(seq) + [None] * n)[:n]

    hex_out, hex_in = [], []
    for _t, _g, _tb, hexs in body_grps:
        # 바깥 육각형 = 더 넓은 쪽
        srt = sorted(hexs, key=lambda h: -(h.width or 0))
        hex_out.append(srt[0])
        hex_in.append(srt[1] if len(srt) > 1 else None)

    return {
        "titles":  _pad([g for _t, g in title_grps]),
        "bodies":  _pad([tb for _t, _g, tb, _h in body_grps]),
        "groups":  _pad([g for _t, g, _tb, _h in body_grps]),
        "hex_out": _pad(hex_out),
        "hex_in":  _pad(hex_in),
        "subs":    _pad([s for _t, s in subs]),
    }


def _remove_non_pdf_shapes(slide):
    """PDF 원본 구조도에 없는 shape 를 XML 에서 완전히 제거합니다.

    삭제 대상 (좌표 기반 식별):
      - shape[10] TEXT_BOX  (L=7.41, T=10.72)  원래 "사업 시행"
      - shape[11] TABLE 2r  (L=13.18, T=14.26) 원래 자산관리자
      - shape[13] LINE      (L=14.44, T=11.77) 자산관리자→차주 세로 화살표
      - shape[14] TEXT_BOX  (L=13.87, T=12.65) 원래 "자산관리"
      - shape[17] TABLE 7r  (L=18.87, T=10.96) 원래 투자자 목록
      - shape[18] LINE      (L=15.70, T=11.07) 투자자→차주 가로 화살표
      - shape[22] TEXT_BOX  (L=15.52, T=11.07) 원래 "Equity 825억원"
      - shape[25] TEXT_BOX  (L=8.26, T=12.74)  원래 "책임준공 확약"
      - shape[27] PICTURE   (shape_type=13)     사업지 이미지
      - shape[29] TEXT_BOX  (L=0.00, T=18.35)  사업명 푸터
    """
    # 좌표로 삭제할 shape 를 식별
    # (left_cm, top_cm, tolerance_cm, shape_type_filter)
    # shape_type_filter: None=모든 타입, 19=TABLE, 17=TEXT_BOX, 13=PICTURE, 9=LINE
    delete_specs = [
        (7.41,  10.72, 0.35, 17),   # [10] "사업 시행"
        (13.18, 14.26, 0.35, 19),   # [11] 자산관리자 TABLE
        (14.44, 11.77, 0.35, 9),    # [13] 자산관리자 세로 LINE
        (13.87, 12.65, 0.35, 17),   # [14] "자산관리"
        (18.87, 10.96, 0.35, 19),   # [17] 투자자 TABLE
        (15.70, 11.07, 0.20, 9),    # [18] 투자자 가로 LINE
        (15.52, 11.07, 0.35, 17),   # [22] "Equity"
        (8.26,  12.74, 0.35, 17),   # [25] "책임준공 확약"
        (0.00,  18.35, 0.20, 17),   # [29] 사업명 푸터
    ]

    to_remove = []
    for shape in slide.shapes:
        l_cm = shape.left / 360000 if shape.left else 0
        t_cm = shape.top  / 360000 if shape.top  else 0
        st   = shape.shape_type

        # PICTURE (사업지 이미지) — shape_type == 13
        if st == 13:
            to_remove.append(shape)
            continue

        for spec_l, spec_t, tol, spec_type in delete_specs:
            if spec_type is not None and st != spec_type:
                continue
            if abs(l_cm - spec_l) <= tol and abs(t_cm - spec_t) <= tol:
                to_remove.append(shape)
                break

    for shape in to_remove:
        sp = shape._element
        sp.getparent().remove(sp)

    if to_remove:
        print(f"[build_slide_7] PDF에 없는 shape {len(to_remove)}개 삭제 완료")


def _fill_sasae_table(slide, fields: list):
    """슬라이드 6 템플릿의 9r×2c TABLE 에 value 값을 씁니다.
    레이블 열(col 0)은 그대로 유지하고, 값 열(col 1) 만 교체합니다."""
    sh = _find_table_by_pos(slide, left_cm=1.09, top_cm=3.84, tol_cm=0.40)
    if sh is None:
        print("[경고] _fill_sasae_table: 9r×2c TABLE 못 찾음")
        return
    tbl = sh.table
    for i, f in enumerate(fields):
        if i >= len(tbl.rows):
            break
        _replace_text_keep_runs(tbl.cell(i, 1).text_frame, f.get("value", ""))


def _fill_entity_tables(slide, entities: dict, tranches: list, total_loan: str):
    """슬라이드 8 투자구조도의 entity TABLE 셀에 실제 기관명/금액을 씁니다.

    PDF 원본 구조도 기준:
      - 신탁사·차주·시공사 3개 entity 박스(2r×1c) 기관명 교체
      - 대주 박스(4r×1c): "Tr.A" / "Tr.B" 만 표기 (금액 제거)
      - 자산관리자·투자자 박스는 _remove_non_pdf_shapes() 에서 삭제
    """
    # ── 3개 entity 박스 (2r×1c): row[0]=레이블 유지, row[1]=기관명 교체 ──
    entity_map = [
        (10.47, 4.98,  entities.get("trustee",     "")),   # 신탁사
        (10.48, 9.55,  entities.get("borrower",    "")),   # 차주
        (10.20, 14.26, entities.get("constructor", "")),   # 시공사
    ]
    for l, t, val in entity_map:
        sh = _find_table_by_pos(slide, l, t)
        if sh:
            _replace_text_keep_runs(sh.table.cell(1, 0).text_frame, val)
        else:
            print(f"[경고] _fill_entity_tables: entity TABLE 못 찾음 ({l}, {t})")

    # ── 대주 박스 (shape[09], 4r×1c, L=18.87 T=5.44) — 금액 제거, Tr 라벨만 ──
    loan_sh = _find_table_by_pos(slide, 18.87, 5.44)
    if loan_sh:
        tbl = loan_sh.table
        # row[0]: 헤더 → 비움 (Bridge Loan 총액 제거)
        _replace_text_keep_runs(tbl.cell(0, 0).text_frame, "")
        # row[1..]: tranche 이름만 (금액 없이)
        for i, tr in enumerate(tranches):
            row_idx = i + 1
            if row_idx < len(tbl.rows):
                _replace_text_keep_runs(
                    tbl.cell(row_idx, 0).text_frame,
                    tr.get('name', '')
                )
        for i in range(len(tranches) + 1, len(tbl.rows)):
            _replace_text_keep_runs(tbl.cell(i, 0).text_frame, "")
    else:
        print("[경고] _fill_entity_tables: 대주 TABLE 못 찾음")


# ════════════════════════════════════════════════════════
# 슬라이드 2 빌더: Executive Summary
# ════════════════════════════════════════════════════════

def _relayout_exec_summary(slide, parts=None):
    """슬라이드 1 Executive Summary 동적 섹션 레이아웃 (이름 기반).

    - 육각형(<>) 좌우폭/left 통일 (바깥/안쪽 각각 동일값, 세로만 조절)
    - 본문 실제 텍스트의 줄 수를 추정 → 본문 높이 산출 →
      꺾쇠 H = 본문높이 + 상하 0.15" 여백 (본문은 꺾쇠 안 세로중앙 정렬)
    - [제목(+부제)+꺾쇠] 블록 단위로 세 섹션을 균등 간격으로 세로 중앙 분포
    - 바깥/안쪽 육각형의 상하 offset(델타)은 원본 비율 보존
    좌표 단위: 인치(EMU = 인치×914400). 모든 대상은 고유 name 으로 매칭.
    """
    IN = 914400
    if parts is None:
        parts = _exec_summary_parts(slide)

    O_L, O_W = int(1.30 * IN), int(8.24 * IN)   # 바깥 육각형 좌/폭
    I_L, I_W = int(1.45 * IN), int(7.94 * IN)   # 안쪽 육각형 좌/폭
    MARGIN = 0.15            # 꺾쇠 내부 상하 여백(인치)
    TITLE_H = 0.40           # 제목 GROUP 높이
    SUB_BLOCK = 0.71         # 섹션1: 제목+부제+여백 (제목top→꺾쇠top)
    TITLE_GAP = 0.45         # 섹션2·3: 제목top→꺾쇠top
    TOP, BOT = 0.84, 7.00    # 사용 가능 세로 범위(로고 아래 ~ 푸터 위)

    def _strip_empty_paras(tb):
        """본문 텍스트프레임의 빈 단락(<a:p> 텍스트 없음)을 제거. 최소 1개는 유지.
        (_replace_text_keep_runs 가 템플릿의 잔여 빈 단락을 남겨 줄 수를 부풀리는 문제 해결)"""
        txbody = tb.text_frame._txBody
        ps = txbody.findall(_qn('a:p'))
        nonempty = [p for p in ps
                    if "".join(t.text or "" for t in p.findall('.//' + _qn('a:t'))).strip()]
        if not nonempty:
            return
        for p in ps:
            txt = "".join(t.text or "" for t in p.findall('.//' + _qn('a:t'))).strip()
            if not txt:
                txbody.remove(p)

    def _est_body(tb):
        """본문 줄 수·행높이·예상높이(인치) 추정 (빈 단락 제거 후)."""
        tf = tb.text_frame
        width_in = (tb.width or 0) / IN
        fpt = 11.0
        done = False
        for p in tf.paragraphs:
            for r in p.runs:
                if r.font.size:
                    fpt = r.font.size.pt
                    done = True
                    break
            if done:
                break
        # 한 줄에 들어가는 '글자폭 단위'. 한글·한자는 전각(1.0), 숫자·영문·기호는 반각(0.55).
        # 금융 문구는 숫자·%·영문이 많아 전각으로만 세면 줄 수가 크게 과대추정된다.
        # 글상자 폭 전체를 쓰지는 못한다(꺾쇠 화살표 안쪽 + 내부 여백) → 0.93 계수.
        cpl = max(1.0, width_in * 0.93 * 72.0 / fpt)

        def _units(s):
            u = 0.0
            for ch in s:
                if ('가' <= ch <= '힣' or '㄰' <= ch <= '㆏'
                        or '一' <= ch <= '鿿'):
                    u += 1.0
                else:
                    u += 0.55
            return u

        lines = 0
        for p in tf.paragraphs:
            s = "".join(r.text for r in p.runs).strip()
            if s:
                lines += max(1, int(-(-_units(s) // cpl)))   # ceil
        lines = max(1, lines)
        line_h = fpt * 1.25 / 72.0
        return lines, round(line_h, 3), round(lines * line_h, 3)

    # ★섹션 = (제목 GROUP, 본문블록 GROUP, 본문 TEXT_BOX) — 이름이 아니라 구조로 찾은 것.
    secs = list(zip(parts["titles"], parts["groups"], parts["bodies"]))

    # 1) 본문 빈 단락 제거 후 추정 → 꺾쇠(본문블록) H, 블록 H
    est, hexH, blockH = [], [], []
    for (_gt, _gb, tb) in secs:
        if tb is not None:
            _strip_empty_paras(tb)
        e = _est_body(tb) if tb is not None else (1, 0.2, 0.2)
        est.append(e)
        h = e[2] + 2 * MARGIN
        hexH.append(h)
        blockH.append(TITLE_GAP + h)

    # 2) 균등 간격 산출 (세로 중앙 분포)
    #    하한 0.30" 이면 내용이 많을 때 3번째 섹션이 슬라이드 밖으로 밀린다 → 0.10" 까지 허용.
    content = sum(blockH)
    gap = max(0.10, (BOT - TOP - content) / 2)
    if TOP + content + 2 * gap > BOT + 0.01:
        print(f"[경고] build_slide_2: 하이라이트 3섹션이 한 장에 안 들어감 "
              f"(필요 {TOP + content + 2 * gap:.2f}\" > 한계 {BOT:.2f}\") "
              f"— 문구를 줄이거나 2페이지로 나눠야 함")

    # 3) 배치 — 본문블록은 GROUP 통째로 옮기고 늘린다.
    #    (그룹 안 자식 좌표를 직접 건드리면 그룹 변환 때문에 어긋난다.)
    report = []
    cur = TOP
    for i, (gt, gb, tb) in enumerate(secs):
        ttop = cur
        hexT = ttop + TITLE_GAP
        if gt is not None:
            gt.top = int(ttop * IN)
        before_h = None
        if gb is not None:
            before_h = round((gb.height or 0) / IN, 2)
            gb.left, gb.width = O_L, O_W
            gb.top, gb.height = int(hexT * IN), int(hexH[i] * IN)
        if tb is not None:
            try:
                tb.text_frame.word_wrap = True
                tb.text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
            except Exception:
                pass
        report.append({
            "sec": f"섹션{i + 1}", "title_top": round(ttop, 2), "hex_top": round(hexT, 2),
            "lines": est[i][0], "body_h": est[i][2],
            "hexH_before": before_h,
            "hexH_after": round(hexH[i], 2),
        })
        cur = hexT + hexH[i] + gap

    print(f"[build_slide_2] 동적 레이아웃 gap={round(gap, 2)}\":")
    for r in report:
        print(f"   {r}")
    return report


def style_table(sh, *, nested=False, has_header=False, label_cols=(0,),
                header_fill=None, value_fill=None, label_fill=None):
    """표 셀 서식 표준 적용 (재사용 헬퍼).

    - 여백: 기본 0 / 좌측정렬→marL=0.3cm / 우측정렬→marR=0.3cm / 가운데→전부 0
    - 글씨: 일반표 10.5pt, 중첩표(nested=True) 10pt
    - Bold/Light: 헤더행(has_header & row0)·구분열(label_cols)=피플폰트 Bold,
                  세부내용=피플폰트 Light  (font.bold 속성은 쓰지 않고 폰트명으로 굵기 표현)
    - 색: header_fill(헤더행)·value_fill(값셀) 지정 시 적용, 미지정이면 기존 유지
          (헤더 08377C, 구분/값 회색 D9D9D9, 강조 3E95BE 등은 호출부에서 지정)
    - 정렬: 가로 기본 가운데(미지정 시), 세로 항상 가운데(MIDDLE)
    - 테두리: 모든 셀 4변 전체, 0.5pt(6350 EMU), A5A5A5(회색 강조3)
    - 행높이: has_header 면 헤더행 0.6cm(중첩 0.55cm)
    """
    tbl = sh.table
    # 자동 밴딩/특수행 스타일 제거 — 값 셀에 의도치 않은 음영(연파랑 줄무늬)이 깔리는 것 방지.
    #   셀 색은 아래에서 직접 지정(헤더 네이비/구분 F2F2F2/값 투명)하므로 자동 스타일은 끈다.
    for attr in ("first_row", "last_row", "first_col", "last_col", "horz_banding", "vert_banding"):
        try:
            setattr(tbl, attr, False)
        except Exception:
            pass
    fpt = 10 if nested else 10.5
    M = Cm(0.3)
    Z = Inches(0)
    BORDER_HEX = "A5A5A5"          # 회색, 강조3 (테마 accent3) — 레이아웃 표 실측값
    BORDER_W = 6350               # 0.5pt (EMU)

    def _set_borders(cell):
        tcPr = cell._tc.get_or_add_tcPr()
        for tag in ("a:lnL", "a:lnR", "a:lnT", "a:lnB"):
            for el in tcPr.findall(_qn(tag)):
                tcPr.remove(el)
        for i, tag in enumerate(("a:lnL", "a:lnR", "a:lnT", "a:lnB")):
            ln = tcPr.makeelement(_qn(tag),
                                  {"w": str(BORDER_W), "cap": "flat", "cmpd": "sng", "algn": "ctr"})
            sf = ln.makeelement(_qn("a:solidFill"), {})
            sf.append(sf.makeelement(_qn("a:srgbClr"), {"val": BORDER_HEX}))
            ln.append(sf)
            ln.append(ln.makeelement(_qn("a:prstDash"), {"val": "solid"}))
            tcPr.insert(i, ln)          # 테두리는 tcPr 최상단(스키마 순서)

    for ri in range(len(tbl.rows)):
        for ci in range(len(tbl.columns)):
            cell = tbl.cell(ri, ci)
            tf = cell.text_frame
            # 가로 정렬: 미지정이면 가운데로 통일
            for p in tf.paragraphs:
                if p.alignment is None:
                    p.alignment = PP_ALIGN.CENTER
            algn = str(tf.paragraphs[0].alignment or "")
            cell.margin_left = M if algn.startswith("LEFT") else Z
            cell.margin_right = M if algn.startswith("RIGHT") else Z
            cell.margin_top = Z
            cell.margin_bottom = Z
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE       # 세로 항상 가운데
            is_header_cell = has_header and ri == 0
            is_bold = is_header_cell or (ci in label_cols)
            fname = "피플폰트 Bold" if is_bold else "피플폰트 Light"
            # 헤더(색칠된 셀) 글자=흰색, 그 외(구분·내용)=검정
            fcolor = _C_WHITE if is_header_cell else RGBColor(0, 0, 0)
            for p in tf.paragraphs:
                for r in p.runs:
                    r.font.size = Pt(fpt)
                    r.font.name = fname
                    r.font.color.rgb = fcolor
            if has_header and ri == 0 and header_fill is not None:
                cell.fill.solid()
                cell.fill.fore_color.rgb = header_fill
            elif ci in label_cols and label_fill is not None:
                cell.fill.solid()
                cell.fill.fore_color.rgb = label_fill
            elif value_fill is not None and ci not in label_cols:
                cell.fill.solid()
                cell.fill.fore_color.rgb = value_fill
            else:
                # 색 미지정 셀(내용/값 셀) = 흰색으로 확실히 칠함 (표 스타일 음영 비침 방지)
                cell.fill.solid()
                cell.fill.fore_color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            _set_borders(cell)                              # 전체 테두리 A5A5A5 0.5pt
    if has_header and len(tbl.rows) > 0:
        tbl.rows[0].height = Cm(0.55 if nested else 0.6)


def clean_bullet(tf, marL=171450, indent=-171450):
    """텍스트프레임 모든 단락의 자동 글머리/레벨/들여쓰기를 정리한다.

    - a:buChar / a:buAutoNum / a:buFont 제거 후 a:buNone 삽입 (자동 글머리 제거)
    - lvl 속성 제거 → 모든 단락 level 0 통일
    - marL/indent 를 동일 값으로 통일 (기본: marL=171450, indent=-171450)
    텍스트·런·폰트·색은 건드리지 않는다.
    """
    A = "http://schemas.openxmlformats.org/drawingml/2006/main"

    def _q(tag):
        return "{%s}%s" % (A, tag)

    for p in tf.paragraphs:
        pPr = p._p.get_or_add_pPr()
        # 들여쓰기/레벨 통일
        pPr.set("marL", str(int(marL)))
        pPr.set("indent", str(int(indent)))
        if "lvl" in pPr.attrib:
            del pPr.attrib["lvl"]
        # 기존 글머리/글머리폰트/자동번호 제거
        for tag in ("buChar", "buAutoNum", "buFont", "buNone"):
            for el in pPr.findall(_q(tag)):
                pPr.remove(el)
        # buNone 삽입 (스키마상 buColor/buSzx/buFont 다음, 단 여기선 다 제거했으므로 끝에 append)
        pPr.append(pPr.makeelement(_q("buNone"), {}))


# 전 슬라이드 공통 푸터 사업명 (딜별 기본값 — business_name 인자로 덮어쓸 수 있음)
_FOOTER_BIZ = "천안 부성2지구 도시개발사업"
_FOOTER_GRAY = RGBColor(0x80, 0x80, 0x80)
_FONT_LIGHT = "피플폰트 Light"


def add_footer(slide, page_num, source=None, business_name=None):
    """전 슬라이드 공통 푸터 (형식 고정).

    - 오른쪽 아래: "{사업명}  |  {page_num}"  (피플폰트 Light 8pt, 808080, 우측정렬)
    - 왼쪽 아래(source 있을 때만): "출처: {source}"  (피플폰트 Light 9pt, 808080, 좌측정렬)
    - page_num 은 PPT 슬라이드 순서(slides 인덱스+1) 권장
    - 두 텍스트박스 모두 내부 여백 0
    기존 푸터성 텍스트박스(name 에 "슬라이드 번호" 포함)는 먼저 삭제 후 새로 생성.
    """
    biz = business_name or _FOOTER_BIZ

    # 1) 기존 푸터성(페이지번호/사업명) 텍스트박스 제거
    for sh in list(slide.shapes):
        if "슬라이드 번호" in (sh.name or ""):
            sh._element.getparent().remove(sh._element)

    def _mk(left, width, text, size_pt, align):
        tb = slide.shapes.add_textbox(Inches(left), Inches(7.16), Inches(width), Inches(0.26))
        tf = tb.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = Inches(0)
        tf.margin_top = tf.margin_bottom = Inches(0)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = tf.paragraphs[0]
        p.alignment = align
        run = p.add_run()
        run.text = text
        run.font.name = _FONT_LIGHT
        run.font.size = Pt(size_pt)
        run.font.color.rgb = _FOOTER_GRAY
        return tb

    # 2) 오른쪽 푸터 (사업명 | 순번) — 우측 끝 ≈ 10.50"
    #    ★쪽번호는 고정 숫자가 아니라 PowerPoint '슬라이드 번호' 필드로 넣는다.
    #      고정 숫자를 쓰면 앞에 슬라이드가 빠지거나 늘 때 실제 위치와 어긋난다
    #      (요약본에서 섹션 divider 를 빼자 4번째 장에 '6' 이 찍히던 문제).
    #      frame_builders._add_combined_footer 와 동일한 방식.
    _tb = _mk(4.83, 5.67, f"{biz}  |  ", 8, PP_ALIGN.RIGHT)
    try:
        _p = _tb.text_frame.paragraphs[0]
        _fld = _p._p.makeelement(_qn('a:fld'), {
            'id': '{B7E3A1C2-0000-4000-9000-000000000002}', 'type': 'slidenum'})
        _rPr = _fld.makeelement(_qn('a:rPr'), {'lang': 'ko-KR', 'sz': '800'})
        _sf = _rPr.makeelement(_qn('a:solidFill'), {})
        _clr = _sf.makeelement(_qn('a:srgbClr'), {'val': '808080'})
        _sf.append(_clr)
        _rPr.append(_sf)
        _lat = _rPr.makeelement(_qn('a:latin'), {'typeface': _FONT_LIGHT})
        _ea = _rPr.makeelement(_qn('a:ea'), {'typeface': _FONT_LIGHT})
        _rPr.append(_lat)
        _rPr.append(_ea)
        _fld.append(_rPr)
        _t = _fld.makeelement(_qn('a:t'), {})
        _t.text = str(page_num)          # 편집 전 표시용 기본값(열면 실제 번호로 갱신)
        _fld.append(_t)
        _p._p.append(_fld)
    except Exception as _e:
        # 필드 생성 실패 시 고정 숫자로 폴백
        print(f"[add_footer] 쪽번호 필드 실패({_e}) — 고정 숫자 사용")
        _tb.text_frame.paragraphs[0].runs[0].text = f"{biz}  |  {page_num}"

    # 3) 왼쪽 출처 (source 지정 시에만)
    if source:
        _mk(0.33, 5.50, f"출처: {source}", 9, PP_ALIGN.LEFT)


def build_slide_2_executive_summary(prs, data: dict,
                                     business_name: str = "",
                                     page_num: int = 2):
    """
    슬라이드 2 (Executive Summary) PPT 슬라이드를 생성합니다.
    레이아웃.pptx index=1 템플릿을 복제해 텍스트만 교체합니다.

    Parameters
    ----------
    prs           : create_presentation_from_template() 로 만든 Presentation 객체
    data          : generate_executive_summary() 반환 data dict
    business_name : 하단 사업명 푸터
    page_num      : 슬라이드 번호
    """
    slide = clone_slide_layout(prs, "executive_summary")

    deal_title = data.get("deal_title", "")

    # ── ★원본 Executive Summary 내용을 '3개 섹션(제목+본문)'으로 요약한 결과 사용 ──
    #   원문에 없는 항목(분양성/만기일자 등) 억지 생성 금지. 구버전(key_points)도 호환 변환.
    sections = data.get("sections")
    if not sections:
        _kp = data.get("key_points", [])
        _ds = (data.get("deal_summary") or "").strip()
        sections = [{"title": "거래 개요", "body": _ds}]
        _pts = [p for p in _kp[:4] if (p.get("description") or "").strip()]
        if _pts:
            sections.append({"title": "핵심 투자 포인트",
                             "body": "\n".join(f"{p.get('title','')} {p.get('description','')}".strip()
                                               for p in _pts)})
        if len(_kp) > 4 and (_kp[4].get("description") or "").strip():
            sections.append({"title": _kp[4].get("title", "만기"),
                             "body": _kp[4].get("description", "")})
    sections = list(sections or [])[:3]
    while len(sections) < 3:
        sections.append({"title": "", "body": ""})

    def _arrowify(body):
        lines = [ln.lstrip("·•-→ ").strip() for ln in str(body or "").split("\n") if ln.strip()]
        return "\n".join(f"→ {ln}" for ln in lines)

    _TITLES = [(sections[0].get("title") or "거래 개요").strip(),
               (sections[1].get("title") or "핵심 투자 포인트").strip(),
               (sections[2].get("title") or "사업 구조").strip()]
    _SEC1_BODY = _arrowify(sections[0].get("body"))
    _SEC2_BODY = _arrowify(sections[1].get("body"))
    _SEC3_BODY = _arrowify(sections[2].get("body"))

    # ★도형은 name·좌표가 아니라 '구조'로 찾는다 (_exec_summary_parts 주석 참고).
    #   템플릿이 바뀌면 좌표 매칭이 조용히 실패해 이전 딜 문구가 그대로 남는다.
    parts = _exec_summary_parts(slide)

    # 본문 박스 교체 (부제 '내용' 글상자는 삭제 대상이라 채우지 않음)
    for _i, _txt in enumerate((_SEC1_BODY, _SEC2_BODY, _SEC3_BODY)):
        _b = parts["bodies"][_i]
        if _b is not None:
            _replace_text_keep_runs(_b.text_frame, _txt)
        else:
            print(f"[경고] build_slide_2: 섹션{_i + 1} 본문 글상자를 못 찾음 "
                  f"— 템플릿 문구가 남을 수 있음")

    # GROUP 제목 = 3개 섹션 제목(내용 기반). 'Executive Summary' 단어 금지.
    for _i, _ttl in enumerate(_TITLES):
        _g = parts["titles"][_i]
        if _g is not None:
            _fill_group_label(_g, _ttl)
        else:
            print(f"[경고] build_slide_2: 섹션{_i + 1} 제목 GROUP 을 못 찾음")

    # ── 수정 2: 섹션 간격 압축 + 육각형 컨테이너 폭 통일 ──
    _relayout_exec_summary(slide, parts=parts)

    # ── 3개 본문 박스 모두 자동 글머리(템플릿 → 등) 제거 + 회색 통일 ──
    #   (내가 넣은 '→ ' 가 유일한 글머리가 되도록. 3섹션 글씨 = 같은 회색으로 통일)
    _BODY_GRAY = RGBColor(0x59, 0x59, 0x59)
    for _b in parts["bodies"]:
        if _b is not None and _b.has_text_frame:
            clean_bullet(_b.text_frame)
            for _p in _b.text_frame.paragraphs:
                for _r in _p.runs:
                    _r.font.color.rgb = _BODY_GRAY

    # ── 사용자 요청: 하늘색 부제 '전부' 삭제 (제목 GROUP 3개는 모두 유지) ──
    #   유지: 제목 GROUP(제목+체크마크), 본문, 육각형
    for sh in parts["subs"]:
        if sh is None:
            continue
        print(f"[build_slide_2] 부제 삭제: name='{sh.name}'")
        sh._element.getparent().remove(sh._element)

    # 공통 푸터 (사업명 | 순번)
    add_footer(slide, page_num, business_name=business_name or None)

    return slide


# ════════════════════════════════════════════════════════
# 슬라이드 5 빌더: 1.1 본건 사모사채 개요
# ════════════════════════════════════════════════════════

def build_slide_5_sasae_overview(prs, data: dict,
                                  business_name: str = "",
                                  page_num: int = 6):
    """
    슬라이드 5 (1.1 본건 사모사채 개요) PPT 슬라이드를 생성합니다.
    레이아웃.pptx index=5 템플릿의 9r×2c TABLE 을 그대로 활용합니다.

    Parameters
    ----------
    prs           : create_presentation_from_template() 로 만든 Presentation 객체
    data          : generate_sasae_overview() 반환 data dict
    business_name : 하단 사업명 푸터
    page_num      : 슬라이드 번호
    """
    # skip_graphic_frames=False → 9r×2c TABLE 포함 그대로 복제
    # ★"bond_overview"(=레이아웃.pptx 5번) 전용 키를 사용한다.
    #   "content"(=9번 빈 본문 템플릿)에는 표가 없어 _fill_sasae_table 이 실패한다.
    slide = clone_slide_layout(prs, "bond_overview", skip_graphic_frames=False)

    intro  = data.get("intro_paragraph", "")
    fields = data.get("fields", [])

    # 인트로 텍스트 (shape[03], T=2.29)
    intro_sh = _find_shape_by_pos(slide, 1.09, 2.29)
    if intro_sh:
        _replace_text_keep_runs(intro_sh.text_frame, intro)
        # ★제목의 내용(인트로)은 항상 9pt 고정(다른 페이지와 통일 — 이 페이지만 10pt이던 문제)
        for _p in intro_sh.text_frame.paragraphs:
            _p.font.size = Pt(9)
            for _r in _p.runs:
                _r.font.size = Pt(9)

    # 기존 9r×2c TABLE 의 값 열(col 1)만 교체
    _fill_sasae_table(slide, fields)

    # ── 표 서식 표준 적용 (일반표 10.5pt) ──
    #   구분열(col0)=Bold, 세부내용(col1)=Light, 가운데정렬셀 여백0·좌측정렬셀 marL=0.3cm
    #   (헤더 행 없는 레이블-값 표 → 색은 기존 무채움 유지)
    sasae_tbl = _find_table_by_pos(slide, 1.09, 3.84, tol_cm=0.40)
    if sasae_tbl is not None:
        style_table(sasae_tbl, nested=False, has_header=False, label_cols=(0,),
                    label_fill=PALETTE["label_gray"])   # 구분열(col0) 밝은회색 F2F2F2 (기초자산개요와 통일)

    # 공통 푸터 (사업명 | 순번)
    add_footer(slide, page_num, business_name=business_name or None)

    return slide


# ════════════════════════════════════════════════════════
# 슬라이드 7 빌더: 2.1 투자구조도
# ════════════════════════════════════════════════════════

def build_slide_7_investment_structure(prs, data: dict,
                                        business_name: str = "",
                                        page_num: int = 8):
    """
    슬라이드 7 (2.1 투자구조도) PPT 슬라이드를 생성합니다.

    [STEP 5-4-H 신규 방식]
    레이아웃 템플릿 다이어그램(박스/표/화살표/이미지)을 전부 버리고,
    헤더·푸터·배경·인트로만 남긴 "빈 베이스" 를 만든 뒤,
    PDF 4번 사진 좌표대로 박스·화살표를 코드로 직접 그립니다.

    STEP 1 (현재): 베이스만 구성 — 아래 6개 shape 만 유지하고 나머지 전부 삭제.
      유지:
        - [0]  "02 금융 개요"  상단 라벨        (L=1.09, T=1.11)
        - [3]  "2.1 투자구조도" 메인 타이틀      (L=1.09, T=1.64)
        - [16] 인트로 텍스트                    (L=1.09, T=2.29)
        - [1]  상단 구분선(배경)                (L=1.09, T=3.87)
        - [2]  하단 구분선(배경)                (L=1.09, T=18.16)
        - [28] 페이지번호(우하단)               (L=18.53, T=18.35)
      RAINFIELD 로고는 슬라이드 마스터/레이아웃 배경에 있어 자동 유지.

    STEP 2~3 (승인 후 예정): add_shape / add_connector 로 박스 4개 + 화살표 4개 직접 그림.

    Parameters
    ----------
    prs           : create_presentation_from_template() 로 만든 Presentation 객체
    data          : generate_investment_structure() 반환 data dict
    business_name : 하단 사업명 푸터 (이 슬라이드에서는 사용 안 함)
    page_num      : 슬라이드 번호
    """
    slide = clone_slide_layout(prs, "investment_structure")

    intro = data.get("intro_paragraph", "")

    # ── 1. 유지할 shape 좌표 화이트리스트 (cm) ──
    # (left_cm, top_cm, tol_cm)
    # 방식 A: 베이스 6개 + 레이아웃 entity 표 4개(신탁사/차주/시공사/대주)를 보존
    keep_specs = [
        (1.09,  1.11, 0.35),   # [0]  "02 금융 개요"
        (1.09,  1.64, 0.35),   # [3]  "2.1 투자구조도"
        (1.09,  2.29, 0.35),   # [16] 인트로 텍스트
        (1.09,  3.87, 0.35),   # [1]  상단 구분선
        (1.09, 18.16, 0.35),   # [2]  하단 구분선
        (18.53, 18.35, 0.35),  # [28] 페이지번호 (우하단)
        (10.47,  4.98, 0.35),  # [5]  신탁사 표 (2r×1c)
        (10.48,  9.55, 0.35),  # [6]  차주 표   (2r×1c)
        (10.20, 14.26, 0.35),  # [7]  시공사 표 (2r×1c)
        (18.87,  5.44, 0.35),  # [9]  대주 표   (4r×1c → Tr.C 삭제 후 3r)
    ]

    def _is_kept(shape):
        l_cm = shape.left / 360000 if shape.left is not None else -999
        t_cm = shape.top  / 360000 if shape.top  is not None else -999
        for spec_l, spec_t, tol in keep_specs:
            if abs(l_cm - spec_l) <= tol and abs(t_cm - spec_t) <= tol:
                return True
        return False

    # ── 2. 화이트리스트 외 shape 전부 XML 에서 삭제 ──
    to_remove = [sh for sh in slide.shapes if not _is_kept(sh)]
    for shape in to_remove:
        sp = shape._element
        sp.getparent().remove(sp)
    print(f"[build_slide_7] 베이스 정리: {len(to_remove)}개 삭제, "
          f"{len(slide.shapes)}개 유지")

    # ── 3. 인트로 텍스트 교체 ──
    intro_sh = _find_shape_by_pos(slide, 1.09, 2.29)
    if intro_sh:
        _replace_text_keep_runs(intro_sh.text_frame, intro)

    # ── 4. 공통 푸터 (사업명 | 순번) — 기존 페이지번호 placeholder 대체 ──
    add_footer(slide, page_num, business_name=business_name or None)

    # ── STEP 5-2 (방식 A): 레이아웃 entity 표 4개를 PDF p4 좌표로 재배치 ──
    entities = data.get("entities", {})

    # 박스별 헤더색 — 신도림 팔레트로 통일 (PALETTE 상수 사용)
    _H_BORROW  = PALETTE["navy_dark"]    # 08377C 차주
    _H_TRUSTEE = PALETTE["blue"]         # 0063A1 신탁사
    _H_CONSTR  = PALETTE["blue"]         # 0063A1 시공사
    _H_LENDER  = PALETTE["maroon"]       # 8C4A59 대주 (유지)
    _BODY_BLACK = RGBColor(0, 0, 0)

    def _set_cell_text_color(cell, rgb):
        for p in cell.text_frame.paragraphs:
            for r in p.runs:
                r.font.color.rgb = rgb

    def _place_table(orig_l_cm, orig_t_cm):
        return _find_table_by_pos(slide, orig_l_cm, orig_t_cm)

    # 원본 좌표로 4개 표를 먼저 확보 (재배치 전에)
    t_trustee = _place_table(10.47, 4.98)
    t_borrow  = _place_table(10.48, 9.55)
    t_constr  = _place_table(10.20, 14.26)
    t_lender  = _place_table(18.87, 5.44)

    def _layout_box(sh, left_in, top_in, w_in, row_h, header_rgb):
        """표를 PDF p4 좌표/크기로 재배치하고 헤더색·헤더 텍스트(흰색) 적용."""
        if sh is None:
            print("[build_slide_7] 경고: entity 표 못 찾음")
            return
        tbl = sh.table
        sh.left, sh.top = Inches(left_in), Inches(top_in)
        sh.width, sh.height = Inches(w_in), Inches(sum(row_h))
        tbl.columns[0].width = Inches(w_in)
        for i, rh in enumerate(row_h):
            if i < len(tbl.rows):
                tbl.rows[i].height = Inches(rh)
        hc = tbl.cell(0, 0)
        hc.fill.solid()
        hc.fill.fore_color.rgb = header_rgb
        _set_cell_text_color(hc, _C_WHITE)

    # 대주: Tr.C 행(4번째) 삭제 → 3r 유지
    if t_lender is not None:
        trs = t_lender.table._tbl.tr_lst
        if len(trs) >= 4:
            t_lender.table._tbl.remove(trs[3])

    # 재배치 (신탁사/시공사/차주 = 2r, 대주 = 3r)
    _layout_box(t_trustee, 4.5, 2.0, 1.5, [0.40, 0.60],        _H_TRUSTEE)
    _layout_box(t_constr,  1.0, 3.5, 1.5, [0.40, 0.60],        _H_CONSTR)
    _layout_box(t_borrow,  4.5, 3.5, 1.5, [0.40, 0.60],        _H_BORROW)
    _layout_box(t_lender,  8.0, 3.2, 1.5, [0.40, 0.50, 0.50],  _H_LENDER)

    # 본문 셀 텍스트 교체 (회사명, 검정)
    def _set_body(sh, row, text):
        if sh is None:
            return
        cell = sh.table.cell(row, 0)
        _replace_text_keep_runs(cell.text_frame, text)
        _set_cell_text_color(cell, _BODY_BLACK)

    _set_body(t_trustee, 1, entities.get("trustee", "신한자산신탁"))
    _set_body(t_borrow,  1, entities.get("borrower", "더함도시개발"))
    _set_body(t_constr,  1, entities.get("constructor", "포스코이앤씨"))
    # 대주: 헤더 "대주", Tr.A / Tr.B (금액 없음)
    if t_lender is not None:
        _replace_text_keep_runs(t_lender.table.cell(0, 0).text_frame, "대주")
        _set_cell_text_color(t_lender.table.cell(0, 0), _C_WHITE)
        _set_body(t_lender, 1, "Tr.A")
        _set_body(t_lender, 2, "Tr.B")

    # ── STEP 3: 화살표(연결선) + 라벨을 PDF p4 배치대로 그림 ──
    _FONT_BOLD = "피플폰트 Bold"   # 라벨 폰트 (레이아웃 따름)
    _C_ARROW = RGBColor(0x40, 0x40, 0x40)   # 진회색

    def _arrow(x1, y1, x2, y2, *, double, bent=False, adj1=None):
        """연결선. double=True 양 끝 화살촉, False 끝점만.
        bent=True 면 꺾인선(ELBOW, prst=bentConnector3), 기본은 직선(line).
        adj1(0~100000) 지정 시 꺾이는 지점을 조절(수평 구간 길이)."""
        cxn = slide.shapes.add_connector(
            MSO_CONNECTOR.ELBOW if bent else MSO_CONNECTOR.STRAIGHT,
            Inches(x1), Inches(y1), Inches(x2), Inches(y2),
        )
        cxn.line.color.rgb = _C_ARROW
        cxn.line.width = Pt(1.5)
        cxn.shadow.inherit = False
        ln = cxn.line._get_or_add_ln()
        if double:
            ln.append(ln.makeelement(
                _qn('a:headEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
        ln.append(ln.makeelement(
            _qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
        # 꺾임 지점 조절 (bentConnector3 의 adj1)
        if bent and adj1 is not None:
            prstGeom = cxn._element.find('.//' + _qn('a:prstGeom'))
            if prstGeom is not None:
                for av in prstGeom.findall(_qn('a:avLst')):
                    prstGeom.remove(av)
                avLst = etree.SubElement(prstGeom, _qn('a:avLst'))
                gd = etree.SubElement(avLst, _qn('a:gd'))
                gd.set('name', 'adj1')
                gd.set('fmla', 'val %d' % int(adj1))
        return cxn

    def _arrow_label(left, top, w, h, text):
        """화살표 옆 라벨 텍스트박스 (맑은 고딕 9pt, 검정, 배경 투명)."""
        tb = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        tf.margin_left = tf.margin_right = Pt(1)
        tf.margin_top = tf.margin_bottom = Pt(0)
        p = tf.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        run = p.add_run()
        run.text = text
        run.font.name = _FONT_BOLD      # 화살표 라벨도 "피플폰트 Bold" (레이아웃 따름)
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0, 0, 0)
        return tb

    def _freeform_arrow(points_in):
        """꺾는 점(인치 좌표 리스트)을 잇는 freeform 꺾은선 + 끝점 화살촉.
        선색 404040, 두께 1.5pt, 채움 없음, tailEnd=triangle."""
        pts = [(int(x * 914400), int(y * 914400)) for (x, y) in points_in]
        fb = slide.shapes.build_freeform(pts[0][0], pts[0][1], scale=1.0)
        fb.add_line_segments(pts[1:], close=False)
        shp = fb.convert_to_shape()
        shp.fill.background()              # 채움 없음(열린 선)
        shp.line.color.rgb = _C_ARROW
        shp.line.width = Pt(1.5)
        shp.shadow.inherit = False
        ln = shp.line._get_or_add_ln()
        ln.append(ln.makeelement(          # 직각 모서리(round join 제거)
            _qn('a:miter'), {'lim': '800000'}))
        ln.append(ln.makeelement(          # 끝점(P4) 화살촉
            _qn('a:tailEnd'), {'type': 'triangle', 'w': 'med', 'len': 'med'}))
        return shp

    # (1) 신탁사 ↔ 차주 : 세로 양방향 (중심 x=5.25, y 3.0→3.5) — 라벨 "신탁계약"만
    _arrow(5.25, 3.0, 5.25, 3.5, double=True)
    _arrow_label(3.30, 3.05, 1.10, 0.40, "신탁계약")              # 세로선 좌측

    # (2) 시공사 ↔ 차주 : 가로 양방향 (y=4.0, x 2.5→4.5)
    _arrow(2.5, 4.0, 4.5, 4.0, double=True)
    _arrow_label(2.75, 3.62, 1.50, 0.30, "공사도급계약")          # 화살표 위

    # (3) 차주 ↔ 대주 : 가로 양방향 (y=4.0, x 6.0→8.0)  ← L6.0/T4.0/W2.0/H0.0
    _arrow(6.0, 4.0, 8.0, 4.0, double=True)
    _arrow_label(6.25, 3.50, 1.50, 0.30, "대출약정")             # 화살표 위

    # (4) 신탁사 → 대주 : freeform 꺾은선(ㄱ자, miter 직각) — 꺾는 점 직접 고정
    #     P1(6.10,2.55) 신탁사 우변서 0.1" 띄움 → P2(7.50,2.55) 수평 →
    #     P3(7.50,3.55) 수직 꺾음 → P4(7.90,3.55) 대주 좌변서 0.1" 띄움(화살촉)
    #     마지막 수평구간 0.40", 양끝 박스와 0.10" 간격
    _freeform_arrow([(6.10, 2.55), (7.50, 2.55), (7.50, 3.55), (7.90, 3.55)])
    _arrow_label(6.30, 2.45, 1.95, 0.30, "담보신탁 우선수익권")    # 꺾은선 위 중앙

    # ── 박스 글씨 12~13pt + 셀 여백 0 (구조도 고정 규칙) ──
    #   신탁사/차주/시공사 = 13pt, 대주(Tr.A·Tr.B) = 12pt
    #   폰트명/색/정렬은 유지(크기만 변경), 여백 0 으로 공간 확보
    def _set_table_font(sh, size_pt):
        if sh is None:
            return
        tbl = sh.table
        for r in range(len(tbl.rows)):
            for c in range(len(tbl.columns)):
                cell = tbl.cell(r, c)
                cell.margin_left = cell.margin_right = Inches(0)
                cell.margin_top = cell.margin_bottom = Inches(0)
                for p in cell.text_frame.paragraphs:
                    for run in p.runs:
                        run.font.size = Pt(size_pt)

    _set_table_font(t_trustee, 13)
    _set_table_font(t_borrow, 13)
    _set_table_font(t_constr, 13)
    _set_table_font(t_lender, 12)

    # ── 다이어그램 전체를 아래로 이동(세로 중앙 정렬) ──
    #   박스4 + 화살표4 + 라벨4 를 동일 offset(+DY)으로만 이동 → 상대위치·간격 유지
    #   헤더/인트로/구분선/로고/푸터는 이동하지 않음
    _DY = 1.0   # 인치 (헤더-푸터 사이 중앙 정렬용)
    _LBLS = {"신탁계약", "공사도급계약", "대출약정", "담보신탁 우선수익권"}
    for sh in slide.shapes:
        tp = str(sh.shape_type)
        is_diag = (sh.has_table
                   or "Connector" in (sh.name or "")
                   or "FREEFORM" in tp
                   or (sh.has_text_frame and sh.text_frame.text.strip() in _LBLS))
        if is_diag and sh.top is not None:
            sh.top = sh.top + int(_DY * 914400)

    return slide
