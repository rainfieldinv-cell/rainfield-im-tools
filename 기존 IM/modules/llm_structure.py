"""
llm_structure.py
─────────────────────────────────────────────────────────
PDF 한 페이지의 원문 텍스트를 LLM(Claude)이 읽어, 가로 A4 제안서 슬라이드 1장의
구조화된 내용(JSON)으로 재구성한다. 규칙기반 추출이 한국어 IM에서 제목·본문·표를
잘못 뽑는 문제(과분할/제목오인식)를 근본 해결하기 위한 경로.

사람이 만든 대전 제안서([[rainfield-daejeon-skeleton]]) 구성 방식을 따른다:
  섹션라벨 / 소제목 / 인트로(별도 글상자) / 표(라벨-값 또는 그리드) / 불릿 / 출처

claude_api.call_claude()를 그대로 재사용(캐시·재시도·비용로깅·숫자환각검증 내장).
"""

import os
import re

from modules.claude_api import call_claude, verify_numbers_in_pdf

# ★★프롬프트 문구를 고치면 이 버전 문자열을 반드시 함께 올릴 것.
#   캐시 키가 (slide_num, 원문, prompt_version) 로만 만들어져 프롬프트 '내용'은 키에 안 들어간다.
#   버전을 안 올리면 프롬프트를 고쳐도 캐시가 옛 결과를 그대로 돌려줘 아무것도 안 바뀐다.
PROMPT_VERSION = "structure_v8"   # v8: 표 3개·16행 상한 + 중첩표/다단헤더 금지(1장 강제)
                                  # v7: 원본 1페이지=슬라이드 1장(꽉 차면 소폭 압축)
                                  # v6: 표 안의 표(구분 라벨=grid title) 일반화

# 레인필드 표준 4섹션. 섹션라벨은 이 이름으로 표기.
SECTION_NAMES = {1: "사모사채 개요", 2: "금융개요", 3: "본건 사업 개요", 4: "Appendix"}

# 원본이 사모사채·유동화 건이 아닐 때 쓰는 섹션1 이름(담보대출·브릿지 등 일반 대출 건).
#   '사모사채 개요'를 그대로 쓰면 사모사채가 전혀 없는 딜에 그 단어가 목차·섹션라벨·
#   소제목에 박혀 사실과 달라진다.
SECTION1_NAME_LOAN = "거래 개요"
SECTION1_SUBTITLE_LOAN = "본 건 거래 개요"

# 사모사채/유동화 건 판정 키워드 (frame_builders._SASAE_KEYS 와 동일 기준)
_SASAE_KEYS = ("사모사채", "유동화", "수익증권", "자산유동화", "ABCP", "ABL")


def _is_sasae_deal(pages) -> bool:
    """문서 전체 원문에 사모사채·유동화 언급이 있으면 True."""
    for p in pages or []:
        raw = p.get("raw_text") or ""
        if any(k in raw for k in _SASAE_KEYS):
            return True
    return False

SYSTEM_PROMPT = """당신은 부동산 금융 IM(PDF)을 가로 A4 제안서 PPT로 재구성하는 전문 애널리스트입니다.
주어진 PDF '한 페이지'의 원문 텍스트를 읽고, 슬라이드 1장의 구조화된 내용을 순수 JSON으로만 출력하세요.

[제안서는 항상 4개 섹션으로 구성됩니다]
  1 = 사모사채 개요   : Executive Summary, 본건 거래/딜 요약, 핵심 투자포인트
  2 = 금융개요        : 금융조건/대출조건(트랜치·금리·LTV·수수료), 투자구조도, 기초자산, 기한이익상실 등 약정조건
  3 = 본건 사업 개요  : 담보/토지 개요, 사업개요(건축·분양), 사업일정·에쿼티, 사업수지, 입지·시장·분양사례 분석, 차주/시행사 개요·재무
  4 = Appendix        : 현장사진, 시공사 개요, 관계사 개요, 사업계획승인서 등 인허가 첨부, 기타 부록
이 페이지가 어느 섹션에 속하는지 1~4 중 하나로 판단하세요.

[출력 JSON 스키마]
{
  "section_num": 1,
  "subtitle": "이 페이지의 소제목(번호 없이 이름만). 예 '담보토지 개요', '건축개요', '주요 금융조건'.",
  "intro": "본문 도입 1~2문장. 페이지번호·머리말·footer·출처는 절대 포함하지 말 것. 없으면 \\"\\".",
  "bullets": ["표로 만들 수 없는 본문/설명 항목들. 각 항목 한 문장. 없으면 빈 배열."],
  "tables": [
    {
      "title": "표 제목(표 위 작은 라벨). 없으면 \\"\\".",
      "kind": "label_value | grid",
      "header": ["grid일 때만 열 제목 배열. label_value면 빈 배열."],
      "rows": [["..."], ["..."]]
    }
  ],
  "source": "출처 문구. 예 '국토부 실거래가, KB부동산'. 없으면 \\"\\"."
}

[가장 중요 — 원본 1페이지 = 슬라이드 딱 1장]
★★이 페이지의 결과물은 **반드시 슬라이드 1장**에 들어가야 합니다. 여러 장으로 나뉘면 실패입니다.
  (원본 12페이지짜리 IM 이면 본문도 12장이어야 합니다. 한 페이지가 2~4장으로 불어나면 안 됩니다.)
★내용은 **빠짐없이 다 담되**, 한 장에 안 들어갈 만큼 꽉 찬 페이지면 **조금 요약**해서 맞추세요.
  - 긴 서술은 짧게 다듬기(뜻과 숫자는 유지)
  - 성격이 같은 세부 행은 묶기(예: 같은 지번 여러 호실 → 한 행에 묶고 합계 유지)
  - 중복 설명·수식어 삭제
  ★단 아래는 절대 빼지 마세요: 금액·금리·수수료·만기·LTV·감정평가액·상환재원·담보·
    총매출/총사업비/사업이익·세대수·면적·확보율·공정률·분양가·주요 일정, 그리고 모든 합계 행.
★표 규모 기준: 이 페이지 **표 최대 2개**, **모든 표의 행 수 합계 11행 이내**(헤더 포함), 최대 8열.
  (11행이 슬라이드 1장의 물리적 한계다. 넘으면 다음 장으로 밀려 원본 1페이지가 2장이 된다.)
  넘칠 것 같으면 위 방법으로 압축하세요. 숫자를 바꾸거나 지어내지는 마세요.
★표 안의 표(중첩표)는 만들지 마세요 — 별도 표로 펴서 쓰세요(한 장에 안 들어가는 주원인).
★2단(다단) 헤더 금지. 헤더는 한 줄로, 'A - B' 처럼 합쳐 쓰세요.
★한 칸에 100자를 넘기지 마세요.
★본문을 bullets로 "요약"하지 마세요. 요약은 오직 intro(1~2문장)에만. 표로 표현되는 내용은 전부 tables로.
  bullets는 표도 아니고 도입문도 아닌 진짜 설명문/주석만(거의 비어 있어야 정상).

[중요 규칙]
1. section_num은 반드시 1~4 정수. subtitle에는 절대 번호(1.1 등)를 붙이지 말고 이름만(번호는 시스템이 부여).
2. label_value: 좌측 항목명(구분) + 우측 값(내용)인 2열 표. rows=[["대지면적","21,665평"], ...]. header는 빈 배열.
   ★★한 구분(라벨)에 여러 세부항목(•, ①②③, 1)2)3), - 등)이 딸리면 **절대 항목마다 행을 나누지 말고**,
     그 구분을 한 행으로 두고 **모든 세부항목을 내용 칸 하나에 줄바꿈(\\n)으로 이어서** 넣으세요.
     세부항목 번호(①②③)를 구분 칸에 넣지 마세요 — 구분 칸에는 반드시 원문의 항목명(예 '주요 채권보전조치',
     '인출선행조건', '기한이익상실 사유', '대주의 의사결정')을 그대로 쓰고, 내용 칸에 그 아래 모든 줄을 넣습니다.
     예) 원문:  주요 채권보전조치 / • A / • B / • C
        → rows에 ["주요 채권보전조치", "• A\\n• B\\n• C"]  (행 1개. A·B·C로 행 3개로 쪼개지 말 것)
     '1) 당연…', '2) 선택적…' 같은 소제목도 내용 칸 안에 줄바꿈으로 함께 넣으세요(구분 칸으로 빼지 말 것).
3. grid: 다열 표(트랜치·Cash-In/Out·재무표처럼 행·열이 격자인 표). header에 열 제목, rows에 데이터행.
   원문 표의 열 구성을 그대로 따르세요. ★단, '구분|내용' 처럼 좌측이 항목명·우측이 설명인 2열은 grid가 아니라 label_value.
4. ★병합/반복: 원문에서 세로로 같은 값이 연속 반복되면(예 '공동주택'이 여러 행) **첫 행에만 값을 쓰고 그 아래 반복 칸은 빈 문자열("")** 로 두세요(시스템이 세로 병합). 가로도 동일 — 합계 라벨은 첫 칸만, 나머지 빈칸.
5. 원문에 없는 숫자·사실을 만들지 마세요(환각 절대 금지). 있는 내용을 그대로 옮기기만.
6. 페이지 번호('3 / 26', '- 5 -'), 회사 로고, 반복 머리말/꼬리말은 버리세요.
7. 표가 18행 이내면 원문 행을 그대로 유지하세요(임의로 줄이지 말 것).
   ★18행을 넘을 때만 성격이 같은 행을 묶어 압축하고, 합계 행과 핵심 지표는 반드시 남기세요.
     묶을 때 라벨에 무엇을 묶었는지 표시하세요(예 '돈암동 628 (304·101·203호)').
8. 첨부 이미지뿐인 페이지(사업계획승인서 등)는 intro 한 줄만, tables/bullets 비움.
9. ★표 바로 아래/옆의 각주·주석(예 '1)', '2)', '주)', '*', '(단위: 백만원)', '※ …')은
   절대 버리지 말고 그 표 다음 순서로 bullets에 그대로 넣으세요. 각주 번호와 내용을 원문 그대로.
   (표 안 값에 붙은 위첨자 표시 '1)' '2)' 도 값에 그대로 유지)
10. ★표 안의 표(구분 라벨의 내용이 그 자체로 여러 행·열의 표인 경우, 예: '배치 계획'=층별 표,
   'Rack 밀도 및 수량'=분류/개수/IT Load 표, '대출 개요'=트랜치별 표):
   - label_value 표에는 그 '구분 라벨' 행을 만들되 내용칸은 비워두고("")
   - 그 표를 별도 grid 로 만들되 **title 에 그 구분 라벨을 그대로**(예 title:"배치 계획").
   → 시스템이 title 과 같은 라벨의 행 '안'에 그 표를 중첩해 넣습니다(표 안의 표 복원).
순수 JSON만 출력하고 그 외 텍스트·설명은 출력하지 마세요."""


# ═══════════════════════════════════════════════════════════
# 요약본(짧은 버전) 전용 프롬프트
#   사용자 지시(2026-08-07): "짧은 버전은 내용을 다 넣지 말고 핵심만 넣어서 절반 페이지로.
#   표도 원본과 100% 똑같이 안 해도 된다 — 핵심만 보고 정리하면 된다."
#   → 원문 재현(SYSTEM_PROMPT)과 정반대 방침. 표를 핵심 행만 추려 슬라이드 1장에 담는다.
#     (긴 버전이 24장이나 되는 이유가 원문 표를 그대로 재현하다 (1/4)(2/4)로 쪼개지는 것이라,
#      표만 추려도 분할이 사라져 자연히 절반이 된다.)
# ═══════════════════════════════════════════════════════════
SUMMARY_PROMPT_VERSION = "summary_v2"   # v2: 표 합계 12행·불릿 3개로 조임(1장 강제)

SUMMARY_SYSTEM_PROMPT = """당신은 부동산 금융 IM(PDF)을 **요약본 제안서 PPT**로 재구성하는 전문 애널리스트입니다.
주어진 PDF '한 페이지'의 원문을 읽고, **슬라이드 딱 1장 분량**의 핵심만 추려 순수 JSON으로 출력하세요.

[이건 '원문 재현'이 아니라 '요약'입니다]
★표를 원문과 똑같이 재현하지 마세요. 투자 판단에 필요한 핵심 행만 골라 간결한 표로 재구성하세요.
★슬라이드 1장에 반드시 들어가야 합니다. 표가 길어져 다음 장으로 넘어가면 실패입니다.

[제안서는 항상 4개 섹션으로 구성됩니다]
  1 = 거래 개요       : Executive Summary, 본건 거래/딜 요약, 핵심 투자포인트
  2 = 금융개요        : 금융조건/대출조건(트랜치·금리·LTV·수수료), 투자구조도, 기초자산, 약정조건
  3 = 본건 사업 개요  : 담보/토지 개요, 사업개요(건축·분양), 사업일정, 사업수지, 입지·시세, 차주/시행사
  4 = Appendix        : 현장사진, 인허가 첨부, 기타 부록
이 페이지가 어느 섹션에 속하는지 1~4 중 하나로 판단하세요.

[출력 JSON 스키마]
{
  "section_num": 1,
  "subtitle": "이 페이지의 소제목(번호 없이 이름만).",
  "intro": "핵심을 요약한 1~2문장. 페이지번호·머리말·footer·출처 제외. 없으면 \\"\\".",
  "bullets": ["표에 담기 어려운 핵심 설명 항목. 최대 4개, 각 1문장. 없으면 빈 배열."],
  "tables": [
    {
      "title": "표 제목. 없으면 \\"\\".",
      "kind": "label_value | grid",
      "header": ["grid일 때만 열 제목. label_value면 빈 배열."],
      "rows": [["..."], ["..."]]
    }
  ],
  "source": "출처 문구. 없으면 \\"\\"."
}

[요약 규칙 — 반드시 지킬 것]
1. **표는 이 페이지 전체에서 최대 2개**, **최대 6열**,
   그리고 **모든 표의 행 수를 합쳐 10행 이내**(헤더 포함).
   ★입력이 원본 여러 페이지를 이어붙인 것일 수 있습니다. 그래도 결과는 10행 이내입니다.
     페이지가 여러 개면 각 페이지에서 가장 중요한 것만 골라 하나의 표로 통합하세요.
   원문 표가 30행이면 핵심 10행으로 추리세요. 세부 항목은 묶거나 버립니다.
   ★한 페이지에 주제가 여러 개면(예: 토지현황+사업일정+사업수지) 각 주제를 표 하나로
     더 압축하거나, 덜 중요한 주제는 bullets 한 줄로 줄이세요. 표를 3개로 늘리지 마세요.
2. **표 안의 표(중첩표) 금지.** 필요하면 별도 표로 펴거나, 핵심 값만 한 칸에 요약해 넣으세요.
3. **2단(다단) 헤더 금지.** 헤더는 반드시 한 줄. 'A - B' 처럼 합쳐 한 칸으로 쓰세요.
4. **긴 서술은 잘라 쓰세요.** 한 칸에 80자 넘지 않게. 법률 장문은 핵심 조건만.
5. bullets 는 최대 3개, 각 1줄. 원문 각주·주석은 핵심이 아니면 버려도 됩니다.
6. ★★이 페이지의 결과물은 **반드시 슬라이드 딱 1장**에 들어가야 합니다.
   표 12행 + 불릿 3줄이 한계선이라고 생각하고, 넘칠 것 같으면 행을 더 줄이세요.
   여러 장으로 나뉘면 요약 실패입니다.
7. 원문에 '-'(값 없음)로 표기된 칸은 '0' 으로 바꾸지 말고 '-' 그대로 두세요.
8. 항목을 묶어 합산할 때, 합산값이 원문 합계와 어긋나더라도 **원문 숫자를 고치지 마세요.**
   맞추려고 세부 값을 ±조정하지 말고, 원문 값을 그대로 쓰고 합계도 원문 그대로 두세요.

[절대 지킬 것 — 요약해도 이것만은]
★★숫자·금액·비율·날짜는 **원문 그대로**. 반올림·환산·추정 금지. 원문에 없는 숫자 생성 절대 금지.
★★아래 핵심 지표가 원문에 있으면 **반드시 표나 intro에 남기세요**:
   대출·발행 금액 / 금리 / 수수료 / 만기·기간 / LTV / 감정평가액 / 상환재원 / 담보·채권보전 /
   총매출·총사업비·사업이익 / 세대수·연면적·대지면적 / 토지 확보율 / 공정률 / 분양가·시세 / 주요 일정
★고유명사(차주·시행사·시공사·신탁사·대주)는 원문 표기 그대로.
★버린 내용이 있어도 남긴 내용은 반드시 원문과 일치해야 합니다.

[기타]
- section_num 은 1~4 정수. subtitle 에 번호(1.1 등) 붙이지 마세요.
- label_value: 좌측 항목명 + 우측 값 2열 표. header 는 빈 배열.
- grid: 다열 격자 표. header 에 열 제목.
- 세로로 같은 값이 반복되면 첫 행에만 쓰고 아래는 빈 문자열("").
- 페이지 번호·로고·반복 머리말/꼬리말은 버리세요.
- 이미지뿐인 페이지는 intro 한 줄만, tables/bullets 비움.
순수 JSON만 출력하고 그 외 텍스트·설명은 출력하지 마세요."""


def structure_page(raw_text: str, page_num: int, *, use_cache: bool = True,
                   verify_numbers: bool = True, summary: bool = False) -> dict:
    """
    PDF 페이지 원문 → 구조화 dict. 실패 시 None.

    summary=True 면 **요약본 전용 프롬프트**를 쓴다(표를 핵심 행만 추려 1장에 담음).
    prompt_version 이 달라 캐시도 판본별로 따로 쌓인다.

    Returns (성공 시): {
        section_label, subtitle, intro, bullets[], tables[], source,
        "_usage": {...}, "_hallucinated": [...]
    }
    """
    text = (raw_text or "").strip()
    if len(text) < 15:
        return None

    user_prompt = f"[PDF 페이지 {page_num} 원문]\n{text}"
    res = call_claude(
        system_prompt=SUMMARY_SYSTEM_PROMPT if summary else SYSTEM_PROMPT,
        user_prompt=user_prompt,
        slide_num=page_num,
        pdf_context=text,
        use_cache=use_cache,
        prompt_version=SUMMARY_PROMPT_VERSION if summary else PROMPT_VERSION,
    )
    if not res.get("ok") or not res.get("data"):
        return None

    data = res["data"]
    data.setdefault("section_num", None)
    data.setdefault("subtitle", "")
    data.setdefault("intro", "")
    data.setdefault("bullets", [])
    data.setdefault("tables", [])
    data.setdefault("source", "")
    data["_usage"] = res.get("usage")

    if verify_numbers:
        chk = verify_numbers_in_pdf(
            {k: v for k, v in data.items() if not k.startswith("_")}, text)
        data["_hallucinated"] = chk.get("hallucinated_numbers", [])
    return data


# ─────────────────────────────────────────────
# 섹션 분류 fallback (LLM이 section_num 못 주거나 범위밖일 때)
# ─────────────────────────────────────────────
_SEC_KEYWORDS = {
    2: ["금융조건", "대출조건", "투자구조", "구조도", "기초자산", "기한이익",
        "트랜치", "tranche", "ltv", "financing", "금리", "수수료"],
    4: ["appendix", "별첨", "현장사진", "시공사", "관계사", "사업계획승인",
        "실시계획", "인허가", "양해각서", "부록"],
    3: ["담보", "토지", "사업개요", "건축개요", "분양", "사업일정", "에쿼티",
        "equity", "사업수지", "입지", "시장", "사례", "차주", "시행사", "재무"],
    1: ["executive", "summary", "사모사채", "딜", "거래개요", "투자포인트"],
}


def _guess_section(struct: dict) -> int:
    blob = " ".join([
        str(struct.get("subtitle", "")),
        str(struct.get("intro", "")),
        " ".join(struct.get("bullets", []) or []),
        " ".join(t.get("title", "") for t in (struct.get("tables") or [])),
    ]).lower()
    for sec in (2, 4, 3, 1):   # 우선순위
        if any(kw in blob for kw in _SEC_KEYWORDS[sec]):
            return sec
    return 3   # 기본: 본건 사업 개요


def _is_structure_page(p: dict) -> bool:
    """투자구조도 페이지인지(병합 대상에서 제외)."""
    st = p.get("_struct") or {}
    blob = " ".join([
        str(st.get("subtitle", "")),
        " ".join(t.get("title", "") for t in (st.get("tables") or [])),
    ])
    return ("구조도" in blob) or bool(p.get("_invest_image") or p.get("_invest_name"))


def _merge_section2_financing(pages: list, *, debug: bool = False) -> None:
    """원본 IM에서 '주요 금융조건'은 하나의 「구분|내용」 표이지만 PDF가 여러 페이지로
       나뉘어 LLM이 페이지별로 쪼갠다. 사람 제안서처럼 섹션2(투자구조도 제외)의 금융조건
       페이지들을 **하나의 연속 label_value 표 + grid 표들**로 병합한다(소제목 '기초자산 개요').

       - 모든 페이지의 label_value 표 rows를 이어붙여 단일 label_value 표 1개 생성(맨 앞)
       - grid 표(대출조건·Cash-In/Out 등)는 순서대로 뒤에 그대로 유지
       → build_structured_slide가 길면 같은 소제목으로 여러 장에 흘림(사람 제안서와 동일).
    """
    fin = [p for p in pages
           if p.get("_sec_int") == 2 and p.get("_struct") and not _is_structure_page(p)]
    if len(fin) < 2:
        return

    lv_rows, grids = [], []
    for p in fin:
        for t in (p["_struct"].get("tables") or []):
            if (t.get("kind") or "") == "label_value":
                lv_rows.extend(t.get("rows") or [])
            else:
                grids.append(t)

    # ★구분(c0) 라벨 중복 제거 — 여러 금융 페이지에 같은 참여기관(차주/연대보증인/대주/시공사/신탁사)이
    #   반복 수록돼 기초자산 표에 두 번 나오던 문제(천안). 첫 등장만 유지(빈 라벨 행은 보존).
    _seen_lab, _dedup = set(), []
    for r in lv_rows:
        lab = str(r[0] if r else "").strip()
        if lab and lab in _seen_lab:
            continue
        if lab:
            _seen_lab.add(lab)
        _dedup.append(r)
    lv_rows = _dedup

    # ★사업명/시행사·차주 행을 표 맨 위에(원본: 구조도 밑부터 기초자산 표 = 사업명부터 시작).
    #   구조도 페이지(fin에서 제외됨)의 label_value 표에서 값을 가져와 신탁사 앞에 prepend.
    biz_top, _seen = [], set()
    for p in pages:
        for t in ((p.get("_struct") or {}).get("tables") or []):
            if (t.get("kind") or "") != "label_value":
                continue
            for r in (t.get("rows") or []):
                lab = str(r[0] if r else "").strip()
                val = str(r[1] if r and len(r) > 1 else "").strip()
                if not val:
                    continue
                if lab == "사업명" and "사업명" not in _seen:
                    biz_top.append(["사업명", val]); _seen.add("사업명")
                elif "시행사" in lab and "차주" in lab and "시행사" not in _seen:
                    biz_top.append([lab, val]); _seen.add("시행사")
    if biz_top:
        biz_top.sort(key=lambda r: 0 if r[0] == "사업명" else 1)
        # 이미 lv_rows에 같은 라벨 있으면 중복 제거
        _exist = {str(r[0]).strip() for r in lv_rows if r}
        biz_top = [r for r in biz_top if r[0] not in _exist]
        lv_rows = biz_top + lv_rows

    # ── ★표 안의 표: grid(대출조건·Cash 등)를 「구분|내용」 표의 해당 '행 셀' 안에 중첩 ──
    #   행 라벨(앵커)에 grid를 매달고, 렌더러가 그 행의 내용칸 위에 grid를 얹는다.
    def _find_row(sub):
        for i, r in enumerate(lv_rows):
            if sub in str(r[0] if r else ""):
                return i
        return None

    # 병합 페이지들의 각주(주N)) 수집 — grid별로 배정하고 남은 건 표 밑 일반 각주로
    all_notes = []
    for p in fin:
        for b in (p["_struct"].get("bullets") or []):
            if str(b).strip() and b not in all_notes:
                all_notes.append(str(b).strip())
        # ★LLM이 bullets로 안 뺀 경우 대비 — raw_text 에서 각주 정의 직접 추출:
        #   '주N) …', '*…'(별표 주석), 'N) …'(주 없는 번호각주, 다음 연속줄 합침 — 천안 LTV 주1) 2줄).
        _lines = (p.get("raw_text", "") or "").splitlines()
        for li, ln in enumerate(_lines):
            s = ln.strip()
            note = None
            if s.startswith("주") and ")" in s:
                head = s[:s.index(")")].replace("주", "").strip()
                after = s[s.index(")") + 1:].strip()
                if head.isdigit() and len(after) >= 3:
                    note = s
            elif s.startswith("*") and len(s) >= 6:
                note = s
            elif (re.match(r"^\d+\)\s*\S", s) and len(s) >= 8
                  and any(k in s for k in ("감정", "탁감", "LTV", "매출액", "별도", "수용재결 필지", "탁상감정"))):
                # ★'N) …'는 '값 각주'(감정가/LTV/매출액 등)일 때만 — 본문 목록('1) 당연 기한이익…')은 제외.
                note = s               # ★원본 마커 그대로 보존(천안 금융조건은 '1)' — '주' 붙이지 말 것)
                nxt = _lines[li + 1].strip() if li + 1 < len(_lines) else ""
                if nxt and not re.match(r"^[\*\d주■▶•]", nxt) and any(k in nxt for k in ("LTV", "기준", "수용", "탁감")):
                    note += " " + nxt
            if note and note not in all_notes:
                all_notes.append(note)

    def _take_note(keys):
        for i, nt in enumerate(all_notes):
            if any(k in nt for k in keys):
                return all_notes.pop(i)
        return ""

    def _insert_before(anchor, keys):
        """앵커 행이 없으면, keys 중 처음 발견되는 행 앞에(없으면 맨 끝) 새 행 삽입."""
        at = None
        for _k in keys:
            at = _find_row(_k)
            if at is not None:
                break
        lv_rows.insert(at if at is not None else len(lv_rows), [anchor, ""])

    def _insert_after(anchor, key):
        at = _find_row(key)
        lv_rows.insert((at + 1) if at is not None else len(lv_rows), [anchor, ""])

    nested, _used = [], set()   # [(anchor_label, grid_tdef, note), ...]
    for g in grids:
        title = (g.get("title") or "")
        hdr = " ".join(str(h) for h in (g.get("header") or []))
        blob = title + " " + hdr
        if "계좌" in blob and "용도" in hdr:
            # 계좌명|계좌주|용도 → '자금관리' 행(원본: 자금관리 안의 표). 없으면 채무불이행 앞에 신설.
            anchor, note = "자금관리", ""
            if _find_row("자금관리") is None:
                _insert_before("자금관리", ("채무불이행", "대주간"))
        elif "자금사용" in blob or "자금 사용" in blob:
            # 천안: 'Cash-in/out'을 담은 '자금사용용도' = 자금용도(텍스트)와 별개 행 → 대주간 의사결정 뒤 신설.
            anchor, note = "자금사용용도", _take_note(("대여금", "정산 방법", "기투입비용 정산 방"))
            if _find_row(anchor) is None:
                _insert_after(anchor, "대주간 의사결정" if _find_row("대주간 의사결정") is not None else "대주")
        elif any(k in blob for k in ("Cash", "cash", "Cash-In", "Cash-Out")):
            # 대전: Cash-in/out은 기존 '자금용도' 행 안에.
            anchor = next((str(r[0]).strip() for r in lv_rows
                           if "자금" in str(r[0] if r else "") and "용도" in str(r[0])), "자금용도")
            note = _take_note(("대여금", "정산 방법", "기투입비용 정산 방"))
            if _find_row(anchor) is None:
                lv_rows.append([anchor, ""])
        else:
            # 구분|대출금액|금리|LTV → 원본 섹션명(예 '금융 조건') 사용. 괄호 접미사(트랜치별 등)는 제거.
            #   참여기관 바로 뒤(상세 조건 앞)에 신설.
            anchor = re.sub(r"\s*\([^)]*\)\s*$", "", title.strip()) or "주요 대출조건"
            note = _take_note(("탁상감정", "감정가", "감정평가", "경일감정"))
            _star = _take_note(("자문수수료", "제반비용 별도", "수수료 및 대출"))   # *…별도 주석
            if _star:
                note = (_star + "\n" + note).strip() if note else _star
            if _find_row(anchor) is None:
                _insert_before(anchor, ("대출기간", "인출일정", "이자지급", "연체", "상환방"))
        if anchor in _used:        # 같은 행에 표 중복(예 금융조건 표 2개) → 첫 표만
            continue
        _used.add(anchor)
        nested.append((anchor, g, note))

    # ★금융과 무관한 각주(매출액·분양가 등 사업수지 주석)가 LLM 오류로 금융 페이지 bullets에 섞여
    #   표 밑에 잘못 찍히던 문제 → 제외(원본 금융조건 표엔 그런 주석 없음).
    all_notes = [n for n in all_notes
                 if not any(k in str(n) for k in ("매출액", "매출 약", "분양가", "8,545", "8545"))]

    merged_tables = []
    if lv_rows:
        # ★다른 본문 표처럼 미니 라벨('주요 금융조건')을 붙임 → 분할되면 (i/n)도 자동 표시
        _lvt = {"title": "주요 금융조건", "kind": "label_value", "header": [], "rows": lv_rows}
        if all_notes:    # '본 금융조건은 당사자들간…' 등 남은 각주 → 표 바로 밑 9pt
            _lvt["_notes"] = list(all_notes)
        merged_tables.append(_lvt)

    base = fin[0]
    # 빨간 글씨/밑줄 합치기(병합된 페이지들 전부)
    reds, uls, fills = [], [], []
    for p in fin:
        reds.extend(p.get("_red_texts") or [])
        uls.extend(p.get("_underline_texts") or [])
        fills.extend(p.get("_filled_texts") or [])
    base["_red_texts"] = list(dict.fromkeys(reds))
    base["_underline_texts"] = list(dict.fromkeys(uls))
    base["_filled_texts"] = list(dict.fromkeys(fills))
    base["_struct"]["tables"] = merged_tables
    base["_struct"]["_nested_grids"] = nested
    # ★금융조건은 표로만 구성 — 떠다니는 불릿/각주 제거(각주는 표 _notes·grid note로 이미 귀속).
    #   매출액 등 오각주가 표 밑에 한 번 더 찍히던 중복 방지.
    base["_struct"]["bullets"] = []
    base["_struct"]["subtitle"] = "기초자산 개요"
    base["_struct"]["bullets"] = all_notes   # grid에 안 붙은 일반 각주 → 표 바로 밑
    base["_struct"]["source"] = ""           # ★대전 원본엔 출처 없음 → 출처 제거
    # ★기초자산(금융조건)은 표 전용 페이지 — PDF 투자구조도 이미지가 붙어 있으면 제거.
    #   투자구조도는 2.1 페이지에 표·선·글상자로 직접 그리므로 원본 이미지는 불필요(중복).
    base["images"] = []
    for q in fin:
        q["images"] = []
    # 나머지 금융 페이지는 제거(병합됨)
    drop = set(id(p) for p in fin[1:])
    pages[:] = [p for p in pages if id(p) not in drop]
    if debug:
        print(f"  [병합] 금융조건 {len(fin)}페이지 → 1개 표"
              f"(label_value {len(lv_rows)}행 + grid {len(grids)}개)")


_TOC_PLAN_SYS = """당신은 부동산 IM 제안서의 본문(섹션3 '본건 사업 개요' + 섹션4 'Appendix')의 목차를
정리하는 편집자입니다. 본문 소제목 목록(번호=원본 순서)을 받아, 아래 규칙대로 섹션 배정 + 압축을
하여 JSON으로만 출력하세요.

[규칙]
1. 각 소제목을 섹션3 또는 섹션4로 정확히 배정한다.
   - 섹션3(본건 사업 개요): 사업개요·분양개요·사업수지·토지(확보/수용/이용)·사업지 전경·입지/조망 분석·시세/분양사례 비교 등.
   - 섹션4(Appendix): 차주(기업)개요·주주구성/역할·재무제표·별첨·인허가 고시·사업계획승인·양해각서·산업단지 현황 등.
2. 원본 순서를 유지(섞지 말 것). 단 섹션4의 주주구성/주주역할은 기업개요/재무제표와 묶어도 됨.
3. 관련된 연속 소제목을 묶어 **각 섹션을 최대 7개 항목으로** 줄인다(7개 초과 금지). 가능하면 더 적게.
4. 제목 형식 2가지: ① 'A 및 B'(둘을 묶을 때) ② 'X (부가설명)'(부가설명 괄호).
   toc_title 은 짧게 — 괄호 부가설명은 toc_title 에서 빼고, 그 페이지 label 에만 괄호 포함.
   예: '천안 불당신도시 시세결과' → toc_title '시세 결과', 페이지 label '시세 결과 (천안 불당신도시)'.
   ★단, '차주개요'와 '재무제표'를 한 묶음으로 묶을 때 toc_title 은 **'차주 개요'** 로만 한다
     (재무제표는 그 묶음의 한 페이지 label 로). '… 및 재무제표' 식으로 적지 말 것.
5. '면책고지/담당자/연락처' 류는 fixed(목차·번호 제외).
6. 묶음 안 원본 소제목들은 각각 별도 페이지로 남고 label(내용 페이지 미니제목)을 가진다.

[출력 JSON]
{
 "section3": [{"toc_title": "...", "pages": [{"src": "원본소제목", "label": "페이지 미니제목"}]}],
 "section4": [{"toc_title": "...", "pages": [{"src": "원본소제목", "label": "..."}]}],
 "fixed": ["원본소제목", ...]
}"""


def _consolidate_sections(pages: list, *, debug: bool = False) -> None:
    """섹션3·4 항목이 7개를 넘으면(천안형) LLM으로 ①정확한 섹션 재배정 ②각 섹션 7개로 압축.
       7개 이하면(대전형) 아무것도 안 함 → 기존 결과 그대로(대전 보호).
       태그: _grp_sec(재배정 섹션)·_grp_seq(섹션 내 묶음 순번)·_grp_title(목차 제목)·
             _grp_label(내용 페이지 미니라벨, 괄호 포함 가능)·_grp_fixed(연락처 등 제외)."""
    cand = [p for p in pages if p.get("_sec_int") in (3, 4) and isinstance(p.get("_struct"), dict)]
    if not cand:
        return
    n3 = sum(1 for p in cand if p.get("_sec_int") == 3)
    n4 = sum(1 for p in cand if p.get("_sec_int") == 4)
    if n3 <= 7 and n4 <= 7:
        return   # 대전형(양 적음) — 손 안 댐
    cand.sort(key=lambda p: p.get("_orig_idx", 0))
    listing = "\n".join(f"{i + 1}. {str((p.get('_struct') or {}).get('subtitle') or '').strip()}"
                        for i, p in enumerate(cand))
    try:
        res = call_claude(system_prompt=_TOC_PLAN_SYS,
                          user_prompt="[본문 소제목(번호=원본순서)]\n" + listing + "\n\n규칙대로 JSON 출력.",
                          slide_num=901, pdf_context=listing, prompt_version="toc_cap7_v2")
    except Exception as exc:
        if debug:
            print(f"  [목차압축] LLM 실패: {exc}")
        return
    data = res.get("data") if res.get("ok") else None
    if not isinstance(data, dict):
        return

    def _norm(s):
        return re.sub(r"\s+", "", str(s or ""))

    by_sub = {}
    for p in cand:
        by_sub.setdefault(_norm((p.get("_struct") or {}).get("subtitle")), []).append(p)
    used = set()

    def _take(src):
        for p in by_sub.get(_norm(src), []):
            if id(p) not in used:
                used.add(id(p))
                return p
        return None
    for secnum, key in ((3, "section3"), (4, "section4")):
        for gi, g in enumerate(data.get(key) or [], start=1):
            toc = str(g.get("toc_title") or "").strip()
            # ★차주개요+재무제표 묶음은 toc='차주 개요'(LLM이 '… 및 재무제표' 붙이면 정리)
            if "차주" in toc and "재무" in toc:
                toc = "차주 개요"
            for pg in (g.get("pages") or []):
                p = _take(pg.get("src"))
                if p is not None:
                    p["_grp_sec"] = secnum
                    p["_grp_seq"] = gi
                    p["_grp_title"] = toc
                    p["_grp_label"] = str(pg.get("label") or pg.get("src") or "").strip()
    for src in (data.get("fixed") or []):
        p = _take(src)
        if p is not None:
            p["_grp_fixed"] = True
    if debug:
        print(f"  [목차압축] 적용 sec3={len(data.get('section3') or [])} "
              f"sec4={len(data.get('section4') or [])} fixed={len(data.get('fixed') or [])}")


def enrich_and_number(pages: list, *, debug: bool = False, pdf_path: str = None,
                      summary: bool = False) -> list:
    """
    각 페이지를 LLM으로 구조화하고, 레인필드 4섹션 체계에 맞춰
    섹션라벨(고정) + 소제목 번호(x.y)를 부여해 page dict에 채워 넣는다.

    채워 넣는 키:
      page["_struct"]       : LLM 구조화 결과(dict) 또는 None(폴백 대상)
      page["section_num"]   : "01"~"04"
      page["section_name"]  : "섹션명"            (divider 제목용 — 번호 없음)
      page["section_title"] : "0N {섹션명}"        (TOC/섹션 인식 로직 호환)
      page["section_label"] : "0N {섹션명}"        (본문 슬라이드 좌상단)
      page["subtitle"]      : "N.k {소제목}"

    ★ 페이지를 섹션(1~4) 순서로 안정정렬한 뒤 번호를 부여한다.
      → divider가 섹션마다 한 번씩, 01~04 순서로만 생성됨(중복·뒤죽박죽 방지).
    """
    # ── 1패스: 각 페이지 구조화 + 섹션 분류 ──
    for idx, p in enumerate(pages):
        raw = p.get("raw_text", "") or ""
        st = (structure_page(raw, p.get("page_num", 0), summary=summary)
              if raw.strip() else None)
        p["_struct"] = st
        if st:
            sec = st.get("section_num")
            try:
                sec = int(sec)
            except (TypeError, ValueError):
                sec = None
            if sec not in (1, 2, 3, 4):
                sec = _guess_section(st)
        else:
            sec = 4   # 구조화 실패(이미지 첨부 등) → Appendix
        # ★'[별첨 N]' 페이지는 무조건 Appendix(섹션4) — 별첨1(본 PF 주요조건)이 참여기관/금융조건을
        #   담아 섹션2로 잘못 분류돼 기초자산에 흡수·누락되던 문제 방지(머리말에 '별첨'이 있을 때만).
        if "별첨" in (raw[:45] if raw else ""):
            sec = 4
        p["_sec_int"] = sec
        p["_orig_idx"] = idx

    # ★Executive Summary = 금융(섹션2) 첫 페이지 '앞'의 연속 페이지 전부 → 섹션1로 묶음.
    #   (천안: 인근분양사례·일몰/호수 조망 등 Exec 연속 페이지가 섹션3으로 새어 본문에 들어가던 문제.
    #    대전: Exec가 1페이지뿐이라 무영향.)
    _fin_idx = min((p.get("_orig_idx", 9999) for p in pages if p.get("_sec_int") == 2),
                   default=None)
    if _fin_idx is not None and _fin_idx > 0:
        for p in pages:
            if p.get("_orig_idx", 9999) < _fin_idx and p.get("_sec_int") != 2:
                p["_sec_int"] = 1

    # ── 섹션 순서로 안정정렬 (섹션 내부는 원본 순서 유지) ──
    pages.sort(key=lambda p: (p.get("_sec_int", 4), p.get("_orig_idx", 0)))

    # ★섹션1(Executive Summary) 여러 페이지 → 하나로 병합(원본: Exec Summary 1개 항목, 본문 아님).
    #   bullets/표는 합치고, 큰 이미지는 제외(사용자: Exec Summary엔 사진 안 넣음).
    _ex = [p for p in pages if p.get("_sec_int") == 1 and isinstance(p.get("_struct"), dict)]
    if len(_ex) > 1:
        base_ex = _ex[0]
        bst = base_ex["_struct"]
        for q in _ex[1:]:
            qst = q.get("_struct") or {}
            bst.setdefault("bullets", []).extend(qst.get("bullets") or [])
            bst.setdefault("tables", []).extend(qst.get("tables") or [])
            q["_struct"] = None
            q["_drop"] = True           # 병합됐으니 별도 페이지로 렌더 안 함
        base_ex["images"] = []          # Exec Summary는 사진 제외(일몰/호수 조감도 등)
        pages[:] = [p for p in pages if not p.get("_drop")]

    # ── 입지분석 페이지에 '사업개요의 위치도'까지 함께(같은 위치 지도 — 사용자 요청) ──
    #   사업개요 위치도 이미지를 입지분석 페이지 images에 추가(복사) → 입지 = 자기 SITE지도 + 사업개요 위치도.
    _ov_loc = None
    for p in pages:
        sub = str((p.get("_struct") or {}).get("subtitle") or "")
        if ("사업 개요" in sub or "사업개요" in sub) and len(p.get("images") or []) >= 2:
            _ov_loc = (p["images"])[-1]      # 위치도(조감도 다음 = 마지막)
            break
    if _ov_loc is not None:
        for p in pages:
            sub = str((p.get("_struct") or {}).get("subtitle") or "")
            if "입지" in sub and (p.get("images")):
                if _ov_loc not in p["images"]:
                    p["images"] = list(p["images"]) + [_ov_loc]
                break

    # ── 금융조건 = 하나의 표(여러 페이지로 쪼개진 것을 다시 합침) ──
    #   ★단 '원본 1페이지 = 슬라이드 1장' 방침에서는 하지 않는다.
    #     3개 페이지의 표를 하나로 합치면 40행이 넘어 결국 (1/4)~(4/4)로 다시 쪼개져
    #     원본 3페이지가 슬라이드 4장이 된다(장수가 불어나는 주원인).
    if os.environ.get("RAINFIELD_ONE_PAGE_ONE_SLIDE", "1") != "1":
        _merge_section2_financing(pages, debug=debug)
    elif debug:
        print("[enrich] 1페이지=1장 방침 → 섹션2 금융조건 페이지 병합 건너뜀")

    # ── 별첨1(본 PF 주요조건) = 기초자산처럼 통합 렌더: 참여기관(label_value) + '주요 금융조건'(중첩표) ──
    #   금융조건 grid가 별도 표로 쪼개져 한 슬라이드에 (1/2)(2/2) 겹쳐 보이던 문제 해결.
    for p in pages:
        st = p.get("_struct")
        if not isinstance(st, dict):
            continue
        if "PF 주요조건" not in str(st.get("subtitle") or "") and "본PF" not in str(st.get("subtitle") or ""):
            continue
        tabs = st.get("tables") or []
        lv = next((t for t in tabs if (t.get("kind") or "") == "label_value"), None)
        grid = next((t for t in tabs if (t.get("kind") or "") == "grid"
                     and any("Tr" in str(h) for h in (t.get("header") or []))), None)
        if lv and grid:
            lv_rows = [list(r) for r in (lv.get("rows") or [])]
            note = ""
            for npart in list(grid.get("_notes") or []):
                note = (note + "\n" + npart).strip() if note else npart
            # ★raw_text의 '*주N)/주N) …' 각주 중, grid가 참조(비고 등에 '주N)')하는 번호의 정의를 표 밑 각주로.
            #   (LLM이 'LTV 비고=주1)'만 남기고 '주1) 매출액 약 8,545억원 기준' 정의를 누락하던 문제)
            _gblob = (" ".join(str(h) for h in (grid.get("header") or []))
                      + " " + " ".join(str(c) for r in (grid.get("rows") or []) for c in r)).replace(" ", "")
            for _ln in (p.get("raw_text", "") or "").splitlines():
                _s = _ln.strip()
                _m = re.match(r"^\*?\s*주\s*(\d+)\s*\)\s*(.+)$", _s)
                if _m and len(_m.group(2).strip()) >= 3:
                    _defn = _s.lstrip("*").strip()
                    if (f"주{_m.group(1)})" in _gblob) and (_defn not in note):
                        note = (note + "\n" + _defn).strip() if note else _defn
            # ★grid 셀의 '주N) 설명…'은 마커만 남김(설명은 표 밑 각주로 이미 감 — 비고칸 LTV '주1)'만)
            for _gr in (grid.get("rows") or []):
                for _ci in range(len(_gr)):
                    _cv = str(_gr[_ci] or "").strip()
                    _tm = re.match(r"^(\*?\s*주\s*\d+\s*\))\s+\S.+$", _cv)
                    if _tm:
                        _gr[_ci] = re.sub(r"\s+", "", _tm.group(1))
            # ★금융구조도 다이어그램(원본은 표 안 그림) → '금융구조도' 행을 참여기관 뒤·금융조건 앞에
            #   삽입하고 페이지 이미지를 그 행 내용칸에 넣음(표안 사진). p['images']서 빼 중복배치 방지.
            _imgs = p.get("images") or []
            _diag = _imgs[0] if _imgs else None
            if _diag and (_diag.get("width", 0) >= _diag.get("height", 1)):
                lv_rows.append(["금융구조도", ""])
                p["images"] = []
            lv_rows.append(["주요 금융조건", ""])
            st["tables"] = [{"kind": "label_value", "title": "본 PF 주요조건",
                             "header": [], "rows": lv_rows}]
            st["_nested_grids"] = [("주요 금융조건", grid, note)]
            if _diag and not (p.get("images")):
                st["_nested_imgs"] = [("금융구조도", _diag)]

    # ── 각주 복원: raw_text의 '주N)' 각주를 살리고(글머리 X), '출처:' 아닌 표각주는 표 밑으로 ──
    _restore_page_notes(pages)

    # ── 차주개요: '주주구성' + '주주 역할' → '주주구성 및 역할' 한 표(업무분담 열, 그룹 세로병합) ──
    _merge_shareholder_role(pages)

    # ── 토지 확보현황: 확보 > 사유지 > {계약,협의} + 국공유지 3단 구분으로 재구성(원본) ──
    _landuse_3level(pages)

    # ── 같은 표(grid)가 '분석 페이지(지도 있음)'와 '전용 표 페이지(지도 없음)' 양쪽에 중복되면,
    #    분석 페이지에선 표 제거(글+지도만 남김). (천안: 인근시세비교=글/지도, 비교대상 매매사례현황=표) ──
    _grid_loc = {}
    for p in pages:
        for t in ((p.get("_struct") or {}).get("tables") or []):
            if (t.get("kind") or "") == "grid" and str(t.get("title") or "").strip():
                _grid_loc.setdefault((str(t["title"]).strip(), len(t.get("rows") or [])), []).append(p)
    for _plist in _grid_loc.values():
        if len(_plist) < 2:
            continue
        _with_img = [p for p in _plist if p.get("images")]
        _no_img = [p for p in _plist if not p.get("images")]
        if _with_img and _no_img:        # 지도 있는 분석 페이지에서 중복 표 제거(전용 표 페이지엔 유지)
            _keys = {(str(t.get("title") or "").strip(), len(t.get("rows") or []))
                     for q in _no_img for t in ((q.get("_struct") or {}).get("tables") or [])}
            for p in _with_img:
                st = p["_struct"]
                st["tables"] = [t for t in (st.get("tables") or [])
                                if (str(t.get("title") or "").strip(), len(t.get("rows") or [])) not in _keys]

    # ── LLM이 통째 누락한 '비교대상 분양/매매사례 현황' 표를 PDF 좌표로 복원(분석문에 묻힌 표).
    #    ★중복표 제거(위) 뒤에 실행 — 불당/호반 매매사례는 제목·행수가 같지만 다른 표이므로 제거 대상 아님.
    _recover_dropped_comparison_tables(pages, pdf_path)

    # ── 비교대상 사례표 구분열 정규화: 'MAIN SUB'(공백압축형)·'MAIN – SUB' → 'MAIN – SUB' 통일 ──
    #    렌더러 _split_grouped_gubun이 대시(–)로 2단 구분(유닛·평형·가격·매매시세·전세시세)을 병합하게 함.
    _normalize_comparison_gubun(pages)

    # ── 비교대상 분양/매매 사례표에 '조감도' 행(단지별 사진) 추가 ──
    _inject_comparison_thumbnails(pages, pdf_path)

    # ── 투자구조도(2.1): 도형으로 직접 그림. 섹션2 맨 앞에 삽입 → 금융조건은 2.2로 밀림 ──
    try:
        from modules.ai_slide_builders import generate_investment_structure
        # 참여기관·트랜치 정보가 있는 섹션1·2 원문을 모아 데이터 생성
        src = "\n".join(p.get("raw_text", "") or "" for p in pages
                        if p.get("_sec_int") in (1, 2))[:6000]
        data, intro = {}, ""
        if src.strip():
            res = generate_investment_structure(src)
            if res.get("ok") and res.get("data"):
                data = res["data"]
                intro = (data.get("intro_paragraph") or "").strip()
        synth = {"_invest_diagram": data, "_invest_intro": intro,
                 "_invest_name": "투자구조도", "_sec_int": 2, "_struct": None,
                 "raw_text": "", "page_num": -1, "tables": [], "images": []}
        insert_at = next((i for i, p in enumerate(pages)
                          if p.get("_sec_int", 4) == 2), len(pages))
        pages.insert(insert_at, synth)
        if debug:
            print(f"  [투자구조도] 2.1 다이어그램 페이지 삽입(섹션2 맨 앞)")
    except Exception as _exc:
        print(f"[enrich] 투자구조도 생성 실패: {_exc}")

    # ── 섹션3·4가 7개 초과면(천안형) LLM으로 섹션 재배정+7개 압축. 7개 이하면(대전형) 손 안 댐 ──
    _consolidate_sections(pages, debug=debug)
    # 압축이 적용됐으면(_grp_seq 태그 존재) 묶음 순서대로 본문 재정렬(섹션1·2는 그대로)
    if any(p.get("_grp_seq") or p.get("_grp_fixed") for p in pages):
        def _ck(p):
            if p.get("_grp_fixed"):
                return (5, 0, p.get("_orig_idx", 0))
            sec = p.get("_grp_sec") or p.get("_sec_int", 4)
            return (sec, p.get("_grp_seq", 0), p.get("_orig_idx", 0))
        pages.sort(key=_ck)

    # ── 2패스: x.y 번호 부여 (묶음=한 번호 공유; 연락처/면책 고지는 번호·목차 제외) ──
    #   ★섹션1 이름은 거래 성격에 맞춘다. 원본에 사모사채·유동화가 없으면(순수 담보대출 등)
    #     '사모사채 개요' 대신 '거래 개요'를 쓴다 — 없는 상품을 목차·라벨에 박지 않기 위함.
    sec_names = dict(SECTION_NAMES)
    if not _is_sasae_deal(pages):
        sec_names[1] = SECTION1_NAME_LOAN
        print(f"[enrich] 원본에 사모사채·유동화 언급 없음 → 섹션1 이름 "
              f"'{SECTION_NAMES[1]}' → '{SECTION1_NAME_LOAN}'")
    _sec1_sub = ("본 건 사모사채 개요" if sec_names[1] == SECTION_NAMES[1]
                 else SECTION1_SUBTITLE_LOAN)

    counters = {1: 0, 2: 0, 3: 0, 4: 0}
    for p in pages:
        sec = p.get("_grp_sec") or p.get("_sec_int", 4)
        st = p.get("_struct")
        name = (p.get("_invest_name")
                or ((st.get("subtitle") if st else "") or "").strip()
                or sec_names[sec])
        # ★섹션1 Executive Summary → 한글 소제목으로 일관(영어 소제목 금지).
        if "executive summary" in name.lower():
            name = _sec1_sub
            if st is not None:
                st["subtitle"] = name
        p["section_num"] = f"0{sec}"
        p["section_name"] = sec_names[sec]
        p["section_title"] = f"0{sec} {sec_names[sec]}"
        p["section_label"] = f"0{sec} {sec_names[sec]}"
        if p.get("_grp_fixed") or any(kw in name for kw in ("면책", "연락처", "담당자")):
            # 면책고지·담당자 연락처 = 고정 페이지: 번호 안 매기고 목차에서 제외
            p["subtitle"] = name
            p["_no_toc"] = True
        elif p.get("_grp_seq"):
            # 압축된 묶음: 'sec.묶음순번 묶음제목' (같은 묶음 페이지끼리 번호 공유)
            grp = (p.get("_grp_title") or name).strip()
            p["subtitle"] = f"{sec}.{p['_grp_seq']} {grp}"
            lbl = (p.get("_grp_label") or "").strip()   # 내용 페이지 미니라벨(괄호 포함 가능)
            if lbl and lbl != grp and st is not None and st.get("tables"):
                t0 = st["tables"][0]
                if isinstance(t0, dict) and not str(t0.get("title") or "").strip():
                    t0["title"] = lbl
        else:
            counters[sec] += 1
            p["subtitle"] = f"{sec}.{counters[sec]} {name}"
        if st is not None:
            st["_final_subtitle"] = p["subtitle"]
            st["_final_section_label"] = p["section_label"]
        if debug:
            print(f"  p{p.get('page_num')}: {p['section_label']} | {p['subtitle']}")

    # ── 짧은 페이지 합치기(한 슬라이드에 들어갈 만한 연속 섹션3/4 페이지) ──
    _pack_short_pages(pages, debug=debug)
    return pages


def _restore_page_notes(pages: list) -> None:
    """각 페이지 각주 정리:
       ① raw_text의 '주N) …' 각주를 살려, prefix 없이 본문불릿(글머리 •)으로 들어간 것을 복원
          (예 '향후 사업 일정은…' → '주 1) 향후 사업 일정은…' → 9pt 각주로 렌더).
       ② '출처:' 형식이 아닌데 source에 들어간 표 각주(예 '금융감독원 전자공시시스템…')는
          해당 표(_notes)로 옮겨 표 바로 밑 9pt 로 — 출처(좌하단)로 안 가게."""
    for p in pages:
        st = p.get("_struct")
        if not isinstance(st, dict):
            continue
        raw = p.get("raw_text", "") or ""

        # ★분석 불릿('✓ …') verbatim 복원 — LLM이 요약·축약하던 문제(산업단지·시세·입지 등 분석문 전문 유지).
        #   원문은 intro 없이 ✓ 불릿만이므로, 복원 시 중복되는 LLM 요약 intro는 제거.
        if "✓" in raw and isinstance(st.get("bullets"), list):
            _checks, _cur = [], None
            for ln in raw.splitlines():
                s = ln.strip()
                if s.startswith("✓"):
                    if _cur:
                        _checks.append(_cur.strip())
                    _cur = s.lstrip("✓ ").strip()
                elif _cur is not None:
                    if (not s) or s[:1] in "[■▶●▪*(" or s.startswith("구분") or re.match(r"^주\s*\d", s):
                        _checks.append(_cur.strip())
                        _cur = None
                    else:
                        _cur += " " + s
            if _cur:
                _checks.append(_cur.strip())
            _checks = [c for c in _checks if len(c) >= 12]
            if _checks:
                _notes_keep = [b for b in st["bullets"]
                               if str(b).lstrip()[:1] in ("주", "*", "※")
                               or re.match(r"^\s*주\s*\d", str(b))]
                st["bullets"] = _checks + _notes_keep
                st["intro"] = ""        # 요약 intro 제거(원본엔 없음 — ✓ 불릿이 본문)

        # ★'*주N) …' / '주N) …' 단독 각주 정의를 그 주N)를 참조하는 표(_notes)로 귀속(원본처럼 표 밑 9pt).
        #   예 차주개요 '*주1) 1주당 5,000원' → 주주구성(투입금/주식수) 표 밑.
        _tabs = st.get("tables") or []
        for ln in raw.splitlines():
            s = ln.strip()
            m = re.match(r"^\*?\s*주\s*(\d+)\s*\)\s*(.+)$", s)
            if not m or len(m.group(2).strip()) < 2:
                continue
            note_def = s.lstrip("*").strip()
            token = f"주{m.group(1)})".replace(" ", "")
            if any(note_def in (t.get("_notes") or []) for t in _tabs):
                continue
            tgt = None
            for t in _tabs:
                blob = (" ".join(str(h) for h in (t.get("header") or []))
                        + " " + " ".join(str(c) for r in (t.get("rows") or []) for c in r))
                if token in blob.replace(" ", ""):
                    tgt = t
                    break
            if tgt is None and any(k in note_def for k in ("1주당", "주당", "주식", "액면")):
                tgt = next((t for t in _tabs if any(k in
                            (" ".join(str(h) for h in (t.get("header") or [])))
                            for k in ("투입금", "주식수", "주식 수"))), None)
            if tgt is not None:
                tgt.setdefault("_notes", [])
                if note_def not in tgt["_notes"]:
                    tgt["_notes"].append(note_def)

        raw_notes = []
        for ln in raw.splitlines():
            s = ln.strip()
            if s.startswith("주") and ")" in s[:6]:
                head = s[:s.index(")")].replace("주", "").strip()
                after = s[s.index(")") + 1:].strip()
                if head.isdigit() and len(after) >= 3:
                    raw_notes.append(s)

        # ⓪ 출처(source)가 비었으면 raw_text의 '(출처 : …)' 줄에서 복구
        #    (LLM이 인근 분양/시세 표의 '출처 : 국토부실거래가…'를 누락하던 문제)
        if not (st.get("source") or "").strip():
            for ln in raw.splitlines():
                s = ln.strip().lstrip("(").strip()
                if s.startswith("출처"):
                    s = s[2:].lstrip(" :：·").rstrip(")").strip()
                    if s:
                        st["source"] = s
                    break

        def _match_raw(text):
            t = str(text).strip()
            for rn in raw_notes:
                body = rn[rn.index(")") + 1:].strip()
                if t and (t[:12] in body or body[:12] in t):
                    return rn
            return None

        # ① prefix 없는 각주 불릿 복원
        if raw_notes and st.get("bullets"):
            nb = []
            for b in st["bullets"]:
                bs = str(b).strip()
                nb.append(_match_raw(bs) if (not bs.startswith("주") and _match_raw(bs)) else b)
            st["bullets"] = nb

        # ② source에 섞여 들어온 주N)·표각주 분리 — 출처(국토부·KB부동산 등)는 source로,
        #    주N)·재무각주(감독원/전자공시)는 해당 표 _notes(표 밑 9pt)로.
        #    (LLM이 '출처 / 주1) / 주2) / 주3)'을 source 한 칸에 몰아넣던 문제 해결)
        src = (st.get("source") or "").strip()
        if src:
            pieces = [x.strip() for x in re.split(r"\s*/\s*|\n", src) if x.strip()]
            note_parts, src_parts = [], []
            for pc in pieces:
                is_ju = bool(re.match(r"^주\s*\d+\s*\)", pc))
                is_fin = any(k in pc for k in ("재무제표", "감독원", "전자공시"))
                (note_parts if (is_ju or is_fin) else src_parts).append(pc)
            tabs = st.get("tables") or []
            if note_parts and tabs:
                fin_t = None
                for t in tabs:
                    blob = (" ".join(str(h) for h in (t.get("header") or []))
                            + " ".join(str(c) for r in (t.get("rows") or []) for c in r))
                    if any(k in blob for k in ("회계연도", "자산총계", "요약")):
                        fin_t = t
                        break
                for npart in note_parts:
                    # ★LLM이 출처에서 '주1)' prefix를 떼고 '금융감독원…'만 넣은 경우 원문(raw_text)서 prefix 복원
                    if not re.match(r"^\*?\s*주\s*\d+\s*\)", npart):
                        _rn = _match_raw(npart)
                        if _rn:
                            npart = _rn
                    tgt = (fin_t if (fin_t and any(k in npart for k in ("재무", "감독원", "전자공시")))
                           else tabs[-1])
                    tgt.setdefault("_notes", [])
                    if npart not in tgt["_notes"]:
                        tgt["_notes"].append(npart)
            elif note_parts:
                st.setdefault("bullets", []).extend(note_parts)
            st["source"] = " ".join(src_parts).strip()

        # ③ 사업수지: 발코니 확장/상가 행 비고가 비면 비율값으로 채움(원본은 비고=비율)
        for t in (st.get("tables") or []):
            hdr = [str(h or "") for h in (t.get("header") or [])]
            if not any("세부" in h for h in hdr) or not any("비고" in h for h in hdr):
                continue
            try:
                bi = next(i for i, h in enumerate(hdr) if "비율" in h)
                gi = next(i for i, h in enumerate(hdr) if "비고" in h)
                si_ = next(i for i, h in enumerate(hdr) if "세부" in h)
            except StopIteration:
                continue
            for r in (t.get("rows") or []):
                if si_ < len(r) and any(k in str(r[si_]) for k in ("발코니", "상가")):
                    while len(r) <= gi:
                        r.append("")
                    if not str(r[gi]).strip() and bi < len(r) and str(r[bi]).strip():
                        r[gi] = r[bi]

        # ④ Equity 투입현황: 합계 행 에쿼티 칸이 비면 에쿼티 열 합으로 채움(원본은 합계도 표기)
        for t in (st.get("tables") or []):
            hdr = [str(h or "") for h in (t.get("header") or [])]
            ei = next((i for i, h in enumerate(hdr) if "에쿼티" in h), None)
            if ei is None:
                continue
            rows = t.get("rows") or []
            tot_i = next((i for i, r in enumerate(rows) if r and "합계" in str(r[0])), None)
            if tot_i is None:
                continue
            tr = rows[tot_i]
            while len(tr) <= ei:
                tr.append("")
            if str(tr[ei]).strip():
                continue
            ssum, ok = 0, False
            for j, r in enumerate(rows):
                if j == tot_i or ei >= len(r):
                    continue
                v = str(r[ei]).replace(",", "").strip()
                if v.isdigit():
                    ssum += int(v); ok = True
            if ok:
                tr[ei] = f"{ssum:,}"

        # ④-2 수용재결 진행일정(구분|진행일정|관계법령): 진행일정 칸에 섞인 '(…)' 조건문은
        #     관계법령 칸 소속(원본). 진행일정 칸의 '(' 로 시작하는 줄을 관계법령 칸으로 이동.
        for t in (st.get("tables") or []):
            hdr = [str(h or "") for h in (t.get("header") or [])]
            pi = next((i for i, h in enumerate(hdr) if "진행일정" in h), None)
            li = next((i for i, h in enumerate(hdr) if "관계법령" in h), None)
            if pi is None or li is None:
                continue
            for r in (t.get("rows") or []):
                if pi >= len(r):
                    continue
                lines = str(r[pi] or "").split("\n")
                moved = [ln for ln in lines if ln.strip().startswith("(")]
                if not moved:
                    continue
                r[pi] = "\n".join(ln for ln in lines if not ln.strip().startswith("(")).strip()
                while len(r) <= li:
                    r.append("")
                r[li] = (str(r[li] or "").strip() + "\n" + "\n".join(moved)).strip()

        # ⑤ 사업수지(구분|세부항목|세대수/내역|…) 표 정규화 — 사람 제안서 구조와 맞춤
        #   · '총 매출/총 비용' 같은 섹션 합계 라벨이 구분(c0)에 있으면 세부항목(c1)으로 옮김
        #     → 구분(매출/비용)은 그 합계행까지 세로로 한 칸, 합계 라벨은 세부항목+세대수내역만 가로병합
        #   · '소계'가 세부항목(c1)에 있으면 세대수/내역(c2)으로 옮김(세부항목은 카테고리로 병합 유지)
        for t in (st.get("tables") or []):
            hdr = [str(h or "") for h in (t.get("header") or [])]
            if not (len(hdr) >= 3 and "구분" in hdr[0]
                    and ("세부" in hdr[1] or "항목" in hdr[1])):
                continue
            for r in (t.get("rows") or []):
                if len(r) < 3:
                    continue
                c0 = str(r[0] or "").strip()
                c1 = str(r[1] or "").strip()
                c2 = str(r[2] or "").strip()
                # 섹션 합계(총 매출/총 비용 등)가 구분열에 → 세부항목열로 이동, 구분은 비워 윗 그룹에 병합
                if c0 and c0.startswith("총") and not c1 and not c2:
                    r[1] = c0
                    r[0] = ""
                # '소계'가 세부항목열에 → 세대수/내역열로(세부항목은 카테고리 병합 유지)
                elif c1 == "소계":
                    r[1] = ""
                    r[2] = ("소계 " + c2).strip() if c2 else "소계"

        # ⑥ 토지개요: 감정평가금액(980억)은 전체 담보(터미널+A구역) 총 감정가 → 원본은 한 칸으로 두 블록에 걸침.
        #    표가 터미널/A구역 페이지로 나뉘므로 '각 구역 블록 첫 행'(터미널·A구역, 국공유지·합계 제외)에
        #    980억을 넣어 두 페이지 모두 그 열이 병합되어 980억이 보이게 한다.
        for t in (st.get("tables") or []):
            hdr = [str(h or "") for h in (t.get("header") or [])]
            if not (any("소재지" in h for h in hdr) and any("감정평가" in h for h in hdr)
                    and any("구역" in h for h in hdr)):
                continue
            ai = next(i for i, h in enumerate(hdr) if "감정평가" in h)
            gi = next(i for i, h in enumerate(hdr) if "구역" in h)
            rows = t.get("rows") or []
            appraisal = ""
            for r in rows:                      # 기존 감정가 값 추출 + 파셀 행에서 비움
                if ai < len(r):
                    v = str(r[ai]).strip()
                    if v and v != "-" and ("억" in v or v.replace(",", "").isdigit()):
                        appraisal = v
                        r[ai] = ""
            if not appraisal:
                continue
            for r in rows:                      # 구역 시작 행(터미널/A구역)에 표기
                gv = str(r[gi]).strip() if gi < len(r) else ""
                if gv and not any(k in gv for k in ("국공유지", "합계", "총")):
                    while len(r) <= ai:
                        r.append("")
                    r[ai] = appraisal
                    break

        # ⑥-2 '본 금융조건은 당사자들간 협의과정…(변경 가능)' 문구는 표 밑 각주(원본). 셀(기타 등) 안에 있으면 _notes로.
        for _t in (st.get("tables") or []):
            for _r in (_t.get("rows") or []):
                for _ci in range(len(_r)):
                    _v = str(_r[_ci] or "")
                    if "당사자들간 협의" not in _v and "심의 과정에서 변경" not in _v:
                        continue
                    _ls = _v.split("\n")
                    _mv = [x for x in _ls if ("당사자들간 협의" in x or "심의 과정에서 변경" in x)]
                    _r[_ci] = "\n".join(x for x in _ls if x not in _mv).strip()
                    _t.setdefault("_notes", [])
                    for _m in _mv:
                        _mm = _m.strip().lstrip("•·-").strip()
                        _mm = "- " + _mm
                        if _mm not in _t["_notes"]:
                            _t["_notes"].append(_mm)

        # ⑥-3 표 _notes 정리: '주 N)'→'주N)' 정규화 후 같은 번호 중복 제거(긴 본문 유지)·번호순 정렬.
        #     (분양사례 표의 셀줄바꿈 잔재 '주 3)'가 '주3)'와 따로 잡혀 중복·역순으로 뜨던 문제)
        for _t in (st.get("tables") or []):
            _nts = _t.get("_notes") or []
            if len(_nts) < 2:
                continue
            _other, _byn = [], {}
            for _n in _nts:
                _s = re.sub(r"^(\*?)\s*주\s*(\d+)\s*\)", r"\1주\2)", str(_n).strip())
                _mn = re.match(r"^\*?주(\d+)\)", _s)
                if not _mn:
                    if _s not in _other:
                        _other.append(_s)
                    continue
                _k = _mn.group(1)
                if _k not in _byn or len(_s) > len(_byn[_k]):
                    _byn[_k] = _s
            _t["_notes"] = _other + [_byn[k] for k in sorted(_byn, key=int)]

        # ⑦ 주N) 각주 중복 제거: bullets(좌하단)와 표 _notes 양쪽에 같은 번호가 있으면, 더 긴 버전을
        #    표 _notes에 남기고 bullets에서는 제거(원본처럼 표 밑에만 1번 — 토지확보 주1/2/3 중복 방지).
        def _jnum(s):
            m = re.match(r"^\*?\s*주\s*(\d+)", str(s).lstrip())
            return m.group(1) if m else None
        _bl = st.get("bullets") or []
        _bl_by = {}
        for _b in _bl:
            _k = _jnum(_b)
            if _k:
                _bl_by[_k] = _b
        if _bl_by:
            _tab_nums = set()        # 표 _notes에 이미 있는 주N) 번호(여기 있는 것만 bullets에서 제거)
            for _t in (st.get("tables") or []):
                _nts = _t.get("_notes") or []
                if not _nts:
                    continue
                for _n in _nts:
                    _k2 = _jnum(_n)
                    if _k2:
                        _tab_nums.add(_k2)
                _t["_notes"] = [(_bl_by[_jnum(_n)] if (_jnum(_n) in _bl_by
                                 and len(str(_bl_by[_jnum(_n)])) > len(str(_n))) else _n)
                                for _n in _nts]
            # ★표 _notes에 '있는 번호'만 bullets에서 제거(중복). 표에 없는 주N)는 bullets에 그대로 유지.
            st["bullets"] = [_b for _b in _bl if _jnum(_b) is None or _jnum(_b) not in _tab_nums]

        # ⑧ 토지 확보현황 표: 각주를 원본 raw '*…' 줄 그대로 사용(LLM 확장본 대신).
        #    원본엔 '*실시계획인가 …'(편입면적 기준) + '*주1)/주2)/주3)' 4줄 — verbatim 보존(사용자 지시).
        _star = [ln.strip() for ln in raw.splitlines()
                 if ln.strip().startswith("*") and len(ln.strip()) >= 6]
        if _star:
            for _t in (st.get("tables") or []):
                _hdr = " ".join(str(h) for h in (_t.get("header") or []))
                if "확보비율" in _hdr and "구역면적" in _hdr:
                    _t["_notes"] = list(_star)


def _recover_comparison_table_by_pos(page):
    """PDF 좌표로 '비교대상 분양/매매사례 현황' 표를 복원(LLM이 통째 누락한 경우).
       단지명(헤더 줄바꿈) 결합, 2단 구분(유닛/평형/매매시세/전세시세)을 'MAIN – SUB'로 재구성."""
    def _clean(s):
        s = re.sub(r"\s+", " ", s).strip()
        return re.sub(r"(\d)\s+(위|층|개동|단지|억원|억|평|%)", r"\1\2", s)
    words = page.get_text("words")
    ty = next((w[1] for w in words if "현황]" in w[4] or "비교대상" in w[4]), None)
    if ty is None:
        return None
    fy = next((w[1] for w in words if w[4] == "페이지" and w[1] > ty + 100), ty + 440)
    tw = [w for w in words if ty + 3 < w[1] < fy - 2]
    tw.sort(key=lambda w: w[1])
    rows, cur, cy = [], [], None
    for w in tw:
        if cy is None or abs(w[1] - cy) <= 4:
            cur.append(w); cy = cy or w[1]
        else:
            rows.append(cur); cur = [w]; cy = w[1]
    if cur:
        rows.append(cur)
    def _rl(r):
        return "".join(w[4] for w in r if w[0] < 150)
    # 칸 중심 = '세대수' 행(칸당 순수 숫자 1개)의 값 x → 지도 라벨 오염 없이 정확. 없으면 숫자 최다 행.
    def _numw(r):
        return [w for w in r if w[0] > 150 and re.match(r"^[\d,]+$", w[4])]
    ref = next((r for r in rows if "세대수" in _rl(r)), None)
    if ref is None:
        ref = max(rows, key=lambda r: len(_numw(r))) if rows else []
    cen = sorted(w[0] for w in _numw(ref))
    if len(cen) < 2:
        return None
    ncol = len(cen)
    b0 = min(cen) - 30

    def colof(x):
        return min(range(ncol), key=lambda i: abs(x - cen[i]))   # 최근접 칸(밴드 경계 오배치 방지)

    bstart = next((i for i, r in enumerate(rows) if any(k in _rl(r) for k in ("조감도", "주소"))), 1)
    # 헤더(단지명): 제목~본문 사이 모든 줄(줄바꿈된 이름 전부)을 칸별 결합(최근접)
    hdr = ["구분"] + [""] * ncol
    for w in sorted([w for r in rows[:bstart] for w in r if w[0] >= b0], key=lambda w: (w[1], w[0])):
        ci = 1 + colof(w[0]); hdr[ci] = (hdr[ci] + " " + w[4]).strip()
    hdr = [_clean(re.sub(r"(써밋)\s+(플레이스)", r"\1\2", h)) for h in hdr]
    AP, APc, PX = ["유닛 구성", "평형"], ["유닛", "평형"], ["매매시세", "전세시세"]
    ap_i = px_i = 0
    out = []
    for r in rows[bstart:]:
        rs = sorted(r, key=lambda w: w[0])
        vals = {}
        for w in rs:
            if w[0] >= b0:
                vals.setdefault(colof(w[0]), []).append(w[4])
        if not any(w[0] < b0 for w in rs):       # 라벨 없는 줄 = 값 줄바꿈 연속 → 직전 행에 병합
            if out:
                for i in range(ncol):
                    if vals.get(i):
                        out[-1][1 + i] = _clean((out[-1][1 + i] + " " + " ".join(vals[i])).strip())
            continue
        sub = " ".join(w[4] for w in rs if 70 <= w[0] < b0).strip()
        vlist = [_clean(" ".join(vals.get(i, []))) for i in range(ncol)]
        if "조감도" in sub:                       # 조감도 행은 썸네일 주입기가 추가
            continue
        if not sub and not any(vlist):
            continue
        if sub in ("공급면적", "공급"):
            ap_i += 1; main = (AP if sub == "공급면적" else APc)[min(ap_i - 1, 1)]
        elif sub in ("전용면적", "전용"):
            main = (AP if sub == "전용면적" else APc)[min(max(ap_i - 1, 0), 1)]
        elif sub == "세대당가격":
            px_i += 1; main = PX[min(px_i - 1, 1)]
        elif sub in ("공급평당가", "가격출처"):
            main = PX[min(max(px_i - 1, 0), 1)]
        elif sub in ("세대", "평당"):
            main = "가격"
        else:
            main = ""
        out.append([(main + " – " + sub) if main else sub] + vlist)
    return hdr, out


def _recover_dropped_comparison_tables(pages: list, pdf_path: str) -> None:
    """raw_text엔 '[비교대상 …현황]'이 있는데 구조화 결과엔 그 표가 없는(LLM 누락) 페이지를
       PDF 좌표로 복원해 추가(천안 불당 매매사례표 — 분석문에 묻혀 통째 누락되던 문제)."""
    if not pdf_path:
        return
    try:
        import fitz
    except Exception:
        return
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return
    _inserts = []   # (원본 index, 새 표 전용 페이지) — 분석/지도 페이지면 표를 별도 슬라이드로
    for idx, p in enumerate(pages):
        st = p.get("_struct")
        if not isinstance(st, dict):
            continue
        raw = p.get("raw_text", "") or ""
        if "비교대상" not in raw or "현황" not in raw:
            continue
        _exist = [t for t in (st.get("tables") or [])
                  if ("사례" in str(t.get("title") or "") and "현황" in str(t.get("title") or ""))]
        if len(_exist) >= 2:
            continue                    # 한 페이지 표 2개(분양 신규/기준공)는 좌표매칭 모호 → 건너뜀
        pi = (p.get("page_num") or 0) - 1
        if pi < 0 or pi >= doc.page_count:
            continue
        try:
            rec = _recover_comparison_table_by_pos(doc[pi])
        except Exception:
            rec = None
        if not rec:
            continue
        hdr, body = rec
        if len(hdr) < 3 or len(body) < 4:
            continue
        if _exist:
            # 표는 있으나 LLM이 단지명(헤더)을 뒤섞어 추출 → 좌표복원 헤더로 이름만 교체(본문 유지)
            t = _exist[0]
            if len(t.get("header") or []) == len(hdr):
                t["header"] = hdr
            continue
        m = re.search(r"비교대상\s*(분양사례|매매사례)\s*현황", raw)
        title = "비교대상 " + (m.group(1) if m else "분양사례") + " 현황"
        new_tbl = {"kind": "grid", "title": title, "header": hdr, "rows": body}
        # ★분석문/지도가 같이 있는 페이지면 표는 '별도 슬라이드'로(원본처럼 표 안 짤리게 — 이어서 한 장에).
        if st.get("bullets") or p.get("images"):
            _np = {
                "page_num": p.get("page_num"),
                "raw_text": "", "images": [], "_sec_int": p.get("_sec_int"),
                # ★정렬키 복사 + _orig_idx 살짝 뒤로 → 후속 정렬(_ck)에서 분석페이지 바로 뒤에 위치
                "_orig_idx": (p.get("_orig_idx", 0) or 0) + 0.5,
                "_struct": {"subtitle": st.get("subtitle"),
                            "section_label": st.get("section_label"),
                            "intro": "", "bullets": [], "source": "",
                            "tables": [new_tbl]},
            }
            for _k in ("_grp_sec", "_grp_seq", "_grp_fixed"):
                if p.get(_k) is not None:
                    _np[_k] = p.get(_k)
            _inserts.append((idx, _np))
        else:
            st.setdefault("tables", []).append(new_tbl)
    for _off, (idx, np_) in enumerate(_inserts):
        pages.insert(idx + 1 + _off, np_)


def _landuse_3level(pages: list) -> None:
    """토지 확보현황 표를 원본처럼 3단 구분으로 재구성:
       확보 > 사유지 > {계약 체결, 협의 완료} + 국공유지 / 소계·수용·합계는 구분 전체 가로병합.
       (현재는 확보 > {계약,협의,국공유지} 2단 → 사유지 중간단 삽입)."""
    for p in pages:
        st = p.get("_struct")
        if not isinstance(st, dict):
            continue
        for t in (st.get("tables") or []):
            hdr = [str(h or "") for h in (t.get("header") or [])]
            if not (any("확보비율" in h for h in hdr) and any("구역면적" in h for h in hdr)):
                continue
            if len(hdr) < 3 or hdr[1].strip() or not hdr[2].strip():
                continue                       # 2단 구분(구분|빈|구역면적…)만 대상(이미 3단이면 skip)
            new_rows, saw_private = [], False
            for r in (t.get("rows") or []):
                c0 = str(r[0] or "").strip()
                c1 = str(r[1] or "").strip()
                rest = list(r[2:])
                if any(k in c1 for k in ("계약", "협의")):       # 사유지 멤버
                    sa = "" if saw_private else "사유지"
                    saw_private = True
                    new_rows.append([c0, sa, c1] + rest)
                elif "국공유지" in c1:                          # 사유지와 동급(col1), 계약칸까지 가로병합
                    new_rows.append([c0, c1, ""] + rest)
                elif "소계" in c1:                              # 소계 → 구분 col0(전체 가로병합)
                    new_rows.append([c1, "", ""] + rest)
                else:                                           # 수용/합계(라벨 c0) 등
                    new_rows.append([c0, "", ""] + rest)
            t["header"] = [hdr[0], "", ""] + hdr[2:]
            t["rows"] = new_rows


def _merge_shareholder_role(pages: list) -> None:
    """차주개요: '주주구성'(주주명·주식수·지분율·투입금·비고) + '주주 역할'(SI/CI/FI 업무분담)
       → '주주구성 및 역할' 한 표로 병합. 업무분담 열을 비고 앞에 삽입, 주주명→그룹 매칭으로 채움
       (그룹 첫 주주만 채워 나머지는 빈칸 → 렌더러가 세로병합). 주주역할 표는 제거(사용자 지시)."""
    def _norm(s):
        return re.sub(r"[㈜()\s,./]", "", str(s or ""))
    for p in pages:
        st = p.get("_struct")
        if not isinstance(st, dict):
            continue
        tabs = st.get("tables") or []

        def _hblob(t):
            return " ".join(str(h) for h in (t.get("header") or []))
        comp = next((t for t in tabs if t.get("kind") == "grid"
                     and "주주명" in _hblob(t) and "업무" not in _hblob(t)), None)
        role = next((t for t in tabs if ("역할" in str(t.get("title") or ""))
                     or ("업무" in _hblob(t) or "분담" in _hblob(t))), None)
        if not comp or not role or comp is role:
            continue
        rhdr = [str(h) for h in (role.get("header") or [])]
        mi = next((i for i, h in enumerate(rhdr) if "주주" in h), 1 if len(rhdr) >= 3 else 0)
        di = next((i for i, h in enumerate(rhdr) if ("업무" in h or "분담" in h)), len(rhdr) - 1)
        groups = []
        for r in (role.get("rows") or []):
            if di < len(r):
                members = [_norm(m) for m in re.split(r"[,/]", str(r[mi] if mi < len(r) else "")) if m.strip()]
                groups.append((members, str(r[di])))
        chdr = [str(h) for h in (comp.get("header") or [])]
        bi = next((i for i, h in enumerate(chdr) if "비고" in h), len(chdr))
        new_hdr = chdr[:bi] + ["업무분담"] + chdr[bi:]
        used, new_rows = set(), []
        for r in (comp.get("rows") or []):
            name = _norm(r[0] if r else "")
            duty = ""
            if any(k in str(r[0] if r else "") for k in ("합계", "소계", "총계")):
                duty = "-"          # 합계행은 '-'(위 그룹 업무분담이 합계까지 세로병합되는 것 방지)
            else:
                for gi, (members, dtext) in enumerate(groups):
                    if gi in used:
                        continue
                    if any(m and (m in name or name in m) for m in members):
                        duty = dtext
                        used.add(gi)
                        break
            new_rows.append(list(r[:bi]) + [duty] + list(r[bi:]))
        comp["header"] = new_hdr
        comp["rows"] = new_rows
        comp["title"] = "주주구성 및 역할"
        st["tables"] = [t for t in tabs if t is not role]


_COMP_MAINS = ["유닛 구성", "매매시세", "전세시세", "유닛", "평형", "가격"]


def _normalize_comparison_gubun(pages: list) -> None:
    """비교대상 분양/매매사례 현황 표 구분열(c0)을 'MAIN – SUB' 형으로 통일.
       LLM이 '유닛 공급'(공백압축)·'유닛 구성 – 공급면적'(대시) 등 제각각으로 줘서 2단 병합이 안 되던 문제.
       알려진 MAIN(유닛/평형/가격/매매시세/전세시세)으로 시작하면 대시로 분리 → 렌더러가 세로병합."""
    for p in pages:
        st = p.get("_struct")
        if not isinstance(st, dict):
            continue
        for t in (st.get("tables") or []):
            title = str(t.get("title") or "")
            if "사례" not in title or "현황" not in title:
                continue
            for r in (t.get("rows") or []):
                if not r:
                    continue
                c0 = str(r[0] or "").strip()
                if not c0 or any(d in c0 for d in ("–", "—")):
                    continue            # 이미 대시형
                for mn in _COMP_MAINS:
                    if c0 != mn and c0.startswith(mn + " "):
                        r[0] = f"{mn} – {c0[len(mn):].strip()}"
                        break


def _inject_comparison_thumbnails(pages: list, pdf_path: str) -> None:
    """비교대상 분양사례/매매사례 현황 표에 '조감도' 행(단지별 작은 사진)을 추가.
       PyMuPDF로 페이지 썸네일들을 y행 그룹·x정렬로 추출 → 표의 데이터 열 수와 매칭해
       g['_thumb_imgs'](왼→오 순서 png bytes)로 저장하고 맨 위에 '조감도' 빈 행을 넣는다."""
    if not pdf_path:
        return
    try:
        import fitz
    except Exception:
        return
    from collections import defaultdict
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return
    for p in pages:
        st = p.get("_struct")
        if not isinstance(st, dict):
            continue
        grids = [t for t in (st.get("tables") or [])
                 if (t.get("kind") or "") == "grid"
                 and any(k in str(t.get("title") or "") for k in ("분양사례 현황", "매매사례 현황"))]
        if not grids:
            continue
        pi = (p.get("page_num") or 0) - 1
        if pi < 0 or pi >= doc.page_count:
            continue
        try:
            info = doc[pi].get_image_info(xrefs=True)
        except Exception:
            continue
        rows = defaultdict(list)
        for im in info:
            b = im.get("bbox") or (0, 0, 0, 0)
            rows[round(b[1] / 20)].append((b[0], im.get("xref")))
        photo_rows = []
        for y in sorted(rows):
            cells = sorted(rows[y])
            if len(cells) < 3:                  # 조감도 행으로 볼 만한 것(사진 3장+)
                continue
            imgs = []
            for _x, xref in cells:
                try:
                    pix = fitz.Pixmap(doc, xref)
                    if (pix.n - pix.alpha) > 3:
                        pix = fitz.Pixmap(fitz.csRGB, pix)
                    imgs.append(pix.tobytes("png"))
                except Exception:
                    imgs.append(None)
            photo_rows.append(imgs)
        if not photo_rows:
            continue
        used = set()
        for g in grids:
            ncol = max([len(g.get("header") or [])]
                       + [len(r) for r in (g.get("rows") or [])] + [0])
            if ncol < 2:
                continue
            data_cols = ncol - 1
            mi = next((i for i, pr in enumerate(photo_rows)
                       if i not in used and len(pr) == data_cols), None)
            if mi is None:
                mi = next((i for i in range(len(photo_rows)) if i not in used), None)
            if mi is None:
                continue
            used.add(mi)
            g["rows"] = ([["조감도"] + ["" for _ in range(ncol - 1)]]
                         + [list(r) for r in (g.get("rows") or [])])
            g["_thumb_imgs"] = photo_rows[mi]


def _pack_short_pages(pages: list, *, debug: bool = False) -> None:
    """연속된 섹션3/4 '짧은' 페이지를 한 슬라이드로 합쳐 페이지 수를 줄인다.
       (build_structured_slide가 넘치면 자동 분할하므로 합쳐도 안전 — 보수적 임계.)
       합칠 때: 첫 페이지 소제목이 슬라이드 상단, 뒤 페이지는 그 표에 미니라벨(소제목)로."""
    def can_lead(p):
        st = p.get("_struct")
        return (isinstance(st, dict) and p.get("_sec_int") in (3, 4)
                and not p.get("_invest_diagram") and not st.get("_nested_grids"))

    def est_h(st):
        h = 0.55 if (st.get("intro") or "").strip() else 0.12
        for t in (st.get("tables") or []):
            rows = t.get("rows") or []
            h += 0.40 + (1 + len(rows)) * 0.26
        h += 0.20 * len([b for b in (st.get("bullets") or []) if str(b).strip()])
        return h

    def can_absorb(p):
        # 뒤따르며 흡수될 작은 페이지(이미지 없음, 표 작음, 연락처/첨부이미지 제외)
        st = p.get("_struct")
        if not (isinstance(st, dict) and not p.get("_invest_diagram")
                and not st.get("_nested_grids") and not p.get("images")):
            return False
        sub = str(p.get("subtitle", ""))
        if any(k in sub for k in ("연락처", "담당자", "승인서", "양해각서", "공문")):
            return False
        return bool(st.get("tables")) and est_h(st) <= 2.7

    out, i = [], 0
    while i < len(pages):
        p = pages[i]
        # 리드 자격 + '표가 있는' 페이지만 뒤 페이지를 흡수(이미지 전용/연락처는 흡수 안 함)
        lead_sub = str(p.get("subtitle", ""))
        lead_ok = (can_lead(p) and bool(p["_struct"].get("tables"))
                   and not any(k in lead_sub for k in ("연락처", "담당자", "승인서", "양해각서", "공문")))
        if not lead_ok:
            out.append(p); i += 1; continue
        group, j = [p], i + 1
        # 큰 페이지(p)가 리드 → 뒤따르는 같은 섹션의 '작은' 페이지들을 흡수
        while (j < len(pages) and pages[j].get("_sec_int") == p.get("_sec_int")
               and can_absorb(pages[j])):
            group.append(pages[j]); j += 1
        if len(group) > 1:
            base = group[0]; bst = base["_struct"]

            def _is_note_b(b):
                s = str(b).lstrip()
                return s.startswith(("주", "*", "※")) and (
                    s[:1] in ("*", "※") or (len(s) > 1 and (s[1].isdigit() or s[1] == " ")))

            # ★base(앞 표)의 각주는 base의 '마지막 표'에 귀속(아래 표로 안 넘어가게)
            _base_notes = [b for b in (bst.get("bullets") or []) if _is_note_b(b)]
            bst["bullets"] = [b for b in (bst.get("bullets") or []) if not _is_note_b(b)]
            if _base_notes and bst.get("tables"):
                bst["tables"][-1].setdefault("_notes", []).extend(_base_notes)

            for q in group[1:]:
                qst = q["_struct"]
                qname = q.get("subtitle", "")          # "3.2 ..."
                qtabs = qst.get("tables") or []
                # 이 페이지(q)의 각주는 이 페이지 '첫 표' 밑에만 붙인다(표별 각주 분리)
                q_notes = [b for b in (qst.get("bullets") or []) if _is_note_b(b)]
                q_prose = [b for b in (qst.get("bullets") or []) if not _is_note_b(b)]
                for k, t in enumerate(qtabs):
                    t2 = dict(t)
                    if k == 0 and not (t2.get("title") or "").strip():
                        t2["title"] = qname             # 뒤 페이지 첫 표에 소제목 라벨
                    if k == 0 and q_notes:
                        t2.setdefault("_notes", []).extend(q_notes)
                    bst.setdefault("tables", []).append(t2)
                # ★흡수된 페이지의 intro(LLM 생성 소개문, 원본엔 없음)는 떠다니는 불릿으로
                #   찍지 않는다(사용자: "적지마" — 표만 깔끔하게). 표 옆 산문(분석)만 유지.
                bst.setdefault("bullets", []).extend(q_prose)
            # ★흡수된 페이지의 빨강/밑줄/포인트색 문구도 base로 합침(각주 빨강 재현 — Equity 주2) 등)
            for _key in ("_red_texts", "_underline_texts", "_filled_texts"):
                _merged = list(base.get(_key) or [])
                for q in group[1:]:
                    _merged.extend(q.get(_key) or [])
                if _merged:
                    base[_key] = list(dict.fromkeys(_merged))
            out.append(base)
            if debug:
                print(f"  [페이지합치기] {[g.get('subtitle') for g in group]} → 1슬라이드")
        else:
            out.append(p)
        i = j
    pages[:] = out
