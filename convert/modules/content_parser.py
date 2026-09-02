"""
content_parser.py
─────────────────────────────────────────────────────────
PDF / Word 문서에서 슬라이드별 내용(제목·부제목·본문)을 추출합니다.

반환 형식:
  [
    {"section_title": "01  사모사채개요",
     "subtitle":      "1.1 현재 자금조달 구조",
     "body_text":     "...본문..."},
    ...
  ]
─────────────────────────────────────────────────────────
"""

import io
import os
import re
from typing import List, Dict

from modules.constants import DEFAULT_TOC_MAP

PageData = Dict[str, str]

# ─────────────────────────────────────────────
# [DEBUG 플래그]
# True 로 바꾸면 각 페이지 파싱 결과가 콘솔에 출력됩니다.
# 배포 시에는 False 로 유지하세요.
# ─────────────────────────────────────────────
DEBUG_PARSER = False


# ─────────────────────────────────────────────
# [확실한 페이지번호·헤더 패턴 — 화이트리스트 원칙]
# 아래 패턴에 완전히 일치하는 줄만 제거합니다.
# 의심스러우면 남깁니다.
# ─────────────────────────────────────────────
_OBVIOUS_FOOTER_RES = [
    re.compile(r'^[-–—]\s*\d{1,3}\s*[-–—]$'),   # "- 3 -"
    re.compile(r'^\d{1,3}\s*/\s*\d{1,3}$'),       # "3 / 25"
    re.compile(r'^페이지\s*\d+$'),                  # "페이지 3"
    re.compile(r'^Page\s+\d+$', re.IGNORECASE),   # "Page 3"
    re.compile(r'^\d{1,3}$'),                      # 단독 숫자 "3"
]

# 목차 페이지 감지 — 이 문자열이 페이지 텍스트(소문자)에 포함되면 목차로 판단
_TOC_MARKERS = ['목차', 'c o n t e n t s', 'contents']

# ──────────────────────────────────────────────────────
# 섹션 자동 분류 — 키워드 기반 (순서대로 매칭, 더 높은 섹션으로만 전진)
# 우선순위 높은 것(04)부터 체크
# ──────────────────────────────────────────────────────
_SECTION_TRIGGERS: list = [
    ("04", "Appendix",       ["별첨", "[별첨"]),
    ("03", "본건 사업 개요", ["사업 개요", "■ 사업 개요", "분양 개요", "분양개요",
                              "토지 확보", "토지 수용", "토지이용", "사업지 전경",
                              "입지 분석", "시세비교",
                              "차주 개요", "기업개요", "재무제표", "인근 산업단지"]),
    ("02", "금융개요",       ["financing terms", "선행조건", "대주간",
                              "terms & conditions"]),
    ("01", "사모사채 개요",  ["executive summary"]),
]

# 불릿 문자로 시작하는 section_title은 헤더가 아닌 본문 첫 줄 — 분류에 사용하지 않음
_BULLET_CHARS = ('▶', '•', '■', '▪', '◆', '→', '☞')

def _classify_section(section_title: str, body_text: str):
    """
    페이지 섹션 번호·섹션명을 반환. 미분류면 ('', '') 반환.

    section_title이 불릿 문자로 시작하거나 60자를 초과하면 헤더가 아닌
    본문 텍스트가 잘못 추출된 것으로 간주해 건너뜁니다.
    section_title에서 매칭 실패 시 body_text 앞 60자를 보조 검사합니다.
    """
    # 1차: section_title — 불릿 시작이나 너무 긴 경우 건너뜀
    title_stripped = section_title.strip()
    title_is_valid_header = (
        bool(title_stripped)
        and not title_stripped[0] in _BULLET_CHARS
        and len(title_stripped) <= 60
    )
    if title_is_valid_header:
        title_lower = title_stripped.lower()
        for sec_num, label, keywords in _SECTION_TRIGGERS:
            if any(kw.lower() in title_lower for kw in keywords):
                return sec_num, label

    # 2차: body_text 앞 60자
    body_head = body_text[:60].lower()
    for sec_num, label, keywords in _SECTION_TRIGGERS:
        if any(kw.lower() in body_head for kw in keywords):
            return sec_num, label
    return "", ""


# ──────────────────────────────────────────────────────
# 공개 API
# ──────────────────────────────────────────────────────

def parse_document(file_path: str) -> List[PageData]:
    """
    PDF 또는 Word(.docx) 파일을 읽어 슬라이드 단위로 분리한 목록을 반환합니다.

    Parameters
    ----------
    file_path : 로컬 파일 경로 (.pdf 또는 .docx)

    Returns
    -------
    list[dict] — 각 항목이 하나의 content 슬라이드에 대응:
        section_title  제목 텍스트 박스에 들어갈 내용
        subtitle       부제목 텍스트 박스에 들어갈 내용
        body_text      본문 텍스트 박스에 들어갈 내용
    """
    ext = os.path.splitext(file_path)[1].lower()
    if ext == ".docx":
        with open(file_path, "rb") as f:
            return _parse_docx_bytes(f.read())
    elif ext == ".pdf":
        with open(file_path, "rb") as f:
            return map_pdf_pages_to_slides(f.read(), debug=DEBUG_PARSER)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext!r}  (.pdf 또는 .docx만 지원)")


def parse_document_from_bytes(data: bytes, filename: str) -> List[PageData]:
    """
    Streamlit 파일 업로더에서 받은 bytes를 바로 파싱합니다.

    Parameters
    ----------
    data     : 업로드된 파일의 bytes
    filename : 원본 파일명 (확장자 판별에 사용)
    """
    ext = os.path.splitext(filename)[1].lower()
    if ext == ".docx":
        return _parse_docx_bytes(data)
    elif ext == ".pdf":
        return map_pdf_pages_to_slides(data, debug=DEBUG_PARSER)
    else:
        raise ValueError(f"지원하지 않는 파일 형식: {ext!r}  (.pdf 또는 .docx만 지원)")


# ──────────────────────────────────────────────────────
# [페이지 매핑 — 핵심 공개 함수]
# ──────────────────────────────────────────────────────

def map_pdf_pages_to_slides(data: bytes, debug: bool = False) -> List[PageData]:
    """
    PDF 각 페이지를 분석해 PPT content 슬라이드 데이터 목록을 반환합니다.
    1:1 매핑 원칙: 각 PDF 콘텐츠 페이지 → 1개 content 슬라이드 데이터.

    자동 제외 규칙:
      - 1페이지: 표지 (PPT는 별도 커버 슬라이드를 사용)
      - 목차 포함 페이지 (목차 / CONTENTS 감지)
      - 텍스트 거의 없는 페이지 (15자 미만)

    섹션 제목 전파:
      - 현재 페이지 헤더가 너무 짧으면 (10자 미만) 이전 페이지의 section_title 유지

    Parameters
    ----------
    data  : PDF 파일 bytes
    debug : True 이면 각 페이지 처리 결과를 콘솔에 출력

    Returns
    -------
    list[PageData]  — build_full_presentation()에 바로 전달 가능한 형식
    """
    import fitz
    import pdfplumber

    fitz_doc     = fitz.open(stream=data, filetype="pdf")
    plumber_pdf  = pdfplumber.open(io.BytesIO(data))
    total        = fitz_doc.page_count
    result: List[PageData] = []
    last_section_title = ""

    # ★발행사 로고·워터마크(여러 페이지 반복 이미지)는 본문에서 통째로 제외한다.
    logo_xrefs = _find_repeated_xrefs(fitz_doc)

    if debug:
        print(f"\n[content_parser] PDF 총 {total}페이지 분석 시작")
        if logo_xrefs:
            print(f"[content_parser] 로고·워터마크로 판단해 제외할 이미지 xref: "
                  f"{sorted(logo_xrefs)}")
        print("─" * 60)

    for i, pdf_page in enumerate(fitz_doc):
        page_num = i + 1
        full_text_lower = (pdf_page.get_text("text") or "").lower().strip()

        # ── 규칙 1: 1페이지는 표지로 간주, 항상 제외 ──────────
        if page_num == 1:
            if debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 표지 (1페이지 자동 제외)")
            continue

        # ── 규칙 2: 목차 페이지 제외 (★제목 위치=맨 위 3줄에 있을 때만) ──
        #   본문에 '목차/contents'를 언급만 한 Appendix 페이지가 통째로
        #   사라지던 문제 방지 — 진짜 목차 페이지는 그 단어가 상단 제목에 있음.
        _head_lower = "\n".join(full_text_lower.splitlines()[:3])
        if any(marker in _head_lower for marker in _TOC_MARKERS):
            if debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 목차 페이지")
            continue

        # ── 페이지 텍스트 추출 (fitz) — 실패 시 그 페이지만 건너뜀 ──
        try:
            pd = _extract_pdf_page(pdf_page, debug=debug)
        except Exception as _pe:
            if debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 텍스트 추출 실패: {_pe}")
            continue
        if pd is None:
            if debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 텍스트 없음")
            continue

        # ── 표 추출 (pdfplumber) — ★'내용 부족' 판단보다 먼저 ──────
        #   표만 있고 글자는 적은 페이지(거래사례·민감도·비교표 등 Appendix)를
        #   표 확인도 안 하고 버리던 버그 수정: 표를 먼저 뽑아서 판단에 반영한다.
        #   ★안전장치: 특정 페이지에서 표 추출이 실패해도 파싱 '전체'가 죽지 않게 감싼다
        #     (이 한 페이지만 표 없이 진행 → 이전엔 여기서 터지면 parsed_pages가 통째로 비어
        #      본문이 하나도 안 나오는 치명적 회귀가 났음).
        try:
            table_data = _extract_page_tables(plumber_pdf.pages[i], debug=debug)
        except Exception as _te:
            if debug:
                print(f"[p{page_num:02d}/{total}] 표 추출 실패(무시): {_te}")
            table_data = {"tables": [], "table_bboxes": [],
                          "text_without_tables": pd.get("body_text", "")}
        pd.update(table_data)   # tables, table_bboxes, text_without_tables 추가

        # ── 규칙 3: 내용 부족 페이지 제외 (★단, 표가 있으면 살린다) ─
        combined = (pd["section_title"] + pd["subtitle"] + pd["body_text"]).strip()
        if len(combined) < 15 and not pd.get("tables"):
            # ★이미지만 있는 페이지(조감도·광역입지 등): 버리면 그림이 사라진다.
            #   큰 이미지가 있으면 직전 본문 페이지(건축개요 등)에 y좌표 순으로 첨부 →
            #   build_structured_slide 의 side_box(표 옆 라벨 박스) 경로로 렌더된다.
            recovered = _extract_page_images(fitz_doc, i, skip_xrefs=logo_xrefs)
            if recovered and result:
                ypos = {}
                for im in pdf_page.get_images(full=True):
                    try:
                        rects = pdf_page.get_image_rects(im[0])
                        if rects:
                            ypos[im[0]] = rects[0].y0
                    except Exception:
                        pass
                recovered.sort(key=lambda im: ypos.get(im.get("xref"), 0.0))
                result[-1].setdefault("images", []).extend(recovered)
                if debug:
                    print(f"[p{page_num:02d}/{total}] 이미지전용 → 직전 페이지에 이미지 "
                          f"{len(recovered)}개 첨부(조감도/광역입지 등)")
            elif debug:
                print(f"[p{page_num:02d}/{total}] SKIP — 내용 부족 ({len(combined)}자)")
            continue

        # ── 섹션 제목 전파 ─────────────────────────────────────
        _SEC_NUM = re.compile(r'^0[1-9][\s\.\-]')
        has_sec_num = bool(_SEC_NUM.match(pd["section_title"].strip()))

        if has_sec_num:
            last_section_title = pd["section_title"]
        elif last_section_title:
            # {**pd, ...} 로 tables 등 기존 키 전부 보존
            pd = {
                **pd,
                "section_title": last_section_title,
                "subtitle":      pd["subtitle"] or pd["section_title"],
            }

        # ── 페이지 이미지/원문/강조 추출 (★각각 감싸 실패해도 페이지는 살린다) ──
        pd["page_num"] = page_num
        try:
            pd["images"] = _extract_page_images(fitz_doc, i, skip_xrefs=logo_xrefs)
        except Exception:
            pd["images"] = []
        try:
            pd["raw_text"] = pdf_page.get_text("text") or ""
        except Exception:
            pd["raw_text"] = pd.get("body_text", "")
        try:
            pd["_red_texts"] = _extract_red_texts(pdf_page)
        except Exception:
            pd["_red_texts"] = []
        try:
            pd["_underline_texts"] = _extract_underline_texts(pdf_page, pd.get("table_bboxes"))
        except Exception:
            pd["_underline_texts"] = []
        try:
            pd["_filled_texts"] = _extract_filled_texts(pdf_page)
        except Exception:
            pd["_filled_texts"] = []

        if debug:
            sec_preview  = pd["section_title"][:30]
            body_preview = pd["body_text"][:40].replace("\n", " ")
            tbl_count    = len(pd.get("tables", []))
            img_count    = len(pd["images"])
            print(f"[p{page_num:02d}/{total}] KEEP  | sec={sec_preview!r:35} | "
                  f"body={len(pd['body_text'])}자 | tables={tbl_count}개 | imgs={img_count}개 | {body_preview!r}")

        result.append(pd)

    plumber_pdf.close()
    fitz_doc.close()

    # ── 2차 패스: 섹션 번호 자동 분류 ─────────────────────
    # 각 페이지에 section_num, section_label 필드 추가
    # 단조 증가 원칙: 한번 높은 섹션으로 올라가면 내려오지 않음
    cur_sec_num   = ""
    cur_sec_label = ""
    if debug:
        print("\n[섹션 자동 분류 — 2차 패스]")
    for idx, pd_item in enumerate(result):
        triggered_num, triggered_label = _classify_section(
            pd_item.get("section_title", ""),
            pd_item.get("body_text", ""),
        )
        # 더 높은 섹션 번호가 감지됐을 때만 업데이트 (단조 증가)
        if triggered_num and (not cur_sec_num or triggered_num > cur_sec_num):
            cur_sec_num   = triggered_num
            cur_sec_label = triggered_label
            if debug:
                title_preview = pd_item.get("section_title", "")[:40]
                print(f"  slide{idx+1:02d}: → [{cur_sec_num}] {cur_sec_label!r}  "
                      f"(trigger={triggered_num!r}, title={title_preview!r})")
        pd_item["section_num"]   = cur_sec_num
        pd_item["section_label"] = cur_sec_label

    if debug:
        print("─" * 60)
        print(f"[content_parser] 최종 content 슬라이드 수: {len(result)}장\n")

    return result


# ──────────────────────────────────────────────────────
# [필터 함수 — 확실한 헤더/푸터만 제거]
# ──────────────────────────────────────────────────────

def clean_page_text(lines: List[str], debug: bool = False) -> List[str]:
    """
    페이지 텍스트 줄 목록에서 확실한 페이지번호·헤더 패턴만 제거합니다.
    화이트리스트 원칙: 의심스러우면 남깁니다.

    제거 대상:
      - "- 3 -", "3 / 25", "페이지 3", "Page 3", 단독 숫자

    Parameters
    ----------
    lines : 텍스트 줄 목록
    debug : True 이면 제거된 줄을 콘솔에 출력

    Returns
    -------
    필터링된 줄 목록
    """
    result = []
    for line in lines:
        stripped = line.strip()
        removed = any(p.fullmatch(stripped) for p in _OBVIOUS_FOOTER_RES)
        if removed:
            if debug:
                print(f"  [clean_page_text 제거] {stripped!r}")
        else:
            result.append(line)
    return result


# ──────────────────────────────────────────────────────
# Word (.docx) 파서
# ──────────────────────────────────────────────────────

def _parse_docx_bytes(data: bytes) -> List[PageData]:
    import docx

    doc = docx.Document(io.BytesIO(data))
    pages: List[PageData] = []

    current_section_title = ""
    current_subtitle      = ""
    body_lines: List[str] = []
    in_page = False

    def _flush():
        nonlocal current_subtitle, body_lines, in_page
        if in_page:
            pages.append({
                "section_title": current_section_title,
                "subtitle":      current_subtitle,
                "body_text":     "\n".join(body_lines).strip(),
            })
        current_subtitle = ""
        body_lines       = []
        in_page          = False

    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style = para.style.name if para.style else ""

        if _is_heading_level(style, 1):
            _flush()
            current_section_title = text

        elif _is_heading_level(style, 2):
            _flush()
            current_subtitle = text
            in_page = True

        elif _is_heading_level(style, 3):
            _flush()
            current_subtitle = text
            in_page = True

        else:
            if not in_page:
                in_page = True
            body_lines.append(text)

    _flush()

    if not pages and current_section_title:
        pages.append({
            "section_title": "",
            "subtitle":      current_section_title,
            "body_text":     "\n".join(body_lines).strip(),
        })

    return pages


def _is_heading_level(style_name: str, level: int) -> bool:
    """'Heading 1', '제목 1', 'heading1' 등 다양한 스타일명을 허용합니다."""
    s = style_name.lower()
    if f"heading {level}" in s:
        return True
    if f"제목 {level}" in s:
        return True
    if f"heading{level}" in s:
        return True
    return False


# ──────────────────────────────────────────────────────
# PDF 파서 — 페이지 단위 이미지 추출
# ──────────────────────────────────────────────────────
_IMG_MIN_PX = 100 * 100   # 100×100px 미만만 아이콘·로고로 간주해 제외
#   (200×200이면 조감도 썸네일·비교표 사진이 통째로 걸러지던 문제 → 100×100로 완화)

# ★발행사 로고·워터마크 제외 기준(2026-08-06)
#   원본 IM 은 페이지마다 발행사 로고와 워터마크를 깔아둔다. 이게 크기 필터를 통과해
#   본문 슬라이드에 '조감도'로 박히는 사고가 있었다(돈암동: 카카오페이증권 로고·워터마크가
#   6개 슬라이드에 삽입, 슬13은 '조감도' 라벨까지 붙음).
#   판정: 같은 xref 가 이 페이지 수 이상에서 쓰이면 = 머리말/꼬리말 장식으로 본다.
#   (조감도·현장사진·지도는 특정 페이지에만 나오므로 안전하다.)
_REPEAT_PAGE_MIN = 3


def _find_repeated_xrefs(fitz_doc, min_pages: int = _REPEAT_PAGE_MIN) -> set:
    """문서 전체에서 min_pages 이상 페이지에 등장하는 이미지 xref 집합(로고·워터마크)."""
    from collections import Counter
    use = Counter()
    try:
        for pg in fitz_doc:
            for x in {im[0] for im in pg.get_images(full=True)}:
                use[x] += 1
    except Exception:
        return set()
    if fitz_doc.page_count < min_pages:      # 페이지가 너무 적으면 판정 불가 → 제외 안 함
        return set()
    return {x for x, c in use.items() if c >= min_pages}


def _extract_red_texts(pdf_page) -> list:
    """페이지에서 '빨간색' 글자 조각을 추출(원본 강조 표시 재현용).
       ★같은 줄에서 가까이 붙은 빨간 span 은 한 구(句)로 합침(예 '매출대비'+'6.7%' →
         '매출대비 6.7%'). 표 칸 사이(간격 큼)는 합치지 않아 셀 단위 매칭이 깨지지 않음."""
    reds = []
    try:
        d = pdf_page.get_text("dict")
    except Exception:
        return reds

    def _is_red(col):
        r = (col >> 16) & 0xFF; g = (col >> 8) & 0xFF; b = col & 0xFF
        return r > 120 and g < 90 and b < 90

    def _flush(cur):
        if cur:
            frag = " ".join(cur).strip()
            if len(frag.replace(" ", "")) >= 2:
                reds.append(frag)

    for blk in d.get("blocks", []):
        for ln in blk.get("lines", []):
            cur, prev_x1 = [], None
            for sp in ln.get("spans", []):
                t = (sp.get("text", "") or "").strip()
                if _is_red(sp.get("color", 0) or 0) and t:
                    x0, _, x1, _ = sp.get("bbox", (0, 0, 0, 0))
                    if cur and prev_x1 is not None and (x0 - prev_x1) > 15:
                        _flush(cur); cur = []      # 간격 크면(다른 칸) 끊음
                    cur.append(t); prev_x1 = x1
                else:
                    _flush(cur); cur, prev_x1 = [], None
            _flush(cur)
    return reds


def _extract_underline_texts(pdf_page, table_bboxes=None) -> list:
    """밑줄 친 글자 조각 추출(원본 강조 재현).
       글자 baseline(y1) 바로 밑의 얇은 가로 rect/line 을, 그 위 글자들의 x범위와 거의
       정확히 일치할 때만 밑줄로 인정. ★표 영역(table_bboxes) 안의 선은 셀 테두리이므로 제외."""
    outs = []
    tboxes = [b for b in (table_bboxes or []) if b and len(b) >= 4]

    def _in_table(x0, x1, y):
        cx = (x0 + x1) / 2.0
        for (bx0, bt, bx1, bb) in tboxes:
            if bx0 - 2 <= cx <= bx1 + 2 and bt - 2 <= y <= bb + 2:
                return True
        return False
    try:
        words = pdf_page.get_text("words")    # (x0,y0,x1,y1,word,block,line,wno)
    except Exception:
        return outs
    segs = []   # (x0, x1, y) 얇은 가로선 후보
    try:
        for dr in pdf_page.get_drawings():
            for it in dr.get("items", []):
                if it[0] == "re":
                    r = it[1]
                    if r.height < 2.5 and r.width > 8:
                        segs.append((r.x0, r.x1, r.y0))
                elif it[0] == "l":
                    p1, p2 = it[1], it[2]
                    if abs(p1.y - p2.y) < 1.0 and abs(p2.x - p1.x) > 8:
                        segs.append((min(p1.x, p2.x), max(p1.x, p2.x), p1.y))
    except Exception:
        return outs
    for (sx0, sx1, sy) in segs:
        cand = [w for w in words
                if abs(w[3] - sy) <= 1.8 and w[2] > sx0 - 1 and w[0] < sx1 + 1]
        if not cand:
            continue
        cand.sort(key=lambda w: w[0])
        # 밑줄이 글자들을 '딱' 감싸야 함(양끝 8px 이내) — 아니면 표 테두리로 보고 제외
        if abs(cand[0][0] - sx0) > 8 or abs(cand[-1][2] - sx1) > 8:
            continue
        frag = " ".join(w[4] for w in cand).strip()
        # ★표 셀 단어 1개(소재지·시공사명 등) 오탐 배제: 여러 단어이거나 8자 이상인 강조구만 인정.
        #   (원본 밑줄은 '포함 기준'·'…담보신탁으로 변경할'처럼 구(句) 단위)
        if (" " not in frag) and len(frag) < 8:
            continue
        if not _in_table(sx0, sx1, sy) or " " in frag:
            outs.append(frag)
            if " " in frag:
                outs.append(frag.replace(" ", ""))
    return list(dict.fromkeys(outs))


def _extract_filled_texts(pdf_page) -> list:
    """원본에서 '포인트 색'으로 칠해진 셀의 텍스트 추출(시세표 공급평단가 주황 등).
       → 렌더 시 그 칸을 하늘색 65%로 재현(원본 색칠 부분 = 하늘 65% 약속).
       헤더(어두운 색)·회색 음영은 제외, 채도 있는 밝은 하이라이트만."""
    outs = []
    try:
        words = pdf_page.get_text("words")
        draws = pdf_page.get_drawings()
    except Exception:
        return outs
    for dr in draws:
        f = dr.get("fill")
        if f is None:
            continue
        try:
            r, g, b = [round(c * 255) for c in f[:3]]
        except Exception:
            continue
        # 포인트(하이라이트): 회색 아님(채널차>20) + 밝음(평균>140). 헤더(어두움)/회색 제외.
        if (max(r, g, b) - min(r, g, b)) <= 20 or (r + g + b) / 3 <= 140:
            continue
        for it in dr.get("items", []):
            if it[0] != "re":
                continue
            rc = it[1]
            if rc.width < 8 or rc.height < 4:
                continue
            inside = [w[4] for w in words
                      if w[0] >= rc.x0 - 1 and w[2] <= rc.x1 + 1
                      and w[1] >= rc.y0 - 1 and w[3] <= rc.y1 + 1]
            t = " ".join(inside).strip()
            if len(t) >= 2:
                outs.append(t)
    return list(dict.fromkeys(outs))


def _extract_page_images(fitz_doc, page_idx: int, skip_xrefs=None) -> list:
    """
    fitz_doc의 page_idx 페이지에서 이미지를 추출합니다.
    _IMG_MIN_PX 미만 소형 이미지(아이콘·로고)는 제외합니다.
    skip_xrefs 에 든 xref(= 여러 페이지에 반복되는 발행사 로고·워터마크)도 제외합니다.

    Returns
    -------
    [{"xref": int, "width": int, "height": int, "ext": str, "data": bytes}, ...]
    """
    import fitz
    page = fitz_doc[page_idx]
    result = []
    seen_xrefs = set()
    skip_xrefs = skip_xrefs or set()
    for img in page.get_images(full=True):
        xref = img[0]
        smask_xref = img[1] if len(img) > 1 else 0   # 별도 투명 마스크(SMask) xref
        if xref in seen_xrefs or xref in skip_xrefs:
            continue
        seen_xrefs.add(xref)
        try:
            # ★Pixmap으로 렌더 + SMask(별도 투명마스크)를 합쳐 '보이는 그대로' 만든다.
            #   (extract_image는 본체만 뽑아 투명영역이 검게 나옴 → 검정 배경 버그의 진짜 원인)
            pix = fitz.Pixmap(fitz_doc, xref)
            if smask_xref:
                try:
                    if pix.alpha:                     # 결합 전 base는 알파가 없어야 함
                        pix = fitz.Pixmap(pix, 0)     # 알파 제거
                    mpix = fitz.Pixmap(fitz_doc, smask_xref)
                    pix = fitz.Pixmap(pix, mpix)      # 마스크를 알파로 결합 → 투명 복원
                except Exception:
                    pass
            # CMYK/기타 → RGB(알파 유지)
            if (pix.n - pix.alpha) >= 4 or (pix.colorspace and pix.colorspace.name not in ("DeviceRGB", "DeviceGray")):
                pix = fitz.Pixmap(fitz.csRGB, pix)
            w, h = pix.width, pix.height
            if w * h < _IMG_MIN_PX:
                continue
            data = pix.tobytes("png")             # 투명이면 알파 보존한 PNG
            result.append({
                "xref":   xref,
                "width":  w,
                "height": h,
                "ext":    "png",
                "data":   data,
            })
        except Exception:
            # Pixmap 실패 시 원본 추출로 폴백(+ 포맷 정규화)
            try:
                base = fitz_doc.extract_image(xref)
                w, h = base["width"], base["height"]
                if w * h < _IMG_MIN_PX:
                    continue
                data, ext = _normalize_image_bytes(base["image"], base["ext"])
                result.append({"xref": xref, "width": w, "height": h, "ext": ext, "data": data})
            except Exception:
                pass
    return result


def _normalize_image_bytes(data: bytes, ext: str):
    """PDF에서 추출한 이미지를 '본래 모습 그대로' 깨끗한 포맷으로 재저장한다.
       - 투명(알파) 이미지 → 투명 유지(RGBA PNG). 흰 배경 강제 X.
       - CMYK 등 PPT가 검게 렌더하는 포맷 → RGB로 변환.
       검정 배경 버그의 원인은 '깨진/미지원 포맷'이므로 깨끗한 PNG로 다시 저장해 해결.
       실패 시 원본 그대로."""
    try:
        from PIL import Image
        import io as _io
        im = Image.open(_io.BytesIO(data))
        has_alpha = (im.mode in ("RGBA", "LA")
                     or (im.mode == "P" and "transparency" in im.info))
        if has_alpha:
            im = im.convert("RGBA")            # 투명 그대로 유지
        elif im.mode == "CMYK":
            im = im.convert("RGB")             # CMYK는 검게 나오므로 RGB로
        elif im.mode not in ("RGB", "L"):
            im = im.convert("RGB")
        else:
            return data, ext                   # RGB/L 정상 포맷은 원본 유지
        out = _io.BytesIO()
        im.save(out, "PNG")                    # 깨끗한 PNG로 재저장(투명이면 알파 보존)
        return out.getvalue(), "png"
    except Exception:
        pass
    return data, ext


# ──────────────────────────────────────────────────────
# PDF 파서 — 페이지 단위 텍스트 추출
# ──────────────────────────────────────────────────────

def _extract_pdf_page(page, debug: bool = False) -> "PageData | None":
    """
    PDF 한 페이지에서 헤더(상단 20%) 텍스트를 제목/부제목으로,
    나머지를 본문으로 추출합니다.
    본문에는 clean_page_text()로 명확한 페이지번호만 제거합니다.
    """
    page_h = page.rect.height
    header_thresh = page_h * 0.20

    lines: List[tuple] = []
    blocks = page.get_text("dict", flags=0)["blocks"]

    for block in blocks:
        if block.get("type") != 0:
            continue
        for line in block["lines"]:
            y0       = line["bbox"][1]
            txt      = ""
            max_size = 0.0
            bold     = False
            for span in line["spans"]:
                t = span["text"].strip()
                if t:
                    txt += t + " "
                    if span["size"] > max_size:
                        max_size = span["size"]
                    if span["flags"] & 16:
                        bold = True
            txt = txt.strip()
            if txt and len(txt) >= 2:
                lines.append((y0, txt, max_size, bold))

    if not lines:
        return None

    lines.sort(key=lambda x: x[0])

    header_lines = [(y, t, sz, b) for y, t, sz, b in lines if y < header_thresh]
    body_lines   = [(y, t, sz, b) for y, t, sz, b in lines if y >= header_thresh]

    # 헤더에서 제목 / 부제목 분리 (가장 큰 폰트 = 섹션 제목)
    section_title = ""
    subtitle      = ""
    if header_lines:
        max_sz = max(sz for _, _, sz, _ in header_lines)
        title_parts    = [t for _, t, sz, _ in header_lines if sz >= max_sz - 1.0]
        subtitle_parts = [t for _, t, sz, _ in header_lines if sz < max_sz - 1.0]
        section_title  = " ".join(title_parts).strip()
        subtitle       = " ".join(subtitle_parts).strip()

    # 본문 — 명확한 페이지번호 패턴만 제거 (화이트리스트 원칙)
    raw_body_texts = [t for _, t, _, _ in body_lines]
    clean_body     = clean_page_text(raw_body_texts, debug=debug)
    body_text      = "\n".join(clean_body).strip()

    return {
        "section_title": section_title,
        "subtitle":      subtitle,
        "body_text":     body_text,
    }


# ──────────────────────────────────────────────────────
# PDF 파서 — pdfplumber 표 추출
# ──────────────────────────────────────────────────────

def _postprocess_table(table: list) -> list:
    """
    추출된 표 데이터를 정제합니다.

    처리 순서:
      1. 각 셀의 None → "", \\n → 공백, 앞뒤 공백 제거
      2. 모든 셀이 빈 행 제거
      3. 행이 1개뿐인 표는 단순 제목 줄로 간주해 버림
      4. 동일 헤더가 반복되는 경우 첫 번째만 유지 (병합 셀 잔재 정리)
    """
    if not table:
        return []

    # 1. 셀 정제
    cleaned = [
        [str(cell).replace('\n', ' ').strip() if cell is not None else ""
         for cell in row]
        for row in table
    ]

    # 2. 모든 셀이 빈 행 제거
    cleaned = [row for row in cleaned if any(c for c in row)]

    # 2b. 빈 열 제거 — 90% 이상 비어있는 열은 phantom column으로 간주
    if cleaned:
        max_cols = max(len(r) for r in cleaned)
        keep = []
        for c in range(max_cols):
            vals = [r[c] if c < len(r) else "" for r in cleaned]
            non_empty = sum(1 for v in vals if v.strip())
            keep.append(non_empty > 0)
        cleaned = [
            [cell for ci, cell in enumerate(row) if ci < len(keep) and keep[ci]]
            for row in cleaned
        ]

    # 3. 행이 1개 이하 → 표가 아님
    if len(cleaned) <= 1:
        return []
    # 열 제거 후 1열만 남아도 버림
    if max((len(r) for r in cleaned), default=0) < 2:
        return []

    # 4. 반복 헤더 제거 (첫 행 = 헤더 기준)
    header = tuple(cleaned[0])
    deduped = [cleaned[0]]
    for row in cleaned[1:]:
        if tuple(row) != header:
            deduped.append(row)
    return deduped


def _is_valid_table(table_data: list) -> bool:
    """
    추출된 표 데이터가 실제 표인지 검사합니다.
    ※ _extract_page_tables 에서 _postprocess_table(빈 행·팬텀 열 제거) 이후의
      "정제된 표" 를 대상으로 호출되므로, 임계치를 관대하게 둔다 (내용 보존 우선).

    필터 조건 (하나라도 해당하면 False):
      - 행 수 < 2        : 헤더 1줄짜리 = 표 아님
      - 열 수 < 2        : 단일 열 = 텍스트박스/라벨 조각
      - 열 수 > 14       : phantom columns (정제 후에도 과다)
      - 불릿 문자 시작 셀 존재 : 본문 단락을 표로 오인식
      - 빈 셀 비율 > 95% : 거의 전부 비어있는 레이아웃 요소
      - 유니크 값 < 2    : 의미 있는 데이터가 사실상 없음
      - 80자 초과 셀 ≥ 8 : 단락 텍스트를 표로 오인식
      - 헤더 유효셀 < 1  : 헤더 행이 전부 비었거나 너무 긺
    """
    if not table_data:
        return False

    rows = len(table_data)
    cols = max((len(row) for row in table_data), default=0)

    # 행/열 수 조건 (완화: 작은 표도 살림)
    if rows < 2:
        return False
    if cols < 2:
        return False
    if cols > 14:
        return False

    # 불릿 단락 오인식 방지 — 단, "•" 를 셀 마커로 쓰는 정상 표(사업개요표 등)는 살림.
    #   좁은 표(≤2열)에서 불릿이 다수 행을 차지하면 본문 불릿 리스트로 간주해 탈락,
    #   3열 이상 데이터 표는 불릿이 있어도 허용.
    BULLET_CHARS = ('▶', '•', '■', '▪', '◆')
    if cols <= 2:
        bullet_rows = sum(
            1 for row in table_data
            if any(str(c or '').strip()[:1] in BULLET_CHARS for c in row)
        )
        if bullet_rows >= max(2, rows * 0.5):
            return False

    total_cells = sum(len(row) for row in table_data)
    if total_cells == 0:
        return False
    empty_cells = sum(
        1 for row in table_data for cell in row
        if not str(cell or "").strip()
    )
    if empty_cells / total_cells > 0.95:        # 85% → 95% 완화
        return False

    flat = [str(c).strip() for row in table_data for c in row if str(c or "").strip()]
    if len(set(flat)) < 2:                      # 3 → 2 완화
        return False

    long_cells = sum(
        1 for row in table_data for cell in row
        if cell and len(str(cell)) > 80          # 50 → 80
    )
    if long_cells > 8:                           # 5 → 8 완화
        return False

    # 첫 행(헤더) 검사: 유효한 헤더 셀이 1개 이상 있어야 함
    header_row = table_data[0]
    valid_header_cells = sum(
        1 for cell in header_row
        if cell and 0 < len(str(cell).strip()) <= 40
    )
    if valid_header_cells < 1:                    # 2 → 1 완화
        return False

    return True


def _bbox_overlap_ratio(bbox_a: tuple, bbox_b: tuple) -> float:
    """두 bbox의 겹침 비율(IoU)을 반환합니다."""
    x0 = max(bbox_a[0], bbox_b[0])
    y0 = max(bbox_a[1], bbox_b[1])
    x1 = min(bbox_a[2], bbox_b[2])
    y1 = min(bbox_a[3], bbox_b[3])
    if x1 <= x0 or y1 <= y0:
        return 0.0
    inter = (x1 - x0) * (y1 - y0)
    area_a = (bbox_a[2] - bbox_a[0]) * (bbox_a[3] - bbox_a[1])
    area_b = (bbox_b[2] - bbox_b[0]) * (bbox_b[3] - bbox_b[1])
    union = area_a + area_b - inter
    return inter / union if union > 0 else 0.0


def _count_filled_cells(tbl: list) -> int:
    """표(2D 리스트)에서 내용이 있는(비어있지 않은) 셀 개수."""
    return sum(1 for row in (tbl or []) for c in row if str(c or "").strip())


def _strategy_lines(page, debug=False):
    """1차: 선(lines) 전략 (기존 동작 보존 — 선 있는 표는 여기서 최다 셀)."""
    out = []
    settings = {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
    try:
        for t in (page.find_tables(table_settings=settings) or []):
            ext = t.extract()
            if ext:
                out.append((ext, tuple(t.bbox), "lines"))
    except Exception as exc:
        if debug:
            print(f"  [lines 오류] {exc}")
    return out


def _strategy_text(page, debug=False):
    """2차: 텍스트 정렬(text) 전략 — 테두리 없는 표 대응."""
    out = []
    settings = {
        "vertical_strategy": "text", "horizontal_strategy": "text",
        "snap_tolerance": 4, "text_tolerance": 3,
    }
    try:
        for t in (page.find_tables(table_settings=settings) or []):
            ext = t.extract()
            if ext:
                out.append((ext, tuple(t.bbox), "text"))
    except Exception as exc:
        if debug:
            print(f"  [text 오류] {exc}")
    return out


def _strategy_pymupdf(page, debug=False):
    """3차: PyMuPDF(fitz) find_tables — 미설치/실패 시 건너뜀(경고만)."""
    out = []
    try:
        import fitz  # PyMuPDF
    except Exception:
        if debug:
            print("  [pymupdf] 미설치 — 3차 전략 건너뜀")
        return out
    doc = None
    try:
        pdfobj = page.pdf
        path = getattr(pdfobj, "path", None)
        if path:
            doc = fitz.open(path)
        else:
            stream = getattr(pdfobj, "stream", None)
            if stream is None:
                return out
            pos = stream.tell()
            stream.seek(0)
            data = stream.read()
            stream.seek(pos)
            doc = fitz.open(stream=data, filetype="pdf")
        fp = doc[page.page_number - 1]
        finder = fp.find_tables()
        for t in getattr(finder, "tables", []):
            ext = t.extract()
            if ext:
                out.append((ext, tuple(t.bbox), "pymupdf"))
    except Exception as exc:
        if debug:
            print(f"  [pymupdf 오류] {exc}")
    finally:
        if doc is not None:
            try:
                doc.close()
            except Exception:
                pass
    return out


def _extract_tables_dual_strategy(plumber_page, debug: bool = False):
    """다단계 표 추출.

    1차 lines → 2차 text → 3차 pymupdf 후보를 모두 모은 뒤,
    같은 영역(bbox 겹침>0.5)에 대해서는 "내용 있는 셀이 가장 많은" 결과만 채택한다.
    (선 있는 표는 lines가 보통 최다 → 기존 동작 보존, 선 없는 표는 text/pymupdf가 보완)

    Returns
    -------
    List[Tuple[list, tuple, str]] — (2D 배열, bbox, 채택전략) 리스트
    """
    candidates = []
    candidates += _strategy_lines(plumber_page, debug)
    candidates += _strategy_text(plumber_page, debug)
    candidates += _strategy_pymupdf(plumber_page, debug)

    if debug:
        from collections import Counter as _Cnt
        print(f"  [후보 전략별] {dict(_Cnt(s for _, _, s in candidates))}")

    # 점수 = 내용 셀 수 + 전략 가점(lines 우대).
    #   선 있는 표는 lines 가 구조(헤더 등)를 더 정확히 잡으므로 근소차에선 lines 채택.
    _BONUS = {"lines": 3, "pymupdf": 1, "text": 0}

    def _score(item):
        ext, _, strat = item
        return _count_filled_cells(ext) + _BONUS.get(strat, 0)

    chosen = []
    for ext, bbox, strat in sorted(candidates, key=_score, reverse=True):
        if any(_bbox_overlap_ratio(bbox, cb) > 0.5 for _, cb, _ in chosen):
            continue
        chosen.append((ext, bbox, strat))
    return chosen


# ──────────────────────────────────────────────────────
# [5섹션 자동 재구성 — 공개 API]
# ──────────────────────────────────────────────────────

_SUBTITLE_NUM_PREFIX = re.compile(r'^\d+\.\d+\s*')
_APPENDIX_KEYWORDS   = ('appendix', '별첨', '부록', '참고')

# "N.N 한글" 패턴만 진짜 소제목으로 허용 (예: "3.1 사업개요")
_VALID_SUBTITLE_RE = re.compile(r'^\d+\.\d+\s+[가-힣]')
# 본문 마커: 이 문자로 시작하면 body text이므로 즉시 제외
_BODY_MARKER_CHARS = set('▶■•▪◆→☞※')


def extract_toc_map(pages: list) -> dict:
    """
    pages에서 'N.N [한글]' 형식의 소제목만 수집해 section_num → subtitles 매핑을 반환합니다.

    필터 규칙:
      - ▶ ■ • 등 본문 마커로 시작하는 줄 제외
      - '^숫자.숫자 [한글]' 패턴에 맞지 않는 줄 제외 (예: "3.1 사업개요")
      → 본문 텍스트·표 내용이 목차에 침입하는 것을 방지

    PDF의 헤더에서 소제목이 제대로 추출되지 않으면 빈 dict를 반환합니다.
    이 경우 build_toc_slide / _build_content_block 이 DEFAULT_TOC_MAP으로 fallback합니다.

    Returns
    -------
    {"01": ["1.1 ...", ...], "02": [...], ...}  — 소제목 없으면 {}
    """
    _SEC_NUM = re.compile(r'^(0[1-9])\b')
    toc: dict  = {}
    seen: dict = {}
    for page in pages:
        subtitle = page.get("subtitle", "").strip()
        sec_num  = page.get("section_num", "").strip()
        if not sec_num:
            m = _SEC_NUM.match(page.get("section_title", "").strip())
            sec_num = m.group(1) if m else ""
        if not sec_num or not subtitle:
            continue
        # 본문 마커로 시작하는 줄 제외
        if subtitle[0] in _BODY_MARKER_CHARS:
            continue
        # "N.N 한글" 패턴이 아닌 줄 제외 (표 헤더, 긴 문장 등)
        if not _VALID_SUBTITLE_RE.match(subtitle):
            continue
        seen.setdefault(sec_num, set())
        if subtitle not in seen[sec_num]:
            seen[sec_num].add(subtitle)
            toc.setdefault(sec_num, []).append(subtitle)
    print(f"[extract_toc_map] 결과: { {k: len(v) for k, v in toc.items()} } "
          f"(비어있으면 DEFAULT_TOC_MAP 사용)")
    return toc


def extract_section_labels(pages: list) -> dict:
    """
    pages에서 section_num → section_label 매핑을 추출합니다.

    Returns
    -------
    {"01": "사모사채 개요", "02": "금융개요", ...}
    """
    labels: dict = {}
    for page in pages:
        num = page.get("section_num", "").strip()
        lbl = page.get("section_label", "").strip()
        if num and lbl and num not in labels:
            labels[num] = lbl
    return labels


_NOISE_PREFIX = re.compile(r'^[\s\d.▶•■▪◆→☞※\[\]]+')
_TABLE_NOISE  = re.compile(r'\s{2,}|\s+구\s+분\s+|\s+내\s+용\s+')


def _clean_subtitle_for_title(raw: str) -> str:
    """
    PDF 원문 subtitle에서 핵심 명사구를 추출합니다.
    - 'N.N ' 숫자 접두사, ■/▶ 불릿, 표 헤더 반복 패턴 제거
    - 최대 12자로 잘라 반환
    """
    s = _SUBTITLE_NUM_PREFIX.sub("", raw).strip()
    s = _NOISE_PREFIX.sub("", s).strip()
    s = _TABLE_NOISE.sub(" ", s).strip()
    # 첫 공백 또는 특수문자 이전까지만 사용 (최대 12자)
    first_break = min(
        (s.find(c) for c in " 　•■▶▪" if c in s),
        default=len(s),
    )
    s = s[:first_break].strip()
    return s[:12] if s else ""


def _make_group2_title(subtitles: list, original_label: str = "") -> str:
    """
    분할된 그룹 2 소제목 목록에서 섹션 제목을 자동 생성합니다.

    1. 'N.N ' 형식 접두사를 제거한 뒤 핵심어 결합 시도
    2. 결과가 너무 짧거나 비어 있으면 original_label + ' (계속)' 반환
    예) ["3.3 분양사례 분석", "3.4 차주 개요"] → "분양사례 분석 및 차주 개요"
    """
    cleaned = [_clean_subtitle_for_title(s) for s in subtitles]
    cleaned = [c for c in cleaned if c]   # 빈 문자열 제거

    if not cleaned:
        return f"{original_label} (계속)" if original_label else "기타"
    if len(cleaned) == 1:
        candidate = cleaned[0]
    elif len(cleaned) == 2:
        candidate = f"{cleaned[0]} 및 {cleaned[1]}"
    else:
        candidate = f"{cleaned[0]} 등"

    # 너무 짧거나(2자 이하) 숫자만 남은 경우 fallback
    if len(candidate.replace(" ", "")) <= 2:
        return f"{original_label} (계속)" if original_label else "기타"
    return candidate


def split_into_5_sections(toc_4_map: dict, section_labels_4: dict = None) -> tuple:
    """
    4섹션 구조를 5섹션으로 자동 재구성합니다.

    분할 기준:
      - Appendix/별첨이 아닌 섹션 중 소제목이 가장 많은 섹션을 절반으로 분할.
      - toc_4_map이 비어 있으면 DEFAULT_TOC_MAP을 기준으로 분할합니다.
      - 동률이면 뒤 번호 우선.
      - 분할 대상 섹션 이후의 모든 섹션 번호를 +1.

    Parameters
    ----------
    toc_4_map        : {"01": [...], "02": [...], "03": [...], "04": [...]}
                       빈 dict도 허용 — DEFAULT_TOC_MAP으로 fallback
    section_labels_4 : {"01": "사모사채 개요", ..., "04": "Appendix"} — 없으면 자동 생략

    Returns
    -------
    (toc_5_map, labels_5, split_info)

    toc_5_map  : {"01": [...], ..., "05": [...]}
    labels_5   : {"01": "title", ..., "05": "title"}
    split_info : {
        "orig_num"       : str   — 분할된 원본 섹션 번호 (예: "03"),
        "new_second_num" : str   — 분할로 생성된 새 번호 (예: "04"),
        "second_subs"    : set   — 새 섹션으로 이동한 소제목 집합 (DEFAULT_TOC_MAP 기준),
        "renumbered"     : dict  — {old_num: new_num} 번호 밀린 섹션 매핑,
    }
    """
    labels = section_labels_4 or {}

    # toc_4_map이 비어 있으면 DEFAULT_TOC_MAP을 기준으로 분할
    effective_toc = toc_4_map if toc_4_map else DEFAULT_TOC_MAP

    # subtitles 없는 섹션(Appendix 등)도 포함 — labels 키까지 합산
    sorted_nums = sorted(set(effective_toc.keys()) | set(labels.keys()))

    if not sorted_nums:
        raise ValueError("toc_4_map이 비어 있고 section_labels_4도 없습니다.")

    # ── 1. 분할 대상 선택 ────────────────────────────────
    # 우선순위: (appendix여부, -소제목수, 번호) 오름차순 → min 선택
    def _priority(num: str):
        lbl         = labels.get(num, "").lower()
        is_appendix = any(k in lbl for k in _APPENDIX_KEYWORDS)
        cnt         = len(effective_toc.get(num, []))
        return (1 if is_appendix else 0, -cnt, num)

    split_target = min(sorted_nums, key=_priority)
    target_subs  = list(effective_toc[split_target])

    # ── 2. 절반 분할 (홀수면 앞쪽이 1개 더) ────────────
    half        = (len(target_subs) + 1) // 2
    subs_first  = target_subs[:half]
    subs_second = target_subs[half:]

    # fallback: 소제목 1개뿐이라 분할 불가 → 마지막 1개를 강제 이동
    if not subs_second:
        subs_first  = target_subs[:-1] if len(target_subs) > 1 else target_subs
        subs_second = target_subs[-1:] if len(target_subs) > 1 else target_subs[-1:]

    # ── 3. 번호 재할당 ────────────────────────────────────
    split_num_int  = int(split_target)
    new_second_num = f"{split_num_int + 1:02d}"
    renumbered     = {
        num: f"{int(num) + 1:02d}"
        for num in sorted_nums
        if int(num) > split_num_int
    }

    # ── 4. toc_5_map 구성 (소제목 번호 재할당 포함) ────────────────────
    _SUB_PREFIX_RE = re.compile(r'^(\d+)\.(\d+)\s*')

    def _renumber_split(subs: list, new_sec: str) -> list:
        """그룹2 소제목: N.M → new_sec.{1,2,...} 재할당 (번호 초기화)"""
        result = []
        for idx, s in enumerate(subs, 1):
            m = _SUB_PREFIX_RE.match(s)
            rest = s[m.end():] if m else s
            result.append(f"{int(new_sec)}.{idx} {rest}".strip())
        return result

    def _renumber_shift(subs: list, old_sec: str, new_sec: str) -> list:
        """이동된 섹션 소제목: N.M → (N+1).M 재할당 (부번호 유지)"""
        result = []
        for s in subs:
            m = _SUB_PREFIX_RE.match(s)
            if m and m.group(1) == str(int(old_sec)):
                rest = s[m.end():]
                result.append(f"{int(new_sec)}.{m.group(2)} {rest}".strip())
            else:
                result.append(s)
        return result

    toc_5_map: dict = {}
    for num in sorted_nums:
        subs = list(effective_toc.get(num, []))   # subtitles 없는 섹션은 []
        if num == split_target:
            toc_5_map[split_target] = subs_first   # 앞 그룹은 원래 번호 유지
        elif num in renumbered:
            new_num = renumbered[num]
            toc_5_map[new_num] = _renumber_shift(subs, num, new_num)
        else:
            toc_5_map[num] = subs
    toc_5_map[new_second_num] = _renumber_split(subs_second, new_second_num)

    # ── 5. labels_5 구성 ─────────────────────────────────
    labels_5: dict = {}
    for num in sorted_nums:
        if num == split_target:
            labels_5[split_target] = labels.get(split_target, f"섹션 {split_target}")
        elif num in renumbered:
            labels_5[renumbered[num]] = labels.get(num, f"섹션 {renumbered[num]}")
        else:
            labels_5[num] = labels.get(num, f"섹션 {num}")
    labels_5[new_second_num] = _make_group2_title(
        subs_second, original_label=labels.get(split_target, "")
    )

    split_info = {
        "orig_num":       split_target,
        "new_second_num": new_second_num,
        "second_subs":    set(subs_second),
        "renumbered":     renumbered,
    }

    # ── 콘솔 보고 ────────────────────────────────────────
    print(f"\n[split_into_5_sections] 분할 대상: [{split_target}] {labels.get(split_target, '?')}")
    print(f"  그룹1 소제목 ({len(subs_first)}개): {subs_first}")
    print(f"  그룹2 소제목 ({len(subs_second)}개): {subs_second}")
    print(f"  그룹2 자동 생성 제목: '{labels_5[new_second_num]}'")
    print(f"  번호 재할당: {renumbered}")
    print(f"\n  최종 5섹션 구조:")
    for k in sorted(toc_5_map):
        print(f"    [{k}] {labels_5.get(k, '?')}: {toc_5_map[k]}")

    return toc_5_map, labels_5, split_info


def remap_pages_for_5sections(pages: list, split_info: dict, labels_5: dict) -> list:
    """
    pages의 section_num / section_label을 5섹션 구조에 맞게 재할당합니다.
    원본 pages 리스트는 변경하지 않고 새 목록을 반환합니다.

    분할 경계 판정 규칙 (페이지 수 기반):
      - orig_num에 속하는 페이지들을 페이지 수 기준으로 절반 분할합니다.
      - 앞 절반(올림)은 orig_num 유지, 나머지는 new_second_num으로 이동합니다.
      - subtitle 텍스트 매칭을 사용하지 않으므로 PDF 본문이 오염될 위험이 없습니다.

    Parameters
    ----------
    pages      : parse_document_from_bytes() 반환값
    split_info : split_into_5_sections() 반환 split_info
    labels_5   : split_into_5_sections() 반환 labels_5

    Returns
    -------
    재할당된 새 pages 목록 (원본 불변)
    """
    orig_num       = split_info["orig_num"]
    new_second_num = split_info["new_second_num"]
    renumbered     = split_info["renumbered"]     # {old: new}

    # orig_num 섹션에 속하는 페이지의 전체 인덱스를 수집
    orig_indices = [i for i, p in enumerate(pages)
                    if p.get("section_num", "").strip() == orig_num]

    # 앞 절반(올림)은 orig_num 유지, 나머지는 new_second_num
    split_at = (len(orig_indices) + 1) // 2
    second_group = set(orig_indices[split_at:])

    remapped = []
    for i, page in enumerate(pages):
        p       = dict(page)   # shallow copy — 원본 불변
        sec_num = p.get("section_num", "").strip()

        if sec_num == orig_num:
            if i in second_group:
                p["section_num"]   = new_second_num
                p["section_label"] = labels_5.get(new_second_num, p.get("section_label", ""))
            else:
                p["section_label"] = labels_5.get(orig_num, p.get("section_label", ""))
        elif sec_num in renumbered:
            new_num = renumbered[sec_num]
            p["section_num"]   = new_num
            p["section_label"] = labels_5.get(new_num, p.get("section_label", ""))

        remapped.append(p)

    return remapped


def _extract_page_tables(plumber_page, debug: bool = False) -> dict:
    """
    pdfplumber로 한 페이지에서 표를 추출합니다.

    Returns
    -------
    dict:
      tables              : List[List[List[str]]] — 페이지 내 표 목록(행×열)
      table_bboxes        : List[Tuple] — 각 표의 (x0, top, x1, bottom) PDF 좌표
      text_without_tables : str — 표 영역·헤더(상단 20%) 제외 본문 텍스트
    """
    tables: list = []
    table_bboxes: list = []

    # ── 다단계 추출 → 후처리(팬텀 열 제거) → 유효성 검사 순 ─────────────
    #   정제(_postprocess_table) 후의 표를 검증해야 phantom column 때문에 멀쩡한
    #   표가 탈락하지 않는다.
    try:
        for extracted, bbox, strat in _extract_tables_dual_strategy(plumber_page, debug=debug):
            processed = _postprocess_table(extracted)
            if not processed:
                continue
            if not _is_valid_table(processed):
                if debug:
                    rows = len(processed)
                    cols = max((len(r) for r in processed), default=0)
                    print(f"  [제외] 유효성 검사 실패 {rows}행×{cols}열 전략={strat} bbox={bbox}")
                continue
            tables.append(processed)
            table_bboxes.append(bbox)
    except Exception as exc:
        if debug:
            print(f"  [_extract_page_tables] 오류: {exc}")

    # ── 표·헤더 영역 제외 텍스트 추출 ──────────────────────
    text_without_tables = ""
    try:
        header_thresh_y = plumber_page.height * 0.20
        words = plumber_page.extract_words() or []
        non_table_words = []
        for word in words:
            # 헤더 영역(상단 20%) 제외
            if word["top"] < header_thresh_y:
                continue
            # 표 영역 제외 (3px 여유)
            in_table = any(
                bbox[0] - 3 <= word["x0"] and word["x1"] <= bbox[2] + 3
                and bbox[1] - 3 <= word["top"] and word["bottom"] <= bbox[3] + 3
                for bbox in table_bboxes
            )
            if not in_table:
                non_table_words.append(word["text"])
        text_without_tables = " ".join(non_table_words)
    except Exception as exc:
        if debug:
            print(f"  [_extract_page_tables] 텍스트 추출 오류: {exc}")

    if debug and tables:
        for idx, tbl in enumerate(tables):
            cols = max((len(r) for r in tbl), default=0)
            print(f"  [표 {idx+1}] {len(tbl)}행 × {cols}열  bbox={table_bboxes[idx]}")

    return {
        "tables":              tables,
        "table_bboxes":        table_bboxes,
        "text_without_tables": text_without_tables,
    }
