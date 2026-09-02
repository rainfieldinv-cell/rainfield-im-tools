import streamlit as st

# 추출 함수 및 UI 컴포넌트 불러오기
from extractors import extract_from_pdf, extract_from_docx, detect_business_name, get_file_type
from ui_components import render_stepper, render_image_gallery, render_text_preview

# PPT 생성 모듈 불러오기
import io
import os
import re
from datetime import datetime
from modules.page_builders import (
    make_output_filename,
    build_full_presentation,
    build_preview_presentation,
)
from modules.content_parser import (
    parse_document_from_bytes,
    extract_toc_map,
    extract_section_labels,
    split_into_5_sections,
    remap_pages_for_5sections,
)

# ─────────────────────────────────────────────
# [페이지 기본 설정]
# - page_title : 브라우저 탭에 표시되는 제목 (여기를 수정하면 탭 제목이 바뀝니다)
# - layout     : "wide" = 화면 전체 너비 사용 / "centered" = 가운데 정렬 좁은 화면
# - initial_sidebar_state : "collapsed" = 사이드바 기본 접힘 / "expanded" = 기본 펼침
# ─────────────────────────────────────────────
# ★'IM 도구' 첫 화면(app.py)이 이미 정해 놓았다. 한 앱에서 두 번은 안 되므로,
#   혼자 돌릴 때만 실행한다.
try:
    st.set_page_config(
        page_title="IM 생성기",
        layout="wide",
        initial_sidebar_state="collapsed",
    )
except Exception:
    pass

# ─────────────────────────────────────────────
# [접근 코드 불러오기]
# - secrets.toml 파일에서 ACCESS_CODE 값을 읽어옵니다
# - 파일이 없거나 키가 없으면 기본값 'rainfield2026' 사용
# - 실제 운영 시에는 secrets.toml 에서만 관리하세요
# ─────────────────────────────────────────────
try:
    CORRECT_CODE = st.secrets["ACCESS_CODE"]
except Exception:
    CORRECT_CODE = "rainfield2026"  # ← 여기를 수정하면 기본 접근 코드가 바뀝니다

# ─────────────────────────────────────────────
# [Anthropic API 키 다리]
# 클라우드에서는 키를 st.secrets["ANTHROPIC_API_KEY"]에 넣는다(Settings→Secrets).
# LLM 엔진(modules/claude_api.py)은 os.environ["ANTHROPIC_API_KEY"]만 읽으므로
# secrets → 환경변수로 넘겨준다. 로컬(.env)에 이미 있으면 건드리지 않는다.
# ─────────────────────────────────────────────
try:
    _k = str(st.secrets.get("ANTHROPIC_API_KEY", "")).strip()
    if _k and not os.environ.get("ANTHROPIC_API_KEY"):
        os.environ["ANTHROPIC_API_KEY"] = _k
except Exception:
    pass

# ─────────────────────────────────────────────
# [세션 상태 초기화]
# - st.session_state : 페이지가 새로고침돼도 값이 유지되는 저장공간
# - logged_in : 로그인 여부 (True = 로그인됨, False = 로그아웃 상태)
# ─────────────────────────────────────────────
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

# ─────────────────────────────────────────────
# [변환 작업 관련 세션 상태 초기화]
# 새로고침해도 데이터가 유지됩니다
# ─────────────────────────────────────────────
if "current_step" not in st.session_state:
    st.session_state.current_step = 1       # 현재 진행 단계 (1~9)
if "uploaded_file" not in st.session_state:
    st.session_state.uploaded_file = None   # 업로드된 파일 객체
if "extracted_data" not in st.session_state:
    st.session_state.extracted_data = None  # 추출된 텍스트·이미지 데이터
if "business_name" not in st.session_state:
    st.session_state.business_name = ""     # 자동 감지된 사업명

# 3단계 관련 세션 상태
if "toc_count" not in st.session_state:
    st.session_state.toc_count = 4          # 목차 개수 (4 또는 5)
if "month_en" not in st.session_state:
    st.session_state.month_en = datetime.now().strftime("%B")  # 현재 월 영문
if "year" not in st.session_state:
    st.session_state.year = str(datetime.now().year)           # 현재 연도
if "cover_image_index" not in st.session_state:
    st.session_state.cover_image_index = 0  # 표지에 쓸 이미지 번호
if "cover_image_bytes" not in st.session_state:
    st.session_state.cover_image_bytes = None  # 선택된 표지 이미지 bytes
if "parsed_pages" not in st.session_state:
    st.session_state.parsed_pages = []         # content_parser 결과
if "pdf_bytes" not in st.session_state:
    st.session_state.pdf_bytes = None          # PDF 원본 bytes(좌표기반 사진/표복원용)
if "ppt_bytes" not in st.session_state:
    st.session_state.ppt_bytes = None          # 생성된 PPT bytes(5단계 내용검수용)

# 섹션 이미지 관련 세션 상태 — 원형 슬롯 3개 (4개 섹션 divider 공통 적용)
if "section_img_idx_list" not in st.session_state:
    st.session_state.section_img_idx_list = [0, 0, 0]
if "section_img_bytes_list" not in st.session_state:
    st.session_state.section_img_bytes_list = [None, None, None]

# 목차 이미지 관련 세션 상태 — 원형 슬롯 1개
if "toc_img_idx" not in st.session_state:
    st.session_state.toc_img_idx = 0
if "toc_img_bytes" not in st.session_state:
    st.session_state.toc_img_bytes = None


# ─────────────────────────────────────────────
# [로그인 화면 함수]
# ─────────────────────────────────────────────
def show_login():
    # 화면 가운데 정렬을 위해 3개 컬럼 중 가운데만 사용
    col1, col2, col3 = st.columns([1, 1.5, 1])

    with col2:
        st.markdown("<br><br>", unsafe_allow_html=True)

        # 로그인 화면 제목 (여기를 수정하면 로그인 화면 제목이 바뀝니다)
        st.markdown(
            "<h2 style='text-align:center;'>IM 생성기</h2>",
            unsafe_allow_html=True,
        )
        st.markdown(
            "<p style='text-align:center; color:gray;'>접근 코드를 입력하세요</p>",
            unsafe_allow_html=True,
        )
        st.markdown("<br>", unsafe_allow_html=True)

        # st.form : 폼 안에서 Enter 키를 누르면 버튼과 동일하게 동작합니다
        with st.form("login_form"):
            # 접근 코드 입력 필드
            # type="password" : 입력값이 ●●● 로 가려집니다
            code_input = st.text_input(
                label="접근 코드",
                type="password",
                placeholder="접근 코드 입력",
                label_visibility="collapsed",
            )

            # 로그인 버튼 (여기를 수정하면 버튼 텍스트가 바뀝니다)
            login_btn = st.form_submit_button("로그인", use_container_width=True, type="primary")

        if login_btn:
            if code_input == CORRECT_CODE:
                # 올바른 코드 → 로그인 성공
                st.session_state.logged_in = True
                st.rerun()  # 화면을 메인으로 전환
            else:
                # 잘못된 코드 → 빨간 오류 메시지
                st.error("❌ 접근 코드가 올바르지 않습니다. 다시 확인해주세요.")


# ─────────────────────────────────────────────
# [1단계: 파일 업로드 화면]
# ─────────────────────────────────────────────
def show_step1():
    st.markdown("## 1단계. 파일 업로드")
    st.markdown(
        "📑 **원본 PDF를 올려주세요 — 권장.** 글·표·사진·페이지를 전부 원본 PDF 기준으로 처리합니다"
        "(네가 보는 페이지 = 내용 페이지가 정확히 일치).  \n"
        "📄 워드(.docx)는 **PDF가 아예 없을 때만** 올리면 돼요(자동으로 PDF로 변환해 사용).  \n"
        "<span style='color:gray; font-size:13px;'>(PDF / Word(.doc·.docx) 지원 · PDF 우선)</span>",
        unsafe_allow_html=True,
    )
    st.markdown("")

    # 파일 업로더 (PDF 우선 — PDF 있으면 PDF로 전부 처리, 없으면 워드를 PDF로 변환)
    uploaded_files = st.file_uploader(
        label="원본 PDF를 여기에 끌어다 놓거나 클릭해서 선택하세요 (PDF 우선 · PDF 없으면 워드)",
        type=["pdf", "docx", "doc"],
        accept_multiple_files=True,
        key="file_uploader_widget",
    )

    if uploaded_files:
        # 워드/PDF 분리 — PDF 있으면 PDF로, 없으면 워드를 PDF로 변환해 처리
        pdf_up = next((f for f in uploaded_files if f.name.lower().endswith(".pdf")), None)
        word_up = next((f for f in uploaded_files
                        if f.name.lower().endswith((".docx", ".doc"))), None)

        # 올린 파일 정보 표시
        total_mb = sum(f.size for f in uploaded_files) / (1024 * 1024)
        if total_mb >= 50:
            st.warning("⚠️ 파일이 큽니다. 추출에 시간이 걸릴 수 있습니다.")

        col1, col2 = st.columns(2)
        with col1:
            st.metric("올린 파일", f"{len(uploaded_files)}개")
        with col2:
            st.metric("합계 크기", f"{total_mb:.2f} MB")
        st.caption("파일: " + " · ".join(f.name for f in uploaded_files))
        if pdf_up:
            if word_up:
                st.info("PDF를 올려서 **PDF 기준**으로 처리합니다(워드는 사용 안 함). "
                        "글·표·사진·페이지가 모두 원본 PDF와 정확히 일치해요.")
            else:
                st.success("원본 PDF로 처리합니다 — 좋아요. 👍")
        elif word_up:
            st.info("PDF가 없어 **워드를 PDF로 변환**해서 처리합니다. "
                    "(원본 PDF가 있으면 그걸 올리는 게 페이지·사진이 더 정확해요)")

        st.markdown("")

        # 추출 시작 버튼 (이미 추출했으면 '다시 추출'로 표시 — 다음 단계 버튼과 헷갈리지 않게)
        _done = bool(st.session_state.get("extracted_data"))
        if st.button("🔄 다시 추출" if _done else "🔍 추출 시작",
                     type=("secondary" if _done else "primary"),
                     use_container_width=False,
                     help="올린 파일에서 글자·사진을 다시 뽑습니다." if _done else None):
            try:
                # ★(A) 원본 PDF 기준 통일: PDF가 있으면 뷰어·파싱·이미지·글 '전부' 원본 PDF로.
                #   → 네가 보는 페이지 = 내용 페이지가 100% 일치(짤림·어긋남 없음).
                #   워드는 'PDF가 없을 때만' 백업으로 PDF 변환해서 사용.
                proc_pdf = None
                proc_from_word = False   # 원본 PDF가 없어 워드→PDF로 대체했는지
                conv_err = None
                if pdf_up is not None:
                    proc_pdf = pdf_up.getvalue()             # PDF 우선(모든 소스 = 이 파일)
                    proc_name = pdf_up.name
                elif word_up is not None:
                    from modules.preview import convert_word_to_pdf
                    with st.spinner("PDF가 없어 워드를 PDF로 변환하는 중입니다..."):
                        proc_pdf, conv_err = convert_word_to_pdf(word_up.getvalue(), word_up.name)
                    proc_from_word = proc_pdf is not None
                    proc_name = word_up.name

                with st.spinner("파일에서 텍스트와 이미지를 추출하는 중입니다..."):
                    if proc_pdf is not None:
                        # 뷰어·파싱·이미지·글 전부 동일 소스(proc_pdf) → 페이지 완벽 정렬
                        result = extract_from_pdf(proc_pdf)
                        st.session_state.pdf_bytes = proc_pdf
                        parse_bytes, parse_name = proc_pdf, "source.pdf"
                        st.session_state.view_pdf_bytes = proc_pdf
                    else:
                        # PDF도 없고 워드 변환도 실패 → python-docx 텍스트만(사진 제한)
                        result = extract_from_docx(word_up.getvalue())
                        st.session_state.pdf_bytes = None
                        st.session_state.view_pdf_bytes = None
                        parse_bytes, parse_name = word_up.getvalue(), word_up.name
                        if conv_err:
                            st.warning(f"워드→PDF 변환 실패 — 텍스트만 추출(사진 제한). 원인: {conv_err}")

                # 추출 결과 저장
                st.session_state.extracted_data = result
                st.session_state.uploaded_file  = proc_name

                # 사업명 자동 감지
                detected = detect_business_name(result["pages_text"])
                st.session_state.business_name = detected

                # 슬라이드 구조 파싱 (content_parser) — 원본 PDF 기준
                try:
                    st.session_state.parsed_pages = parse_document_from_bytes(
                        parse_bytes, parse_name
                    )
                except Exception:
                    st.session_state.parsed_pages = []

                # 어떤 파일에서 뽑았는지 기록(화면 표시용)
                _src = "워드(.docx)" if proc_from_word else "PDF"
                st.session_state.src_info = {"text": _src, "image": _src}

                # 안내: PDF+워드 둘 다 올리면 'PDF 기준'으로만 처리(워드 미사용)
                if pdf_up is not None and word_up is not None:
                    st.session_state.conv_warn = ("PDF를 올려서 **PDF 기준**으로 처리했어요"
                                                  "(워드는 사용 안 함 — 페이지가 원본과 정확히 일치).")
                elif word_up is not None and proc_pdf is None:
                    st.session_state.conv_warn = (f"워드→PDF 변환 실패로 텍스트만 추출됐어요"
                                                  f"(사진 제한). 원인: {conv_err or '알 수 없음'}")
                else:
                    st.session_state.conv_warn = None

                # 새 문서이므로 이전 목차 편집·이미지 상태 초기화(3단계 새로 시작)
                for _k in ("toc_edit", "toc_hashtags", "toc_page_cache", "toc_view_page",
                           "toc_png_cache", "view_pdf_bytes", "highlight_cards", "hl_img_render",
                           "hl_done_render", "hl_view_page", "layout_preview_render",
                           "_hl_autofill_msg", "_toc_autofill_msg",
                           "hl_es_pages_str", "hl_es_idx"):
                    st.session_state.pop(_k, None)
                _clear_toc_widget_state()

                # 2단계 없이 1단계에 머물며 사업명 입력받음
                st.rerun()

            except ValueError as e:
                # 비밀번호 걸린 PDF
                st.error(f"🔒 {e}")
            except RuntimeError as e:
                # 손상된 파일
                st.error(f"❌ {e}")
            except Exception as e:
                st.error(f"❌ 예상치 못한 오류가 발생했습니다: {e}")

    # ────────────────────────────────────────
    # 추출 완료 후 — 사업명 입력 + 다음 단계 (2단계 없이 여기서 처리)
    # ────────────────────────────────────────
    data = st.session_state.get("extracted_data")
    if data:
        st.markdown("---")
        st.markdown("### 📌 추출 완료")
        _src = st.session_state.get("src_info") or {}
        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("총 페이지 수", f"{data['total_pages']}페이지")
        with c2:
            st.metric("추출된 이미지 수", f"{len(data['images'])}개")
        with c3:
            st.metric("처리 방식", _src.get("text") or data["file_type"].upper())
        # 어디서 뽑았는지 명확히
        if _src:
            st.caption(f"📄 글·표·사진 모두 → **{_src.get('text','-')}** 에서 추출했습니다. "
                       "(원본 PDF 기준이라 네가 보는 페이지와 내용이 정확히 일치)")

        _cw = st.session_state.get("conv_warn")
        if _cw:
            st.warning(f"⚠️ {_cw}")

        for w in data.get("warnings", []):
            st.warning(f"⚠️ {w}")

        st.markdown("")
        if not st.session_state.business_name:
            st.warning("⚠️ 사업명을 자동으로 인식하지 못했습니다. 직접 입력해주세요.")
        st.session_state.business_name = st.text_input(
            label="사업명",
            value=st.session_state.business_name,
            placeholder="예) 천안 부성2지구 도시개발사업",
            help="자동 감지된 사업명입니다. 잘못된 경우 직접 수정해주세요.",
        )

        st.markdown("")
        _n1, _ns, _n2 = st.columns([2, 4, 2])
        with _n2:
            _dis = not bool(st.session_state.business_name.strip())
            if _dis:
                st.caption("사업명을 입력해주세요.")
            if st.button("다음 단계 →", use_container_width=True, type="primary",
                         disabled=_dis):
                st.session_state.current_step = 2
                st.rerun()


def _make_divider_miniature(w_px: int = 540, h_px: int = 374) -> bytes:
    """섹션 divider 슬라이드 레이아웃을 단순화한 PNG 미니어처를 반환합니다.
    ①②③ 표시로 각 원형 슬롯 위치를 직관적으로 안내합니다."""
    import io as _mio
    from PIL import Image as _Img, ImageDraw as _Draw, ImageFont as _Font

    SLIDE_W = 27.52
    SLIDE_H = 19.05
    sx, sy = w_px / SLIDE_W, h_px / SLIDE_H

    img  = _Img.new("RGB", (w_px, h_px), (248, 248, 248))
    draw = _Draw.Draw(img)
    draw.rectangle([0, 0, w_px - 1, h_px - 1], outline=(200, 200, 200), width=2)

    # 텍스트 영역 힌트 (섹션 번호+제목 위치)
    draw.rectangle(
        [int(0.5 * sx), int(7.5 * sy), int(13 * sx), int(10.5 * sy)],
        fill=(225, 225, 225), outline=(190, 190, 190),
    )

    # 3개 원형 슬롯 (outer oval 위치 기준) — _DIV_OVAL_PAIRS와 동일 순서
    _OVALS  = [(16.3273, 0.7346, 9.70, 9.70),
               (17.3779, 8.2234, 8.70, 8.70),
               (13.5919, 6.1749, 6.80, 6.80)]
    _LABELS = ["①", "②", "③"]
    _GREEN  = (146, 208, 80)

    try:
        _fnt_path = os.path.join(os.path.dirname(__file__), "fonts", "PEOPLEFONTB.TTF")
        _base_fnt = _Font.truetype(_fnt_path, size=28)
    except Exception:
        _base_fnt = _Font.load_default()

    for (ol, ot, ow, oh), label in zip(_OVALS, _LABELS):
        cx = int((ol + ow / 2) * sx)
        cy = int((ot + oh / 2) * sy)
        rx = int(ow / 2 * sx)
        ry = int(oh / 2 * sy)
        draw.ellipse([cx - rx, cy - ry, cx + rx, cy + ry],
                     fill=(220, 220, 220), outline=_GREEN, width=5)
        draw.text((cx, cy), label, font=_base_fnt, fill=(50, 50, 50), anchor="mm")

    buf = _mio.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


def _make_toc_miniature(w_px: int = 540, h_px: int = 374) -> bytes:
    """목차 슬라이드 레이아웃을 단순화한 PNG 미니어처를 반환합니다.
    ① 표시로 원형 슬롯 위치를 직관적으로 안내합니다."""
    import io as _mio
    from PIL import Image as _Img, ImageDraw as _Draw, ImageFont as _Font

    SLIDE_W = 27.52
    SLIDE_H = 19.05
    sx, sy = w_px / SLIDE_W, h_px / SLIDE_H

    img  = _Img.new("RGB", (w_px, h_px), (248, 248, 248))
    draw = _Draw.Draw(img)
    draw.rectangle([0, 0, w_px - 1, h_px - 1], outline=(200, 200, 200), width=2)

    # 오른쪽 목차 항목 힌트 (4개 행)
    for row in range(4):
        y_top = int((2.0 + row * 4.1) * sy)
        draw.rectangle(
            [int(7.5 * sx), y_top, int(26.5 * sx), int(y_top + 3.2 * sy)],
            fill=(225, 225, 225), outline=(190, 190, 190),
        )

    # 1개 원형 슬롯 — TOC oval 위치 (left=0.84, top=6.95, w=5.26, h=5.30)
    _GREEN = (146, 208, 80)
    ol, ot, ow, oh = 0.8409, 6.9520, 5.2600, 5.3000
    cx = int((ol + ow / 2) * sx)
    cy = int((ot + oh / 2) * sy)
    rx = int(ow / 2 * sx)
    ry = int(oh / 2 * sy)

    try:
        _fnt_path = os.path.join(os.path.dirname(__file__), "fonts", "PEOPLEFONTB.TTF")
        _base_fnt = _Font.truetype(_fnt_path, size=28)
    except Exception:
        _base_fnt = _Font.load_default()

    draw.ellipse([cx - rx, cy - ry, cx + rx, cy + ry],
                 fill=(220, 220, 220), outline=_GREEN, width=5)
    draw.text((cx, cy), "①", font=_base_fnt, fill=(50, 50, 50), anchor="mm")

    buf = _mio.BytesIO()
    img.save(buf, format="PNG")
    buf.seek(0)
    return buf.getvalue()


# ─────────────────────────────────────────────
# [3단계: 레이아웃 및 표지 미리 생성 화면]
# ─────────────────────────────────────────────
def show_step3():
    data = st.session_state.extracted_data
    if data is None:
        st.warning("추출 데이터가 없습니다. 1단계로 돌아가주세요.")
        if st.button("← 1단계로"):
            st.session_state.current_step = 1
            st.rerun()
        return

    st.markdown("## 5단계. 레이아웃 및 표지 미리 생성")
    st.caption("날짜·표지 이미지를 지정하고 표지·목차·섹션을 미리 확인하세요. "
               "(목차 형식 4/5개는 3단계 '목차 구성'에서 정합니다)")
    st.markdown("")

    # 목차 형식(4/5)은 4단계 목차 구성에서 정한 값을 그대로 사용 (여기서 다시 고르지 않음)
    st.caption(f"현재 목차 형식: **{st.session_state.toc_count}개 형식** (3단계에서 설정)")

    st.markdown("---")

    # ────────────────────────────────────────
    # (B) 날짜 정보 자동 입력
    # ────────────────────────────────────────
    st.markdown("### 📅 날짜 정보")

    col_month, col_year, _ = st.columns([2, 2, 4])
    with col_month:
        # 여기를 수정하면 영문 월 기본값이 바뀝니다
        month_input = st.text_input(
            "영문 월",
            value=st.session_state.month_en,
            placeholder="예) March",
            help="표지에 표시될 영문 월 이름입니다.",
        )
        st.session_state.month_en = month_input

    with col_year:
        # 여기를 수정하면 연도 기본값이 바뀝니다
        year_input = st.text_input(
            "연도",
            value=st.session_state.year,
            placeholder="예) 2026",
            help="표지에 표시될 연도입니다.",
        )
        st.session_state.year = year_input

    st.markdown("---")

    # ────────────────────────────────────────
    # (C) 표지 이미지 선택
    # ────────────────────────────────────────
    st.markdown("### 🖼️ 표지 이미지 선택")

    images = data.get("images", [])

    if not images:
        st.warning("추출된 이미지가 없습니다. 표지 이미지 없이 생성됩니다.")
        cover_image_bytes = None
    else:
        # 가장 큰 이미지를 자동 추천 (조감도일 가능성 높음)
        largest = max(images, key=lambda x: x["width"] * x["height"])
        default_idx = largest["index"] + 1  # 1부터 시작하는 번호

        st.caption(f"총 {len(images)}개 이미지 추출됨 — 가장 큰 이미지(#{default_idx})를 자동 추천합니다.")

        # ★섹션/목차 이미지 선택과 동일하게 expander(접기/펼치기)로 통일
        with st.expander("표지 이미지 선택"):
            # 이미지 미리보기 — 처음 8장 + 나머지는 expander
            cols = st.columns(4)
            for i, img_data in enumerate(images[:8]):
                with cols[i % 4]:
                    st.image(img_data["pil_image"], use_container_width=True)
                    st.caption(f"#{img_data['index']+1} ({img_data['width']}×{img_data['height']})")

            if len(images) > 8:
                with st.expander(f"나머지 {len(images)-8}개 이미지 더 보기"):
                    _more_cols = st.columns(4)
                    for i, img_data in enumerate(images[8:]):
                        with _more_cols[i % 4]:
                            st.image(img_data["pil_image"], use_container_width=True)
                            st.caption(f"#{img_data['index']+1} ({img_data['width']}×{img_data['height']})")

            # 이미지 번호 입력
            img_num = st.number_input(
                "표지에 사용할 이미지 번호",
                min_value=1,
                max_value=len(images),
                value=min(default_idx, len(images)),
                step=1,
                help="위 갤러리에서 원하는 이미지 번호를 입력하세요.",
            )
            st.session_state.cover_image_index = int(img_num) - 1

            # 선택된 이미지 미리보기
            selected = images[st.session_state.cover_image_index]
            st.markdown("**선택된 표지 이미지:**")
            st.image(selected["pil_image"], width=500)

            # PIL Image → bytes 변환 후 세션에 저장
            import io as _io
            buf = _io.BytesIO()
            selected["pil_image"].save(buf, format="PNG")
            cover_image_bytes = buf.getvalue()
            st.session_state.cover_image_bytes = cover_image_bytes

    st.markdown("---")

    # ────────────────────────────────────────
    # (D-2) 목차 이미지 선택 (선택 사항)
    # ────────────────────────────────────────
    st.markdown("### 📌 목차 이미지 선택 (선택 사항)")
    st.caption("목차 페이지의 원형 자리(1개)에 들어갈 이미지를 선택하세요.")

    if not images:
        st.info("추출된 이미지가 없어 목차 이미지를 선택할 수 없습니다.")
    else:
        import io as _io
        with st.expander("목차 이미지 선택 (원형 자리 1개)"):
            _tgcols = st.columns(3)
            for _i, _img in enumerate(images[:9]):
                with _tgcols[_i % 3]:
                    st.image(_img["pil_image"], use_container_width=True)
                    st.caption(f"#{_img['index']+1}")
            if len(images) > 9:
                with st.expander(f"나머지 {len(images)-9}개 이미지 더 보기"):
                    _more_tgcols = st.columns(3)
                    for _i, _img in enumerate(images[9:]):
                        with _more_tgcols[_i % 3]:
                            st.image(_img["pil_image"], use_container_width=True)
                            st.caption(f"#{_img['index']+1}")

            st.markdown("")

            st.image(_make_toc_miniature(),
                     caption="원형 슬롯 위치 — ① 좌측 중앙",
                     use_container_width=True)

            st.markdown("")

            _toc_prev = st.session_state.toc_img_idx
            _toc_col, _ = st.columns([1, 2])
            with _toc_col:
                st.caption("원형 자리 1번 (좌측 중앙)")
                _toc_sel_num = st.number_input(
                    "이미지 번호 (0 = 선택 안 함)",
                    min_value=0,
                    max_value=len(images),
                    value=_toc_prev,
                    step=1,
                    key="toc_oval_slot_0",
                )
                st.session_state.toc_img_idx = int(_toc_sel_num)
                if _toc_sel_num > 0:
                    _toc_sel = images[int(_toc_sel_num) - 1]
                    st.image(_toc_sel["pil_image"], use_container_width=True)
                    st.caption(f"선택됨: #{int(_toc_sel_num)}")
                    _tbuf = _io.BytesIO()
                    _toc_sel["pil_image"].save(_tbuf, format="PNG")
                    st.session_state.toc_img_bytes = _tbuf.getvalue()
                else:
                    st.caption("(선택 안 함)")
                    st.session_state.toc_img_bytes = None

    st.markdown("---")

    # ────────────────────────────────────────
    # (D) 섹션 이미지 선택 (선택 사항)
    # ────────────────────────────────────────
    st.markdown("### 📌 섹션 이미지 선택 (선택 사항)")
    st.caption("섹션 구분 페이지의 원형 자리(3개)에 들어갈 이미지를 선택하세요. 4개 섹션(01~04) 모두 동일하게 적용됩니다.")

    if not images:
        st.info("추출된 이미지가 없어 섹션 이미지를 선택할 수 없습니다.")
    else:
        import io as _io
        with st.expander("섹션 이미지 선택 (원형 자리 3개)"):
            # 이미지 갤러리 — 처음 9장 + 나머지는 expander
            _gcols = st.columns(3)
            for _i, _img in enumerate(images[:9]):
                with _gcols[_i % 3]:
                    st.image(_img["pil_image"], use_container_width=True)
                    st.caption(f"#{_img['index']+1}")
            if len(images) > 9:
                with st.expander(f"나머지 {len(images)-9}개 이미지 더 보기"):
                    _more_gcols = st.columns(3)
                    for _i, _img in enumerate(images[9:]):
                        with _more_gcols[_i % 3]:
                            st.image(_img["pil_image"], use_container_width=True)
                            st.caption(f"#{_img['index']+1}")

            st.markdown("")

            # 슬롯 위치 안내 미니어처
            st.image(_make_divider_miniature(),
                     caption="원형 슬롯 위치 — ① 우측 상단 / ② 우측 하단 / ③ 중앙 좌측",
                     use_container_width=True)

            st.markdown("")

            # 3개 슬롯 번호 입력 + 미리보기 (가로 3열)
            _prev_list = st.session_state.section_img_idx_list
            _new_idx_list   = []
            _new_bytes_list = []

            _scols = st.columns(3)
            _slot_labels = ["원형 자리 1번 (우측 상단)", "원형 자리 2번 (우측 하단)", "원형 자리 3번 (중앙 좌측)"]
            for _slot in range(3):
                with _scols[_slot]:
                    st.caption(_slot_labels[_slot])
                    _prev = _prev_list[_slot] if _slot < len(_prev_list) else 0
                    _sel_num = st.number_input(
                        f"이미지 번호 (0 = 선택 안 함)",
                        min_value=0,
                        max_value=len(images),
                        value=_prev,
                        step=1,
                        key=f"sec_oval_slot_{_slot}",
                    )
                    _new_idx_list.append(int(_sel_num))
                    if _sel_num > 0:
                        _sel = images[int(_sel_num) - 1]
                        st.image(_sel["pil_image"], use_container_width=True)
                        st.caption(f"선택됨: #{int(_sel_num)}")
                        _buf = _io.BytesIO()
                        _sel["pil_image"].save(_buf, format="PNG")
                        _new_bytes_list.append(_buf.getvalue())
                    else:
                        st.caption("(선택 안 함)")
                        _new_bytes_list.append(None)

            st.session_state.section_img_idx_list   = _new_idx_list
            st.session_state.section_img_bytes_list = _new_bytes_list

    st.markdown("---")

    # ────────────────────────────────────────
    # (E) 표지·목차·섹션 미리보기
    # ────────────────────────────────────────
    st.markdown("### 🚀 표지·목차·섹션 미리보기")
    st.caption("표지 1장 + 목차 1장 + 섹션 구분 4장 = 총 6장을 이미지로 보여줍니다. (정식 완성본 아님 · 화면 확인용)")

    # ★목차 단계의 '소제목 순서대로 미리보기'처럼 — 체크박스를 켜면 바로 그 자리에 표시.
    _lp_on = st.checkbox("📖 표지·목차·섹션 미리보기", key="layout_preview_on",
                         help="켜면 표지·목차·섹션 6장을 이미지로 바로 보여줍니다. "
                              "설정(표지/목차/섹션 이미지)을 바꾸면 '🔄 갱신'을 누르세요.")
    if _lp_on:
        _need = st.session_state.get("layout_preview_render") is None
        if st.button("🔄 갱신 (다시 생성)", key="layout_preview_refresh"):
            _need = True
        if _need:
            try:
                _toc_cnt, _toc_map = _toc_edit_to_maps()   # 목차 제목/소제목 반영
                with st.spinner("미리보기 PPT를 생성하는 중입니다..."):
                    preview_bytes = build_preview_presentation(
                        business_name=st.session_state.business_name,
                        year=st.session_state.year,
                        month_en=st.session_state.month_en,
                        cover_image_bytes=st.session_state.cover_image_bytes,
                        section_image_bytes_list=st.session_state.section_img_bytes_list,
                        toc_image_bytes_list=[st.session_state.toc_img_bytes],
                        toc_count=_toc_cnt,
                        toc_map=_toc_map,
                    )
                from modules.preview import ppt_to_images
                with st.spinner("슬라이드를 이미지로 변환하는 중... (LibreOffice)"):
                    _imgs, _err = ppt_to_images(preview_bytes)
                st.session_state.layout_preview_render = {"imgs": _imgs, "err": _err}
            except Exception as e:
                st.session_state.layout_preview_render = {
                    "imgs": [], "err": f"미리보기 생성 실패: {e}"}

        _lrend = st.session_state.get("layout_preview_render")
        if _lrend and _lrend.get("err"):
            st.error(f"이미지 변환 실패 — {_lrend['err']}")
        elif _lrend and _lrend.get("imgs"):
            st.success(f"✅ 미리보기 {len(_lrend['imgs'])}장 (표지·목차·섹션)")
            _imgs = _lrend["imgs"]
            # 라벨: 표지 · 목차 · 섹션 01,02,… (섹션 수는 목차 형식에 따라 4~5)
            _labels = ["표지", "목차"] + [f"섹션 {i + 1:02d}" for i in range(len(_imgs) - 2)]
            for _i in range(0, len(_imgs), 2):     # 가로 2장씩
                _row = _imgs[_i:_i + 2]
                _cs = st.columns(2)
                for _j, _png in enumerate(_row):
                    with _cs[_j]:
                        _lab = _labels[_i + _j] if _i + _j < len(_labels) else f"{_i + _j + 1}장"
                        st.image(_png, use_container_width=True)
                        st.caption(_lab)
        else:
            st.warning("표시할 이미지가 없습니다.")

    st.markdown("---")

    # ────────────────────────────────────────
    # (F) 하단 네비게이션
    # ────────────────────────────────────────
    col_prev, col_space, col_next = st.columns([2, 4, 2])

    with col_prev:
        if st.button("← 이전 단계", use_container_width=True):
            st.session_state.current_step = 4   # → 배치 확인(4단계)
            st.rerun()

    with col_next:
        if st.button("다음 단계 →", use_container_width=True, type="primary"):
            st.session_state.current_step = 6   # → 전체 PPT 생성(6단계)
            st.rerun()


# ─────────────────────────────────────────────
# [4단계: 전체 PPT 생성 및 다운로드]
# ─────────────────────────────────────────────
def show_step4():
    data = st.session_state.extracted_data
    if data is None:
        st.warning("추출 데이터가 없습니다. 1단계로 돌아가주세요.")
        if st.button("← 1단계로"):
            st.session_state.current_step = 1
            st.rerun()
        return

    st.markdown("## 6단계. 전체 PPT 생성")
    st.caption("지금까지 설정한 내용으로 완성된 제안서 PPT를 생성합니다.")
    st.markdown("")

    # ────────────────────────────────────────
    # (A) 파싱된 슬라이드 구성 요약
    # ────────────────────────────────────────
    pages = st.session_state.parsed_pages
    st.markdown("### 📋 슬라이드 구성 미리보기")

    if not pages:
        st.warning("⚠️ 문서에서 구조화된 내용을 찾지 못했습니다. 본문 슬라이드 없이 표지·목차·연락처만 생성됩니다.")
    else:
        st.success(f"총 **{len(pages)}개** 내용 슬라이드가 감지되었습니다.")

        # 섹션별 슬라이드 수 표시
        from collections import Counter
        section_counts = Counter(p.get("section_title", "(섹션 없음)") for p in pages)
        for sec, cnt in section_counts.items():
            st.markdown(f"- **{sec or '(섹션 없음)'}** — {cnt}개 슬라이드")

        # 상세 펼치기
        with st.expander("슬라이드 상세 내용 보기"):
            for i, page in enumerate(pages):
                st.markdown(f"**[{i+1}] {page.get('section_title', '')} / {page.get('subtitle', '')}**")
                body = page.get("body_text", "").strip()
                if body:
                    st.caption(body[:200] + ("..." if len(body) > 200 else ""))
                st.markdown("---")

    st.markdown("---")

    # ────────────────────────────────────────
    # (B) 생성 설정 요약
    # ────────────────────────────────────────
    st.markdown("### ⚙️ 생성 설정 확인")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("사업명", st.session_state.business_name or "(미입력)")
    with col2:
        st.metric("날짜", f"{st.session_state.month_en} {st.session_state.year}")
    with col3:
        cover_ok = st.session_state.cover_image_bytes is not None
        st.metric("표지 이미지", "선택됨" if cover_ok else "없음")

    st.markdown("---")

    # ────────────────────────────────────────
    # (C) PPT 생성 버튼
    # ────────────────────────────────────────
    st.markdown("### 🚀 PPT 생성")

    st.caption("소제목마다 지정한 페이지들을 묶어 요약합니다. "
               "핵심 수치(금액·금리·만기·LTV·감정가·상환재원 등)는 그대로 유지됩니다.")

    if st.button("생성", type="primary", use_container_width=False):
        try:
            with st.spinner("PPT를 생성하는 중입니다... 잠시만 기다려주세요."):
                # ── 목차: 4단계에서 짠 사용자 목차를 최우선 사용 ──
                toc_choice = st.session_state.toc_count
                _u_cnt, _u_map = _toc_edit_to_maps()
                if _u_map and _u_map.get("_labels"):
                    final_pages     = pages
                    final_toc_map   = _u_map                 # 사용자 제목·소제목
                    final_toc_count = _u_cnt if _u_cnt in (4, 5) else toc_choice
                elif pages and toc_choice == 5:
                    try:
                        _toc4   = extract_toc_map(pages)
                        _lbl4   = extract_section_labels(pages)
                        _toc5, _lbl5, _split = split_into_5_sections(_toc4, _lbl4)
                        _pages5 = remap_pages_for_5sections(pages, _split, _lbl5)
                        final_pages     = _pages5
                        final_toc_map   = dict(_toc5)
                        final_toc_map["_labels"] = _lbl5
                        final_toc_count = 5
                    except Exception as _e:
                        st.warning(f"⚠️ 5섹션 자동 분할 실패({_e}). 4섹션으로 생성합니다.")
                        final_pages     = pages
                        final_toc_map   = None
                        final_toc_count = 4
                else:
                    final_pages     = pages
                    final_toc_map   = None
                    final_toc_count = toc_choice

                # ── 하이라이트: 2단계에서 직접 쓴 카드 3개 → Executive Summary 섹션 ──
                #   (LLM 자동생성 아님 — 수동 카드를 그대로 넣는다)
                _es_sections = None
                _cards = st.session_state.get("highlight_cards")
                if _cards and any((c.get("title") or c.get("content")) for c in _cards):
                    _es_sections = _cards_to_sections(_cards)

                # ── 본문: 4단계 목차 + 5단계 배치 순서대로 (섹션4개+1.1+2.1 무조건) ──
                _toc = st.session_state.get("toc_edit") or {}
                _toc_groups = _toc.get("groups") or None
                _full_text = "\n".join(
                    (st.session_state.get("extracted_data") or {}).get("pages_text", []))

                # ★목차 기준 생성에선 enrich_and_number(페이지 드롭·재정렬·재번호)를 쓰지 않는다.
                #   그걸 쓰면 네가 소제목에 적은 페이지 번호가 사라져 내용을 못 찾는다.
                #   대신 build_content_from_toc가 '소제목에 지정한 페이지'만 개별 구조화한다.
                os.environ["RAINFIELD_LLM"] = "0"
                if not _toc_groups:
                    # (구) 자동 흐름에서만 enrich 사용
                    os.environ["RAINFIELD_LLM"] = "1"
                    try:
                        from modules.llm_structure import enrich_and_number
                        _pdf_b = st.session_state.get("pdf_bytes")
                        _pdf_path = None
                        if _pdf_b:
                            import tempfile
                            _tf = tempfile.NamedTemporaryFile(suffix=".pdf", delete=False)
                            _tf.write(_pdf_b); _tf.close(); _pdf_path = _tf.name
                        enrich_and_number(final_pages, pdf_path=_pdf_path)
                    except Exception as _llm_e:
                        st.warning(f"⚠️ 본문 LLM 구조화 일부 실패({_llm_e}) — 기본 경로로 진행합니다.")

                # ── 요약본 한 종류만 만든다 ────────────────────
                #   (긴버전·짧은버전을 따로 냈더니 내용이 거의 같아 나눌 이유가 없었음)
                os.environ["RAINFIELD_COMPACT"] = "1"
                ppt_bytes = build_full_presentation(
                    business_name=st.session_state.business_name,
                    year=st.session_state.year,
                    month_en=st.session_state.month_en,
                    pages=final_pages,
                    cover_image_bytes=st.session_state.cover_image_bytes,
                    executive_summary_sections=_es_sections,   # 수동 하이라이트 카드
                    section_image_bytes_list=st.session_state.section_img_bytes_list,
                    toc_count=final_toc_count,
                    toc_map=final_toc_map,
                    toc_image_bytes_list=[st.session_state.toc_img_bytes],
                    exec_summary_data=None,
                    toc_groups=_toc_groups,   # ★네 목차·배치로 본문 생성
                    full_text=_full_text,     # 1.1/2.1 자동추출용 원문
                )
                os.environ.pop("RAINFIELD_COMPACT", None)

            # ★5단계 내용검수에서 쓰도록 생성 PPT 보관
            st.session_state.ppt_bytes = ppt_bytes

            # ★결과를 세션에 저장한다 — 다운로드 버튼을 누르면 Streamlit 이 페이지를
            #   다시 그리므로, 결과가 이 블록 안 지역변수면 그때 사라진다.
            #   버튼은 아래(생성 블록 바깥)에서 세션 값으로 그린다.
            try:
                from pptx import Presentation as _Prs
                _n = len(_Prs(io.BytesIO(ppt_bytes)).slides)
            except Exception:
                _n = 0
            st.session_state.ppt_result = {
                "bytes": ppt_bytes,
                "slides": _n,
                "filename": make_output_filename(st.session_state.business_name),
            }

        except Exception as e:
            st.session_state.pop("ppt_result", None)
            st.error(f"❌ PPT 생성 중 오류가 발생했습니다: {e}")
            import traceback
            st.code(traceback.format_exc(), language="text")

    # ────────────────────────────────────────
    # (C-2) 다운로드 — 생성 블록 '바깥'에서 세션 값으로 그린다.
    #        그래야 한쪽을 받아도 나머지 버튼이 그대로 남는다.
    # ────────────────────────────────────────
    _res = st.session_state.get("ppt_result")
    if _res:
        st.success(f"생성 완료 — {_res['slides']}장" if _res["slides"] else "생성 완료")
        st.download_button(
            label="다운로드",
            data=_res["bytes"],
            file_name=_res["filename"],
            mime="application/vnd.openxmlformats-officedocument."
                 "presentationml.presentation",
            type="primary",
            key="dl_ppt",
        )

    st.markdown("---")

    # ────────────────────────────────────────
    # (D) 하단 네비게이션
    # ────────────────────────────────────────
    if st.button("← 이전 단계 (설정 변경)", use_container_width=False):
        st.session_state.current_step = 5   # → 레이아웃/표지(5단계)
        st.rerun()


# ─────────────────────────────────────────────
# [3단계: 목차 구성] — 자동 추출 목차 편집 + 원본 항목(해시태그) 매치. (읽기/편집만, 이미지·재배치 X)
# ─────────────────────────────────────────────
def _clean_toc_text(t: str) -> str:
    return re.sub(r"\s+", " ", (t or "").strip())


def _is_toc_heading(t):
    """목차 제목/소제목다운 '짧은 문구'인지 — 긴 본문 문장/서술형은 제외."""
    t = (t or "").strip()
    if not (0 < len(t) <= 45):
        return False
    if re.search(r"(습니다|입니다|합니다|이다|였다|된다|한다|하였다)\.?$", t):
        return False
    return True


def _parse_pages(raw, total=0):
    """적은 '순서 그대로' 페이지 리스트로. 정렬하지 않는다.
       '8,9,11,10,12,16-19' → [8,9,11,10,12,16,17,18,19].
       범위 'a-b'/'a~b'는 방향대로 펼침(19-16 → 19,18,17,16).
       중복(뒤에 또 나온 페이지)·범위밖은 제거. 빈 값이면 []."""
    out = []

    def _add(n):
        if n >= 1 and (total <= 0 or n <= total) and n not in out:
            out.append(n)

    for tok in re.split(r"[,\s]+", (raw or "").strip()):
        if not tok:
            continue
        m = re.match(r"^(\d+)\s*[-~]\s*(\d+)$", tok)   # 범위: 8-12 / 8~12 / 12-8
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            step = 1 if a <= b else -1
            for n in range(a, b + step, step):       # 방향(오름/내림) 유지
                _add(n)
            continue
        try:
            _add(int(tok))
        except ValueError:
            continue
    return out                                        # ★정렬 안 함 — 순서 보존


def _pages_to_str(pages):
    """[4,5] → '4,5' (입력칸 기본값용)."""
    return ",".join(str(p) for p in (pages or []))


def _build_toc_from_parsed(parsed_pages):
    """parsed_pages로 목차 자동 추출 시도(best-effort) → groups.
       ※IM마다 섹션/소제목 형식이 제각각이고 러닝헤더(사업명)가 섞여 신뢰도 낮음 → '시도'용.
       여러 페이지에 반복되는 제목(러닝헤더)은 빈도로 걸러낸다."""
    from collections import OrderedDict, Counter
    freq = Counter(_clean_toc_text(pd.get("section_title") or pd.get("section_label"))
                   for pd in (parsed_pages or []))
    npages = max(1, len(parsed_pages or []))

    def _bad(t):
        return (not _is_toc_heading(t)) or freq.get(t, 0) > max(3, npages * 0.4)

    sections = OrderedDict()
    for pd in (parsed_pages or []):
        sec = _clean_toc_text(pd.get("section_title") or pd.get("section_label"))
        sub = _clean_toc_text(pd.get("subtitle"))
        if _bad(sec):
            sec = ""
        sub = sub if _is_toc_heading(sub) else ""
        if sec:
            sections.setdefault(sec, [])
            if sub and sub != sec and sub not in sections[sec]:
                sections[sec].append(sub)
    return [{"title": k, "subs": [{"text": s, "pages": [], "fixed": False} for s in v]}
            for k, v in sections.items()]


def _adjust_groups(groups, n):
    """대분류 개수를 정확히 n개로(부족하면 빈 그룹 추가, 많으면 뒤에서 제거)."""
    while len(groups) < n:
        groups.append({"title": "", "subs": [{"text": "", "pages": [], "fixed": False}]})
    while len(groups) > n:
        groups.pop()
    return groups


def _clear_toc_widget_state():
    """목차 편집 위젯 상태 초기화(구조 변경 후 재시딩용)."""
    for k in list(st.session_state.keys()):
        if k.startswith(("toctitle_", "tocsub_", "tocpage_", "tocfix_")):
            del st.session_state[k]


# 표준 목차(거의 고정) — 화면에서 수정 가능.
#   ★페이지는 '소제목(sub)' 단위로 넣는다: sub["pages"] (순서·겹침 가능).
#     → 어떤 페이지가 어느 소제목 내용인지 명확 → 소제목별 내용 슬라이드가 정확히 생성.
#   fixed=True 소제목(1.1 사모사채 개요 / 2.1 금융 구조도)은 원본 없이 자동 생성(페이지 불필요).
_DEFAULT_TOC = [
    {"title": "사모사채 개요",
     "subs": [{"text": "본건 사모사채 개요", "pages": [], "fixed": True}]},
    {"title": "금융개요",
     "subs": [{"text": "금융 구조도", "pages": [], "fixed": True},
              {"text": "본건 기초자산 금융조건", "pages": [], "fixed": False}]},
    {"title": "사업개요",
     "subs": [{"text": "사업개요", "pages": [], "fixed": False}]},
    {"title": "Appendix",
     "subs": [{"text": "", "pages": [], "fixed": False}]},
]


def _init_toc_edit(parsed, auto=False):
    """목차 편집 상태 초기화 — 기본은 표준 목차(수정 가능). auto=True면 원본에서 자동 추출 시도."""
    import copy
    groups = _build_toc_from_parsed(parsed) if auto else copy.deepcopy(_DEFAULT_TOC)
    fmt = max(4, min(5, len(groups) or 4))
    _adjust_groups(groups, fmt)
    st.session_state.toc_edit = {"format": fmt, "groups": groups}
    st.session_state.toc_count = fmt


def _build_keywords_from_parsed(parsed):
    """★키워드 방식: 파싱된 각 내용 페이지에서 '진짜 내용 제목(키워드)+그 페이지 번호'를 뽑는다.
       - 매 페이지 반복되는 머리말/꼬리말(사업명·Strictly Private·페이지번호)을 자동 감지해 걸러내고,
         그 페이지의 '첫 진짜 제목' 줄을 키워드로 쓴다(광역입지·선임차인 등).
       - 페이지 매핑은 parsed_pages의 실제 page_num이라 100% 정확(LLM 추측 아님).
       반환: [{"label": "표시용(항목 · Np)", "base": "항목", "pages": [..]}]"""
    from collections import Counter
    rows = [(p.get("page_num"),
             p.get("raw_text") or p.get("subtitle") or p.get("section_title") or "")
            for p in (parsed or []) if p.get("page_num") is not None]
    if not rows:
        return []

    # 1) 여러 페이지에 반복되는 줄 = 머리말/꼬리말 → 제외 집합
    freq = Counter()
    for _, txt in rows:
        for s in {re.sub(r"\s+", " ", ln.strip()) for ln in txt.splitlines() if ln.strip()}:
            freq[s] += 1
    npg = len(rows)
    common = {s for s, c in freq.items() if c >= max(2, int(npg * 0.3))}

    def _title_of(txt):
        for ln in txt.splitlines():
            s = re.sub(r"\s+", " ", ln.strip())
            if not s or s in common or len(s) < 2:
                continue
            low = s.lower()
            if "strictly" in low or "confidential" in low:
                continue
            if re.fullmatch(r"[\d\s|·.\-]+", s):          # 페이지번호/구분선만
                continue
            return s
        return ""

    kws = []
    for pn, txt in rows:
        base = _title_of(txt)
        if len(base) > 26:
            cut = base[:26]
            sp = cut.rfind(" ")
            base = (cut[:sp] if sp > 8 else cut).strip() + "…"
        if not base:
            base = f"{pn}페이지"
        if kws and kws[-1]["base"] == base:
            kws[-1]["pages"].append(pn)
        else:
            kws.append({"base": base, "pages": [pn]})
    for k in kws:
        k["label"] = f"{k['base']}  ·  {_pages_to_str(k['pages'])}p"
    return kws


def _init_toc_edit_from_llm():
    """IM 원문 전체를 LLM으로 읽어 목차(대분류+소제목) 구성 → toc_edit.
       반환: (성공?, 안내메시지). 실패해도 기존 목차는 유지."""
    pages_text = (st.session_state.get("extracted_data") or {}).get("pages_text", [])
    full_text = "\n".join(pages_text)
    if not full_text.strip():
        return False, "추출된 텍스트가 없어 목차를 자동 추출할 수 없습니다. 1단계에서 워드/PDF를 올려주세요."
    try:
        from modules.ai_slide_builders import generate_toc
        res = generate_toc(full_text)
    except Exception as e:
        return False, f"목차 자동 추출 실패: {e}"
    if not res.get("ok"):
        return False, f"목차 자동 추출 실패: {res.get('error', '알 수 없음')}"
    secs = (res.get("data") or {}).get("sections") or []
    if not secs:
        return False, "IM에서 목차를 추출하지 못했습니다. 기본 목차에서 직접 수정해주세요."
    groups = []
    for si, sec in enumerate(secs[:5]):
        title = (sec.get("title") or "").strip()
        subs_raw = [str(x).strip() for x in (sec.get("subtitles") or []) if str(x).strip()]
        # 고정(자동생성) 소제목: 섹션1의 1.1(본건 사모사채 개요), 섹션2의 2.1(금융 구조도)
        subs = [{"text": t, "pages": [],
                 "fixed": ((si == 0 and j == 0) or (si == 1 and j == 0))}
                for j, t in enumerate(subs_raw)]
        if not subs:
            subs = [{"text": "", "pages": [], "fixed": (si in (0, 1))}]
        groups.append({"title": title, "subs": subs})
    fmt = max(4, min(5, len(groups) or 4))
    _adjust_groups(groups, fmt)
    st.session_state.toc_edit = {"format": fmt, "groups": groups}
    st.session_state.toc_count = fmt
    return True, "IM 전체를 읽어 목차를 구성했어요. 확인하고 필요하면 수정하세요."


def _pdf_page_count(pdf_bytes):
    """PDF 페이지 수(실패 시 -1) — 사진 PDF와 글자 PDF 정렬 확인용."""
    try:
        import fitz
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            return doc.page_count
        finally:
            doc.close()
    except Exception:
        return -1


def _render_pdf_page(pdf_bytes, page_num, zoom=2.0):
    """PDF 한 페이지를 PNG bytes로 렌더 — PyMuPDF(fitz), 자금판(전) 방식.
       외부 프로그램(LibreOffice/poppler) 없이 빠르게 렌더. 반환: (png|None, err|None)."""
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        try:
            pix = doc[page_num - 1].get_pixmap(matrix=fitz.Matrix(zoom, zoom))
            return pix.tobytes("png"), None
        finally:
            doc.close()
    except Exception as e:
        return None, f"이미지 렌더 실패: {e}"


def _toc_page_png(view_pdf, page):
    """미리보기용 페이지 PNG(렌더 결과를 세션에 캐시 → 재실행 때 재렌더 안 함).
       실패하면 None. 새 문서 업로드 시 toc_png_cache는 초기화된다."""
    cache = st.session_state.setdefault("toc_png_cache", {})
    if page not in cache:
        png, _err = _render_pdf_page(view_pdf, page)
        cache[page] = png
    return cache.get(page)


def _render_pages_grid(view_pdf, pages, cols=5):
    """페이지 목록을 가로 최대 cols개(기본 5) 그리드로 렌더 — 화면폭에 맞춰 축소.
       칸은 항상 cols개 만들어 이미지 크기를 일정하게 유지한다."""
    for i in range(0, len(pages), cols):
        row = pages[i:i + cols]
        cells = st.columns(cols)
        for j, p in enumerate(row):
            with cells[j]:
                png = _toc_page_png(view_pdf, p)
                if png:
                    st.image(png, use_container_width=True)
                    st.caption(f"{i + j + 1}번째 · {p}p")
                else:
                    st.caption(f"⚠️ {p}p 렌더 실패")


def show_step_toc():
    st.markdown("## 3단계. 목차 구성")
    st.caption("왼쪽에서 원본 IM을 표지부터 넘겨보고, 오른쪽에서 목차 제목·소제목을 직접 만드세요. "
               "소제목마다 **원본 IM의 몇 페이지인지** 번호로 지정하면 됩니다. "
               "(IM마다 목차/소제목 형식이 제각각이라 자동추출은 부정확 → 페이지 번호로 매칭)")

    parsed = st.session_state.get("parsed_pages") or []
    pages_text = (st.session_state.get("extracted_data") or {}).get("pages_text", [])
    total = len(pages_text)
    pdf_bytes = st.session_state.get("pdf_bytes")           # 글자/항목 소스(워드 변환)
    view_pdf = st.session_state.get("view_pdf_bytes") or pdf_bytes   # 사진(원본 PDF 우선)

    # ★파싱 결과가 비어있으면(세션 초기화 등) PDF로 즉석 재파싱 → 키워드·생성 복구
    if not parsed and pdf_bytes:
        try:
            with st.spinner("원본을 다시 읽어 페이지·키워드를 준비하는 중..."):
                parsed = parse_document_from_bytes(pdf_bytes, "source.pdf")
            st.session_state.parsed_pages = parsed
        except Exception as _pe:
            st.warning(f"파싱 재시도 실패: {_pe} — 1단계에서 PDF를 다시 올려주세요.")
            parsed = []
    if "toc_edit" not in st.session_state:
        _init_toc_edit(parsed)
    toc = st.session_state.toc_edit

    bc1, bc2, _bc = st.columns([1.4, 1, 2.6])
    with bc1:
        if st.button("🤖 IM 전체 읽어 목차 자동 추출", use_container_width=True,
                     help="IM 원문 전체를 LLM으로 읽어 대분류·소제목을 구성합니다(그 뒤 직접 수정)."):
            with st.spinner("IM 전체를 읽어 목차를 구성하는 중..."):
                _ok, _msg = _init_toc_edit_from_llm()
            _clear_toc_widget_state()
            st.session_state["_toc_autofill_msg"] = ("success" if _ok else "warning", _msg)
            st.rerun()
    with bc2:
        if st.button("🧹 기본 목차로 되돌리기", use_container_width=True,
                     help="표준 목차(사모사채 개요/금융개요/사업개요/Appendix)로 초기화합니다."):
            _init_toc_edit(parsed, auto=False)
            _clear_toc_widget_state()
            st.session_state.pop("_toc_autofill_msg", None)
            st.rerun()
    _tmsg = st.session_state.get("_toc_autofill_msg")
    if _tmsg:
        (st.success if _tmsg[0] == "success" else st.warning)(_tmsg[1])

    # ★왼쪽 PDF 칸만 화면 상단에 고정(sticky) — 오른쪽 목차를 스크롤해도 페이지가 안 사라지게.
    #   :has()로 '왼쪽 칸'만 지목 → 소제목 줄 컬럼엔 영향 없음. 미지원 브라우저는 자동으로 원래대로.
    st.markdown(
        "<style>"
        "div[data-testid='stColumn']:has(#toc-left-anchor),"
        "div[data-testid='column']:has(#toc-left-anchor){"
        "position:sticky; top:3.2rem; align-self:flex-start;"
        "max-height:calc(100vh - 4rem); overflow-y:auto; z-index:5;}"
        "</style>", unsafe_allow_html=True)

    img_col, form_col = st.columns([1, 1])

    # ── 왼쪽: 원본 IM 전체 페이지 뷰어 (표지부터 한 장씩, 버튼 없이 바로 표시) ──
    with img_col:
        st.markdown("<div id='toc-left-anchor'></div>", unsafe_allow_html=True)
        st.markdown("#### 📄 원본 IM (표지부터 한 장씩)")
        if not pdf_bytes or total == 0:
            st.warning("원본 PDF가 없어 이미지를 못 띄웁니다. 1단계에서 PDF(또는 워드→자동 PDF변환)로 올리세요.")
        else:
            cur = min(max(1, st.session_state.get("toc_view_page", 1)), total)
            pc1, pc2, pc3 = st.columns([1, 2, 1])
            with pc1:
                if st.button("◀ 이전", use_container_width=True, disabled=(cur <= 1)):
                    st.session_state.toc_view_page = cur - 1
                    st.rerun()
            with pc2:
                st.markdown(f"<div style='text-align:center; padding-top:6px;'>"
                            f"<b>{cur}</b> / {total} 페이지</div>", unsafe_allow_html=True)
            with pc3:
                if st.button("다음 ▶", use_container_width=True, disabled=(cur >= total)):
                    st.session_state.toc_view_page = cur + 1
                    st.rerun()
            # 현재 페이지를 fitz로 즉시 렌더(캐시) — 생성 버튼·다운로드 없이 바로 표시
            #   ★어떤 오류가 나도 오른쪽 폼·아래 버튼이 사라지지 않게 전부 감싼다.
            try:
                cache = st.session_state.setdefault("toc_page_cache", {})
                if cur not in cache or not isinstance(cache.get(cur), tuple):
                    cache[cur] = _render_pdf_page(view_pdf, cur)
                png, err = cache[cur]
            except Exception as _e:
                png, err = None, f"이미지 표시 오류: {_e}"
            if err:
                st.error(err)
            elif png:
                # 페이지 전체를 한눈에 보이게 — 잘리지 않도록 높이 제한 없이 그대로 표시
                st.image(png, use_container_width=True)

    # ── 오른쪽: 목차 편집 폼 (컬럼 2중첩 방지 → 세로로 쌓음) ──
    with form_col:
        st.markdown("#### 📝 목차 구성 편집")
        fmt_sel = st.radio("목차 형식 (대분류 개수)", [4, 5],
                           index=(0 if toc["format"] == 4 else 1),
                           horizontal=True, format_func=lambda x: f"{x}형식",
                           key="toc_fmt_radio")
        if fmt_sel != toc["format"]:
            toc["format"] = fmt_sel
            _adjust_groups(toc["groups"], fmt_sel)
            st.session_state.toc_count = fmt_sel
            _clear_toc_widget_state()
            st.rerun()

        _cur = min(max(1, st.session_state.get("toc_view_page", 1)), total or 1)
        # ★키워드 방식: 파싱된 페이지에서 '키워드(항목)+페이지'를 뽑아 소제목에서 골라 연결.
        _kw_list = _build_keywords_from_parsed(parsed)
        _kw_labels = [k["label"] for k in _kw_list]
        _kw_pages = {k["label"]: k["pages"] for k in _kw_list}
        if _kw_labels:
            st.caption(f"🔖 **키워드 {len(_kw_labels)}개** 자동 추출됨 — 각 소제목에서 "
                       f"**키워드를 골라 연결**하면 페이지가 자동으로 채워집니다.")
        else:
            st.caption(f"소제목마다 원본 페이지를 넣으세요. 지금 왼쪽 페이지: **{_cur}p** "
                       f"(범위 8-12 · 재배치 8,9,11,10 · 겹침 가능)")
        # ★기본은 키워드만(깔끔). 번호로 직접 넣고 싶을 때만 켜기 → '페이지 직접' 칸 표시
        _manual_mode = st.checkbox("✏️ 페이지 번호로 직접 입력 (키워드 대신)",
                                   key="toc_manual_mode",
                                   help="켜면 소제목마다 페이지 직접 입력칸이 생깁니다. "
                                        "평소엔 꺼두고 키워드만 골라 쓰면 깔끔해요.")

        pending = {"del": None, "add": None}
        for gi, g in enumerate(toc["groups"]):
            with st.container(border=True):
                tk = f"toctitle_{gi}"
                st.session_state.setdefault(tk, g.get("title", ""))
                g["title"] = st.text_input(f"목차 {gi + 1} 제목", key=tk,
                                           placeholder="예: 4 리스크분석 / Executive Summary")

                # ★소제목 1개 = 한 줄([번호|소제목|페이지|고정|삭제]) — 세로 공간 절약(스크롤 최소화)
                for si, sub in enumerate(g["subs"]):
                    is_fixed = bool(sub.get("fixed", False))
                    sc0, sc1, sc2, sc3, sc4 = st.columns([0.6, 3.5, 4.3, 1.2, 0.7])
                    with sc0:
                        st.markdown(f"<div style='padding-top:8px;color:#888;'>"
                                    f"{gi + 1}.{si + 1}</div>", unsafe_allow_html=True)
                    with sc1:
                        sk = f"tocsub_{gi}_{si}"
                        st.session_state.setdefault(sk, sub.get("text", ""))
                        sub["text"] = st.text_input(
                            f"소제목 {gi + 1}.{si + 1}", key=sk,
                            placeholder="소제목", label_visibility="collapsed")
                    chosen = []
                    with sc2:
                        # ── ★키워드 선택(고르면 페이지 자동). 고정 소제목은 불필요 ──
                        if is_fixed:
                            sub["pages"] = []
                            rok = f"tocpgro_{gi}_{si}"
                            st.session_state[rok] = "🔒 자동생성(원본 불필요)"
                            st.text_input("키워드", key=rok, disabled=True,
                                          label_visibility="collapsed")
                        else:
                            kwk = f"tockw_{gi}_{si}"
                            # 현재 키워드 목록에 없는 이전 선택은 제거(옵션 불일치 크래시 방지)
                            if kwk in st.session_state:
                                st.session_state[kwk] = [x for x in st.session_state[kwk]
                                                         if x in _kw_labels]
                            else:
                                st.session_state[kwk] = [x for x in (sub.get("_keywords") or [])
                                                         if x in _kw_labels]
                            chosen = st.multiselect(
                                "키워드", options=_kw_labels, key=kwk,
                                placeholder="키워드 선택… (여러 개 가능)",
                                label_visibility="collapsed")
                            sub["_keywords"] = chosen
                    with sc3:
                        fk = f"tocfix_{gi}_{si}"
                        st.session_state.setdefault(fk, is_fixed)
                        sub["fixed"] = st.checkbox("🔒고정", key=fk,
                            help="1.1 사모사채 개요·2.1 금융 구조도처럼 원본 없이 자동 생성(페이지 불필요)")
                    with sc4:
                        if st.button("🗑", key=f"tocdel_{gi}_{si}", help="이 소제목 삭제"):
                            pending["del"] = (gi, si)
                    # ── 페이지 결정 ──
                    #   키워드 선택 → 그 페이지 자동. 아니면 '직접 입력 모드'일 때만 입력칸(깔끔).
                    if not is_fixed:
                        pk = f"tocpage_{gi}_{si}"
                        if chosen:
                            _pgs = []
                            for _lab in chosen:
                                for _p in _kw_pages.get(_lab, []):
                                    if _p not in _pgs:
                                        _pgs.append(_p)
                            sub["pages"] = _pgs
                            st.session_state[pk] = _pages_to_str(_pgs)   # 직접칸과 동기화
                            st.caption(f"　└ 📄 {_pages_to_str(_pgs)}p · 키워드 "
                                       f"{len(chosen)}개 연결(자동)")
                        elif _manual_mode:
                            st.session_state.setdefault(pk, _pages_to_str(sub.get("pages")))
                            _p_raw = st.text_input(
                                f"페이지 직접 {gi + 1}.{si + 1}", key=pk,
                                placeholder="페이지 직접: 8,9,11",
                                label_visibility="collapsed")
                            sub["pages"] = _parse_pages(_p_raw, total)
                        else:
                            # 깔끔: 입력칸 없음. 이미 정해진 페이지만 작게 표시.
                            _prev = st.session_state.get(pk, _pages_to_str(sub.get("pages")))
                            sub["pages"] = _parse_pages(_prev, total)
                            if sub["pages"]:
                                st.caption(f"　└ 📄 {_pages_to_str(sub['pages'])}p")
                if st.button("➕ 소제목 추가", key=f"tocadd_{gi}"):
                    pending["add"] = gi

        if pending["del"] is not None:
            gi, si = pending["del"]
            toc["groups"][gi]["subs"].pop(si)
            _clear_toc_widget_state()
            st.rerun()
        if pending["add"] is not None:
            toc["groups"][pending["add"]]["subs"].append({"text": "", "pages": [], "fixed": False})
            _clear_toc_widget_state()
            st.rerun()

    # (소제목 순서대로 미리보기는 제거 — 다음 '배치 확인' 단계가 그 미리보기 역할을 함)

    # ── 저장 요약 + 네비게이션 (전체 폭) ──
    st.markdown("---")
    total_subs = sum(len(g["subs"]) for g in toc["groups"])
    with_page = sum(1 for g in toc["groups"] for s in g["subs"] if s.get("pages"))
    fixed_cnt = sum(1 for g in toc["groups"] for s in g["subs"] if s.get("fixed"))
    st.success(f"저장됨 · 대분류 {toc['format']}개 · 소제목 {total_subs}개 · "
               f"페이지 넣은 소제목 {with_page}개 · 고정 {fixed_cnt}건")
    nc1, _ns, nc2 = st.columns([2, 4, 2])
    with nc1:
        if st.button("← 이전 단계", use_container_width=True):
            st.session_state.current_step = 2   # → 하이라이트(2단계)
            st.rerun()
    with nc2:
        if st.button("다음 단계 →", use_container_width=True, type="primary"):
            st.session_state.current_step = 4   # → 배치 확인(4단계)
            st.rerun()


def _toc_edit_to_maps():
    """4단계 목차(toc_edit) → (toc_count, toc_map).
       toc_map = {"01":[소제목...], ..., "_labels":{"01":제목,...}}
       → 미리보기/최종 생성의 목차 슬라이드·섹션 divider가 같은 제목을 쓰게 한다.
       toc_edit이 없으면 (None, None)."""
    toc = st.session_state.get("toc_edit")
    if not toc:
        return None, None
    count = toc.get("format", 4)
    labels, tmap = {}, {}
    for gi, g in enumerate(toc.get("groups", [])):
        num = f"0{gi + 1}"
        title = (g.get("title") or "").strip()
        if title:
            labels[num] = title
        raw_subs = [(s.get("text") or "").strip() for s in g.get("subs", [])]
        raw_subs = [s for s in raw_subs if s]
        if raw_subs:
            # 자동 번호: 목차1 밑 → 1.1 / 목차2 밑 → 2.1, 2.2, 2.3 …
            tmap[num] = [f"{gi + 1}.{j + 1} {t}" for j, t in enumerate(raw_subs)]
    if labels:
        tmap["_labels"] = labels
    return count, tmap


def _find_toc_pages(pages_text):
    """원본 텍스트에서 '목차' 페이지(1-based) 추정 — 상단에 '목차/CONTENTS' 표기가 있는 페이지."""
    hits = []
    for i, t in enumerate(pages_text or [], start=1):
        head = (t or "")[:300]
        low = head.lower()
        if "목차" in head or "contents" in low or "table of contents" in low:
            hits.append(i)
    return hits


# ─────────────────────────────────────────────
# [3단계: 배치 확인] — 목차(대분류)에 넣은 원본 페이지를 순서대로 이미지로 확인하고,
#   여기서 바로 페이지를 수정하면 아래 이미지가 다시 그려진다(생성·다운로드 아님, 화면 확인용).
#   4단계와 같은 위젯 키(tocpage_{gi}_{si})를 써서 수정이 서로 동기화된다.
# ─────────────────────────────────────────────
def show_step_arrange():
    st.markdown("## 4단계. 배치 확인")
    st.caption("목차(대분류)에 넣은 원본 페이지를 목차 순서대로 이미지로 확인합니다. "
               "잘못됐거나 바꾸고 싶으면 각 목차의 페이지 칸에서 바로 고치세요 — "
               "고치면 아래 이미지가 다시 그려집니다. (다운로드 아님 · 화면 확인용)")

    toc = st.session_state.get("toc_edit")
    if not toc:
        st.warning("먼저 3단계에서 목차를 구성해주세요.")
        if st.button("← 3단계로"):
            st.session_state.current_step = 3   # → 목차 구성(3단계)
            st.rerun()
        return

    pages_text = (st.session_state.get("extracted_data") or {}).get("pages_text", [])
    total = len(pages_text)
    view = st.session_state.get("view_pdf_bytes") or st.session_state.get("pdf_bytes")

    # ── 상단: 소제목 → 배치 페이지 매핑 요약 ──
    st.markdown("#### 🗂️ 현재 배치 요약 (어떤 소제목에 몇 페이지를 넣었는지)")
    for gi, g in enumerate(toc["groups"]):
        gtitle = (g.get("title") or "").strip() or f"목차 {gi + 1}"
        st.markdown(f"**목차 {gi + 1}. {gtitle}**")
        for si, sub in enumerate(g["subs"]):
            stext = (sub.get("text") or "").strip() or "(소제목 없음)"
            lab = f"{gi + 1}.{si + 1} {stext}"
            if sub.get("fixed"):
                st.caption(f"　└ {lab} — 🔒 자동 생성(원본 없음)")
            elif sub.get("pages"):
                st.caption(f"　└ {lab} — 📄 {_pages_to_str(sub['pages'])}p ({len(sub['pages'])}장)")
            else:
                st.caption(f"　└ {lab} — ⚪ 페이지 미지정")

    st.markdown("---")

    if not view:
        st.warning("원본 PDF가 없어 이미지를 못 띄웁니다. 1단계에서 PDF를 올리세요.")
    else:
        # ── 소제목별: 페이지 수정칸 + 배치 순서대로 이미지 ──
        with st.spinner("배치한 원본 페이지를 순서대로 렌더링하는 중..."):
            for gi, g in enumerate(toc["groups"]):
                gtitle = (g.get("title") or "").strip() or f"목차 {gi + 1}"
                st.markdown(f"### {gi + 1}. {gtitle}")
                for si, sub in enumerate(g["subs"]):
                    stext = (sub.get("text") or "").strip() or "(소제목 없음)"
                    label = f"{gi + 1}.{si + 1} {stext}"
                    if sub.get("fixed"):
                        st.markdown(f"**{label}**　🔒 자동 생성(원본 없음)")
                        continue
                    st.markdown(f"**{label}**")
                    # 인라인 수정칸 — 4단계와 같은 key로 동기화(고치면 rerun되어 이미지 갱신)
                    pk = f"tocpage_{gi}_{si}"
                    st.session_state.setdefault(pk, _pages_to_str(sub.get("pages")))
                    _raw = st.text_input(
                        "원본 페이지 (순서대로! 예: 8,9,11,10,16-19)",
                        key=pk, placeholder="예: 8  또는  8,9,11,10",
                        label_visibility="collapsed")
                    sub["pages"] = _parse_pages(_raw, total)
                    pages = sub["pages"]
                    if not pages:
                        st.caption("⚪ 페이지 미지정")
                        continue
                    st.caption(f"이 순서로 → {_pages_to_str(pages)} ({len(pages)}장)")
                    _render_pages_grid(view, pages)
                st.markdown("")

    # ── 네비게이션 ──
    st.markdown("---")
    st.info("여기서 확인·수정한 배치 순서가 마지막 단계에서 그대로 PPT 본문 순서가 됩니다.")
    nc1, _ns, nc2 = st.columns([2, 4, 2])
    with nc1:
        if st.button("← 이전 단계", use_container_width=True):
            st.session_state.current_step = 3   # → 목차 구성(3단계)
            st.rerun()
    with nc2:
        if st.button("다음 단계 →", use_container_width=True, type="primary"):
            st.session_state.current_step = 5   # → 레이아웃/표지(5단계)
            st.rerun()


# ─────────────────────────────────────────────
# [4단계: 하이라이트(Executive Summary) 구성] — 원본 ES 페이지 확인 + 카드 3개 편집
# ─────────────────────────────────────────────
def _find_es_pages(pages_text):
    """원본에서 'Executive Summary' 표기가 있는 페이지(1-based)를 추정 — 보통 시작 1페이지.
       ★여러 페이지로 이어지는 ES는 자동 경계 감지가 러닝헤더 때문에 불안정하므로,
         여기선 '표기가 있는 페이지'만 반환하고 실제 범위는 사용자가 화면에서 조정한다."""
    hits = []
    for i, t in enumerate(pages_text or [], start=1):
        head = (t or "")[:400]
        if "executive summary" in head.lower() or \
           re.search(r"(^|\n)\s*(투자|사업|핵심)?\s*요약\s*(\n|$)", head):
            hits.append(i)
    return hits


def _init_highlight_cards():
    """카드 3개 빈 값 — 하이라이트 내용은 사용자가 '직접' 입력한다(자동 추출 안 함)."""
    st.session_state.highlight_cards = [
        {"title": "", "use_sub": False, "subtitle": "", "content": ""} for _ in range(3)
    ]


def _es_only_text(pages_text):
    """오직 'Executive Summary' 페이지 텍스트만 이어붙여 반환.
       ★사용자가 화면에서 지정한 ES 페이지 범위(hl_es_pages_str)를 우선 사용,
         없으면 자동 감지. 전체 문서 요약이 아님. 페이지가 없으면 '' 반환."""
    _raw = (st.session_state.get("hl_es_pages_str") or "").strip()
    es = _parse_pages(_raw, len(pages_text)) if _raw else _find_es_pages(pages_text)
    if not es:
        return ""
    return "\n".join((pages_text[p - 1] or "") for p in es
                     if 1 <= p <= len(pages_text))


def _autofill_highlight_from_es():
    """'Executive Summary' 항목 내용만 LLM으로 3개 카드(각 3~5줄)로 정리해 채운다.
       반환: (성공?, 안내메시지). ES를 못 찾으면 (False, 안내)."""
    pages_text = (st.session_state.get("extracted_data") or {}).get("pages_text", [])
    es_text = _es_only_text(pages_text)
    if not es_text.strip():
        return False, ("이 IM에서 'Executive Summary' 항목을 못 찾았어요. "
                       "(간혹 없는 딜이 있습니다) 카드를 직접 입력해주세요.")
    try:
        from modules.ai_slide_builders import generate_executive_summary
        res = generate_executive_summary(es_text)
    except Exception as e:
        return False, f"자동 추출 실패: {e}"
    if not res.get("ok"):
        return False, f"자동 추출 실패: {res.get('error', '알 수 없음')}"
    secs = (res.get("data") or {}).get("sections") or []
    cards = [{"title": "", "use_sub": False, "subtitle": "", "content": ""} for _ in range(3)]
    for i, sec in enumerate(secs[:3]):
        sub = (sec.get("subtitle") or "").strip()
        cards[i] = {
            "title":    (sec.get("title") or "").strip(),
            "use_sub":  bool(sub),
            "subtitle": sub,
            "content":  (sec.get("content") or sec.get("body") or "").strip(),
        }
    st.session_state.highlight_cards = cards
    _used = (st.session_state.get("hl_es_pages_str") or "").strip() \
        or _pages_to_str(_find_es_pages(pages_text))
    return True, f"'Executive Summary' 항목({_used}p)을 3개 카드로 정리했어요. 확인·수정하세요."


def _clear_highlight_widget_state():
    for k in list(st.session_state.keys()):
        if k.startswith(("hlt_", "hls_", "hlc_", "hlu_")):
            del st.session_state[k]


def show_step_highlight():
    st.markdown("## 2단계. 하이라이트 (Executive Summary) 구성")
    st.caption("제안서 표지 다음에 들어가는 '핵심 요약(Executive Summary)' 3줄 카드를 만드는 단계입니다.")
    with st.expander("📖 사용법 (처음이면 눌러보세요)", expanded=False):
        st.markdown(
            "**순서**\n"
            "1. **왼쪽에서 Executive Summary 페이지를 확인**하세요. 자동으로 찾아서 보여줍니다.\n"
            "2. ES가 **여러 페이지**면(가끔 5장 이상) 왼쪽 위 칸에 **`2-5`처럼 범위를 고치세요**. "
            "한 페이지면 그대로 두면 됩니다.\n"
            "3. **`🤖 Executive Summary 자동 추출`** 버튼을 누르면, **왼쪽에 보이는 그 페이지의 "
            "Executive Summary 내용만** 읽어 오른쪽 카드 3개로 정리해 줍니다.\n"
            "4. 오른쪽 카드를 **직접 수정**하세요(제목·내용). 이게 최종 하이라이트가 됩니다.\n\n"
            "**중요**\n"
            "- **왼쪽에 보이는 페이지 = 자동 추출이 읽는 페이지**입니다. 둘은 항상 같습니다.\n"
            "- **페이지 범위를 먼저 맞춘 뒤** 자동 추출 버튼을 누르세요. "
            "이미 추출한 뒤 범위를 바꿨다면 **버튼을 다시** 누르면 새 범위로 다시 정리합니다.\n"
            "- 한 페이지에 ES와 다른 내용이 섞여 있어도, **Executive Summary 부분만** 골라 정리합니다.")

    pages_text = (st.session_state.get("extracted_data") or {}).get("pages_text", [])
    view_pdf = st.session_state.get("view_pdf_bytes") or st.session_state.get("pdf_bytes")
    total = len(pages_text)
    if "highlight_cards" not in st.session_state:
        _init_highlight_cards()
    cards = st.session_state.highlight_cards

    _hb1, _hb2, _hb3 = st.columns([2, 2, 3])
    with _hb1:
        if st.button("🤖 Executive Summary 자동 추출", use_container_width=True,
                     help="IM의 'Executive Summary' 항목만 읽어 3개 카드로 정리합니다(그 뒤 직접 수정)."):
            with st.spinner("'Executive Summary' 항목을 읽어 정리하는 중..."):
                _ok, _msg = _autofill_highlight_from_es()
            # ★사용자가 고른 ES 페이지 유지(rerun에 위젯값이 자동감지값으로 리셋되던 문제 방지)
            st.session_state["_es_pages_saved"] = (
                st.session_state.get("hl_es_pages_str") or "").strip()
            _clear_highlight_widget_state()
            st.session_state["_hl_autofill_msg"] = ("success" if _ok else "warning", _msg)
            st.rerun()
    with _hb2:
        if st.button("🧹 카드 3개 비우기", use_container_width=True,
                     help="입력한 카드 내용을 모두 지우고 처음부터 다시 씁니다."):
            _init_highlight_cards()
            _clear_highlight_widget_state()
            st.session_state.pop("_hl_autofill_msg", None)
            st.rerun()
    _amsg = st.session_state.get("_hl_autofill_msg")
    if _amsg:
        (st.success if _amsg[0] == "success" else st.warning)(_amsg[1])

    img_col, form_col = st.columns([1, 1])

    # ── 왼쪽: 'Executive Summary' 페이지만 보여줌(전체 X). 범위는 자동감지+사용자 조정 ──
    with img_col:
        st.markdown("#### 📄 원본 Executive Summary")
        if not view_pdf or total == 0:
            st.warning("원본 PDF가 없어 이미지를 못 띄웁니다. 1단계에서 워드/PDF를 올려주세요.")
        else:
            npage = _pdf_page_count(view_pdf)
            if npage <= 0:
                npage = total
            _detected = [p for p in _find_es_pages(pages_text) if 1 <= p <= npage]
            # ES 페이지 범위 입력(자동감지값 기본 · 여러 장이면 '2-5'처럼)
            _default = _pages_to_str(_detected) if _detected else ""
            # ★한번 고친 값이 있으면(백업) 그걸 우선 — 추출 후 자동감지값으로 되돌아가지 않게
            st.session_state.setdefault(
                "hl_es_pages_str", st.session_state.get("_es_pages_saved") or _default)
            _raw = st.text_input(
                "Executive Summary 페이지 (여러 장이면 2-5 · 한 장이면 3)",
                key="hl_es_pages_str",
                placeholder="예: 3  또는  2-5",
                help="자동으로 찾은 값이에요. ES가 여러 페이지면 여기서 범위를 고치세요(예: 2-5).")
            st.session_state["_es_pages_saved"] = (_raw or "").strip()   # 현재 값 백업(유지용)
            es_pages = _parse_pages(_raw, npage)
            if _detected:
                st.caption(f"자동 감지: {_pages_to_str(_detected)}p (필요하면 위에서 조정)")
            if not es_pages:
                st.warning("표시할 Executive Summary 페이지가 없어요. 위에 페이지를 적거나, "
                           "없는 딜이면 오른쪽 카드를 직접 입력하세요.")
            else:
                idx = min(max(0, st.session_state.get("hl_es_idx", 0)), len(es_pages) - 1)
                if len(es_pages) > 1:
                    pc1, pc2, pc3 = st.columns([1, 2, 1])
                    with pc1:
                        if st.button("◀ 이전", key="hl_prev", use_container_width=True,
                                     disabled=(idx <= 0)):
                            st.session_state.hl_es_idx = idx - 1
                            st.rerun()
                    with pc2:
                        st.markdown(f"<div style='text-align:center; padding-top:6px;'>"
                                    f"ES <b>{es_pages[idx]}</b>p ({idx + 1}/{len(es_pages)})</div>",
                                    unsafe_allow_html=True)
                    with pc3:
                        if st.button("다음 ▶", key="hl_next", use_container_width=True,
                                     disabled=(idx >= len(es_pages) - 1)):
                            st.session_state.hl_es_idx = idx + 1
                            st.rerun()
                png = _toc_page_png(view_pdf, es_pages[idx])
                if png:
                    st.image(png, use_container_width=True)
                else:
                    st.caption(f"⚠️ {es_pages[idx]}p 렌더 실패")

    # ── 오른쪽: 카드 3개 편집 (컬럼 2중첩 방지 → 세로) ──
    with form_col:
        st.markdown("#### ✅ 하이라이트 카드 3개")
        st.caption("카드 간격은 다음 단계에서 일정하게 배치되고, 높이는 내용 양에 따라 자동 조절됩니다.")
        for i, c in enumerate(cards):
            with st.container(border=True):
                st.markdown(f"**카드 {i + 1}**")
                tk = f"hlt_{i}"
                st.session_state.setdefault(tk, c.get("title", ""))
                c["title"] = st.text_input("✓ 제목", key=tk,
                                           placeholder="예: 낮은 인허가 리스크")
                uk = f"hlu_{i}"
                st.session_state.setdefault(uk, bool(c.get("use_sub")))
                c["use_sub"] = st.checkbox("부제목 넣기", key=uk)
                if c["use_sub"]:
                    sk = f"hls_{i}"
                    st.session_state.setdefault(sk, c.get("subtitle", ""))
                    c["subtitle"] = st.text_input("부제목 (하늘색 줄)", key=sk,
                                                  placeholder="예: ‘25년 8월 실시계획인가 완료")
                ck = f"hlc_{i}"
                st.session_state.setdefault(ck, c.get("content", ""))
                c["content"] = st.text_area("내용", key=ck, height=110,
                                            placeholder="카드 본문 (여러 줄 입력 가능)")

    # ── 하이라이트 완성본 미리보기 (버튼 · 화면 확인용) ──
    st.markdown("---")
    if st.checkbox("🖼️ 하이라이트 완성본 미리보기", key="hl_preview_on",
                   help="작성한 카드 3개로 Executive Summary 슬라이드를 만들어 이미지로 보여줍니다. "
                        "(다운로드 아님 · 화면 확인용 · 수정하면 다시 생성 눌러 갱신)"):
        if st.button("완성본 이미지 생성/갱신", key="hl_preview_gen"):
            try:
                with st.spinner("완성본 슬라이드를 만드는 중..."):
                    ppt_b = _build_highlight_ppt_bytes(
                        cards, st.session_state.get("business_name", ""))
                from modules.preview import ppt_to_images
                with st.spinner("슬라이드를 이미지로 변환하는 중... (LibreOffice)"):
                    imgs, err = ppt_to_images(ppt_b, max_pages=1)
                st.session_state.hl_done_render = {"imgs": imgs, "err": err}
            except Exception as e:
                st.session_state.hl_done_render = {"imgs": [], "err": f"완성본 생성 실패: {e}"}
        rend = st.session_state.get("hl_done_render")
        if not rend:
            st.info("‘완성본 이미지 생성/갱신’을 누르면 완성된 Executive Summary가 여기에 표시됩니다.")
        elif rend.get("err"):
            st.error(f"이미지 생성 실패 — {rend['err']}")
        elif rend.get("imgs"):
            st.markdown("#### ✅ 완성본 Executive Summary")
            _ic1, _ic2, _ic3 = st.columns([1, 3, 1])
            with _ic2:
                st.image(rend["imgs"][0], use_container_width=True)

    # ── 저장 요약 + 네비게이션 ──
    st.markdown("---")
    filled = sum(1 for c in cards if (c.get("title") or c.get("content")))
    withsub = sum(1 for c in cards if c.get("use_sub") and c.get("subtitle"))
    st.success(f"저장됨 · 작성된 카드 {filled}/3개 · 부제목 사용 {withsub}개")

    nc1, _n, nc2 = st.columns([2, 4, 2])
    with nc1:
        if st.button("← 이전 단계", use_container_width=True):
            st.session_state.current_step = 1
            st.rerun()
    with nc2:
        if st.button("다음 단계 →", use_container_width=True, type="primary"):
            st.session_state.current_step = 3   # → 목차 구성(3단계)
            st.rerun()


# ─────────────────────────────────────────────
# [5단계: 하이라이트 완성본 확인] — 카드 3개로 ES 슬라이드 생성 → 이미지로 확인(다운로드 X)
# ─────────────────────────────────────────────
def _cards_to_sections(cards):
    """편집 카드 → build_executive_summary_slide가 받는 sections 형식으로 변환(읽기만)."""
    out = []
    for c in (cards or [])[:3]:
        out.append({
            "title":    (c.get("title") or "").strip(),
            "subtitle": ((c.get("subtitle") or "").strip() if c.get("use_sub") else ""),
            "content":  (c.get("content") or "").strip(),
        })
    return out


def _build_highlight_ppt_bytes(cards, business_name=""):
    """카드 3개로 Executive Summary 슬라이드 1장짜리 PPT bytes 생성(확인용).
       ★기존 build_executive_summary_slide를 그대로 호출 — 본문 생성 로직 미수정."""
    import io as _io
    from modules.page_builders import (
        create_presentation_from_template,
        build_executive_summary_slide,
        finalize_presentation,
    )
    prs = create_presentation_from_template()
    template_count = len(prs.slides)
    build_executive_summary_slide(prs, _cards_to_sections(cards), business_name)
    finalize_presentation(prs, template_count)   # 템플릿 슬라이드 제거 → ES 1장만 남김
    buf = _io.BytesIO()
    prs.save(buf)
    return buf.getvalue()


def show_step_highlight_preview():
    st.markdown("## 3단계. 하이라이트 완성본 확인")
    st.caption("2단계에서 만든 카드 3개로 Executive Summary 슬라이드를 만들어 이미지로 보여줍니다. "
               "(화면 확인용 · 다운로드 없음 · 수정은 이전 단계에서)")

    cards = st.session_state.get("highlight_cards")
    if not cards:
        st.warning("2단계에서 하이라이트 카드를 먼저 작성해주세요.")
        if st.button("← 2단계로"):
            st.session_state.current_step = 2
            st.rerun()
        return

    # 현재 카드 내용(읽기 전용 확인)
    with st.expander("현재 카드 내용 (읽기 전용)", expanded=False):
        for i, c in enumerate(cards, start=1):
            sub = (f"  ·  부제목: {c.get('subtitle')}"
                   if c.get("use_sub") and c.get("subtitle") else "")
            st.markdown(f"**카드 {i}. {c.get('title') or '(제목 없음)'}**{sub}")
            if c.get("content"):
                st.caption(c["content"][:250])

    if st.button("🖼️ 완성본 이미지 생성", type="primary"):
        try:
            with st.spinner("완성본 슬라이드를 만드는 중..."):
                ppt_b = _build_highlight_ppt_bytes(
                    cards, st.session_state.get("business_name", ""))
            from modules.preview import ppt_to_images
            with st.spinner("슬라이드를 이미지로 변환하는 중... (LibreOffice)"):
                imgs, err = ppt_to_images(ppt_b, max_pages=1)
            st.session_state.hl_done_render = {"imgs": imgs, "err": err}
        except Exception as e:
            st.session_state.hl_done_render = {"imgs": [], "err": f"완성본 생성 실패: {e}"}

    rend = st.session_state.get("hl_done_render")
    if not rend:
        st.info("‘🖼️ 완성본 이미지 생성’을 누르면 완성된 Executive Summary가 여기에 표시됩니다.")
    elif rend.get("err"):
        st.error(f"이미지 생성 실패 — {rend['err']}")
    elif rend.get("imgs"):
        st.markdown("#### ✅ 완성본 Executive Summary")
        # 사진이 너무 커서 가운데 좁게(약 60%) 표시
        _ic1, _ic2, _ic3 = st.columns([1, 3, 1])
        with _ic2:
            st.image(rend["imgs"][0], use_container_width=True)
    else:
        st.warning("표시할 이미지가 없습니다.")

    st.markdown("---")
    nc1, _n, nc2 = st.columns([2, 4, 2])
    with nc1:
        if st.button("← 이전 (수정하러 가기)", use_container_width=True):
            st.session_state.pop("hl_done_render", None)   # 수정하면 완성본은 다시 만들도록
            st.session_state.current_step = 2
            st.rerun()
    with nc2:
        if st.button("다음 단계 →", use_container_width=True, type="primary"):
            st.session_state.current_step = 4
            st.rerun()


# ─────────────────────────────────────────────
# ['처음으로' — 현재 작업 전체 초기화 후 1단계로]
# ─────────────────────────────────────────────
def _reset_to_start():
    """로그인 상태만 남기고 모든 세션 상태(업로드·추출·선택·생성 PPT 등)를 비운 뒤 1단계로.
       삭제된 키들은 스크립트 상단의 기본값 초기화 블록에서 다음 rerun 때 재생성된다.
       (메모는 memos.json에 저장돼 있어 세션 초기화와 무관하게 그대로 유지됨)"""
    for k in list(st.session_state.keys()):
        if k != "logged_in":
            del st.session_state[k]
    st.session_state.current_step = 1
    st.rerun()


def _confirm_home_dialog():
    """st.dialog 지원 시 확인 팝업으로 '처음으로' 확인/취소."""
    @st.dialog("처음으로")
    def _dlg():
        st.write("정말 처음으로 가시겠어요? 현재 작업이 초기화됩니다.")
        c1, c2 = st.columns(2)
        if c1.button("확인", type="primary", use_container_width=True):
            _reset_to_start()
        if c2.button("취소", use_container_width=True):
            st.rerun()        # 팝업만 닫고 현재 상태 유지
    _dlg()


# ─────────────────────────────────────────────
# [변환 작업 탭 전체 라우터]
# current_step 값에 따라 각 단계 화면을 표시합니다
# ─────────────────────────────────────────────
def show_conversion_tab():
    # 상단: '처음으로' 버튼 + Stepper
    _has_dialog = hasattr(st, "dialog")
    _hc1, _hc2 = st.columns([5, 1])
    with _hc2:
        if st.button("🏠 처음으로", use_container_width=True, help="현재 작업을 초기화하고 1단계로 돌아갑니다."):
            if _has_dialog:
                _confirm_home_dialog()
            else:
                st.session_state._show_home_confirm = True

    # st.dialog 미지원 환경: 인라인 확인 영역
    if not _has_dialog and st.session_state.get("_show_home_confirm"):
        st.warning("정말 처음으로 가시겠어요? 현재 작업이 초기화됩니다.")
        _cc1, _cc2 = st.columns(2)
        if _cc1.button("확인", type="primary", use_container_width=True):
            _reset_to_start()
        if _cc2.button("취소", use_container_width=True):
            st.session_state._show_home_confirm = False
            st.rerun()

    # 상단 Stepper 항상 표시
    render_stepper(st.session_state.current_step)

    if st.session_state.current_step == 1:
        show_step1()
    elif st.session_state.current_step == 2:
        show_step_highlight()
    elif st.session_state.current_step == 3:
        show_step_toc()
    elif st.session_state.current_step == 4:
        show_step_arrange()
    elif st.session_state.current_step == 5:
        show_step3()
    elif st.session_state.current_step == 6:
        show_step4()
    else:
        st.info(f"📌 {st.session_state.current_step}단계는 추후 구현 예정입니다.")


# ─────────────────────────────────────────────
# [메인 화면 함수]
# ─────────────────────────────────────────────
def show_main():
    # ── 사이드바: 로그아웃 버튼 ──
    with st.sidebar:
        st.markdown("### IM 생성기")
        st.markdown("---")
        # 로그아웃 버튼 (여기를 수정하면 로그아웃 버튼 텍스트가 바뀝니다)
        if st.button("🚪 로그아웃", use_container_width=True):
            # 로그아웃 시 모든 작업 데이터 초기화
            st.session_state.logged_in = False
            st.session_state.current_step = 1
            st.session_state.uploaded_file = None
            st.session_state.extracted_data = None
            st.session_state.business_name = ""
            st.session_state.toc_count = 4
            st.session_state.month_en = datetime.now().strftime("%B")
            st.session_state.year = str(datetime.now().year)
            st.session_state.cover_image_index = 0
            st.session_state.cover_image_bytes = None
            st.session_state.parsed_pages = []
            st.session_state.section_img_idx_list   = [0, 0, 0]
            st.session_state.section_img_bytes_list = [None, None, None]
            st.session_state.toc_img_idx   = 0
            st.session_state.toc_img_bytes = None
            st.rerun()  # 로그인 화면으로 전환

    # ── 메인 상단 제목 ──
    # 여기를 수정하면 메인 화면 상단 제목이 바뀝니다
    st.markdown(
        "<h1 style='text-align:center;'>IM 생성기</h1>",
        unsafe_allow_html=True,
    )
    st.markdown("---")

    # ── 변환 작업 화면 ──
    show_conversion_tab()

    # ── 하단 푸터 ──
    # 여기를 수정하면 하단 저작권 문구가 바뀝니다
    st.markdown("<br><br>", unsafe_allow_html=True)
    st.markdown(
        "<p style='text-align:center; color:lightgray; font-size:12px;'>"
        "ⓒ 2026 Rainfield Investment Advisory"
        "</p>",
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# [화면 분기]
# - 비밀번호 제거: 접근 코드 없이 항상 메인 화면 표시
#   (통합 대시보드에서만 비밀번호로 접근을 통제합니다)
# - 다시 로그인 화면을 켜려면 아래 두 줄을 지우고 원래 분기(if ...: show_main() else: show_login())로 되돌리면 됩니다.
# ─────────────────────────────────────────────
show_main()
