"""
ui_components.py
─────────────────────────────────────────────────────────
화면에서 반복적으로 사용하는 UI 컴포넌트 모음.
app.py 에서 import 해서 사용합니다.
─────────────────────────────────────────────────────────
"""

import io
import streamlit as st
from PIL import Image


# ─────────────────────────────────────────────
# [9단계 이름 정의]
# 여기를 수정하면 Stepper에 표시되는 단계 이름이 바뀝니다
# ─────────────────────────────────────────────
STEP_NAMES = [
    "파일 업로드",        # 1단계 (업로드 + 추출 + 사업명)
    "하이라이트 구성",    # 2단계 (ES 카드 3개 입력 + 완성본 미리보기 버튼)
    "목차 구성",          # 3단계 (원본 페이지 보며 목차·소제목 작성 + 미리보기)
    "배치 확인",          # 4단계 (배치한 원본 페이지를 순서대로 이미지로 확인·수정)
    "레이아웃 선택",      # 5단계 (날짜·표지 설정 + 표지·목차·섹션 미리보기)
    "이미지 선택",        # 6단계 (PPT 생성·다운로드)
]


# ─────────────────────────────────────────────
# [Stepper 컴포넌트]
# 상단에 진행 단계를 시각적으로 표시합니다
# ─────────────────────────────────────────────
def render_stepper(current_step: int):
    """
    9단계 진행 표시바를 화면 상단에 그립니다.

    current_step: 현재 진행 중인 단계 번호 (1~9)
    - 완료된 단계: 초록색 ✓
    - 현재 단계: 파란색 강조
    - 미완료 단계: 회색
    """

    # CSS 스타일 정의
    # 여기를 수정하면 Stepper 색상과 모양이 바뀝니다
    st.markdown("""
    <style>
    .stepper-wrap {
        display: flex;
        align-items: center;
        justify-content: center;
        flex-wrap: wrap;
        gap: 2px;
        padding: 12px 0 16px 0;
    }
    .step-item {
        display: flex;
        flex-direction: column;
        align-items: center;
        min-width: 72px;
        max-width: 90px;
    }
    .step-circle {
        width: 32px;
        height: 32px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 13px;
        font-weight: bold;
        margin-bottom: 4px;
    }
    .step-done   { background:#1a7f37; color:white; }
    .step-active { background:#1A2B5E; color:white; box-shadow: 0 0 0 3px #cdd4e6; }
    .step-todo   { background:#e5e7eb; color:#9ca3af; }
    .step-label {
        font-size: 10px;
        text-align: center;
        line-height: 1.3;
        word-break: keep-all;
    }
    .step-label-done   { color: #1a7f37; font-weight: 600; }
    .step-label-active { color: #1A2B5E; font-weight: 700; }
    .step-label-todo   { color: #9ca3af; }
    .step-connector {
        width: 20px;
        height: 2px;
        margin-bottom: 20px;
    }
    .connector-done { background: #1a7f37; }
    .connector-todo { background: #e5e7eb; }
    </style>
    """, unsafe_allow_html=True)

    # HTML 조합
    html = '<div class="stepper-wrap">'

    for i, name in enumerate(STEP_NAMES):
        step_num = i + 1

        if step_num < current_step:
            # 완료된 단계
            circle_class = "step-circle step-done"
            label_class  = "step-label step-label-done"
            circle_text  = "✓"
            connector_class = "step-connector connector-done"
        elif step_num == current_step:
            # 현재 단계
            circle_class = "step-circle step-active"
            label_class  = "step-label step-label-active"
            circle_text  = str(step_num)
            connector_class = "step-connector connector-todo"
        else:
            # 미완료 단계
            circle_class = "step-circle step-todo"
            label_class  = "step-label step-label-todo"
            circle_text  = str(step_num)
            connector_class = "step-connector connector-todo"

        html += f"""
        <div class="step-item">
            <div class="{circle_class}">{circle_text}</div>
            <div class="{label_class}">{name}</div>
        </div>
        """

        # 마지막 단계 뒤에는 연결선 없음
        if step_num < len(STEP_NAMES):
            html += f'<div class="{connector_class}"></div>'

    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)
    st.markdown("---")


# ─────────────────────────────────────────────
# [이미지 갤러리 컴포넌트]
# 추출된 이미지를 그리드 형태로 표시합니다
# ─────────────────────────────────────────────
def render_image_gallery(images: list):
    """
    추출된 이미지 목록을 한 줄에 4개씩 썸네일 갤러리로 표시합니다.

    images: extractors.py 에서 반환하는 images 리스트
    각 항목: {'index': int, 'page': int, 'pil_image': PIL.Image, 'width': int, 'height': int}
    """

    st.markdown("### 🖼️ 추출된 이미지 갤러리")

    if not images:
        st.info("추출된 이미지가 없습니다.")
        return

    # 이미지가 20개 이상이면 안내 메시지 표시
    # 여기를 수정하면 안내 기준 개수가 바뀝니다
    if len(images) >= 20:
        st.warning(f"이미지가 {len(images)}개 있습니다. 스크롤해서 확인하세요.")

    st.caption(f"총 {len(images)}개 이미지 추출됨")

    # 한 줄에 표시할 이미지 수 (여기를 수정하면 열 수가 바뀝니다)
    COLS_PER_ROW = 4

    # 이미지를 COLS_PER_ROW 개씩 묶어서 행(row) 단위로 처리
    for row_start in range(0, len(images), COLS_PER_ROW):
        row_images = images[row_start: row_start + COLS_PER_ROW]
        cols = st.columns(COLS_PER_ROW)

        for col_idx, img_data in enumerate(row_images):
            with cols[col_idx]:
                pil_img = img_data["pil_image"]
                page    = img_data["page"]
                w       = img_data["width"]
                h       = img_data["height"]
                num     = img_data["index"] + 1  # 1부터 시작하는 번호

                # 썸네일 크기로 리사이즈 (원본 비율 유지)
                # 여기를 수정하면 썸네일 가로 최대 크기가 바뀝니다
                THUMB_WIDTH = 200
                ratio = THUMB_WIDTH / w if w > THUMB_WIDTH else 1.0
                thumb = pil_img.resize(
                    (int(w * ratio), int(h * ratio)),
                    Image.LANCZOS
                )

                st.image(thumb, use_container_width=True)
                st.caption(f"이미지 #{num}\n페이지 {page} · {w}×{h}")


# ─────────────────────────────────────────────
# [페이지별 텍스트 미리보기 컴포넌트]
# ─────────────────────────────────────────────
def render_text_preview(pages_text: list):
    """
    페이지별 텍스트를 접었다 펼 수 있는 expander 형태로 표시합니다.

    pages_text: 페이지별 텍스트 문자열 리스트
    """

    st.markdown("### 📄 페이지별 텍스트 미리보기")

    total = len(pages_text)

    if total == 0:
        st.info("추출된 텍스트가 없습니다.")
        return

    # 10페이지 이상이면 기본 접힘 상태로 시작
    # 여기를 수정하면 기본 펼침 기준 페이지 수가 바뀝니다
    DEFAULT_OPEN_THRESHOLD = 10
    default_expanded = total < DEFAULT_OPEN_THRESHOLD

    # 전체 펼치기/접기 토글 버튼
    col_l, col_r, _ = st.columns([1, 1, 4])
    with col_l:
        if st.button("전체 펼치기", key="expand_all"):
            st.session_state["text_expanded"] = True
    with col_r:
        if st.button("전체 접기", key="collapse_all"):
            st.session_state["text_expanded"] = False

    # 토글 상태 반영
    if "text_expanded" in st.session_state:
        default_expanded = st.session_state["text_expanded"]

    st.markdown("")

    for i, text in enumerate(pages_text):
        page_num   = i + 1
        char_count = len(text.strip())
        label      = f"페이지 {page_num}  (글자 수: {char_count:,}자)"

        with st.expander(label, expanded=default_expanded):
            if text.strip():
                st.text_area(
                    label=f"page_{page_num}",
                    value=text,
                    height=300,
                    disabled=True,          # 읽기 전용
                    label_visibility="collapsed",
                    key=f"text_area_page_{page_num}",
                )
            else:
                st.caption("이 페이지에는 텍스트가 없습니다.")
