# -*- coding: utf-8 -*-
"""IM 회사양식 — 원본 IM(PDF)을 내용 그대로 두고 회사 양식만 입혀 워드로.

바꾸는 것   : 표지 · 위 머리말 · 아래 회사 로고 · 글꼴(Pretendard) · 표 머리글 색(회사 네이비) · 맨 끝 연락처 쪽
더하는 것   : Ⅰ. 사모사채 개요 (원본 하이라이트 뒤)
안 바꾸는 것: 내용 · 원본 1쪽 = 결과 1쪽
규격        : 회사 양식(A4·여백·머리말/꼬리말) 고정. 원본 쪽이 크면 비율 그대로 줄여 넣는다.
받는 것     : **PDF 만**(2026-10-08 사용자 확정 — 워드 원본은 결과가 엉망이라, 워드뿐이면 PDF 로 저장해 올리라고 안내)

양식은 사용자가 만든 `양식\\형식_1.docx` → 딜 내용을 지운 `양식\\회사양식_틀.docx`. 엔진은 pdf_input.py · engine.py.

★세 도구와 한 앱에서 돈다(IM\\app.py 가 runpy 로 실행). 화면 상태 이름은 전부 'cf_' 로 시작.
★입력칸 값은 'w_cf_…'(입력칸) 와 'cf_…'(보관) 를 따로 둔다. 스트림릿은 화면에서 사라진 입력칸의
  값을 지우므로, 입력칸 이름으로만 두면 2단계에서 적은 값이 3단계에서 버튼을 누르는 순간 날아간다.

혼자 돌릴 때 :  streamlit run app.py --server.port 8620
"""
import os
import re

import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
STEP_NAMES = ["원본 올리기", "표지·사모사채 개요", "만들기"]
BOND_ROWS = ["사모사채명", "사채유형", "발행인", "기초자산", "발행금액",
             "발행일", "만기일", "금융조건", "이자지급주기"]
PHOTO_COLS = 6          # 사진 고르기 — 한 줄에 몇 장(원본 사진 전부를 한 화면에 펼친다)
SS = st.session_state

try:      # 통합 앱이 먼저 불렀으면 두 번 부를 수 없다
    st.set_page_config(page_title="IM 회사양식", page_icon="📝", layout="wide")
except Exception:
    pass


# ── 단계 표시줄 (다른 IM 도구와 같은 모양) ──────────
def render_stepper(current):
    st.markdown("""
    <style>
    .stp-wrap{display:flex;align-items:center;justify-content:center;flex-wrap:wrap;
              gap:2px;padding:12px 0 16px 0;}
    .stp-item{display:flex;flex-direction:column;align-items:center;min-width:82px;max-width:110px;}
    .stp-circle{width:32px;height:32px;border-radius:50%;display:flex;align-items:center;
                justify-content:center;font-size:13px;font-weight:bold;margin-bottom:4px;}
    .stp-done{background:#1a7f37;color:#fff;}
    .stp-active{background:#08377C;color:#fff;box-shadow:0 0 0 3px #cdd4e6;}
    .stp-todo{background:#e5e7eb;color:#9ca3af;}
    .stp-label{font-size:11px;text-align:center;line-height:1.3;word-break:keep-all;}
    .stp-l-done{color:#1a7f37;font-weight:600;}
    .stp-l-active{color:#08377C;font-weight:700;}
    .stp-l-todo{color:#9ca3af;}
    .stp-conn{width:26px;height:2px;margin-bottom:20px;}
    .stp-c-done{background:#1a7f37;} .stp-c-todo{background:#e5e7eb;}
    </style>
    """, unsafe_allow_html=True)
    html = '<div class="stp-wrap">'
    for i, name in enumerate(STEP_NAMES):
        n = i + 1
        if n < current:
            cc, lc, tx, conn = "stp-done", "stp-l-done", "✓", "stp-c-done"
        elif n == current:
            cc, lc, tx, conn = "stp-active", "stp-l-active", str(n), "stp-c-todo"
        else:
            cc, lc, tx, conn = "stp-todo", "stp-l-todo", str(n), "stp-c-todo"
        html += (f'<div class="stp-item"><div class="stp-circle {cc}">{tx}</div>'
                 f'<div class="stp-label {lc}">{name}</div></div>')
        if n < len(STEP_NAMES):
            html += f'<div class="stp-conn {conn}"></div>'
    st.markdown(html + "</div>", unsafe_allow_html=True)
    st.markdown("---")


def _goto(n):
    SS["cf_step"] = n
    st.rerun()


def _field(label, store, default="", help=None):
    """보관값(store)을 바탕으로 입력칸을 그리고, 바뀐 값을 다시 보관한다."""
    SS.setdefault(store, default)
    val = st.text_input(label, value=SS[store], key="w_" + store, help=help)
    SS[store] = val
    return val


def _prepare(name, data):
    """원본에서 표지 제목·사진·하이라이트 쪽·끝쪽을 뽑아 보관한다(올릴 때 한 번)."""
    import fitz
    import engine
    import pdf_input
    from datetime import date
    doc = fitz.open(stream=data, filetype="pdf")
    tall = sum(1 for p in doc if p.rect.height >= p.rect.width)
    info = {"pages": doc.page_count,
            "orient": "세로" if tall == doc.page_count else "가로" if tall == 0 else f"섞임(세로 {tall}쪽)",
            "text_per_page": pdf_input.text_per_page(data),
            "last_contact": pdf_input.last_page_is_contact(doc),
            "highlight": pdf_input.find_highlight_pages(data)}
    SS["cf_info"] = info
    title = (pdf_input.suggest_title(data) + ["", ""])[:2]
    SS["cf_title1"], SS["cf_title2"] = title
    base = re.sub(r"^\[?IM\]?[\s\-_]*", "", os.path.splitext(name)[0]).strip() or title[0]
    SS["cf_header"] = f"{base} 사모사채 제안서 {date.today().year}"
    SS["cf_date"] = engine.default_date_text()
    SS["cf_photos"] = pdf_input.all_photos(data)
    SS["cf_photo_pick"] = 0 if SS["cf_photos"] else -1
    SS["cf_highlight"] = info["highlight"]
    SS["cf_drop_cover"] = True
    SS["cf_drop_tail"] = bool(info["last_contact"])
    for k in ("cf_photo_custom", "cf_out", "cf_rep", "cf_build_err", "cf_fname", "cf_thumbs", "cf_edge_pages"):
        SS.pop(k, None)


SS.setdefault("cf_step", 1)
step = SS["cf_step"]

st.markdown("""
<style>
.block-container{padding-top:1.6rem;max-width:100%;
                 padding-left:2.2rem;padding-right:2.2rem;}
h1,h2,h3{color:#08377C;}
div.stButton>button{border-radius:4px;border:1px solid #08377C;color:#08377C;
                    background:#fff;font-weight:700;}
div.stButton>button:hover{background:#eef3fa;color:#08377C;border-color:#08377C;}
div.stButton>button[kind="primary"]{background:#1A2B5E;color:#fff;border-color:#1A2B5E;}
div.stButton>button[kind="primary"]:hover{background:#08377C;border-color:#08377C;}
div.stDownloadButton>button{background:#08377C;color:#fff;border:1px solid #08377C;
                            border-radius:4px;font-weight:700;}
div.stDownloadButton>button:hover{background:#0063A1;border-color:#0063A1;color:#fff;}
hr{border-color:#dfe6f0;}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='margin:0 0 .1rem 0;line-height:1.1;'>IM 회사양식</h1>",
            unsafe_allow_html=True)
st.caption("받은 IM(PDF)을 **내용 그대로** 두고, 회사 양식(표지·머리말·로고·글꼴·표 머리글 색·연락처)만 "
           "입혀 **워드**로 만듭니다. 사모사채 개요도 넣습니다.")
render_stepper(step)


# ══════════════════════════════════════════════════
# 1단계 — 원본 올리기 (PDF 만)
# ══════════════════════════════════════════════════
if step == 1:
    st.markdown("### 1단계 · 원본 IM(PDF) 올리기")
    st.caption("받은 IM 의 **PDF** 를 올리세요(워드는 받지 않습니다). 원본 한 쪽이 결과 한 쪽으로 그대로 옮겨지고, "
               "쪽 크기만 회사 양식(A4)에 맞춥니다.")

    # 두 상자를 같은 높이로(사용자 요청) — st.columns 의 알림 상자는 내용 길이대로 높이가 달라져 직접 그린다
    st.markdown("""
    <div style="display:flex;gap:16px;align-items:stretch;margin:4px 0 12px 0;">
      <div style="flex:1;background:#e8f7ec;border-radius:8px;padding:14px 18px;color:#14532d;">
        <b>그대로 두는 것</b>
        <ul style="margin:8px 0 0 0;">
          <li>본문 내용·표·그림</li>
          <li>원본 1쪽 = 결과 1쪽</li>
          <li>(원본 쪽이 회사 양식보다 크면 비율 그대로 줄여서 넣습니다)</li>
        </ul>
      </div>
      <div style="flex:1;background:#e8f1fb;border-radius:8px;padding:14px 18px;color:#0b3a6e;">
        <b>회사 양식으로 바꾸는 것</b>
        <ul style="margin:8px 0 0 0;">
          <li>표지</li>
          <li>위쪽 머리말 · 아래쪽 로고 · 쪽번호 (원본 증권사 것을 빼고 회사 것으로)</li>
          <li>원본 담당자 연락처(이름·전화·메일)는 뺍니다</li>
          <li>글꼴(Pretendard) · 표 머리글 색(회사 네이비)</li>
          <li>맨 끝 연락처 쪽</li>
        </ul>
        <div style="margin-top:8px;"><b>더하는 것</b> : Ⅰ. 사모사채 개요 (원본 하이라이트 바로 뒤)</div>
      </div>
    </div>
    """, unsafe_allow_html=True)

    # ★PDF 만 받는다는 것을 접지 말고 맨 앞에(사용자 지적 2026-10-08) — 접힌 칸 속에 두면 사람들이
    #   못 보고 정리 안 된 워드를 그대로 올리려 한다.
    st.warning(
        "📄 **PDF 파일만 올릴 수 있습니다. 워드(.docx·.doc)는 받지 않습니다.**\n\n"
        "받은 IM 이 **워드뿐이라면, 워드에서 PDF 로 저장한 뒤 그 PDF 를 올려 주세요.**\n"
        "1. 워드 파일을 워드로 엽니다.\n"
        "2. 키보드 **F12** → 아래 **파일 형식** 칸에서 **PDF (\\*.pdf)** 선택 → **저장**.\n"
        "3. 같은 폴더에 생긴 **PDF** 를 아래에 올립니다.\n\n"
        "**왜 PDF 만?** 받은 워드 원본은 서식이 제각각이고 정리가 안 된 경우가 많아 결과가 깨집니다. "
        "워드 → PDF 는 누구나 30초면 바꿀 수 있고, PDF 는 보이는 모양이 고정돼 있어 원본 그대로 옮기기 좋습니다.")
    with st.expander("🚫 올려도 바꿀 수 없는 PDF"):
        st.markdown(
            "글자를 마우스로 드래그해도 선택되지 않는 PDF(스캔본·글자를 그림으로 바꿔 저장한 것)는 "
            "워드 글자로 바꿀 수 없습니다. 보낸 쪽에 원본(워드나 글자가 살아 있는 PDF)을 다시 받아 주세요.")

    up = st.file_uploader("원본 IM — PDF 만", type=["pdf"], key="w_cf_up",
                          help="워드 파일은 올라가지 않습니다. 워드에서 F12 → PDF 로 저장한 뒤 올려 주세요.")
    if up is not None and SS.get("cf_name") != up.name:
        SS["cf_name"] = up.name
        SS["cf_bytes"] = up.getvalue()
        SS.pop("cf_err", None)
        with st.spinner("원본을 살펴보는 중…"):
            try:
                _prepare(up.name, SS["cf_bytes"])
            except Exception as e:
                SS["cf_info"] = None
                SS["cf_err"] = f"PDF 를 열지 못했습니다 : {e}"

    if SS.get("cf_err"):
        st.error(SS["cf_err"])

    info = SS.get("cf_info")
    scanned = bool(info and info.get("text_per_page", 999) < 100)
    if info:
        st.success(f"올린 파일 — **{SS['cf_name']}**")
        m1, m2, m3 = st.columns(3)
        m1.metric("쪽수", f"{info['pages']}쪽")
        m2.metric("쪽 방향", info["orient"])
        m3.metric("원본 사진", f"{len(SS.get('cf_photos') or [])}장")
        if scanned:
            st.error(
                "🚫 **이 PDF 는 글자를 읽을 수 없습니다** (쪽당 글자 "
                f"{info['text_per_page']:.0f}자 — 보통 IM 은 600~1,100자).\n\n"
                "스캔본이거나 글자를 그림으로 바꿔 저장한 PDF 입니다. 겉으로는 글자가 또렷해도 컴퓨터는 "
                "그림으로만 봅니다(실제 사례: 신사동 평화빌딩 PDF). 이런 PDF 는 워드로 바꿀 수 없습니다.\n\n"
                "**워드 원본**을 받아 PDF 로 저장(F12)해서 올리거나, 보낸 쪽에 원본을 다시 요청해 주세요.")
        else:
            st.markdown("---")
            _b1, _b2, _b3 = st.columns([2, 4, 2])
            if _b3.button("다음 단계 →", type="primary", use_container_width=True):
                _goto(2)


# ══════════════════════════════════════════════════
# 2단계 — 표지 · 사모사채 개요
# ══════════════════════════════════════════════════
elif step == 2:
    st.markdown("### 2단계 · 표지 · 사모사채 개요")
    if not SS.get("cf_info"):
        st.warning("먼저 1단계에서 원본을 올려 주세요.")
    else:
        st.markdown("#### 표지 · 머리말")
        st.caption("원본 표지에서 제목을 찾아 미리 채워 두었습니다. 고칠 곳만 고치세요. "
                   "제목이 길면 표지의 제목 칸이 알아서 늘어납니다.")
        a, b = st.columns(2)
        with a:
            _field("표지 제목 (첫째 줄)", "cf_title1")
            _field("표지 제목 (둘째 줄 · 없으면 비워 두세요)", "cf_title2")
        with b:
            _field("위쪽 머리말 문구", "cf_header",
                   help="본문 모든 쪽 맨 위에 들어가는 문구입니다. 예) 신촌지역 브릿지 대출 사모사채 제안서 2026")
            _field("표지 날짜", "cf_date", help="영어로 적습니다. 예) October 2026")

        st.markdown("#### 표지 사진")
        photos = SS.get("cf_photos") or []
        pick = SS.get("cf_photo_pick", -1)
        if photos:
            st.caption(f"원본에 든 사진 **{len(photos)}장 전부**입니다(표지에 쓸 만한 것부터). 사진은 **자르지 않고** "
                       "원본 비율 그대로 크기만 줄여 표지 제목과 Disclaimer 사이에 넣고, 화질은 또렷하게 다듬습니다.")
            for start in range(0, len(photos), PHOTO_COLS):
                cols = st.columns(PHOTO_COLS)
                for i, col in zip(range(start, min(start + PHOTO_COLS, len(photos))), cols):
                    p = photos[i]
                    with col:
                        st.image(p["blob"], use_container_width=True)
                        picked = pick == i
                        if st.button("✅ 선택됨" if picked else f"쓰기 ({p['page']}쪽)", key=f"w_cf_pick_{i}",
                                     type="primary" if picked else "secondary", use_container_width=True):
                            SS["cf_photo_pick"] = i
                            st.rerun()
        else:
            st.caption("원본에서 사진을 찾지 못했습니다.")
        o1, o2 = st.columns([1, 2])
        with o1:
            if st.button("✅ 사진 없이" if pick == -1 else "사진 없이 만들기", key="w_cf_pick_none",
                         type="primary" if pick == -1 else "secondary", use_container_width=True):
                SS["cf_photo_pick"] = -1
                st.rerun()
        with o2:
            mine = st.file_uploader("다른 사진 직접 올리기 (jpg·png)", type=["jpg", "jpeg", "png"],
                                    key="w_cf_photo_up")
            if mine is not None:
                SS["cf_photo_custom"] = mine.getvalue()
                SS["cf_photo_pick"] = -2
            if SS.get("cf_photo_pick") == -2 and SS.get("cf_photo_custom"):
                st.image(SS["cf_photo_custom"], caption="✅ 직접 올린 사진을 씁니다", width=260)

        st.markdown("---")
        st.markdown("#### Ⅰ. 사모사채 개요")
        st.caption("왼쪽 구분은 항상 같고, 오른쪽 내용만 이번 딜에 맞게 적어 주세요. 비워 둔 칸은 빈칸으로 들어갑니다. "
                   "원본 하이라이트(Executive Summary 등) 바로 뒤에 들어갑니다.")
        for row in BOND_ROWS:
            k1, k2 = st.columns([1, 4])
            k1.markdown(f"<div style='padding-top:8px;font-weight:700;color:#08377C;'>{row}</div>",
                        unsafe_allow_html=True)
            with k2:
                SS.setdefault(f"cf_bond_{row}", "")
                SS[f"cf_bond_{row}"] = st.text_input(row, value=SS[f"cf_bond_{row}"],
                                                     key=f"w_cf_bond_{row}", label_visibility="collapsed")

    st.markdown("---")
    p1, _sp, p3 = st.columns([2, 4, 2])
    if p1.button("← 이전 단계", use_container_width=True):
        _goto(1)
    if SS.get("cf_info") and p3.button("다음 단계 →", type="primary", use_container_width=True):
        _goto(3)


# ══════════════════════════════════════════════════
# 3단계 — 만들기
# ══════════════════════════════════════════════════
elif step == 3:
    st.markdown("### 3단계 · 만들기")
    info = SS.get("cf_info")
    if not SS.get("cf_bytes") or not info:
        st.warning("먼저 1단계에서 원본을 올려 주세요.")
    else:
        import engine
        import pdf_input
        n = info["pages"]
        filled = [r for r in BOND_ROWS if (SS.get(f"cf_bond_{r}") or "").strip()]
        st.write(f"**{SS.get('cf_name', '')}** · {n}쪽 · 표지 제목 : **{SS.get('cf_title1', '')} "
                 f"{SS.get('cf_title2', '')}** · 사모사채 개요 : {len(filled)}/{len(BOND_ROWS)}칸 채움")

        # ★회사 사람들이 처음 봐도 알게 자세히(사용자 지적 2026-10-08) — '마지막 쪽 빼기' 를 '하나 빠뜨린
        #   체크' 로 보고 본문 쪽(성수동 66쪽 인허가 일정)을 체크한 일이 있었다. 쪽 그림을 옆에 보여 준다.
        st.markdown("#### 원본에서 뺄 쪽")
        st.info(
            "**✅ 체크하지 않아도 자동으로 빼는 것** — 따로 하실 일 없습니다.\n"
            "- 원본 모든 쪽의 **위쪽 머리말·아래쪽 꼬리말**(증권사 이름·사업명·로고·'Confidential' 문구)\n"
            "- 원본 **쪽번호** (결과물에는 회사 양식 쪽번호가 새로 들어갑니다)\n"
            "- 원본 **담당자 연락처**(이름·직급·전화·메일 묶음) — 어느 쪽에 있든\n\n"
            "※ 본문 안의 표·그림은 손대지 않습니다. 그림 **안에** 그려진 증권사 로고(금융구조도의 참여사 표시 등)도 "
            "딜 내용이라 그대로 둡니다.\n\n"
            "**☑️ 아래 두 칸은 '쪽을 통째로' 빼는 것**이라 쪽 그림을 보고 직접 정해 주세요.")
        if "cf_edge_pages" not in SS:
            SS["cf_edge_pages"] = (pdf_input.page_png(SS["cf_bytes"], 0, zoom=0.3),
                                   pdf_input.page_png(SS["cf_bytes"], n - 1, zoom=0.3))
        e1, e2 = st.columns(2)
        with e1:
            ic, tc = st.columns([1, 3])
            ic.image(SS["cf_edge_pages"][0], caption="원본 1쪽", use_container_width=True)
            with tc:
                SS["cf_drop_cover"] = st.checkbox("원본 1쪽(증권사 표지)을 빼고 회사 표지로 바꾸기",
                                                  value=SS.get("cf_drop_cover", True), key="w_cf_drop_cover")
                st.caption("거의 모든 IM 은 1쪽이 증권사 표지라 처음부터 체크돼 있습니다. "
                           "1쪽이 표지가 아니라 본문이면 체크를 풀어 주세요.")
        with e2:
            ic, tc = st.columns([1, 3])
            ic.image(SS["cf_edge_pages"][1], caption=f"원본 마지막 쪽({n}쪽)", use_container_width=True)
            with tc:
                SS["cf_drop_tail"] = st.checkbox(
                    f"원본 마지막 쪽({n}쪽)을 통째로 빼기 — 증권사 연락처만 있는 쪽일 때만",
                    value=SS.get("cf_drop_tail", False), key="w_cf_drop_tail")
                st.caption(
                    ("🔎 이 쪽에서 메일·전화번호를 찾아 **연락처 쪽으로 보여 체크해 두었습니다.** "
                     if info.get("last_contact") else
                     "🔎 이 쪽은 **연락처 쪽으로 보이지 않아 체크하지 않았습니다.** ")
                    + "IM 끝에 증권사 담당자 연락처나 'End of Document' 만 있는 쪽이 있으면 그 쪽을 통째로 빼고 "
                    "회사 연락처 쪽으로 바꾸는 칸입니다. **마지막 쪽이 본문(표·설명)이면 체크하지 마세요** — "
                    "체크하면 그 쪽 내용이 통째로 빠집니다. (본문 쪽 안의 담당자 연락처는 체크하지 않아도 위처럼 "
                    "자동으로 지웁니다.)")

        st.markdown("#### 하이라이트 쪽 (사모사채 개요가 이 뒤에 들어갑니다)")
        found = info.get("highlight") or []
        st.caption(("원본에서 하이라이트로 보이는 쪽을 찾았습니다 : **" + ", ".join(f"{p}쪽" for p in found) + "**. "
                    if found else "원본에서 하이라이트 쪽을 찾지 못했습니다(그러면 표지 바로 다음 쪽에 넣습니다). ")
                   + "아래에서 쪽 그림을 보고 맞게 고르세요. 하이라이트 마지막 쪽에 여백이 넉넉하면 그 여백에 이어서, "
                   "꽉 차 있으면 다음 쪽에 넣습니다.")
        if "cf_thumbs" not in SS:
            SS["cf_thumbs"] = [pdf_input.page_png(SS["cf_bytes"], i, zoom=0.35) for i in range(1, min(n, 9))]
        tcols = st.columns(len(SS["cf_thumbs"]) or 1)
        for i, (col, png) in enumerate(zip(tcols, SS["cf_thumbs"])):
            col.image(png, caption=f"{i + 2}쪽", use_container_width=True)
        SS["cf_highlight"] = st.multiselect("하이라이트 쪽", list(range(2, n + 1)),
                                            default=[p for p in SS.get("cf_highlight", []) if 2 <= p <= n],
                                            format_func=lambda p: f"{p}쪽", key="w_cf_highlight")

        st.markdown("#### 파일 이름")
        _base = os.path.splitext(SS.get("cf_name") or "IM")[0]
        SS.setdefault("cf_fname", f"{_base}_회사양식")
        f1, f2 = st.columns([5, 1])
        with f1:
            _field("파일명", "cf_fname", help="내려받을 파일 이름입니다. 확장자(.docx)는 자동으로 붙습니다.")
        f2.markdown("<div style='padding-top:34px;color:#5b6b85;'>.docx</div>", unsafe_allow_html=True)

        if not engine.template_path():
            st.error("회사 양식 파일(양식\\회사양식_틀.docx)을 찾지 못해 만들 수 없습니다. 관리자에게 알려 주세요.")
        elif st.button("📝 회사 양식 워드 만들기", type="primary"):
            for k in ("cf_out", "cf_rep", "cf_build_err"):
                SS.pop(k, None)
            pick = SS.get("cf_photo_pick", -1)
            photos = SS.get("cf_photos") or []
            photo = (SS.get("cf_photo_custom") if pick == -2 else
                     photos[pick]["blob"] if pick is not None and 0 <= pick < len(photos) else None)
            bond = {r: SS.get(f"cf_bond_{r}", "") for r in BOND_ROWS}
            try:
                bar = st.progress(0.0, text="만드는 중입니다… 이 화면을 닫지 말고 기다려 주세요.")

                def _prog(v, msg):
                    bar.progress(min(max(v, 0.0), 1.0), text=f"만드는 중입니다… {msg}")
                body, rep1 = pdf_input.convert(SS["cf_bytes"], drop_cover=SS.get("cf_drop_cover", True),
                                               drop_last=SS.get("cf_drop_tail", False), progress=_prog)
                _prog(0.98, "표지·사모사채 개요·연락처 넣는 중…")
                place = pdf_input.bond_place(rep1, SS.get("cf_highlight") or [])
                out, rep2 = engine.build(body, title_lines=[SS.get("cf_title1", ""), SS.get("cf_title2", "")],
                                         header_text=SS.get("cf_header", ""), date_text=SS.get("cf_date", ""),
                                         photo=photo, bond=bond, bond_place=place)
                bar.empty()
                SS["cf_out"] = out
                SS["cf_rep"] = {**rep1, **rep2}
            except Exception:
                import traceback
                SS["cf_build_err"] = traceback.format_exc()

        if SS.get("cf_build_err"):
            st.error("만들지 못했습니다. 아래 내용을 관리자에게 보내 주세요.")
            st.code(SS["cf_build_err"])

        if SS.get("cf_out"):
            rep = SS.get("cf_rep") or {}
            st.success("✅ 다 만들었습니다. 아래 버튼으로 내려받으세요.")
            nm = re.sub(r'[\\/:*?"<>|]', "_", (SS.get("cf_fname") or "").strip())[:80] or f"{_base}_회사양식"
            if nm.lower().endswith(".docx"):
                nm = nm[:-5]
            st.download_button("⬇️ 워드 내려받기", data=SS["cf_out"], file_name=f"{nm}.docx",
                               mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            pages = rep.get("쪽") or []
            small = [p for p in pages if p.get("비율") and p["비율"] < 0.999]
            done = [f"원본 {len(pages)}쪽을 한 쪽씩 그대로 옮겼습니다"
                    + (f" (회사 양식보다 커서 {min(p['비율'] for p in small) * 100:.0f}~"
                       f"{max(p['비율'] for p in small) * 100:.0f}% 로 줄인 쪽 {len(small)}개)." if small else "."),
                    f"표 머리글 칸 {rep.get('머리글 칸', 0)}개를 회사 네이비로 바꿨습니다.",
                    f"사모사채 개요 : {rep.get('개요 자리', '')}."]
            if rep.get("원본 표지 뺌"):
                done.append("원본 표지를 빼고 회사 표지를 넣었습니다.")
            if rep.get("원본 끝쪽 뺌"):
                done.append("원본 마지막 쪽(증권사 연락처)을 빼고 회사 연락처 쪽을 넣었습니다.")
            if rep.get("머리말·꼬리말 지운 쪽"):
                done.append(f"쪽마다 되풀이되던 증권사 머리말·로고·쪽번호를 {rep['머리말·꼬리말 지운 쪽']}쪽에서 지웠습니다.")
            if rep.get("연락처 지운 쪽"):
                done.append("원본 담당자 연락처(이름·전화·메일)를 지웠습니다 : "
                            + ", ".join(f"{p}쪽" for p in rep["연락처 지운 쪽"]) + ".")
            if rep.get("그림으로 넣은 쪽"):
                done.append("⚠️ 워드로 바꾸지 못해 **원본 쪽을 그림으로** 넣은 쪽 : "
                            + ", ".join(f"{p}쪽" for p in rep["그림으로 넣은 쪽"]) + " (고칠 수 없으니 확인해 주세요).")
            st.markdown("\n".join("- " + t for t in done))
            with st.expander("⚠️ 받은 뒤 확인해 주세요"):
                st.markdown(
                    "- **글꼴 Pretendard** 가 깔린 컴퓨터에서 열어야 같은 모양으로 보입니다. 밖으로 보낼 때는 "
                    "워드에서 **PDF 로 저장해서** 보내면 어디서나 똑같이 보입니다.\n"
                    "- PDF 를 워드로 바꾸는 과정에서 **표·줄바꿈이 원본과 조금 다를 수 있습니다.** 한 번 훑어봐 주세요.\n"
                    "- 원본의 **기울임(이탤릭) 글씨**는 Pretendard 에 기울임 글꼴이 없어 다른 글꼴로 보일 수 있습니다.\n"
                    "- 원본 그림 **안에** 박힌 증권사 로고(구조도 그림 등)는 그림이라 바꿀 수 없습니다.")

    st.markdown("---")
    p1, _sp, _p3 = st.columns([2, 4, 2])
    if p1.button("← 이전 단계", use_container_width=True):
        _goto(2)
