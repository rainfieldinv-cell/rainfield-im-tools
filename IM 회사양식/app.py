# -*- coding: utf-8 -*-
"""IM 회사양식 — 원본 IM(워드·PDF)을 내용 그대로 두고 회사 양식만 입혀 워드로.

바꾸는 것   : 표지 · 위 머리말 · 아래 회사 로고 · 글꼴 · 맨 끝 연락처 페이지
더하는 것   : 사모사채 개요
안 바꾸는 것: 내용 · 글씨 크기 · 쪽 방향(세로는 세로, 가로는 가로)
결과        : 워드(.docx). 원본이 PDF 여도 먼저 워드로 바꾼 뒤 같은 처리를 한다.

★지금은 틀(화면)만 있다. 사용자가 회사 양식 워드를 만들어 주면 그걸 보고
  3단계의 실제 만들기를 붙인다(2026-10-06).

★세 도구와 한 앱에서 돈다(IM\\app.py 가 runpy 로 실행). 화면 상태 이름이 다른 도구와
  겹치면 값이 섞이므로 전부 'cf_' 를 앞에 붙인다(요약본도 'step' 을 쓴다).

혼자 돌릴 때 :  streamlit run app.py --server.port 8620
"""
import io
import os
import zipfile

import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))
STEP_NAMES = ["원본 올리기", "사모사채 개요", "만들기"]

# 사모사채 개요 — 구분 칸은 딜과 상관없이 늘 같다(천안·대전·넷마블 회사 제안서 실측).
#   내용 칸만 딜마다 채운다.
BOND_ROWS = ["사모사채명", "사채유형", "발행인", "기초자산", "발행금액",
             "발행일", "만기일", "금융조건", "이자지급주기"]

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
    st.session_state["cf_step"] = n
    st.rerun()


def _read_source(name, data):
    """올린 파일의 종류·쪽수·쪽 방향을 잰다. 내용은 건드리지 않는다."""
    ext = os.path.splitext(name)[1].lower()
    info = {"kind": "PDF" if ext == ".pdf" else "워드", "pages": None, "orient": "?"}
    if ext == ".pdf":
        import fitz
        doc = fitz.open(stream=data, filetype="pdf")
        info["pages"] = doc.page_count
        tall = sum(1 for p in doc if p.rect.height >= p.rect.width)
        info["orient"] = ("세로" if tall == doc.page_count else
                          "가로" if tall == 0 else f"섞임(세로 {tall}쪽)")
    else:
        import docx
        d = docx.Document(io.BytesIO(data))
        tall = sum(1 for s in d.sections if s.page_height >= s.page_width)
        n = len(d.sections)
        info["orient"] = ("세로" if tall == n else "가로" if tall == 0 else "섞임")
        info["tables"] = len(d.tables)
    return info


st.session_state.setdefault("cf_step", 1)
step = st.session_state["cf_step"]

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
hr{border-color:#dfe6f0;}
</style>
""", unsafe_allow_html=True)

st.markdown("<h1 style='margin:0 0 .1rem 0;line-height:1.1;'>IM 회사양식</h1>",
            unsafe_allow_html=True)
st.caption("받은 IM 을 **내용 그대로** 두고, 회사 양식(표지·머리말·로고·글꼴·연락처)만 "
           "입혀 **워드**로 만듭니다. 사모사채 개요도 넣습니다.")
render_stepper(step)


# ══════════════════════════════════════════════════
# 1단계 — 원본 올리기
# ══════════════════════════════════════════════════
if step == 1:
    st.markdown("### 1단계 · 원본 IM 올리기")
    st.caption("받은 IM 을 그대로 올리세요. **워드(.docx)와 PDF 둘 다** 됩니다. "
               "세로 문서는 세로 그대로, 가로 문서는 가로 그대로 나옵니다.")

    c1, c2 = st.columns(2)
    c1.success("**그대로 두는 것**\n\n- 본문 내용·표·그림\n- 글씨 크기\n- 쪽 방향(세로/가로)")
    c2.info("**회사 양식으로 바꾸는 것**\n\n- 표지\n- 위쪽 머리말\n- 아래쪽 로고(증권사 로고 자리)\n"
            "- 글꼴\n- 맨 끝 연락처 페이지\n\n**더하는 것** : 사모사채 개요")

    with st.expander("📎 워드와 PDF 중 무엇을 올릴까요?"):
        st.markdown(
            "- **워드 원본이 있으면 워드를 올려 주세요.** 가장 깔끔하게 나옵니다.\n"
            "- **PDF 밖에 없으면 PDF 를 올리면 됩니다.** 도구가 먼저 워드로 바꾼 뒤 양식을 입힙니다. "
            "이때 표나 줄바꿈이 원본과 조금 다를 수 있어, 받은 뒤 한 번 훑어봐 주세요.\n"
            "- 글자를 마우스로 드래그해도 선택되지 않는 PDF(스캔본)는 바꿀 수 없습니다.\n"
            "- 옛날 워드 형식(.doc)은 워드에서 열어 **F12 → 파일 형식 'Word 문서(*.docx)'** 로 "
            "저장한 뒤 올려 주세요.")

    up = st.file_uploader("원본 IM (워드 또는 PDF)", type=["docx", "pdf"], key="cf_up")
    if up is not None and st.session_state.get("cf_name") != up.name:
        st.session_state["cf_name"] = up.name
        st.session_state["cf_bytes"] = up.getvalue()
        for k in ("cf_err", "cf_repair"):
            st.session_state.pop(k, None)
        try:
            st.session_state["cf_info"] = _read_source(up.name, up.getvalue())
        except zipfile.BadZipFile:
            # ★실제로 있었다 — 부산 남포동 워드 원본은 파일 끝이 잘려 있어 못 열었다.
            #   먼저 스스로 고쳐 보고(docx_repair), 그래도 안 될 때만 사람에게 부탁한다.
            from docx_repair import repair_docx, missing_summary
            fixed, missing = repair_docx(up.getvalue())
            info = None
            if fixed is not None:
                try:
                    info = _read_source(up.name, fixed)
                except Exception:
                    info = None
            if info is not None:
                st.session_state["cf_bytes"] = fixed          # 이후 단계는 고친 파일로 진행
                st.session_state["cf_info"] = info
                st.session_state["cf_repair"] = missing_summary(missing)
            else:
                st.session_state["cf_info"] = None
                st.session_state["cf_err"] = "broken"
        except Exception as e:
            st.session_state["cf_info"] = None
            st.session_state["cf_err"] = f"파일을 열지 못했습니다 : {e}"

    err = st.session_state.get("cf_err")
    if err == "broken":
        st.error(
            "**이 워드 파일은 깨져 있어서 열 수 없습니다.** 자동으로 고쳐 보았지만 본문까지 "
            "잘려 나가 되살리지 못했습니다.\n\n"
            "**왜 이런가요?** 워드 파일(.docx)은 본문·서식·그림 같은 여러 조각을 하나로 묶은 "
            "압축 파일이고, 맨 끝에 '어느 조각이 어디 있는지' 적힌 **목차**가 붙어 있습니다. "
            "메일로 주고받거나 내려받는 도중에 **파일 끝부분이 잘리면** 이 목차가 사라져 "
            "프로그램이 파일을 열지 못합니다. 워드 프로그램은 이런 파일도 스스로 고쳐서 열어 주기 때문에, "
            "내 컴퓨터에서는 멀쩡해 보일 수 있습니다.\n\n"
            "**이렇게 해 주세요**\n"
            "1. 이 파일을 **워드에서 열어** 주세요. '복구할까요?' 라고 물으면 **예**를 누릅니다.\n"
            "2. **F12** → 파일 형식 **'Word 문서(*.docx)'** → 저장.\n"
            "3. 새로 저장한 파일을 여기에 다시 올려 주세요.\n\n"
            "같은 IM 의 **PDF** 가 있으면 PDF 를 올리는 것이 더 정확합니다. 보낸 쪽에 원본을 "
            "다시 받을 수 있으면 그게 가장 좋습니다.")
    elif err:
        st.error(err)

    repair = st.session_state.get("cf_repair")
    if repair is not None:
        if repair:
            st.warning(
                "🛠️ **깨진 워드 파일이라 자동으로 고쳐서 열었습니다.** 다만 **되살리지 못한 부분**이 있습니다 : "
                f"**{repair}**\n\n"
                "**왜 이런가요?** 이 파일은 받는 도중에 **파일 끝부분이 잘려 나간** 상태입니다. 워드 파일은 "
                "본문·서식·그림 같은 조각을 묶어 둔 것인데, 잘려 나간 뒤쪽에 있던 조각은 파일 안에 **아예 "
                "남아 있지 않아** 어떤 프로그램으로도 되살릴 수 없습니다(워드에서 다시 저장해도 마찬가지입니다).\n\n"
                "**그래서 결과물이 원본과 이렇게 다를 수 있습니다**\n"
                "- 서식이 빠지면 글머리표(✓ 등)가 번호로 바뀌거나 줄 간격이 넓어져 쪽이 밀립니다.\n"
                "- 빠진 그림은 그 자리가 비어 나옵니다.\n\n"
                "**더 정확하게 하려면** 같은 IM 의 **PDF** 를 올리거나, 보낸 쪽에 **원본 워드를 다시 요청**해 주세요. "
                "이대로 계속 진행해도 됩니다.")
        else:
            st.info("🛠️ 깨진 워드 파일이라 자동으로 고쳐서 열었습니다. 빠진 내용은 없습니다.")

    info = st.session_state.get("cf_info")
    if info:
        st.success(f"올린 파일 — **{st.session_state['cf_name']}**")
        m1, m2, m3 = st.columns(3)
        m1.metric("종류", info["kind"])
        m2.metric("쪽 방향", info["orient"])
        if info["pages"] is not None:
            m3.metric("쪽수", f"{info['pages']}쪽")
        else:
            m3.metric("표 개수", f"{info.get('tables', 0)}개")
        if info["kind"] == "PDF":
            st.caption("PDF 는 3단계에서 먼저 워드로 바꾼 뒤 양식을 입힙니다.")

        st.markdown("---")
        _b1, _b2, _b3 = st.columns([2, 4, 2])
        if _b3.button("다음 단계 →", type="primary", use_container_width=True):
            _goto(2)


# ══════════════════════════════════════════════════
# 2단계 — 사모사채 개요
# ══════════════════════════════════════════════════
elif step == 2:
    st.markdown("### 2단계 · 사모사채 개요")
    st.caption("결과 문서에 들어갈 **사모사채 개요** 표입니다. 왼쪽 구분은 항상 같고, "
               "오른쪽 내용만 이번 딜에 맞게 적어 주세요. 비워 둔 칸은 빈칸으로 들어갑니다.")
    st.info("🛠️ 원본에서 내용을 자동으로 채워 넣는 기능은 회사 양식을 받은 뒤 붙일 예정입니다. "
            "지금은 직접 적어 주세요.")

    for row in BOND_ROWS:
        k1, k2 = st.columns([1, 4])
        k1.markdown(f"<div style='padding-top:8px;font-weight:700;color:#08377C;'>{row}</div>",
                    unsafe_allow_html=True)
        k2.text_input(row, key=f"cf_bond_{row}", label_visibility="collapsed")

    st.markdown("---")
    p1, _sp, p3 = st.columns([2, 4, 2])
    if p1.button("← 이전 단계", use_container_width=True):
        _goto(1)
    if p3.button("다음 단계 →", type="primary", use_container_width=True):
        _goto(3)


# ══════════════════════════════════════════════════
# 3단계 — 만들기
# ══════════════════════════════════════════════════
elif step == 3:
    st.markdown("### 3단계 · 만들기")
    if not st.session_state.get("cf_bytes"):
        st.warning("먼저 1단계에서 원본을 올려 주세요.")
    else:
        info = st.session_state.get("cf_info") or {}
        st.write(f"**{st.session_state.get('cf_name', '')}** · {info.get('kind', '')} · "
                 f"{info.get('orient', '')}")
        filled = [r for r in BOND_ROWS if (st.session_state.get(f"cf_bond_{r}") or "").strip()]
        st.write(f"사모사채 개요 : {len(filled)}/{len(BOND_ROWS)}칸 채움")

        _base = os.path.splitext(st.session_state.get("cf_name") or "IM")[0]
        st.session_state.setdefault("cf_fname", f"{_base}_회사양식")
        f1, f2 = st.columns([5, 1])
        f1.text_input("파일명", key="cf_fname",
                      help="내려받을 파일 이름입니다. 확장자(.docx)는 자동으로 붙습니다.")
        f2.markdown("<div style='padding-top:34px;color:#5b6b85;'>.docx</div>",
                    unsafe_allow_html=True)

        st.button("📝 회사 양식 워드 만들기", type="primary", disabled=True)
        st.warning("**준비 중** — 회사 양식 워드를 받으면 여기서 바로 만들어지게 연결합니다.")

    st.markdown("---")
    p1, _sp, _p3 = st.columns([2, 4, 2])
    if p1.button("← 이전 단계", use_container_width=True):
        _goto(2)
