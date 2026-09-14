# -*- coding: utf-8 -*-
"""IM Decalcomania — 원본 IM(PDF) 을 생긴 그대로 PPT 로. 단계별 진행.

지금은 두 단계다. 나머지는 이 사이에 붙일 예정이다.
    1 원본 업로드(올리면 바로 PPT 로 변환)
      →  (하이라이트)  →  (금융구조도)  →  (내용 배치)  →  2 생성

★올리는 즉시 변환한다. 예전엔 '변환' 단계를 따로 두고 버튼을 누르게 했는데,
  한 번 더 누르는 것뿐이라 번거로웠다.
★결과 미리보기(파워포인트로 슬라이드를 그림으로 뽑기)는 뺐다. 스트림릿은 화면을
  딴 갈래에서 그리는데 파워포인트를 부르려면 그 갈래에서 CoInitialize 를 먼저 해야 해
  'CoInitialize가 호출되지 않았습니다' 로 실패했고, 웹(리눅스)에는 파워포인트가 아예 없다.
  → 결과는 내려받아 파워포인트에서 직접 확인한다.

혼자 돌릴 때 :  streamlit run app.py --server.port 8610
"""
import importlib
import os
import re
import tempfile
import time

import streamlit as st

import pdf_to_ppt

HERE = os.path.dirname(os.path.abspath(__file__))
STEP_NAMES = ["원본 업로드", "생성"]

# ★스트림릿은 app.py 만 다시 읽는다. 옆 파일(pdf_to_ppt.py)을 고쳐도 이미 읽어 둔 것을
#   계속 써서 **고친 내용이 반영되지 않는다.** 실제로 이것 때문에, 표로 바꾼 뒤에도
#   화면에서는 옛 방식(도형+글상자) 결과가 계속 나왔다.
#   → 엔진 파일이 바뀌었으면 다시 읽는다.
#   ★기준값을 **세션이 아니라 엔진 자체에** 적어 둔다. 세션에 두면 브라우저 창마다
#     따로 놀아서, 어떤 창은 옛 엔진으로 계속 돈다(그래서 고친 게 반영이 안 됐다).
_ENGINE = os.path.join(HERE, "pdf_to_ppt.py")
ENGINE_TIME = "?"
try:
    _mt = os.path.getmtime(_ENGINE)
    if getattr(pdf_to_ppt, "__loaded_at__", None) != _mt:
        importlib.reload(pdf_to_ppt)
        pdf_to_ppt.__loaded_at__ = _mt
    ENGINE_TIME = time.strftime("%m-%d %H:%M", time.localtime(_mt))
except Exception:
    pass

try:      # 나중에 'IM 도구' 첫 화면에 붙일 때를 위해 — 두 번 부르면 안 된다
    st.set_page_config(page_title="IM Decalcomania", page_icon="📄", layout="wide")
except Exception:
    pass


# ── 단계 표시줄 (요약본·변환기와 같은 모양) ──────────
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
    st.session_state["step"] = n
    st.rerun()


st.session_state.setdefault("step", 1)
step = st.session_state["step"]

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
div.stDownloadButton>button:hover{background:#0063A1;border-color:#0063A1;}
hr{border-color:#dfe6f0;}
</style>
""", unsafe_allow_html=True)

st.markdown(
    "<h1 style='margin:0 0 .1rem 0;line-height:1.1;'>IM"
    "<span style='font-size:.34em;font-weight:600;color:#08377C;"
    "letter-spacing:.16em;margin-left:.55rem;vertical-align:.35em;'>"
    "DECALCOMANIA</span></h1>", unsafe_allow_html=True)
# ★엔진이 언제 고친 것인지 화면에 찍는다. '고쳤다는데 그대로 나온다' 싶으면
#   이 시각이 최신인지 먼저 보면 된다(서버가 옛 엔진을 물고 있으면 여기서 드러난다).
st.caption(f"원본 IM(PDF) 을 **생긴 그대로** 찍어 고칠 수 있는 PPT 로　·　엔진 {ENGINE_TIME}")
render_stepper(step)


# ══════════════════════════════════════════════════
# 1단계 — 원본 업로드 (올리면 바로 변환)
# ══════════════════════════════════════════════════
if step == 1:
    st.markdown("### 1단계 · 원본 IM 업로드")
    st.caption("PDF 를 올리면 **바로 PPT 로 바뀝니다.** 따로 누를 것은 없습니다. "
               "표의 칸 색·선 색·선 굵기와 자리는 원본을 그대로 따르고, "
               "글꼴만 회사 글꼴로 바꿉니다.")
    st.info(f"글꼴 **{pdf_to_ppt.FONT_BODY} / {pdf_to_ppt.FONT_BOLD}** · "
            f"글자 크기 **{pdf_to_ppt.FONT_SIZE:g}pt 고정** · **전체 쪽**을 옮깁니다.")

    up = st.file_uploader("원본 IM PDF", type=["pdf"])

    def _convert_now():
        """지금 들고 있는 PDF 를 변환해 결과를 담아 둔다."""
        for k in ("ppt", "info", "err", "fname"):
            st.session_state.pop(k, None)
        bar = st.progress(0.0, text="원본을 읽는 중...")
        out = tempfile.NamedTemporaryFile(suffix=".pptx", delete=False).name
        try:
            info = pdf_to_ppt.convert(
                st.session_state["pdf_bytes"], out,
                progress=lambda i, n, p: bar.progress(
                    i / n, text=f"{p}쪽 옮기는 중... ({i}/{n})"))
            with open(out, "rb") as f:
                st.session_state["ppt"] = f.read()
            st.session_state["info"] = info
        except Exception:
            import traceback
            st.session_state["err"] = traceback.format_exc()
        bar.empty()

    # ★올리는 즉시 변환한다. 다시 그릴 때마다 또 돌지 않도록 파일 이름으로 가른다.
    if up is not None and st.session_state.get("pdf_name") != up.name:
        st.session_state["pdf_name"] = up.name
        st.session_state["pdf_bytes"] = up.getvalue()
        _convert_now()
        st.rerun()

    if st.session_state.get("err"):
        st.error("변환하지 못했습니다.")
        st.code(st.session_state["err"])

    info = st.session_state.get("info")
    if info:
        st.success(f"변환 완료 — **{st.session_state.get('pdf_name','')}** · "
                   f"{info['count']}쪽")
        # ★같은 파일은 다시 올려도 변환이 또 돌지 않는다(이름이 같으니까).
        #   도구를 고친 뒤에는 이 버튼을 눌러야 새 결과가 나온다.
        if st.button("🔁 다시 변환하기",
                     help="도구를 고쳤을 때 누르세요. 같은 파일을 새로 변환합니다."):
            _convert_now()
            st.rerun()
        c1, c2, c3 = st.columns(3)
        c1.metric("쪽수", f"{info['count']}쪽")
        c2.metric("슬라이드 크기",
                  f"{info['size'][0]:.2f} × {info['size'][1]:.2f} in")
        c3.metric("원본과 크기", "같음")
        with st.expander("쪽마다 무엇이 몇 개 들어갔는지 보기"):
            st.dataframe(
                [{"쪽": s["쪽"], "네모(표 선·칸 배경)": s["네모"],
                  "그림": s["그림"], "글상자": s["글상자"]}
                 for s in info["pages"]],
                hide_index=True, use_container_width=True)

        st.markdown("---")
        _b1, _b2, _b3 = st.columns([2, 4, 2])
        if _b3.button("다음 단계 →", type="primary", use_container_width=True):
            _goto(2)


# ══════════════════════════════════════════════════
# 2단계 — 생성(내려받기)
# ══════════════════════════════════════════════════
elif step == 2:
    st.markdown("### 2단계 · 생성")
    if not st.session_state.get("ppt"):
        st.warning("먼저 1단계에서 원본을 올려주세요.")
    else:
        info = st.session_state.get("info") or {}
        st.write(f"**{st.session_state.get('pdf_name','')}** · "
                 f"{info.get('count', '?')}쪽")

        _base = os.path.splitext(st.session_state.get("pdf_name") or "변환본")[0]
        st.session_state.setdefault("fname", f"{_base}_PPT변환")
        f1, f2 = st.columns([5, 1])
        f1.text_input("파일명", key="fname",
                      help="내려받을 파일 이름입니다. 확장자(.pptx)는 자동으로 붙습니다.")
        f2.markdown("<div style='padding-top:34px;color:#5b6b85;'>.pptx</div>",
                    unsafe_allow_html=True)

        nm = re.sub(r'[\\/:*?"<>|]', "_",
                    (st.session_state.get("fname") or "").strip())[:80]
        if nm.lower().endswith(".pptx"):
            nm = nm[:-5]
        st.download_button("다운로드", data=st.session_state["ppt"],
                           file_name=f"{nm or _base}.pptx",
                           mime="application/vnd.openxmlformats-officedocument."
                                "presentationml.presentation",
                           type="primary")
        st.caption("⚠️ 글꼴은 파일에 담기지 않습니다. 피플폰트가 깔린 컴퓨터에서 열어야 "
                   "원본과 같은 모습으로 보입니다.")

    st.markdown("---")
    p1, _sp, _p3 = st.columns([2, 4, 2])
    if p1.button("← 이전 단계", use_container_width=True):
        _goto(1)
