# -*- coding: utf-8 -*-
"""레인필드 IM 도구 — 처음 화면.

IM 관련 도구 두 개를 한 주소에서 쓴다. (계약서 도구와 같은 방식)
  📑 IM 변환기  → convert/app.py       원본 IM(PDF) → 회사 양식 제안서 전체
  📄 IM 요약본  → summary/summary_app.py  원본 IM(PDF) → 4~5장 요약본

★두 도구의 코드는 각자 폴더에 **그대로** 둔다. 합치면서 코드를 뜯어고치면
  잘 돌던 것이 깨지기 쉬워서다. 대신 고를 때마다 그 폴더만 파이썬 경로에
  올려서 실행한다.
"""
import os
import runpy
import sys

import streamlit as st

HERE = os.path.dirname(os.path.abspath(__file__))

st.set_page_config(page_title="Rainfield IM 도구", layout="wide")

st.markdown(
    """
    <style>
    .home-hero { text-align:center; padding: 6px 0 2px; }
    .home-hero h1 { margin-bottom: 6px; }
    .home-hero p { color:#666; font-size:15px; margin-top:0; }
    .tool-card { text-align:center; padding: 6px 4px 2px; }
    .tool-card .emoji { font-size: 46px; line-height: 1.1; }
    .tool-card .name { font-size: 20px; font-weight: 700; color:#1A2B5E; margin-top: 6px; }
    .tool-card .desc { color:#555; font-size: 14px; margin: 10px 0 4px; min-height: 66px; }
    .tool-card .need { color:#888; font-size: 13px; margin-bottom: 10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

TOOLS = {
    "convert": {
        "emoji": "📑",
        "name": "IM 변환기",
        "desc": "원본 IM(PDF)을 읽어 <b>회사 양식 제안서 한 벌</b>로 만듭니다. "
                "표지부터 연락처까지 20장 안팎으로 나옵니다.",
        "need": "필요한 파일 : 원본 IM(PDF) 1개",
        "dir": "convert",
        "file": "app.py",
    },
    "summary": {
        "emoji": "📄",
        "name": "IM 요약본",
        "desc": "원본 IM(PDF)을 <b>4~5장짜리 요약본</b>으로 줄입니다. "
                "먼저 보고할 때 쓰고, 금융구조도도 함께 만들 수 있습니다.",
        "need": "필요한 파일 : 원본 IM(PDF) 1개",
        "dir": "summary",
        "file": "summary_app.py",
    },
}

# 두 도구에 이름이 같은 파일이 있다(extractors.py 는 같지만 claude_api.py 는 다르다).
# 도구를 바꿀 때 앞서 불러둔 것이 남아 있으면 **엉뚱한 쪽이 불린다** → 지우고 다시 읽는다.
_SHARED_NAMES = ("claude_api", "extractors", "engine_bits", "diagram",
                 "summary_pipeline", "ui_components", "modules")


def open_tool(key: str):
    st.session_state["tool"] = key


def go_home():
    st.session_state["tool"] = None


def render_home():
    """처음 화면 — 카드 두 개 중 하나를 고릅니다."""
    st.markdown(
        '<div class="home-hero"><h1>레인필드 IM 도구</h1>'
        "<p>무엇을 할지 아래에서 골라주세요.</p></div>",
        unsafe_allow_html=True,
    )
    st.write("")

    _l, mid, _r = st.columns([0.5, 3, 0.5])
    with mid:
        cols = st.columns(2, gap="large")
        for col, (key, t) in zip(cols, TOOLS.items()):
            with col:
                with st.container(border=True):
                    st.markdown(
                        f'<div class="tool-card">'
                        f'<div class="emoji">{t["emoji"]}</div>'
                        f'<div class="name">{t["name"]}</div>'
                        f'<div class="desc">{t["desc"]}</div>'
                        f'<div class="need">{t["need"]}</div>'
                        f"</div>",
                        unsafe_allow_html=True,
                    )
                    st.button(
                        f'{t["emoji"]}  {t["name"]} 시작하기',
                        key=f"open_{key}", type="primary",
                        use_container_width=True,
                        on_click=open_tool, args=(key,),
                    )

        st.write("")
        st.caption(
            "⚠️ 두 기능 모두 **원본을 읽어 초안을 만들어 주는 것**입니다. "
            "숫자와 문구는 반드시 담당자가 원본과 대조해 확인하세요."
        )


def _prepare_path(folder: str):
    """고른 도구의 폴더만 파이썬 경로 맨 앞에 둔다."""
    for t in TOOLS.values():
        p = os.path.join(HERE, t["dir"])
        while p in sys.path:
            sys.path.remove(p)
    sys.path.insert(0, os.path.join(HERE, folder))


def render_tool(key: str):
    """고른 도구 화면을 그립니다. 맨 위에 처음으로 돌아가는 버튼."""
    back, title = st.columns([0.16, 0.84], vertical_alignment="center")
    back.button("◀ 처음 화면", key="go_home", on_click=go_home,
                use_container_width=True)
    other = "summary" if key == "convert" else "convert"
    title.button(
        f'{TOOLS[other]["emoji"]}  {TOOLS[other]["name"]} 로 바로 가기',
        key=f"switch_{other}", on_click=open_tool, args=(other,),
    )
    st.divider()

    t = TOOLS[key]
    if st.session_state.get("_loaded_tool") != key:
        for name in list(sys.modules):
            if name in _SHARED_NAMES or name.split(".")[0] in _SHARED_NAMES:
                del sys.modules[name]
        st.session_state["_loaded_tool"] = key
    _prepare_path(t["dir"])

    path = os.path.join(HERE, t["dir"], t["file"])
    cwd = os.getcwd()
    try:
        os.chdir(os.path.join(HERE, t["dir"]))   # 도구가 자기 폴더 기준으로 파일을 찾는다
        runpy.run_path(path, run_name="__main__")
    finally:
        os.chdir(cwd)


tool = st.session_state.get("tool")
if tool in TOOLS:
    render_tool(tool)
else:
    render_home()
