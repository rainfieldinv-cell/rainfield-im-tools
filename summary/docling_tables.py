# -*- coding: utf-8 -*-
"""Docling 으로 원본 IM 의 표를 읽는다 — **내 컴퓨터에서만** 쓰는 기능.

왜 웹에는 못 올리나
  설치 용량 1.3GB, 처음 돌 때 내려받는 모델 500MB, 돌 때 메모리 1.3GB 를 쓴다.
  스트림릿 무료 웹은 메모리 1GB 가 한도라 올리면 그 자리에서 죽는다.
  그래서 이 파일은 설치본이 있으면 켜지고, 없으면 조용히 꺼진다(웹에서는 항상 꺼짐).

한글 경로 주의 (여기서 한 번 막혔다)
  Docling 의 PDF 파서는 C++ 이라 경로에 한글이 있으면 자기 자료 파일을 못 연다.
      RuntimeError: filename does not exists: ...\\pdf_resources/glyphs//standard/additional.dat
  우리 폴더는 '종합\\자동화' 라 한글이 섞인다.
  → 설치본을 영문 이름(C:\\dlenv)으로 이어놓고, 읽을 PDF 도 영문 임시폴더에 복사해서 넘긴다.

지금 방식과의 차이
  · 지금(pdfplumber·PyMuPDF) : 표의 **선**을 보고 격자를 만든다. 선이 없거나 흐리면 놓친다.
  · Docling                  : 표 생김새를 학습한 모델이 칸을 잡는다. 대신 훨씬 느리다.
  글자는 둘 다 원문에서 그대로 가져온다(OCR 을 끄므로 한글이 깨지지 않는다).
"""
import os
import sys
import json
import hashlib
import subprocess

_HERE = os.path.dirname(os.path.abspath(__file__))
_INSTALL = os.path.join(os.path.dirname(_HERE), "_doclingenv")   # 실제 설치본
_ASCII = r"C:\dlenv"        # C++ 파서가 열 수 있게 영문 이름으로 이어 놓은 자리
_WORK = r"C:\dltmp"         # 읽을 PDF 를 잠깐 두는 곳(여기도 영문이어야 한다)
_WORKER = os.path.join(_HERE, "_docling_worker.py")


def _is_ascii(p):
    try:
        p.encode("ascii")
        return True
    except UnicodeEncodeError:
        return False


def env_path():
    """Docling 설치본을 **영문 경로**로 가리킨다. 없으면 None."""
    if not os.path.isdir(_INSTALL):
        return None
    if _is_ascii(_INSTALL):
        return _INSTALL
    if os.name != "nt":
        return _INSTALL
    if os.path.isdir(_ASCII):
        return _ASCII
    try:        # 폴더 바로가기(junction) — 관리자 권한 없이 만들 수 있다
        subprocess.run(["cmd", "/c", "mklink", "/J", _ASCII, _INSTALL],
                       capture_output=True, timeout=30)
    except Exception as e:
        print(f"[docling] 영문 경로 잇기 실패: {e}")
    return _ASCII if os.path.isdir(_ASCII) else None


def available():
    """이 컴퓨터에서 Docling 을 쓸 수 있나."""
    env = env_path()
    return bool(env and os.path.isdir(os.path.join(env, "docling")))


def why_not():
    """못 쓰는 이유 한 줄(화면에 그대로 보여준다)."""
    if not os.path.isdir(_INSTALL):
        return ("이 컴퓨터에는 Docling이 설치돼 있지 않습니다. "
                "웹에서는 용량·메모리 한도 때문에 쓸 수 없고, 내 컴퓨터에서만 됩니다.")
    if not available():
        return "Docling 설치본은 있는데 영문 경로로 잇지 못했습니다(경로에 한글)."
    return ""


# 읽는 방식을 고치면 이 숫자를 올린다(예전에 읽어둔 결과를 다시 쓰지 않도록).
_VERSION = 2


def _cache_path(pdf_bytes):
    h = hashlib.sha1(pdf_bytes + str(_VERSION).encode()).hexdigest()[:16]
    d = os.path.join(_WORK, "cache")
    os.makedirs(d, exist_ok=True)
    return os.path.join(d, f"{h}.json")


def read_tables(pdf_bytes, max_tables=80, timeout=1800, use_cache=True):
    """PDF 에서 표를 읽어 pdf_tables 와 **같은 모양**의 목록으로 돌려준다.

    반환: (표목록, 알림글)
    """
    env = env_path()
    if not env:
        return [], why_not()

    cache = _cache_path(pdf_bytes)
    if use_cache and os.path.exists(cache):
        try:
            with open(cache, encoding="utf-8") as f:
                return json.load(f)[:max_tables], "이미 읽어둔 결과를 씁니다."
        except Exception:
            pass

    os.makedirs(_WORK, exist_ok=True)
    src = os.path.join(_WORK, "in.pdf")          # ★영문 경로여야 파서가 연다
    with open(src, "wb") as f:
        f.write(pdf_bytes)
    out = os.path.join(_WORK, "out.json")
    if os.path.exists(out):
        os.remove(out)

    env_vars = dict(os.environ)
    # 설치본을 맨 앞에 둔다(Docling 이 들고 온 numpy·torch 가 먼저 잡히도록).
    env_vars["PYTHONPATH"] = os.pathsep.join(
        [env] + [p for p in (env_vars.get("PYTHONPATH") or "").split(os.pathsep) if p])
    env_vars["PYTHONIOENCODING"] = "utf-8"

    try:
        r = subprocess.run([sys.executable, _WORKER, src, out],
                           capture_output=True, text=True, encoding="utf-8",
                           errors="replace", env=env_vars, timeout=timeout)
    except subprocess.TimeoutExpired:
        return [], f"Docling이 {timeout // 60}분 안에 끝나지 않았습니다."

    if not os.path.exists(out):
        tail = (r.stderr or r.stdout or "").strip().splitlines()[-4:]
        return [], "Docling 실행 실패:\n" + "\n".join(tail)

    with open(out, encoding="utf-8") as f:
        tables = json.load(f)
    try:
        with open(cache, "w", encoding="utf-8") as f:
            json.dump(tables, f, ensure_ascii=False)
    except Exception:
        pass
    return tables[:max_tables], ""
