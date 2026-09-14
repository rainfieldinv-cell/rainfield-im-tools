# -*- coding: utf-8 -*-
"""워드(.doc/.docx) → PDF. LibreOffice 를 불러서 바꾼다.

★왜 PDF 로 바꾸나
  이 도구들은 전부 PDF 를 읽게 만들어져 있다(글자 하나하나의 자리를 재야 해서).
  워드는 그 자리 정보가 없으므로, 먼저 PDF 로 바꾼 뒤 읽는다.

★그래서 워드는 **권하지 않는다.**
  LibreOffice 가 워드를 PDF 로 바꿀 때 줄바꿈·표 높이·글꼴이 원본 워드와 조금 달라진다.
  PDF 원본이 있으면 그걸 쓰는 편이 훨씬 정확하다.

★웹(리눅스)에서는 packages.txt 에 `libreoffice-writer` 가 있어야 돌아간다.
  없으면 조용히 실패하지 않고 **이유를 돌려준다** — 앱이 죽으면 안 되기 때문이다.
"""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

_후보 = ("soffice", "libreoffice",
         "/usr/bin/soffice", "/usr/bin/libreoffice",
         "/usr/lib/libreoffice/program/soffice",
         r"C:\Program Files\LibreOffice\program\soffice.exe")


def 찾기():
    """LibreOffice 실행 파일을 찾는다. 없으면 None."""
    for c in _후보:
        p = shutil.which(c) if os.sep not in c else (c if os.path.exists(c) else None)
        if p:
            return p
    return None


def 있나():
    return 찾기() is not None


def 바꾸기(word_bytes, filename="doc.docx", timeout=240):
    """워드 바이트 → (pdf_bytes | None, 안 된 이유 | None)."""
    if not word_bytes:
        return None, "워드 파일이 비어 있습니다."
    soffice = 찾기()
    if not soffice:
        return None, ("이 서버에 LibreOffice 가 없어 워드를 PDF 로 바꾸지 못했습니다. "
                      "PDF 로 올려 주세요.")

    ext = ".doc" if filename.lower().endswith(".doc") else ".docx"
    tmp = tempfile.mkdtemp(prefix="rf_word_")
    try:
        src = os.path.join(tmp, "src" + ext)
        with open(src, "wb") as f:
            f.write(word_bytes)
        env = dict(os.environ)
        env["HOME"] = tmp                     # 리눅스에서 프로필 쓸 곳
        # ★부를 때마다 전용 프로필을 쓴다. 안 그러면 앞 변환이 남긴 잠금 때문에
        #   아무 말 없이 실패한다(원본 convert 도구에서 겪은 문제).
        prof = Path(tmp, "lo_profile").as_uri()
        try:
            proc = subprocess.run(
                [soffice, "-env:UserInstallation=" + prof,
                 "--headless", "--norestore", "--nofirststartwizard",
                 "--convert-to", "pdf", "--outdir", tmp, src],
                capture_output=True, timeout=timeout, env=env)
        except subprocess.TimeoutExpired:
            return None, f"워드→PDF 변환이 {timeout}초를 넘겼습니다. 파일이 크면 생깁니다."
        except Exception as e:
            return None, f"LibreOffice 를 실행하지 못했습니다: {e}"

        pdf = os.path.join(tmp, "src.pdf")
        if not os.path.exists(pdf):
            why = ((proc.stderr or b"").decode("utf-8", "ignore").strip()
                   or (proc.stdout or b"").decode("utf-8", "ignore").strip()
                   or "LibreOffice 가 PDF 를 만들지 못했습니다.")
            return None, why[:300]
        with open(pdf, "rb") as f:
            return f.read(), None
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# 화면에 한 줄로 붙일 안내 — 세 도구가 같은 문구를 쓴다
안내 = ("※ 워드(.docx)도 올릴 수 있지만 **PDF 를 권합니다** — "
        "워드는 PDF 로 바꾸는 과정에서 줄바꿈·표 모양이 조금 달라질 수 있습니다.")
