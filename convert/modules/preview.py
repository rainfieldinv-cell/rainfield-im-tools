"""7단계 '전체 미리보기' — 생성 PPT를 슬라이드별 이미지로 변환(LibreOffice → PDF → PNG).

★읽기 전용: 본문 생성/검수 로직과 무관. 변환 실패 시 예외를 잡아 (이미지없음, 에러메시지)를
  돌려주어 앱이 절대 죽지 않게 한다(Streamlit Cloud=리눅스 기준).

서버 의존성(packages.txt): libreoffice, fonts-nanum(한글), poppler-utils(pdf2image용).
파이썬 의존성(requirements.txt): pdf2image.
"""
import io
import os
import shutil
import tempfile
import subprocess


def _find_soffice():
    """LibreOffice 실행 파일 경로 탐색(리눅스 우선, 로컬 대비 폴백)."""
    for cand in ("soffice", "libreoffice",
                 "/usr/bin/soffice", "/usr/bin/libreoffice",
                 "/usr/lib/libreoffice/program/soffice"):
        p = shutil.which(cand) if os.sep not in cand else (cand if os.path.exists(cand) else None)
        if p:
            return p
    return None


def ppt_to_images(ppt_bytes: bytes, max_pages=None, dpi: int = 120):
    """PPT bytes → (images: list[bytes(PNG)], error: str|None).

    max_pages: 앞 N페이지만 변환(테스트 모드). None이면 전체.
    실패하면 ([], '사람이 읽을 수 있는 원인') 을 반환 — 호출부에서 메시지만 표시하면 됨.
    """
    if not ppt_bytes:
        return [], "변환할 PPT가 없습니다. 먼저 4단계에서 PPT를 생성하세요."

    soffice = _find_soffice()
    if not soffice:
        return [], ("LibreOffice(soffice)가 설치되어 있지 않습니다. "
                    "packages.txt에 'libreoffice' 추가 후 재배포가 필요합니다.")

    tmp = tempfile.mkdtemp(prefix="rf_preview_")
    try:
        pptx_path = os.path.join(tmp, "deck.pptx")
        with open(pptx_path, "wb") as f:
            f.write(ppt_bytes)

        # 1) PPT → PDF (LibreOffice headless). HOME을 임시폴더로 줘야 프로필 생성 충돌이 없음.
        env = dict(os.environ)
        env["HOME"] = tmp
        try:
            proc = subprocess.run(
                [soffice, "--headless", "--norestore", "--nofirststartwizard",
                 "--convert-to", "pdf", "--outdir", tmp, pptx_path],
                capture_output=True, timeout=240, env=env,
            )
        except subprocess.TimeoutExpired:
            return [], "LibreOffice 변환 시간 초과(240초). 페이지 수가 많으면 '앞 3페이지만'으로 먼저 시도하세요."
        except Exception as e:
            return [], f"LibreOffice 실행 실패: {e}"

        pdf_path = os.path.join(tmp, "deck.pdf")
        if not os.path.exists(pdf_path):
            err = (proc.stderr or b"").decode("utf-8", "ignore").strip()
            return [], f"PDF 변환 실패: {err[:300] or 'soffice가 PDF를 생성하지 못했습니다.'}"

        # 2) PDF → PNG (pdf2image + poppler)
        try:
            from pdf2image import convert_from_path
        except Exception as e:
            return [], f"pdf2image 임포트 실패(requirements.txt 확인): {e}"

        last = max_pages if (isinstance(max_pages, int) and max_pages > 0) else None
        try:
            pages = convert_from_path(pdf_path, dpi=dpi, first_page=1, last_page=last)
        except Exception as e:
            return [], f"PDF→이미지 변환 실패(poppler-utils 필요): {e}"

        images = []
        for img in pages:
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            images.append(buf.getvalue())

        if not images:
            return [], "변환된 이미지가 없습니다(빈 PDF)."
        return images, None

    except Exception as e:                                # 어떤 예외든 앱은 살린다
        return [], f"미리보기 생성 중 오류: {e}"
    finally:
        shutil.rmtree(tmp, ignore_errors=True)


# ─────────────────────────────────────────────
# [원본 PDF 페이지 렌더] — ppt_to_images와 동일하게 pdf2image 사용(원본 목차 이미지 확인용).
#   ★기존 ppt_to_images는 건드리지 않고, 원본 PDF 페이지 렌더 전용 함수를 별도 추가.
# ─────────────────────────────────────────────
def pdf_pages_to_images(pdf_bytes: bytes, pages=None, dpi: int = 120):
    """원본 PDF의 지정 페이지를 PNG 이미지로 변환 → (images: list[bytes], error: str|None).

    pages: 1-based 페이지 번호 리스트(None이면 전체). 실패해도 앱이 죽지 않게 graceful 처리.
    """
    if not pdf_bytes:
        return [], "원본 PDF가 없습니다. 1단계에서 PDF를 함께 올려주세요."
    try:
        from pdf2image import convert_from_bytes
    except Exception as e:
        return [], f"pdf2image 임포트 실패(requirements.txt 확인): {e}"

    images = []
    try:
        if not pages:
            for im in convert_from_bytes(pdf_bytes, dpi=dpi):
                buf = io.BytesIO()
                im.save(buf, format="PNG")
                images.append(buf.getvalue())
        else:
            for p in pages:
                for im in convert_from_bytes(pdf_bytes, dpi=dpi, first_page=p, last_page=p):
                    buf = io.BytesIO()
                    im.save(buf, format="PNG")
                    images.append(buf.getvalue())
    except Exception as e:
        return [], f"PDF→이미지 변환 실패(poppler-utils 필요): {e}"

    if not images:
        return [], "변환된 이미지가 없습니다(페이지 번호를 확인하세요)."
    return images, None


# ─────────────────────────────────────────────
# [워드 → PDF 변환] — 자금판(전)처럼 워드를 LibreOffice로 PDF로 바꿔 페이지·레이아웃·이미지를 살린다.
#   (python-docx 직독은 '1페이지'로 뭉치고 좌표가 없어 IM 재현이 안 됨 → PDF 변환이 필요)
# ─────────────────────────────────────────────
def _salvage_docx_zip(word_bytes: bytes):
    """손상된 docx(중앙 디렉토리/EOCD가 없거나 깨진 zip) 복구 시도.
       파일 안에 남아있는 로컬 파일 헤더(PK0304)들을 순서대로 읽어 새 zip으로 재조립한다.
       반환: (새 docx bytes|None, 살린 항목 수). 본문(word/document.xml)이 없으면 실패로 본다."""
    import struct, zlib
    try:
        import zipfile as _zf
    except Exception:
        return None, 0
    data = word_bytes
    entries = []          # (name, raw_uncompressed_bytes)
    i = 0
    while True:
        j = data.find(b"PK\x03\x04", i)
        if j < 0 or j + 30 > len(data):
            break
        try:
            (ver, flag, method, mt, md, crc,
             csize, usize, fnl, efl) = struct.unpack("<HHHHHIIIHH", data[j + 4:j + 30])
        except struct.error:
            break
        name_end = j + 30 + fnl
        body = name_end + efl
        # 데이터 디스크립터(크기가 헤더에 없음)나 크기 이상 → 재조립 불가 항목, 중단
        if (flag & 0x08) or csize == 0 and usize == 0 and method != 0:
            break
        comp = data[body:body + csize]
        if len(comp) < csize:                       # 파일이 여기서 잘림
            break
        name = data[j + 30:name_end].decode("utf-8", "replace")
        try:
            raw = comp if method == 0 else zlib.decompress(comp, -15)
            entries.append((name, raw))
        except Exception:
            pass                                    # 이 조각만 버리고 계속
        i = body + csize

    if not entries or not any(n == "word/document.xml" for n, _ in entries):
        return None, 0

    import io as _io
    out = _io.BytesIO()
    seen = set()
    with _zf.ZipFile(out, "w", _zf.ZIP_DEFLATED) as z:
        for name, raw in entries:
            if name in seen:
                continue
            seen.add(name)
            z.writestr(name, raw)
    return out.getvalue(), len(seen)


def convert_word_to_pdf(word_bytes: bytes, filename: str = "doc.docx"):
    """워드(.doc/.docx) → PDF bytes. 반환: (pdf_bytes: bytes|None, error: str|None).
       LibreOffice(soffice) 사용. 실패해도 앱이 죽지 않게 graceful 처리.
       soffice가 파일을 못 열면(손상 docx) zip 재조립으로 한 번 더 시도한다."""
    if not word_bytes:
        return None, "워드 파일이 없습니다."
    soffice = _find_soffice()
    if not soffice:
        return None, "LibreOffice(soffice)가 없어 워드→PDF 변환을 못 합니다(로컬 실행 시 LibreOffice 필요)."

    ext = ".doc" if filename.lower().endswith(".doc") else ".docx"

    def _run(src_bytes):
        """단일 변환 시도 → (pdf_bytes|None, reason|None)."""
        tmp = tempfile.mkdtemp(prefix="rf_word_")
        try:
            src = os.path.join(tmp, "src" + ext)
            with open(src, "wb") as f:
                f.write(src_bytes)
            env = dict(os.environ)
            env["HOME"] = tmp
            # ★호출마다 '전용 프로필' 강제 → 이전 변환의 프로필 잠금에 의한 무음 실패 차단.
            from pathlib import Path
            user_inst = Path(tmp, "lo_profile").as_uri()
            try:
                proc = subprocess.run(
                    [soffice, "-env:UserInstallation=" + user_inst,
                     "--headless", "--norestore", "--nofirststartwizard",
                     "--convert-to", "pdf", "--outdir", tmp, src],
                    capture_output=True, timeout=240, env=env,
                )
            except subprocess.TimeoutExpired:
                return None, "워드→PDF 변환 시간 초과(240초). 파일이 크거나 이미지가 많으면 발생합니다."
            except Exception as e:
                return None, f"LibreOffice 실행 실패: {e}"

            pdf_path = os.path.join(tmp, "src.pdf")
            if not os.path.exists(pdf_path):
                err = (proc.stderr or b"").decode("utf-8", "ignore").strip()
                out = (proc.stdout or b"").decode("utf-8", "ignore").strip()
                reason = err or out or "soffice가 PDF를 생성하지 못했습니다."
                return None, reason[:300]
            with open(pdf_path, "rb") as f:
                return f.read(), None
        except Exception as e:
            return None, f"워드→PDF 변환 중 오류: {e}"
        finally:
            shutil.rmtree(tmp, ignore_errors=True)

    # 1차 시도(원본 그대로)
    pdf, reason = _run(word_bytes)
    if pdf is not None:
        return pdf, None

    # 2차 시도: docx가 손상돼 못 열린 경우 zip 재조립본으로 재시도
    if ext == ".docx":
        salvaged, n = _salvage_docx_zip(word_bytes)
        if salvaged is not None:
            pdf2, reason2 = _run(salvaged)
            if pdf2 is not None:
                return pdf2, None
            return None, (f"워드→PDF 변환 실패 — {reason}. "
                          f"손상 복구({n}개 조각)로 재시도했으나 실패: {reason2}")

    return None, f"워드→PDF 변환 실패 — {reason}"
