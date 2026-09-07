# -*- coding: utf-8 -*-
"""원본 IM(PDF)의 표를 **그대로** 읽어낸다.

요약하지 않는다. 원문 표의 행·열·병합·글자를 있는 그대로 가져와서
PPT 표로 다시 그릴 수 있는 형태로 준다.

왜 좌표를 직접 재는가
  pdfplumber 의 extract() 는 병합된 칸이 있는 줄에서 None 을 돌려준다.
  (돈암동 사업수지 표의 '비고' 열이 통째로 비어 나오던 원인)
  → 셀 사각형을 모아 격자를 만들고, 각 칸이 몇 행·몇 열을 먹는지 직접 센다.
"""
import re

import pdfplumber

# 선 조각 하나하나를 열 경계로 보면 9열짜리가 25열로 쪼개진다.
# 가까운 선끼리 하나로 합치도록 넉넉히 준다.
TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 6,
    "join_tolerance": 6,
    "intersection_tolerance": 6,
    "text_tolerance": 2,
}
_TOL = 3.0          # 같은 선으로 볼 좌표 오차(포인트)
_MIN_ROWS = 2
_MIN_COLS = 2


def _axis(values):
    """비슷한 좌표끼리 묶어 경계선 목록을 만든다."""
    out = []
    for v in sorted(values):
        if not out or abs(v - out[-1]) > _TOL:
            out.append(v)
    return out


def _idx(v, axis):
    best, bd = 0, 1e9
    for i, a in enumerate(axis):
        d = abs(a - v)
        if d < bd:
            best, bd = i, d
    return best


def _clean(s):
    if not s:
        return ""
    s = s.replace("\r", "")
    s = re.sub(r"[ \t]+", " ", s)
    return "\n".join(ln.strip() for ln in s.split("\n")).strip()


def read_table(page, tbl):
    """표 하나를 {rows, cols, cells:[{r,c,rs,cs,text}]} 로 만든다."""
    rects = [c for c in tbl.cells if c]
    if not rects:
        return None
    xs = _axis([v for c in rects for v in (c[0], c[2])])
    ys = _axis([v for c in rects for v in (c[1], c[3])])
    # 3pt 도 안 되는 칸은 선이 두 겹으로 그어진 것 → 진짜 행/열이 아니다
    xs = [v for i, v in enumerate(xs)
          if i == 0 or i == len(xs) - 1 or xs[i + 1] - v > 3.0]
    ys = [v for i, v in enumerate(ys)
          if i == 0 or i == len(ys) - 1 or ys[i + 1] - v > 3.0]
    n_col, n_row = len(xs) - 1, len(ys) - 1
    if n_row < _MIN_ROWS or n_col < _MIN_COLS:
        return None

    cells = []
    for (x0, y0, x1, y1) in rects:
        c0, c1 = _idx(x0, xs), _idx(x1, xs)
        r0, r1 = _idx(y0, ys), _idx(y1, ys)
        cs, rs = max(1, c1 - c0), max(1, r1 - r0)
        if c0 >= n_col or r0 >= n_row:
            continue
        # 칸 사각형을 직접 잘라 글자를 뽑는다(병합된 칸도 이러면 잡힌다)
        try:
            txt = page.crop((x0, y0, x1, y1), strict=False).extract_text()
        except Exception:
            txt = ""
        cells.append({"r": r0, "c": c0, "rs": min(rs, n_row - r0),
                      "cs": min(cs, n_col - c0), "text": _clean(txt)})
    if not cells:
        return None

    # ★칸 사각형이 아예 없는 자리(=선이 덜 그려진 열)도 글자를 뽑는다.
    #   돈암동 사업수지의 '비고' 열이 통째로 비어 나오던 원인이 이것이다.
    taken = set()
    for c in cells:
        for rr in range(c["r"], c["r"] + c["rs"]):
            for cc in range(c["c"], c["c"] + c["cs"]):
                taken.add((rr, cc))
    # ★한 칸씩 자르면 안 된다. 세로로 이어진 빈 자리는 **원래 하나의 병합 칸**이라,
    #   줄마다 자르면 같은 글자가 여러 번 복사된다(비중 '6.9%' 가 두 번 나오던 원인).
    #   → 이어진 구간을 통째로 한 칸으로 만든다. 그래야 병합이 그대로 살아난다.
    for c in range(n_col):
        r = 0
        while r < n_row:
            if (r, c) in taken:
                r += 1
                continue
            r2 = r
            while r2 + 1 < n_row and (r2 + 1, c) not in taken:
                r2 += 1
            try:
                txt = page.crop((xs[c], ys[r], xs[c + 1], ys[r2 + 1]),
                                strict=False).extract_text()
            except Exception:
                txt = ""
            cells.append({"r": r, "c": c, "rs": r2 - r + 1, "cs": 1,
                          "text": _clean(txt)})
            r = r2 + 1

    cells.sort(key=lambda x: (x["r"], x["c"]))
    # 열 폭·행 높이 비율 — PPT 표를 원본과 같은 모양으로 그리는 데 쓴다
    col_ratio = [round(xs[i + 1] - xs[i], 2) for i in range(n_col)]
    row_ratio = [round(ys[i + 1] - ys[i], 2) for i in range(n_row)]
    return {"rows": n_row, "cols": n_col, "cells": cells,
            "col_ratio": col_ratio, "row_ratio": row_ratio,
            "bbox": tuple(round(v, 1) for v in tbl.bbox)}


def _title_for(page, tbl, fallback):
    """표 바로 위/왼쪽에 있는 글자를 표 제목으로 삼는다."""
    x0, y0, x1, y1 = tbl.bbox
    words = page.extract_words()
    # ★이 회사 IM 은 표 이름을 **왼쪽 여백**에 적는다('사업일정', '예상 사업수지').
    #   위쪽부터 보면 바로 위 표의 마지막 줄을 제목으로 집는다 → 왼쪽을 먼저 본다.
    left = [w for w in words
            if x0 - 95 <= w["x1"] <= x0 + 2 and y0 - 6 <= w["top"] <= y1]
    best = " ".join(w["text"] for w in sorted(left, key=lambda w: (w["top"], w["x0"])))
    if not best:
        above = [w for w in words
                 if y0 - 26 <= w["bottom"] <= y0 + 2
                 and w["x1"] >= x0 - 6 and w["x0"] <= x1]
        best = " ".join(w["text"] for w in sorted(above, key=lambda w: w["x0"]))
    best = _clean(best)
    return best[:40] or fallback


def read_pdf_tables(pdf_path_or_bytes, max_tables=60):
    """PDF 전체에서 표를 읽어 목록으로 준다.

    반환: [{"key","title","page","rows","cols","cells"}...]  (원문 순서)
    """
    import io as _io
    raw = (pdf_path_or_bytes if isinstance(pdf_path_or_bytes, (bytes, bytearray))
           else open(pdf_path_or_bytes, "rb").read())

    # ★먼저 PyMuPDF 로 읽는다(격자를 훨씬 잘 잡는다). 못 찾으면 pdfplumber 로.
    try:
        got = _mupdf_tables(raw, max_tables)
        if got:
            return got
    except Exception as e:
        print(f"[표읽기] PyMuPDF 실패({type(e).__name__}) — pdfplumber 로 넘어감")

    src = _io.BytesIO(raw)
    out = []
    with pdfplumber.open(src) as doc:
        for pno, page in enumerate(doc.pages, 1):
            try:
                found = page.find_tables(TABLE_SETTINGS)
            except Exception:
                continue
            for ti, tbl in enumerate(found, 1):
                info = read_table(page, tbl)
                if not info:
                    continue
                if not any(c["text"] for c in info["cells"]):
                    continue                      # 글자 없는 껍데기 표
                info["page"] = pno
                info["title"] = _title_for(page, tbl, f"{pno}쪽 표{ti}")
                info["key"] = f"표 {pno}-{ti}"
                out.append(info)
                if len(out) >= max_tables:
                    return out
    return out


# ── PyMuPDF 로 읽기 ─────────────────────────────────
#  pdfplumber 는 선을 조각조각 보아 어떤 문서에서는 격자가 크게 틀어진다
#  (헌인마을 IM: 13행 7열짜리를 22행 7열로, 일부 열은 선 자체를 못 찾음).
#  PyMuPDF 의 lines_strict 는 같은 표를 제대로 잡았다 → 이쪽을 먼저 쓴다.
#  PyMuPDF 는 이미 깔려 있어 무게가 늘지 않는다.
def _mupdf_tables(pdf_bytes, max_tables=60):
    import fitz
    out = []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        for pno in range(doc.page_count):
            page = doc[pno]
            try:
                found = page.find_tables(strategy="lines_strict").tables
            except Exception:
                continue
            for ti, t in enumerate(found, 1):
                info = _from_mupdf(page, t)
                if info:
                    info["page"] = pno + 1
                    info["key"] = f"표 {pno + 1}-{ti}"
                    info["title"] = _mupdf_title(page, t, f"{pno + 1}쪽 표{ti}")
                    out.append(info)
                    if len(out) >= max_tables:
                        return out
    finally:
        doc.close()
    return out


def _from_mupdf(page, t):
    """PyMuPDF 표 → {rows, cols, cells}. 빈 줄·빈 칸은 접고 병합은 살린다."""
    grid = t.extract()
    if not grid:
        return None
    n_r = len(grid)
    n_c = max(len(r) for r in grid)
    grid = [list(r) + [None] * (n_c - len(r)) for r in grid]

    def cell(r, c):
        v = grid[r][c]
        return _clean(v) if v else ""

    keep_r = [r for r in range(n_r) if any(cell(r, c) for c in range(n_c))]
    keep_c = [c for c in range(n_c) if any(cell(r, c) for r in range(n_r))]
    if len(keep_r) < _MIN_ROWS or len(keep_c) < _MIN_COLS:
        return None

    # 세로로 이어지는 같은 글자는 원래 한 칸(병합)이었던 것
    txt = [[cell(r, c) for c in keep_c] for r in keep_r]
    R, C = len(txt), len(txt[0])
    used = [[False] * C for _ in range(R)]
    cells = []
    for r in range(R):
        for c in range(C):
            if used[r][c] or not txt[r][c]:
                continue
            rs = 1
            while r + rs < R and not txt[r + rs][c]:
                rs += 1
            cs = 1
            while c + cs < C and not txt[r][c + cs] and all(
                    not txt[r + k][c + cs] for k in range(rs)):
                cs += 1
            for rr in range(r, r + rs):
                for cc in range(c, c + cs):
                    used[rr][cc] = True
            cells.append({"r": r, "c": c, "rs": rs, "cs": cs, "text": txt[r][c]})
    if not cells:
        return None
    return {"rows": R, "cols": C, "cells": cells,
            "col_ratio": [1.0] * C, "row_ratio": [1.0] * R,
            "bbox": tuple(round(v, 1) for v in t.bbox)}


def _mupdf_title(page, t, fallback):
    x0, y0, x1, y1 = t.bbox
    words = page.get_text("words")          # (x0,y0,x1,y1,word,...)
    left = [w for w in words if x0 - 95 <= w[2] <= x0 + 2 and y0 - 6 <= w[1] <= y1]
    best = " ".join(w[4] for w in sorted(left, key=lambda w: (w[1], w[0])))
    if not best:
        above = [w for w in words
                 if y0 - 26 <= w[3] <= y0 + 2 and w[2] >= x0 - 6 and w[0] <= x1]
        best = " ".join(w[4] for w in sorted(above, key=lambda w: w[0]))
    return _clean(best)[:40] or fallback
