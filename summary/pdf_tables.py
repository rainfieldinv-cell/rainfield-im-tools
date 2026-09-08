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


def _line_cover(page, x0, x1):
    """가로선을 y 자리별로 묶어, 표 폭을 얼마나 덮는지 비율(0~1)로 준다.

    한 줄이 여러 조각으로 그려져 있어(칸마다 따로) 조각 하나만 보면
    '짧은 선' 으로 보인다. 겹치는 구간을 합쳐서 재야 한다.
    """
    W = max(1.0, x1 - x0)
    by = {}
    for e in page.horizontal_edges:
        a, b = max(x0, e["x0"]), min(x1, e["x1"])
        if b - a <= 1:
            continue
        by.setdefault(round(e["top"] / 1.5), []).append((a, b))
    out = {}
    for key, segs in by.items():
        segs.sort()
        cov, cur = 0.0, None
        for a, b in segs:
            if cur is None or a > cur[1]:
                if cur:
                    cov += cur[1] - cur[0]
                cur = [a, b]
            else:
                cur[1] = max(cur[1], b)
        if cur:
            cov += cur[1] - cur[0]
        out[key * 1.5] = cov / W
    return out


def _vlines_in(page, y_top, y_bot):
    """그 띠를 절반 넘게 지나는 세로선들의 x 목록."""
    band = max(1.0, y_bot - y_top)
    out = []
    for e in page.vertical_edges:
        if min(y_bot, e["bottom"]) - max(y_top, e["top"]) >= band * 0.5:
            out.append(e["x0"])
    return out


def _header_cells(page, xs, y_top, y_bot):
    """머리글 줄을 칸으로 만들어 준다.

    ★반드시 칸을 박아 넣어야 한다. 빈자리로 두면 아래 '빈자리 채우기' 가
      머리글부터 본문 끝까지 한 칸으로 이어 붙인다('비 고' 가 9줄을 먹던 원인).
    세로선이 없는 자리는 원본에서 이어진 칸이므로 그대로 이어 붙인다.
    """
    n_col = len(xs) - 1
    vx = _vlines_in(page, y_top, y_bot)
    out, c = [], 0
    while c < n_col:
        c2 = c
        while c2 + 1 < n_col and not any(abs(x - xs[c2 + 1]) <= 6 for x in vx):
            c2 += 1
        try:
            txt = page.crop((xs[c], y_top, xs[c2 + 1], y_bot),
                            strict=False).extract_text() or ""
        except Exception:
            txt = ""
        out.append({"r": 0, "c": c, "rs": 1, "cs": c2 - c + 1,
                    "text": _clean(txt)})
        c = c2 + 1
    return out


def _has_col_lines(page, xs, y_top, y_bot):
    """그 띠 안에 표의 **열 경계선**이 실제로 그어져 있나.

    머리글 줄인지, 표 위에 얹힌 설명글인지 가르는 잣대다.
    ('▪ 토지 확보현황 : 96.9% ...', '(단위:억원)' 같은 줄을 머리글로
     잘못 삼아 표에 끌어들이던 것을 막는다.)
    """
    inner = xs[1:-1]
    if not inner:
        return False
    band = max(1.0, y_bot - y_top)
    hit = 0
    for x in inner:
        for e in page.vertical_edges:
            if abs(e["x0"] - x) > 6:
                continue
            if min(y_bot, e["bottom"]) - max(y_top, e["top"]) >= band * 0.5:
                hit += 1
                break
    return hit >= max(1, (len(inner) + 1) // 2)


def _header_above(page, xs, ys):
    """표 바로 위에 붙어 있는 **머리글 줄**을 찾아 그 윗선 y 를 준다.

    돈암동 사업일정 표가 이랬다. 머리글(구분/일정/내용/비고)이 본문과
    0.5pt 떨어져 그려져 있어 find_tables 가 다른 덩어리로 보고 통째로 버렸다
    (본문만 '17.04 부터 시작하는 표로 잡힘).
    """
    x0, x1, top = xs[0], xs[-1], ys[0]
    gaps = [ys[i + 1] - ys[i] for i in range(len(ys) - 1)]
    h = sorted(gaps)[len(gaps) // 2] if gaps else 12.0
    cover = _line_cover(page, x0, x1)
    # ★0.95 로 본다. 0.8 로 잡으면 칸마다 그려진 **글자 배경 상자**의 밑선
    #   (사업일정 머리글 안쪽 상자, 폭의 90%)까지 진짜 줄로 쳐서 막혀 버린다.
    FULL = 0.95
    cand = [y for y, c in cover.items()
            if c >= FULL and top - 3.0 * h <= y <= top - 3.0]
    if not cand:
        return None
    y = max(cand)
    # 머리글과 본문 사이에 또 다른 줄이 있으면 '한 줄' 이 아니다
    if any(y + 2 < z < top - 2 and c >= FULL for z, c in cover.items()):
        return None
    if not _has_col_lines(page, xs, y, top):
        return None
    try:
        txt = page.crop((x0, y, x1, top), strict=False).extract_text() or ""
    except Exception:
        txt = ""
    return y if txt.strip() else None


def _restore_merges(page, xs, ys, cells, bbox):
    """원본에 **한 칸으로 그려진 상자**를 보고, 쪼개진 칸을 도로 합친다.

    돈암동 사업일정의 '진행 경과' 는 6줄을 합친 한 칸인데, 그 안에 글자마다
    작은 상자가 또 그려져 있다. pdfplumber 는 그 작은 상자의 선까지 격자로 쳐서
    '진행' 과 '경과' 를 다른 줄로 갈라 놓았다.
    바깥 상자만 보면 원래 한 칸이라는 것을 알 수 있다.
    """
    n_col, n_row = len(xs) - 1, len(ys) - 1
    bx0, by0, bx1, by1 = bbox
    boxes = []
    for r in page.rects:
        a, b, c, d = r["x0"], r["top"], r["x1"], r["bottom"]
        if c - a < 3 or d - b < 3:                 # 선처럼 납작한 것은 상자가 아니다
            continue
        if a < bx0 - 3 or c > bx1 + 3 or b < by0 - 3 or d > by1 + 3:
            continue
        boxes.append((a, b, c, d))
    if not boxes or len(boxes) > 400:              # 도형이 수만 개인 문서에서 멈추지 않게
        return cells

    # 다른 상자 **안에** 들어 있는 상자는 글자 배경이다 → 바깥 것만 남긴다
    outer = []
    for i, (a, b, c, d) in enumerate(boxes):
        inner = False
        for j, (p, q, r2, s) in enumerate(boxes):
            if j == i:
                continue
            if (p <= a + 1 and q <= b + 1 and r2 >= c - 1 and s >= d - 1
                    and (r2 - p) * (s - q) > (c - a) * (d - b) + 1):
                inner = True
                break
        if not inner:
            outer.append((a, b, c, d))

    spans = []
    for (a, b, c, d) in outer:
        c0, c1 = _idx(a, xs), _idx(c, xs)
        r0, r1 = _idx(b, ys), _idx(d, ys)
        cs, rs = c1 - c0, r1 - r0
        if rs <= 1 and cs <= 1:
            continue
        if rs >= n_row and cs >= n_col:            # 표 전체 테두리
            continue
        if abs(xs[c0] - a) > 3 or abs(xs[min(c1, n_col)] - c) > 3:
            continue                               # 격자에 안 맞으면 건너뛴다
        if abs(ys[r0] - b) > 3 or abs(ys[min(r1, n_row)] - d) > 3:
            continue
        spans.append((r0, c0, min(rs, n_row - r0), min(cs, n_col - c0),
                      (a, b, c, d)))
    if not spans:
        return cells

    covered = set()
    for r0, c0, rs, cs, _ in spans:
        for rr in range(r0, r0 + rs):
            for cc in range(c0, c0 + cs):
                covered.add((rr, cc))
    out = [c for c in cells if (c["r"], c["c"]) not in covered]
    for r0, c0, rs, cs, box in spans:
        try:
            txt = page.crop(box, strict=False).extract_text() or ""
        except Exception:
            txt = ""
        out.append({"r": r0, "c": c0, "rs": rs, "cs": cs, "text": _clean(txt)})
    return out


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
    if len(xs) - 1 < _MIN_COLS or len(ys) - 1 < _MIN_ROWS:
        return None
    # ★머리글 줄이 따로 떨어져 있으면 끌어와서 첫 줄로 붙인다
    _head = _header_above(page, xs, ys)
    if _head is not None:
        ys = [_head] + ys
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

    # ★끌어온 머리글 줄은 칸으로 박아 넣는다(빈자리로 두면 아래로 이어 붙는다)
    if _head is not None:
        cells += _header_cells(page, xs, ys[0], ys[1])

    # ★원본에 한 칸으로 그려진 상자를 보고 쪼개진 칸을 도로 합친다.
    #   (빈자리 채우기보다 **먼저** 해야 한다. 안 그러면 줄마다 잘라 넣은 뒤라
    #    '진행' / '경과' 가 이미 다른 줄로 굳어 버린다.)
    cells = _restore_merges(page, xs, ys, cells, tbl.bbox)

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
            # 줄마다 따로 잘라 본다.
            #   전부 같은 글자 → 원래 하나였던 병합 칸
            #   서로 다르면   → 줄마다 값이 있는 것(비고 열처럼) → 따로 넣는다
            parts = []
            for rr in range(r, r2 + 1):
                try:
                    parts.append(_clean(page.crop(
                        (xs[c], ys[rr], xs[c + 1], ys[rr + 1]),
                        strict=False).extract_text()))
                except Exception:
                    parts.append("")
            uniq = {p for p in parts if p}
            if len(uniq) <= 1:
                try:
                    txt = page.crop((xs[c], ys[r], xs[c + 1], ys[r2 + 1]),
                                    strict=False).extract_text()
                except Exception:
                    txt = ""
                cells.append({"r": r, "c": c, "rs": r2 - r + 1, "cs": 1,
                              "text": _clean(txt)})
            else:
                for k, val in enumerate(parts):
                    cells.append({"r": r + k, "c": c, "rs": 1, "cs": 1,
                                  "text": val})
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

    # ★어느 한쪽이 항상 옳지 않다.
    #   헌인마을 IM: PyMuPDF 가 제대로, pdfplumber 는 격자가 틀어짐
    #   돈암동  IM: pdfplumber 가 제대로, PyMuPDF 는 페이지를 표 하나로 뭉갬
    #   → 둘 다 읽어서 **쪽마다 더 잘 나온 쪽**을 고른다.
    try:
        a = _mupdf_tables(raw, 400)
    except Exception as e:
        print(f"[표읽기] PyMuPDF 실패({type(e).__name__})")
        a = []
    try:
        b = _plumber_tables(_io.BytesIO(raw), 400)
    except Exception as e:
        print(f"[표읽기] pdfplumber 실패({type(e).__name__})")
        b = []
    return _pick_better(a, b, max_tables)


def _score(tables):
    """쪽 하나의 결과가 얼마나 잘 나왔는지. 낮을수록 좋다.

    둘 다 실패할 때의 모습이 '여러 표를 하나로 뭉치고 칸 하나에 글을 몰아넣는 것'
    이므로, **뭉친 칸이 적고 표가 잘게 나뉜 쪽**을 좋게 본다.
    """
    if not tables:
        return (0, 10 ** 6, 10 ** 6)
    cells = [c for t in tables for c in t["cells"]]
    if not cells:
        return (0, 10 ** 6, 10 ** 6)
    # ① 표가 잘게 나뉜 쪽이 좋다(둘 다 실패할 때의 모습이 '여러 표를 하나로 뭉치기')
    # ② 같으면 군더더기 행이 적은 쪽(선이 겹쳐 없는 행이 생기는 것을 피한다)
    # ③ 그래도 같으면 글이 한 칸에 몰리지 않은 쪽
    giant = sum(1 for c in cells
                if len(c["text"]) > 80 or c["text"].count(chr(10)) >= 4)
    return (-len(tables), sum(t["rows"] for t in tables), giant)


def _pick_better(a, b, max_tables):
    """쪽 번호별로 더 잘 나온 쪽을 골라 합친다."""
    pages = sorted({t["page"] for t in a} | {t["page"] for t in b})
    out = []
    for pno in pages:
        ta = [t for t in a if t["page"] == pno]
        tb = [t for t in b if t["page"] == pno]
        out += ta if _score(ta) <= _score(tb) else tb
    out.sort(key=lambda t: (t["page"], t.get("bbox", (0, 0))[1]))
    for i, t in enumerate(out, 1):
        t["key"] = f"표 {t['page']}-{i}"
    return out[:max_tables]


def _plumber_tables(src, max_tables=400):
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
                info["title"] = _tidy_title(
                    _title_for(page, tbl, ""), info, f"{pno}쪽 표{ti}")
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
                    info["title"] = _tidy_title(
                        _mupdf_title(page, t, ""), info,
                        f"{pno + 1}쪽 표{ti}")
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


# ── 제목 다듬기 ─────────────────────────────────────
#  왼쪽 여백 글자를 전부 긁으면 옆 표의 라벨까지 붙는다
#  ('구분 진행 경과 사업일정 추진 계획' 처럼).
#  표 안에 이미 있는 말은 빼고, 짧은 덩어리 하나만 남긴다.
_STOP = ("구분", "구 분", "내용", "내 용", "비고", "비 고", "일정", "일 정",
         "금액", "합계", "진행", "경과", "추진", "계획", "단위")


def _tidy_title(raw, table_info, fallback):
    words = [w for w in (raw or "").split() if w]
    words = [w for w in words if w not in _STOP]
    inside = {c["text"].strip() for c in (table_info or {}).get("cells", [])}
    words = [w for w in words if w not in inside]
    out = " ".join(words).strip(" :·-—")
    return out[:34] or fallback
