# -*- coding: utf-8 -*-
"""Docling 으로 PDF 표를 읽어 JSON 으로 뱉는 **별도 프로세스**.

왜 따로 도는가
  Docling 은 torch·numpy 를 자기 것으로 들고 온다. 앱과 같은 프로세스에서
  불러오면 이미 올라와 있는 numpy 와 부딪혀 앱 전체가 죽을 수 있다.
  그래서 아예 다른 파이썬 프로세스로 돌리고 결과만 JSON 으로 받는다.

쓰는 법 (docling_tables.py 가 알아서 부른다)
  python _docling_worker.py <읽을PDF> <결과JSON>
"""
import sys, os, json, io

sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

_MIN_ROWS = 2
_MIN_COLS = 2


def _clean(s):
    import re
    if not s:
        return ""
    s = s.replace("\r", "")
    s = re.sub(r"[ \t]+", " ", s)
    return "\n".join(ln.strip() for ln in s.split("\n")).strip()


class _Box:
    """제목 찾는 함수에 넘길 때만 쓰는 껍데기(.bbox 만 본다)."""

    def __init__(self, bbox):
        self.bbox = bbox


def _titles(pdf_path, tables):
    """표 위·왼쪽 글자에서 제목을 뽑는다. 지금 방식(pdf_tables)과 같게 맞춘다."""
    out = [t["_fallback"] for t in tables]
    try:
        import fitz
        sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
        import pdf_tables as _pt
    except Exception as e:                       # 못 불러와도 표는 살린다
        print(f"[docling] 제목 찾기 건너뜀({type(e).__name__})", file=sys.stderr)
        return out
    doc = fitz.open(pdf_path)
    try:
        for i, t in enumerate(tables):
            pno = t["page"] - 1
            if not (0 <= pno < doc.page_count):
                continue
            raw = _pt._mupdf_title(doc[pno], _Box(t["bbox"]), "")
            out[i] = _pt._tidy_title(raw, t, t["_fallback"])
    finally:
        doc.close()
    return out


def _one_table(tbl, page_h):
    """Docling 표 하나 → {rows, cols, cells:[{r,c,rs,cs,text}], bbox, col_ratio}."""
    data = tbl.data
    raw = list(getattr(data, "table_cells", None) or [])
    if not raw:
        return None
    n_r = int(getattr(data, "num_rows", 0) or 0)
    n_c = int(getattr(data, "num_cols", 0) or 0)
    if n_r < _MIN_ROWS or n_c < _MIN_COLS:
        return None

    cells, seen = [], set()
    for c in raw:
        r0 = int(c.start_row_offset_idx)
        c0 = int(c.start_col_offset_idx)
        rs = max(1, int(c.end_row_offset_idx) - r0)
        cs = max(1, int(c.end_col_offset_idx) - c0)
        if (r0, c0) in seen:                     # 같은 자리를 두 번 주는 일이 있다
            continue
        seen.add((r0, c0))
        cells.append({"r": r0, "c": c0, "rs": rs, "cs": cs,
                      "text": _clean(getattr(c, "text", ""))})
    if not any(x["text"] for x in cells):
        return None

    # 열 너비 — 열마다 **가장 왼쪽~가장 오른쪽** 을 재서 실제 폭을 쓴다.
    #   칸 하나하나의 폭을 평균 내면 글자가 짧은 칸 때문에 열이 좁아져서
    #   '3,887' 같은 숫자가 '3,88 / 7' 로 접혔다.
    lo = [None] * n_c
    hi = [None] * n_c
    for c in raw:
        bb = getattr(c, "bbox", None)
        if bb is None or int(c.end_col_offset_idx) - int(c.start_col_offset_idx) != 1:
            continue
        i = int(c.start_col_offset_idx)
        if not (0 <= i < n_c):
            continue
        l, r = float(bb.l), float(bb.r)
        lo[i] = l if lo[i] is None else min(lo[i], l)
        hi[i] = r if hi[i] is None else max(hi[i], r)
    ratio = [(hi[i] - lo[i]) if (lo[i] is not None and hi[i] is not None
                                 and hi[i] > lo[i]) else 0.0
             for i in range(n_c)]
    if not any(ratio):
        ratio = [1.0] * n_c
    else:
        avg = sum(v for v in ratio if v) / max(1, sum(1 for v in ratio if v))
        ratio = [v if v else avg for v in ratio]

    # 표 자리 — 위에서부터 재는 좌표(pdfplumber 와 같은 기준)로 바꾼다.
    bbox = (0.0, 0.0, 0.0, 0.0)
    prov = list(getattr(tbl, "prov", None) or [])
    page = 1
    if prov:
        p = prov[0]
        page = int(getattr(p, "page_no", 1) or 1)
        bb = getattr(p, "bbox", None)
        if bb is not None:
            l, t, r, b = float(bb.l), float(bb.t), float(bb.r), float(bb.b)
            if str(getattr(bb, "coord_origin", "")).endswith("BOTTOMLEFT"):
                t, b = page_h - t, page_h - b     # 아래가 0인 좌표 → 위가 0인 좌표
            bbox = (round(l, 1), round(min(t, b), 1),
                    round(r, 1), round(max(t, b), 1))
    return {"page": page, "rows": n_r, "cols": n_c, "cells": cells,
            "col_ratio": ratio, "row_ratio": [1.0] * n_r, "bbox": bbox}


def main():
    pdf_path, out_path = sys.argv[1], sys.argv[2]

    from docling.document_converter import DocumentConverter, PdfFormatOption
    from docling.datamodel.base_models import InputFormat
    from docling.datamodel.pipeline_options import PdfPipelineOptions

    opt = PdfPipelineOptions()
    # ★OCR 은 끈다. 글자가 살아 있는 PDF 에 OCR 을 걸면 한글이 깨진다
    #   ('구분'→'7是', '선순위'→'享'). 표 모양만 Docling 의 표 모델에 맡기고
    #   글자는 원문 그대로 가져온다.
    opt.do_ocr = False
    opt.do_table_structure = True
    opt.table_structure_options.do_cell_matching = True

    conv = DocumentConverter(format_options={
        InputFormat.PDF: PdfFormatOption(pipeline_options=opt)})
    doc = conv.convert(pdf_path).document

    heights = {}
    for no, pg in (getattr(doc, "pages", None) or {}).items():
        try:
            heights[int(no)] = float(pg.size.height)
        except Exception:
            pass

    out, per_page = [], {}
    for tbl in (getattr(doc, "tables", None) or []):
        prov = list(getattr(tbl, "prov", None) or [])
        pno = int(getattr(prov[0], "page_no", 1)) if prov else 1
        info = _one_table(tbl, heights.get(pno, 842.0))
        if not info:
            continue
        per_page[info["page"]] = per_page.get(info["page"], 0) + 1
        info["key"] = f"표 {info['page']}-{per_page[info['page']]}"
        info["_fallback"] = f"{info['page']}쪽 표{per_page[info['page']]}"
        out.append(info)

    out.sort(key=lambda t: (t["page"], t["bbox"][1]))
    for t, title in zip(out, _titles(pdf_path, out)):
        t["title"] = title
        t.pop("_fallback", None)

    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False)
    print(f"[docling] 표 {len(out)}개", file=sys.stderr)


if __name__ == "__main__":
    main()
