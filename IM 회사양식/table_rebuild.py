"""망가진 표를 원본 PDF 의 선·칠·글 자리로 다시 짠다(2026-10-15).

★사용자: '표 칸 합치기·구조가 가장 중요 — 어떤 수를 써서라도 원본과 똑같이'.
  pdf2docx 는 머리글이 두 줄이고 칸이 합쳐진 표(내곡동 '서울 주요 권역 고급주택 실거래비교')에서
  머리글 칸이 어긋나고 아래 줄 숫자가 머리글 칸으로 올라갔다.
  → PyMuPDF find_tables 가 원본 선으로 찾은 칸 격자(합친 칸 포함)를 그대로 워드 표로 만든다.
  글은 원본 글 조각(크기·굵기·색)을, 칠은 원본 칠한 네모를, 테두리는 원본 선(색·굵기·점선)을 칸마다 옮긴다.

쓰는 곳: pdf2docx 결과 표가 원본 칸과 글이 안 맞을 때만(맞는 표는 그대로 둔다).
안쪽 표가 든 표·그림이 든 표는 건드리지 않는다(따로 규칙이 있다).
"""
import copy
import re

import fitz
from docx.oxml.ns import qn
from lxml import etree

BULLETS = "•▪■□◦○●➢✓※▶▷◆◇◼◻-·∙ㆍ"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def _norm(t):
    return re.sub(r"\s+", "", t or "")


def _tw(pt):
    return int(round(pt * 20))


def _hex(c):
    if c is None:
        return None
    if isinstance(c, int):
        return "%06X" % c
    r, g, b = (max(0, min(255, int(round(v * 255)))) for v in c[:3])
    return "%02X%02X%02X" % (r, g, b)


def _font_name(name):
    name = (name or "").split("+")[-1]
    return re.sub(r"(?i)[-,]?(bold|regular|light|medium)$", "", name) or "Malgun Gothic"


# ──────────────────────────────────────────────
# 원본 읽기
# ──────────────────────────────────────────────
def _page_art(page):
    """쪽의 선(가늘고 긴 것)·칠한 네모. ([(rect, 색hex, 굵기pt, 점선?)], [(rect, 색hex)])"""
    lines, fills = [], []
    for dr in page.get_drawings():
        dashed = bool(dr.get("dashes")) and dr.get("dashes") not in ("[] 0", "[]0", "[] 0.0")
        for it in dr["items"]:
            if it[0] == "re":
                r = fitz.Rect(it[1])
            elif it[0] == "l":
                r = fitz.Rect(it[1], it[2]).normalize()
            elif it[0] == "qu":
                r = it[1].rect
            else:
                continue
            if dr.get("fill") is not None and r.width > 2.5 and r.height > 2.5:
                fills.append((r, _hex(dr["fill"])))
            thin = min(r.width, r.height)
            if dr.get("fill") is not None and thin <= 1.6 and max(r.width, r.height) >= 2:
                lines.append((r, _hex(dr["fill"]), max(thin, 0.25), dashed))
            elif dr.get("color") is not None and (it[0] == "l" or dr.get("fill") is None):
                w = max(dr.get("width") or 0.5, 0.25)
                if it[0] == "l":
                    lines.append((r + ((-w / 2, 0, w / 2, 0) if r.width < 0.5 else (0, -w / 2, 0, w / 2)),
                                  _hex(dr["color"]), w, dashed))
                else:                                          # 테두리만 그린 네모 → 네 변
                    for e in (fitz.Rect(r.x0, r.y0 - w / 2, r.x1, r.y0 + w / 2),
                              fitz.Rect(r.x0, r.y1 - w / 2, r.x1, r.y1 + w / 2),
                              fitz.Rect(r.x0 - w / 2, r.y0, r.x0 + w / 2, r.y1),
                              fitz.Rect(r.x1 - w / 2, r.y0, r.x1 + w / 2, r.y1)):
                        lines.append((e, _hex(dr["color"]), w, dashed))
    # 그림 조각으로 그린 점선(내곡동 우발채무 '소계' 굵은 점선 네모)도 선으로
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"])
        thin = min(r.width, r.height)
        if 0 < thin <= 3.5 and max(r.width, r.height) >= 2 and r.get_area() < 400:
            if r.width >= r.height:
                lines.append((fitz.Rect(r.x0, (r.y0 + r.y1) / 2 - thin / 2, r.x1, (r.y0 + r.y1) / 2 + thin / 2),
                              "000000", thin, True))
            else:
                lines.append((fitz.Rect((r.x0 + r.x1) / 2 - thin / 2, r.y0, (r.x0 + r.x1) / 2 + thin / 2, r.y1),
                              "000000", thin, True))
    return lines, fills


def _spans(page, clip):
    """clip 안 글 줄들 [(bbox, [span…])] — 가로쓰기만. span 에 글자별 자리(chars)도 담는다."""
    out = []
    for b in page.get_text("rawdict", clip=clip)["blocks"]:
        for l in b.get("lines", []):
            if l.get("dir") != (1.0, 0.0):
                continue
            sp = []
            for s in l["spans"]:
                chars = [(c["c"], fitz.Rect(c["bbox"])) for c in s["chars"]]
                text = "".join(c for c, _ in chars)
                if text:
                    sp.append({"text": text, "font": s["font"], "size": s["size"], "flags": s.get("flags", 0),
                               "color": s.get("color", 0), "chars": chars})
            if sp and "".join(s["text"] for s in sp).strip():
                out.append((fitz.Rect(l["bbox"]), sp))
    return out


def _pieces(text_lines, keyf):
    """글 줄을 글자 자리마다 keyf(점)→칸 열쇠로 나눈다. [(열쇠, bbox, [span…])]
    ★한 줄로 읽힌 머리글 '공급면적 실거래가 공급평당 …' 이 칸 하나에 몰렸다(내곡동 31쪽)."""
    out = []
    for r, sp in text_lines:
        cur_key, cur_sp, cur_box = None, [], None
        pieces = []
        for s in sp:
            if s is None:
                continue
            part = None
            for ch, cb in s.get("chars") or [(s["text"], r)]:
                k = keyf(fitz.Point((cb.x0 + cb.x1) / 2, (cb.y0 + cb.y1) / 2)) if ch.strip() else cur_key
                if k != cur_key and ch.strip():
                    if cur_sp:
                        pieces.append((cur_key, cur_box, cur_sp))
                    cur_key, cur_sp, cur_box, part = k, [], None, None
                if part is None:
                    part = dict(s, text="", chars=[])
                    cur_sp.append(part)
                part["text"] += ch
                part["chars"].append((ch, cb))
                if ch.strip():
                    cur_box = fitz.Rect(cb) if cur_box is None else cur_box | cb
            part = None
        if cur_sp:
            pieces.append((cur_key, cur_box, cur_sp))
        out.extend((k, bx, sps) for k, bx, sps in pieces if bx is not None)
    return out


def _cell_of(r, cells):
    c = fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)
    for k, cr in enumerate(cells):
        if cr.contains(c):
            return k
    return None


# ──────────────────────────────────────────────
# 표 격자
# ──────────────────────────────────────────────
def _grid(tab):
    """find_tables 표 → (xs, ys, 칸들[(rect, r0, c0, r1, c1)])."""
    cells = [fitz.Rect(c) for c in tab.cells if c]
    xs = sorted({round(v, 1) for c in cells for v in (c.x0, c.x1)})
    ys = sorted({round(v, 1) for c in cells for v in (c.y0, c.y1)})

    def snap(vals, v):
        return min(range(len(vals)), key=lambda i: abs(vals[i] - v))
    # 아주 가까운 선(1pt 안)은 하나로
    def merge(vals):
        out = []
        for v in vals:
            if out and v - out[-1] < 1.0:
                continue
            out.append(v)
        return out
    xs, ys = merge(xs), merge(ys)
    spec = []
    for c in cells:
        r0, r1 = snap(ys, c.y0), snap(ys, c.y1)
        c0, c1 = snap(xs, c.x0), snap(xs, c.x1)
        if r1 > r0 and c1 > c0:
            spec.append((c, r0, c0, r1, c1))
    return xs, ys, spec


# ──────────────────────────────────────────────
# 칸 안 글 → 문단
# ──────────────────────────────────────────────
def _paragraphs(lines, cell):
    """칸 안 글 줄 → 문단 묶음. 줄 머리가 글머리이거나, 앞 줄이 짧게 끝났거나, 줄 사이가 크게 벌어지면 새 문단."""
    lines = sorted(lines, key=lambda x: (round(x[0].y0), x[0].x0))
    # 같은 높이(같은 줄) 조각은 한 줄로
    rows = []
    for r, sp in lines:
        if rows and abs(rows[-1][0].y0 - r.y0) < 2 and abs(rows[-1][0].y1 - r.y1) < 3:
            rows[-1] = (rows[-1][0] | r, rows[-1][1] + [None] + sp)
        else:
            rows.append((fitz.Rect(r), list(sp)))
    paras = []
    width = cell.width
    for i, (r, sp) in enumerate(rows):
        txt = "".join(s["text"] for s in sp if s).strip()
        new = not paras
        if paras:
            pr, psp = paras[-1][-1]
            gap = r.y0 - pr.y1
            h = max(pr.height, 1)
            if txt[:1] in BULLETS or re.match(r"^(\(?\d{1,2}[).]|[①-⑳]|주\d|\*|※)", txt):
                new = True
            elif cell.x1 - pr.x1 > width * 0.2 and r.x0 - cell.x0 < width * 0.3:
                new = True                                     # 앞 줄이 짧게 끝남
            elif gap > h * 0.8:
                new = True
            elif abs(((r.x0 + r.x1) / 2) - ((cell.x0 + cell.x1) / 2)) < 2 and \
                    abs(((pr.x0 + pr.x1) / 2) - ((cell.x0 + cell.x1) / 2)) < 2:
                new = True                                     # 가운데 맞춘 줄은 줄마다 따로(원본 줄바꿈 그대로)
        if new:
            paras.append([(r, sp)])
        else:
            paras[-1].append((r, sp))
    return paras


def _align(par, cell):
    lg = [r.x0 - cell.x0 for r, _ in par]
    rg = [cell.x1 - r.x1 for r, _ in par]
    if len(par) >= 2 and all(g < 6 for g in rg[:-1]) and all(abs(g - lg[0]) < 3 for g in lg[1:]):
        return "both", lg[0]
    if all(abs(a - b) < max(3, cell.width * 0.04) for a, b in zip(lg, rg)) and min(lg) > 2:
        return "center", 0
    if all(b < 6 for b in rg) and min(lg) > 8:
        return "right", 0
    return "left", min(lg)


def _run(text, s):
    r = etree.Element(qn("w:r"))
    rpr = etree.SubElement(r, qn("w:rPr"))
    fn = _font_name(s["font"])
    etree.SubElement(rpr, qn("w:rFonts"), {qn("w:ascii"): fn, qn("w:hAnsi"): fn, qn("w:eastAsia"): fn, qn("w:cs"): fn})
    bold = "bold" in s["font"].lower() or bool(s.get("flags", 0) & 16)
    etree.SubElement(rpr, qn("w:b"), {qn("w:val"): "1" if bold else "0"})
    etree.SubElement(rpr, qn("w:i"), {qn("w:val"): "0"})
    etree.SubElement(rpr, qn("w:color"), {qn("w:val"): _hex(s.get("color", 0)) or "000000"})
    sz = str(max(2, int(round(s["size"] * 2))))
    etree.SubElement(rpr, qn("w:sz"), {qn("w:val"): sz})
    etree.SubElement(rpr, qn("w:szCs"), {qn("w:val"): sz})
    t = etree.SubElement(r, qn("w:t"))
    t.text = text
    t.set(XML_SPACE, "preserve")
    return r


def _par_el(par, cell):
    p = etree.Element(qn("w:p"))
    ppr = etree.SubElement(p, qn("w:pPr"))
    jc, left = _align(par, cell)
    sizes = [s["size"] for _, sp in par for s in sp if s]
    size = max(sizes) if sizes else 9
    if len(par) >= 2:
        pitch = (par[-1][0].y0 - par[0][0].y0) / (len(par) - 1)
        line = int(round(pitch / (size * 1.5625) * 240))
        line = max(200, min(line, 480))
    else:
        line = 240
    etree.SubElement(ppr, qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0", qn("w:line"): str(line),
                                            qn("w:lineRule"): "auto"})
    etree.SubElement(ppr, qn("w:ind"), {qn("w:left"): str(_tw(max(0, left)) if jc in ("left", "both") else 0),
                                        qn("w:right"): "0", qn("w:firstLine"): "0"})
    etree.SubElement(ppr, qn("w:jc"), {qn("w:val"): jc})
    first = True
    for k, (r, sp) in enumerate(par):
        if k:                                                  # 줄 사이 = 띄어쓰기(원본 줄바꿈은 워드가 다시)
            last_t = p.findall(f".//{qn('w:t')}")
            if last_t and not (last_t[-1].text or "").endswith(" "):
                last_t[-1].text = (last_t[-1].text or "") + " "
        for s in sp:
            if s is None:                                      # 같은 줄 떨어진 조각 사이
                last_t = p.findall(f".//{qn('w:t')}")
                if last_t and not (last_t[-1].text or "").endswith(" "):
                    last_t[-1].text = (last_t[-1].text or "") + " "
                continue
            text = s["text"]
            if first:
                text = text.lstrip()
            if not text:
                continue
            first = False
            p.append(_run(text, s))
    ts = p.findall(f".//{qn('w:t')}")
    if ts:
        ts[-1].text = (ts[-1].text or "").rstrip()
    for t in p.findall(f".//{qn('w:t')}"):                   # 빈 글 조각은 뺀다
        if not t.text and len(p.findall(f".//{qn('w:t')}")) > 1:
            r = t.getparent()
            r.getparent().remove(r)
    return p


# ──────────────────────────────────────────────
# 테두리·칠
# ──────────────────────────────────────────────
def _edge_border(edge, horizontal, lines):
    """칸 한 변(edge rect)에 겹친 원본 선 → (색, 굵기pt, 점선?) 또는 None."""
    L = edge.width if horizontal else edge.height
    best = None
    cov = 0.0
    for r, col, w, dashed in lines:
        if horizontal:
            if r.width < r.height or abs((r.y0 + r.y1) / 2 - edge.y0) > max(1.6, w):
                continue
            ov = min(r.x1, edge.x1) - max(r.x0, edge.x0)
        else:
            if r.height < r.width or abs((r.x0 + r.x1) / 2 - edge.x0) > max(1.6, w):
                continue
            ov = min(r.y1, edge.y1) - max(r.y0, edge.y0)
        if ov > 0:
            cov += ov
            if best is None or ov > best[0]:
                best = (ov, col, w, dashed)
    if best is None or cov < L * 0.5:
        return None
    return best[1:]


def _borders_el(rect, lines):
    b = etree.Element(qn("w:tcBorders"))
    for side, edge, hor in (("top", fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0), True),
                            ("left", fitz.Rect(rect.x0, rect.y0, rect.x0, rect.y1), False),
                            ("bottom", fitz.Rect(rect.x0, rect.y1, rect.x1, rect.y1), True),
                            ("right", fitz.Rect(rect.x1, rect.y0, rect.x1, rect.y1), False)):
        got = _edge_border(edge, hor, lines)
        if got is None:
            etree.SubElement(b, qn("w:" + side), {qn("w:val"): "nil"})
        else:
            col, w, dashed = got
            etree.SubElement(b, qn("w:" + side), {qn("w:val"): "dashed" if dashed else "single",
                                                  qn("w:sz"): str(max(2, min(48, int(round(w * 8))))),
                                                  qn("w:space"): "0", qn("w:color"): col or "000000"})
    return b


def _fill_of(rect, fills):
    best, bc = 0.0, None
    a = rect.get_area()
    if not a:
        return None
    for r, col in fills:
        if col in (None, "FFFFFF"):
            ov = (r & rect).get_area() / a if r.intersects(rect) else 0
            if ov > best and col == "FFFFFF":
                best, bc = ov, None
            continue
        ov = (r & rect).get_area() / a if r.intersects(rect) else 0
        if ov > best:
            best, bc = ov, col
    return bc if best >= 0.6 else None


# ──────────────────────────────────────────────
# 선 → 칸 격자(합친 칸 포함)
# ──────────────────────────────────────────────
TOL = 1.5            # 선이 '닿았다' 고 보는 거리(pt)


def _cluster(vals, tol=1.2):
    out = []
    for v in sorted(vals):
        if out and v - out[-1][-1] <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(g) / len(g) for g in out]


def _segs(lines, fills, box):
    """표 범위 안 선 조각 [(rect, 'h'|'v', 진짜선?)] — 칠한 네모 가장자리도(칸 나눔 자리로만) 넣는다."""
    big = box + (-2, -2, 2, 2)
    out = []
    for r, col, w, dashed in lines:
        if not big.intersects(r) and not big.contains(r):
            continue
        o = "h" if r.width >= r.height else "v"
        if o == "h" and not (big.y0 <= (r.y0 + r.y1) / 2 <= big.y1) or                 o == "v" and not (big.x0 <= (r.x0 + r.x1) / 2 <= big.x1):
            continue
        r = fitz.Rect(max(r.x0, big.x0), max(r.y0, big.y0), min(r.x1, big.x1), min(r.y1, big.y1))             if o == "h" else fitz.Rect(r.x0, max(r.y0, big.y0), r.x1, min(r.y1, big.y1))
        if o == "h":
            r = fitz.Rect(r.x0, (r.y0 + r.y1) / 2 - 0.2, r.x1, (r.y0 + r.y1) / 2 + 0.2) if r.height <= 0 else r
        if (r.width if o == "h" else r.height) < 1:
            continue
        out.append((r, o, True))
    for r, col in fills:
        if col in (None, "FFFFFF") or not (big & r).get_area() or r.get_area() > box.get_area() * 0.9:
            continue
        r = r & big
        for e, o in ((fitz.Rect(r.x0, r.y0, r.x1, r.y0), "h"), (fitz.Rect(r.x0, r.y1, r.x1, r.y1), "h"),
                     (fitz.Rect(r.x0, r.y0, r.x0, r.y1), "v"), (fitz.Rect(r.x1, r.y0, r.x1, r.y1), "v")):
            out.append((e, o, False))
    return out


def _mid(r, o):
    return (r.y0 + r.y1) / 2 if o == "h" else (r.x0 + r.x1) / 2


def _components(segs):
    """닿아 있는 선끼리 묶음(안쪽 표 = 바깥 선과 떨어진 섬)."""
    n = len(segs)
    par = list(range(n))

    def f(a):
        while par[a] != a:
            par[a] = par[par[a]]
            a = par[a]
        return a
    infl = [s[0] + (-TOL, -TOL, TOL, TOL) for s in segs]
    order = sorted(range(n), key=lambda i: infl[i].x0)
    for a_i, a in enumerate(order):
        for b in order[a_i + 1:]:
            if infl[b].x0 > infl[a].x1:
                break
            if infl[a].intersects(infl[b]):
                par[f(a)] = f(b)
    groups = {}
    for i in range(n):
        groups.setdefault(f(i), []).append(segs[i])
    return list(groups.values())


def _bbox(segs):
    r = fitz.Rect(segs[0][0])
    for s in segs[1:]:
        r |= s[0]
    return r


def _intervals_reach(rs, lo, hi, tol=3.0):
    """세로 선 조각들(같은 x)이 이어져 lo 근처 또는 hi 근처에 닿는가."""
    iv = sorted((r.y0, r.y1) for r in rs)
    merged = []
    for a, b in iv:
        if merged and a <= merged[-1][1] + TOL * 2:
            merged[-1][1] = max(merged[-1][1], b)
        else:
            merged.append([a, b])
    return any(a <= lo + tol or b >= hi - tol for a, b in merged)


def _split_islands(segs, box):
    """(바깥 선들, [안쪽 표 선 묶음…], [안쪽 표 자리…]).
    바깥 표의 칸 줄(세로선)은 표 맨 위나 맨 아래까지 닿는다. 칸 가운데에만 떠 있는 세로선 묶음 + 거기 닿은 가로선 = 안쪽 표
    (내곡동 '신탁사' 칸 안 [구분|신탁 BL] 표 — 오른쪽은 바깥 표 선에 붙어 있어도 안쪽 표다)."""
    vs = [s for s in segs if s[1] == "v" and s[2]]
    hs = [s for s in segs if s[1] == "h"]
    xs_all = _cluster([_mid(r, "v") for r, o, _ in vs])
    X = [box.x0, box.x1]
    for x in xs_all:
        rs = [r for r, o, _ in vs if abs(_mid(r, "v") - x) <= TOL]
        cov = sum(r.height for r in rs)
        if _intervals_reach(rs, box.y0, box.y1) or cov >= box.height * 0.6:   # 소계 점선 네모(내곡동 우발채무)도 바깥 칸 줄
            X.append(x)
    X = sorted(_cluster(X))
    on_X = lambda x: any(abs(x - v) <= TOL for v in X)
    foreign = [s for s in vs if not on_X(_mid(s[0], "v"))]
    if not foreign:
        return segs, [], []
    regions = []
    for xl, xr in zip(X, X[1:]):
        fv = [s for s in foreign if xl < _mid(s[0], "v") < xr]
        if not fv:
            continue
        hh = []
        for r, o, real in hs:
            c = fitz.Rect(max(r.x0, xl), r.y0, min(r.x1, xr), r.y1)
            if c.width > 1:
                hh.append((c, o, real))
        items = fv + hh
        comps = _components(items)
        for comp in comps:
            if not any(it in fv for it in comp):
                continue
            nh = len(_cluster([_mid(r, o) for r, o, real in comp if o == "h" and real]))
            if nh < 2:
                continue
            fvs = [it for it in comp if it in fv]
            y0 = min(r.y0 for r, o, _ in fvs)
            y1 = max(r.y1 for r, o, _ in fvs)
            hb = [r for r, o, _ in comp if o == "h" and y0 - TOL <= _mid(r, "h") <= y1 + TOL]
            if not hb:
                continue
            R = fitz.Rect(min(r.x0 for r in hb), min(_mid(r, "h") for r in hb),
                          max(r.x1 for r in hb), max(_mid(r, "h") for r in hb))
            if R.height > 8 and R.width > 20:
                regions.append(R)
    if not regions:
        return segs, [], []
    def inside(r, R):
        """선 조각이 안쪽 표 것인가 — 안쪽 표가 바깥 표 밖으로 삐져나와도(방배동 금융조건) 그 조각까지."""
        Rb = R + (-TOL, -TOL, TOL, TOL)
        if Rb.contains(r):
            return True
        if r.width >= r.height:                                # 가로 조각: 높이가 R 안이고 길이 반 넘게 R 과 겹침
            m = (r.y0 + r.y1) / 2
            ov = min(r.x1, R.x1) - max(r.x0, R.x0)
            return Rb.y0 < m < Rb.y1 and ov > r.width * 0.5
        m = (r.x0 + r.x1) / 2
        ov = min(r.y1, R.y1) - max(r.y0, R.y0)
        return R.x0 + TOL < m and ov > r.height * 0.5 and r.y0 >= Rb.y0 and r.y1 <= Rb.y1 and m > R.x1 - TOL
    outer = [sg for sg in segs if not any(inside(sg[0], R) for R in regions)]
    islands = []
    for R in regions:
        isl = []
        for r, o, real in segs:
            c = r & (R + (-TOL, -TOL, TOL, TOL))
            if c.is_empty and not (R + (-TOL, -TOL, TOL, TOL)).contains(r):
                continue
            if o == "h" and (R.y0 - TOL <= _mid(r, "h") <= R.y1 + TOL):
                c = fitz.Rect(max(r.x0, R.x0), r.y0, min(r.x1, R.x1), r.y1)
            elif o == "v" and (R.x0 - TOL <= _mid(r, "v") <= R.x1 + TOL):
                c = fitz.Rect(r.x0, max(r.y0, R.y0), r.x1, min(r.y1, R.y1))
            else:
                continue
            if (c.width if o == "h" else c.height) > 1:
                isl.append((c, o, real))
        islands.append(isl)
    return outer, islands, regions


def _cells_from_segs(segs, box):
    """선 → 칸들 [(rect, r0, c0, r1, c1)], xs, ys."""
    reals_v = [_mid(r, o) for r, o, real in segs if o == "v" and real]
    reals_h = [_mid(r, o) for r, o, real in segs if o == "h" and real]

    def keep(seg):                                   # 칠 가장자리는 진짜 선 자리와 맞을 때만, 그 선 자리로 붙인다
        r, o, real = seg
        if real:
            return seg
        m = _mid(r, o)
        near = [v for v in (reals_v if o == "v" else reals_h) if abs(m - v) <= TOL]
        if not near:
            return None
        v = min(near, key=lambda q: abs(q - m))
        return ((fitz.Rect(v - 0.2, r.y0, v + 0.2, r.y1) if o == "v" else fitz.Rect(r.x0, v - 0.2, r.x1, v + 0.2)),
                o, False)
    segs = [k for k in (keep(s) for s in segs) if k]
    xs = _cluster([box.x0, box.x1] + [_mid(r, o) for r, o, _ in segs if o == "v"])
    ys = _cluster([box.y0, box.y1] + [_mid(r, o) for r, o, _ in segs if o == "h"])
    xs = [x for x in xs if box.x0 - 2 <= x <= box.x1 + 2]
    ys = [y for y in ys if box.y0 - 2 <= y <= box.y1 + 2]
    vs = [(r, _mid(r, "v")) for r, o, _ in segs if o == "v"]
    hs = [(r, _mid(r, "h")) for r, o, _ in segs if o == "h"]

    def has_v(x, y0, y1):
        cov = sum(max(0.0, min(r.y1, y1) - max(r.y0, y0)) for r, m in vs if abs(m - x) <= TOL)
        return cov >= (y1 - y0) * 0.6

    def has_h(y, x0, x1):
        cov = sum(max(0.0, min(r.x1, x1) - max(r.x0, x0)) for r, m in hs if abs(m - y) <= TOL)
        return cov >= (x1 - x0) * 0.6
    nx, ny = len(xs) - 1, len(ys) - 1
    if nx < 1 or ny < 1:
        return [], xs, ys
    seen = [[False] * nx for _ in range(ny)]
    cells = []
    for j in range(ny):
        for i in range(nx):
            if seen[j][i]:
                continue
            i1 = i
            while i1 + 1 < nx and not seen[j][i1 + 1] and not has_v(xs[i1 + 1], ys[j], ys[j + 1]):
                i1 += 1
            j1 = j
            while j1 + 1 < ny and not any(seen[j1 + 1][c] for c in range(i, i1 + 1)) \
                    and not any(has_h(ys[j1 + 1], xs[c], xs[c + 1]) for c in range(i, i1 + 1)) \
                    and not any(has_v(xs[c + 1], ys[j1 + 1], ys[j1 + 2]) for c in range(i, i1)):
                j1 += 1
            for jj in range(j, j1 + 1):
                for ii in range(i, i1 + 1):
                    seen[jj][ii] = True
            cells.append((fitz.Rect(xs[i], ys[j], xs[i1 + 1], ys[j1 + 1]), j, i, j1 + 1, i1 + 1))
    return cells, xs, ys


# ──────────────────────────────────────────────
# 워드 표 만들기
# ──────────────────────────────────────────────
def _empty_p(line="120"):
    p = etree.Element(qn("w:p"))
    pp = etree.SubElement(p, qn("w:pPr"))
    etree.SubElement(pp, qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0", qn("w:line"): line,
                                           qn("w:lineRule"): "auto"})
    return p


def _is_lone_bullet(r, sp):
    t = "".join(s["text"] for s in sp if s).strip()
    return len(t) == 1 and (t in "•▪■□◦○●➢✓▶▷◆◇◼◻·∙ㆍ" or 0xF000 <= ord(t) <= 0xF0FF)   # '-' 는 표 값(빈 값 표시)이라 뺀다


def _center(r):
    return fitz.Point((r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2)


class _Abort(Exception):
    pass


def _pic_par(part, png, rect, cell):
    import io
    inline = part.new_pic_inline(io.BytesIO(png), int(rect.width * 12700), int(rect.height * 12700))
    p = etree.Element(qn("w:p"))
    pp = etree.SubElement(p, qn("w:pPr"))
    etree.SubElement(pp, qn("w:spacing"), {qn("w:before"): "0", qn("w:after"): "0"})
    lg, rg = rect.x0 - cell.x0, cell.x1 - rect.x1
    jc = "center" if abs(lg - rg) < max(4, cell.width * 0.05) else ("right" if rg < lg else "left")
    etree.SubElement(pp, qn("w:jc"), {qn("w:val"): jc})
    r = etree.SubElement(p, qn("w:r"))
    etree.SubElement(r, qn("w:drawing")).append(inline)
    return p


def _build(page, segs, box, lines, fills, text_lines, old_pr=None, nested=False, pics=(), part=None):
    outer, islands, ibox = _split_islands(segs, box)            # 안쪽 표 안의 안쪽 표도(방배동 금융조건)
    cells, xs, ys = _cells_from_segs(outer, box)
    if len(cells) < 2:
        return None
    rects = [c for c, *_ in cells]
    if ibox:                                                   # 안쪽 표 칠은 바깥 칸 칠로 안 본다
        fills_out = [(r, c) for r, c in fills
                     if not any((r & b).get_area() > r.get_area() * 0.5 for b in ibox if r.intersects(b))]
    else:
        fills_out = fills
    per_cell, inner_of, pic_of = {}, {}, {}
    for k_i, b in enumerate(ibox):
        k = _cell_of(b, rects)
        if k is not None:
            inner_of.setdefault(k, []).append(k_i)
    my_pics = [(r, png) for r, png in pics if box.intersects(r) and (r & box).get_area() > 25]
    for g_i, (r, png) in enumerate(my_pics):
        if any((b + (-2, -2, 2, 2)).contains(r) for b in ibox):
            continue                                           # 안쪽 표가 가진다
        k = next((k for k, cr in enumerate(rects) if (cr + (-2, -2, 2, 2)).contains(r)), None)
        if k is None:
            raise _Abort("그림이 칸 경계에 걸침")                 # 칸에 안 들어가는 그림 → 이 표는 pdf2docx 것 그대로
        pic_of.setdefault(k, []).append(g_i)

    def keyf(pt):
        for r, _ in my_pics:
            if r.contains(pt):
                return ("g", 0)                                # 그림 위 글은 그림이 가진다
        for k_i, b in enumerate(ibox):
            if b.contains(pt):
                return ("i", k_i)                              # 안쪽 표 글은 안쪽 표가 가진다
        for k, cr in enumerate(rects):
            if cr.contains(pt):
                return ("c", k)
        return None
    for key, bx, sp in _pieces(text_lines, keyf):
        if key and key[0] == "c":
            per_cell.setdefault(key[1], []).append((bx, sp))
    # 옆에 글 없는 글머리 혼자(원본 찌꺼기 '•', 내곡동 신탁사·납입비율 칸)는 뺀다 — 칸이 쪼개지거나 빈 줄이 생겼다
    for k, items in per_cell.items():
        per_cell[k] = [(bx, sp) for bx, sp in items
                       if not _is_lone_bullet(bx, sp)
                       or any(o is not bx and abs(o.y0 - bx.y0) < 3 and o.x0 > bx.x0 for o, _ in items)]

    tbl = etree.Element(qn("w:tbl"))
    tpr = copy.deepcopy(old_pr) if old_pr is not None else etree.Element(qn("w:tblPr"))
    for tag in ("w:tblW", "w:tblBorders", "w:tblCellMar", "w:tblLayout", "w:tblLook", "w:tblStyle", "w:jc"):
        for e in tpr.findall(qn(tag)):
            tpr.remove(e)
    lead = [e for e in tpr if etree.QName(e).localname in ("tblStyle", "tblpPr", "tblOverlap", "bidiVisual",
                                                            "tblStyleRowBandSize", "tblStyleColBandSize")]
    tw = etree.Element(qn("w:tblW"), {qn("w:w"): str(_tw(xs[-1] - xs[0])), qn("w:type"): "dxa"})
    tpr.insert(len(lead), tw)
    last = tw
    if nested:                                                 # 안쪽 표는 칸 가운데(양옆 여백 같게, 사용자 20번)
        jc = etree.Element(qn("w:jc"), {qn("w:val"): "center"})
        tw.addnext(jc)
        last = jc
    for name in ("tblCellSpacing", "tblInd", "shd"):
        e = tpr.find(qn("w:" + name))
        if e is not None:
            last = e
    lay = etree.Element(qn("w:tblLayout"), {qn("w:type"): "fixed"})
    last.addnext(lay)
    cm = etree.Element(qn("w:tblCellMar"))
    for side in ("left", "right"):
        etree.SubElement(cm, qn("w:" + side), {qn("w:w"): "0", qn("w:type"): "dxa"})
    lay.addnext(cm)
    tbl.append(tpr)
    grid = etree.SubElement(tbl, qn("w:tblGrid"))
    for a, b in zip(xs, xs[1:]):
        etree.SubElement(grid, qn("w:gridCol"), {qn("w:w"): str(max(1, _tw(b - a)))})
    owner = {}
    for k, (c, r0, c0, r1, c1) in enumerate(cells):
        for rr in range(r0, r1):
            for cc in range(c0, c1):
                owner.setdefault((rr, cc), k)
    nrows, ncols = len(ys) - 1, len(xs) - 1
    for rr in range(nrows):
        tr = etree.SubElement(tbl, qn("w:tr"))
        trp = etree.SubElement(tr, qn("w:trPr"))
        etree.SubElement(trp, qn("w:trHeight"), {qn("w:val"): str(max(10, _tw(ys[rr + 1] - ys[rr]))),
                                                 qn("w:hRule"): "atLeast"})
        cc = 0
        while cc < ncols:
            k = owner.get((rr, cc))
            if k is None:
                rect, r0, c0, r1, c1 = fitz.Rect(xs[cc], ys[rr], xs[cc + 1], ys[rr + 1]), rr, cc, rr + 1, cc + 1
            else:
                rect, r0, c0, r1, c1 = cells[k]
            span = max(1, c1 - cc)
            tc = etree.SubElement(tr, qn("w:tc"))
            tcp = etree.SubElement(tc, qn("w:tcPr"))
            etree.SubElement(tcp, qn("w:tcW"), {qn("w:w"): str(_tw(xs[cc + span] - xs[cc])), qn("w:type"): "dxa"})
            if span > 1:
                etree.SubElement(tcp, qn("w:gridSpan"), {qn("w:val"): str(span)})
            if r1 - r0 > 1:
                etree.SubElement(tcp, qn("w:vMerge"), {qn("w:val"): "restart"} if rr == r0 else {})
            tcp.append(_borders_el(rect, lines))
            f = _fill_of(rect, fills_out)
            if f:
                etree.SubElement(tcp, qn("w:shd"), {qn("w:val"): "clear", qn("w:color"): "auto", qn("w:fill"): f})
            etree.SubElement(tcp, qn("w:vAlign"), {qn("w:val"): "center"})
            if k is not None and rr == r0:
                items = []                                     # 글 문단과 안쪽 표를 위→아래 순서로
                txt = [x for x in per_cell.get(k, []) if not (inner_of.get(k) and _is_lone_bullet(*x))]
                for par in (_paragraphs(txt, rect) if txt else []):
                    items.append((par[0][0].y0, 0, par))
                for k_i in inner_of.get(k, []):
                    items.append((ibox[k_i].y0, 1, k_i))
                for g_i in pic_of.get(k, []):
                    items.append((my_pics[g_i][0].y0, 2, g_i))
                items.sort(key=lambda x: (x[0], x[1]))
                for y, kind, val in items:
                    if kind == 0:
                        tc.append(_par_el(val, rect))
                    elif kind == 2:
                        tc.append(_pic_par(part, my_pics[val][1], my_pics[val][0], rect))
                    else:
                        b = ibox[val]
                        sub = _build(page, islands[val], b, lines, fills,
                                     [x for x in text_lines if b.intersects(x[0])], nested=True,
                                     pics=[g for g in my_pics if (b + (-2, -2, 2, 2)).contains(g[0])], part=part)
                        if sub is not None:
                            tc.append(sub)
                            tc.append(_empty_p("40"))          # 워드: 칸 안 표 뒤에는 문단이 꼭 있어야 한다
            if tc[-1].tag != qn("w:p"):
                tc.append(_empty_p())
            cc += span
    return tbl


def _all_text(el):
    return _norm("".join(t.text or "" for t in el.iter(qn("w:t"))))


def _missing(a, b):
    """a 글자 중 b 에 없는 비율(글자 수 기준)."""
    from collections import Counter
    ca, cb = Counter(a), Counter(b)
    miss = sum(max(0, n - cb[c]) for c, n in ca.items())
    return miss / max(1, len(a))


def _page_pics(page):
    """쪽 그림 [(자리, png)] — 칸에 넣을 것(작은 점선 조각은 선으로 따로 다룸)."""
    out = []
    for info in page.get_image_info():
        r = fitz.Rect(info["bbox"]) & page.rect
        if r.get_area() < 400 or min(r.width, r.height) <= 3.5:
            continue
        out.append((r, page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=r).tobytes("png")))
    return out


def rebuild_page(page, groups, part=None):
    """한 쪽(pdf2docx 구역 묶음)의 표를 원본 선·칠·글로 다시 짠다. 바꾼 표 [(쪽 안 번호, 칸 수)]."""
    tops = [el for grp in groups for el in grp if el.tag == qn("w:tbl")]
    if not tops:
        return []
    try:
        tabs = page.find_tables().tables
    except Exception:
        return []
    pics = _page_pics(page) if part is not None else []
    big_imgs = [r for r, _ in pics] if part is not None else \
        [fitz.Rect(i["bbox"]) for i in page.get_image_info()
         if fitz.Rect(i["bbox"]).get_area() >= 400 and min(fitz.Rect(i["bbox"]).width, fitz.Rect(i["bbox"]).height) > 3.5]
    lines, fills = _page_art(page)
    done, done_boxes = [], []
    for ti, tab in enumerate(tabs):
        tb = fitz.Rect(tab.bbox)
        if tab.row_count < 2 or tb.width < 30:
            continue
        if any((d + (-2, -2, 2, 2)).contains(tb) for d in done_boxes):
            continue                                           # 이미 다시 짠 큰 표 안의 표
        if part is None and any((ir & tb).get_area() > 25 for ir in big_imgs if ir.intersects(tb)):
            continue                                           # 그림을 넣을 수 없으면 그림 든 표는 그대로
        text_lines = _spans(page, tb + (-1, -1, 1, 1))
        in_pic = lambda pt: any(r.contains(pt) for r in big_imgs)
        src_text = _norm("".join(ch for _, sp in text_lines for s in sp for ch, cb in s.get("chars", [])
                                 if tb.contains(_center(cb)) and not in_pic(_center(cb))))
        src_all = _norm("".join(ch for _, sp in text_lines for s in sp for ch, cb in s.get("chars", [])
                                if tb.contains(_center(cb))))
        cands = []
        for t in tops:
            if t.getparent() is None:
                continue
            wt = _all_text(t)
            if not wt or _missing(wt, src_all) >= 0.1:
                continue
            # ★글자 수만 비교하면 '2 주요 금융조건' 제목 상자까지 큰 표 것으로 잡혀 지워졌다(이천) → 칸 글이 통째로 원본 표 안에 있어야
            cell_txt = [x for x in (_norm("".join(w.text or "" for w in tc.iter(qn("w:t"))))
                                    for tc in t.iter(qn("w:tc")) if tc.find(".//" + qn("w:tbl")) is None) if x]
            if not cell_txt:
                continue
            hit = sum(1 for x in cell_txt if x in src_all)
            if hit >= len(cell_txt) * 0.6 and hit > len(cell_txt) / 2:
                cands.append(t)
        if not cands:
            continue
        try:
            new = _build(page, _segs(lines, fills, tb), tb, lines, fills, text_lines,
                         old_pr=cands[0].find(qn("w:tblPr")), pics=pics, part=part)
        except _Abort:
            continue
        if new is None:
            continue
        # 다시 짠 표에서 원본 글이 빠지면 쓰지 않는다(pdf2docx 표와 비교하면 안 된다 — 그쪽이 글을 두 번 넣기도 했다)
        if _missing(src_text, _all_text(new)) > 0.02:
            continue
        done_boxes.append(tb)
        cands[0].addprevious(new)
        for t in cands:
            t.getparent().remove(t)
        done.append((ti + 1, len(new.findall(f".//{qn('w:tc')}"))))
    return done
