# -*- coding: utf-8 -*-
"""원본 PDF를 **생긴 그대로** 편집 가능한 PPT 로 옮긴다.

요약본·변환기와 목적이 다르다.
  · 요약본/변환기 : 원본에서 **내용만** 뽑아 회사 양식으로 **다시 그린다**.
  · 여기          : 원본의 **모습 그대로** 옮긴다. 자리·표 색깔·표 선을 원본대로 두고,
                    글꼴만 회사 글꼴(피플폰트)로, 크기는 9pt 로 고정한다.

★표는 **진짜 PPT 표**로 만든다. 글자가 표 칸 안에 들어가야 나중에 고치기 쉽다.

  처음엔 PyMuPDF 가 찾아준 표를 그대로 PPT 표로 만들었다가 크게 깨졌다.
  이 회사 IM 은 **큰 표(구분/내용) 안에 작은 표가 또 들어 있는데**, PPT 표는 중첩이
  안 돼서 작은 표 글자가 큰 칸 하나에 전부 몰리고 행 높이가 폭발했다.
  → 큰 표와 그 안의 작은 표를 **하나의 격자로 합쳐서** 만든다.
     ① 그 자리에 그어진 선을 모두 모아 가로·세로 격자선을 만들고
     ② **사이에 선이 없는 이웃 칸끼리 합친다**(=원본에서 한 칸이었던 것).
     이러면 중첩 없이 표 하나로 원본 모양이 그대로 나온다.
     큰 표의 '구분' 칸이 여러 줄을 먹는 것도, 선이 없으니 저절로 합쳐진다.

  표가 아닌 곳(제목 띠·글머리 문단 등)은 예전처럼 네모 도형 + 글상자로 놓는다.

원본 PDF 의 생김새(실측 — [IM].pdf 12쪽)
  · 그림 요소가 **전부 네모(re)** 다. 선(l)·곡선(c) 이 하나도 없다.
    얇은 네모 = 표의 선, 두꺼운 네모 = 칸 배경.

쌓는 순서(뒤 → 앞) : ① 표 밖 네모  ② 표  ③ 그림  ④ 표 밖 글자
"""
import io
import math
import re

import fitz
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Emu, Pt

# ── 고정 규칙(사용자 확정) ───────────────────────────
FONT_BOLD = "피플폰트 Bold"
FONT_BODY = "피플폰트 Light"
FONT_SIZE = 10.0                # 전부 10pt 로 고정한다(사용자 확정, 2026-09-17)

# ★표 머리글·진하게 칠한 자리는 **전부 이 남색**으로 통일한다(사용자 확정).
#   원본 IM 들을 재보니 칠한 색이 두 무리로 딱 갈린다 — 밝기 0.35 이하(1,365칸,
#   글씨는 거의 다 흰색)와 0.71 이상(연한 배경, 흰 글씨 0칸). 그 사이는 아예 없다.
#   그래서 '진하다' 는 밝기 0.5 로 가른다. 이 남색 위 글씨는 흰색으로 박는다.
NAVY = "08377C"        # 표 머리글·진한 칸 (사용자 확정, 파워포인트에서 고른 남색)
WHITE = "FFFFFF"
LINE = "A5A5A5"        # 표 테두리 — 파워포인트 '회색, 강조 3'. 검정으로 두면 안 된다
_DARK_MAX = 0.5


def _lum(hx):
    """색이 얼마나 밝은가(0=검정, 1=흰색)."""
    try:
        r, g, b = (int(hx[i:i + 2], 16) / 255.0 for i in (0, 2, 4))
    except (ValueError, TypeError):
        return 1.0
    return 0.299 * r + 0.587 * g + 0.114 * b


def _is_dark(hx):
    return bool(hx) and _lum(hx) < _DARK_MAX

# ★피플폰트에 **없는 글자**가 있다(실측: ▪ ㎡ ㈜ – ％ …). 그대로 두면 파워포인트가
#   제멋대로 다른 글꼴로 바꿔 그리고, 그 글꼴은 컴퓨터마다 다르다.
#   → 없는 글자만 **원본과 같은 맑은 고딕으로 콕 집어 지정**한다.
#   ★대체 글꼴은 **한 번으로 부족하다.** 맑은 고딕에도 없는 글자가 있다(실측):
#     ▪(돈암동 글머리)·➢(헌인 글머리)·✓ → 맑은 고딕 ✗ / Segoe UI Symbol ✓
#     ㈜·※·①·㎡      → 맑은 고딕 ✓ / Segoe UI Symbol ✗
#   그래서 **피플폰트 → 맑은 고딕 → Segoe UI Symbol** 순으로 넘긴다.
FONT_FALLBACK = ["맑은 고딕", "Segoe UI Symbol"]
_FONT_FILES = {
    FONT_BODY: "PEOPLEFONTL.TTF",
    FONT_BOLD: "PEOPLEFONTB.TTF",
    "맑은 고딕": r"C:\Windows\Fonts\malgun.ttf",
    "Segoe UI Symbol": r"C:\Windows\Fonts\seguisym.ttf",
}
_FONT_CACHE = {}
_GLYPH_CACHE = {}

# ★두 기준을 갈라야 한다. 하나로 쓰면 **양쪽정렬 문서에서 낱말마다 글상자가 쪼개진다**
#   (헌인마을 IM: 낱말 사이를 공백 글자가 아니라 자리로 벌려 놓음 → 3~6pt).
_SPACE_PT = 2.6     # 이보다 벌어지면 글자 사이에 **띄어쓰기**를 넣는다
# ★같은 줄에 있는 것은 한 글상자로 묶는다. 원본 PDF 는 '※' 와 그 뒤 문장을 **아예 다른
#   조각**으로 담아 놓기도 해서(헌인마을 각주: ※ x48~55, 본문 x66~450, 같은 y),
#   조각마다 글상자를 만들면 각주가 두 조각으로 쪼개진다.
#   이보다 크게 벌어졌을 때만 딴 것으로 본다(좌우 두 단 배치 등).
_GAP_PT = 60.0
_SPACE_W = 2.5      # 9pt 글자에서 띄어쓰기 하나의 대략 폭(pt)
_MIN_RECT = 0.3     # 이보다 작은 네모는 버린다(pt)
THIN_PT = 2.2       # 이보다 얇은 네모는 '선' 으로 본다
_TOL = 2.0          # 같은 격자선으로 볼 좌표 오차(pt)
_COVER = 0.55       # 이 정도는 덮어야 '그 변에 선이 있다' 고 본다


# ── 색 ──────────────────────────────────────────────
def _hex(c):
    """fitz 색(0~1 실수 3개) → 'RRGGBB'. 색이 없으면 None."""
    if not c:
        return None
    try:
        r, g, b = (int(round(max(0.0, min(1.0, v)) * 255)) for v in c[:3])
    except Exception:
        return None
    return "%02X%02X%02X" % (r, g, b)


def _is_pua(text):
    """쓰기 나름 영역(U+E000~F8FF) 글자인가 — 이때만 원래 기호 글꼴을 살린다.

    이 원본의 글머리는 글꼴이 Wingdings 지만 글자는 이미 진짜 ▪(U+25AA) 라,
    Wingdings 로 되돌리면 그 글꼴에 없는 자리라 엉뚱한 기호가 된다(실제로 깨졌음).
    """
    return any(0xE000 <= ord(c) <= 0xF8FF for c in text)


# ── 네모 ────────────────────────────────────────────
def _rects(page):
    """이 쪽에 그려진 네모·선을 **그린 순서대로** 모은다(같은 것은 한 번만).

    ★네모(re)만 읽으면 안 된다. IM 을 만든 프로그램에 따라 표의 선을 **선(l)** 으로
      그리기도 한다(헌인마을 IM: re 316 · l 1,812). 그러면 표 격자를 하나도 못 만들어
      전부 도형+글상자로 나온다. 테두리만 있는 네모도 네 변을 선으로 풀어서 넣는다.
    """
    out, seen = [], set()

    def _add(r, color, filled, line_w, dash=False):
        r = fitz.Rect(r).normalize()
        if (r.width < _MIN_RECT and r.height < _MIN_RECT) or not color:
            return
        key = (round(r.x0, 1), round(r.y0, 1), round(r.x1, 1),
               round(r.y1, 1), color, filled)
        if key in seen:
            return
        seen.add(key)
        out.append({"r": r, "hex": color, "filled": filled,
                    "w": min(r.width, r.height),
                    "thin": min(r.width, r.height) <= THIN_PT,
                    "line_w": line_w, "dash": dash})

    for g in page.get_drawings():
        fill = _hex(g.get("fill"))
        stroke = _hex(g.get("color"))
        width = g.get("width") or 0.0
        # 점선으로 그은 선인가 — 원본 표는 행 구분선을 점선으로 긋는다(57쪽 중 54쪽)
        dashes = (g.get("dashes") or "").strip()
        dash = bool(dashes) and dashes not in ("[] 0", "[]0", "[ ] 0")
        # ★곡선이 섞인 그림에서는 선을 읽지 않는다. 글자를 도형으로 바꿔 놓은 것이라
        #   (헌인마을 IM 의 'Strictly Confidential' = 항목 391개짜리 곡선 그림)
        #   그 직선 조각을 표 선으로 읽으면 낙서가 된다. 그 자리는 그림으로 넣는다.
        artwork = any(it[0] in ("c", "qu") for it in g["items"])
        for it in g["items"]:
            if artwork and it[0] == "l":
                continue
            if it[0] == "re":
                r = fitz.Rect(it[1]).normalize()
                if fill:
                    _add(r, fill, True, width)
                elif stroke:
                    # 테두리만 있는 네모 → 네 변을 각각 선으로
                    w = max(0.3, width or 0.5)
                    _add((r.x0, r.y0 - w / 2, r.x1, r.y0 + w / 2), stroke, True, w, dash)
                    _add((r.x0, r.y1 - w / 2, r.x1, r.y1 + w / 2), stroke, True, w, dash)
                    _add((r.x0 - w / 2, r.y0, r.x0 + w / 2, r.y1), stroke, True, w, dash)
                    _add((r.x1 - w / 2, r.y0, r.x1 + w / 2, r.y1), stroke, True, w, dash)
            elif it[0] == "l":
                p1, p2 = it[1], it[2]
                color = stroke or fill
                if not color:
                    continue
                w = max(0.3, width or 0.5)
                x0, x1 = sorted((p1.x, p2.x))
                y0, y1 = sorted((p1.y, p2.y))
                if x1 - x0 <= 0.8:                      # 세로선
                    _add((x0 - w / 2, y0, x0 + w / 2, y1), color, True, w, dash)
                elif y1 - y0 <= 0.8:                    # 가로선
                    _add((x0, y0 - w / 2, x1, y0 + w / 2), color, True, w, dash)
                # 비스듬한 선은 표 격자가 아니다 → 건너뛴다
    return out


# ── 선 찾기(빠르게) ─────────────────────────────────
class _Lines:
    """얇은 네모를 '가로선/세로선' 으로 갈라 자리별로 담아 둔다.

    쪽마다 네모가 1,000개씩 되므로 매번 전부 훑으면 느리다 → 자리로 미리 나눠 담는다.
    """

    def __init__(self, rects):
        self.v, self.h = {}, {}
        for it in rects:
            if not it["thin"]:
                continue
            r = it["r"]
            if r.width <= THIN_PT and r.height > THIN_PT:          # 세로선
                x = (r.x0 + r.x1) / 2
                self.v.setdefault(int(round(x / _TOL)), []).append(
                    (r.y0, r.y1, it["hex"], max(0.25, r.width),
                     it.get("dash", False)))
            elif r.height <= THIN_PT and r.width > THIN_PT:        # 가로선
                y = (r.y0 + r.y1) / 2
                self.h.setdefault(int(round(y / _TOL)), []).append(
                    (r.x0, r.x1, it["hex"], max(0.25, r.height),
                     it.get("dash", False)))

    @staticmethod
    def _cover(bucket, key, a, b):
        """a~b 중 선이 덮은 비율(0~1)과 (색, 굵기). 도막은 이어 붙여서 잰다."""
        segs = []
        for k in (key - 1, key, key + 1):
            for (p, q, c, w, dsh) in bucket.get(k, ()):
                lo, hi = max(p, a), min(q, b)
                if hi - lo > 0.3:
                    segs.append((lo, hi, c, w, dsh))
        if not segs:
            return 0.0, None
        segs.sort()
        cov, cur = 0.0, None
        for (lo, hi, _c, _w, _d) in segs:
            if cur is None or lo > cur[1]:
                if cur:
                    cov += cur[1] - cur[0]
                cur = [lo, hi]
            else:
                cur[1] = max(cur[1], hi)
        if cur:
            cov += cur[1] - cur[0]
        real = [s for s in segs if s[2].upper() != "FFFFFF"] or segs
        best = max(real, key=lambda s: s[1] - s[0])
        return cov / max(1e-6, b - a), (best[2], best[3], best[4])

    def spine_x(self, box, need=0.95):
        """표 전체를 세로로 가로지르는 '뼈대' 세로선 x 목록."""
        out = []
        for k in self.v:
            x = k * _TOL
            if box.x0 - _TOL <= x <= box.x1 + _TOL:
                c, _ = self._cover(self.v, k, box.y0, box.y1)
                if c >= need:
                    out.append(x)
        return out

    def spine_y(self, box, need=0.95):
        out = []
        for k in self.h:
            y = k * _TOL
            if box.y0 - _TOL <= y <= box.y1 + _TOL:
                c, _ = self._cover(self.h, k, box.x0, box.x1)
                if c >= need:
                    out.append(y)
        return out

    @staticmethod
    def _hit(bucket, key, a, b):
        """a~b 구간에 선이 그어져 있나 → (색, 굵기)

        ★도막을 **이어 붙여서** 재야 한다. 이 원본은 한 줄로 보이는 선을 행마다
          따로 그려 놓는다(예: 32pt 짜리 칸의 옆선이 15.6pt 두 도막). 도막 하나씩만
          보면 '칸 높이의 55%에 못 미친다'고 버려서 테두리가 통째로 빠진다
          (10쪽에서만 43개, 12쪽 합쳐 85개가 이렇게 빠졌다).
        """
        need = (b - a) * _COVER
        segs = []
        for k in (key - 1, key, key + 1):
            for (p, q, c, w, dsh) in bucket.get(k, ()):
                lo, hi = max(p, a), min(q, b)
                if hi - lo > 0.3:
                    segs.append((lo, hi, c, w, dsh))
        if not segs:
            return None
        segs.sort()
        cov, cur = 0.0, None
        for (lo, hi, _c, _w, _d) in segs:
            if cur is None or lo > cur[1]:
                if cur:
                    cov += cur[1] - cur[0]
                cur = [lo, hi]
            else:
                cur[1] = max(cur[1], hi)
        if cur:
            cov += cur[1] - cur[0]
        if cov < need:
            return None
        # ★같은 자리에 선을 **두 번** 그려 놓은 원본이 있다 — 회색 점선(929292 0.48pt)
        #   위에 흰 선(FFFFFF 0pt)을 겹쳐 긋는다(10쪽 자금조달 표). 흰 선을 고르면
        #   테두리가 배경에 묻혀 행 구분선이 통째로 사라진다 → 색 있는 도막을 먼저 본다.
        real = [s for s in segs if s[2].upper() != "FFFFFF"] or segs
        best = max(real, key=lambda s: s[1] - s[0])   # 색·굵기는 가장 긴 도막의 것
        return (best[2], best[3], best[4])

    def vline(self, x, y0, y1):
        """x 자리에 y0~y1 을 반 넘게 지나는 세로선이 있나 → (색, 굵기)"""
        return self._hit(self.v, int(round(x / _TOL)), y0, y1)

    def hline(self, y, x0, x1):
        return self._hit(self.h, int(round(y / _TOL)), x0, x1)

    def xs(self, region):
        out = []
        for k, segs in self.v.items():
            x = k * _TOL
            if region.x0 - _TOL <= x <= region.x1 + _TOL and \
                    any(min(q, region.y1) - max(p, region.y0) > 1
                        for p, q, _, _, _ in segs):
                out.append(x)
        return out

    def ys(self, region):
        out = []
        for k, segs in self.h.items():
            y = k * _TOL
            if region.y0 - _TOL <= y <= region.y1 + _TOL and \
                    any(min(q, region.x1) - max(p, region.x0) > 1
                        for p, q, _, _, _ in segs):
                out.append(y)
        return out


def _axis(values, lo, hi, tol=_TOL):
    """비슷한 좌표를 하나로 묶어 격자선 목록을 만든다(양 끝 포함)."""
    out = []
    for v in sorted(list(values) + [lo, hi]):
        if v < lo - tol or v > hi + tol:
            continue
        if not out or v - out[-1] > tol:
            out.append(v)
    if out:
        out[0], out[-1] = lo, hi
    return out


# ── 표 자리 찾기 ────────────────────────────────────
def _regions(page):
    """표가 있는 자리. 겹치는 것만 합친다(포함 관계는 따로 둔다)."""
    try:
        boxes = [fitz.Rect(t.bbox) for t in
                 page.find_tables(strategy="lines_strict").tables]
    except Exception:
        return []
    changed = True
    while changed:
        changed = False
        out = []
        for b in boxes:
            for i, m in enumerate(out):
                if m.intersects(b) and not (_inside(b, m) or _inside(m, b)):
                    out[i] = m | b
                    changed = True
                    break
            else:
                out.append(b)
        boxes = out
    return boxes


def _inside(r, box, tol=1.5):
    """r 이 box 안에 통째로 들어 있나."""
    return (r.x0 >= box.x0 - tol and r.x1 <= box.x1 + tol
            and r.y0 >= box.y0 - tol and r.y1 <= box.y1 + tol)


def _same_spot(a, b, least=0.8):
    """둘이 **사실상 같은 자리**인가(겹친 넓이 ÷ 합친 넓이).

    '안에 든 표'와 헷갈리면 안 된다 — 안쪽 표는 한쪽만 100% 겹치므로 이 값이
    낮게 나온다(바깥의 78% 를 채우는 안쪽 표도 0.78). 실제로 잡아야 할 쌍둥이는
    1pt 어긋난 정도라 0.99 가 넘는다.
    """
    inter = abs((a & b).get_area())
    if inter <= 0:
        return False
    return inter / (abs(a.get_area()) + abs(b.get_area()) - inter) > least


def _table_groups(page):
    """[(바깥 표, [그 안에 든 표들])] — 표 안의 표를 **따로** 가려낸다.

    ★원본을 만든 파일에서는 안쪽 표가 별개 개체였을 것이다(바깥 표에는 그 자리에
      가로선이 아예 없다 = 통짜 칸 위에 표를 얹은 모양). PDF 에는 '개체' 개념이
      없어 확인할 길이 없지만, 따로 만들어 두면 나중에 따로 옮기고 고칠 수 있다.
    """
    boxes = _regions(page)
    boxes.sort(key=lambda r: -abs(r.get_area()))
    taken, out = [False] * len(boxes), []
    for i, b in enumerate(boxes):
        if taken[i]:
            continue
        taken[i] = True
        inners = []
        for j, c in enumerate(boxes):
            if j == i or taken[j]:
                continue
            if _inside(c, b) and abs(c.get_area()) < abs(b.get_area()) * 0.98:
                taken[j] = True
                inners.append(c)
        out.append((b, inners))
    return out


def _grow(box, rects, limit):
    """안쪽 표 범위를 **맞닿은 선까지 넓힌다.**

    ★PyMuPDF 가 잡아 주는 표 범위가 실제보다 작을 때가 있다
      (4쪽: x 195~387 로 잡혔는데 실제 표는 비고 칸까지 550). 그대로 쓰면 그 바깥 선이
      바깥 표에도 안쪽 표에도 안 들어가 통째로 사라진다(선 검사에서 2,402pt 누락).

    ★단, **표 안쪽에서 뻗어 나가는 선만** 인정한다. 위·아래 테두리에 걸친 선까지
      끌어들이면 바깥 표의 칸막이를 먹어 버린다(3쪽에서 바깥 표의 구분/내용 칸막이
      x=136.8 을 삼켜 표가 통째로 깨졌다).
    """
    box = fitz.Rect(box)
    cap = abs(limit.get_area()) * 0.9
    changed = True
    while changed:
        changed = False
        for it in rects:
            if not it["thin"]:
                continue
            r = it["r"]
            if _inside(r, box):
                continue
            if r.width > r.height:                      # 가로선 → 좌우로 넓힌다
                y = (r.y0 + r.y1) / 2
                if not (box.y0 + 1.0 < y < box.y1 - 1.0):
                    continue                            # 테두리에 걸친 선은 남의 것
                if min(r.x1, box.x1) - max(r.x0, box.x0) < -2.0:
                    continue                            # 닿지도 않는다
            else:                                       # 세로선 → 위아래로 넓힌다
                x = (r.x0 + r.x1) / 2
                if not (box.x0 + 1.0 < x < box.x1 - 1.0):
                    continue
                if min(r.y1, box.y1) - max(r.y0, box.y0) < -2.0:
                    continue
            nb = box | r
            if not _inside(nb, limit) or abs(nb.get_area()) > cap:
                continue        # 바깥 표를 통째로 삼키면 안 된다
            box, changed = nb, True
            break
    return box


def _more_inners(rects, box):
    """PyMuPDF 가 **아예 못 잡은** 안쪽 표를 찾아낸다.

    ★7쪽 '예상 사업수지' 가 그랬다 — find_tables 는 토지현황·사업일정만 잡아 주고
      이 표는 못 잡아서 바깥 표에 통째로 붙어 있었다.
    바깥 표의 **뼈대 선**(표 전체를 가로·세로로 가로지르는 선)만으로 큰 칸을 나누고,
    그 큰 칸 안에 격자를 이룰 만큼 선이 들어 있으면 그게 안쪽 표다.
    (바깥 격자를 통째로 쓰면 안 된다 — 안쪽 표의 선까지 이미 격자선으로 섞여 있어서
     '칸 안쪽' 이라는 것이 없어진다.)
    """
    # ★선이 행마다 도막으로 그려져 있어 '통째로 긴 네모' 를 찾으면 하나도 안 나온다.
    #   도막을 이어 붙여 덮은 비율로 재야 뼈대 선이 잡힌다.
    lines = _Lines(rects)
    xs = _axis(lines.spine_x(box), box.x0, box.x1)
    ys = _axis(lines.spine_y(box), box.y0, box.y1)
    if len(xs) < 3 and len(ys) < 3:
        return []                       # 뼈대가 없으면 나눌 것도 없다
    found = []
    for c in range(len(xs) - 1):
        for r in range(len(ys) - 1):
            cell = fitz.Rect(xs[c], ys[r], xs[c + 1], ys[r + 1])
            # ★큰 칸의 **테두리 선은 빼고** 재야 한다. 안 그러면 범위가 칸 전체가 되어
            #   바깥 표의 칸막이까지 안쪽 표로 딸려 가 표가 깨진다.
            ins = [it for it in rects if it["thin"] and _inside(it["r"], cell)
                   and not _on_cell_edge(it["r"], cell)]
            if len(ins) < 6:
                continue
            n_h = len({round(((i["r"].y0 + i["r"].y1) / 2) / _TOL)
                       for i in ins if i["r"].width > i["r"].height})
            n_v = len({round(((i["r"].x0 + i["r"].x1) / 2) / _TOL)
                       for i in ins if i["r"].height >= i["r"].width})
            # ★세로선을 2개 이상 요구하면 안 된다. 안쪽 표의 좌·우 테두리가 큰 칸의
            #   테두리와 겹치면 안 세어져, 2열짜리 표(신탁사)는 세로선이 1개가 된다.
            if n_h < 2 or n_v < 1:          # 진짜 격자여야 한다
                continue
            bb = fitz.Rect(ins[0]["r"])
            for it in ins[1:]:
                bb = bb | it["r"]
            # ★진짜 안쪽 표는 **자기 머리글**(진하게 칠한 행)을 가지고 있다.
            #   그냥 표의 일부(열 묶음)에는 머리글이 없다. 실측이 아주 깨끗하다 —
            #   7·8·16쪽의 진짜 안쪽 표는 진한 칸이 4~18개, 18쪽 사업수지의 열 묶음
            #   네 곳은 전부 0개다.
            dark = any(f["filled"] and not f["thin"] and _is_dark(f["hex"])
                       and _inside(f["r"], bb) for f in rects)
            # 그래도 큰 칸에 **딱 붙어 있으면** 안쪽 표가 아니다 — 바깥 표의 머리글
            #   행일 뿐이다(38쪽: 사방 여백 0). 진짜 안쪽 표는 좌우가 2pt 이상 떨어져
            #   있다(7·16쪽 2.9·2.5pt, 8쪽은 위아래까지 18·52pt).
            inset = sum(1 for d in (bb.x0 - cell.x0, bb.y0 - cell.y0,
                                    cell.x1 - bb.x1, cell.y1 - bb.y1) if d >= 2.0)
            #   ★머리글로 가려낼 때는 **가로를 거의 다 채워야** 한다. 큰 칸 왼쪽에
            #     라벨 열이 통째로 남아 있으면(이천 3쪽: 왼쪽 여백 90pt, 가로 81%)
            #     그건 그냥 표의 오른쪽 덩어리지 안쪽 표가 아니다.
            if bb.width > 20 and bb.height > 20 and (
                    inset >= 3
                    or (dark and inset >= 2 and bb.width >= cell.width * 0.95)):
                found.append(bb)
    return found


def _on_cell_edge(r, cell, tol=2.0):
    """그 선이 큰 칸의 테두리에 놓여 있나(= 바깥 표의 선이다)."""
    if r.width > r.height:
        y = (r.y0 + r.y1) / 2
        return abs(y - cell.y0) <= tol or abs(y - cell.y1) <= tol
    x = (r.x0 + r.x1) / 2
    return abs(x - cell.x0) <= tol or abs(x - cell.x1) <= tol


def _chart_areas(page):
    """**차트 자리**를 찾는다 — 잘게 그린 곡선이 수백 개씩 몰려 있는 넓은 곳.

    ★차트의 막대·축·눈금·데이터 글자를 하나하나 옮기면 숫자가 뒤엉킨다
      (헌인 24쪽 한 장에 곡선만 6,386개). 차트는 PPT 표·글상자로 되살릴 수 없으니
      그 자리를 원본에서 통째로 오려 그림으로 얹는다.
    ★반드시 **표를 찾기 전에** 걸러야 한다 — 차트의 눈금선을 격자로 보고 표를 만들어
      버리기 때문이다(24쪽 차트 한가운데에 가짜 표가 생겼다).
    ★기준(곡선 100개·100x60pt 이상)은 IM 6개를 전부 재서 정했다. 걸리는 건
      24쪽 차트 하나뿐이고 나머지 5개 문서는 0건이다 — 멀쩡한 도형을 그림으로
      바꿔 버리는 일이 없다.
    """
    try:
        got = page._im_charts
    except AttributeError:
        got = None
    if got is not None:
        return got

    spots = []
    for g in page.get_drawings():
        r = fitz.Rect(g["rect"])
        if r.is_empty:
            continue
        nc = sum(1 for it in g["items"] if it[0] in ("c", "qu"))
        near = fitz.Rect(r.x0 - 12, r.y0 - 12, r.x1 + 12, r.y1 + 12)
        hit = [s for s in spots if s["r"].intersects(near)]
        if hit:
            first = hit[0]
            for s in hit[1:]:               # 여러 덩이를 잇는 도형이면 합친다
                first["r"] |= s["r"]
                first["c"] += s["c"]
                spots.remove(s)
            first["r"] |= r
            first["c"] += nc
        else:
            spots.append({"r": fitz.Rect(r), "c": nc})

    out = [fitz.Rect(s["r"].x0 - 2, s["r"].y0 - 2,
                     s["r"].x1 + 2, s["r"].y1 + 2) & page.rect
           for s in spots
           if s["c"] >= 100 and s["r"].width >= 100 and s["r"].height >= 60]
    try:
        page._im_charts = out
    except Exception:
        pass
    return out


def _drop_backing(page, rects):
    """**그림 뒤에 깔아 놓은 받침판**을 버린다.

    ★원본은 사진·도면 뒤에 그 크기 그대로 색판을 깔아 놓기도 한다(방배동 6쪽:
      도면 두 장 뒤에 노란 판 FFFF00). 원본에서는 그림이 정확히 덮어 **절대 안 보인다.**
      그런데 그걸 칸 배경색으로 잡으면 칸 전체가 노랗게 칠해져, 그림보다 넓은 만큼
      노란 띠가 삐져나온다(사용자 지적).
    ★그림과 서로 90% 넘게 겹치는 **두꺼운 색판**만 버린다. 선(얇은 것)은 표 격자일 수
      있으니 건드리지 않는다. 실측: IM 7개에서 33개가 걸리고 전부 안 보이는 판이다.
    """
    try:
        imgs = [fitz.Rect(x["bbox"]) for x in page.get_image_info()]
    except Exception:
        return rects
    imgs = [r for r in imgs if r.width > 10 and r.height > 10]
    if not imgs:
        return rects
    out = []
    for it in rects:
        r = it["r"]
        if it["thin"] or not it["filled"] or r.width < 10 or r.height < 10:
            out.append(it)
            continue
        hidden = False
        for im in imgs:
            inter = r & im
            if inter.is_empty:
                continue
            if inter.get_area() > r.get_area() * 0.9 \
                    and inter.get_area() > im.get_area() * 0.9:
                hidden = True
                break
        if not hidden:
            out.append(it)
    return out


def _dash_lines(page):
    """**그림 조각으로 그린 파선**을 격자선 후보로 바꾼다.

    ★원본은 강조 테두리를 투명 마스크가 붙은 작은 그림 조각을 줄줄이 늘어놓아
      그리기도 한다(38쪽 '소계' 열 좌우, 조각 100개가 넘는다). 그 자리를 선으로
      세지 않으면 칸이 **옆으로 합쳐져** 옆 열 값까지 한 칸에 딸려 들어간다
      (4Q 칸이 소계·잔액까지 삼켜 '대전도안2- ▲7,474 3,586' 이 한 칸이 됐다).
    """
    out = []
    for im in page.get_images(full=True):
        if not (im[1] if len(im) > 1 else 0):    # 투명 마스크가 있는 조각만
            continue
        try:
            rs = page.get_image_rects(im[0])
        except Exception:
            continue
        for r in rs:
            r = fitz.Rect(r)
            if r.width >= 20 or r.height >= 40 or r.width < 0.3 or r.height < 0.3:
                continue
            if abs(r.width - r.height) < 1.0:
                continue                       # 네모난 조각은 선이 아니다(모서리 등)
            # ★조각 폭이 3pt 쯤이라 그대로는 '선' 으로 안 잡힌다(THIN_PT 2.2).
            #   가운데를 지나는 얇은 선으로 바꿔 넣는다.
            if r.width < r.height:                          # 세로 조각
                cx = (r.x0 + r.x1) / 2
                rr = fitz.Rect(cx - 1.0, r.y0, cx + 1.0, r.y1)
            else:                                           # 가로 조각
                cy = (r.y0 + r.y1) / 2
                rr = fitz.Rect(r.x0, cy - 1.0, r.x1, cy + 1.0)
            out.append({"r": rr, "hex": "000000", "filled": True,
                        "w": 2.0, "thin": True,
                        "line_w": min(r.width, r.height), "dash": True})
    return out


def _banded(page, rects, charts, taken):
    """**세로선이 없어 못 찾은 표**를 찾는다 — 가로선만으로 띠를 이룬 표.

    ★원본에는 세로선 없이 가로선만으로 나뉜 표가 있다(방배동 6쪽 '담보 위치도':
      광역위치도 머리글 + 지도 + 세부위치도 머리글 + 지도). PyMuPDF 의 find_tables 는
      세로선이 없으면 아예 못 찾아서, 머리글 띠는 도형·지도는 그림으로 흩어진다.
    ★그냥 '가로선이 여러 개' 로만 잡으면 **표지의 장식선**까지 표가 된다(돈암동 1쪽).
      그래서 **머리글 띠(진하게 칠한 가로 막대)가 2개 이상** 있는 것만 표로 본다.
      실측: IM 7개에서 10곳이 걸리고 표지·차트는 하나도 안 걸린다.
    """
    hs = [it["r"] for it in rects
          if it["thin"] and it["r"].height <= 2.5 and it["r"].width >= 100]
    bars = [it["r"] for it in rects
            if it["filled"] and not it["thin"] and _is_dark(it["hex"])]
    groups = {}
    for r in hs:
        groups.setdefault((round(r.x0 / 3), round(r.x1 / 3)), []).append(r)

    out = []
    seen = set()
    for v in groups.values():
        if len(v) < 3:
            continue
        ys = sorted({round((r.y0 + r.y1) / 2, 1) for r in v})
        if len(ys) < 3:
            continue
        reg = fitz.Rect(min(r.x0 for r in v), ys[0],
                        max(r.x1 for r in v), ys[-1])
        if reg.height < 40:
            continue
        # ★겹침은 **양쪽으로** 봐야 한다. '이 자리가 이미 표에 반 넘게 덮였나' 만 보면,
        #   작은 표 여러 개를 **통째로 감싸는 큰 띠**가 그냥 통과한다. 그러면 진짜 표
        #   위에 20행짜리 빈 표가 얹혀 화면이 엉망이 된다(방배동 14쪽: 표 3개를 덮는
        #   가짜 표가 둘이나 생겼다).
        #   ★이미 표가 있는 자리와 **조금이라도 겹치면 만들지 않는다.**
        #     실측이 아주 깨끗하다 — IM 7개에서 후보 20여 개 중 **진짜 띠 표 6개는
        #     겹침이 전부 0%** 이고, 나머지는 전부 기존 표 안에 든 '열 묶음'(겹침 38~100%)이다.
        #     예전 기준(50%)으로는 29·38% 짜리 가짜가 통과해, 그게 커져서 20행짜리
        #     빈 표가 진짜 표 셋을 통째로 덮었다(방배동 14쪽).
        if any((reg & b).get_area() > reg.get_area() * 0.05 for b in taken):
            continue
        if any((reg & c).get_area() > reg.get_area() * 0.3 for c in charts):
            continue
        if sum(1 for b in bars
               if _inside(b, reg) and b.width > reg.width * 0.8) >= 2:
            # 같은 자리를 두 번 넣지 않는다(x 묶음이 달라도 범위가 같을 수 있다)
            key = tuple(round(v) for v in reg)
            if key not in seen:
                seen.add(key)
                out.append(reg)
    return out


def _plan(page, txt=()):
    """이 쪽에 그릴 표를 **그릴 순서대로** 늘어놓는다(바깥 먼저, 안쪽 나중).

    안쪽 표에 속한 선·색은 바깥 표에서 빼 둔다. 그래야 바깥은 그 자리가 통짜 빈 칸이 되고,
    안쪽 표를 그 위에 얹을 수 있다.
    """
    rects = _drop_backing(page, _rects(page))
    charts = _chart_areas(page)
    # ★밑줄은 표 격자선이 아니다. 격자로 쓰면 **문단 한가운데서 행이 잘려**, 그 문단
    #   첫 줄만 16pt 짜리 칸에 갇히고 나머지는 다른 칸으로 흩어진다(7쪽 '본건 사업'
    #   밑줄이 y=398 에 행 경계를 만들었다).
    ul = _underlines(rects, txt) if txt else set()
    grid = [it for n, it in enumerate(rects) if n not in ul] + _dash_lines(page)
    plan = []
    groups = list(_table_groups(page))
    # 세로선이 없어 find_tables 가 못 찾은 '가로 띠 표' 도 후보에 넣는다
    groups += [(r, []) for r in
               _banded(page, grid, charts, [b for b, _ in groups])]
    placed = []                     # 자리를 잡은 바깥 표(**넓힌 뒤** 기준)
    for (box, inners) in groups:
        # 차트 자리에 생긴 '표'는 눈금선을 격자로 본 것이다 — 표가 아니다
        if any((box & c).get_area() > box.get_area() * 0.5 for c in charts):
            continue
        # ★바깥 표 범위도 넓혀야 한다. PyMuPDF 가 표를 잘라서 주기도 한다
        #   (천안 3쪽: 실제 표는 x 542 까지인데 456 까지만 잡아 줘 **맨 오른쪽 열이
        #    통째로 빠졌다**).
        box = _grow(box, rects, page.rect)
        # ★같은 표인지는 **넓히고 나서야** 드러난다. find_tables 가 표를 좁게 잡아 줄
        #   때가 있는데(PyMuPDF 1.28: 37쪽 표를 x 41.9~552.6 대신 148.8~505.3 으로
        #   준다), 그러면 남은 왼쪽 띠를 _banded 가 **별개 표**로 잡는다. 둘 다 _grow
        #   로 표 전체까지 넓어져 **같은 자리에 표가 둘** 생기고, 글자는 먼저 그린
        #   쪽이 다 가져가 뒤엣것은 **텅 빈 유령 표**로 남는다(3·10·15·34·37쪽).
        # → 넓힌 뒤 거의 같은 자리면 버린다. '안에 든 표'(겹침이 한쪽만 100%)는
        #   여기서 안 걸린다 — 그건 inners 로 따로 처리한다.
        if any(_same_spot(box, k) for k in placed):
            continue
        placed.append(box)
        inners = [_grow(b, rects, box) for b in inners]
        keep = [it for it in grid if not any(_inside(it["r"], b) for b in inners)]
        o_lines = _Lines(keep)
        o_fills = [x for x in keep if not x["thin"]]
        o_grid = _cells_of(box, o_lines, o_fills, txt, [b for b in inners])
        for _ in range(3):          # 못 잡은 안쪽 표까지 찾아낸다(중첩 대비 몇 번)
            extra = _more_inners(keep, box)
            if not extra:
                break
            inners += extra
            keep = [it for it in grid
                    if not any(_inside(it["r"], b) for b in inners)]
            o_lines = _Lines(keep)
            o_fills = [x for x in keep if not x["thin"]]
            o_grid = _cells_of(box, o_lines, o_fills, txt, [b for b in inners])
        # ★안쪽 표 **안에 또 안쪽 표**가 있을 수 있다(방배동 3쪽: 바깥 표 → 금융조건
        #   내용 블록 → 그 안의 '조달금액' 격자, 세 겹).
        #   ① **큰 것부터 그린다** — 작은(더 안쪽) 표가 위에 얹혀야 안 가려진다.
        #      순서를 안 맞추면 큰 표가 나중에 그려져 안쪽 표를 통째로 덮고,
        #      그 머리글이 텅 빈 것처럼 보인다.
        #   ② 큰 표는 자기 안에 든 표 자리를 **skip 으로 비워 둔다** — 안 그러면
        #      안쪽 표의 글자(구분·대출금·금리·수수료·All-in·LTV)를 큰 표가 가져가
        #      안쪽 표가 빈 채로 남는다.
        inners = sorted(inners, key=lambda b: -b.get_area())

        # ★**표가 실제로 만들어지는 자리만** 안쪽 표로 친다. 격자를 못 만든 자리까지
        #   비워 두면 그 자리 글자가 어느 표에도 못 들어가 표 위에 떠 버린다
        #   (천안 21쪽: 17pt 짜리 한 줄이라 표가 안 됐는데 바깥이 비워 둬서 글자가 떴다).
        ok = []
        for b in inners:
            sub = [it for it in grid if _inside(it["r"], b)]
            if _cells_of(b, _Lines(sub), [x for x in sub if not x["thin"]], txt):
                ok.append(b)
        if len(ok) != len(inners):          # 실패한 자리는 바깥 표가 도로 가져간다
            inners = ok
            keep = [it for it in grid
                    if not any(_inside(it["r"], b) for b in inners)]
            o_lines = _Lines(keep)
            o_fills = [x for x in keep if not x["thin"]]
            o_grid = _cells_of(box, o_lines, o_fills, txt, list(inners))

        subs = []
        for b in inners:
            sub = [it for it in grid if _inside(it["r"], b)]
            s_lines = _Lines(sub)
            deeper = [o for o in inners
                      if o is not b and _inside(o, b) and o.get_area() < b.get_area()]
            s_grid = _cells_of(b, s_lines, [x for x in sub if not x["thin"]],
                               txt, deeper)
            if s_grid:
                subs.append({"grid": s_grid, "lines": s_lines, "box": b,
                             "fills": [x for x in sub if not x["thin"]],
                             "skip": deeper})
        if o_grid:
            plan.append({"grid": o_grid, "lines": o_lines, "box": box,
                         "fills": o_fills,
                         "skip": [s["box"] for s in subs]})
        plan += subs
    # ★'표를 그린 자리' 만 도형에서 뺀다. 표를 못 만든 자리(노란 제목 띠 등)까지
    #   빼 버리면 그 도형이 통째로 사라진다.
    return rects, plan, [t["box"] for t in plan]


def _cells_of(region, lines, fills=(), txt=(), skip=()):
    """격자를 만들고, **사이에 선이 없는 이웃 칸끼리 합쳐** 진짜 칸을 만든다.

    ★칸 경계가 **선이 아니라 배경색으로만** 나뉘어 있는 표가 있다(천안 3쪽 머리글 행:
      남색 배경 1F3864 의 가장자리가 경계이고 선은 아예 없다). 선만 보면 머리글 행이
      아래 행과 합쳐지고, 배경이 칸의 30%밖에 안 덮어 색도 안 칠해진다
      (→ 흰 글자가 흰 바탕에 묻혀 머리글이 통째로 안 보였다).
    → 배경 네모의 가장자리도 격자선 후보로 넣고, **배경색이 다르면 합치지 않는다.**

    반환 (xs, ys, cells) · cells = [(r0, c0, rs, cs)] · 못 만들면 None
    """
    big = [f for f in fills
           if f["r"].width > 3 and f["r"].height > 3 and _inside(f["r"], region)]

    def _changes(v, vertical):
        """그 자리에서 **색이 실제로 바뀌는가.**

        ★원본은 같은 색 배경 띠를 줄마다 따로 그린다(8쪽 '구분' 열의 DBE5F1 이
          411.2~428.5, 428.5~445.7 두 조각). 조각 경계를 격자선으로 넣으면 행이
          엉뚱한 데서 잘리고, 그 아래 글이 **칸 밖으로 밀려나** 표 위에 겹친다
          (분양촉진책 Trigger 행이 490 이 아니라 445.7 에서 끊겼다).
        → 경계 양쪽 색을 실제로 찍어 보고 **다를 때만** 격자선으로 삼는다.
        """
        for f in big:
            r = f["r"]
            p = (r.y0 + r.y1) / 2 if vertical else (r.x0 + r.x1) / 2
            if vertical and abs(r.x0 - v) > 0.6 and abs(r.x1 - v) > 0.6:
                continue
            if not vertical and abs(r.y0 - v) > 0.6 and abs(r.y1 - v) > 0.6:
                continue
            if vertical:
                a = fitz.Rect(v - 2.5, p - 1, v - 0.5, p + 1)
                b = fitz.Rect(v + 0.5, p - 1, v + 2.5, p + 1)
            else:
                a = fitz.Rect(p - 1, v - 2.5, p + 1, v - 0.5)
                b = fitz.Rect(p - 1, v + 0.5, p + 1, v + 2.5)
            if _cell_fill_hex(big, a) != _cell_fill_hex(big, b):
                return True
        return False

    def _axis2(line_pos, fill_pos, lo, hi, vertical):
        """선 자리를 기준으로 삼고, **선에서 멀리 떨어진** 배경 경계만 더한다.

        ★배경 네모의 가장자리는 실제 선에서 조금 안쪽으로 들어가 있을 때가 많다.
          그대로 다 넣으면 얇은 칸이 하나 더 생겨 **줄이 두 겹으로 보인다**(사용자 지적).
          실측: 헌인 4쪽은 선에서 4.7~6.3pt(덧붙은 것), 천안 3쪽 머리글은 21.8pt(진짜 경계).
          → 10pt 넘게 떨어진 것만, 그리고 **색이 실제로 바뀌는 자리만** 진짜 경계로 본다.
        """
        base = _axis(line_pos, lo, hi)
        add = [v for v in fill_pos
               if all(abs(v - a) > 10.0 for a in base) and _changes(v, vertical)]
        return _axis(list(base) + add, lo, hi)

    def _drop_slivers(a, least=3.0):
        """눈에 안 보일 만큼 얇은 칸은 없앤다.

        ★안 없애면 표 맨 아래에 3.2pt 짜리 **빈 행이 하나 더** 생긴다(사용자 지적).
        ★기준을 5pt 로 올려 보았지만 **되돌렸다**(2026-09-17).
          4pt 짜리 빈 행 5개(방배동 17쪽 시공사~사업진행일정 사이 등)는 없어지지만,
          그 행의 **테두리도 같이 빠져** 선이 13개 사라진다(이천 28→35, 넷마블 26→32).
          빈 행 하나보다 빠진 선이 더 눈에 띈다. 열(xs)에 쓰면 더 나쁘다 — 3~5pt 짜리
          좁은 열은 진짜 칸막이라 대전 못넣은선이 0 → 19 로 튄다.
        """
        out = list(a)
        # 가장자리(맨 위·맨 아래)의 얇은 띠는 표 범위를 넓히다 생긴 군더더기다 → 더 넉넉히
        while len(out) > 2 and out[1] - out[0] < 5.0:
            out.pop(1)
        while len(out) > 2 and out[-1] - out[-2] < 5.0:
            out.pop(-2)
        i = 1
        while i < len(out) - 1:
            if out[i + 1] - out[i] < least and len(out) > 2:
                out.pop(i)
                continue
            i += 1
        return out

    xs = _drop_slivers(_axis2(
        lines.xs(region), [f["r"].x0 for f in big] + [f["r"].x1 for f in big],
        region.x0, region.x1, True))
    ys = _drop_slivers(_axis2(
        lines.ys(region), [f["r"].y0 for f in big] + [f["r"].y1 for f in big],
        region.y0, region.y1, False))
    n_col, n_row = len(xs) - 1, len(ys) - 1
    if n_col < 1 or n_row < 1 or n_col * n_row < 2 or n_col * n_row > 4000:
        return None
    # ★1행 2열(라벨 | 내용)짜리도 진짜 표다(헌인 19쪽 '광역 입지 분석').
    #   예전엔 배경 띠 조각이 가짜 행을 만들어 준 덕에 통과했을 뿐이다.
    #   단 **쪽 제목 띠**('1 Executive summary')와 갈라야 한다 — 그것도 선이 있고
    #   두 칸이다. 다른 점은 높이다: 제목 띠는 글 한 줄이라 26pt, 진짜 라벨 표는
    #   내용이 여러 줄이라 130pt.
    if n_row < 2 and ys[-1] - ys[0] < 40.0:
        return None
    if n_col < 2 and xs[-1] - xs[0] < 40.0:
        return None

    # 칸마다 배경색을 미리 구해 둔다(색이 다르면 합치지 않는다)
    bg = {}
    for r in range(n_row):
        for c in range(n_col):
            bg[(r, c)] = _cell_fill_hex(
                big, fitz.Rect(xs[c], ys[r], xs[c + 1], ys[r + 1]))

    # ★어느 칸에 글자가 있는지 미리 표시해 둔다.
    #   글이 있는 행끼리 세로로 합치면, 아래쪽 글을 제자리에 놓으려고 위 여백을 줘야 하는데
    #   파워포인트가 그 여백을 첫 행에 몰아넣어 표가 통째로 늘어난다.
    #   → 합치지 않으면 글이 제 행에 들어가고 여백도 필요 없다(테두리가 없어 보기엔 같다).
    import bisect
    has = {}
    for ln in txt:
        for ch in ln["chars"]:
            if not ch["c"].strip():
                continue
            cx = (ch["bbox"].x0 + ch["bbox"].x1) / 2
            cy = (ch["bbox"].y0 + ch["bbox"].y1) / 2
            if not (xs[0] - 1 <= cx <= xs[-1] + 1 and ys[0] - 1 <= cy <= ys[-1] + 1):
                continue
            c = min(max(bisect.bisect_right(xs, cx) - 1, 0), n_col - 1)
            r = min(max(bisect.bisect_right(ys, cy) - 1, 0), n_row - 1)
            has[(r, c)] = True

    # ★'이어진 것끼리 다 묶기(합집합 찾기)' 로 하면 안 된다. 옆 칸을 통해 **간접적으로
    #   줄줄이 연결**되어, 선이 분명히 있는 자리인데도 합쳐진다(헌인마을 11쪽: 표 오른쪽
    #   끝의 얇은 칸이 다리가 되어 점선 4줄이 통째로 한 칸이 됐다).
    #   → 왼쪽 위부터 훑으며 **오른쪽으로·아래로 갈 수 있는 만큼만** 넓혀 네모를 만든다.
    taken = [[False] * n_col for _ in range(n_row)]
    cells = []
    for r in range(n_row):
        for c in range(n_col):
            if taken[r][c]:
                continue
            cs = 1
            while (c + cs < n_col and not taken[r][c + cs]
                   and bg[(r, c)] == bg[(r, c + cs)]
                   and not lines.vline(xs[c + cs], ys[r], ys[r + 1])):
                cs += 1
            def _in_skip(rr):
                """그 행이 안쪽 표 자리에 걸쳐 있나."""
                cy = (ys[rr] + ys[rr + 1]) / 2
                w = max(1.0, xs[c + cs] - xs[c])
                return any(b.y0 - 1 <= cy <= b.y1 + 1
                           and (min(b.x1, xs[c + cs]) - max(b.x0, xs[c])) > w * 0.5
                           for b in skip)

            rs = 1
            base_skip = _in_skip(r)
            while r + rs < n_row:
                if any(taken[r + rs][k] or bg[(r, c)] != bg[(r + rs, k)]
                       for k in range(c, c + cs)):
                    break
                if any(lines.hline(ys[r + rs], xs[k], xs[k + 1])
                       for k in range(c, c + cs)):
                    break
                # ★안쪽 표가 놓인 자리와 그 밖은 한 칸으로 합치지 않는다.
                #   합치면 칸이 안쪽 표까지 덮어, 그 안의 글이 위로 올라가 표 뒤에 가려진다
                #   (4쪽 각주가 그렇게 사라졌다).
                if _in_skip(r + rs) != base_skip:
                    break
                rs += 1
            for rr in range(r, r + rs):
                for cc in range(c, c + cs):
                    taken[rr][cc] = True
            cells.append((r, c, rs, cs))
    return xs, ys, cells


def _rectangles(members):
    """한 덩어리를 **네모 조각들로 나눈다.** PPT 표는 네모만 합칠 수 있다.

    ★계단 모양이라고 낱칸으로 부수면 안 된다(예전에 그렇게 했다가 5쪽 '인출후행' 칸의
      글이 세로선을 따라 토막났다). 위쪽 줄글은 통짜 한 칸, 아래쪽 작은 표는 제 칸으로
      나뉘는 모양이라 덩어리가 계단이 되는데, 줄마다 이어진 만큼 묶고 위아래로 같은
      범위끼리 다시 묶으면 원본 모양 그대로 네모 조각이 된다.
    """
    by_row = {}
    for (r, c) in members:
        by_row.setdefault(r, []).append(c)
    runs = []                                   # (행, 시작열, 끝열)
    for r in sorted(by_row):
        cols = sorted(by_row[r])
        s = p = cols[0]
        for c in cols[1:]:
            if c == p + 1:
                p = c
                continue
            runs.append((r, s, p))
            s = p = c
        runs.append((r, s, p))
    out, open_runs = [], {}                     # (시작열, 끝열) → (시작행, 끝행)
    for r in sorted({x[0] for x in runs}):
        here = {(s, e) for (rr, s, e) in runs if rr == r}
        for key in list(open_runs):
            if key in here and open_runs[key][1] == r - 1:
                open_runs[key] = (open_runs[key][0], r)
            else:
                r0, r1 = open_runs.pop(key)
                out.append((r0, key[0], r1 - r0 + 1, key[1] - key[0] + 1))
        for key in here:
            open_runs.setdefault(key, (r, r))
    for key, (r0, r1) in open_runs.items():
        out.append((r0, key[0], r1 - r0 + 1, key[1] - key[0] + 1))
    return out


# ── 글자 ────────────────────────────────────────────
def _font_of(part, size_pt):
    """그 조각에 쓸 글꼴 이름과 크기."""
    if _is_pua(part["text"]):
        return part["font"] or FONT_BODY, part["size"]
    return (FONT_BOLD if part["bold"] else FONT_BODY), size_pt


# 단위 표시 — '(단위 : 평, 억원)' '[단위:천원]'
_UNIT_RE = re.compile(r"[\(\[（［]\s*단위")


def _is_unit(chars):
    """단위 표시인가. **표 칸에 넣지 않고 따로 글상자로** 둔다(사용자 확정 규칙).

    원본에서 단위는 표 **오른쪽 위**에 따로 적혀 있다. 표 칸에 넣으면 옆 제목과
    한 문단으로 붙어 버린다('▪ 조합원 모집현황      (단위 : 평, 천원)').
    """
    t = "".join(c["c"] for c in chars).strip()
    m = _UNIT_RE.search(t)
    return bool(m) and m.start() == 0


def _cut_unit(chars):
    """한 줄 안에 제목과 단위가 같이 있으면 **갈라 놓는다.**

    사이가 공백 글자로 채워져 있어 자리로는 안 갈라진다
    ('■ 주요 재무현황            (단위 : 백만원)' 이 한 줄).
    """
    text = "".join(c["c"] for c in chars)
    m = _UNIT_RE.search(text)
    if not m or m.start() == 0:
        return [chars]
    head, tail = chars[:m.start()], chars[m.start():]
    while head and not head[-1]["c"].strip():
        head.pop()
    return [head, tail] if head else [chars]


def _lines_of(page):
    """글자를 줄 단위로, 단 **글자 하나하나의 자리까지** 들고 온다.

    ★왜 글자 단위인가 — PDF 가 옆 칸의 두 값을 **한 조각으로 묶어 놓을 때가 있다**
      (실측: '- 161,634,100' 이 x 251~313 한 조각). 조각째로 칸을 고르면 두 값이
      한 칸에 들어가 위아래로 쌓인다. 글자 단위로 잘라야 각자 제 칸에 들어간다.
    """
    out = []
    for b in page.get_text("rawdict")["blocks"]:
        if b["type"] != 0:
            continue
        for ln in b["lines"]:
            chars = []
            for sp in ln["spans"]:
                name = sp.get("font") or ""
                meta = {"hex": "%06X" % (int(sp.get("color", 0)) & 0xFFFFFF),
                        "bold": ("bold" in name.lower()
                                 or "black" in name.lower()),
                        "font": name,
                        "size": sp.get("size", FONT_SIZE)}
                for ch in sp.get("chars", ()):
                    chars.append(dict(meta, c=ch["c"],
                                      bbox=fitz.Rect(ch["bbox"])))
            while chars and not chars[0]["c"].strip():      # 줄 앞뒤 공백은 버린다
                chars.pop(0)
            # ★줄 끝 공백은 버리되 **있었다는 것은 남긴다.** 원본이 낱말 사이에서
            #   줄을 바꿀 때 그 공백이 줄 끝에 남는다('… 경부 고속도로 및 ').
            #   그냥 버리면 문단을 다시 흘릴 때 '및용인서울' 로 붙어 버린다.
            tail_sp = bool(chars) and not chars[-1]["c"].strip()
            while chars and not chars[-1]["c"].strip():
                chars.pop()
            if not chars:
                continue
            for k, part in enumerate(_cut_unit(chars)):
                out.append({"chars": part, "bbox": _bbox_of(part),
                            "dir": ln.get("dir", (1.0, 0.0)),
                            "tail_sp": tail_sp,
                            "unit": _is_unit(part)})
    return out


def _underlines(rects, txt_lines):
    """밑줄로 그려진 얇은 가로선을 찾아 **글자에 밑줄 서식**으로 붙인다.

    ★도형으로 그대로 그리면 안 된다. 피플폰트가 원본 글꼴보다 좁아 글자는 짧아지는데
      밑줄은 원래 길이 그대로라 **글자 뒤로 길게 삐져나온다**(사용자 지적).
      서식으로 넣으면 글자 길이를 따라간다.
    반환: 그리지 말아야 할 네모의 자리(집합)
    """
    # ★같은 높이로 이어진 토막은 **한 줄로 이어 붙여서** 길이를 잰다.
    #   원본은 표 행 구분선을 두세 토막으로 나눠 그린다(16쪽 공정현황 줄: 42.6~106.2
    #   과 106.7~552.7). 토막 하나만 보면 그게 마침 글자 폭 안에 들어와 **표 선이
    #   글자 밑줄로 바뀐다.** 이어 붙이면 표 폭 전체라 밑줄이 아님을 알 수 있다.
    cand = [(n, it["r"]) for n, it in enumerate(rects)
            if it["thin"] and it["r"].height <= 1.6 and it["r"].width >= 4]
    cand.sort(key=lambda t: ((t[1].y0 + t[1].y1) / 2, t[1].x0))
    groups = []
    for n, r in cand:
        y = (r.y0 + r.y1) / 2
        for g in groups:
            if abs(g["y"] - y) <= 0.8 and r.x0 <= g["x1"] + 1.5 \
                    and r.x1 >= g["x0"] - 1.5:
                g["x0"] = min(g["x0"], r.x0)
                g["x1"] = max(g["x1"], r.x1)
                g["ns"].append(n)
                break
        else:
            groups.append({"y": y, "x0": r.x0, "x1": r.x1, "ns": [n]})

    # 표 격자의 세로선 — 여기에 물려 있는 가로선은 테두리지 밑줄이 아니다
    vlines = [it["r"] for it in rects
              if it["thin"] and it["r"].width <= 1.6 and it["r"].height >= 3]

    skip = set()
    for g in groups:
        # ★안쪽 표 윗변은 그 위 글줄과 **폭이 똑같을 수 있다**(8쪽 분양촉진책:
        #   글줄 110.9~549.8, 표 윗변 110.9~549.7). 폭으로는 못 가른다.
        #   세로선이 이 선 한가운데에 물려 있으면 표 테두리다.
        if any(g["x0"] + 1.0 < (v.x0 + v.x1) / 2 < g["x1"] - 1.0
               and v.y0 - 1.5 <= g["y"] <= v.y1 + 1.5 for v in vlines):
            continue
        for ln in txt_lines:
            b = ln["bbox"]
            # 글줄 바로 밑에 있나 — ★글줄 상자는 글꼴의 내려긋는 부분까지 품고 있어서
            #   밑줄이 상자 **안쪽**에 들어와 있을 때가 있다(7쪽 '본건 사업' 밑줄이
            #   글줄 아래끝보다 0.65pt 위였다). 아래쪽 4분의 1 구간까지 인정한다.
            if not (b.y1 - b.height * 0.25 <= g["y"] <= b.y1 + 2.5):
                continue
            # 표 테두리와 가르려면 **글자 폭 안에** 들어와야 한다
            if g["x0"] < b.x0 - 2.0 or g["x1"] > b.x1 + 2.0:
                continue
            hit = [c for c in ln["chars"]
                   if g["x0"] - 0.5 <= (c["bbox"].x0 + c["bbox"].x1) / 2
                   <= g["x1"] + 0.5]
            if not hit:
                continue
            for c in hit:
                c["under"] = True
            skip.update(g["ns"])
            break
    return skip


def _parts_of(chars):
    """글자들을 색·굵기가 같은 것끼리 묶어 run 으로 만든다.

    ★사이가 벌어진 곳에는 **띄어쓰기를 넣는다.** 이 원본은 낱말 사이를 공백 글자가
      아니라 자리로 벌려 놓은 데가 있어서(예: '돈암동 628'), 그냥 이으면 붙어 버린다.
    """
    parts, prev = [], None
    for c in chars:
        pad = ""
        if prev is not None and c["c"].strip():
            d = c["bbox"].x0 - prev
            if d > _SPACE_PT:
                # 벌어진 만큼 띄어쓰기를 넣는다(하나만 넣으면 '※ 본문' 이 붙어 버린다)
                pad = " " * max(1, min(24, int(round(d / _SPACE_W))))
        key = (c["hex"], c["bold"], c["font"], c["size"], c.get("under", False))
        if parts and parts[-1]["key"] == key:
            parts[-1]["text"] += pad + c["c"]
            parts[-1]["bbox"] = parts[-1]["bbox"] | c["bbox"]
        else:
            parts.append({"key": key, "text": pad + c["c"],
                          "bbox": fitz.Rect(c["bbox"]), "hex": c["hex"],
                          "bold": c["bold"], "font": c["font"],
                          "size": c["size"], "under": c.get("under", False)})
        prev = c["bbox"].x1
    return parts


def _bbox_of(chars):
    bx = chars[0]["bbox"]
    for c in chars[1:]:
        bx = bx | c["bbox"]
    return bx


def _split_chunks(chars):
    """표 밖 글자: **사이가 벌어진 곳**을 끊어 글상자를 나눈다(들여쓰기 보존)."""
    out, cur, prev = [], [], None
    for c in chars:
        if cur and prev is not None and c["c"].strip() and \
                c["bbox"].x0 - prev > _GAP_PT:
            out.append(cur)
            cur = []
        cur.append(c)
        prev = c["bbox"].x1
    if cur:
        out.append(cur)
    return [g for g in out if "".join(c["c"] for c in g).strip()]


def _set_typeface(run, name):
    """한글은 동아시아 글꼴(ea)까지 지정해야 파워포인트가 제대로 쓴다."""
    rPr = run._r.get_or_add_rPr()
    for tag in ("a:latin", "a:ea", "a:cs"):
        for e in rPr.findall(qn(tag)):
            rPr.remove(e)
    for tag in ("a:latin", "a:ea", "a:cs"):
        rPr.append(parse_xml(f'<{tag} {nsdecls("a")} typeface="{name}"/>'))


def _has_glyph(name, ch):
    """그 글꼴에 이 글자가 실제로 들어 있나."""
    if ch.isspace():
        return True
    key = (name, ch)
    if key in _GLYPH_CACHE:
        return _GLYPH_CACHE[key]
    if name not in _FONT_CACHE:
        import os as _os
        rel = _FONT_FILES.get(name, "")
        path = rel if _os.path.isabs(rel) else _os.path.join(
            _os.path.dirname(_os.path.abspath(__file__)), "fonts", rel)
        try:
            _FONT_CACHE[name] = fitz.Font(fontfile=path)
        except Exception:
            _FONT_CACHE[name] = None          # 글꼴 파일이 없으면 검사를 건너뛴다
    font = _FONT_CACHE[name]
    ok = True if font is None else bool(font.has_glyph(ord(ch)))
    _GLYPH_CACHE[key] = ok
    return ok


def _split_by_glyph(text, name):
    """피플폰트에 없는 글자만 떼어 대체 글꼴로 넘긴다 → [(글자들, 글꼴)]"""
    if name not in (FONT_BODY, FONT_BOLD):    # 기호 글꼴 등은 그대로 둔다
        return [(text, name)]

    def _pick(ch):
        if _has_glyph(name, ch):
            return name
        for alt in FONT_FALLBACK:             # 맑은 고딕 → Segoe UI Symbol 순
            if _has_glyph(alt, ch):
                return alt
        return FONT_FALLBACK[0]

    out, cur, cur_font = [], "", None
    for ch in text:
        f = _pick(ch)
        if cur and f != cur_font:
            out.append((cur, cur_font))
            cur = ""
        cur += ch
        cur_font = f
    if cur:
        out.append((cur, cur_font))
    return out


def _para_font_el(p_el, name, size_pt, color=None):
    """문단(XML) 자체의 기본 글꼴 — **글자가 없는 칸도 이걸로 정해 둬야 한다.**

    안 정하면 빈 칸은 테마 기본 글꼴(맑은 고딕/Calibri)이 되고, 거기에 타이핑하면
    엉뚱한 글꼴로 입력된다(실측: 빈 문단 311개).
    """
    sz = int(round(size_pt * 100))
    # 남색 칸은 **빈 칸에도** 흰 글씨를 박아 둔다 — 안 그러면 거기 타이핑할 때
    # 검은 글씨가 나와 남색에 묻힌다.
    faces = (f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else "") + \
            (f'<a:latin typeface="{name}"/><a:ea typeface="{name}"/>'
             f'<a:cs typeface="{name}"/>')
    pPr = p_el.find(qn("a:pPr"))
    if pPr is None:
        pPr = parse_xml(f'<a:pPr {nsdecls("a")}/>')
        p_el.insert(0, pPr)
    for e in pPr.findall(qn("a:defRPr")):
        pPr.remove(e)
    pPr.append(parse_xml(f'<a:defRPr {nsdecls("a")} sz="{sz}">{faces}</a:defRPr>'))
    for e in p_el.findall(qn("a:endParaRPr")):
        p_el.remove(e)
    p_el.append(parse_xml(f'<a:endParaRPr {nsdecls("a")} lang="ko-KR" sz="{sz}">'
                          f'{faces}</a:endParaRPr>'))


def _para_font(para, name, size_pt, color=None):
    _para_font_el(para._p, name, size_pt, color)


def _fit_overflow(table, xs, cells, size_pt):
    """칸 폭을 넘치는 글자를 **그 칸만** 줄여서 한 줄에 들어가게 한다.

    ★파워포인트 표 칸은 `wrap="none"` 을 **무시하고 무조건 줄을 접는다**(글상자와 다르다).
      그래서 글자가 칸보다 넓으면 두 줄이 되고, 그 행이 두 배로 부풀어 표가 통째로 늘어난다.
      9pt 일 때는 대부분 칸에 들어갔는데 10pt 로 올리자 103칸이 새로 넘쳐
      7개 문서에서 표 28개가 부풀었다(천안 16쪽 14pt 행 → 26pt).
    → 넘치는 칸의 글자만 폭에 맞게 낮춘다. 표 전체를 낮추지 않는다.
    """
    n_col = len(xs) - 1
    for (r0, c0, rs, cs) in cells:
        try:
            cell = table.cell(r0, c0)
        except Exception:
            continue
        x0, x1 = xs[c0], xs[min(c0 + cs, n_col)]
        # ★칸 여백을 빼야 한다. 원본 자리를 맞추려고 좌우 여백을 준 칸이 있는데
        #   (오른쪽 정렬 숫자 칸 등) 그걸 안 빼면 '들어간다' 고 잘못 재서 그 칸이
        #   두 줄로 접힌다(헌인 34쪽: 폭 62·여백 3.9 인데 글자 61 → 실제로는 넘침).
        try:
            ml = (cell.margin_left or 0) / 12700.0
            mr = (cell.margin_right or 0) / 12700.0
        except Exception:
            ml = mr = 0.0
        room = (x1 - x0) - ml - mr - 1.5
        if room <= 2:
            continue
        for para in cell.text_frame.paragraphs:
            runs = para.runs
            if not runs:
                continue
            txt = "".join(r.text for r in runs)
            if not txt.strip():
                continue
            cur = runs[0].font.size.pt if runs[0].font.size else size_pt
            w = _text_w(txt, cur)
            if w <= room:
                continue
            # 폭에 맞을 만큼만 낮춘다. 너무 작아지지는 않게 바닥을 둔다.
            small = max(size_pt * 0.72, cur * room / w)
            for r in runs:
                r.font.size = Pt(small)
            _para_font_el(para._p, FONT_BODY, small)


def _fit_thin_rows(gframe, ys, size_pt):
    """원본이 아주 얇게 그린 행을 그 높이 그대로 지킨다.

    ★파워포인트 표의 행 높이에는 **최소값**이 있다 — 글자 크기의 1.2배
      (9pt 면 10.8pt, 10pt 면 12pt). 원본이 그보다 얇게 그린 행은 전부 그 최소값으로
      부풀어 표가 통째로 늘어난다(7쪽에서 빈 행 4개가 정확히 22.5pt 를 밀어 올려
      맨 아랫줄이 푸터를 덮었다).
    → 그런 행은 **그 행 안의 글자만** 행 높이에 맞게 낮춘다. 빈 칸도 글자 든 칸도 똑같이.
      ★글자 든 칸을 빼 두면 안 된다 — 9pt 일 때는 10.8pt 보다 얇은 행이 거의 다 빈 행이라
      티가 안 났지만, 10pt(최소 12pt)로 올리자 **글자 든 얇은 행**이 걸려 7개 문서에서
      표 28개가 부풀었다(최대 +69pt).
    """
    tbl = gframe._element.graphic.graphicData.tbl
    trs = tbl.findall(qn("a:tr"))
    for r, tr in enumerate(trs):
        if r + 1 >= len(ys):
            break
        h = ys[r + 1] - ys[r]
        if h >= size_pt * 1.2:
            continue
        small = max(1.0, h / 1.2 - 0.2)
        sz = str(int(round(small * 100)))
        for tc in tr.findall(qn("a:tc")):
            body = tc.find(qn("a:txBody"))
            if body is None:
                continue
            for p_el in body.findall(qn("a:p")):
                runs = p_el.findall(qn("a:r"))
                if not runs:
                    _para_font_el(p_el, FONT_BODY, small)
                    continue
                # 글자가 있는 칸 — 그 글자들을 행 높이에 맞게 낮춘다
                for r_el in runs:
                    rPr = r_el.find(qn("a:rPr"))
                    if rPr is None:
                        rPr = parse_xml(f'<a:rPr {nsdecls("a")} lang="ko-KR"/>')
                        r_el.insert(0, rPr)
                    rPr.set("sz", sz)
                for e in p_el.findall(qn("a:endParaRPr")):
                    e.set("sz", sz)
                pPr = p_el.find(qn("a:pPr"))
                if pPr is not None:
                    for d in pPr.findall(qn("a:defRPr")):
                        d.set("sz", sz)
                    # 줄 간격도 같이 낮춰야 행이 안 부푼다
                    for ls in pPr.findall(qn("a:lnSpc")):
                        pPr.remove(ls)
                    pPr.insert(0, parse_xml(
                        f'<a:lnSpc {nsdecls("a")}><a:spcPts val="{sz}"/></a:lnSpc>'))


def _fill_para(para, parts, size_pt, force_hex=None):
    base = FONT_BODY
    for part in parts:
        name, size = _font_of(part, size_pt)
        base = name
        for text, face in _split_by_glyph(part["text"], name):
            run = para.add_run()
            run.text = text
            run.font.size = Pt(size)
            run.font.color.rgb = RGBColor.from_string(force_hex or part["hex"])
            if part.get("under"):
                run.font.underline = True
            _set_typeface(run, face)
    _para_font(para, base, size_pt, force_hex)


def _align_of(bbox, x0, x1):
    """글자 한 줄이 칸 안 어디에 붙어 있었는지 보고 정렬을 정한다."""
    left, right = bbox.x0 - x0, x1 - bbox.x1
    wide = bbox.width > (x1 - x0) * 0.85      # 칸을 거의 꽉 채우면 가운데인지 알 수 없다
    if not wide and abs(left - right) <= max(3.0, (x1 - x0) * 0.06):
        return PP_ALIGN.CENTER
    return PP_ALIGN.LEFT if left <= right else PP_ALIGN.RIGHT


def _text_w(text, size_pt):
    """피플폰트로 이 글이 몇 pt 폭인지. 글꼴 파일을 못 읽으면 대략치."""
    f = _FONT_CACHE.get(FONT_BODY)
    if f is None:
        _has_glyph(FONT_BODY, "가")
        f = _FONT_CACHE.get(FONT_BODY)
    if f is None:
        return len(text) * size_pt * 0.6
    try:
        return f.text_length(text, fontsize=size_pt)
    except Exception:
        return len(text) * size_pt * 0.6


# 새 항목이 시작된다는 표시 — 이런 줄은 앞줄에 이어 붙이지 않는다
_ITEM_RE = re.compile(
    r"^\s*(?:[•▪➢✓※◦‣·\-–—]|\(?\d{1,2}\)|\d{1,2}\.|주\s*\d{1,2}\)|"
    r"[①-⑳]|[가-힣]\)|[a-zA-Z]\))\s")


def _join_space(p):
    """문단을 이을 때 **원본 줄 끝에 있던 공백을 되살린다.**

    ★원본은 낱말 사이에서 줄을 바꿀 때 그 공백을 줄 끝에 남긴다('… 경부 고속도로 및 ').
      줄 앞뒤 공백을 버리고 그냥 이어 붙이면 '및용인서울' 로 낱말이 붙어 버린다.
      낱말 중간에서 끊긴 줄('… 쾌적하며 자' + '연 친화적인 …')에는 공백이 없으므로
      그대로 붙는다 — 원본에 있던 대로만 따라간다.
    """
    its = p.get("items") or []
    if not its or not its[-1].get("tail_sp") or not its[-1]["chars"]:
        return
    last = its[-1]["chars"][-1]
    sp = dict(last, c=" ",
              bbox=fitz.Rect(last["bbox"].x1, last["bbox"].y0,
                             last["bbox"].x1 + 1.0, last["bbox"].y1))
    it = dict(its[-1])                      # 원본 조각은 건드리지 않는다
    it["chars"] = it["chars"] + [sp]
    p["items"] = its[:-1] + [it]


def _reflow(rows, x0, x1, size_pt):
    """원본 줄바꿈을 지키지 말고 **다시 흘려야** 하는 문단은 한 문단으로 합친다.

    ★원본은 글자를 오른쪽 끝까지 채워 놓았는데, 피플폰트가 더 좁아 같은 글이 짧게 끝난다.
      원본 줄바꿈을 그대로 지키면 오른쪽에 빈자리가 남아 어색하다(사용자 지적).
    ★단, **줄 수가 늘지 않을 때만** 합친다. 늘면 행이 부풀어 표가 밀린다.
    """
    if len(rows) < 2:
        return rows
    right = max(r["bbox"].x1 for r in rows)
    out = [dict(rows[0])]
    out[-1]["last_x1"] = rows[0]["bbox"].x1
    out[-1]["last_x0"] = rows[0]["bbox"].x0
    for row in rows[1:]:
        p = out[-1]
        # ★'왼쪽이 나란한가' 는 **바로 앞줄**과 견줘야 한다. 누적 상자로 재면
        #   글머리표가 붙은 첫 줄(x=60)이 왼쪽 끝으로 남아, 들여쓴 본문 줄(x=73)이
        #   셋째 줄부터 13pt 어긋난 것으로 잡혀 **문단이 매번 끊겼다**(7·9쪽).
        dx = row["bbox"].x0 - p["last_x0"]
        # ★'아직 아무것도 안 붙인 첫 줄인가' 는 **붙인 줄 수**로 봐야 한다.
        #   조각 수로 세면 굵은 글씨가 섞인 첫 줄이 조각 여럿으로 쪼개져 있어
        #   '첫 줄이 아니다' 가 되고, 글머리표만큼 들여쓴 둘째 줄이 안 붙는다
        #   (7쪽 '…까지 본건 사업' / '의 Tr.B EXIT 분양률…').
        head = p.get("n", 1) == 1
        same_left = abs(dx) <= 2.0 or (head and 0.0 <= dx <= 40.0)
        # 문단은 **앞줄이 오른쪽 끝까지 갔을 때만** 이어진다(누적 상자로 재면
        # 중간에 짧게 끝난 줄을 건너뛰고 엉뚱한 줄까지 이어 붙인다)
        prev_full = p["last_x1"] >= right - 8.0
        # ★다음 줄이 글머리·번호로 시작하면 **새 항목**이다(주1)·주2) 각주 4줄이
        #   한 문단으로 합쳐져 표가 부풀었다). 이어 붙이지 않는다.
        # ★조각 사이는 **띄어서** 이어 봐야 한다. 원본은 '※' 를 본문과 다른 조각으로
        #   담는데, 그냥 붙이면 '※EXIT 분양률…' 이 되어 글머리 규칙(`※ `)에 안 걸리고
        #   앞 각주에 이어 붙는다(4쪽 각주 2·3번이 한 줄로 붙었다).
        starts_item = bool(_ITEM_RE.match(
            " ".join("".join(c["c"] for c in s["chars"]) for s in row["items"])))
        if same_left and prev_full and not starts_item:
            if p.get("n", 1) == 1 and dx > 2.0:
                p["hang"] = dx          # 글머리표 들여쓰기(둘째 줄부터)
            _join_space(p)              # 원본 줄 끝 공백을 되살린다
            p["items"] = p["items"] + row["items"]
            p["bbox"] = p["bbox"] | row["bbox"]
            p["n"] = p.get("n", 1) + 1
            # ★앞줄이 오른쪽 끝에 **딱 맞춰** 끝났으면 원본이 양쪽 정렬한 문단이다.
            #   왼쪽 정렬로 놓으면 오른쪽 끝이 들쭉날쭉해져 원본과 달라 보인다.
            if p["last_x1"] >= right - 2.0:
                p["justify"] = True
            p["last_x1"] = row["bbox"].x1
            p["last_x0"] = row["bbox"].x0
            continue
        out.append(dict(row))
        out[-1]["last_x1"] = row["bbox"].x1
        out[-1]["last_x0"] = row["bbox"].x0
    # 합친 문단이 원래 줄 수 안에 들어가는지 확인 — 넘치면 원래대로 둔다
    for p in out:
        if p.get("n", 1) < 2:
            continue
        txt = "".join(c["c"] for s in p["items"] for c in s["chars"])
        # 여유를 넉넉히 둔다 — 파워포인트의 줄바꿈은 내 계산보다 이르게 일어난다
        avail = max(1.0, (x1 - x0) - 8.0)
        need = int(_text_w(txt, size_pt) * 1.08 / avail) + 1
        if need > p["n"]:
            return rows                 # 줄이 늘어난다 → 원본 줄바꿈 유지
    return out


def _cell_align(boxes, x0, x1):
    """칸 하나의 정렬 — 여러 줄이면 **줄마다 따로 보면 안 된다.**

    글머리 문단처럼 줄마다 길이가 다른 것을 한 줄씩 재면 어떤 줄은 가운데로 튄다.
    → 왼쪽 끝·오른쪽 끝·가운데 중 **가장 가지런한 쪽**을 그 칸의 정렬로 삼는다.
    """
    if len(boxes) < 2:
        return _align_of(boxes[0], x0, x1)
    lefts = [b.x0 for b in boxes]
    rights = [b.x1 for b in boxes]
    mids = [(b.x0 + b.x1) / 2 for b in boxes]
    sl = max(lefts) - min(lefts)
    sr = max(rights) - min(rights)
    sc = max(mids) - min(mids)
    # ★세 값이 똑같으면 '가지런함' 으로는 못 가른다. 세로로 한 글자씩 쌓아 놓은
    #   라벨(사/유/지, 진행/경과)이 그렇다 — 줄마다 폭이 같아 셋 다 0 이 된다.
    #   예전엔 이때 그냥 왼쪽을 골라서 가운데 라벨이 전부 왼쪽으로 붙었다.
    #   → 이럴 땐 **칸 안 어디에 있는지(위치)** 로 정한다.
    if abs(sl - sr) < 0.5 and abs(sl - sc) < 0.5:
        whole = boxes[0]
        for b in boxes[1:]:
            whole = whole | b
        return _align_of(whole, x0, x1)
    best = min(sl, sr, sc)
    if best == sc and sc + 0.5 < min(sl, sr):
        return PP_ALIGN.CENTER
    return PP_ALIGN.LEFT if sl <= sr else PP_ALIGN.RIGHT


# ── 표 칸 테두리 ────────────────────────────────────
_LN_ORDER = ("a:lnL", "a:lnR", "a:lnT", "a:lnB", "a:lnTlToBr", "a:lnBlToTr",
             "a:cell3D", "a:noFill", "a:solidFill", "a:gradFill", "a:blipFill",
             "a:pattFill", "a:grpFill", "a:headers", "a:extLst")


def _set_border(cell, tag, rgb, width_pt, dash=False):
    """칸 한 변의 테두리. rgb 가 None 이면 '선 없음'.

    원본이 점선으로 그은 자리는 점선으로 — IM 은 표 행 구분선을 아주 촘촘한
    점선(0.48/0.48)으로 긋는다(57쪽 중 54쪽).
    """
    tcPr = cell._tc.get_or_add_tcPr()
    for e in tcPr.findall(qn(tag)):
        tcPr.remove(e)
    if rgb is None:
        body, w = "<a:noFill/>", 0
    else:
        body = (f'<a:solidFill><a:srgbClr val="{rgb}"/></a:solidFill>'
                + ('<a:prstDash val="sysDot"/>' if dash else ""))
        w = max(1, int(round(width_pt * 12700)))
    ln = parse_xml(f'<{tag} {nsdecls("a")} w="{w}" cap="flat" cmpd="sng" '
                   f'algn="ctr">{body}</{tag}>')
    tcPr.insert_element_before(ln, *_LN_ORDER[_LN_ORDER.index(tag) + 1:])


def _cell_fill_hex(fills, rect):
    """칸 배경색 — 그 칸을 덮는 두꺼운 네모 중 **마지막에 그린 것**."""
    cx, cy = (rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2
    area = max(1.0, rect.get_area())
    hit = None
    for it in fills:
        r = it["r"]
        if not (r.x0 - 1 <= cx <= r.x1 + 1 and r.y0 - 1 <= cy <= r.y1 + 1):
            continue
        inter = r & rect
        if inter.is_empty or inter.get_area() < area * 0.55:
            continue
        hit = it
    return hit["hex"] if hit else None


def _strip_table_style(gframe):
    """기본 표 스타일(파란 줄무늬)을 걷어낸다. 색은 우리가 원본대로 칠한다."""
    tbl = gframe._element.graphic.graphicData.tbl
    tblPr = tbl.find(qn("a:tblPr"))
    if tblPr is None:
        return
    for e in tblPr.findall(qn("a:tableStyleId")):
        tblPr.remove(e)
    for attr in ("firstRow", "firstCol", "lastRow", "lastCol",
                 "bandRow", "bandCol"):
        tblPr.set(attr, "0")


def _put_table(slide, xs, ys, cells, lines, fills, txt_lines, used, size_pt,
               skip=()):
    """격자 하나를 **진짜 PPT 표**로 만든다. 쓴 글자 줄은 used 에 담는다."""
    n_col, n_row = len(xs) - 1, len(ys) - 1
    gframe = slide.shapes.add_table(
        n_row, n_col, Pt(xs[0]), Pt(ys[0]),
        Pt(xs[-1] - xs[0]), Pt(ys[-1] - ys[0]))
    table = gframe.table
    _strip_table_style(gframe)
    table.first_row = table.first_col = False
    table.last_row = table.last_col = False
    table.horz_banding = table.vert_banding = False
    for c in range(n_col):
        table.columns[c].width = Emu(int(Pt(xs[c + 1] - xs[c])))
    for r in range(n_row):
        table.rows[r].height = Emu(int(Pt(ys[r + 1] - ys[r])))

    for (r0, c0, rs, cs) in cells:                 # 합치기가 먼저다
        if rs > 1 or cs > 1:
            try:
                table.cell(r0, c0).merge(table.cell(r0 + rs - 1, c0 + cs - 1))
            except Exception:
                pass

    navy_at = set()               # 남색으로 칠한 자리(합쳐져 가려진 칸까지)
    for (r0, c0, rs, cs) in cells:
        try:
            cell = table.cell(r0, c0)
        except Exception:
            continue
        x0, y0, x1, y1 = xs[c0], ys[r0], xs[c0 + cs], ys[r0 + rs]
        box = fitz.Rect(x0, y0, x1, y1)

        # ★진하게 칠한 칸(머리글)은 **전부 남색**으로 맞춘다. 연한 색은 원본 그대로.
        rgb = _cell_fill_hex(fills, box)
        ink = None
        if rgb:
            if _is_dark(rgb):
                rgb, ink = NAVY, WHITE
                navy_at.update((r0 + a, c0 + b)
                               for a in range(rs) for b in range(cs))
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor.from_string(rgb)
        else:
            cell.fill.background()

        for side, tag in (("L", "a:lnL"), ("R", "a:lnR"),
                          ("T", "a:lnT"), ("B", "a:lnB")):
            if side == "L":
                got = lines.vline(x0, y0, y1)
            elif side == "R":
                got = lines.vline(x1, y0, y1)
            elif side == "T":
                got = lines.hline(y0, x0, x1)
            else:
                got = lines.hline(y1, x0, x1)
            # ★테두리는 **굵기 0.5pt·색 A5A5A5 로 통일**한다(사용자 확정).
            #   원본 색(검정 등)을 그대로 쓰면 안 된다. 점선 여부만 원본을 따른다.
            _set_border(cell, tag, LINE if got else None, 0.5,
                        bool(got) and len(got) > 2 and got[2])

        tf = cell.text_frame
        # ★칸 안에서 **자동 줄바꿈을 끈다.** 원본의 줄바꿈은 이미 줄마다 문단으로 넣었다.
        #   켜 두면 피플폰트가 원본 글꼴보다 조금 넓어 한 줄이 두 줄로 접히고, 그만큼
        #   행이 부풀어 표가 통째로 늘어난다(헌인마을 103개 표 중 9개, 최대 +244pt).
        tf.word_wrap = False
        cell.margin_left = cell.margin_right = Pt(0)
        cell.margin_top = cell.margin_bottom = Pt(0)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE

        # ★글자 하나하나를 보고 **이 칸 안에 있는 글자만** 가져온다.
        #   줄째로 가져오면 PDF 가 한 조각으로 묶어 놓은 옆칸 값까지 딸려 온다.
        picked = []
        for i, ln in enumerate(txt_lines):
            if ln.get("unit"):          # 단위는 표에 넣지 않는다 — 따로 글상자로
                continue
            b = ln["bbox"]
            if b.y1 < y0 - 1 or b.y0 > y1 + 1 or b.x1 < x0 - 1 or b.x0 > x1 + 1:
                continue
            done = used.setdefault(i, set())
            idx = [k for k, c in enumerate(ln["chars"])
                   if k not in done
                   and x0 - 0.5 <= (c["bbox"].x0 + c["bbox"].x1) / 2 <= x1 + 0.5
                   and y0 - 1 <= (c["bbox"].y0 + c["bbox"].y1) / 2 <= y1 + 1
                   # 안쪽 표 자리는 그 표가 가져간다 — 바깥 표가 삼키면 안 된다
                   and not any(_inside(c["bbox"], s) for s in skip)]
            if idx and "".join(ln["chars"][k]["c"] for k in idx).strip():
                picked.append((i, idx))
        if not picked:
            _para_font(tf.paragraphs[0], FONT_BODY, size_pt, ink)  # 빈 칸도 글꼴 고정
            continue

        # ★글이 칸 위쪽에서 한참 내려와 있으면 **칸에 넣지 않는다.**
        #   파워포인트는 합쳐진 칸의 위 여백을 **첫 행 하나에 몰아넣어서**, 8pt 짜리
        #   행이 98pt 로 부풀고 표가 통째로 늘어난다(헌인 11쪽 +126pt).
        #   그런 글은 제자리에 글상자로 놓는다(빈 칸 위라 보기에는 똑같다).
        #   단, **자리 잡기(위 여백)를 실제로 쓸 때만** 그렇게 한다. 아래에서 판단한다.
        _first_h = ys[min(r0 + 1, n_row)] - ys[r0]

        # 같은 높이에 있는 것끼리 한 문단으로 묶는다(위아래로 쌓이지 않게)
        #   ★아직 '이 글자를 썼다' 고 표시하지 않는다. 칸에 넣기로 확정한 뒤에 표시한다.
        seg = []
        for (i, idx) in picked:
            chars = [txt_lines[i]["chars"][k] for k in idx]
            # 그 줄을 끝까지 가져왔을 때만 '줄 끝 공백' 표시를 물려받는다
            seg.append({"chars": chars, "bbox": _bbox_of(chars),
                        "tail_sp": txt_lines[i].get("tail_sp")
                        and idx[-1] == len(txt_lines[i]["chars"]) - 1})
        # ★같은 높이끼리 줄로 묶되 **바로 앞 줄하고만 견주면 안 된다** — '※' 가 글줄보다
        #   조금 아래라 순서가 뒤집혀 엉뚱한 줄에 붙는다(4쪽 각주가 그렇게 깨졌다).
        seg.sort(key=lambda s: s["bbox"].y0)
        rows = []
        for s in seg:
            b = s["bbox"]
            for row in rows:
                if min(row["bbox"].y1, b.y1) - max(row["bbox"].y0, b.y0) > \
                        min(row["bbox"].height, b.height) * 0.5:
                    row["items"].append(s)
                    row["bbox"] = row["bbox"] | b
                    break
            else:
                rows.append({"cy": (b.y0 + b.y1) / 2, "items": [s],
                             "bbox": fitz.Rect(b)})
        rows.sort(key=lambda r: r["bbox"].y0)
        for row in rows:
            row["items"].sort(key=lambda s: s["bbox"].x0)
            row["cy"] = (row["bbox"].y0 + row["bbox"].y1) / 2

        # 줄 간격은 원본 그대로. ★반드시 못 박아야 한다 — 안 그러면 글꼴 높이에 따라
        #   행이 저절로 커져서 표가 아래로 밀리고 다음 표를 덮는다.
        cys = [r["cy"] for r in rows]
        gg = sorted(cys[k + 1] - cys[k] for k in range(len(cys) - 1))
        gap = gg[len(gg) // 2] if gg else None
        want = gap if (gap and gap > size_pt) else size_pt * 1.25
        # ★그 칸 높이 안에 반드시 들어가게 눌러 준다. 안 그러면 파워포인트가 행을 키워
        #   표가 통째로 늘어나고 아래(푸터·다음 표)를 덮는다(7쪽에서 실제로 그랬다).
        room = max(1.0, (y1 - y0) - 1.0) / len(rows)
        spacing = Pt(max(7.0, min(want, room)))
        align = _cell_align([r["bbox"] for r in rows], x0, x1)
        # ★'한 줄 높이' 는 **다시 흘리기 전에** 재야 한다. 다시 흘리고 나면 여러 줄을
        #   합친 문단 하나가 90pt 짜리 '한 줄' 로 잡혀서, 아래의 세로 자리 판단
        #   (간격이 들쭉날쭉한가·위아래 여백이 다른가)이 통째로 무너진다.
        #   그래서 긴 문단이 든 칸이 전부 가운데로 몰렸다(7·8쪽).
        hs = sorted(r["bbox"].height for r in rows)
        line_h = hs[len(hs) // 2]
        # ★가운데 정렬된 칸(구분 열 라벨 '시공사/롯데건설㈜/자금대여')은 흘리면 안 된다.
        #   일부러 줄을 나눠 놓은 것이라 이어 붙이면 엉뚱한 자리에서 끊긴다.
        flowed = False
        if align == PP_ALIGN.LEFT:
            new_rows = _reflow(rows, x0, x1, size_pt)
            if new_rows is not rows:
                rows = new_rows
                tf.word_wrap = True     # 다시 흘릴 때만 자동 줄바꿈을 켠다
                flowed = True

        # ★여백도 원본대로 준다. 정렬만 맞추고 여백을 아무렇게나 주면 글자가 옆으로
        #   밀린다(45,003 은 오른쪽 정렬이 맞는데도 원본 여백 9.9pt 를 안 줘서 9.9pt 어긋났다).
        wide = max(r["bbox"].width for r in rows)
        room = max(0.0, (x1 - x0) - wide - 1.0)
        if align == PP_ALIGN.LEFT:
            pad = min(max(0.0, min(r["bbox"].x0 for r in rows) - x0), room)
            cell.margin_left = Pt(pad)
        elif align == PP_ALIGN.RIGHT:
            pad = min(max(0.0, x1 - max(r["bbox"].x1 for r in rows)), room)
            cell.margin_right = Pt(pad)
        else:
            pad = 0.0

        # ★줄이 여러 개면 **원본 세로 자리 그대로** 놓는다.
        #   가운데 맞춤 + 일정한 줄 간격으로 두면, 줄 사이에 안쪽 표가 끼어 있는 칸에서
        #   글이 가운데로 몰려 표와 겹친다(3쪽 '차주/자금용도/채권보전/상환재원').
        gaps = [0.0]
        if len(rows) > 1:
            gaps += [max(0.0, rows[k]["bbox"].y0 - rows[k - 1]["bbox"].y1)
                     for k in range(1, len(rows))]
        # ★'간격이 들쭉날쭉한가' 를 간격끼리 비교하면 안 된다 — 줄이 둘뿐이면
        #   간격이 하나라 늘 '고르다' 가 나온다(3쪽 담보현황·금융개요 칸이 그랬다).
        far = len(gaps) > 1 and max(gaps[1:]) > line_h
        # ★**위아래 여백이 크게 다르면** 가운데로 맞추면 안 된다. 글 위(또는 아래)에
        #   안쪽 표가 얹혀 있는 칸이 그렇다. **줄이 하나여도 마찬가지다** — 7쪽
        #   '토지 확보현황' 한 줄이 세로로 긴 칸 가운데로 내려가 표를 덮었다.
        over = rows[0]["bbox"].y0 - y0
        under = y1 - rows[-1]["bbox"].y1
        if abs(over - under) > line_h:
            far = True
        # 간격이 고른 칸은 예전대로(가운데 맞춤) — 그래야 행 높이가 안 늘어난다.
        if far:
            # ★줄 간격은 **글자 한 줄 높이**여야 한다. 빈 자리(안쪽 표가 낀 곳)에서
            #   뽑은 큰 값을 줄 간격으로 쓰면 105pt 같은 값이 나와 글이 엉뚱한 데 놓인다.
            spacing = Pt(max(size_pt * 1.2, line_h))
            top = max(0.0, rows[0]["bbox"].y0 - y0)
            # ★위 여백이 첫 행보다 크면 파워포인트가 **그 행 하나를** 그만큼 키워
            #   표가 통째로 늘어난다. 이럴 땐 여백 대신 **아래쪽 맞춤**으로 놓는다.
            #   (글이 안쪽 표 밑에 있는 칸이라 아래로 붙이면 원본 자리와 거의 같다.)
            if top > _first_h:
                under = max(0.0, y1 - rows[-1]["bbox"].y1)
                _last_h = ys[min(r0 + rs, n_row)] - ys[min(r0 + rs - 1,
                                                          n_row - 1)]
                if under <= _last_h:
                    cell.vertical_anchor = MSO_ANCHOR.BOTTOM
                    cell.margin_bottom = Pt(under)
                    top = 0.0
                else:                       # 그래도 안 되면 제자리에 글상자로
                    _para_font(tf.paragraphs[0], FONT_BODY, size_pt, ink)
                    continue
            else:
                cell.vertical_anchor = MSO_ANCHOR.TOP
            # ★파워포인트는 글자 높이를 내 계산보다 크게 잡는다(첫 줄 여유 등).
            #   딱 맞게 채우면 행이 조금씩 밀려 표가 통째로 늘어난다
            #   (헌인 11쪽 위여백 143pt 짜리 칸 → 표가 +126pt). 여유를 두고 잡는다.
            need = (top + len(rows) * spacing.pt + sum(gaps)) * 1.12 + size_pt
            room = (y1 - y0) - 0.5
            if need > room and need > 0:
                # ★위 여백을 줄이면 글이 **위로 올라가 안쪽 표를 덮는다**
                #   (4쪽 각주: 위여백 146.8 → 128.2pt 로 줄어 표 마지막 행과 겹쳤다).
                #   글이 칸 아래에 붙어 있으면 **아래쪽 맞춤**이 정확하다 — 위 여백을
                #   계산할 필요가 없어 원본 자리를 그대로 지킨다.
                under2 = max(0.0, y1 - rows[-1]["bbox"].y1)
                if under2 * 2 < top:
                    cell.vertical_anchor = MSO_ANCHOR.BOTTOM
                    cell.margin_top = Pt(0)
                    cell.margin_bottom = Pt(under2)
                    top = 0.0
                else:                                   # 넘치면 통째로 줄인다
                    f = room / need
                    top *= f
                    gaps = [g * f for g in gaps]
                    spacing = Pt(max(size_pt * 1.05, spacing.pt * f))
            if top > 0:
                cell.vertical_anchor = MSO_ANCHOR.TOP
                cell.margin_top = Pt(top)
        else:
            gaps = [0.0] * len(rows)
            # ★다시 흘린 문단은 **줄 수가 줄어든다**(피플폰트가 원본보다 좁다).
            #   가운데 맞춤으로 두면 줄어든 만큼 위에 빈자리가 생겨 원본과 달라진다.
            #   원본에서 글이 칸 위에 붙어 있었으면 그대로 위에서부터 놓는다.
            if flowed and over <= line_h * 1.5:
                cell.vertical_anchor = MSO_ANCHOR.TOP
                cell.margin_top = Pt(max(0.0, over))

        for (i, idx) in picked:        # 여기서 확정 — 이 글자들은 칸이 가져간다
            used.setdefault(i, set()).update(idx)

        for k, row in enumerate(rows):
            para = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
            # 원본이 양쪽 정렬한 문단은 양쪽 정렬로 — 오른쪽 끝이 원본처럼 가지런해진다
            para.alignment = PP_ALIGN.JUSTIFY if row.get("justify") else align
            para.line_spacing = spacing
            para.space_before = Pt(gaps[k] if k < len(gaps) else 0.0)
            para.space_after = Pt(0)
            # 줄마다 더 들어간 만큼은 그 줄에만 준다(글머리 문단의 둘째 줄 들여쓰기 등)
            if align == PP_ALIGN.LEFT:
                extra = (row["bbox"].x0 - x0) - pad
                attr = "marL"
            elif align == PP_ALIGN.RIGHT:
                extra = (x1 - row["bbox"].x1) - pad
                attr = "marR"
            else:
                extra, attr = 0.0, None
            # 다시 흘린 문단은 **둘째 줄부터 들여쓰기**(글머리표 옆에 맞춘다)
            hang = row.get("hang", 0.0)
            if hang > 2.0 and align == PP_ALIGN.LEFT:
                pPr = para._p.get_or_add_pPr()
                pPr.set("marL", str(int(round((extra + hang) * 12700))))
                pPr.set("indent", str(-int(round(hang * 12700))))
            elif attr and extra > 0.5:
                pPr = para._p.get_or_add_pPr()
                pPr.set(attr, str(int(round(extra * 12700))))
                pPr.set("indent", "0")
            chars = []
            for s in row["items"]:
                chars.extend(s["chars"])
            _fill_para(para, _parts_of(chars), size_pt, ink)

    # 합쳐져 가려진 칸까지 전부 — 나중에 합치기를 풀어도 글꼴이 안 튀게 한다
    for no, cell in enumerate(table.iter_cells()):
        ink = WHITE if (no // n_col, no % n_col) in navy_at else None
        for para in cell.text_frame.paragraphs:
            if not para.runs:
                _para_font(para, FONT_BODY, size_pt, ink)
    _fit_overflow(table, xs, cells, size_pt)   # 칸 넘치는 글자만 그 칸에서 줄인다
    _fit_thin_rows(gframe, ys, size_pt)
    return gframe


# ── 표 밖 요소 ──────────────────────────────────────
def _merge_rects(items):
    """붙어 있는 **같은 색 네모**를 하나로 합친다.

    ★원본은 배경 띠를 줄마다 따로 그려 놓는다(헌인 3쪽: 회색 F5F5F5 가 13조각).
      그대로 옮기면 클릭할 때마다 빈 상자처럼 잡혀 거슬린다(사용자 지적).
      보이는 모습은 그대로면서 조각 수만 줄인다.
    """
    out = []
    for it in items:
        r = it["r"]
        done = False
        for p in out:
            if p["hex"] != it["hex"] or p["filled"] != it["filled"]:
                continue
            q = p["r"]
            if abs(q.x0 - r.x0) <= 0.6 and abs(q.x1 - r.x1) <= 0.6 \
                    and r.y0 - q.y1 <= 0.6 and q.y0 - r.y1 <= 0.6:
                p["r"] = q | r          # 위아래로 이어진 같은 폭
                done = True
                break
            if abs(q.y0 - r.y0) <= 0.6 and abs(q.y1 - r.y1) <= 0.6 \
                    and r.x0 - q.x1 <= 0.6 and q.x0 - r.x1 <= 0.6:
                p["r"] = q | r          # 좌우로 이어진 같은 높이
                done = True
                break
        if not done:
            out.append(dict(it))
    return out


def _put_rects(slide, rects, regions, skip=(), txt_lines=()):
    def in_region(r):
        cx, cy = (r.x0 + r.x1) / 2, (r.y0 + r.y1) / 2
        return any(g.x0 - 2 <= cx <= g.x1 + 2 and g.y0 - 2 <= cy <= g.y1 + 2
                   for g in regions)

    def header(r):
        """표 밖에 있는 **머리글 띠**인가 — 진하게 칠했고 그 위 글씨가 흰색이다.

        ★진한 네모를 다 남색으로 바꾸면 안 된다. 실측하니 표 밖 진한 네모 456개 중
          420개는 글씨가 없다(차트 막대·빨간 표시·로고). 흰 글씨가 얹힌 36개만 머리글이다.
        """
        got = [c["hex"] for l in txt_lines for c in l["chars"]
               if c["c"].strip()
               and r.x0 <= (c["bbox"].x0 + c["bbox"].x1) / 2 <= r.x1
               and r.y0 <= (c["bbox"].y0 + c["bbox"].y1) / 2 <= r.y1]
        return bool(got) and all(_lum(h) > 0.85 for h in got)

    keep = [it for no, it in enumerate(rects)
            if no not in skip and not in_region(it["r"])]  # 밑줄은 글자 서식으로
    n = 0
    for it in _merge_rects(keep):
        r = it["r"]
        if it["filled"] and _is_dark(it["hex"]) and header(r):
            it["hex"] = NAVY
        sh = slide.shapes.add_shape(
            MSO_SHAPE.RECTANGLE, Pt(r.x0), Pt(r.y0),
            Pt(max(r.width, _MIN_RECT)), Pt(max(r.height, _MIN_RECT)))
        sh.shadow.inherit = False
        if it["filled"]:
            sh.fill.solid()
            sh.fill.fore_color.rgb = RGBColor.from_string(it["hex"])
            sh.line.fill.background()
        else:
            sh.fill.background()
            sh.line.color.rgb = RGBColor.from_string(it["hex"])
            sh.line.width = Pt(max(0.25, it["line_w"] or 0.5))
        tf = sh.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = 0
        tf.margin_top = tf.margin_bottom = 0
        _para_font(tf.paragraphs[0], FONT_BODY, FONT_SIZE)
        n += 1
    return n


def _image_bytes(doc, xref):
    """그림 하나를 PNG/원본 바이트로. 반투명(워터마크)이면 투명도를 살린다."""
    info = doc.extract_image(xref)
    if not info:
        return None
    raw, smask = info.get("image"), info.get("smask")
    if not smask:
        return raw
    try:
        base = fitz.Pixmap(raw)
        if base.colorspace and base.colorspace.n > 3:
            base = fitz.Pixmap(fitz.csRGB, base)
        return fitz.Pixmap(base, fitz.Pixmap(
            doc.extract_image(smask)["image"])).tobytes("png")
    except Exception:
        return raw


def _put_images(slide, page, boxes=(), inside=None, txt_lines=()):
    """★글자 블록에서만 찾으면 kakaopay 워터마크가 통째로 빠진다.

    ★표 안에 든 그림(조감도 등)은 표 **위에**, 그 밖의 그림(지도 등)은 표 **뒤에** 놓는다.
      원본에서 그림 자리가 표 위쪽까지 걸쳐 있을 때가 있는데(천안 3쪽 지도), 전부 표
      위에 놓으면 그 그림이 표 머리글을 덮어 버린다.
    """
    doc = page.parent
    n = 0
    # 이 쪽 그림들이 놓인 자리 — '큰 그림 위에 얹힌 작은 그림' 을 가려내는 데 쓴다
    spots = []
    try:
        spots = [fitz.Rect(x["bbox"]) for x in page.get_image_info()]
    except Exception:
        pass

    def on_big(r):
        """큰 그림 위에 얹힌 작은 라벨 그림인가.

        ★이런 그림은 투명 마스크(smask)로 모양을 내는데, 합성이 어긋나면 **새까맣게**
          나온다(27쪽 지도 위 '워너청담'·'아스턴 55' 라벨). 그 자리를 원본에서 그대로
          오려 넣으면 지도 배경째 정확히 재현된다. 실측: IM 6개 중 27쪽 한 곳뿐이다.
        """
        return r.width < 200 and any(
            o.width > r.width * 1.6 and (o & r).get_area() > r.get_area() * 0.9
            for o in spots)

    for im in page.get_images(full=True):
        xref = im[0]
        smask = im[1] if len(im) > 1 else 0
        try:
            rects = page.get_image_rects(xref)
        except Exception:
            rects = []
        if not rects:
            continue
        data = None
        for r in rects:
            r = fitz.Rect(r)
            if r.width < 1 or r.height < 1:
                continue
            if inside is not None:
                in_tab = any(_inside(r, b, tol=2.0) for b in boxes)
                if in_tab != inside:
                    continue
            # ★그림 **맨 윗줄에 본문 글줄이 얹혀 있으면** 그림 속에 같은 글자가 박혀
            #   있는 것이다 — 원본은 그림을 잘라서(clip) 놓았는데 통째로 넣으면
            #   그림 속 글자와 본문 글상자가 겹쳐 이중으로 보인다(28쪽 '총 29 세대…').
            #   같은 이유로, 글줄이 그림 **위쪽 테두리를 가로질러** 끝나면 원본은 그
            #   글줄 아래부터 그림을 보여 준다(27쪽 지도: 자리는 y=60.8 부터인데 실제로는
            #   제목 '(2) 주요 고급주택 개발 사례' 아래 y=77 부터 그려진다).
            #   ★이 계산은 아래 '표와 겹칠 때 오려 넣기' **앞**에 있어야 한다 —
            #     오려 넣기는 그 자리를 통째로 찍기 때문에 글자까지 함께 찍힌다.
            top_cut = 0.0
            for l in txt_lines:
                b = l["bbox"]
                if b.x1 < r.x0 + 2 or b.x0 > r.x1 - 2:
                    continue
                if 0.5 < b.y1 - r.y0 <= r.height * 0.15 and b.y0 < r.y0 + 2:
                    top_cut = max(top_cut, b.y1 - r.y0 + 1.0)   # 테두리를 가로지름
                elif r.y0 - 1.0 <= b.y0 <= r.y0 + 5.0 \
                        and b.width >= r.width * 0.5:
                    top_cut = max(top_cut, b.y1 - r.y0 + 1.0)   # 그림 속에 박힌 글
            if not 0 < top_cut < r.height * 0.4:
                top_cut = 0.0

            # ★원본은 그림을 잘라서 보여 주기도 한다. 자리만 보고 통째로 넣으면
            #   표 첫 행까지 그림이 걸쳐져 머리글을 덮는다(천안 3쪽 지도).
            #   표와 겹치면 겹치는 만큼 잘라 낸 모습을 원본에서 떠서 넣는다.
            cut = None
            for b in boxes:
                if r.y0 < b.y0 < r.y1 and r.x1 > b.x0 and r.x0 < b.x1:
                    cut = b.y0 if cut is None else min(cut, b.y0)
            if cut is not None and cut - r.y0 > 2:
                try:
                    vis = fitz.Rect(r.x0, r.y0 + top_cut, r.x1, cut)
                    pm = page.get_pixmap(clip=vis, matrix=fitz.Matrix(3, 3))
                    slide.shapes.add_picture(
                        io.BytesIO(pm.tobytes("png")), Pt(vis.x0), Pt(vis.y0),
                        Pt(vis.width), Pt(vis.height))
                    n += 1
                    continue
                except Exception:
                    pass
            # ★투명 마스크(smask)로 모양을 낸 **아주 작은 조각**은 합성이 어긋나면
            #   통째로 시커멓게 찍힌다. 원본이 굵은 파선을 이런 조각을 줄줄이 늘어놓아
            #   그리는데(38쪽 '소계' 열 강조 테두리, 조각 100개가 넘는다), 그러면
            #   대시 사이가 메워져 **실선**처럼 보인다 → 그 자리를 원본에서 오려 넣는다.
            tiny_mask = bool(smask) and r.width < 20 and r.height < 40
            if on_big(r) or tiny_mask:
                try:
                    pm = page.get_pixmap(clip=r, matrix=fitz.Matrix(4, 4))
                    slide.shapes.add_picture(
                        io.BytesIO(pm.tobytes("png")), Pt(r.x0), Pt(r.y0),
                        Pt(r.width), Pt(r.height))
                    n += 1
                    continue
                except Exception:
                    pass
            if data is None:
                data = _image_bytes(doc, xref)
                if not data:
                    break
            try:
                pic = slide.shapes.add_picture(
                    io.BytesIO(data), Pt(r.x0), Pt(r.y0),
                    Pt(r.width), Pt(r.height))
                if top_cut > 0:
                    pic.crop_top = top_cut / r.height
                    pic.top = Pt(r.y0 + top_cut)
                    pic.height = Pt(r.height - top_cut)
                n += 1
            except Exception:
                pass
    return n


def _paragraphs(items):
    """줄들을 **문단 단위로 묶는다.**

    ★이 PDF 는 문단을 안 묶어 준다 — 줄 하나가 덩어리 하나로 온다.
      그대로 두면 세 줄짜리 글머리 문단이 글상자 세 개로 쪼개진다.
      규칙은 실측으로 뚜렷하다 : 같은 문단 안 줄 간격 3~4pt, 문단 사이 21pt.
    """
    items = sorted(items, key=lambda s: (s["bbox"].y0, s["bbox"].x0))
    hs = sorted(s["bbox"].height for s in items)
    h = hs[len(hs) // 2] if hs else 12.0
    limit = h * 0.65
    out = []
    for s in items:
        b = s["bbox"]
        if out and not s.get("unit") and not out[-1][-1].get("unit"):
            prev = out[-1][-1]["bbox"]
            below = b.y0 - prev.y0 > prev.height * 0.5      # 같은 높이 줄은 딴 것
            gap = b.y0 - prev.y1
            over = (min(b.x1, prev.x1) - max(b.x0, prev.x0))
            same_col = over > min(b.width, prev.width) * 0.5
            # ★간격만으로는 못 가른다 — 표지 제목 사이도 본문 줄 사이와 똑같이 벌어져 있다
            #   (둘 다 줄높이의 0.24~0.30배). 확실한 차이는 두 가지다.
            #   ① **제목은 글자가 크다**(표지 26.7pt·23.9pt vs 본문 12~14pt).
            #      큰 글자는 문단이 아니라 제목이므로 묶지 않는다.
            #      (예전엔 '앞줄이 오른쪽 끝까지 꽉 차야 한다' 로 갈랐는데, 그러면 각주처럼
            #       뒷줄이 더 긴 목록이 줄마다 쪼개졌다 — 사용자 지적.)
            #   ② 이어지는 글은 **왼쪽이 나란하다**(67.9 → 67.9). 첫 줄만 글머리표 때문에
            #      조금 왼쪽에서 시작한다(13pt). 가운데 정렬된 표지 글은 33pt씩 어긋난다.
            body = max(prev.height, b.height) <= h * 1.35
            dx = b.x0 - prev.x0
            head = len(out[-1]) == 1
            # 첫 줄만 글머리표 때문에 왼쪽에서 시작한다(실측 13pt·20pt) → 넉넉히 허용
            same_left = abs(dx) <= 2.0 or (head and 0.0 <= dx <= 40.0)
            if below and -1.0 <= gap <= limit and same_col and body and same_left:
                out[-1].append(s)
                continue
        out.append([s])
    return out


def _put_charts(slide, page, areas, txt_lines, used):
    n = 0
    for r in areas:
        try:
            pm = page.get_pixmap(clip=r, matrix=fitz.Matrix(4, 4), alpha=True)
            slide.shapes.add_picture(io.BytesIO(pm.tobytes("png")),
                                     Pt(r.x0), Pt(r.y0), Pt(r.width), Pt(r.height))
            n += 1
        except Exception:
            continue
        for i, l in enumerate(txt_lines):   # 그 안 글자는 이미 그림에 들어 있다
            b = fitz.Rect(l["bbox"])
            if (b & r).get_area() > b.get_area() * 0.5:
                used.setdefault(i, set()).update(range(len(l["chars"])))
    return n


def _put_artwork(slide, page, txt_lines, used=None, charts=()):
    """글자를 도형으로 바꿔 놓은 그림을 **그 자리만 오려 그림으로** 넣는다.

    선으로 다시 그릴 수 없는 것들이다(도장·워터마크 글씨 등). 빼면 사라지고,
    선으로 그리면 낙서가 된다 → 원본을 그 자리만 오려서 그림으로 얹는다.
    차트는 _put_charts 가 통째로 떴으니 여기서는 건너뛴다.
    """
    spots = []
    for g in page.get_drawings():
        if not any(it[0] in ("c", "qu") for it in g["items"]):
            continue
        r = fitz.Rect(g["rect"])
        if r.width < 2 or r.height < 2:
            continue
        if any((r & c).get_area() > r.get_area() * 0.5 for c in charts):
            continue
        for i, s in enumerate(spots):
            if s.intersects(fitz.Rect(r.x0 - 3, r.y0 - 3, r.x1 + 3, r.y1 + 3)):
                spots[i] = s | r
                break
        else:
            spots.append(r)

    n = 0
    for r in spots:
        if any(fitz.Rect(l["bbox"]).intersects(r) for l in txt_lines):
            continue            # 도장 글씨 등은 글자를 두 번 그리게 되므로 건너뛴다
        try:
            pm = page.get_pixmap(clip=r, matrix=fitz.Matrix(4, 4), alpha=True)
            slide.shapes.add_picture(io.BytesIO(pm.tobytes("png")),
                                     Pt(r.x0), Pt(r.y0),
                                     Pt(r.width), Pt(r.height))
            n += 1
        except Exception:
            pass
    return n


def _put_textboxes(slide, txt_lines, used, size_pt):
    """표에 안 들어간 글자를 **문단째로** 글상자에 넣는다."""
    left_lines = []
    for i, line in enumerate(txt_lines):
        done = used.get(i, ())
        rest = [c for k, c in enumerate(line["chars"]) if k not in done]
        if not rest or not "".join(c["c"] for c in rest).strip():
            continue
        left_lines.append({"chars": rest, "bbox": _bbox_of(rest),
                           "dir": line["dir"], "unit": line.get("unit")})

    # ★같은 줄에 있는 조각은 **한 글상자로 묶는다.** 원본 PDF 는 '※' 와 그 뒤 문장을
    #   따로 담아 놓기도 해서, 조각마다 글상자를 만들면 각주가 두 동강 난다.
    # ★먼저 **같은 높이끼리 줄로 묶고** 그 줄 안에서 가로로 줄 세운다.
    #   그냥 (위→아래, 왼→오른쪽) 으로 정렬하면 안 된다 — 글머리표 '➢' 가 글줄보다
    #   1.9pt 아래에 있어서 순서가 뒤집히고, 그 바람에 글머리표가 문단에서 떨어져 나갔다.
    left_lines.sort(key=lambda s: s["bbox"].y0)
    bands = []
    for s in left_lines:
        b = s["bbox"]
        for band in bands:
            if min(band["bbox"].y1, b.y1) - max(band["bbox"].y0, b.y0) > \
                    min(band["bbox"].height, b.height) * 0.5:
                band["items"].append(s)
                band["bbox"] = band["bbox"] | b
                break
        else:
            bands.append({"bbox": fitz.Rect(b), "items": [s]})

    merged = []
    for band in bands:
        row = sorted(band["items"], key=lambda s: s["bbox"].x0)
        cur = None
        for s in row:
            if cur is not None and not s.get("unit") and not cur.get("unit") \
                    and 0 <= s["bbox"].x0 - cur["bbox"].x1 <= _GAP_PT:
                cur["chars"] = cur["chars"] + s["chars"]
                cur["bbox"] = cur["bbox"] | s["bbox"]
                continue
            cur = dict(s)
            merged.append(cur)
    left_lines = merged

    n = 0
    for group in _paragraphs(left_lines):
        # 한 줄짜리는 사이가 벌어진 곳에서 쪼갠다(제목 띠의 번호·오른쪽 단위 표시 등)
        pieces = ([[{"chars": c, "bbox": _bbox_of(c), "dir": group[0]["dir"]}]
                   for c in _split_chunks(group[0]["chars"])]
                  if len(group) == 1 else [group])
        for lines in pieces:
            x0 = min(l["bbox"].x0 for l in lines)
            x1 = max(l["bbox"].x1 for l in lines)
            y0, y1 = lines[0]["bbox"].y0, lines[-1]["bbox"].y1
            # 글꼴이 바뀌면 글자 폭이 조금 달라진다 → 오른쪽에 여유를 준다
            tb = slide.shapes.add_textbox(Pt(x0 - 1.0), Pt(y0 - 1.5),
                                          Pt(x1 - x0 + 14), Pt(y1 - y0 + 3))
            tf = tb.text_frame
            tf.margin_left = tf.margin_right = 0
            tf.margin_top = tf.margin_bottom = 0
            tf.word_wrap = False
            tf.auto_size = MSO_AUTO_SIZE.NONE
            tf.vertical_anchor = MSO_ANCHOR.MIDDLE
            cys = [(l["bbox"].y0 + l["bbox"].y1) / 2 for l in lines]
            gg = sorted(cys[k + 1] - cys[k] for k in range(len(cys) - 1))
            pitch = gg[len(gg) // 2] if gg else None
            for k, l in enumerate(lines):
                para = tf.paragraphs[0] if k == 0 else tf.add_paragraph()
                para.alignment = PP_ALIGN.LEFT
                para.space_before = Pt(0)
                para.space_after = Pt(0)
                if pitch:
                    para.line_spacing = Pt(pitch)
                # 둘째 줄부터 들여쓰기 — 글머리표 옆에 맞춰 붙는 모양을 그대로
                marL = int(round((l["bbox"].x0 - x0) * 12700))
                if marL > 0:
                    pPr = para._p.get_or_add_pPr()
                    pPr.set("marL", str(marL))
                    pPr.set("indent", "0")
                _fill_para(para, _parts_of(l["chars"]), size_pt)
            dx, dy = lines[0]["dir"]
            if abs(dy) > 0.01:
                tb.rotation = (-math.degrees(math.atan2(dy, dx))) % 360
            n += 1
    return n


# ── 한 쪽 옮기기 ────────────────────────────────────
def _put_page(prs, page, size_pt):
    slide = prs.slides.add_slide(prs.slide_layouts[6])     # 6 = 빈 화면
    txt = _lines_of(page)
    rects, plan, boxes = _plan(page, txt)
    ul = _underlines(rects, txt)            # 밑줄 → 글자 서식으로

    # 차트 자리는 통째로 그림이 될 곳이다 — 그 안 도형은 따로 그리지 않는다
    charts = _chart_areas(page)
    if charts:
        ul = set(ul) | {no for no, it in enumerate(rects)
                        if any((it["r"] & c).get_area() > it["r"].get_area() * 0.5
                               or it["r"].get_area() < 1 and c.contains(it["r"].tl)
                               for c in charts)}

    n_rect = _put_rects(slide, rects, boxes, ul, txt)
    n_img = _put_images(slide, page, boxes, False, txt)     # 표 뒤(지도 등)
    used = {}                       # 줄 번호 → 이미 표에 넣은 글자 자리
    n_img += _put_charts(slide, page, charts, txt, used)    # 차트는 통째로 그림
    for t in plan:                  # 바깥 먼저, 안쪽 나중(안쪽이 위에 얹힌다)
        xs, ys, cells = t["grid"]
        _put_table(slide, xs, ys, cells, t["lines"], t["fills"], txt, used,
                   size_pt, skip=t["skip"])
    n_img += _put_images(slide, page, boxes, True, txt)      # 표 위(조감도 등)
    n_img += _put_artwork(slide, page, txt, used, charts)  # 도형으로 된 글씨(도장 등)
    n_txt = _put_textboxes(slide, txt, used, size_pt)
    return {"네모": n_rect, "표": len(plan), "그림": n_img, "글상자": n_txt}


# ── 바깥에서 부르는 것 ──────────────────────────────
def convert(pdf_bytes, out_path, size_pt=FONT_SIZE, pages=None, progress=None):
    """PDF → PPTX. pages 는 1부터 세는 쪽 번호 목록(None 이면 전부)."""
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    try:
        want = list(range(1, doc.page_count + 1)) if not pages else list(pages)
        want = [p for p in want if 1 <= p <= doc.page_count]
        if not want:
            raise ValueError("옮길 쪽이 없습니다.")

        prs = Presentation()
        r0 = doc[want[0] - 1].rect
        prs.slide_width = Emu(int(Pt(r0.width)))
        prs.slide_height = Emu(int(Pt(r0.height)))

        stats = []
        for n, pno in enumerate(want, 1):
            st = _put_page(prs, doc[pno - 1], size_pt)
            st["쪽"] = pno
            stats.append(st)
            if progress:
                progress(n, len(want), pno)
        prs.save(out_path)
        return {"pages": stats, "size": (r0.width / 72.0, r0.height / 72.0),
                "count": len(want)}
    finally:
        doc.close()


if __name__ == "__main__":
    import sys
    src, dst = sys.argv[1], sys.argv[2]
    want = [int(a) for a in sys.argv[3:]] or None
    info = convert(open(src, "rb").read(), dst, pages=want,
                   progress=lambda i, n, p: print(f"  {i}/{n} — {p}쪽"))
    print(f"쪽크기 {info['size'][0]:.2f} x {info['size'][1]:.2f} in")
    for s in info["pages"]:
        print(f"  {s['쪽']:>3}쪽 : 표 {s['표']:>2} · 네모 {s['네모']:>4} · "
              f"그림 {s['그림']:>2} · 글상자 {s['글상자']:>4}")
    print("완료 →", dst)
