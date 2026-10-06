"""로컬 검증용: PPTX → 슬라이드별 PNG (PowerPoint COM 자동화).
사용: python _render_pptx.py <input.pptx> <out_dir> [slide_start] [slide_end]
"""
import sys, os
import win32com.client


def render(pptx_path, out_dir, s0=1, s1=None):
    os.makedirs(out_dir, exist_ok=True)
    pptx_path = os.path.abspath(pptx_path)
    out_dir = os.path.abspath(out_dir)
    ppt = win32com.client.Dispatch("PowerPoint.Application")
    # WithWindow=False 로 창 없이 열기(일부 버전은 무시). 오류 시 그냥 연다.
    try:
        pres = ppt.Presentations.Open(pptx_path, WithWindow=False)
    except Exception:
        pres = ppt.Presentations.Open(pptx_path)
    n = pres.Slides.Count
    s1 = n if s1 is None else min(s1, n)
    paths = []
    for i in range(s0, s1 + 1):
        p = os.path.join(out_dir, f"slide_{i:02d}.png")
        pres.Slides(i).Export(p, "PNG", 1600, 900)
        paths.append(p)
    pres.Close()
    try:
        ppt.Quit()
    except Exception:
        pass
    print(f"OK: {len(paths)}장 → {out_dir}")
    for p in paths:
        print(" ", p)


if __name__ == "__main__":
    inp = sys.argv[1]
    out = sys.argv[2] if len(sys.argv) > 2 else "_render_out"
    a = int(sys.argv[3]) if len(sys.argv) > 3 else 1
    b = int(sys.argv[4]) if len(sys.argv) > 4 else None
    render(inp, out, a, b)
