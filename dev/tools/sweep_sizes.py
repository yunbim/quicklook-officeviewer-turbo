"""Sweep preview-window sizes and report the resulting content geometry.

The Word preview handler does not lay the page out the way we assumed. A 1130x760 window
gives clipped, oversized text; the very same 1130-wide window at 2600 tall gives a clean
page with visible margins. Something in the handler keys off the window box, so measure it
instead of guessing.

For each size this records the content box (first/last inked column and row), the implied
side margins, mean ink, and how many full-width bands appear - a real Word page gutter shows
up as a band of rows inked edge to edge.

Usage:
    python sweep_sizes.py [file.docx] [--sizes 1130x700,1130x900,...]
"""

import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = sys.executable

sys.path.insert(0, HERE)
from analyze_render import analyse  # noqa: E402

SHOTS = os.path.join(ROOT, "shots")


def _run(cmd):
    res = subprocess.run(cmd, capture_output=True)
    return (res.stdout or b"") + (res.stderr or b"")


def kill_word():
    import time
    _run(["taskkill", "/F", "/IM", "WINWORD.EXE"])
    for _ in range(60):
        out = _run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"])
        if not any(ln.startswith(b'"WINWORD.EXE"') for ln in out.split(b"\r\n")):
            return
        time.sleep(0.25)


def render(fixture, out, size):
    res = subprocess.run(
        [PY, os.path.join(HERE, "render_handler.py"), fixture, out,
         "--size", size, "--wait", "9"],
        capture_output=True,
    )
    return (res.stdout or b"").decode("utf-8", "replace")


def main():
    fixture = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") \
        else os.path.join(ROOT, "fixtures", "view-none.docx")

    sizes = ["1130x700", "1130x900", "1130x1200", "1130x1600", "1130x2000", "1130x2600"]
    if "--sizes" in sys.argv:
        sizes = sys.argv[sys.argv.index("--sizes") + 1].split(",")

    print(f"fixture: {os.path.relpath(fixture, ROOT)}\n")
    print(f"{'window':>12} {'content x':>16} {'Lmargin':>9} {'Rmargin':>9} "
          f"{'content y':>16} {'ink':>8} {'gut':>4}")

    rows = []
    for size in sizes:
        kill_word()
        tag = size.replace("x", "_")
        out = os.path.join(SHOTS, f"sweep-{tag}.png")
        render(fixture, out, size)
        import io
        import contextlib
        buf = io.StringIO()
        with contextlib.redirect_stdout(buf):
            info = analyse(out)
        w, h = info["width"], info["height"]
        if info["left"] is None:
            print(f"{size:>12} {'(blank render)':>16}")
            continue
        lm = info["left"] / w
        rm = (w - 1 - info["right"]) / w
        print(f"{size:>12} {info['left']:>6}..{info['right']:<8} {lm:>8.1%} {rm:>8.1%} "
              f"{info['top']:>6}..{info['bottom']:<8} {info['mean_ink']:>8.4f} "
              f"{len(info['gutters']):>4}")
        rows.append((size, lm, rm, info["mean_ink"]))

    print("\nreading: a *centred* page has Lmargin ~= Rmargin. Lmargin >> Rmargin means the")
    print("page is wider than the window (clipped on the right) - the Web-Layout look.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
