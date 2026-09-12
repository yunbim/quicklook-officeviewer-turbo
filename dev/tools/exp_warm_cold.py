"""Controlled experiment: does a *warm* Office host break page layout, and does re-fitting fix it?

Background
----------
A cold Word preview render shows a proper paginated page (paper edge, wide side margin,
no text clipping). A *second* preview - one that rides on the still-running WINWORD.EXE -
looks like a Web Layout: content starts at the very left, overflows the right edge, and the
page structure is gone. QuickLook's own keep-alive/warmup optimisation puts every preview
after the first into exactly that warm state, which is the regression the user is seeing.

This script runs the same file, same window size, four times, varying only the host state:

  1. cold   - WINWORD killed first
  2. warm   - immediately re-render, so the handler reuses the running WINWORD
  3. refit  - warm, but re-apply SetWindow/SetRect after DoPreview (candidate fix)
  4. recold - kill again to prove the first reading was not a fluke

Usage:
    python exp_warm_cold.py [file.docx] [--size WxH]
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
    """tasklist/taskkill speak the console OEM codepage (cp936 here), so never decode as
    UTF-8 - just read the bytes and look for ASCII markers."""
    res = subprocess.run(cmd, capture_output=True)
    return (res.stdout or b"") + (res.stderr or b"")


def word_pids():
    out = _run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"])
    return [ln for ln in out.split(b"\r\n") if ln.startswith(b'"WINWORD.EXE"')]


def kill_word():
    _run(["taskkill", "/F", "/IM", "WINWORD.EXE"])
    # Wait for the handle to actually drop; otherwise the next LocalServer activation can
    # reattach to the dying process and the "cold" reading is silently a warm one.
    import time
    for _ in range(60):
        if not word_pids():
            return
        time.sleep(0.25)


def render(fixture, out, size, refit=False):
    cmd = [PY, os.path.join(HERE, "render_handler.py"), fixture, out, "--size", size,
           "--wait", "7"]
    if refit:
        cmd.append("--refit")
    res = subprocess.run(cmd, capture_output=True)
    for line in (res.stdout or b"").decode("utf-8", "replace").splitlines():
        if line.startswith(("host window", "DoPreview", "captured", "refit", "Initialize")):
            print("     " + line)


def main():
    fixture = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") \
        else os.path.join(ROOT, "fixtures", "view-none.docx")
    size = "1130x760"
    if "--size" in sys.argv:
        size = sys.argv[sys.argv.index("--size") + 1]

    print(f"fixture: {os.path.relpath(fixture, ROOT)}   window: {size}\n")

    steps = [
        ("1 cold", True, False, "cold-start.docx"),
        ("2 warm", False, False, "warm-reuse.docx"),
        ("3 refit", False, True, "warm-refit.docx"),
        ("4 recold", True, False, "cold-again.docx"),
    ]

    results = []
    for label, cold, refit, name in steps:
        if cold:
            kill_word()
        out = os.path.join(SHOTS, "exp-" + name.replace(".docx", ".png"))
        print(f"[{label}] cold_kill={cold} refit={refit}  word_running_before={bool(word_pids())}")
        render(fixture, out, size, refit=refit)
        info = analyse(out, label=f"  -> {label}")
        results.append((label, info))
        print()

    print("=" * 72)
    print(f"{'step':10} {'content x':>16} {'left margin':>12} {'mean ink':>9} {'gutters':>8}")
    for label, info in results:
        if info["left"] is None:
            print(f"{label:10} {'(blank)':>16}")
            continue
        lm = f"{info['left'] / info['width']:.1%}"
        print(f"{label:10} {info['left']:>6}..{info['right']:<8} {lm:>12} "
              f"{info['mean_ink']:>9.4f} {len(info['gutters']):>8}")


if __name__ == "__main__":
    sys.exit(main())
