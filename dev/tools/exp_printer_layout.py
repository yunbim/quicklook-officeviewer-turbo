"""Does the *default printer* decide how the Word preview lays the page out?

The fixtures proved that w:view / w:zoom have no effect - four different settings.xml files
render pixel-identically. Yet some earlier runs produced a completely different, much better
fitted page. The only machine state that changed between those runs was the default printer
(switched while probing printer DCs). Word's page layout is driven by the printer driver's
DEVMODE, so this is a strong candidate: a virtual printer reporting clean A4 gives a sane
page; a real driver reporting a different / unavailable paper size gives the clipped mess.

Runs the same file at the same size once per printer, killing WINWORD between runs so each
one starts cold with the new default. Leaves the default printer exactly as it found it.

Usage:
    python exp_printer_layout.py [file.docx] [--size 1130x760]
"""

import os
import subprocess
import sys
import time

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)
PY = sys.executable

sys.path.insert(0, HERE)
from analyze_render import analyse  # noqa: E402

SHOTS = os.path.join(ROOT, "shots")

PRINTERS = [
    "Microsoft Print to PDF",
    "HP LaserJet Pro MFP M125-M126 PCLmS",
    "HP Laser MFP 1136-1139 1188",
    "OneNote (Desktop)",
]


def _run(cmd):
    res = subprocess.run(cmd, capture_output=True)
    return (res.stdout or b"") + (res.stderr or b"")


def kill_word():
    _run(["taskkill", "/F", "/IM", "WINWORD.EXE"])
    for _ in range(60):
        out = _run(["tasklist", "/FI", "IMAGENAME eq WINWORD.EXE", "/FO", "CSV", "/NH"])
        if not any(ln.startswith(b'"WINWORD.EXE"') for ln in out.split(b"\r\n")):
            return
        time.sleep(0.25)


def get_default():
    out = subprocess.run([PY, os.path.join(HERE, "set_default_printer.py"), "--list"],
                         capture_output=True)
    for line in (out.stdout or b"").decode("utf-8", "replace").splitlines():
        if line.startswith("default:"):
            return line.split(":", 1)[1].strip().strip("'")
    return None


def set_default(name):
    subprocess.run([PY, os.path.join(HERE, "set_default_printer.py"), "--default", name],
                   capture_output=True)


def main():
    fixture = sys.argv[1] if len(sys.argv) > 1 and not sys.argv[1].startswith("--") \
        else os.path.join(ROOT, "fixtures", "view-none.docx")
    size = "1130x760"
    if "--size" in sys.argv:
        size = sys.argv[sys.argv.index("--size") + 1]

    original = get_default()
    print(f"fixture {os.path.relpath(fixture, ROOT)}  window {size}")
    print(f"default printer before: {original!r}\n")
    print(f"{'printer':40} {'content x':>16} {'Lmargin':>9} {'ink':>8}")
    print("-" * 78)

    try:
        for name in PRINTERS:
            set_default(name)
            actual = get_default()
            if actual != name:
                print(f"{name:40} (could not set, default is {actual!r})")
                continue
            kill_word()
            tag = "".join(c if c.isalnum() else "_" for c in name)[:28]
            out = os.path.join(SHOTS, f"prn-{tag}.png")
            subprocess.run(
                [PY, os.path.join(HERE, "render_handler.py"), fixture, out,
                 "--size", size, "--wait", "8"],
                capture_output=True,
            )
            import contextlib
            import io
            buf = io.StringIO()
            with contextlib.redirect_stdout(buf):
                info = analyse(out)
            if info["left"] is None:
                print(f"{name:40} {'(blank)':>16}")
                continue
            lm = info["left"] / info["width"]
            print(f"{name:40} {info['left']:>6}..{info['right']:<8} {lm:>8.1%} "
                  f"{info['mean_ink']:>8.4f}")
    finally:
        if original:
            set_default(original)
        print(f"\ndefault printer restored to: {get_default()!r}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
