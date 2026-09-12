"""Verify the copy-based Protected View preview, end to end, on the real install.

The behaviour under test (replacing the old confirmation dialog):

  1. A document carrying a Zone.Identifier mark ("downloaded from the Internet") previews
     immediately - no dialog - by being rendered from an unblocked *temporary copy*.
  2. The ORIGINAL file is left completely alone: its Zone.Identifier must still be there
     afterwards. This is the whole point of the change; the old code unblocked it in place.
  3. The preview shows a notice strip saying it is a copy, with a button that unblocks the
     source file on explicit request.
  4. A normal (unmarked) document is completely unaffected: no notice, no temp copy.

Three scenarios, each starting from a cold QuickLook:

  T1  plain document            -> renders, NO notice strip, NO temp copy
  T2  zone-marked document      -> renders, notice strip present, original STILL marked
  T3  (continues from T2)       -> click the button; original becomes unblocked, strip turns green

T3 drives a real mouse click at the button's screen position, located from the captured
pixels - the same pixels the user would look at.

Usage:
    python verify_protected_copy.py
"""

import ctypes
import os
import shutil
import subprocess
import sys
import time
from ctypes import wintypes

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(HERE)

sys.path.insert(0, HERE)
from capture_preview import (  # noqa: E402
    INSTALL_DIR,
    EXE,
    RECT,
    capture,
    kill_quicklook,
    nonwhite_ratio,
    quicklook_pids,
    write_png,
)

USER32 = ctypes.WinDLL("user32", use_last_error=True)
USER32.SetCursorPos.argtypes = [ctypes.c_int, ctypes.c_int]
USER32.mouse_event.argtypes = [
    wintypes.DWORD, wintypes.DWORD, wintypes.DWORD, wintypes.DWORD, ctypes.c_void_p]

CONFIG = os.path.join(INSTALL_DIR, "UserData", "QuickLook.Plugin.OfficeViewer.config")
SHOTS = os.path.join(ROOT, "shots")
ZONE_DIR = os.path.join(ROOT, "zone-test")
ZONE_FILE = os.path.join(ZONE_DIR, "downloaded-from-internet.docx")
PLAIN_FILE = os.path.join(ZONE_DIR, "plain-local.docx")
CACHE_DIR = os.path.join(os.environ.get("TEMP", ""), "QuickLook", "OfficeViewer")

# The banner colours are hard-coded in ProtectedViewBanner.cs; the capture must show them.
STRIP_AMBER = (0xFF, 0xF4, 0xCE)
STRIP_GREEN = (0xE7, 0xF6, 0xE9)
DIALOG_TITLE = "PROTECTED VIEW"

# The window runs at 200% DPI, so the strip sits ~62..147px below the top edge. Scan a band
# that covers it at any DPI scale rather than hard-coding those rows.
TOP_SCAN = 320

WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


# --------------------------------------------------------------------------- windows


def all_windows_of(pid):
    found = []

    def cb(hwnd, _):
        wpid = wintypes.DWORD()
        USER32.GetWindowThreadProcessId(hwnd, ctypes.byref(wpid))
        if wpid.value == pid:
            n = USER32.GetWindowTextLengthW(hwnd)
            title = ctypes.create_unicode_buffer(n + 1)
            USER32.GetWindowTextW(hwnd, title, n + 1)
            cls = ctypes.create_unicode_buffer(256)
            USER32.GetClassNameW(hwnd, cls, 256)
            r = RECT()
            USER32.GetWindowRect(hwnd, ctypes.byref(r))
            found.append({
                "hwnd": hwnd, "title": title.value, "class": cls.value,
                "visible": bool(USER32.IsWindowVisible(hwnd)),
                "left": r.left, "top": r.top,
                "w": r.right - r.left, "h": r.bottom - r.top,
            })
        return True

    USER32.EnumWindows(WNDENUMPROC(cb), 0)
    return found


def find_dialogs():
    seen = []
    for pid in quicklook_pids():
        for w in all_windows_of(pid):
            if w["visible"] and (w["class"] == "#32770"
                                 or DIALOG_TITLE in (w["title"] or "").upper()):
                seen.append(w)
    return seen


def wait_preview(timeout=18.0):
    """Poll for the settled preview window, watching for a dialog at the same time."""
    deadline = time.time() + timeout
    dialogs = []
    while time.time() < deadline:
        dialogs += find_dialogs()
        candidates = []
        for pid in quicklook_pids():
            candidates += [
                w for w in all_windows_of(pid)
                if w["visible"] and w["w"] >= 900 and w["h"] >= 500
                and (w["title"] or "").startswith("QuickLook")
            ]
        if candidates:
            candidates.sort(key=lambda w: w["w"] * w["h"], reverse=True)
            return candidates[0], dialogs
        time.sleep(0.5)
    return None, dialogs


# --------------------------------------------------------------------------- pixels


def count_near(rows, width, target, tol=12, y0=0, y1=None):
    """Approximate count of pixels close to `target`.

    Sampled on a 2x2 grid and scaled by 4: the thresholds used below are three orders of
    magnitude away from the sampled totals, so the approximation cannot change a verdict -
    but it does keep the scan fast on a 2x-DPI capture (2289x1591).
    """
    y1 = min(y1 if y1 is not None else len(rows), len(rows))
    tr, tg, tb = target
    hits = 0
    for y in range(y0, y1, 2):
        row = rows[y]
        for x in range(0, width, 2):
            i = x * 3
            if abs(row[i] - tr) <= tol and abs(row[i + 1] - tg) <= tol \
                    and abs(row[i + 2] - tb) <= tol:
                hits += 1
    return hits * 4


def strip_rows(rows, width, target, tol=12, min_frac=0.5):
    """Rows where `target` covers most of the width -> the notice strip's vertical extent."""
    rows_hit = []
    for y in range(min(TOP_SCAN, len(rows))):
        row = rows[y]
        hits = 0
        for x in range(0, width, 4):
            i = x * 3
            if abs(row[i] - target[0]) <= tol and abs(row[i + 1] - target[1]) <= tol \
                    and abs(row[i + 2] - target[2]) <= tol:
                hits += 1
        if hits >= (width // 4) * min_frac:
            rows_hit.append(y)
    return (min(rows_hit), max(rows_hit)) if rows_hit else None


def find_button(rows, width, y0, y1):
    """Bounding box of the near-white button face inside the notice strip."""
    xs, ys = [], []
    for y in range(y0, y1 + 1):
        row = rows[y]
        for x in range(width):
            i = x * 3
            if row[i] > 250 and row[i + 1] > 250 and row[i + 2] > 250:
                xs.append(x)
                ys.append(y)
    if not xs:
        return None
    return (min(xs), min(ys), max(xs), max(ys))


def click(x, y):
    USER32.SetCursorPos(int(x), int(y))
    time.sleep(0.15)
    USER32.mouse_event(0x0002, 0, 0, 0, None)  # LEFTDOWN
    time.sleep(0.06)
    USER32.mouse_event(0x0004, 0, 0, 0, None)  # LEFTUP
    time.sleep(0.4)


# --------------------------------------------------------------------------- scenarios


def cached_copies():
    if not os.path.isdir(CACHE_DIR):
        return []
    return [f for f in os.listdir(CACHE_DIR) if f.lower().endswith((".docx", ".doc"))]


def launch(path):
    kill_quicklook()  # also needed because SettingHelper caches the config per process
    subprocess.Popen(
        [EXE, path], cwd=INSTALL_DIR,
        creationflags=0x00000008 | 0x00000200,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )


def open_preview(path, out_png):
    launch(path)
    win, dialogs = wait_preview()
    if win is None:
        print(f"  !! no settled preview window (dialogs seen: {len(dialogs)})")
        return None, dialogs, None, None, None

    USER32.SetForegroundWindow(win["hwnd"])
    time.sleep(3.0)
    rows, width, height = capture(win["hwnd"], "printwindow")
    if rows:
        write_png(out_png, width, height, rows)
    print(f"  window {win['w']}x{win['h']} title={win['title']!r} "
          f"nonwhite={nonwhite_ratio(rows, width, height):.3f}")
    return win, dialogs, rows, width, height


def scenario_plain():
    print("\n=== T1  plain document (no zone mark) ===")
    if os.path.exists(PLAIN_FILE):
        os.remove(PLAIN_FILE)
    shutil.copy2(os.path.join(ROOT, "fixtures", "view-none.docx"), PLAIN_FILE)
    print(f"  zone marker: {'present' if os.path.exists(PLAIN_FILE + ':Zone.Identifier') else 'absent'}")

    before = set(cached_copies())
    win, dialogs, rows, width, height = open_preview(
        PLAIN_FILE, os.path.join(SHOTS, "copy-t1-plain.png"))
    if win is None:
        kill_quicklook()
        return {"ok": False, "why": "no window"}

    strip = strip_rows(rows, width, STRIP_AMBER) if rows else None
    amber = count_near(rows, width, STRIP_AMBER, y1=TOP_SCAN) if rows else 0
    during = set(cached_copies())
    print(f"  amber strip rows: {strip}   amber px in top 160: {amber}")
    print(f"  temp copies created: {sorted(during - before) or 'none'}")

    kill_quicklook()
    ok = (not dialogs and rows is not None
          and nonwhite_ratio(rows, width, height) > 0.02
          and strip is None and amber < 500 and not (during - before))
    return {"ok": ok}


def scenario_marked():
    print("\n=== T2  zone-marked document (Protected View) ===")
    if os.path.exists(ZONE_FILE):
        os.remove(ZONE_FILE)
    shutil.copy2(os.path.join(ROOT, "fixtures", "view-none.docx"), ZONE_FILE)
    if os.path.exists(ZONE_FILE + ":Zone.Identifier"):
        os.remove(ZONE_FILE + ":Zone.Identifier")
    with open(ZONE_FILE + ":Zone.Identifier", "w", encoding="utf-8") as f:
        f.write("[ZoneTransfer]\r\nZoneId=3\r\n")
    print(f"  zone marker before: {'present' if os.path.exists(ZONE_FILE + ':Zone.Identifier') else 'MISSING'}")

    before = set(cached_copies())
    win, dialogs, rows, width, height = open_preview(
        ZONE_FILE, os.path.join(SHOTS, "copy-t2-marked.png"))
    if win is None:
        kill_quicklook()
        return {"ok": False, "why": "no window", "win": None}

    strip = strip_rows(rows, width, STRIP_AMBER) if rows else None
    amber = count_near(rows, width, STRIP_AMBER, y1=TOP_SCAN) if rows else 0
    during = set(cached_copies())
    still_marked = os.path.exists(ZONE_FILE + ":Zone.Identifier")
    print(f"  amber strip rows: {strip}   amber px in top 160: {amber}")
    print(f"  temp copies during preview: {sorted(during) or 'none'}")
    print(f"  ORIGINAL still zone-marked after preview: {still_marked}")

    ok = (not dialogs and rows is not None
          and nonwhite_ratio(rows, width, height) > 0.02
          and strip is not None and amber > 3000
          and bool(during) and still_marked)
    return {"ok": ok, "win": win, "rows": rows, "width": width,
            "strip": strip, "still_marked": still_marked, "dialogs": dialogs}


def scenario_click(state):
    print("\n=== T3  click the button -> unblock the source file ===")
    if state.get("win") is None or state.get("rows") is None:
        print("  skipped: T2 produced no window")
        return {"ok": False}

    win, rows, width = state["win"], state["rows"], state["width"]
    y0, y1 = state["strip"]
    box = find_button(rows, width, y0, y1)
    if box is None:
        print("  !! could not locate the button face inside the strip")
        return {"ok": False}

    bx0, by0, bx1, by1 = box
    px, py = (bx0 + bx1) // 2, (by0 + by1) // 2
    sx, sy = win["left"] + px, win["top"] + py
    print(f"  button box in window: ({bx0},{by0})-({bx1},{by1}) -> click at "
          f"window({px},{py}) = screen({sx},{sy})")

    USER32.SetForegroundWindow(win["hwnd"])
    time.sleep(0.5)
    click(sx, sy)
    time.sleep(2.5)

    after_rows, w2, h2 = capture(win["hwnd"], "printwindow")
    if after_rows:
        write_png(os.path.join(SHOTS, "copy-t3-after-click.png"), w2, h2, after_rows)

    green = count_near(after_rows, w2, STRIP_GREEN, y1=TOP_SCAN) if after_rows else 0
    unblocked = not os.path.exists(ZONE_FILE + ":Zone.Identifier")
    print(f"  green strip px in top 160: {green}")
    print(f"  source file zone mark after click: "
          f"{'REMOVED (unblocked)' if unblocked else 'still present'}")

    kill_quicklook()
    ok = unblocked and green > 3000
    return {"ok": ok, "green": green, "unblocked": unblocked}


def main():
    os.makedirs(ZONE_DIR, exist_ok=True)
    os.makedirs(SHOTS, exist_ok=True)
    with open(CONFIG, encoding="utf-8") as f:
        print("config:", f.read())

    t1 = scenario_plain()
    t2 = scenario_marked()
    t3 = scenario_click(t2)

    print("\n" + "=" * 72)
    print(f"  T1 plain: no notice, no copy, renders       -> {'PASS' if t1['ok'] else 'FAIL'}")
    print(f"  T2 marked: renders + notice + original kept -> {'PASS' if t2['ok'] else 'FAIL'}")
    print(f"  T3 button: source unblocked + strip green   -> {'PASS' if t3['ok'] else 'FAIL'}")
    print("=" * 72)

    kill_quicklook()
    return 0 if (t1["ok"] and t2["ok"] and t3["ok"]) else 1


if __name__ == "__main__":
    sys.exit(main())
