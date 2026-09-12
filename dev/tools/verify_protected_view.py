"""Verify the OfficeViewer Protected View behaviour end to end, on the real install.

The requirement: a document downloaded from the Internet (i.e. carrying a
Zone.Identifier alternate data stream) must preview immediately instead of raising the
"PROTECTED VIEW / Be careful - files from the Internet can contain viruses" dialog that
asks the user to confirm unblocking.

Testing this needs care, because "no dialog appeared" is only meaningful if the harness can
actually see dialogs. So this runs two scenarios:

  A) AlwaysUnblockProtectedView = True   -> EXPECT no dialog, document renders, ADS stripped
  B) AlwaysUnblockProtectedView = False  -> EXPECT the dialog to be detected (control)

Scenario B is the control: if it does not find the dialog, the detector is broken and
scenario A proves nothing. The config is restored to True (the requested setting) on exit.

Usage:
    python verify_protected_view.py [source.docx]
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
CONFIG = os.path.join(INSTALL_DIR, "UserData", "QuickLook.Plugin.OfficeViewer.config")
ZONE_DIR = os.path.join(ROOT, "zone-test")
ZONE_FILE = os.path.join(ZONE_DIR, "downloaded-from-internet.docx")

DIALOG_TITLE = "PROTECTED VIEW"
ZONE_STREAM = ZONE_FILE + ":Zone.Identifier"


# --------------------------------------------------------------------------- helpers


def write_config(always_unblock, warmup=True):
    body = (
        '<?xml version="1.0" encoding="utf-8"?><Settings>'
        f"<WarmUpAfterFirstPreview>{warmup}</WarmUpAfterFirstPreview>"
        "<WarmUpAtStartup>False</WarmUpAtStartup>"
        "<HandlerIdleTimeoutSeconds>1800</HandlerIdleTimeoutSeconds>"
        f"<AlwaysUnblockProtectedView>{always_unblock}</AlwaysUnblockProtectedView>"
        "</Settings>"
    )
    with open(CONFIG, "w", encoding="utf-8") as f:
        f.write(body)


def mark_zone_blocked(path):
    """Give the file the marker Explorer/Office use for 'downloaded from the Internet'."""
    with open(path + ":Zone.Identifier", "w", encoding="utf-8") as f:
        f.write("[ZoneTransfer]\r\nZoneId=3\r\n")
    return zone_identifiers(path)


def zone_identifiers(path):
    """List the zone streams currently attached to a file."""
    return [s for s in ("Zone.Identifier",) if os.path.exists(path + ":" + s)]


WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


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
                "hwnd": hwnd,
                "title": title.value,
                "class": cls.value,
                "visible": bool(USER32.IsWindowVisible(hwnd)),
                "w": r.right - r.left,
                "h": r.bottom - r.top,
            })
        return True

    USER32.EnumWindows(WNDENUMPROC(cb), 0)
    return found


def detect_dialog(timeout):
    """Poll for a modal PROTECTED VIEW dialog (class #32770) on any QuickLook process."""
    deadline = time.time() + timeout
    seen = []
    while time.time() < deadline:
        for pid in quicklook_pids():
            for w in all_windows_of(pid):
                if w["visible"] and w["class"] == "#32770":
                    seen.append(w)
                elif w["visible"] and DIALOG_TITLE in (w["title"] or "").upper():
                    seen.append(w)
        if seen:
            return seen
        time.sleep(0.4)
    return []


def preview_window(timeout):
    deadline = time.time() + timeout
    while time.time() < deadline:
        for pid in quicklook_pids():
            for w in all_windows_of(pid):
                if w["visible"] and w["w"] >= 900 and w["h"] >= 500 \
                        and (w["title"] or "").startswith("QuickLook"):
                    return w
        time.sleep(0.5)
    return None


# --------------------------------------------------------------------------- scenarios


def run_scenario(label, always_unblock, out_png):
    print(f"\n=== {label} (AlwaysUnblockProtectedView={always_unblock}) ===")
    write_config(always_unblock)

    # Re-arm the marker: scenario A strips it, so B must start from a blocked file again.
    if os.path.exists(ZONE_FILE):
        os.remove(ZONE_FILE)
    shutil.copy2(os.path.join(ROOT, "fixtures", "view-none.docx"), ZONE_FILE)
    os.remove(ZONE_FILE + ":Zone.Identifier") if os.path.exists(ZONE_FILE + ":Zone.Identifier") else None
    zones = mark_zone_blocked(ZONE_FILE)
    print(f"  zone marker before preview: {zones or 'NONE'}")

    kill_quicklook()  # needed anyway: SettingHelper caches the XML per process
    subprocess.Popen(
        [EXE, ZONE_FILE], cwd=INSTALL_DIR,
        creationflags=0x00000008 | 0x00000200,
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )

    dialogs = detect_dialog(14.0)
    win = preview_window(16.0)

    result = {"label": label, "dialogs": dialogs, "win": win}
    if dialogs:
        d = dialogs[0]
        print(f"  DIALOG DETECTED: class={d['class']!r} title={d['title']!r} "
              f"{d['w']}x{d['h']}")
    else:
        print("  no dialog detected")

    if win:
        print(f"  preview window: {win['w']}x{win['h']} title={win['title']!r}")
        rect = RECT()
        USER32.GetWindowRect(win["hwnd"], ctypes.byref(rect))
        USER32.SetForegroundWindow(win["hwnd"])
        time.sleep(0.7)
        rows, w, h = capture(win["hwnd"], "printwindow")
        if rows:
            write_png(out_png, w, h, rows)
            result["nonwhite"] = nonwhite_ratio(rows, w, h)
            print(f"  captured {w}x{h} nonwhite={result['nonwhite']:.3f} -> {out_png}")
    else:
        print("  no settled preview window appeared (expected if a modal dialog is up)")

    time.sleep(1.0)
    left = zone_identifiers(ZONE_FILE)
    result["zone_after"] = left
    print(f"  zone marker after preview: {left or 'NONE (unblocked)'}")

    kill_quicklook()
    return result


def main():
    original_config = None
    if os.path.exists(CONFIG):
        with open(CONFIG, encoding="utf-8") as f:
            original_config = f.read()
    os.makedirs(ZONE_DIR, exist_ok=True)

    print(f"config file: {CONFIG}")
    print(f"test file:   {ZONE_FILE}")

    try:
        a = run_scenario("SCENARIO A  (wanted behaviour)", True,
                         os.path.join(SHOTS := os.path.join(ROOT, "shots"),
                                      "protected-a-unblocked.png"))
        b = run_scenario("SCENARIO B  (control)", False,
                         os.path.join(SHOTS, "protected-b-dialog.png"))
    finally:
        write_config(True)  # leave the requested setting in place
        kill_quicklook()
        print("\nconfig restored to AlwaysUnblockProtectedView=True")

    print("\n" + "=" * 70)
    ok_a = not a["dialogs"] and a["win"] is not None and not a["zone_after"]
    ok_b = bool(b["dialogs"])
    print(f"  A: no popup + rendered + unblocked  -> {'PASS' if ok_a else 'FAIL'}")
    print(f"  B: control popup detected           -> {'PASS' if ok_b else 'FAIL'}")
    if not ok_b:
        print("  !! control failed - scenario A proves nothing, the detector needs fixing")
    return 0 if (ok_a and ok_b) else 1


if __name__ == "__main__":
    sys.exit(main())
