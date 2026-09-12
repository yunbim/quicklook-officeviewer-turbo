"""Drive the installed QuickLook through its real preview pipeline and screenshot the result.

This exists because "the Word preview looks like Web Layout" is a claim about *pixels*, and
the only honest way to settle it is to render a document and look. Reading the plugin source
cannot answer it: the OfficeViewer plugin forwards to the Shell preview handler
(WINWORD.EXE), so the layout is decided inside Office, not in our code.

Usage:
    python capture_preview.py <file> <out.png> [--wait 6] [--size 1200x800]

What it does:
    1. kills any running QuickLook, then launches `QuickLook.exe <file>` (the app accepts a
       path argument and toggles the preview window for it)
    2. polls for the preview window belonging to that process
    3. captures it with PrintWindow(PW_RENDERFULLCONTENT) and, as a cross-check, a plain
       screen BitBlt of the same rectangle
    4. writes both as 24-bit PNGs (stdlib only - no Pillow on this machine)
"""

import ctypes
import os
import struct
import subprocess
import sys
import time
import zlib
from ctypes import wintypes

user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)


def make_dpi_aware():
    """Tell Windows this process wants real pixels, not virtualised ones.

    Must run before any window or DC is created. Without it the process is DPI-unaware and
    Windows virtualises every coordinate: on a 200%-scaled display GetWindowRect reports
    *half* the real size and PrintWindow captures only the top-left quarter of the window.

    That is not a cosmetic detail - it silently turned a correctly laid out, fully visible
    notice strip into "the button is missing", and it makes any click computed from a
    capture land in the wrong place.
    """
    try:
        if user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):  # PER_MONITOR_AWARE_V2
            return "per-monitor-v2"
    except Exception:
        pass
    try:
        shcore = ctypes.WinDLL("shcore", use_last_error=True)
        if shcore.SetProcessDpiAwareness(2) == 0:  # PROCESS_PER_MONITOR_DPI_AWARE
            return "per-monitor"
    except Exception:
        pass
    try:
        if user32.SetProcessDPIAware():
            return "system"
    except Exception:
        pass
    return "none"


DPI_AWARENESS = make_dpi_aware()

INSTALL_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Programs", "QuickLook")
EXE = os.path.join(INSTALL_DIR, "QuickLook.exe")

PW_RENDERFULLCONTENT = 0x00000002
SRCCOPY = 0x00CC0020
DIB_RGB_COLORS = 0


# --------------------------------------------------------------------------- win32 helpers


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", ctypes.c_long),
        ("biHeight", ctypes.c_long),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", ctypes.c_long),
        ("biYPelsPerMeter", ctypes.c_long),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


def write_png(path, width, height, rgb_rows):
    """Minimal 24-bit PNG encoder. rgb_rows: list of bytes, 3 bytes per pixel, top-down."""

    def chunk(tag, data):
        body = tag + data
        return (
            struct.pack(">I", len(data))
            + body
            + struct.pack(">I", zlib.crc32(body) & 0xFFFFFFFF)
        )

    raw = b"".join(b"\x00" + row for row in rgb_rows)  # filter type 0 per scanline
    header = struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0)  # 8-bit truecolour
    png = (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", header)
        + chunk(b"IDAT", zlib.compress(raw, 6))
        + chunk(b"IEND", b"")
    )
    with open(path, "wb") as f:
        f.write(png)


def dib_to_rgb(mem_dc, bmp, width, height):
    """Read a 32-bit top-down DIB back as packed 24-bit RGB rows."""

    info = BITMAPINFO()
    info.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    info.bmiHeader.biWidth = width
    info.bmiHeader.biHeight = -height  # negative => top-down
    info.bmiHeader.biPlanes = 1
    info.bmiHeader.biBitCount = 32
    info.bmiHeader.biCompression = 0

    buf = ctypes.create_string_buffer(width * height * 4)
    got = gdi32.GetDIBits(mem_dc, bmp, 0, height, buf, ctypes.byref(info), DIB_RGB_COLORS)
    if got == 0:
        return None

    rows = []
    for y in range(height):
        bgra = bytearray(buf.raw[y * width * 4 : (y + 1) * width * 4])
        rgb = bytearray(width * 3)
        rgb[0::3] = bgra[2::4]  # R
        rgb[1::3] = bgra[1::4]  # G
        rgb[2::3] = bgra[0::4]  # B
        rows.append(bytes(rgb))
    return rows


def capture(hwnd, mode):
    rect = RECT()
    user32.GetWindowRect(hwnd, ctypes.byref(rect))
    width = rect.right - rect.left
    height = rect.bottom - rect.top
    if width <= 0 or height <= 0:
        return None, 0, 0

    screen_dc = user32.GetDC(0) if mode == "screen" else user32.GetWindowDC(hwnd)
    mem_dc = gdi32.CreateCompatibleDC(screen_dc)
    bmp = gdi32.CreateCompatibleBitmap(screen_dc, width, height)
    gdi32.SelectObject(mem_dc, bmp)

    if mode == "printwindow":
        user32.PrintWindow(hwnd, mem_dc, PW_RENDERFULLCONTENT)
    else:
        gdi32.BitBlt(mem_dc, 0, 0, width, height, screen_dc, 0, 0, SRCCOPY)

    rows = dib_to_rgb(mem_dc, bmp, width, height)

    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mem_dc)
    if mode == "screen":
        user32.ReleaseDC(0, screen_dc)
    else:
        user32.ReleaseDC(hwnd, screen_dc)

    return rows, width, height


def nonwhite_ratio(rows, width, height):
    if not rows:
        return 0.0
    step = 4  # sample every 4th pixel, plenty for a sanity check
    hits = total = 0
    for y in range(0, height, step):
        row = rows[y]
        for x in range(0, width, step):
            px = row[x * 3 : x * 3 + 3]
            total += 1
            if px[0] < 240 or px[1] < 240 or px[2] < 240:
                hits += 1
    return hits / max(total, 1)


# --------------------------------------------------------------------------- window plumbing

WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HWND, wintypes.LPARAM)


def windows_of(pid):
    found = []

    def cb(hwnd, _):
        wpid = wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(wpid))
        if wpid.value == pid and user32.IsWindowVisible(hwnd):
            length = user32.GetWindowTextLengthW(hwnd)
            title = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, title, length + 1)
            cls = ctypes.create_unicode_buffer(256)
            user32.GetClassNameW(hwnd, cls, 256)
            rect = RECT()
            user32.GetWindowRect(hwnd, ctypes.byref(rect))
            found.append(
                {
                    "hwnd": hwnd,
                    "title": title.value,
                    "class": cls.value,
                    "w": rect.right - rect.left,
                    "h": rect.bottom - rect.top,
                }
            )
        return True

    user32.EnumWindows(WNDENUMPROC(cb), 0)
    return found


def quicklook_pids():
    out = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq QuickLook.exe", "/FO", "CSV", "/NH"],
        capture_output=True,
    ).stdout.decode("utf-8", errors="ignore")
    return [int(line.split('","')[1]) for line in out.splitlines() if line.startswith('"')]


def kill_quicklook(timeout=15.0):
    """Kill QuickLook and wait for the single-instance mutex to actually clear.

    Without this wait the freshly launched process finds the *dying* instance still holding
    the mutex, forwards the path to it over the named pipe and exits - so the pid we spawned
    never owns a window and the capture silently finds nothing.
    """
    subprocess.run(["taskkill", "/IM", "QuickLook.exe", "/F"], capture_output=True)
    deadline = time.time() + timeout
    while time.time() < deadline:
        if not quicklook_pids():
            return True
        time.sleep(0.5)
    return False


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 1

    target = os.path.abspath(sys.argv[1])
    out = os.path.abspath(sys.argv[2])
    wait = 6.0
    if "--wait" in sys.argv:
        wait = float(sys.argv[sys.argv.index("--wait") + 1])

    if not os.path.exists(target):
        print("no such file:", target)
        return 1

    kill_quicklook()
    proc = subprocess.Popen(
        [EXE, target],
        cwd=INSTALL_DIR,
        creationflags=0x00000008 | 0x00000200,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    print(f"launched QuickLook pid={proc.pid} for {os.path.basename(target)}")

    # The preview window belongs to whichever process ended up owning the single instance,
    # so poll every QuickLook process rather than just the pid we spawned. QuickLook first
    # shows a small placeholder and only later resizes to the plugin's preferred size, so
    # wait for a large window instead of accepting the first one that appears.
    deadline = time.time() + wait
    win = None
    while time.time() < deadline:
        time.sleep(0.75)
        candidates = []
        for pid in quicklook_pids():
            candidates += [
                w for w in windows_of(pid)
                if w["w"] >= 900 and w["h"] >= 500 and w["title"].startswith("QuickLook")
            ]
        if candidates:
            candidates.sort(key=lambda w: w["w"] * w["h"], reverse=True)
            win = candidates[0]
            break

    if win is None:
        print("no settled preview window appeared within", wait, "s")
        proc.terminate()
        return 2

    # Let the handler paint its first page before grabbing pixels.
    time.sleep(3.0)
    print(f"window: hwnd={win['hwnd']} class={win['class']!r} title={win['title']!r} "
          f"{win['w']}x{win['h']}")

    user32.SetForegroundWindow(win["hwnd"])
    time.sleep(0.6)

    for mode, suffix in (("printwindow", ""), ("screen", ".screen")):
        rows, width, height = capture(win["hwnd"], mode)
        if rows is None:
            print(f"{mode}: capture failed")
            continue
        path = out if not suffix else out.replace(".png", suffix + ".png")
        write_png(path, width, height, rows)
        print(f"{mode}: {width}x{height} nonwhite={nonwhite_ratio(rows, width, height):.3f} -> {path}")

    proc.terminate()
    return 0


if __name__ == "__main__":
    sys.exit(main())
