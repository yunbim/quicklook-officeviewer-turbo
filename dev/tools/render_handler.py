"""Render an Office file through its Shell preview handler into a PNG, with full control.

Why this exists: the QuickLook plugin forwards Office rendering to the Shell preview handler
(WINWORD.EXE / EXCEL.EXE / POWERPNT.EXE). Anything we might "fix" about the view lives inside
that handler, so it has to be studied in isolation - with a window size we choose, on a file
we control - instead of inferred from source.

The host window here is a plain Win32 window (not WPF), so PrintWindow captures the
cross-process child window the handler creates without the layered-window airspace problem
that makes screen grabs useless inside QuickLook.

Usage:
    python render_handler.py <file> <out.png> [--size 1200x800] [--wait 6]
                             [--clsid {84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}] [--keep]

Defaults to the Microsoft Word previewer. The file is compiled against the repo's own
IPreviewHandler / IInitializeWithFile vtable layouts, so the call order matches the plugin.
"""

import ctypes
import ctypes.wintypes as wt
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from bench_com_activation import (  # noqa: E402
    GUID,
    IID_IClassFactory,
    IID_IUnknown,
    _vcall,
    com_release,
)
from capture_preview import write_png  # noqa: E402

ole32 = ctypes.WinDLL("ole32", use_last_error=True)
user32 = ctypes.WinDLL("user32", use_last_error=True)
gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)

CLSCTX_LOCAL_SERVER = 0x4
COINIT_APARTMENTTHREADED = 0x2
S_FALSE = 1

IID_IINITIALIZEWITHFILE = "{b7d14566-0509-4cce-a71f-0a554233bd9b}"
IID_IPREVIEWHANDLER = "{8895b1c6-b41f-4c1c-a562-0d564250836f}"

WORD_PREVIEWER = "{84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}"

# vtable slots, counting from the three IUnknown entries
SLOT_CREATE_INSTANCE = 3
SLOT_LOCK_SERVER = 4
SLOT_QI, SLOT_ADDREF, SLOT_RELEASE = 0, 1, 2
SLOT_INITIALIZE_WITH_FILE = 3          # IInitializeWithFile
SLOT_SET_WINDOW = 3                    # IPreviewHandler
SLOT_SET_RECT = 4
SLOT_DO_PREVIEW = 5
SLOT_UNLOAD = 6

PW_RENDERFULLCONTENT = 0x00000002

WNDPROC = ctypes.WINFUNCTYPE(
    ctypes.c_long, wt.HWND, ctypes.c_uint, wt.WPARAM, wt.LPARAM
)


class WNDCLASSEX(ctypes.Structure):
    _fields_ = [
        ("cbSize", ctypes.c_uint),
        ("style", ctypes.c_uint),
        ("lpfnWndProc", WNDPROC),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wt.HINSTANCE),
        ("hIcon", wt.HANDLE),
        ("hCursor", wt.HANDLE),
        ("hbrBackground", wt.HANDLE),
        ("lpszMenuName", wt.LPCWSTR),
        ("lpszClassName", wt.LPCWSTR),
        ("hIconSm", wt.HANDLE),
    ]


class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]


class BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wt.DWORD),
        ("biWidth", ctypes.c_long),
        ("biHeight", ctypes.c_long),
        ("biPlanes", wt.WORD),
        ("biBitCount", wt.WORD),
        ("biCompression", wt.DWORD),
        ("biSizeImage", wt.DWORD),
        ("biXPelsPerMeter", ctypes.c_long),
        ("biYPelsPerMeter", ctypes.c_long),
        ("biClrUsed", wt.DWORD),
        ("biClrImportant", wt.DWORD),
    ]


class BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", BITMAPINFOHEADER), ("bmiColors", wt.DWORD * 3)]


# Handles are pointer-sized; without explicit prototypes ctypes assumes c_int and truncates
# them on 64-bit, which shows up as "CreateWindowEx failed: 0".
kernel32.GetModuleHandleW.argtypes = [wt.LPCWSTR]
kernel32.GetModuleHandleW.restype = wt.HINSTANCE
user32.RegisterClassExW.argtypes = [ctypes.POINTER(WNDCLASSEX)]
user32.RegisterClassExW.restype = wt.ATOM
user32.CreateWindowExW.restype = wt.HWND
user32.CreateWindowExW.argtypes = [
    wt.DWORD, wt.LPCWSTR, wt.LPCWSTR, wt.DWORD,
    ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
    wt.HWND, wt.HANDLE, wt.HINSTANCE, ctypes.c_void_p,
]
user32.UpdateWindow.argtypes = [wt.HWND]
user32.DestroyWindow.argtypes = [wt.HWND]
user32.DefWindowProcW.argtypes = [wt.HWND, ctypes.c_uint, wt.WPARAM, wt.LPARAM]
user32.DefWindowProcW.restype = ctypes.c_long
user32.GetClientRect.argtypes = [wt.HWND, ctypes.POINTER(RECT)]
user32.GetDC.argtypes = [wt.HWND]
user32.GetDC.restype = wt.HDC
user32.ReleaseDC.argtypes = [wt.HWND, wt.HDC]
user32.PrintWindow.argtypes = [wt.HWND, wt.HDC, ctypes.c_uint]
gdi32.CreateCompatibleDC.argtypes = [wt.HDC]
gdi32.CreateCompatibleDC.restype = wt.HDC
gdi32.CreateCompatibleBitmap.argtypes = [wt.HDC, ctypes.c_int, ctypes.c_int]
gdi32.CreateCompatibleBitmap.restype = wt.HBITMAP
gdi32.SelectObject.argtypes = [wt.HDC, wt.HGDIOBJ]
gdi32.SelectObject.restype = wt.HGDIOBJ
gdi32.DeleteObject.argtypes = [wt.HGDIOBJ]
gdi32.DeleteDC.argtypes = [wt.HDC]
gdi32.GetDIBits.argtypes = [
    wt.HDC, wt.HBITMAP, ctypes.c_uint, ctypes.c_uint,
    ctypes.c_void_p, ctypes.POINTER(BITMAPINFO), ctypes.c_uint,
]
gdi32.GetDIBits.restype = ctypes.c_int


def dump_child_windows(root):
    """Print every descendant window and its rect.

    This is the decisive check for "the handler re-laid out but only painted a corner":
    if the child window really is small, the handler shrank its own surface; if it is
    full-size, we are only looking at an unfinished paint.
    """
    user32.GetClassNameW.argtypes = [wt.HWND, ctypes.c_wchar_p, ctypes.c_int]
    user32.GetWindowRect.argtypes = [wt.HWND, ctypes.POINTER(RECT)]
    user32.IsWindowVisible.argtypes = [wt.HWND]

    found = []

    @ctypes.WINFUNCTYPE(ctypes.c_int, wt.HWND, wt.LPARAM)
    def cb(child, _):
        buf = ctypes.create_unicode_buffer(256)
        user32.GetClassNameW(child, buf, 256)
        wr = RECT()
        user32.GetWindowRect(child, ctypes.byref(wr))
        found.append((child, buf.value, wr.left, wr.top, wr.right - wr.left,
                      wr.bottom - wr.top, bool(user32.IsWindowVisible(child))))
        return 1

    user32.EnumChildWindows.argtypes = [wt.HWND, cb.__class__, wt.LPARAM]
    user32.EnumChildWindows(root, cb, 0)
    print(f"child windows of host hwnd={root}: {len(found)}")
    for h, cls, l, t, w, hh, vis in found:
        print(f"    hwnd={h:<9} {cls:28} at ({l},{t}) size {w}x{hh} visible={vis}")


def pump_messages(seconds):
    """Pump the thread's message queue. The handler's child window needs this to be created
    and painted; without it DoPreview appears to do nothing."""
    deadline = time.time() + seconds
    msg = wt.MSG()
    while time.time() < deadline:
        while user32.PeekMessageW(ctypes.byref(msg), None, 0, 0, 1):
            user32.TranslateMessage(ctypes.byref(msg))
            user32.DispatchMessageW(ctypes.byref(msg))
        time.sleep(0.02)


def screenshot(hwnd, out_path):
    rect = RECT()
    user32.GetClientRect(hwnd, ctypes.byref(rect))
    width, height = rect.right, rect.bottom

    # Ask the window to render itself, children included. PrintWindow travels cross-process,
    # which is what makes it work for the handler's out-of-process child window.
    wnd_dc = user32.GetDC(hwnd)
    mem_dc = gdi32.CreateCompatibleDC(wnd_dc)
    bmp = gdi32.CreateCompatibleBitmap(wnd_dc, width, height)
    gdi32.SelectObject(mem_dc, bmp)
    user32.PrintWindow(hwnd, mem_dc, PW_RENDERFULLCONTENT)

    info = BITMAPINFO()
    info.bmiHeader.biSize = ctypes.sizeof(BITMAPINFOHEADER)
    info.bmiHeader.biWidth = width
    info.bmiHeader.biHeight = -height
    info.bmiHeader.biPlanes = 1
    info.bmiHeader.biBitCount = 32
    buf = ctypes.create_string_buffer(width * height * 4)
    gdi32.GetDIBits(mem_dc, bmp, 0, height, buf, ctypes.byref(info), 0)

    rows = []
    for y in range(height):
        bgra = bytearray(buf.raw[y * width * 4 : (y + 1) * width * 4])
        rgb = bytearray(width * 3)
        rgb[0::3] = bgra[2::4]
        rgb[1::3] = bgra[1::4]
        rgb[2::3] = bgra[0::4]
        rows.append(bytes(rgb))

    gdi32.DeleteObject(bmp)
    gdi32.DeleteDC(mem_dc)
    user32.ReleaseDC(hwnd, wnd_dc)
    write_png(out_path, width, height, rows)

    nonwhite = sum(
        1
        for y in range(0, height, 4)
        for x in range(0, width, 4)
        if rows[y][x * 3] < 240 or rows[y][x * 3 + 1] < 240 or rows[y][x * 3 + 2] < 240
    )
    return width, height, nonwhite / max((height // 4 + 1) * (width // 4 + 1), 1)


def create_host_window(width, height):
    inst = kernel32.GetModuleHandleW(None)
    cls_name = "QLPreviewProbeHost"

    def wndproc(hwnd, msg, wparam, lparam):
        if msg == 0x0002:  # WM_DESTROY
            user32.PostQuitMessage(0)
            return 0
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    proc = WNDPROC(wndproc)
    wc = WNDCLASSEX()
    wc.cbSize = ctypes.sizeof(WNDCLASSEX)
    wc.style = 0
    wc.lpfnWndProc = proc
    wc.hInstance = inst
    wc.hbrBackground = 5 + 1  # COLOR_WINDOW + 1
    wc.lpszClassName = cls_name
    if not user32.RegisterClassExW(ctypes.byref(wc)):
        err = ctypes.get_last_error()
        if err != 1410:  # ERROR_CLASS_ALREADY_EXISTS
            raise OSError(f"RegisterClassEx failed: {err}")

    # Borderless popup on purpose: no caption or frame means the client area is the whole
    # bitmap (so pixel analysis sees only document pixels), and window size is not clamped
    # to the desktop's maximum tracking size - which matters when the page needs to be shown
    # wider than the monitor.
    WS_POPUP = 0x80000000
    WS_VISIBLE = 0x10000000
    hwnd = user32.CreateWindowExW(
        0x00000008,  # WS_EX_TOPMOST - keep it unobstructed for the capture
        cls_name, "Office preview handler probe",
        WS_POPUP | WS_VISIBLE,
        40, 40, width, height, None, None, inst, None,
    )
    if not hwnd:
        raise OSError(f"CreateWindowEx failed: {ctypes.get_last_error()}")

    # Swap in a plain background so a blank render is obvious rather than ambiguous.
    user32.UpdateWindow(hwnd)
    return hwnd, proc


def main():
    if len(sys.argv) < 3:
        print(__doc__)
        return 1

    path = os.path.abspath(sys.argv[1])
    out = os.path.abspath(sys.argv[2])
    width, height = 1200, 800
    wait = 6.0
    clsid = WORD_PREVIEWER
    if "--size" in sys.argv:
        width, height = (int(v) for v in sys.argv[sys.argv.index("--size") + 1].split("x"))
    if "--wait" in sys.argv:
        wait = float(sys.argv[sys.argv.index("--wait") + 1])
    if "--clsid" in sys.argv:
        clsid = sys.argv[sys.argv.index("--clsid") + 1]

    if not os.path.exists(path):
        print("no such file:", path)
        return 1

    hr = ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)
    if hr < 0 and hr != S_FALSE:
        print(f"CoInitializeEx failed: 0x{hr & 0xFFFFFFFF:08x}")
        return 1

    hwnd, wndproc = create_host_window(width, height)  # noqa: F841 - keep proc alive
    print(f"host window {width}x{height} hwnd={hwnd}")

    clsid_guid = GUID(clsid)
    p_factory = ctypes.c_void_p()
    hr = ole32.CoGetClassObject(
        ctypes.byref(clsid_guid), CLSCTX_LOCAL_SERVER, None,
        ctypes.byref(IID_IClassFactory), ctypes.byref(p_factory),
    )
    if hr < 0:
        print(f"CoGetClassObject failed: 0x{hr & 0xFFFFFFFF:08x}")
        return 1

    # Keep the server alive for the duration, and - importantly - release it again.
    _vcall(p_factory, SLOT_LOCK_SERVER, ctypes.c_int, [ctypes.c_int], 1)

    p_unk = ctypes.c_void_p()
    hr = _vcall(
        p_factory, SLOT_CREATE_INSTANCE, ctypes.c_int,
        [ctypes.c_void_p, ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)],
        None, ctypes.byref(IID_IUnknown), ctypes.byref(p_unk),
    )
    if hr < 0:
        print(f"CreateInstance failed: 0x{hr & 0xFFFFFFFF:08x}")
        return 1

    # QI on IUnknown (slot 0) for the two interfaces we need.
    def query(ptr, iid_str):
        out_ptr = ctypes.c_void_p()
        g = GUID(iid_str)
        rc = _vcall(
            ptr, SLOT_QI, ctypes.c_int,
            [ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)],
            ctypes.byref(g), ctypes.byref(out_ptr),
        )
        return (out_ptr if rc >= 0 else None), rc

    p_file_init, rc = query(p_unk, IID_IINITIALIZEWITHFILE)
    print(f"QI IInitializeWithFile -> 0x{rc & 0xFFFFFFFF:08x}")
    p_handler, rc = query(p_unk, IID_IPREVIEWHANDLER)
    print(f"QI IPreviewHandler     -> 0x{rc & 0xFFFFFFFF:08x}")
    if not p_file_init or not p_handler:
        return 1

    hr = _vcall(
        p_file_init, SLOT_INITIALIZE_WITH_FILE, ctypes.c_int,
        [wt.LPCWSTR, ctypes.c_uint], path, 0,
    )
    print(f"Initialize('{os.path.basename(path)}') -> 0x{hr & 0xFFFFFFFF:08x}")
    if hr < 0:
        print("  (handler refused the file - this is what Protected View looks like)")
        return 3

    rect = RECT(0, 0, width, height)

    # --grow-then-shrink models what QuickLook actually does: the WinForms host is created and
    # activated while the WPF layout is still settling, so the handler may be shown a rect that
    # is *not* the final one. If the handler bakes its zoom into the first rect, every later
    # resize keeps the wrong zoom - which is exactly the "content is too big and clipped"
    # symptom. Start it on a roomy rect, then shrink to the real one and see whether it re-fits.
    start_rect = rect
    if "--grow-then-shrink" in sys.argv:
        start_rect = RECT(0, 0, max(width, 1400), max(height, 2000))

    hr = _vcall(
        p_handler, SLOT_SET_WINDOW, ctypes.c_int,
        [wt.HWND, ctypes.POINTER(RECT)], hwnd, ctypes.byref(start_rect),
    )
    print(f"SetWindow({start_rect.right}x{start_rect.bottom}) -> 0x{hr & 0xFFFFFFFF:08x}")

    hr = _vcall(p_handler, SLOT_DO_PREVIEW, ctypes.c_int, [])
    print(f"DoPreview -> 0x{hr & 0xFFFFFFFF:08x}")

    pump_messages(wait)

    if "--grow-then-shrink" in sys.argv:
        print("grow-then-shrink: shrinking window to the real size and re-fitting")
        user32.SetWindowPos(hwnd, None, 40, 40, width, height, 0x0004 | 0x0010)
        _vcall(p_handler, SLOT_SET_WINDOW, ctypes.c_int,
               [wt.HWND, ctypes.POINTER(RECT)], hwnd, ctypes.byref(rect))
        _vcall(p_handler, SLOT_SET_RECT, ctypes.c_int, [ctypes.POINTER(RECT)],
               ctypes.byref(rect))
        user32.InvalidateRect(hwnd, None, True)
        pump_messages(3.0)

    if "--refit" in sys.argv:
        # Re-apply the rect after the preview has started, in case the handler only lays out
        # once (on DoPreview) and ignores later SetRect calls.
        print("refit: re-applying SetWindow/SetRect")
        _vcall(
            p_handler, SLOT_SET_WINDOW, ctypes.c_int,
            [wt.HWND, ctypes.POINTER(RECT)], hwnd, ctypes.byref(rect),
        )
        user32.InvalidateRect(hwnd, None, True)
        pump_messages(3.0)
        _vcall(p_handler, SLOT_SET_RECT, ctypes.c_int, [ctypes.POINTER(RECT)],
               ctypes.byref(rect))
        pump_messages(4.0)
        # Force the whole tree to repaint; a partially painted re-layout is otherwise
        # indistinguishable from a genuinely small render.
        user32.RedrawWindow(hwnd, None, None,
                            0x0001 | 0x0100 | 0x0080 | 0x0400)  # INVALIDATE|UPDATENOW|ALLCHILDREN|ERASE
        pump_messages(3.0)

    if "--children" in sys.argv:
        dump_child_windows(hwnd)

    w, h, ratio = screenshot(hwnd, out)
    print(f"captured {w}x{h} nonwhite={ratio:.3f} -> {out}")

    # Tear down in the reverse order, and always pair LockServer(TRUE) with FALSE so we do
    # not leave an Office process pinned in the background.
    _vcall(p_handler, SLOT_UNLOAD, ctypes.c_int, [])
    com_release(p_handler)
    com_release(p_file_init)
    com_release(p_unk)
    _vcall(p_factory, SLOT_LOCK_SERVER, ctypes.c_int, [ctypes.c_int], 0)
    com_release(p_factory)
    user32.DestroyWindow(hwnd)
    ole32.CoUninitialize()
    return 0


if __name__ == "__main__":
    sys.exit(main())
