"""Time opening a device context for each installed printer.

This is the check that decides the Word preview layout question. Word builds page layout from
the default printer's driver + DEVMODE; if that device cannot be created, Word falls back to a
continuous, non-paginated rendering - which is exactly what the Office preview handler shows.
Reproducing the failure at the Win32 level (CreateDCW) puts it beyond doubt and identifies
*which* printer is at fault.

Usage: python probe_printer_dc.py
"""

import ctypes
import ctypes.wintypes as wt
import time

gdi32 = ctypes.WinDLL("gdi32", use_last_error=True)
winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)

gdi32.CreateDCW.argtypes = [wt.LPCWSTR, wt.LPCWSTR, wt.LPCWSTR, ctypes.c_void_p]
gdi32.CreateDCW.restype = wt.HDC
gdi32.DeleteDC.argtypes = [wt.HDC]

winspool.OpenPrinterW.argtypes = [wt.LPCWSTR, ctypes.POINTER(wt.HANDLE), ctypes.c_void_p]
winspool.OpenPrinterW.restype = wt.BOOL
winspool.GetDefaultPrinterW.argtypes = [wt.LPWSTR, ctypes.POINTER(wt.DWORD)]
winspool.GetDefaultPrinterW.restype = wt.BOOL


def default_printer():
    size = wt.DWORD(0)
    winspool.GetDefaultPrinterW(None, ctypes.byref(size))
    if size.value == 0:
        return None
    buf = ctypes.create_unicode_buffer(size.value)
    winspool.GetDefaultPrinterW(buf, ctypes.byref(size))
    return buf.value


def printers():
    needed = wt.DWORD(0)
    returned = wt.DWORD(0)
    winspool.EnumPrintersW(2, None, 4, None, 0, ctypes.byref(needed), ctypes.byref(returned))
    if needed.value == 0:
        return []
    buf = ctypes.create_string_buffer(needed.value)
    winspool.EnumPrintersW(2, None, 4, buf, needed.value, ctypes.byref(needed),
                           ctypes.byref(returned))

    class PRINTER_INFO_4(ctypes.Structure):
        _fields_ = [("pPrinterName", wt.LPWSTR), ("pServerName", wt.LPWSTR),
                    ("Attributes", wt.DWORD)]

    arr = ctypes.cast(buf, ctypes.POINTER(PRINTER_INFO_4))
    return [arr[i].pPrinterName for i in range(returned.value)]


HORZSIZE, VERTSIZE = 4, 6
HORZRES, VERTRES = 8, 10
LOGPIXELSX, LOGPIXELSY = 88, 90
PHYSICALWIDTH, PHYSICALHEIGHT = 110, 111
PHYSICALOFFSETX, PHYSICALOFFSETY = 112, 113

gdi32.GetDeviceCaps.argtypes = [wt.HDC, ctypes.c_int]
gdi32.GetDeviceCaps.restype = ctypes.c_int

CAPS = [
    ("paper (0.1mm)  ", HORZSIZE, VERTSIZE),
    ("resolvable px  ", HORZRES, VERTRES),
    ("dpi            ", LOGPIXELSX, LOGPIXELSY),
    ("physical px    ", PHYSICALWIDTH, PHYSICALHEIGHT),
    ("physical offset", PHYSICALOFFSETX, PHYSICALOFFSETY),
]


def describe_dc(dc):
    for label, cap_x, cap_y in CAPS:
        print(f"        {label} x={gdi32.GetDeviceCaps(dc, cap_x):<7d} "
              f"y={gdi32.GetDeviceCaps(dc, cap_y)}")


def main():
    current = default_printer()
    print(f"default printer: {current!r}")
    print()

    for name in printers():
        marker = "  <- DEFAULT" if name == current else ""

        handle = wt.HANDLE()
        t0 = time.perf_counter()
        opened = bool(winspool.OpenPrinterW(name, ctypes.byref(handle), None))
        t_open = (time.perf_counter() - t0) * 1000

        t0 = time.perf_counter()
        dc = gdi32.CreateDCW("WINSPOOL", name, None, None)
        t_dc = (time.perf_counter() - t0) * 1000
        err = ctypes.get_last_error()

        print(f"{name}{marker}")
        print(f"    OpenPrinter   {'ok' if opened else 'FAILED'}  {t_open:8.1f} ms")
        print(f"    CreateDC      {'ok' if dc else f'FAILED ({err})'}  {t_dc:8.1f} ms")
        if dc:
            describe_dc(dc)
            gdi32.DeleteDC(dc)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
