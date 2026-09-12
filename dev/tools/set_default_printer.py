"""Read / change the Windows default printer, and enumerate installed printers.

Why this matters here: Word's page layout (and therefore both Word's paged export and the
Shell Word preview handler's rendering) depends on a usable default printer device context.
This machine's default printer sits on an unreachable network port, so the default device is
a prime suspect for "no page layout anywhere". This tool makes that testable and reversible.

Uses the documented Win32 APIs (GetDefaultPrinterW / SetDefaultPrinterW / EnumPrintersW) via
ctypes rather than rundll32 or WMI, which are blocked in this environment.

Usage:
    python set_default_printer.py                      # list printers + default
    python set_default_printer.py --default NAME       # switch default (reversible)
    python set_default_printer.py --save FILE          # remember the current default
    python set_default_printer.py --restore FILE       # put it back
"""

import ctypes
import ctypes.wintypes as wt
import json
import sys

winspool = ctypes.WinDLL("winspool.drv", use_last_error=True)
winspool.GetDefaultPrinterW.argtypes = [wt.LPWSTR, ctypes.POINTER(wt.DWORD)]
winspool.GetDefaultPrinterW.restype = wt.BOOL
winspool.SetDefaultPrinterW.argtypes = [wt.LPWSTR]
winspool.SetDefaultPrinterW.restype = wt.BOOL


def get_default():
    size = wt.DWORD(0)
    winspool.GetDefaultPrinterW(None, ctypes.byref(size))
    if size.value == 0:
        return None
    buf = ctypes.create_unicode_buffer(size.value)
    if not winspool.GetDefaultPrinterW(buf, ctypes.byref(size)):
        return None
    return buf.value


def set_default(name):
    ok = winspool.SetDefaultPrinterW(name)
    if not ok:
        raise OSError(f"SetDefaultPrinterW failed: {ctypes.get_last_error()}")
    return get_default()


class PRINTER_INFO_4(ctypes.Structure):
    _fields_ = [
        ("pPrinterName", wt.LPWSTR),
        ("pServerName", wt.LPWSTR),
        ("Attributes", wt.DWORD),
    ]


def list_printers():
    needed = wt.DWORD(0)
    returned = wt.DWORD(0)
    winspool.EnumPrintersW(
        2, None, 4, None, 0, ctypes.byref(needed), ctypes.byref(returned)
    )
    if needed.value == 0:
        return []
    buf = ctypes.create_string_buffer(needed.value)
    if not winspool.EnumPrintersW(
        2, None, 4, buf, needed.value, ctypes.byref(needed), ctypes.byref(returned)
    ):
        return []
    arr = ctypes.cast(buf, ctypes.POINTER(PRINTER_INFO_4))
    return [arr[i].pPrinterName for i in range(returned.value)]


def main():
    argv = sys.argv[1:]

    if "--save" in argv:
        path = argv[argv.index("--save") + 1]
        with open(path, "w", encoding="utf-8") as f:
            json.dump({"default": get_default()}, f)
        print(f"saved default {get_default()!r} -> {path}")
        return 0

    if "--restore" in argv:
        path = argv[argv.index("--restore") + 1]
        with open(path, encoding="utf-8") as f:
            name = json.load(f)["default"]
        after = set_default(name)
        print(f"restored default -> {after!r}")
        return 0

    if "--default" in argv:
        name = argv[argv.index("--default") + 1]
        before = get_default()
        after = set_default(name)
        print(f"default printer: {before!r} -> {after!r}")
        return 0

    print("installed printers:")
    for name in list_printers():
        print("   -", name)
    print(f"default: {get_default()!r}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
