"""Measure the real cold/warm cost of the COM activation that QuickLook's
OfficeViewer plugin performs for every Office preview.

PreviewHandlerHost.Open() runs on the WPF UI thread and does:

    CoGetClassObject(clsid, CLSCTX_LOCAL_SERVER, ...)   <-- IClassFactory (cached in
                                                            static HandlerFactories,
                                                            LockServer(TRUE) => never
                                                            released)
    factory.CreateInstance(null, IID_IUnknown, ...)     <-- new handler instance
    handler.Initialize via IInitializeWithStream/Item/File
    handler.SetWindow(hwnd, rect) / DoPreview()         <-- blocking, UI thread

On this machine the Office preview-handler CLSIDs are registered as
LocalServer32 pointing at the full Office executables (EXCEL.EXE / WINWORD.EXE /
POWERPNT.EXE) with **no AppID / no DllSurrogate**. So CLSCTX_LOCAL_SERVER launches
the whole Office application - not a lightweight prevhost.exe surrogate, contrary
to the comment in the plugin source.

This probe times that activation with raw COM (ctypes). It does NOT call
SetWindow()/DoPreview(), so no preview UI is shown and no document is parsed: we
isolate the activation term, i.e. the cost the plugin pays *before* Office even
starts reading the file.

Cleanup: LockServer(FALSE) is always issued before release so no Office process is
left pinned by this probe.
"""

import ctypes
import json
import sys
import time

ole32 = ctypes.WinDLL("ole32", use_last_error=True)

CLSCTX_INPROC_SERVER = 0x1
CLSCTX_LOCAL_SERVER = 0x4
COINIT_APARTMENTTHREADED = 0x2


class GUID(ctypes.Structure):
    _fields_ = [
        ("Data1", ctypes.c_ulong),
        ("Data2", ctypes.c_ushort),
        ("Data3", ctypes.c_ushort),
        ("Data4", ctypes.c_ubyte * 8),
    ]

    def __init__(self, s):
        super().__init__()
        if ole32.CLSIDFromString(ctypes.c_wchar_p(s), ctypes.byref(self)) < 0:
            raise ValueError(f"bad GUID {s!r}")

    def __str__(self):
        d4 = "".join(f"{b:02X}" for b in self.Data4)
        return f"{{{self.Data1:08X}-{self.Data2:04X}-{self.Data3:04X}-{d4[:4]}-{d4[4:]}}}"


IID_IClassFactory = GUID("{00000001-0000-0000-C000-000000000046}")
IID_IUnknown = GUID("{00000000-0000-0000-C000-000000000046}")

HANDLERS = [
    ("Excel (.xlsx)", "{00020827-0000-0000-C000-000000000046}"),
    ("Word  (.docx)", "{84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}"),
    ("PPT   (.pptx)", "{65235197-874B-4A07-BDC5-E65EA825B718}"),
]

OFFICE_EXES = {"excel.exe", "winword.exe", "powerpnt.exe", "vpreview.exe", "prevhost.exe"}


def _vcall(ptr, index, restype, argtypes, *args):
    vtable = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
    proto = ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)
    return proto(vtable[index])(ptr, *args)


def com_release(ptr):
    if ptr:
        _vcall(ptr, 2, ctypes.c_ulong, [])


def enumerate_procs():
    import ctypes.wintypes as wt

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)

    class PROCESSENTRY32(ctypes.Structure):
        _fields_ = [
            ("dwSize", wt.DWORD),
            ("cntUsage", wt.DWORD),
            ("th32ProcessID", wt.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)),
            ("th32ModuleID", wt.DWORD),
            ("cntThreads", wt.DWORD),
            ("th32ParentProcessID", wt.DWORD),
            ("pcPriClassBase", ctypes.c_long),
            ("dwFlags", wt.DWORD),
            ("szExeFile", ctypes.c_char * 260),
        ]

    snap = k32.CreateToolhelp32Snapshot(0x2, 0)
    entry = PROCESSENTRY32()
    entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
    names = set()
    if k32.Process32First(snap, ctypes.byref(entry)):
        while True:
            names.add(entry.szExeFile.decode("mbcs", "ignore").lower())
            if not k32.Process32Next(snap, ctypes.byref(entry)):
                break
    k32.CloseHandle(snap)
    return names


def activate(clsid_str):
    """CoGetClassObject + CreateInstance, keeping the factory alive."""
    clsid = GUID(clsid_str)
    iid = IID_IClassFactory
    p_factory = ctypes.c_void_p()

    before = enumerate_procs() & OFFICE_EXES

    t0 = time.perf_counter()
    hr = ole32.CoGetClassObject(
        ctypes.byref(clsid), CLSCTX_LOCAL_SERVER, None, ctypes.byref(iid), ctypes.byref(p_factory)
    )
    t1 = time.perf_counter()
    if hr < 0 or not p_factory.value:
        return None, {
            "ok": False,
            "hr": hex(hr & 0xFFFFFFFF),
            "total_ms": round((t1 - t0) * 1000, 1),
        }

    _vcall(p_factory, 4, ctypes.c_int, [ctypes.c_int], 1)  # LockServer(TRUE)

    p_inst = ctypes.c_void_p()
    iid_unk = IID_IUnknown
    t2 = time.perf_counter()
    try:
        hr2 = _vcall(
            p_factory,
            3,  # IClassFactory::CreateInstance
            ctypes.c_int,
            [ctypes.c_void_p, ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)],
            None,
            ctypes.byref(iid_unk),
            ctypes.byref(p_inst),
        )
    except OSError as exc:
        print(f"    CreateInstance raised: {exc}", file=sys.stderr)
        hr2 = -1
    t3 = time.perf_counter()

    time.sleep(0.5)
    after = enumerate_procs() & OFFICE_EXES

    if p_inst.value:
        com_release(p_inst)

    return (p_factory.value, clsid), {
        "ok": hr2 >= 0 and bool(p_inst.value),
        "getclassobject_ms": round((t1 - t0) * 1000, 1),
        "createinstance_ms": round((t3 - t2) * 1000, 1),
        "total_ms": round((t3 - t0) * 1000, 1),
        "started": sorted(after - before),
    }


def main():
    ole32.CoInitializeEx(None, COINIT_APARTMENTTHREADED)
    live = []
    results = []
    try:
        print("### PHASE 1 - COLD (no Office process running yet)\n")
        for label, clsid in HANDLERS:
            handle, r = activate(clsid)
            if handle:
                live.append(handle)
            r["target"], r["phase"] = label, "cold"
            results.append(r)
            print(f"  {label:16s} total={r.get('total_ms'):>8} ms  "
                  f"getClassObject={r.get('getclassobject_ms')} createInstance={r.get('createinstance_ms')}  "
                  f"started={r.get('started')}")
            time.sleep(1.5)

        print("\n### PHASE 2 - WARM (server process already alive)\n")
        for label, clsid in HANDLERS:
            _, r = activate(clsid)
            r["target"], r["phase"] = label, "warm"
            results.append(r)
            print(f"  {label:16s} total={r.get('total_ms'):>8} ms  "
                  f"getClassObject={r.get('getclassobject_ms')} createInstance={r.get('createinstance_ms')}  "
                  f"started={r.get('started')}")
            time.sleep(1.0)
    finally:
        # Release the LockServer pins so the servers can shut down again.
        for ptr, _clsid in live:
            if ptr:
                _vcall(ptr, 4, ctypes.c_int, [ctypes.c_int], 0)  # LockServer(FALSE)
                com_release(ptr)
        ole32.CoUninitialize()

    with open("com_activation_results.json", "w", encoding="utf-8") as fh:
        json.dump(results, fh, indent=2, ensure_ascii=False)
    print("\n(wrote com_activation_results.json)")


if __name__ == "__main__":
    main()
