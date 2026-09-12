"""Two measurements for the optimization trade-off table.

1) Warm-server memory cost: keep the handler alive (as the plugin does via the
   static HandlerFactories + LockServer(TRUE) that is never undone), and report
   the resident memory of the Office process that gets pinned. This is the RAM
   price of turning a ~850 ms cold preview into a ~2 ms warm one.

2) Registry lookup cost: time the exact lookups OfficeViewer.CanHandle performs
   per preview (ShellExRegister.GetPreviewHandlerGUID + CLSIDRegister.GetName),
   to show how small that term is relative to the COM activation term.
"""

import ctypes
import json
import sys
import time

sys.path.insert(0, __file__.rsplit("\\", 1)[0])
from bench_com_activation import (  # noqa: E402
    GUID,
    IID_IClassFactory,
    IID_IUnknown,
    CLSCTX_LOCAL_SERVER,
    com_release,
    _vcall,
    enumerate_procs,
    ole32,
)

import winreg  # noqa: E402

PREVIEW_KEY = r"shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}"

# --- 1) registry lookup cost ------------------------------------------------


def lookup_once(ext):
    """Mirror of ShellExRegister.GetPreviewHandlerGUID + the CLSID name probe."""
    guid = None
    try:
        with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, ext) as k:
            try:
                with winreg.OpenKey(k, PREVIEW_KEY) as sk:
                    guid = winreg.QueryValueEx(sk, "")[0]
            except OSError:
                progid = winreg.QueryValueEx(k, "")[0]
                with winreg.OpenKey(
                    winreg.HKEY_CLASSES_ROOT, progid + "\\" + PREVIEW_KEY
                ) as sk:
                    guid = winreg.QueryValueEx(sk, "")[0]
    except OSError:
        return None
    if not guid:
        return None
    # CLSIDRegister.GetName -> Registry.GetValue(HKCR\CLSID\{guid}, "", null)
    try:
        name = winreg.QueryValue(winreg.HKEY_CLASSES_ROOT, "CLSID\\" + guid.strip())
    except OSError:
        name = None
    return guid.strip(), name


def bench_registry(ext=".xlsx", n=200):
    lookup_once(ext)  # warm
    t0 = time.perf_counter()
    for _ in range(n):
        lookup_once(ext)
    dt = (time.perf_counter() - t0) / n
    return dt * 1000


# --- 2) warm-server memory cost -------------------------------------------


def process_memory(name):
    """Resident working set + private bytes for processes whose exe matches."""
    import ctypes.wintypes as wt

    k32 = ctypes.WinDLL("kernel32", use_last_error=True)
    psapi = ctypes.WinDLL("psapi", use_last_error=True)
    PROCESS_QUERY_INFORMATION = 0x0400
    PROCESS_VM_READ = 0x0010

    class PROCESSENTRY32(ctypes.Structure):
        _fields_ = [
            ("dwSize", wt.DWORD), ("cntUsage", wt.DWORD), ("th32ProcessID", wt.DWORD),
            ("th32DefaultHeapID", ctypes.POINTER(ctypes.c_ulong)), ("th32ModuleID", wt.DWORD),
            ("cntThreads", wt.DWORD), ("th32ParentProcessID", wt.DWORD),
            ("pcPriClassBase", ctypes.c_long), ("dwFlags", wt.DWORD),
            ("szExeFile", ctypes.c_char * 260),
        ]

    class PROCESS_MEMORY_COUNTERS(ctypes.Structure):
        _fields_ = [
            ("cb", wt.DWORD), ("PageFaultCount", wt.DWORD),
            ("PeakWorkingSetSize", ctypes.c_size_t), ("WorkingSetSize", ctypes.c_size_t),
            ("QuotaPeakPagedPoolUsage", ctypes.c_size_t), ("QuotaPagedPoolUsage", ctypes.c_size_t),
            ("QuotaPeakNonPagedPoolUsage", ctypes.c_size_t), ("QuotaNonPagedPoolUsage", ctypes.c_size_t),
            ("PagefileUsage", ctypes.c_size_t), ("PeakPagefileUsage", ctypes.c_size_t),
        ]

    snap = k32.CreateToolhelp32Snapshot(0x2, 0)
    entry = PROCESSENTRY32()
    entry.dwSize = ctypes.sizeof(PROCESSENTRY32)
    out = []
    if k32.Process32First(snap, ctypes.byref(entry)):
        while True:
            exe = entry.szExeFile.decode("mbcs", "ignore").lower()
            if exe == name:
                h = k32.OpenProcess(PROCESS_QUERY_INFORMATION | PROCESS_VM_READ, False, entry.th32ProcessID)
                if h:
                    pmc = PROCESS_MEMORY_COUNTERS()
                    pmc.cb = ctypes.sizeof(PROCESS_MEMORY_COUNTERS)
                    if psapi.GetProcessMemoryInfo(h, ctypes.byref(pmc), pmc.cb):
                        out.append(
                            {
                                "pid": entry.th32ProcessID,
                                "working_set_mb": round(pmc.WorkingSetSize / 1048576, 1),
                                "private_mb": round(pmc.PagefileUsage / 1048576, 1),
                            }
                        )
                    k32.CloseHandle(h)
            if not k32.Process32Next(snap, ctypes.byref(entry)):
                break
    k32.CloseHandle(snap)
    return out


def hold_and_measure(clsid_str, exe_name, label):
    clsid = GUID(clsid_str)
    iid = IID_IClassFactory
    p_factory = ctypes.c_void_p()
    hr = ole32.CoGetClassObject(
        ctypes.byref(clsid), CLSCTX_LOCAL_SERVER, None, ctypes.byref(iid), ctypes.byref(p_factory)
    )
    if hr < 0:
        return {"label": label, "error": hex(hr & 0xFFFFFFFF)}
    _vcall(p_factory, 4, ctypes.c_int, [ctypes.c_int], 1)  # LockServer(TRUE) - never undone, as in the plugin

    iid_unk = IID_IUnknown
    p_inst = ctypes.c_void_p()
    _vcall(
        p_factory, 3, ctypes.c_int,
        [ctypes.c_void_p, ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)],
        None, ctypes.byref(iid_unk), ctypes.byref(p_inst),
    )

    result = {"label": label}
    for wait in (3, 10):
        time.sleep(wait if wait == 3 else 7)
        mem = process_memory(exe_name)
        result[f"alive_after_{wait}s"] = mem

    # release the pin so the server can shut down
    if p_inst.value:
        com_release(p_inst)
    _vcall(p_factory, 4, ctypes.c_int, [ctypes.c_int], 0)
    com_release(p_factory.value)

    time.sleep(8)
    result["after_release_alive"] = process_memory(exe_name)
    return result


def main():
    part1 = {"registry_lookup_ms_per_call": round(bench_registry(".xlsx"), 4)}
    print("### registry lookup cost (per preview, .xlsx)")
    print(json.dumps(part1, indent=2))

    ole32.CoInitializeEx(None, 0x2)
    part2 = []
    try:
        part2.append(
            hold_and_measure("{00020827-0000-0000-C000-000000000046}", "excel.exe", "Excel previewer pinned")
        )
        part2.append(
            hold_and_measure("{84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}", "winword.exe", "Word previewer pinned")
        )
        part2.append(
            hold_and_measure("{65235197-874B-4A07-BDC5-E65EA825B718}", "powerpnt.exe", "PowerPoint previewer pinned")
        )
    finally:
        ole32.CoUninitialize()

    print("\n### resident cost of keeping the Office COM server pinned")
    print(json.dumps(part2, indent=2, ensure_ascii=False))

    with open("warm_server_cost.json", "w", encoding="utf-8") as fh:
        json.dump({"registry": part1, "warm_server": part2}, fh, indent=2, ensure_ascii=False)


if __name__ == "__main__":
    main()
