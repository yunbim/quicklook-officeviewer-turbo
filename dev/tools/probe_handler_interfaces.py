"""Which initialization interfaces do the Office preview handlers actually expose?

PreviewHandlerHost.Open() initializes in this priority order:
    IInitializeWithStream -> IInitializeWithItem -> IInitializeWithFile

For an out-of-process (LocalServer32) handler, IInitializeWithStream means the
host hands it a managed IStream that is COM-marshalled across the process
boundary: every Read/Seek becomes an IPC round trip. IInitializeWithItem /
IInitializeWithFile instead let the Office process open the file itself with its
own I/O stack. So the question "does the handler even implement stream
initialization?" decides whether the plugin's chosen priority is optimal or
counter-productive.

This probe only calls QueryInterface - it never calls Initialize/SetWindow/
DoPreview, so nothing is rendered and no document is touched.
"""

import ctypes
import json
import sys

sys.path.insert(0, __file__.rsplit("\\", 1)[0])
from bench_com_activation import (  # noqa: E402
    GUID,
    IID_IClassFactory,
    IID_IUnknown,
    CLSCTX_LOCAL_SERVER,
    CLSCTX_INPROC_SERVER,
    com_release,
    _vcall,
    ole32,
)

INTERFACES = {
    "IInitializeWithStream": "{b824b49d-22ac-4161-ac8a-9916e8fa3f7f}",
    "IInitializeWithItem": "{7f73be3f-fb79-493c-a6c7-7ee14e245841}",
    "IInitializeWithFile": "{b7d14566-0509-4cce-a71f-0a554233bd9b}",
    "IPreviewHandler": "{8895b1c6-b41f-4c1c-a562-0d564250836f}",
    "IObjectWithSite": "{fc4801a3-2ba9-11cf-a229-00aa003d7352}",
    "IPersistStream": "{00000109-0000-0000-C000-000000000046}",
}

HANDLERS = [
    ("Excel (.xlsx)", "{00020827-0000-0000-C000-000000000046}"),
    ("Word  (.docx)", "{84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}"),
    ("PPT   (.pptx)", "{65235197-874B-4A07-BDC5-E65EA825B718}"),
]


def create(clsid_str, ctx):
    clsid = GUID(clsid_str)
    iid = IID_IClassFactory
    p_factory = ctypes.c_void_p()
    hr = ole32.CoGetClassObject(
        ctypes.byref(clsid), ctx, None, ctypes.byref(iid), ctypes.byref(p_factory)
    )
    if hr < 0 or not p_factory.value:
        return None, None, hr
    iid_unk = IID_IUnknown
    p_inst = ctypes.c_void_p()
    hr2 = _vcall(
        p_factory,
        3,
        ctypes.c_int,
        [ctypes.c_void_p, ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)],
        None,
        ctypes.byref(iid_unk),
        ctypes.byref(p_inst),
    )
    if hr2 < 0 or not p_inst.value:
        com_release(p_factory.value)
        return None, None, hr2
    return p_factory.value, p_inst.value, 0


def query_interfaces(p_unk):
    out = {}
    for name, iid_str in INTERFACES.items():
        iid = GUID(iid_str)
        ppv = ctypes.c_void_p()
        hr = _vcall(
            p_unk,
            0,  # IUnknown::QueryInterface
            ctypes.c_int,
            [ctypes.POINTER(GUID), ctypes.POINTER(ctypes.c_void_p)],
            ctypes.byref(iid),
            ctypes.byref(ppv),
        )
        out[name] = hr >= 0
        if ppv.value:
            com_release(ppv.value)
    return out


def main():
    ole32.CoInitializeEx(None, 0x2)  # STA
    report = {}
    try:
        for label, clsid in HANDLERS:
            entry = {}

            # (a) what happens with plain CLSCTX_INPROC_SERVER?
            fac_ip, inst_ip, hr_ip = create(clsid, CLSCTX_INPROC_SERVER)
            entry["inproc_clsctx_ok"] = inst_ip is not None
            entry["inproc_hr"] = hex(hr_ip & 0xFFFFFFFF) if hr_ip < 0 else "S_OK"
            if inst_ip:
                com_release(inst_ip)
            if fac_ip:
                com_release(fac_ip)

            # (b) the path the plugin actually takes
            fac, inst, hr = create(clsid, CLSCTX_LOCAL_SERVER)
            if not inst:
                entry["local_server_error"] = hex(hr & 0xFFFFFFFF)
            else:
                entry["interfaces"] = query_interfaces(inst)
                _vcall(fac, 4, ctypes.c_int, [ctypes.c_int], 0)  # LockServer(FALSE)
                com_release(inst)
                com_release(fac)
            report[label] = entry
    finally:
        ole32.CoUninitialize()

    print(json.dumps(report, indent=2))
    print()
    for label, e in report.items():
        print(f"{label}")
        print(f"   CLSCTX_INPROC_SERVER usable : {e.get('inproc_clsctx_ok')} ({e.get('inproc_hr')})")
        ifs = e.get("interfaces")
        if ifs:
            for name, ok in ifs.items():
                print(f"   {'YES' if ok else ' no'}  {name}")
            supported = [n for n, ok in ifs.items() if ok]
            best = next(
                (n for n in ("IInitializeWithStream", "IInitializeWithItem", "IInitializeWithFile")
                 if n in supported),
                None,
            )
            print(f"   -> plugin would use: {best}")
        print()


if __name__ == "__main__":
    main()
