"""Probe how Office / shell preview handlers are registered on this machine.

Reads HKCR (merged HKLM+HKCU view) for the shell preview-handler subkey of a
given file extension and resolves the CLSID registration chain so we can tell
whether the handler is hosted in-process or out-of-process (prevhost.exe / a
dedicated surrogate EXE).
"""

import sys
import winreg

PREVIEW_HANDLER_KEY = "{8895b1c6-b41f-4c1c-a562-0d564250836f}"

NAMES = {
    "{84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}": "Microsoft Word Previewer",
    "{00020827-0000-0000-C000-000000000046}": "Microsoft Excel Previewer",
    "{65235197-874B-4A07-BDC5-E65EA825B718}": "Microsoft PowerPoint Previewer",
    "{21E17C2F-AD3A-4b89-841F-09CFE02D16B7}": "Microsoft Visio Previewer",
    "{6d2b5079-2f0b-48dd-ab7f-97cec514d30b}": "prevhost.exe surrogate AppID",
}


def reg_value(root, path, name=""):
    try:
        with winreg.OpenKey(root, path) as k:
            val, _ = winreg.QueryValueEx(k, name)
            return val
    except OSError:
        return None


def subkeys(root, path):
    out = []
    try:
        with winreg.OpenKey(root, path) as k:
            i = 0
            while True:
                try:
                    out.append(winreg.EnumKey(k, i))
                    i += 1
                except OSError:
                    break
    except OSError:
        pass
    return out


def find_preview_clsid(ext):
    """Mirror ShellExRegister.GetPreviewHandlerGUID()."""
    # 1) direct on the extension key
    v = reg_value(winreg.HKEY_CLASSES_ROOT, rf"{ext}\shellex\{PREVIEW_HANDLER_KEY}")
    if v:
        return v, f"HKCR\\{ext}\\shellex\\{PREVIEW_HANDLER_KEY}"
    # 2) on the ProgID referenced by the extension key
    progid = reg_value(winreg.HKEY_CLASSES_ROOT, ext)
    if progid:
        v = reg_value(winreg.HKEY_CLASSES_ROOT, rf"{progid}\shellex\{PREVIEW_HANDLER_KEY}")
        if v:
            return v, f"HKCR\\{progid}\\shellex\\{PREVIEW_HANDLER_KEY}"
    return None, None


def describe_clsid(clsid):
    base = f"CLSID\\{clsid}"
    info = {
        "name": reg_value(winreg.HKEY_CLASSES_ROOT, base),
        "appid": reg_value(winreg.HKEY_CLASSES_ROOT, base, "AppID"),
        "inproc": reg_value(winreg.HKEY_CLASSES_ROOT, base + r"\InprocServer32"),
        "inproc_apartment": reg_value(winreg.HKEY_CLASSES_ROOT, base + r"\InprocServer32", "ThreadingModel"),
        "local": reg_value(winreg.HKEY_CLASSES_ROOT, base + r"\LocalServer32"),
        "subkeys": subkeys(winreg.HKEY_CLASSES_ROOT, base),
        "progid": reg_value(winreg.HKEY_CLASSES_ROOT, base + r"\ProgID"),
    }
    # surrogate registration lives under the AppID
    if info["appid"]:
        info["appid_dllsurrogate"] = reg_value(winreg.HKEY_CLASSES_ROOT, f"AppID\\{info['appid']}", "DllSurrogate")
        info["appid_runas"] = reg_value(winreg.HKEY_CLASSES_ROOT, f"AppID\\{info['appid']}", "RunAs")
        info["appid_name"] = reg_value(winreg.HKEY_CLASSES_ROOT, f"AppID\\{info['appid']}")
    return info


def main():
    exts = sys.argv[1:] or [".xlsx", ".xls", ".xlsb", ".docx", ".pptx", ".pptm", ".vsdx"]
    for ext in exts:
        clsid, where = find_preview_clsid(ext)
        print("=" * 78)
        print(f"{ext}")
        if not clsid:
            print("  no preview handler registered  -> OfficeViewer plugin CANNOT handle it")
            continue
        clsid = clsid.strip()
        print(f"  handler CLSID : {clsid}   ({NAMES.get(clsid, '?')})")
        print(f"  declared at   : {where}")
        info = describe_clsid(clsid)
        print(f"  CLSID name    : {info['name']}")
        print(f"  InprocServer32: {info['inproc']}   (ThreadingModel={info['inproc_apartment']})")
        print(f"  LocalServer32 : {info['local']}")
        print(f"  AppID         : {info['appid']}   ({NAMES.get(str(info['appid']).strip(), '?')})")
        if info.get("appid_dllsurrogate") is not None:
            print(f"    -> AppID DllSurrogate = '{info['appid_dllsurrogate']}'  (empty => default surrogate prevhost.exe)")
        if info.get("appid_runas"):
            print(f"    -> AppID RunAs = {info['appid_runas']!r}   ('Interactive User' => cross-session/slow)")
        print(f"  CLSID subkeys : {info['subkeys']}")


if __name__ == "__main__":
    main()
