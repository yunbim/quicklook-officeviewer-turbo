"""Dump the Word per-user settings that could drive the preview handler's view mode.

Word's preview handler is WINWORD.EXE itself, so its rendering choices are likely governed by
the same per-user state as Word proper - not by anything the QuickLook plugin can pass in.
The interesting candidates are:

  HKCU\\Software\\Microsoft\\Office\\<ver>\\Word\\Data\\PreferredView
      The value behind the well-known "Word always starts in Web Layout" problem.
  HKCU\\...\\Word\\Options - any view / zoom / preview / protected-view related value.

Usage: python probe_word_view.py
"""

import winreg

VERSIONS = ["16.0", "15.0", "14.0", "12.0"]

INTERESTING = ("view", "zoom", "preview", "pane", "protect", "reading", "layout")


def dump(hive, path):
    try:
        key = winreg.OpenKey(hive, path)
    except OSError:
        return False

    print(f"=== {path} ===")
    count = winreg.QueryInfoKey(key)[1]
    if count == 0:
        print("   (no values)")

    for i in range(count):
        name, value, kind = winreg.EnumValue(key, i)
        if isinstance(value, bytes):
            shown = value.hex()
        else:
            shown = repr(value)
        print(f"   {name:32s} type={kind:<3d} {shown}")

    winreg.CloseKey(key)
    return True


def main():
    for ver in VERSIONS:
        base = rf"Software\Microsoft\Office\{ver}\Word"
        for sub in ("Data", "Options"):
            path = base + "\\" + sub
            if not dump(winreg.HKEY_CURRENT_USER, path):
                continue

    # The preview handlers' own registration, for cross-reference with the CLSID forensics.
    print()
    for ext in (".docx", ".doc", ".xlsx", ".pptx"):
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CLASSES_ROOT,
                ext + r"\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}",
            )
            guid = winreg.QueryValueEx(key, "")[0]
            winreg.CloseKey(key)
            print(f"{ext:6s} preview handler {guid}")
        except OSError as e:
            print(f"{ext:6s} -> {e}")


if __name__ == "__main__":
    main()
