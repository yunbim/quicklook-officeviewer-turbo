"""Read or write HKCU\\Software\\Microsoft\\Office\\16.0\\Word\\Options\\PreferredView.

Word's per-user "preferred view" is the documented cause of Word opening (and, if the preview
handler shares the setting, previewing) documents in Web Layout instead of Print Layout.
That makes it the cheapest possible lever for "make Word previews look like paper", so it is
worth testing directly instead of reasoned about.

WdViewType reference: 1=Normal 2=Outline 3=Print 4=PrintPreview 5=Master 6=Web 7=Reading

Usage:
    python set_word_preferred_view.py            # print current value
    python set_word_preferred_view.py 3          # set to Print Layout
"""

import sys
import winreg

KEY = r"Software\Microsoft\Office\16.0\Word\Options"
NAME = "PreferredView"


def read_value():
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, KEY) as key:
            return winreg.QueryValueEx(key, NAME)[0]
    except OSError:
        return None


def write_value(value):
    with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, KEY, 0, winreg.KEY_SET_VALUE) as key:
        winreg.SetValueEx(key, NAME, 0, winreg.REG_DWORD, value)


def main():
    if len(sys.argv) == 1:
        current = read_value()
        print(f"PreferredView = {current!r}" if current is not None else "PreferredView = <absent>")
        return 0

    value = int(sys.argv[1])
    previous = read_value()
    write_value(value)
    print(f"PreferredView {previous!r} -> {value}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
