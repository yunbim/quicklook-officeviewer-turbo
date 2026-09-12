"""End-to-end check of the Office warm-up / idle-reclaim optimisation.

It writes the plugin settings, restarts QuickLook, then samples the process table to show
that QuickLook pre-activates the Office preview hosts in the background and later releases
them once they have been idle.

    python verify_office_warmup.py [--warmup on|off] [--idle SECONDS] [--watch SECONDS]

What to look for:
    * with warm-up on  -> EXCEL.EXE / WINWORD.EXE / POWERPNT.EXE appear a few seconds after
      QuickLook starts, WITHOUT the user previewing anything;
    * after the idle timeout elapses -> the same processes disappear again, which is the
      half of the lifecycle the stock plugin was missing.
"""

import argparse
import os
import subprocess
import sys
import time
import xml.etree.ElementTree as ET

INSTALL_DIR = os.path.join(os.environ["LOCALAPPDATA"], "Programs", "QuickLook")
EXE = os.path.join(INSTALL_DIR, "QuickLook.exe")
CONFIG = os.path.join(INSTALL_DIR, "UserData", "QuickLook.Plugin.OfficeViewer.config")

OFFICE_EXES = {"EXCEL.EXE", "WINWORD.EXE", "POWERPNT.EXE", "OUTLOOK.EXE", "VPREVIEW.EXE"}


def _decode(raw):
    """tasklist emits the console codepage (GBK on zh-CN Windows), not UTF-8."""
    for enc in ("utf-8", "gbk", "cp936", "cp1252"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="ignore")


def office_procs():
    """Return {name: (pid, rss_kb)} for running Office processes."""
    try:
        res = subprocess.run(
            ["tasklist", "/FO", "CSV", "/NH"],
            capture_output=True, timeout=30,
        )
    except Exception:
        return {}

    found = {}
    for line in _decode(res.stdout).splitlines():
        parts = [p.strip('"') for p in line.split('","')]
        if len(parts) < 5:
            continue
        name = parts[0].strip('"').upper()
        if name.upper() in OFFICE_EXES:
            digits = "".join(ch for ch in parts[4] if ch.isdigit())
            found[name.upper()] = (parts[1].strip('"'), int(digits or 0))
    return found


def write_config(warmup, idle_seconds):
    os.makedirs(os.path.dirname(CONFIG), exist_ok=True)
    settings = {
        "WarmUpAtStartup": "True" if warmup else "False",
        "WarmUpAfterFirstPreview": "True",
        "HandlerIdleTimeoutSeconds": str(idle_seconds),
    }

    root = ET.Element("Settings")
    for k, v in settings.items():
        ET.SubElement(root, k).text = v

    with open(CONFIG, "w", encoding="utf-8") as f:
        f.write('<?xml version="1.0" encoding="utf-8"?>')
        f.write(ET.tostring(root, encoding="unicode"))

    print(f"settings -> {CONFIG}")
    for k, v in settings.items():
        print(f"    {k} = {v}")


def restart_quicklook():
    subprocess.run(["taskkill", "/IM", "QuickLook.exe", "/F"], capture_output=True)
    time.sleep(2)
    subprocess.Popen([EXE], cwd=INSTALL_DIR,
                     creationflags=0x00000008 | 0x00000200,
                     stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    print("restarted QuickLook")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--warmup", choices=["on", "off"], default="on")
    ap.add_argument("--idle", type=int, default=60,
                    help="HandlerIdleTimeoutSeconds to write (default 60)")
    ap.add_argument("--watch", type=int, default=180,
                    help="how long to keep sampling, in seconds")
    args = ap.parse_args()

    print("=== baseline (before restart) ===")
    for name, (pid, rss) in sorted(office_procs().items()):
        print(f"    {name} pid={pid} rss={rss / 1024:.0f} MB")
    if not office_procs():
        print("    (no Office processes)")

    write_config(args.warmup == "on", args.idle)
    restart_quicklook()

    print(f"\n=== sampling for {args.watch}s ===")
    t0 = time.time()
    previous = set()
    events = []

    while time.time() - t0 < args.watch:
        now = set(office_procs())
        for name in sorted(now - previous):
            events.append((time.time() - t0, "STARTED", name))
        for name in sorted(previous - now):
            events.append((time.time() - t0, "EXITED ", name))
        previous = now
        time.sleep(2)

    if not events:
        print("    no transitions observed")
    for at, what, name in events:
        print(f"    t+{at:6.1f}s  {what}  {name}")

    print("\n=== final state ===")
    for name, (pid, rss) in sorted(office_procs().items()):
        print(f"    {name} pid={pid} rss={rss / 1024:.0f} MB")
    if not office_procs():
        print("    (no Office processes)")


if __name__ == "__main__":
    sys.exit(main())
