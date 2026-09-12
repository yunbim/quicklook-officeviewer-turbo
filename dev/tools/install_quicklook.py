"""Install the QuickLook 4.5.0 release and swap in the optimized OfficeViewer plugin.

Usage:
    python install_quicklook.py extract      # unzip the release into the install dir
    python install_quicklook.py patch        # back up + replace the OfficeViewer plugin
    python install_quicklook.py extra        # add plugins that only exist after 4.5.0
    python install_quicklook.py verify       # report install state
    python install_quicklook.py revert       # restore the stock OfficeViewer plugin
"""

import hashlib
import os
import shutil
import sys
import zipfile

PROJECT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ZIP = os.path.join(PROJECT, "dist", "QuickLook-4.5.0.zip")

INSTALL_DIR = os.path.join(
    os.environ["LOCALAPPDATA"], "Programs", "QuickLook"
)

BUILD_PLUGIN_ROOT = os.path.join(PROJECT, "repo", "Build", "Release", "QuickLook.Plugin")

BUILT_PLUGIN = os.path.join(
    BUILD_PLUGIN_ROOT, "QuickLook.Plugin.OfficeViewer", "QuickLook.Plugin.OfficeViewer.dll",
)

PLUGIN_REL = os.path.join(
    "QuickLook.Plugin", "QuickLook.Plugin.OfficeViewer",
    "QuickLook.Plugin.OfficeViewer.dll",
)

# Plugins added to the repository after the 4.5.0 release. They are built from source and
# dropped in alongside the shipped ones; QuickLook's plugin manager isolates a plugin that
# fails to load, so a bad one cannot take the host down with it.
EXTRA_PLUGINS = [
    "QuickLook.Plugin.BinaryViewer",
    "QuickLook.Plugin.ChmViewer",
    "QuickLook.Plugin.DbViewer",
    "QuickLook.Plugin.DumpViewer",
    "QuickLook.Plugin.PrefetchViewer",
]


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def do_extract():
    if not os.path.isfile(ZIP):
        sys.exit(f"release archive not found: {ZIP}")
    os.makedirs(INSTALL_DIR, exist_ok=True)
    with zipfile.ZipFile(ZIP) as z:
        z.extractall(INSTALL_DIR)
    print(f"extracted {ZIP}")
    print(f"     -> {INSTALL_DIR}")


def do_patch():
    target = os.path.join(INSTALL_DIR, PLUGIN_REL)
    if not os.path.isfile(target):
        sys.exit(f"installed plugin not found: {target}")
    if not os.path.isfile(BUILT_PLUGIN):
        sys.exit(f"built plugin not found: {BUILT_PLUGIN}")

    backup = target + ".stock-4.5.0"
    if not os.path.exists(backup):
        shutil.copy2(target, backup)
        print(f"backed up stock plugin -> {os.path.basename(backup)}")
    else:
        print("stock backup already present, keeping it")

    before = sha256(target)
    shutil.copy2(BUILT_PLUGIN, target)
    after = sha256(target)

    print(f"backup (stock 4.5.0) sha256 : {sha256(backup)}")
    print(f"previous installed  sha256 : {before}")
    print(f"now installed       sha256 : {after}")
    print("PATCHED" if before != after else "WARNING: identical, patch did not apply")
    print(f"source of the new dll      : {BUILT_PLUGIN}")


def do_revert():
    target = os.path.join(INSTALL_DIR, PLUGIN_REL)
    backup = target + ".stock-4.5.0"
    if not os.path.isfile(backup):
        sys.exit(f"no stock backup to restore at {backup}")
    shutil.copy2(backup, target)
    print(f"restored stock plugin, sha256 {sha256(target)}")


def do_extra():
    dest_root = os.path.join(INSTALL_DIR, "QuickLook.Plugin")
    if not os.path.isdir(dest_root):
        sys.exit(f"install dir not populated: {dest_root}")

    for name in EXTRA_PLUGINS:
        src = os.path.join(BUILD_PLUGIN_ROOT, name)
        if not os.path.isdir(src):
            print(f"SKIP {name}: not built")
            continue

        dst = os.path.join(dest_root, name)
        os.makedirs(dst, exist_ok=True)
        shutil.copytree(src, dst, dirs_exist_ok=True)
        count = sum(len(files) for _, _, files in os.walk(dst))
        print(f"installed {name} ({count} files)")


def do_verify():
    exe = os.path.join(INSTALL_DIR, "QuickLook.exe")
    print(f"install dir : {INSTALL_DIR}")
    print(f"  exists    : {os.path.isdir(INSTALL_DIR)}")
    print(f"  QuickLook.exe : {os.path.isfile(exe)}")
    print(f"  portable.lock : {os.path.isfile(os.path.join(INSTALL_DIR, 'portable.lock'))}")

    native = ["QuickLook.Native32.dll", "QuickLook.Native64.dll", "QuickLook.Common.dll"]
    for n in native:
        print(f"  {n}: {os.path.isfile(os.path.join(INSTALL_DIR, n))}")

    plugin_root = os.path.join(INSTALL_DIR, "QuickLook.Plugin")
    plugins = sorted(
        d for d in os.listdir(plugin_root)
        if os.path.isdir(os.path.join(plugin_root, d))
    )
    print(f"  plugins   : {len(plugins)}")
    for p in plugins:
        print(f"      {p}")

    target = os.path.join(INSTALL_DIR, PLUGIN_REL)
    backup = target + ".stock-4.5.0"
    print()
    print(f"OfficeViewer.dll          : {os.path.isfile(target)}")
    if os.path.isfile(target):
        print(f"  installed sha256        : {sha256(target)}")
    if os.path.isfile(BUILT_PLUGIN):
        print(f"  our build sha256        : {sha256(BUILT_PLUGIN)}")
    if os.path.isfile(backup):
        print(f"  stock backup sha256     : {sha256(backup)}")
        if os.path.isfile(target):
            patched = sha256(target) != sha256(backup)
            print(f"  optimization applied    : {patched}")


if __name__ == "__main__":
    action = sys.argv[1] if len(sys.argv) > 1 else "verify"
    {
        "extract": do_extract,
        "patch": do_patch,
        "extra": do_extra,
        "verify": do_verify,
        "revert": do_revert,
    }.get(action, do_verify)()
