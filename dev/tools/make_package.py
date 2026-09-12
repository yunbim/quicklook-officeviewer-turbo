#!/usr/bin/env python3
"""Assemble the portable, redistributable package for QuickLook Office 极速预览版 (Turbo).

    python tools/make_package.py

Produces:
    dist/QuickLook-OfficeViewer-Turbo/                    full package (folder)
    dist/QuickLook-OfficeViewer-Turbo-<ver>.zip            full package (zip)
    dist/QuickLook-OfficeViewer-Turbo-plugin-only.zip      tiny, plugin + install.bat only
    dist/quicklook-officeviewer-turbo.bundle              git bundle of the patched branch

Layout of the full package (root is deliberately "obvious at a glance"):

    README.md  INSTALL.txt  CHANGELOG.md  LICENSE-GPL.txt  MANIFEST.txt
    install.bat  rollback.bat
    plugin/          <- the ONE file that gets deployed
    plugin-stock/    <- official 4.5.0 DLL, for exact rollback
    config/          <- default settings template
    dev/             <- src / patch / docs / tools (developers only)
"""

import hashlib
import os
import shutil
import subprocess
import sys
import zipfile

VERSION = "4.5.0-turbo.3"
PKG_NAME = "QuickLook-OfficeViewer-Turbo"
BRANCH = "office-preview-patched"
# The patch is generated for the whole branch, not just its tip commit, so adding a second
# commit on top of the branch must not silently shrink the patch to that one commit.
BASE_REF = "master"

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
REPO = os.path.join(ROOT, "repo")
DIST = os.path.join(ROOT, "dist")
PKG = os.path.join(DIST, PKG_NAME)
PLUGIN_ZIP_NAME = PKG_NAME + "-plugin-only"

PLUGIN_REL = "QuickLook.Plugin/QuickLook.Plugin.OfficeViewer"
DLL_NAME = "QuickLook.Plugin.OfficeViewer.dll"

SRC_FILES = [
    "QuickLook.Common/ExtensionMethods/WindowInteropHelperExtension.cs",
    f"{PLUGIN_REL}/OfficePreviewHostPool.cs",
    f"{PLUGIN_REL}/ProtectedViewPreview.cs",
    f"{PLUGIN_REL}/ProtectedViewBanner.cs",
    f"{PLUGIN_REL}/Plugin.cs",
    f"{PLUGIN_REL}/PreviewHandlerHost.cs",
    f"{PLUGIN_REL}/PreviewPanel.cs",
    f"{PLUGIN_REL}/UiText.cs",
]

DOCS = [
    "QuickLook-Office预览性能深度分析.md",
    "QuickLook-Office优化-实现与部署.md",
    "REPORT-protected-view-copy-preview.md",
]

TOOL_FILES = [
    "analyze_render.py",
    "bench_com_activation.py",
    "capture_preview.py",
    "exp_printer_layout.py",
    "exp_warm_cold.py",
    "install_quicklook.py",
    "make_view_fixture.py",
    "probe_handler_interfaces.py",
    "probe_office_preview.py",
    "probe_printer_dc.py",
    "probe_warm_server_cost.py",
    "probe_word_view.py",
    "render_handler.py",
    "set_default_printer.py",
    "set_word_preferred_view.py",
    "sweep_sizes.py",
    "verify_office_warmup.py",
    "verify_protected_copy.py",
    "verify_protected_view.py",
    "publish_github.sh",
    "make_package.py",
]


def log(msg):
    print(msg, flush=True)


def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def is_junk(rel):
    """Paths that must never end up in the package (editor/bytecode/VCS noise)."""
    parts = rel.replace("\\", "/").split("/")
    return (
        "__pycache__" in parts
        or ".git" in parts
        or rel.endswith(".pyc")
        or rel.endswith(".pyo")
        or ".DS_Store" in parts
        # lives in packaging/ so the publishing text stays in one place, but it belongs on
        # the GitHub Release page, not inside the package the release ships
        or rel.replace("\\", "/") == "RELEASE-NOTES.md"
    )


def strip_junk(root):
    """Delete junk that a previous verification run may have left inside the package."""
    removed = []
    for dirpath, dirnames, filenames in os.walk(root, topdown=False):
        for fn in filenames:
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, root)
            if is_junk(rel):
                os.remove(full)
                removed.append(rel.replace("\\", "/"))
        for dn in list(dirnames):
            full = os.path.join(dirpath, dn)
            if os.path.isdir(full) and not os.listdir(full):
                os.rmdir(full)
    return removed


def copy(src, dst):
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    shutil.copy2(src, dst)
    return dst


def run(cmd, cwd=None):
    res = subprocess.run(cmd, cwd=cwd, capture_output=True, text=True)
    if res.returncode != 0:
        raise RuntimeError(f"{' '.join(cmd)} failed:\n{res.stdout}\n{res.stderr}")
    return res.stdout.strip()


def normalize_windows_scripts(root):
    """Windows-facing scripts shipped in the package need two byte-level fixes.

    * ``.bat`` / ``.cmd`` / ``.ps1`` must use CRLF: a batch file with bare LF works in most
      cases but is not what cmd.exe expects, and it is not worth betting a first-run
      experience on.
    * PowerShell 5.1 only decodes a ``.ps1`` as UTF-8 when it starts with a BOM. Without
      one it falls back to the ANSI code page, which turns the Chinese half of the
      installer's bilingual messages into mojibake.
    """
    touched = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames.sort()
        for fn in sorted(filenames):
            ext = os.path.splitext(fn)[1].lower()
            if ext not in (".bat", ".cmd", ".ps1"):
                continue
            full = os.path.join(dirpath, fn)
            data = open(full, "rb").read()
            data = data.replace(b"\r\n", b"\n").replace(b"\r", b"\n").replace(b"\n", b"\r\n")
            if ext == ".ps1" and not data.startswith(b"\xef\xbb\xbf"):
                data = b"\xef\xbb\xbf" + data
            open(full, "wb").write(data)
            touched.append(os.path.relpath(full, root).replace("\\", "/"))
    return touched


def zip_tree(src, archive, prefix):
    if os.path.exists(archive):
        os.remove(archive)
    with zipfile.ZipFile(archive, "w", zipfile.ZIP_DEFLATED, compresslevel=9) as zf:
        for dirpath, dirnames, filenames in os.walk(src):
            dirnames.sort()
            for fn in sorted(filenames):
                full = os.path.join(dirpath, fn)
                rel = os.path.relpath(full, src)
                if is_junk(rel):
                    continue
                zf.write(full, os.path.join(prefix, rel).replace("\\", "/"))
    return archive


def main():
    if not os.path.isdir(REPO):
        sys.exit("repo/ not found - run this from the quicklook workspace.")
    static = os.path.join(ROOT, "packaging")
    if not os.path.isfile(os.path.join(static, "README.md")):
        sys.exit(f"hand-written package files missing: {static}\\README.md")
    if os.path.isdir(PKG):
        shutil.rmtree(PKG)
    os.makedirs(PKG)
    log(f"assembling {PKG}")

    # 0. hand-written files (README / INSTALL / CHANGELOG / install.bat / rollback.bat)
    log("[0/7] hand-written package files")
    for dirpath, dirnames, filenames in os.walk(static):
        dirnames[:] = sorted(d for d in dirnames if d != "__pycache__")
        for fn in sorted(filenames):
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, static)
            if is_junk(rel):
                continue
            copy(full, os.path.join(PKG, rel))
            log(f"  {rel.replace(os.sep, '/')}")

    for rel in strip_junk(PKG):
        log(f"  (stripped junk) {rel}")

    for rel in normalize_windows_scripts(PKG):
        log(f"  (normalized) {rel}")

    # 1. source ---------------------------------------------------------------
    log("[1/7] source files -> dev/src/")
    for rel in SRC_FILES:
        src = os.path.join(REPO, *rel.split("/"))
        if not os.path.isfile(src):
            sys.exit(f"  missing source: {src}")
        copy(src, os.path.join(PKG, "dev", "src", *rel.split("/")))
        log(f"  dev/src/{rel}")

    # 2. patch ----------------------------------------------------------------
    log("[2/7] patch vs upstream -> dev/patch/")
    patch_dir = os.path.join(PKG, "dev", "patch")
    os.makedirs(patch_dir, exist_ok=True)
    commit = run(["git", "rev-parse", "--short", BRANCH], cwd=REPO)
    base = run(["git", "rev-parse", BASE_REF], cwd=REPO)
    patch_text = run(["git", "format-patch", f"{base}..{BRANCH}", "--stdout"], cwd=REPO)
    patch_path = os.path.join(patch_dir, "0001-office-preview-optimization-protected-view.patch")
    with open(patch_path, "w", encoding="utf-8", newline="\n") as fh:
        fh.write(patch_text + "\n")
    log(f"  0001-...patch  ({len(patch_text)} chars, {base[:12]}..{commit})")
    with open(os.path.join(patch_dir, "BASE-COMMIT.txt"), "w", encoding="utf-8") as fh:
        fh.write(
            "Apply the patch on top of this upstream commit:\n\n"
            f"  upstream   : https://github.com/QL-Win/QuickLook.git\n"
            f"  base commit: {base}\n"
            f"  patched    : {commit} (branch {BRANCH})\n\n"
            "  git clone https://github.com/QL-Win/QuickLook.git\n"
            "  cd QuickLook && git checkout " + base + "\n"
            "  git apply /path/to/0001-office-preview-optimization-protected-view.patch\n"
        )
    log(f"  BASE-COMMIT.txt (base {base[:12]})")

    # 3. binaries -------------------------------------------------------------
    log("[3/7] binaries -> plugin/  plugin-stock/")
    built = os.path.join(REPO, "Build", "Release", "QuickLook.Plugin",
                         "QuickLook.Plugin.OfficeViewer", DLL_NAME)
    if not os.path.isfile(built):
        sys.exit(f"  build output missing: {built}\n  build the plugin first.")
    copy(built, os.path.join(PKG, "plugin", DLL_NAME))
    log(f"  plugin/{DLL_NAME}  ({os.path.getsize(built)} bytes) "
        f"sha256 {sha256(built)[:16]}")

    install = os.path.join(os.environ.get("LOCALAPPDATA", ""), "Programs", "QuickLook")
    pdir = os.path.join(install, "QuickLook.Plugin", "QuickLook.Plugin.OfficeViewer")
    stock_src = os.path.join(pdir, DLL_NAME + ".stock-4.5.0")
    if os.path.isfile(stock_src):
        copy(stock_src, os.path.join(PKG, "plugin-stock", DLL_NAME))
        log(f"  plugin-stock/{DLL_NAME}  ({os.path.getsize(stock_src)} bytes) "
            f"sha256 {sha256(stock_src)[:16]}")
    else:
        log("  (official stock DLL not found - plugin-stock/ will be MISSING)")

    # 4. config ---------------------------------------------------------------
    # The settings template is hand-written in packaging/config/ and was already copied by
    # step 0 - deliberately NOT taken from the machine's installed config, because a template
    # must list every key (including the ones this machine never happened to write).
    log("[4/7] config template -> config/")
    cfg_dst = os.path.join(PKG, "config", "QuickLook.Plugin.OfficeViewer.config")
    if not os.path.isfile(cfg_dst):
        sys.exit(f"  missing hand-written template: {cfg_dst}")
    log(f"  config/...  ({os.path.getsize(cfg_dst)} bytes, hand-written template)")

    # 5. docs -----------------------------------------------------------------
    log("[5/7] docs -> dev/docs/")
    for d in DOCS:
        p = os.path.join(ROOT, d)
        if os.path.isfile(p):
            copy(p, os.path.join(PKG, "dev", "docs", d))
            log(f"  dev/docs/{d}")
        else:
            log(f"  (missing, skipped) {d}")

    # 6. tools ----------------------------------------------------------------
    log("[6/7] tools -> dev/tools/")
    for t in TOOL_FILES:
        p = os.path.join(ROOT, "tools", t)
        if os.path.isfile(p):
            copy(p, os.path.join(PKG, "dev", "tools", t))
    sm_src = os.path.join(ROOT, "tools", "shortcut-maker", "Program.cs")
    sm_proj = os.path.join(ROOT, "tools", "shortcut-maker", "ShortcutMaker.csproj")
    for p in (sm_src, sm_proj):
        if os.path.isfile(p):
            copy(p, os.path.join(PKG, "dev", "tools", "shortcut-maker", os.path.basename(p)))
    log(f"  dev/tools/  {len(os.listdir(os.path.join(PKG, 'dev', 'tools')))} entries")

    # 7. manifest -------------------------------------------------------------
    # Format is exactly `sha256sum` compatible: "<64 hex><2 spaces><path>".
    # Do NOT put the size between hash and path - that breaks `sha256sum -c`.
    log("[7/7] manifest")
    rows = []
    for dirpath, dirnames, filenames in os.walk(PKG):
        dirnames.sort()
        for fn in sorted(filenames):
            full = os.path.join(dirpath, fn)
            rel = os.path.relpath(full, PKG).replace("\\", "/")
            if rel == "MANIFEST.txt" or is_junk(rel):
                continue
            rows.append((rel, os.path.getsize(full), sha256(full)))
    lines = [
        f"# {PKG_NAME} {VERSION}",
        "# sha256 manifest - verify from inside this directory with:",
        "#   sha256sum -c MANIFEST.txt        (Linux / Git Bash / macOS)",
        "#   certutil -hashfile <file> SHA256 (Windows, one file at a time)",
        f"# files: {len(rows)}   total bytes: {sum(r[1] for r in rows)}",
    ]
    lines += [f"{h}  {rel}" for rel, sz, h in rows]
    with open(os.path.join(PKG, "MANIFEST.txt"), "w", encoding="utf-8", newline="\n") as fh:
        fh.write("\n".join(lines) + "\n")
    log(f"  MANIFEST.txt ({len(rows)} files, sha256sum-compatible)")

    # full zip ----------------------------------------------------------------
    zip_path = zip_tree(PKG, os.path.join(DIST, f"{PKG_NAME}-{VERSION}.zip"), PKG_NAME)
    log(f"\nfull package zip : {zip_path}  ({os.path.getsize(zip_path)} bytes)")

    # plugin-only zip ---------------------------------------------------------
    # What a normal user actually wants: the DLL + the two .bat files.
    log("plugin-only zip  : assembling")
    tree = os.path.join(DIST, "_plugin_only")
    if os.path.isdir(tree):
        shutil.rmtree(tree)
    members = [
        "INSTALL.txt", "install.bat", "rollback.bat",
        f"plugin/{DLL_NAME}", f"plugin-stock/{DLL_NAME}",
    ]
    for rel in members:
        src = os.path.join(PKG, *rel.split("/"))
        if not os.path.isfile(src):
            sys.exit(f"  plugin-only: missing {rel}")
        copy(src, os.path.join(tree, *rel.split("/")))
    plugin_zip = zip_tree(tree, os.path.join(DIST, f"{PLUGIN_ZIP_NAME}.zip"), PLUGIN_ZIP_NAME)
    shutil.rmtree(tree)
    log(f"plugin-only zip  : {plugin_zip}  ({os.path.getsize(plugin_zip)} bytes)")

    # git bundle --------------------------------------------------------------
    bundle = os.path.join(DIST, "quicklook-officeviewer-turbo.bundle")
    if os.path.exists(bundle):
        os.remove(bundle)
    try:
        run(["git", "bundle", "create", bundle, BRANCH], cwd=REPO)
        log(f"git bundle       : {bundle}  ({os.path.getsize(bundle)} bytes)")
    except RuntimeError as exc:
        log(f"git bundle       : SKIPPED ({exc})")

    log("\npackage contents:")
    for rel, sz, h in rows:
        log(f"  {rel:72} {sz:>9}")


if __name__ == "__main__":
    main()
