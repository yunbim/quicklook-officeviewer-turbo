# QuickLook Office Turbo

**English** | [中文](README.md)

Makes [QuickLook](https://github.com/QL-Win/QuickLook) jump from *"half a second of stall on every Space"* to *"instant"* when you page through a folder of Office documents, and shows Internet-downloaded Protected View files straight away — no confirmation dialog, and your original file is never touched.

> This is an **unofficial, plugin-level patch** for the official **4.5.0** release of
> [QL-Win/QuickLook](https://github.com/QL-Win/QuickLook). It replaces exactly **one file**;
> `QuickLook.exe` and the other 24 plugins stay exactly as upstream ships them, so it can be
> dropped into a stock installation.
>
> **Keywords**: QuickLook Office preview slow · faster Word Excel PowerPoint preview · Protected View popup · Zone.Identifier · preview handler warm activation

---

## ⚡ Install in 30 seconds

> Requires the official [QuickLook 4.5.0](https://github.com/QL-Win/QuickLook/releases) to be
> installed first. This package contains **only the plugin**, not the host application.

1. Grab **`...-plugin-only.zip`** (≈30 KB) from this repo's
   **[Releases page](https://github.com/yunbim/quicklook-officeviewer-turbo/releases)**.
2. Unzip it anywhere and **double-click `install.bat`**. It automatically locates QuickLook,
   stops it, backs up your current DLL, replaces it, **verifies the result byte-for-byte**,
   and restarts QuickLook.
3. Done. Press Space on any Office file.

**Nothing extra to install**: the script uses the PowerShell that ships with Windows (since 7).
No Python. No administrator rights either, unless QuickLook lives under `Program Files`.
If QuickLook is somewhere unusual the script asks you for the path; you can also run
`install.bat "D:\Apps\QuickLook"`.

**To go back to stock**: double-click `rollback.bat`.

<details>
<summary>Prefer doing it by hand? It really is one file.</summary>

1. Right-click the tray icon → **Quit** to exit QuickLook.
2. Copy `plugin\QuickLook.Plugin.OfficeViewer.dll` over:

   ```
   %LOCALAPPDATA%\Programs\QuickLook\QuickLook.Plugin\QuickLook.Plugin.OfficeViewer\
   ```

3. Start QuickLook again.

That is the only file — the official package already ships `UnblockZoneIdentifier.dll`, so
**do not** overwrite that one.
</details>

---

## What it fixes

### A. Every Office preview costs an Office launch

QuickLook's Office plugin **does not parse documents itself** — it hands the file to the
**Windows Shell preview handler**. The handler Office registers points its `LocalServer32`
straight at the Office executable itself (`WINWORD.EXE` / `EXCEL.EXE` / `POWERPNT.EXE`),
with **no `DllSurrogate`**.

So **previewing one Office file = starting one Office process**: measured at
**482–863 ms**. Tolerable for a single file; but the real use case is arrowing through a
folder of documents — and then every keystroke costs half a second, which adds up to tens of
seconds of pure waiting.

### B. Office processes are pinned in memory forever

The stock plugin calls `LockServer(true)` on the COM class factory and **never releases it**,
so the Word/Excel/PowerPoint processes stay resident at roughly **355 MB** until you log out
or kill them by hand.

### C. The viewer stutters every time you close it

`PreviewHandlerHost.Dispose()` ran a blocking `GC.Collect()` on the **WPF UI thread**.

### D. Downloaded files need a dialog — and it rewrites your file

Documents that arrive from a browser, mail client or cloud drive carry a `Zone.Identifier`
mark ("came from the Internet"), and **Office's preview handler refuses to load them**. The
stock plugin therefore shows a Yes/No confirmation, and it acts by **permanently deleting the
`Zone.Identifier` from your original file** — so your file is silently modified.

### Before / after

| Scenario | Stock | Turbo |
|---|---|---|
| First preview of an Office type (cold start, architectural) | 482–863 ms | 482–863 ms (**same**) |
| **Every later preview of the same type** | **482–863 ms, every time** | **≈2 ms** |
| Paging through 20 Excel files | ≈10–17 s of waiting | first ≈0.5–0.9 s, **the rest effectively instant** |
| Mixed Word / Excel / PowerPoint | cold start again per type | after the first one, the rest are **warmed up in the background** |
| Memory after closing the preview | Word/Excel/PPT resident ≈355 MB | released after an idle timeout, **nothing left behind** |
| Closing the viewer | blocking GC stall | removed |
| Internet-downloaded files | dialog + your file rewritten | renders directly + notice strip + opt-in unblock |

> **Why is the first preview still slow?** Because "previewing Excel" *is* "starting an Excel
> process" at the architecture level, and a plugin cannot remove process start-up time. What
> Turbo does is pay that cost **once instead of on every file**, plus it gives you a switch to
> warm everything at startup.

---

## How you use it

**There is nothing new to learn.** Keep pressing Space:

- **Paging through documents**: the very first one is still cold (architectural); after that the
  remaining Office handlers are warmed in the background, so consecutive previews are near-instant.
- **Want even the first one instant?** Set `WarmUpAtStartup` to `True`
  (cost: ≈355 MB resident, reclaimed after the idle timeout).
- **Downloaded documents**: they just render. An amber strip states that you are looking at a
  copy and that your original is untouched. To stop it entering Protected View in Word as well,
  click **"Unblock the original file"** on the right of that strip.
- **Ordinary documents**: byte-for-byte the stock experience — no copy, no strip.

> Config changes need a **QuickLook restart** (settings are cached in memory).
> The tray icon only exists **while running**, and Windows 11 hides it under the `^` overflow.

---

## How it works

Four performance changes plus one Protected View change, **all inside the OfficeViewer
plugin**. No QuickLook core file is touched, so `QuickLook.exe` never needs rebuilding.

### Performance: `OfficePreviewHostPool.cs` (new)

The key is COM **`IClassFactory` keep-alive**: after `LockServer(true)` the class factory keeps
its process alive with it, so as long as the factory is still there the next `CoCreateInstance`
is a **warm in-process activation (≈2 ms)** instead of a fresh Office launch. **That is where
"instant" comes from.**

The stock plugin only ever kept the factory alive and never let go. Turbo adds an **idle
watchdog**: after `HandlerIdleTimeoutSeconds` (default 1800 s) without a preview it calls
`LockServer(false)` and drops the factory reference, and the Office process then exits on its
own.

> **Important constraint**: COM class factories are **apartment-bound (STA)** and must be
> released on the STA that created them. The watchdog thread only *decides* that the pool is
> idle; the actual release is marshalled back to the **WPF UI thread** through the `Dispatcher`.

The other three:

| Change | File | What |
|---|---|---|
| Drop the blocking GC | `PreviewHandlerHost.cs` | Remove `GC.Collect()` from `Dispose()` |
| Skip interfaces that cannot work | `PreviewHandlerHost.cs` | Measured: the three Office handlers implement **only `IInitializeWithFile`**, so the first two probes always fail. The working initialisation interface is remembered per CLSID; non-Office handlers keep the original fallback order |
| Two-stage background warm-up | `Plugin.cs` | Warm-up runs at `DispatcherPriority.Background`, i.e. only once the window is up and the message queue is idle — invisible to you |

### Protected View: three files

| File | Responsibility | Key point |
|---|---|---|
| `ProtectedViewPreview.cs` | Lifecycle of the copy | `ZoneIdentifierManager.IsZoneBlocked()` → copy into `%TEMP%\QuickLook\OfficeViewer\` (hash in the name to avoid collisions) → write an empty `Zone.Identifier` **on the copy only** → render → delete on close; stale copies are swept at start-up |
| `ProtectedViewBanner.cs` | Notice strip + button | Hand-drawn amber `Border` (**independent of QuickLook's theme**, legible in both light and dark); the button unblocks in place and turns the whole strip green |
| `Plugin.cs` | Branching | Marked file → copy path + notice strip; unmarked file → the original path, identical to stock |

**The confirmation dialog is deleted in code** — not "the switch is now on by default". There is
no `MessageBox` left in `Plugin.cs`.

### One honest trade-off

The strip makes the **Protected View** preview area about **43 px** shorter (it occupies its own
row). QuickLook hosts the Office view in a `WindowsFormsHost`, which paints over any WPF sibling
it overlaps — so the strip **cannot be an overlay**. **Ordinary documents are unaffected.**

---

## Settings

Domain `QuickLook.Plugin.OfficeViewer`, file:
`%LOCALAPPDATA%\Programs\QuickLook\UserData\QuickLook.Plugin.OfficeViewer.config`

| Key | Default | Meaning |
|---|---|---|
| `WarmUpAfterFirstPreview` | `True` | After the first preview, warm the other Office handlers in the background |
| `WarmUpAtStartup` | `False` | Warm every Office handler at start-up (cost: resident memory) |
| `HandlerIdleTimeoutSeconds` | `1800` | Release the Office processes after this much idle time; `0` = never (stock behaviour) |
| `LogPreviewTiming` | `False` | Append each preview's duration to `UserData\QuickLook.Exception.log` |
| `UiLanguage` | `auto` | Language of the notice strip: `auto` (follow the Windows UI language) / `zh` / `en` |

Merge into the existing `<Settings>` — do not replace the whole file (a template lives in
`config/`):

```xml
<?xml version="1.0" encoding="utf-8"?>
<Settings>
  <WarmUpAfterFirstPreview>True</WarmUpAfterFirstPreview>
  <WarmUpAtStartup>False</WarmUpAtStartup>
  <HandlerIdleTimeoutSeconds>1800</HandlerIdleTimeoutSeconds>
  <UiLanguage>auto</UiLanguage>
</Settings>
```

---

## Verify it yourself

The scripts under `dev/tools/` do not "look and see if something rendered" — they **read real
pixels and real processes** (Python 3, standard library only):

```bash
python dev/tools/verify_office_warmup.py --warmup on --idle 60 --watch 190  # cold/warm cost + do the processes really exit?
python dev/tools/verify_protected_copy.py                                   # the three copy-preview scenarios, incl. a real mouse click
```

The latter's pass criteria include **"is the Zone.Identifier still on the original file?"** —
a test that only checks "it rendered" cannot tell whether your file was modified.

---

## Rollback

| Goal | How |
|---|---|
| Undo this patch | Double-click `rollback.bat` (or restore `DLL.official-backup` by hand) |
| Back to stock | Copy `plugin-stock\QuickLook.Plugin.OfficeViewer.dll` over |
| Turn warm-up off | `WarmUpAtStartup=False`, `WarmUpAfterFirstPreview=False` |
| Restore stock memory behaviour | `HandlerIdleTimeoutSeconds=0` |
| Uninstall everything | Quit QuickLook, delete `%LOCALAPPDATA%\Programs\QuickLook` |

---

## Limitations (stated honestly)

1. **The first cold preview still costs ~840 ms** unless `WarmUpAtStartup` is on. Architectural.
2. **QuickLook's core is untouched, and there is no Word page/paper view.** The Office preview
   handler renders as a **continuous flow and draws no page boundaries**, so pagination is not
   reachable on this path; it would need "convert to PDF + PdfiumViewer", a separate project.
3. **The copy-preview cost**: about 43 px of preview height on protected documents (see above).
   Ordinary documents are unaffected.
4. **The temporary copy is plaintext in `%TEMP%`**, deleted when the preview closes; if the
   process is killed mid-preview it can linger, and the next start-up sweeps stale copies.
5. **The native libraries `QuickLook.Native*.dll` come from the official release**, not from a
   local build; this package does not include them — keep the official installation.

---

## Repository layout

```
QuickLook-OfficeViewer-Turbo/
├── README.md            Chinese README (this file is README.en.md)
├── INSTALL.txt          Plain install instructions (also inside the plugin-only zip)
├── CHANGELOG.md         Version history
├── LICENSE-GPL.txt      Full upstream GPL-3.0 text (this patch is released under it too)
├── MANIFEST.txt         sha256 manifest of every file
├── install.bat          ★ Double-click to install - the entry point
├── install.ps1            The install logic (uses the PowerShell Windows already has)
├── rollback.bat         ★ Double-click to restore stock
├── plugin/              ★ The one file that gets deployed
├── plugin-stock/        Official 4.5.0 DLL, for an exact rollback
├── config/              Default settings template
└── dev/                 Development & forensics material — safe to ignore
    ├── src/             Patched source, keeping upstream-relative paths
    ├── patch/           git patch vs upstream + base commit
    ├── docs/            Deep analysis, implementation notes, verification report
    └── tools/           Forensics / benchmark / verification scripts
```

---

## Verification

```bash
sha256sum -c MANIFEST.txt          # Linux / Git Bash / macOS
certutil -hashfile <file> SHA256   # Windows, one file at a time
```

The patch in `patch/` was validated against a **clean upstream source tree**: applying it
produces the same 8 files as `dev/src/` (line endings aside), and it passes a **reverse-apply
check** (`git apply -R --check` is clean), proving it is exactly "the complete change set
relative to upstream" with nothing smuggled in.

The patch also passes an **independent rebuild check**: applied to a pristine upstream tree
extracted with `git archive`, `dotnet build` succeeds directly and the result has the same
34816-byte size as the DLL in this package (byte-level hashes differ, which is normal — .NET
assemblies embed a timestamp/MVID and are not deterministic by default).

---

## Origin and licence

- **Upstream**: [QL-Win/QuickLook](https://github.com/QL-Win/QuickLook) — the universal
  press-Space-to-preview tool for Windows, **GPL-3.0**.
- This package is an **unofficial plugin-level patch**. It is not affiliated with upstream and
  is not endorsed by it; get the original from the upstream release page.
- All changes here are released under **GPL-3.0** as well (see `LICENSE-GPL.txt`). The DLL in
  `plugin/` is built from `dev/src/`.
- If you redistribute, ship the source and keep this licence notice.
