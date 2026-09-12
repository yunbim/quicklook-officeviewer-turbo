# =====================================================================
#  QuickLook Office 极速预览版 (Turbo) - installer  /  安装脚本
#
#  Replaces exactly ONE file:
#      QuickLook.Plugin\QuickLook.Plugin.OfficeViewer\QuickLook.Plugin.OfficeViewer.dll
#  The current DLL is backed up first, so rollback.ps1 can undo everything.
#
#  Normally you do NOT run this directly - double-click  install.bat  instead.
#  Manual use:
#      powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1
#      powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 "D:\Apps\QuickLook"
#      powershell -NoProfile -ExecutionPolicy Bypass -File install.ps1 -DryRun
# =====================================================================

[CmdletBinding()]
param(
    [Parameter(Position = 0)][string]$QuickLookPath,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

# ---- language: follow the Windows UI language, overridable ------------------
try { $script:IsZh = (Get-UICulture).TwoLetterISOLanguageName -eq 'zh' }
catch { $script:IsZh = $false }

function T([string]$zh, [string]$en) { if ($script:IsZh) { $zh } else { $en } }
function Say([string]$m, [string]$c = 'Gray') { Write-Host $m -ForegroundColor $c }
function Ok([string]$m) { Say "  [ok] $m" 'Green' }
function Step([string]$m) { Say "  [..] $m" 'Cyan' }
function Bad([string]$m) { Say "  [X]  $m" 'Red' }
function Warn([string]$m) { Say "  [!]  $m" 'Yellow' }

$DllName = 'QuickLook.Plugin.OfficeViewer.dll'
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Src = Join-Path $Here "plugin\$DllName"

Say ''
Say '  ==========================================================' 'White'
Say (T '    QuickLook Office 极速预览版  -  安装' '    QuickLook Office Turbo  -  installer') 'White'
Say '  ==========================================================' 'White'
Say ''

if (-not (Test-Path -LiteralPath $Src -PathType Leaf)) {
    Bad (T "找不到 plugin\$DllName" "Cannot find plugin\$DllName")
    Bad (T '请在解压后的完整文件夹里运行本脚本。' 'Run this from inside the extracted package folder.')
    exit 2
}

# ---- 1. locate QuickLook ----------------------------------------------------
function Find-QuickLook([string]$explicit) {
    if ($explicit) {
        if (-not (Test-Path -LiteralPath $explicit)) { return $null }
        return (Resolve-Path -LiteralPath $explicit).Path.TrimEnd('\')
    }

    $candidates = New-Object System.Collections.Generic.List[string]
    if ($env:LOCALAPPDATA) { $candidates.Add((Join-Path $env:LOCALAPPDATA 'Programs\QuickLook')) }
    if (${env:ProgramFiles}) { $candidates.Add((Join-Path ${env:ProgramFiles} 'QuickLook')) }
    if (${env:ProgramFiles(x86)}) { $candidates.Add((Join-Path ${env:ProgramFiles(x86)} 'QuickLook')) }

    foreach ($c in $candidates) {
        if ($c -and (Test-Path -LiteralPath (Join-Path $c 'QuickLook.exe'))) {
            return $c.TrimEnd('\')
        }
    }
    return $null
}

$ql = Find-QuickLook $QuickLookPath

while (-not $ql) {
    Say ''
    Warn (T '没有在常规位置找到 QuickLook。' 'QuickLook was not found in the usual locations.')
    Say  (T '请粘贴你的 QuickLook 文件夹完整路径（里面有 QuickLook.exe），回车确认。' `
             'Paste the FULL path of your QuickLook folder - the one containing QuickLook.exe - then press Enter.')
    Say  (T '示例：D:\Apps\QuickLook      （直接回车放弃）' 'Example: D:\Apps\QuickLook      (just Enter to abort)')
    Say ''
    $answer = Read-Host (T '  路径' '  Path')
    if ([string]::IsNullOrWhiteSpace($answer)) {
        Bad (T '已取消。' 'Aborted - no path given.')
        exit 2
    }
    $ql = Find-QuickLook $answer.Trim('"')
    if (-not $ql) { Bad (T "该路径下没有 QuickLook.exe：$answer" "No QuickLook.exe in: $answer") }
}

Ok (T "找到 QuickLook：$ql" "QuickLook found at: $ql")

$pluginDir = Join-Path $ql 'QuickLook.Plugin\QuickLook.Plugin.OfficeViewer'
$target = Join-Path $pluginDir $DllName

if (-not (Test-Path -LiteralPath $target -PathType Leaf)) {
    Bad (T '这个 QuickLook 没有 OfficeViewer 插件目录：' 'This QuickLook has no OfficeViewer plugin folder:')
    Say "       $pluginDir"
    Bad (T '请先安装 / 修复官方 QuickLook 4.5.0，再运行本脚本。' `
           'Install / repair the official QuickLook 4.5.0 first, then retry.')
    exit 2
}
Ok (T 'OfficeViewer 插件目录存在。' 'OfficeViewer plugin folder present.')

if ($DryRun) {
    Say ''
    Warn (T 'DRY RUN：只检查，不做任何改动。' 'DRY RUN: checked only, nothing was changed.')
    Say  (T "将要替换：$target" "Would replace: $target")
    exit 0
}

# ---- 2. stop QuickLook (the DLL is locked while it runs) --------------------
$wasRunning = $false
$procs = @(Get-Process -Name QuickLook -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) {
    $wasRunning = $true
    Step (T '正在关闭 QuickLook，以便替换 DLL…' 'Stopping QuickLook so the DLL can be replaced...')
    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
}
else {
    Ok (T 'QuickLook 当前未运行。' 'QuickLook is not running.')
}

# ---- 3. back up the current DLL once ---------------------------------------
$backup = "$target.official-backup"
if (Test-Path -LiteralPath $backup) {
    Ok (T '沿用上次安装时创建的备份。' 'Keeping the backup created by an earlier install.')
}
else {
    Copy-Item -LiteralPath $target -Destination $backup -Force
    Ok (T "已备份当前 DLL  ->  $DllName.official-backup" "Backed up your current DLL  ->  $DllName.official-backup")
}

# ---- 4. install -------------------------------------------------------------
Copy-Item -LiteralPath $Src -Destination $target -Force

# ---- 5. verify --------------------------------------------------------------
$srcHash = (Get-FileHash -LiteralPath $Src -Algorithm SHA256).Hash
$dstHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
if ($srcHash -ne $dstHash) {
    Bad (T '复制完成但校验不一致，请重试或运行 rollback.ps1。' `
           'Copied, but the verification hash does not match. Retry, or run rollback.ps1.')
    exit 3
}
Ok (T '安装完成，已逐字节校验。' 'Installed and verified (byte-identical to the package).')

# ---- 6. bring QuickLook back ------------------------------------------------
if ($wasRunning) {
    Start-Process -FilePath (Join-Path $ql 'QuickLook.exe') -WorkingDirectory $ql
    Ok (T 'QuickLook 已重新启动。' 'QuickLook restarted.')
}

Say ''
Say '  ----------------------------------------------------------' 'DarkGray'
Say (T '  完成。对着 Word / Excel / PowerPoint 文件按空格即可。' `
       '  DONE.  Press SPACE on any Word / Excel / PowerPoint file.') 'White'
Say (T '   * 第 2 次及以后的预览接近瞬开。' '   * 2nd and later previews are near-instant.')
Say (T '   * 下载来的文件直接出图，不弹窗，且不改动你的原文件。' `
       '   * Internet-downloaded files render directly, no popup, original untouched.')
Say (T '   * 想还原官方原版：双击 rollback.bat' '   * To undo: double-click rollback.bat')
Say '  ----------------------------------------------------------' 'DarkGray'
Say ''
exit 0
