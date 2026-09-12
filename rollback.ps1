# =====================================================================
#  QuickLook Office 极速预览版 (Turbo) - rollback  /  还原官方原版
#
#  Restores the ORIGINAL QuickLook.Plugin.OfficeViewer.dll, i.e. undoes
#  install.ps1 completely. Uses, in order of preference:
#
#    1. the backup made at install time  (DLL.official-backup)
#    2. the stock 4.5.0 DLL shipped in this package  (plugin-stock\)
#
#  Normally you do NOT run this directly - double-click  rollback.bat  instead.
#  Manual use:
#      powershell -NoProfile -ExecutionPolicy Bypass -File rollback.ps1
#      powershell -NoProfile -ExecutionPolicy Bypass -File rollback.ps1 "D:\Apps\QuickLook"
# =====================================================================

[CmdletBinding()]
param(
    [Parameter(Position = 0)][string]$QuickLookPath,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

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
$Stock = Join-Path $Here "plugin-stock\$DllName"

Say ''
Say '  ==========================================================' 'White'
Say (T '    QuickLook Office Turbo  -  还原官方原版' '    QuickLook Office Turbo  -  rollback') 'White'
Say '  ==========================================================' 'White'
Say ''

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
             'Paste the FULL path of your QuickLook folder, then press Enter.')
    Say  (T '（直接回车放弃）' '(just Enter to abort)')
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
    exit 2
}

# ---- pick the source of truth ----------------------------------------------
$backup = "$target.official-backup"
if (Test-Path -LiteralPath $backup) {
    $source = $backup
    $origin = T '安装时创建的备份' 'the backup taken at install time'
}
elseif (Test-Path -LiteralPath $Stock) {
    $source = $Stock
    $origin = T '本包内的官方 4.5.0 DLL' 'the stock 4.5.0 DLL shipped in this package'
}
else {
    Bad (T '没有可还原的来源：既没有安装时的备份，也没有 plugin-stock\ 下的官方 DLL。' `
           'Nothing to roll back from: neither an install-time backup nor plugin-stock\ was found.')
    Bad (T '请重新下载官方 QuickLook 4.5.0 并恢复插件目录。' `
           'Download the official QuickLook 4.5.0 and restore the plugin folder.')
    exit 2
}

if ($DryRun) {
    Say ''
    Warn (T 'DRY RUN：只检查，不做任何改动。' 'DRY RUN: checked only, nothing was changed.')
    Say  (T "将用 $origin 还原：$target" "Would restore $target from $origin")
    exit 0
}

Step (T "正在用 $origin 还原…" "Restoring from $origin...")

# ---- stop QuickLook ---------------------------------------------------------
$wasRunning = $false
$procs = @(Get-Process -Name QuickLook -ErrorAction SilentlyContinue)
if ($procs.Count -gt 0) {
    $wasRunning = $true
    Step (T '正在关闭 QuickLook…' 'Stopping QuickLook...')
    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Milliseconds 800
}
else {
    Ok (T 'QuickLook 当前未运行。' 'QuickLook is not running.')
}

Copy-Item -LiteralPath $source -Destination $target -Force

$srcHash = (Get-FileHash -LiteralPath $source -Algorithm SHA256).Hash
$dstHash = (Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash
if ($srcHash -ne $dstHash) {
    Bad (T '还原后校验不一致，请重试。' 'Restored file does not match the source. Try again.')
    exit 3
}
Ok (T '官方原版 DLL 已还原并校验通过。' 'Original DLL restored and verified.')

if ($wasRunning) {
    Start-Process -FilePath (Join-Path $ql 'QuickLook.exe') -WorkingDirectory $ql
    Ok (T 'QuickLook 已重新启动。' 'QuickLook restarted.')
}

Say ''
Say '  ----------------------------------------------------------' 'DarkGray'
Say (T '  完成。QuickLook 已回到官方的 OfficeViewer 插件。' `
       '  DONE.  QuickLook is back to the official OfficeViewer plugin.') 'White'
Say (T '  提示：配置文件里多出来的 WarmUp* / HandlerIdleTimeoutSeconds 这几个键' `
       '  Tip: the extra WarmUp* / HandlerIdleTimeoutSeconds keys left in the config')
Say (T '       对官方插件无影响，不删也没关系。' `
       '       are simply ignored by the official plugin - no need to remove them.')
Say '  ----------------------------------------------------------' 'DarkGray'
Say ''
exit 0
