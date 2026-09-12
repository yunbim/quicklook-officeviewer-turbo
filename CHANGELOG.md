# CHANGELOG

本项目为 [QL-Win/QuickLook](https://github.com/QL-Win/QuickLook) 的**非官方插件级补丁**，
版本号格式为 `<上游版本>-turbo.<修订>`。

---

## 4.5.0-turbo.3 — 2026-09-12

**这一版改「怎么拿到、怎么装上」和「界面文案」，预览行为本身没有变化。**

### 新增：界面文案中英文两版

- 新增 `UiText.cs`：本补丁自己绘制的全部文字（受保护视图提示条、按钮、标题栏前缀）
  集中在这里，中英各一份。
- 默认**跟随 Windows 界面语言**（中文系统中文、其他系统英文）；可用新设置项
  `UiLanguage`（`auto` / `zh` / `en`）强制指定。设置项本身有中文说明。
- 新增 English README（`README.en.md`），中文 README 顶部有语言切换链接。

### 安装：双击一个文件，零依赖

- 新增 `install.bat` / `install.ps1` 与 `rollback.bat` / `rollback.ps1`：
  自动定位 QuickLook 安装目录 → 关闭正在运行的 QuickLook → 备份当前 DLL →
  替换 → **逐字节校验** → 自动重启。QuickLook 位置不常规时会提示粘贴路径，
  也支持 `install.ps1 "D:\Apps\QuickLook"` 与 `-DryRun`。
- `.bat` 只做入口（ASCII，避免控制台编码问题），逻辑放在 Windows 自带的
  PowerShell 里；两个 `.ps1` 带 UTF-8 BOM，保证中文提示在 PowerShell 5.1 下不乱码。
- **两个脚本都在沙盒目录里实测过**：安装后 DLL 哈希 == 包内哈希，备份哈希 ==
  官方原版哈希；`rollback.ps1` 能把 DLL 还原成官方原版哈希；`-DryRun` 不改动任何文件。
- 明确并写进文档的关键事实：**全包只有一个文件需要部署**
  （`QuickLook.Plugin.OfficeViewer.dll`）。官方发布包里的
  `UnblockZoneIdentifier.dll` 原装已有，**不应**覆盖。

### 仓库结构重排，README 精简

- 根目录变成「一眼能看懂该干什么」：`README.md` / `README.en.md` / `INSTALL.txt` /
  `install.bat` / `rollback.bat` / `plugin/` / `plugin-stock/` / `config/`，
  开发材料全部下沉到 `dev/`。原先平铺的 `src/ patch/ docs/ tools/ install/ bin/ bin-stock/`
  已被 `dev/` 收拢。
- `README.md` 把「30 秒安装」提到最前，深度分析下沉到 `dev/docs/`。
- 旧版 `install/install.py` / `install/install.ps1` 移除，功能由 `install.ps1` 覆盖；
  安装完整性校验仍可用 `dev/tools/install_quicklook.py verify`。
- `config/` 改为手写模板，列出全部 5 个键（含 `UiLanguage`）和中文注释，不再从本机
  已安装配置里抄（那样会漏掉本机没写过的键）。

### 发布

- 新增 **GitHub Release**，附三个资产：单独一个
  `QuickLook.Plugin.OfficeViewer.dll`（功能全在这一个文件里，直接拿走就能用）、
  `...-plugin-only.zip`（约 30 KB）与 `...-turbo.3.zip`（完整包，含源码 / 补丁 / 文档 / 工具）。

---

## 4.5.0-turbo.2 — 2026-09-12

**改名为「极速预览版（Turbo）」**，并把文档重心改为如实描述它解决的两件事：

1. **连续快速预览多个 Office 文件时的性能问题** —— 目标是"连续秒开"；
2. **受保护视图（Protected View）文档的预览方式** —— 直接出图，且不改动原文件。

### 新增：受保护视图「副本预览」

- 带 `Zone.Identifier`（来自 Internet）的文档不再弹确认框，改为**渲染临时副本**，
  原文件零改动。
- 预览页顶部新增提示条（`ProtectedViewBanner`）：琥珀色说明 + 「解除源文件保护视图」
  按钮；点击后摘除**源文件**标识，提示条变绿。
- 移除原版 `Plugin.cs` 中的 Yes/No 阻塞式 `MessageBox`（代码级删除，非"改默认值"）。
- 新增 `ProtectedViewPreview.cs`：副本创建、去区域标识、临时目录清理与陈旧副本清扫。

### 修复

- 取证脚本 DPI 感知缺陷：屏幕缩放 200% 时，未声明 DPI 感知的抓图程序只拿到窗口
  左上角 1/4，导致版式误判。已在验证工具中修正（`SetProcessDpiAwarenessContext`）。
- 上一版文档中"Office 处理器走 `prevhost.exe` DllSurrogate"的描述错误，已更正
  （实测 `LocalServer32` 直接指向 Office 主程序）。

### 打包

- 包名改为 `QuickLook-OfficeViewer-Turbo`。
- `MANIFEST.txt` 改为**严格 `sha256sum -c` 兼容格式**（此前误将文件大小插入哈希与路径
  之间，会破坏校验）。

---

## 4.5.0-turbo.1 — 2026-09-12

### 新增：Office 预览性能优化（O1–O4）

- **O1 工厂保活 + 空闲回收**：新增 `OfficePreviewHostPool.cs`。热激活 ≈2 ms 的来源；
  同时修复 `LockServer(true)` 从不配对释放导致的 Word/Excel/PPT **永久常驻 ≈355 MB**。
  空闲超过 `HandlerIdleTimeoutSeconds` 后归还进程（看门狗判定、UI 线程执行释放，
  遵守 COM 类工厂的 STA 亲和性）。
- **O2** 删除 `PreviewHandlerHost.Dispose()` 中在 UI 线程执行的阻塞式 `GC.Collect()`。
- **O3** 按 CLSID 记住有效的 `IInitialize*` 方式，跳过注定失败的两级探测
  （Office 处理器只实现 `IInitializeWithFile`）。
- **O4** 两级后台预热（`WarmUpAfterFirstPreview` 默认开、`WarmUpAtStartup` 默认关），
  把 480~860 ms 的冷激活移出用户可见路径。
- 新增每次预览的耗时输出（`LogPreviewTiming` 可落盘）。

### 构建适配

- `QuickLook.Common/ExtensionMethods/WindowInteropHelperExtension.cs`：上游使用 C#14
  的 `extension(...)` 块语法，.NET SDK 9 无法编译 → 改写为**行为等价**的经典扩展方法。
  属纯构建适配，无功能变化。

### 说明

- 未改动任何 QuickLook 核心文件（`PluginManager` / `ViewerWindow` / `App`），
  因此可直接与官方 4.5.0 主程序混用，无需重编 `QuickLook.exe`。
