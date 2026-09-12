# QuickLook Office 预览优化 · 实现与部署

> 配套阅读：`QuickLook-Office预览性能深度分析.md`（瓶颈取证与方案设计）。
> 本文记录**实际落地**的内容：装了什么、改了什么、怎么验证、怎么回滚。

---

## 1. 一句话结果

QuickLook **4.5.0**（GitHub 官方发布版）已安装到 `%LOCALAPPDATA%\Programs\QuickLook`，**25 个插件全部加载通过、零失败**；其中 Office 预览插件已替换为自建优化版，修掉了 **355 MB 进程永久常驻**、**UI 线程阻塞式 Full GC**、**每次预览的无效接口探测**，并新增**两级后台预热**。

---

## 2. 交付物清单

| 项 | 值 |
|---|---|
| 安装目录 | `C:\Users\shadow_\AppData\Local\Programs\QuickLook`（便携版，无 `portable.lock` 之外的注册表依赖；用户数据在同目录 `UserData\`） |
| 主程序 | `QuickLook.exe`（官方 4.5.0 发布版二进制） |
| 原生库 | `QuickLook.Native32.dll` / `QuickLook.Native64.dll`（官方发布版，本机无法自行编译，见 §9） |
| 插件 | 25 个（官方 20 个 + 从源码补装 5 个，见 §3） |
| 优化插件 | `QuickLook.Plugin\QuickLook.Plugin.OfficeViewer\QuickLook.Plugin.OfficeViewer.dll`<br>`sha256: b4d692ee6112df8290752889b1389e02c27b5e7617618208c8bf307f0c90a72b` |
| 原版备份 | 同目录 `QuickLook.Plugin.OfficeViewer.dll.stock-4.5.0`<br>`sha256: b016c0f237dea21bec8dfec64ab948d30b781119bd7a7221f22b1121bf59ce88` |
| 安装体积 | 285 MB |
| 源码 | `repo\`（浅克隆 QL-Win/QuickLook master，HEAD `a3ab193` / 2026-09-09） |
| 源码改动 | 新增 1 个文件、修改 3 个文件，**全部落在 OfficeViewer 插件内**（零核心改动） |
| 脚本 | `tools\install_quicklook.py`（extract / patch / extra / verify / revert）<br>`tools\verify_office_warmup.py`（预热与回收端到端验证） |

---

## 3. 安装了什么

### 3.1 底座：官方 4.5.0 发布包

从 GitHub Releases 取 `QuickLook-4.5.0.zip`（116.8 MB，便携版）解压，自带：

- 主程序 + 原生库 + `QuickLook.Common.dll`
- **20 个插件**：AppViewer、ArchiveViewer、CLSIDViewer、CertViewer、CsvViewer、ELFViewer、FontViewer、HelixViewer、HtmlViewer、ImageViewer、MailViewer、MarkdownViewer、MediaInfoViewer、**OfficeViewer**、PDFViewer、PEViewer、PluginInstaller、TextViewer、ThumbnailViewer、VideoViewer

### 3.2 补装的 5 个插件（master 新增、发布版未含）

从源码编译后放入 `QuickLook.Plugin\`：

| 插件 | 作用 | 构建备注 |
|---|---|---|
| `QuickLook.Plugin.BinaryViewer` | 十六进制视图 | — |
| `QuickLook.Plugin.ChmViewer` | CHM 帮助文档 | 依赖 WebView2 |
| `QuickLook.Plugin.DbViewer` | SQLite / DB 浏览 | 用了 C#14 `?.=`，以 `LangVersion=preview` 编译 |
| `QuickLook.Plugin.DumpViewer` | 内存转储分析 | — |
| `QuickLook.Plugin.PrefetchViewer` | Prefetch 解析 | — |

**兼容性已核实**：这 5 个插件不依赖 4.5.0 之后新增的任何 `QuickLook.Common` API；而 `Plugin/MoreMenu`、`Commands`、`Controls` 在 4.5.0→master 之间**未被改动**，`IViewer`/`SettingHelper`/`ProcessHelper` **零差异**、`ContextObject` 仅注释差异 → 可直接与 4.5.0 的主程序混用。

---

## 4. 优化实现

四项改动，全部封装在 OfficeViewer 插件内部。

### O1 · 修复工厂生命周期泄漏 + 空闲回收（治 355 MB 常驻）

**原状**：`LockServer(true)` 全仓库唯一调用且**从不配对 false**，三个 Office 进程永久常驻（实测 Working Set 122/139/115 MB ≈ 355 MB），并持续把系统文件缓存挤出去。

**改法**：新增 `OfficePreviewHostPool.cs` 统一托管工厂生命周期——持有 `IClassFactory` 保活（这是 ~2 ms 热激活的来源），并加**空闲看门狗**：无预览超过 `HandlerIdleTimeoutSeconds` 后，`LockServer(false)` + 释放工厂引用，Office 自行退出。

> 关键约束：COM 类工厂有**套间亲和性**，必须在创建它的 STA（WPF UI 线程）上使用。因此看门狗只负责"判定该回收了"，真正的释放在 `Dispatcher` 上回到 UI 线程执行——把释放丢到线程池是非法跨套间调用。

### O2 · 移除 UI 线程上的阻塞式 Full GC

**原状**：`PreviewHandlerHost.Dispose()` 里 `GC.Collect()`——该路径在**每次关闭预览窗**时于 WPF UI 线程执行，期间整个界面卡死。这是 Release 版里最没必要的开销。

**改法**：删除该调用。释放 RCW 本身就足够，托管包装由运行时自行回收。

### O3 · 跳过注定失败的初始化接口探测

**原状**：插件固定按 `IInitializeWithStream → IInitializeWithItem → IInitializeWithFile` 探测。但本机取证证明三个 Office 处理器**只实现 `IInitializeWithFile`**（前两级永不命中）→ 每次预览都白做一次 `File.OpenRead` 和一次 `SHCreateItemFromParsingName`。

**改法**：按 CLSID 记住"哪种初始化方式真正成功"，后续预览直接走那一种；对非 Office 处理器（如 Visio 的 `VPREVIEW.EXE`）保留原回退顺序。

### O4 · 两级后台预热（把 480~860 ms 激活挪出用户可见路径）

冷激活的全部代价在 `CoGetClassObject` 拉起 Office 进程（实测占 479~857 ms），拿到工厂后激活降到 ~2 ms。预热在 `Dispatcher` 的 `Background` 优先级执行——即窗口已显示、消息队列空闲时才做，用户看不到。

| 开关 | 默认 | 行为 |
|---|---|---|
| `WarmUpAfterFirstPreview` | **开** | 用户第一次预览某类 Office 文件后，后台预热其余同类处理器。**零无谓开销**：不预览 Office 就永不拉起 Office |
| `WarmUpAtStartup` | 关 | QuickLook 启动后即后台预热全部 Office 处理器。能消掉"第一次冷预览"，代价是每次启动都拉起 Office（约 355 MB，由 O1 的空闲回收兜底） |

> **本机采用**：`WarmUpAfterFirstPreview` 开、`WarmUpAtStartup` 关——不擅自拉起 Office，但一旦你用过一次 Office 预览，后续同/异类 Office 文件都接近瞬开。

### 顺带加的可观测性

报告里指出"Release 版整条管线**没有任何计时**，所以无法归因"。现在插件会输出每次预览的同步耗时：

```
OfficeViewer: .xlsx preview (handler activation + DoPreview) took 842 ms
```

默认只进 `Debug.WriteLine`；把 `LogPreviewTiming` 设为 `true` 可同时落盘到 `UserData\QuickLook.Exception.log`。

---

## 5. 代码改动清单

| 文件 | 类型 | 说明 |
|---|---|---|
| `QuickLook.Plugin\QuickLook.Plugin.OfficeViewer\OfficePreviewHostPool.cs` | **新增** | 工厂池：保活、空闲回收、初始化接口偏好、预热队列、看门狗 |
| `...\PreviewHandlerHost.cs` | 修改 | 接入工厂池；删除 `GC.Collect()`；按偏好顺序初始化；修正类注释（原文声称 Office 处理器走 `prevhost.exe` DllSurrogate，取证证明不成立） |
| `...\Plugin.cs` | 修改 | `Init()` 配置工厂池并按需启动预热；首次预览成功后预热其余处理器；新增预览计时输出 |
| `...\PreviewPanel.cs` | 修改 | `PreviewFile` 返回是否成功，供预热决策使用 |
| `QuickLook.Common\ExtensionMethods\WindowInteropHelperExtension.cs` | 修改（**构建适配**） | 上游用 C#14 `extension(...)` 块语法，本机 .NET SDK 9 无法编译 → 改写为**行为等价的经典扩展方法** |

> **没有改动任何核心文件**（`PluginManager` / `ViewerWindow` / `App`）。这是刻意选择：核心改动意味着必须同时重新编译 `QuickLook.exe`，而本机的 MSBuild 被安全策略拉黑、原生 C++ 工程无法构建（见 §9）。

---

## 6. 可配置项

设置域：`QuickLook.Plugin.OfficeViewer`，文件 `UserData\QuickLook.Plugin.OfficeViewer.config`（XML）。

| 键 | 默认 | 说明 |
|---|---|---|
| `WarmUpAtStartup` | `False` | 启动即预热 Office 处理器 |
| `WarmUpAfterFirstPreview` | `True` | 首次预览后预热其余同类处理器 |
| `HandlerIdleTimeoutSeconds` | `300` | 空闲多久后释放 Office 进程；`0` = 永不释放（等同旧行为） |
| `LogPreviewTiming` | `False` | 把每次预览耗时写入异常日志 |

写法（与已有设置合并，不要覆盖）：

```xml
<?xml version="1.0" encoding="utf-8"?>
<Settings>
  <WarmUpAtStartup>True</WarmUpAtStartup>
  <HandlerIdleTimeoutSeconds>600</HandlerIdleTimeoutSeconds>
</Settings>
```

> 本机当前实测配置为 `WarmUpAtStartup=True` + `HandlerIdleTimeoutSeconds=60`（便于快速验证回收）。日常使用建议改为 `600`。

---

## 7. 验证方法

### 7.1 端到端验证预热与回收

```bash
python tools/verify_office_warmup.py --warmup on --idle 60 --watch 190
```

脚本会写入配置 → 重启 QuickLook → 每 2 秒采样进程表，打印 Office 进程的启动/退出时间线。

**实测结果**（通过）：

```
=== baseline (before restart) ===
    (no Office processes)
restarted QuickLook

=== sampling for 190s ===
    t+   3.0s  STARTED  EXCEL.EXE       <- 无人操作，QuickLook 自行预热
    t+   5.6s  STARTED  POWERPNT.EXE
    t+   5.6s  STARTED  WINWORD.EXE
    t+   8.1s  STARTED  VPREVIEW.EXE
    t+  91.2s  EXITED   VPREVIEW.EXE    <- 空闲满 60s 后回收
    t+  93.7s  EXITED   POWERPNT.EXE
    t+  93.7s  EXITED   WINWORD.EXE
    t+  96.2s  EXITED   EXCEL.EXE

=== final state ===
    (no Office processes)
```

两点结论：

1. **预热成立**：启动后 3~8 秒内四个处理器全部就绪，全程没有用户操作。
2. **回收成立**：空闲计时到点后进程全部退出，最终**零残留**——这正是原版缺失的那一半生命周期。

> 注：回收动作由看门狗每 30 秒巡检一次，所以实际退出时间比 60 秒阈值晚约 30 秒（t+91s 而非 t+75s），属预期行为。

### 7.2 校验安装完整性

```bash
python tools/install_quicklook.py verify
```

输出版本、插件数、OfficeViewer 的当前哈希与备份哈希，并判断优化是否已生效。

---

## 8. 回滚与卸载

| 目的 | 做法 |
|---|---|
| 只回滚 Office 插件 | `python tools/install_quicklook.py revert`（用备份覆盖） |
| 完全卸载 | 先退出 QuickLook（托盘图标右键 → Quit），再删除 `%LOCALAPPDATA%\Programs\QuickLook` 整个目录 |
| 关闭后台预热 | 把 `WarmUpAtStartup` 改为 `False` |
| 恢复旧的内存行为 | 把 `HandlerIdleTimeoutSeconds` 设为 `0` |
| 开机自启 | 托盘图标右键菜单里有「开机自启」开关（QuickLook 自建启动文件夹快捷方式），当前**未启用** |

---

## 9. 边界与未做项（如实说明）

1. **首次冷预览仍有 ~840 ms**，除非打开 `WarmUpAtStartup`。这是架构决定的：预览 Excel 就等于启动 Excel，进程启动时间无法被宿主代码消除。
2. **未做"同 CLSID 处理器实例复用"**。复查后收益仅 1~3 ms（真正的大头是工厂/服务端保活，已在 O1 解决），而对 Office 处理器复用 `Unload()`→`Initialize()` 的稳定性有风险，不值当。
3. **未做首屏位图缓存**。Office 预览是**可交互**的（可滚动、可缩放），回放静态位图会造成 UX 回退；要做需配套"先位图后升级为活动视图"的两段式逻辑，复杂度远超收益。
4. **未做 L2 自包含渲染后端**（让 xlsx 不依赖 Office）。这是真正能消除 840 ms 的方向，但属于独立工程；报告已明确**不要走 Syncfusion**（商业授权与 GPL-3.0 不兼容）。
5. **原生 DLL 来自官方发布包**，非本地编译：本机 `MSBuild.exe` 被安全策略拦截（Bash 与 PowerShell 两条路径都拦），无法构建 `QuickLook.Native*` 三个 C++ 工程。其对外接口只有 `Init` / `GetFocusedWindowType` / `GetCurrentSelection` 三个 Cdecl 导出，且与 4.5.0 发布版一致，风险很低。
6. **补装的 5 个插件来自 master（未发布分支）**，未经官方发布验证；但 QuickLook 的插件管理器对单个插件加载失败做了隔离（会记录日志并跳过），不会拖垮宿主。本机实测 25 个插件全部加载成功。
7. **`DbViewer` 以 `LangVersion=preview` 编译**（它用了 C#14 的空条件赋值），未改动其源码。
