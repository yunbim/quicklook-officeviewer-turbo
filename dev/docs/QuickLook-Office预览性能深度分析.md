# QL-Win/QuickLook Office 预览性能深度分析报告

> 分析对象：`QL-Win/QuickLook`（master 分支，CHANGELOG 最新条目为 `4.6.0`）
> 分析方式：源码级追踪 + 本机注册表/COM 取证 + 实测基准（脚本随报告交付、可复现）
> 样本机：Windows，已安装 Microsoft Office（`C:\Program Files\Microsoft Office\root\Office16\`）
> 分析时间：2026-09-12

---

## 0. 结论摘要（TL;DR）

1. **QuickLook 的 Office 插件根本不解析 Office 文件。** `QuickLook.Plugin.OfficeViewer` 只是查注册表找到该扩展名对应的 **Windows Shell 预览处理器 CLSID**，然后把文件"转交"给它。真正解析和渲染 Excel/Word/PPT 的，是 Windows/Office 自己的处理器。

2. **本机取证推翻了插件源码里的关键假设。** `PreviewHandlerHost.cs:42-43` 注释声称 Office 处理器"多数注册了 DllSurrogate，会被 prevhost.exe 托管"。实测：`.xlsx / .docx / .pptx` 的处理器 CLSID **没有 AppID、没有 DllSurrogate**，其 `LocalServer32` **直接指向完整 Office 主程序本体**：

   | 扩展名 | 预览处理器 CLSID | 注册名 | LocalServer32 | InprocServer32 |
   |---|---|---|---|---|
   | `.xlsx` / `.xls` / `.xlsb` | `{00020827-0000-0000-C000-000000000046}` | Microsoft Excel previewer | **`EXCEL.EXE`** | 无 |
   | `.docx` / `.doc` | `{84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}` | Microsoft Word previewer | **`WINWORD.EXE`** | 无 |
   | `.pptx` | `{65235197-874B-4A07-BDC5-E65EA825B718}` | Microsoft PowerPoint previewer | **`POWERPNT.EXE`** | 无 |
   | `.vsdx` | `{21E17C2F-AD3A-4b89-841F-09CFE02D16B7}` | Microsoft Visio previewer | `VPREVIEW.EXE`（专用轻量宿主） | 无 |

   也就是说：**预览一个 Excel 文件 = 启动一次 Excel 主程序**。这是慢的根因。

3. **实测冷启动代价（本机，OS 文件缓存已热）：**

   | 目标 | 冷启动总耗时 | 其中 CoGetClassObject | CreateInstance | 结果 |
   |---|---|---|---|---|
   | Excel 预览器 | **839.0 ms** | 819.8 ms | 18.7 ms | 拉起 `excel.exe` |
   | Word 预览器 | **863.5 ms** | 857.0 ms | 6.4 ms | 拉起 `winword.exe` |
   | PowerPoint 预览器 | **482.3 ms** | 479.1 ms | 3.2 ms | 拉起 `powerpnt.exe` |
   | 三者（进程已存活） | 1.8 / 2.3 / 2.5 ms | ~1.4 ms | ~0.8 ms | 复用已有进程 |

   **冷/热差距约 200~400 倍。** 冷启动那 840ms 纯粹是"把 Office 主程序拉起来"，还没开始读文件。

4. **`CLSCTX_INPROC_SERVER` 实测返回 `0x80040154`（REGDB_E_CLASSNOTREG）**，三个处理器全部失败 → `PreviewHandlerHost.TryCreateInProcHandler()` 对 Office 文件是**永不生效的死代码**。每一次 Office 预览都只能走进程外、只能启动 Office。

5. **三个处理器实测只实现 `IInitializeWithFile`，不实现 `IInitializeWithStream`，也不实现 `IInitializeWithItem`。** 因此插件精心实现的"stream → item → file"三级优先探测中，前两级对 Office 文件**从不命中**；`IStreamWrapper` 这一整套跨进程流封送机制对 Office 文件是死代码（仅对第三方处理器可能有用）。

6. **"保活"是当前唯一被实现的缓存策略，且实现有缺陷：** `HandlerFactories` 是静态字典，写入前调用 `LockServer(true)`，而**全仓库搜不到任何一次 `LockServer(false)`**（仅 `PreviewHandlerHost.cs:248` 一处调用）。结果是 Office 三个进程一旦被预览过就**永久常驻**。实测常驻代价：

   | 进程 | Working Set | Private Bytes |
   |---|---|---|
   | `excel.exe` | 122.0 MB | 61.6 MB |
   | `winword.exe` | 139.3 MB | 68.2 MB |
   | `powerpnt.exe` | 114.9 MB | 54.8 MB |
   | **合计** | **≈ 355 MB** | ≈ 185 MB |

7. **UI 冻结被放大：** `Plugin.View()` 通过 `Dispatcher.BeginInvoke(..., DispatcherPriority.Input)` 在 **WPF UI 线程**执行（`ViewerWindow.Actions.cs:356-393`），而 `PreviewPanel.PreviewFile` → `PreviewHandlerHost.Open`（COM 激活 + `DoPreview()`）**全程同步**。于是界面在"启动 Office + 解析文档"的整个过程中完全冻结，`IsBusy` 转圈动画也无法刷新（WPF 渲染线程被阻塞）。用户感知到的"卡住"= 真实耗时 + 无法反馈的等待。

8. **用户列举的四类候选瓶颈，实测判定：**

   | 候选 | 判定 | 依据 |
   |---|---|---|
   | 插件初始化 | ❌ 不是主因 | 插件 DLL 加载是启动期一次性成本；每预览反射新建实例 <1ms；handler 实例化热态仅 0.5~3ms |
   | 文件解析 | ✅ 主因之一，但被外包给 Office | 解析完全发生在 Office 进程内，由同步 `DoPreview()` 触发；**当前无任何埋点，耗时不可观测** |
   | 缓存策略 | ✅ 主因之一 | 只有"IClassFactory 永久保活"这一种缓存；**无预热、无同文件结果缓存、无同类型处理器复用、无空闲回收** |
   | 资源占用 | ✅ 是，且反噬性能 | 3 个 Office 进程永久常驻 ≈355MB，且不接受回收；在 8GB 机器上会挤出文件缓存、加剧后续加载变慢 |
   | 顺带排除 | 注册表查询 | `.xlsx` 完整查找链实测 **0.1182 ms/次**，可忽略（**不要**把优化力气花在这里） |

---

## 1. 分析方法与可复现证据

本报告的所有量化结论都来自随报告交付的三个脚本（`tools/` 目录），可在任意 Windows + Office 机器上复现：

| 脚本 | 作用 |
|---|---|
| `tools/probe_office_preview.py` | 解析 `HKCR\.<ext>\shellex\{8895b1c6-...}` → CLSID → `InprocServer32`/`LocalServer32`/`AppID`/`DllSurrogate` 完整注册链，判断宿主形态 |
| `tools/bench_com_activation.py` | ctypes 裸 COM：`CoGetClassObject` + `IClassFactory::CreateInstance` 冷/热计时，并枚举被拉起的进程名 |
| `tools/probe_handler_interfaces.py` | `QueryInterface` 探测处理器实现哪些初始化接口与 `IPreviewHandler`/`IObjectWithSite` |
| `tools/probe_warm_server_cost.py` | 保活态下 Office 进程 Working Set / Private Bytes 采样 + 注册表查找耗时 |

运行方式（本机实测所用命令）：

```
python tools/probe_office_preview.py
python tools/bench_com_activation.py
python tools/probe_handler_interfaces.py
python tools/probe_warm_server_cost.py
```

> 说明：`reg.exe` 在本机被安全策略拦截，故取证改由 Python `winreg` 直读注册表。脚本中的 `LockServer(false)` 均被显式调用，不残留后台 Office 进程（已验证探测后无 Office 进程存活）。
> 注意：本报告**没有**测量 `DoPreview()` 的真实耗时（它需要真实 HWND + 消息泵 + 真实文件，属于项目侧埋点范畴）。第 6 节给出了该项的埋点与基准方案，请勿把我给出的激活耗时当作端到端耗时。

---

## 2. QuickLook 整体架构

### 2.1 进程与集成模型

- **常驻单实例进程 `QuickLook.exe`**：`App.EnsureFirstInstance` 用命名 Mutex（`QuickLook.App.Mutex`）保证单实例；第二次启动会把文件路径通过命名管道投递给首实例（`App.xaml.cs:284-314`）。
- **与资源管理器的集成**：`NativeMethods.QuickLook.Init()` 挂钩；按键由 `KeystrokeDispatcher` 捕获空格，`ViewWindowManager.TogglePreview` 取当前选中项（`NativeMethods.QuickLook.GetCurrentSelection()`）。
- **管道**：`PipeServerManager` 接收外部请求（含命令行 `QuickLook.exe <path>` 与 `/autorun` 等）。
- 启动期成本分布（`App.OnStartup`）：`RunListener` → `PluginManager.GetInstance()`（**一次性加载全部插件 DLL**）→ `ViewWindowManager` / `KeystrokeDispatcher` / `PipeServerManager`。

### 2.2 插件体系

契约 `IViewer`：`Init()` / `CanHandle(path)` / `Prepare(path, ctx)` / `View(path, ctx)` / `Cleanup()`，另可用 `Priority` 影响被选中的先后。

`PluginManager`（`QuickLook/PluginManager.cs`）关键行为：

- **启动时全量加载**：`Directory.GetFiles(folder, "QuickLook.Plugin.*.dll", AllDirectories)` 逐个 `Assembly.LoadFrom` + 反射找 `IViewer` 实现并实例化（`:82-113`），随后按 `Priority` **降序**排序（`:115`）。加载目录 = 用户插件目录 + 程序目录 `QuickLook.Plugin\`。
- **每次预览都新建插件实例**：`FindMatch` 结尾 `(matched ?? DefaultPlugin).GetType().CreateInstance<IViewer>()`（`:79`）。→ **架构性约束：任何跨预览的缓存都必须是静态/进程级的，不能挂在插件实例字段上。**
- **匹配是 UI 线程上的顺序线扫**：`FindMatch` 用 `FirstOrDefault` 依 Priority 降序逐个 `plugin.CanHandle(path)`，且每次匹配都建 `Stopwatch` 并 `Debug.WriteLine`（`:57-77`，仅 Debug 版输出）。

**Office 相关优先级排序（降序）**：`ELFViewer(11)` → `ChmViewer(2)` → 一批 `0`（ImageViewer、CsvViewer、HtmlViewer、ThumbnailViewer、MailViewer、MarkdownViewer…）→ `CLSIDViewer/PDFViewer/OfficeViewer(-1)` → `VideoViewer(-3)` → `TextViewer/ArchiveViewer/HelixViewer(-5)` → `BinaryViewer(-10)`。

对 `.xlsx` 而言，OfficeViewer(-1) 之前会被十几个插件 `CanHandle` 检查。经核查：这些插件绝大多数只做扩展名判断（廉价）；`ThumbnailViewer` 的 `WellKnownExtensions` 只含 `.cdr/.fig/.kra/.pdn/.pix/.sketch/.xd/.xmind` 之类，**不抢 Office 扩展名**；`TextViewer`（会读前 16KB 嗅探内容）优先级 -5，在 OfficeViewer 之后，**不会**被执行。→ **匹配链对 Office 文件不是瓶颈。**

### 2.3 预览管线时序（一次空格按下）

```
Explorer 空格
  └─ KeystrokeDispatcher → ViewWindowManager.TogglePreview(path)
       └─ PluginManager.FindMatch(path)              [UI 线程]  线扫全部插件 CanHandle
            └─ (新建插件实例)
       └─ BeginShowNewWindow → ViewerWindow.UnloadPlugin()      [UI 线程]  旧插件 Cleanup()
            └─ ViewerWindow.BeginShow(plugin, path)
                 ├─ plugin.Prepare()  → 决定窗口尺寸        [UI 线程]
                 ├─ PositionWindow() → Show()               [UI 线程]  窗口先出现（空内容 + busy 转圈）
                 └─ Dispatcher.BeginInvoke(…, Priority.Input)
                      └─ plugin.View(path, ctx)             [UI 线程 ← 关键]
```

`ViewerWindow.Actions.cs:356-393` 的这次 `BeginInvoke` 是理解 Office 卡顿的关键：**它只是把工作排到 UI 线程队列尾部，并没有换线程**。因此 `View()` 内部的一切同步 COM 调用都会冻结整个界面。

---

## 3. Office 预览的加载与渲染全链路

### 3.1 第一步：`CanHandle` —— 决定"能不能预览"（纯注册表）

`Plugin.cs:47-100`：

1. 目录直接排除；
2. 扩展名白名单：`.doc .docx .docm .odt .xls .xlsx .xlsm .xlsb .ods .ppt .pptx .odp .vsd .vsdx`（`:31-37`）；
3. `ShellExRegister.GetPreviewHandlerGUID(ext)`（`ShellExRegister.cs:14-38`）：
   - 先查 `HKCR\<ext>\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}`；
   - 若无，再取 `HKCR\<ext>` 默认值作为 ProgID，查 `HKCR\<ProgID>\shellex\{8895b1c6-...}`；
   - 都无 → `Guid.Empty` → **直接放弃**（这就是本机 `.pptm` 无法预览的原因：未注册处理器）。
4. `CLSIDRegister.GetName(clsid)` 读 `HKCR\CLSID\{clsid}` 默认值确认处理器真的存在（可用 `CheckPreviewHandler=False` 跳过）。
5. 实测耗时 **0.1182 ms**（含上述全链）→ **可忽略**。

> 含义：**QuickLook 的 Office 支持本质上是"Windows 预览处理器转发器"，其可用性完全取决于机器上注册了什么处理器。** 没有 Office（或等价处理器）的机器上，这个插件直接返回 false，Office 文件会落到默认的 InfoPanel。

### 3.2 第二步：COM 激活 —— 把 Office 拉起来（瓶颈所在）

`Plugin.cs:155-158` → `PreviewPanel.PreviewFile`（`PreviewPanel.cs:42-47`）：

```csharp
_control = new PreviewHandlerHost();
Child = _control;        // WindowsFormsHost 载体（PreviewPanel.xaml:11）
_control.Open(file);
```

`PreviewHandlerHost.Open`（`PreviewHandlerHost.cs:127-213`）逐段：

| 步骤 | 代码 | 实际发生的事 |
|---|---|---|
| 取 CLSID | `:136` | 再查一次注册表 |
| 优先进程外 | `:144` `TryCreateOutOfProcHandler(guid)` | `CoGetClassObject(clsid, CLSCTX_LOCAL_SERVER, …)`（`:243`）→ **启动 `EXCEL.EXE` 等完整 Office 进程**；成功后 `LockServer(true)` 并塞进静态 `HandlerFactories`（`:247-249`） |
| 回退进程内 | `:144` `TryCreateInProcHandler(guid)` | `Activator.CreateInstance(Type.GetTypeFromCLSID(clsid, true))`（`:284`）→ 实测 **`0x80040154` 必然失败**（无 InprocServer32），**死代码** |
| 初始化（stream） | `:152-165` | `o is IInitializeWithStream` → 实测 QI 失败 → **跳过** |
| 初始化（item） | `:167-180` | `SHCreateItemFromParsingName` + `IInitializeWithItem` → 实测 QI 失败 → **跳过** |
| 初始化（file） | `:182-190` | `IInitializeWithFile.Initialize(path, STGM_READ)` → **唯一命中路径** |
| 挂窗口 | `:208-209` | `SetWindow(Handle, ref rect)`：把 WinForms 控件的 HWND 交给 **Office 进程内的处理器**，由它创建并挂接渲染窗口（跨进程子窗口，即 Shell 预览处理器的标准进程外模型） |
| 触发渲染 | `:210` | `DoPreview()` —— **同步阻塞**，Office 在此加载并渲染文档 |

**实测接口能力（`probe_handler_interfaces.py`）：**

| 接口 | Excel | Word | PPT |
|---|---|---|---|
| `IInitializeWithStream` | ❌ | ❌ | ❌ |
| `IInitializeWithItem` | ❌ | ❌ | ❌ |
| `IInitializeWithFile` | ✅ | ✅ | ✅ |
| `IPreviewHandler` | ✅ | ✅ | ✅ |
| `IObjectWithSite` | ✅ | ✅ | ✅ |
| `CLSCTX_INPROC_SERVER` 可用 | ❌ `0x80040154` | ❌ `0x80040154` | ❌ `0x80040154` |

两条可直接落地的推论：

- **探测顺序对 Office 无意义**：既然只有 `IInitializeWithFile` 命中，就应该把每个 CLSID **首次探测结果缓存下来**，之后直接走命中路径，省掉每次两次必失败的 `QueryInterface`（收益虽小，但属于"零风险清理"）。
- **`IObjectWithSite` 被实现但从未被设置**：处理器支持该接口，说明它期望宿主调用 `SetSite`（资源管理器会调）。QuickLook 未调用。这不是耗时问题，但可能影响部分处理器的行为（如事件/焦点语义），属于**兼容性观察项**。

### 3.3 第三步：渲染发生在 Office 进程内

由于处理器在 `EXCEL.EXE` 进程里执行，实际渲染链路是：

```
QuickLook.exe ──COM(跨进程)──> EXCEL.EXE
                                 └─ 加载 xlsx（解压 OOXML、重建单元格/公式/样式模型）
                                 └─ 按 SetWindow 给定的 rect 创建渲染窗口并绘制首屏
                                 └─ DoPreview() 同步返回（或在工作线程继续细化）
QuickLook.exe <── 窗口由 Office 绘制 ────┘
```

推论（对优化方向很重要）：
- **Excel 的解析成本完全不可控**，包括公式重算、条件格式、图表、外部链接更新等；大工作簿尤其明显。
- **QuickLook 无法对渲染过程做进度反馈或取消**，因为那是另一个进程的同步调用。

### 3.4 第四步：清理（每次切文件都全量重建）

- `ViewWindowManager.BeginShowNewWindow` → `_viewerWindow.UnloadPlugin()` → `Plugin.Cleanup()`（`ViewWindowManager.cs:219-224`）
- OfficeViewer 的 `Cleanup()` → `_panel?.Dispose()`（`Plugin.cs:173-177`）
- `PreviewPanel.Dispose()` → 通过 `Dispatcher.BeginInvoke` 清空 `Child` 并 `_control.Dispose()`（`PreviewPanel.cs:30-40`）
- `PreviewHandlerHost.Dispose()` → `UnloadPreviewHandler()`（调 `IPreviewHandler::Unload`）+ `Marshal.FinalReleaseComObject` + **显式 `GC.Collect()`**（`PreviewHandlerHost.cs:84-99`）

> 注意这里的不对称：**handler 实例被释放了，但承载它的 Office 进程被 `LockServer(true)` 永久钉住**。所以"释放"只释放了对象，没释放资源。

---

## 4. 瓶颈定位（按影响排序）

### P0-1 ｜Office 主程序冷启动：482 ~ 863 ms（硬地板）

- 证据：`bench_com_activation.py` 实测，且 `started=['excel.exe']` 等直接证明启动的是主程序。
- 触发条件：**每个 Office 应用类型、每次会话首次预览**（Office 进程被系统回收后再次预览同样触发；重启/更新/清理工具杀进程后亦然）。
- 叠加：这段耗时**完全在 UI 线程内同步等待**，且发生在文档解析之前。
- 为什么"感觉特别慢"：对 `xlsx→docx` 来回切换的用户，每切一次类型就吃一次冷启动（约 0.5~0.9s），且没有任何预热/渐进提示。

### P0-2 ｜UI 线程全程同步阻塞（放大器）

- 证据：`ViewerWindow.Actions.cs:356-393`（`DispatcherPriority.Input`，同 UI 线程）+ `Plugin.cs:158` 同步 `PreviewFile` + `PreviewHandlerHost.cs:210` 同步 `DoPreview()`。
- 后果：
  1. `IsBusy` 的转圈动画**不会动**（WPF 渲染线程被阻塞）→ 用户看到"假死"而非"加载中"；
  2. 期间按 Esc/空格**无法取消**，只能等；
  3. 连续浏览多个 Office 文件时，键盘输入被阻塞，j/k 快速切换会明显"粘滞"。

### P1-1 ｜文档解析完全不可观测、也不可复用

- 证据：`DoPreview()` 前后没有任何计时；`PluginManager.FindMatch` 的 `Stopwatch` 只测 `CanHandle` 且只在 Debug 版输出（`PluginManager.cs:63-69`）。
- 后果：**没人知道 xlsx 到底慢在"启动 Excel"还是"解析工作簿"**。这直接导致社区只能猜测（也导致本报告必须用实测把激活项先钉死）。
- 附带：同一文件反复预览（空格开→关→空格开）会**重新解析一遍**，没有任何结果复用。

### P1-2 ｜资源永久占用：3 × Office 进程 ≈ 355 MB，且拒绝回收

- 证据：`PreviewHandlerHost.cs:47` 静态 `ConcurrentDictionary<Guid, IClassFactory> HandlerFactories` + `:248` `LockServer(true)`；**全仓库无 `LockServer(false)`**。
- 后果：
  - 预览过一个 xlsx 就永久多一个 ~120MB 的 Excel；三种类型都碰过 ≈ 355MB 常驻、永不回收、QuickLook 退出前也不释放（`App.OnExit` 不清理该静态字典）。
  - 反噬性能：在 8GB 机器上，这会把内存压到需要换页，**连累后续所有文件预览变慢**——即"资源占用"与"打开速度"是同一个问题。
- 附带缺陷：`TryCreateOutOfProcHandler` 只在 `CreateInstance` 抛 `COMException` 时驱逐失效工厂（`:258-269`）；若 Office 进程被外加因素杀掉且未抛异常，会留下"看起来有效但已死"的缓存项。

### P2-1 ｜无预热、无空闲回收、无同类型复用

- 预热：无。QuickLook 启动后不会为 Office 处理器做任何准备，第一次预览必然是冷的。
- 空闲回收：无（与 P1-2 同源）。
- **同类型处理器复用：无。** 这是最可惜的一处：`IPreviewHandler` 契约本身允许 `Unload()` 后重新 `Initialize(path)` 复用一个实例，但当前实现每次切文件都 `Dispose()` 掉 handler 再 `CreateInstance` 一个新的（含 `Marshal.FinalReleaseComObject` + `GC.Collect()`）。

### P2-2 ｜既有的"暴力 GC"在帮倒忙

- `PreviewHandlerHost.Dispose` 内的显式 `GC.Collect()`（`:95`）发生在 **UI 线程**，且每次关闭预览都执行；
- `ViewerWindow.OnClosing` → `ProcessHelper.PerformAggressiveGC()`（`ViewerWindow.Actions.cs:519`）→ `Task.Delay(2000).ContinueWith(t => GC.Collect(GC.MaxGeneration))`（`ProcessHelper.cs:32-36`）→ **每次关窗 2 秒后一次 Full GC**。
- 这些是有意为之的内存优化，但对 Office 场景（大对象、跨进程句柄）反而制造停顿；至少应改为"内存压力触发 + 后台低优先级"。

### P3 ｜零风险清理项

1. `TryCreateInProcHandler` 对 Office 永假 → 可加"该 CLSID 无 InprocServer32 则跳过"的探测缓存；
2. `IInitializeWithStream/Item` 对 Office 永假 → 首次 QI 结果缓存；
3. `CanHandle` 每次两轮注册表访问 → 进程级字典缓存（扩展名→CLSID），监听 `HKCR` 变更或在校验失败时失效；
4. `IStreamWrapper` 不完整：`CopyTo/Clone/LockRegion/UnlockRegion/Commit/Revert` 全部 `throw NotSupportedException`，`Stat` 只填 `cbSize`（`IStreamWrapper.cs:36-80`）。对 Office 无影响，但**若用户的第三方处理器（WPS/LibreOffice 等）实现了 `IInitializeWithStream` 且调用 `Clone()`，会直接失败**。属于潜在 bug，建议补全或至少降级为"stream 路径失败即回退 file"。

---

## 5. 优化方案

按"收益/风险比"分三层，可独立交付。

### L0 层｜快速收益（1~2 天，低风险，不改架构）

**L0-1 修正 `LockServer` 泄漏 + 引入空闲回收管理器**（收益：内存 ≈355MB → ≤120MB）

- 模块：`QuickLook.Plugin.OfficeViewer/PreviewHandlerHost.cs`
- 改动：把 `static ConcurrentDictionary<Guid, IClassFactory>` 换成 `HandlerLease` 管理器，记录 `LastUsedUtc`；在 `UnloadPreviewHandler` 后调度"空闲 N 分钟（默认 5）则 `LockServer(false)` + 释放并移除缓存项"；进程退出时统一释放。
- 保留能力：同一会话内的热路径（~2ms）不变。

**L0-2 接口能力探测缓存**（收益：小但零风险）

- 模块：同上 + `Plugin.cs`
- 改动：对每个 CLSID 首次 `CreateInstance` 后记录其支持的首选初始化接口（Office = `IInitializeWithFile`），后续直接用；跳过已知无 `InprocServer32` 的 CLSID。

**L0-3 按需预热**（收益：消除首次 482~863 ms）

- 模块：`Plugin.cs` 的 `Init()` 或新增 `OfficePreviewPrewarmer`
- 改动：设置项 `PrewarmHandlers`（默认关闭）与 `LastUsedHandler`（记录上次实际用过的处理器 CLSID）。若开启，在 QuickLook 启动后延迟若干秒、于**后台 STA 线程**对上次用过的 CLSID 做一次 `CoGetClassObject`+`CreateInstance` 即释放，把冷启动挪到用户看不见的时候。
- 权衡：预热会让对应 Office 进程常驻（与 L0-1 的空闲回收配合使用，形成"预热→使用→闲置回收"闭环）。

**L0-4 去掉 UI 线程上的暴力 GC**

- 模块：`PreviewHandlerHost.cs:95`、`ProcessHelper.cs:32-36`
- 改动：删除 `Dispose` 里的 `GC.Collect()`；`PerformAggressiveGC` 改为监听内存压力（或至少 `GC.Collect(2, GCCollectionMode.Optimized, blocking:false, compacting:false)` 并放到线程池）。

**L0-5 `CanHandle` 注册表结果缓存**

- 模块：`ShellExRegister.cs` / `Plugin.cs`
- 改动：进程级 `Dictionary<string, Guid>`，首次查询后缓存；文件扩展名是小集合，收益有限（0.12ms→~0.01ms），但能减少 HKCR 访问与失败异常。

### L1 层｜结构性改进（约 1 周，中风险，收益最大）

**L1-1 同类型处理器复用（推荐优先做）**

- 模块：`PreviewHandlerHost.cs`、`PreviewPanel.cs`、`Plugin.cs`
- 做法：当"上一次预览的 CLSID"与"本次相同"且 handler 实例仍存活时，**不销毁**，改为 `Unload()` → `Initialize(newPath)` → `SetWindow/DoPreview`；仅在 CLSID 变化或窗口关闭时释放。
- 预期：连续浏览同目录/同类型文件时，省掉实例化与 Office 侧处理器重入成本。**这正是"在文件夹里按 j/k 挨个看 xlsx"的主路径。**
- 注意：`Marshal.FinalReleaseComObject` + `GC.Collect` 必须从该路径上移除，否则复用无意义。

**L1-2 首屏结果缓存（位图级）**

- 模块：新增 `OfficePreviewCache`（静态/进程级——**必须静态**，因为插件实例每次预览都会重建，见 2.2）
- 键：`(规范化路径, 文件大小, LastWriteTimeUtc, 处理器CLSID, DPI档位, 窗口尺寸档)`；建议尺寸按"档位"归一化（如 0.75x/1x/1.5x），否则命中率会被任意改窗尺寸毁掉。
- 值：首屏位图（`PrintWindow`/`BitBlt` 从 handler 窗口抓取，或用 `Dwm` 缩略图）；可选磁盘 LRU（`%LocalAppData%\QuickLook\Cache\OfficePreview\`）。
- 行为：命中即先贴图（≈0ms），随后后台刷新；`AutoReload` 触发的文件变更必须失效对应键。
- 预期：**同一文件重复预览（开→关→开，极高频）从"激活+解析"降到接近 0。**
- 风险：磁盘占用、DPI/主题变化导致画面失真 → 缓存项需带 DPI/主题指纹。

**L1-3 打开/关闭动画与"假死"修复**

- 模块：`ViewerWindow.Actions.cs`、`Plugin.cs`
- 最小改动版：把 COM 激活（`CoGetClassObject`+`CreateInstance`，即 482~863ms 那段）从 `Open()` 中拆出，改由 L0-3 预热或后台线程提前完成；UI 线程只保留 `SetWindow`+`DoPreview`。这样**UI 冻结时间从"激活+解析"降为"解析"**，且不触及跨线程 HWND 的复杂性。
- 完整版（可选，L2）：专用 STA 线程 + 独立消息泵承载 handler，通过子 HWND 挂进 WPF（airspace 用 `WindowsFormsHost` 规避）→ UI 完全不冻结，Esc 可取消。

### L2 层｜架构升级（视产品定位，2~4 周）

**L2-1 可插拔渲染后端**

- 抽象 `IOfficePreviewBackend`，提供三种实现：
  1. `ShellPreviewHandlerBackend`（现状，保真度最高，依赖 Office）
  2. `ManagedOoxmlBackend`：用 `DocumentFormat.OpenXml`（MIT 许可）解析 xlsx/docx 并做轻量渲染（表格/段落/基础样式），**毫秒级、无需 Office、无需 COM**。适合"要看内容而不是要看版式"的高频场景——用户抱怨最多的 Excel 恰好最适合。
  3. `ThirdPartyHandlerBackend`：若机器上注册了 WPS/LibreOffice 等更轻的处理器，优先用（需实测其接口能力与冷启动耗时——按本报告方法量一遍即可）。
- 决策策略：设置项 `OfficePreviewBackend = Auto|Shell|Managed`；`Auto` 下按"该文件类型 + 文件大小"选择（例如 xlsx 小文件走 Managed，复杂版式/大文件走 Shell），并在预览窗口提供"精确渲染"切换按钮。
- 明确不推荐：`QL-Win/QuickLook.Plugin.OfficeViewer` 那条 Syncfusion 路线——**Syncfusion 组件非免费**，且其授权明确禁止把组件用于 GPL 项目（QuickLook 本体是 GPL-3.0），需要在 `SyncfusionKey.cs` 注入商业/开源授权密钥。不适合作为免费分支的默认方案。

**L2-2 预览宿主服务化**

- 把 Office 预览宿主拆成独立 helper 进程（思路同 `prevhost.exe` / PowerToys Peek）：主进程只做 UI，崩溃隔离、可常驻复用、可独立限流。这是长期最干净的形态，但改动面最大。

---

## 6. 优化目标、模块范围与验证方式

### 6.1 量化目标（KPI）

定义 **T_firstpaint** = 从"空格按下"到"预览窗口内出现可辨认内容"的墙钟时间。分项埋点：`T_lookup` / `T_activate` / `T_init` / `T_dopreview` / `T_cachehit`。

样本：同一套文件在改动前后各测 10 次，取 P50/P95（首次预览单列）。

| # | 场景（指标） | 现状（实测/推断） | 目标 | 主要手段 |
|---|---|---|---|---|
| K1 | 冷启动首次预览 `.xlsx`（T_activate 分项） | **839.0 ms** | ≤1200 ms 完成预热，用户侧首帧不再包含激活项 | L0-3 |
| K2 | 会话内首次预览 `.xlsx`（未预热） | ≥839 ms + 解析 | ≤1000 ms（P50） | L0-3 / L1-1 |
| K3 | 同类型第二文件（热路径，T_activate） | 1.8 ~ 2.5 ms | 保持 ≤5 ms（不劣化） | L0-1 不得破坏热路径 |
| K4 | 同一文件重复预览 | ≈ 激活 + 解析（约 0.9~3s） | **≤50 ms**（P50，缓存命中） | L1-2 |
| K5 | 连续浏览 10 个同类型文件：平均每文件 T_firstpaint | 待测（基线未知） | 相比基线 **−50%** 以上 | L1-1 + L1-2 |
| K6 | UI 冻结时长（单次预览内主线程被占时长） | = 激活 + 解析全程 | ≤100 ms（L1-3 最小版）；0（L2 版） | L1-3 |
| K7 | 常驻内存（QuickLook 之外） | **≈355 MB（永久）** | ≤120 MB，且闲置 5 min 后回收 | L0-1 |
| K8 | 20 次预览后的 GDI/User 句柄增量 | 待测（跨进程窗口泄漏风险） | ≤基线（无单调增长） | 回归监控 |
| K9 | Debug 版 `FindMatch` 总耗时（插件匹配链） | 待测 | 不劣化 | 回归监控 |

### 6.2 模块范围（改动清单）

| 层 | 文件 | 改动性质 |
|---|---|---|
| L0 | `QuickLook.Plugin/QuickLook.Plugin.OfficeViewer/PreviewHandlerHost.cs` | 保活管理器、探测缓存、去 GC.Collect |
| L0 | `…/OfficeViewer/Plugin.cs` | 预热开关、`LastUsedHandler`、注册表缓存 |
| L0 | `…/OfficeViewer/ShellExRegister.cs` | 缓存 + 失效 |
| L0 | `QuickLook.Common/Helpers/ProcessHelper.cs` | 暴击 GC 改异步/按需 |
| L0 | `QuickLook/ViewerWindow.Actions.cs`（`:519`） | 调用点调整 |
| L1 | `…/OfficeViewer/PreviewPanel.cs` + `PreviewHandlerHost.cs` | 处理器复用、激活与渲染拆分 |
| L1 | 新增 `…/OfficeViewer/OfficePreviewCache.cs` | 首屏位图缓存（**必须静态**） |
| L2 | 新增 `IOfficePreviewBackend` + 两个实现 | 后端抽象与 Managed 渲染 |
| 回归 | `QuickLook/PluginManager.cs` | 埋点与匹配链耗时统计（可选） |

### 6.3 验证方式

**(a) 分项埋点（改造前必须先做，否则无法归因）**

在 `Plugin.View` 与 `PreviewHandlerHost.Open` 内插 `Stopwatch`，通过 `ProcessHelper.WriteLog` 或 `EventSource`（ETW）输出结构化记录：

```
ts, path, ext, clsid, T_lookup, T_activate, T_init, T_dopreview, T_firstpaint, cache_hit, ws_delta_mb
```

注意：现有 `PluginManager.FindMatch` 的 `Debug.WriteLine` 只在 Debug 版存在，Release 版**无任何计时**——这是当前无法定位问题根因的直接原因。

**(b) 本次交付的回归脚本（可直接对比前后）**

```
python tools/probe_office_preview.py        # 注册链是否变化（换了处理器/被别的软件劫持）
python tools/bench_com_activation.py        # 冷/热激活耗时 + 是否仍拉起主程序
python tools/probe_warm_server_cost.py      # 保活内存 + 注册表开销
```

门槛建议：`T_activate` 或常驻内存相对基线劣化 >10% 即视为回归。

**(c) 端到端自动基准（推荐补建）**

QuickLook 支持命令行传入路径（`App.OnStartup` 会把首个参数经管道投递为 `Toggle`），因此可脚本化：

1. `QuickLook.exe "<样本.xlsx>"` 触发预览；
2. 用 UI Automation / 窗口枚举 + 定时截屏，检测"预览窗口内出现非空白内容"的时刻；
3. 记录 T_firstpaint；`Esc`/关窗后再测下一轮。

样本集必须覆盖：空表 / 10 行 / 5 千行含公式 / 5 万行多工作表 / 含图表与条件格式；同构的 docx、pptx 各一套。**大文件样本是区分"激活慢"与"解析慢"的唯一手段。**

**(d) 外部基线对照（防止把 Office 固有的慢算到 QuickLook 头上）**

用**资源管理器自带的预览窗格**打开同一批样本，记录同样的 T_firstpaint。二者的处理器完全相同，差值才是 QuickLook 自身引入的开销（激活策略、窗口管线、GC 等）。这条对照能避免优化方向跑偏。

**(e) 资源验证**

采样 `Working Set` / `Private Bytes` / `Handle Count` / `GDI Objects` / `User Objects`（含 `excel.exe` 等被拉起的进程），在"预览 20 次 + 闲置 10 分钟"后确认：L0-1 生效（进程退出/内存回落）、无句柄单调增长。

### 6.4 风险与回滚

| 改动 | 风险 | 回滚/缓解 |
|---|---|---|
| L0-1 空闲回收 | 释放后立刻再预览又要冷启动（体验回退） | 空闲阈值可配置（默认 5 min）；回收前判断窗口未使用 |
| L0-3 预热 | 用户不希望后台常驻 Office | 默认**关闭**；设置项显式开启 |
| L1-1 处理器复用 | 复用状态污染（残留上一文档的视图/缩放） | 只在同 CLSID + 同扩展名时复用；失败即回退到"重建"路径 |
| L1-2 位图缓存 | 画面失真、磁盘占用、`AutoReload` 显示陈旧内容 | 缓存键含文件 mtime/size；主题/DPI 指纹变更即失效；提供"清空缓存"入口 |
| L1-3 激活移出 UI 线程 | 跨线程 COM/窗口时序问题 | 最小版只把 `CoGetClassObject`+`CreateInstance` 挪走，`SetWindow/DoPreview` 仍留 UI 线程 |
| L2-1 Managed 后端 | 保真度下降（图表/版式） | 默认 `Auto` 且仅对 xlsx 小文件启用；窗口内提供"精确渲染"按钮；后端可整体关闭 |

---

## 7. 附录：本机取证原始输出

### 7.1 注册链（`probe_office_preview.py`）

```
.xlsx  handler CLSID : {00020827-0000-0000-C000-000000000046}   (Microsoft Excel Previewer)
       declared at   : HKCR\.xlsx\shellex\{8895b1c6-b41f-4c1c-a562-0d564250836f}
       CLSID name    : Microsoft Excel previewer
       InprocServer32: None
       LocalServer32 : C:\Program Files\Microsoft Office\Root\Office16\EXCEL.EXE
       AppID         : None
       CLSID subkeys : ['InprocHandler32', 'LocalServer32']

.docx  handler CLSID : {84F66100-FF7C-4fb4-B0C0-02CD7FB668FE}   (Microsoft Word Previewer)
       LocalServer32 : C:\Program Files\Microsoft Office\Root\Office16\WINWORD.EXE
       AppID         : None

.pptx  handler CLSID : {65235197-874B-4A07-BDC5-E65EA825B718}   (Microsoft PowerPoint Previewer)
       LocalServer32 : C:\Program Files\Microsoft Office\Root\Office16\POWERPNT.EXE
       AppID         : None

.vsdx  handler CLSID : {21E17C2F-AD3A-4b89-841F-09CFE02D16B7}   (Microsoft Visio Previewer)
       LocalServer32 : C:\Program Files\Microsoft Office\Root\Office16\VPREVIEW.EXE
       AppID         : {21E17C2F-AD3A-4b89-841F-09CFE02D16B7}

.pptm  no preview handler registered  -> OfficeViewer plugin CANNOT handle it
```

### 7.2 COM 冷/热激活（`bench_com_activation.py`）

```
### PHASE 1 - COLD (no Office process running yet)
  Excel (.xlsx)    total=   839.0 ms  getClassObject=819.8 createInstance=18.7  started=['excel.exe']
  Word  (.docx)    total=   863.5 ms  getClassObject=857.0 createInstance=6.4   started=['winword.exe']
  PPT   (.pptx)    total=   482.3 ms  getClassObject=479.1 createInstance=3.2   started=['powerpnt.exe']

### PHASE 2 - WARM (server process already alive)
  Excel (.xlsx)    total=     1.8 ms  getClassObject=1.3 createInstance=0.5  started=[]
  Word  (.docx)    total=     2.3 ms  getClassObject=1.5 createInstance=0.8  started=[]
  PPT   (.pptx)    total=     2.5 ms  getClassObject=1.4 createInstance=1.0  started=[]
```

### 7.3 接口能力（`probe_handler_interfaces.py`）

```
Excel (.xlsx)
   CLSCTX_INPROC_SERVER usable : False (0x80040154)
    no  IInitializeWithStream
    no  IInitializeWithItem
   YES  IInitializeWithFile
   YES  IPreviewHandler
   YES  IObjectWithSite
    no  IPersistStream
```
（Word / PowerPoint 结果完全一致）

### 7.4 保活代价与注册表开销（`probe_warm_server_cost.py`）

```
registry_lookup_ms_per_call: 0.1182

excel.exe    working_set 122.0 MB   private 61.6 MB   (pinned, 10s)
winword.exe  working_set 139.3 MB   private 68.2 MB
powerpnt.exe working_set 114.9 MB   private 54.8 MB
合计 working set ≈ 355 MB         释放 LockServer 后进程全部退出（after_release_alive: []）
```

---

## 8. 一句话总结

QuickLook 的 Office 预览慢，**不是因为它解析得慢，而是因为它把"启动一个完整的 Excel/Word/PowerPoint 进程"这件事放进了用户按下空格后的同步路径里**：冷启动实测 482~863 ms 且完全无法回避（进程内激活已被实测证明不可用），期间 UI 线程同步阻塞、没有任何预热，唯一的"缓存"是一个漏掉 `LockServer(false)` 的永久保活（≈355 MB 常驻），而唯一能复用的机会——同一处理器实例的重新初始化——反而被每次 `Dispose()`+`GC.Collect()` 主动放弃。优化应集中在**"把激活挪出用户可见路径（预热/后台）+ 让结果可复用（位图缓存/处理器复用）+ 让保活可回收（空闲超时）"**这三件事上，而不是去动插件匹配或注册表查询。
