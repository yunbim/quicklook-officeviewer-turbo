# QuickLook Office 极速预览版（Turbo）

[English](README.en.md) | **中文**

让 [QuickLook](https://github.com/QL-Win/QuickLook) **连续预览 Office 文件时从「每按一次空格卡半秒」变成「秒开」**，
并让**下载来的受保护视图文档直接出图**（不弹确认框、不改动你的原文件）。

> 这是 [QL-Win/QuickLook](https://github.com/QL-Win/QuickLook) 官方 **4.5.0** 的**插件级补丁（非官方）**。
> 只替换**一个文件**，主程序 `QuickLook.exe` 与其余 24 个插件保持官方原样，可与官方版混用。
>
> **搜索关键词**：QuickLook Office 预览慢 · Protected View 弹窗 · Zone.Identifier · preview handler 预热 · Office preview slow · faster Word Excel PowerPoint preview

---

## ⚡ 30 秒安装（推荐）

> 前置：先装官方的 [QuickLook 4.5.0](https://github.com/QL-Win/QuickLook/releases)。
> 本包**只含插件**，不含主程序。

1. 到本仓库的 **[Releases 页面](https://github.com/yunbim/quicklook-officeviewer-turbo/releases)** 下载
   **`...-plugin-only.zip`**（约 30 KB，只含插件与安装脚本）。
   > 也可以直接下载那个**光秃秃的 `QuickLook.Plugin.OfficeViewer.dll`** ——
   > 全部功能都在这一个文件里，不需要任何安装脚本。
2. 解压到任意文件夹，**双击 `install.bat`**。它自动完成：定位 QuickLook → 关闭它 →
   备份你当前的 DLL → 替换 → **逐字节校验** → 重新启动。
3. 完事。对着 Office 文件按空格即可。

**不需要装任何额外东西**：用的是 Windows 自带的 PowerShell（Win7 起就有），不需要 Python。
也不需要管理员权限（除非 QuickLook 装在 `Program Files` 下）。
QuickLook 位置不常规时脚本会提示你粘贴路径，也可直接 `install.bat "D:\Apps\QuickLook"`。

**还原官方原版**：双击 `rollback.bat`。

<details>
<summary>不想跑脚本？手动替换（就一步）</summary>

1. 托盘图标右键 → **Quit**，退出 QuickLook。
2. 把 `plugin\QuickLook.Plugin.OfficeViewer.dll` 覆盖到：

   ```
   %LOCALAPPDATA%\Programs\QuickLook\QuickLook.Plugin\QuickLook.Plugin.OfficeViewer\
   ```

3. 重新打开 QuickLook。

就这一个文件 —— 官方包里的 `UnblockZoneIdentifier.dll` 原装已有，**不要**覆盖它。
</details>

---

## 它解决什么

### A. 连续预览 Office 文件，每次都要等半秒

QuickLook 的 Office 插件**自己不解析文档**，它把文件转交给 **Windows Shell 预览处理器**。
而 Office 注册的处理器，其 `LocalServer32` **直接指向 Office 主程序本体**
（`WINWORD.EXE` / `EXCEL.EXE` / `POWERPNT.EXE`），**没有 `DllSurrogate`**。

后果：**每预览一个 Office 文件 = 启动一次 Office 进程**，实测 **482 ~ 863 ms**。
只看一个文件还能忍；但真实用法是在装满文档的文件夹里按方向键连续翻看 ——
于是每次翻页都卡半秒，翻 20 个文件就是十几秒的纯等待。

### B. Office 进程被永久钉在内存里

原版调用 COM 类工厂的 `LockServer(true)` 后**从不配对释放**，导致 Word/Excel/PPT
**永久常驻约 355 MB**，直到注销或手动结束进程。

### C. 每关一次预览窗，界面卡一下

`PreviewHandlerHost.Dispose()` 在 **WPF UI 线程**上跑了一次阻塞式 `GC.Collect()`。

### D. 下载来的文档必须点弹窗才能看，而且会改你的文件

浏览器 / 邮件 / 网盘下载的文档带 `Zone.Identifier` 标记（"来自 Internet"），
**Office 处理器拒绝加载这类文件**。原插件因此弹一个 Yes/No 确认框，
而它是靠**永久删除原文件的 `Zone.Identifier`** 来实现的 —— **你的原文件被改动了**。

### 效果对照

| 场景 | 原版 | 极速版 |
|---|---|---|
| 第一次预览某类 Office 文件（冷启动，架构决定） | 482 ~ 863 ms | 482 ~ 863 ms（**相同**） |
| **之后每一次预览同类文件** | **每次都是 482 ~ 863 ms** | **≈ 2 ms** |
| 连续翻看 20 个 Excel | ≈ 10 ~ 17 s 纯等待 | 首个约 0.5 ~ 0.9 s，**其余接近瞬开** |
| Excel / Word / PPT 混着翻 | 每换一种再冷启动 | 首次后**后台把其余类型也预热好** |
| 预览关闭后内存 | 常驻 ≈ 355 MB | 空闲超时自动归还，**零残留**（时长可配） |
| 关闭预览窗 | 阻塞 GC，界面卡顿 | 已移除 |
| 下载来的文档 | 弹框 + 改你原文件 | 直接出图 + 提示条 + 按需解保护 |

> **为什么第一次还是慢？** 因为"预览 Excel"在架构上就等于"启动 Excel 进程"，
> 插件消除不了进程启动时间。极速版做的是把这次成本**从第 2 次开始彻底付掉**，
> 并额外提供「连第一次也预热」的开关。

---

## 使用方式

**没有需要你记的新操作。** 照常按空格预览：

- **连续翻看**：第一次仍是冷的（架构决定）；用过一个之后后台就把其余 Office 处理器
  也热起来了，之后连续预览接近瞬开。
- **想让"连第一次也秒开"**：把 `WarmUpAtStartup` 改成 `True`
  （代价：常驻约 355 MB，空闲超时后自动回收）。
- **下载来的文档**：直接就能看，顶部琥珀色提示条说明这是副本、原文件未被修改。
  想让以后用 Word 打开它也不进保护视图，就点提示条右边的 **「解除源文件保护视图」**。
- **普通文档**：体验与官方版完全相同 —— 不复制、不显示提示条。

> 改了配置要**重启 QuickLook** 才生效（配置被缓存在内存里）。
> 托盘图标只在**运行中**出现，Windows 11 默认收进时钟旁的 `^` 溢出区。

---

## 实现原理

四项性能改动 + 一项受保护视图改动，**全部封装在 OfficeViewer 插件内部**，
不改任何 QuickLook 核心文件，所以 `QuickLook.exe` 无需重编。

### 性能：`OfficePreviewHostPool.cs`（新增）

核心是 COM 的 **`IClassFactory` 保活**：`LockServer(true)` 后类工厂带着它的进程一起活着，
只要工厂还在，下一次 `CoCreateInstance` 就是**同进程内的热激活（≈2 ms）**，
而不是重新启动 Office。**"秒开"就是这么来的。**

原版只保活、从不释放。现在加了**空闲看门狗**：无预览超过 `HandlerIdleTimeoutSeconds`
（默认 1800 秒）后 `LockServer(false)` 并释放工厂引用，Office 进程随之自行退出。

> **关键约束**：COM 类工厂有**套间亲和性（STA）**，必须在创建它的那个 STA 上释放。
> 看门狗线程只负责"判定空闲"，真正的释放通过 `Dispatcher` **回到 WPF UI 线程**执行。

其余三项：

| 改动 | 文件 | 内容 |
|---|---|---|
| 去掉 UI 线程阻塞 GC | `PreviewHandlerHost.cs` | 删除 `Dispose()` 里的 `GC.Collect()` |
| 跳过注定失败的接口探测 | `PreviewHandlerHost.cs` | 实测三个 Office 处理器**只实现 `IInitializeWithFile`**，前两级必失败。按 CLSID 记住成功过的初始化接口；非 Office 处理器仍保留原回退顺序 |
| 两级后台预热 | `Plugin.cs` | 预热在 `Dispatcher` 的 `Background` 优先级执行，窗口已显示、消息队列空闲时才做，用户看不到 |

### 受保护视图：四个文件

| 文件 | 职责 | 关键点 |
|---|---|---|
| `ProtectedViewPreview.cs` | 副本生命周期 | `ZoneIdentifierManager.IsZoneBlocked()` 判定 → 复制到 `%TEMP%\QuickLook\OfficeViewer\`（文件名含哈希避免碰撞）→ **只对副本**写空 `Zone.Identifier` → 渲染 → 关闭时删除；启动时清扫陈旧副本 |
| `ProtectedViewBanner.cs` | 提示条 + 按钮 | 自绘琥珀色 `Border`（**不依赖 QuickLook 主题**，浅/深色下都清晰）；按钮点击后原地解保护并把整条变绿 |
| `UiText.cs` | 中英文两版文案 | 本补丁自己绘制的文字全部集中在这里。默认跟随 **Windows 界面语言**（中文系统中文、其他系统英文），可用 `UiLanguage` 强制 |
| `Plugin.cs` | 分支决策 | 带标识 → 副本路径 + 提示条；不带标识 → 原路径，与官方版完全一致 |

**确认弹窗是代码级删除的** —— 不是"帮你把开关默认打开"，`Plugin.cs` 里已经没有那段
`MessageBox` 逻辑了。

### 一个如实的取舍

提示条会让**受保护文档**的可视区矮约 **43 px**（占独立一行）。原因是宿主用
`WindowsFormsHost` 承载 Office 视图，它会盖住与之重叠的 WPF 元素，
所以提示条**做不成悬浮层**。**普通文档不受影响。**

---

## 可配置项

设置域 `QuickLook.Plugin.OfficeViewer`，文件：
`%LOCALAPPDATA%\Programs\QuickLook\UserData\QuickLook.Plugin.OfficeViewer.config`

| 键 | 默认 | 说明 |
|---|---|---|
| `WarmUpAfterFirstPreview` | `True` | 首次预览后后台预热其余同类 Office 处理器 |
| `WarmUpAtStartup` | `False` | 启动即预热全部 Office 处理器（代价：常驻内存） |
| `HandlerIdleTimeoutSeconds` | `1800` | 空闲多久后释放 Office 进程；`0` = 永不释放（等同原版行为） |
| `LogPreviewTiming` | `False` | 把每次预览耗时写入 `UserData\QuickLook.Exception.log` |
| `UiLanguage` | `auto` | 受保护视图提示条的语言：`auto` 跟随 Windows 界面语言 / `zh` 始终中文 / `en` 始终英文 |

写法（**合并**进已有 `<Settings>`，不要整段覆盖；仓库 `config/` 下有模板）：

```xml
<?xml version="1.0" encoding="utf-8"?>
<Settings>
  <WarmUpAfterFirstPreview>True</WarmUpAfterFirstPreview>
  <WarmUpAtStartup>False</WarmUpAtStartup>
  <HandlerIdleTimeoutSeconds>1800</HandlerIdleTimeoutSeconds>
  <LogPreviewTiming>False</LogPreviewTiming>
  <UiLanguage>auto</UiLanguage>
</Settings>
```

---

## 怎么自己验证它生效了

`dev/tools/` 下的脚本不是"看一眼能不能显示"，而是**读真实像素 / 读真实进程**
（需要 Python 3，只用标准库）：

```bash
python dev/tools/verify_office_warmup.py --warmup on --idle 60 --watch 190  # 冷/热耗时 + 空闲后进程是否真的退出
python dev/tools/verify_protected_copy.py                                   # 副本预览三场景（含真实鼠标点击）
```

后者的判据包含 **"原文件的区域标识是否仍然存在"** —— 只测"能预览"是测不出有没有动到用户文件的。

---

## 回滚

| 目的 | 做法 |
|---|---|
| 撤销本补丁 | 双击 `rollback.bat`（或手动用 `DLL.official-backup` 覆盖） |
| 回官方原版 | 用 `plugin-stock\QuickLook.Plugin.OfficeViewer.dll` 覆盖 |
| 关闭预热 | `WarmUpAtStartup=False`、`WarmUpAfterFirstPreview=False` |
| 恢复原版内存行为 | `HandlerIdleTimeoutSeconds=0` |
| 完全卸载 | 退出 QuickLook，删除 `%LOCALAPPDATA%\Programs\QuickLook` |

---

## 边界与未做项（如实说明）

1. **首次冷预览仍有约 840 ms**，除非开 `WarmUpAtStartup`。架构决定，无法在插件内消除。
2. **没有改 QuickLook 核心，也没有做 Word 纸张分页视图**。Office 预览处理器是
   **连续流式渲染、不画页边界**的，此链路上做不到分页；要做得走"转 PDF + PdfiumViewer"，
   属独立工程，本包未做。
3. **副本预览的可见代价**：受保护文档预览区矮约 43 px（见上）。普通文档不受影响。
4. **临时副本明文落在 `%TEMP%`**，预览结束即清理；若预览期间被强杀可能残留，
   下次启动会清扫陈旧副本。
5. **原生库 `QuickLook.Native*.dll` 来自官方发布包**，非本地编译；本包不含，沿用官方安装。

---

## 目录结构

```
QuickLook-OfficeViewer-Turbo/
├── README.md            本文（中文）
├── README.en.md         English README
├── INSTALL.txt          纯安装说明（精简包里也有）
├── CHANGELOG.md         版本变更
├── LICENSE-GPL.txt      上游 GPL-3.0 全文（本改动同样以此发布）
├── MANIFEST.txt         全部文件 sha256 校验清单
├── install.bat          ★ 双击安装（零依赖）—— 入口，逻辑在 install.ps1
├── install.ps1            安装逻辑（用 Windows 自带的 PowerShell）
├── rollback.bat         ★ 双击还原官方原版 —— 入口，逻辑在 rollback.ps1
├── plugin/              ★ 唯一需要部署的文件
├── plugin-stock/        官方 4.5.0 原版 DLL（用于精确回滚）
├── config/              默认配置模板
└── dev/                 开发与取证材料，普通用户可忽略
    ├── src/             改动后的源码（保留仓库内相对路径）
    ├── patch/           相对上游的 git 补丁 + 基准提交
    ├── docs/            深度分析、实现与部署、验证报告
    └── tools/           取证 / 基准 / 验证脚本
```

---

## 校验

```bash
sha256sum -c MANIFEST.txt          # Linux / Git Bash / macOS
certutil -hashfile <file> SHA256   # Windows 单文件
```

`patch/` 里的补丁在**纯净的上游源码树**上验证过：打上之后产出的 8 个文件与 `dev/src/`
内容一致（仅换行符不同），并通过**反向应用校验**（`git apply -R --check` 干净通过），
证明它恰好等于"相对上游的完整改动集"，没有夹带无关内容。

补丁还通过了**独立可重建验证**：在 `git archive` 出来的**纯净上游源码树**上 `git apply`
之后，`dotnet build` **直接编译通过**，产物与包内 DLL 同为 34816 字节
（字节级哈希不同属正常 —— .NET 程序集默认嵌入时间戳/MVID，非确定性构建）。

---

## 出处与许可

- **原项目**：[QL-Win/QuickLook](https://github.com/QL-Win/QuickLook) —— Windows 通用文件快速预览工具（按空格预览），**GPL-3.0**。
- 本包是**非官方插件级补丁**，与上游无隶属关系，上游未背书本改动；如需原版请从上游发布页获取。
- 全部改动同样以 **GPL-3.0** 发布（见 `LICENSE-GPL.txt`）。`plugin/` 中的 DLL 由 `dev/src/` 编译而来。
- 再分发请一并提供源码并保留本许可声明。
