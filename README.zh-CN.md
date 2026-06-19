# QuickLook PowerPoint Native Viewer

[English](README.md)

这是一个实验性的 QuickLook 插件，用于通过 Windows Shell Preview Handler 或本机 Microsoft PowerPoint 渲染引擎预览 PowerPoint 文件，而不是使用 Syncfusion。

这个原型的目标是验证一种更稳的 PPT/PPTX 预览路径，尤其适用于 `QuickLook.Plugin.OfficeViewer` 预览效果失真或直接报错的文件。

## 当前功能

- 支持 PowerPoint 格式：`.ppt`、`.pptx`、`.pptm`、`.pot`、`.potx`、`.potm`、`.pps`、`.ppsx`、`.ppsm`
- 优先尝试系统已注册的 Windows Shell Preview Handler，整体思路借鉴 PowerToys Peek
- 当系统预览处理器不可用或初始化失败时，回退到 Microsoft PowerPoint COM 自动化
- 在回退路径中将演示文稿导出为临时 PDF
- 在回退路径中通过 WebView2 显示 PDF，使用系统已安装的 Microsoft Edge WebView2 Runtime
- 当当前后端支持可选择文字时，可以复制文字，例如 Shell Preview Handler 或 WebView2 PDF 查看器
- 当源文件路径、大小、修改时间不变时，复用持久 PDF 缓存
- 预览控件卸载时删除 WebView2 会话数据
- 预览窗口打开时会尝试主动置前
- 不修改原始 PowerPoint 文件
- 可与原版 OfficeViewer 插件并排安装
- 使用 `Priority = 100`，让本插件优先接管 PowerPoint 文件，Word/Excel 仍可由其他插件处理

## 当前限制

- Shell Preview Handler 的效果取决于本机 Windows 注册的预览处理器
- PDF 回退路径需要 Windows 上已安装并完成 COM 注册的 Microsoft PowerPoint
- 某个文件首次进入 PDF 回退路径时可能慢于 Syncfusion，因为需要启动 PowerPoint 并生成 PDF
- 同一个文件未变化时，再次进入 PDF 回退路径会优先使用持久 PDF 缓存，速度应更快
- 当前实现仍是沙箱原型，不是正式稳定版
- 需要 Microsoft Edge WebView2 Runtime；现代 Windows 通常已经自带

## 启动加载页面

当插件可以使用 Shell Preview Handler 时，应尽量避免出现 PowerPoint 启动页。如果插件回退到启动 PowerPoint，QuickLook 可能会显示下面这个黑色加载页：

![Starting PowerPoint loading screen](docs/images/starting-powerpoint-stuck.png)

这个页面只应在首次进入 PDF 回退路径时出现，表示插件正在启动 PowerPoint 并生成 PDF 预览。如果 Shell Preview Handler 可用，或已有 PDF 缓存，这个页面应明显缩短或不明显出现。

## 安装

下载或构建：

```text
dist/QuickLook.Plugin.PowerPointNativeViewer.qlplugin
```

可以双击 `.qlplugin` 让 QuickLook 安装，也可以手动解压到 QuickLook 用户插件目录：

```text
%LOCALAPPDATA%\Packages\21090PaddyXu.QuickLook_egxr34yet59cg\LocalCache\Roaming\pooi.moe\QuickLook\QuickLook.Plugin\QuickLook.Plugin.PowerPointNativeViewer
```

安装后需要重启 QuickLook。

## 回滚

删除下面这个目录，然后重启 QuickLook：

```text
%LOCALAPPDATA%\Packages\21090PaddyXu.QuickLook_egxr34yet59cg\LocalCache\Roaming\pooi.moe\QuickLook\QuickLook.Plugin\QuickLook.Plugin.PowerPointNativeViewer
```

这个原型不会修改原版 OfficeViewer 插件。

持久 PDF 缓存位于：

```text
%LOCALAPPDATA%\QuickLook.PowerPointNativeViewer\pdf-cache
```

## 构建说明

项目目标框架为 .NET Framework 4.6.2，并引用 QuickLook 的 `QuickLook.Common.dll` 和 Microsoft WebView2 SDK 程序集。

本地构建前，请先把 `QuickLook.Common.dll` 复制到：

```text
lib/QuickLook.Common.dll
```

还需要把 WebView2 SDK 程序集放到：

```text
lib/webview2/Microsoft.Web.WebView2.Core.dll
lib/webview2/Microsoft.Web.WebView2.Wpf.dll
lib/webview2/runtimes/win-x64/native/WebView2Loader.dll
```

然后运行：

```powershell
MSBuild.exe QuickLook.Plugin.PowerPointNativeViewer.sln /p:Configuration=Release
powershell -ExecutionPolicy Bypass -File scripts/pack-zip.ps1
```

## 状态

原型阶段。当前仓库用于记录 Shell Preview Handler 优先、PowerPoint PDF/WebView2 回退的预览实验，以及测试过程中观察到的已知行为。
