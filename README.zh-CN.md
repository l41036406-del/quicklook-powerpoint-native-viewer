# QuickLook PowerPoint Native Viewer

[English](README.md)

这是一个实验性的 QuickLook 插件，用于通过本机 Microsoft PowerPoint 渲染引擎预览 PowerPoint 文件，而不是使用 Syncfusion。

这个原型的目标是验证一种更稳的 PPT/PPTX 预览路径，尤其适用于 `QuickLook.Plugin.OfficeViewer` 预览效果失真或直接报错的文件。

## 当前功能

- 支持 PowerPoint 格式：`.ppt`、`.pptx`、`.pptm`、`.pot`、`.potx`、`.potm`、`.pps`、`.ppsx`、`.ppsm`
- 使用 Microsoft PowerPoint COM 自动化，以只读方式打开源演示文稿
- 将演示文稿导出为临时 PDF
- 通过 WebView2 显示 PDF，使用系统已安装的 Microsoft Edge WebView2 Runtime
- 通过内嵌 WebView2 PDF 查看器支持 PDF 文字选择和复制
- 当源文件路径、大小、修改时间不变时，复用持久 PDF 缓存
- 预览控件卸载时删除 WebView2 会话数据
- 预览窗口打开时会尝试主动置前
- 不修改原始 PowerPoint 文件
- 可与原版 OfficeViewer 插件并排安装
- 使用 `Priority = 100`，让本插件优先接管 PowerPoint 文件，Word/Excel 仍可由其他插件处理

## 当前限制

- 需要 Windows 上已安装并完成 COM 注册的 Microsoft PowerPoint
- 首次预览某个文件时可能慢于 Syncfusion，因为需要启动 PowerPoint 并生成 PDF
- 同一个文件未变化时，再次预览会优先使用持久 PDF 缓存，速度应更快
- 当前实现仍是沙箱原型，不是正式稳定版
- 需要 Microsoft Edge WebView2 Runtime；现代 Windows 通常已经自带

## 启动加载页面

插件启动 PowerPoint 时，QuickLook 可能会显示下面这个黑色加载页：

![Starting PowerPoint loading screen](docs/images/starting-powerpoint-stuck.png)

这个页面是某个文件首次启动阶段的预期现象，表示插件正在启动 PowerPoint 并生成 PDF 预览。如果已有缓存，这个页面应明显缩短或不明显出现。

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

原型阶段。当前仓库用于记录 PowerPoint 原生 PDF/WebView2 预览实验，以及测试过程中观察到的已知行为。
