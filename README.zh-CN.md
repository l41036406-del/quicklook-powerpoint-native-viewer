# QuickLook PowerPoint Native Viewer

[English](README.md)

一个 QuickLook 插件，通过本机安装的 Microsoft PowerPoint 引擎，把幻灯片渲染成图片来预览 PowerPoint 文件。

它是 `QuickLook.Plugin.OfficeViewer`（Syncfusion）的高保真替代方案，适用于在原插件里预览失真或直接报错的 PPT/PPTX 文件——同时**保持 QuickLook 正常的上/下键文件切换可用**。

## 当前版本能力

- 支持 PowerPoint 格式：`.ppt`、`.pptx`、`.pptm`、`.pot`、`.potx`、`.potm`、`.pps`、`.ppsx`、`.ppsm`
- 通过 Microsoft PowerPoint COM 自动化（以只读方式打开）把幻灯片渲染成图片，并在纯 WPF 查看器中显示，视觉效果与 PowerPoint 本体一致
- **上/下键切换文件可用**——预览过程中按上/下键，会像其他 QuickLook 预览一样切换到资源管理器里的上一个/下一个文件（已端到端实测验证）。图片界面不会抢占前台焦点，资源管理器始终掌控选中项
- 预览内提供上一页/下一页**幻灯片**按钮，并预渲染下一张以加快翻页
- 不修改原始 PowerPoint 文件（只读打开）
- 可与原版 OfficeViewer 插件并排安装
- 使用 `Priority = 100`，优先接管 PowerPoint 文件，Word/Excel 仍由其他插件处理
- 预览关闭时清理临时渲染目录

## 限制

- 默认图片模式把幻灯片渲染为图片，因此**无法选择或复制幻灯片中的文字**
- 需要 Windows 上已安装并完成 COM 注册的 Microsoft PowerPoint
- 某个未缓存或较大文件首次预览时，仍需几秒钟启动 PowerPoint 并渲染
- 沙箱构建，供个人使用，尚非正式稳定版

## 可选 / 保留代码（非默认）

仓库还保留了两条更早的预览路径，供将来实现**可复制文字**模式使用，但**默认图片查看器并不使用它们**：

- Shell Preview Handler 宿主（类似 PowerToys Peek 的系统预览）
- 把 PowerPoint 导出为 PDF 后用 WebView2 显示，并带一个按文件路径/大小/修改时间为键的持久 PDF 缓存

启用这些会重新引入原生 / WebView2 子窗口——而这正是之前破坏上/下键切换的原因，所以默认特意让它们处于休眠状态。其中 PDF 路径还需要 Microsoft Edge WebView2 Runtime。

## 加载页面

打开文件时，QuickLook 会在 PowerPoint 启动、首张幻灯片渲染期间短暂显示一个 “Starting PowerPoint…” 页面。这是文件首次打开时的预期现象，之后翻页会更快。

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

本插件不会修改原版 OfficeViewer 插件。

持久 PDF 缓存（仅在使用非默认的 PDF 模式时才会生成）位于：

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

可用的沙箱构建。默认的图片预览路径已验证：PowerPoint 幻灯片能正确渲染，且预览时 QuickLook 上/下键文件切换可用。可复制文字的 Shell/PDF 路径仍属实验性，默认禁用。
