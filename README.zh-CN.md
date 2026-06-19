# QuickLook PowerPoint Native Viewer

[English](README.md)

一个 QuickLook 插件，通过本机安装的 Microsoft PowerPoint 引擎，把幻灯片渲染成图片来预览 PowerPoint 文件。

它是 `QuickLook.Plugin.OfficeViewer`（Syncfusion）的高保真替代方案，适用于在原插件里预览失真或直接报错的 PPT/PPTX 文件——同时**保持 QuickLook 正常的上/下键文件切换可用**。

## 当前版本能力

- 支持 PowerPoint 格式：`.ppt`、`.pptx`、`.pptm`、`.pot`、`.potx`、`.potm`、`.pps`、`.ppsx`、`.ppsm`
- 通过 Microsoft PowerPoint COM 自动化（以只读方式打开）把幻灯片渲染成图片，并在纯 WPF 查看器中显示，视觉效果与 PowerPoint 本体一致
- **上/下键切换文件可用**——预览过程中按上/下键，会像其他 QuickLook 预览一样切换到资源管理器里的上一个/下一个文件（已端到端实测验证）。图片界面不会抢占前台焦点，资源管理器始终掌控选中项
- 预览内提供上一页/下一页**幻灯片**按钮，并预渲染下一张以加快翻页
- 单一图片预览路径——不依赖 WebView2 / PDF / Shell 预览处理器
- 不修改原始 PowerPoint 文件（只读打开）
- 可与原版 OfficeViewer 插件并排安装
- 使用 `Priority = 100`，优先接管 PowerPoint 文件，Word/Excel 仍由其他插件处理
- 预览关闭时清理临时渲染目录

## 限制

- 幻灯片以图片形式渲染，因此**无法选择或复制幻灯片中的文字**
- 需要 Windows 上已安装并完成 COM 注册的 Microsoft PowerPoint
- 较大文件首次预览时，仍需几秒钟启动 PowerPoint 并渲染
- 沙箱构建，供个人使用，尚非正式稳定版

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

临时渲染图片写在 `%TEMP%\QuickLook.PowerPointNativeViewer` 下，预览关闭时会自动清理（插件启动时也会清理过期目录）。

## 构建说明

项目目标框架为 .NET Framework 4.6.2，并引用 QuickLook 的 `QuickLook.Common.dll`。

本地构建前，请先把 `QuickLook.Common.dll` 复制到：

```text
lib/QuickLook.Common.dll
```

然后运行：

```powershell
MSBuild.exe QuickLook.Plugin.PowerPointNativeViewer.sln /p:Configuration=Release
powershell -ExecutionPolicy Bypass -File scripts/pack-zip.ps1
```

## 状态

可用的沙箱构建。图片预览路径已验证：PowerPoint 幻灯片能正确渲染，且预览时 QuickLook 上/下键文件切换可用。早期的 PDF/WebView2 与 Shell Preview Handler 实验已移除，现仅保留这一单一图片模式。
