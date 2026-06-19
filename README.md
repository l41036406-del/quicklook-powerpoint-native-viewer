# QuickLook PowerPoint Native Viewer

[中文说明](README.zh-CN.md)

Experimental QuickLook plugin for previewing PowerPoint files through the Windows Shell Preview Handler or the local Microsoft PowerPoint rendering engine instead of Syncfusion.

This is a prototype created to test a safer preview path for PPT/PPTX files that render poorly or fail in `QuickLook.Plugin.OfficeViewer`.

## Current Features

- Handles PowerPoint formats: `.ppt`, `.pptx`, `.pptm`, `.pot`, `.potx`, `.potm`, `.pps`, `.ppsx`, `.ppsm`
- Tries the registered Windows Shell Preview Handler first, following the same general idea used by PowerToys Peek
- Falls back to Microsoft PowerPoint COM automation when the system preview handler is unavailable or fails
- Exports the presentation to a temporary PDF in the fallback path
- Displays fallback PDFs through WebView2, using the installed Microsoft Edge WebView2 Runtime
- Supports text selection/copy when the active backend exposes selectable text, such as the Shell Preview Handler or WebView2 PDF viewer
- Reuses a persistent PDF cache when the source file path, size, and modified time are unchanged
- Deletes WebView2 session data when the preview control unloads
- Tries to bring the QuickLook preview window to the front when it opens
- Leaves the original PowerPoint file unchanged
- Installs side-by-side with the original OfficeViewer plugin
- Uses `Priority = 100` so it can take over PowerPoint files while Word/Excel remain handled by other plugins

## Current Limitations

- Shell Preview Handler behavior depends on the preview handler registered on the local Windows system
- Requires Microsoft PowerPoint to be installed and COM-registered on Windows for the PDF fallback path
- First fallback startup for a file can be slower than Syncfusion because PowerPoint must be launched and a PDF must be generated
- Subsequent fallback previews of the same unchanged file should open faster from the persistent PDF cache
- The current implementation is a sandbox prototype, not a polished release
- Requires Microsoft Edge WebView2 Runtime, which is usually already installed on modern Windows systems

## Startup Loading Screen

When the plugin can use the Shell Preview Handler, this PowerPoint startup page should be avoided. If the plugin falls back to launching PowerPoint, QuickLook may show this black loading page:

![Starting PowerPoint loading screen](docs/images/starting-powerpoint-stuck.png)

This screen is expected only during the PDF fallback path for the first startup of a file. It indicates that the plugin is starting PowerPoint and generating the PDF preview. If the Shell Preview Handler works or a cached PDF is available, this screen should appear for a much shorter time or not be noticeable.

## Install

Download or build:

```text
dist/QuickLook.Plugin.PowerPointNativeViewer.qlplugin
```

Then install it in QuickLook, or extract it into QuickLook's user plugin directory:

```text
%LOCALAPPDATA%\Packages\21090PaddyXu.QuickLook_egxr34yet59cg\LocalCache\Roaming\pooi.moe\QuickLook\QuickLook.Plugin\QuickLook.Plugin.PowerPointNativeViewer
```

Restart QuickLook after installing.

## Rollback

Remove this folder and restart QuickLook:

```text
%LOCALAPPDATA%\Packages\21090PaddyXu.QuickLook_egxr34yet59cg\LocalCache\Roaming\pooi.moe\QuickLook\QuickLook.Plugin\QuickLook.Plugin.PowerPointNativeViewer
```

The original OfficeViewer plugin is not modified by this prototype.

The persistent PDF cache is stored under:

```text
%LOCALAPPDATA%\QuickLook.PowerPointNativeViewer\pdf-cache
```

## Build Notes

This project targets .NET Framework 4.6.2 and references QuickLook's `QuickLook.Common.dll` plus Microsoft WebView2 SDK assemblies.

Before building locally, copy `QuickLook.Common.dll` into:

```text
lib/QuickLook.Common.dll
```

Also place the WebView2 SDK assemblies in:

```text
lib/webview2/Microsoft.Web.WebView2.Core.dll
lib/webview2/Microsoft.Web.WebView2.Wpf.dll
lib/webview2/runtimes/win-x64/native/WebView2Loader.dll
```

Then run:

```powershell
MSBuild.exe QuickLook.Plugin.PowerPointNativeViewer.sln /p:Configuration=Release
powershell -ExecutionPolicy Bypass -File scripts/pack-zip.ps1
```

## Status

Prototype. Published for tracking the current Shell Preview Handler plus PowerPoint PDF/WebView2 fallback experiment and its known behavior.
