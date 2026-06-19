# QuickLook PowerPoint Native Viewer

[中文说明](README.zh-CN.md)

Experimental QuickLook plugin for previewing PowerPoint files through the local Microsoft PowerPoint rendering engine instead of Syncfusion.

This is a prototype created to test a safer preview path for PPT/PPTX files that render poorly or fail in `QuickLook.Plugin.OfficeViewer`.

## Current Features

- Handles PowerPoint formats: `.ppt`, `.pptx`, `.pptm`, `.pot`, `.potx`, `.potm`, `.pps`, `.ppsx`, `.ppsm`
- Uses Microsoft PowerPoint COM automation to open the source presentation read-only
- Renders slides to PNG images and displays them in a pure WPF viewer
- Avoids the WPS/Shell native preview child window so QuickLook Up/Down file navigation can keep working
- Provides previous/next slide buttons inside the PowerPoint preview
- Keeps experimental Shell Preview Handler and PDF/WebView2 code in the repository for future selectable-text modes
- Reuses a persistent PDF cache when the source file path, size, and modified time are unchanged
- Deletes WebView2 session data when the preview control unloads
- Tries to bring the QuickLook preview window to the front when it opens
- Leaves the original PowerPoint file unchanged
- Installs side-by-side with the original OfficeViewer plugin
- Uses `Priority = 100` so it can take over PowerPoint files while Word/Excel remain handled by other plugins

## Current Limitations

- Requires Microsoft PowerPoint to be installed and COM-registered on Windows
- Default preview is image-based, so text selection/copy is not available yet
- First startup for a file can still take time because PowerPoint must be launched
- The current implementation is a sandbox prototype, not a polished release
- Requires Microsoft Edge WebView2 Runtime, which is usually already installed on modern Windows systems

## Startup Loading Screen

When the plugin launches PowerPoint, QuickLook may show this loading page:

![Starting PowerPoint loading screen](docs/images/starting-powerpoint-stuck.png)

This screen is expected during the first startup of a file. It indicates that the plugin is starting PowerPoint and rendering the first slide preview.

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

Prototype. Published for tracking the current PowerPoint image preview experiment and its known behavior.
