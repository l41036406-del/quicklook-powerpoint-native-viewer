# QuickLook PowerPoint Native Viewer

[中文说明](README.zh-CN.md)

Experimental QuickLook plugin for previewing PowerPoint files through the local Microsoft PowerPoint rendering engine instead of Syncfusion.

This is a prototype created to test a safer preview path for PPT/PPTX files that render poorly or fail in `QuickLook.Plugin.OfficeViewer`.

## Current Features

- Handles PowerPoint formats: `.ppt`, `.pptx`, `.pptm`, `.pot`, `.potx`, `.potm`, `.pps`, `.ppsx`, `.ppsm`
- Uses Microsoft PowerPoint COM automation to open the source presentation read-only
- Exports the current slide to PNG and displays it in QuickLook
- Provides previous/next page navigation
- Pre-renders the next page in the background
- Deletes the session image cache when the preview control unloads
- Leaves the original PowerPoint file unchanged
- Installs side-by-side with the original OfficeViewer plugin
- Uses `Priority = 100` so it can take over PowerPoint files while Word/Excel remain handled by other plugins

## Current Limitations

- Requires Microsoft PowerPoint to be installed and COM-registered on Windows
- Preview is image-based, so text selection and copying are not supported yet
- Startup can be slower than Syncfusion because PowerPoint must be launched
- The current implementation is a sandbox prototype, not a polished release
- The next planned direction is `PowerPoint -> temporary PDF -> WebView2/PDF.js` to support selectable text

## Startup Loading Screen

When the plugin launches PowerPoint, QuickLook may show this black loading page:

![Starting PowerPoint loading screen](docs/images/starting-powerpoint-stuck.png)

This screen is expected during startup. It indicates that the plugin is starting PowerPoint and preparing the first rendered slide.

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

## Build Notes

This project targets .NET Framework 4.6.2 and references QuickLook's `QuickLook.Common.dll`.

Before building locally, copy `QuickLook.Common.dll` into:

```text
lib/QuickLook.Common.dll
```

Then run:

```powershell
MSBuild.exe QuickLook.Plugin.PowerPointNativeViewer.sln /p:Configuration=Release
powershell -ExecutionPolicy Bypass -File scripts/pack-zip.ps1
```

## Status

Prototype. Published for tracking the current PowerPoint-native PNG preview experiment and its known behavior.
