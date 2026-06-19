# QuickLook PowerPoint Native Viewer

[中文说明](README.zh-CN.md)

A QuickLook plugin that previews PowerPoint files by rendering their slides to images through the locally installed Microsoft PowerPoint engine.

It is a higher-fidelity alternative to `QuickLook.Plugin.OfficeViewer` (Syncfusion) for PPT/PPTX files that render poorly or fail there — while keeping QuickLook's normal Up/Down file navigation working.

## Capabilities (current version)

- Handles PowerPoint formats: `.ppt`, `.pptx`, `.pptm`, `.pot`, `.potx`, `.potm`, `.pps`, `.ppsx`, `.ppsm`
- Renders slides to images via Microsoft PowerPoint COM automation (opened read-only) and shows them in a pure WPF viewer, so visual fidelity matches PowerPoint itself
- **Up/Down file navigation works** — while previewing, pressing Up/Down switches to the previous/next file in Explorer just like any other QuickLook preview (verified end-to-end). The image surface does not steal foreground focus, so Explorer stays in control of the selection
- Previous/next **slide** buttons inside the preview, and pre-renders the next slide for faster paging
- Leaves the original PowerPoint file untouched (read-only)
- Installs side-by-side with the original OfficeViewer plugin
- Uses `Priority = 100` so it takes over PowerPoint files, while Word/Excel stay with other plugins
- Cleans up its temporary render folder when the preview closes

## Limitations

- The default image mode renders slides as pictures, so **slide text cannot be selected or copied**
- Requires Microsoft PowerPoint to be installed and COM-registered on Windows
- The first preview of an uncached or large file still needs a few seconds while PowerPoint starts and renders
- Sandbox build intended for personal use, not yet a polished release

## Optional / retained code (not the default)

The repository also keeps two earlier preview paths for a possible future *selectable-text* mode. They are **not** used by the default image viewer:

- A Shell Preview Handler host (PowerToys Peek-style system preview)
- A PowerPoint-to-PDF path shown in WebView2, with a persistent PDF cache keyed by file path/size/modified-time

Enabling these would re-introduce native/WebView2 child windows, which is exactly what previously broke Up/Down navigation, so they are kept dormant by design. The PDF path additionally needs the Microsoft Edge WebView2 Runtime.

## Loading screen

When opening a file, QuickLook briefly shows a "Starting PowerPoint…" page while PowerPoint launches and the first slide renders. This is expected on the first open of a file; subsequent paging is faster.

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

The original OfficeViewer plugin is not modified by this plugin.

The persistent PDF cache (only created if the non-default PDF mode is used) is stored under:

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

Working sandbox build. The default image-preview path is verified: PowerPoint slides render correctly and QuickLook Up/Down file navigation works while previewing. The selectable-text (Shell/PDF) paths remain experimental and disabled by default.
