# QuickLook PowerPoint Native Viewer

[中文说明](README.zh-CN.md)

A QuickLook plugin that previews PowerPoint files by rendering their slides to images through the locally installed Microsoft PowerPoint engine.

It is a higher-fidelity alternative to `QuickLook.Plugin.OfficeViewer` (Syncfusion) for PPT/PPTX files that render poorly or fail there — while keeping QuickLook's normal Up/Down file navigation working.

## Capabilities (current version)

- Handles PowerPoint formats: `.ppt`, `.pptx`, `.pptm`, `.pot`, `.potx`, `.potm`, `.pps`, `.ppsx`, `.ppsm`
- Renders slides to images via Microsoft PowerPoint COM automation (opened read-only) and shows them in a pure WPF viewer, so visual fidelity matches PowerPoint itself
- **Up/Down file navigation works** — while previewing, pressing Up/Down switches to the previous/next file in Explorer just like any other QuickLook preview (verified end-to-end). The image surface does not steal foreground focus, so Explorer stays in control of the selection
- Previous/next **slide** buttons inside the preview, and pre-renders the next slide for faster paging
- Single, image-only preview path — no WebView2/PDF/Shell-handler dependencies
- Leaves the original PowerPoint file untouched (read-only)
- Installs side-by-side with the original OfficeViewer plugin
- Uses `Priority = 100` so it takes over PowerPoint files, while Word/Excel stay with other plugins
- Cleans up its temporary render folder when the preview closes

## Limitations

- Slides are rendered as images, so **slide text cannot be selected or copied**
- Requires Microsoft PowerPoint to be installed and COM-registered on Windows
- The first preview of a large file still needs a few seconds while PowerPoint starts and renders
- Sandbox build intended for personal use, not yet a polished release

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

Temporary render images are written under `%TEMP%\QuickLook.PowerPointNativeViewer` and are cleaned up automatically when a preview closes (and on plugin startup for stale folders).

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

Working sandbox build. The image-preview path is verified: PowerPoint slides render correctly and QuickLook Up/Down file navigation works while previewing. The earlier PDF/WebView2 and Shell Preview Handler experiments have been removed in favor of this single image-only mode.
