# Release Notes

## v0.1.0-sandbox

Initial sandbox prototype.

- Adds a QuickLook plugin that previews PowerPoint files using Microsoft PowerPoint COM automation.
- Renders slides as PNG images instead of using Syncfusion.
- Supports previous/next slide navigation.
- Cleans up the temporary session image folder when the preview unloads.
- Includes the current `.qlplugin` package under `dist/`.
- Documents the startup loading page shown while PowerPoint is launching.

Known limitations:

- Requires local Microsoft PowerPoint.
- Does not support text selection or copying yet.
- Startup may show a black "Starting PowerPoint..." loading page.
- Intended for testing, not production use.

## v0.2.0-sandbox

Switches the prototype from PNG slide rendering to a temporary PDF preview path.

- Exports the source presentation to `preview.pdf` through Microsoft PowerPoint.
- Displays the temporary PDF with WebView2.
- Enables PDF text selection/copy through the embedded WebView2 PDF viewer.
- Closes PowerPoint after PDF generation instead of keeping it open during preview.
- Tracks the PowerPoint process created by the plugin and cleans it up if COM `Quit()` does not exit cleanly.
- Packages WebView2 runtime loader and managed SDK assemblies with the `.qlplugin`.

Known limitations:

- Still requires local Microsoft PowerPoint.
- First preview can be slower because the whole temporary PDF is generated up front.
- WebView2 Runtime must be available on the machine.

## v0.3.0-sandbox

Keeps `v0.2.0-sandbox` as the second version and adds usability optimizations.

- Adds a persistent PDF cache keyed by source file path, file size, and modified time.
- Reuses the cached PDF on subsequent previews of the same unchanged presentation.
- Keeps WebView2 session data temporary and deletes it when the preview unloads.
- Cleans persistent cache folders older than 14 days during plugin initialization.
- Tries to bring the QuickLook preview window to the foreground when the control loads and when the PDF opens.

Known limitations:

- First preview of a changed or uncached file still needs PowerPoint startup and PDF generation.
- Foreground activation depends on Windows focus rules, so it is a best-effort improvement rather than a hard guarantee.

## v0.4.0-sandbox

Adds a PowerToys Peek-inspired Shell Preview Handler path before the PDF fallback.

- Checks the Windows registry for the file type's registered Shell Preview Handler.
- Hosts the Shell Preview Handler inside a WPF `HwndHost` child window.
- Initializes handlers through `IInitializeWithFile`, `IInitializeWithStream`, or `IInitializeWithItem`.
- Uses the system preview path first, then falls back to the existing cached PDF/WebView2 pipeline if the handler is unavailable or fails.
- Keeps the persistent PDF cache and WebView2 cleanup behavior from `v0.3.0-sandbox`.
- Leaves plain Up/Down arrow keys for QuickLook file navigation instead of forwarding them to the Shell Preview Handler.

Known limitations:

- Shell Preview Handler quality, speed, and text selection behavior depend on the local Office/WPS/Windows preview handler registration.
- If the system handler hangs inside COM initialization, fallback cannot always interrupt it immediately.
- PDF fallback still requires local PowerPoint.

## v0.5.0-sandbox

Switches the default preview path back to a PowerPoint-rendered image surface.

- Uses a pure WPF image viewer by default, avoiding the WPS/Shell native child window that blocked QuickLook Up/Down file navigation.
- Renders the first slide quickly through PowerPoint COM and pre-renders the next slide.
- Keeps the PDF/WebView2 and Shell Preview Handler code in the repository for future selectable-text modes, but does not use them as the default viewer.

Known limitations:

- Default mode is image-based, so slide text is not selectable.
- Requires local PowerPoint.

## v0.5.1-sandbox

Verifies and documents the image-preview navigation behavior; no functional code changes.

- Verified end-to-end that QuickLook Up/Down file navigation works while previewing PowerPoint in the default image mode: pressing Up/Down switches the previewed file (slide image included), and the foreground window stays on Explorer throughout.
- Confirmed the root cause of the earlier navigation failures: the previous Shell Preview Handler / WebView2 child windows stole foreground focus from Explorer, so QuickLook could no longer read the moved selection. The pure WPF image surface does not steal focus, which restores navigation.
- Updated the plugin manifest description to match the current image-rendering default.

Known limitations:

- Default mode is image-based, so slide text is not selectable.
- Requires local PowerPoint.
- Switching to a not-yet-cached large presentation still incurs PowerPoint startup latency.

## v0.5.2-sandbox

Removes the unused alternate preview paths and their caches, leaving a single image-only viewer.

- Removed the PowerPoint-to-PDF / WebView2 preview mode (`PowerPointPdfPreviewControl`).
- Removed the Shell Preview Handler host (`ShellPreviewHandlerHost`) that only fed the PDF path.
- Dropped the now-unused `ExportPdf` helper, the WebView2 SDK references, and the WebView2 files from the `.qlplugin` package.
- Deleted the experimental persistent PDF cache and the temporary WebView2 session data.
- Result: the plugin DLL and package are much smaller, and there is exactly one preview path (image rendering) with working Up/Down navigation.

Known limitations:

- Slides are rendered as images, so slide text is not selectable.
- Requires local PowerPoint.
- The first preview of a large file still incurs PowerPoint startup latency.
