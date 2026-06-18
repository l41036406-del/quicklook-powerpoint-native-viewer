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
