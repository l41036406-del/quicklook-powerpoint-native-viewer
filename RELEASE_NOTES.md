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
