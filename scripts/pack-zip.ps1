$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$release = Join-Path $root 'bin\Release'
$out = Join-Path $root 'QuickLook.Plugin.PowerPointNativeViewer.qlplugin'
$zip = Join-Path $root 'QuickLook.Plugin.PowerPointNativeViewer.zip'

Remove-Item $out -ErrorAction SilentlyContinue
Remove-Item $zip -ErrorAction SilentlyContinue

$files = @(
    (Join-Path $release 'QuickLook.Plugin.PowerPointNativeViewer.dll'),
    (Join-Path $release 'QuickLook.Plugin.Metadata.config')
)

Compress-Archive -Path $files -DestinationPath $zip
Move-Item $zip $out
Write-Host $out
