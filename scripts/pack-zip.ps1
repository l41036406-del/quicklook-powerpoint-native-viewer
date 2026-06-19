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

$staging = Join-Path $root 'obj\package'
Remove-Item $staging -Recurse -Force -ErrorAction SilentlyContinue
New-Item -ItemType Directory -Path $staging | Out-Null

foreach ($file in $files) {
    Copy-Item -LiteralPath $file -Destination $staging -Force
}

Compress-Archive -Path (Join-Path $staging '*') -DestinationPath $zip
Move-Item $zip $out
Write-Host $out
