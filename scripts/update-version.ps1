$ErrorActionPreference = 'Stop'

$root = Split-Path -Parent $PSScriptRoot
$versionFile = Join-Path $root 'GitVersion.cs'

@'
using System.Reflection;

[assembly: AssemblyInformationalVersion("0.1.0-sandbox")]
'@ | Set-Content -LiteralPath $versionFile -Encoding UTF8
