[xml]$vsix = Get-Content $PSScriptRoot\..\src\GtmExtension\source.extension.vsixmanifest
$env:VSIX_VERSION = $vsix.PackageManifest.Metadata.Identity.Version
