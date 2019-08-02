[xml]$vsix = Get-Content "$PSScriptRoot\..\src\GtmExtension\source.extension.vsixmanifest"
$version = $vsix.PackageManifest.Metadata.Identity.Version
Write-Host "##vso[task.setvariable variable=vsix.version]$version"
