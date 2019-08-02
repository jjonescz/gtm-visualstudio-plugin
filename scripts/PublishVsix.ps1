$VisualStudio = & "${env:ProgramFiles(x86)}\Microsoft Visual Studio\Installer\vswhere.exe" -latest -prerelease -format json | ConvertFrom-Json
$VsixPublisher = Join-Path -Path $VisualStudio.installationPath -ChildPath "VSSDK\VisualStudioIntegration\Tools\Bin\VsixPublisher.exe" -Resolve
& $VsixPublisher publish -payload $env:vsix -publishManifest "$PSScriptRoot\vsix-publish.json" -personalAccessToken $env:MARKETPLACE_TOKEN
