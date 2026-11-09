param()
$ErrorActionPreference = 'Stop'
$Project = Split-Path -Parent $PSScriptRoot
$TestDir = Join-Path ([IO.Path]::GetTempPath()) ('wps-install-test-' + [guid]::NewGuid())
$Config = Join-Path $TestDir 'WPS config with spaces/publish.xml'
$AddinDir = Split-Path -Parent $Config
$Version = (Get-Content (Join-Path $Project 'package.json') -Raw | ConvertFrom-Json).version
$Destination = Join-Path $AddinDir "WpsPresentationTableStyles_$Version"
function Assert([bool]$Condition, [string]$Message) { if (-not $Condition) { throw $Message } }
try {
    New-Item -ItemType Directory -Force -Path $AddinDir | Out-Null
    Set-Content $Config '<jsplugins><!--keep--><jspluginonline name="Other" type="et" url="https://example.com"/><jspluginonline name="WpsPresentationTableStyles" type="wpp" url="http://127.0.0.1:3989/"/></jsplugins>' -Encoding UTF8
    & (Join-Path $Project 'install.ps1') -Config $Config
    [xml]$Registration = Get-Content $Config -Raw
    Assert ($Registration.jsplugins.jsplugin.version -eq $Version) 'Version mismatch'
    Assert ($Registration.jsplugins.jsplugin.url -eq (Split-Path $Destination -Leaf)) 'Not a local registration'
    Assert ($Registration.jsplugins.jspluginonline.name -eq 'Other') 'Other plugin changed'
    Assert (-not (Test-Path (Join-Path $Destination 'index.html'))) 'Unexpected online entry'
    Assert (Test-Path (Join-Path $Destination 'editor.js')) 'Missing resource'
    Assert ((Get-ChildItem $AddinDir -Filter 'publish.xml.bak-*').Count -eq 1) 'Missing config backup'
    & (Join-Path $Project 'install.ps1') -Config $Config
    Assert ((Get-ChildItem $AddinDir -Directory -Filter "WpsPresentationTableStyles_$Version.bak-*").Count -eq 1) 'Missing resource backup'
    & (Join-Path $Project 'uninstall.ps1') -Config $Config
    & (Join-Path $Project 'uninstall.ps1') -Config $Config
    Assert (-not (Test-Path $Destination)) 'Uninstall left managed resources'
    [xml]$Registration = Get-Content $Config -Raw
    Assert ($Registration.jsplugins.jspluginonline.name -eq 'Other') 'Uninstall removed other plugin'
    New-Item -ItemType Directory -Path $Destination | Out-Null
    Set-Content (Join-Path $Destination 'keep') 'unmanaged'
    $Rejected = $false
    try { & (Join-Path $Project 'install.ps1') -Config $Config } catch { $Rejected = $true }
    Assert $Rejected 'Unmanaged destination was accepted'
    Assert ((Get-Content (Join-Path $Destination 'keep')) -eq 'unmanaged') 'Unmanaged data changed'
    Write-Host 'PowerShell install/reinstall/uninstall and ownership checks passed.'
} finally { if (Test-Path $TestDir) { Remove-Item -LiteralPath $TestDir -Recurse -Force } }
