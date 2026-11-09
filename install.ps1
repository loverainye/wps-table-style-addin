param([switch]$Remove, [string]$Config)
$ErrorActionPreference = 'Stop'
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Package = Get-Content (Join-Path $Root 'package.json') -Raw -Encoding UTF8 | ConvertFrom-Json
$Version = [string]$Package.version
$Name = 'WpsPresentationTableStyles'
$Marker = '.wps-table-style-install.json'
$Stamp = Get-Date -Format 'yyyyMMdd-HHmmss-fffffff'
if ($Version -notmatch '^[A-Za-z0-9.-]+$') { throw '加载项版本号格式错误。' }
if (-not $Config) {
  $Roaming = if ($env:APPDATA) { $env:APPDATA } else { [Environment]::GetFolderPath('ApplicationData') }
  $Candidates = @((Join-Path $Roaming 'Kingsoft/wps/jsaddons/publish.xml'), (Join-Path $Roaming 'Kingsoft/WPS Office/jsaddons/publish.xml'))
  $Config = $Candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
  if (-not $Config) { $Config = $Candidates[0] }
}
$Config = [IO.Path]::GetFullPath($Config)
$AddinDir = Split-Path -Parent $Config
$Destination = Join-Path $AddinDir ("{0}_{1}" -f $Name, $Version)
function Test-Owned([string]$Directory) {
  if (-not (Test-Path -LiteralPath $Directory -PathType Container)) { return $false }
  if ((Get-Item -LiteralPath $Directory).Attributes -band [IO.FileAttributes]::ReparsePoint) { return $false }
  try { return (Get-Content (Join-Path $Directory $Marker) -Raw -Encoding UTF8 | ConvertFrom-Json).name -eq $Name } catch { return $false }
}
$Xml = New-Object System.Xml.XmlDocument
$Xml.PreserveWhitespace = $true
if (Test-Path -LiteralPath $Config) {
  $Settings = New-Object System.Xml.XmlReaderSettings
  $Settings.DtdProcessing = [System.Xml.DtdProcessing]::Prohibit
  $Reader = [System.Xml.XmlReader]::Create($Config, $Settings)
  try { $Xml.Load($Reader) } finally { $Reader.Dispose() }
} else { [void]$Xml.AppendChild($Xml.CreateElement('jsplugins')) }
if ($Xml.DocumentElement.Name -ne 'jsplugins') { throw '已有 publish.xml 根节点未知，未修改。' }
$Entries = @($Xml.DocumentElement.ChildNodes | Where-Object { $_.NodeType -eq [Xml.XmlNodeType]::Element -and $_.GetAttribute('name') -eq $Name })
foreach ($Entry in $Entries) { [void]$Xml.DocumentElement.RemoveChild($Entry) }
function Save-Registration {
  New-Item -ItemType Directory -Force -Path $AddinDir | Out-Null
  $Pending = Join-Path $AddinDir ('.publish-' + [guid]::NewGuid() + '.tmp')
  try {
    $Xml.Save($Pending)
    if (Test-Path -LiteralPath $Config) {
      Copy-Item -LiteralPath $Config -Destination "$Config.bak-$Stamp"
      [IO.File]::Replace($Pending, $Config, [System.Management.Automation.Language.NullString]::Value)
    } else { [IO.File]::Move($Pending, $Config) }
  } finally { if (Test-Path -LiteralPath $Pending) { Remove-Item -LiteralPath $Pending -Force } }
}
if ($Remove) {
  if ($Entries.Count -gt 0) { Save-Registration }
  if (Test-Path -LiteralPath $AddinDir) {
    Get-ChildItem -LiteralPath $AddinDir -Directory -Filter ($Name + '_*') | Where-Object { $_.Name -notlike '*.bak-*' -and (Test-Owned $_.FullName) } | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Recurse -Force }
  }
  Write-Host '已移除本加载项，保留其他加载项和备份。请重新打开 WPS 演示。'
  return
}
foreach ($File in @('js/main.js','ribbon.xml','editor.html','editor.js','editor.css','table-style.svg')) {
  if (-not (Test-Path -LiteralPath (Join-Path (Join-Path $Root 'dist') $File) -PathType Leaf)) { throw "缺少构建资源：$File" }
}
if ((Test-Path -LiteralPath $Destination) -and -not (Test-Owned $Destination)) { throw "目标目录不属于本安装程序，未覆盖：$Destination" }
New-Item -ItemType Directory -Force -Path $AddinDir | Out-Null
$Staged = Join-Path $AddinDir ('.' + $Name + '-install-' + [guid]::NewGuid())
$Previous = $null
$Moved = $false
try {
  New-Item -ItemType Directory -Path $Staged | Out-Null
  Get-ChildItem -LiteralPath (Join-Path $Root 'dist') | Where-Object { $_.Name -ne 'index.html' -and $_.Extension -ne '.map' } | Copy-Item -Destination $Staged -Recurse -Force
  Get-ChildItem -LiteralPath $Staged -Recurse -File -Filter '*.map' | Remove-Item -Force
  @{ name=$Name; version=$Version } | ConvertTo-Json | Set-Content (Join-Path $Staged $Marker) -Encoding UTF8
  if (Test-Path -LiteralPath $Destination) { $Previous = "$Destination.bak-$Stamp"; Move-Item -LiteralPath $Destination -Destination $Previous }
  Move-Item -LiteralPath $Staged -Destination $Destination
  $Moved = $true
  $Plugin = $Xml.CreateElement('jsplugin')
  foreach ($Pair in @{name=$Name;type='wpp';url=(Split-Path $Destination -Leaf);version=$Version;debug='';enable='enable_dev';install='null';customDomain=''}.GetEnumerator()) { $Plugin.SetAttribute($Pair.Key, [string]$Pair.Value) }
  [void]$Xml.DocumentElement.AppendChild($Plugin)
  Save-Registration
} catch {
  if ($Moved -and (Test-Path -LiteralPath $Destination)) { Remove-Item -LiteralPath $Destination -Recurse -Force }
  if ($Previous) { Move-Item -LiteralPath $Previous -Destination $Destination }
  throw
} finally { if (Test-Path -LiteralPath $Staged) { Remove-Item -LiteralPath $Staged -Recurse -Force } }
Write-Host "已安装 Windows 离线加载项：$Destination"
Write-Host '请保存工作并重新打开 WPS 演示。'
