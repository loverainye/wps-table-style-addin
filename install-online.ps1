param([switch]$Remove, [string]$Config)
$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
$Version = '1.0'; $Base = "https://github.com/loverainye/wps-table-style-addin/releases/download/v$Version"
$Temp = Join-Path ([IO.Path]::GetTempPath()) ("wps-table-style-$([guid]::NewGuid())")
New-Item -ItemType Directory -Path $Temp | Out-Null
try {
  $Asset = "wps-table-style-addin-windows-$Version.zip"
  Invoke-WebRequest -UseBasicParsing "$Base/$Asset" -OutFile (Join-Path $Temp $Asset)
  Expand-Archive (Join-Path $Temp $Asset) -DestinationPath $Temp -Force
  $Install = Get-ChildItem $Temp -Filter install.ps1 -Recurse | Select-Object -First 1
  if (-not $Install) { throw '安装包中缺少 install.ps1。' }
  & $Install.FullName @PSBoundParameters
} finally { Remove-Item $Temp -Recurse -Force -ErrorAction SilentlyContinue }
