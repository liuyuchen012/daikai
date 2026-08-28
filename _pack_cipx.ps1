$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$projDir = (Resolve-Path 'AgoraIn.ClassIslandPlugin').Path
$src = Join-Path $projDir 'bin\Release\net8.0'
$dstDir = Join-Path $projDir 'cipx'
$dst = Join-Path $dstDir 'AgoraIn.ClassIslandPlugin.cipx'
New-Item -ItemType Directory -Force -Path $dstDir | Out-Null
if (Test-Path $dst) { Remove-Item -Force $dst }
[System.IO.Compression.ZipFile]::CreateFromDirectory($src, $dst)
Get-Item $dst | Select-Object FullName, Length, LastWriteTime | Format-List
