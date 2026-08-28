Add-Type -AssemblyName System.IO.Compression.FileSystem
$path = "AgoraIn.ClassIslandPlugin\bin\Release\net8.0\AgoraIn.ClassIslandPlugin.cipx"
$z = [System.IO.Compression.ZipFile]::OpenRead((Resolve-Path $path))
$z.Entries | ForEach-Object { $_.FullName }
$z.Dispose()
