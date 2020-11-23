[CmdletBinding()]
param (
  [Parameter()]
  [string]
  $Path = 'cellmapperdata.db'
)

$tempDir = Join-Path $env:TEMP ([System.IO.Path]::GetRandomFileName())
$tempFile = Join-Path $tempDir cellmapper.ab
$temptar = Join-Path $tempDir 'cellmapper.tar.gz'
mkdir -force $tempDir

adb wait-for-device

adb backup -f "$tempFile" -noapk cellmapper.net.cellmapper

$infile = [System.IO.File]::OpenRead($tempFile)
$infile.Seek(24, [System.IO.SeekOrigin]::Begin)

$header = @(0x1f, 0x8b, 0x08, 0x00, 0x00, 0x00, 0x00, 0x00)
$outfile = [System.IO.File]::OpenWrite($temptar)
$outfile.Write($header, 0, $header.Length)

$infile.CopyTo($outfile)
$infile.Close()
$outfile.Close()

tar -zxvf "$temptar" -C "$tempdir" apps/cellmapper.net.cellmapper/db/cellmapperdata.db

Move-Item (Join-Path $tempDir apps/cellmapper.net.cellmapper/db/cellmapperdata.db) $Path -Force

Remove-Item $tempDir -Recurse -force
