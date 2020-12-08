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

if ((get-command sqlite3).CommandType -ne 'Application') {
  Write-Warning "Install the sqlite executable on the path, and the script will attempt to repair any potential errors in the database."
  exit
}
else {
  $output = sqlite3 cellmapperdata.db "pragma integrity_check" 2>&1
  if ($output | where-object { $_ -is [System.Management.Automation.ErrorRecord] }) {
    sqlite3 cellmapperdata.db ".recover" | sqlite3 recovered.db
    Move-Item Recovered.db cellmapperdata.db -Force
  }

}