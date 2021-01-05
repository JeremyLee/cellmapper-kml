[CmdletBinding()]
param (
  $filterENBs = $null,
  $filename = $null,
  [Nullable[datetime]]$enbsSeenSince = $null,
  [switch]$noCircles,
  [switch]$noLines,
  [switch]$noPoints,
  [switch]$noTowers,
  [switch]$noToMove,
  [switch]$refreshCalculated
)

#.\Get-CellmapperDB.ps1
. .\helpers.ps1

ImportInstall-Module PSSqlite
Add-Type -AssemblyName 'system.drawing'



$timingAdvances = @{
  "310-260" = @{
    0  = 150
    41 = { param($ta) ($ta - 20) * 150 }
  }
  "310-120" = @{
    0  = 150
    41 = { param($ta) ($ta - 20) * 150 }
  }
  ""        = @{
    0 = 150
  }
}

$taRegex = [regex]'&LTE_TA=(?<TA>[0-9]+)&'
$bandRegex = [regex]'&INFO_BAND_NUMBER=(?<Band>[0-9]+)&'
Write-Host "Reading DB"

# Filter Types:
#  - eNBs seens since date
#  - eNBs
# Both work the same. If the date is specified, it simply generates a list of eNBs and passes it to the second filter.
    
if ($null -ne $filterENBs -and $null -ne $enbsSeenSince) {
  throw 'Can''t specify both $filterENBs and $enbsSeenSince.'
  return
}

$dbNames = Get-ChildItem -Path $PSScriptRoot -Filter '*.db' | foreach-object { $_.FullName }

if ($null -ne $enbsSeenSince) {
  # Generate a list of eNBs from all the DBs
  $filterENBs = @()
  foreach ($db in $dbNames) {
    $enbs = Invoke-SqliteQuery -Database $db -Query "select (CID >> 8) as eNB from data
    where date > @Since and CID <> 0 and Latitude <> 0.0 and Longitude <> 0.0
    group by (CID >> 8)" -SqlParameters @{Since = $enbsSeenSince } -ErrorAction Stop
    $filterENBs += $enbs | ForEach-Object { $_.eNB }
  }
}

$data = [System.Collections.ArrayList]::new()
if ($null -ne $filterENBs) {
  $filter = "(CID >> 8) in ($($filterENBs -join ','))"
}
else {
  $filter = "CID <> 0 and Latitude <> 0.0 and Longitude <> 0.0"
}
$totalCount = 0
foreach ($db in $dbNames) {
  Write-Host "Reading $db"
  $totalCount += (Invoke-SqliteQuery -Database $db -Query "select count() as count from data").count
  $dbData = Invoke-SqliteQuery -Database $db -Query "select
   (CID >> 8)          as eNB,
   (MCC || '-' || MNC) as MCCMNC,
   -1                  as TimingAdvance,
   -1                  as Band,
   Date,
   CID,
   Latitude,
   Longitude,
   Signal,
   extraData from data
    where $filter
    group by Latitude,Longitude,Altitude,CID
    having min(rowid)
    order by date" -ErrorAction Stop
  if ($dbData -isnot [array]) {
    $dbData = @($dbData)
  }
    $data.AddRange($dbData)
  
  }


Write-Host "Parsing $($data.Count) data points"

foreach ($point in $data) {

  if ($point.extraData -match $bandRegex) {
    $point.Band = [int]$matches.Band
  }
  
  if ($point.extraData -match $taRegex) {
    $point.TimingAdvance = [int]$matches.TA

    if (($point.TimingAdvance % 78) -eq 0) {
      $point.TimingAdvance = $point.TimingAdvance / 78
    }
    elseif (($point.TimingAdvance % 144) -eq 0) {
      $point.TimingAdvance = $point.TimingAdvance / 144
    }
    elseif (($point.TimingAdvance % 150) -eq 0) {
      $point.TimingAdvance = $point.TimingAdvance / 150
    }
    
    $carrierTA = $timingAdvances[$point.MCCMNC]
    if (-not $carrierTA) {
      $carrierTA = $timingAdvances['']
    }
    if ($carrierTA) {
      if (-not $carrierTA[$point.Band]) {
        $point.TimingAdvance = $point.TimingAdvance * $carrierTA[0]
      }
      elseif ( $carrierTA[$point.Band] -is [scriptblock]) {
        $point.TimingAdvance = $carrierTA[$point.Band].InvokeReturnAsIs($point.TimingAdvance)
      }
      else {
        $point.TimingAdvance = $point.TimingAdvance * $carrierTA[$point.Band]
      }
    }
  }
}

Write-Host "Using $($data.Count) points of $totalCount"

$providerGroups = $data | Group-Object -Property MCCMNC

foreach ($provider in $providerGroups) {
  $eNBs = $provider.Group | Group-Object -Property eNB | Sort-Object { [int]$_.Name }

  $resultLocated = @($kmlHeader.Replace('My Places.kml', "$($provider.Name) - Located"))
  $resultCalculated = @($kmlHeader.Replace('My Places.kml', "$($provider.Name) - Calculated"))
  $resultMissing = @()
  $resultLocated += Get-LineStyles
  $resultCalculated += Get-LineStyles

  foreach ($enb in $eNBs) {
    $enbFolder = Get-eNBFolder -group $enb -towers $towers -noCircles:$noCircles -noLines:$noLines -noPoints:$noPoints -noTowers:$noTowers -noToMove:$noToMove -refreshCalculated:$refreshCalculated
    if ($enbFolder.Status -eq 'Located') {
      $resultLocated += $enbFolder.XML
    }
    elseif ($enbFolder.Status -eq 'Calculated') {
      $resultCalculated += $enbFolder.XML
    }
    else {
      $resultMissing += $enbFolder.XML
    }
  }
  $resultCalculated += $resultMissing # The ones missing from CellMapper should go at the bottom.

  $resultLocated += $kmlFooter
  $resultCalculated += $kmlFooter

  $resultLocated = [string]::Join("`r`n", $resultLocated)
  $resultCalculated = [string]::Join("`r`n", $resultCalculated)

  $resultLocated | Out-File "$($provider.Name) - Located$filename.kml"
  Write-Host "Created $($provider.Name) - Located$filename.kml"
  $resultCalculated | Out-File "$($provider.Name) - Calculated$filename.kml"
  Write-Host "Created $($provider.Name) - Calculated$filename.kml"

}

Save-TowerCache

Write-Host "Saved $GLOBAL:requestsSaved requests by getting a list of towers instead of requesting each tower."