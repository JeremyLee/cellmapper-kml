[CmdletBinding()]
param (
  $filterENBs = $null,
  [Nullable[datetime]]$enbsSeenSince = $null
)

#.\Get-CellmapperDB.ps1
. .\helpers.ps1

ImportInstall-Module PSSqlite
Add-Type -AssemblyName 'system.drawing'



$timingAdvances = @{
  "310-260" = @{
    0  = 144
    41 = { param($ta) ($ta - 20) * 144 }
  }
  "310-120" = @{
    0  = 144
    41 = { param($ta) ($ta - 20) * 144 }
  }
  ""        = @{
    0 = 144
  }
}

$taRegex = [regex]'&LTE_TA=(?<TA>[0-9]+)&'
$bandRegex = [regex]'&INFO_BAND_NUMBER=(?<Band>[0-9]+)&'
if ($true) {
  Write-Host "Reading DB"

  try {
    
    $rawData = Invoke-SqliteQuery -Database .\cellmapperdata.db -Query "select * from data
    where extraData like '%LTE_TA=%'
    group by Latitude,Longitude,Altitude,CID
    having min(rowid)
    order by date" -ErrorAction Stop
  }
  catch {
    if ((get-command sqlite3).CommandType -ne 'Application') {
      Write-Error "Error loading the database file. Install the sqlite executable on the path, and the script will attempt an autorepair."
      exit
    }
    else {
      sqlite3 cellmapperdata.db ".recover" | sqlite3 recovered.db

      $rawData = Invoke-SqliteQuery -Database .\recovered.db -Query "select * from data
      where extraData like '%LTE_TA=%'
      group by Latitude,Longitude,Altitude,CID
      having min(rowid)
      order by date"
    }
  }

  Write-Host "Filtering data points"
  $data = @()

  if ($null -ne $enbsSeenSince -and $null -eq $filterENBs) {
    $filterENBs = $rawData | Where-Object { $_.Date -gt $enbsSeenSince.Value } | ForEach-Object { $_.CID -shr 8 } | Group-Object | ForEach-Object { $_.Name }
  }

  if ($filterENBs -is [int]) {
    $partiallyFiltered = $rawData | where-object { $filterENBs -eq ($_.CID -shr 8) }
  }
  elseif ($filterENBs -is [array]) {
    $partiallyFiltered = $rawData | where-object { $filterENBs -contains ($_.CID -shr 8) }
  }
  else {
    $partiallyFiltered = $rawData
  }
  
}
foreach ($point in $partiallyFiltered) {
  if ($point.extraData -match $taRegex) {
    # Only use data with a Timing Advance    
    $current = @{
      Date          = $point.Date
      Signal        = $point.Signal
      MCCMNC        = "$($point.MCC)-$($point.MNC)" 
      eNB           = $point.CID -shr 8
      CID           = $point.CID
      Latitude      = $point.latitude
      Longitude     = $point.longitude
      TimingAdvance = [int]$matches.TA
    }

    if (($current.TimingAdvance % 78) -eq 0) {
      $current.TimingAdvance = $current.TimingAdvance / 78
    }
    elseif (($current.TimingAdvance % 144) -eq 0) {
      $current.TimingAdvance = $current.TimingAdvance / 144
    }
    
    if ($point.extraData -match $bandRegex) {
      $current['Band'] = [int]$matches.Band
    }

    $carrierTA = $timingAdvances[$current.MCCMNC]
    if (-not $carrierTA) {
      $carrierTA = $timingAdvances['']
    }
    if ($carrierTA) {
      if (-not $carrierTA[$current.Band]) {
        $current.TimingAdvance = $current.TimingAdvance * $carrierTA[0]
      }
      elseif ( $carrierTA[$current.Band] -is [scriptblock]) {
        $current.TimingAdvance = $carrierTA[$current.Band].InvokeReturnAsIs($current.TimingAdvance)
      }
      else {
        $current.TimingAdvance = $current.TimingAdvance * $carrierTA[$current.Band]
      }
    }

    $data += [pscustomobject]$current
  }
}

Write-Host "Using $($data.Count) points of $($rawData.Count)"

$providerGroups = $data | Group-Object -Property MCCMNC

foreach ($provider in $providerGroups) {
  $eNBs = $provider.Group | Group-Object -Property eNB | Sort-Object { [int]$_.Name }

  $resultLocated = @($kmlHeader.Replace('My Places.kml', "$($provider.Name) - Located"))
  $resultCalculated = @($kmlHeader.Replace('My Places.kml', "$($provider.Name) - Calculated"))
  $resultMissing = @()
  foreach ($enb in $eNBs) {
    $enbFolder = Get-eNBFolder -group $enb -towers $towers
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

  $resultLocated | Out-File "$($provider.Name) - Located.kml"
  Write-Host "Created $($provider.Name) - Located.kml"
  $resultCalculated | Out-File "$($provider.Name) - Calculated.kml"
  Write-Host "Created $($provider.Name) - Calculated.kml"

}

Write-Host "Saved $GLOBAL:requestsSaved requests by getting a list of towers instead of requesting each tower."