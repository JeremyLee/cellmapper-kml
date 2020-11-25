#.\Get-CellmapperDB.ps1
. .\helpers.ps1

ImportInstall-Module PSSqlite
Add-Type -AssemblyName 'system.drawing'

$rawData = Invoke-SqliteQuery -Database .\cellmapperdata.db -Query 'select * from data'

$data = @()
$taRegex = [regex]'&LTE_TA=(?<TA>[0-9]+)&'
foreach ($point in $rawData) {
  if ($point.extraData.Contains('LTE_TA') -and $point.extraData -match $taRegex) {
    # Only use data with a Timing Advance

    $data += [pscustomobject]@{
      Date          = $point.Date
      MCCMNC        = "$($point.MCC)-$($point.MNC)" 
      eNB           = $point.CID -shr 8
      Latitude      = $point.latitude
      Longitude     = $point.longitude
      TimingAdvance = $matches.TA
    }
  }
}

$providerGroups = $data | Group-Object -Property MCCMNC

foreach ($provider in $providerGroups) {
  $eNBs = $provider.Group | Group-Object -Property eNB | Sort-Object { [int]$_.Name }
  [System.Drawing.PointF]$boundsSW = [System.Drawing.PointF]::Empty
  [System.Drawing.PointF]$boundsNE = [System.Drawing.PointF]::Empty

  #Get-BoundsFromPoints -points $provider.Group -boundNE ([ref]$boundsNE) -boundSW ([ref]$boundsSW)
  $boundsNE.X = $boundsNE.X + 0.1
  $boundsNE.Y = $boundsNE.Y + 0.1
  $boundsSW.X = $boundsSW.X - 0.1
  $boundsSW.Y = $boundsSW.Y - 0.1

  #$towers = Get-Towers -mcc $provider.Group[0].MCCMNC.Split('-')[0] -mnc $provider.Group[0].MCCMNC.Split('-')[1] -boundNE $boundsNE -boundSW $boundsSW

  $result = @($kmlHeader)
  foreach ($enb in $eNBs) {
    $result += Get-eNBFolder -group $enb -towers $towers
  }
  $result += $kmlFooter

  $result = [string]::Join("`r`n", $result)

  $result | Out-File "$($provider.Name).kml"
}