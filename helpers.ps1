function ImportInstall-Module ($moduleName) {
  if (-not (Get-Module $moduleName)) {
    if (-not (Get-Module -ListAvailable $moduleName)) {
      Install-Module $moduleName -Scope CurrentUser -Force
    }
    Import-Module $moduleName
  }
}

function rad2deg ($angle) {
  return $angle * (180 / [Math]::PI);
}
function deg2rad ($angle) {
  return $angle * ([Math]::PI / 180);
}

function Get-Circlecoordinates($lat, $long, $meter) {
  # convert coordinates to radians
  $lat1 = deg2rad($lat);
  $long1 = deg2rad($long);
  $d_rad = $meter / 6378137;
 
  $coordinatesList = @();
  # loop through the array and write path linestrings
  for ($i = 0; $i -le 360; $i += 9) {
    $radial = deg2rad($i);
    $lat_rad = [math]::asin([math]::sin($lat1) * [math]::cos($d_rad) + [math]::cos($lat1) * [math]::sin($d_rad) * [math]::cos($radial));
    $dlon_rad = [math]::atan2([math]::sin($radial) * [math]::sin($d_rad) * [math]::cos($lat1), [math]::cos($d_rad) - [math]::sin($lat1) * [math]::sin($lat_rad));
    $lon_rad = (($long1 + $dlon_rad + [math]::PI) % (2 * [math]::PI)) - [math]::PI;
    $coordinatesList += "$(rad2deg($lon_rad)),$(rad2deg($lat_rad)),0"
  }
  return [string]::join(' ', $coordinatesList)
}

$lastDigitRegex = [regex]'\d(?= - )'

function Get-eNBFolder($group) {
  $tower = Get-Tower -points $group.Group -mcc $group.Group[0].MCCMNC.Split('-')[0] -mnc $group.Group[0].MCCMNC.Split('-')[1] -siteID $group.Group[0].eNB
  $towerStatus = $null
  if (-not $tower) {
    $towerStatus = 'Missing'
    $desc = $group.Name
  }
  elseif ($tower.towerMover) {
    $towerStatus = 'Located'
    $desc = $group.Name + " - Located"
  }
  else {
    $towerStatus = 'Calculated'
    $desc = $group.Name + " - Calculated"
  }
  $parts = @(
    "
		<Folder>
			<name>$desc ($($group.Group.Count))</name>
			<open>0</open>
    "
  )
  if ($tower) {
    $parts += "
  <Placemark>
    <name>CellMapper Location - $towerStatus</name>
    <visibility>0</visibility>
    <styleUrl>#m_ylw-pushpin</styleUrl>
    <Point>
      <gx:drawOrder>1</gx:drawOrder>
      <coordinates>$($tower.longitude),$($tower.latitude),0</coordinates>
    </Point>
  </Placemark>"
    $parts += "
  <Placemark>
    <name>Move This - $($group.Name)</name>
    <visibility>0</visibility>
    <styleUrl>#m_ylw-pushpin</styleUrl>
    <Point>
      <gx:drawOrder>1</gx:drawOrder>
      <coordinates>$($tower.longitude),$($tower.latitude + 0.0005),0</coordinates>
    </Point>
  </Placemark>"
  }
  else {
    $bounds = GetBoundsFromPoints $group.Group
    $bounds = [System.Drawing.PointF]::new((($bounds[0].X + $bounds[1].X) / 2), (($bounds[0].Y + $bounds[1].Y) / 2))
    $parts += "
  <Placemark>
    <name>Move This - $($group.Name)</name>
    <visibility>0</visibility>
    <styleUrl>#m_ylw-pushpin</styleUrl>
    <Point>
      <gx:drawOrder>1</gx:drawOrder>
      <coordinates>$($bounds.X),$($bounds.Y),0</coordinates>
    </Point>
  </Placemark>"
  }

  $circles = @()
  $pointFolders = @{}
  $signals = $group.Group | ForEach-Object { $_.Signal } | Sort-Object -Descending
  $ratios = @(
    (($signals[0] * 3 + $signals[$signals.count - 1]) / 4),
    (($signals[0] + $signals[$signals.count - 1]) / 2),
    (($signals[0] + $signals[$signals.count - 1] * 3) / 4))
  
  foreach ($point in $group.Group) {
    $folderName = Get-PointFolderName $point
    if (-not $pointFolders[$folderName]) {
      $pointFolders[$folderName] = @()
    }
    $pointFolders[$folderName] += Get-PointPlacemark -point $point -signalRatios $ratios
    $circles += Get-CirclePlacemark $point
  }

  $parts += "
  <Folder>
    <name>Points</name>
    <open>0</open>  
  "
  $sortedPointFolders = $pointFolders.Keys | Sort-object { $_[$_.IndexOf(' ') - 1] }, { $_ }
  foreach ($pf in $sortedPointFolders) {
    $parts += "
    <Folder>
      <name>$pf</name>
      <open>0</open>  
    "
    $parts += $pointFolders[$pf]
    $parts += "</Folder>"
  }
  $parts += "</Folder>"

  
  $parts += "
  <Folder>
    <name>Circles</name>
    <open>0</open>  
  "
  $parts += $circles
  $parts += "</Folder>"
  
  $parts += "</Folder>"
  
  [pscustomobject]@{
    XML    = [string]::Join("`r`n", $parts)
    Status = $towerStatus
  }
}

function Get-PointPlacemark($point, $signalRatios = @(-84, -102, -111)) {
  $style = "#m_grn-dot"
  
  if ($point.signal -le $signalRatios[2]) {
    $style = "#m_red-dot"
  }
  elseif ($point.signal -lt $signalRatios[1]) {
    $style = "#m_org-dot"
  }
  elseif ($point.signal -lt $signalRatios[0]) {
    $style = "#m_ylw-dot"
  }

  "
<Placemark>
  <name>$($point.Signal)</name>
  <visibility>0</visibility>
  <description>$($point.Date.ToString('yyyy-MM-dd HH.mm.ss'))
RSRP:	$($point.Signal)
TA:	$($point.TimingAdvance)</description>
  <styleUrl>$style</styleUrl>
  <Point>
    <gx:drawOrder>1</gx:drawOrder>
    <coordinates>$($point.longitude),$($point.latitude),0</coordinates>
  </Point>
</Placemark>"
}

function Get-PointFolderName($point) {
  $sectorNumber = $point.CID -band 0xFF # Keep last byte
  
  "$sectorNumber - Band $($point.Band)"
}

function Get-BoundsFromPoints($points, [ref]$boundNE, [ref]$boundSW) {
  $val = GetBoundsFromPoints $points
  
  $boundNE.Value = $val[0]
  $boundSW.Value = $val[1]
}
function GetBoundsFromPoints($points) {

  $tempSW = $tempNE = [system.drawing.pointf]::new($points[0].longitude, $points[0].latitude)

  foreach ($point in $points) {
    if ($point.longitude -gt $tempNE.X) {
      $tempNE.X = $point.longitude
    }
    if ($point.longitude -lt $tempSW.X ) {
      $tempSW.x = $point.longitude
    }

    if ($point.latitude -gt $tempNE.Y) {
      $tempNE.Y = $point.latitude
    }

    if ($point.latitude -lt $tempSW.Y) {
      $tempSW.Y = $point.latitude
    }
  }
  
  @($tempNE, $tempSW)
}

$global:lastRequest = [datetime]::MinValue

$delayMS = 50
$useragent = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.66 Safari/537.36'


function Request($url, [switch]$reset) {
  $timeSinceLastRequest = ([datetime]::Now - $global:lastRequest).TotalMilliseconds

  if ($timeSinceLastRequest -lt $delayMS) {
    Write-Host "Sleeping $($delayMS - $timeSinceLastRequest)ms"
    Start-Sleep -Milliseconds($delayMS - $timeSinceLastRequest)
  }

  if ($global:session -and -not $reset) {
    Invoke-WebRequest -uri $url -WebSession $global:session -UserAgent $useragent
  }
  else {
    Invoke-WebRequest -uri $url -SessionVariable 'tempsession' -UserAgent $useragent
    $global:session = $tempsession
    if ($reset) {
      Write-Host 'Creating New Session'
    }
  }
  $global:lastRequest = [datetime]::Now
}

function Get-Towers($mcc, $mnc, [System.Drawing.PointF]$boundNE, [System.Drawing.PointF]$boundSW ) {
  $savedProgress = $ProgressPreference
  $ProgressPreference = 'SilentlyContinue'
  $results = Request "https://api.cellmapper.net/v6/getTowers?MCC=$mcc&MNC=$mnc&RAT=LTE&boundsNELatitude=$($boundNE.Y)&boundsNELongitude=$($boundNE.X)&boundsSWLatitude=$($boundSW.Y)&boundsSWLongitude=$($boundSW.X)&filterFrequency=false&showOnlyMine=false&showUnverifiedOnly=false&showENDCOnly=false"

  if ($results.Content -is [byte[]]) {
    $results = Request -reset "https://api.cellmapper.net/v6/getTowers?MCC=$mcc&MNC=$mnc&RAT=LTE&boundsNELatitude=$($boundNE.Y)&boundsNELongitude=$($boundNE.X)&boundsSWLatitude=$($boundSW.Y)&boundsSWLongitude=$($boundSW.X)&filterFrequency=false&showOnlyMine=false&showUnverifiedOnly=false&showENDCOnly=false"
  }

  $results = $results.Content | ConvertFrom-Json
  $ProgressPreference = $savedProgress
  return $results.responseData
}

$towerCache = @{}
$GLOBAL:requestsSaved = 0

function Get-Tower($mcc, $mnc, $siteID, $points) {
  if ($towerCache[$siteID.ToString()]) {
    $GLOBAL:requestsSaved += 1
    return $towerCache[$siteID.ToString()]
  }

  [System.Drawing.PointF]$boundsSW = [System.Drawing.PointF]::Empty
  [System.Drawing.PointF]$boundsNE = [System.Drawing.PointF]::Empty

  Get-BoundsFromPoints -points $points -boundNE ([ref]$boundsNE) -boundSW ([ref]$boundsSW)
  $boundsNE.X = $boundsNE.X + 0.04
  $boundsNE.Y = $boundsNE.Y + 0.04
  $boundsSW.X = $boundsSW.X - 0.04
  $boundsSW.Y = $boundsSW.Y - 0.04

  $towers = Get-Towers -mcc $mcc -mnc $mnc -boundNE $boundsNE -boundSW $boundsSW

  foreach ($tower in $towers) {
    $towerCache[$tower.siteID.ToString()] = $tower
  }

  if ($towerCache[$siteID.ToString()]) {
    return $towerCache[$siteID.ToString()]
  }


  $results = Request "https://api.cellmapper.net/v6/getSite?MCC=$mcc&MNC=$mnc&Site=$siteID&RAT=LTE"
  if ($results.Content -is [byte[]]) {
    $results = Request -reset "https://api.cellmapper.net/v6/getSite?MCC=$mcc&MNC=$mnc&Site=$siteID&RAT=LTE"

  }
  
  $results = $results.Content | ConvertFrom-Json
  return $results.responseData
}

function Get-CirclePlacemark($point) {
  $ta = $point.TimingAdvance
  if ($ta -eq 0) {
    $ta = 5
  }
  "
  <Placemark>
    <name>$($point.Date.ToString('yyyy-MM-dd HH.mm.ss'))</name>
    <visibility>0</visibility>
    <styleUrl>#inline0</styleUrl>
    <LineString>
      <tessellate>1</tessellate>
      <coordinates>
        $((Get-CircleCoordinates -lat $point.Latitude -long $point.Longitude -meter $ta))
      </coordinates>
    </LineString>
  </Placemark>"
}

$kmlHeader = '<?xml version="1.0" encoding="UTF-8"?>
<kml xmlns="http://www.opengis.net/kml/2.2" xmlns:gx="http://www.google.com/kml/ext/2.2" xmlns:kml="http://www.opengis.net/kml/2.2" xmlns:atom="http://www.w3.org/2005/Atom">
<Document>
	<name>My Places.kml</name>
	<Style id="inline0">
		<LineStyle>
			<color>ff0000ff</color>
			<width>2</width>
		</LineStyle>
	</Style>
	<Style id="s_ylw-pushpin">
		<IconStyle>
			<scale>1.1</scale>
			<Icon>
				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>
			</Icon>
			<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>
		</IconStyle>
	</Style>
	<Style id="s_ylw-pushpin_hl">
		<IconStyle>
			<scale>1.3</scale>
			<Icon>
				<href>http://maps.google.com/mapfiles/kml/pushpin/ylw-pushpin.png</href>
			</Icon>
			<hotSpot x="20" y="2" xunits="pixels" yunits="pixels"/>
		</IconStyle>
	</Style>
	<StyleMap id="m_ylw-pushpin">
		<Pair>
			<key>normal</key>
			<styleUrl>#s_ylw-pushpin</styleUrl>
		</Pair>
		<Pair>
			<key>highlight</key>
			<styleUrl>#s_ylw-pushpin_hl</styleUrl>
		</Pair>
  </StyleMap>
  
  <Style id="sn_grn-dot">
    <IconStyle>
      <color>ff00FF21</color>
      <scale>1.2</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <Style id="sh_grn-dot">
    <IconStyle>
      <color>ff00FF21</color>
      <scale>1.4</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <StyleMap id="m_grn-dot">
    <Pair>
      <key>normal</key>
      <styleUrl>#sn_grn-dot</styleUrl>
    </Pair>
    <Pair>
      <key>highlight</key>
      <styleUrl>#sh_grn-dot</styleUrl>
    </Pair>
  </StyleMap>
  
  <Style id="sn_red-dot">
    <IconStyle>
      <color>ff00007f</color>
      <scale>1.2</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <Style id="sh_red-dot">
    <IconStyle>
      <color>ff00007f</color>
      <scale>1.4</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <StyleMap id="m_red-dot">
    <Pair>
      <key>normal</key>
      <styleUrl>#sn_red-dot</styleUrl>
    </Pair>
    <Pair>
      <key>highlight</key>
      <styleUrl>#sh_red-dot</styleUrl>
    </Pair>
  </StyleMap>
  
  <Style id="sn_org-dot">
    <IconStyle>
      <color>ff00aaff</color>
      <scale>1.2</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <Style id="sh_org-dot">
    <IconStyle>
      <color>ff00aaff</color>
      <scale>1.4</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <StyleMap id="m_org-dot">
    <Pair>
      <key>normal</key>
      <styleUrl>#sn_org-dot</styleUrl>
    </Pair>
    <Pair>
      <key>highlight</key>
      <styleUrl>#sh_org-dot</styleUrl>
    </Pair>
  </StyleMap>
  
  <Style id="sn_ylw-dot">
    <IconStyle>
      <color>ff00ffff</color>
      <scale>1.2</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <Style id="sh_ylw-dot">
    <IconStyle>
      <color>ff00ffff</color>
      <scale>1.4</scale>
      <Icon>
        <href>http://maps.google.com/mapfiles/kml/shapes/shaded_dot.png</href>
      </Icon>
    </IconStyle>
  </Style>
  <StyleMap id="m_ylw-dot">
    <Pair>
      <key>normal</key>
      <styleUrl>#sn_ylw-dot</styleUrl>
    </Pair>
    <Pair>
      <key>highlight</key>
      <styleUrl>#sh_ylw-dot</styleUrl>
    </Pair>
  </StyleMap>'
  
$kmlFooter = '
  </Document>
  </kml>
  '

