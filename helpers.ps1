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

function Get-eNBFolder($group) {
  $tower = Get-Tower -points $group.Group -mcc $group.Group[0].MCCMNC.Split('-')[0] -mnc $group.Group[0].MCCMNC.Split('-')[1] -siteID $group.Group[0].eNB

  if(-not $tower){
    $desc = $group.Name
  } elseif ($tower.towerMover) {
    $desc = $group.Name + " - Located"
  } else {
    $desc = $group.Name + " - Calculated"
  }
  $parts = @(
    "
		<Folder>
			<name>$desc ($($group.Group.Count))</name>
			<open>0</open>
    "
  )

  foreach ($point in $group.Group) {
    $parts += Get-CirclePlacemark $point
  }

  $parts += "</Folder>"

  [string]::Join("`r`n", $parts)
  
}

function Get-BoundsFromPoints($points, [ref]$boundNE, [ref]$boundSW){

  $tempSW = $tempNE = [system.drawing.pointf]::new($points[0].longitude, $points[0].latitude)

  foreach ($point in $points) {
    if($point.longitude -gt $tempNE.X){
      $tempNE.X = $point.longitude
    }
    if($point.longitude -lt $tempSW.X ){
      $tempSW.x = $point.longitude
    }

    if($point.latitude -gt $tempNE.Y){
      $tempNE.Y = $point.latitude
    }

    if($point.latitude -lt $tempSW.Y){
      $tempSW.Y = $point.latitude
    }
  }

  $boundNE.Value = $tempNE
  $boundSW.Value = $tempSW
}

function Get-Towers($mcc, $mnc, [System.Drawing.PointF]$boundNE, [System.Drawing.PointF]$boundSW ){
  $results = Invoke-WebRequest "https://api.cellmapper.net/v6/getTowers?MCC=$mcc&MNC=$mnc&RAT=LTE&boundsNELatitude=$($boundNE.Y)&boundsNELongitude=$($boundNE.X)&boundsSWLatitude=$($boundSW.Y)&boundsSWLongitude=$($boundSW.X)&filterFrequency=false&showOnlyMine=false&showUnverifiedOnly=false&showENDCOnly=false"

  if($results.Content -is [byte[]]){
    Write-Host 'Too many requests to CellMapper. Please go to https://www.cellmapper.net/map and check the captcha when prompted.'
    $z = Read-Host -Prompt "Press Enter to continue"
    $results = Invoke-WebRequest "https://api.cellmapper.net/v6/getTowers?MCC=$mcc&MNC=$mnc&RAT=LTE&boundsNELatitude=$($boundNE.Y)&boundsNELongitude=$($boundNE.X)&boundsSWLatitude=$($boundSW.Y)&boundsSWLongitude=$($boundSW.X)&filterFrequency=false&showOnlyMine=false&showUnverifiedOnly=false&showENDCOnly=false"

  }

  $results = $results.Content | ConvertFrom-Json
  return $results.responseData
}

$towerCache = @{}

function Get-Tower($mcc, $mnc, $siteID, $points){
  if($towerCache[$siteID.ToString()]){
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

  if($towerCache[$siteID.ToString()]){
    return $towerCache[$siteID.ToString()]
  }


  $results = Invoke-WebRequest "https://api.cellmapper.net/v6/getSite?MCC=$mcc&MNC=$mnc&Site=$siteID&RAT=LTE"
  if($results.Content -is [byte[]]){
    Write-Host 'Too many requests to CellMapper. Please go to https://www.cellmapper.net/map and check the captcha when prompted.'
    $z = Read-Host -Prompt "Press Enter to continue"
    $results = Invoke-WebRequest "https://api.cellmapper.net/v6/getSite?MCC=$mcc&MNC=$mnc&Site=$siteID&RAT=LTE"

  }
  $results = $results.Content | ConvertFrom-Json
  return $results.responseData
}

function Get-CirclePlacemark($point) {
  "
  <Placemark>
    <name>$($point.Date.ToString('yyyy-MM-dd HH.mm.ss'))</name>
    <styleUrl>#inline0</styleUrl>
    <LineString>
      <tessellate>1</tessellate>
      <coordinates>
        $((Get-CircleCoordinates -lat $point.Latitude -long $point.Longitude -meter $point.TimingAdvance))
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
	<StyleMap id="m_ylw-pushpin">
		<Pair>
			<key>normal</key>
			<styleUrl>#s_ylw-pushpin</styleUrl>
		</Pair>
		<Pair>
			<key>highlight</key>
			<styleUrl>#s_ylw-pushpin_hl</styleUrl>
		</Pair>
  </StyleMap>'
  
$kmlFooter = '
  </Document>
  </kml>
  '

