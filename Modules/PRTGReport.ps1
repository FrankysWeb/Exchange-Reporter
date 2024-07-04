$PRTGreport = Generate-ReportHeader "PRTGReport.png" "$l_prtg_header"

$PRTGsettingshash = $inifile["PRTG-Report"]
$PRTGsettings = convert-hashtoobject $PRTGsettingshash

$prtgserver = ($PRTGsettings| Where-Object {$_.Setting -eq "PRTGServer"}).Value
$prtguser = ($PRTGsettings| Where-Object {$_.Setting -eq "PRTGUser"}).Value
$prtgpassword = ($PRTGsettings | Where-Object {$_.Setting -eq "PRTGPassword"}).Value

$serverlist = ($PRTGsettings | Where-Object {$_.Setting -match "Node"}).Value

$edate = $end | get-date -Format "yyyy-MM-dd-00-00-00"
$sdate = $start | get-date -Format "yyyy-MM-dd-00-00-00"

$sensorlist = (invoke-webrequest -uri "$prtgserver/api/table.csv?content=sensors&columns=device,sensor,objid&username=$prtguser&password=$prtgpassword").content | ConvertFrom-Csv

$serverid = 1
foreach ($server in $serverlist)
{
	$cells =""
	#$PRTGreport += Generate-HTMLTable "$server" $cells
	$serversensors = $sensorlist | where {$_.gerät -match "$server"}
	
	foreach ($serversensor in $serversensors)
	{
		$sensorid = $serversensor.id
		$sensorname = $serversensor.Sensor
		$filename = "$tmpdir\" + "PRTG_" + "$serverid" + "_" + "$SensorID" + ".png"
		#$sensordata = (invoke-webrequest -uri "$prtgserver/api/historicdata.csv?id=$sensorid&avg=86400&sdate=$sdate&edate=$edate&username=$prtguser&password=$prtgpassword").content | ConvertFrom-Csv | select -Property * -ExcludeProperty *RAW* | ConvertTo-Html -Fragment -PreContent "<h3 style=`"text-align:center; font-family:calibri; color: `#0072C6;`">$sensorname</h3>"
		#$PRTGreport += $sensordata
		
		$sensordata = (invoke-webrequest -uri "$prtgserver/api/historicdata.csv?id=$sensorid&avg=86400&sdate=$sdate&edate=$edate&username=$prtguser&password=$prtgpassword").content | ConvertFrom-Csv | select -Property * -ExcludeProperty *RAW*
		
		$cells = $sensordata | Convertto-csv -NoTypeInformation | select -First 1
		[array]$cells = $cells.Replace("`"","").split(",") #"
		$PRTGreport += Generate-HTMLTable "$sensorname" $cells
		
		$cells = $sensordata | Convertto-csv -NoTypeInformation
		$last = $cells.Length
		$cells = $cells[1..$last]
		foreach ($cell in $cells)
			{
				[array]$cell = $cell.Replace("`"","").split(",") #"
				$PRTGreport += New-HTMLTableLine $cell
			}
		
		$img = (invoke-webrequest -uri "$prtgserver/chart.png?id=$sensorid&avg=0&sdate=$sdate&edate=$edate&width=850&height=270&graphstyling=baseFontSize='12'%20showLegend='1'&graphid=-1&username=$prtguser&password=$prtgpassword").content | Set-Content $filename -Encoding Byte
		$PRTGreport += Include-HTMLInlinePictures "$filename"
	}
	
	$serverid++
	
	$PRTGreport += End-HTMLTable
}

$PRTGreport | set-content "$tmpdir\PRTGreport.html"
$PRTGreport | add-content "$tmpdir\report.html"