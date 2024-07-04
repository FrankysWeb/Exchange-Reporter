$exreport = Generate-ReportHeader "serverinfo.png" "$l_serv_header"

$cells=@("$l_serv_name","$l_serv_roles","$l_serv_edition","$l_serv_version")
$exreport += Generate-HTMLTable "$l_serv_header1" $cells

foreach ($exserver in $exservers)
	{
		$name = $exserver.Name
		$roles = $exserver.ServerRole
		$edition = $exserver.edition
		$version = $exserver.AdminDisplayVersion

		$cells=@("$name","$roles","$edition","$version")
		$exreport += New-HTMLTableLine $cells
	}

$exreport += End-HTMLTable

$cells=@("$l_serv_server","$l_serv_validuntil","$l_serv_state","$l_serv_applicant")
$exreport += Generate-HTMLTable "$l_serv_header2" $cells
Foreach ($casserver in $casservers) 
	{
		$servername = $casserver.Name
		$iiscertlist = Get-ExchangeCertificate -Server $servername | Where {$_.Services -match "IIS"}
		Foreach ($iiscert in $iiscertlist) 
			{
				$iiscertdate = $iiscert.NotAfter
				$iiscertdate = get-date $iiscertdate -UFormat "%d.%m.%Y %R"
				$iiscertsubject = $iiscert.Subject
				if ($iiscert.Status -eq "Valid")
					{
						$iiscertstatus = "<font color=`"#008B00`">$l_serv_valid</font>"
					}
				else
					{
						$iiscertstatus = "<font color=`"#CD0000`">$l_serv_invalid</font>"
					}
	
				$cells=@("$servername","$iiscertdate","$iiscertstatus","$iiscertsubject")
				$exreport += New-HTMLTableLine $cells
			}
	}
$exreport += End-HTMLTable

$cells=@("$l_serv_name","$l_serv_manufacturer","$l_serv_model","$l_serv_os","$l_serv_ram","$l_serv_lastboot")
$exreport += Generate-HTMLTable "$l_serv_header3" $cells
foreach ($exserver in $exservers)
	{
		$computername = $exserver.name
		$computerSystem = get-wmiobject Win32_ComputerSystem -ComputerName $computername
		$computerOS = get-wmiobject Win32_OperatingSystem -ComputerName $computername

		$hardware = $computerSystem.Manufacturer
		$model = $computerSystem.Model
		$os = $computerOS.caption + ", SP: " + $computerOS.ServicePackMajorVersion
		$os = $os.replace("Microsoft Windows ","")
		$ram = $computerSystem.TotalPhysicalMemory/1gb
		$ram = [System.Math]::Round($ram, 2)
		$lastboot = $computerOS.ConvertToDateTime($computerOS.LastBootUpTime)
		$lastboot = get-date $lastboot -UFormat "%d.%m.%Y %R" 

		$cells=@("$computername","$hardware","$model","$os","$ram","$lastboot")
		$exreport += New-HTMLTableLine $cells
	}
$exreport += end-htmltable

#Eventlog

Foreach ($exserver in $exservers) 
	{
		$eventsrv = $exserver.name
		$cells=@("$l_serv_source","$l_serv_timestamp","$l_serv_id","$l_serv_count","$l_serv_message")
		$exreport += Generate-HTMLTable "$eventsrv - $l_serv_header4" $cells
 
		$eventgroups = Get-WinEvent -ComputerName $eventsrv -FilterHashtable @{Logname="application";StartTime = [datetime]$start;level=2} -ea 0| where {$_.providername -match "exchange" | select message,id,timecreated} | Group-Object id
 
		if ($eventgroups)
			{
				Foreach ($eventgroup in $eventgroups) 
					{
						$event = $eventgroup.Group | select -first 1
						$eventcount = $eventgroup.count
						$eventsource = $event.providername
						$eventid = $event.id
						$eventtime = $event.TimeCreated
						$eventtime = $eventtime | get-date -format "dd.MM.yy hh:mm:ss"
						$eventmessage = $event.Message
						$eventmeslength = $eventmessage.Length
						if ($eventmeslength -gt 200)
							{
								$eventcontent = $eventmessage.Substring(0,200)
								$eventcontent = $eventcontent + "..."
							}
						else
							{
								$eventcontent = $eventmessage
							}
							
						$cells=@("$eventsource","$eventtime","$eventid","$eventcount","$eventcontent")
						$exreport += New-HTMLTableLine $cells
					}
			}
		else
			{
				$cells=@("$l_serv_noerror")
				$exreport += New-HTMLTableLine $cells
			}
		$exreport += end-htmltable
	}

$exreport | set-content "$tmpdir\serverinfo.html"
$exreport | add-content "$tmpdir\report.html"