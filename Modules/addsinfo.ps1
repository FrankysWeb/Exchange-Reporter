$admodule = get-module -ListAvailable | where {$_.name -match "ActiveDirectory"}
if ($admodule)
	{
		import-module activedirectory
	}
else
	{
		if ($errorlog -match "yes")
			{
				"Fehler: PowerShell ActiveDirectory Modul nicht gefunden" | add-content "$installpath\ErrorLog.txt"
			}
		exit 0
	}
	
$adreport = Generate-ReportHeader "addsinfo.png" "$l_adds_header"

$cells=@("$l_adds_fname","$l_adds_sversion","$l_adds_sname","$l_adds_ffl","$l_adds_sites")
$adreport += Generate-HTMLTable "$l_adds_overview" $cells

$schemaversion = (Get-ADObject (get-adrootdse).schemaNamingContext -Property objectVersion).objectVersion

	if ($schemaversion -eq "13")
		{
			$schemaname = "Windows Server 2000"
		}

	if ($schemaversion -eq "30")
		{
			$schemaname = "Windows Server 2003"
		}
	
	if ($schemaversion -eq "31")
		{
			$schemaname = "Windows Server 2003 R2"
		}	

		if ($schemaversion -eq "44")
		{
			$schemaname = "Windows Server 2008"
		}
	
	if ($schemaversion -eq "47")
		{
			$schemaname = "Windows Server 2008 R2"
		}

	if ($schemaversion -eq "52")
		{
			$schemaname = "Windows Server 2012 Beta"
		}	
	
	if ($schemaversion -eq "56")
		{
			$schemaname = "Windows Server 2012"
		}
	
	if ($schemaversion -eq "69")
		{
			$schemaname = "Windows Server 2012 R2"
		}
	if ($schemaversion -eq "87")
		{
			$schemaname = "Windows Server 2016"
		}	
	if ($schemaversion -ge "88")
		{
			$schemaname = "Windows Server 2019/2022"
		}	

$adforest = get-adforest
$forestmode = $adforest.ForestMode
[string]$forestsites = $adforest.sites
$forestsites = $forestsites.Replace(" ",", ")
$forestname = $adforest.rootdomain

$cells=@("$forestname","$schemaversion","$schemaname","$forestmode","$forestsites")
$adreport += New-HTMLTableLine $cells

$adreport += End-HTMLTable

$cells=@("$l_adds_dname","$l_adds_nbtname","$l_adds_tdomain","$l_adds_dfl","$l_adds_dc")
$adreport += Generate-HTMLTable "$l_adds_adoverview" $cells

$addomains = $adforest.domains
foreach ($addomain in $addomains)
	{
		$adds = get-addomain $addomain
		$domainname = $addomain
		$netbiosname = $adds.netbiosname
		$parentdomain = $adds.parentdomain
		if (!$parentdomain)
			{
				$parentdomain = "$l_adds_nothing"
			}
		$domainmode = $adds.DomainMode
		[string]$adcontrollers = $adds.ReplicaDirectoryServers
		$adcontrollers = $adcontrollers.Replace(" ",", ")

		$cells=@("$domainname","$netbiosname","$parentdomain","$domainmode","$adcontrollers")
		$adreport += New-HTMLTableLine $cells
}

$adreport += End-HTMLTable

$cells=@("$l_adds_name","$l_adds_contributer","$l_adds_model","$l_adds_os","$l_adds_ram","$l_adds_uptime")
$adreport += Generate-HTMLTable "$l_adds_dcoverview" $cells

foreach ($domaincontroller in $domaincontrollers)
	{
		try {
		$computername = $domaincontroller.name
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
		$adreport += New-HTMLTableLine $cells
		}
		catch {
		$cells=@("WMI Error")
		$adreport += New-HTMLTableLine $cells
		}
}

$adreport += end-htmltable


Foreach ($domaincontroller in $domaincontrollers) 
	{
		try {
		$eventsrv = $domaincontroller.name
		$cells=@("$l_adds_source","$l_adds_timestamp","$l_adds_id","$l_adds_count","$l_adds_message")
		$adreport += Generate-HTMLTable "$eventsrv - $l_adds_replerror" $cells
 
		$eventgroups = Get-WinEvent -ComputerName $eventsrv -FilterHashtable @{Logname="*replication*";StartTime = [datetime]$start;level=2,3} -ea 0| select message,id,timecreated | Group-Object id
 
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
						$adreport += New-HTMLTableLine $cells
					}
			}
		else
			{
				$cells=@("$l_adds_noerror")
				$adreport += New-HTMLTableLine $cells
			}
			
		}
		catch {
			$cells=@("WMI Error")
			$adreport += New-HTMLTableLine $cells
		}
		$adreport += end-htmltable
	}

$adreport | set-content "$tmpdir\adreport.html"
$adreport | add-content "$tmpdir\report.html"