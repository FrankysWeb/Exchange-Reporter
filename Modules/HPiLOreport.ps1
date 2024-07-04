$iloreport = Generate-ReportHeader "hpiloreport.png" "$l_ilo_header"

try
	{
		import-module HPiLOCmdlets -ea 0
	}
Catch
	{
		if ($errorlog -match "yes")
			{
				
				$error[0] | add-content "$installpath\ErrorLog.txt"
			}
	}

$iLOsettingshash = $inifile["HPiLO"]
$iLOsettings = convert-hashtoobject $iLOsettingshash	

$iLOIPs = ($iLOsettings| Where-Object {$_.Setting -eq "iLOIPs"}).Value
$iLOUser = ($iLOsettings| Where-Object {$_.Setting -eq "iLOUser"}).Value
$iLOPass = ($iLOsettings | Where-Object {$_.Setting -eq "iLOPassword"}).Value

[array]$iLOIPs = $iLOIPS.split(",")

$cells=@("$l_ilo_hostname","$l_ilo_ip","$l_ilo_version","$l_ilo_srvtype","$l_ilo_sn","$l_ilo_uptime")
$iloreport += Generate-HTMLTable "$l_ilo_t1header" $cells	

foreach ($iLOIP in $iLOIPs)
	{
		$ilosrvinfo = Get-HPiLOServerInfo -Server $iLOIP -Username $iLOUser -Password $iLOPass
		$uptimemin = (Get-HPiLOPowerOnTime -Server $iLOIP -Username $iLOUser -Password $iLOPass).SERVER_POWER_ON_MINUTES
		$uptimedays = ([timespan]::fromminutes($uptimemin)).Days
		[string]$uptime = "$uptimedays" + " $l_ilo_days"
		
		$srvhostname = $ilosrvinfo.HOSTNAME
		$hpiloip = $ilosrvinfo.IP
		
		$XML = New-Object XML 
		$XML.Load("https://$iLOIP/xmldata?item=All") 
		$ServerType = $XML.RIMP.HSI.SPN
		$SN = $XML.RIMP.HSI.SBSN
		$ILOType = $XML.RIMP.MP.PN
		
		$cells=@("$srvhostname","$hpiloip","$ilotype","$servertype","$sn","$uptime")
		$iloreport += New-HTMLTableLine $cells
		
	}
$iloreport += End-HTMLTable

foreach ($iLOIP in $iLOIPs)
	{
		$ilosrvinfo = Get-HPiLOServerInfo -Server $iLOIP -Username $iLOUser -Password $iLOPass
		$serverpower = Get-HPiLOPowerReading -Server $iLOIP -Username $iLOUser -Password $iLOPass
		$healthsum = Get-HPiLOHealthSummary -Server $iLOIP -Username $iLOUser -Password $iLOPass
		$imllog = (Get-HPiLOIML -Server $iLOIP -Username $iLOUser -Password $iLOPass).EVENT | select -last 10
		
		$biosstate = $healthsum.BIOS_HARDWARE_STATUS
		$fanstate = $healthsum.FANS.STATUS
		$ramstate = $healthsum.MEMORY_STATUS
		$nicstate = $healthsum.NETWORK_STATUS
		$psustate = $healthsum.POWER_SUPPLIES.STATUS
		$storstate = $healthsum.STORAGE_STATUS
		$cpustate =  $healthsum.PROCESSOR_STATUS
		$tempstate = $healthsum.TEMPERATURE_STATUS
		$curtemp = ($ilosrvinfo.TEMP | where {$_.LOCATION -match "Ambient"}).CURRENTREADING
		$powerstate = $serverpower.PRESENT_POWER_READING
		
		$srvhostname = $ilosrvinfo.HOSTNAME
		$firmwares = $ilosrvinfo.FirmwareInfo

		$cells=@("$l_ilo_firmware","$l_ilo_version")
		$iloreport += Generate-HTMLTable "$l_ilo_firmware $srvhostname" $cells
		foreach ($firmware in $firmwares)
			{
				$fwname = $firmware.FIRMWARE_NAME
				$fwversion = $firmware.FIRMWARE_VERSION
				
				$cells=@("$fwname","$fwversion")
				$iloreport += New-HTMLTableLine $cells
			}
		$iloreport += End-HTMLTable
		
		$cells=@("$l_ilo_component","$l_ilo_state")
		$iloreport += Generate-HTMLTable "$l_ilo_hardwarestate $srvhostname" $cells
			$cells=@("$l_ilo_bios","$biosstate")
			$iloreport += New-HTMLTableLine $cells
			$cells=@("$l_ilo_fans","$fanstate")
			$iloreport += New-HTMLTableLine $cells
			$cells=@("$l_ilo_ram","$ramstate")
			$iloreport += New-HTMLTableLine $cells
			$cells=@("$l_ilo_network","$nicstate")
			$iloreport += New-HTMLTableLine $cells			
			$cells=@("$l_ilo_poweravail","$psustate")
			$iloreport += New-HTMLTableLine $cells
			$cells=@("$l_ilo_powercoms","$powerstate")
			$iloreport += New-HTMLTableLine $cells			
			$cells=@("$l_ilo_cpu","$cpustate")
			$iloreport += New-HTMLTableLine $cells
			$cells=@("$l_ilo_storage","$storstate")
			$iloreport += New-HTMLTableLine $cells
			$cells=@("$l_ilo_temp","$tempstate ($curtemp)")
			$iloreport += New-HTMLTableLine $cells

			$iloreport += End-HTMLTable
			
		$cells=@("$l_ilo_date","$l_ilo_message","$l_ilo_type")
		$iloreport += Generate-HTMLTable "$l_ilo_t2header $srvhostname" $cells
		foreach ($event in $imllog)
			{
				if ($event -ne '')
					{
						[datetime]$eventdate = $event.LAST_UPDATE
						$eventdatum = get-date $eventdate -Format "dd.MM.yyyy HH:mm"
						$eventdes = $event.DESCRIPTION
						$eventtyp = $event.SEVERITY

						$cells=@("$eventdatum","$eventdes","$eventtyp")
						$iloreport += New-HTMLTableLine $cells
					}
			}

		$iloreport += End-HTMLTable
	}

$iloreport| set-content "$tmpdir\pfreport.html"
$iloreport| add-content "$tmpdir\report.html"