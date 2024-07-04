$easreport = Generate-ReportHeader "easreport.png" "$l_eas_header"

$cells=@("$l_eas_mbx","$l_eas_id","$l_eas_model","$l_eas_type","$l_eas_displayname","$l_eas_os","$l_eas_firstsync","$l_eas_lastsync	","$l_eas_inactivesince")
$easreport += Generate-HTMLTable "$l_eas_t1header" $cells

try 
{
$timenow = get-date
$easmailboxes = Get-CASMailbox -Resultsize Unlimited -wa 0 -ea 0 | Where {$_.HasActiveSyncDevicePartnership}
if ($easmailboxes)
{
foreach ($easmailbox in $easmailboxes)
	{
		$easdevices = Get-ActiveSyncDeviceStatistics -Mailbox $easmailbox.Identity -ea 0 -wa 0
		if ($easdevices)
		{
			foreach ($easdevice in $easdevices)
				{
		
					$easmbxname = $easmailbox.Name
					$easlastsync = $easdevice.LastSuccessSync
					$easfirstsync = $easdevice.FirstSyncTime		
					$easdeviceid = $easdevice.DeviceID
					$easdevicemodel = $easdevice.DeviceModel
					$easdevicetype = $easdevice.DeviceType
					$easdeviceos = $easdevice.DeviceOS
					$easdevicefrname = $easdevice.DeviceFriendlyName
					

			
					if ($easlastsync -and $easfirstsync -and $timenow)
						{
							$inactivetimestamp = $timenow - $easlastsync
							$daysinactive = $inactivetimestamp.days
		
							$easlastsync = $easlastsync | get-date -format "dd.MM.yyyy HH:mm"
							$easfirstsync = $easfirstsync | get-date -format "dd.MM.yyyy HH:mm"
		
							if ($daysinactive -gt 60)
								{
									[string]$daysinactivestring = $daysinactive
									$daysinactivestring= "$daysinactivestring" + " Tage"
									$cells=@("$easmbxname","$easdeviceid","$easdevicemodel","$easdevicetype","$easdevicefrname","$easdeviceos","$easfirstsync","$easlastsync","$daysinactivestring")
									$easreport += New-HTMLTableLine $cells
								}
							else
								{
									$accells=@("$easmbxname","$easdeviceid","$easdevicemodel","$easdevicetype","$easdevicefrname","$easdeviceos","$easfirstsync","$easlastsync")
									$accreport += New-HTMLTableLine $accells
								}
						}
				}
		}
	}
}	
$easreport += End-HTMLTable
}
catch
{
}

$cells=@("$l_eas_mbx","$l_eas_id","$l_eas_model","$l_eas_type","$l_eas_displayname","$l_eas_os","$l_eas_firstsync","$l_eas_lastsync")
$easreport += Generate-HTMLTable "$l_eas_t2header	" $cells
$easreport += $accreport
$easreport += End-HTMLTable

$cells=@("$l_eas_ios","$l_eas_android","$l_eas_winphone","$l_eas_blackberry","$l_eas_outlookapp","$l_eas_otherversion")
$easreport += Generate-HTMLTable "$l_eas_t3header" $cells

if ($emsversion -match "2010")
	{ 
		$activesyncos = Get-ActiveSyncDevice | where { $_.DistinguishedName -NotLike "*,CN=ExchangeDeviceClasses,*" } | select deviceos
	}
	
if ($emsversion -match "2013")
	{
		$activesyncos = Get-MobileDevice | select deviceos
	}

if ($emsversion -ge "2016")
	{
		$activesyncos = Get-MobileDevice | select deviceos
	}

if ($activesyncos)
	{
		$iOS = ($activesyncos | where {$_.deviceos -like "iOS*"}).count
		$android = ($activesyncos | where {$_.deviceos -like "Android*"}).count
		$wp = ($activesyncos | where {$_.deviceos -like "Windows*"}).count
		$bb = ($activesyncos | where {$_.deviceos -like "Black*"}).count
		$oapp = ($activesyncos | where {$_.deviceos -like "Outlook*"}).count
		$otheros = ($activesyncos | where {$_.deviceos -notlike "iOS*" -and $_.deviceos -notlike "Android*" -and $_.deviceos -notlike "Windows*" -and $_.deviceos -notlike "Black*" -and $_.deviceos -notlike "Outlook*"}).count

		$cells=@("$iOS","$android","$wp","$bb","$oapp","$otheros")
		$easreport += New-HTMLTableLine $cells
		$easreport += End-HTMLTable

		$eascltvalues += [ordered]@{"$l_eas_ios"=$iOS;"$l_eas_android"=$android;"$l_eas_winphone"=$wp;"$l_eas_blackberry"=$bb;"$l_eas_outlookapp"=$oapp;"$l_eas_otherversion"=$otheros}
		new-cylinderchart 500 400 Betriebssystem Name "$l_eas_count" $eascltvalues "$tmpdir\easclients.png"

		$easreport += Include-HTMLInlinePictures "$tmpdir\easclients.png"
	}
else
	{
		$cells=@("$l_eas_nodevice")
		$easreport += New-HTMLTableLine $cells
		$easreport += End-HTMLTable
	}

$easreport | set-content "$tmpdir\easreport.html"
$easreport | add-content "$tmpdir\report.html"