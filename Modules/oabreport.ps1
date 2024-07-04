$oabreport = Generate-ReportHeader "oabreport.png" "$l_oab_header"

#OAB Report für Exchange 2010
if ($emsversion -match "2010")
	{
		$cells=@("$l_oab_srv","$l_oab_oabname","$l_oab_sizekb")
		$oabreport += Generate-HTMLTable "$l_oab_oaboverview" $cells

		$OABs = Get-OfflineAddressBook | sort name
		foreach ($OAB in $OABs)
			{
				foreach ($OABVirtualDirectory in $OAB.VirtualDirectories)
					{
						$OABVirtualDirectory = Get-OabVirtualDirectory -Identity $OABVirtualDirectory
						$LocalOABPath = "$($OABVirtualDirectory.Path)\$($OAB.Guid)"
						#$BaseUNCPath = "\\$($OABVirtualDirectory.Server)\C$"
						#$UNCOABPath = $LocalOABPath.Replace("C:",$BaseUNCPath)
						
						$RemoteOABPath = $LocalOABPath.Replace(":\","$\")
						[string]$UNCOABPath = "\\" + $OABVirtualDirectory.Server.Name + "\" + "$RemoteOABPath"
						
						$OABItems = Get-ChildItem -Path $UNCOABPath
						[long]$TotalBytes = ($OABItems | Measure-Object -Property Length -Sum).Sum;
						[long]$TotalKBytes = $TotalBytes/1024;
						$cells=@("$OABVirtualDirectory","$OAB","$TotalKBytes")
						$oabreport += New-HTMLTableLine $cells
	  
						$oabserver = $OABVirtualDirectory.Server.name
						$oabname = $oabserver + " " + $oab.name
						$oabvalues += [ordered]@{$oabname=$TotalKBytes}
					}
			}
		$oabreport += End-HTMLTable

		new-cylinderchart 500 400 $l_oab_oabs $l_oab_name $l_oab_size $oabvalues "$tmpdir\oabstat.png"
		$oabreport += Include-HTMLInlinePictures "$tmpdir\oabstat.png"
	}

#OAB Report für Exchange 2013 / 2016
if ($emsversion -match "2013" -or $emsversion -match "2016")
	{
		$cells=@("$l_oab_oabname","$l_oab_entrycount","$l_oab_lastgerate")
		$oabreport += Generate-HTMLTable "$l_oab_oaboverview" $cells

		$OABs = Get-OfflineAddressBook | sort name
		foreach ($OAB in $OABs)
			{
				$oabname = $oab.name
				$oabrecords = $oab.LastNumberOfRecords
				$oablastgen = $oab.LastTouchedTime
				if ($oablastgen)
					{
						$oablastgen = $oablastgen | get-date -format "dd.MM.yyyy HH:mm"
					}
				else
					{
						$oablastgen = "Nie"
					}
			
				$cells=@("$oabname","$oabrecords ","$oablastgen")
				$oabreport += New-HTMLTableLine $cells

			}
		$oabreport += End-HTMLTable
	
		$cells=@("$l_oab_mbxname","$l_oab_db","$l_oab_oabmbx")
		$oabreport += Generate-HTMLTable "$l_oab_oabmbx" $cells
	
		$OABMailboxes = Get-Mailbox -Arbitration | where {$_.PersistedCapabilities -like "*oab*"}
		foreach ($OABMailbox in $OABMailboxes)
			{
				$oabmailboxname = $oabmailbox.name
				$oabmailboxdb = $oabmailbox.database
				$oabmailboxsize = ($OABMailbox | Get-MailboxStatistics).totalitemsize.value
			
				$cells=@("$oabmailboxname","$oabmailboxdb","$oabmailboxsize")
				$oabreport += New-HTMLTableLine $cells
			}

	$oabreport += End-HTMLTable
	}

$oabreport | set-content "$tmpdir\oabreport.html"
$oabreport | add-content "$tmpdir\report.html"