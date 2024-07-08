$pfreport = Generate-ReportHeader "pfreport.png" "$l_pf_header"

if ($emsversion -match "2010")
	{
		$cells=@("$l_pf_db","$l_pf_server","$l_pf_size","$l_pf_lastbackup")
		$pfreport += Generate-HTMLTable "$l_pf_header2" $cells
	
		$pfdbs = get-PublicFolderDatabase -Status
		foreach ($pfdb in $pfdbs)
			{
				$pfdbname = $pfdb.name
				$pfdbserver = $pfdb.server
				$pfdbsize = $pfdb.databasesize
				$pflastbackup = $pfdb.LastFullBackup
				if ($pflastbackup)
					{
						$pflastbackup = get-date $pflastbackup -UFormat "%d.%m.%Y %R"
					}
				else
					{
						$pflastbackup = "Nie"
					}
					
				$cells=@("$pfdbname","$pfdbserver","$pfdbsize","$pflastbackup")
				$pfreport += New-HTMLTableLine $cells
			}
		$pfreport += End-HTMLTable
		
		$cells=@("$l_pf_name","$l_pf_db","$l_pf_size","$l_pf_elementcount")
		$pfreport += Generate-HTMLTable "$l_pf_header3" $cells
		
		$pfs = Get-PublicFolderStatistics -resultsize unlimited -ea 0 | sort Totalitemsize -Descending | select -First 200
		foreach ($pf in $pfs)
			{
				$pfname = $pf.AdminDisplayName
				$pfdb = $pf.DatabaseName
				$pfsize = $pf.TotalItemSize
				$pfitemcount = $pf.ItemCount
				
				$cells=@("$pfname","$pfdb","$pfsize","$pfitemcount")
				$pfreport += New-HTMLTableLine $cells
								
			}
		$pfreport += End-HTMLTable
	}

if ($emsversion -match "2013" -or $emsversion -match "2016" -or $emsversion -match "2019")
	{
		$cells=@("$l_pf_mbx","$l_pf_server","$l_pf_size","$l_pf_db")
		$pfreport += Generate-HTMLTable "$l_pf_header4" $cells
		
		$pfmbxs = get-mailbox -PublicFolder
		foreach ($pfmbx in $pfmbxs)
			{
				$pfmbxname = $pfmbx.name
				$pfmbxserver = $pfmbx.servername
				$pfmbxsize = (Get-MailboxStatistics $pfmbx).totalitemsize.value
				$pfmbxdatabase = $pfmbx.database.name
				$cells=@("$pfmbxname","$pfmbxserver","$pfmbxsize","$pfmbxdatabase")
				$pfreport += New-HTMLTableLine $cells
			}
		$pfreport += End-HTMLTable
		
		$cells=@("$l_pf_name","$l_pf_db","$l_pf_size","$l_pf_elementcount")
		$pfreport += Generate-HTMLTable "$l_pf_header3" $cells
		
		$pfs = Get-PublicFolderStatistics -resultsize unlimited -ea 0| sort Totalitemsize -Descending | select -First 200
		foreach ($pf in $pfs)
			{
				$pfname = $pf.name
				$pfid = $pf.EntryId
				$pfdb = (get-publicfolder $pfid).contentmailboxname
				$pfsize = $pf.TotalItemSize
				$pfitemcount = $pf.ItemCount
				
				$cells=@("$pfname","$pfdb","$pfsize","$pfitemcount")
				$pfreport += New-HTMLTableLine $cells
								
			}
		$pfreport += End-HTMLTable
	}

$pfreport | set-content "$tmpdir\pfreport.html"
$pfreport | add-content "$tmpdir\report.html"
