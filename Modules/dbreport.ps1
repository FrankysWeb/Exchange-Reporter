$dbreport = Generate-ReportHeader "dbreport.png" "$l_db_header"

$cells=@("$l_db_dbname","$l_db_servername","$l_db_size","$l_db_FreeSpace","$l_db_mbxcount","$l_db_lastbackup")
$dbreport += Generate-HTMLTable "$l_db_overview" $cells

$databases = get-mailboxdatabase -status | sort
$currentDate = Get-Date

foreach ($database in $databases)
	{
		$dbname = $database.name
		$dbserver = $database.server
		$dbsize = $database.DatabaseSize
		$dbFreeSpace = $database.AvailableNewMailboxSpace
		$pf = (Get-MailboxDatabase "$database" | get-mailbox -ResultSize Unlimited).count
		$lastbackup = $database.LastFullBackup
		if ($lastbackup)
			{
                #Angepasster Aufruf
				if (($lastbackup - $currentDate).Days -eq 0)
					{
						$lastbackup = "<font color=`"#008B00`">" + (get-date $lastbackup -UFormat "%d.%m.%Y %R") + "</font>"
					}
                elseif (($lastbackup - $currentDate).Days -ge -1) 
					{
						$lastbackup = "<font color=`"#E59400`">" + (get-date $lastbackup -UFormat "%d.%m.%Y %R") + "</font>"
					}
                else 
					{
						$lastbackup = "<font color=`"#CD0000`">" + (get-date $lastbackup -UFormat "%d.%m.%Y %R") + "</font>"
					}
 
			}
		else
			{
				$lastbackup = "<font color=`"#CD0000`">Nie</font>"
			}
 
		$cells=@("$dbname","$dbserver","$dbsize","$dbFreeSpace","$pf","$lastbackup")
		$dbreport += New-HTMLTableLine $cells
 
		$dbsizeges = $dbsizeges + $database.databasesize
		$gespf = $gespf + $pf
 
		$dbsizegb=[double]$dbsize/1024/1024/1024
		$dbvalues += @{$dbname=$dbsizegb}
}

$dbreport += End-HTMLTable

$cells=@("$l_db_dbcount","$l_db_overalldbsize","$l_db_overallmbxcount")
$dbreport += Generate-HTMLTable "$l_db_summary" $cells

$anzdb = $databases.count

$cells=@("$anzdb","$dbsizeges","$gespf")
$dbreport += New-HTMLTableLine $cells
$dbreport += End-HTMLTable

new-cylinderchart 500 400 Datenbanken Name $l_db_size $dbvalues "$tmpdir\dbstat.png"

$dbreport += Include-HTMLInlinePictures "$tmpdir\dbstat.png"

$dbreport | set-content "$tmpdir\dbreport.html"
$dbreport | add-content "$tmpdir\report.html"