$dagreport = Generate-ReportHeader "dagreport.png" "$l_dag_header"

$cells=@("$l_dag_db","$l_dag_activeon","$l_dag_activeok","$l_dag_copycount","$l_dag_dbstateactive","$l_dag_dbstaepassive","$l_dag_indexstate")
$dagreport += Generate-HTMLTable $l_dag_dboverview $cells

$databases = get-mailboxdatabase
foreach ($database in $databases)
	{
		$activationpref = $database.ActivationPreference
		$activationpref = $activationpref | where {$_.Value -eq 1}
		$dbshouldonserver = $activationpref.Key
		$dbshouldonserver = $dbshouldonserver.Name
  
		$dbisonserver = $database.Server
  
		$dbonrightserver = $dbisonserver -match $dbshouldonserver
  
		$dbcopycount = $database.ActivationPreference | Measure-Object | ForEach-Object {$_.Count}
		#$dbcopycount = $dbcopycount - 1
  
		$activedbstate = get-mailboxdatabasecopystatus $database -active
		$activedbstate = $activedbstate.Status
  
		$dbstates = get-mailboxdatabasecopystatus $database
  
		#Passive Status
		$passivecopies = $dbstates | where {$_.Status -notmatch "$activedbstate"}
		$passivedbstate = "OK"
		foreach ($copy in $passivecopies)
			{
				$copy = $copy.Status
				if ($copy -notmatch "Healthy")
					{
						$passivedbstate = "Error"
					}
			}

		#Index Status	
		$indexstates = $dbstates
		$dbindexstate = "OK"
		foreach ($state in $indexstates)
			{
				$state = $state.ContentIndexState
				if ($state -match "NotApplicable")
					{
						$dbindexstate = "Not Applicable"
					}
				elseif ($state -notmatch "Healthy") 
					{
						$dbindexstate = "Error"
					}
			}
		
		$cells=@("$database","$dbisonserver","$dbonrightserver","$dbcopycount","$activedbstate","$passivedbstate","$dbindexstate")
		$dagreport += New-HTMLTableLine $cells
  
	}
$dagreport += End-HTMLTable

$cells=@("$l_dag_failserver","$l_dag_date")
$dagreport += Generate-HTMLTable "$l_dag_dagfailover" $cells

foreach ($exserver in $exservers)
	{
		$failoverevents = Get-WinEvent -ComputerName $exserver -FilterHashtable @{logname='system'; id=1135; starttime=$start} -erroraction SilentlyContinue
		if ($failoverevents)
			{
				foreach ($failoverevent in $failoverevents)
					{
						$failovertime = $failoverevent.TimeCreated
						$failovertime = $failovertime | get-date -format "dd.MM.yy HH:mm:ss"
						$failedserver = $failoverevent |  %{([xml]$_.ToXml()).Event.EventData.Data}
						$failedserver = $failedserver.Firstchild
						$failedserver = $failedserver.Value
						$cells=@("$failedserver","$failovertime")
						$dagreport += New-HTMLTableLine $cells
					}
			}
		else
			{
				$cells=@("$l_dag_dagnofailover")
				$dagreport += New-HTMLTableLine $cells
			}
	}
$dagreport += End-HTMLTable

$dagreport | set-content "$tmpdir\dagreport.html"
$dagreport | add-content "$tmpdir\report.html"