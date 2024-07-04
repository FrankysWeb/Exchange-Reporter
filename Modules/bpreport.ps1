$bpreport = Generate-ReportHeader "bpreport.png" "$l_bp_header "

# Übersicht Fehler
$cells=@("$l_bp_srvname","$l_bp_svcname","$l_bp_errorcount")
$bpreport += Generate-HTMLTable "$l_bp_t1header" $cells

$start = (get-date).adddays(-$reportinterval)


foreach ($exserver in $exservers)
 {
  $eventsrv = $exserver.name
  $bperrors = Get-WinEvent -ComputerName $eventsrv -FilterHashtable @{Logname="application";StartTime = [datetime]$start;level=2} -ea 0| where {$_.providername -match "exchange" -and $_.ID -match "1500*"} | select message,id,timecreated,providername
   
  if ($bperrors)
   {
    $bperrorgroups = $bperrors | group providername
    foreach ($bperrorgroup in $bperrorgroups)
     {
      $providername = $bperrorgroup.name
	  $errorcount = $bperrorgroup.count
	  
	  $cells=@("$eventsrv","$providername","$errorcount")
      $bpreport += New-HTMLTableLine $cells
     }
	 
	$bperrormessages = $bperrors | select message -Unique
	foreach ($bperrormessage in $bperrormessages)
	 {
	  $message = $bperrormessage.message
	  $cells=@("$eventsrv","$message")
      $bpdetailreport += New-HTMLTableLine $cells
	 }
   }

 }

$bpreport += End-HTMLTable

# Übersicht Warnungen
$cells=@("$l_bp_srvname","$l_bp_svcname","$l_bp_warncount")
$bpreport += Generate-HTMLTable "$l_bp_t2header" $cells

foreach ($exserver in $exservers)
 {
  $eventsrv = $exserver.name
  $bpwarnings = Get-WinEvent -ComputerName $eventsrv -FilterHashtable @{Logname="application";StartTime = [datetime]$start;level=3} -ea 0| where {$_.providername -match "exchange" -and $_.ID -match "1500*"} | select message,id,timecreated,providername

  if ($bpwarnings)
   {
    $bpwarninggroups = $bpwarnings | group providername
    foreach ($bpwarninggroup in $bpwarninggroups)
     {
      $providername = $bpwarninggroup.name
	  $warningcount = $bpwarninggroup.count
	  
	  $cells=@("$eventsrv","$providername","$warningcount")
      $bpreport += New-HTMLTableLine $cells
     }
	 
    $bpwarningmessages = $bpwarnings | select message -Unique
	foreach ($bpwarningmessage in $bpwarningmessages)
	 {
	  $message = $bpwarningmessage.message
	  $cells=@("$eventsrv","$message")
      $bpdetailreport += New-HTMLTableLine $cells
	 }
   }
 }

$bpreport += End-HTMLTable

#Details
$cells=@("$l_bp_srvname","$l_bp_discription")
$bpreport += Generate-HTMLTable "$l_bp_t3header" $cells

$bpreport += $bpdetailreport

$bpreport += End-HTMLTable

$bpreport | set-content "$tmpdir\serverinfo.html"
$bpreport | add-content "$tmpdir\report.html"