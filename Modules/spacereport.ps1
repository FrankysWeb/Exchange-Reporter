$spacereport = Generate-ReportHeader "spacereport.png" "$l_space_header"

$exsrvcount = 1
foreach ($exserver in $exservers)
 {
  $exvolcount = 1
  $computername = $exserver.name
  $cells=@("$l_space_drive","$l_space_name","$l_space_size","$l_space_free")
  $spacereport += Generate-HTMLTable "$computername $l_space_header" $cells
  $volumes = Get-WmiObject win32_volume -computername $computername | where {$_.Drivetype -match "3" -and $_.SystemVolume -match "False" -and $_.capacity -ne 314568704 -and $_.Label -notmatch "Wiederherstellung" -and $_.Label -notmatch "Restore"} | sort caption
  foreach ($volume in $volumes)
   {
    $filename = "$exsrvcount" + "_" + "$exvolcount" + ".png"
    $volsize = [long]($volume.Capacity / 1GB)
	$volsizestring = "$volsize $l_space_GB"
	$volfree = [long]($volume.FreeSpace / 1GB)
	if ($volfree -ge 30) {
		$volfreestring = "<font color=`"#008B00`">$volfree $l_space_GB</font>"
	}
	else {
		$volfreestring = "<font color=`"#CD0000`">$volfree $l_space_GB</font>"
	}
	[long]$volused = $volsize - $volfree
	$volname = $volume.Label
	$volid = $volume.name
	
	$chartdata = @{$l_space_free=$volfree; $l_space_used=$volused}
	new-piechart "150" "150" "$volname $volid" $chartdata "$tmpdir\$filename"
	
	$cells=@($volid,$volname,$volsizestring,$volfreestring)
    $spacereport += New-HTMLTableLine $cells
	
	 $exvolcount = $exvolcount + 1
	 
   } 
  $spacereport += End-HTMLTable
  $spacereport += Include-HTMLInlinePictures "$tmpdir\$exsrvcount*.png"
  
  $exsrvcount = $exsrvcount + 1  
 }

#Domain Controller
 
$dcsrvcount = 1

if ($exserver.name -notmatch $domaincontroller.name) {
foreach ($domaincontroller in $domaincontrollers)
 {
  $dcvolcount = 1
  $computername = $domaincontroller.name
  $cells=@("$l_space_drive","$l_space_name","$l_space_size","$l_space_free")
  $spacereport += Generate-HTMLTable "$computername $l_space_header" $cells
  $volumes = Get-WmiObject win32_volume -computername $computername | where {$_.Drivetype -match "3" -and $_.SystemVolume -match "False" -and $_.capacity -ne 314568704 -and $_.Label -notmatch "Wiederherstellung" -and $_.Label -notmatch "Restore"} | sort caption
  foreach ($volume in $volumes)
   {
    $filename = "dc" + "$dcsrvcount" + "_" + "$dcvolcount" + ".png"
    $volsize = [long]($volume.Capacity/1073741824)
	$volsizestring = "$volsize $l_space_GB"
    $volfree = [long]($volume.FreeSpace/1073741824)
	if ($volfree -ge 20) {
		$volfreestring = "<font color=`"#008B00`">$volfree $l_space_GB</font>"
	}
	else {
		$volfreestring = "<font color=`"#CD0000`">$volfree $l_space_GB</font>"
	}
    [long]$volused = $volsize - $volfree
    $volid = $volume.name
    $volname = $volume.label
    $chartdata = @{$l_space_free=$volfree; $l_space_used=$volused}
   
    new-piechart "150" "150" "$volname $volid" $chartdata "$tmpdir\$filename"
    
    $cells=@($volid,$volname,$volsizestring,$volfreestring)
    $spacereport += New-HTMLTableLine $cells

    $dcvolcount= $dcvolcount + 1
   } 
  $spacereport += End-HTMLTable
  $spacereport += Include-HTMLInlinePictures "$tmpdir\dc$dcsrvcount*.png"
  
  $dcsrvcount = $dcsrvcount + 1  
 }
}

$spacereport | set-content "$tmpdir\spacereport.html"
$spacereport | add-content "$tmpdir\report.html"

