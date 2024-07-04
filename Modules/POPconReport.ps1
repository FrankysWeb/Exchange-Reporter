$POPconInstallPath = "C:\Program Files (x86)\POPcon"

$popconreport = Generate-ReportHeader "POPconReport.png" "$l_pc_header"

#Übersicht

$cells=@("$l_pc_servicedname","$l_pc_servicename","$l_pc_state")
$popconreport += Generate-HTMLTable "$l_pc_header2" $cells

$pcservice = Get-Service POPcon
$pcprocess = Get-Process POPconSrv

$pcservicedname = $pcservice.DisplayName
$pcservicename = $pcservice.name
$pcstate = $pcservice.status

$cells=@("$pcservicedname","$pcservicename","$pcstate")
$popconreport += New-HTMLTableLine $cells
$popconreport += End-HTMLTable

#Verzeichnisse

$pcbadmailcount = (Get-ChildItem $POPconInstallPath\BADMAIL | where {$_.name -notmatch "was_ist"}).count
$pcpickupcount = (Get-ChildItem $POPconInstallPath\PICKUP | where {$_.name -notmatch "was_ist"}).count
$pctoolargecount = (Get-ChildItem $POPconInstallPath\TOOLARGE | where {$_.name -notmatch "was_ist"}).count

$cells=@("$l_pc_dirName","$l_pc_dirItems")
$popconreport += Generate-HTMLTable "$l_pc_header3" $cells
$cells=@("$l_pc_dirBadmail","$pcbadmailcount")
$popconreport += New-HTMLTableLine $cells
$cells=@("$l_pc_dirPickup","$pcpickupcount")
$popconreport += New-HTMLTableLine $cells
$cells=@("$l_pc_dirToolarge","$pctoolargecount")
$popconreport += New-HTMLTableLine $cells
$popconreport += End-HTMLTable

#Konfiguration

$cells=@("$l_pc_settingname","$l_pc_value")
$popconreport += Generate-HTMLTable "$l_pc_header4" $cells


$pclizenz = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\registration).Name
$pcpostmaster = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\SMTP).Postmaster
$pcserver = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\SMTP).Server

$pcerr1 = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\SMTP).React_on_Exchange_Err1
$pcerr2 = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\SMTP).React_on_Exchange_Err2
$pcerr3 = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\SMTP).React_on_Exchange_Err3

$pcwait = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\Schedule).Schedule_Waittime

$pcwait = switch ($pcwait) 
    { 
        0 {"$l_pc_continusly"}
		1 {"$pcwait $l_pc_minute"}
		default {"$pcwait $l_pc_minutes"}
	}

$pcerr1 = switch ($pcerr1) 
    { 
        0 {"$l_pc_action2"} 
        1 {"$l_pc_action1"}
		2 {"$l_pc_action3"}
		3 {"$l_pc_action4"}
	}

$pcerr2 = switch ($pcerr2) 
    { 
        0 {"$l_pc_action2"} 
        1 {"$l_pc_action1"}
		2 {"$l_pc_action3"}
		3 {"$l_pc_action4"}
	}
	
$pcerr3 = switch ($pcerr3) 
    { 
        0 {"$l_pc_action2"} 
        1 {"$l_pc_action1"}
		2 {"$l_pc_action3"}
		3 {"$l_pc_action4"}
	}

$cells=@("$l_pc_license","$pclizenz")
$popconreport += New-HTMLTableLine $cells

$cells=@("$l_pc_postmaster","$pcpostmaster")
$popconreport += New-HTMLTableLine $cells

$cells=@("$l_pc_exserver","$pcserver")
$popconreport += New-HTMLTableLine $cells

$cells=@("$l_pc_userunknown","$pcerr1")
$popconreport += New-HTMLTableLine $cells

$cells=@("$l_pc_msgrefued","$pcerr2")
$popconreport += New-HTMLTableLine $cells

$cells=@("$l_pc_addrmissing","$pcerr3")
$popconreport += New-HTMLTableLine $cells

$cells=@("$l_pc_wait","$pcwait")
$popconreport += New-HTMLTableLine $cells

$popconreport += End-HTMLTable

#Postfächer

$cells=@("$l_pc_mbxdname","$l_pc_mbxsmtp","$l_pc_mbxuser")
$popconreport += Generate-HTMLTable "$l_pc_header5" $cells

$popmbxcount = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\POP3).nPOP3

$i = 1
do
{
	$mbxdname = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\POP3\$i).displayname
	$mbxsmtp = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\POP3\$i).server
	$mbxuser = (Get-ItemProperty HKLM:\Software\Wow6432Node\POPcon\POP3\$i).username
	
	$cells=@("$mbxdname","$mbxsmtp","$mbxuser")
	$popconreport += New-HTMLTableLine $cells
	
	$i++
}
while ($i -le $popmbxcount)

$popconreport += End-HTMLTable

#Log

$cells=@("$l_pc_logentry")
$popconreport += Generate-HTMLTable "$l_pc_header6" $cells

$pclogfile = get-content $POPconInstallPath\POPconSrv.log | select -last 50
foreach ($pclogentry in $pclogfile)
{
	$pclogentry = $pclogentry.replace("-","")
	$cells=@("$pclogentry")
	$popconreport += New-HTMLTableLine $cells
}
$popconreport += End-HTMLTable

$popconreport | set-content "$tmpdir\pfreport.html"
$popconreport | add-content "$tmpdir\report.html"