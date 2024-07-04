$esareport = Generate-ReportHeader "esareport.png" "$l_esa_header"

#Module importieren und Verbindung aufbauen

import-module Posh-SSh

$ESAsettingshash = $inifile["CiscoESA"]
$ESAsettings = convert-hashtoobject $ESAsettingshash	

$ESAIPs = ($ESAsettings| Where-Object {$_.Setting -eq "ESAIPs"}).Value
$ESAUser = ($ESAsettings| Where-Object {$_.Setting -eq "ESAUser"}).Value
$ESAPass = ($ESAsettings | Where-Object {$_.Setting -eq "ESAPassword"}).Value

$ESApasspassword = $ESAPass | ConvertTo-SecureString -AsPlainText -Force
$Creds= New-Object System.Management.Automation.PSCredential -ArgumentList $ESAUser, $ESApasspassword

[array]$ESAIPs = $ESAIPS.split(",")

foreach ($esa in $ESAIPs)
{

$sshsession = New-SSHSession -ComputerName $esa -Credential $creds -AcceptKey:$true

#Status holen

$statusobj = Invoke-SSHCommand -Index 0 -Command "status"

$statusstring = $statusobj.output.Replace(" ","").replace(",","")

[INT]$index = $statusstring.IndexOf("Upsince:") + 8
$Uptime = $statusstring.Substring($index).Split()[0]
[INT]$index = $statusstring.IndexOf("Lastcounterreset:") + 17
$lastcounterreset = $statusstring.Substring($index).Split()[0]
[INT]$index = $statusstring.IndexOf("Systemstatus:") + 13
$systemstatus = $statusstring.Substring($index).Split()[0]
[INT]$index = $statusstring.IndexOf("OldestMessage:") + 14
$oldestmessage = $statusstring.Substring($index).Split()[0]

#Messages.In.Quarantine
$pattern = '(\d+)'
$statusstring = $statusobj.output.Replace(" ",".").replace(",","")
$firststring = "Messages.In.Quarantine"
$secondstring = "Kilobytes.In.Quarantine"
[INT]$firststringlength = $firststring.Length
[INT]$from = $statusstring.IndexOf($firststring) + $firststringlength
[Int]$to = $statusstring.IndexOf($secondstring) - $from
$matches = $statusstring.Substring($from,$to)
$matches = $matches | Select-String -AllMatches $pattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
$current = $matches[0]
 #write-host "$firststring Current:$current"

 $cells=@("$l_esa_systemstate","$l_esa_uptime","$l_esa_lastcounterreset","$l_esa_oldestmail","$l_esa_quranmail")
 $esareport += Generate-HTMLTable "$l_esa_stats $esa" $cells
 $cells=@("$systemstatus","$uptime","$lastcounterreset","$oldestmessage","$current")
 $esareport += New-HTMLTableLine $cells
 $esareport += End-HTMLTable

 $cells=@(" ","$l_esa_counterstats","$l_esa_uptime","$l_esa_lifetime")
 $esareport += Generate-HTMLTable "$l_esa_stats $esa" $cells
 
#Messages Received
$pattern = '(\d+)'
$statusstring = $statusobj.output.Replace(" ",".").replace(",","")
$firststring = "Messages.Received"
$secondstring = "Recipients.Received"
[INT]$firststringlength = $firststring.Length
[INT]$from = $statusstring.IndexOf($firststring) + $firststringlength
[Int]$to = $statusstring.IndexOf($secondstring) - $from
[INT]$firststringlength = $firststring.Length
[INT]$from = $statusstring.IndexOf($firststring) + $firststringlength
[Int]$to = $statusstring.IndexOf($secondstring) - $from
$matches = $statusstring.Substring($from,$to)
$matches = $matches | Select-String -AllMatches $pattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
$reset = $matches[0]
$uptime = $matches[1]
$lifetime = $matches[2]
 #write-host "$firststring Reset:$reset Uptime:$uptime Lifetime:$lifetime"

	$cells=@("M$l_esa_mailreceived","$reset","$uptime","$lifetime")
	$esareport += New-HTMLTableLine $cells


#Recipients Received
$pattern = '(\d+)'
$statusstring = $statusobj.output.Replace(" ",".").replace(",","")
$firststring = "Recipients.Received"
$secondstring = "Rejection"
[INT]$firststringlength = $firststring.Length
[INT]$from = $statusstring.IndexOf($firststring) + $firststringlength
[Int]$to = $statusstring.IndexOf($secondstring) - $from
$matches = $statusstring.Substring($from,$to)
$matches = $matches | Select-String -AllMatches $pattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
$reset = $matches[0]
$uptime = $matches[1]
$lifetime = $matches[2]
 #write-host "$firststring Reset:$reset Uptime:$uptime Lifetime:$lifetime"

	$cells=@("$l_esa_resreceived","$reset","$uptime","$lifetime")
	$esareport += New-HTMLTableLine $cells

#Rejected.Recipients
$pattern = '(\d+)'
$statusstring = $statusobj.output.Replace(" ",".").replace(",","")
$firststring = "Rejected.Recipients"
$secondstring = "Dropped.Messages"
[INT]$firststringlength = $firststring.Length
[INT]$from = $statusstring.IndexOf($firststring) + $firststringlength
[Int]$to = $statusstring.IndexOf($secondstring) - $from
$matches = $statusstring.Substring($from,$to)
$matches = $matches | Select-String -AllMatches $pattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
$reset = $matches[0]
$uptime = $matches[1]
$lifetime = $matches[2]
 #write-host "$firststring Reset:$reset Uptime:$uptime Lifetime:$lifetime"

	$cells=@("$l_esa_rejrec","$reset","$uptime","$lifetime")
	$esareport += New-HTMLTableLine $cells

#Dropped.Messages
$pattern = '(\d+)'
$statusstring = $statusobj.output.Replace(" ",".").replace(",","")
$firststring = "Dropped.Messages"
$secondstring = "Queue"
[INT]$firststringlength = $firststring.Length
[INT]$from = $statusstring.IndexOf($firststring) + $firststringlength
[Int]$to = $statusstring.IndexOf($secondstring) - $from
$matches = $statusstring.Substring($from,$to)
$matches = $matches | Select-String -AllMatches $pattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
$reset = $matches[0]
$uptime = $matches[1]
$lifetime = $matches[2]
 #write-host "$firststring Reset:$reset Uptime:$uptime Lifetime:$lifetime"

	$cells=@("$l_esa_maildrop","$reset","$uptime","$lifetime")
	$esareport += New-HTMLTableLine $cells

#Soft.Bounced.Events
$pattern = '(\d+)'
$statusstring = $statusobj.output.Replace(" ",".").replace(",","")
$firststring = "Soft.Bounced.Events"
$secondstring = "Completion"
[INT]$firststringlength = $firststring.Length
[INT]$from = $statusstring.IndexOf($firststring) + $firststringlength
[Int]$to = $statusstring.IndexOf($secondstring) - $from
$matches = $statusstring.Substring($from,$to)
$matches = $matches | Select-String -AllMatches $pattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
$reset = $matches[0]
$uptime = $matches[1]
$lifetime = $matches[2]
 #write-host "$firststring Reset:$reset Uptime:$uptime Lifetime:$lifetime"

	$cells=@("$l_esa_softbounced","$reset","$uptime","$lifetime")
	$esareport += New-HTMLTableLine $cells

#Completed.Recipients
$pattern = '(\d+)'
$statusstring = $statusobj.output.Replace(" ",".").replace(",","")
$firststring = "Completed.Recipients"
$secondstring = "Current"
[INT]$firststringlength = $firststring.Length
[INT]$from = $statusstring.IndexOf($firststring) + $firststringlength
[Int]$to = $statusstring.IndexOf($secondstring) - $from
$matches = $statusstring.Substring($from,$to)
$matches = $matches | Select-String -AllMatches $pattern | Select-Object -ExpandProperty Matches | Select-Object -ExpandProperty Value
$reset = $matches[0]
$uptime = $matches[1]
$lifetime = $matches[2]
 #write-host "$firststring Reset:$reset Uptime:$uptime Lifetime:$lifetime"

	$cells=@("$l_esa_completed","$reset","$uptime","$lifetime")
	$esareport += New-HTMLTableLine $cells

$esareport += End-HTMLTable

$removessh = Get-SSHSession | Remove-SSHSession

 #write-host "Uptime: $uptime"
 #write-host "Systemstatus: $systemstatus"
 #write-host "Lastcounterreset: $lastcounterreset"
 #write-host "Oldestmessage: $oldestmessage"

}

$esareport | set-content "$tmpdir\esareport.html"
$esareport | add-content "$tmpdir\report.html"
