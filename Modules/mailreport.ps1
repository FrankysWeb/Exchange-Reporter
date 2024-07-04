$mailreport = Generate-ReportHeader "mailreport.png" "$l_mail_header"

$cells=@("$l_mail_sendcount","$l_mail_reccount","$l_mail_volsend","$l_mail_volrec")
$mailreport += Generate-HTMLTable "$l_mail_header2 $ReportInterval $l_mail_days" $cells

$mailexclude = ($excludelist | where {$_.setting -match "mailreport"}).value
if ($mailexclude)
	{
		[array]$mailexclude = $mailexclude.split(",")
	}

if ($emsversion -match "2016" -or $emsversion -match "2019")
{
 $transportservers = Get-TransportService
 $SendMails = Get-TransportService | Get-MessageTrackingLog -Start $Start -end $End -EventId Send -ea 0 -resultsize unlimited | where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} | select sender,Recipients,timestamp,totalbytes,clienthostname
 $ReceivedMails = Get-TransportService | Get-MessageTrackingLog -Start $Start -end $End -EventId Receive -ea 0 -resultsize unlimited | where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} | select sender,Recipients,timestamp,totalbytes,serverhostname
}

if ($emsversion -match "2013")
{
 $transportservers = Get-TransportService
 $SendMails = Get-TransportService | Get-MessageTrackingLog -Start $Start -end $End -EventId Send -ea 0 -resultsize unlimited | where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} | select sender,Recipients,timestamp,totalbytes,clienthostname
 $ReceivedMails = Get-TransportService | Get-MessageTrackingLog -Start $Start -end $End -EventId Receive -ea 0 -resultsize unlimited | where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} | select sender,Recipients,timestamp,totalbytes,serverhostname
}

if ($emsversion -match "2010")
{
 $transportservers = Get-TransportServer
 $SendMails = Get-TransportServer | Get-MessageTrackingLog -Start $Start -end $End -EventId Send -ea 0 -resultsize unlimited | where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} | select sender,Recipients,timestamp,totalbytes,clienthostname
 $ReceivedMails = Get-TransportServer | Get-MessageTrackingLog -Start $Start -end $End -EventId Receive -ea 0 -resultsize unlimited | where {$_.Recipients -notmatch "HealthMailbox" -and $_.Sender -notmatch "MicrosoftExchange" -and $_.source -match "SMTP"} | select sender,Recipients,timestamp,totalbytes,serverhostname
}

if ($mailexclude)
	{
		foreach ($entry in $mailexclude) {$SendMails = $SendMails | where {$_.sender -notmatch $entry -and $_.recipients -notmatch $entry}}
		foreach ($entry in $mailexclude) {$ReceivedMails = $ReceivedMails | where {$_.sender -notmatch $entry -and $_.recipients -notmatch $entry}}
	}

#Total

$totalsendmail = $sendmails | measure-object Totalbytes -sum
$totalreceivedmail = $receivedmails  | measure-object Totalbytes -sum

$totalsendvol = $totalsendmail.sum
$totalreceivedvol = $totalreceivedmail.sum
$totalsendvol = $totalsendvol / 1024 /1024
$totalreceivedvol = $totalreceivedvol / 1024 /1024
$totalsendvol = [System.Math]::Round($totalsendvol , 2)
$totalreceivedvol  = [System.Math]::Round($totalreceivedvol , 2)

$totalsendcount = $totalsendmail.count
$totalreceivedcount = $totalreceivedmail.count

$totalmail = @{$l_mail_send=$totalsendcount}
$totalmail +=@{$l_mail_received=$totalreceivedcount}

new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_count" $totalmail "$tmpdir\totalmailcount.png"

$totalmail = @{$l_mail_send=$totalsendvol}
$totalmail +=@{$l_mail_received=$totalreceivedvol}

new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_size" $totalmail "$tmpdir\totalmailvol.png"

$cells=@("$totalsendcount","$totalreceivedcount","$totalsendvol","$totalreceivedvol")
$mailreport += New-HTMLTableLine $cells
$mailreport += End-HTMLTable
$mailreport += Include-HTMLInlinePictures "$tmpdir\totalmail*.png"

#Je Server
if ($transportservers.count -gt 1)
{
$cells=@("$l_mail_servername","$l_mail_overallcount","$l_mail_overallvolume","$l_mail_sendcount","$l_mail_reccount","$l_mail_volsend","$l_mail_volrec")
$mailreport += Generate-HTMLTable "$l_mail_header2 $ReportInterval $l_mail_days $l_mail_perserver" $cells

$perserverstats  = @()
foreach ($transportserver in $transportservers)
	{
		$tpsname = $transportserver.name
		$tpssend = $sendmails | where {$_.Clienthostname -match "$tpsname"} | measure-object Totalbytes -sum
		$tpsreceive = $ReceivedMails | where {$_.serverhostname -match "$tpsname"} | measure-object Totalbytes -sum
		$tpssendcount = $tpssend.count
		$tpsreceivecount = $tpsreceive.count
		
		$tpssendvol = $tpssend.sum
		$tpssendvol = $tpssendvol / 1024 / 1024
		$tpssendvol = [System.Math]::Round($tpssendvol , 2)
		$tpsreceivevol = $tpsreceive.sum
		$tpsreceivevol = $tpsreceivevol / 1024 /1024
		$tpsreceivevol = [System.Math]::Round($tpsreceivevol , 2)
		

		$tpstotalvol = $tpsreceivevol + $tpssendvol
		$tpstotalcount = $tpsreceivecount + $tpssendcount
		
		$cells=@("$tpsname","$tpstotalcount","$tpstotalvol","$tpssendcount","$tpsreceivecount","$tpssendvol","$tpsreceivevol")
		$mailreport += New-HTMLTableLine $cells
		
		$perserverstats += new-object PSObject -property @{Name="$tpsname";TotalCount=$tpstotalcount;SendCount=$tpssendcount;ReceiveCount=$tpsreceivecount;ToltalVolume=$tpstotalvol;SendVolume=$tpssendvol;Receivevolume=$tpsreceivevol}
	}
$mailreport += End-HTMLTable

foreach ($tpserver in $perserverstats)
	{
		$tpsname = $tpserver.name
		$tpstotalvol = $tpserver.ToltalVolume
		$tpstotalcount = $tpserver.TotalCount		
		$tpssendvol = $tpserver.SendVolume
		$tpsreceivedvol = $tpserver.Receivevolume
		$tpssendcount = $tpserver.SendCount
		$tpsreceivedcount = $tpserver.ReceiveCount
		
		$tpsvoldata += [ordered]@{$tpsname=$tpstotalvol}
		$tpscountdata += [ordered]@{$tpsname=$tpstotalcount}

		$tpsrscountdata += [ordered]@{"$tpsname $l_mail_send"=$tpssendcount}
		#$tpsrscountdata += @{"$tpsname $l_mail_received"=$tpsreceivedcount}
		
		$tpsrsvoldata += [ordered]@{"$tpsname $l_mail_send"=$tpssendvol}
		#$tpsrsvoldata += @{"$tpsname $l_mail_received"=$tpsreceivedvol}
	
	}
	
foreach ($tpserver in $perserverstats)
	{
		$tpsname = $tpserver.name
		$tpsreceivedvol = $tpserver.Receivevolume
		$tpsreceivedcount = $tpserver.ReceiveCount
		


		#$tpsrscountdata += @{"$tpsname $l_mail_send"=$tpssendcount}
		$tpsrscountdata += [ordered]@{"$tpsname $l_mail_received"=$tpsreceivedcount}
		
		#$tpsrsvoldata += @{"$tpsname $l_mail_send"=$tpssendvol}
		$tpsrsvoldata += [ordered]@{"$tpsname $l_mail_received"=$tpsreceivedvol}
	}
		
new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_size $l_mail_overall" $tpsvoldata "$tmpdir\pertpsvol.png"
new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_count $l_mail_overall" $tpscountdata "$tmpdir\pertpscount.png"
new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_size" $tpsrsvoldata "$tmpdir\pertpsvolrs.png"
new-cylinderchart 500 400 "$l_mail_overallcount" Mails "$l_mail_coun" $tpsrscountdata "$tmpdir\pertpscountrs.png"

$mailreport += Include-HTMLInlinePictures "$tmpdir\pertps*.png"
}
$total += new-object PSObject -property @{Name="$name";Volume=$volume}


#days

$cells=@("$l_mail_date","$l_mail_sendcount","$l_mail_reccount","$l_mail_volsend","$l_mail_volrec")
$mailreport += Generate-HTMLTable "$l_Mail_overviewperday" $cells

$daycounter = 1
do
 {
 $dayendcounter = $daycounter - 1
 $daystart = (Get-Date -Hour 00 -Minute 00 -Second 00).AddDays(-$daycounter)
 $dayend = (Get-Date -Hour 00 -Minute 00 -Second 00).AddDays(-$dayendcounter)
  
  $DayReceivedMails = $ReceivedMails | where {$_.timestamp -ge $daystart -and $_.timestamp -le $dayend}
  $DaySendMails = $sendmails | where {$_.timestamp -ge $daystart -and $_.timestamp -le $dayend}
  
  $daytotalsendmail = $daysendmails | measure-object Totalbytes -sum
  $daytotalreceivedmail = $dayreceivedmails  | measure-object Totalbytes -sum
  
  $daytotalsendvol = $daytotalsendmail.sum
  $daytotalreceivedvol = $daytotalreceivedmail.sum
  $daytotalsendvol = $daytotalsendvol / 1024 /1024
  $daytotalreceivedvol = $daytotalreceivedvol / 1024 /1024
  $daytotalsendvol = [System.Math]::Round($daytotalsendvol , 2)
  $daytotalreceivedvol  = [System.Math]::Round($daytotalreceivedvol , 2)
  
  $daytotalsendcount = $daytotalsendmail.count
  $daytotalreceivedcount = $daytotalreceivedmail.count
  
  $day = $daystart | get-date -Format "dd.MM.yy"
  
  $daystotalmailvol +=[ordered]@{$day=$daytotalreceivedvol}
  $daystotalmailcount +=[ordered]@{$day=$daytotalreceivedcount}
  
  $cells=@("$day","$daytotalsendcount","$daytotalreceivedcount","$daytotalsendvol","$daytotalreceivedvol")
  $mailreport += New-HTMLTableLine $cells
  
 $daycounter++
 }
 while ($daycounter -le $reportinterval)

 new-cylinderchart 500 400 "$l_mail_daycount" Mails "$l_mail_count" $daystotalmailcount "$tmpdir\dailymailcount.png"
 new-cylinderchart 500 400 "$l_mail_daysize" Mails "$l_mail_size" $daystotalmailvol "$tmpdir\dailymailvol.png"
 
$mailreport += End-HTMLTable
$mailreport += Include-HTMLInlinePictures "$tmpdir\dailymail*.png"

$sendstat = $SendMails | select sender,totalbytes
$receivedstat = $receivedMails | select sender,totalbytes

$sendmails = $sendmails.sender
$ReceivedMails = $ReceivedMails.Recipients

$topsenders = $sendmails | Group-Object –noelement | Sort-Object Count -descending | Select-Object -first $DisplayTop
$toprecipients = $ReceivedMails | Group-Object –noelement | Sort-Object Count -descending | Select-Object -first $DisplayTop

$cells=@("$l_mail_sender","$l_mail_count")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_sender ($l_mail_count)" $cells
foreach ($topsender in $topsenders)
{
 $tsname = $topsender.name
 $tscount = $topsender.count
 
 $cells=@("$tsname","$tscount")
 $mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable

$cells=@("$l_mail_recipient","$l_mail_count")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_recipient ($l_mail_count)" $cells
foreach ($toprecipient in $toprecipients)
{
 $trname = $toprecipient.name
 $trcount = $toprecipient.count
 
 $cells=@("$trname","$trcount")
 $mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable

#--------------
#Sender
$cells=@("$l_mail_sender","$l_mail_sizemb")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_sender ($l_mail_size)" $cells

$sendstatgroup = $sendstat | group sender
$total  = @()
foreach ($group in $sendstatgroup)
	{
		$name = ($group.Group | select -first 1).sender
		$volume = ($group.Group | measure totalbytes -Sum).Sum
		$total += new-object PSObject -property @{Name="$name";Volume=$volume}
	}
$toptensendersvol = $total | sort volume -descending | select -first $DisplayTop

foreach ($topsender in $toptensendersvol)
{
 $tsname = $topsender.name
 $tsvolume= $topsender.volume
 $tsvolume = $tsvolume / 1024 /1024
 $tsvolume = [System.Math]::Round($tsvolume , 2)
 $cells=@("$tsname","$tsvolume")
 $mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable

#Recipient
$cells=@("$l_mail_recipient","$l_mail_sizemb")
$mailreport += Generate-HTMLTable "Top $DisplayTop $l_mail_recipient ($l_mail_size)" $cells

$receivedstatgroup = $receivedstat | group sender
$total  = @()
foreach ($group in $receivedstatgroup)
	{
		$name = ($group.Group | select -first 1).sender
		$volume = ($group.Group | measure totalbytes -Sum).Sum
		$total += new-object PSObject -property @{Name="$name";Volume=$volume}
	}
$toptenrecipientsvol = $total | sort volume -descending | select -first $DisplayTop

foreach ($toprecipient in $toptenrecipientsvol)
{
 $trname = $toprecipient.name
 $trvolume = $toprecipient.Volume
 $trvolume = $trvolume / 1024 /1024
 $trvolume = [System.Math]::Round($trvolume , 2)
 $cells=@("$trname","$trvolume")
 $mailreport += New-HTMLTableLine $cells
}
$mailreport += End-HTMLTable

#---------------------------

#Durchschnitt
try
{
$usercount = (get-mailbox -resultsize unlimited | select alias).count

$dsend = $totalsendcount / $usercount
 $dsend = [System.Math]::Round($dsend , 2)
$dreceived = $totalreceivedcount / $usercount
 $dreceived= [System.Math]::Round($dreceived , 2)
$dsendvol = $totalsendvol / $usercount
 $dsendvol= [System.Math]::Round($dsendvol , 2)
$dreceivedvol = $totalreceivedvol / $usercount
 $dreceivedvol = [System.Math]::Round($dreceivedvol  , 2)
$dmailsizesend = $totalsendvol / $totalsendcount
 $dmailsizesend= [System.Math]::Round($dmailsizesend , 2)
$dmailsizereceived = $totalreceivedvol / $totalreceivedcount
 $dmailsizereceived= [System.Math]::Round($dmailsizereceived , 2)

$cells=@("$l_mail_average","$l_mail_value")
$mailreport += Generate-HTMLTable "$l_mail_averagevalue" $cells

 $cells=@("$l_mail_avmbxsendcount","$dsend")
 $mailreport += New-HTMLTableLine $cells
 
 $cells=@("$l_mail_avmbxreccount","$dreceived")
 $mailreport += New-HTMLTableLine $cells
 
 $cells=@("$l_mail_avmbxsendsize","$dsendvol MB")
 $mailreport += New-HTMLTableLine $cells
 
 $cells=@("$l_mail_avmbxrecsize","$dreceivedvol MB")
 $mailreport += New-HTMLTableLine $cells

 $cells=@("$l_mail_avmailsendsize","$dmailsizesend MB")
 $mailreport += New-HTMLTableLine $cells
 
 $cells=@("$l_mail_avmailrecsize","$dmailsizereceived MB")
 $mailreport += New-HTMLTableLine $cells
 
$mailreport += End-HTMLTable
}
catch
{
}
 
$mailreport | set-content "$tmpdir\mailreport.html"
$mailreport | add-content "$tmpdir\report.html"