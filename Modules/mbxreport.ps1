$warninglevel = 200
#-----------------------------------------------------------
$mbxreport = Generate-ReportHeader "mbxreport.png" "$l_mbx_header"

$cells=@("$l_mbx_name","$l_mbx_size","$l_mbx_database")
$mbxreport += Generate-HTMLTable "$l_mbx_topmbx ($DisplayTop)" $cells

[array]$exservers = get-exchangeserver

[array]$allmbxstats = @()
[array]$mailboxes = @()
Write-Host "$(Get-Date -Format 'yyyy.MM.dd HH:mm:ss.ffff')    Query MBStats..."
foreach ($exserver in $exservers)
	{
        Start-Job -Name "$($exserver.Name)_MBStats" -ScriptBlock {
            $exserver = $using:exserver
            $SMTPServer = $exserver.FQDN

            Add-Pssnapin *exchange*

            $allmbxstats = Get-MailboxStatistics -Server $SMTPServer -ea 0 -wa 0

            Return $allmbxstats
	    } | Out-Null
	}

While (Get-Job -State "Running")
	{
		Write-Host "$(Get-Date -Format 'yyyy.MM.dd HH:mm:ss.ffff')    Waiting for MBStats $((Get-Job -State "Completed").Count) / $((Get-Job).Count) ..."
        Start-Sleep 15
	}

    $allmbxstats = Get-Job -Name "*MBStats*" | Receive-Job
    Write-Host "$(Get-Date -Format 'yyyy.MM.dd HH:mm:ss.ffff')    Returned $($allmbxstats.Count) MBStats."
    Get-Job | Remove-Job

    $mailboxes = $allmbxstats

$mbxexclude = ($excludelist | where {$_.setting -match "mbxreport"}).value
if ($mbxexclude)
	{
		[array]$mbxexclude = $mbxexclude.split(",")
		foreach ($entry in $mbxexclude) {$mailboxes = $mailboxes.where({$_.displayname -notmatch $entry -or $_.alias -notmatch $entry;})}
		$mailboxes = $mailboxes | sort Totalitemsize -Descending | select -First $DisplayTop
	}
else
	{
		$mailboxes = $mailboxes | sort Totalitemsize -Descending | select -First $DisplayTop
	}

foreach ($mailbox in $mailboxes)
{

 $mbxname = $mailbox.displayname
 $mbxsize = $mailbox.totalitemsize
 $db = $mailbox.database

 $cells=@("$mbxname","$mbxsize","$db")
 $mbxreport += New-HTMLTableLine $cells
 
}
$mbxreport += End-HTMLTable

$databases = get-mailboxdatabase

$mbxlimits = @()
foreach ($database in $databases)
	{
		$mbxdatabase = $database.name
		$dblimit = $database.ProhibitSendQuota
		if ($dblimit -match "Unlimited")
			{
				$dblimitstate = "Inactive"
				$dblimitvalue = "Unlimited"
			}
		else
			{
				$dblimitstate = "Active"
				$dblimitvalue = $dblimit.value.toMB()
			}
		$mailboxesindb = $allmbx.Where({$_.Database -eq $database.Name;}) | sort
		foreach ($mailbox in $mailboxesindb)
			{
				$mbxname = $mailbox.name
				$mbxalias = $mailbox.alias
				$mbxsize = ($allmbxstats.Where({$_.MailboxGuid -eq $mailbox.ExchangeGuid -and $_.Database -eq $mailbox.database;})).TotalItemSize
				If ($mbxsize.count -gt 1) 
					{
						$mbxsize = ($allmbxstats.Where({$_.MailboxGuid -eq $mailbox.ExchangeGuid -and $_.Database -eq $mailbox.database -and $_.DisconnectDate -eq $null;})).TotalItemSize
					}
				if (!$mbxsize -or $Null -eq $mbxsize) 
					{
						$mbxsize = 0
					}
				else
					{
                        try
                            {
						    $mbxsize = $mbxsize.Value.toMB()
                            }
                        catch
                            {
						    $mbxsize = "$mbxsize"
                            }
					}
				$mbxlimit = $mailbox.ProhibitSendQuota
				$mbxdefault = $mailbox.UseDatabaseQuotaDefaults
				if ($mbxlimit -match "Unlimited")
					{
						$mbxlimitstate = "Inactive"
						$mbxlimitvalue = "Unlimited"
					}
				else
					{
						$mbxlimitstate = "Active"
                        try
                            {
						    $mbxlimitvalue = $mbxlimit.Value.toMB()
                            }
                        catch
                            {
						    $mbxlimitvalue = "$mbxlimit"
                            }
					}
			[array]$mbxlimits  += new-object PSObject -property @{Mailbox="$mbxname";DBlimit="$dblimitstate";DBLimitValue="$dblimitvalue";MBXLimit="$mbxlimitstate";MBXLimitValue="$mbxlimitvalue";MBXSize="$mbxsize";MBXAlias="$mbxalias";MBXUseDBDefault="$mbxdefault";Database=$mbxdatabase}
			}
	}

$reportlimits = @()
foreach ($mailbox in $mbxlimits)
	{
		$mbxname = $mailbox.mailbox
		$mbxalias = $mailbox.mbxalias
		[double]$mbxsize = $mailbox.mbxsize
		$mbxdatabase = $mailbox.database
		
		$mbxlimit = $mailbox.mbxlimit
		$dblimit = $mailbox.dblimit
		$mbxusedbdefault = $mailbox.MBXUseDBDefault
		
		#es gilt das Limit der Datenbank
		if ($mbxusedbdefault -eq "True" -and $dblimit -eq "Active")
			{
			    [double]$limitsize = $mailbox.dblimitvalue
				$warningactive = $mbxsize -ge ($limitsize - $warninglevel)
				$limittype = "Database"
				[array]$reportlimits += new-object PSObject -property @{Mailbox="$mbxname";MBXAlias="$mbxalias";LimitType="$Limittype";LimitSize="$limitsize";MailboxSize="$mbxsize";WarningActive="$warningactive";Database=$mbxdatabase}
			}
		#es gilt das Limit des Postfachs
		if ($mbxusedbdefault -eq "False" -and $mbxlimit -eq "Active")
			{
			    [double]$limitsize = $mailbox.MBXLimitValue
				$warningactive = $mbxsize -ge ($limitsize - $warninglevel)
				$limittype = "Mailbox"
				[array]$reportlimits += new-object PSObject -property @{Mailbox="$mbxname";MBXAlias="$mbxalias";LimitType="$Limittype";LimitSize="$limitsize";MailboxSize="$mbxsize";WarningActive="$warningactive";Database=$mbxdatabase}
			}
	}
$reportlimits = $reportlimits.where({$_.WarningActive -match "True";})

$cells=@("$l_mbx_name","$l_mbx_size","$l_mbx_limit","$l_mbx_database","$l_mbx_limittype")
$mbxreport += Generate-HTMLTable "$l_mbx_mbxlimit" $cells
if ($reportlimits)
	{
		foreach ($mbx in $reportlimits)
			{
				$mbxname = $mbx.mailbox
				$mbxsize = $mbx.mailboxsize
				$mbxlimit = $mbx.limitsize
				$mbxdb = $mbx.database
				$limittype = $mbx.limittype
				$cells=@("$mbxname","$mbxsize","$mbxlimit","$mbxdb","$limittype")
				$mbxreport += New-HTMLTableLine $cells
			}
	}
else
	{
		$cells=@("$l_mbx_nolimit")
		$mbxreport += New-HTMLTableLine $cells
	}
		
$mbxreport += End-HTMLTable

#Getrennte Mailboxen

$cells=@("$l_mbx_name","$l_mbx_database","$l_mbx_size","$l_mbx_disconnected","$l_mbx_id")
$mbxreport += Generate-HTMLTable "$l_mbx_dismbx" $cells

$dismbxs = $allmbxstats.Where({ $_.DisconnectDate -ne $null;}) | select displayName,Identity,disconnectdate,database,totalitemsize
foreach ($dismbx in $dismbxs)
	{
		$dismbxname = $dismbx.displayname
		$disdb = $dismbx.database
		$dissize = $dismbx.totalitemsize
		[string]$disdate = $dismbx.disconnectdate | get-date -UFormat %d.%m.%Y
		$disid = $dismbx.Identity
		
		$cells=@("$dismbxname","$disdb ","$dissize","$disdate","$disid")
		$mbxreport += New-HTMLTableLine $cells
	}
$mbxreport += End-HTMLTable


#Inaktive Mailboxen

$cells=@("$l_mbx_name","$l_mbx_database","$l_mbx_size","$l_mbx_lastlogin","$l_mbx_lastloginfrom")
$mbxreport += Generate-HTMLTable "$l_mbx_maybeinactive" $cells

$logonstats = $allmbxstats.where({$_.lastlogontime -lt ((get-date).adddays(-120));}) | select displayname,database,totalitemsize,LastLoggedOnUserAccount,lastlogontime | sort lastlogontime
foreach ($entry in $logonstats)
	{
		$ianame = $entry.displayname
		$iadb = $entry.database
		$iasize = $entry.totalitemsize
		$iall = $entry.lastlogontime
		$iauser = $entry.LastLoggedOnUserAccount
		if (!$iall)
			{
				$iastate = "$l_mbx_userdeactivated"
			}
		if (!$iall -and $iaobj -notmatch "Disabled")
			{
				$iastate = "$l_mbx_unknown"
			}
		if ($iall)
			{
				$iastate = $iall
			} 
		
		$cells=@("$ianame","$iadb","$iasize","$iastate","$iauser")
		$mbxreport += New-HTMLTableLine $cells

	}
$mbxreport += End-HTMLTable

$mbxreport | set-content "$tmpdir\mbxreport.html"
$mbxreport | add-content "$tmpdir\report.html"
