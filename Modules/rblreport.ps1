#------------------------------------------------------------------------------
#Benutze Exchange Default Domain um MX Eintrag zu ermitteln:

$RBLUseExchangeDefaultDomain = $True		# $True oder $False

#MX Record wenn $RBLUseExchangeDefaultDomain = $False

$CustomMXRecord = "frankysweb.de"

#Liste der Blacklisten die Aabgefragt werden sollen:

$blacklistServers = @(
    'b.barracudacentral.org'
    'spam.rbl.msrbl.net'
    'zen.spamhaus.org'
    'bl.deadbeef.com'
    'bl.spamcop.net'
    'blackholes.five-ten-sg.com'
    'blacklist.woody.ch'
    'bogons.cymru.com'
    'cbl.abuseat.org'
    'cdl.anti-spam.org.cn'
    'combined.abuse.ch'
    'combined.rbl.msrbl.net'
    'db.wpbl.info'
    'dnsbl-1.uceprotect.net'
    'dnsbl-2.uceprotect.net'
    'dnsbl-3.uceprotect.net'
    'dnsbl.cyberlogic.net'
    'dnsbl.inps.de'
    'dnsbl.njabl.org'
    'dnsbl.sorbs.net'
    'drone.abuse.ch'
    'drone.abuse.ch'
    'duinv.aupads.org'
    'dul.dnsbl.sorbs.net'
    'dul.ru'
    'dyna.spamrats.com'
    'http.dnsbl.sorbs.net'
    'images.rbl.msrbl.net'
    'ips.backscatterer.org'
    'ix.dnsbl.manitu.net'
    'korea.services.net'
    'misc.dnsbl.sorbs.net'
    'noptr.spamrats.com'
    'ohps.dnsbl.net.au'
    'omrs.dnsbl.net.au'
    'orvedb.aupads.org'
    'osps.dnsbl.net.au'
    'osrs.dnsbl.net.au'
    'owfs.dnsbl.net.au'
    'owps.dnsbl.net.au'
    'pbl.spamhaus.org'
    'phishing.rbl.msrbl.net'
    'probes.dnsbl.net.au'
    'proxy.bl.gweep.ca'
    'proxy.block.transip.nl'
    'psbl.surriel.com'
    'rbl.interserver.net'
    'rdts.dnsbl.net.au'
    'relays.bl.gweep.ca'
    'relays.bl.kundenserver.de'
    'relays.nether.net'
    'residential.block.transip.nl'
    'ricn.dnsbl.net.au'
    'rmst.dnsbl.net.au'
    'sbl.spamhaus.org'
    'short.rbl.jp'
    'smtp.dnsbl.sorbs.net'
    'socks.dnsbl.sorbs.net'
    'spam.abuse.ch'
    'spam.dnsbl.sorbs.net'
    'spam.spamrats.com'
    'spamlist.or.kr'
    'spamrbl.imp.ch'
    't3direct.dnsbl.net.au'
    'tor.dnsbl.sectoor.de'
    'torserver.tor.dnsbl.sectoor.de'
    'ubl.lashback.com'
    'ubl.unsubscore.com'
    'virbl.bit.nl'
    'virus.rbl.jp'
    'virus.rbl.msrbl.net'
    'web.dnsbl.sorbs.net'
    'wormrbl.imp.ch'
    'xbl.spamhaus.org'
    'zombie.dnsbl.sorbs.net'
)

#------------------------------------------------------------------------------


$rblreport = Generate-ReportHeader "rblreport.png" "$l_rbl_header"

if ($RBLUseExchangeDefaultDomain -eq $True)
{
	$domainname = (Get-AcceptedDomain | where {$_.Default -eq "True"}).Domainname.Domain
}
else 
{
	$domainname = $CustomMXRecord
}

$IPs = @()
$mxrecords = Resolve-DnsName $domainname -Type MX -DnsOnly -server 8.8.8.8 -ea 0 | where {$_.section -match "Answer"}
foreach ($mxrecord in $mxrecords)
	{
		$MXName = $mxrecord.nameexchange
		$ARecords = Resolve-DnsName $mxname -server 8.8.8.8 | where {$_.section -match "Answer"}
		$IPs += $ARecords.IPAddress
	}


$blacklist = @()
foreach ($IP in $IPs)
{
	$reversedIP = ($IP -split '\.')[3..0] -join '.'
	foreach ($server in $blacklistServers)
	{
    $fqdn = "$reversedIP.$server"

    try
		{
			$null = [System.Net.Dns]::GetHostEntry($fqdn)
			$blacklist += new-object PSObject -property @{Blacklist="$server";IP="$IP";Blacklisted="<b><font color=`"#CD0000`">$l_rbl_yes</font></b>"}
		}
    catch 
		{
			$blacklist += new-object PSObject -property @{Blacklist="$server";IP="$IP";Blacklisted="<font color=`"#008B00`">$l_rbl_no</font>"}
		}
	}
}

$blacklistgroups = $blacklist | group IP
foreach ($blacklistgroup in $blacklistgroups)
	{
		$TestedIP = $blacklistgroup.Name
		$entrys = $blacklistgroup.group | sort Blacklist
		$cells=@("$l_rbl_blname","$l_rbl_ip","$l_rbl_bllist")
		$RBLreport += Generate-HTMLTable "$l_rbl_header2 $TestedIP" $cells
		foreach ($entry in $entrys)
			{
				$blname = $entry.Blacklist
				$blip = $entry.IP
				$bllist = $entry.blacklisted
				$cells=@("$blname","$blip","$bllist")
				$RBLreport += New-HTMLTableLine $cells
			}
		$RBLreport += End-HTMLTable
	}
	
$RBLreport | set-content "$tmpdir\rblreport.html"
$RBLreport | add-content "$tmpdir\report.html"