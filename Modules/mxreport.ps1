#------------------------------------------------------------------------------
#Benutze akzeptierte Domains der Exchange Organisation:

$UseExchangeAcceptedDomains = $True		# $True oder $False

#Liste der Domains wenn $UseExchangeAcceptedDomains = $false:

$DomainsToTest = @(
    'frankysweb.de'
    'frankysweb.com'
	'etc.foo'
	)

#------------------------------------------------------------------------------
$DNSServer = "8.8.8.8"

$mxreport = Generate-ReportHeader "mxreport.png" "$l_mx_header"

$mxlist = @()
if ($UseExchangeAcceptedDomains -eq $True)
{
	$maildomains = Get-AcceptedDomain
}
else
{
	$maildomains = $DomainsToTest
}

foreach ($maildomain in $maildomains)
	{
		
		if ($UseExchangeAcceptedDomains -eq $True)
		{
			$domainname = $maildomain.domainname
		}
		else
		{
			$domainname = $maildomain
		}
		
			$mxrecords = Resolve-DnsName $domainname -server $DNSServer -Type MX -DnsOnly -ea 0| where {$_.section -match "Answer"}
            $SPFRecord = Resolve-DnsName $domainname -server $DNSServer -Type TXT -ea 0 | where {$_.strings -match "SPF"}


		if ($SPFRecord)
			{
				$SPF = $SPFRecord.Strings
			}
		else
			{
				$SPF = "$l_mx_notavail"
			}
		
		foreach ($mxrecord in $mxrecords)
			{
				$MXName = $mxrecord.nameexchange
				$ARecords = Resolve-DnsName $mxname -server $DNSServer | where {$_.section -match "Answer"}
				
				foreach ($ARecord in $ARecords)
					{
						$MXIP = $ARecord.IPAddress							
						$MXPTR = Resolve-DnsName $MXIP -server $DNSServer | where {$_.section -match "Answer"}
						$MXPTRName = $MXPTR.NameHost
						
						$smtp = New-Object Net.Sockets.TcpClient
						try 
							{
								$smtp.Connect("$MXIP", 25)
							}
						catch {}
						if($smtp.Connected)
							{
								$SMTPConnect = "$l_mx_success"
								$smtp.Close()
							}
						else
							{
								$SMTPConnect = "$l_mx_nosuccess"
							}
							
						$mxlist += new-object PSObject -property @{Domainname="$domainname";SPF="$SPF";MXRecord="$MXName";MXIPAddress="$MXIP";MXPTRName="$MXPTRName";SMTPConnect="$SMTPConnect";ForRevCheck="$ForRevCheck"}
					}
			}	
	}

$cells=@("$l_mx_domname","$l_mx_mxrec","$l_mx_mxip","$l_mx_mxrevname","$l_mx_spfrec","$l_mx_smtpcheck","$l_mx_forrevtest")
$mxreport += Generate-HTMLTable "$l_mx_mxsettings" $cells

foreach ($mxentry in $mxlist)
	{
		$domainname = $mxentry.Domainname
		$SPF = $mxentry.SPF
		$MXName = $mxentry.MXRecord
		$MXIP = $mxentry.MXIPAddress
		$MXPTRName = $mxentry.MXPTRName
		$SMTPConnect = $mxentry.SMTPConnect
		$ForRevCheck = $mxentry.ForRevCheck
		
		$cells=@("$domainname","$MXName","$MXIP","$MXPTRName","$SPF","$SMTPConnect","$ForRevCheck")
		$mxreport += New-HTMLTableLine $cells		
	}
$mxreport += End-HTMLTable
	
$mxreport | set-content "$tmpdir\mxreport.html"
$mxreport | add-content "$tmpdir\report.html"