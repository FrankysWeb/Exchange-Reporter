$pspkimodule = get-module pspki -ListAvailable
if (!$pspkimodule)
	{
		"PSPKI PowerShell Module not found: Please read Exchange Reportter manual" | add-content "$installpath\ErrorLog.txt"
		exit
	}

try
	{
		import-module pspki -ea 0
	}
Catch
	{
		if ($errorlog -match "yes")
			{
				
				$error[0] | add-content "$installpath\ErrorLog.txt"
			}
	}

$careport = Generate-ReportHeader "careport.png" "$l_ca_header"	
$cells=@("$l_ca_caname","$l_ca_hname","$l_ca_avail","$l_ca_svcstate","$l_ca_type")
$careport += Generate-HTMLTable "$l_ca_castate" $cells

try
{
	$cas = get-ca
}
catch
{
	if ($errorlog -match "yes")
		{
			$error[0] | add-content "$installpath\ErrorLog.txt"
		}
}
foreach ($ca in $cas)
	{
		$caname = $ca.displayname
		$cahostname = $ca.computername
		$caaccess = $ca.IsAccessible
		$caservice = $ca.servicestatus
		$catype = $ca.type
		
		$cells=@("$caname","$cahostname","$caaccess","$caservice","$catype")
		$careport += New-HTMLTableLine $cells
	}
$careport += end-htmltable

#CA Zertifikat
$cells=@("$l_ca_caname","$l_ca_hname","$l_ca_thumbprint","$l_ca_cahash","$l_ca_keyl","$l_ca_validfrom","$l_ca_validto")
$careport += Generate-HTMLTable "$l_ca_t1header" $cells

foreach ($ca in $cas)
	{
		try 
		{
		$caname = $ca.displayname
		$cahostname = $ca.computername
		$rootcerttb = $ca.Certificate.Thumbprint
		$rootcertal = $ca.Certificate.SignatureAlgorithm.FriendlyName
		$rootcertkeysize = $ca.Certificate.PublicKey.Key.keysize
		$rootcertstart = $ca.Certificate.NotBefore | get-date -Format "dd.MM.yyyy HH:mm"
		$rootcertend = $ca.Certificate.NotAfter | get-date -Format "dd.MM.yyyy HH:mm"
		
		$cells=@("$caname","$cahostname","$rootcerttb","$rootcertal","$rootcertkeysize","$rootcertstart","$rootcertend")
		$careport += New-HTMLTableLine $cells
		}
		catch
		{
		}
	}
$careport += end-htmltable

#CA Sperrlisten
$cells=@("$l_ca_caname","$l_ca_hname","$l_ca_crl","$l_ca_lastupdate","$l_ca_nextupdate")
$careport += Generate-HTMLTable "$l_ca_t2header" $cells

foreach ($ca in $cas)
	{
		try
		{
		$caname = $ca.displayname
		$cahostname = $ca.computername
		$crltype = $ca.basecrl.type
		$crllastupdate = $ca.basecrl.thisupdate | get-date -Format "dd.MM.yyyy HH:mm"
		$crlnextupdate = $ca.basecrl.nextupdate | get-date -Format "dd.MM.yyyy HH:mm"
		
		$cells=@("$caname","$cahostname","$crltype ","$crllastupdate","$crlnextupdate")
		$careport += New-HTMLTableLine $cells
		
		$crltype = $ca.deltacrl.type
		$crllastupdate = $ca.deltacrl.thisupdate | get-date -Format "dd.MM.yyyy HH:mm"
		$crlnextupdate = $ca.deltacrl.nextupdate | get-date -Format "dd.MM.yyyy HH:mm"
		
		$cells=@("$caname","$cahostname","$crltype","$crllastupdate","$crlnextupdate")
		$careport += New-HTMLTableLine $cells
		}
		catch
		{
		}
	}
$careport += end-htmltable

#CA ACL
$cells=@("$l_ca_caname","$l_ca_hname","$l_ca_causer","$l_ca_caperm","$l_ca_type")
$careport += Generate-HTMLTable "$l_ca_t3header" $cells

foreach ($ca in $cas)
	{
		$caname = $ca.displayname
		$cahostname = $ca.computername
		$acllist = (Get-CAACL $ca).access
		foreach ($acl in $acllist)
			{
				$acluser = $acl.IdentityReference
				$aclright = $acl.CertificationAuthorityRights
				$acltype = $acl.AccessControlType
				$cells=@("$caname","$cahostname","$acluser","$aclright","$acltype")
				$careport += New-HTMLTableLine $cells
			}
	}
$careport += end-htmltable

#CA Templates
$cells=@("$l_ca_caname","$l_ca_templatename","$l_ca_sversion","$l_ca_tmplversion","$l_ca_autoenrole")
$careport += Generate-HTMLTable "$l_ca_t4header" $cells

foreach ($ca in $cas)
	{
		$caname = $ca.displayname
		$templatelist = (Get-CATemplate $ca).Templates
		foreach ($template in $templatelist)
			{
				$templatename = $template.displayname
				$templateschema = $template.schemaversion
				$templateca = $template.supportedCA
				$templateae = $template.AutoenrollmentAllowed
				$cells=@("$caname","$templatename","$templateschema","$templateca","$templateae ")
				$careport += New-HTMLTableLine $cells
			}
	}
$careport += end-htmltable

#Zertifikate abfragen
#--------------------------------------------------------------------------------------
$issuedcerts = get-ca | Get-IssuedRequest
$CertStates = @() 
foreach ($issuedcert in $issuedcerts)
	{
		$CertData = ($issuedcert | Receive-Certificate).GetRawCertData()
		$TempCert = new-object system.security.cryptography.x509certificates.x509certificate2
		$TempCert.Import($CertData)
		$cn = $TempCert.SubjectName.Name
		$expire = $TempCert.NotAfter | get-date
		$thumbprint = $tempCert.thumbprint
		if ($TempCert.Extensions.oid.value -contains "2.5.29.17")
			{
				$SANs = ($TempCert.Extensions | Where-Object {$_.Oid.value -eq "2.5.29.17"}).format(1)

				if ($SANs -match "DNS-Name")
					{
						$SANs = $SANs.trim().replace("DNS-Name=","").replace("`r`n",", ")
					}
				if ($SANs -match "Prinzipalname")
					{
						$SANs = $SANs.replace("Prinzipalname=","").replace("Anderer Name:","").trim()
					}
				if ($SANs -match "RFC822-Name")
					{
						$SANs = $SANs.replace("RFC822-Name=","").replace("`r`n",", ")
					}
			}
		else
			{
				$SANs = $NULL
			}
		$CertStates += new-object PSObject -property @{CN="$cn";SubjectAlternateNames="$SANs";ExpireDate=$expire;Thumbprint=$thumbprint}
	}

$files += Get-ChildItem "$Installpath\Images\ca_report_cert.png" | Where {-NOT $_.PSIsContainer} | foreach {$_.fullname}
$careport += Generate-ReportHeader "ca_report_cert.png" "$l_ca_header2"	

#Abgelaufene Zertifikate
$cells=@("$l_ca_cn","$l_ca_altname","$l_ca_thumbprint","$l_ca_expiredate","$l_ca_expiretime")
$careport += Generate-HTMLTable "$l_ca_t5header" $cells

$datetoday = get-date
$expiredcerts = $certstates | where {$_.ExpireDate -le $datetoday} | sort expiredate -Descending
foreach ($cert in $expiredcerts)
	{
		try
		{
		$expiredate = $cert.expiredate | get-date -format "dd.MM.yyyy"
		$expiretime = $cert.expiredate | get-date -format "HH:mm"
		$thumbprint = $cert.thumbprint
		$cn = $cert.cn
		$san = $cert.SubjectAlternateNames
		
		$cells=@("$cn","$san","$thumbprint","$expiredate","$expiretime")
		$careport += New-HTMLTableLine $cells	
		}
		catch
		{
		}
	}
$careport += end-htmltable
$expiredcerts = $NULL

#In 30 Tagen ablaufende Zertifikate
$cells=@("$l_ca_cn","$l_ca_altname","$l_ca_thumbprint","$l_ca_expiredate","$l_ca_expiretime")
$careport += Generate-HTMLTable "$l_ca_t6header" $cells

$30days = (get-date).AddDays(+30)
$now = get-date
$expiredcerts = $CertStates | where {$_.expiredate -ge $now -and $_.expiredate -le $30days} | sort expiredate -Descending	
foreach ($cert in $expiredcerts)
	{
		try
		{
		$expiredate = $cert.expiredate | get-date -format "dd.MM.yyyy"
		$expiretime = $cert.expiredate | get-date -format "HH:mm"
		$cn = $cert.cn
		$thumbprint = $cert.thumbprint
		$san = $cert.SubjectAlternateNames
		
		$cells=@("$cn","$san","$thumbprint","$expiredate","$expiretime")
		$careport += New-HTMLTableLine $cells
		}
		catch
		{
		}
	}
$expiredcerts = $NULL
$careport += end-htmltable

$careport | add-content "$tmpdir\report.html"