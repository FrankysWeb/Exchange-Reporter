$URL = (Get-OwaVirtualDirectory | select -first 1).ExternalUrl.Authority
#$URL = "www.frankysweb.de"		#Here you can override the Default URL (OWA URL) by a custum URL
$APIEntryPoint = "https://api.ssllabs.com/api/v2"
$SSLReportTimeOut = 10	#Minutes
$ClearCache = "on"
[System.Net.ServicePointManager]::SecurityProtocol = @("Tls12","Tls11","Tls")

#--------------------------------------------------------------
$SSLReport = Generate-ReportHeader "SSLLabsReport.png" "$l_SSLReport_header"

$SSLReportStartTime = get-date
$webrequest =  "$APIEntryPoint" + "/analyze?host=$URL&publish=off&startNew=$ClearCache"

$StartNewReport =  Invoke-WebRequest -Uri $webrequest
$ReportStatus = ($StartNewReport.Content | ConvertFrom-Json).status

do
	{
		start-sleep -seconds 60
		$webrequest =  "$APIEntryPoint" + "/analyze?host=$URL&publish=off"
		$StartNewReport =  Invoke-WebRequest -Uri $webrequest
		$ReportStatus = ($StartNewReport.Content | ConvertFrom-Json).status
		$Timer = ((get-date) - $SSLReportStartTime).Minutes
	}
while ($ReportStatus -notmatch "READY" -and $Timer -le $SSLReportTimeOut)

if ($timer -ge $SSLReportTimeOut)
	{
		if ($errorlog -match "yes")
			{
				
				"Qualys SSL Labs Report Timeout after $SSLReportTimeOut Minutes" | add-content "$installpath\ErrorLog.txt"
				exit 0
			}
	}
else
	{
		$Report = $StartNewReport.Content | ConvertFrom-Json
		
		$TestedServer = $Report.host
		
		$ReportEndpoints =  $Report.Endpoints
		foreach ($ReportEndpoint in $ReportEndpoints)
			{
				
				#Rating
				$EndpointGrade = $ReportEndpoint.grade
				
				#Testduration
				$Testduration = $ReportEndpoint.duration / 1000
				$Testduration = [System.Math]::Round($Testduration , 2)
				
				#Detailed Report
				$EndpointIP = $ReportEndpoint.ipaddress
					$webrequest =  "$APIEntryPoint" + "/getEndpointData?host=$URL&s=$EndpointIP"
					$GetDetailedReport = Invoke-WebRequest -Uri $webrequest
					$DetailedReport = ($GetDetailedReport.content | ConvertFrom-Json).Details
										
					#HTML Report	
					$cells=@("$l_SSLReport_TestedServer","$l_SSLReport_EndPointIP","$l_SSLReport_OverallRating","$l_SSLReport_TestDuration")
					$SSLReport += Generate-HTMLTable "$l_SSLReport_overview $EndpointIP" $cells
				
					$cells=@("$TestedServer","$EndpointIP","$EndpointGrade","$Testduration")
					$SSLReport += New-HTMLTableLine $cells
					$SSLReport += End-HTMLTable
					
					#Supported Protocols
					$cells=@("$l_SSLReport_Name","$l_SSLReport_Version")
					$SSLReport += Generate-HTMLTable "$l_SSLReport_ProtocolHeader" $cells
					
					$Protocols = $DetailedReport.protocols
					foreach ($Protocol in $Protocols)
						{
							$protocolname = $Protocol.name
							$protocolversion = $Protocol.version
							
							$cells=@("$protocolname","$protocolversion")
							$SSLReport += New-HTMLTableLine $cells
						}
					$SSLReport += End-HTMLTable
					
					#Supported Cipher Suites
					$cells=@("$l_SSLReport_Name","$l_SSLReport_CipherStrength")
					$SSLReport += Generate-HTMLTable "$l_SSLReport_SupSuites" $cells
					
					$Suites = $DetailedReport.suites.list
					foreach ($Suite in $Suites)
						{
							$SuiteName = $Suite.name
							$cipherStrength = $suite.cipherStrength
														
							$cells=@("$SuiteName","$cipherStrength")
							$SSLReport += New-HTMLTableLine $cells
						}
					$SSLReport += End-HTMLTable					
					
					#Handshake Simulationen
					$cells=@("$l_SSLReport_Name","$l_SSLReport_Protocol","$l_SSLReport_Suite")
					$SSLReport += Generate-HTMLTable "$l_SSLReport_HSSimulation" $cells
					
					$Simulations = $DetailedReport.sims.results
					foreach ($Simulation in $Simulations)
						{
							$simname = $simulation.client.name + " " + $simulation.client.version
							
							$simresult = $simulation.errorCode
							if ($simresult -eq 0)
								{
									$simprotocol = $simulation.protocolId
									$simsuite = $simulation.suiteId
							
									$simprotocolname = ($DetailedReport.protocols | where {$_.id -match $simprotocol}).name + " " + ($DetailedReport.protocols | where {$_.id -match $simprotocol}).version
									$simsuitename = ($DetailedReport.suites.list | where {$_.id -match $simsuite}).name
									
									$cells=@("$simname","$simprotocolname","$simsuitename")
									$SSLReport += New-HTMLTableLine $cells
								}
							else
								{
									$simprotocolname = "$l_SSLReport_Mismatch"
									$simsuitename = "$l_SSLReport_NA"
									
									$cells=@("$simname","$simprotocolname","$simsuitename")
									$SSLReport += New-HTMLTableLine $cells
								}
						}
					$SSLReport += End-HTMLTable
					
					#Protocol Details
					
					$SecureRenegotiationSupport = switch ($DetailedReport.renegSupport) 
						{ 
							1 {"$l_SSLReport_status_cliinsecure"}
							2 {"$l_SSLReport_status_securesup"}
							4 {"$l_SSLReport_status_clisecure"}
							8 {"$l_SSLReport_status_srvsecure"}
							default {"$l_SSLReport_status_unknown"}
						}
					
					$poodle = switch ($DetailedReport.poodle) 
						{ 
							True {"$l_SSLReport_status_Vulnerable"}
							False {"$l_SSLReport_status_nVulnerable"}
							default {"$l_SSLReport_status_unknown"}
						}
					$poodleTls = switch ($DetailedReport.poodleTls) 
						{ 
							-3 {"$l_SSLReport_status_timeout"}
							-2 {"$l_SSLReport_status_tlsnsupport"}
							-1 {"$l_SSLReport_status_testfailed"}
							0 {"$l_SSLReport_status_unknown"}
							1 {"$l_SSLReport_status_nVulnerable"}
							2 {"$l_SSLReport_status_Vulnerable"}
							default {"$l_SSLReport_status_unknown"}
						}
					$rc4 = switch ($DetailedReport.supportsRc4)
						{ 
							True {"$l_SSLReport_status_rc4sup"}
							False {"$l_SSLReport_status_rc4nup"}
							default {"$l_SSLReport_status_unknown"}
						}					
					
					$openSslCcs = switch ($DetailedReport.openSslCcs)
						{
							-1 {"$l_SSLReport_status_testfailed"}
							0 {"$l_SSLReport_status_unknown"}
							1 {"$l_SSLReport_status_nVulnerable"}
							2 {"$l_SSLReport_status_posvul"}
							3 {"$l_SSLReport_status_vulexp"}
							default {"$l_SSLReport_status_unknown"}		
						}
					
					$forwardSecrecy = switch ($DetailedReport.forwardSecrecy)
						{ 
							1 {"$l_SSLReport_status_one"}
							2 {"$l_SSLReport_status_modern"}
							4 {"$l_SSLReport_status_all"}
							default {"$l_SSLReport_status_unknown"}
						}
					
			}

}

$SSLReport | add-content "$tmpdir\sslreport.html"
$SSLReport | add-content "$tmpdir\report.html"