$healthreport = Generate-ReportHeader "HealthChecker.png" "$l_health_header"

#Download HealthChecker
invoke-webrequest -Uri "https://github.com/microsoft/CSS-Exchange/releases/latest/download/HealthChecker.ps1" -outfile "$installpath\Modules\3rdParty\HealthChecker.ps1"

#Run HealthChecker Script
$ExecuteHealthChecker = Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | %{. "$installpath\Modules\3rdParty\HealthChecker.ps1" -Server $_.Name -OutputFilePath $tmpdir} | Remove-WriteConsole

#Get HealthChecker XMLs
$HealthCheckerXMLs = Get-ChildItem "$tmpdir\HealthCheck*.xml" | foreach {$_.fullname}

#Import XMLs and format Output
foreach ($HealthCheckerXML in $HealthCheckerXMLs) {
	$HealthCheckerXML = Import-Clixml "$HealthCheckerXML"

	$HCServername = $HealthCheckerXML.HealthCheckerExchangeServer.ServerName

	$cells=@("$l_health_name","$l_health_message")
	$healthreport += Generate-HTMLTable "$l_health_header2 $HCServername" $cells

	$ServerDetails = $HealthCheckerXML.HtmlServerValues.ServerDetails
	foreach ($line in $ServerDetails) {
		$LineName = $line.name
		$LineClass = $line.class
		$LineValue = $line.DetailValue
		
		if ($LineClass -match "Red") {
			$LineValueStr = "<font color=`"#CD0000`">$LineValue</font>"
			}
		elseif ($LineClass -match "Yellow") {
			$LineValueStr = "<font color=`"#E59400`">$LineValue</font>"
			}
		elseif ($LineClass -match "Green") {
			$LineValueStr = "<font color=`"#008B00`">$LineValue</font>"
			}
		else {
			$LineValueStr = "$LineValue"
			}
		$cells=@("$LineName","$LineValueStr")
		$healthreport += New-HTMLTableLine $cells
	}

	$healthreport += End-HTMLTable
}

$healthreport | set-content "$tmpdir\healthreport.html"
$healthreport | add-content "$tmpdir\report.html"