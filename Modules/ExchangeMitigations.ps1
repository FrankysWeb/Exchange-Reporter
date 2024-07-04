$em = Generate-ReportHeader "ExchangeMitigations.png" "$l_em_header"

if ($emsversion -match "2016" -or $emsversion -match "2019") {

#EM enabled on Org Level
$OrgName = (Get-OrganizationConfig).Name
$OrgEnabled = Get-OrganizationConfig | select MitigationsEnabled
if ($OrgEnabled.MitigationsEnabled -eq $true) {
	$OrgEnabledStr = "<font color=`"#008B00`">$l_em_org_enabled</font>"
}
else {
	$OrgEnabledStr = "<font color=`"#E59400`">$l_em_org_disabled</font>"
}

$cells=@("$l_em_name","$l_em_state")
$em  += Generate-HTMLTable "$l_em_t1header" $cells

$cells=@("$OrgName","$OrgEnabledStr")
	$em += New-HTMLTableLine $cells

#EM enabled on Server Level
$ServersEnabled = Get-ExchangeServer | select Name,MitigationsEnabled
foreach ($ServerEnabled in $ServersEnabled) {
	if ($OrgEnabled.MitigationsEnabled -eq $true -and $ServerEnabled.MitigationsEnabled -eq $True) {
		$ServerEnabledStr = "<font color=`"#008B00`">$l_em_server_enabled</font>"
	}
	if ($OrgEnabled.MitigationsEnabled -eq $true -and $ServerEnabled.MitigationsEnabled -eq $false) {
		$ServerEnabledStr = "<font color=`"#E59400`">$l_em_server_disabled</font>"
	}
	if ($OrgEnabled.MitigationsEnabled -eq $false -and $ServerEnabled.MitigationsEnabled -eq $True) {
		$ServerEnabledStr = "<font color=`"#008B00`">$l_em_org_enabled</font>"
	}
	if ($OrgEnabled.MitigationsEnabled -eq $false -and $ServerEnabled.MitigationsEnabled -eq $False) {
		$ServerEnabledStr = "<font color=`"#008B00`">$l_em_org_disabled</font>"
	}
	$EMServername = $ServerEnabled.Name
	$cells=@()
	$cells=@("$EMServerName","$ServerEnabledStr")
	$em += New-HTMLTableLine $cells
}

$em += End-HTMLTable

#Test Mitigation Service
function Test-MitigationServiceConnectivity{

	$ConfigurationCloudEndpoint="https://officeclient.microsoft.com/getexchangemitigations"

	try{
		[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
		$result = Invoke-RestMethod -Method Get -Uri $ConfigurationCloudEndpoint
		return $True
	}
	catch{
		return $False
	}
}

$emtest = Test-MitigationServiceConnectivity
if ($emtest -eq $True) {
	$emtestresult = "<font color=`"#008B00`">$l_em_test_success</font>"
}
else {
	$emtestresult = "<font color=`"#008B00`">$l_em_test_failed</font>"
}

$cells=@("$l_em_test","$l_em_teststate")
$em  += Generate-HTMLTable "$l_em_t3header" $cells

$cells=@("$l_em_testname","$emtestresult")
	$em += New-HTMLTableLine $cells
	
$em += End-HTMLTable

#Get Mitigations
$cells=@("$l_em_Server","$l_em_ID","$l_em_Type","$l_em_desc","$l_em_mitstatus")
$em  += Generate-HTMLTable "$l_em_t2header" $cells

$Mitigations = . "$Exinstall\Scripts\Get-Mitigations.ps1"

foreach ($Mitigation in $Mitigations) {
	$MitServer = $Mitigation.Server
	$MitID = $Mitigation.ID
	$MitType = $Mitigation.Type
	$MitDesc = $Mitigation.Description
	$MitStatus = $Mitigation.Status
	if ($MitStatus -match "Applied") {
		$MitStatusStr = "<font color=`"#008B00`">$l_em_mit_status_applied</font>"
	}
	else {
		$MitStatusStr = "<font color=`"#008B00`">$l_em_mit_status_notapplied</font>"
	}
	
	$cells=@()
	$cells=@("$MitServer","$MitID","$MitType","$MitDesc","$MitStatusStr")
	$em += New-HTMLTableLine $cells
}
$em += End-HTMLTable

$em | set-content "$tmpdir\ExchangeMitigations.html"
$em | add-content "$tmpdir\report.html"

}
else {
	$cells=@("")
	$em  += Generate-HTMLTable "$l_em_t4header" $cells
	$em += End-HTMLTable
}