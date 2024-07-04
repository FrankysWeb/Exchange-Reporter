$NSPreport = Generate-ReportHeader "nospamproxy.png" "$l_nsp_header"

$NSPsettingshash = $inifile["NoSpamProxy"]
$NSPsettings = convert-hashtoobject $NSPsettingshash

$nspserver = ($nspsettings| Where-Object {$_.Setting -eq "NSPServer"}).Value
$nspuser = ($nspsettings| Where-Object {$_.Setting -eq "NSPUser"}).Value
$nsppassword = ($nspsettings | Where-Object {$_.Setting -eq "NSPPassword"}).Value

$NSPSecpassword = $nsppassword | ConvertTo-SecureString -AsPlainText -Force
$NSPCreds = New-Object System.Management.Automation.PSCredential -ArgumentList $nspuser, $NSPSecpassword

$NSPPSSession = New-PSSession -ComputerName $nspserver -Credential $NSPCreds

#Get data from NSP
try {
	$timespan = New-TimeSpan -Days $reportinterval

	$NSPInboundSuccess = Invoke-Command -Session $NSPPSSession -ScriptBlock { Get-NspMessageTrack -Status Success -Age $args[0] -Directions FromExternal} -argumentlist $timespan
	$NSPOutboundSuccess = Invoke-Command -Session $NSPPSSession -ScriptBlock { Get-NspMessageTrack -Status Success -Age $args[0] -Directions FromLocal} -argumentlist $timespan
	$NSPInboundPermBlocked = Invoke-Command -Session $NSPPSSession -ScriptBlock { Get-NspMessageTrack -Status PermanentlyBlocked -Age $args[0] -Directions FromExternal} -argumentlist $timespan
	$NSPOutboundPending = Invoke-Command -Session $NSPPSSession -ScriptBlock { Get-NspMessageTrack -Status DeliveryPending -Age $args[0] -Directions FromLocal} -argumentlist $timespan

	$NSPLicense = Invoke-Command -Session $NSPPSSession -ScriptBlock { Get-NspLicense | select License }

	$NSPServices = Invoke-Command -Session $NSPPSSession -ScriptBlock { get-service netatwork* | select displayname,status }
	
	$NSPCertificate = Invoke-Command -Session $NSPPSSession -ScriptBlock {
		$NSPCurrentCert = Get-NspReceiveConnector
		$NSPCurrentCertTP = $NSPCurrentCert.TlsCertificate.Thumbprint.ToUpper()
		$NSPCurrentCertData = Get-ChildItem "Cert:\LocalMachine\My" | where {$_.Thumbprint -match $NSPCurrentCertTP}
		
		$NSPCurrentCertValidUntil = $NSPCurrentCertData.NotAfter
		$NSPDaysRemain = ($NSPCurrentCertValidUntil - (get-date)).Days
		$NSPCurrentCertFN = $NSPCurrentCertData.FriendlyName
		$NSPCurrentCertSubject = $NSPCurrentCertData.Subject.split("=")[1]
		
		$NSPCertObject = New-Object PSObject
		$NSPCertObject | Add-Member NoteProperty Thumbprint $NSPCurrentCertTP
		$NSPCertObject | Add-Member NoteProperty ValidUntil $NSPCurrentCertValidUntil
		$NSPCertObject | Add-Member NoteProperty FriendlyName $NSPCurrentCertFN
		$NSPCertObject | Add-Member NoteProperty Subject $NSPCurrentCertSubject
		$NSPCertObject | Add-Member NoteProperty DaysRemain $NSPDaysRemain
		
		return $NSPCertObject
		}
		
	$NSPIssues = Invoke-Command -Session $NSPPSSession -ScriptBlock { Get-NspIssue }
		
	$NSPLargeFiles = Invoke-Command -Session $NSPPSSession -ScriptBlock {
		$LargeFiles = Get-NspLargeFile
		if (!$LargeFiles) { $LargeFiles = 0 } else { $LargeFiles = $LargeFiles.count }
		return $LargeFiles
		}
	
	$NSPLargeFilesVolume = Invoke-Command -Session $NSPPSSession -ScriptBlock {
		$LargeFilesVolume = Get-NspLargeFile
		if (!$LargeFilesVolume ) { $LargeFilesVolume  = 0 } else { $LargeFilesVolume  = [System.Math]::Round(($LargeFilesVolume.FileSize | measure -sum).Sum / 1024 / 1024, 2)}
		return $LargeFilesVolume
		}

	$NSPEvents =  Invoke-Command -Session $NSPPSSession -ScriptBlock {
		Get-EventLog -LogName "Net at Work Mail Gateway" -after  $args[0] | where {$_.entrytype -match "Warning" -or $_.entrytype -match "Error"}
	} -argumentlist $Start
		
	$CleanUp = Remove-PSsession $NSPPSSession
}
catch {
}

#Format data

try {
	if (!$NSPInboundSuccess) { $NSPInboundSuccess = 0 } else { $NSPInboundSuccess = $NSPInboundSuccess.count }
	if (!$NSPOutboundSuccess) { $NSPOutboundSuccess = 0 } else { $NSPOutboundSuccess = $NSPOutboundSuccess.count }
	if (!$NSPInboundPermBlocked) { $NSPInboundPermBlocked = 0 } else { $NSPInboundPermBlocked = $NSPInboundPermBlocked.count }
	if (!$NSPOutboundPending) { $NSPOutboundPending = 0 } else { $NSPOutboundPending = $NSPOutboundPending.count }
}
catch {
}

#Add data to Report

$cells=@("$l_nsp_InboundSuccess","$l_nsp_OutboundSucess","$l_nsp_InboundBlocked","$l_nsp_Pending")
$NSPreport += Generate-HTMLTable "$l_nsp_Mailstats" $cells

$cells=@($NSPInboundSuccess,$NSPOutboundSuccess,$NSPInboundPermBlocked,$NSPOutboundPending)
$NSPreport += New-HTMLTableLine $cells
$NSPreport += End-HTMLTable

$filename = "nspmailstats.png"
$chartdata = @{$l_nsp_InboundSuccess=$NSPInboundSuccess; $l_nsp_OutboundSucess=$NSPOutboundSuccess; $l_nsp_InboundBlocked=$NSPInboundPermBlocked; $l_nsp_Pending=$NSPOutboundPending} 
new-piechart "300" "300" "E-Mail Statistik" $chartdata "$tmpdir\$filename"

$NSPreport += Include-HTMLInlinePictures "$tmpdir\$filename"

$cells=@("$l_nsp_Count","$l_nsp_Volume")
$NSPreport += Generate-HTMLTable "$l_nsp_LargeFiles" $cells

$cells=@("$NSPLargeFiles","$NSPLargeFilesVolume MB")
$NSPreport += New-HTMLTableLine $cells

$NSPreport += End-HTMLTable

#Events

$cells=@("$l_nsp_Source","$l_nsp_TimeStamp","$l_nsp_ID","$l_nsp_Count","$l_nsp_Message")
$NSPreport += Generate-HTMLTable "$l_nsp_headerEvents" $cells

if ($NSPEvents) {
	$NSPEventGroups = $NSPEvents | group InstanceId | sort count -Descending
	Foreach ($NSPEventGroup in $NSPEventGroups) {
		$event = $NSPEventGroup.Group | select -first 1
		$eventcount = $NSPEventGroup.count
		$eventsource = $event.Source
		$eventid = $event.InstanceId
		$eventtime = $event.TimeGenerated
		$eventtime = $eventtime | get-date -format "dd.MM.yy hh:mm:ss"
		$eventmessage = $event.Message
		$eventmeslength = $eventmessage.Length
		if ($eventmeslength -gt 200) {
			$eventcontent = $eventmessage.Substring(0,200)
			$eventcontent = $eventcontent + "..."
		}
		else {
			$eventcontent = $eventmessage
		}
		$cells=@("$eventsource","$eventtime","$eventid","$eventcount","$eventcontent")
		$NSPreport += New-HTMLTableLine $cells
	}
}
else {
	$cells=@("$l_nsp_noerror")
	$NSPreport += New-HTMLTableLine $cells
}
$NSPreport += end-htmltable

#Issues

$cells=@("$l_nsp_Severity","$l_nsp_ReportedOn","$l_nsp_Role","$l_nsp_IssueText")
$NSPreport += Generate-HTMLTable "$l_nsp_headerIssues" $cells

if ($NSPIssues) {
	foreach ($NSPIssue in $NSPIssues) {
		$IssueSeverity = $NSPIssue.Severity
		$IssueReportedOn = $NSPIssue.ReportedOn | get-date -format "dd.MM.yy hh:mm:ss" 
		$IssueRole = $NSPIssue.Role
		$IssueText = $NSPIssue.Text
		
		$cells=@("$IssueSeverity","$IssueReportedOn","$IssueRole","$IssueText")
		$NSPreport += New-HTMLTableLine $cells
	}
}
else {
	$cells=@("$l_nsp_Bestform")
	$NSPreport += New-HTMLTableLine $cells
}

$NSPreport += end-htmltable

#Certificate

$cells=@("$l_nsp_FriedlyName","$l_nsp_Thumbprint","$l_nsp_ValidUntil","$l_nsp_Subject","$l_nsp_RemainingDays")
$NSPreport += Generate-HTMLTable "$l_nsp_headerCertificate" $cells

if ($NSPCertificate) {
	$NSPCertTP = $NSPCertificate.Thumbprint
	$NSPVaildUntil = $NSPCertificate.ValidUntil | get-date -format "dd.MM.yy hh:mm:ss"
	$NSPFriendlyname = $NSPCertificate.FriendlyName
	$NSPSubject = $NSPCertificate.Subject
	$NSPRemainingDays = $NSPCertificate.DaysRemain
	if ($NSPRemainingDays -ge 30) {
		$NSPRemainingDaysText = "<font color=`"#008B00`">$NSPRemainingDays $l_nsp_days</font>"
	}
	else {
		$NSPRemainingDaysText = "<font color=`"#CD0000`">$NSPRemainingDays $l_nsp_days</font>"
	}
	
	$cells=@("$NSPFriendlyname","$NSPCertTP","$NSPVaildUntil","$NSPSubject","$NSPRemainingDaysText")
	$NSPreport += New-HTMLTableLine $cells
}
else {
	$cells=@("$l_nsp_NoCertificateFound")
	$NSPreport += New-HTMLTableLine $cells
}

$NSPreport += end-htmltable

#Services

$cells=@("$l_nsp_ServiceName","$l_nsp_ServiceStatus")
$NSPreport += Generate-HTMLTable "$l_nsp_ServiceHeader" $cells

if ($NSPServices) {
	foreach ($NSPService in $NSPServices) {
		$NSPServiceName = $NSPService.Displayname
		$NSPServiceStatus = $NSPService.Status
		if ($NSPServiceStatus -match "Running") {
			$NSPServiceStatusString = "<font color=`"#008B00`">$l_nsp_ServiceStatusOK</font>"
		}
		else {
			$NSPServiceStatusString = "<font color=`"#CD0000`">$l_nsp_ServiceStatusNOK</font>"
		}
		$cells=@("$NSPServiceName","$NSPServiceStatusString")
		$NSPreport += New-HTMLTableLine $cells
	}
}
else {
 	$cells=@("$l_nsp_ServiceNotFound")
	$NSPreport += New-HTMLTableLine $cells
}
$NSPreport += end-htmltable

#Lizenz

$cells=@("$l_nsp_LicenseName","$l_nsp_LicenseStatus")
$NSPreport += Generate-HTMLTable "$l_nsp_LicenseHeader" $cells

if ($NSPLicense) {
	[string]$NSPAntispam = $NSPLicense.license.AntiSpam
	[string]$NSPDisclaimer = $NSPLicense.license.Disclaimer
	[string]$NSPCrypto = $NSPLicense.license.Cryptography
	[string]$NSPLargeFile = $NSPLicense.license.LargeFileTransfer
	[string]$NSPServiceContract = $NSPLicense.license.ServiceContractExpiresOn
	[string]$NSPCyren = $NSPLicense.license.CyrenServicesEnabledUntil
	
	$cells=@("$l_nsp_LicenseAntiSpam","$NSPAntispam")
	$NSPreport += New-HTMLTableLine $cells
	
	$cells=@("$l_nsp_LicenseDisclaimer","$NSPDisclaimer")
	$NSPreport += New-HTMLTableLine $cells
	
	$cells=@("$l_nsp_LicenseCrypto","$NSPCrypto")
	$NSPreport += New-HTMLTableLine $cells
	
	$cells=@("$l_nsp_LicenseLFT","$NSPLargeFile")
	$NSPreport += New-HTMLTableLine $cells

	$cells=@("$l_nsp_LicenseService","$NSPServiceContract")
	$NSPreport += New-HTMLTableLine $cells

	$cells=@("$l_nsp_LicenseCyren","$NSPCyren")
	$NSPreport += New-HTMLTableLine $cells	
}
else {
 	$cells=@("$l_nsp_LicenseNotFound")
	$NSPreport += New-HTMLTableLine $cells
}
$NSPreport += end-htmltable


#Finish Report

$NSPreport | set-content "$tmpdir\NSPreport.html"
$NSPreport | add-content "$tmpdir\report.html"


#Exchange Reporter Server
# Set-Item WSMan:\localhost\Client\TrustedHosts -Value "*" -Force

