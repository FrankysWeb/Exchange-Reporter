function Get-MapiHttpProtocolLog {
    <#
    .SYNOPSIS
        Identifies and reports which Outlook client versions are being used to access Exchange 2016/2019 with MAPIoverHTTP.
    .DESCRIPTION
        Get-MapiHttpProtocolLog is an advanced PowerShell function that parses Exchange Server MapiHttp
        logs to determine what Outlook client versions are being used to access the Exchange Server.
    .PARAMETER LogFile
        The path to the Exchange MapiHttp log files.
    .EXAMPLE
         Get-MapiHttpProtocolLog -LogFile 'C:\Program Files\Microsoft\Exchange Server\V15\Logging\MapiHttp\Mailbox\MapiHttp_2021052018-1.LOG'
    .EXAMPLE
         Get-ChildItem -Path '\\servername\c$\Program Files\Microsoft\Exchange Server\V15\Logging\MapiHttp\Mailbox\*.log' |
         Get-MapiHttpProtocolLog |
         Out-GridView -Title 'Outlook Client Versions'
    .INPUTS
        String
    .OUTPUTS
        PSCustomObject
    .NOTES
        Author:  Mike F Robbins
        Website: http://mikefrobbins.com
        Twitter: @mikefrobbins
		
		Modified by: Frank Zoechling
		Website: https://www.frankysweb.de
		Twitter: @FrankysWeb
    #>
        [CmdletBinding()]
        param (
            [Parameter(Mandatory,
                       ValueFromPipeline)]
            [ValidateScript({
                Test-Path -Path $_ -PathType Leaf -Include '*.log'
            })]
            [string[]]$LogFile
        )
        PROCESS {
            foreach ($file in $LogFile) {
                $Headers = (Get-Content -Path $file -TotalCount 6 | Where-Object {$_ -like '#Fields*'}) -replace '#Fields: ' -split ','
                Import-Csv -Header $Headers -Path $file |
                Where-Object {$_.RequestType -eq 'Connect' -and $_.'ClientSoftware' -eq 'outlook.exe'} |
                Select-Object -Unique -Property @{label='User';expression={$_.'AuthenticatedUser'}},
                                                @{label='Version';expression={Get-OutlookVersion -OutlookBuild $_.'ClientSoftwareVersion'}},
                                                ClientSoftwareVersion
            }
        }
    }

    function Get-OutlookVersion {
	    <#
		.SYNOPSIS
			Outlook Version Table Function
		.DESCRIPTION
			This function is used to determine the Outlook version friendly name.
			Sadly Outlook 2016 / 2019 are sharing the same build numbers. So it's not
			possible to get the exact versions from protocol log.
		.INPUTS
			String
		.OUTPUTS
			String
		.NOTES
			Office Builds: https://docs.microsoft.com/de-de/officeupdates/update-history-office-2019
		#>
        param (
            [string]$OutlookBuild
        )
        switch ($OutlookBuild) {
            # Outlook 2016 / Outlook 365 / Outlook 2019
            {$_ -ge '16.0.11001.20097'} {'Outlook 2016 / 2019 / C2R'; break}
            {$_ -ge '16.0.4229.1003'} {'Outlook 2016 / 2019 / C2R'; break}
            # Outlook 2013
            {$_ -ge '15.0.4569.1506'} {'Outlook 2013 SP1'; break}
            {$_ -ge '15.0.4420.1017'} {'Outlook 2013 RTM'; break}
            # Outlook 2010
            {$_ -ge '14.0.7015.1000'} {'Outlook 2010 SP2'; break}
            {$_ -ge '14.0.6025.1000'} {'Outlook 2010 SP1'; break}
            {$_ -ge '14.0.4734.1000'} {'Outlook 2010 RTM'; break}
            # Outlook 2007
            {$_ -ge '12.0.6606.1000'} {'Outlook 2007 SP3'; break}
            {$_ -ge '12.0.6423.1000'} {'Outlook 2007 SP2'; break}
            {$_ -ge '12.0.6212.1000'} {'Outlook 2007 SP1'; break}
            {$_ -ge '12.0.4518.1014'} {'Outlook 2007 RTM'; break}
            # Outlook 2003
            {$_ -ge '11.0.8303.0'} {'Outlook 2003 SP3'; break}
            {$_ -ge '11.0.8000.0'} {'Outlook 2003 SP2'; break}
            {$_ -ge '11.0.6352.0'} {'Outlook 2003 SP1'; break}
            {$_ -ge '11.0.5604.0'} {'Outlook 2003'; break}
            # Noch älter
            Default {'Unknown Version'}
        }
    }	

function Get-MapiHttpProtocolLogFolder {
	$ExchangeServers = (Get-ExchangeServer | where {$_.AdminDisplayVersion.Major -ge 15}).Name
	foreach ($ExchangeServer in $ExchangeServers) {
		$ExchangeInstallpath = $exinstall.Replace(":","$")  
		$MapiLoggingFolder += "\\" + $ExchangeServer + "\" + $ExchangeInstallpath + "Logging\MapiHttp\Mailbox"
		if (test-path $MapiLoggingFolder) {
			return $MapiLoggingFolder
		}
	}
}

function Get-MapiHttpProtocolLogFiles {
	$MapiLogs = @()
	$MapiLogFolders = Get-MapiHttpProtocolLogFolder
	$StartDate = (Get-Date).AddDays(-$reportinterval)
	foreach ($MapiLogFolder in $MapiLogFolders) {
		$MapiLogs += (Get-ChildItem -Path $MapiLogFolder -Filter *.log | where {$_.LastWriteTime -gt $StartDate} | select fullname).Fullname
		return $MapiLogs
	}
}

function Report-OutlookVersions {
	$OutlookVersions = @()
	$OutookUserConnects = Get-MapiHttpProtocolLogFiles | Get-MapiHttpProtocolLog | group User
	foreach ($OutookUserConnect in $OutookUserConnects) {
		$OutlookVersions += $OutookUserConnect.Group | select -Unique
	}
	return $OutlookVersions
}

$GetOutlookVersions = Report-OutlookVersions
$OutlookVersions = $GetOutlookVersions | group version | sort name

$clientinfo = Generate-ReportHeader "clientinfo.png" "$l_client_header"

$cells=@("$l_client_Oversion","$l_client_Loghits")
$clientinfo += Generate-HTMLTable "$l_client_t1header" $cells

$cells=@()
foreach ($OutlookVersion in $OutlookVersions) {
	$OName = $OutlookVersion.Name
	$OCount = $OutlookVersion.Count
	$cells=@("$OName","$OCount")
	$clientinfo += New-HTMLTableLine $cells
}

$clientinfo += End-HTMLTable

$clientinfo | set-content "$tmpdir\clientreport.html"
$clientinfo | add-content "$tmpdir\report.html"