#Requires -RunAsAdministrator
#--------------------------------------------------------------------------------------
# Exchange Reporter 3.13
# for Exchange Server 2010/2013/2016/2019
# www.frankysweb.de
#
# Generating Exchange Reports
# by Frank Zoechling
#
#--------------------------------------------------------------------------------------

Param(
[Parameter(Mandatory=$false)][string]$Installpath = $PSScriptRoot,
[Parameter(Mandatory=$false)][string]$ExchangeVersion,
[Parameter(Mandatory=$false)][string]$ConfigFile = "settings.ini"
)

#Konsole Header
#--------------------------------------------------------------------------------------

$reporterversion = "3.13"

clear-host
if ($ExchangeVersion)
	{
		$EMSVersion = $ExchangeVersion
	}
$otitle = $host.ui.RawUI.WindowTitle
$host.ui.RawUI.WindowTitle = "Exchange Reporter $reporterversion - www.FrankysWeb.de"


write-host "
------------------------------------------------------------------------------------------"
write-host "
   _____         _                             ______                      _            
  |  ___|       | |                            | ___ \                    | |           
  | |____  _____| |__   __ _ _ __   __ _  ___  | |_/ /___ _ __   ___  _ __| |_ ___ _ __ 
  |  __\ \/ / __| '_ \ / _`` | '_ \ / _`` |/ _ \ |    // _ \ '_ \ / _ \| '__| __/ _ \ '__|
  | |___>  < (__| | | | (_| | | | | (_| |  __/ | |\ \  __/ |_) | (_) | |  | ||  __/ |   
  \____/_/\_\___|_| |_|\__,_|_| |_|\__, |\___| \_| \_\___| .__/ \___/|_|  \__ \___|_|   
                                    __/ |                | |                            
                                   |___/                 |_|                                                 
" -foregroundcolor cyan
write-host "
			for Exchange Server 2010 / 2013 / 2016 / 2019
							 
                                     www.FrankysWeb.de

                                       Version: $reporterversion

------------------------------------------------------------------------------------------
"
#Prüfen ob PowerShell 4.0 vorhanden
#--------------------------------------------------------------------------------------

write-host " Checking Powershell Version:" -nonewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
$psversion = (get-host).version.major

if ($psversion -ge "4")
	{
		
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "OK (PowerShell $psversion)" -foregroundcolor green
	}
else
	{
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error" -foregroundcolor red
		exit 0
		write-host ""
	}

#Laden der Funktionen aus "Include-Functions.ps1"
#--------------------------------------------------------------------------------------

write-host " Loading functions from Include-Functions.ps1:" -nonewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
$functionfile = test-path "$installpath\Includes\Include-Functions.ps1"
if ($functionfile)
	{
		. "$installpath\Includes\Include-Functions.ps1"
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Done" -foregroundcolor green
	}
else
	{
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error (not found)" -foregroundcolor red
		exit 0
		write-host ""
 }

# settings.ini einlesen
#--------------------------------------------------------------------------------------

try 
	{
		write-host " Loading settings from $ConfigFile`:" -nonewline
		$origpos = $host.UI.RawUI.CursorPosition
		$origpos.X = 70
		$globalsettingsfile = "$installpath\$ConfigFile"
		$inifile = get-inicontent "$globalsettingsfile"
	}
Catch
	{
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error" -foregroundcolor red
		exit 0
		write-host ""
	}
	$host.UI.RawUI.CursorPosition = $origpos
	write-host "Done" -foregroundcolor green

# Settings verarbeiten
#--------------------------------------------------------------------------------------

 # INI sections
$activemoduleshash = $inifile["Modules"]
$3rdPartyactivemoduleshash = $inifile["3rdPartyModules"]
$reportsettingshash = $inifile["Reportsettings"]
$reportsettings = convert-hashtoobject $reportsettingshash
$languagehash = $inifile["LanguageSettings"]
$excludehash = $inifile["ExcludeList"]
$languagesettings = convert-hashtoobject $languagehash
$excludelist = convert-hashtoobject $excludehash
$activemodules = convert-hashtoobject $activemoduleshash
$excludelist = $excludelist | where {$_.setting -notmatch "Comment" -and $_.setting -notmatch ";"}
$activemodules = $activemodules | where {$_.setting -notmatch "Comment" -and $_.setting -notmatch ";"} | sort setting
$3rdPartyactivemodules = convert-hashtoobject $3rdPartyactivemoduleshash
$3rdPartyactivemodules = $3rdPartyactivemodules | where {$_.setting -notmatch "Comment" -and $_.setting -notmatch ";"} | sort setting

#Einstellungen:
#--------------------------------------------------------------------------------------

$ReportInterval = ($reportsettings | Where-Object {$_.Setting -eq "Interval"}).Value
$CleanTMPFolder = ($reportsettings | Where-Object {$_.Setting -eq "CleanTMPFolder"}).Value
$Errorlog = ($reportsettings | Where-Object {$_.Setting -eq "WriteErrorLog"}).Value
$AddPDFFileToMail = ($reportsettings | Where-Object {$_.Setting -eq "AddPDFFileToMail"}).Value
$SMTPAuth = ($reportsettings | Where-Object {$_.Setting -eq "SMTPServerAuth"}).Value

if ($SMTPAuth -match "yes")
	{
		$SMTPServerUser = ($reportsettings | Where-Object {$_.Setting -eq "SMTPServerUser"}).Value
		$SMTPServerPass = ($reportsettings | Where-Object {$_.Setting -eq "SMTPServerPass"}).Value
  
		$secpasswd = ConvertTo-SecureString $SMTPServerPass -AsPlainText -Force
		$smtpcreds = New-Object System.Management.Automation.PSCredential ($SMTPServerUser, $secpasswd)
	}
	
$Recipient = ($reportsettings | Where-Object {$_.Setting -eq "Recipient"}).Value
[array]$Recipient = $Recipient.split(",")
$Sender = ($reportsettings | Where-Object {$_.Setting -eq "Sender"}).Value
$Mailserver = ($reportsettings | Where-Object {$_.Setting -eq "Mailserver"}).Value
$Subject = ($reportsettings | Where-Object {$_.Setting -eq "Subject"}).Value
[int]$DisplayTop = ($reportsettings | Where-Object {$_.Setting -eq "DisplayTop"}).Value
$language = ($languagesettings | Where-Object {$_.Setting -eq "Language"}).Value

# Errorlog schreiben
#--------------------------------------------------------------------------------------    
if ($errorlog -match "yes")
	{
		$logtime = get-date
		"-Start-- $logtime ----------------------------------------------------------------------------------" | add-content "$installpath\ErrorLog.txt"
	}

# Sprache anzeigen
#--------------------------------------------------------------------------------------
try 
	{
		write-host " Setting Report Language:" -nonewline
		$origpos = $host.UI.RawUI.CursorPosition
		$origpos.X = 70

	}
Catch
	{
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error" -foregroundcolor red
		if ($errorlog -match "yes")
			{
				
				$error[0] | add-content "$installpath\ErrorLog.txt"
			}
		exit 0
		write-host ""
	}
	$host.UI.RawUI.CursorPosition = $origpos
	write-host "$language" -foregroundcolor green

#Lade Exchange Snapin
#-------------------------------------------------------------------------------------- 
write-host " Loading Exchange Management SnapIn:" -nonewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
try
	{
		. "$installpath\Includes\Include-ExchangeSnapins.ps1"
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Done" -foregroundcolor green
	}
Catch
	{
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error" -foregroundcolor red
		if ($errorlog -match "yes")
			{
				
				$error[0] | add-content "$installpath\ErrorLog.txt"
			}
		exit 0
		write-host ""
	}
	
# Exchange Version ermitteln
#--------------------------------------------------------------------------------------
write-host " Checking Exchange Management Shell:" -nonewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
	if (!$emsversion)
		{
			$emsversion = Get-ExchangeVersionByRegistry
		}
	if ($emsversion -match "2010")
		{
			$host.UI.RawUI.CursorPosition = $origpos
			write-host "OK (Exchange 2010)" -foregroundcolor green
		}
	if ($emsversion -match "2013")
		{
			$host.UI.RawUI.CursorPosition = $origpos
			write-host "OK (Exchange 2013)" -foregroundcolor green
		}
	if ($emsversion -match "2016")
		{
			$host.UI.RawUI.CursorPosition = $origpos
			write-host "OK (Exchange 2016)" -foregroundcolor green
		}
	if ($emsversion -match "2019")
		{
			$host.UI.RawUI.CursorPosition = $origpos
			write-host "OK (Exchange 2019)" -foregroundcolor green
		}
	if (!$emsversion)
		{
			$host.UI.RawUI.CursorPosition = $origpos
			write-host "Error (EMS not found)" -foregroundcolor red
			exit 0
			write-host ""
		}
	if ($emsversion -notmatch "2010" -and $emsversion -notmatch "2013" -and $emsversion -notmatch "2016" -and $emsversion -notmatch "2019")
		{
			$host.UI.RawUI.CursorPosition = $origpos
			write-host "Error (Wrong EMS Version)" -foregroundcolor red
		if ($errorlog -match "yes")
			{
				
				"Wrong EMS Version" | add-content "$installpath\ErrorLog.txt"
			}
			exit 0
			write-host ""
		}

#Temporäres Verzeichnis erstellen
#--------------------------------------------------------------------------------------

write-host " Generating temp. Directory:" -nonewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
 if (test-path "$installpath\TEMP") {Remove-Item "$installpath\TEMP" -Force -Recurse}
 $tmpdir = New-Item "$installpath\TEMP" -Type directory -ea 0
 $tmpdir = $tmpdir.fullname
 if ($tmpdir)
	{
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Done" -foregroundcolor green
	}
 else
	{
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error" -foregroundcolor red
		if ($errorlog -match "yes")
			{
				
				$error[0] | add-content "$installpath\ErrorLog.txt"
			}
		exit 0
		write-host ""
	}

#Lade Assembly
#--------------------------------------------------------------------------------------
write-host " Loading .NET Assemblies:" -nonewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
try
	{
		[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") 
		[void][Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization")
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Done" -foregroundcolor green
	}
Catch
	{
		if ($errorlog -match "yes")
			{
				
				$error[0] | add-content "$installpath\ErrorLog.txt"
			}
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error" -foregroundcolor red
		exit 0
		write-host ""
 }
 
#Häufig genutzte Variablen
#--------------------------------------------------------------------------------------
write-host " Loading global Variables:" -nonewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
try 
	{
		$mbxservers = Get-MailboxServer -ea 0
		if ($emsversion -match "2010")
			{
				$casservers = Get-ClientAccessServer -ea 0
			}
		if ($emsversion -match "2013")
			{
				$casservers = Get-ClientAccessServer -ea 0
			}
		if ($emsversion -match "2016")
			{
				$casservers = Get-ClientAccessService -ea 0
			}
		if ($emsversion -match "2019")
			{
				$casservers = Get-ClientAccessService -ea 0
			}
		
		$exservers = get-exchangeserver -ea 0 | where {$_.admindisplayversion.major -ge 14}
        $exdomains = $exservers | select domain -Unique
        foreach ($exdomain in $exdomains)
            {
                $domainname = $exdomain.domain
                $domaincontrollers = Get-DomainController -domain $domainname -ea 0
            }
		$orgname = (Get-OrganizationConfig).Name
		#$emsversion = Get-ExchangeVersionByRegistry
		$files = @()
		$host.UI.RawUI.CursorPosition = $origpos
		$modulpath = "$installpath" +"\modules"
		$languagefilepath = "$installpath" + "\Language\" + "$language"
		$Start = (Get-Date -Hour 00 -Minute 00 -Second 00).AddDays(-$ReportInterval)
		$End = (Get-Date -Hour 00 -Minute 00 -Second 00)
		$today = get-date | convert-date
		write-host "Done" -foregroundcolor green
		$entireforrest = Set-ADServerSettings -ViewEntireForest $True
	}
Catch
	{
		if ($errorlog -match "yes")
			{
				$error[0] | add-content "$installpath\ErrorLog.txt"
			}
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error" -foregroundcolor red
		exit 0
		write-host ""
	}
write-host ""
write-host "------------------------------------------------------------------------------------------"
write-host ""

#MODULE
#--------------------------------------------------------------------------------------

#HTML Datei vorbereiten
$htmlheader = new-htmlheader ExchangeReporter
$htmlheader | set-content "$tmpdir\report.html"

foreach ($activemodule in $activemodules)
	{
		$module = $activemodule.Value
		write-host " Working on Module '$module':" -nonewline
		$origpos = $host.UI.RawUI.CursorPosition
		$origpos.X = 70
		try 
			{
				. "$languagefilepath\$module"
				. "$modulpath\$module"
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Done" -foregroundcolor green
			}
		Catch
			{
				if ($errorlog -match "yes")
					{
						$module | add-content "$installpath\ErrorLog.txt"
						$error[0] | add-content "$installpath\ErrorLog.txt"
					}
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Error" -foregroundcolor red
				write-host ""
			}

	}

foreach ($3rdPartyactivemodule in $3rdPartyactivemodules)
	{
		$module = $3rdPartyactivemodule.Value
		write-host " Working on 3rd Party Module '$module':" -nonewline
		$origpos = $host.UI.RawUI.CursorPosition
		$origpos.X = 70
		try 
			{
				. "$modulpath\3rdParty\$module"
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Done" -foregroundcolor green
			}
		Catch
			{
				if ($errorlog -match "yes")
					{
						$module | add-content "$installpath\ErrorLog.txt"
						$error[0] | add-content "$installpath\ErrorLog.txt"
					}
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Error" -foregroundcolor red
				write-host ""
			}

	}

# Report vorbereiten
#--------------------------------------------------------------------------------------

Generate-ReportFooter | add-content "$tmpdir\report.html"
$mailbody = get-content "$tmpdir\report.html" | out-string

foreach ($activemodule in $activemodules)
	{
		$module = $activemodule.Value
		$pngfile = $module.replace(".ps1",".png")
		$files += Get-ChildItem "$Installpath\Images\$pngfile" -Recurse | Where {-NOT $_.PSIsContainer} | foreach {$_.fullname}
	}
	
foreach ($3rdPartyactivemodule in $3rdPartyactivemodules)
	{
		$module = $3rdPartyactivemodule.Value
		$pngfile = $module.replace(".ps1",".png")
		$files += Get-ChildItem "$Installpath\Images\$pngfile" -Recurse | Where {-NOT $_.PSIsContainer} | foreach {$_.fullname}
	}

$files += Get-ChildItem "$Installpath\Images\reportheader.png" | Where {-NOT $_.PSIsContainer} | foreach {$_.fullname}
$files += Get-ChildItem "$Installpath\TEMP\*.png" | Where {-NOT $_.PSIsContainer} | foreach {$_.fullname}
	
# PDF File erzeugen
#--------------------------------------------------------------------------------------

if ($AddPDFFileToMail -match "yes")
	{
	write-host " Saving PDF Report:" -nonewline
	$origpos = $host.UI.RawUI.CursorPosition
	$origpos.X = 70
	if (test-path "C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe")
	{
	try
		{
		$pdfreport = $mailbody.replace("cid:","")
		$pdfreport | set-content "$installpath\TEMP\PDFReport.htm"
		$pdfpath = "$installpath\TEMP"
		$pdffile = "$installpath\TEMP\Report.pdf"
		foreach ($file in $files)
			{
				copy-item $file -Destination $pdfpath -force -ea 0
			}
		$pdf = &"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" --quiet --enable-local-file-access "$installpath\TEMP\PDFReport.htm" "$installpath\TEMP\Report.pdf" | out-null
		$files += "$installpath\TEMP\Report.pdf"
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Done" -foregroundcolor green
		}
	Catch
		{
			if ($errorlog -match "yes")
				{
					$error[0] | add-content "$installpath\ErrorLog.txt"
				}
			$host.UI.RawUI.CursorPosition = $origpos
			write-host "Error" -foregroundcolor red
			write-host ""
		}
	}
	else
	{
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error (WKHTML not found)" -foregroundcolor red
		write-host ""
	}
	}
	
# Report per Mail verschicken
#--------------------------------------------------------------------------------------
write-host ""
write-host " Sending Report:" -nonewline
$origpos = $host.UI.RawUI.CursorPosition
$origpos.X = 70
try 
	{
		if ($SMTPAuth -match "yes")
			{
				send-mailmessage -encoding UTF8 -from "Exchange Reporter - www.FrankysWeb.de <$sender>" -to "$Recipient"  -subject "$subject" -smtpserver $mailserver -BodyAsHtml -Body $mailbody -Attachments $files -Credential $smtpcreds
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Done" -foregroundcolor green
			}
		else
			{
				send-mailmessage -encoding UTF8 -from "Exchange Reporter - www.FrankysWeb.de <$sender>" -to $Recipient  -subject "$subject" -smtpserver $mailserver -BodyAsHtml -Body $mailbody -Attachments $files
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Done" -foregroundcolor green
				write-host ""
			}
   }
Catch
   {
		if ($errorlog -match "yes")
			{
				$error[0] | add-content "$installpath\ErrorLog.txt"
			}
		$host.UI.RawUI.CursorPosition = $origpos
		write-host "Error" -foregroundcolor red
		write-host ""
   }

# Report per FTP Hochladen
#--------------------------------------------------------------------------------------

if ($FTPUpload -match "yes")
	{
		write-host ""
		write-host "------------------------------------------------------------------------------------------"
		write-host " Uploading files to FTP Server:" -nonewline
		$origpos = $host.UI.RawUI.CursorPosition
		$origpos.X = 70
		try
			{
			$FTPServer = ($reportsettings | Where-Object {$_.Setting -eq "FTPServer"}).Value
			$FTPUser = ($reportsettings | Where-Object {$_.Setting -eq "FTPUser"}).Value
			$FTPPass = ($reportsettings | Where-Object {$_.Setting -eq "FTPPass"}).Value
			$FTPLocalFolder = ($reportsettings | Where-Object {$_.Setting -eq "FTPLocalFolder"}).Value

			$webclient = New-Object System.Net.WebClient 
			$webclient.Credentials = New-Object System.Net.NetworkCredential($FTPUser,$FTPPass)  
 
			foreach($item in (dir $FTPLocalFolder)){ 
			"Uploading $item..." 
			$uri = New-Object System.Uri($FTPServer+$item.Name) 
			$webclient.UploadFile($uri, $item.FullName) 
			} 
			
			write-host "Done" -foregroundcolor green
			write-host ""
			}
		Catch
			{
				if ($errorlog -match "yes")
					{
						$error[0] | add-content "$installpath\ErrorLog.txt"
					}
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Error" -foregroundcolor red
				write-host ""
			}		
	}

# Aufräumen
#--------------------------------------------------------------------------------------
if ($CleanTMPFolder -match "yes")
	{
		write-host ""
		write-host "------------------------------------------------------------------------------------------"
		write-host " Cleaning up temp. Directory:" -nonewline
		$origpos = $host.UI.RawUI.CursorPosition
		$origpos.X = 70
		try 
			{
				$delTMPdir = remove-item $tmpdir -recurse -force
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Done" -foregroundcolor green
				write-host ""
			}
		Catch
			{
				if ($errorlog -match "yes")
					{
						$error[0] | add-content "$installpath\ErrorLog.txt"
					}
				$host.UI.RawUI.CursorPosition = $origpos
				write-host "Error" -foregroundcolor red
				write-host ""
			}
	}

	
# Errorlog schliessen
#--------------------------------------------------------------------------------------

if ($errorlog -match "yes")
	{
		$logtime = get-date
		"-End--- $logtime ----------------------------------------------------------------------------------" | add-content "$installpath\ErrorLog.txt"
	}

$host.ui.RawUI.WindowTitle = $otitle