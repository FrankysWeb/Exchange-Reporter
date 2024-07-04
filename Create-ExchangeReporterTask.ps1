#--------------------------------------------------------------------------------------
# Exchange Reporter Task
# für Exchange Server 2010/2013/2016/2019
# www.frankysweb.de
#
# erstellt geplanten Task für Exchange Reporter
# by Frank Zöchling
#
#--------------------------------------------------------------------------------------

Param([Parameter(Mandatory=$false)][string]$Installpath = $PSScriptRoot)

clear-host
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
                for Exchange Server 2010 / 2013 / 2016 / 2019 / Office365
							 
                                     www.FrankysWeb.de

                                       Version: 3.12

------------------------------------------------------------------------------------------
"

#Laden der Funktionen aus "Include-Functions.ps1"
#--------------------------------------------------------------------------------------

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
		write-host " Loading settings from settings.ini:" -nonewline
		$origpos = $host.UI.RawUI.CursorPosition
		$origpos.X = 70
		$globalsettingsfile = "$installpath\settings.ini"
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

write-host "
-----------------------------------------------------------------------------------------

 What do you want to do?

 1 - Create / refresh Task
 2 - Delete Task
"
while ($aktion -notmatch 1 -and $aktion -notmatch 2)
	{
		$aktion = read-host " Action"
	}

$filepath = "$installpath" + "\New-ExchangeReport.ps1"

if ($aktion -match 1)
	{
	
		$reportsettingshash = $inifile["Reportsettings"]
		$reportsettings = convert-hashtoobject $reportsettingshash
		$reportinterval = ($reportsettings | Where-Object {$_.Setting -eq "Interval"}).Value
		
		if ($reportinterval -match 1)
			{
				write-host ""
				$zeitpunkt = read-host " Please specify Starttime (Example: 22:00)"
				$username = read-host " Please enter Task User (Domain\User)"
				$SecurePassword = read-host " Password" -AsSecureString
				$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
				$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

				$startTime = "$zeitpunkt" | get-date -format s
				
				$taskService = New-Object -ComObject Schedule.Service 
				$taskService.Connect() 
  
				$rootFolder = $taskService.GetFolder($NULL) 
  
				$taskDefinition = $taskService.NewTask(0) 
  
				$registrationInformation = $taskDefinition.RegistrationInfo 
  
				$registrationInformation = $taskDefinition.RegistrationInfo 
				$registrationInformation.Description = "Exchange Reporter Task - www.FrankysWeb.de"
				$registrationInformation.Author = $username
  
				$taskPrincipal = $taskDefinition.Principal 
				$taskPrincipal.LogonType = 1 
				$taskPrincipal.UserID = $username
				$taskPrincipal.RunLevel = 0 
  
				$taskSettings = $taskDefinition.Settings 
				$taskSettings.StartWhenAvailable = $true
				$taskSettings.RunOnlyIfNetworkAvailable = $true
				$taskSettings.Priority = 7 
  
				$taskTriggers = $taskDefinition.Triggers 
  
				$executionTrigger = $taskTriggers.Create(2)  
				$executionTrigger.StartBoundary = $startTime
  
				$taskAction = $taskDefinition.Actions.Create(0) 
				$taskAction.Path = "powershell.exe"
				$taskAction.Arguments = "-Command `"&'$installpath\New-ExchangeReport.ps1' -installpath '$installpath'`""

				$job = $rootFolder.RegisterTaskDefinition("Exchange-Reporter (www.FrankysWeb.de)" , $taskDefinition, 6, $username, $password, 1) 
			}
		if ($reportinterval -match 7)
			{
				write-host ""
				$zeitpunkt = read-host " Please specify Starttime (Example: 22:00)"
				$username = read-host " Please enter Task User (Domain\User)"
				$SecurePassword = read-host " Password" -AsSecureString
				$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecurePassword)
				$Password = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

				$startTime = "$zeitpunkt" | get-date -format s
				
				$taskService = New-Object -ComObject Schedule.Service 
				$taskService.Connect() 
				
				$rootFolder = $taskService.GetFolder($NULL) 
  
				$taskDefinition = $taskService.NewTask(0) 
  
				$registrationInformation = $taskDefinition.RegistrationInfo 
				$registrationInformation.Description = "Exchange Reporter Task - www.FrankysWeb.de"
				$registrationInformation.Author = $username
  
				$taskPrincipal = $taskDefinition.Principal 
				$taskPrincipal.LogonType = 1 
				$taskPrincipal.UserID = $username
				$taskPrincipal.RunLevel = 0 
  
				$taskSettings = $taskDefinition.Settings 
				$taskSettings.StartWhenAvailable = $true
				$taskSettings.RunOnlyIfNetworkAvailable = $true
				$taskSettings.Priority = 7 
  
				$taskTriggers = $taskDefinition.Triggers 
  
				$executionTrigger = $taskTriggers.Create(3)   
				$executionTrigger.DaysOfWeek = 2
				$executionTrigger.StartBoundary = $startTime
  
				$taskAction = $taskDefinition.Actions.Create(0) 
				$taskAction.Path = "powershell.exe"
				$taskAction.Arguments = "-Command `"&'$installpath\New-ExchangeReport.ps1' -installpath '$installpath'`""

				$job = $rootFolder.RegisterTaskDefinition("Exchange-Reporter (www.FrankysWeb.de)" , $taskDefinition, 6, $username, $password, 1) 
				
			}
			
		if ($reportinterval -notmatch 1 -and $reportinterval -notmatch 7)
			{
				write-host ""
				write-host " Reportinterval (settings.ini) is not 1 (Daily) or" -foregroundcolor yellow
				write-host " 7 (weekly), Could not create task, please do it manualy"  -foregroundcolor yellow
				write-host ""
				
				
			}
		
		if ($job)
			{
				write-host ""
				write-host " Task created!" -foregroundcolor green
				write-host ""
			}
	}
	
if ($aktion -match 2)
	{
		$job = invoke-command {SCHTASKS /Delete /TN "Exchange-Reporter (www.FrankysWeb.de)" /F} -ea 0 | out-null
		write-host ""
		write-host " Task deleted!" -foregroundcolor green
		write-host ""
	}
