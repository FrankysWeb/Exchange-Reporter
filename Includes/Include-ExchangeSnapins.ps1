# Die Funktion "Remove-WriteConsole" dient nur dazu die Ausgabe der Exchange Shell
# unterdrücken
#--------------------------------------------------------------------------------------

function Remove-WriteConsole
{
	[CmdletBinding(DefaultParameterSetName = 'FromPipeline')]
	param(
	[Parameter(ValueFromPipeline = $true, ParameterSetName = 'FromPipeline')]
	[object] $InputObject,

	[Parameter(Mandatory = $true, ParameterSetName = 'FromScriptblock', Position = 0)]
	[ScriptBlock] $ScriptBlock
	)

	begin
		{
			function Cleanup
				{
					remove-item function:\write-host -ea 0
					remove-item function:\write-verbose -ea 0
				}

			function ReplaceWriteConsole([string] $Scope)
				{
					Invoke-Expression "function ${scope}:Write-Host { }"
					Invoke-Expression "function ${scope}:Write-Verbose { }"
				}

			Cleanup

			if($pscmdlet.ParameterSetName -eq 'FromPipeline')
				{
					ReplaceWriteConsole -Scope 'global'
				}
		}

	process
		{
			if($pscmdlet.ParameterSetName -eq 'FromScriptBlock')
				{
					. ReplaceWriteConsole -Scope 'local'
					& $scriptblock
				}
			else
				{
					$InputObject
				}
		}

	end
		{
			Cleanup
		}  
}

# Lade Exchange Snapins und Verbinde zu Exchange Server
#--------------------------------------------------------------------------------------

if ((Get-PSSnapin | where {$_.name}) -notmatch "Microsoft.Exchange.Management.PowerShell")
	{
		if ($emsversion -match "2010")
			{
				$repspath = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v14\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
				$snapins = . $repspath | Remove-WriteConsole
				$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
			}
		if ($emsversion -match "2013")
			{
				$repspath = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
				$snapins = . $repspath | Remove-WriteConsole
				$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
			}
		if ($emsversion -match "2016")
			{
				$repspath = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
				$snapins = . $repspath | Remove-WriteConsole
				$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
			}
		if ($emsversion -match "2019")
			{
				$repspath = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
				$snapins = . $repspath | Remove-WriteConsole
				$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
			}
		if ($emsversion -notmatch "2010" -and $emsversion -notmatch "2013" -and $emsversion -notmatch "2016" -and $emsversion -notmatch "2019")
			{
				$version = (Get-ChildItem HKLM:\SOFTWARE\Microsoft\ExchangeServer\v1* -ea 0  | sort -Descending | select -first 1).pschildname
				$repspath = (get-itemproperty HKLM:\SOFTWARE\Microsoft\ExchangeServer\$version\Setup).MsiInstallPath + "bin\RemoteExchange.ps1"
				$snapins = . $repspath | Remove-WriteConsole
				$connect = Connect-ExchangeServer -auto | Remove-WriteConsole
			}
	}
