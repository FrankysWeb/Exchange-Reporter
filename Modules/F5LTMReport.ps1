$f5settingshash = $inifile["F5-BigIP-LTM"]
$f5settings = convert-hashtoobject $f5settingshash

$icontrol = Add-PSSnapin iControlSnapIn -ea 0

$F5LTMIP = ($f5settings| Where-Object {$_.Setting -eq "F5LTMIP"}).Value
$F5LTMUser = ($f5settings| Where-Object {$_.Setting -eq "F5LTMUser"}).Value
$F5LTMPass = ($f5settings | Where-Object {$_.Setting -eq "F5LTMPassword"}).Value
$F5VirtualServer = ($f5settings | Where-Object {$_.Setting -eq "F5VirtualServer"}).Value

$F5LTMPass = $F5LTMPass | ConvertTo-SecureString -AsPlainText -Force
$Creds= New-Object System.Management.Automation.PSCredential -ArgumentList $F5LTMUser , $F5LTMPass

$initicontrol = Initialize-F5.iControl -HostName $F5LTMIP -Credentials $Creds

$f5report = Generate-ReportHeader "f5ltmreport.png" "$l_ltm_header"

$cells=@("$l_ltm_name","$l_ltm_ip","$l_ltm_version","$l_ltm_hotfix","$l_ltm_edition","$l_ltm_uptime")
$f5report += Generate-HTMLTable "$l_ltm_t1header" $cells

$f5info = Get-F5.ProductInformation
 $f5version = $f5info.ProductVersion
 $f5hotfix = $f5info.PackageEdition
$f5sysinfo = Get-F5.SystemInformation
 $f5name = $f5sysinfo.Hostname
 $f5edition = $f5sysinfo.ProductCategory
$f5uptime = (Get-F5.SystemUptime) / 3600 / 24
 $f5uptime = [string]$f5uptime
 $f5uptime = $f5uptime.split(".")[0]

$cells=@("$f5name","$F5LTMIP","$f5version","$f5hotfix","$f5edition","$f5uptime")
$f5report += New-HTMLTableLine $cells

$f5report += End-HTMLTable

$cells=@("$l_ltm_vs","$l_ltm_avail","$l_ltm_vsactive","$l_ltm_state")
$f5report += Generate-HTMLTable "$l_ltm_t2header" $cells

$f5vservers = Get-F5.LTMVirtualServer
foreach ($f5vs in $f5vservers)
	{
		$vsname = $f5vs.name
		[string]$vsav = $f5vs.Availability
		$vsav = $vsav.split("_")[2]
		[string]$vsen = $f5vs.enabled
		$vsen = $vsen.split("_")[2]
		$vsstatus = $f5vs.status
		
		$cells=@("$vsname","$vsav","$vsen","$vsstatus")
		$f5report += New-HTMLTableLine $cells
		
	}
$f5report += End-HTMLTable
	
$cells=@("$l_ltm_poolname","$l_ltm_membercount","$l_ltm_avail","$l_ltm_vsactive","$l_ltm_state")
$f5report += Generate-HTMLTable "$l_ltm_t3header" $cells

$f5pools = Get-F5.LTMPool
foreach ($f5pool in $f5pools)
	{
		$poolname = $f5pool.name
		[string]$poolav = $f5pool.Availability
		$poolav = $poolav.split("_")[2]
		[string]$poolen = $f5pool.enabled
		$poolen = $poolen.split("_")[2]
		$poolstatus = $f5pool.status
		$poolmembercount = $f5pool.membercount
		
		$cells=@("$poolname","$poolmembercount","$poolav","$poolen","$poolstatus")
		$f5report += New-HTMLTableLine $cells
	}
$f5report += End-HTMLTable

$cells=@("$l_ltm_member","$l_ltm_port","$l_ltm_avail","$l_ltm_member","$l_ltm_state","$l_ltm_pool")
$f5report += Generate-HTMLTable "$l_ltm_t3header" $cells

foreach ($f5pool in $f5pools)
{
	$poolname = $f5pool.name
	$f5poolmembers = Get-F5.LTMPoolmember $poolname
		foreach ($f5poolmember in $f5poolmembers)
			{
				$membername = $f5poolmember.address
				$memberport = $f5poolmember.port
				[string]$memberav = $f5poolmember.Availability
				$memberav =$memberav.split("_")[2]
				[string]$memberen = $f5poolmember.enabled
				$memberen = $memberen.split("_")[2]
				$memberstatus = $f5poolmember.status
		
				$cells=@("$membername","$memberport","$memberav","$memberen","$memberstatus","$poolname")
				$f5report += New-HTMLTableLine $cells
			}
}
$f5report += End-HTMLTable

$f5report | set-content "$tmpdir\f5report.html"
$f5report | add-content "$tmpdir\report.html"