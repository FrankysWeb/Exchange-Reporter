$Kempsettingshash = $inifile["Kemp-Loadmaster"]
$Kempsettings = convert-hashtoobject $Kempsettingshash

$kempmodule = Import-Module -Name Kemp.LoadBalancer.Powershell

$KempIP = ($Kempsettings| Where-Object {$_.Setting -eq "KempLMIP"}).Value
$KempUser = ($Kempsettings| Where-Object {$_.Setting -eq "KempLMUser"}).Value
$KempPass = ($Kempsettings | Where-Object {$_.Setting -eq "KempLMPassword"}).Value

$KempPass = $KempPass | ConvertTo-SecureString -AsPlainText -Force
$KempCreds= New-Object System.Management.Automation.PSCredential -ArgumentList $KempUser , $KempPass

try { 
 $initlb = Initialize-LmConnectionParameters -Address $KempIP -LBPort 443 -Credential $KempCreds
}
catch {
 $error[0] | add-content "$installpath\ErrorLog.txt"
 exit 0
}

$kempreport = Generate-ReportHeader "kempreport.png" "$l_kemp_header"

$cells=@("$l_kemp_Nickname","$l_kemp_VSIP","$l_kemp_VSStatus")
$kempreport += Generate-HTMLTable "$l_kemp_vsheader" $cells

$vservers = (Get-AdcVirtualService).data.vs | sort vsaddress
foreach ($vserver in $vservers)
	{
		$nickname = $vserver.nickname
		$vsip = $vserver.VSAddress
		
		if (!$vsip)
			{
				$mastervs = $vserver.MasterVS
				$vsip = ($vservers | where {$_.MasterVS -eq $mastervs}).VSAddress
			}
		
		$vsstate = $vserver.status
		
		if ($vsstate -match "Unchecked") 
			{
				$vsstatus = "<font color=`"#E59400`">" + "$l_kemp_Unchecked" + "</font>"
			}
		if ($vsstate -match "Down") 
			{
				$vsstatus = "<font color=`"#CD0000`">" + "$l_kemp_Down" + "</font>"
			}
		if ($vsstate -match "Up") 
			{
				$vsstatus = "<font color=`"#008B00`">" + "$l_kemp_Up" + "</font>"
			}
		

		$cells=@("$nickname","$vsip","$vsstatus")
		$kempreport += New-HTMLTableLine $cells
	}
	
$kempreport += End-HTMLTable

$cells=@("$l_kemp_rsip","$l_kemp_rsport","$l_kemp_RSStatus","$l_kemp_Limit","$l_kemp_weight","$l_kemp_active")
$kempreport += Generate-HTMLTable "$l_kemp_rsoverview" $cells

$VSIndexes = $vservers.index
foreach ($VSIndex in $VSIndexes)
	{
		$rserver = (Get-AdcRealServer -VSIndex $VSIndex).data.rs
		$rsnickname = $rserver.DnsName

			$rsid = $rserver.RsIndex
			$rsstate = $rserver.Status
			$rsip = $rserver.Addr
			$rsport = $rserver.Port
			$rslimit = $rserver.Limit
			$rsactive = $rserver.Enable
			$rscritical = $rserver.Critical
			$rsweight = $rserver.Weight
			
			if ($rsstate -match "Unchecked") 
			{
				$rsstatus = "<font color=`"#E59400`">" + "$l_kemp_Unchecked" + "</font>"
			}
			if ($rsstate -match "Down") 
			{
				$rsstatus = "<font color=`"#CD0000`">" + "$l_kemp_Down" + "</font>"
			}
			if ($rsstate -match "Up") 
			{
				$rsstatus = "<font color=`"#008B00`">" + "$l_kemp_Up" + "</font>"
			}
			
			if ($rsactive -match "Y")
				{
					$rsactive = $l_kemp_yes
				}
			if ($rsactive -match "N")
				{
					$rsactive = $l_kemp_no
				}
			
			if ($rsip)
				{
					$cells=@("$rsip","$rsport","$rsstatus","$rslimit","$rsweight","$rsactive")
					$kempreport += New-HTMLTableLine $cells
				}
	}

$kempreport += End-HTMLTable

$kempreport | set-content "$tmpdir\kempreport.html"
$kempreport | add-content "$tmpdir\report.html"