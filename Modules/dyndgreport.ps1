$dyndgreport = Generate-ReportHeader "dyndgreport.png" "$l_dyndg_header"

$cells=@("$l_dyndg_name","$l_dyndg_email","$l_dyndg_hasmember","$l_dyndg_membercount")
$dyndgreport += Generate-HTMLTable "$l_dyndg_t1header" $cells

$dyngroups = Get-DynamicDistributionGroup -resultsize unlimited | sort name
foreach ($dyngroup in $dyngroups)
	{
		$dyndgname = $dyngroup.name
		$dyndgmail = [string]$dyngroup.PrimarySmtpAddress
		$dyndgmembers = Get-Recipient -RecipientPreviewFilter $dyngroup.RecipientFilter -OrganizationalUnit $dyngroup.RecipientContainer -resultsize unlimited | sort name
		if ($dyndgmembers)
			{
				$dyndghasmember = "$l_dyndg_memberyes"
				$dynmemcount = $dyndgmembers.count
				
				$memcells=@("Name","Typ")
				$dyndgmemberreport += Generate-HTMLTable "$l_dyndg_t2header $dyndgname" $memcells
				
				foreach ($dyndgmember in $dyndgmembers)
					{
						$memname = $dyndgmember.name
						$memtyp = $dyndgmember.RecipientType
						
						$memcells=@("$memname","$memtyp")
						$dyndgmemberreport += New-HTMLTableLine $memcells
					}
				$dyndgmemberreport += End-HTMLTable
			}
		else
			{
				$dyndghasmember = "$l_dyndg_memberno"
				$dynmemcount = "0"
			}
		$cells=@("$dyndgname","$dyndgmail","$dyndghasmember","$dynmemcount")
		$dyndgreport += New-HTMLTableLine $cells
	}

$dyndgreport += End-HTMLTable

$dyndgreport += $dyndgmemberreport

$dyndgreport | set-content "$tmpdir\dyndgreport.html"
$dyndgreport | add-content "$tmpdir\report.html"