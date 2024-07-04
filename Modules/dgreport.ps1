# Tage die eine Verteilerliste nicht benutzt wurde
#---------------------------------------------------
$unuseddays = 14
#---------------------------------------------------

$dgreport = Generate-ReportHeader "dgreport.png" "$l_dg_header"

$cells=@("$l_dg_name","$l_dg_email","$l_dg_member")
$dgreport += Generate-HTMLTable "$l_dg_t1header $unuseddays $l_dg_t1header2" $cells

$end = get-date 
$dgstart = $end.addDays(-$unuseddays)

if ($emsversion -match "2010")
	{
		$distributiongroups = Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySMTPAddress | Sort-Object PrimarySMTPAddress
		$counts = Get-TransportServer | Get-MessageTrackingLog -EventId Expand -ResultSize Unlimited -start $dgstart -end $end | Sort-Object RelatedRecipientAddress | Group-Object RelatedRecipientAddress | Sort-Object Name | Select-Object @{label="PrimarySmtpAddress"; expression={$_.Name}}, Count
	}
	
if ($emsversion -match "2013")
	{
		$distributiongroups = Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySMTPAddress | Sort-Object PrimarySMTPAddress
		$counts = Get-Transportservice | Get-MessageTrackingLog -EventId Expand -ResultSize Unlimited -start $dgstart -end $end | Sort-Object RelatedRecipientAddress | Group-Object RelatedRecipientAddress | Sort-Object Name | Select-Object @{label="PrimarySmtpAddress"; expression={$_.Name}}, Count
	}
	
if ($emsversion -match "2016" -or $emsversion -match "2019")
	{
		$distributiongroups = Get-DistributionGroup -ResultSize Unlimited | Select-Object PrimarySMTPAddress | Sort-Object PrimarySMTPAddress
		$counts = Get-Transportservice | Get-MessageTrackingLog -EventId Expand -ResultSize Unlimited -start $dgstart -end $end | Sort-Object RelatedRecipientAddress | Group-Object RelatedRecipientAddress | Sort-Object Name | Select-Object @{label="PrimarySmtpAddress"; expression={$_.Name}}, Count
	}

if ($distributiongroups -and $counts)
	{
		$unuseddls = Compare-Object $distributiongroups $counts -syncWindow 1000 -Property PrimarySmtpAddress -PassThru | Where-Object {$_.SideIndicator -eq '<='} | Select-Object -Property PrimarySmtpAddress | sort
	}

if ($unuseddls)
	{
		foreach ($unuseddl in $unuseddls)
			{
				[string]$smtpaddress = $unuseddl.primarysmtpaddress
				$dg = get-distributiongroup $smtpaddress -ResultSize Unlimited
				$dgname = $dg.displayname
				$members = Get-DistributionGroupMember -Identity $dg -ResultSize Unlimited | select name | foreach {$_.name}
				if ($members)
					{
						$hasmembers = "$l_dg_memberyes"
					}
				else
					{
						$hasmembers = "$l_dg_memberno"
					}
				$cells=@("$dgname","$smtpaddress","$hasmembers")
				$dgreport += New-HTMLTableLine $cells
			}
		
	}
else
	{
		$cells=@("$l_dg_nounuseddg")
	}

$dgreport += End-HTMLTable	

$dgreport| set-content "$tmpdir\dgreport.html"
$dgreport| add-content "$tmpdir\report.html"