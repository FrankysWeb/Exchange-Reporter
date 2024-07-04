$redirectreport = Generate-ReportHeader "redirectreport.png" "$l_redir_header "

$cells=@("$l_redir_mbx","$l_redir_rulename","$l_redir_type","$l_redir_targetaddr","$l_redir_active")
$redirectreport += Generate-HTMLTable "$l_redir_header2" $cells

$rules = Get-Mailbox -resultsize unlimited | ForEach-Object {Get-InboxRule -Mailbox $PSItem.Id} | where {$_.forwardto -or $_.redirectto}
foreach ($rule in $rules)
	{
		$mbxname = $rule.mailboxownerid.name
		$rulename = $rule.name
		$ruleactive = $rule.Enabled
		if ($rule.forwardto -and !$rule.redirectto)
			{
				$type = "$l_redir_forward"
				$target = $rule.ForwardTo.displayname
			}
		if ($rule.redirectto -and !$rule.forwardto)
			{
				$type = "$l_redir_redir"
				$target = $rule.RedirectTo.displayname
			}
		if ($rule.redirectto -and $rule.forwardto)
			{
				$type = "$l_redir_forandredir"
				$target = $rule.RedirectTo.displayname
				$target += $rule.ForwardTo.displayname
			}
 
		$cells=@("$mbxname","$rulename","$type","$target","$ruleactive")
		$redirectreport += New-HTMLTableLine $cells
	} 
 
$redirectreport += End-HTMLTable

$cells=@("$l_redir_mbx","$l_redir_forwardto","$l_redir_targetaddr")
$redirectreport += Generate-HTMLTable "$l_redir_header3" $cells

$rules = get-mailbox -resultsize unlimited | where {$_.ForwardingAddress -ne $NULL} | Sort-Object -Property Name
foreach ($rule in $rules)
	{
		$mbxname = $rule.Name
		if ($rule.DeliverToMailboxAndForward -match "False")
			{
				$type = "$l_redir_onlytarget"
			}
		else
			{
				$type = "$l_redir_mbxandtarget"
			}
			
		$canname = $rule.ForwardingAddress
		$dn = ConvertFrom-Canonical $canname
		try 
			{
				$adobj = Get-ADObject $dn -Properties proxyaddresses
				[string]$target = ($adobj.proxyaddresses | Select-String SMTP -CaseSensitive:$true)
				$target = $target.replace("SMTP:","")
				if (!$target) {$target = "$dn"}
			}
		catch
			{
				$target = "$dn"
			}
		$cells=@("$mbxname","$type","$target")
		$redirectreport += New-HTMLTableLine $cells
	}
$redirectreport += End-HTMLTable
	
$redirectreport | add-content "$tmpdir\report.html"