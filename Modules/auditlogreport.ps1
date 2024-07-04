$auditlogreport = Generate-ReportHeader "auditlogreport.png" "$l_audit_header"

$cells=@("$l_audit_objmod","$l_audit_caller","$l_audit_cmdlet","$l_audit_parameter","$l_audit_propmod","$l_audit_date")
$auditlogreport += Generate-HTMLTable "$l_audit_header2" $cells

$adminauditlog = Search-AdminAuditLog -StartDate $start -EndDate $end | where {$_.Succeeded -match "True"}
foreach ($event in $adminauditlog)
	{
		try 
			{
				$ObjectModified = (get-aduser (ConvertFrom-Canonical $event.ObjectModified)).Name
			}
		catch
			{
				$ObjectModified = $event.ObjectModified
			}
		try
			{
				$Caller = (get-aduser (ConvertFrom-Canonical $event.caller)).Name
			}
		catch
			{
				$Caller = $event.caller
			}
		$CmdletName = $event.CmdletName
		$CmdLetparameter = $event.CmdletParameters
		$ModifiedProperties = $event.ModifiedProperties
		$rundate = $event.rundate
		
		$cells=@("$ObjectModified","$Caller","$CmdletName","$CmdLetparameter","$ModifiedProperties","$rundate")
		$auditlogreport += New-HTMLTableLine $cells
		
	}

$auditlogreport += End-HTMLTable
	
$auditevents = @()
$mailboxaudits = Search-MailboxAuditLog
foreach ($mailboxaudit in $mailboxaudits)
	{
		$mailboxid = $mailboxaudit.Identity
		
		$auditlogentries = @()
		$auditlogentries = get-mailbox "$mailboxid" | Search-MailboxAuditLog -LogonTypes Delegate -StartDate $start -enddate $end -ShowDetails

		if ($($auditlogentries.Count) -gt 0)
			{
				foreach ($entry in $auditlogentries)
				{
					$reportObj = New-Object PSObject
					$reportObj | Add-Member NoteProperty -Name "Mailbox" -Value $entry.MailboxResolvedOwnerName
					$reportObj | Add-Member NoteProperty -Name "MailboxUPN" -Value $entry.MailboxOwnerUPN
					$reportObj | Add-Member NoteProperty -Name "Timestamp" -Value $entry.LastAccessed
					$reportObj | Add-Member NoteProperty -Name "AccessedBy" -Value $entry.LogonUserDisplayName
					$reportObj | Add-Member NoteProperty -Name "Operation" -Value $entry.Operation
					$reportObj | Add-Member NoteProperty -Name "Result" -Value $entry.OperationResult
					$reportObj | Add-Member NoteProperty -Name "Folder" -Value $entry.FolderPathName
						if ($entry.ItemSubject)
							{
								$reportObj | Add-Member NoteProperty -Name "Subject Lines" -Value $entry.ItemSubject
							}
							else
							{
								$reportObj | Add-Member NoteProperty -Name "Subject Lines" -Value $entry.SourceItemSubjectsList
							}

					$auditevents += $reportObj
				}
				
			}
	}

$cells=@("$l_audit_Mailbox","$l_audit_UPN","$l_audit_Timestamp","$l_audit_AccessedBy","$l_audit_Operation","$l_auditResult","$l_audit_Folder")
$auditlogreport += Generate-HTMLTable "$l_audit_header3" $cells

if ($auditevents)
	{
		foreach ($auditevent in $auditevents)
			{
				$mailbox = $auditevent.Mailbox
				$mailboxupn = $auditevent.MailboxUPN
				$timestamp = $auditevent.Timestamp
				$accessedby = $auditevent.AccessedBy
				$operation = $auditevent.Operation
				$result = $auditevent.result
				$folder = $auditevent.folder
		
				$cells=@("$Mailbox","$mailboxupn","$timestamp","$accessedby","$operation","$result","$folder")
				$auditlogreport += New-HTMLTableLine $cells
			}
	}
else
	{
		$cells=@("$l_audit_noevent")
		$auditlogreport += New-HTMLTableLine $cells
	}
$auditlogreport += End-HTMLTable
	
$auditlogreport | set-content "$tmpdir\auditlogreport.html"
$auditlogreport | add-content "$tmpdir\report.html"