$updatereport = Generate-ReportHeader "updatereport.png" "$l_update_header"

[System.Net.ServicePointManager]::SecurityProtocol = @("Tls12","Tls11","Tls")

$onlineversion = Invoke-WebRequest "https://www.frankysweb.de/download/exchangereporter.txt"
$onlineversion = $onlineversion.content

$cells=@("$l_update_name","$l_update_instversion","$l_update_onlineversion")
$updatereport += Generate-HTMLTable "$l_update_exrep" $cells

$cells=@("$l_update_exrep","$reporterversion","$onlineversion")
$updatereport+= New-HTMLTableLine $cells
$updatereport += End-HTMLTable
 
foreach ($exserver in $exservers)
	{
		$servername = $exserver.name
		
		if ($env:Computername -match $exserver.name)
			{
				$Searcher = New-Object -ComObject Microsoft.Update.Searcher 
				$winupdates = $Searcher.Search("IsInstalled=0 and Type='Software'")
				$updates = $winupdates.updates | select title
			}
		else
		{
		$pssession = new-pssession -computername $servername
		if ($pssession)
			{
			$updates = invoke-command -session $pssession -scriptblock {
				$Searcher = New-Object -ComObject Microsoft.Update.Searcher 
				$winupdates = $Searcher.Search("IsInstalled=0 and Type='Software'")
				$winupdates.updates | select title} -ea 0
			}
		}
		
		$cells=@("Name")
		$updatereport += Generate-HTMLTable "$l_update_availupdates $servername" $cells

	if ($updates)
		{
			foreach ($update in $updates)
				{
					$winupdatename = $update.title
					$cells=@("$winupdatename")
					$updatereport+= New-HTMLTableLine $cells
				}
		}
	else
		{
			$cells=@("$l_update_noavailupdates")
			$updatereport+= New-HTMLTableLine $cells
		}
	$updatereport += End-HTMLTable
	}
	
#FW Updates
$feedsdownload = invoke-webrequest "https://www.frankysweb.de/feed/atom/" -outfile "$tmpdir\feedtemp.htm"
[xml]$feeds = Get-Content "$tmpdir\feedtemp.htm" -Encoding UTF8
$delfeedtemp = remove-item "$tmpdir\feedtemp.htm" -force

$postlist = @()
foreach ($feed in $feeds.feed.entry)
 {
	$pubdate = $feed.published | get-date
	$ptitle = $feed.title.innertext
	$url = $feed.id
	
	$postlist += new-object PSObject -property @{Published=$pubdate;Title="$ptitle";URL="$url"}
 }

$postlist = $postlist | sort published -Descending
$postlist = $postlist | where {$_.published -ge $start}

if ($postlist)
 {
	$cells=@("$l_update_date","$l_update_link","$l_update_title")
	$updatereport += Generate-HTMLTable "$l_update_header2" $cells

	foreach ($post in $postlist)
		{
			$published = $post.published | get-date -Format "dd.MM.yyyy"
			$link = "<a href=`"" + $post.url + "`">$l_update_link</a>" 
			$posttitle = $post.title

			$cells=@("$published","$link","$posttitle")
			$updatereport += New-HTMLTableLine $cells
		}

	$updatereport += End-HTMLTable
 }

$updatereport | add-content "$tmpdir\report.html"