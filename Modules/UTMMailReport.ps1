$UTMreport = Generate-ReportHeader "UTMMailreport.png" "$l_UTM_header"

$SophosUTMsettingshash = $inifile["SophosUTM-Mailreport"]
$SophosUTMsettings = convert-hashtoobject $SophosUTMsettingshash	

$UTMMailboxName = ($SophosUTMsettings | Where-Object {$_.Setting -eq "E-Mail-Address"}).Value
$UTMusername = ($SophosUTMsettings | Where-Object {$_.Setting -eq "Username"}).Value
$UTMpassword = ($SophosUTMsettings | Where-Object {$_.Setting -eq "Password"}).Value
$UTMdomain = ($SophosUTMsettings | Where-Object {$_.Setting -eq "Domain"}).Value

$apiload = Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)

$creds = New-Object System.Net.NetworkCredential("$UTMusername","$UTMpassword","$UTMdomain")
$service.Credentials = $creds  

#EWS URL Option 1 Autodiscover
$service.AutodiscoverUrl($UTMMailboxName,{$true})

#EWS URL Option 2 Hardcoded  
#$uri=[system.URI] "https://outlook.frankysweb.local/ews/Exchange.asmx"
#$service.Url = $uri  

$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,"$UTMMailboxName")   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)

$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)

$findItemsResults = $Inbox.FindItems($Sfha,$ivItemView)

$Sophosmail = $findItemsResults | where {$_.subject -match "INFO-721"} | select -first 1
$Sophosmail.Load()
$Sophosmail.Body.Text | set-content "$tmpdir\UTMreport.html"

$html = Get-Content "$tmpdir\UTMreport.html"  | out-string 
$extract = [regex]::match($html,'(?<=\<a name=\"mailsec\"\>).+(?=\<a name=\"appctrl\"\>)',"singleline").value

$extract = $extract.replace("Verdana","calibri")
$extract | set-content "$tmpdir\UTMreport.html"

(Get-Content "$tmpdir\UTMreport.html" | Select-Object -Skip 15) | set-content "$tmpdir\UTMreport.html"
(Get-Content "$tmpdir\UTMreport.html" | Select-Object -Skiplast 9) | set-content "$tmpdir\UTMreport.html"

$UTMreport | add-content "$tmpdir\report.html"
Get-Content "$tmpdir\UTMreport.html" | add-content "$tmpdir\report.html"