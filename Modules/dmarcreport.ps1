$DMARCUseExchangeDefaultDomain = $true

#You can specify your Domains here, if you don't wan't to use Exchange Default Accepted Domain (set $DMARCUseExchangeDefaultDomain = $false)
$CustomDMARCDomains = @(
    'microsoft.com'
    'google.de'
    )

#--------------------------------------------------------------

$DMARCsettingshash = $inifile["DMARC"]
$DMARCsettings = convert-hashtoobject $DMARCsettingshash    

$MailboxName = ($DMARCsettings | Where-Object {$_.Setting -eq "RUA-Address"}).Value
$username = ($DMARCsettings | Where-Object {$_.Setting -eq "Username"}).Value
$password = ($DMARCsettings | Where-Object {$_.Setting -eq "Password"}).Value
$domain = ($DMARCsettings | Where-Object {$_.Setting -eq "Domain"}).Value
$folder = ($DMARCsettings | Where-Object {$_.Setting -eq "ArchiveFolder"}).Value

$dmarcreport = Generate-ReportHeader "dmarcreport.png" "$l_dmarc_header"

$cells=@("$l_dmarc_domain","$l_dmarc_entryfordomain")
$dmarcreport += Generate-HTMLTable "$l_dmarc_header2" $cells

if ($DMARCUseExchangeDefaultDomain -eq $True)
{
    $dmarcdomainnames = (Get-AcceptedDomain | where {$_.Default -eq "True"}).Domainname.Domain
}
else 
{
    $dmarcdomainnames = $Customdmarcdomains
}

foreach ($dmarcdomainname in $dmarcdomainnames)
    {
        $dmarcdnsname = "_dmarc." + "$dmarcdomainname"
        $dnsentry = Resolve-DnsName $dmarcdnsname <#new# -server 8.8.8.8 #/new#> -Type TXT -ea 0
        $dnsdmarcentry = $dnsentry.strings
        
        $cells=@("$dmarcdnsname","$dnsdmarcentry")
        $dmarcreport += New-HTMLTableLine $cells
    }
$dmarcreport += End-HTMLTable

#Get DMARC Reports mailbox

Add-Type -Path "C:\Program Files\Microsoft\Exchange\Web Services\2.2\Microsoft.Exchange.WebServices.dll"

$ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2010_SP2
$service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)


$creds = New-Object System.Net.NetworkCredential("$username","$password","$domain")
$service.Credentials = $creds  

#CAS URL Option 1 Autodiscover
$service.AutodiscoverUrl($MailboxName,{$true})

#CAS URL Option 2 Hardcoded  
#$uri=[system.URI] "https://outlook.frankysweb.local/ews/Exchange.asmx"
#$service.Url = $uri  

$Sfha = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.EmailMessageSchema]::HasAttachments, $true)
$folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,"$MailboxName")   
$Inbox = [Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$folderid)  

$ivItemView = New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)
$downloadDirectory = "$installpath\Temp"

$findItemsResults = $Inbox.FindItems($Sfha,$ivItemView)
foreach($miMailItems in $findItemsResults.Items){
    $miMailItems.Load()
    foreach($attach in $miMailItems.Attachments){
        if ($attach.name -match "zip"<#new#> -or $attach.name -match "gz"<#/new#>)
            {
                $attach.Load()
                $fiFile = new-object System.IO.FileStream(($downloadDirectory + "\" + $attach.Name.ToString()), [System.IO.FileMode]::Create)
                $fiFile.Write($attach.Content, 0, $attach.Content.Length)
                $fiFile.Close()
            }
    }
}
  
$fvFolderView =  New-Object Microsoft.Exchange.WebServices.Data.FolderView(1000)  
$fvFolderView.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Shallow;
$SfSearchFilter = new-object Microsoft.Exchange.WebServices.Data.SearchFilter+IsEqualTo([Microsoft.Exchange.WebServices.Data.FolderSchema]::DisplayName,"$folder")
$findFolderResults = $Inbox.FindFolders($SfSearchFilter,$fvFolderView)  
    
$ivItemView =  New-Object Microsoft.Exchange.WebServices.Data.ItemView(1000)  
$fiItems = $null    
do{    
    $fiItems = $Inbox.FindItems($Sfha,$ivItemView)   
        foreach($Item in $fiItems.Items){      
 
            $move = $Item.Move($findFolderResults.Folders[0].Id)  
        }    
        $ivItemView.Offset += $fiItems.Items.Count    
    }while($fiItems.MoreAvailable -eq $true)  

#Extract ZIP Files

Add-Type -AssemblyName System.IO.Compression.FileSystem
function Unzip-File
{
    param([string]$zipfile, [string]$outpath)
    try 
        {
            [System.IO.Compression.ZipFile]::ExtractToDirectory($zipfile, $outpath)
        }
    catch
        {
        }
}

$zipfiles =  Get-ChildItem "$Installpath\Temp\*.zip" -Recurse | Where {-NOT $_.PSIsContainer} | foreach {$_.fullname}
foreach ($zipfile in $zipfiles)
    {
        $unzip = Unzip-File $zipfile "$Installpath\Temp"
    }

<#new#>
#Extract GZ Files

Function DeGZip-File{
    Param(
        $infile,
        $outfile       
        )

    $input = New-Object System.IO.FileStream $inFile, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::Read)
    $output = New-Object System.IO.FileStream $outFile, ([IO.FileMode]::Create), ([IO.FileAccess]::Write), ([IO.FileShare]::None)
    $gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)

    $buffer = New-Object byte[](1024)
    while($true){
        $read = $gzipstream.Read($buffer, 0, 1024)
        if ($read -le 0){break}
        $output.Write($buffer, 0, $read)
        }

    $gzipStream.Close()
    $output.Close()
    $input.Close()
}

$gzfiles =  Get-ChildItem "$Installpath\Temp\*.gz" -Recurse | Where {-NOT $_.PSIsContainer} | foreach {$_.fullname}
foreach ($gzfile in $gzfiles)
    {
        $gzoutfile = $gzfile -replace '.xml.gz','.xml'
        DeGZip-File $gzfile $gzoutfile
    }
<#/new#>

#Getting DMARC Entrys vom XML Files
    
$dmarcobject = @()
$xmlfiles =  Get-ChildItem "$Installpath\Temp\*.xml" -Recurse | Where {-NOT $_.PSIsContainer} | foreach {$_.fullname}
foreach ($xmlfile in $xmlfiles)
    {
        $xmldata = [xml](Get-Content $xmlfile)

        $dmarcorgname = $xmldata.feedback.report_metadata.org_name
        $dmarcmail    = $xmldata.feedback.report_metadata.email
        $dmarcentrys  = $xmldata.feedback.record
        
        foreach ($dmarcentry in $dmarcentrys)
            {
                $dmarcdom = $dmarcentry.identifiers.header_from
                <#new#>
                $dmarcauthdkimdomain = $dmarcentry.auth_results.dkim.domain
                $dmarcauthdkimresult = $dmarcentry.auth_results.dkim.result
                $dmarcauthspfdomain = $dmarcentry.auth_results.spf.domain
                $dmarcauthspfresult = $dmarcentry.auth_results.spf.result
                if ($dmarcdom -eq $dmarcauthdkimdomain) { $dmarcalignmentdkim = "pass" } else { $dmarcalignmentdkim = "fail" }
                if ($dmarcdom -eq $dmarcauthspfdomain) { $dmarcalignmentspf = "pass" } else { $dmarcalignmentspf = "fail" }
                <#/new#>
            
                $dmarcrows = $dmarcentry.row
                foreach ($dmarcrow in $dmarcrows)
                {
                    $dmarcip = $dmarcrow.source_ip
                    $dmarcipcount = $dmarcrow.count
                    
                    <#new#>
                    $dmarcpolicy_evaluateddkim = $dmarcrow.policy_evaluated.dkim
                    $dmarcpolicy_evaluatedspf = $dmarcrow.policy_evaluated.spf

                    [array]$dmarcobject  += new-object PSObject -property @{Organisation=$dmarcorgname;OrgMail=$dmarcmail;Domain=$dmarcdom;IP=$dmarcip;Count=$dmarcipcount;DKIMD=$dmarcauthdkimdomain;DKIMR=$dmarcauthdkimresult;DKIMA=$dmarcalignmentdkim;DKIMP=$dmarcpolicy_evaluateddkim;SPFD=$dmarcauthspfdomain;SPFR=$dmarcauthspfresult;SPFA=$dmarcalignmentspf;SPFP=$dmarcpolicy_evaluatedspf} 
                    #<#/new#>[array]$dmarcobject  += new-object PSObject -property @{Organisation=$dmarcorgname;OrgMail=$dmarcmail;Domain=$dmarcdom;IP=$dmarcip;Count=$dmarcipcount}
    
                }
            }
    }

$cells=@("$l_dmarc_orgname","$l_dmarc_orgmailaddr","$l_dmarc_domain","$l_dmarc_ip","$l_dmarc_count"<#new#>,"DKIM Domain","DKIM Authentication","DKIM Alignment","DKIM Policy","SPF Domain","SPF Authentication","SPF Alignment","SPF Policy"<#/new#>)
$dmarcreport += Generate-HTMLTable "$l_dmarc_header3" $cells

if ($dmarcobject)
{
    foreach ($dmarcentry in $dmarcobject)
        {
            $orgname = $dmarcentry.organisation
            $orgmail = $dmarcentry.orgmail
            $count = $dmarcentry.count
            $ip = $dmarcentry.ip
            $domain = $dmarcentry.domain
            $ip
            <#new#>
            $dkimd = $dmarcentry.dkimd
            $dkimr = $dmarcentry.dkimr
            $dkima = $dmarcentry.dkima
            $dkimp = $dmarcentry.dkimp
            $spfd = $dmarcentry.spfd
            $spfr = $dmarcentry.spfr
            $spfa = $dmarcentry.spfa
            $spfp = $dmarcentry.spfp
            $cells=@("$orgname","$orgmail","$domain","$ip","$count","$dkimd","$dkimr","$dkima","$dkimp","$spfd","$spfr","$spfa","$spfp")
            #<#/new#>$cells=@("$orgname","$orgmail","$domain","$ip","$count")
            $dmarcreport += New-HTMLTableLine $cells
        }
}
else
    {
        $cells=@("$l_dmarc_noentry")
        $dmarcreport += New-HTMLTableLine $cells
    }


$dmarcreport += End-HTMLTable

$dmarcreport | set-content "$tmpdir\dmarcreport.html"
$dmarcreport | add-content "$tmpdir\report.html" 
