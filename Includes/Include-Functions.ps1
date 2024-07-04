#HTML Tabellen
#--------------------------------------------------------------------------------------

# HTML Header
function New-HTMLHeader ($title)
 {
  if ($language -match "DE")
  {
  $HTMLHeader = "
   <html>
	<head>
	 <title>$title</title>
	</head>
	<body>
    <div align=`"center`">
  <table style=`"width: 90%; text-align: left; margin-left: auto; margin-right: auto; border-collapse:collapse; font-family:calibri;`" border=`"0`" cellpadding=`"0`" cellspacing=`"0`">
  <tr>
  <td style=`"background-color: `#0072C6; color: `#ffffff; font-weight: bold; border:solid `#0072C6; border-width: 3px 3px 3px 3px; width: 100px`"><img src=`"cid:reportheader.png`" alt=`"picture`"></td>
  <td style=`"background-color: `#0072C6; color: `#ffffff; font-weight: bold; border:solid `#0072C6; border-width: 3px 3px 3px 3px; font-size: 20px`">Exchange Report</td>
  </tr>
  </table>
    <table align=`"center`" cellspacing=`"0`" style=`"width: 80%`">
     <td style=`"background-color: `#F8F8F8; color: `#585858;`">
	 
	 Exchange Report für $orgname<br>
	 Datum: $today<br>
     erstellt von Exchange Reporter<br>
     Version: $reporterversion<br>
		www.frankysweb.de
	 
	 </tr></td>
   </table>
   &nbsp;
   "#"
   return $HTMLHeader
   }
   else
   {
     $HTMLHeader = "
   <html>
	<head>
	 <title>$title</title>
	</head>
	<body>
    <div align=`"center`">
  <table style=`"width: 90%; text-align: left; margin-left: auto; margin-right: auto; border-collapse:collapse; font-family:calibri;`" border=`"0`" cellpadding=`"0`" cellspacing=`"0`">
  <tr>
  <td style=`"background-color: `#0072C6; color: `#ffffff; font-weight: bold; border:solid `#0072C6; border-width: 3px 3px 3px 3px; width: 100px`"><img src=`"cid:reportheader.png`" alt=`"picture`"></td>
  <td style=`"background-color: `#0072C6; color: `#ffffff; font-weight: bold; border:solid `#0072C6; border-width: 3px 3px 3px 3px; font-size: 20px`">Exchange Report</td>
  </tr>
  </table>
    <table align=`"center`" cellspacing=`"0`" style=`"width: 80%`">
     <td style=`"background-color: `#F8F8F8; color: `#585858;`">
	 
	 Exchange Report for $orgname<br>
	 Date: $today<br>
     created with Exchange Reporter<br>
     Version: $reporterversion<br>
		www.frankysweb.de
	 
	 </tr></td>
   </table>
   &nbsp;
   "#"
   return $HTMLHeader
   }
 }

# Neue Tabelle inkl Überschrift
function Generate-HTMLTable ($headline, $cells)
 {
  $HTMLTable = "
  <h3 style=`"text-align:center; font-family:calibri; color: `#0072C6;`">$headline</h3>
<table style=`"width: 80%; text-align: left; margin-left: auto; margin-right: auto; border-collapse:collapse; font-family:calibri;`" border=`"0`" cellpadding=`"0`" cellspacing=`"0`">
	<tr>"#"
     foreach ($cell in $cells)
	  {
	   $HTMLTable += "<td style=`"background-color: `#0072C6; color: `#ffffff; font-weight: bold; border:solid `#0072C6; border-width: 3px 3px 3px 3px;`">$cell</td>"
	  }
  $HTMLTable += "</tr>"
	return $HTMLTable
 }

# Neue Tabellenzeile
function New-HTMLTableLine ($cells)
 {
  $NewTableLine = "<tr>"
     foreach ($cell in $cells)
	  {
	   $NewTableLine += "<td style=`"background-color: `#F8F8F8; color: `#585858; border:solid `#0072C6; border-width: 1px 1px 1px 1px;`">$cell</td>"
	  }
  $NewTableLine += "</tr>"
	return $NewTableLine
 }

# Tabellenende
function End-HTMLTable ()
 {
  $TableEnd = "</table>"
  return $TableEnd
 }
 
# HTML Ende
function End-HTML ()
 {
  $HTMLEnd = "
   </body>
   </html>
   "
   return $HTMLEnd
 }
 
# HTML Inline Pictures
function Include-HTMLInlinePictures ($picid)
{
 $htmlreport = "
  <table align=`"center`" cellspacing=`"0`" style=`"width: 80%`">
   <tr><td><center>
  "
   
  $Pics = Get-ChildItem "$picid" -name
   Foreach ($pic in $Pics) 
   {
    $htmlreport += "<img src=`"cid:$pic`" alt=`"picture`">"
   }

   $htmlreport += "  
    </center></tr></td>
   </table>
   "
  
  return $HTMLReport
}

# Report Header
function Generate-ReportHeader ($Image, $Headline)
 {
  $ReportHeader = "<br>"
  $ReportHeader += "<table style=`"width: 90%; text-align: left; margin-left: auto; margin-right: auto; border-collapse:collapse; font-family:calibri;`" border=`"0`" cellpadding=`"0`" cellspacing=`"0`">"
  $ReportHeader += "<tr>
  <td style=`"background-color: `#0072C6; color: `#ffffff; font-weight: bold; border:solid `#0072C6; border-width: 3px 3px 3px 3px; width: 100px`"><img src=`"cid:$image`" alt=`"picture`"></td>
  <td style=`"background-color: `#0072C6; color: `#ffffff; font-weight: bold; border:solid `#0072C6; border-width: 3px 3px 3px 3px; font-size: 20px`">$headline</td>
  </tr>
  </table>
  " #"
	return $ReportHeader
 }

# Report Footer
function Generate-ReportFooter ()
 {
 if ($language -match "DE")
 {
 $footer = "
 &nbsp;
 <table align=`"center`" cellspacing=`"0`" style=`"width: 80%`"><tr>
     <td style=`"background-color: `#F8F8F8; color: `#585858;`">
	 
     erstellt von Exchange Reporter $reporterversion<br>
	 Frank Zöchling www.FrankysWeb.de
	
	 </td>
	 </tr>
	 <tr>
	 <td style=`"background-color: `#F8F8F8; color: `#585858;`">
	 
	 <center>
	 
	 <hr>
	 Spenden Sie etwas für die Weiterentwicklung, wenn Ihnen Exchange Reporter gefällt:<br>
	 <a href=`"https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=frank%40frankysweb%2ede&lc=DE&no_note=0&currency_code=EUR&bn=PP%2dDonationsBF%3abtn_donate_SM%2egif%3aNonHostedGuest`">Spenden via Paypal</a>
	 
	 </center>
	 </td>
	 </tr>
   </table>
   &nbsp;
   "#"
   return $footer
   }
   else
   {
    $footer = "
     &nbsp;
     <table align=`"center`" cellspacing=`"0`" style=`"width: 80%`"><tr>
     <td style=`"background-color: `#F8F8F8; color: `#585858;`">
	 
     created with Exchange Reporter $reporterversion<br>
	 Frank Zoechling www.FrankysWeb.de
	
	 </td>
	 </tr>
	 <tr>
	 <td style=`"background-color: `#F8F8F8; color: `#585858;`">
	 
	 <center>
	 
	 <hr>
	 If you like Exchange Reporter, please donate:<br>
	 <a href=`"https://www.paypal.com/cgi-bin/webscr?cmd=_donations&business=frank%40frankysweb%2ede&lc=DE&no_note=0&currency_code=EUR&bn=PP%2dDonationsBF%3abtn_donate_SM%2egif%3aNonHostedGuest`">Donate with Paypal</a>
	 
	 </center>
	 </td>
	 </tr>
   </table>
   &nbsp;
   "#"
   return $footer
   }
 }

#Grafiken
#--------------------------------------------------------------------------------------

# new-piechart
function new-piechart ($width, $height, $charttitle, $data, $pngpath)
{
 $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
 $Chart.Width = $width
 $Chart.Height = $height
 $Chart.Left = 40 
 $Chart.Top = 30
 $Chart.Palette = 'brightpastel'
 $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
 $Chart.ChartAreas.Add($ChartArea)
 [void]$Chart.Titles.Add("$charttitle") 
 [void]$Chart.Series.Add("Data") 
 $Chart.Series["Data"].Points.DataBindXY($data.Keys, $data.Values)
 $Chart.Series["Data"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::Doughnut
 $Chart.Series["Data"]["PieLabelStyle"] = "Inside" 
 $Chart.Series["Data"]["PieLineColor"] = "Black"
 $Chart.SaveImage("$pngpath", "PNG")
}

# new-cylinderchart 

function new-cylinderchart ($width, $height, $charttitle, $xtitle, $ytitle, $data, $pngpath)
{
 $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart 
 $Chart.Width = $width
 $Chart.Height = $height
 $Chart.Left = 40
 $Chart.Top = 30
 $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea 
 $Chart.ChartAreas.Add($ChartArea)
 [void]$Chart.Titles.Add("$charttitle")
 $ChartArea.AxisX.Title = $xtitle 
 $ChartArea.AxisX.Interval = 1
 $ChartArea.AxisY.Title = $ytitle
 [void]$Chart.Series.Add("Data")
 $chart.Series["Data"].color = "#0072C6"
 $Chart.Series["Data"].Points.DataBindXY($data.Keys, $data.Values)
 $Chart.SaveImage("$pngpath", "PNG")
}

# New-StackedColumnChart
function New-StackedColumnChart
{
    param($width, $height, $charttitle, $DataSeries1, $DataSeries2, $pngpath, $xtitle, $ytitle, $mode)

    [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.DataVisualization") | Out-Null
    $Chart = New-object System.Windows.Forms.DataVisualization.Charting.Chart
    $Chart.Width = $width
    $Chart.Height = $height
    $Chart.Left = 40
    $Chart.Top = 30
    $ChartArea = New-Object System.Windows.Forms.DataVisualization.Charting.ChartArea
    $Chart.ChartAreas.Add($ChartArea)
    [void]$Chart.Titles.Add("$charttitle")
    $ChartArea.AxisX.Title = $xtitle
    $ChartArea.AxisX.Interval = 1
    $ChartArea.AxisY.Title = $ytitle
    [void]$Chart.Series.Add("DataSeries1")
    $Chart.Series["DataSeries1"].Points.DataBindXY($DataSeries1.Keys, $DataSeries1.Values)
    if ( $mode -eq 0 ) { $chart.Series["DataSeries1"].color = "#0072C6" } #no limit (blue)
    else { $chart.Series["DataSeries1"].color = "#59bd1f" } #limited (green)
    $Chart.Series["DataSeries1"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::StackedColumn
    [void]$Chart.Series.Add("DataSeries2")
    $Chart.Series["DataSeries2"].Points.DataBindXY($DataSeries2.Keys, $DataSeries2.Values)
    $Chart.Series["DataSeries2"].Color = "Red"
    $Chart.Series["DataSeries2"].ChartType = [System.Windows.Forms.DataVisualization.Charting.SeriesChartType]::StackedColumn
    $Chart.SaveImage("$pngpath", "PNG")
}

# read INI-File function
#--------------------------------------------------------------------------------------

function Get-IniContent ($filePath)
{
    $ini = @{}
    switch -regex -file $FilePath
    {
        "^\[(.+)\]" # Section
        {
            $section = $matches[1]
            $ini[$section] = @{}
            $CommentCount = 0
        }
        "^(;.*)$" # Comment
        {
            $value = $matches[1]
            $CommentCount = $CommentCount + 1
            $name = "Comment" + $CommentCount
            $ini[$section][$name] = $value
        } 
        "(.+?)\s*=(.*)" # Key
        {
            $name,$value = $matches[1..2]
            $ini[$section][$name] = $value
        }
    }
    return $ini
}

# Hashtable to Object converter function
#--------------------------------------------------------------------------------------

function convert-hashtoobject ($hash)
{
$newobject = @()
ForEach ($Key in $hash.Keys) {
   $Obj =  New-Object PSObject
   Add-Member -InputObject $Obj -Name 'Setting' -Value $Key -MemberType NoteProperty
   Add-Member -InputObject $Obj -Name 'Value' -Value $hash.$Key -MemberType NoteProperty
   $newobject += $Obj	
}
return $newobject
}

# Ermittle Exchange Managment Shell Version by Registry
#--------------------------------------------------------------------------------------

function Get-ExchangeVersionByRegistry ()
	{
		$version = (Get-ExchangeServer | where {$_.serverrole -match "Mailbox"} | sort admindisplayversion | select AdminDisplayVersion -first 1).AdminDisplayVersion
		$majorversion = $version.major
		$minorversion = $version.minor
		
		if ($majorversion -match "14")
			{
				$version = "2010"
			}
		if ($majorversion -match "15" -and $minorversion -match "0")
			{
				$version = "2013"
			}
		if ($majorversion -match "15" -and $minorversion -match "1")
			{
				$version = "2016"
			}
		if ($majorversion -match "15" -and $minorversion -match "2")
			{
				$version = "2019"
			}
		return $version
	}

# Date Converter
#--------------------------------------------------------------------------------------	

function convert-date
{

   [cmdletbinding()]
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline=$true)]
        $str
    )

try 
  {
	if ($language -match "DE")
		{
			[string]$result = get-date $str -UFormat %d.%m.%Y
		}
	if ($language -match "EN")
		{
			[string]$result = get-date $str -UFormat %m/%d/%Y
		}
	if ($language -notmatch "EN" -and $language -notmatch "DE")
		{
			[string]$result = get-date $str -UFormat %m/%d/%Y
		}
	
  }
Catch [system.exception]
  {
   $result = $str
  }
Finally
  {
   #
  }
return $result
}

# ConvertFrom-Canonical Canonical Name Converter
#--------------------------------------------------------------------------------------	

function ConvertFrom-Canonical
{
param([string]$canoincal=(trow '$Canonical is required!'))
    $obj = $canoincal.Replace(',','\,').Split('/')
    [string]$DN = "CN=" + $obj[$obj.count - 1]
    for ($i = $obj.count - 2;$i -ge 1;$i--){$DN += ",OU=" + $obj[$i]}
    $obj[0].split(".") | ForEach-Object { $DN += ",DC=" + $_}
    return $dn
}