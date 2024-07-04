# Module for displaying the current status of client access licences
# Leslie Heinz 2020-10-17 v1.1.2

#Table
$LicenceReport = Generate-ReportHeader "LicenceReport.png" "$l_licence_header"

$LicenceSettingshash = $inifile["Licence"]
$Licencesettings = convert-hashtoobject $LicenceSettingshash

$OwnedStdCAL = ($Licencesettings | Where-Object {$_.Setting -eq "StdCAL"}).Value
$OwnedEntCAL = ($Licencesettings | Where-Object {$_.Setting -eq "EntCAL"}).Value

if ( $OwnedStdCAL -eq 0 -and $OwnedEntCAL -eq 0 ) { $mode = 0 } #no licence limitation
else { $mode = 1 }

$EntCAL=(Get-ExchangeServerAccessLicenseUser -LicenseName (Get-ExchangeServerAccessLicense |
    Where-Object {($_.UnitLabel -eq "CAL") -and ($_.LicenseName -like "*Enterprise*")}).licenseName -WarningAction silentlyContinue |
    Measure-Object | Select-Object Count).count

$StdCAL=(Get-ExchangeServerAccessLicenseUser -LicenseName (Get-ExchangeServerAccessLicense |
    Where-Object {($_.UnitLabel -eq "CAL") -and ($_.LicenseName -like "*Standard*")}).licenseName -WarningAction silentlyContinue |
    Measure-Object | Select-Object Count).count

$cells=@("Exchange Standard CALs","Exchange Enterprise CALs")
$LicenceReport += Generate-HTMLTable "$l_licence_TableHeadline" $cells

if ( $mode -eq 0 ) { $cells=@("$StdCAL","$EntCAL") }
else  { $cells=@("$StdCAL / $OwnedStdCAL","$EntCAL / $OwnedEntCAL") }

$LicenceReport += New-HTMLTableLine $cells
$LicenceReport += End-HTMLTable

#Chart
$hashtable1 = @{}
$hashtable2 = @{}

$diff = $StdCAL - $OwnedStdCAL
if ($diff -lt 0)
  {
      $hashtable1["Standard-CAL"] = $StdCAL
      $hashtable2["Standard-CAL"] = 0
  }
  elseif($diff -eq $StdCAL) #no limit
  {
      $hashtable1["Standard-CAL"] = $StdCAL
      $hashtable2["Standard-CAL"] = 0
  }
  else
  {
      $hashtable1["Standard-CAL"] = $OwnedStdCAL
      $hashtable2["Standard-CAL"] = $diff
  }

$diff = $EntCAL - $OwnedEntCAL
if ($diff -lt 0)
  {
      $hashtable1["Enterprise-CAL"] = $EntCAL
      $hashtable2["Enterprise-CAL"] = 0
  }
  elseif($diff -eq $EntCAL) #no limit
  {
      $hashtable1["Enterprise-CAL"] = $EntCAL
      $hashtable2["Enterprise-CAL"] = 0
  }
  else
  {
      $hashtable1["Enterprise-CAL"] = $OwnedEntCAL
      $hashtable2["Enterprise-CAL"] = $diff
  }

new-StackedColumnchart "500" "400" "CAL Overview" $hashtable1 $hashtable2 "$tmpdir\licence.png" "Name" "Count" $mode
$LicenceReport += Include-HTMLInlinePictures "$tmpdir\licence.png"

$LicenceReport | set-content "$tmpdir\licencereport.html"
$LicenceReport | add-content "$tmpdir\report.html"

#Cleanup
Remove-Variable mode, hashtable1, hashtable2, diff, EntCAL, OwnedEntCAL, StdCAL, OwnedStdCAL
