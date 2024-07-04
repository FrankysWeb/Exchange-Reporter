$vmwarereport = Generate-ReportHeader "vmwarereport.png" "$l_vm_header"

$vcsettingshash = $inifile["VMWareReport"]
$vcsettings = convert-hashtoobject $vcsettingshash

$viserver = ($vcsettings| Where-Object {$_.Setting -eq "vCenterServer"}).Value
$viuser = ($vcsettings| Where-Object {$_.Setting -eq "vCenterUser"}).Value
$vipassword = ($vcsettings | Where-Object {$_.Setting -eq "vCenterPassword"}).Value

$vmlist = ($vcsettings | Where-Object {$_.Setting -match "VM"}).Value

$vipasspassword = $vipassword | ConvertTo-SecureString -AsPlainText -Force
$Creds= New-Object System.Management.Automation.PSCredential -ArgumentList $viuser, $vipasspassword

$loadpowercli = Get-Module –ListAvailable VM* | Import-Module | Out-null
if ($loadpowercli)
{
 $loadpowercli = Get-Module –ListAvailable VM* | Import-Module
}
else
{
 write-host "$l_vm_clierror" -foregroundcolor red
 exit 0
}
$PowerCLIConfig = Set-PowerCLIConfiguration -Scope Session -ProxyPolicy NoProxy -InvalidCertificateAction Ignore -ParticipateInCEIP $false -Confirm:$false  | Out-null 
$viconnect = connect-viserver $viserver -credential $creds -force -wa 0 -ea 0

if (!$viconnect)
{
 write-host "$l_vm_conerror" -foregroundcolor red
 exit 0
}

$cells=@("$l_vm_vm","$l_vm_host","$l_vm_vmhardware","$l_vm_ram","$l_vm_cpu","$l_vm_status","$l_vm_snapshotcount")
$vmwarereport += Generate-HTMLTable "$l_vm_header1" $cells

foreach ($server in $vmlist)
{
 $vm = get-vm $server
 
 $vmname = $vm.Name
 $vmhost = $vm.vmhost
 $vmhardware = $vm.hardwareversion
 $vmram = $vm.MemoryGB
 $vmcpu = $vm.NumCpu
 $vmstate = $vm.powerstate
 $snapcount = (Get-Snapshot $vm).count
 
 $cells=@("$vmname","$vmHost","$vmhardware","$vmram","$vmcpu","$vmstate","$snapcount")
 $vmwarereport += New-HTMLTableLine $cells
}
$vmwarereport += End-HTMLTable

$dsvmcount = 1
foreach ($server in $vmlist)
{
  $vm = get-vm $server
 
 $vmname = $vm.Name
 $vmhost = $vm.vmhost
 
 $cells=@("$l_vm_vm","$l_vm_host","$l_vm_datastore","$l_vm_datastorecapacity","$l_vm_datastorefree","$l_vm_datastoreformat")
 $vmwarereport += Generate-HTMLTable "Datatores $vmname Übersicht" $cells
 
 $vmdatastores = get-datastore -vm $server
 foreach ($vmdatastore in $vmdatastores)
 {
 $dsname = $vmdatastore.Name
 $dscapacity = $vmdatastore.CapacityGB
 $dscapacity  = [System.Math]::Round($dscapacity , 2)
 $dsfree = $vmdatastore.FreeSpaceGB
 $dsfree  = [System.Math]::Round($dsfree , 2)
 $dsformat = $vmdatastore.type
 
 $cells=@("$vmname","$vmHost","$dsname","$dscapacity","$dsfree","$dsformat")
 $vmwarereport += New-HTMLTableLine $cells
 
 $dsused = $dscapacity - $dsfree
 $filename = "vm" + "$dsvmcount" + "_" + "$dsname" + ".png"
 $chartdata = @{Frei=$dsfree; Belegt=$dsused}
 new-piechart "150" "150" "$dsname" $chartdata "$tmpdir\$filename"
 }
   $vmwarereport += End-HTMLTable
   $vmwarereport += Include-HTMLInlinePictures "$tmpdir\vm$dsvmcount*.png"
   $dsvmcount ++
}

$vmwarereport | set-content "$tmpdir\vmwarereport.html"
$vmwarereport | add-content "$tmpdir\report.html"