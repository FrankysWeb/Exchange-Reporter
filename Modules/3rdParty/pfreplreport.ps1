#
# Modul PFReplReport von Ralph Andreas Altermann & Florian Heller, Alfru GmbH - IT.Systems, Rellingen
# Version 0.9 Beta - 08.2015
#
# Für Exchange 2010 und älter
#
# Dieser Report zeigt alle Öffentlichen Ordner (IPM_SUBTREE und NON_IPM_SUBTREE) sowie dessen Replikationsstatus zwischen den Datenbanken an.
# Zusätzlich gibt es die Option alle Ergebnisse in eine CSV exportieren zu lassen.
#
# Wir verwenden Teile aus Franks pfreport.ps1, damit kann der Gesamt-Bericht sehr groß bzw. lang werden, sodass das Modul "pfreport.ps1"  ggf.
# bei Bedraf temporär auskommentiert werden sollte.
# 
# 
# Dieser Report gliedert sich in vier Abschnitte: 1. Public Folders (IPM_SUBTREE) nach Größe
#                                                 2. Public Folders (IPM_SUBTREE) nach Namen
#                                                 3. Public Folders (NON_IPM_SUBTREE) Replikationsstatus
#                                                 4. Public Folders (IPM_SUBTREE) Replikationsstatus
#
# Es gibt drei Ausgabemöglichkeiten:  (1) nur HTML, (2) HTML & CSV oder (3) nur CSV - (0) überspringt dieses Script komplett
#
# Geplant ist die Anpassung der Spalten im HTML-Output, ein Items (Anzahl) Vergleich und die Unterdrückung der Exchange Powershell Warnings
# wenn es nicht als Task läuft.
#
# Für Fehlerhinweise und Anregungen sind wir dankbar und Ihr könnt Euch an Frank oder uns unter support@alfru.net wenden.
#------------------------------------------------------------------------------------------------------------------------------------------------------

$PFRepOut = 0     

$Date = get-Date -format yyyymmdd
$Clock = Get-Date -format HHmmss
$LoPath= $InstallPath+'\Export'$ErrorActionPreference = 'SilentlyContinue'

if($PFRepOut -notmatch "0")
{

$pfreplreport = Generate-ReportHeader "pfreplreport.png" "Alle Öffentliche und System Ordner mit deren Replikationsstatus"


if ($emsversion -match "2010")
	{
		$cells=@("Datenbank","Server","Größe","Letzte Vollsicherung")
		$pfreplreport += Generate-HTMLTable "Öffentliche Ordner Datenbanken" $cells
	
		$pfdbs = get-PublicFolderDatabase -Status
		foreach ($pfdb in $pfdbs)
			{
				$pfdbname = $pfdb.name
				$pfdbserver = $pfdb.server
				$pfdbsize = $pfdb.databasesize
				$pflastbackup = $pfdb.LastFullBackup
				if ($pflastbackup)
					{
						$pflastbackup = get-date $pflastbackup -UFormat "%d.%m.%Y %R"
					}
				else
					{
						$pflastbackup = "Nie"
					}
					
				$cells=@("$pfdbname","$pfdbserver","$pfdbsize","$pflastbackup")
				$pfreplreport += New-HTMLTableLine $cells
			}
		$pfreplreport += End-HTMLTable


#------Beginn HTML Export-----------------------------------------------------

    if ($PfRepOut -lt "3") 
       #--------Export HTML nach Größe ohne System Ordner---------------------
       {
		
		$cells=@("Name","Größe","Anzahl Elemente")
		$pfreplreport += Generate-HTMLTable "Öffentliche Ordner - sortiert nach Größe" $cells
		
		$pfs = Get-PublicFolder -Identity "\" -Recurse -Resultsize unlimited | Get-PublicFolderStatistics | sort Totalitemsize -Descending
		foreach ($pf in $pfs)
			{
				$pfname = $pf.AdminDisplayName
				$pfsize = $pf.TotalItemSize
				$pfitemcount = $pf.ItemCount
				
				$cells=@("$pfname","$pfsize","$pfitemcount")
				$pfreplreport += New-HTMLTableLine $cells
								
			}
		$pfreplreport += End-HTMLTable

        #--------Export HTML nach Name ohne System Ordner---------------------

        $cells=@("Name","Größe","Anzahl Elemente")
		$pfreplreport += Generate-HTMLTable "Öffentliche Ordner - sortiert nach Namen" $cells
		
		$pfs = Get-PublicFolder -Identity "\" -Recurse -Resultsize unlimited | Get-PublicFolderStatistics | sort Name
		foreach ($pf in $pfs)
			{
				$pfname = $pf.AdminDisplayName
				$pfsize = $pf.TotalItemSize
				$pfitemcount = $pf.ItemCount
				
				$cells=@("$pfname","$pfsize","$pfitemcount")
				$pfreplreport += New-HTMLTableLine $cells
								
			}
		$pfreplreport += End-HTMLTable

        #--------Export HTML System Ordner -----------------------------------
        
        $cells=@("Name","Größe","Anzahl Elemente")
		$pfreplreport += Generate-HTMLTable "System Ordner - Inhalte" $cells
		
		$pfs = Get-PublicFolder -Identity "\NON_IPM_SUBTREE" -Recurse -Resultsize unlimited | Get-PublicFolderStatistics | sort Name 
		
            foreach ($pf in $pfs)
        	{
				$pfname = $pf.AdminDisplayName
				$pfsize = $pf.TotalItemSize
				$pfitemcount = $pf.ItemCount
				
				$cells=@("$pfname","$pfsize","$pfitemcount")
				$pfreplreport += New-HTMLTableLine $cells
			}
        $pfreplreport += End-HTMLTable

        #-----------Export HTML Replikationstatus: Öffentliche Ordner --------

        $cells=@("Name","Datenbanken")
		$pfreplreport += Generate-HTMLTable "Replikationsstatus: Öffentliche Ordner" $cells
        
        $pffs = Get-PublicFolder -Identity "\" -Recurse -Resultsize unlimited 
    
            foreach ($pff in $pffs)
            {
                $pfname = ($pff).name 
                $pfrepl = ($pff).replicas -split(",") -Join(" | ")

                $cells=@("$pfname","$pfrepl")
				$pfreplreport += New-HTMLTableLine $cells
								
	        }
		$pfreplreport += End-HTMLTable

        #------------Export HTML Replikationstatus: System Ordner ------------

        $cells=@("Name","Datenbanken")
		$pfreplreport += Generate-HTMLTable "Replikationsstatus: System Ordner" $cells
        
        $pffs = Get-PublicFolder -Identity "\NON_IPM_SUBTREE" -Recurse -Resultsize unlimited 
    
            foreach ($pff in $pffs)
            {
                $pfname = ($pff).name 
                $pfrepl = ($pff).replicas -split(",") -Join(" | ")

                $cells=@("$pfname","$pfrepl")
				$pfreplreport += New-HTMLTableLine $cells
								
	        }
		$pfreplreport += End-HTMLTable
       }

#------Ende HTML Export-------------------------------------------------------
#------Begin CSV Export-------------------------------------------------------

       if ($PfRepOut -gt "1")
       #-------Export CSV Öffentlich Ordner(Size) ----------------------------
       {

            New-Item -Path $LoPath -ItemType directory -ErrorAction SilentlyContinue

	        $Cells=@("Pfad zur Datei - Öffentliche Ordner (Size)")
            $Outputcsv='PubFolderSize'+'_'+$Date+'_'+$Clock
            $pfreplreport += Generate-HTMLTable "Export-Datei" $cells
    	
              Get-PublicFolder -Identity "\" -recurse -Resultsize unlimited | Get-PublicFolderStatistics | sort TotalItemSize -Descending | select-Object Name,TotalItemSize,ItemCount | Export-CSV $LoPath\$Outputcsv.csv -Delimiter ";"  -NoTypeInformation
       	    
            $Exportpath = $LoPath+'\'+$Outputcsv+'.csv'
            
            $cells=@("$Exportpath")
        	$pfreplreport += New-HTMLTableLine $cells
	        $pfreplreport += End-HTMLTable

            #-------Export CSV System Ordner (Size) ---------------------------

	        $Cells=@("Pfad zu Datei - System Ordner (Size)")
            $Outputcsv='SysFolderSize'+'_'+$Date+'_'+$Clock
            $pfreplreport += Generate-HTMLTable "Export-Datei" $cells
    	
              Get-PublicFolder -Identity "\NON_IPM_SUBTREE" -recurse -Resultsize unlimited | Get-PublicFolderStatistics | sort TotalItemSize -Descending | select-Object Name,TotalItemSize,ItemCount | Export-CSV $LoPath\$Outputcsv.csv -Delimiter ";"  -NoTypeInformation
       	    
            $Exportpath = $LoPath+'\'+$Outputcsv+'.csv'

            $cells=@("$Exportpath")
        	$pfreplreport += New-HTMLTableLine $cells
	        $pfreplreport += End-HTMLTable

            #-------Export CSV System Ordner (Replikation)---------------------

            $PFsarray =@()            $PFrepl = @()            $PFDBName =@()            $PFDBMax = (Get-PublicFolderDatabase).Count            $PFDBName = (Get-PublicFolderDatabase).AdminDisplayname 

	        $Cells=@("Pfad zu Datei - System Ordner (Replikation)")
            $Outputcsv='SysFolderRepl'+'_'+$Date+'_'+$Clock
            $pfreplreport += Generate-HTMLTable "Export-Datei" $cells

            $Pffs = Get-PublicFolder -Identity "\NON_IPM_SUBTREE" -Recurse -Resultsize unlimited 

            foreach ($pff in $pffs)
            {
                $pfname = ($pff).name 
                $pfpath = ($pff).parentpath
                $pfrepl = ($pff).replicas -Split(",")
  
                $PFObject = new-object PSObject
  
                $PFObject | add-member -membertype NoteProperty -name "Ordner" -Value $pfname 
                $PFObject | add-member -membertype NoteProperty -name "Pfad"  -Value $pfpath
                
                for($PFCol=0;$PFCol -lt $PFDBMax; $PFCol++)
                {
                    if($pfrepl.Contains($PFDBName[$PFCol]))
                    {
                        $PFObject | add-member -membertype NoteProperty -name "DB[$PFCol]" -Value $PFDBName[$PFCol] # Folder ist in DB vorhanden
                    }
                    else
                    {
                        $PFObject | add-member -membertype NoteProperty -name "DB[$PFCol]" -Value "" # Folder existiert nicht in dieser DB
                    }
                }                
                $PFsarray += $PFObject
            }  

            $PFsarray | Export-Csv $LoPath\$Outputcsv.csv -Delimiter ";"  -NoTypeInformation 

            
            $Exportpath = $LoPath+'\'+$Outputcsv+'.csv'

            $cells=@("$Exportpath")
        	$pfreplreport += New-HTMLTableLine $cells
	        $pfreplreport += End-HTMLTable

            #-------Export CSV Öffentliche Ordner (Replikation) ---------------------

            $PFsarray =@()            $PFrepl = @()            $PFDBName =@()            $PFDBMax = (Get-PublicFolderDatabase).Count            $PFDBName = (Get-PublicFolderDatabase).AdminDisplayname 

            $Cells=@("Pfad zu Datei - Öffentliche Ordner (Replikation)")
            $Outputcsv='PubFolderRepl'+'_'+$Date+'_'+$Clock
            $pfreplreport += Generate-HTMLTable "Export-Datei" $cells

            $Pffs = Get-PublicFolder -Identity "\" -Recurse -Resultsize unlimited 

            foreach ($pff in $pffs)
            {
                $pfname = ($pff).name 
                $pfpath = ($pff).parentpath
                $pfrepl = ($pff).replicas -Split(",")
  
                $PFObject = new-object PSObject
  
                $PFObject | add-member -membertype NoteProperty -name "Ordner" -Value $pfname 
                $PFObject | add-member -membertype NoteProperty -name "Pfad"  -Value $pfpath
                
                for($PFCol=0;$PFCol -lt $PFDBMax; $PFCol++)
                {
                    if($pfrepl.Contains($PFDBName[$PFCol]))
                    {
                        $PFObject | add-member -membertype NoteProperty -name "DB[$PFCol]" -Value $PFDBName[$PFCol] # Folder ist in DB vorhanden
                    }
                    else
                    {
                        $PFObject | add-member -membertype NoteProperty -name "DB[$PFCol]" -Value "" # Folder existiert nicht in dieser DB
                    }
                }                
                $PFsarray += $PFObject
            }  

            $PFsarray | Export-Csv $LoPath\$Outputcsv.csv -Delimiter ";"  -NoTypeInformation

            $Exportpath = $LoPath+'\'+$Outputcsv+'.csv'

            $cells=@("$Exportpath")
        	$pfreplreport += New-HTMLTableLine $cells
	        $pfreplreport += End-HTMLTable

         } 
#-----------Ende CSV Export-----------------------------------------------------------
       }

$pfreplreport += End-HTMLTable
$pfreplreport | set-content "$tmpdir\pfreplreport.html"
$pfreplreport | add-content "$tmpdir\report.html"
}