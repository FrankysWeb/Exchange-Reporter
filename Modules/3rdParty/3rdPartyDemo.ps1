#Modul überschrift erzeugen:
# Die PNG Datei (Logo vor Überschrift) sollte eine größe von 31x31px nicht überschreiten, die PNG-Datei muss wie das Modul benannt werden
# Modulname: 3rdPartydemo.ps1 (\Modules\3rdParty\)
# PNG Datei Name: 3rdPartyDemo.png (\Images\3rdParty\)
# Farbcode für Hintergrund der PNG Datei: Rot=0, Grün=114, Blau=198

$3rdPartyDemo = Generate-ReportHeader "3rdPartyDemo.png" "Mein eigenes Modul"

#Neue Tabelle erzeugen:

$cells=@("Spalte 1","Spalte 2","Spalte 3","usw","usw")
$3rdPartyDemo += Generate-HTMLTable "Überschrift der Tabelle" $cells

#Neue Zeile mit Daten der Tabelle hinzufügen:

$cells=@("Zeile1 Spalte 1","Zeile 1 Spalte 2"," Zeile 1 Spalte 3","usw","usw")
$3rdPartyDemo += New-HTMLTableLine $cells

#Noch eine Zeile hinzufügen

$cells=@("Zeile2 Spalte 1","Zeile 2 Spalte 2"," Zeile 2 Spalte 3","usw","usw")
$3rdPartyDemo += New-HTMLTableLine $cells

#Tabelle abschließen

$3rdPartyDemo += End-HTMLTable

#Neues Balkendiagram erzeugen

$values = @{"Balken1"=10;"Balken2"=20;"Balken3"=30}
new-cylinderchart 500 400 Balken Name "Anzahl" $outlookcltvalues "$tmpdir\diagramname.png"

#Neues Kuchendiagram erzeugen

$values = @{Frei=50; Belegt=50}
new-piechart "150" "150" "Name" $values "$tmpdir\kuchendiagramname.png"

#Grafik in Report einfügen

$3rdPartyDemo += Include-HTMLInlinePictures "$tmpdir\diagramname.png"
$3rdPartyDemo += Include-HTMLInlinePictures "$tmpdir\kuchendiagramname.png"

#Report erzeugen und Modul beenden
# 3rdPartyDemo.html für Fehleranalyse
# Report.html ist der komplette Report der verschickt wird

$3rdPartyDemo | set-content "$tmpdir\3rdPartyDemo.html"
$3rdPartyDemo | add-content "$tmpdir\report.html"