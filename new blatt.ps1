Set-Location $PSScriptRoot
$Language = "EN"

if ($Language -eq "DE"){
  $Blatttext = "Blatt"
  $Vorlagentext = "Vorlage"
}
elseif ($Language -eq "EN") {
  $Blatttext = "Sheet"
  $Vorlagentext = "Template"
} 

$Blätter = Get-ChildItem -Directory ("{0} [0-9][0-9]" -f $Blatttext)


$Vorlage = Get-ChildItem -Directory ("*{0}*"  -f $Vorlagentext)


foreach ($Blatt in $Blätter) { 
  $Blatt.Name -match "(\d+)">$null
  Add-Member -Force -InputObject $Blatt -NotePropertyName Nummer -NotePropertyValue ([int]($matches[0])) 
}

$Blattnummer = ($Blätter.Nummer | Measure-Object -Maximum).Maximum + 1

$Blattname = "{0} {1:00}" -f $Blatttext,$Blattnummer

New-Item -Name $Blattname -ItemType Directory

Copy-Item -Path $Vorlage\* -Destination $Blattname -Recurse

$Dateien = Get-ChildItem $Blattname

foreach ($Datei in $Dateien) {
  (Get-Content $Datei) -replace $Vorlagentext, $Blattnummer | Set-Content $Datei
  Rename-Item $Datei -NewName ($Datei.Name -replace $Vorlagentext, ("{0:00}" -f $Blattnummer))
}
