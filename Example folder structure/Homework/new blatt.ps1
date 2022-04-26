# This file is just here to show the folder/file structure
# To be sure to get the most current version of this script,
# use the "new blatt.ps1" in the root directory of the 
# "new_blatt"-github  repository (https://github.com/r0uv3n/new_blatt)
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
