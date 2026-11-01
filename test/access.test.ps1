# Test de outils/exporter-access.ps1 avec un faux Access.Application (pwsh, sans Access).
$ErrorActionPreference = 'Stop'
$ici = Split-Path -Parent $MyInvocation.MyCommand.Path
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("acc-" + [guid]::NewGuid().ToString('N'))
$src = Join-Path $tmp 'Qualite'
New-Item -ItemType Directory -Force -Path (Join-Path $src 'Sous') | Out-Null
[System.IO.File]::WriteAllBytes((Join-Path $src 'Base NC.accdb'), [byte[]](1, 2, 3))
[System.IO.File]::WriteAllBytes((Join-Path $src 'Sous/ancienne.mdb'), [byte[]](4, 5, 6))
[System.IO.File]::WriteAllBytes((Join-Path $src 'notes.txt'), [byte[]](7))
$avant = (Get-FileHash (Join-Path $src 'Base NC.accdb')).Hash

function Module($nom, $type, $code) {
  $cm = [pscustomobject]@{ Code = $code }
  $cm | Add-Member -MemberType ScriptProperty -Name CountOfLines -Value { if ($this.Code) { ($this.Code -split "`r`n").Count } else { 0 } }
  $cm | Add-Member -MemberType ScriptMethod -Name Lines -Value { param($a, $b) ($this.Code -split "`r`n")[($a - 1)..($a + $b - 2)] -join "`r`n" }
  [pscustomobject]@{ Name = $nom; Type = $type; CodeModule = $cm }
}
$ouvertures = New-Object System.Collections.ArrayList
$faux = [pscustomobject]@{ AutomationSecurity = 1; VBE = $null; Ouvertes = $ouvertures }
$faux | Add-Member -MemberType ScriptMethod -Name OpenCurrentDatabase -Value {
  param($chemin, $excl)
  if ($this.AutomationSecurity -ne 3) { throw 'macros non désactivées avant ouverture' }
  if (-not (Test-Path $chemin)) { throw 'copie absente' }
  [void]$this.Ouvertes.Add($chemin)
  $mods = @((Module 'modExport' 1 "Option Explicit`r`nPrivate Declare Function GetTickCount Lib `"kernel32`" () As Long"), (Module 'Form_Saisie' 100 "Private Sub Form_Load()`r`n  Set o = CreateObject(`"Outlook.Application`")`r`nEnd Sub"), (Module 'Vide' 1 ''))
  $this.VBE = [pscustomobject]@{ ActiveVBProject = [pscustomobject]@{ VBComponents = $mods } }
}
$faux | Add-Member -MemberType ScriptMethod -Name CloseCurrentDatabase -Value { $this.VBE = $null }

$sortie = Join-Path $tmp 'export'
& (Join-Path $ici '../outils/exporter-access.ps1') -Dossier $src -Sortie $sortie -Application $faux | Out-Null

$attendus = @('Base NC.accdb/modExport.bas', 'Base NC.accdb/Form_Saisie.cls', (Join-Path 'Sous_ancienne.mdb' 'modExport.bas'))
foreach ($a in $attendus) { if (-not (Test-Path (Join-Path $sortie $a))) { throw "manquant : $a" } }
if (Test-Path (Join-Path $sortie 'Base NC.accdb/Vide.bas')) { throw 'module vide exporté' }
$bas = Get-Content -Raw (Join-Path $sortie 'Base NC.accdb/modExport.bas')
if ($bas -notmatch 'Attribute VB_Name = "modExport"' -or $bas -notmatch 'Declare Function GetTickCount') { throw 'contenu exporté incorrect' }
if ($faux.Ouvertes.Count -ne 2) { throw "2 ouvertures attendues, $($faux.Ouvertes.Count)" }
foreach ($o in $faux.Ouvertes) { if ($o -like "$src*") { throw "l'original a été ouvert : $o" }; if (Test-Path $o) { throw "copie temporaire non supprimée : $o" } }
if ((Get-FileHash (Join-Path $src 'Base NC.accdb')).Hash -ne $avant) { throw 'original modifié' }
if (-not (Test-Path (Join-Path $sortie '_bilan-export.txt'))) { throw 'bilan absent' }
Write-Output "SORTIE=$sortie"
Write-Output 'OK'
