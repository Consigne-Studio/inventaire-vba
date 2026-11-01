<#
.SYNOPSIS
  Exporte le code VBA des bases Access (.accdb, .mdb) d'un dossier, pour l'inventaire des macros.

.DESCRIPTION
  Access est le seul à pouvoir lire le code de ses bases. Ce script, lancé sur un poste où Access est installé :
    1. copie chaque base dans un dossier temporaire (l'original n'est jamais ouvert ni modifié) ;
    2. ouvre la copie avec les macros désactivées (AutomationSecurity = 3 : rien ne s'exécute à l'ouverture) ;
    3. écrit chaque module en .bas, chaque module de classe, de formulaire ou d'état en .cls ;
    4. supprime la copie.
  Lancez ensuite inventaire-vba sur le dossier de sortie.

.EXAMPLE
  powershell -ExecutionPolicy Bypass -File exporter-access.ps1 -Dossier "\\serveur\Qualite" -Sortie "C:\Inventaire\access"
#>
param(
  [Parameter(Mandatory = $true)][string]$Dossier,
  [string]$Sortie = (Join-Path (Get-Location) 'export-access'),
  [object]$Application = $null  # pour les tests : objet qui imite Access.Application
)

$ErrorActionPreference = 'Stop'
$extensions = @('*.accdb', '*.mdb')
$types = @{ 1 = 'bas'; 2 = 'cls'; 3 = 'cls'; 100 = 'cls' }  # module, classe, formulaire MSForms, document (formulaire / état Access)

function Nom-Sur([string]$nom) { return ($nom -replace '[\\/:*?"<>|]', '_') }

function Exporter-Base([object]$app, [string]$base, [string]$cible) {
  $tmp = Join-Path ([System.IO.Path]::GetTempPath()) ("inventaire-vba-" + [guid]::NewGuid().ToString('N') + [System.IO.Path]::GetExtension($base))
  Copy-Item -LiteralPath $base -Destination $tmp
  $n = 0
  try {
    $app.AutomationSecurity = 3
    $app.OpenCurrentDatabase($tmp, $false)
    try {
      $projet = $app.VBE.ActiveVBProject
      if ($null -eq $projet) { return 0 }
      New-Item -ItemType Directory -Force -Path $cible | Out-Null
      foreach ($c in $projet.VBComponents) {
        $lignes = $c.CodeModule.CountOfLines
        if ($lignes -lt 1) { continue }
        $ext = $types[[int]$c.Type]
        if (-not $ext) { $ext = 'cls' }
        $code = "Attribute VB_Name = `"$($c.Name)`"`r`n" + $c.CodeModule.Lines(1, $lignes)
        [System.IO.File]::WriteAllText((Join-Path $cible ((Nom-Sur $c.Name) + '.' + $ext)), $code, [System.Text.Encoding]::UTF8)
        $n++
      }
    } finally { $app.CloseCurrentDatabase() }
  } finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }
  return $n
}

$fichiers = Get-ChildItem -LiteralPath $Dossier -Recurse -File -Include $extensions -ErrorAction SilentlyContinue
if (-not $fichiers) { Write-Host "Aucune base .accdb ou .mdb dans $Dossier"; exit 0 }

$creee = $false
if ($null -eq $Application) {
  try { $Application = New-Object -ComObject Access.Application; $creee = $true }
  catch { Write-Error "Microsoft Access est introuvable sur ce poste : lancez ce script sur un poste où Access est installé."; exit 1 }
}

$racine = (Resolve-Path -LiteralPath $Dossier).Path.TrimEnd('\', '/')
$bilan = @()
try {
  foreach ($f in $fichiers) {
    $relatif = $f.FullName.Substring($racine.Length).TrimStart('\', '/')
    $cible = Join-Path $Sortie (Nom-Sur $relatif)
    try {
      $n = Exporter-Base $Application $f.FullName $cible
      $bilan += "$n module(s)`t$relatif"
      Write-Host ("{0,4} module(s)  {1}" -f $n, $relatif)
    } catch {
      $bilan += "ERREUR`t$relatif`t$($_.Exception.Message)"
      Write-Warning "$relatif : $($_.Exception.Message) (base protégée par mot de passe, .accde compilée ou déjà ouverte en exclusif)"
    }
  }
} finally {
  if ($creee) { $Application.Quit(); [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($Application) }
}
New-Item -ItemType Directory -Force -Path $Sortie | Out-Null
[System.IO.File]::WriteAllLines((Join-Path $Sortie '_bilan-export.txt'), [string[]]$bilan, [System.Text.Encoding]::UTF8)
Write-Host "`nCode exporté dans $Sortie. Étape suivante : inventaire-vba `"$Sortie`""
