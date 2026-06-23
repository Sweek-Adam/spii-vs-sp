# =====================================================================
# ajouter-winpython-au-path.ps1
# ---------------------------------------------------------------------
# Ajoute WinPython au PATH de TON compte utilisateur (pas besoin d'admin).
# À lancer UNE SEULE FOIS, après avoir décompressé WinPython.
#
# Après ça, la commande `python` fonctionnera dans n'importe quel terminal,
# sans passer par le "WinPython Command Prompt".
#
# UTILISATION :
#   1. Place ce fichier où tu veux.
#   2. Clic droit dessus -> "Exécuter avec PowerShell"
#      (ou : ouvre PowerShell et lance  .\ajouter-winpython-au-path.ps1)
#   3. Indique le dossier de ton WinPython quand c'est demandé.
#   4. FERME puis ROUVRE tes terminaux pour que le PATH soit pris en compte.
# =====================================================================

Write-Host ""
Write-Host "=== Ajout de WinPython au PATH utilisateur ===" -ForegroundColor Cyan
Write-Host ""

# 1. Demander le dossier WinPython
$wpRoot = Read-Host "Colle le chemin du dossier WinPython (ex: C:\WPy64-31360)"
$wpRoot = $wpRoot.Trim('"').Trim()

if (-not (Test-Path $wpRoot)) {
    Write-Host "X Dossier introuvable : $wpRoot" -ForegroundColor Red
    Write-Host "  Verifie le chemin et relance le script."
    exit 1
}

# 2. Trouver automatiquement le sous-dossier Python et le dossier Scripts.
#    Selon la version de WinPython, le dossier Python s'appelle :
#      - "python"                (versions récentes, ~2024+)
#      - "python-3.x.x.amd64"    (versions plus anciennes)
#    Et le dossier des scripts (pip) est soit "scripts" à la racine,
#    soit "Scripts" dans le dossier python.
$pythonDir = $null

# a) Version récente : dossier nommé exactement "python"
$candidat = Join-Path $wpRoot "python"
if (Test-Path (Join-Path $candidat "python.exe")) {
    $pythonDir = Get-Item $candidat
}

# b) Sinon, ancienne convention : "python-*"
if (-not $pythonDir) {
    $pythonDir = Get-ChildItem -Path $wpRoot -Directory |
        Where-Object { $_.Name -like "python-*" } |
        Select-Object -First 1
}

if (-not $pythonDir) {
    Write-Host "X Aucun dossier Python trouve dans $wpRoot" -ForegroundColor Red
    Write-Host "  (ni 'python', ni 'python-*'). Es-tu sur du dossier racine WinPython ?"
    exit 1
}

$pythonPath = $pythonDir.FullName

# Dossier Scripts : essayer "scripts" à la racine WinPython, puis "Scripts"
# dans le dossier python.
$scriptsPath = $null
foreach ($cand in @((Join-Path $wpRoot "scripts"), (Join-Path $pythonPath "Scripts"))) {
    if (Test-Path $cand) { $scriptsPath = $cand; break }
}

Write-Host "Python detecte : $pythonPath" -ForegroundColor Green
if ($scriptsPath) {
    Write-Host "Scripts (pip)  : $scriptsPath" -ForegroundColor Green
}

# 3. Recuperer le PATH utilisateur actuel
$currentPath = [Environment]::GetEnvironmentVariable("Path", "User")
if ($null -eq $currentPath) { $currentPath = "" }

# 4. Construire la liste des chemins a ajouter (en evitant les doublons)
$cheminsCibles = @($pythonPath)
if ($scriptsPath) { $cheminsCibles += $scriptsPath }

$aAjouter = @()
foreach ($p in $cheminsCibles) {
    if ((Test-Path $p) -and ($currentPath -notlike "*$p*")) {
        $aAjouter += $p
    }
}

if ($aAjouter.Count -eq 0) {
    Write-Host ""
    Write-Host "i WinPython est deja dans ton PATH. Rien a faire." -ForegroundColor Yellow
    exit 0
}

# 5. Ajouter au PATH utilisateur (sans toucher au PATH systeme = pas d'admin)
$nouveauPath = $currentPath.TrimEnd(';')
foreach ($p in $aAjouter) {
    $nouveauPath = "$nouveauPath;$p"
}
$nouveauPath = $nouveauPath.TrimStart(';')

[Environment]::SetEnvironmentVariable("Path", $nouveauPath, "User")

Write-Host ""
Write-Host "OK ! Ajoute au PATH utilisateur :" -ForegroundColor Green
foreach ($p in $aAjouter) { Write-Host "   + $p" }
Write-Host ""
Write-Host "IMPORTANT : ferme puis rouvre tes terminaux (PowerShell, cmd)" -ForegroundColor Yellow
Write-Host "pour que le changement prenne effet."
Write-Host ""
Write-Host "Ensuite, teste avec :  python --version"
Write-Host ""
