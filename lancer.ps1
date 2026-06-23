# =====================================================================
# lancer.ps1 — Lance le script SPII vs SP
# ---------------------------------------------------------------------
# Double-clique ce fichier (ou clic droit -> "Exécuter avec PowerShell")
# pour générer le fichier Excel sans taper de commande.
#
# Le chemin de Python est lu depuis config.toml (champ python_exe sous
# la section [chemins]). Renseigne-le avant le premier lancement.
# =====================================================================

# Se placer dans le dossier de CE script (donc à côté de spii_v2.py et config.toml)
Set-Location -Path $PSScriptRoot

Write-Host ""
Write-Host "=== Lancement de SPII vs SP ===" -ForegroundColor Cyan
Write-Host ""

# --- 1. Vérifier la présence de config.toml ---
$configPath = Join-Path $PSScriptRoot "config.toml"
if (-not (Test-Path $configPath)) {
    Write-Host "X config.toml introuvable a cote de ce script." -ForegroundColor Red
    Write-Host "  Copie config.toml.exemple en config.toml et renseigne-le."
    Write-Host ""
    Read-Host "Appuie sur Entree pour fermer"
    exit 1
}

# --- 2. Lire le champ python_exe depuis config.toml ---
# (extraction ciblée d'une ligne 'python_exe = "..."' ou "'...'")
$python = $null
foreach ($ligne in Get-Content $configPath) {
    # Ignorer les commentaires
    $sansCommentaire = $ligne -replace '#.*$', ''
    if ($sansCommentaire -match '^\s*python_exe\s*=\s*["''](.+?)["'']\s*$') {
        $python = $matches[1].Trim()
        break
    }
}

# --- 3. Vérifier que le chemin est renseigné et valide ---
if ([string]::IsNullOrWhiteSpace($python)) {
    Write-Host "X Le champ 'python_exe' est vide dans config.toml." -ForegroundColor Red
    Write-Host "  Renseigne le chemin complet vers python.exe de WinPython, ex. :"
    Write-Host "  python_exe = 'C:\...\WinPython\WPy64-xxxxx\python\python.exe'"
    Write-Host ""
    Read-Host "Appuie sur Entree pour fermer"
    exit 1
}

if (-not (Test-Path $python)) {
    Write-Host "X Le chemin 'python_exe' de config.toml est introuvable :" -ForegroundColor Red
    Write-Host "  $python"
    Write-Host "  Verifie le chemin dans config.toml."
    Write-Host ""
    Read-Host "Appuie sur Entree pour fermer"
    exit 1
}

Write-Host "Python utilise : $python" -ForegroundColor Green
Write-Host ""

# --- 4. Lancer le script principal ---
& $python "spii_v2.py"
$code = $LASTEXITCODE

Write-Host ""
if ($code -eq 0) {
    Write-Host "Termine." -ForegroundColor Green
} else {
    Write-Host "Le script s'est termine avec une erreur (code $code)." -ForegroundColor Yellow
}

# Garder la fenêtre ouverte pour lire les messages
Write-Host ""
Read-Host "Appuie sur Entree pour fermer"
