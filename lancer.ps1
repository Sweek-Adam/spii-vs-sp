# =====================================================================
# lancer.ps1 — Lance le script SPII vs SP
# ---------------------------------------------------------------------
# Double-clique ce fichier (ou clic droit -> "Exécuter avec PowerShell")
# pour générer le fichier Excel sans taper de commande.
#
# Prérequis : WinPython ajouté au PATH (voir ajouter-winpython-au-path.ps1),
# OU adapte la variable $python ci-dessous vers le python de WinPython.
# =====================================================================

# Se placer dans le dossier de CE script (donc à côté de spii_v2.py)
Set-Location -Path $PSScriptRoot

# --- Choix de l'interpréteur Python ---
# Par défaut : utilise le "python" du PATH (si tu l'as ajouté).
$python = "python"

# Si WinPython n'est PAS dans ton PATH, décommente la ligne suivante et
# adapte le chemin vers le python.exe de ton WinPython :
# $python = "C:\WPy64-31360\python-3.13.6.amd64\python.exe"

Write-Host ""
Write-Host "=== Lancement de SPII vs SP ===" -ForegroundColor Cyan
Write-Host ""

# Vérifier que Python est accessible
try {
    & $python --version | Out-Null
} catch {
    Write-Host "X Python introuvable." -ForegroundColor Red
    Write-Host "  Soit ajoute WinPython au PATH (ajouter-winpython-au-path.ps1),"
    Write-Host "  soit edite ce script et renseigne le chemin complet dans la"
    Write-Host "  variable \$python."
    Write-Host ""
    Read-Host "Appuie sur Entree pour fermer"
    exit 1
}

# Lancer le script principal
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
