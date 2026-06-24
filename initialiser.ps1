# =====================================================================
# initialiser.ps1 - Assistant d'installation du projet SPII vs SP
# ---------------------------------------------------------------------
# Ce script prepare automatiquement ce qui peut l'etre :
#   1. Verifie / telecharge WinPython (Python portable, sans admin)
#   2. Detecte le python.exe et le renseigne dans config.toml
#   3. Installe les dependances Python (pandas, openpyxl, requests, truststore)
#   4. Cree config.toml et secrets.toml a partir des modeles .exemple
#   5. Autorise l'execution des scripts PowerShell (ExecutionPolicy)
#   6. Affiche un recapitulatif et les etapes manuelles restantes
#
# UTILISATION : clic droit -> "Executer avec PowerShell"
#               (ou .\initialiser.ps1 dans un terminal)
# =====================================================================

# Toujours travailler dans le dossier de ce script
Set-Location -Path $PSScriptRoot

# Version de WinPython a telecharger si absent (stable, avec wheels dispo)
$WP_URL  = "https://github.com/winpython/winpython/releases/download/17.4.20260511final/WinPython64-3.14.5.0dot.exe"
$WP_EXE  = "WinPython64-3.14.5.0dot.exe"
$WP_DIR  = Join-Path $PSScriptRoot "WinPython"

# Petites fonctions d'affichage
function Titre($t)   { Write-Host ""; Write-Host "=== $t ===" -ForegroundColor Cyan }
function OK($t)      { Write-Host "  [OK] $t" -ForegroundColor Green }
function Info($t)    { Write-Host "  [i]  $t" -ForegroundColor Gray }
function Warn($t)    { Write-Host "  [!]  $t" -ForegroundColor Yellow }
function Erreur($t)  { Write-Host "  [X]  $t" -ForegroundColor Red }

# --- Fonctions du questionnaire interactif ---

function Demander($libelle, $defaut = "") {
    # Pose une question ; si l'utilisateur laisse vide, renvoie le defaut.
    if ($defaut) {
        $rep = Read-Host "  $libelle [$defaut]"
        if ([string]::IsNullOrWhiteSpace($rep)) { return $defaut }
        return $rep.Trim()
    }
    $rep = Read-Host "  $libelle"
    return $rep.Trim()
}

function Echapper-Toml($v) {
    # En TOML guillemets simples, l'apostrophe ne peut pas etre echappee :
    # on la remplace par un accent (rare dans les chemins/emails, mais prudent).
    return ($v -replace "'", [char]0x2019)
}

function Ecrire-SansBOM($chemin, $contenu) {
    # Ecrit un fichier en UTF-8 SANS BOM. Le BOM (3 octets invisibles ajoutes
    # par Set-Content -Encoding UTF8 sous Windows PowerShell) fait echouer la
    # lecture TOML cote Python ("Invalid statement" en ligne 1).
    $utf8SansBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($chemin, $contenu, $utf8SansBom)
}

function Lire-Exemple($cle, $defaut = "") {
    # Lit une valeur (cle = "...") depuis config.toml.exemple s'il existe.
    # Sert a reprendre les champs fixes (url, sp_field, pi_field) sans les
    # redemander a l'utilisateur. Renvoie $defaut si introuvable.
    $ex = Join-Path $PSScriptRoot "config.toml.exemple"
    if (Test-Path $ex) {
        foreach ($ligne in Get-Content $ex) {
            $sans = $ligne -replace '#.*$', ''
            if ($sans -match ("^\s*" + [regex]::Escape($cle) + "\s*=\s*[""'](.+?)[""']\s*$")) {
                return $matches[1].Trim()
            }
        }
    }
    return $defaut
}

function Construire-Config($jira, $chemins, $ressources) {
    # Reconstruit un config.toml complet a partir des reponses.
    $sb = [System.Text.StringBuilder]::new()
    [void]$sb.AppendLine("# Configuration SPII vs SP - genere par initialiser.ps1")
    [void]$sb.AppendLine("# Chemins Windows en guillemets SIMPLES ' ' (antislash litteral).")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("[jira]")
    [void]$sb.AppendLine("email    = '$(Echapper-Toml $jira.email)'")
    [void]$sb.AppendLine("url      = '$(Echapper-Toml $jira.url)'")
    [void]$sb.AppendLine("sp_field = '$(Echapper-Toml $jira.sp_field)'")
    [void]$sb.AppendLine("pi_field = '$(Echapper-Toml $jira.pi_field)'")
    [void]$sb.AppendLine("projet   = '$(Echapper-Toml $jira.projet)'")
    [void]$sb.AppendLine("prefixe_feature = '$(Echapper-Toml $jira.prefixe_feature)'")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("[chemins]")
    [void]$sb.AppendLine("csv            = '$(Echapper-Toml $chemins.csv)'")
    [void]$sb.AppendLine("dossier_sortie = '$(Echapper-Toml $chemins.dossier_sortie)'")
    [void]$sb.AppendLine("python_exe     = '$(Echapper-Toml $chemins.python_exe)'")
    [void]$sb.AppendLine("")
    [void]$sb.AppendLine("# Equipe : Nom complet = Role (PO, SM, BA, DEV, QA)")
    [void]$sb.AppendLine("# Le nom doit correspondre EXACTEMENT a la colonne Ressource du CSV.")
    [void]$sb.AppendLine("[ressources]")
    foreach ($p in $ressources) {
        $nom = $p.nom -replace '"', "'"
        [void]$sb.AppendLine(('"{0}" = "{1}"' -f $nom, $p.role))
    }
    return $sb.ToString()
}

function Questionnaire-Config {
    Titre "Questionnaire : config.toml"
    Info "Laisse vide pour garder la valeur entre [crochets] quand proposee."
    Write-Host ""

    Write-Host "  -- Jira --" -ForegroundColor White
    # url, sp_field et pi_field sont repris du fichier config.toml.exemple
    # (valeurs fixes pour l'instance) : on ne les redemande pas.
    $jira = @{
        email    = Demander "Email Jira (ex. nom.prenom@imsa.msa.fr)"
        url      = Lire-Exemple "url"      "https://imsa.atlassian.net"
        sp_field = Lire-Exemple "sp_field" "customfield_10024"
        pi_field = Lire-Exemple "pi_field" "customfield_11400"
        projet   = Demander "Cle projet (ex. LIEVRE)" "LIEVRE"
        prefixe_feature = Lire-Exemple "prefixe_feature" "TCRE"
    }
    Info ("url / sp_field / pi_field / prefixe repris de l'exemple : {0}, {1}, {2}, {3}" -f $jira.url, $jira.sp_field, $jira.pi_field, $jira.prefixe_feature)

    Write-Host ""
    Write-Host "  -- Chemins --" -ForegroundColor White
    Info "Astuce : Maj + clic droit sur un fichier -> 'Copier en tant que chemin'."
    $chemins = @{
        csv            = (Demander "Chemin du CSV a lire - ex. C:\data\y25VUE9.csv") -replace '"', ''
        dossier_sortie = (Demander "Dossier de sortie - emplacement de la génération") -replace '"', ''
        python_exe     = ""
    }

    Write-Host ""
    Write-Host "  -- Equipe (ressources) --" -ForegroundColor White
    Info "Pour chaque membre : tape d'abord le NOM Prenom, puis Entree ;"
    Info "le script demande ensuite son role."
    Info "Le nom doit correspondre EXACTEMENT a la colonne Ressource du CSV."
    Info "Roles attendus : PO, SM, BA, DEV, QA."
    Info "Quand tu as fini, laisse le nom VIDE et appuie sur Entree."
    $ressources = @()
    while ($true) {
        Write-Host ""
        $nom = Demander "Nom Prenom (ou Entree pour terminer)"
        if ([string]::IsNullOrWhiteSpace($nom)) { break }
        $role = (Demander ("Role de {0} (PO/SM/BA/DEV/QA)" -f $nom)).ToUpper()
        $ressources += @{ nom = $nom; role = $role }
        OK ("Ajoute : {0} = {1}" -f $nom, $role)
    }
    if ($ressources.Count -eq 0) {
        Warn "Aucune ressource saisie - tu pourras les ajouter dans config.toml."
    }

    return @{ jira = $jira; chemins = $chemins; ressources = $ressources }
}

$etapes_manuelles = @()

Write-Host ""
Write-Host "###############################################" -ForegroundColor Cyan
Write-Host "#  Initialisation du projet SPII vs SP        #" -ForegroundColor Cyan
Write-Host "###############################################"

# ---------------------------------------------------------------------
# 0. Autoriser l'execution des scripts PowerShell (pour ce compte)
# ---------------------------------------------------------------------
Titre "Autorisation des scripts PowerShell"
try {
    $pol = Get-ExecutionPolicy -Scope CurrentUser
    if ($pol -in @("Restricted", "Undefined", "AllSigned")) {
        Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned -Force
        OK "ExecutionPolicy reglee sur RemoteSigned (compte utilisateur)"
    } else {
        OK "ExecutionPolicy deja permissive ($pol)"
    }
} catch {
    Warn "Impossible de regler l'ExecutionPolicy : $($_.Exception.Message)"
    $etapes_manuelles += "Lancer : Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned"
}

# ---------------------------------------------------------------------
# 1. WinPython : present ? sinon, le telecharger
# ---------------------------------------------------------------------
Titre "Python portable (WinPython)"

function Trouver-PythonExe {
    # Cherche un python.exe sous un dossier WinPython/WPy du projet.
    $candidats = Get-ChildItem -Path $PSScriptRoot -Recurse -Filter "python.exe" -ErrorAction SilentlyContinue |
        Where-Object { $_.FullName -like "*WinPython*" -or $_.FullName -like "*WPy*" }
    if (-not $candidats) { return $null }
    $prefere = $candidats | Where-Object { $_.Directory.Name -eq "python" } | Select-Object -First 1
    if ($prefere) { return $prefere.FullName }
    return ($candidats | Select-Object -First 1).FullName
}

$python = Trouver-PythonExe

if ($python) {
    OK "WinPython deja present"
    Info $python
} else {
    Warn "WinPython introuvable dans le projet."
    $rep = Read-Host "  Telecharger WinPython maintenant ? (O/N)"
    if ($rep -match '^[OoYy]') {
        try {
            $dest = Join-Path $PSScriptRoot $WP_EXE
            Info "Telechargement en cours (~17 Mo)..."
            [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
            Invoke-WebRequest -Uri $WP_URL -OutFile $dest -UseBasicParsing
            OK "Telecharge : $WP_EXE"
            Info "Decompression dans $WP_DIR ..."
            New-Item -ItemType Directory -Force -Path $WP_DIR | Out-Null
            # Installeurs recents = Inno Setup (/VERYSILENT /DIR=). Anciens = 7-Zip (-o -y).
            try {
                & $dest "/VERYSILENT" "/DIR=$WP_DIR" | Out-Null
                Start-Sleep -Seconds 3
            } catch { }
            if (-not (Trouver-PythonExe)) {
                try { & $dest "-o$WP_DIR" "-y" | Out-Null; Start-Sleep -Seconds 3 } catch { }
            }
            $python = Trouver-PythonExe
            if ($python) {
                OK "WinPython installe"
                Info $python
                Remove-Item $dest -ErrorAction SilentlyContinue
            } else {
                Erreur "Decompression faite mais python.exe introuvable."
                Warn "Lance l'exe telecharge a la main pour l'extraire dans $WP_DIR,"
                Warn "puis relance ce script."
                $etapes_manuelles += "Extraire WinPython manuellement, puis relancer ce script."
            }
        } catch {
            Erreur "Echec du telechargement : $($_.Exception.Message)"
            Warn "Sur reseau d'entreprise, le proxy peut bloquer. Telecharge WinPython"
            Warn "manuellement (voir GUIDE_WINDOWS.md) puis relance ce script."
            $etapes_manuelles += "Telecharger WinPython manuellement (voir GUIDE_WINDOWS.md)"
        }
    } else {
        Info "Telechargement ignore."
        $etapes_manuelles += "Installer WinPython (voir GUIDE_WINDOWS.md)"
    }
}

# ---------------------------------------------------------------------
# 2. Fichiers de configuration (questionnaire interactif ou copie modele)
# ---------------------------------------------------------------------
Titre "Fichiers de configuration"

$reponses = $null

function Copier-Exemple($exemple, $cible) {
    if (Test-Path $cible) {
        OK "$cible existe deja (non ecrase)"
    } elseif (Test-Path $exemple) {
        Copy-Item $exemple $cible
        OK "$cible cree a partir de $exemple"
        $script:etapes_manuelles += "Remplir $cible avec tes vraies valeurs"
    } else {
        Warn "$exemple introuvable, impossible de creer $cible"
    }
}

# Proposer le questionnaire (sauf si config.toml existe deja)
$faire_questionnaire = $false
if (Test-Path "config.toml") {
    OK "config.toml existe deja"
    $r = Read-Host "  Le RECREER via le questionnaire ? (efface l'actuel) (O/N)"
    if ($r -match '^[OoYy]') { $faire_questionnaire = $true }
} else {
    $r = Read-Host "  Remplir la config maintenant via un questionnaire ? (O/N)"
    if ($r -match '^[OoYy]') { $faire_questionnaire = $true }
}

if ($faire_questionnaire) {
    $reponses = Questionnaire-Config

    $contenu = Construire-Config $reponses.jira $reponses.chemins $reponses.ressources
    Ecrire-SansBOM (Join-Path $PSScriptRoot "config.toml") $contenu
    OK "config.toml genere"

    Write-Host ""
    Write-Host "  -- Token Jira (secrets.toml) --" -ForegroundColor White
    Info "Generer un token : https://id.atlassian.com/manage-profile/security/api-tokens"
    $token = Demander "Colle ton token Jira (laisse vide pour le mettre plus tard)"
    if ([string]::IsNullOrWhiteSpace($token)) {
        $token = "REMPLACE_PAR_TON_TOKEN"
        $etapes_manuelles += "Mettre ton token Jira dans secrets.toml"
    }
    $token = $token -replace '"', ''
    $sec = "[jira]`napi_token = ""$token""`n"
    Ecrire-SansBOM (Join-Path $PSScriptRoot "secrets.toml") $sec
    OK "secrets.toml genere"
} else {
    Info "Questionnaire ignore - copie des modeles a la place."
    Copier-Exemple "config.toml.exemple"  "config.toml"
    Copier-Exemple "secrets.toml.exemple" "secrets.toml"
}

# ---------------------------------------------------------------------
# 3. Renseigner python_exe dans config.toml (si trouve)
# ---------------------------------------------------------------------
if ($python -and (Test-Path "config.toml")) {
    Titre "Configuration du chemin Python"
    $cfgPath = Join-Path $PSScriptRoot "config.toml"
    $contenu = Get-Content $cfgPath -Raw
    if ($contenu -match "(?m)^\s*python_exe\s*=") {
        $val = $python -replace '\$', '$$$$'
        $nouveau = $contenu -replace "(?m)^\s*python_exe\s*=.*", "python_exe     = '$val'"
        Ecrire-SansBOM $cfgPath $nouveau
        OK "python_exe renseigne dans config.toml"
        Info "  -> $python"
    } else {
        Info "Ligne python_exe absente de config.toml. A ajouter manuellement :"
        Info "  python_exe = '$python'"
        $etapes_manuelles += "Ajouter dans config.toml : python_exe = '$python'"
    }
}

# ---------------------------------------------------------------------
# 4. Installer les dependances Python
# ---------------------------------------------------------------------
if ($python) {
    Titre "Installation des dependances Python"
    try {
        Info "pip install pandas openpyxl requests truststore ..."
        & $python -m pip install --quiet pandas openpyxl requests truststore
        if ($LASTEXITCODE -eq 0) {
            OK "Dependances installees"
        } else {
            Warn "pip s'est termine avec un avertissement (code $LASTEXITCODE)"
            $etapes_manuelles += "Verifier : $python -m pip install pandas openpyxl requests truststore"
        }
    } catch {
        Erreur "Echec pip : $($_.Exception.Message)"
        $etapes_manuelles += "Installer les dependances manuellement"
    }
} else {
    Warn "Python absent : dependances non installees."
}

# ---------------------------------------------------------------------
# 5. Recapitulatif
# ---------------------------------------------------------------------
Titre "Recapitulatif"

Write-Host ""
Write-Host "  Etat du projet :" -ForegroundColor White
if ($python)                    { OK "Python pret" } else { Erreur "Python a installer" }
if (Test-Path "config.toml")    { OK "config.toml present" } else { Erreur "config.toml manquant" }
if (Test-Path "secrets.toml")   { OK "secrets.toml present" } else { Erreur "secrets.toml manquant" }
if (Test-Path "spii_v2.py")     { OK "spii_v2.py present" } else { Erreur "spii_v2.py manquant" }

if ($etapes_manuelles.Count -gt 0) {
    Write-Host ""
    Write-Host "  >> Il te reste a faire MANUELLEMENT :" -ForegroundColor Yellow
    $n = 1
    foreach ($e in $etapes_manuelles) {
        Write-Host "     $n. $e" -ForegroundColor Yellow
        $n++
    }
    Write-Host ""
    Write-Host "  Pense a verifier config.toml et secrets.toml (chemins, infos" -ForegroundColor Yellow
    Write-Host "  Jira, token) avant de lancer." -ForegroundColor Yellow
} else {
    Write-Host ""
    OK "Tout est pret ! Tu peux lancer le projet avec : .\lancer.ps1"
}

Write-Host ""
Read-Host "Appuie sur Entree pour fermer"
