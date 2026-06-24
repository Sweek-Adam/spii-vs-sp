# SPII vs SP — Génération du suivi de consommation

Script Python qui génère un classeur Excel de suivi de consommation des features
(par défaut « TCRE »), à partir d'un export CSV et de l'API Jira (Story Points,
titres, Planning Interval et statut).

Il produit :

- un onglet **Stats** de synthèse : par feature, les Story Points, la conso
  totale, le ratio jours/SP, la ventilation par rôle (PO/SM, BA, Dévs, QA), le
  statut, le Planning Interval, des liens vers Jira et vers le détail ; plus un
  tableau de moyennes par complexité et un nuage de points SP vs jours ;
- un onglet **Suivi_Features_<projet>** : conso mensuelle par feature, avec PI,
  statut et un dégradé de couleur sur le total ;
- un onglet **par collaborateur** : ses lignes de conso avec dégradé ;
- un onglet **par feature** (conso > 0) : tableaux par profil et par
  collaborateur (jours **et %**) avec camemberts, et liens de navigation.

Le tout via **openpyxl** (pas besoin qu'Excel soit ouvert, multiplateforme
Mac/Windows).

## Prérequis

- Python 3.11+ (le lecteur TOML est intégré ; sur 3.10, ajouter `tomli`)
- Dépendances : `pip install pandas openpyxl requests truststore`
  - `truststore` permet, sur réseau d'entreprise avec proxy SSL, d'utiliser les
    certificats du système pour joindre Jira (sinon erreur de certificat).

## Configuration (avant le premier lancement)

Le script lit deux fichiers, à placer à côté de `spii_v2.py`. Aucun des deux
n'est versionné (ils contiennent des données locales / un secret) — pars des
modèles fournis :

1. **`config.toml`** — copie `config.toml.exemple` et renseigne :
   - `[jira]` : `email`, `url`, `sp_field`, `pi_field`, `projet`,
     `prefixe_feature` (le préfixe des features à suivre, ex. `TCRE`) ;
   - `[chemins]` : `csv` (export à lire), `dossier_sortie`, `python_exe`
     (chemin du Python à utiliser, surtout pour le lancement Windows) ;
   - `[ressources]` : la liste des collaborateurs et leur rôle (le nom doit être
     **identique** à la colonne Ressource du CSV).

2. **`secrets.toml`** — copie `secrets.toml.exemple` et mets ton token Jira.
   ⚠ **Ne jamais committer ce fichier** (il est dans le `.gitignore`).

## Lancer

```
python spii_v2.py
```

Le script lit le CSV, interroge Jira, et écrit un fichier nommé d'après le
projet et horodaté (ex. `SPII_vs_SP_LIEVRE_2026-06-23_10h38.xlsx`) dans le
dossier de sortie indiqué en config. Aucun fichier existant n'est modifié ; le
fichier généré s'ouvre automatiquement à la fin.

## Installation et lancement sur Windows

Sans droits admin, on utilise WinPython (Python portable). Deux scripts
PowerShell facilitent tout :

- **`initialiser.ps1`** — assistant d'installation : télécharge WinPython si
  besoin, installe les dépendances, propose un questionnaire pour remplir la
  config et le token, et règle l'autorisation d'exécution des scripts.
- **`lancer.ps1`** — lance la génération (clic droit → « Exécuter avec
  PowerShell »), en lisant le chemin Python depuis `config.toml`.

Voir **`GUIDE_WINDOWS.md`** pour la procédure détaillée pas à pas.

## Ce qui n'est pas versionné

Voir `.gitignore` : le token (`secrets.toml`), la config locale (`config.toml`),
les données source (CSV, Excel) et les fichiers générés restent en local.
