# SPII vs SP — Génération du suivi de consommation

Script Python qui génère un classeur Excel de suivi de consommation par TCRE,
à partir d'un export CSV et de l'API Jira (Story Points + titres).

Génère, par TCRE, des onglets détaillés avec camemberts ; un onglet Stats de
synthèse (ratios jours/SP, nuage de points, liens de navigation) ; et un onglet
par collaborateur. Le tout via **openpyxl** (pas besoin qu'Excel soit ouvert,
multiplateforme Mac/Windows).

## Prérequis

- Python 3.11+ (ou 3.10 avec le paquet `tomli`)
- Dépendances : `pip install pandas openpyxl requests`
- Sur réseau d'entreprise avec proxy SSL : ajouter `truststore`
  (`pip install truststore`) pour que Python utilise les certificats du système

## Configuration (avant le premier lancement)

Le script lit deux fichiers de config, à placer à côté de `spii_v2.py`. Aucun
des deux contenant des secrets n'est versionné — tu pars des modèles fournis :

1. **`config.toml`** — copie `config.toml.exemple` (ou crée le tien) et
   renseigne : le chemin du CSV à lire, le dossier de sortie, les paramètres
   Jira (email, url, projet), et la liste des ressources avec leur rôle.

2. **`secrets.toml`** — copie `secrets.toml.exemple` en `secrets.toml` et mets
   ton token Jira. ⚠ **Ce fichier ne doit jamais être committé** (il est dans
   le `.gitignore`).

## Lancer

```
python spii_v2.py
```

Le script lit le CSV, interroge Jira, et écrit un fichier horodaté
(ex. `SPII_vs_SP_2026-06-23_10h38.xlsx`) dans le dossier de sortie indiqué en
config. Aucun fichier existant n'est modifié.

## Installation sur Windows

Voir `GUIDE_WINDOWS.md` (procédure avec WinPython portable, sans droits admin).

## Ce qui n'est pas versionné

Voir `.gitignore` : le token (`secrets.toml`), la config locale (`config.toml`),
les données source (CSV, Excel) et les fichiers générés restent en local.
