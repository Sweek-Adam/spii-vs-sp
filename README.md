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

En plus des features `<prefixe>-xxx`, certaines **catégories nommées** repérées
dans le libellé du livrable (« Correction SOLAU », « Analyse N2 ») reçoivent leur
propre onglet détaillé, avec le même contenu (profils, collaborateurs, %,
camemberts). La liste de ces catégories est définie en haut de `spii_v2.py`
(`CATEGORIES_LIVRABLE`) et facile à étendre.

Les tableaux principaux (Stats, Suivi_Features, collaborateurs) ont des **filtres
automatiques** sur leurs en-têtes pour trier et filtrer facilement.

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

### Format du CSV attendu

Le CSV doit contenir les colonnes `Ressource`, `Projet`, `Livrable`, `Type`,
puis une colonne **`Mois de référence`** suivie de **colonnes mensuelles** au
format `M/AAAA` (ex. `4/2025`, `5/2025`, …). La période (année, mois de départ et
**nombre de mois**) est **détectée automatiquement** : pas besoin que le CSV
commence en janvier ni qu'il couvre exactement 12 mois. Les consommations doivent
être de type `consomme` (les autres types sont ignorés).

### Fusionner plusieurs exports consécutifs (optionnel)

Si tu as plusieurs exports qui se suivent dans le temps (ex. une année puis la
suivante) et que tu veux une vue continue, le script **`fusionner_csv.py`** les
concatène en un seul CSV :

```
python fusionner_csv.py export_2025.csv export_2026.csv -o fusion.csv
```

Donne les fichiers **du plus ancien au plus récent**. Les lignes sont reliées par
leur clé (Ressource + Projet + Livrable + Type), et si deux exports partagent des
mois, le plus récent l'emporte. Le fichier `fusion.csv` produit se donne ensuite
à `spii_v2.py` comme n'importe quel CSV (il gère un nombre de mois quelconque).

### Absences (optionnel)

Si le champ `absences` de `[chemins]` pointe vers le fichier Excel
« CREA - Absences AAAA.xlsx », le rapport ajoute un onglet **Absences** (total
annuel par personne, détail par personne/période, totaux par PI) et reprend, dans
chaque onglet collaborateur, un encart des jours d'absence/présence de l'année.
Les noms sont rapprochés automatiquement (accents, casse et ordre des mots
tolérés). Laisse le champ vide pour ignorer cette partie. Le module
`absences.py` doit être présent à côté de `spii_v2.py`.

Le fichier d'absences vit sur SharePoint/OneDrive, qui ne se télécharge pas par
une URL : il faut le **chemin local de ta copie synchronisée**. Pour le trouver :
ouvre le fichier depuis l'app OneDrive/Explorateur, fais un clic droit dessus →
« Copier le chemin d'accès » (Windows) ou « Copier » puis colle-le. Il ressemble
à `C:\Users\TonNom\OneDrive - MSA\...\CREA - Absences 2026.xlsx`. Colle ce chemin
dans le champ `absences`. Si le fichier n'est pas encore synchronisé en local,
ouvre-le une fois depuis OneDrive pour déclencher la synchro.

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
