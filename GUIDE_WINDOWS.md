# Guide d'installation — SPII vs SP (V2) sur Windows

Ce guide explique comment faire tourner `spii_v2.py` sur une machine Windows.
Le script lui-même est multiplateforme : il n'y a rien à modifier dans le code,
seulement la configuration à adapter.

---

## 1. Récupérer le projet depuis GitHub

Deux méthodes selon que Git est installé ou non sur ta machine.

### Méthode A — avec Git (`git clone`)

Si Git est disponible (teste avec `git --version` dans un terminal), ouvre un
terminal dans le dossier où tu veux mettre le projet, puis :

```
git clone https://github.com/Sweek-Adam/spii-vs-sp.git
cd spii-vs-sp
```

Tu obtiens un dossier `spii-vs-sp` contenant tout le projet, et tu pourras
récupérer les futures mises à jour avec un simple `git pull`.

### Méthode B — sans Git (téléchargement ZIP)

Si Git n'est pas installé (fréquent sans droits admin) :

1. Va sur la page du dépôt : https://github.com/Sweek-Adam/spii-vs-sp
2. Clique sur le bouton vert **« Code »** → **« Download ZIP »**.
3. Décompresse le ZIP où tu veux (clic droit → « Extraire tout »).
4. Tu obtiens un dossier contenant le projet (le nom peut finir par `-main`).

> Différence : avec le ZIP, tu auras le projet mais pas le lien Git. Pour
> récupérer une future mise à jour, il faudra re-télécharger le ZIP. Avec
> `git clone`, un `git pull` suffit.

---

## Voie rapide : le script d'initialisation

**Étape préalable (une seule fois) : autoriser l'exécution des scripts.**
Avant le tout premier lancement, Windows bloque par défaut l'exécution des
scripts PowerShell. Voici comment l'autoriser pour ton compte (sans droits
admin) :

1. Ouvre le menu **Démarrer**, tape **PowerShell**.
2. Clique sur **Windows PowerShell** pour l'ouvrir (un terminal bleu s'ouvre).
3. Copie-colle cette commande, puis appuie sur **Entrée** :

   ```
   Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
   ```

4. Si une confirmation est demandée, tape **O** (ou **Y**) puis **Entrée**.

Ce réglage est durable : tu n'auras à le faire qu'une seule fois sur ton poste.
Il autorise les scripts locaux (comme `initialiser.ps1`) tout en continuant de
bloquer les scripts non signés téléchargés d'Internet.

> Le script `initialiser.ps1` tente lui aussi de faire ce réglage, mais selon
> ta configuration il se peut qu'il ne puisse pas se lancer du tout tant que la
> commande ci-dessus n'a pas été passée — d'où cette étape préalable.

**Lancer l'initialisation.**
Une fois cela fait, le plus simple est de lancer le script
**`initialiser.ps1`** (clic droit → **« Exécuter avec PowerShell »**). Il
automatise presque tout le reste de ce guide :

- télécharge WinPython s'il est absent (Python portable, sans admin) ;
- installe les dépendances Python ;
- **propose un questionnaire** pour remplir `config.toml` et `secrets.toml`
  (infos Jira, chemins, équipe membre par membre, token) — tu peux le passer
  et éditer les fichiers à la main si tu préfères ;
- renseigne le chemin Python dans `config.toml` ;
- règle l'autorisation d'exécution des scripts PowerShell ;
- affiche un récapitulatif et la liste de ce qu'il te reste à faire à la main.

À la fin, il te restera essentiellement à **remplir `config.toml` et
`secrets.toml`** avec tes vraies valeurs (chemins, infos Jira, token) — voir le
chapitre 4 pour le détail de chaque champ.

> Si le script ne peut pas tout faire (par ex. téléchargement bloqué par le
> proxy d'entreprise), il te le dit clairement et les chapitres ci-dessous
> détaillent la procédure manuelle correspondante.

---

## 2. Installer Python en version portable (WinPython, sans droits admin)

Comme tu n'as pas les droits administrateur, on utilise **WinPython** : une
distribution Python **portable** qui se décompresse dans un simple dossier,
sans installation système, sans toucher au PATH, sans admin. Elle inclut déjà
pip.

> Si tu as utilisé `initialiser.ps1` (voie rapide ci-dessus), WinPython est
> déjà installé — tu peux sauter directement au chapitre 4 pour remplir tes
> fichiers de configuration.

### Télécharger

1. **Lien direct (recommandé)** — télécharge l'exécutable testé pour ce projet,
   **WinPython64-3.14.5.0dot** (Python 3.14.5, 64 bits, ~17 Mo) :
   https://github.com/winpython/winpython/releases/download/17.4.20260511final/WinPython64-3.14.5.0dot.exe

   Si ce lien ne fonctionne plus (nouvelle version publiée entre-temps), passe
   par la page officielle : https://winpython.github.io/ ou la liste des
   releases https://github.com/winpython/winpython/releases — et prends une
   version **stable** en **64 bits**, variante **"dot"** (la plus légère).

   > ⚠ **N'installe PAS la toute dernière version de Python** (ni une version
   > marquée `a`, `b` ou `rc` = alpha/beta/release candidate, ex. `3.15.0b1`).
   > Les bibliothèques comme **pandas** et **numpy** ne publient leurs versions
   > pré-compilées (« wheels ») qu'avec un délai. Sur une version Python trop
   > récente, `pip install pandas` tente de **compiler** depuis les sources et
   > échoue (il faudrait un compilateur C++ / Visual Studio que tu n'as pas).
   > Reste une version mineure ou deux derrière la toute dernière : par exemple,
   > si la dernière est 3.15, prends une 3.14 ou 3.13 stable.

### Décompresser

2. Double-clique sur le fichier `.exe` téléchargé. **Ce n'est pas un
   installeur** : il décompresse simplement un dossier à l'emplacement que tu
   choisis. Aucun droit admin requis.
3. ⚠ **Choisis un emplacement avec un chemin court** (moins de ~37 caractères).
   Par exemple `C:\WPy` ou `C:\Outils\WPy` plutôt qu'un dossier profondément
   imbriqué — WinPython le recommande pour éviter des soucis.

### Repérer le « WinPython Command Prompt »

4. Dans le dossier décompressé, tu trouveras un fichier nommé
   **`WinPython Command Prompt.exe`**. C'est LUI qu'on utilisera pour lancer
   toutes les commandes : il ouvre un terminal où `python` et `pip` pointent
   automatiquement vers ce Python portable, sans rien configurer.

   Vérifie en l'ouvrant et en tapant :

   ```
   python --version
   ```

   Tu dois voir `Python 3.13.x` (ou la version que tu as prise).

> Astuce : tu peux placer le dossier WinPython où tu veux — disque local,
> lecteur réseau, voire clé USB. Tout reste contenu dedans.

---

## 3. Installer les dépendances

Ouvre le **« WinPython Command Prompt.exe »** et tape :

```
pip install pandas openpyxl requests truststore
```

Note : WinPython récent embarque souvent déjà pandas et openpyxl (c'est une
distribution orientée science des données). Dans ce cas pip te dira simplement
« Requirement already satisfied » — c'est normal, rien à faire de plus.

Le paquet **`truststore`** est important sur le réseau d'entreprise : il permet
à Python d'utiliser les certificats de Windows pour se connecter à Jira à
travers le proxy de sécurité (sinon les appels Jira échouent avec une erreur
de certificat SSL).

Tu n'as **pas** besoin d'installer `tomli` : WinPython est en Python 3.13, qui
a déjà le lecteur TOML intégré.

---

## 4. Préparer les fichiers de configuration

Dans le même dossier que `spii_v2.py`, tu dois avoir :

- **config.toml** — équipe, chemins, paramètres Jira
- **secrets.toml** — ton token Jira (à NE PAS partager / versionner)

Pour config.toml, pars de `config.toml.exemple` fourni : renomme-le en
`config.toml`, puis renseigne chaque champ **dans l'ordre du fichier**, comme
détaillé ci-dessous.

> ⚠ **Chemins Windows** : l'antislash `\` est un caractère spécial en TOML.
> Mets tous les chemins entre **guillemets simples** `'...'` — l'antislash est
> alors pris littéralement, sans rien doubler. (Avec des guillemets doubles, il
> faudrait écrire `\\` partout.)

### Section `[jira]`

- **`email`** — ton adresse de connexion Jira (celle de ton compte Atlassian).
- **`url`** — l'adresse de ton instance Jira, ex. `https://imsa.atlassian.net`.
- **`sp_field`** — l'identifiant du champ Story Points (de la forme
  `customfield_XXXXX`). Si tu ne le connais pas, demande-le à ton admin Jira.
- **`projet`** — le préfixe des tickets à comptabiliser (ex. `LIEVRE`), utilisé
  pour filtrer les liens dans Jira.

```toml
[jira]
email    = "prenom.nom@exemple.com"
url      = "https://votre-instance.atlassian.net"
sp_field = "customfield_XXXXX"
projet   = "PROJET"
```

### Section `[chemins]`

- **`csv`** — chemin complet du fichier export SPII à lire (il doit exister).
- **`dossier_sortie`** — dossier où écrire le fichier généré (créé
  automatiquement s'il n'existe pas). Le nom du fichier produit est
  automatique et horodaté (ex. `SPII_vs_SP_2026-06-23_10h38.xlsx`).
- **`python_exe`** — chemin complet vers le `python.exe` de WinPython.
  Sert **uniquement** au script de lancement `lancer.ps1` (voir chapitre 5).
  Astuce : dans l'explorateur, Maj + clic droit sur `python.exe` →
  « Copier en tant que chemin d'accès ».

```toml
[chemins]
csv            = 'C:\Users\TonNom\Documents\export.csv'
dossier_sortie = 'C:\Users\TonNom\Documents\Sorties'
python_exe     = 'C:\Users\TonNom\Documents\spii-vs-sp\WinPython\WPy64-31450\python\python.exe'
```

### Section `[ressources]`

Cette section liste les collaborateurs à suivre, avec leur rôle. **Deux points
importants :**

- Le **nom de chaque collaborateur doit être identique au nom exact tel qu'il
  apparaît dans l'export SPII** (le fichier CSV). La moindre différence (accent,
  espace, ordre prénom/nom) fait que la personne n'est pas reconnue et sa
  consommation ignorée.
- Indique le **rôle** de chacun, parmi : `PO`, `SM`, `BA`, `DEV`, `QA`.

```toml
[ressources]
"Nom Prenom1" = "PO"
"Nom Prenom2" = "SM"
"Nom Prenom3" = "DEV"
"Nom Prenom4" = "QA"
```

### Fichier `secrets.toml` — ton token Jira

Le script se connecte à Jira avec un **token d'API** (pas ton mot de passe).
Pour le générer :

1. Va sur **https://id.atlassian.com/manage-profile/security/api-tokens**
   (connecté avec ton compte Atlassian / Jira).
2. Clique sur **« Create API token »** (Créer un token d'API).
3. Donne-lui un nom parlant, par exemple `spii-vs-sp`.
4. Choisis une date d'expiration (de 1 jour à 1 an — Atlassian impose une
   limite, note la date pour penser à le renouveler).
5. Clique **« Create »**, puis **« Copy to clipboard »**.

⚠ Le token n'est affiché **qu'une seule fois** : copie-le tout de suite. Si tu
fermes la fenêtre sans le copier, il faudra en générer un nouveau.

Crée ensuite `secrets.toml` (à partir de `secrets.toml.exemple`) et colle le
token :

```toml
[jira]
api_token = "colle_ton_token_ici"
```

---

## 5. Lancer le script

Le projet fournit **`lancer.ps1`** : clic droit dessus → **« Exécuter avec
PowerShell »**, et c'est parti. Il lit le chemin de Python depuis `config.toml`,
lance la génération, et garde la fenêtre ouverte à la fin pour que tu lises les
messages.

**Prérequis** : le champ `python_exe` doit être renseigné dans `config.toml`
(section `[chemins]`). Mets-y le chemin complet vers le `python.exe` de
WinPython, par exemple :

```toml
[chemins]
python_exe = 'C:\Users\TonNom\Documents\spii-vs-sp\WinPython\WPy64-31450\python\python.exe'
```

> Astuce pour trouver ce chemin : dans l'explorateur, ouvre le dossier
> `WinPython\WPy...\python`, fais Maj + clic droit sur `python.exe` →
> « Copier en tant que chemin d'accès », et colle-le (entre guillemets simples)
> dans config.toml.
>
> Si le champ est vide ou le chemin faux, `lancer.ps1` s'arrête avec un message
> t'indiquant quoi corriger.

**Avant le tout premier lancement**, Windows bloque par défaut l'exécution des
scripts PowerShell. Autorise-les pour ton compte (sans droits admin) en lançant
une fois cette commande dans PowerShell :

```
Set-ExecutionPolicy -Scope CurrentUser -ExecutionPolicy RemoteSigned
```

Réponds `O` (ou `Y`) si une confirmation est demandée. Ce réglage est durable
(à faire une seule fois) et sûr : il autorise les scripts locaux comme
`lancer.ps1` tout en bloquant les scripts non signés téléchargés d'Internet.
Ensuite, tu peux lancer `lancer.ps1` autant de fois que tu veux.

Le script :
1. lit le CSV,
2. interroge Jira en parallèle,
3. génère un fichier horodaté (ex. `SPII_vs_SP_2026-06-23_10h38.xlsx`) dans le
   dossier de sortie indiqué en config (créé s'il n'existe pas).

Aucun fichier existant n'est modifié. Le fichier généré s'ouvre automatiquement
à la fin.

---

## 6. En cas de souci

- **`python n'est pas reconnu`** → tu n'utilises pas le « WinPython Command
  Prompt ». Ouvre bien `WinPython Command Prompt.exe` depuis le dossier
  WinPython : c'est lui qui connaît le bon `python`.
- **`CSV introuvable`** → vérifie le chemin `csv` dans config.toml.
  Astuce : dans l'explorateur Windows, Maj + clic droit sur le fichier →
  « Copier en tant que chemin d'accès » te donne le chemin exact.
- **`Token Jira manquant`** → renseigne `api_token` dans secrets.toml.
- **Erreur de module (`No module named ...`)** → la dépendance n'est pas
  installée dans CE WinPython (refais l'étape 3 depuis le WinPython Command
  Prompt).
- **`pip install pandas` échoue** avec `Unknown compiler` / `vswhere.exe` /
  `metadata-generation-failed` → ta version de Python est trop récente, pip
  essaie de compiler faute de wheels. Reprends un WinPython en **Python 3.14
  ou 3.13 stable** (voir l'avertissement au chapitre 2).
- **Comportements bizarres / chemins longs** → si tu as décompressé WinPython
  dans un dossier très profond, déplace-le vers un chemin court (ex. `C:\WPy`).
