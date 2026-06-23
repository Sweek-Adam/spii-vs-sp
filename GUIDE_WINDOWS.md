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

## 2. Installer Python en version portable (WinPython, sans droits admin)

Comme tu n'as pas les droits administrateur, on utilise **WinPython** : une
distribution Python **portable** qui se décompresse dans un simple dossier,
sans installation système, sans toucher au PATH, sans admin. Elle inclut déjà
pip.

### Télécharger

1. Va sur la page officielle : https://winpython.github.io/
   (ou les téléchargements : https://github.com/winpython/winpython/releases)
2. Prends la dernière version, en **64 bits**, variante **"dot"** (la plus
   légère : Python seul, suffisant pour ce script). Le fichier ressemble à
   `Winpython64-3.13.x.0dot.exe`.

### Décompresser

3. Double-clique sur le fichier `.exe` téléchargé. **Ce n'est pas un
   installeur** : il décompresse simplement un dossier à l'emplacement que tu
   choisis. Aucun droit admin requis.
4. ⚠ **Choisis un emplacement avec un chemin court** (moins de ~37 caractères).
   Par exemple `C:\WPy` ou `C:\Outils\WPy` plutôt qu'un dossier profondément
   imbriqué — WinPython le recommande pour éviter des soucis.

### Repérer le « WinPython Command Prompt »

5. Dans le dossier décompressé, tu trouveras un fichier nommé
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

### (Optionnel mais pratique) Ajouter WinPython au PATH

Par défaut, `python` ne fonctionne que dans le « WinPython Command Prompt ». Si
tu veux pouvoir taper `python` dans **n'importe quel** terminal (PowerShell, cmd)
— pratique notamment pour lancer le script via un raccourci — tu peux ajouter
WinPython au PATH de ton compte utilisateur. **Sans droits admin.**

Le projet fournit un script qui fait ça pour toi : **`ajouter-winpython-au-path.ps1`**

1. Clic droit sur `ajouter-winpython-au-path.ps1` → **« Exécuter avec PowerShell »**.
2. Quand c'est demandé, colle le chemin de ton dossier WinPython (ex. `C:\WPy64-31360`).
3. Le script détecte automatiquement le bon sous-dossier Python et l'ajoute à
   ton PATH utilisateur.
4. **Ferme puis rouvre tes terminaux** pour que le changement prenne effet.
5. Teste dans un terminal normal : `python --version` doit répondre.

> Si PowerShell refuse d'exécuter le script (« exécution de scripts désactivée »),
> ouvre PowerShell et lance d'abord cette commande (sans admin, pour ta session
> uniquement) :
> ```
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```
> puis relance le script.

Une fois le PATH configuré, tu n'es plus obligée de passer par le « WinPython
Command Prompt » — un terminal classique suffit pour les étapes suivantes.

---

## 3. Pas d'environnement virtuel à créer

Avec WinPython, **tu n'as pas besoin de créer un environnement virtuel** : la
distribution portable est déjà un environnement Python isolé et autonome. Tout
ce que tu installes avec pip reste contenu dans le dossier WinPython, sans
affecter le reste de la machine.

Il suffit d'utiliser le **« WinPython Command Prompt.exe »** (repéré à l'étape
précédente) pour toutes les commandes qui suivent.

---

## 4. Installer les dépendances

Ouvre le **« WinPython Command Prompt.exe »** et tape :

```
pip install pandas openpyxl requests
```

Note : WinPython récent embarque souvent déjà pandas et openpyxl (c'est une
distribution orientée science des données). Dans ce cas pip te dira simplement
« Requirement already satisfied » — c'est normal, rien à faire de plus.

Tu n'as **pas** besoin d'installer `tomli` : WinPython est en Python 3.13, qui
a déjà le lecteur TOML intégré.

---

## 5. Préparer les fichiers de configuration

Dans le même dossier que `spii_v2.py`, tu dois avoir :

- **config.toml** — équipe, chemins, projet Jira (partageable)
- **secrets.toml** — ton token Jira (à NE PAS partager / versionner)

Pour config.toml, pars de `config.toml.exemple` fourni :
renomme-le en `config.toml` et adapte les chemins de la section
`[chemins]` : `csv` (le fichier à lire) et `dossier_sortie` (où écrire).

### ⚠ Chemins Windows dans le TOML

L'antislash `\` est un caractère spécial en TOML. Le plus simple : utilise des
**guillemets simples**, l'antislash est alors pris tel quel :

```toml
[chemins]
csv            = 'C:\Users\TonNom\Documents\export.csv'
dossier_sortie = 'C:\Users\TonNom\Documents\Sorties'
```

Pour secrets.toml :

```toml
[jira]
api_token = "colle_ton_token_ici"
```

---

## 6. Lancer le script

### Méthode simple — le script `lancer.ps1`

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

### Méthode manuelle — en ligne de commande

Sinon, en ligne de commande, en appelant directement le python.exe de
WinPython (remplace par ton chemin) :

```
& 'C:\...\WinPython\WPy64-31450\python\python.exe' spii_v2.py
```

(ou simplement `python spii_v2.py` si tu as ajouté WinPython au PATH — voir la
section optionnelle 2.)

### Dans les deux cas

Le script :
1. lit le CSV,
2. interroge Jira en parallèle,
3. génère un fichier horodaté (ex. `SPII_vs_SP_2026-06-23_10h38.xlsx`) dans le
   dossier de sortie indiqué en config (créé s'il n'existe pas).

Aucun fichier existant n'est modifié. Le fichier généré s'ouvre automatiquement
à la fin.

---

## 7. En cas de souci

- **`python n'est pas reconnu`** → tu n'utilises pas le « WinPython Command
  Prompt ». Ouvre bien `WinPython Command Prompt.exe` depuis le dossier
  WinPython : c'est lui qui connaît le bon `python`.
- **`CSV introuvable`** → vérifie le chemin `csv` dans config.toml.
  Astuce : dans l'explorateur Windows, Maj + clic droit sur le fichier →
  « Copier en tant que chemin d'accès » te donne le chemin exact.
- **`Token Jira manquant`** → renseigne `api_token` dans secrets.toml.
- **Erreur de module (`No module named ...`)** → la dépendance n'est pas
  installée dans CE WinPython (refais l'étape 4 depuis le WinPython Command
  Prompt).
- **Comportements bizarres / chemins longs** → si tu as décompressé WinPython
  dans un dossier très profond, déplace-le vers un chemin court (ex. `C:\WPy`).

---

## Note sur l'affichage Excel

Le fichier généré est identique quelle que soit la plateforme. C'est Excel qui
l'affiche, et Excel Windows a tendance à mieux rendre certains détails
(étiquettes de camembert, graduations du nuage de points, hyperliens) que sa
version Mac. Si tu avais des petits écarts d'affichage sur Mac, ils devraient
être au moins aussi bons, voire meilleurs, sur Windows.
