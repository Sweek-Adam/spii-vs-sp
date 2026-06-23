# Guide d'installation — SPII vs SP (V2) sur Windows

Ce guide explique comment faire tourner `spii_v2.py` sur une machine Windows.
Le script lui-même est multiplateforme : il n'y a rien à modifier dans le code,
seulement la configuration à adapter.

---

## 1. Installer Python en version portable (WinPython, sans droits admin)

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

---

## 2. Pas d'environnement virtuel à créer

Avec WinPython, **tu n'as pas besoin de créer un environnement virtuel** : la
distribution portable est déjà un environnement Python isolé et autonome. Tout
ce que tu installes avec pip reste contenu dans le dossier WinPython, sans
affecter le reste de la machine.

Il suffit d'utiliser le **« WinPython Command Prompt.exe »** (repéré à l'étape
précédente) pour toutes les commandes qui suivent.

---

## 3. Installer les dépendances

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

## 4. Préparer les fichiers de configuration

Dans le même dossier que `spii_v2.py`, tu dois avoir :

- **config.toml** — équipe, chemins, projet Jira (partageable)
- **secrets.toml** — ton token Jira (à NE PAS partager / versionner)

Pour config.toml, pars de `config.toml.exemple` fourni :
renomme-le en `config.toml` et adapte les deux chemins de la section
`[chemins]` (excel et csv).

### ⚠ Chemins Windows dans le TOML

L'antislash `\` est un caractère spécial en TOML. Le plus simple : utilise des
**guillemets simples**, l'antislash est alors pris tel quel :

```toml
[chemins]
excel = 'C:\Users\TonNom\Documents\SPII vs SP.xlsx'
csv   = 'C:\Users\TonNom\Documents\y25VUE9.csv'
```

Pour secrets.toml :

```toml
[jira]
api_token = "colle_ton_token_ici"
```

---

## 5. Lancer le script

Dans le **« WinPython Command Prompt.exe »**, place-toi dans le dossier qui
contient `spii_v2.py` (avec la commande `cd`), puis lance :

```
cd C:\chemin\vers\ton\projet
python spii_v2.py
```

Le script :
1. lit le CSV,
2. interroge Jira en parallèle,
3. génère un nouveau fichier **« <nom>_V2.xlsx »** à côté de ton Excel d'origine.

Ton fichier Excel d'origine n'est jamais modifié.

> Astuce confort : tu peux créer un petit fichier `lancer.bat` à côté du script
> contenant une seule ligne — le chemin complet vers le python de WinPython
> suivi de spii_v2.py — pour tout lancer d'un double-clic. Demande-le moi si tu
> veux que je te le prépare.

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
- **Comportements bizarres / chemins longs** → si tu as décompressé WinPython
  dans un dossier très profond, déplace-le vers un chemin court (ex. `C:\WPy`).

---

## Note sur l'affichage Excel

Le fichier généré est identique quelle que soit la plateforme. C'est Excel qui
l'affiche, et Excel Windows a tendance à mieux rendre certains détails
(étiquettes de camembert, graduations du nuage de points, hyperliens) que sa
version Mac. Si tu avais des petits écarts d'affichage sur Mac, ils devraient
être au moins aussi bons, voire meilleurs, sur Windows.
