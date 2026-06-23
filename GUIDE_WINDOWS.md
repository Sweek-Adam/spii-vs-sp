# Guide d'installation — SPII vs SP (V2) sur Windows

Ce guide explique comment faire tourner `spii_v2.py` sur une machine Windows.
Le script lui-même est multiplateforme : il n'y a rien à modifier dans le code,
seulement la configuration à adapter.

---

## 1. Installer Python

1. Télécharge Python depuis https://www.python.org/downloads/windows/
   (prends la version 3.11 ou plus récente — recommandé).
2. À l'installation, **coche la case « Add Python to PATH »** en bas de la
   première fenêtre. C'est important, sinon la commande `python` ne sera pas
   reconnue.
3. Vérifie l'installation en ouvrant l'invite de commandes (PowerShell ou cmd) :

   ```
   python --version
   ```

   Tu dois voir quelque chose comme `Python 3.12.x`.

---

## 2. Créer un environnement virtuel (recommandé)

Dans le dossier de ton projet (celui qui contient `spii_v2.py`), ouvre un
terminal et tape :

```
python -m venv .venv
.venv\Scripts\activate
```

Une fois activé, ton invite affiche `(.venv)` au début de la ligne.

---

## 3. Installer les dépendances

Avec l'environnement activé :

```
pip install pandas openpyxl requests
```

Si tu es sur **Python 3.10 ou antérieur**, ajoute aussi :

```
pip install tomli
```

(Python 3.11+ a déjà le lecteur TOML intégré, rien à installer.)

---

## 4. Préparer les fichiers de configuration

Dans le même dossier que `spii_v2.py`, tu dois avoir :

- **config.toml** — équipe, chemins, projet Jira (partageable)
- **secrets.toml** — ton token Jira (à NE PAS partager / versionner)

Pour config.toml, pars de `config.windows.exemple.toml` fourni :
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

Toujours avec l'environnement activé, dans le dossier du projet :

```
python spii_v2.py
```

Le script :
1. lit le CSV,
2. interroge Jira en parallèle,
3. génère un nouveau fichier **« <nom>_V2.xlsx »** à côté de ton Excel d'origine.

Ton fichier Excel d'origine n'est jamais modifié.

---

## 6. En cas de souci

- **`python n'est pas reconnu`** → Python n'a pas été ajouté au PATH
  (réinstalle en cochant la case, ou utilise `py` au lieu de `python`).
- **`CSV introuvable`** → vérifie le chemin `csv` dans config.toml.
  Astuce : dans l'explorateur Windows, Maj + clic droit sur le fichier →
  « Copier en tant que chemin d'accès » te donne le chemin exact.
- **`Token Jira manquant`** → renseigne `api_token` dans secrets.toml.
- **Erreur de module (`No module named ...`)** → l'environnement virtuel n'est
  pas activé, ou la dépendance n'est pas installée (refais l'étape 3).

---

## Note sur l'affichage Excel

Le fichier généré est identique quelle que soit la plateforme. C'est Excel qui
l'affiche, et Excel Windows a tendance à mieux rendre certains détails
(étiquettes de camembert, graduations du nuage de points, hyperliens) que sa
version Mac. Si tu avais des petits écarts d'affichage sur Mac, ils devraient
être au moins aussi bons, voire meilleurs, sur Windows.
