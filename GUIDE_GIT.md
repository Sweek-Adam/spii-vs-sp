# Guide — Mettre le projet sous Git et le pousser sur GitHub

Guide pas-à-pas pour publier ce projet sur GitHub (dépôt public).
Aucune connaissance Git préalable requise.

---

## ⚠ Avant tout : la règle de sécurité n°1

Ce dépôt sera **public** (visible par tous sur Internet). Trois fichiers ne
doivent **JAMAIS** y apparaître. Ils sont déjà protégés par le `.gitignore`,
mais tu dois le **vérifier** à l'étape 4 :

- `secrets.toml` — ton token Jira
- `config.toml` — ta config locale (vrais chemins, email, noms)
- les fichiers `.csv`, `.xlsx`, `.xlsm` — tes données

Si l'un d'eux apparaît dans `git status`, **arrête-toi** et préviens-moi.

---

## 1. Installer Git (si ce n'est pas déjà fait)

- **Mac** : ouvre le Terminal et tape `git --version`. S'il n'est pas installé,
  macOS te proposera de l'installer automatiquement.
- **Windows** : télécharge Git depuis https://git-scm.com/download/win
  (ou, en portable sans admin, Git est inclus dans certaines distributions).

Vérifie :

```
git --version
```

## 2. Configurer ton identité Git (une seule fois par machine)

```
git config --global user.name "Ton Nom"
git config --global user.email "ton-email@exemple.com"
```

> Astuce : pour un dépôt public, tu peux utiliser l'email « noreply » que
> GitHub fournit (dans Settings → Emails → "Keep my email addresses private"),
> pour ne pas exposer ton vrai email dans l'historique des commits.

## 3. Initialiser le dépôt local

Place-toi dans le dossier du projet (celui qui contient `spii_v2.py`) :

```
cd /chemin/vers/ton/projet
git init
```

Cela crée un sous-dossier caché `.git` : ton dépôt local est né.

## 4. ⚠ ÉTAPE DE SÉCURITÉ : vérifier ce qui sera versionné

```
git add .
git status
```

`git status` affiche la liste des fichiers qui seront committés (en vert).

**Vérifie cette liste attentivement.** Tu DOIS y voir :
- `spii_v2.py`, `README.md`, `GUIDE_WINDOWS.md`, `GUIDE_GIT.md`
- `config.toml.exemple`, `secrets.toml.exemple`, `.gitignore`

Tu ne dois **PAS** y voir :
- `secrets.toml`, `config.toml`
- aucun fichier `.csv`, `.xlsx`, `.xlsm`

Si un fichier interdit apparaît : tape `git rm --cached <fichier>` pour le
retirer de la zone de commit (sans le supprimer du disque), et vérifie le
`.gitignore`.

## 5. Premier commit

Une fois la liste validée :

```
git commit -m "Initial commit : générateur SPII vs SP (V2 openpyxl)"
```

## 6. Créer le dépôt sur GitHub

1. Va sur https://github.com et connecte-toi (crée un compte si besoin).
2. Clique sur le **+** en haut à droite → **New repository**.
3. Donne un nom (ex. `spii-vs-sp`), choisis **Public**.
4. **Ne coche RIEN** (pas de README, pas de .gitignore, pas de licence) : ton
   projet local les a déjà. Un dépôt vide évite les conflits.
5. Clique **Create repository**.

GitHub affiche alors une page avec des commandes. Utilise celles de la section
**"…or push an existing repository from the command line"**.

## 7. Relier le local à GitHub et pousser

GitHub te donne deux lignes à copier (remplace l'URL par la tienne) :

```
git remote add origin https://github.com/TonPseudo/spii-vs-sp.git
git branch -M main
git push -u origin main
```

GitHub te demandera de t'authentifier. La première fois, le plus simple est de
suivre l'invite (connexion via navigateur), ou d'utiliser un **token d'accès
personnel** (Settings → Developer settings → Personal access tokens) en guise
de mot de passe.

## 8. Vérifier

Recharge la page de ton dépôt GitHub : tes fichiers doivent y être.
**Vérifie une dernière fois** qu'il n'y a ni `secrets.toml`, ni `config.toml`,
ni données.

---

## Pour la suite : workflow quotidien

À chaque modification du code que tu veux enregistrer :

```
git add .
git status                          # toujours vérifier avant de committer
git commit -m "Décris ton changement"
git push
```

## En cas de souci

- **« fatal: not a git repository »** → tu n'es pas dans le bon dossier, ou
  `git init` n'a pas été fait. Refais l'étape 3.
- **« remote origin already exists »** → le lien existe déjà ; passe directement
  au `git push`.
- **Tu as committé un secret par erreur** → ne te contente pas de le supprimer
  dans un nouveau commit (il reste dans l'historique). Préviens-moi : il faut
  réécrire l'historique ET régénérer le token immédiatement.

---

## Si tu hésites entre public et privé

Tu as choisi public (utile pour montrer ton travail, par exemple dans une
optique d'évolution vers le PM/produit). C'est un bon choix tant que les données
internes sont neutralisées — ce qui est le cas dans les fichiers versionnés
(emails, URL Jira, noms d'équipe et projet sont génériques dans les exemples).

Si un jour tu préfères basculer en privé : sur GitHub, Settings du dépôt →
section "Danger Zone" → "Change repository visibility".
