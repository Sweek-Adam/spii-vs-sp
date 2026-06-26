#!/usr/bin/env python3
# =====================================================================
# fusionner_csv.py — Concatène plusieurs exports CSV consécutifs en un
# seul CSV continu, prêt à être donné à spii_v2.py.
# ---------------------------------------------------------------------
# Chaque export a le format attendu par spii_v2.py :
#   colonnes Ressource, Projet, Livrable, Type, puis "Mois de référence"
#   suivie de 11 colonnes "M/AAAA".
#
# Le script :
#   - détecte la période de chaque export (année / mois de départ variables) ;
#   - construit une timeline continue couvrant TOUS les mois rencontrés ;
#   - relie les lignes d'un export à l'autre par la clé
#     (Ressource, Projet, Livrable, Type) ;
#   - en cas de mois présent dans plusieurs exports (chevauchement), la
#     valeur du CSV le PLUS RÉCENT l'emporte (écrase) ;
#   - écrit un seul CSV au même format, avec autant de colonnes mensuelles
#     que la période totale.
#
# Le fichier produit se donne directement à spii_v2.py (champ 'csv' de
# config.toml) : celui-ci gère un nombre de mois quelconque.
#
# Usage :
#   python fusionner_csv.py export1.csv export2.csv [...] -o fusion.csv
#   python fusionner_csv.py *.csv -o fusion.csv
# =====================================================================

import argparse
import csv
import re
import sys
from datetime import date

COL_MOIS_REF = "Mois de référence"
# Colonnes non-mensuelles qui identifient une ligne (et qu'on recopie telles
# quelles dans la sortie).
COLS_CLE = ["Ressource", "Projet", "Livrable", "Type"]


def _lire_csv(chemin):
    """Lit un CSV en détectant le séparateur ; renvoie (entetes, lignes)."""
    # On lit en latin1 comme spii_v2.py (exports Windows/Excel FR).
    with open(chemin, "r", encoding="latin1", newline="") as f:
        echantillon = f.read(4096)
        f.seek(0)
        try:
            dialecte = csv.Sniffer().sniff(echantillon, delimiters=";,\t")
            sep = dialecte.delimiter
        except csv.Error:
            sep = ";"  # défaut raisonnable pour un export FR
        lecteur = csv.reader(f, delimiter=sep)
        lignes = list(lecteur)
    if not lignes:
        raise ValueError(f"Fichier vide : {chemin}")
    return lignes[0], lignes[1:], sep


def detecter_mois(entetes, chemin):
    """Repère les colonnes de mois et calcule leur date (1er du mois).

    Renvoie (idx0, dates) où idx0 est l'index de 'Mois de référence' et dates
    la liste des 12 dates correspondantes. Même logique que spii_v2.py.
    """
    if COL_MOIS_REF not in entetes:
        raise ValueError(
            f"Colonne '{COL_MOIS_REF}' introuvable dans {chemin}. "
            f"En-têtes : {entetes}")
    i0 = entetes.index(COL_MOIS_REF)
    cols_mois = entetes[i0:i0 + 12]
    if len(cols_mois) < 12:
        raise ValueError(
            f"Seulement {len(cols_mois)} colonne(s) de mois dans {chemin}, "
            f"il en faut 12.")
    # La 2e colonne (ex. "4/2025") donne le mois suivant le mois de référence.
    m = re.match(r"\s*(\d{1,2})\s*/\s*(\d{4})\s*", str(cols_mois[1]))
    if not m:
        raise ValueError(
            f"Format inattendu pour la 2e colonne de mois dans {chemin} : "
            f"'{cols_mois[1]}'. Attendu 'M/AAAA' (ex. '4/2025').")
    mois2, annee2 = int(m.group(1)), int(m.group(2))
    ref_mois, ref_annee = mois2 - 1, annee2
    if ref_mois == 0:
        ref_mois, ref_annee = 12, annee2 - 1
    dates = []
    y, mth = ref_annee, ref_mois
    for _ in range(12):
        dates.append(date(y, mth, 1))
        mth += 1
        if mth == 13:
            mth, y = 1, y + 1
    return i0, dates


def _cle_ligne(ligne, idx_cle):
    """Clé d'identité d'une ligne (Ressource, Projet, Livrable, Type)."""
    return tuple(str(ligne[i]).strip() if i < len(ligne) else "" for i in idx_cle)


def _conv_num(v):
    """Convertit une valeur en float (virgule FR tolérée). Vide/illisible -> 0."""
    if v is None:
        return 0.0
    s = str(v).strip().replace("\xa0", "").replace(" ", "")
    if not s:
        return 0.0
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0


def fusionner(chemins, sortie):
    """Fusionne les exports (ordre = du plus ancien au plus récent fourni)."""
    # On lit tous les exports et on mémorise, pour chaque ligne-clé, la valeur
    # de chaque mois (date -> valeur). L'ordre de traitement fait que le dernier
    # fichier écrit en dernier : le plus récent gagne sur les mois en double.
    #
    # data[cle] = {"infos": [Ressource, Projet, Livrable, Type],
    #              "mois": {date: valeur}}
    data = {}
    tous_les_mois = set()

    for chemin in chemins:
        entetes, lignes, _ = _lire_csv(chemin)
        idx_cle = [entetes.index(c) if c in entetes else -1 for c in COLS_CLE]
        if -1 in idx_cle:
            manquantes = [c for c, i in zip(COLS_CLE, idx_cle) if i == -1]
            raise ValueError(
                f"Colonnes manquantes dans {chemin} : {manquantes}")
        i0, dates = detecter_mois(entetes, chemin)
        tous_les_mois.update(dates)

        for ligne in lignes:
            if not any(str(x).strip() for x in ligne):
                continue  # ligne vide
            cle = _cle_ligne(ligne, idx_cle)
            entree = data.setdefault(
                cle, {"infos": [str(ligne[i]).strip() if i >= 0 and i < len(ligne)
                                else "" for i in idx_cle],
                      "mois": {}})
            # Écrase les valeurs des mois de cet export (le plus récent gagne).
            for j, d in enumerate(dates):
                col = i0 + j
                val = ligne[col] if col < len(ligne) else ""
                entree["mois"][d] = _conv_num(val)

    if not tous_les_mois:
        raise ValueError("Aucun mois détecté dans les fichiers fournis.")

    mois_tries = sorted(tous_les_mois)
    n_mois = len(mois_tries)

    # En-têtes de sortie : colonnes-clé + "Mois de référence" + (n-1) "M/AAAA".
    entetes_sortie = list(COLS_CLE) + [COL_MOIS_REF] + [
        f"{d.month}/{d.year}" for d in mois_tries[1:]]

    with open(sortie, "w", encoding="latin1", newline="") as f:
        ecrivain = csv.writer(f, delimiter=";")
        ecrivain.writerow(entetes_sortie)
        for cle in sorted(data.keys()):
            entree = data[cle]
            valeurs = [entree["mois"].get(d, 0.0) for d in mois_tries]
            # Formatage FR : virgule décimale, sans .0 inutile sur les entiers
            valeurs_fmt = []
            for v in valeurs:
                if v == int(v):
                    valeurs_fmt.append(str(int(v)))
                else:
                    valeurs_fmt.append(str(v).replace(".", ","))
            ecrivain.writerow(entree["infos"] + valeurs_fmt)

    return mois_tries, n_mois, len(data)


def main():
    parseur = argparse.ArgumentParser(
        description="Fusionne plusieurs exports CSV consécutifs en un seul.")
    parseur.add_argument("fichiers", nargs="+",
                         help="CSV à fusionner, DU PLUS ANCIEN AU PLUS RÉCENT "
                              "(le plus récent gagne sur les mois en double).")
    parseur.add_argument("-o", "--sortie", default="fusion.csv",
                         help="Fichier CSV de sortie (défaut : fusion.csv).")
    args = parseur.parse_args()

    try:
        mois_tries, n_mois, n_lignes = fusionner(args.fichiers, args.sortie)
    except (ValueError, FileNotFoundError) as e:
        print(f"❌ {e}")
        sys.exit(1)

    debut = mois_tries[0].strftime("%m/%Y")
    fin = mois_tries[-1].strftime("%m/%Y")
    print(f"✅ Fusion écrite : {args.sortie}")
    print(f"   Période continue : {debut} -> {fin}  ({n_mois} mois)")
    print(f"   Lignes (features/livrables) : {n_lignes}")
    print(f"\n   Donne ce fichier à spii_v2.py (champ 'csv' de config.toml) :")
    print(f"   il gère un nombre de mois quelconque.")


if __name__ == "__main__":
    main()
