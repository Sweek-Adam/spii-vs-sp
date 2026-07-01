#!/usr/bin/env python3
# =====================================================================
# absences.py — Lecture du fichier Excel "CREA - Absences AAAA.xlsx"
# ---------------------------------------------------------------------
# Adapté du script d'origine (main.py) pour être appelé par spii_v2.py :
#   - pas de mode interactif (aucune question posée) ;
#   - pas de dépendance externe (jours fériés FR calculés directement) ;
#   - renvoie des DataFrames prêts à écrire dans le classeur.
#
# Structure attendue du fichier (par feuille / semestre) :
#   ligne 2 (idx 1) : mois (cellules fusionnées -> ffill)
#   ligne 3 (idx 2) : périodes / PI (ffill)
#   ligne 4 (idx 3) : jours de semaine (S/D = week-end à ignorer)
#   ligne 5 (idx 4) : numéros de jour
#   col A : Nom Prénom | col B : équipe | col C : type
#   cellules : codes d'absence (A, 1/2A, TP, ...) valorisés en jours
# =====================================================================

import re
from datetime import date, datetime, timedelta

import pandas as pd

# Codes de congé -> valeur en jours d'absence
LEAVE_MAPPING = {
    "A": 1.0, "1/2A": 0.5, "TP": 1.0, "1/2TP": 0.5, "D": 1.0, "F": 1.0,
    "P": 1.0, "CP": 1.0, "CONGÉ": 1.0, "CA": 1.0,
}

MONTH_MAPPING = {
    "JANVIER": 1, "FEVRIER": 2, "MARS": 3, "AVRIL": 4, "MAI": 5, "JUIN": 6,
    "JUILLET": 7, "AOUT": 8, "SEPTEMBRE": 9, "OCTOBRE": 10, "NOVEMBRE": 11,
    "DECEMBRE": 12,
}


def _paques(annee):
    """Dimanche de Pâques (algorithme de Butcher)."""
    a = annee % 19
    b = annee // 100
    c = annee % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    l = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * l) // 451
    mois = (h + l - 7 * m + 114) // 31
    jour = ((h + l - 7 * m + 114) % 31) + 1
    return date(annee, mois, jour)


def feries_fr(annee):
    """Jours fériés français (métropole) pour une année."""
    p = _paques(annee)
    return {
        date(annee, 1, 1), date(annee, 5, 1), date(annee, 5, 8),
        date(annee, 7, 14), date(annee, 8, 15), date(annee, 11, 1),
        date(annee, 11, 11), date(annee, 12, 25),
        p + timedelta(days=1),    # lundi de Pâques
        p + timedelta(days=39),   # Ascension
        p + timedelta(days=50),   # lundi de Pentecôte
    }


def _traiter_feuille(file_path, sheet_name, annee, fr_holidays):
    """Extrait les enregistrements jour par jour d'une feuille (semestre)."""
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    months = df.iloc[1, :].ffill()
    periods = df.iloc[2, :].ffill().apply(
        lambda x: str(x).strip() if pd.notna(x) else x)
    day_names = df.iloc[3, :]
    days = df.iloc[4, :]

    day_numbers = pd.to_numeric(days, errors="coerce")
    date_col_indices = day_numbers.index[day_numbers.notna()].tolist()

    name_col_index, team_col_index = 0, 1
    data_rows = df.iloc[5:, :]

    # Colonnes "nb jrs travaillés" à ignorer
    ignore = [idx for idx in date_col_indices
              if str(periods[idx]).upper() == "NB JRS TRAVAILLÉS"]

    records = []
    for _, row in data_rows.iterrows():
        person = str(row[name_col_index]).strip() if pd.notna(row[name_col_index]) else ""
        team = str(row[team_col_index]).strip() if pd.notna(row[team_col_index]) else "Sans Équipe"
        if person == "":
            continue
        for col_idx in date_col_indices:
            if col_idx in ignore:
                continue
            day_label = str(day_names[col_idx]).strip().upper()
            if day_label in ("S", "D"):
                continue  # week-end
            month_name = str(months[col_idx]).strip().upper().replace("\xa0", " ")
            if month_name not in MONTH_MAPPING:
                continue
            day_num = int(days[col_idx])
            current_date = date(annee, MONTH_MAPPING[month_name], day_num)
            if current_date in fr_holidays:
                continue  # férié
            raw = row[col_idx]
            code = str(raw).strip().upper() if pd.notna(raw) else ""
            val_abs = LEAVE_MAPPING.get(code, 0.0)
            val_pres = 1.0 - val_abs
            if val_abs > 0:
                records.append({"Personne": person, "Equipe": team,
                                "Date": current_date, "Periode": periods[col_idx],
                                "Type": "Absence", "Valeur": val_abs})
            if val_pres > 0:
                records.append({"Personne": person, "Equipe": team,
                                "Date": current_date, "Periode": periods[col_idx],
                                "Type": "Présence", "Valeur": val_pres})
    return records


def charger_absences(file_path, feuilles=None):
    """Lit le fichier d'absences et renvoie un dict de DataFrames :
      - 'par_personne_periode' : Absence/Présence par équipe, personne, période
      - 'total_annuel'         : Absence/Présence par équipe, personne (année)
      - 'par_pi'               : Absence/Présence par groupe de période (PI)
      - 'brut'                 : le détail jour par jour (pour rapprochements)
    Renvoie None si le fichier est illisible ou vide.
    """
    # Année depuis le nom de fichier (…AAAA.xlsx), sinon année courante
    m = re.search(r"(\d{4})", str(file_path))
    annee = int(m.group(1)) if m else datetime.now().year
    fr_holidays = feries_fr(annee)

    if feuilles is None:
        # Feuilles par défaut : les deux semestres de l'année détectée
        feuilles = [f"1er SEMESTRE {annee}", f"2eme SEMESTRE {annee}"]

    tous = []
    for sh in feuilles:
        try:
            tous.extend(_traiter_feuille(file_path, sh, annee, fr_holidays))
        except Exception as e:
            print(f"   [absences] feuille '{sh}' ignorée : {e}")

    if not tous:
        return None

    df = pd.DataFrame(tous)
    df = df.drop_duplicates(subset=["Personne", "Date", "Type"])
    # Groupe de période : "PI 6 - Itération IP" -> "PI 6"
    df["Groupe_Periode"] = df["Periode"].apply(
        lambda x: str(x).split(" - ")[0] if " - " in str(x) else str(x))

    def pivot(cle):
        g = df.groupby(cle + ["Type"])["Valeur"].sum().unstack(fill_value=0)
        for col in ("Absence", "Présence"):
            if col not in g.columns:
                g[col] = 0
        return g[["Absence", "Présence"]].reset_index()

    par_personne_periode = pivot(["Equipe", "Personne", "Periode"]).rename(
        columns={"Absence": "Jours d'absence", "Présence": "Jours de présence"}
    ).sort_values(["Equipe", "Personne", "Periode"])

    total_annuel = pivot(["Equipe", "Personne"]).rename(
        columns={"Absence": "Total absences", "Présence": "Total présences"}
    ).sort_values(["Equipe", "Personne"])
    total_annuel["Total jours"] = (total_annuel["Total absences"]
                                   + total_annuel["Total présences"])

    par_pi = pivot(["Groupe_Periode"]).rename(
        columns={"Absence": "Total absences", "Présence": "Total présences"}
    ).sort_values(["Groupe_Periode"])

    # Dates de début / fin observées (1er et dernier jour ouvré avec activité)
    # par groupe de période (PI) et par sous-période (itération).
    bornes_pi = (df.groupby("Groupe_Periode")["Date"].agg(["min", "max"])
                 .to_dict("index"))
    bornes_periode = (df.groupby("Periode")["Date"].agg(["min", "max"])
                      .to_dict("index"))

    # On ajoute les colonnes Début / Fin au tableau par_pi.
    par_pi["Début"] = par_pi["Groupe_Periode"].map(
        lambda g: bornes_pi.get(g, {}).get("min"))
    par_pi["Fin"] = par_pi["Groupe_Periode"].map(
        lambda g: bornes_pi.get(g, {}).get("max"))
    # Réordonner : PI, Début, Fin, puis les totaux.
    par_pi = par_pi[["Groupe_Periode", "Début", "Fin",
                     "Total absences", "Total présences"]]

    return {
        "par_personne_periode": par_personne_periode,
        "total_annuel": total_annuel,
        "par_pi": par_pi,
        "bornes_pi": bornes_pi,           # {PI: {'min': date, 'max': date}}
        "bornes_periode": bornes_periode,  # {sous-période: {'min':.., 'max':..}}
        "brut": df,
    }


if __name__ == "__main__":
    import sys
    res = charger_absences(sys.argv[1] if len(sys.argv) > 1 else "absences.xlsx")
    if res is None:
        print("Aucune donnée d'absence extraite.")
    else:
        for nom, d in res.items():
            if nom == "brut":
                continue
            print(f"\n=== {nom} ({len(d)} lignes) ===")
            print(d.head(10).to_string(index=False))
