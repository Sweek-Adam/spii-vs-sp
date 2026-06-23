"""
SPII vs SP — Version 2 (openpyxl, rapide)
==========================================
Flux entièrement en mémoire, Excel jamais ouvert :
  1. Lecture CSV (avec conversion des décimales à virgule 0,125 -> 0.125)
  2. Construction du modèle métier en mémoire
  3. Appels Jira en parallèle (titres + Story Points)
  4. Écriture du fichier .xlsx final d'un coup (openpyxl)

Génère un NOUVEAU fichier .xlsx ; l'original n'est pas touché.
Le fichier de sortie peut rester fermé : openpyxl écrit sur disque directement.

Config :  config.toml (équipe, chemins, projet Jira) + secrets.toml (token)
Lance  :  python spii_v2.py
"""

import os
import re
import sys
import time
import unicodedata
from datetime import date
from concurrent.futures import ThreadPoolExecutor

# tomllib est natif depuis Python 3.11 ; sinon fallback sur le paquet 'tomli'
try:
    import tomllib  # Python >= 3.11
except ModuleNotFoundError:  # Python <= 3.10
    import tomli as tomllib  # pip install tomli

import pandas as pd
import requests
from requests.auth import HTTPBasicAuth

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side
from openpyxl.chart import PieChart, ScatterChart, Reference, Series
from openpyxl.chart.axis import ChartLines
from openpyxl.chart.label import DataLabelList
from openpyxl.utils import get_column_letter

# =====================================================================
# CONFIGURATION
# =====================================================================
_BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(_BASE_DIR, "config.toml")
SECRETS_PATH = os.path.join(_BASE_DIR, "secrets.toml")

# 12 colonnes de mois : "Mois de référence" = janvier 2026, puis 2/2026 -> 12/2026
MOIS_COLS = ["Mois de référence"] + [f"{m}/2026" for m in range(2, 13)]
# En-têtes affichés dans les onglets : dates 01/01/2026 -> 01/12/2026
ENTETES_MOIS = [date(2026, m, 1) for m in range(1, 13)]

# --- Styles openpyxl ---
FONT_TITRE   = Font(name="Arial", size=16, bold=True, color="003366")
FONT_BOLD    = Font(name="Arial", bold=True)
FILL_ENTETE  = PatternFill("solid", fgColor="E6F0FA")
FILL_TOTAL   = PatternFill("solid", fgColor="F0F0F0")
_thin        = Side(style="thin")
BORDER_ALL   = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)


def charger_config():
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(f"Config introuvable : {CONFIG_PATH}")
    if not os.path.exists(SECRETS_PATH):
        raise FileNotFoundError(f"Secrets introuvable : {SECRETS_PATH}")
    with open(CONFIG_PATH, "rb") as f:
        cfg = tomllib.load(f)
    with open(SECRETS_PATH, "rb") as f:
        cfg.setdefault("jira", {})["api_token"] = \
            tomllib.load(f).get("jira", {}).get("api_token", "")
    return cfg


# =====================================================================
# 1 & 2. LECTURE CSV + MODÈLE MÉTIER (en mémoire)
# =====================================================================
def conv_num(v):
    """float robuste, gère la virgule décimale française ('0,125' -> 0.125)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", "."))
    except ValueError:
        return 0.0


def _norm(s):
    """Normalise pour comparaison : minuscules, sans accents ni espaces autour.
    'Consommé', ' CONSOMME ', 'consomme' -> 'consomme'
    """
    s = str(s).strip().lower()
    return "".join(c for c in unicodedata.normalize("NFD", s)
                   if unicodedata.category(c) != "Mn")


def _labels_cat_pourcentage():
    """Étiquettes de camembert affichant la CATÉGORIE + le POURCENTAGE.
    Désactive explicitement nom de série, valeur brute et clé de légende
    (sinon openpyxl affiche 'Série1; Catégorie; Valeur; %').
    """
    dl = DataLabelList()
    dl.showCatName = True       # ex. "Nom Prénom"
    dl.showPercent = True       # ex. "100%"
    dl.showVal = False          # retire la valeur brute (0,5)
    dl.showSerName = False      # retire "Série1"
    dl.showLegendKey = False
    dl.showBubbleSize = False
    dl.separator = "; "         # séparateur entre catégorie et pourcentage
    return dl


def construire_modele(csv_path, dict_param):
    df = pd.read_csv(csv_path, sep=None, engine="python",
                     encoding="latin1", index_col=False)

    features_conso = {}   # code -> [12 mois]
    collab_data = {}      # nom -> {role, rows, total}
    stats = {}            # code -> ventilation rôle

    for _, row in df.iterrows():
        ressource = str(row["Ressource"]).strip()
        ressource_maj = ressource.upper()
        if ressource_maj not in dict_param:
            continue
        # On ne garde QUE la consommation réelle (exclut Affecte, RAF, Disponible)
        if _norm(row["Type"]) != "consomme":
            continue
        # Exclure les lignes de total agrégé du CSV (Projet == "TOTAL") : elles
        # sont la somme des lignes détaillées du collaborateur. Les inclure
        # provoquerait un double comptage de la consommation.
        if _norm(row["Projet"]) == "total":
            continue
        role = dict_param[ressource_maj].upper()

        livrable = str(row["Livrable"]) if pd.notna(row["Livrable"]) else ""
        match = re.search(r"TCRE-\d{3,}", livrable, re.IGNORECASE)
        valeurs = [conv_num(row[c]) for c in MOIS_COLS]
        total_ligne = sum(valeurs)

        est_tcre = bool(match)
        code = match.group(0).upper() if est_tcre else ""

        if est_tcre:
            features_conso.setdefault(code, [0.0] * 12)
            for i in range(12):
                features_conso[code][i] += valeurs[i]
            s = stats.setdefault(code, {"total": 0.0, "po_sm": 0.0,
                                        "ba": 0.0, "dev": 0.0, "qa": 0.0})
            s["total"] += total_ligne
            if role in ("PO", "SM"):
                s["po_sm"] += total_ligne
            elif role == "BA":
                s["ba"] += total_ligne
            elif role in ("DEV", "DEVS"):
                s["dev"] += total_ligne
            elif role == "QA":
                s["qa"] += total_ligne

        nom_onglet = ressource[:30]
        c = collab_data.setdefault(nom_onglet, {"role": dict_param[ressource_maj],
                                                "rows": [], "total": 0.0})
        feat_root = code if est_tcre else (
            "" if str(row["Projet"]).upper() == "TOTAL" else "Autre / Hors Feature")
        c["rows"].append([row["Projet"], livrable, feat_root] + valeurs + [total_ligne])
        c["total"] += total_ligne

    return {"features_conso": features_conso,
            "collab_data": collab_data, "stats": stats}


# =====================================================================
# 3. JIRA (parallèle)
# =====================================================================
def recuperer_jira(tcre_list, cfg):
    auth = HTTPBasicAuth(cfg["jira"]["email"], cfg["jira"]["api_token"])
    session = requests.Session()
    session.auth = auth
    url_base = cfg["jira"]["url"]
    sp_field = cfg["jira"]["sp_field"]
    projet = cfg["jira"]["projet"]

    def un_tcre(tcre):
        titre = "Titre introuvable"
        try:
            r = session.get(f"{url_base}/rest/api/3/issue/{tcre}",
                            params={"fields": "summary"}, timeout=30)
            if r.status_code == 200:
                titre = r.json().get("fields", {}).get("summary", "Sans titre")
        except Exception as e:
            print(f"   Erreur titre {tcre} : {e}")
        jql = f'issue in linkedIssues("{tcre}") AND key ~ "{projet}-"'
        total_sp = 0.0
        try:
            r = session.get(f"{url_base}/rest/api/3/search/jql",
                            params={"jql": jql, "fields": sp_field,
                                    "maxResults": 100}, timeout=30)
            if r.status_code == 200:
                for issue in r.json().get("issues", []):
                    total_sp += float(issue["fields"].get(sp_field) or 0)
        except Exception as e:
            print(f"   Erreur SP {tcre} : {e}")
        return tcre, {"titre": titre, "sp": total_sp}

    with ThreadPoolExecutor(max_workers=8) as ex:
        return dict(ex.map(un_tcre, tcre_list))


# =====================================================================
# 4. ÉCRITURE OPENPYXL
# =====================================================================
def _style_entete(ws, cell_range):
    for row in ws[cell_range]:
        for c in row:
            c.font = FONT_BOLD
            c.fill = FILL_ENTETE


def _bordures(ws, r1, c1, r2, c2):
    for r in range(r1, r2 + 1):
        for c in range(c1, c2 + 1):
            ws.cell(row=r, column=c).border = BORDER_ALL


def _couleur_degrade(t):
    """Interpole vert (t=0) -> jaune (t=0.5) -> rouge (t=1). Renvoie un hex 'RRGGBB'."""
    t = max(0.0, min(1.0, t))
    vert, jaune, rouge = (99, 190, 123), (255, 235, 132), (248, 105, 107)
    if t <= 0.5:
        f, a, b = t / 0.5, vert, jaune
    else:
        f, a, b = (t - 0.5) / 0.5, jaune, rouge
    rgb = tuple(int(round(a[i] + (b[i] - a[i]) * f)) for i in range(3))
    return f"{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


def _appliquer_degrade(ws, col_letter, row_min, row_max):
    """Colore le fond des cellules d'une colonne selon leur valeur (vert->rouge).
    Couleur figée (cohérent avec le choix valeurs figées de la V2).
    """
    vals = []
    for r in range(row_min, row_max + 1):
        v = ws[f"{col_letter}{r}"].value
        vals.append(v if isinstance(v, (int, float)) else None)
    nums = [v for v in vals if v is not None]
    if not nums:
        return
    vmin, vmax = min(nums), max(nums)
    etendue = (vmax - vmin) or 1.0
    for i, v in enumerate(vals):
        if v is not None:
            hexc = _couleur_degrade((v - vmin) / etendue)
            ws[f"{col_letter}{row_min + i}"].fill = PatternFill("solid", fgColor=hexc)


def ecrire_classeur(modele, jira, sortie_path):
    wb = Workbook()
    wb.remove(wb.active)  # retire la feuille vide par défaut

    features = modele["features_conso"]
    collab = modele["collab_data"]
    stats = modele["stats"]

    # Tri des TCRE par numéro
    def num(code):
        m = re.search(r"TCRE-(\d+)", code, re.IGNORECASE)
        return int(m.group(1)) if m else 0
    codes_tries = sorted(features.keys(), key=num)

    # TCRE qui auront un onglet dédié (conso > 0) : sert aux liens internes.
    codes_avec_onglet = {c for c in codes_tries if stats[c]["total"] > 0}

    # Style de lien interne (bleu souligné, comme un hyperlien classique)
    FONT_LIEN = Font(name="Arial", color="0563C1", underline="single")

    # (Onglet Paramètres supprimé : la source des rôles est config.toml)

    # --- Suivi_Features ---
    ws_feat = wb.create_sheet("Suivi_Features")
    ws_feat.append(["Code Feature"] + ENTETES_MOIS + ["Total Consommé"])
    for code in codes_tries:
        mois = features[code]
        ws_feat.append([code] + mois + [sum(mois)])
    _style_entete(ws_feat, "A1:N1")
    # Dégradé vert->rouge sur la colonne Total Consommé (N)
    if codes_tries:
        _appliquer_degrade(ws_feat, "N", 2, 1 + len(codes_tries))

    # --- Onglets collaborateurs ---
    for nom, data in collab.items():
        ws = wb.create_sheet(nom[:31])  # limite Excel : 31 car.
        ws.append(["Projet", "Livrable d'origine", "Feature (Racine)"]
                  + ENTETES_MOIS + ["Total"])
        for r in data["rows"]:
            ws.append(r)
        _style_entete(ws, "A1:P1")
        # Rôle en R1 (colonne 18)
        ws.cell(row=1, column=17, value="Rôle :").font = FONT_BOLD
        ws.cell(row=1, column=18, value=data["role"])
        # Dégradé vert->rouge sur la colonne Total (P) — UNIQUEMENT les lignes détail
        n_detail = len(data["rows"])
        if n_detail >= 1:
            _appliquer_degrade(ws, "P", 2, 1 + n_detail)

        # Ligne TOTAL globale calculée (somme des lignes détail), pour affichage
        # uniquement : elle n'alimente NI les cumuls NI le dégradé.
        if n_detail >= 1:
            row_total = 2 + n_detail  # juste après la dernière ligne détail
            # Somme colonne par colonne : 12 mois (D..O = idx 3..14) + Total (P = idx 15)
            sommes = [0.0] * 13  # 12 mois + total
            for r in data["rows"]:
                for j in range(13):           # r[3..15] = 12 mois + total
                    sommes[j] += conv_num(r[3 + j])
            ws.cell(row=row_total, column=1, value="TOTAL").font = FONT_BOLD
            for j, val in enumerate(sommes):
                c = ws.cell(row=row_total, column=4 + j, value=round(val, 3))
                c.font = FONT_BOLD
            for col in range(1, 17):          # fond gris clair sur A..P
                ws.cell(row=row_total, column=col).fill = FILL_TOTAL

    # --- Stats ---
    ws_stats = wb.create_sheet("Stats")
    ws_stats.append(["Feature", "Story points", "Total consommé",
                     "Ratio total consommé / story points",
                     "Conso PO / SM", "Conso BA", "Conso Dévs", "Conso QA",
                     "Titre"])
    _style_entete(ws_stats, "A1:I1")
    r = 2
    # Lignes de données + collecte (sp, total) par feature pour la synthèse
    sp_total_par_feature = []  # [(sp, total), ...]
    for code in codes_tries:
        s = stats[code]
        sp = jira.get(code, {}).get("sp", 0.0)
        titre = jira.get(code, {}).get("titre", "")
        ratio = (s["total"] / sp) if sp else 0.0
        ws_stats.append([code, sp, s["total"], ratio,
                         s["po_sm"], s["ba"], s["dev"], s["qa"], titre])
        # Lien interne vers l'onglet TCRE, seulement s'il existe (conso > 0)
        if code in codes_avec_onglet:
            cell = ws_stats.cell(row=r, column=1)
            cell.hyperlink = f"#'{code}'!A1"
            cell.font = FONT_LIEN
        sp_total_par_feature.append((sp, s["total"]))
        r += 1
    derniere = r - 1

    # Ligne TOTAL
    ws_stats.cell(row=r, column=1, value="TOTAL").font = FONT_BOLD
    for col in (2, 3, 5, 6, 7, 8):
        L = get_column_letter(col)
        ws_stats.cell(row=r, column=col, value=f"=SUM({L}2:{L}{derniere})").font = FONT_BOLD
    for c in range(1, 9):
        ws_stats.cell(row=r, column=c).fill = FILL_TOTAL
    total_row = r

    # Ligne MOYENNE J/SP TYPOLOGIE (moyenne globale = total conso / total SP)
    avg_row = total_row + 1
    somme_sp = sum(sp for sp, _ in sp_total_par_feature)
    somme_total = sum(t for _, t in sp_total_par_feature)
    moy_globale = (somme_total / somme_sp) if somme_sp else 0.0
    ws_stats.cell(row=avg_row, column=1, value="MOYENNE J/SP TYPOLOGIE").font = FONT_BOLD
    cell_moy = ws_stats.cell(row=avg_row, column=4, value=round(moy_globale, 3))
    cell_moy.number_format = '0.000" jour(s)"'
    for c in range(1, 9):
        ws_stats.cell(row=avg_row, column=c).fill = PatternFill("solid", fgColor="E1EBF5")

    # --- Tableau de synthèse : moyenne jours réels par complexité (SP) ---
    start_recap = derniere + 5
    ws_stats.cell(row=start_recap, column=1, value="Complexité (SP)")
    ws_stats.cell(row=start_recap, column=2, value="Moyenne Jours Réels")
    ws_stats.cell(row=start_recap, column=3, value="Nb Features")
    for c in range(1, 4):
        ws_stats.cell(row=start_recap, column=c).font = FONT_BOLD
        ws_stats.cell(row=start_recap, column=c).fill = PatternFill("solid", fgColor="D2E1F0")

    # Regroupement par SP unique (> 0), calculé en Python
    par_sp = {}  # sp -> [liste des totaux]
    for sp, total in sp_total_par_feature:
        if isinstance(sp, (int, float)) and sp > 0:
            par_sp.setdefault(sp, []).append(total)

    recap_first = start_recap + 1
    for i, sp_val in enumerate(sorted(par_sp.keys())):
        totaux = par_sp[sp_val]
        moyenne = sum(totaux) / len(totaux)
        rr = recap_first + i
        ws_stats.cell(row=rr, column=1, value=sp_val)
        ws_stats.cell(row=rr, column=2, value=round(moyenne, 2)).number_format = '0.00'
        ws_stats.cell(row=rr, column=3, value=len(totaux))
    recap_last = recap_first + len(par_sp) - 1

    # --- Graphique nuage de points : SP (X) vs Total consommé (Y), 1 point/TCRE ---
    if derniere >= 2:
        scatter = ScatterChart()
        scatter.title = "Features TCRE : Story Points vs Jours réels"
        scatter.x_axis.title = "Story Points"
        scatter.y_axis.title = "Total consommé (jours)"
        scatter.height, scatter.width = 9, 14
        # Graduations : marques majeures vers l'extérieur + quadrillage majeur,
        # et on force l'affichage des étiquettes de graduation sur les 2 axes.
        scatter.x_axis.majorTickMark = "out"
        scatter.y_axis.majorTickMark = "out"
        scatter.x_axis.minorTickMark = "out"
        scatter.y_axis.minorTickMark = "out"
        scatter.x_axis.majorGridlines = ChartLines()
        scatter.y_axis.majorGridlines = ChartLines()
        scatter.x_axis.tickLblPos = "low"
        scatter.y_axis.tickLblPos = "nextTo"
        scatter.x_axis.delete = False  # ne pas masquer l'axe (défaut openpyxl)
        scatter.y_axis.delete = False
        xref = Reference(ws_stats, min_col=2, min_row=2, max_row=derniere)  # SP
        yref = Reference(ws_stats, min_col=3, min_row=2, max_row=derniere)  # Total
        serie = Series(yref, xref, title="Features TCRE")
        serie.marker.symbol = "circle"
        serie.graphicalProperties.line.noFill = True  # points seuls, pas de ligne
        scatter.series.append(serie)
        ws_stats.add_chart(scatter, f"E{start_recap}")

    # --- Onglets par TCRE (avec camemberts) ---
    # On ne crée un onglet dédié que si le TCRE a une consommation > 0.
    # (Les lignes restent présentes dans Stats et Suivi_Features.)
    for idx, code in enumerate(codes_tries):
        s = stats[code]
        if s["total"] <= 0:
            continue
        titre = jira.get(code, {}).get("titre", "")
        ws = wb.create_sheet(code)

        # Titre
        ws.cell(row=1, column=1, value=f"{code} - {titre}").font = FONT_TITRE
        ws.row_dimensions[1].height = 30

        # Lien retour vers l'onglet Stats (ligne 2, colonne A)
        lien = ws.cell(row=2, column=1, value="← Retour vers Stats")
        lien.hyperlink = "#'Stats'!A1"
        lien.font = FONT_LIEN

        # Tableau 1 : par profil
        ws.cell(row=4, column=1, value="Profil / Rôle")
        ws.cell(row=4, column=2, value="Consommation (Jours)")
        profils = [("PO / SM", s["po_sm"]), ("BA", s["ba"]),
                   ("Développeurs", s["dev"]), ("QA", s["qa"])]
        for i, (lab, val) in enumerate(profils):
            ws.cell(row=5 + i, column=1, value=lab)
            ws.cell(row=5 + i, column=2, value=round(val, 3))
        ws.cell(row=9, column=1, value="TOTAL")
        ws.cell(row=9, column=2, value=round(sum(v for _, v in profils), 3))
        _style_entete(ws, "A4:B4")
        ws.cell(row=9, column=1).font = FONT_BOLD
        ws.cell(row=9, column=2).font = FONT_BOLD
        ws.cell(row=9, column=1).fill = FILL_TOTAL
        ws.cell(row=9, column=2).fill = FILL_TOTAL
        _bordures(ws, 4, 1, 9, 2)

        # Camembert 1 : par profil
        pie1 = PieChart()
        pie1.title = "Répartition par Profil"
        pie1.height, pie1.width = 6, 11
        data = Reference(ws, min_col=2, min_row=5, max_row=8)
        cats = Reference(ws, min_col=1, min_row=5, max_row=8)
        pie1.add_data(data, titles_from_data=False)
        pie1.set_categories(cats)
        pie1.dataLabels = _labels_cat_pourcentage()
        ws.add_chart(pie1, "E4")

        # Tableau 2 : par collaborateur
        ws.cell(row=12, column=1, value="Collaborateur")
        ws.cell(row=12, column=2, value="Profil")
        ws.cell(row=12, column=3, value="Consommation")
        _style_entete(ws, "A12:C12")
        indiv = []
        for nom, cdata in collab.items():
            t = sum(rr[-1] for rr in cdata["rows"]
                    if str(rr[2]).upper() == code)
            if t > 0:
                indiv.append((nom, cdata["role"], round(t, 3)))
        row_i = 13
        for nom, role, t in indiv:
            ws.cell(row=row_i, column=1, value=nom)
            ws.cell(row=row_i, column=2, value=role)
            ws.cell(row=row_i, column=3, value=t)
            row_i += 1
        if indiv:
            ws.cell(row=row_i, column=1, value="TOTAL INDIVIDUEL").font = FONT_BOLD
            ws.cell(row=row_i, column=3,
                    value=round(sum(t for _, _, t in indiv), 3)).font = FONT_BOLD
            for c in range(1, 4):
                ws.cell(row=row_i, column=c).fill = FILL_TOTAL
            _bordures(ws, 12, 1, row_i, 3)

            # Camembert 2 : par collaborateur
            pie2 = PieChart()
            pie2.title = "Part par Collaborateur"
            pie2.height, pie2.width = 6.5, 11
            data2 = Reference(ws, min_col=3, min_row=13, max_row=row_i - 1)
            cats2 = Reference(ws, min_col=1, min_row=13, max_row=row_i - 1)
            pie2.add_data(data2, titles_from_data=False)
            pie2.set_categories(cats2)
            pie2.dataLabels = _labels_cat_pourcentage()
            ws.add_chart(pie2, "E18")

        ws.column_dimensions["A"].width = 22
        ws.column_dimensions["B"].width = 18
        ws.column_dimensions["C"].width = 14

    # Placer l'onglet Stats en première position
    if "Stats" in wb.sheetnames:
        wb._sheets.remove(ws_stats)
        wb._sheets.insert(0, ws_stats)
        wb.active = 0  # onglet actif à l'ouverture = Stats

    wb.save(sortie_path)


# =====================================================================
# MAIN
# =====================================================================
def chrono(nom, fn):
    t0 = time.perf_counter()
    res = fn()
    print(f"   ⏱ {nom} : {time.perf_counter() - t0:.2f}s")
    return res


def main():
    cfg = charger_config()
    token = cfg["jira"]["api_token"]
    if not token or token == "REMPLACE_PAR_TON_NOUVEAU_TOKEN":
        print(f"❌ Token Jira manquant dans {SECRETS_PATH}")
        return

    dict_param = {str(n).strip().upper(): str(r).strip()
                  for n, r in cfg.get("ressources", {}).items()}
    # normpath : tolère les séparateurs / ou \ quelle que soit la plateforme
    csv_path = os.path.normpath(cfg["chemins"]["csv"])
    excel_src = os.path.normpath(cfg["chemins"]["excel"])

    if not os.path.exists(csv_path):
        print(f"❌ CSV introuvable : {csv_path}")
        print("   Vérifie le chemin 'csv' dans config.toml.")
        return

    # Fichier de sortie : à côté de l'Excel d'origine, suffixé _V2
    base = os.path.splitext(os.path.basename(excel_src))[0]
    dossier_sortie = os.path.dirname(excel_src) or _BASE_DIR
    sortie = os.path.join(dossier_sortie, f"{base}_V2.xlsx")

    print("Construction du modèle métier...")
    modele = chrono("Lecture CSV + modèle",
                    lambda: construire_modele(csv_path, dict_param))

    codes = list(modele["features_conso"].keys())
    print(f"Appels Jira ({len(codes)} TCRE, parallèle)...")
    jira = chrono("Jira", lambda: recuperer_jira(codes, cfg))

    print(f"Écriture du classeur openpyxl -> {os.path.basename(sortie)}...")
    chrono("Écriture openpyxl",
           lambda: ecrire_classeur(modele, jira, sortie))

    print(f"\n✅ Terminé : {sortie}")


if __name__ == "__main__":
    main()
