# ============================================================
#  OUTIL D'AUDIT ET CONTRÔLE DES PRIX FOURNISSEUR
#  Consolidation automatique de factures PDF + détection
#  des écarts de prix par rapport au catalogue de référence
# ============================================================
#
#  CONTEXTE
#  --------
#  Conçu pour un opérateur multi-sites (~400 établissements).
#  Chaque mois, le fournisseur envoie un PDF par site.
#  Ce script :
#    1. Analyse tous les PDFs (extraction de tableaux + fallback OCR)
#    2. Rapproche chaque ligne produit avec le catalogue de prix
#    3. Signale les écarts (surfacturation / sous-facturation)
#    4. Consolide le tout dans un classeur Excel mis en forme
#       avec un récapitulatif par site et un rapport d'anomalies colorisé
#
#  STACK  : Python · pdfplumber · pytesseract · pandas · xlsxwriter
#  AUTEUR : [Votre nom] — Analyste de données
#  USAGE  : Renseigner les trois chemins ci-dessous, puis exécuter.
# ============================================================

# ---------- Dépendances ----------
# !pip install pdfplumber pandas xlsxwriter pytesseract pdf2image pillow
# !apt-get install -y poppler-utils tesseract-ocr tesseract-ocr-fra -qq

import os
import re
import math
import pdfplumber
import pandas as pd
from pdf2image import convert_from_path
import pytesseract
from difflib import get_close_matches


# ============================================================
#  CONFIGURATION  — renseigner ces trois chemins avant d'exécuter
# ============================================================
DOSSIER_FACTURES = "/chemin/vers/factures"         # dossier contenant les PDF
FICHIER_SORTIE   = "factures_consolidees.xlsx"     # nom du classeur de sortie
FICHIER_REFERENCE= "/chemin/vers/reference.xlsx"  # Excel avec 2 onglets :
                                                   #   "Sites"  : nom site | code analytique | agrément | RO
                                                   #   "Prix"   : désignation | prix HT


# ============================================================
#  1. FONCTIONS UTILITAIRES
# ============================================================

def nettoyer_nombre(val):
    """Normalise une valeur brute (cellule ou chaîne) en float Python.
    Gère le format numérique européen (virgule décimale),
    les symboles monétaires et les négatifs entre parenthèses type (12,50).
    """
    if val is None:
        return 0
    val = str(val).replace("€", "").replace(" ", "").replace(",", ".")
    if val.startswith("(") and val.endswith(")"):
        val = "-" + val[1:-1]
    try:
        return float(val)
    except ValueError:
        return 0


def nombre_sur(val):
    """Convertit en float en renvoyant 0 pour NaN / Inf / non-numérique."""
    try:
        v = float(val)
        return 0 if (math.isnan(v) or math.isinf(v)) else v
    except (TypeError, ValueError):
        return 0


def rendre_unique(colonnes):
    """Déduplique une liste de noms de colonnes en ajoutant _1, _2, …
    Nécessaire quand pdfplumber produit des en-têtes répétés (cellules fusionnées).
    """
    vus, nouvelles = {}, []
    for col in colonnes:
        if col not in vus:
            vus[col] = 0
            nouvelles.append(col)
        else:
            vus[col] += 1
            nouvelles.append(f"{col}_{vus[col]}")
    return nouvelles


def nettoyer_colonnes(df):
    """Supprime les espaces dans les noms de colonnes ; retire les colonnes sans nom."""
    df.columns = [str(c).strip() if c else f"*col*{i}" for i, c in enumerate(df.columns)]
    return df.loc[:, ~df.columns.str.startswith("*col*")]


# ============================================================
#  2. HELPERS REGEX — extraction de valeurs scalaires depuis le texte brut
# ============================================================

def extraire_ligne(pattern, texte):
    """Renvoie le premier groupe capturé d'une recherche regex, ou ''."""
    match = re.search(pattern, texte, re.IGNORECASE)
    return match.group(1).strip() if match else ""


def extraire_montant(texte, mot_cle):
    """Extrait un montant monétaire suivant un mot-clé dans le texte brut.
    Repli sur le premier token numérique après la position du mot-clé.
    """
    pattern = rf"{mot_cle}.*?(-?[\d.,()]+)"
    match = re.search(pattern, texte, re.IGNORECASE | re.DOTALL)
    val = match.group(1) if match else "0"
    if not match:
        idx = texte.lower().find(mot_cle.lower())
        if idx != -1:
            nombres = re.findall(r"(-?[\d.,()]+)", texte[idx:])
            val = nombres[0] if nombres else "0"
    val = val.replace("€", "").replace(" ", "").replace(",", ".")
    if val.startswith("(") and val.endswith(")"):
        val = "-" + val[1:-1]
    try:
        return float(val)
    except ValueError:
        return 0


def extraire_nom_site(texte):
    """Extrait le nom du site de livraison depuis l'en-tête de la facture.
    Modèle : 'Livraison : <nom du site> -'
    Adapter la regex à la mise en page PDF de votre fournisseur.
    """
    match = re.search(r"Livraison\s*:\s*(.+?)\s*-", texte, re.IGNORECASE)
    if match:
        return match.group(1).strip()
    match2 = re.search(r"Livraison\s*:\s*(.+)", texte, re.IGNORECASE)
    return match2.group(1).strip() if match2 else ""


# ============================================================
#  3. CORRESPONDANCE FLOUE — tolère les variantes OCR / orthographiques
# ============================================================

def correspondre_site(nom, noms_reference):
    """Renvoie le nom de site le plus proche dans la liste de référence.
    Essaie : exacte → inclusion partielle → difflib flou (seuil 0.60).
    """
    nom = str(nom).strip()
    if nom in noms_reference:
        return nom
    for ref in noms_reference:
        if nom in ref or ref in nom:
            return ref
    matches = get_close_matches(nom, noms_reference, n=1, cutoff=0.60)
    return matches[0] if matches else None


def correspondre_produit(designation, noms_catalogue):
    """Renvoie le produit catalogue le plus proche.
    Essaie : exacte (insensible à la casse) → inclusion → flou (seuil 0.70).
    """
    desig_maj = str(designation).strip().upper()
    for ref in noms_catalogue:
        if desig_maj == ref.upper():
            return ref
    for ref in noms_catalogue:
        if desig_maj in ref.upper() or ref.upper() in desig_maj:
            return ref
    matches = get_close_matches(desig_maj, [r.upper() for r in noms_catalogue], n=1, cutoff=0.70)
    if matches:
        idx = [r.upper() for r in noms_catalogue].index(matches[0])
        return noms_catalogue[idx]
    return None


# ============================================================
#  4. PARSING PDF
# ============================================================

# Paramètres pdfplumber — extraction vectorielle d'abord, repli sur alignement texte
PARAMETRES_TABLEAU = {
    "vertical_strategy":   "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance":      10,
    "join_tolerance":      10,
    "edge_min_length":     3,
    "min_words_vertical":  1,
    "min_words_horizontal":1,
    "intersection_tolerance": 10,
}

PARAMETRES_REPLI = {
    "vertical_strategy":   "text",
    "horizontal_strategy": "text",
    "snap_tolerance":       5,
    "join_tolerance":       5,
    "min_words_vertical":   2,
    "min_words_horizontal": 1,
}

EN_TETES_ATTENDUS  = {"Réf.", "Désignation", "Bio", "Quantité", "Prix HT", "Montant € HT"}
MOTS_CLES_PIED    = ["sous-total", "total ht", "total ttc", "total tva"]


def contient_donnees(liste_tableaux):
    """Renvoie True si au moins un tableau contient une cellule non vide."""
    return any(
        any(str(c).strip() for c in ligne if c is not None)
        for tableau in liste_tableaux for ligne in tableau
    )


def extraire_tableaux_page(page):
    """Extrait les tableaux de lignes produit depuis une page pdfplumber.
    Renvoie une liste de DataFrames (un par tableau trouvé).
    """
    resultats = []
    for parametres in [PARAMETRES_TABLEAU, PARAMETRES_REPLI]:
        tableaux_bruts = page.extract_tables(parametres)
        if not tableaux_bruts:
            continue
        for t in tableaux_bruts:
            if not t or len(t) < 2:
                continue
            # Identifier la ligne d'en-tête
            idx_entete = 0
            for i, ligne in enumerate(t):
                ligne_plate = [str(c).strip() if c else "" for c in ligne]
                if any(c in EN_TETES_ATTENDUS for c in ligne_plate):
                    idx_entete = i
                    break
            df_temp = pd.DataFrame(t[idx_entete + 1:], columns=t[idx_entete])
            df_temp = nettoyer_colonnes(df_temp)
            resultats.append(df_temp)
        if resultats:
            break   # s'arrêter après la première stratégie fructueuse
    return resultats


# ============================================================
#  5. FALLBACK OCR — pour les PDFs images ou scannés
# ============================================================

def extraire_texte_ocr(chemin, langue="fra"):
    """Rastérise toutes les pages à 300 dpi et les soumet à l'OCR.
    Changer langue= selon la langue de vos factures (ex. 'eng' pour anglais).
    """
    images = convert_from_path(chemin, dpi=300)
    return "\n".join(pytesseract.image_to_string(img, lang=langue) for img in images)


def parser_texte_ocr(texte):
    """Parseur heuristique pour la sortie OCR.
    Suppose que chaque ligne produit commence par une référence alphanumérique
    de 6-10 caractères, suivie de la désignation, du flag Bio (Oui/Non),
    du prix unitaire et de la quantité.
    Adapter les regex à la mise en page PDF de votre fournisseur.
    """
    lignes = []
    for ligne in texte.splitlines():
        ligne = ligne.strip()
        if not ligne:
            continue
        m = re.match(r"^([A-Z0-9]{6,10})\s+(.+)", ligne)
        if not m:
            continue
        ref, reste = m.group(1).strip(), m.group(2).strip()
        bio_match = re.search(r"\b(Oui|Non)\b", reste, re.IGNORECASE)
        if not bio_match:
            continue
        bio   = bio_match.group(1)
        desig = reste[:bio_match.start()].strip()
        apres = re.sub(r"[€|/\\%]", " ", reste[bio_match.end():].strip())
        apres = re.sub(r"\s+", " ", apres).strip()
        tokens = re.findall(r"\d+[.,]\d+|\d+", apres)
        if len(tokens) < 2:
            continue

        # Distinguer prix unitaire et quantité selon le format du token
        prix, qte = 0, 0
        for tok in tokens:
            if prix == 0 and ("." in tok or "," in tok):
                prix = nettoyer_nombre(tok)
            elif prix > 0 and qte == 0 and "." not in tok and "," not in tok:
                qte = nettoyer_nombre(tok)
                break
        if prix == 0 or qte == 0:
            prix, qte = nettoyer_nombre(tokens[0]), nettoyer_nombre(tokens[1])

        montant = round(prix * qte, 2)

        # Vérification croisée avec la colonne montant si disponible (3e token)
        if len(tokens) >= 3:
            montant_ocr = nettoyer_nombre(tokens[2])
            if montant_ocr > 0 and abs(montant - montant_ocr) / montant_ocr > 0.05:
                # OCR a probablement mal lu le premier chiffre du prix — on tente en le retirant
                prix_alt    = nettoyer_nombre(str(tokens[0]).replace(",", ".")[1:])
                montant_alt = round(prix_alt * qte, 2)
                if prix_alt > 0 and abs(montant_alt - montant_ocr) / montant_ocr <= 0.05:
                    prix, montant = prix_alt, montant_alt

        lignes.append({
            "Réf.":         ref,
            "Désignation":  desig,
            "Bio":          bio,
            "Prix HT":      prix,
            "Quantité":     qte,
            "Montant € HT": montant,
        })
    return pd.DataFrame(lignes) if lignes else pd.DataFrame()


# ============================================================
#  6. BOUCLE DE TRAITEMENT PRINCIPALE
# ============================================================

# ── Chargement du référentiel sites ─────────────────────────
try:
    df_ref = pd.read_excel(FICHIER_REFERENCE, sheet_name="Sites")
    df_ref.columns = [c.strip() for c in df_ref.columns]
    df_ref = df_ref.rename(columns={"Nom": "Nom site"})
    df_ref["Nom site"]  = df_ref["Nom site"].astype(str).str.strip()
    df_ref["Agrément"]  = pd.to_numeric(df_ref.get("Agrément", 0), errors="coerce").fillna(0).astype(int)
    print(f"📋 Référentiel sites chargé : {len(df_ref)} entrées")
except Exception as e:
    df_ref = pd.DataFrame(columns=["Nom site", "Code analytique", "Agrément", "RO"])
    print(f"⚠️  Référentiel sites non chargé : {e}")

# ── Chargement du catalogue de prix ─────────────────────────
try:
    df_catalogue = pd.read_excel(FICHIER_REFERENCE, sheet_name="Prix")
    df_catalogue.columns = [c.strip() for c in df_catalogue.columns]

    renommage = {}
    for col in df_catalogue.columns:
        cl = col.lower()
        if "désignation" in cl or "designation" in cl:
            renommage[col] = "Désignation"
        elif "prix" in cl and "ht" in cl:
            renommage[col] = "Prix HT catalogue"
    df_catalogue = df_catalogue.rename(columns=renommage)

    if "Désignation" not in df_catalogue.columns or "Prix HT catalogue" not in df_catalogue.columns:
        raise ValueError(f"Colonnes attendues introuvables. Colonnes trouvées : {list(df_catalogue.columns)}")

    df_catalogue = df_catalogue.dropna(subset=["Désignation", "Prix HT catalogue"])
    df_catalogue["Désignation"]       = df_catalogue["Désignation"].astype(str).str.strip()
    df_catalogue["Prix HT catalogue"] = df_catalogue["Prix HT catalogue"].apply(nettoyer_nombre)
    noms_catalogue  = df_catalogue["Désignation"].tolist()
    dict_catalogue  = dict(zip(df_catalogue["Désignation"], df_catalogue["Prix HT catalogue"]))
    print(f"💰 Catalogue prix chargé : {len(dict_catalogue)} produits")
except Exception as e:
    noms_catalogue, dict_catalogue = [], {}
    print(f"⚠️  Catalogue prix non chargé : {e}")

# ── Traitement de chaque facture PDF ────────────────────────
nb_factures = nb_avoirs = 0
total_ht    = total_ttc = 0.0
donnees     = []

COLONNES_SORTIE = [
    "Nom Facture", "Nom site", "Code analytique", "Agrément", "RO",
    "Type", "Réf.", "Désignation", "Bio", "Quantité",
    "Prix HT", "Prix catalogue", "Écart prix", "Montant € HT",
]

for fichier in sorted(os.listdir(DOSSIER_FACTURES)):
    if not fichier.lower().endswith(".pdf"):
        continue

    chemin = os.path.join(DOSSIER_FACTURES, fichier)
    print(f"\n📄 Traitement : {fichier}")

    texte_brut, tableaux = "", []

    with pdfplumber.open(chemin) as pdf:
        for page in pdf.pages:
            texte_brut += (page.extract_text() or "") + "\n"
            tableaux.extend(extraire_tableaux_page(page))

    # Fallback OCR si aucun tableau utilisable
    if not tableaux or not contient_donnees(tableaux):
        print("  🔎 Aucun tableau exploitable — tentative OCR…")
        texte_ocr = extraire_texte_ocr(chemin)
        df_ocr    = parser_texte_ocr(texte_ocr)
        if not df_ocr.empty:
            texte_brut, tableaux = texte_ocr, [df_ocr]
            print(f"  ✅ OCR : {len(df_ocr)} lignes extraites")
        else:
            print(f"  ⚠️  OCR : aucune ligne produit trouvée dans {fichier}")
            continue

    if not tableaux:
        print(f"  ⚠️  Aucun tableau trouvé dans {fichier}")
        continue

    # Fusionner les tableaux de toutes les pages
    df = pd.concat(tableaux, ignore_index=True)
    df.columns = rendre_unique(df.columns)
    df = df.dropna(how="all")
    df = df[df.apply(lambda row: any(str(v).strip() for v in row), axis=1)]

    for col in ["Quantité", "Prix HT", "Montant € HT"]:
        if col in df.columns:
            df[col] = df[col].apply(nettoyer_nombre)

    # Supprimer les lignes de pied de tableau / totaux
    df = df[~df.apply(
        lambda row: any(mc in " ".join(str(v).lower() for v in row) for mc in MOTS_CLES_PIED),
        axis=1
    )].reset_index(drop=True)

    # ── Jointure catalogue de prix ──
    if dict_catalogue and "Désignation" in df.columns:
        def obtenir_prix_catalogue(desig):
            corresp = correspondre_produit(desig, noms_catalogue)
            return dict_catalogue.get(corresp) if corresp else None

        df["Prix catalogue"] = df["Désignation"].apply(obtenir_prix_catalogue)
        df["Écart prix"] = df.apply(
            lambda r: round(nombre_sur(r["Prix HT"]) - nombre_sur(r["Prix catalogue"]), 4)
            if pd.notna(r.get("Prix catalogue")) else None,
            axis=1
        )
    else:
        df["Prix catalogue"] = None
        df["Écart prix"]     = None

    # Métadonnées au niveau document
    ht_doc    = df["Montant € HT"].apply(nombre_sur).sum() if "Montant € HT" in df.columns else extraire_montant(texte_brut, "Total HT")
    ttc_doc   = extraire_montant(texte_brut, "Total TTC")
    type_doc  = "Avoir" if re.search(r"\bAvoir\b", texte_brut[:500], re.IGNORECASE) else "Facture"
    nom_site  = extraire_nom_site(texte_brut)

    if type_doc == "Avoir":
        nb_avoirs += 1
    else:
        nb_factures += 1
    total_ht  += ht_doc
    total_ttc += ttc_doc

    taux_tva = sorted({
        t.replace(",", ".")
        for t in re.findall(r"(\d+[.,]\d+|\d+)\s*%", texte_brut)
        if t.replace(",", ".") in {"5.5", "10", "20"}
    })

    df["Nom Facture"] = fichier
    df["Nom site"]    = nom_site
    df["Type"]        = type_doc
    df["Taux TVA"]    = ", ".join(f"{t}%" for t in taux_tva) if taux_tva else "5.5%"
    df["Total TTC"]   = ttc_doc

    print(f"  ✅ {len(df)} lignes — {type_doc} — {nom_site} — HT : {ht_doc:.2f} € — TTC : {ttc_doc:.2f} €")
    donnees.append(df)


# ============================================================
#  7. EXPORT EXCEL
# ============================================================

if not donnees:
    print("\n⚠️  Aucune facture traitée.")
else:
    df_tout = pd.concat(donnees, ignore_index=True)

    # ── Rapprochement flou des noms de sites ──
    noms_ref = df_ref["Nom site"].tolist()
    df_tout["_nom_ref"] = df_tout["Nom site"].apply(
        lambda n: correspondre_site(str(n).strip(), noms_ref)
    )
    sans_corresp = df_tout[df_tout["_nom_ref"].isna()]["Nom site"].unique()
    if len(sans_corresp):
        print(f"\n  ⚠️  Sites sans correspondance référentiel : {list(sans_corresp)}")

    # Intégration des colonnes de référence
    df_tout = df_tout.drop(columns=["Code analytique", "Agrément", "RO"], errors="ignore")
    df_tout = df_tout.merge(
        df_ref[["Nom site", "Code analytique", "Agrément", "RO"]].rename(columns={"Nom site": "_nom_ref"}),
        on="_nom_ref", how="left"
    ).drop(columns=["_nom_ref"], errors="ignore")

    # Rapport produits sans correspondance catalogue
    if dict_catalogue and "Désignation" in df_tout.columns:
        sans_cat = df_tout[df_tout["Prix catalogue"].isna()]["Désignation"].dropna().unique()
        if len(sans_cat):
            print(f"\n  ⚠️  Produits sans correspondance catalogue ({len(sans_cat)}) :")
            for d in sans_cat:
                print(f"     - {d}")

    cols_existantes = [c for c in COLONNES_SORTIE if c in df_tout.columns]
    autres_cols     = [c for c in df_tout.columns if c not in COLONNES_SORTIE]
    df_tout         = df_tout[cols_existantes + autres_cols]

    with pd.ExcelWriter(
        FICHIER_SORTIE,
        engine="xlsxwriter",
        engine_kwargs={"options": {"nan_inf_to_errors": True}}
    ) as writer:
        wb = writer.book
        ws = wb.add_worksheet("Factures")
        writer.sheets["Factures"] = ws

        # ── Définition des formats ────────────────────────────
        fmt_titre      = wb.add_format({"bold": True, "font_size": 13, "font_name": "Arial"})
        fmt_entete     = wb.add_format({"bold": True, "bg_color": "#D9E1F2", "border": 1, "font_name": "Arial", "text_wrap": True, "valign": "vcenter"})
        fmt_cellule    = wb.add_format({"border": 1, "font_name": "Arial"})
        fmt_nombre     = wb.add_format({"border": 1, "num_format": "#,##0.00", "font_name": "Arial"})
        fmt_euro       = wb.add_format({"border": 1, "num_format": '#,##0.00 "€"', "font_name": "Arial"})
        fmt_syn_titre  = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#1F4E79", "font_color": "#FFFFFF", "border": 1, "font_name": "Arial", "align": "center"})
        fmt_syn_label  = wb.add_format({"bold": True, "bg_color": "#D6E4F0", "border": 1, "font_name": "Arial"})
        fmt_syn_val    = wb.add_format({"border": 1, "font_name": "Arial", "align": "center"})

        # Formats conditionnels pour les écarts de prix
        fmt_ecart_ok   = wb.add_format({"border": 1, "num_format": '#,##0.0000 "€"', "font_name": "Arial", "bg_color": "#E2EFDA"})
        fmt_ecart_sup  = wb.add_format({"border": 1, "num_format": '#,##0.0000 "€"', "font_name": "Arial", "bg_color": "#FCE4D6", "font_color": "#C00000"})
        fmt_ecart_inf  = wb.add_format({"border": 1, "num_format": '#,##0.0000 "€"', "font_name": "Arial", "bg_color": "#FFF2CC", "font_color": "#7F6000"})
        fmt_ecart_na   = wb.add_format({"border": 1, "font_name": "Arial", "font_color": "#AAAAAA", "italic": True, "align": "center"})
        fmt_catalogue  = wb.add_format({"border": 1, "num_format": '#,##0.0000 "€"', "font_name": "Arial", "bg_color": "#EDF2FB"})
        fmt_cat_na     = wb.add_format({"border": 1, "font_name": "Arial", "font_color": "#AAAAAA", "italic": True, "align": "center"})

        COL_SYN = len(cols_existantes) + 2   # panneau de synthèse après les colonnes de données

        # ── Panneau de synthèse ───────────────────────────────
        ws.merge_range(0, COL_SYN, 0, COL_SYN + 1, "Synthèse", fmt_syn_titre)
        ws.write(1, COL_SYN, "Nombre de factures",    fmt_syn_label); ws.write(1, COL_SYN + 1, nb_factures, fmt_syn_val)
        ws.write(2, COL_SYN, "Nombre d'avoirs",       fmt_syn_label); ws.write(2, COL_SYN + 1, nb_avoirs,   fmt_syn_val)

        taux_set = set()
        for t in df_tout.get("Taux TVA", pd.Series(dtype=str)).dropna().unique():
            for tx in str(t).split(","):
                taux_set.add(tx.strip())
        ws.write(4, COL_SYN, "Taux TVA appliqués", fmt_syn_label)
        ws.write(4, COL_SYN + 1, " / ".join(sorted(taux_set)), fmt_syn_val)

        # Statistiques de contrôle des prix
        if "Écart prix" in df_tout.columns:
            ecarts = df_tout["Écart prix"].dropna()
            ws.write(6,  COL_SYN, "── Contrôle prix catalogue ──",            fmt_syn_label)
            ws.write(7,  COL_SYN, "Lignes conformes (écart ≤ 0,001 €)",       fmt_syn_label); ws.write(7,  COL_SYN + 1, int((ecarts.abs() <= 0.001).sum()),  fmt_syn_val)
            ws.write(8,  COL_SYN, "Lignes facturées > catalogue",              fmt_syn_label); ws.write(8,  COL_SYN + 1, int((ecarts > 0.001).sum()),         fmt_syn_val)
            ws.write(9,  COL_SYN, "Lignes facturées < catalogue",              fmt_syn_label); ws.write(9,  COL_SYN + 1, int((ecarts < -0.001).sum()),        fmt_syn_val)
            ws.write(10, COL_SYN, "Lignes sans correspondance catalogue",       fmt_syn_label); ws.write(10, COL_SYN + 1, int(df_tout["Prix catalogue"].isna().sum()), fmt_syn_val)

        # Récapitulatif par site
        LIG_SYN = 12
        fmt_st  = wb.add_format({"bold": True, "font_size": 11, "bg_color": "#1F4E79", "font_color": "#FFFFFF", "border": 1, "font_name": "Arial", "align": "center"})
        fmt_sh  = wb.add_format({"bold": True, "bg_color": "#D6E4F0", "border": 1, "font_name": "Arial"})
        fmt_sc  = wb.add_format({"border": 1, "font_name": "Arial"})
        fmt_sa  = wb.add_format({"border": 1, "font_name": "Arial", "font_color": "#C00000", "italic": True})
        fmt_se  = wb.add_format({"border": 1, "num_format": '#,##0.00 "€"', "font_name": "Arial"})
        fmt_sae = wb.add_format({"border": 1, "num_format": '#,##0.00 "€"', "font_name": "Arial", "font_color": "#C00000", "italic": True})
        fmt_tot = wb.add_format({"bold": True, "bg_color": "#EBF1DE", "border": 1, "num_format": '#,##0.00 "€"', "font_name": "Arial"})
        fmt_tl  = wb.add_format({"bold": True, "bg_color": "#EBF1DE", "border": 1, "font_name": "Arial"})

        ws.merge_range(LIG_SYN, COL_SYN, LIG_SYN, COL_SYN + 2, "Par site", fmt_st)
        ws.write(LIG_SYN + 1, COL_SYN,     "Nom site",   fmt_sh)
        ws.write(LIG_SYN + 1, COL_SYN + 1, "Total HT",   fmt_sh)
        ws.write(LIG_SYN + 1, COL_SYN + 2, "Total TTC",  fmt_sh)

        r       = LIG_SYN + 2
        ht_fac  = ht_av  = 0.0
        ttc_fac = ttc_av = 0.0

        for site in sorted(df_tout["Nom site"].dropna().unique()):
            for type_doc in ["Facture", "Avoir"]:
                sous = df_tout[(df_tout["Nom site"] == site) & (df_tout["Type"] == type_doc)]
                if sous.empty:
                    continue
                ht  = sous["Montant € HT"].apply(nombre_sur).sum()
                ttc = sous.drop_duplicates("Nom Facture")["Total TTC"].apply(nombre_sur).sum()
                label  = f"{site}  [{type_doc}]"
                avoir  = type_doc == "Avoir"
                ws.write(r, COL_SYN,     label,              fmt_sa if avoir else fmt_sc)
                ws.write_number(r, COL_SYN + 1, nombre_sur(ht),  fmt_sae if avoir else fmt_se)
                ws.write_number(r, COL_SYN + 2, nombre_sur(ttc), fmt_sae if avoir else fmt_se)
                if not avoir:
                    ht_fac += ht; ttc_fac += ttc
                else:
                    ht_av  += ht; ttc_av  += ttc
                r += 1

        ws.write(r, COL_SYN, "Total Factures",                               fmt_tl)
        ws.write_number(r, COL_SYN + 1, nombre_sur(ht_fac),         fmt_tot)
        ws.write_number(r, COL_SYN + 2, nombre_sur(ttc_fac),        fmt_tot); r += 1
        ws.write(r, COL_SYN, "Net (Factures - Avoirs)",                       fmt_tl)
        ws.write_number(r, COL_SYN + 1, nombre_sur(ht_fac + ht_av),   fmt_tot)
        ws.write_number(r, COL_SYN + 2, nombre_sur(ttc_fac + ttc_av), fmt_tot)

        ws.set_column(COL_SYN,     COL_SYN,     24)
        ws.set_column(COL_SYN + 1, COL_SYN + 2, 16)

        # ── Tableau de données principal ──────────────────────
        ws.write(0, 0, "Consolidation des factures", fmt_titre)
        LIG_ENTETE = 2
        ws.set_row(LIG_ENTETE, 30)
        for c, nom_col in enumerate(cols_existantes):
            ws.write(LIG_ENTETE, c, nom_col, fmt_entete)

        COLS_EURO     = {"Montant € HT", "Prix HT"}
        COLS_NOMBRE   = {"Quantité"}
        COLS_CAT      = {"Prix catalogue"}
        COLS_ECART    = {"Écart prix"}

        for idx_ligne, (_, ligne) in enumerate(df_tout.iterrows()):
            el = LIG_ENTETE + 1 + idx_ligne
            for idx_col, nom_col in enumerate(cols_existantes):
                val = ligne.get(nom_col, "")
                if nom_col in COLS_EURO:
                    ws.write_number(el, idx_col, nombre_sur(val), fmt_euro)
                elif nom_col in COLS_NOMBRE:
                    ws.write_number(el, idx_col, nombre_sur(val), fmt_nombre)
                elif nom_col in COLS_CAT:
                    if pd.isna(val) or val is None:
                        ws.write(el, idx_col, "—", fmt_cat_na)
                    else:
                        ws.write_number(el, idx_col, nombre_sur(val), fmt_catalogue)
                elif nom_col in COLS_ECART:
                    if pd.isna(val) or val is None:
                        ws.write(el, idx_col, "—", fmt_ecart_na)
                    else:
                        v = nombre_sur(val)
                        fmt = fmt_ecart_ok if abs(v) <= 0.001 else (fmt_ecart_sup if v > 0 else fmt_ecart_inf)
                        ws.write_number(el, idx_col, v, fmt)
                else:
                    ws.write(el, idx_col, "" if pd.isna(val) else str(val), fmt_cellule)

        largeurs_col = {
            "Nom Facture": 30, "Nom site": 28, "Code analytique": 16,
            "Agrément": 10, "RO": 20, "Type": 10,
            "Réf.": 14, "Désignation": 30, "Bio": 6, "Quantité": 10,
            "Prix HT": 12, "Prix catalogue": 16, "Écart prix": 14, "Montant € HT": 14,
        }
        for idx_col, nom_col in enumerate(cols_existantes):
            ws.set_column(idx_col, idx_col, largeurs_col.get(nom_col, 15))

    print(f"\n✅ Excel généré : {FICHIER_SORTIE}")
    print(f"   → {len(df_tout)} lignes | {nb_factures} factures | {nb_avoirs} avoirs")
    print(f"   → Total HT : {total_ht:.2f} € | Total TTC : {total_ttc:.2f} €")
    if "Écart prix" in df_tout.columns:
        anomalies = (df_tout["Écart prix"].dropna().abs() > 0.001).sum()
        print(f"   → Anomalies prix détectées : {anomalies} lignes")
