# Outil d'audit et contrôle des prix fournisseur

> **Consolidation automatique de factures PDF avec détection des écarts de prix**

## Contexte

Outil conçu pour un opérateur multi-sites gérant environ 400 établissements.
Chaque mois, le fournisseur transmet une facture PDF par site.
Ce script consolide l'ensemble des factures, rapproche chaque ligne produit
d'un catalogue de prix de référence, et produit un classeur Excel mis en forme
avec récapitulatif par site et signalement colorisé des anomalies.

**Problème résolu :** Le contrôle manuel sur des centaines de PDFs était
chronophage et source d'erreurs. Une surfacturation pouvait passer inaperçue
pendant plusieurs mois.

---

## Fonctionnalités

- **Extraction PDF** via pdfplumber avec fallback OCR (pytesseract) pour les documents scannés
- **Correspondance floue** sur les noms de sites et les désignations produits — tolère les variantes OCR et les écarts orthographiques
- **Détection des écarts de prix** — trois couleurs : conforme (vert), surfacturé (rouge), sous-facturé (orange)
- **Export Excel mis en forme** avec récapitulatif par site, panneau de synthèse et comptage des anomalies
- **Gestion automatique des avoirs** — évite les doubles comptages dans les totaux

---

## Stack technique

| Bibliothèque | Rôle |
|---|---|
| `pdfplumber` | Extraction de tableaux depuis les PDF vectoriels |
| `pytesseract` + `pdf2image` | Fallback OCR pour les PDF images |
| `pandas` | Transformation et jointure des données |
| `xlsxwriter` | Export Excel mis en forme avec couleurs conditionnelles |
| `difflib` | Correspondance floue sur les chaînes de caractères |

---

## Fonctionnement

```
PDFs (un par site)
      │
      ▼
┌──────────────────────────┐
│  1. Extraction tableaux  │  pdfplumber (vectoriel) → fallback OCR
└──────────────────────────┘
      │
      ▼
┌──────────────────────────┐
│  2. Nettoyage / filtrage │  normalisation des nombres, suppression des totaux
└──────────────────────────┘
      │
      ▼
┌──────────────────────────┐
│  3. Jointure floue        │  noms de sites → référentiel
│                           │  désignations  → catalogue de prix
└──────────────────────────┘
      │
      ▼
┌──────────────────────────┐
│  4. Détection des écarts  │  prix facturé vs prix catalogue
└──────────────────────────┘
      │
      ▼
┌──────────────────────────┐
│  5. Export Excel          │  détail ligne par ligne + synthèse par site
└──────────────────────────┘
```

---

## Configuration

Trois chemins à renseigner en haut du script :

```python
DOSSIER_FACTURES  = "/chemin/vers/factures"        # dossier contenant les PDF
FICHIER_SORTIE    = "factures_consolidees.xlsx"    # nom du classeur de sortie
FICHIER_REFERENCE = "/chemin/vers/reference.xlsx" # Excel avec 2 onglets :
                                                   #   "Sites" — nom site, code analytique, agrément, RO
                                                   #   "Prix"  — désignation, prix HT
```

---

## Structure du classeur Excel de sortie

**Tableau principal** — une ligne par ligne de facture, avec les colonnes :
`Nom Facture · Nom site · Code analytique · Agrément · RO · Type · Réf. · Désignation · Bio · Quantité · Prix HT · Prix catalogue · Écart prix · Montant € HT`

**Colorisation de la colonne Écart prix :**
- 🟢 Vert — conforme (|écart| ≤ 0,001 €)
- 🔴 Rouge — surfacturé (facturé > catalogue)
- 🟡 Orange — sous-facturé (facturé < catalogue)
- Gris — aucune correspondance catalogue trouvée

**Panneau de synthèse (colonne droite) :**
- Nombre de factures / avoirs
- Taux de TVA détectés
- Statistiques de contrôle prix (conforme / surfacturé / sous-facturé / sans correspondance)
- Totaux HT et TTC par site

---

## Limitations et pistes d'évolution

- La mise en page PDF est supposée homogène pour un fournisseur donné — les regex sont à adapter pour d'autres fournisseurs
- La qualité de l'OCR dépend de la résolution de numérisation (300 dpi recommandé)
- Pistes d'évolution : automatisation par e-mail (déclenchement sur pièce jointe PDF), connecteur Power BI, support multi-fournisseurs

---

## Auteur

Analyste de données — opérateur multi-sites  
Outils : Python · Power BI · Power Automate · Excel
