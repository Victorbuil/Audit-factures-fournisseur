# 🔍 Outil d'Audit & Contrôle des Prix Fournisseur

## 📊 Vue d'ensemble du projet

![Audit des Factures Fournisseur](./images/audit-portfolio.png)

---

## 🎯 Le Problème

Le contrôle manuel sur des centaines de PDFs fournisseurs est **chronophage** et **source d'erreurs**. Une surfacturation pouvait passer inaperçue pendant des mois, impactant directement la rentabilité opérationnelle.

---

## ✅ La Solution

Un **script Python unique** qui :
- ✨ Ingère automatiquement tous les PDFs fournisseurs
- 🔗 Rapproche chaque ligne du catalogue de prix contractuel
- 📈 Exporte un **rapport d'anomalies colorisé** et exploitable
- 🚀 100% automatisé vs processus manuel

---

## 📈 Impact Business

| Métrique | Avant | Après |
|----------|-------|-------|
| ⏱️ Temps de traitement | 20h/mois | 2h/mois |
| 🏢 Couverture | 1-2 sites | ~400 sites |
| ✅ Taux d'automatisation | 0% | 100% |
| 🎯 Granularité du contrôle | Manuelle | Ligne à ligne |

---

## 🛠️ Stack Technique

- **Python** : orchestration et traitement
- **pdfplumber** : extraction PDF
- **pytesseract / pdf2image** : OCR et reconnaissance
- **pandas** : consolidation données
- **xlsxwriter** : génération rapports
- **difflib / regex** : rapprochement floue

---

## 📋 Pipeline de Traitement

```
PDF → OCR → Jointage Floue → Détection Écarts → Rapport Excel
```

### Étapes principales :
1. **EXTRACTION PDF** - Extraction et structure des données
2. **FALLBACK OCR** - Reconnaissance optique pour PDFs complexes
3. **JOINTURE FLOUE** - Rapprochement intelligent avec catalogue
4. **DÉTECTION ÉCARTS** - Identification des anomalies
5. **RAPPORT EXCEL** - Export colorisé et pivot

---

## 🎯 Cas d'Usage

✅ **Audit & Contrôle interne**  
✅ **Détection de surfacturations**  
✅ **Vérification conformité catalogues prix**  
✅ **Analyse des écarts fournisseurs**  
✅ **Sécurisation des dépenses**  

---

## 👤 Profil & Objectif

**Rôle** : Contrôleur de Gestion / FP&A Analyst  

Développement d'un outil **fiable, rapide et scalable** pour piloter les achats et garantir l'équité des prix sur l'ensemble du réseau.

---

## 📚 Tags

`#Audit` `#ControlInterne` `#Automation` `#Python` `#FPA` `#Procurement` `#DataAnalysis`

---

**Dernier push** : 2026-06-08  
**Statut** : ✅ En production
