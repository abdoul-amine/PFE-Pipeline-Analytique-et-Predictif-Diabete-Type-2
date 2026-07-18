# PFE Diabète Type 2 — Institut Pasteur du Maroc

Dashboard analytique et prédictif interactif pour l'identification des facteurs de risque environnementaux du diabète de type 2, réalisé dans le cadre d'un projet de fin d'études (PFE) à l'Institut Pasteur du Maroc, sur des données de patients à Casablanca et Jadida.

## 📊 Aperçu

L'application Streamlit propose :

- Statistiques descriptives et tests univariés (OR bruts, test exact de Fisher, chi²)
- Régression logistique multivariée (OR ajustés)
- Analyse en composantes principales (ACP)
- Heatmap des corrélations entre facteurs environnementaux
- Stratification des résultats par ville (Casablanca / Jadida)
- Modélisation prédictive (Random Forest, XGBoost, etc.) avec courbes ROC et matrices de confusion
- Interprétabilité des modèles via SHAP
- Génération d'un mini-rapport et export Excel des tableaux

## 📁 Structure du projet

```
.
├── app.py                  # Application Streamlit (point d'entrée)
├── requirements.txt         # Dépendances Python
├── data/                    # Jeux de données (CSV, Excel)
├── images/                  # Graphiques générés par le pipeline (PNG)
├── exports/                 # Rapports et exports générés (mini-rapport, etc.)
├── notebooks/               # Notebooks d'exploration / modélisation ML
└── utils/                   # Modules utilitaires (réservé à une future modularisation)
```

## 🚀 Installation

```bash
git clone <url-du-repo>
cd pfe-diabete
python -m venv .venv
source .venv/bin/activate      # Windows : .venv\Scripts\activate
pip install -r requirements.txt
```

## ▶️ Lancer l'application

```bash
streamlit run app.py
```

Deux modes de fonctionnement sont disponibles :

- **Mode rapide** : utilise directement les fichiers déjà présents dans `data/` et `images/` (`diabete_dashboard_ready.csv` ou `diabete_clean_general.csv`, `Tableaux_complets_PFE_v2.xlsx`, images pré-générées).
- **Mode auto** : à partir d'un fichier Excel brut à 4 feuilles (JM / CM / JT / CT), le pipeline complet (nettoyage, tests statistiques, ACP, modèles ML, SHAP) est exécuté à la volée.

## 🧰 Stack technique

Python · Streamlit · Pandas · NumPy · Scikit-learn · SciPy · Statsmodels · SHAP · Matplotlib · OpenPyXL

## 👤 Contexte

Projet de fin d'études (PFE) — Data Science, encadré par Pr. Mohamed Lazaar et Dr. Salsabil Hamdi, Institut Pasteur du Maroc.

## 📄 Licence

À définir selon les besoins du projet (ex. usage académique uniquement, données non redistribuables sans autorisation).
