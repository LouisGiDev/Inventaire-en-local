# Inventaire Parc Informatique (Local, sans serveur)

Ce projet est 100% local : pas de serveur web.  
La page `inventaire.html` sert uniquement d'interface (filtrer / supprimer / exporter en CSV).  
Python synchronise ensuite le CSV avec le fichier Excel `.xlsx`.

## Fichiers importants
- `inventaire.html` : interface web (à ouvrir en local)
- `a-ouvrir-avant-de-scanner.py` : ajoute des lignes dans le `.xlsx` via scan console
- `inventaire_parc_informatique.xlsx` : votre base de données
- `excel_to_csv.py` : export Excel -> CSV
- `sync_csv_to_excel.py` : sync CSV -> Excel (mise à jour + suppression)
- `mdp_excel.txt` : mot de passe optionnel (dernière ligne) pour protéger la feuille
- `requirements_local.txt` : dépendances Python

Le fichier `inventaire_export.csv` est généré automatiquement quand tu exécutes `excel_to_csv.py` (tu peux le supprimer et le regénérer).

## 1) Installer les dépendances Python
Dans `Code`, lance :

```powershell
python -m pip install -r requirements_local.txt
```

## 2) Rendre l’interface utilisable (Excel -> CSV)
```powershell
python excel_to_csv.py
```

Ça génère `inventaire_export.csv` dans le même dossier.

## 3) Utiliser l’interface web (sans serveur)
Ouvre `inventaire.html` (double-clic dans Edge/Chrome).

1. Clique sur **“Charger le fichier CSV”**
2. Sélectionne `inventaire_export.csv`
3. Utilise la recherche / filtres
4. Supprime si besoin
5. Clique sur **“Exporter”** : ça télécharge un nouveau CSV (ex: `inventaire_export_<timestamp>.csv`)

## 4) Synchroniser le CSV vers l’Excel (CSV -> XLSX)
Renomme le CSV téléchargé en `inventaire_maj.csv` (ou passe directement son nom au script), puis :

```powershell
python sync_csv_to_excel.py inventaire_maj.csv
```

Le script :
- met à jour les lignes existantes par `Code-barres`
- supprime les lignes Excel absentes du CSV
- ajoute les nouveaux codes

## 5) Mettre à jour le CSV pour la prochaine fois
Après sync, relance :

```powershell
python excel_to_csv.py
```

Puis recharge le CSV dans `inventaire.html`.

## Notes
- Si le fichier Excel est ouvert pendant la sync : `openpyxl` peut échouer. Ferme Excel puis relance.
- Le code-barres est normalisé en majuscules dans tous les scripts.
- Si tu cliques sur **“Tout supprimer”**, le CSV exporté ne contient que l'entête : le script de sync videra alors l'Excel.

