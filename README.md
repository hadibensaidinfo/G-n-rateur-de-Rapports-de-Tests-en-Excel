# Générateur de Résultats de Tests en Excel

Ce projet est un script Python qui analyse un fichier xUnit XML contenant les résultats des cas de test, puis génère un fichier Excel avec les données formatées, comprenant une mise en forme colorée et des statistiques.

## Fonctionnalités

- Lecture et analyse des fichiers xUnit XML.
- Génération d'un fichier Excel contenant :
  - Les chemins des cas de test.
  - Les noms des cas de test.
  - Le statut de chaque cas (`Passed` ou `Failed`).
  - Les messages d'erreur associés (le cas échéant).
- Mise en forme colorée des lignes :
  - **Vert** : Cas de test réussis.
  - **Rouge** : Cas de test échoués.
  - **Bleu** : Titres des colonnes.
- Ajout d'une feuille contenant des statistiques sur les cas de test.

## Structure du projet

- **`generate_test_results_excel.py`** : Script principal.
- **`xUnit.xml`** : Exemple de fichier xUnit XML (à fournir pour l'exécution).

## Prérequis

- Python 3.7 ou plus récent.
- Bibliothèques Python :
  - `openpyxl`
  - (Facultatif) Installer les dépendances via `pip` :
    ```bash
    pip install openpyxl
    ```

## Comment exécuter le script ?

1. Assurez-vous que le fichier xUnit XML est prêt et disponible (exemple : `xUnit.xml`).
2. Lancez le script avec Python :
   ```bash
   python generate_test_results_excel.py
