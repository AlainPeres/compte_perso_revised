# Bank Statement Processor (Compte Perso)

Ce projet permet d'automatiser le traitement et l'analyse des relevés bancaires au format CSV. Il fusionne plusieurs fichiers, génère des synthèses mensuelles et exporte les données vers Excel.

## Fonctionnalités

- **Fusion Intelligente (`read_csv_1.py`)** : Combine plusieurs relevés CSV en un seul fichier `fusion.csv`.
  - **Concaténation des Libellés** : Récupère l'intégralité des informations de l'opération (Type, Libellé détaillé, Catégorie, Nature) pour ne perdre aucune donnée.
  - **Gestion des Doublons** : En cas de chevauchement, les données du fichier le plus récent sont prioritaires.
- **Synthèse Mensuelle** : Regroupe les transactions par mois et calcule l'évolution du solde.
- **Support des Accents** : Gestion native du format UTF-8 (BOM) pour une compatibilité parfaite entre les exports bancaires et Excel.
- **Robustesse** : Gestion des erreurs de permission si le fichier Excel est déjà ouvert.
- **Statistiques et Catégorisation** : Extrait automatiquement les catégories et génère un onglet **Statistiques** sous forme de **matrice Postes × Mois**.
  - Chaque ligne correspond à un poste de dépense, chaque colonne à un mois.
  - Une colonne **Moyenne mensuelle** est calculée uniquement sur les **mois complets** (mois ayant un successeur dans l'export) pour éviter les biais de début/fin d'export.
- **Export Excel** : Génère un fichier `.xlsx` avec :
  - Un onglet **Synthèse** pour le suivi du solde mensuel.
  - Un onglet **Statistiques** (matrice Postes × Mois + Moyenne) calqué sur le format de `compte_commun.xlsx`.
  - Un onglet par mois pour le détail exhaustif des transactions.

## Installation

Ce projet utilise [uv](https://github.com/astral-sh/uv) pour la gestion des dépendances.

1. Installez `uv` si ce n'est pas déjà fait.
2. Clonez le dépôt ou copiez les fichiers.
3. Synchronisez l'environnement : `uv sync`

## Utilisation

Placez vos fichiers CSV dans le dossier `DATA_DIR` (par défaut : `D:\Documents\formalités\compte perso`).

Lancez le script de fusion et traitement :

```bash
uv run read_csv_1.py
```

### Format du CSV attendu
Le script est optimisé pour les exports bancaires récents :
- **Séparateur** : `;`
- **Encodage** : `utf-8-sig` (supporté nativement pour les accents).
- **Structure** : Le script extrait automatiquement la date, le montant et toutes les colonnes de description disponibles.

## Structure du Projet

- `read_csv_1.py` : Script principal de fusion et d'analyse.
- `pyproject.toml` : Configuration des dépendances (`openpyxl`).
- `fusion.csv` : Fichier intermédiaire regroupant toutes les données dédoublonnées.
- `compte_perso.xlsx` : Fichier Excel final généré.

## Développement

```bash
# Ajouter une dépendance
uv add <package>

# Lancer les tests
uv run python -m unittest discover
```
