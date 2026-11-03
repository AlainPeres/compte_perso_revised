# Bank Statement Processor (Compte Perso)

Ce projet permet d'automatiser le traitement et l'analyse des relevés bancaires au format CSV. Il fusionne plusieurs fichiers, génère des synthèses mensuelles et exporte les données vers Excel.

## Fonctionnalités

- **Fusion Intelligente (`read_csv.py`)** : Combine plusieurs relevés CSV en un seul fichier `fusion.csv`.
  - **Concaténation des Libellés** : Récupère l'intégralité des informations de l'opération (Type, Libellé détaillé, Catégorie, Nature) pour ne perdre aucune donnée.
  - **Gestion des Doublons** : En cas de chevauchement, les données du fichier le plus récent sont prioritaires.
- **Mois complets uniquement** : Seuls les mois entièrement écoulés sont traités — le mois en cours est exclu de la synthèse, des statistiques et des onglets mensuels. Le solde affiché est celui à la fin du dernier mois complet.
- **Synthèse Mensuelle** : Regroupe les transactions par mois complet et calcule l'évolution du solde. La première colonne est intitulée **Mois complet**.
- **Support des Accents** : Gestion native du format UTF-8 (BOM) pour une compatibilité parfaite entre les exports bancaires et Excel.
- **Robustesse** : Gestion des erreurs de permission si le fichier Excel est déjà ouvert.
- **Statistiques et Catégorisation** : Catégorise automatiquement chaque transaction par mots-clés et génère un onglet **Statistiques** sous forme de **matrice Rubrique × Mois**.
  - Chaque ligne correspond à un poste de dépense. Les colonnes sont ordonnées : **Poste de dépense**, puis **Moyenne**, puis les mois du plus récent au plus ancien.
  - La **Moyenne** est calculée sur les **6 derniers mois complets** (ou moins si l'historique est plus court).
  - La catégorisation repose sur un dictionnaire `CATEGORIES` (dans `read_csv.py`) qui associe une rubrique à une liste de mots-clés recherchés dans le libellé. Les rubriques spécifiques sont vérifiées avant les rubriques génériques :
    - **Retraite** : `AGIRC`, `CNAVTS`.
    - **Téléphone** : `SFR`, `FREE`, `BOUYGUES`.
    - **Virement vers CC** : `PERMANENT FOURNIER`.
    - **Alimentation**, **Restauration**, **Culture**, **Internet**, **Logement**, **Santé**, **Transports**, **Salaires**, **Impôts**, **Retraits** : chacune sa propre liste de mots-clés (voir `CATEGORIES`).
    - **Virements** (générique) : tout libellé contenant `VIREMENT` non capté par une règle plus spécifique, avec deux sous-cas : `SDC` → Charges copropriété, `VITAL` → Virement enfants.
    - Si aucun mot-clé ne correspond (ou libellé vide) : **Autre**.
- **Export Excel** : Génère un fichier `.xlsx` avec :
  - Un onglet **Synthèse** pour le suivi du solde par mois complet.
  - Un onglet **Statistiques** (matrice Rubrique × Mois + Moyenne).
  - Un onglet par mois complet pour le détail exhaustif des transactions.

## Installation

Ce projet utilise [uv](https://github.com/astral-sh/uv) pour la gestion des dépendances.

1. Installez `uv` si ce n'est pas déjà fait.
2. Clonez le dépôt ou copiez les fichiers.
3. Synchronisez l'environnement : `uv sync`

## Utilisation

Placez vos fichiers CSV dans le dossier `DATA_DIR` (par défaut : `D:\Documents\formalités\compte perso`).

Lancez le script de fusion et traitement :

```bash
uv run read_csv.py
```

### Format du CSV attendu

Le script est optimisé pour les exports bancaires récents :

- **Séparateur** : `;`
- **Encodage** : `utf-8-sig` (supporté nativement pour les accents).
- **Structure** : Le script extrait automatiquement la date, le montant et toutes les colonnes de description disponibles.

## Structure du Projet

- `read_csv.py` : Script principal de fusion et d'analyse.
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
