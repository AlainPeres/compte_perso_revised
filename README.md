# Bank Statement Processor (Compte Perso)

Ce projet permet d'automatiser le traitement et l'analyse des relevés bancaires au format CSV. Il génère des synthèses mensuelles et exporte les données vers Excel pour un suivi facilité.

## Fonctionnalités

- **Détection Automatique** : Identifie automatiquement le relevé CSV le plus récent dans le répertoire.
- **Synthèse Mensuelle** : Regroupe les transactions par mois (Année-Mois).
- **Suivi du Solde** : Calcule le solde à la fin de chaque mois en remontant à partir du solde actuel extrait du relevé.
- **Export Excel** : Génère un fichier `.xlsx` avec :
  - Un onglet **Summary** pour la vue d'ensemble.
  - Un onglet par mois pour le détail des transactions.

- **Fusion et Dédoublonage (read_csv_1.py)** : Combine plusieurs relevés CSV en un seul fichier `fusion.csv`.
  - Pour chaque date, les données du fichier le plus récent sont prioritaires.
  - Évite les doublons en cas de chevauchement entre plusieurs relevés.

## Installation

Ce projet utilise [uv](https://github.com/astral-sh/uv) pour la gestion des dépendances et de l'environnement virtuel.

1. Installez `uv` si ce n'est pas déjà fait.
2. Clonez le dépôt ou copiez les fichiers.
3. Le projet nécessite `openpyxl` pour l'export Excel.

## Utilisation

Placez vos fichiers CSV (relevés bancaires) dans le dossier :
`D:\Documents\formalités\compte perso`

Lancez l'un des deux scripts selon votre besoin :

### 1. Traitement du fichier le plus récent uniquement
```bash
uv run python read_csv.py
```

### 2. Fusion de tous les fichiers et traitement (Recommandé)
```bash
uv run python read_csv_1.py
```

### Format du CSV attendu
Le script est adapté au format des exports "Compte Perso" :
- **Pas d'en-tête** : Le fichier contient directement les données.
- **Séparateur** : `;`
- **Encodage** : `latin-1`
- **Structure des colonnes** :
  - Colonne 0 : Date (JJ/MM/AAAA)
  - Colonne 1 : Montant (ex: `-12,50`)
  - Colonne 2 : Type (ex: `Carte`, `Virement`)
  - Colonne 4 : Libellé
- **Solde** : Le solde final est extrait automatiquement de la dernière ligne du fichier le plus récent.

## Structure du Projet

- `read_csv.py` : Script pour traiter uniquement le fichier CSV le plus récent.
- `read_csv_1.py` : Script pour fusionner tous les CSV et les traiter globalement.
- `pyproject.toml` : Configuration des dépendances.
- `fusion.csv` : Fichier généré par `read_csv_1.py` regroupant toutes les données.
- `compte_perso.xlsx` : Fichier final généré dans le dossier `D:\Documents\formalités\compte perso`.

## Développement

Pour ajouter des fonctionnalités ou exécuter les tests :

```bash
# Ajouter une dépendance
uv add <package>

# Lancer les tests (si implémentés)
uv run python -m unittest discover
```
