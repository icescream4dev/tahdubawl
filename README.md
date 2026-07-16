# Tahdubawl

**Tirage au sort sur fichier Excel — quotas emboîtés, offline, antivirus-safe.**

Un exécutable Windows standalone qui tire au sort des gagnants dans un fichier Excel selon des règles de quotas configurables, et produit un fichier texte avec les adresses mail prêtes à copier dans Outlook.

## Principe

1. L'utilisateur place son fichier Excel et édite `config.yaml`
2. Il double-clique sur `tahdubawl.exe`
3. Le résultat est dans `resultats/resultat.txt` — 1 seul fichier, paquets séparés, prêt à copier dans Outlook

## Règles de quotas emboîtées

Les règles sont **emboîtées** : la dernière règle est la plus **externe**, elle partitionne tout le monde. L'avant-dernière partitionne chaque sous-groupe, et ainsi de suite jusqu'à la première qui est la plus **interne** (elle s'applique dans tous les sous-groupes).

```
config.yaml :

  - column: F      ← règle 1 — la + interne (50% de femmes partout)
    op: "="
    value: "Femme"
    percent: 50

  - column: T      ← règle 2 — intermédiaire (20% PACA dans chaque groupe)
    op: "="
    value: "PACA"
    percent: 20

  - column: C      ← règle 3 — la + externe (30% >180cm, partitionne d'abord)
    op: ">"
    value: 180
    percent: 30
```

**Avec 100 gagnants, ça donne (si le dataset le permet) :**

```
100 gagnants
├── 30 >180cm          (règle 3 — 30%)
│   ├── 6 PACA          (règle 2 — 20% de 30)
│   │   ├── 3 femmes    (règle 1 — 50% de 6)
│   │   └── 3 hommes
│   └── 24 hors PACA
│       ├── 12 femmes   (règle 1 — 50% de 24)
│       └── 12 hommes
└── 70 ≤180cm
    ├── 14 PACA         (règle 2 — 20% de 70)
    │   ├── 7 femmes
    │   └── 7 hommes
    └── 56 hors PACA
        ├── 28 femmes
        └── 28 hommes
```

### Repêchage

Si une feuille de l'arbre n'a pas assez de candidats pour sa cible, le programme repêche automatiquement dans le groupe parent (le frère de la bifurcation). Le nombre de gagnants est **toujours** `min(winners, nombre_total_de_lignes)`.

### Opérateurs supportés

`=`, `>`, `<`, `>=`, `<=`, `!=`

Les opérateurs numériques gèrent la virgule comme séparateur décimal (Excel français).

## Format de sortie

```text
=== TAHDUBAWL — RÉSULTATS ===
Date : 2026-07-16 14:30
Dataset : data.xlsx (58 432 lignes)
Gagnants demandés : 100 — Tirés : 100

⚠️  Règles adaptées :
  - Colonne T = PACA (20% demandé) → 18% appliqué (seulement 90 candidats dispos)

=== Paquet 1 ===
toto@mail.com
tata@mail.com
...

=== Paquet 2 ===
...
```

## Stack technique

| Choix | Décision | Pourquoi |
|---|---|---|
| Langage | Python 3.11+ | openpyxl natif, pas de dépendance réseau |
| Packaging | Nuitka (compilation C) | Binaire natif, pas flaggé par les AV comme PyInstaller |
| Fallback | Python embeddable + .bat | Signé Microsoft, zéro false positive |
| Lecture Excel | openpyxl (read_only) | Optimisé pour gros fichiers (60k+ lignes) |
| Config | YAML | Lisible, standard |
| Aléatoire | `secrets.SystemRandom()` | Non-reproductible, entropie OS |

## Usage

```yaml
# config.yaml
input:
  file: data.xlsx
  sheet: 0

email_column: D    # Lettre de la colonne contenant les emails

winners: 100
batch_size: 50     # Nombre d'adresses par paquet

rules: []          # Ajouter les règles ici
```

## Structure du projet

```
tahdubawl/
├── tahdubawl.py           # Point d'entrée
├── config_loader.py       # Chargement YAML
├── excel_reader.py        # Lecture Excel (openpyxl)
├── rules_engine.py        # Moteur de règles — arbre + repêchage
├── draw.py                # Extraction emails
├── output_writer.py       # Écriture resultat.txt
├── config.yaml            # Config exemple commentée
├── requirements.txt       # openpyxl, PyYAML
├── build.bat              # Build Windows (Nuitka)
└── LISEZ-MOI.txt          # Instructions utilisateur
```
