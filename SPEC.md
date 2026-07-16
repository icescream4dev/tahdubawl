# Spec fonctionnelle — Tahdubawl

## Règles de quotas emboîtées

### Principe

Les règles sont lues de haut en bas dans `config.yaml`, mais appliquées de **bas en haut** :

- La **dernière règle** est la plus **externe** : elle partitionne tout le dataset en MATCH / RESTE
- L'**avant-dernière** s'applique **à l'intérieur** de chaque groupe créé par la dernière
- La **première règle** est la plus **interne** : elle s'applique dans **tous** les sous-groupes

### Exemple chiffré

Dataset : 1000 personnes, 500 femmes, 200 PACA, 300 >180cm
Règles :
1. `F = Femme → 50%`
2. `T = PACA → 20%`
3. `C > 180 → 30%`

Objectif : 100 gagnants.

```
Étape 1 — Règle 3 (C > 180, 30% de 100 = 30)
  100 personnes → 30 tirées du groupe >180cm, 70 du reste

Étape 2 — Règle 2 (T = PACA, 20%) appliquée dans CHAQUE groupe
  Dans le groupe >180cm (30 pers.) : 20% × 30 = 6 PACA, 24 non-PACA
  Dans le groupe ≤180cm (70 pers.) : 20% × 70 = 14 PACA, 56 non-PACA

Étape 3 — Règle 1 (F = Femme, 50%) appliquée dans CHAQUE sous-groupe
  >180cm + PACA (6)      : 3 femmes, 3 hommes
  >180cm + non-PACA (24)  : 12 femmes, 12 hommes
  ≤180cm + PACA (14)      : 7 femmes, 7 hommes
  ≤180cm + non-PACA (56)  : 28 femmes, 28 hommes
```

### Algorithme

**Phase 1 — Construction de l'arbre**

```
build_tree(pool, rules, target):
    si rules est vide → retourne feuille(pool, target)
    règle = rules[-1]  # la plus externe
    match = {r ∈ pool | r satisfait la règle}
    reste = pool \ match
    n_match = min(target × percent%, |match|)
    n_reste = target - n_match
    retourne nœud(
        match_child = build_tree(match, rules[:-1], n_match),
        rest_child  = build_tree(reste, rules[:-1], n_reste)
    )
```

**Phase 2 — Tirage + repêchage**

```
fill_tree(noeud):
    si noeud est une feuille → tirage aléatoire dans pool → retourne
    fill_tree(match_child)
    fill_tree(rest_child)
    si match_child n'a pas rempli sa cible → repêcher dans rest_child.pool non utilisé
    si rest_child n'a pas rempli sa cible → repêcher dans match_child.pool non utilisé
    si toujours pas assez → repêcher dans tout le pool du nœud parent
```

### Garanties

- `len(résultat) = min(winners, nombre_total_de_lignes)` **toujours**
- Une même personne n'est jamais tirée deux fois
- Le tirage est non-reproductible (`secrets.SystemRandom()`)
- Si une règle est impossible (ex: 20% PACA mais seulement 5 PACA dans le pool), elle est réduite et un avertissement est ajouté au fichier de sortie

## Format du fichier de config

```yaml
input:
  file: data.xlsx       # Chemin du fichier
  sheet: 0              # 0 = 1ère feuille

email_column: D          # Lettre de la colonne email

winners: 100             # Nombre de gagnants
batch_size: 50           # Emails par paquet Outlook

rules:
  - column: F            # Lettre de colonne
    op: "="              # =, >, <, >=, <=, !=
    value: "Femme"       # Valeur à comparer
    percent: 50          # Pourcentage cible
```

### Règles de validation

- `column` doit exister dans le fichier Excel
- `op` doit être parmi `=`, `>`, `<`, `>=`, `<=`, `!=`
- `percent` doit être entre 0 et 100 (exclus)
- `email_column` doit exister dans le fichier Excel
- `winners` doit être ≥ 1
- `batch_size` doit être ≥ 1

## Format de sortie

Un seul fichier `resultats/resultat.txt` :

```
=== TAHDUBAWL — RÉSULTATS ===
Date : ...
Dataset : ... (N lignes)
Gagnants demandés : X — Tirés : Y

⚠️  Règles adaptées :
  - Colonne X = Y (Z% demandé) → W% appliqué (raison)

=== Paquet 1 ===
email1@domaine.com
email2@domaine.com

=== Paquet 2 ===
...

=== FIN ===
```

## Contraintes techniques

- **Offline** : aucune dépendance réseau
- **Windows** : exécutable unique (Nuitka) ou Python embeddable
- **Antivirus-safe** : pas PyInstaller (flaggué), Nuitka ou embeddable signé Microsoft
- **60k+ lignes** : openpyxl en mode `read_only`, pas de chargement complet en mémoire
- **Virgule française** : `3,14` accepté comme séparateur décimal
- **Colonnes par lettre** : A, B, C... (Excel) pour simplifier la config utilisateur
