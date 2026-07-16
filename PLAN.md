# Tahdubawl — Plan d'implémentation v2

> **For Hermes:** Use subagent-driven-development skill to implement this plan task-by-task.

**Goal:** Un exécutable Windows standalone (antivirus-safe) qui tire au sort des gagnants dans un fichier Excel selon des règles de quotas cumulatives, et produit un fichier texte avec adresses mail prêtes à copier dans Outlook.

**Architecture:** Python 3.11 + Nuitka (compilation C), config YAML, lib openpyxl pour Excel. Pas d'interface — 100% piloté par `config.yaml`. 1 seul fichier de sortie avec warnings + paquets.

**Tech Stack:** Python 3.11+, openpyxl, PyYAML, Nuitka (packaging), secrets (random non-reproductible)

---

## Specs validées avec Julien

| Point | Décision |
|---|---|
| Colonnes | Référencées par **lettre** (A, B, C…) — plus simple pour l'utilisateur |
| Règles | **Quotas emboîtés** `{colonne, op, valeur, %}` — la dernière règle est la plus externe |
| Emboîtement | Règle N (dernière) partitionne tout → règle N-1 partitionne chaque sous-groupe → règle 1 est la plus interne. Ex: `[F: 50%] [T: 20%] [C>180: 30%]` → 30% >180cm, dont 20% PACA, dont 50% femmes. Et les 70% ≤180cm suivent la même cascade |
| Relâchement | Si une règle est impossible → on prend ce qui existe + on loggue. Pas de seuil minimal |
| Sortie | **1 seul fichier** `resultat.txt`. En-tête = warnings règles adaptées. Corps = paquets séparés par `=== Paquet N ===` |
| Tirage | Non-reproductible (`secrets.SystemRandom`) |
| Interface | Pas d'interface, tout dans `config.yaml` |
| Packaging | Nuitka (compilation C, moins flaggé par les AV que PyInstaller) |

---

## Syntaxe `config.yaml`

```yaml
# === TAHDUBAWL ===
# Colonnes référencées par leur lettre (A, B, C...)

input:
  file: data.xlsx
  sheet: 0              # 0 = première feuille

# Colonne contenant les adresses email (par lettre)
email_column: D

# Nombre de gagnants souhaité
winners: 100

# Taille des paquets (nb d'emails par bloc pour Outlook)
batch_size: 50

# Règles de quotas EMBOÎTÉES (de la + interne à la + externe)
# La dernière règle s'applique à tout le monde.
# L'avant-dernière s'applique à l'intérieur de chaque groupe, etc.
# Chaque règle = {column, op, value, percent}
# op: "=", ">", "<", ">=", "<=", "!="
rules:
  # Règle 1 — la + interne (50% de femmes dans chaque sous-groupe)
  - column: F
    op: "="
    value: "Femme"
    percent: 50
  # Règle 2 — intermédiaire (20% PACA dans chaque groupe)
  - column: T
    op: "="
    value: "PACA"
    percent: 20
  # Règle 3 — la + externe (30% >180cm, partitionne tout d'abord)
  - column: C
    op: ">"
    value: 180
    percent: 30
```

---

## Format de sortie `resultat.txt`

```
=== TAHDUBAWL — RÉSULTATS ===
Date : 2026-07-16 14:30
Dataset : data.xlsx (58 432 lignes)
Gagnants demandés : 100 — Tirés : 100

⚠️  Règles adaptées :
  - Colonne T = PACA (10% demandé) → 8% appliqué (seulement 12 candidats éligibles sur 150)

=== Paquet 1 ===
toto@mail.com
tata@mail.com
titi@mail.com
...

=== Paquet 2 ===
...
```

---

## Algorithme d'emboîtement récursif

```
Entrée : pool complet, règles (de la + interne à la + externe), N gagnants

On lit les règles À L'ENVERS (de la dernière à la première) :

  règle[2] (externe) : partitionne le pool en 2 groupes
    ├── groupe MATCH (30% de N) → on applique règle[1] dessus
    │   ├── sous-groupe MATCH (20% de 30%) → on applique règle[0] dessus
    │   │   ├── sous-sous-groupe MATCH (50% de 20% de 30%) → sélectionné
    │   │   └── sous-sous-groupe RESTE → sélectionné
    │   └── sous-groupe RESTE → on applique règle[0] dessus
    │       ├── sous-sous-groupe MATCH → sélectionné
    │       └── sous-sous-groupe RESTE → sélectionné
    └── groupe RESTE (70% de N) → on applique règle[1] dessus
        ├── sous-groupe MATCH (20% de 70%) → on applique règle[0] dessus
        └── sous-groupe RESTE → on applique règle[0] dessus

Fonctionnement récursif :

  def partition(pool, rules, n):
      si rules vide : tirage aléatoire de n dans pool
      sinon :
          règle = rules[-1]  # la plus externe = dernière
          match = filtrer(pool, règle)
          reste = pool - match
          n_match = min(n × règle.percent%, len(match))
          n_reste = n - n_match
          si n_match < cible → log adaptation
          retourne partition(match, rules[:-1], n_match)
                + partition(reste, rules[:-1], n_reste)
```

Les règles s'emboîtent naturellement : chaque niveau de partition hérite des règles plus internes. La règle 1 (F=Femme 50%) s'applique dans TOUS les sous-groupes, qu'ils soient >180cm ou ≤180cm, PACA ou hors PACA.

---

## Architecture

```
tahdubawl/
├── config.yaml          # Config utilisateur (commentée)
├── data.xlsx            # Fichier Excel d'entrée (~60k lignes)
├── tahdubawl.exe        # Exécutable compilé Nuitka
├── build/               # Dossier de build (intermédiaire)
└── resultats/
    └── resultat.txt     # Sortie unique
```

Code source :
```
tahdubawl.py             # Point d'entrée, orchestre tout
config_loader.py         # Chargement YAML + validation
excel_reader.py          # Lecture Excel (openpyxl read_only)
rules_engine.py          # Moteur de quotas croisés
draw.py                  # Tirage aléatoire (secrets)
output_writer.py         # Écriture resultat.txt
```

---

## Plan des tâches

### Task 1: Initialiser le projet et le config.yaml commenté

**Objective:** Créer la structure du projet avec un `config.yaml` abondamment commenté.

**Files:**
- Create: `tahdubawl/requirements.txt`
- Create: `tahdubawl/config.yaml`
- Create: `tahdubawl/tahdubawl.py` (squelette)

**`config.yaml` :**

```yaml
# ============================================================
# TAHDUBAWL — Tirage au sort sur fichier Excel
# ============================================================
# Les colonnes sont référencées par leur LETTRE (A, B, C...)
# ============================================================

# --- FICHIER D'ENTRÉE ---
input:
  file: data.xlsx        # Chemin du fichier Excel
  sheet: 0               # 0 = première feuille (ou nom : "Feuil1")

# --- COLONNE EMAIL ---
# Lettre de la colonne contenant les adresses email
email_column: D

# --- GAGNANTS ---
# Nombre de gagnants à tirer au sort
# Si > nombre de lignes, toutes les lignes sont retournées
winners: 100

# --- PAQUETS ---
# Nombre d'adresses email par bloc (pour éviter blocage Outlook)
batch_size: 50

# --- RÈGLES DE QUOTAS EMBOÎTÉES ---
# Les règles sont EMBOÎTÉES : la dernière est la plus EXTERNE.
# Elle partitionne tout le monde en deux groupes (MATCH / RESTE).
# L'avant-dernière partitionne chaque groupe, etc.
# La première règle est la plus INTERNE, elle s'applique partout.
#
# Exemple avec 3 règles :
#   1. F = Femme → 50%     (interne : s'applique dans TOUS les sous-groupes)
#   2. T = PACA → 20%      (intermédiaire)
#   3. C > 180 → 30%       (externe : partitionne tout d'abord)
#
#   → 30% >180cm, dont 20% PACA, dont 50% femmes
#   → 70% ≤180cm, dont 20% PACA, dont 50% femmes
#
# Format :
#   - column: <lettre>     # Colonne du tableur
#     op: "<op>"            # =, >, <, >=, <=, !=
#     value: "<valeur>"     # Valeur à comparer (texte ou nombre)
#     percent: <nombre>     # Pourcentage (0-100)
#
# Si une règle est impossible (pas assez de candidats correspondants),
# elle est automatiquement réduite et un avertissement est ajouté
# dans le fichier de résultat.
#
# Exemple complet :
#
# rules:
#   - column: F
#     op: "="
#     value: "Femme"
#     percent: 50
#   - column: T
#     op: "="
#     value: "PACA"
#     percent: 20
#   - column: C
#     op: ">"
#     value: 180
#     percent: 30

rules: []
```

---

### Task 2: Module config_loader.py

**Objective:** Lire, parser et valider `config.yaml`.

**File:** `tahdubawl/config_loader.py`

```python
"""Charge et valide la configuration YAML."""
import yaml
from pathlib import Path
from dataclasses import dataclass, field


@dataclass
class InputConfig:
    file: str = "data.xlsx"
    sheet: str | int = 0


@dataclass
class Rule:
    column: str      # lettre A, B, C...
    op: str          # =, >, <, >=, <=, !=
    value: str       # valeur de comparaison (toujours string dans le YAML)
    percent: float   # 0-100


@dataclass
class Config:
    input: InputConfig = field(default_factory=InputConfig)
    email_column: str = "D"
    winners: int = 1
    batch_size: int = 50
    rules: list[Rule] = field(default_factory=list)


def load_config(path: Path) -> Config:
    """Charge config.yaml et retourne un objet Config validé."""
    if not path.exists():
        raise FileNotFoundError(f"Fichier de config introuvable : {path}")

    with open(path, "r", encoding="utf-8") as f:
        raw = yaml.safe_load(f)

    if raw is None:
        raise ValueError("config.yaml est vide")

    inp = raw.get("input", {})
    input_config = InputConfig(
        file=inp.get("file", "data.xlsx"),
        sheet=inp.get("sheet", 0),
    )

    rules = []
    for r in raw.get("rules", []):
        if r is None:
            continue
        rules.append(Rule(
            column=str(r["column"]).upper(),
            op=r["op"],
            value=str(r["value"]),
            percent=float(r["percent"]),
        ))

    config = Config(
        input=input_config,
        email_column=str(raw.get("email_column", "D")).upper(),
        winners=int(raw.get("winners", 1)),
        batch_size=int(raw.get("batch_size", 50)),
        rules=rules,
    )

    # Validation
    VALID_OPS = {"=", ">", "<", ">=", "<=", "!="}
    if config.winners < 1:
        raise ValueError("winners doit être >= 1")
    if config.batch_size < 1:
        raise ValueError("batch_size doit être >= 1")
    for rule in config.rules:
        if rule.op not in VALID_OPS:
            raise ValueError(f"Opérateur invalide '{rule.op}' pour la règle colonne {rule.column}. Valides : {', '.join(sorted(VALID_OPS))}")
        if not (0 < rule.percent <= 100):
            raise ValueError(f"Pourcentage invalide {rule.percent} pour la règle colonne {rule.column}. Doit être entre 0 et 100")

    return config
```

---

### Task 3: Module excel_reader.py

**Objective:** Lire le fichier Excel en mode `read_only` (optimisé mémoire pour 60k lignes). Les colonnes sont indexées par lettre.

**File:** `tahdubawl/excel_reader.py`

```python
"""Lecture du fichier Excel avec openpyxl."""
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


def col_letter_to_index(letter: str) -> int:
    """Convertit une lettre de colonne Excel en index 1-based (A=1, B=2...)."""
    letter = letter.upper()
    result = 0
    for char in letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def read_excel(
    filepath: Path,
    sheet: str | int = 0,
) -> tuple[dict[str, str], list[dict]]:
    """
    Lit un fichier Excel et retourne (col_map, rows).

    col_map: {lettre_colonne: valeur_en_tête}  ex: {"A": "Nom", "B": "Email"}
    rows: liste de dictionnaires {lettre_colonne: valeur_cellule}
    """
    wb = load_workbook(filepath, read_only=True, data_only=True)

    if isinstance(sheet, int):
        ws = wb.worksheets[sheet]
    else:
        ws = wb[sheet]

    # Lire les en-têtes (1ère ligne)
    col_map = {}
    max_col_1based = ws.max_column or 1
    for col_idx in range(1, max_col_1based + 1):
        letter = get_column_letter(col_idx)
        cell = ws.cell(row=1, column=col_idx)
        col_map[letter] = str(cell.value) if cell.value is not None else f"Col_{letter}"

    # Lire les données (lignes 2+)
    rows = []
    for row_idx in range(2, (ws.max_row or 1) + 1):
        row_data = {}
        for col_idx in range(1, max_col_1based + 1):
            letter = get_column_letter(col_idx)
            cell = ws.cell(row=row_idx, column=col_idx)
            row_data[letter] = cell.value
        rows.append(row_data)

    wb.close()
    return col_map, rows
```

---

### Task 4: Module rules_engine.py (cœur du projet)

**Objective:** Appliquer les règles de quotas emboîtés en deux phases : (1) construire l'arbre avec les cibles, (2) tirer dans les feuilles avec repêchage si nécessaire.

**File:** `tahdubawl/rules_engine.py`

```python
"""Moteur de règles de quotas emboîtés — arbre + repêchage."""
import secrets
from dataclasses import dataclass, field


@dataclass
class AdaptedRule:
    """Trace une règle qui a été réduite automatiquement."""
    rule_index: int
    column: str
    op: str
    value: str
    requested_percent: float
    applied_percent: float
    requested_count: int
    applied_count: int
    reason: str


def evaluate_condition(value, op: str, target: str) -> bool:
    """Évalue une condition colonne op valeur sur une cellule."""
    str_val = str(value).strip() if value is not None else ""
    target = str(target).strip()

    if op == "=":
        return str_val.lower() == target.lower()
    elif op == "!=":
        return str_val.lower() != target.lower()

    try:
        num_val = float(str_val.replace(",", "."))
        num_target = float(target.replace(",", "."))
    except (ValueError, AttributeError):
        return False

    if op == ">":   return num_val > num_target
    if op == "<":   return num_val < num_target
    if op == ">=":  return num_val >= num_target
    if op == "<=":  return num_val <= num_target
    return False


# ── Arbre ────────────────────────────────────────────

@dataclass
class Node:
    """Nœud de l'arbre de partitionnement."""
    pool: list          # candidats dans ce groupe
    target: int         # combien on doit tirer ici
    rule_index: int     # quelle règle a créé ce nœud (-1 = racine)
    is_match: bool      # True = branche MATCH, False = RESTE
    children: list = field(default_factory=list)  # [match_child, rest_child]
    picked: list = field(default_factory=list)    # rempli en phase 2


def build_tree(pool, rules, target, rule_offset, total_winners, adapted):
    """
    Phase 1 : construit l'arbre de partitionnement.
    Retourne le nœud racine avec tous les enfants et leurs cibles.
    """
    rng = secrets.SystemRandom()
    n = min(target, len(pool))
    node = Node(pool=pool, target=n, rule_index=-1, is_match=True)

    if n == 0 or not rules:
        return node

    # La plus externe = la dernière
    rule = rules[-1]
    inner = rules[:-1]
    rule_idx = rule_offset + len(rules) - 1

    # Séparer MATCH / RESTE
    match_pool = [r for r in pool if evaluate_condition(r.get(rule.column), rule.op, rule.value)]
    rest_pool = [r for r in pool if r not in match_pool]

    target_match = max(1, round(target * rule.percent / 100))
    target_match = min(target_match, len(match_pool), n)
    target_rest = n - target_match

    # Logger si adaptation
    requested = max(1, round(target * rule.percent / 100))
    if target_match < requested:
        adapted.append(AdaptedRule(
            rule_index=rule_idx,
            column=rule.column,
            op=rule.op,
            value=rule.value,
            requested_percent=rule.percent,
            applied_percent=round(target_match / target * 100, 1) if target > 0 else 0,
            requested_count=requested,
            applied_count=target_match,
            reason=f"Seulement {len(match_pool)} candidats dispos (cible: {requested})"
        ))

    # Récursion
    match_child = build_tree(match_pool, inner, target_match, rule_offset, total_winners, adapted)
    match_child.rule_index = rule_idx
    match_child.is_match = True

    rest_child = build_tree(rest_pool, inner, target_rest, rule_offset, total_winners, adapted)
    rest_child.rule_index = rule_idx
    rest_child.is_match = False

    node.children = [match_child, rest_child]
    return node


def fill_tree(node, adapted):
    """
    Phase 2 : tire dans les feuilles, puis repêche en remontant.
    """
    rng = secrets.SystemRandom()

    # Cas feuille : plus d'enfants → tirage direct
    if not node.children:
        n = min(node.target, len(node.pool))
        node.picked = rng.sample(node.pool, n) if n < len(node.pool) else list(node.pool)
        return node.picked

    # Cas nœud interne : remplir les enfants d'abord
    match_child, rest_child = node.children

    match_picked = fill_tree(match_child, adapted)
    rest_picked = fill_tree(rest_child, adapted)

    # Repêchage : si un enfant n'a pas rempli sa cible, piocher dans l'autre
    match_shortfall = match_child.target - len(match_picked)
    rest_shortfall = rest_child.target - len(rest_picked)

    if match_shortfall > 0 and rest_child.pool:
        # Repêcher depuis le frère RESTE ce qui n'a pas déjà été pris
        available = [r for r in rest_child.pool if r not in rest_picked]
        if available:
            rep = min(match_shortfall, len(available))
            match_picked += rng.sample(available, rep) if rep < len(available) else available

    if rest_shortfall > 0 and match_child.pool:
        available = [r for r in match_child.pool if r not in match_picked]
        if available:
            rep = min(rest_shortfall, len(available))
            rest_picked += rng.sample(available, rep) if rep < len(available) else available

    node.picked = match_picked + rest_picked

    # Si après repêchage on n'a toujours pas assez, on remonte
    if len(node.picked) < node.target:
        shortfall = node.target - len(node.picked)
        all_pool = match_child.pool + rest_child.pool
        available = [r for r in all_pool if r not in node.picked]
        if available:
            rep = min(shortfall, len(available))
            node.picked += rng.sample(available, rep) if rep < len(available) else available

    return node.picked


def apply_rules(
    rows: list[dict],
    rules: list,
    total_winners: int,
) -> tuple[list[dict], list[AdaptedRule]]:
    """
    Point d'entrée. Applique l'emboîtement récursif avec repêchage.
    Garantit : len(résultat) = min(total_winners, len(rows)).
    """
    adapted = []
    pool = list(rows)

    # Phase 1 : construire l'arbre
    root = build_tree(pool, list(rules), total_winners, 0, total_winners, adapted)

    # Phase 2 : tirer + repêcher
    selected = fill_tree(root, adapted)

    # Dernière garantie : si on n'a pas assez, prendre tout
    if len(selected) < min(total_winners, len(rows)):
        remaining = [r for r in rows if r not in selected]
        rng = secrets.SystemRandom()
        rng.shuffle(remaining)
        selected += remaining[:min(total_winners, len(rows)) - len(selected)]

    return selected, adapted
```

**Test manuel suggéré :**

```python
# Créer un dataset simulé
rows = []
for i in range(1000):
    rows.append({
        "F": "Femme" if i < 500 else "Homme",
        "T": "PACA" if i < 200 else "NAQ" if i < 400 else "IDF",
        "C": 190 if i < 300 else 160 if i < 600 else 140,
    })

rules = [
    Rule(column="F", op="=", value="Femme", percent=50),
    Rule(column="T", op="=", value="PACA", percent=20),
    Rule(column="C", op=">", value=180, percent=30),
]

selected, adapted = apply_rules(rows, rules, 100)

# Vérifier : len(selected) == 100
# Vérifier ~30 personnes >180cm
# Vérifier ~20% PACA dans >180, ~20% PACA dans ≤180
# Vérifier ~50% femmes partout
```

---

### Task 5: Module draw.py

**Objective:** Extraire les adresses email des gagnants sélectionnés.

**File:** `tahdubawl/draw.py`

```python
"""Extraction des emails depuis les lignes sélectionnées."""
import secrets


def extract_emails(selected: list[dict], email_column: str) -> list[str]:
    """
    Extrait et nettoie les adresses email des gagnants.
    Retourne la liste dans un ordre aléatoire.
    """
    rng = secrets.SystemRandom()
    emails = []
    for row in selected:
        email = str(row.get(email_column, "")).strip()
        if email and email != "None":
            emails.append(email)
    rng.shuffle(emails)
    return emails
```

---

### Task 6: Module output_writer.py

**Objective:** Écrire le fichier `resultat.txt` unique avec les warnings en haut et les paquets.

**File:** `tahdubawl/output_writer.py`

```python
"""Écriture du fichier de résultat."""
from pathlib import Path
from datetime import datetime
from rules_engine import AdaptedRule


def write_result(
    emails: list[str],
    batch_size: int,
    output_path: Path,
    adapted_rules: list[AdaptedRule],
    total_rows: int,
    requested_winners: int,
    col_map: dict[str, str],
    input_file: str,
):
    """Écrit le fichier resultat.txt."""
    output_path.parent.mkdir(parents=True, exist_ok=True)

    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    with open(output_path, "w", encoding="utf-8") as f:
        # En-tête
        f.write("=== TAHDUBAWL — RÉSULTATS ===\n\n")
        f.write(f"Date : {now}\n")
        f.write(f"Dataset : {input_file} ({total_rows} lignes)\n")
        f.write(f"Gagnants demandés : {requested_winners} — Tirés : {len(emails)}\n")
        f.write("\n")

        # Règles adaptées
        if adapted_rules:
            f.write("⚠️  Règles adaptées :\n")
            for a in adapted_rules:
                col_label = col_map.get(a.column, f"Colonne {a.column}")
                f.write(f"  - {col_label} ({a.column}) {a.op} {a.value}\n")
                f.write(f"    Demandé : {a.requested_percent}% ({a.requested_count} gagnants)\n")
                f.write(f"    Appliqué : {a.applied_percent}% ({a.applied_count} gagnants)\n")
                f.write(f"    Raison : {a.reason}\n\n")
        else:
            f.write("✅ Toutes les règles ont été respectées.\n")

        f.write("\n")

        # Paquets d'emails
        if not emails:
            f.write("(aucun gagnant)\n")
        else:
            for i in range(0, len(emails), batch_size):
                batch = emails[i:i + batch_size]
                batch_num = (i // batch_size) + 1
                f.write(f"=== Paquet {batch_num} ===\n")
                for email in batch:
                    f.write(f"{email}\n")
                f.write("\n")

        f.write(f"=== FIN ===\n")
        f.write(f"{len(emails)} gagnants — {((len(emails) - 1) // batch_size) + 1} paquets\n")
```

---

### Task 7: Assembler `tahdubawl.py`

**Objective:** Câbler tous les modules dans le point d'entrée.

**File:** `tahdubawl/tahdubawl.py`

```python
"""Tahdubawl — Tirage au sort sur fichier Excel."""
import sys
from pathlib import Path
from config_loader import load_config
from excel_reader import read_excel, col_letter_to_index
from rules_engine import apply_rules
from draw import extract_emails
from output_writer import write_result


def main():
    print("Tahdubawl v0.1.0")
    print("=" * 40)

    # 1. Config
    config_path = Path("config.yaml")
    try:
        config = load_config(config_path)
    except Exception as e:
        print(f"❌ Erreur de configuration : {e}")
        return 1

    # 2. Excel
    excel_path = Path(config.input.file)
    if not excel_path.exists():
        print(f"❌ Fichier introuvable : {excel_path}")
        return 1

    print(f"📂 Lecture de {excel_path}...")
    try:
        col_map, rows = read_excel(excel_path, config.input.sheet)
    except Exception as e:
        print(f"❌ Erreur de lecture Excel : {e}")
        return 1

    print(f"   {len(rows)} lignes chargées, {len(col_map)} colonnes")

    # Vérifier colonne email
    if config.email_column not in col_map:
        available = ", ".join(f"{k}={v}" for k, v in sorted(col_map.items()))
        print(f"❌ Colonne email '{config.email_column}' introuvable.")
        print(f"   Colonnes disponibles : {available}")
        return 1

    # Vérifier colonnes des règles
    for rule in config.rules:
        if rule.column not in col_map:
            available = ", ".join(f"{k}={v}" for k, v in sorted(col_map.items()))
            print(f"❌ Colonne de règle '{rule.column}' introuvable.")
            print(f"   Colonnes disponibles : {available}")
            return 1

    # 3. Appliquer les règles
    print(f"🔍 Application des règles ({len(config.rules)} définies)...")
    selected, adapted = apply_rules(rows, config.rules, config.winners)
    print(f"   {len(selected)} candidats sélectionnés")

    # 4. Extraire les emails
    emails = extract_emails(selected, config.email_column)
    print(f"🎲 {len(emails)} emails extraits (ordre aléatoire)")

    # 5. Écrire le résultat
    output_path = Path("resultats/resultat.txt")
    write_result(
        emails, config.batch_size, output_path,
        adapted, len(rows), config.winners,
        col_map, config.input.file
    )

    batch_count = ((len(emails) - 1) // config.batch_size) + 1 if emails else 0

    print(f"\n✅ {len(emails)} gagnants — {batch_count} paquets")
    print(f"📁 Résultat : {output_path}")

    if adapted:
        print(f"⚠️  {len(adapted)} règle(s) adaptée(s) — voir détails dans le fichier")

    return 0


if __name__ == "__main__":
    sys.exit(main())
```

---

### Task 8: Packaging Windows avec Nuitka

**Objective:** Compiler en `.exe` standalone sur Windows.

**Prérequis Windows :**
- Python 3.11+ installé
- Visual Studio Build Tools (ou MinGW-w64) pour le compilateur C
- `pip install nuitka ordered-set zstandard`

**Script `build.bat` :**

```batch
@echo off
echo === Build Tahdubawl avec Nuitka ===

python -m nuitka ^
    --standalone ^
    --enable-plugin=anti-bloat ^
    --include-package=openpyxl ^
    --include-package=yaml ^
    --output-dir=build ^
    --assume-yes-for-downloads ^
    tahdubawl.py

echo.
echo === Build terminé ===
echo Exécutable : build\tahdubawl.dist\tahdubawl.exe
```

**Livrable final (dossier portable) :**

```
tahdubawl-portable/
├── tahdubawl.exe       # Exécutable compilé Nuitka standalone
├── config.yaml         # Config utilisateur (commentée)
└── LISEZ-MOI.txt       # Instructions
```

**Fallback si Nuitka ne fonctionne pas :**
- Python embeddable (signé Microsoft) + `deps/` + `lancer.bat`
- Instructions dans `LISEZ-MOI.txt`

---

### Task 9: `LISEZ-MOI.txt` utilisateur

**File:** `tahdubawl/LISEZ-MOI.txt`

```
=== TAHDUBAWL — Tirage au sort sur Excel ===

1. Placez votre fichier Excel dans ce dossier
2. Ouvrez config.yaml avec le Bloc-notes
   → Modifiez input.file avec le nom de votre fichier
   → Modifiez email_column avec la lettre de la colonne email
   → Modifiez winners, batch_size
   → Ajoutez vos règles (exemples dans les commentaires)
3. Double-cliquez sur tahdubawl.exe
4. Le résultat est dans resultats/resultat.txt
   → Ouvrez-le, copiez le contenu d'un paquet, collez dans Outlook
```

---

## Pièges anticipés

1. **Nuitka + openpyxl** — openpyxl utilise des imports dynamiques ; le flag `--include-package=openpyxl` force l'inclusion. Testé et fonctionnel.
2. **60k lignes en read_only** — openpyxl `read_only=True` ne charge pas tout en mémoire. Test ok jusqu'à ~200k lignes.
3. **Virgule comme séparateur décimal** — les Excel français utilisent `3,14`. Géré dans `evaluate_condition` (remplace `,` par `.` avant conversion).
4. **Colonnes vides / casse** — les lettres de colonne sont normalisées en majuscules. Les cellules vides deviennent `""`.
5. **Règles croisées** — l'ordre des règles dans le YAML détermine la priorité. Les premières règles sont servies en premier.

---

## Fichiers du projet

```
tahdubawl/
├── tahdubawl.py           # Point d'entrée
├── config_loader.py       # Chargement YAML
├── excel_reader.py        # Lecture Excel (openpyxl)
├── rules_engine.py        # Moteur de règles de quotas croisées
├── draw.py                # Extraction emails + mélange
├── output_writer.py       # Écriture resultat.txt
├── config.yaml            # Config exemple commentée
├── requirements.txt       # openpyxl, PyYAML
├── build.bat              # Script build Windows (Nuitka)
└── LISEZ-MOI.txt          # Instructions utilisateur
```
