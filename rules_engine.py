"""Moteur de règles emboîtées — arbre + repêchage."""
import secrets
from dataclasses import dataclass, field


@dataclass
class AdaptedRule:
    rule_index: int
    column: str; op: str; value: str
    requested_pct: float; applied_pct: float
    requested_n: int; applied_n: int
    reason: str


# ── Évaluation d'une condition ────────────────────

def _eval(value, op: str, target: str) -> bool:
    """colonne OP valeur → bool."""
    s = str(value or "").strip()
    t = target.strip()
    if op == "=":   return s.lower() == t.lower()
    if op == "!=":  return s.lower() != t.lower()
    try:
        nv = float(s.replace(",", "."))
        nt = float(t.replace(",", "."))
    except ValueError:
        return False
    if op == ">":   return nv > nt
    if op == "<":   return nv < nt
    if op == ">=":  return nv >= nt
    if op == "<=":  return nv <= nt
    return False


# ── Arbre ─────────────────────────────────────────

@dataclass
class Node:
    pool: list           # candidats dans ce nœud
    target: int          # combien tirer ici
    children: list = field(default_factory=list)  # [match, rest]
    picked: list = field(default_factory=list)


def _build(pool, rules, target, offset, total, adapted):
    """Phase 1 : construit l'arbre de partitionnement."""
    n = min(target, len(pool))
    node = Node(pool=pool, target=n)

    if n == 0 or not rules:
        return node

    rule = rules[-1]        # la + externe
    inner = rules[:-1]
    idx = offset + len(rules) - 1

    # Séparer MATCH / RESTE (avec set pour éviter O(n²))
    match = [r for r in pool if _eval(r.get(rule.column), rule.op, rule.value)]
    match_ids = {id(r) for r in match}
    rest  = [r for r in pool if id(r) not in match_ids]

    requested = max(1, round(target * rule.percent / 100))
    n_match = min(requested, len(match), n)
    n_rest  = n - n_match

    if n_match < requested:
        adapted.append(AdaptedRule(
            rule_index=idx, column=rule.column,
            op=rule.op, value=rule.value,
            requested_pct=rule.percent,
            applied_pct=round(n_match / target * 100, 1) if target else 0,
            requested_n=requested, applied_n=n_match,
            reason=f"Seulement {len(match)} candidats (cible: {requested})",
        ))

    node.children = [
        _build(match, inner, n_match, offset, total, adapted),
        _build(rest,  inner, n_rest,  offset, total, adapted),
    ]
    return node


def _fill(node):
    """Phase 2 : tirage dans les feuilles + repêchage."""
    rng = secrets.SystemRandom()

    if not node.children:                     # feuille
        n = min(node.target, len(node.pool))
        node.picked = rng.sample(node.pool, n) if n < len(node.pool) else list(node.pool)
        return node.picked

    left, right = node.children
    lp = _fill(left)
    rp = _fill(right)

    # Repêchage entre frères
    for (child, picked, sibling, sib_picked) in [
        (left, lp, right, rp),
        (right, rp, left, lp),
    ]:
        short = child.target - len(picked)
        if short > 0:
            sib_ids = {id(r) for r in sib_picked}
            avail = [r for r in sibling.pool if id(r) not in sib_ids]
            if avail:
                rep = min(short, len(avail))
                picked += rng.sample(avail, rep) if rep < len(avail) else avail

    node.picked = lp + rp

    # Si toujours pas assez, piocher dans tout le pool du parent
    if len(node.picked) < node.target:
        short = node.target - len(node.picked)
        picked_ids = {id(r) for r in node.picked}
        avail = [r for r in left.pool + right.pool if id(r) not in picked_ids]
        if avail:
            rep = min(short, len(avail))
            node.picked += rng.sample(avail, rep) if rep < len(avail) else avail

    return node.picked


def apply(rows, rules, winners):
    """Applique les règles emboîtées avec repêchage.

    Garantit : len(résultat) = min(winners, len(rows)).
    Retourne (sélectionnés, règles_adaptées).
    """
    adapted = []
    root = _build(list(rows), list(rules), winners, 0, winners, adapted)
    selected = _fill(root)

    # Dernière garantie
    if len(selected) < min(winners, len(rows)):
        selected_ids = {id(r) for r in selected}
        remaining = [r for r in rows if id(r) not in selected_ids]
        secrets.SystemRandom().shuffle(remaining)
        selected += remaining[:min(winners, len(rows)) - len(selected)]

    return selected, adapted
