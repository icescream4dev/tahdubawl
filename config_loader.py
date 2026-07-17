"""Chargement et validation de la configuration YAML."""
import yaml
from pathlib import Path
from dataclasses import dataclass, field

# ── Opérateurs valides ─────────────────────────────
VALID_OPS = frozenset({"=", ">", "<", ">=", "<=", "!="})


@dataclass
class Rule:
    column: str       # lettre A, B, C...
    op: str           # =, >, <, >=, <=, !=
    value: str        # valeur de comparaison
    percent: float    # 0-100


@dataclass
class Config:
    excel_file: str = "data.xlsx"
    sheet: str | int = 0
    email_column: str = "D"
    winners: int = 1
    batch_size: int = 50
    rules: list[Rule] = field(default_factory=list)


def load(path: Path) -> Config:
    """Charge config.yaml et retourne un objet Config validé."""
    if not path.exists():
        raise FileNotFoundError(f"Config introuvable : {path}")

    with open(path, encoding="utf-8") as f:
        raw = yaml.safe_load(f) or {}

    inp = raw.get("input", {})
    cfg = Config(
        excel_file=inp.get("file", "data.xlsx"),
        sheet=inp.get("sheet", 0),
        email_column=str(raw.get("email_column", "D")).upper(),
        winners=int(raw.get("winners", 1)),
        batch_size=int(raw.get("batch_size", 50)),
        rules=[Rule(
            column=str(r["column"]).upper(),
            op=r["op"],
            value=str(r["value"]),
            percent=float(r["percent"]),
        ) for r in raw.get("rules") or []],
    )

    # Validation
    if cfg.winners < 1:
        raise ValueError("winners doit être >= 1")
    if cfg.batch_size < 1:
        raise ValueError("batch_size doit être >= 1")
    for r in cfg.rules:
        if r.op not in VALID_OPS:
            raise ValueError(
                f"Opérateur invalide '{r.op}' (colonne {r.column}). "
                f"Valides : {', '.join(sorted(VALID_OPS))}"
            )
        if not 0 < r.percent <= 100:
            raise ValueError(
                f"Pourcentage invalide {r.percent} (colonne {r.column}). "
                f"Doit être entre 0 et 100"
            )

    return cfg
