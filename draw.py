"""Extraction des emails et mélange aléatoire."""
import secrets


def emails(rows, column):
    """Extrait les adresses email, nettoie, mélange."""
    rng = secrets.SystemRandom()
    result = []
    for r in rows:
        e = str(r.get(column, "")).strip()
        if e and e != "None":
            result.append(e)
    rng.shuffle(result)
    return result
