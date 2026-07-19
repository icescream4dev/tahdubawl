"""Écriture du fichier de résultat."""
from pathlib import Path
from datetime import datetime


def write(emails, batch_size, output, adapted, total_rows, winners, col_map, excel_file):
    """Écrit resultat.txt — warnings en haut, paquets en dessous."""
    output.parent.mkdir(parents=True, exist_ok=True)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    batch_count = ((len(emails) - 1) // batch_size) + 1 if emails else 0

    with open(output, "w", encoding="utf-8") as f:
        f.write("=== TAHDUBAWL — RÉSULTATS ===\n\n")
        f.write(f"Date    : {now}\n")
        f.write(f"Dataset : {excel_file} ({total_rows} lignes)\n")
        f.write(f"Gagnants demandés : {winners} — Tirés : {len(emails)}\n\n")

        if adapted:
            f.write("⚠️  Règles adaptées :\n")
            for a in adapted:
                label = col_map.get(a.column, f"Colonne {a.column}")
                f.write(f"  - {label} ({a.column}) {a.op} {a.value}\n")
                f.write(f"    Contexte : {a.path}\n")
                f.write(f"    Demandé : {a.requested_pct}% ({a.requested_n} pers.)\n")
                f.write(f"    Appliqué : {a.applied_pct}% ({a.applied_n} pers.)\n")
                f.write(f"    Raison : {a.reason}\n\n")
        else:
            f.write("✅ Toutes les règles ont été respectées.\n\n")

        for i in range(0, len(emails), batch_size):
            batch = emails[i:i + batch_size]
            f.write(f"=== Paquet {(i // batch_size) + 1} ===\n")
            f.write("; ".join(batch) + "\n")
            f.write("\n")

        f.write(f"=== FIN ===\n")
        f.write(f"{len(emails)} gagnants — {batch_count} paquets\n")
