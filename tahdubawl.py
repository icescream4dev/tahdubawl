"""Tahdubawl — Tirage au sort sur fichier Excel."""
import sys
from pathlib import Path
from config_loader import load as load_config
from excel_reader import read as read_excel
from rules_engine import apply as apply_rules
from draw import emails as extract_emails
from output_writer import write as write_result


def main():
    print("Tahdubawl v0.1.0")
    print("=" * 40)

    # ── 1. Config ──────────────────────────────────
    config_path = Path("config.yaml")
    try:
        cfg = load_config(config_path)
    except Exception as e:
        print(f"❌ Erreur config : {e}")
        return 1

    # ── 2. Excel ───────────────────────────────────
    xlsx = Path(cfg.excel_file)
    if not xlsx.exists():
        print(f"❌ Fichier introuvable : {xlsx}")
        return 1

    print(f"📂 {xlsx}...")
    try:
        col_map, rows = read_excel(xlsx, cfg.sheet)
    except Exception as e:
        print(f"❌ Erreur lecture : {e}")
        return 1
    print(f"   {len(rows)} lignes, {len(col_map)} colonnes")

    # Vérifier colonnes
    for col in [cfg.email_column] + [r.column for r in cfg.rules]:
        if col not in col_map:
            cols = ", ".join(f"{k}={v}" for k, v in sorted(col_map.items()))
            print(f"❌ Colonne '{col}' absente. Disponibles : {cols}")
            return 1

    # ── 3. Règles + tirage ─────────────────────────
    if cfg.rules:
        print(f"🔍 {len(cfg.rules)} règle(s)...")
    selected, adapted = apply_rules(rows, cfg.rules, cfg.winners)

    # ── 4. Emails ──────────────────────────────────
    result = extract_emails(selected, cfg.email_column)
    print(f"🎲 {len(result)} gagnants")

    # ── 5. Sortie ──────────────────────────────────
    out = Path("resultats/resultat.txt")
    write_result(result, cfg.batch_size, out, adapted,
                 len(rows), cfg.winners, col_map, cfg.excel_file)

    batches = ((len(result) - 1) // cfg.batch_size) + 1 if result else 0
    print(f"\n✅ {len(result)} gagnants — {batches} paquets")
    print(f"📁 {out}")

    if adapted:
        print(f"⚠️  {len(adapted)} règle(s) adaptée(s)")

    return 0


if __name__ == "__main__":
    sys.exit(main())
