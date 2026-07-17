# Tests — Tahdubawl

Pour générer les datasets et lancer tous les tests :

```bash
cd tests
python3 generate.py    # génère les data.xlsx + copie les configs
./run_all.sh           # lance tahdubawl sur chaque scénario
```

ou sur Windows :

```batch
cd tests
python generate.py
run_all.bat
```

## Scénarios

| # | Nom | Dataset | Règles | Ce qui est testé |
|---|---|---|---|---|
| 01 | Standard | 5000 lignes, 50/50 H/F, 5 régions, tailles variées | 3 règles (F=Femme 50%, T=PACA 20%, C>180 30%) | Cas nominal |
| 02 | Repêchage | 5000 lignes, seulement 2% PACA | Mêmes règles que 01 | Adaptation automatique |
| 03 | Plus de gagnants | 100 lignes, winners=500 | 1 règle | Toutes les lignes retournées |
| 04 | Sans règle | 10 000 lignes | Aucune | Tirage 100% aléatoire |
| 05 | Numérique | 5000 lignes, âges 18-80 | C>50 → 40%, C<25 → 15% | Opérateurs >, < |
| 06 | Stress | 20 000 lignes, 8 colonnes | 5 règles emboîtées | Performance + emboîtement profond |
