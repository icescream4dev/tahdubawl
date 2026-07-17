#!/bin/bash
# Lance tous les scénarios de test avec le binaire compilé (ou python)
set -e

BIN="${TAHDUBAWL_BIN:-python3 ../tahdubawl.py}"
PASS=0
FAIL=0

for d in scenario_0*; do
    echo "=== $d ==="
    cd "$d"
    rm -rf resultats
    if timeout 30 $BIN > /dev/null 2>&1; then
        WINNERS=$(grep "Tirés :" resultats/resultat.txt 2>/dev/null | grep -oP '\d+' | head -1)
        ADAPTED=$(grep -c "Règles adaptées" resultats/resultat.txt 2>/dev/null || true)
        if [ -n "$WINNERS" ] && [ "$WINNERS" -gt 0 ]; then
            echo "  ✅ $WINNERS gagnants (adaptations: ${ADAPTED:-0})"
            PASS=$((PASS + 1))
        else
            echo "  ❌ Pas de gagnants"
            FAIL=$((FAIL + 1))
        fi
    else
        echo "  ❌ Timeout ou erreur"
        FAIL=$((FAIL + 1))
    fi
    cd ..
done

echo ""
echo "=== $PASS/$((PASS+FAIL)) scénarios OK ==="
[ "$FAIL" -eq 0 ] || exit 1
