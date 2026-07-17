# Build Windows — Tahdubawl

## Prérequis

- **Python 3.11+** installé (https://python.org)
- **Visual Studio Build Tools** ou **MinGW-w64** (compilateur C)
- Dépendances Python :

```bash
pip install -r requirements.txt
```

## Build

```batch
build.bat
```

→ Produit `build\tahdubawl.dist\` avec `tahdubawl.exe` et toutes les dépendances.

## Dossier portable

```batch
mkdir tahdubawl-portable
xcopy /E build\tahdubawl.dist\* tahdubawl-portable\
copy config.yaml tahdubawl-portable\
copy ..\LISEZ-MOI.txt tahdubawl-portable\
```

Ou utiliser le script `build.bat` qui propose de le faire automatiquement.

## Alternative : Python embeddable (fallback AV)

Si Nuitka est bloqué par l'antivirus, utiliser Python embeddable :

1. Télécharger https://www.python.org/downloads/windows/ → "Windows embeddable package (64-bit)"
2. Extraire dans `tahdubawl-portable/`
3. Installer les dépendances dans le dossier embeddable :

```batch
python-embed\python.exe -m pip install openpyxl PyYAML --target python-embed\Lib
```

4. Créer `tahdubawl-portable\lancer.bat` :

```batch
@echo off
python-embed\python.exe tahdubawl.py
pause
```

5. Copier les sources Python (`tahdubawl.py`, `config_loader.py`, etc.) dans le dossier.
