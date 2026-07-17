# Build Windows — Tahdubawl

## Prérequis

- **Python 3.11+** (https://python.org/downloads/)
- **Visual Studio Build Tools** ou **MinGW-w64** (compilateur C)
  - VS Build Tools : https://visualstudio.microsoft.com/downloads/#build-tools-for-visual-studio-2022
  - Cochez "Desktop development with C++"

## Build offline (recommandé)

Toutes les dépendances sont dans `wheels/`. Aucune connexion réseau nécessaire.

```batch
build.bat
```

→ Le script installe les dépendances depuis `wheels/`, puis Nuitka, puis compile.

## Build online (si réseau dispo)

```batch
pip install -r requirements.txt
python -m nuitka --standalone --include-package=openpyxl --include-package=yaml --output-dir=..\build --assume-yes-for-downloads ..\tahdubawl.py
```

## Résultat

`..\build\tahdubawl.dist\` contient `tahdubawl.exe` et toutes les dépendances compilées.

## Dossier portable livrable

```batch
mkdir ..\tahdubawl-portable
xcopy /E ..\build\tahdubawl.dist\* ..\tahdubawl-portable\
copy ..\config.yaml ..\tahdubawl-portable\
copy LISEZ-MOI.txt ..\tahdubawl-portable\
```

Ou copier manuellement :
- `tahdubawl.exe` (et tous les `.dll`/`.pyd` du dossier `dist`)
- `config.yaml` (depuis la racine du projet)
- `LISEZ-MOI.txt`

## Alternative : Python embeddable (fallback si Nuitka bloqué par l'AV)

1. Télécharger Python embeddable : https://www.python.org/downloads/windows/ → "Windows embeddable package (64-bit)"
2. Extraire dans `tahdubawl-portable/python-embed/`
3. Modifier `python-embed/python312._pth` (ou similaire) pour ajouter `Lib` et `..\` au path
4. Installer les dépendances :

```batch
python-embed\python.exe -m pip install --no-index --find-links=..\build\wheels openpyxl PyYAML
```

5. Copier les sources Python (`*.py` de la racine) dans le dossier portable
6. Créer `lancer.bat` :

```batch
@echo off
python-embed\python.exe tahdubawl.py
pause
```
