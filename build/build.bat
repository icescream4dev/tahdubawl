@echo off
echo === Tahdubawl — Build Windows (offline) ===
echo.

REM Étape 1 : installer les dépendances depuis les wheels locales (pas de réseau)
echo [1/3] Installation des dépendances (offline)...
pip install --no-index --find-links=wheels openpyxl PyYAML ordered-set zstandard et-xmlfile

if %errorlevel% neq 0 (
    echo ERREUR : échec de l'installation des dépendances
    pause
    exit /b 1
)

REM Étape 2 : installer Nuitka depuis le tarball local (ou PyPI si réseau dispo)
echo [2/3] Installation de Nuitka...
pip install wheels/nuitka-4.1.3.tar.gz 2>nul
if %errorlevel% neq 0 (
    echo Essai via PyPI...
    pip install nuitka
)

REM Étape 3 : compiler
echo [3/3] Compilation avec Nuitka...
python -m nuitka ^
    --standalone ^
    --include-package=openpyxl ^
    --include-package=yaml ^
    --output-dir=..\build ^
    --assume-yes-for-downloads ^
    ..\tahdubawl.py

if %errorlevel% neq 0 (
    echo ERREUR : échec de la compilation
    pause
    exit /b 1
)

echo.
echo === Build terminé ===
echo Exécutable : ..\build\tahdubawl.dist\tahdubawl.exe
echo.
echo Pour créer le dossier portable (dans le dossier parent) :
echo   mkdir ..\tahdubawl-portable
echo   xcopy /E ..\build\tahdubawl.dist\* ..\tahdubawl-portable\
echo   copy ..\config.yaml ..\tahdubawl-portable\
echo   copy LISEZ-MOI.txt ..\tahdubawl-portable\
pause
