@echo off
echo === Tahdubawl — Build Windows avec Nuitka ===
echo.

REM Prérequis :
REM   - Python 3.11+ installé
REM   - Visual Studio Build Tools (ou MinGW-w64)
REM   - pip install nuitka ordered-set zstandard

python -m nuitka ^
    --standalone ^
    --enable-plugin=anti-bloat ^
    --include-package=openpyxl ^
    --include-package=yaml ^
    --output-dir=build ^
    --assume-yes-for-downloads ^
    tahdubawl.py

echo.
echo === Build terminé ===
echo Exécutable : build\tahdubawl.dist\tahdubawl.exe
echo.
echo Pour créer le dossier portable :
echo   mkdir tahdubawl-portable
echo   copy build\tahdubawl.dist\tahdubawl.exe tahdubawl-portable\
echo   copy config.yaml tahdubawl-portable\
echo   copy LISEZ-MOI.txt tahdubawl-portable\
pause
