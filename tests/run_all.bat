@echo off
REM Lance tous les scénarios de test
setlocal enabledelayedexpansion

set PASS=0
set FAIL=0

for /d %%d in (scenario_0*) do (
    echo === %%d ===
    cd %%d
    if exist resultats rmdir /s /q resultats
    ..\..\build\tahdubawl.dist\tahdubawl.exe > nul 2>&1
    if !errorlevel! equ 0 (
        echo   [OK]
        set /a PASS+=1
    ) else (
        echo   [FAIL]
        set /a FAIL+=1
    )
    cd ..
)

echo.
echo === %PASS%/%PASS%+%FAIL% scenarios OK ===
pause
