@echo off
setlocal enabledelayedexpansion

REM AGRIECO PRO - Generation locale du classeur Sprint 1
REM Ce script est destine aux utilisateurs Windows non techniques.

set "SCRIPT_DIR=%~dp0"
cd /d "%SCRIPT_DIR%"

echo ============================================================
echo AGRIECO PRO - Installation Sprint 1
echo ============================================================
echo.

echo Verification de Python...
where python >nul 2>nul
if %ERRORLEVEL% EQU 0 (
    set "PYTHON_CMD=python"
) else (
    where py >nul 2>nul
    if %ERRORLEVEL% EQU 0 (
        set "PYTHON_CMD=py"
    ) else (
        echo [ERREUR] Python n'est pas installe ou n'est pas disponible dans le PATH.
        echo.
        echo Installez Python depuis https://www.python.org/downloads/windows/
        echo Pendant l'installation, cochez l'option "Add python.exe to PATH".
        echo Relancez ensuite ce fichier build.bat.
        echo.
        pause
        exit /b 1
    )
)

echo Python detecte : !PYTHON_CMD!
echo.

echo Verification du script de generation...
if not exist "scripts\generate_workbook.py" (
    echo [ERREUR] Le fichier scripts\generate_workbook.py est introuvable.
    echo Verifiez que vous avez bien telecharge tout le depot GitHub.
    echo.
    pause
    exit /b 1
)

echo Generation du classeur AGRIECO PRO...
!PYTHON_CMD! "scripts\generate_workbook.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERREUR] La generation du classeur a echoue.
    echo Verifiez les messages affiches ci-dessus, puis reessayez.
    echo.
    pause
    exit /b 1
)

echo.
if exist "AGRIECO_PRO_V1.xlsx" if exist "AGRIECO_PRO_V1.xlsm" (
    echo [OK] Les classeurs ont ete crees avec succes :
    echo   - AGRIECO_PRO_V1.xlsx
    echo   - AGRIECO_PRO_V1.xlsm
    echo.
    echo Vous pouvez maintenant ouvrir AGRIECO_PRO_V1.xlsm dans Microsoft Excel.
    echo Consultez docs\GUIDE_UTILISATEUR_INSTALLATION_SPRINT1.md pour la suite.
    echo.
    pause
    exit /b 0
)

echo [ERREUR] La commande s'est terminee, mais les fichiers attendus sont introuvables.
echo Fichiers attendus : AGRIECO_PRO_V1.xlsx et AGRIECO_PRO_V1.xlsm
echo.
pause
exit /b 1
