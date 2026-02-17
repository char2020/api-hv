@echo off
chcp 65001 >nul
REM Cambiar al directorio donde está el script
cd /d "%~dp0"
echo.
echo ========================================
echo   SUBIR API A GITHUB - char2020/api-hv
echo ========================================
echo.

set GITHUB_USER=char2020
set GITHUB_REPO=api-hv
if "%GITHUB_TOKEN%"=="" (
    set /p GITHUB_TOKEN="Ingresa tu token de GitHub (char2020): "
)
if "%GITHUB_TOKEN%"=="" (
    echo ❌ Token requerido. Ejemplo: set GITHUB_TOKEN=tu_token
    pause
    exit /b 1
)

echo 🔄 Configurando Git...
if not exist .git (
    echo 📦 Inicializando repositorio Git...
    git init
    git branch -M main
)

REM Configurar remote para char2020/api-hv
git remote remove origin 2>nul
git remote add origin https://%GITHUB_TOKEN%@github.com/%GITHUB_USER%/%GITHUB_REPO%.git

echo.
echo 📋 Agregando archivos de la API a Git...
git add app.py requirements.txt render.yaml README.md .gitignore templates/hv.docx templates/cobro_*.docx templates/contrato*.docx ANALISIS-DATOS.md

REM Verificar si hay cambios para subir
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo.
    echo ⚠️ No hay cambios para subir
    echo 💡 Todos los archivos ya están actualizados
    pause
    exit /b
)

echo.
echo 📊 Archivos que se subirán:
git diff --cached --name-status

echo.
echo 💾 Haciendo commit...
set /p COMMIT_MSG="Ingresa el mensaje del commit (o Enter para 'Actualizar API'): "
if "%COMMIT_MSG%"=="" set COMMIT_MSG=Actualizar API
git commit -m "%COMMIT_MSG%"

if %errorlevel% neq 0 (
    echo ⚠️ Error al hacer commit
    pause
    exit /b
)

echo.
echo 🚀 Subiendo a GitHub...
git push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ✅ ¡API subida exitosamente!
    echo 🔗 https://github.com/%GITHUB_USER%/%GITHUB_REPO%
) else (
    echo.
    echo ❌ Error al subir. Verifica el token y permisos.
)

echo.
pause
