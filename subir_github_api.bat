@echo off
chcp 65001 >nul
REM Cambiar al directorio donde está el script
cd /d "%~dp0"
echo.
echo ========================================
echo   SUBIR API A GITHUB
echo ========================================
echo.

set GITHUB_USER=char2020
set GITHUB_REPO=api-hv

echo 🔄 Configurando Git...
REM Verificar si ya existe un repositorio Git
if not exist .git (
    echo 📦 Inicializando repositorio Git...
    git init
    git branch -M main
)

REM Intentar obtener token de variable de entorno primero
if "%GITHUB_TOKEN%"=="" (
    REM Intentar extraer token del remote del proyecto principal (directorio padre)
    pushd ..
    for /f "tokens=*" %%i in ('git config --get remote.origin.url 2^>nul') do set EXISTING_URL=%%i
    popd
    if defined EXISTING_URL (
        echo 🔍 Intentando usar token del proyecto principal...
        REM Extraer token de la URL si existe (formato: https://TOKEN@github.com/...)
        echo %EXISTING_URL% | findstr /C:"@" >nul
        if %errorlevel% equ 0 (
            for /f "tokens=2 delims=@" %%a in ("%EXISTING_URL%") do (
                for /f "tokens=1 delims=/" %%b in ("%%a") do (
                    set GITHUB_TOKEN=%%b
                    echo ✅ Token encontrado y configurado automáticamente
                )
            )
        )
    )
)

REM Si aún no hay token, pedirlo al usuario
if "%GITHUB_TOKEN%"=="" (
    set /p GITHUB_TOKEN="Ingresa tu token de GitHub: "
)
if "%GITHUB_TOKEN%"=="" (
    echo ❌ Token requerido
    echo 💡 Puedes configurarlo como variable de entorno: set GITHUB_TOKEN=tu_token
    pause
    exit /b
)

REM Configurar remote (crear si no existe, actualizar si existe)
git remote remove origin 2>nul
git remote add origin https://%GITHUB_TOKEN%@github.com/%GITHUB_USER%/%GITHUB_REPO%.git

echo.
echo 📋 Agregando archivos de la API a Git...
REM Agregar archivos principales de la API
git add app.py requirements.txt render.yaml README.md .gitignore templates/hv.docx templates/cobro_*.docx ANALISIS-DATOS.md

REM Verificar si hay cambios para subir
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo.
    echo ⚠️ No hay cambios para subir
    echo 💡 Todos los archivos ya están actualizados
    echo.
    echo 🔍 Verificando si hay commits pendientes...
    git log origin/main..HEAD --oneline 2>nul
    if %errorlevel% neq 0 (
        echo 📝 No hay commits locales pendientes
    )
    pause
    exit /b
)

echo.
echo 📊 Archivos que se subirán:
git diff --cached --name-status

echo.
echo 💾 Haciendo commit...
set /p COMMIT_MSG="Ingresa el mensaje del commit (o presiona Enter para 'Actualizar API de generacion de Word'): "
if "%COMMIT_MSG%"=="" set COMMIT_MSG=Actualizar API de generacion de Word
git commit -m "%COMMIT_MSG%"

if %errorlevel% neq 0 (
    echo.
    echo ⚠️ Error al hacer commit
    echo 💡 Puede que no haya cambios para commitear
    pause
    exit /b
)

echo.
echo 🚀 Subiendo a GitHub...
REM Intentar push, si falla puede ser porque el repositorio no existe en GitHub
git push -u origin main

if %errorlevel% equ 0 (
    echo.
    echo ✅ ¡API subida exitosamente!
    echo 🔗 Verifica en: https://github.com/%GITHUB_USER%/%GITHUB_REPO%
    echo.
    echo 📝 IMPORTANTE: Si el repositorio no existe en GitHub:
    echo    1. Ve a https://github.com/new
    echo    2. Crea un repositorio llamado: %GITHUB_REPO%
    echo    3. NO inicialices con README, .gitignore o licencia
    echo    4. Ejecuta este script nuevamente
) else (
    echo.
    echo ❌ Error al subir.
    echo.
    echo 💡 Verifica:
    echo    - Que el token sea correcto
    echo    - Que el repositorio exista en GitHub: https://github.com/%GITHUB_USER%/%GITHUB_REPO%
    echo    - Si el repositorio no existe, créalo primero en GitHub
    echo    - Que tengas cambios para subir
    echo.
    echo 📝 Para crear el repositorio:
    echo    1. Ve a https://github.com/new
    echo    2. Nombre: %GITHUB_REPO%
    echo    3. Descripción: API para generar documentos Word desde hojas de vida
    echo    4. Público o Privado (tu elección)
    echo    5. NO marques "Add a README file"
    echo    6. Ejecuta este script nuevamente
)

echo.
pause

