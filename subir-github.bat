@echo off
echo ========================================
echo Subiendo API a GitHub
echo ========================================
echo.

cd api

echo Verificando estado de Git...
git status

echo.
echo Agregando archivos...
git add app.py
git add requirements.txt
git add render.yaml
git add README.md
git add .gitignore
git add templates/hv.docx

echo.
echo Ingresa el mensaje del commit (o presiona Enter para usar mensaje por defecto):
set /p commit_msg="Mensaje: "
if "%commit_msg%"=="" set commit_msg="Actualizar API de generacion de Word"

echo.
echo Haciendo commit...
git commit -m "%commit_msg%"

echo.
echo Subiendo a GitHub...
git push origin main

echo.
echo ========================================
echo ¡Subida completada!
echo ========================================
pause

