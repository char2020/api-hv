@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo.
echo ========================================
echo   OBTENER TOKEN DE GOOGLE DRIVE
echo ========================================
echo.
echo 1. Iniciando la API (se abrirá otra ventana)...
start "API Drive" cmd /k "cd /d %~dp0 && python app.py"
echo.
echo 2. Esperando que la API arranque (5 seg)...
timeout /t 5 /nobreak >nul
echo.
echo 3. Abriendo navegador para autorizar...
start http://localhost:5000/get-drive-token
echo.
echo ========================================
echo   INSTRUCCIONES
echo ========================================
echo.
echo - En el navegador que se abrió, inicia sesión con Google
echo   (la cuenta donde quieres guardar los archivos en Drive)
echo.
echo - Haz clic en "Permitir" para autorizar la aplicación
echo.
echo - Se creará el archivo token.json en esta carpeta
echo.
echo - Copia TODO el contenido de token.json
echo.
echo - En Render: Environment -^> Añadir variable:
echo   Nombre: GOOGLE_DRIVE_CREDENTIALS
echo   Valor: (pega el JSON copiado)
echo.
echo - Guarda y espera el redespliegue
echo.
echo ========================================
echo.
echo La API sigue corriendo. Presiona Ctrl+C en la ventana
echo de la API para detenerla cuando termines.
echo.
pause
