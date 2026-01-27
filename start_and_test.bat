@echo off
chcp 65001 >nul
echo ========================================
echo   PRUEBA LOCAL - CUENTA DE COBRO
echo ========================================
echo.

echo 🔍 Verificando dependencias...
python -c "import flask; import flask_cors; import docx" 2>nul
if %errorlevel% neq 0 (
    echo ❌ Faltan dependencias. Instalando...
    pip install -r requirements.txt
    if %errorlevel% neq 0 (
        echo ❌ Error al instalar dependencias
        pause
        exit /b 1
    )
)

echo.
echo 🚀 Iniciando servidor Flask en segundo plano...
start /B python app.py > server.log 2>&1

echo ⏳ Esperando a que el servidor inicie...
timeout /t 5 /nobreak >nul

echo.
echo 🧪 Ejecutando prueba...
python test_cuenta_cobro.py

echo.
echo 📋 Verificando si se generó el archivo...
if exist "test_cuenta_cobro.docx" (
    echo ✅ Archivo generado: test_cuenta_cobro.docx
    echo 📦 Tamaño: 
    dir test_cuenta_cobro.docx | findstr "test_cuenta_cobro.docx"
) else (
    echo ❌ No se generó el archivo
)

echo.
echo 🛑 Deteniendo servidor...
taskkill /F /IM python.exe /FI "WINDOWTITLE eq *app.py*" 2>nul
if %errorlevel% neq 0 (
    echo 💡 Si el servidor sigue corriendo, ciérralo manualmente
)

echo.
pause

