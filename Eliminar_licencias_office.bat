@echo off
chcp 65001 > nul
cls
color 0A

:: Verificar si el script se ejecuta con permisos de administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo ❌ Se requieren permisos de administrador. Reintentando...
    timeout /t 2 >nul
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs -WindowStyle Normal"
    exit
)

echo.
echo =====================================
echo      🗑️  ELIMINANDO LICENCIAS DE OFFICE
echo =====================================
echo.

:: Buscar y ejecutar ospp.vbs en versiones conocidas de Office
set "osppPath="

for %%a in (4,5,6) do (
    if exist "%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs" (
        set "osppPath=%ProgramFiles%\Microsoft Office\Office1%%a\ospp.vbs"
    )
    if exist "%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs" (
        set "osppPath=%ProgramFiles(x86)%\Microsoft Office\Office1%%a\ospp.vbs"
    )
)

if not defined osppPath (
    echo ❌ ERROR: No se encontró ospp.vbs. Office puede no estar instalado.
    pause
    exit /b
)

:: Obtener las claves de producto activas correctamente
echo [1/3] Obteniendo licencias activas...
setlocal enabledelayedexpansion
set "licencias="

:: Extraer los últimos 5 caracteres de cada clave
for /f "tokens=8" %%B in ('cscript //Nologo "%osppPath%" /dstatus ^| findstr /b /c:"Last 5"') do (
    set "licencias=!licencias! %%B"
)

:: Verificar si se encontraron claves reales
if "%licencias%"=="" (
    echo ❌ No se encontraron licencias activas.
    pause
    exit /b
)

:: Mostrar las claves detectadas antes de eliminarlas
echo =====================================
echo  🔎 Se encontraron las siguientes licencias activas:
for %%C in (%licencias%) do echo - %%C
echo =====================================
set /p confirm="¿Deseas continuar y eliminarlas? (S/N): "
if /i "%confirm%" neq "S" exit /b

:: Eliminar cada licencia encontrada y guardar en variable
echo [2/3] Eliminando licencias...
set "eliminadas="
for %%C in (%licencias%) do (
    echo 🔴 Eliminando clave: %%C...
    cscript //Nologo "%osppPath%" /unpkey:%%C > nul 2>&1
    set "eliminadas=!eliminadas! %%C "
)

:: Forzar actualización de estado de activación
echo [3/3] Actualizando estado de activación...
cscript //Nologo "%osppPath%" /act > nul 2>&1

:: Mostrar las licencias eliminadas
echo =====================================
echo  ✅ LICENCIAS ELIMINADAS:
for %%D in (%eliminadas%) do echo - %%D
echo =====================================

echo ✅ PROCESO COMPLETADO
pause

