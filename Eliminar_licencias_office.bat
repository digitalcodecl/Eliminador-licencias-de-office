@echo off
chcp 65001 > nul
cls

:: ===================================================
::  ELIMINADOR DE LICENCIAS DE OFFICE
::  Creado por: Digitalcode SPA, Chile
::  Fecha de creación: 28 de febrero de 2025
::  Descripción: Script para eliminar licencias de Office en Windows
:: ===================================================

:: Verificar permisos de administrador
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo.
    echo ERROR: Se requieren permisos de administrador. Reintentando...
    timeout /t 2 >nul
    powershell -Command "Start-Process cmd -ArgumentList '/c \"%~f0\"' -Verb RunAs -WindowStyle Normal"
    exit
)

:inicio
cls
echo.
echo ===================================================
echo               ELIMINADOR DE LICENCIAS DE OFFICE
echo ===================================================
echo Creado por: Digitalcode SPA, Chile
echo Fecha de creación: 28 de febrero de 2025
echo.

:: Buscar ospp.vbs en Program Files y Program Files (x86)
set "osppPath="
for %%d in ("%ProgramFiles%" "%ProgramFiles(x86)%") do (
    for /f "delims=" %%p in ('where /r "%%~d\Microsoft Office" ospp.vbs 2^>nul') do (
        if not defined osppPath set "osppPath=%%p"
    )
)

if not defined osppPath (
    echo.
    echo ERROR: No se encontró ospp.vbs. Office puede no estar instalado.
    echo.
    pause
    exit /b
)

:: Mostrar la ruta encontrada
echo Ruta de ospp.vbs detectada: %osppPath%
echo.

:: Obtener las licencias en el sistema
:get_licenses
cls
echo.
echo [1/3] Obteniendo licencias en el sistema...
echo.
setlocal enabledelayedexpansion
set "licenciasNombres="
set "licenciasClave="
set "licenciasEstado="
set "index=0"

:: Extraer nombres de licencia
for /f "tokens=3*" %%A in ('cscript //Nologo "%osppPath%" /dstatus ^| findstr /C:"LICENSE NAME"') do (
    set /a index+=1
    set "licenciasNombres[!index!]=%%A %%B"
)

:: Extraer claves
set "index=0"
for /f "delims=" %%X in ('cscript //Nologo "%osppPath%" /dstatus ^| findstr /C:"Last 5 characters"') do (
    set /a index+=1
    set "lineaClave=%%X"
    for /f "tokens=*" %%Y in ("!lineaClave!") do (
        set "clave=%%Y"
        set "clave=!clave:~-5!"
        set "clave=!clave: =!"
        set "licenciasClave[!index!]=!clave!"
    )
)

:: Extraer estados
set "index=0"
for /f "tokens=3" %%E in ('cscript //Nologo "%osppPath%" /dstatus ^| findstr /C:"LICENSE STATUS"') do (
    set /a index+=1
    set "estado=%%E"
    if "!estado!"=="LICENSED" (
        set "licenciasEstado[!index!]=ACTIVADA"
    ) else (
        set "licenciasEstado[!index!]=NO ACTIVADA"
    )
)

:: Verificar si hay licencias
if "%index%"=="0" (
    echo No se encontraron licencias en el sistema.
    echo.
    pause
    exit /b
)

:: Imprimir licencias
echo.
echo ===================================================
echo         LICENCIAS ENCONTRADAS EN EL SISTEMA
echo ===================================================
echo.

for /L %%i in (1,1,%index%) do (
    echo [%%i]  !licenciasNombres[%%i]!
    echo      Clave: !licenciasClave[%%i]!
    echo      Estado: !licenciasEstado[%%i]!
    echo --------------------------------------------------
)

:: Pedir selección de licencia
echo.
:seleccionar
set /p seleccion="Ingrese el número de la licencia que desea eliminar (o presione Enter para salir): "

if "%seleccion%"=="" (
    echo.
    echo No se seleccionó ninguna licencia. Saliendo...
    echo.
    pause
    exit /b
)

if %seleccion% GTR %index% (
    echo.
    echo ERROR: Selección inválida. Inténtelo de nuevo.
    echo.
    goto seleccionar
)

:: Confirmar eliminación
echo.
set /p confirmar="¿Seguro que desea eliminar esta licencia? (S/N): "
if /i "%confirmar%" neq "S" goto seleccionar

:: Eliminar licencia
cls
echo.
echo ===================================================
echo           ELIMINANDO LICENCIA SELECCIONADA...
echo ===================================================
echo.
set "claveEliminar=!licenciasClave[%seleccion%]!"
echo Eliminando clave: %claveEliminar%...
echo.
cscript //Nologo "%osppPath%" /unpkey:%claveEliminar% > nul 2>&1

:: Actualizar estado
echo.
echo [3/3] Actualizando estado de activación...
echo.
cscript //Nologo "%osppPath%" /act > nul 2>&1

:: Confirmación
echo.
echo ===================================================
echo               LICENCIA ELIMINADA
echo ===================================================
echo.
echo !licenciasNombres[%seleccion%]!
echo Clave eliminada: %claveEliminar%
echo --------------------------------------------------
echo.

:: Repetir
set /p continuar="¿Desea eliminar otra licencia? (S/N): "
if /i "%continuar%"=="S" goto get_licenses

echo.
echo PROCESO COMPLETADO
pause
