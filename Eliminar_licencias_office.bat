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

:: Buscar y ejecutar ospp.vbs en versiones conocidas de Office
set "osppPath="
set "officeVersion="

for %%a in (16,15,14,24) do (
    if exist "%ProgramFiles%\Microsoft Office\Office%%a\ospp.vbs" (
        set "osppPath=%ProgramFiles%\Microsoft Office\Office%%a\ospp.vbs"
        set "officeVersion=Office %%a (64-bit)"
    )
    if exist "%ProgramFiles(x86)%\Microsoft Office\Office%%a\ospp.vbs" (
        set "osppPath=%ProgramFiles(x86)%\Microsoft Office\Office%%a\ospp.vbs"
        set "officeVersion=Office %%a (32-bit)"
    )
)

if not defined osppPath (
    echo.
    echo ERROR: No se encontró ospp.vbs. Office puede no estar instalado.
    echo.
    pause
    exit /b
)

:: Mostrar la versión de Office detectada
echo Versión de Office detectada: %officeVersion%
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

:: Extraer claves y estados
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

:: Verificar si hay licencias en el sistema
if "%index%"=="0" (
    echo No se encontraron licencias en el sistema.
    echo.
    pause
    exit /b
)

:: Imprimir licencias con separación clara
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

:: Si el usuario presiona Enter sin escribir nada, salir del script
if "%seleccion%"=="" (
    echo.
    echo No se seleccionó ninguna licencia. Saliendo...
    echo.
    pause
    exit /b
)

:: Verificar que la selección sea válida
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

:: Proceso de eliminación
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

:: Forzar actualización de estado de activación sin ejecutar de nuevo todo el script
echo.
echo [3/3] Actualizando estado de activación...
echo.
cscript //Nologo "%osppPath%" /act > nul 2>&1

:: Confirmación de eliminación
echo.
echo ===================================================
echo               LICENCIA ELIMINADA
echo ===================================================
echo.
echo !licenciasNombres[%seleccion%]!
echo Clave eliminada: %claveEliminar%
echo --------------------------------------------------
echo.

:: Preguntar si desea eliminar otra licencia
set /p continuar="¿Desea eliminar otra licencia? (S/N): "
if /i "%continuar%"=="S" goto get_licenses

echo.
echo PROCESO COMPLETADO
pause
