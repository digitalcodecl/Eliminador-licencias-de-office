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
echo      🗑️  ELIMINADOR DE LICENCIAS DE OFFICE
echo =====================================
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
    echo ❌ ERROR: No se encontró ospp.vbs. Office puede no estar instalado.
    pause
    exit /b
)

:: Mostrar la versión de Office detectada
echo 🔍 Versión de Office detectada: %officeVersion%
echo.

:: Obtener las licencias activas con su nombre y clave
echo [1/3] Obteniendo licencias activas...
setlocal enabledelayedexpansion
set "licencias="
set "licenciasNombres="

:: Extraer nombres de licencia (LICENSE NAME)
for /f "tokens=3*" %%A in ('cscript //Nologo "%osppPath%" /dstatus ^| findstr /C:"LICENSE NAME"') do (
    set "licenciasNombres=!licenciasNombres! %%A %%B|"
)

:: Extraer la línea completa que contiene "Last 5 characters"
for /f "delims=" %%X in ('cscript //Nologo "%osppPath%" /dstatus ^| findstr /C:"Last 5 characters"') do (
    set "lineaClave=%%X"
)

:: Extraer solo los últimos 5 caracteres de la línea capturada, eliminando espacios
for /f "tokens=*" %%Y in ("%lineaClave%") do (
    set "clave=%%Y"
    set "clave=!clave:~-5!"
    set "clave=!clave: =!"  :: Elimina espacios en blanco
    set "licencias=!licencias!!clave!|"
)

:: Verificar si se encontraron claves reales
if "%licencias%"=="" (
    echo ❌ No se encontraron licencias activas.
    pause
    exit /b
)

:: Mostrar las licencias detectadas antes de eliminarlas
echo =====================================
echo  🔎 Licencias activas encontradas:
set "count=1"
for /f "delims=|" %%C in ("%licencias%") do (
    for /f "delims=|" %%N in ("%licenciasNombres%") do (
        echo [%count%] %%N - Clave: %%C
        set /a count+=1
    )
)
echo =====================================
set /p confirm="¿Deseas continuar y eliminarlas? (S/N): "
if /i "%confirm%" neq "S" exit /b

:: Eliminar cada licencia encontrada y guardar en variable
echo [2/3] Eliminando licencias...
set "eliminadas="
for /f "delims=|" %%C in ("%licencias%") do (
    if not "%%C"=="" (
        echo 🔴 Eliminando clave: %%C...
        cscript //Nologo "%osppPath%" /unpkey:%%C > nul 2>&1
        set "eliminadas=!eliminadas!%%C|"
    )
)

:: Forzar actualización de estado de activación
echo [3/3] Actualizando estado de activación...
cscript //Nologo "%osppPath%" /act > nul 2>&1

:: Mostrar las licencias eliminadas
echo =====================================
echo  ✅ LICENCIAS ELIMINADAS:
set "count=1"
for /f "delims=|" %%D in ("%eliminadas%") do (
    for /f "delims=|" %%N in ("%licenciasNombres%") do (
        echo [%count%] %%N - Clave: %%D
        set /a count+=1
    )
)
echo =====================================

echo ✅ PROCESO COMPLETADO
pause
