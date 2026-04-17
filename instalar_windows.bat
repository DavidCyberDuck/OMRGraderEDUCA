@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

echo.
echo ╔══════════════════════════════════════════╗
echo ║  Calificador OMR — Instalación Windows  ║
echo ╚══════════════════════════════════════════╝
echo.
echo Este asistente instalará todo lo necesario:
echo   • Python 3.10+
echo   • Poppler  (lectura de PDFs)
echo   • Librerías de Python
echo.
pause

:: ── Step 1: Python ────────────────────────────────────────────────────────────
echo [1/3] Buscando Python 3.10+...
set PYTHON=
for %%C in (python3.12 python3.11 python3.10 python3 python) do (
    where %%C >nul 2>&1
    if !errorlevel! == 0 (
        %%C -c "import sys; assert sys.version_info >= (3,10)" >nul 2>&1
        if !errorlevel! == 0 (
            set PYTHON=%%C
            goto :python_found
        )
    )
)

echo.
echo     Python 3.10+ no encontrado.
echo     Abriendo la pagina de descarga en tu navegador...
echo.
start https://www.python.org/downloads/
echo     IMPORTANTE: Durante la instalacion de Python, marca la casilla:
echo     "Add Python to PATH"  (esta al inicio del instalador)
echo.
echo     Cuando termines de instalar Python, presiona cualquier tecla.
pause >nul

set PYTHON=
for %%C in (python3.12 python3.11 python3.10 python3 python) do (
    where %%C >nul 2>&1
    if !errorlevel! == 0 (
        %%C -c "import sys; assert sys.version_info >= (3,10)" >nul 2>&1
        if !errorlevel! == 0 (
            set PYTHON=%%C
            goto :python_found
        )
    )
)

echo.
echo ERROR: Python no se detecta. Asegurate de marcar "Add Python to PATH"
echo        durante la instalacion y luego vuelve a ejecutar este instalador.
pause
exit /b 1

:python_found
for /f "tokens=*" %%V in ('%PYTHON% --version') do echo     ✓ %%V

:: ── Step 2: Poppler ───────────────────────────────────────────────────────────
echo.
echo [2/3] Instalando Poppler (necesario para leer PDFs)...

where pdftoppm >nul 2>&1
if %errorlevel% == 0 (
    echo     ✓ Poppler ya instalado
    goto :poppler_done
)

:: Try winget
where winget >nul 2>&1
if %errorlevel% == 0 (
    winget install --id oscarblancartesarabia.poppler -e --silent
    goto :poppler_done
)

:: Try choco
where choco >nul 2>&1
if %errorlevel% == 0 (
    choco install poppler -y
    goto :poppler_done
)

:: Manual fallback
echo.
echo     No se encontro un gestor de paquetes (winget / chocolatey).
echo.
echo     Descarga Poppler manualmente:
echo     https://github.com/oschwartz10612/poppler-windows/releases
echo.
echo     Extrae el ZIP y agrega la carpeta bin\ a tu variable PATH.
echo     Cuando termines, presiona cualquier tecla para continuar.
start https://github.com/oschwartz10612/poppler-windows/releases
pause >nul

:poppler_done

:: ── Step 3: Python deps ───────────────────────────────────────────────────────
echo.
echo [3/3] Instalando librerias de Python...
%PYTHON% -m pip install --upgrade pip --quiet
%PYTHON% -m pip install -r omr_app\requirements.txt
if %errorlevel% neq 0 (
    echo.
    echo ERROR instalando librerias. Revisa el mensaje de arriba.
    pause
    exit /b 1
)
echo     ✓ Librerias instaladas

:: ── Create launcher ───────────────────────────────────────────────────────────
echo.
echo Creando lanzador...
(
    echo @echo off
    echo cd /d "%%~dp0"
    echo %PYTHON% omr_app\main.py
    echo pause
) > "Abrir Calificador.bat"
echo     ✓ Lanzador creado: "Abrir Calificador.bat"

:: ── Done ──────────────────────────────────────────────────────────────────────
echo.
echo ╔══════════════════════════════════════════╗
echo ║          ¡Instalacion completa!          ║
echo ╚══════════════════════════════════════════╝
echo.
echo Para abrir el programa, haz doble clic en:
echo.
echo     "Abrir Calificador.bat"
echo.
echo (Lo encontraras en esta misma carpeta)
echo.
pause
