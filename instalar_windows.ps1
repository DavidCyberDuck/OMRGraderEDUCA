# Calificador OMR — Instalador para Windows
# Invocado automaticamente por instalar_windows.bat

Set-Location $PSScriptRoot
$ErrorActionPreference = "Stop"

function Pause-Key($msg = "Presiona cualquier tecla para continuar...") {
    Write-Host ""
    Write-Host $msg
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}

Clear-Host
Write-Host ""
Write-Host "=========================================="
Write-Host "  Calificador OMR - Instalacion Windows  "
Write-Host "=========================================="
Write-Host ""
Write-Host "Este asistente instalara todo lo necesario:"
Write-Host "  - Python 3.10+"
Write-Host "  - Poppler (lectura de PDFs)"
Write-Host "  - Librerias de Python"
Write-Host ""
Pause-Key "Presiona cualquier tecla para comenzar..."

# ── Step 1: Python ─────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "[1/3] Buscando Python 3.10+..."

function Find-Python {
    foreach ($cmd in @("py", "python", "python3")) {
        try {
            $ver = & $cmd -c "import sys; print(sys.version_info >= (3,10))" 2>$null
            if ($ver -eq "True") { return $cmd }
        } catch {}
    }
    return $null
}

$PYTHON = Find-Python

if (-not $PYTHON) {
    Write-Host ""
    Write-Host "    Python 3.10+ no encontrado."
    Write-Host "    Abriendo la pagina de descarga en tu navegador..."
    Write-Host ""
    Start-Process "https://www.python.org/downloads/"
    Write-Host "    IMPORTANTE: Durante la instalacion de Python, marca la casilla:"
    Write-Host '    "Add Python to PATH"  (esta al inicio del instalador)'
    Write-Host ""
    Pause-Key "    Cuando termines de instalar Python, presiona cualquier tecla."

    $PYTHON = Find-Python

    if (-not $PYTHON) {
        Write-Host ""
        Write-Host "ERROR: Python no se detecta."
        Write-Host "Asegurate de marcar 'Add Python to PATH' durante la instalacion"
        Write-Host "y luego vuelve a ejecutar este instalador."
        Pause-Key
        exit 1
    }
}

$pyVer = & $PYTHON --version 2>&1
Write-Host "    OK  $pyVer"

# ── Step 2: Poppler ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "[2/3] Verificando Poppler (necesario para leer PDFs)..."

$popplerOk = $null -ne (Get-Command "pdftoppm" -ErrorAction SilentlyContinue)

if ($popplerOk) {
    Write-Host "    OK  Poppler ya instalado"
} else {
    $winget = $null -ne (Get-Command "winget" -ErrorAction SilentlyContinue)
    $choco  = $null -ne (Get-Command "choco"  -ErrorAction SilentlyContinue)

    if ($winget) {
        Write-Host "    Instalando via winget..."
        winget install --id oschwartz10612.poppler -e --silent
    } elseif ($choco) {
        Write-Host "    Instalando via chocolatey..."
        choco install poppler -y
    } else {
        Write-Host ""
        Write-Host "    No se encontro winget ni chocolatey."
        Write-Host "    Abriendo pagina de descarga manual de Poppler..."
        Write-Host ""
        Write-Host "    1. Descarga el ZIP de la pagina que se abrira"
        Write-Host "    2. Extrae el ZIP (ej. C:\poppler)"
        Write-Host "    3. Agrega la carpeta bin\ al PATH de Windows"
        Write-Host "       (Busca 'Variables de entorno' en el menu Inicio)"
        Write-Host ""
        Start-Process "https://github.com/oschwartz10612/poppler-windows/releases"
        Pause-Key "    Cuando termines, presiona cualquier tecla para continuar."
    }
}

# ── Step 3: Python deps ────────────────────────────────────────────────────────
Write-Host ""
Write-Host "[3/3] Instalando librerias de Python..."

& $PYTHON -m pip install --upgrade pip --quiet
& $PYTHON -m pip install -r "omr_app\requirements.txt"

if ($LASTEXITCODE -ne 0) {
    Write-Host ""
    Write-Host "ERROR instalando librerias. Revisa el mensaje de arriba."
    Pause-Key
    exit 1
}
Write-Host "    OK  Librerias instaladas"

# ── Create launcher ────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "Creando lanzador..."

$launcher = "Abrir Calificador.bat"
@"
@echo off
cd /d "%~dp0"
$PYTHON omr_app\main.py
pause
"@ | Set-Content -Encoding ASCII $launcher

Write-Host "    OK  Lanzador creado: `"$launcher`""

# ── Done ───────────────────────────────────────────────────────────────────────
Write-Host ""
Write-Host "=========================================="
Write-Host "         Instalacion completa!            "
Write-Host "=========================================="
Write-Host ""
Write-Host "Para abrir el programa, haz doble clic en:"
Write-Host ""
Write-Host "    `"Abrir Calificador.bat`""
Write-Host ""
Write-Host "(Lo encontraras en esta misma carpeta)"
Write-Host ""
Pause-Key
