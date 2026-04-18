#!/bin/bash
# Calificador OMR — Instalador para macOS
# Double-click this file in Finder to run.

cd "$(dirname "$BASH_SOURCE")"

# ── Dialog helpers ─────────────────────────────────────────────────────────────
msg() {
    osascript -e "display dialog \"$1\" buttons {\"Continuar\"} default button \"Continuar\" with title \"Calificador OMR — Instalación\" with icon note" 2>/dev/null \
    || echo -e "\n$1\n[Presiona Enter para continuar]" && read
}

msg_ok() {
    osascript -e "display dialog \"$1\" buttons {\"¡Listo!\"} default button \"¡Listo!\" with title \"Calificador OMR — Instalación\" with icon note" 2>/dev/null \
    || echo -e "\n$1"
}

msg_err() {
    osascript -e "display alert \"Error\" message \"$1\" as critical" 2>/dev/null \
    || echo -e "\nERROR: $1"
    exit 1
}

# ── Welcome ────────────────────────────────────────────────────────────────────
osascript -e 'display dialog "Bienvenido al instalador del Calificador de Exámenes OMR.\n\nEste asistente instalará todo lo necesario:\n  • Python 3.10+\n  • Homebrew\n  • Poppler (lectura de PDFs)\n  • Librerías de Python\n\nPresiona Continuar para comenzar." buttons {"Continuar"} default button "Continuar" with title "Calificador OMR — Instalación" with icon note' 2>/dev/null

echo ""
echo "╔══════════════════════════════════════════╗"
echo "║   Calificador OMR — Instalación macOS   ║"
echo "╚══════════════════════════════════════════╝"
echo ""

# ── Step 1: Python ─────────────────────────────────────────────────────────────
echo "[1/4] Buscando Python 3.10+..."
PYTHON=""
for cmd in python3.12 python3.11 python3.10 python3; do
    if command -v "$cmd" &>/dev/null; then
        if "$cmd" -c "import sys; assert sys.version_info >= (3,10)" 2>/dev/null; then
            PYTHON="$cmd"
            break
        fi
    fi
done

if [ -z "$PYTHON" ]; then
    osascript -e 'display dialog "Python 3.10 o superior no está instalado.\n\nSe abrirá la página de descarga. Instala la versión más reciente (3.12 recomendado) y luego regresa aquí y haz clic en Continuar." buttons {"Abrir python.org"} default button "Abrir python.org" with title "Calificador OMR — Instalación" with icon caution' 2>/dev/null
    open "https://www.python.org/downloads/"
    osascript -e 'display dialog "Una vez que hayas instalado Python, cierra el instalador de Python y regresa aquí." buttons {"Ya lo instalé"} default button "Ya lo instalé" with title "Calificador OMR — Instalación" with icon note' 2>/dev/null

    for cmd in python3.12 python3.11 python3.10 python3; do
        if command -v "$cmd" &>/dev/null; then
            if "$cmd" -c "import sys; assert sys.version_info >= (3,10)" 2>/dev/null; then
                PYTHON="$cmd"
                break
            fi
        fi
    done

    if [ -z "$PYTHON" ]; then
        msg_err "Python no se detectó. Por favor reinicia la Terminal y vuelve a ejecutar el instalador."
    fi
fi

echo "    ✓ $($PYTHON --version)"

# ── Step 2: Homebrew ───────────────────────────────────────────────────────────
echo "[2/4] Verificando Homebrew..."
BREW=$(command -v brew 2>/dev/null)

if [ -z "$BREW" ]; then
    msg "Homebrew no está instalado. Se instalará ahora. Esto puede tardar unos minutos — mira la Terminal."
    echo "    Instalando Homebrew..."
    /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
    for candidate in /opt/homebrew/bin/brew /usr/local/bin/brew; do
        [ -f "$candidate" ] && BREW="$candidate" && break
    done
    [ -n "$BREW" ] && eval "$($BREW shellenv)"
fi

if [ -z "$BREW" ]; then
    msg_err "No se pudo instalar Homebrew. Visita https://brew.sh e inténtalo manualmente."
fi
echo "    ✓ Homebrew $(${BREW} --version | head -1)"

# ── Step 3: Poppler ────────────────────────────────────────────────────────────
echo "[3/4] Instalando Poppler (lectura de PDFs)..."
if ! command -v pdftoppm &>/dev/null; then
    $BREW install poppler
fi
echo "    ✓ Poppler instalado"

# ── Step 4: Python deps ────────────────────────────────────────────────────────
echo "[4/4] Instalando librerías de Python..."
$PYTHON -m pip install --upgrade pip --quiet
$PYTHON -m pip install -r omr_app/requirements.txt
echo "    ✓ Librerías instaladas"

# ── Create launcher ────────────────────────────────────────────────────────────
LAUNCHER="Abrir Calificador.command"
cat > "$LAUNCHER" <<LAUNCHEREOF
#!/bin/bash
cd "\$(dirname "\$BASH_SOURCE")"

echo ""
echo "================================"
echo "  Calificador OMR"
echo "================================"
echo ""
echo "Verificando actualizaciones..."

if command -v git &>/dev/null && [ -d ".git" ]; then
    if git fetch origin main --quiet 2>/dev/null; then
        BEHIND=\$(git rev-list HEAD..origin/main --count 2>/dev/null || echo 0)
        if [ "\$BEHIND" -gt "0" ]; then
            echo "  \$BEHIND actualización(es) disponible(s). Actualizando..."
            git pull origin main --quiet
            $PYTHON -m pip install -r omr_app/requirements.txt --quiet
            echo "  Actualizado correctamente."
        else
            echo "  Ya tienes la versión más reciente."
        fi
    else
        echo "  Sin conexión — omitiendo verificación."
    fi
else
    echo "  (Sin repositorio git — descarga ZIP detectada)"
fi

echo ""
$PYTHON omr_app/main.py
LAUNCHEREOF
chmod +x "$LAUNCHER"
echo ""
echo "    ✓ Lanzador creado: '$LAUNCHER'"

# ── Done ───────────────────────────────────────────────────────────────────────
echo ""
echo "╔══════════════════════════════════════════╗"
echo "║          ¡Instalación completa!          ║"
echo "╚══════════════════════════════════════════╝"
echo ""

osascript -e 'display dialog "¡Instalación completa!\n\nPara abrir el programa, haz doble clic en:\n\n     \"Abrir Calificador.command\"\n\nEncontrarás ese archivo en esta misma carpeta." buttons {"¡Listo!"} default button "¡Listo!" with title "Calificador OMR — Instalación" with icon note' 2>/dev/null \
|| echo "Abre 'Abrir Calificador.command' para iniciar el programa."
