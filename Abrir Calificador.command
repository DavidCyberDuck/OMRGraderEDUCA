#!/bin/bash
cd "$(dirname "$BASH_SOURCE")"

echo ""
echo "================================"
echo "  Calificador OMR"
echo "================================"
echo ""
echo "Verificando actualizaciones..."

if ! command -v git &>/dev/null; then
    echo "  Git no está instalado."
    read -rp "  ¿Instalar Git para recibir actualizaciones automáticas? (s/n): " resp
    if [[ "$resp" =~ ^[ssSy] ]]; then
        echo "  Instalando Git via Homebrew..."
        if ! command -v brew &>/dev/null; then
            /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
            for candidate in /opt/homebrew/bin/brew /usr/local/bin/brew; do
                [ -f "$candidate" ] && eval "$($candidate shellenv)" && break
            done
        fi
        brew install git
        echo "  Git instalado. Reinicia el lanzador para activar actualizaciones."
    else
        echo "  De acuerdo, continuando sin Git."
    fi
elif [ ! -d ".git" ]; then
    echo "  (Sin repositorio git — descarga ZIP detectada)"
    read -rp "  ¿Convertir en repositorio git para recibir actualizaciones? (s/n): " resp
    if [[ "$resp" =~ ^[ssSy] ]]; then
        git clone --no-checkout https://github.com/DavidCyberDuck/OMRGraderEDUCA.git .git_tmp 2>/dev/null \
        && mv .git_tmp/.git . && rm -rf .git_tmp \
        && git reset HEAD --quiet
        echo "  Repositorio configurado. Reinicia el lanzador para activar actualizaciones."
    fi
else
    if git fetch origin main --quiet 2>/dev/null; then
        BEHIND=$(git rev-list HEAD..origin/main --count 2>/dev/null || echo 0)
        if [ "$BEHIND" -gt "0" ]; then
            echo "  $BEHIND actualización(es) disponible(s). Actualizando..."
            git pull origin main --quiet
            python3 -m pip install -r omr_app/requirements.txt --quiet
            echo "  Actualizado correctamente."
        else
            echo "  Ya tienes la versión más reciente."
        fi
    else
        echo "  Sin conexión — omitiendo verificación."
    fi
fi

echo ""
python3 omr_app/main.py
