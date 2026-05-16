#!/bin/bash
# ============================================================
# Build script — Word Tools para Mac (.app / .dmg)
# Requisitos: pip install pyinstaller PyQt6 python-docx requests
# ============================================================

set -e

echo ""
echo " Construindo Word Tools para Mac..."
echo ""

# Limpa builds anteriores
rm -rf dist/WordTools.app dist/WordTools build/

# Gera o .app
pyinstaller \
    --noconfirm \
    --windowed \
    --onedir \
    --name "WordTools" \
    --add-data "app:app" \
    main.py

echo ""
if [ -d "dist/WordTools.app" ]; then
    echo " [OK] Build concluído: dist/WordTools.app"
else
    echo " [ERRO] Build falhou."
    exit 1
fi

# Cria .dmg se hdiutil estiver disponível
echo ""
echo " Criando .dmg..."
hdiutil create \
    -volname "WordTools" \
    -srcfolder "dist/WordTools.app" \
    -ov -format UDZO \
    "dist/WordTools.dmg"

echo ""
echo " [OK] Instalador: dist/WordTools.dmg"
echo ""
