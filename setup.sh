#!/bin/bash
# One-time setup for LabelMaker.
# Run this after cloning the repo to ~/Desktop/LabelMaker/

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "==============================="
echo "  LabelMaker Setup"
echo "==============================="
echo ""

# 1. Create virtual environment
if [ ! -d "venv" ]; then
    echo "Creating virtual environment..."
    python3 -m venv venv
else
    echo "Virtual environment already exists."
fi

# 2. Install dependencies
echo "Installing dependencies..."
source venv/bin/activate
pip install --quiet numbers-parser python-docx openpyxl

# 3. Copy the shortcut to the Desktop
echo "Placing Make Labels.command on your Desktop..."
cp "Make Labels.command" ~/Desktop/
chmod +x ~/Desktop/"Make Labels.command"

# 4. Analyze the template if the .docx exists
TEMPLATE="$HOME/Desktop/File Folder Labels.docx"
PROFILE="$HOME/Desktop/label_format.json"

if [ -f "$TEMPLATE" ]; then
    echo "Analyzing label template..."
    python -m label_maker.template_analyzer
    echo ""
else
    echo ""
    echo "NOTE: Template not found at:"
    echo "  $TEMPLATE"
    echo ""
    echo "Place your Avery 5026 template there, then run:"
    echo "  cd ~/Desktop/LabelMaker && source venv/bin/activate"
    echo "  python -m label_maker.template_analyzer"
    echo ""
fi

echo "==============================="
echo "  Setup complete!"
echo "==============================="
echo ""
echo "To generate labels, double-click 'Make Labels.command' on your Desktop."
echo ""
