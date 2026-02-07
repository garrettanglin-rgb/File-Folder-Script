#!/bin/bash
# Double-click this file from Finder to generate Avery labels.

cd ~/Desktop/LabelMaker || { echo "ERROR: ~/Desktop/LabelMaker not found."; sleep 5; exit 1; }

source venv/bin/activate

echo "==============================="
echo "  Label Maker"
echo "==============================="
echo ""

python -m label_maker.generate_labels

echo ""
echo "-------------------------------"
echo "Done. This window will close in 5 seconds..."
sleep 5
