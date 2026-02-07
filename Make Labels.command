#!/bin/bash
# Double-click this file from Finder to generate Avery labels.

cd ~/Desktop/LabelMaker || {
    echo "ERROR: ~/Desktop/LabelMaker not found."
    echo "Press any key to close..."
    read -n 1
    exit 1
}

source venv/bin/activate

echo "==============================="
echo "  Label Maker"
echo "==============================="
echo ""

# Check that the spreadsheet exists
if [ ! -f ~/Desktop/"Case List.numbers" ]; then
    echo "ERROR: Could not find ~/Desktop/Case List.numbers"
    echo ""
    echo "Press any key to close..."
    read -n 1
    exit 1
fi

# Check that the template exists
if [ ! -f ~/Desktop/"File Folder Labels.docx" ]; then
    echo "ERROR: Could not find ~/Desktop/File Folder Labels.docx"
    echo ""
    echo "Press any key to close..."
    read -n 1
    exit 1
fi

# Generate the formatting profile if it doesn't exist yet
if [ ! -f ~/Desktop/label_format.json ]; then
    echo "First run — analyzing template formatting..."
    python -m label_maker.template_analyzer
    echo ""
fi

# Run the generator and capture exit status
python -m label_maker.generate_labels
STATUS=$?

echo ""
if [ $STATUS -ne 0 ]; then
    echo "==============================="
    echo "  Something went wrong."
    echo "==============================="
    echo ""
    echo "Press any key to close..."
    read -n 1
else
    echo "-------------------------------"
    echo "Done. This window will close in 5 seconds..."
    sleep 5
fi
