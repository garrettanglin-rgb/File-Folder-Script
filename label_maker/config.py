from pathlib import Path

# Path to the Numbers spreadsheet containing label data
NUMBERS_SPREADSHEET_PATH = Path("~/Desktop/Case List.numbers").expanduser()

# Path to the existing Avery template Word document (Avery 5026)
AVERY_TEMPLATE_PATH = Path("~/Desktop/Avery5026ExtraLargeFileFolderLabels-2.docx").expanduser()

# ---------------------------------------------------------------------------
# Field mapping: which spreadsheet columns go on which label lines.
#
# Each entry maps a label line number (1-based) to either:
#   - A single column name (string)  → that column fills the whole line
#   - A tuple of two column names    → first on the left, second on the right
#
# Column names must match your spreadsheet headers exactly.
# The line number determines which template formatting is applied.
# ---------------------------------------------------------------------------
FIELD_MAPPING = {
    1: ("Client", "File Number"),  # Left: Client, Right: File Number
    2: "Case",                     # Column D — full line
    3: "Phone/Email:",             # Column E — full line
}

# ---------------------------------------------------------------------------
# Formatting overrides per label line.
#
# Use this to override any formatting detected from the template.
# Keys are line numbers (matching FIELD_MAPPING).
# Values are dicts with any of: bold, italic, underline, font_name, font_size_pt
# Only properties listed here are overridden; everything else comes from the template.
# ---------------------------------------------------------------------------
LINE_FORMAT_OVERRIDE = {
    1: {"font_name": "Times New Roman", "font_size_pt": 12, "bold": True},
    2: {"font_name": "Times New Roman", "font_size_pt": 12, "bold": False},
    3: {"font_name": "Times New Roman", "font_size_pt": 12, "bold": False},
}

# ---------------------------------------------------------------------------
# Left indent for the colored tab area on physical labels.
#
# Avery file folder labels have a colored strip on the left edge.
# This indent (in points) pushes all text to the right so it prints
# on the white area of the label, not over the colored tab.
#
# Set to 0 to disable.  Adjust the value if text overlaps or is too
# far from the edge.  1 inch = 72 points.
# ---------------------------------------------------------------------------
LABEL_LEFT_INDENT_PT = 32   # ~0.44 inches — clears the Avery 5026 color tab

# ---------------------------------------------------------------------------
# Maximum characters per label line.
#
# Long values (e.g. case names) can wrap to the next line and push other
# fields down.  Set a character limit per line number to truncate with "…".
# Only lines listed here are truncated; omitted lines have no limit.
# Set to None or 0 to disable for a specific line.
# ---------------------------------------------------------------------------
LINE_MAX_CHARS = {
    2: 40,   # Case — keep on one line
}
