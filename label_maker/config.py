from pathlib import Path

# Path to the Numbers spreadsheet containing label data
NUMBERS_SPREADSHEET_PATH = Path("~/Desktop/Case List.numbers").expanduser()

# Path to the existing Avery template Word document (Avery 5026)
AVERY_TEMPLATE_PATH = Path("~/Desktop/File Folder Labels.docx").expanduser()

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
    1: {"bold": True},     # Client + File Number — bold
    2: {"bold": False},    # Case — not bold
}
