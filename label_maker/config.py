from pathlib import Path

# Path to the Numbers spreadsheet containing label data
NUMBERS_SPREADSHEET_PATH = Path("~/Desktop/Case List.numbers").expanduser()

# Path to the existing Avery template Word document (Avery 5026)
AVERY_TEMPLATE_PATH = Path("~/Desktop/File Folder Labels.docx").expanduser()

# ---------------------------------------------------------------------------
# Field mapping: which spreadsheet columns go on which label lines.
#
# Each entry maps a label line number (1-based) to a column header name
# that must match your spreadsheet exactly.  Only the columns listed here
# will appear on the label — everything else is ignored.
#
# The line number also determines which formatting from the template
# profile is applied (line 1 gets the first line's font/size/bold, etc.).
# ---------------------------------------------------------------------------
FIELD_MAPPING = {
    1: "File Number",   # Column B — label line 1
    2: "Client",        # Column C — label line 2
    3: "Case",          # Column D — label line 3
    4: "Email/Phone",   # Column E — label line 4
}
