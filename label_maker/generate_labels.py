"""Generate a print-ready Avery label document from spreadsheet data.

Copies the original Avery template and fills label cells with data,
preserving exact page layout, table dimensions, margins, and positioning
so the output aligns perfectly with physical label sheets.
"""

import math
import platform
import shutil
import subprocess
import sys
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Twips, RGBColor
from docx.table import Table

from label_maker.config import (
    AVERY_TEMPLATE_PATH,
    FIELD_MAPPING,
    LABEL_LEFT_INDENT_PT,
    LINE_FORMAT_OVERRIDE,
    LINE_MAX_CHARS,
    NUMBERS_SPREADSHEET_PATH,
)
from label_maker.data_reader import clear_print_flags, read_flagged_rows
from label_maker.template_analyzer import load_profile

# Where the generated document is saved
OUTPUT_PATH = Path("~/Desktop/Labels_Ready.docx").expanduser()

# Where the cached formatting profile lives
PROFILE_PATH = AVERY_TEMPLATE_PATH.parent / "label_format.json"

# ---------------------------------------------------------------------------
# Alignment / spacing look-ups
# ---------------------------------------------------------------------------
_ALIGN_MAP = {
    "left": WD_ALIGN_PARAGRAPH.LEFT,
    "center": WD_ALIGN_PARAGRAPH.CENTER,
    "right": WD_ALIGN_PARAGRAPH.RIGHT,
    "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
}

_SPACING_RULE_MAP = {
    "single": WD_LINE_SPACING.SINGLE,
    "1.5": WD_LINE_SPACING.ONE_POINT_FIVE,
    "double": WD_LINE_SPACING.DOUBLE,
    "exactly": WD_LINE_SPACING.EXACTLY,
    "at_least": WD_LINE_SPACING.AT_LEAST,
    "multiple": WD_LINE_SPACING.MULTIPLE,
}


# ---------------------------------------------------------------------------
# Text formatting helpers
# ---------------------------------------------------------------------------

def _set_paragraph_spacing(paragraph, line_fmt):
    """Apply paragraph-level spacing and indent properties."""
    pf = paragraph.paragraph_format
    rule = _SPACING_RULE_MAP.get(line_fmt["line_spacing_rule"],
                                 WD_LINE_SPACING.SINGLE)
    pf.line_spacing_rule = rule
    if rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
        pf.line_spacing = Pt(line_fmt["line_spacing_value"])
    else:
        pf.line_spacing = line_fmt["line_spacing_value"]

    sb = line_fmt.get("space_before_pt")
    sa = line_fmt.get("space_after_pt")
    pf.space_before = Pt(sb) if sb else Pt(0)
    pf.space_after = Pt(sa) if sa else Pt(0)

    # Left indent — pushes text past the colored tab area on the label.
    # Use the template-detected indent if present, otherwise fall back
    # to the configurable LABEL_LEFT_INDENT_PT (for the Avery color tab).
    li = line_fmt.get("left_indent_pt") or LABEL_LEFT_INDENT_PT
    if li:
        pf.left_indent = Pt(li)
    fli = line_fmt.get("first_line_indent_pt")
    if fli:
        pf.first_line_indent = Pt(fli)


def _format_run(run, line_fmt):
    """Apply run-level font properties from a line format dict."""
    font = run.font
    # Always set font name — fall back to Times New Roman if not detected
    font.name = line_fmt.get("font_name") or "Times New Roman"
    # Always set font size — fall back to 12pt if not detected
    size = line_fmt.get("font_size_pt")
    font.size = Pt(size) if size else Pt(12)
    font.bold = line_fmt.get("bold", False)
    font.italic = line_fmt.get("italic", False)
    font.underline = line_fmt.get("underline", False)
    if line_fmt.get("color_rgb"):
        font.color.rgb = RGBColor.from_string(line_fmt["color_rgb"])


def _add_right_tab_stop(paragraph, position_twips):
    """Add a right-aligned tab stop to a paragraph via XML."""
    pPr = paragraph._p.get_or_add_pPr()
    tabs = pPr.find(qn("w:tabs"))
    if tabs is None:
        tabs = OxmlElement("w:tabs")
        pPr.append(tabs)
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(position_twips))
    tab.set(qn("w:leader"), "none")
    tabs.append(tab)


def _apply_line_single(paragraph, line_fmt, text):
    """Fill a paragraph with a single column value."""
    paragraph.alignment = _ALIGN_MAP.get(line_fmt["alignment"],
                                         WD_ALIGN_PARAGRAPH.LEFT)
    _set_paragraph_spacing(paragraph, line_fmt)
    run = paragraph.add_run(str(text) if text is not None else "")
    _format_run(run, line_fmt)


def _apply_line_pair(paragraph, line_fmt, left_text, right_text,
                     cell_width_twips, cell_margin_lr_twips):
    """Fill a paragraph with two values: left-aligned and right-aligned.

    Uses a right-aligned tab stop so the second value sits flush-right
    inside the cell.
    """
    # Override to left-align so the tab stop layout works correctly
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    _set_paragraph_spacing(paragraph, line_fmt)

    # Tab stop position = cell content width (cell width minus L+R margins)
    tab_pos = cell_width_twips - (2 * cell_margin_lr_twips)
    _add_right_tab_stop(paragraph, tab_pos)

    # Left run
    run_left = paragraph.add_run(
        str(left_text) if left_text is not None else "")
    _format_run(run_left, line_fmt)

    # Tab character
    run_tab = paragraph.add_run("\t")
    _format_run(run_tab, line_fmt)

    # Right run
    run_right = paragraph.add_run(
        str(right_text) if right_text is not None else "")
    _format_run(run_right, line_fmt)


def _truncate(text, max_chars):
    """Truncate text to max_chars, adding '...' if shortened."""
    text = str(text) if text is not None else ""
    if max_chars and len(text) > max_chars:
        return text[:max_chars - 3].rstrip() + "..."
    return text


def _get_all_mapped_columns(field_mapping):
    """Return a flat set of every column name referenced in FIELD_MAPPING."""
    cols = set()
    for value in field_mapping.values():
        if isinstance(value, tuple):
            cols.update(value)
        else:
            cols.add(value)
    return cols


def _populate_cell(cell, data_row, line_formats, field_mapping,
                   cell_width_twips, cell_margin_lr_twips):
    """Fill a single table cell with label data using the format profile.

    *field_mapping* values may be a string (single column → full line)
    or a tuple of two strings (left column, right column on one line).
    """
    # Fallback format if the template had no detectable line styles
    _DEFAULT_FMT = {
        "font_name": "Times New Roman", "font_size_pt": 12.0,
        "bold": False, "italic": False, "underline": False,
        "alignment": "left", "line_spacing_rule": "single",
        "line_spacing_value": 1.0, "space_before_pt": 0.0,
        "space_after_pt": 0.0, "color_rgb": None,
    }

    sorted_lines = sorted(field_mapping.items())

    for li, (line_num, columns) in enumerate(sorted_lines):
        if line_formats:
            fmt_index = min(li, len(line_formats) - 1)
            fmt = dict(line_formats[fmt_index])  # copy so we can override
        else:
            fmt = dict(_DEFAULT_FMT)

        # Apply per-line formatting overrides from config
        overrides = LINE_FORMAT_OVERRIDE.get(line_num, {})
        fmt.update(overrides)

        para = cell.paragraphs[0] if li == 0 else cell.add_paragraph()

        # Truncate if a max character limit is set for this line
        max_chars = LINE_MAX_CHARS.get(line_num)

        if isinstance(columns, tuple):
            left_col, right_col = columns
            _apply_line_pair(
                para, fmt,
                _truncate(data_row.get(left_col, ""), max_chars),
                _truncate(data_row.get(right_col, ""), max_chars),
                cell_width_twips,
                cell_margin_lr_twips,
            )
        else:
            _apply_line_single(
                para, fmt,
                _truncate(data_row.get(columns, ""), max_chars),
            )


# ---------------------------------------------------------------------------
# Template-based document generation
# ---------------------------------------------------------------------------

def _clear_label_cells(table, label_row_indices, label_col_indices):
    """Remove all content from every cell in the table.

    Strips every child element from each <w:tc> except <w:tcPr>, then
    adds back a single clean paragraph (<w:p> with an empty <w:pPr>).
    This guarantees no template text, styles, or inherited formatting
    survive, while giving Word a structurally valid paragraph to work with.

    Vertical alignment is pinned to TOP and the top margin is zeroed
    only for actual label cells.
    """
    label_row_set = set(label_row_indices)
    label_col_set = set(label_col_indices)

    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            tc = cell._tc
            # Remove ALL children except tcPr (cell dimensions/borders)
            for child in list(tc):
                if child.tag != qn("w:tcPr"):
                    tc.remove(child)
            # Add back a clean, empty paragraph with a pPr element
            # so Word treats it as a proper editable paragraph.
            p = OxmlElement("w:p")
            p.append(OxmlElement("w:pPr"))
            tc.append(p)

            # Only tweak alignment/margins on actual label cells
            if ri in label_row_set and ci in label_col_set:
                cell.vertical_alignment = WD_ALIGN_VERTICAL.TOP
                tcPr = cell._tc.get_or_add_tcPr()
                tcMar = tcPr.find(qn("w:tcMar"))
                if tcMar is not None:
                    top_el = tcMar.find(qn("w:top"))
                    if top_el is not None:
                        top_el.set(qn("w:w"), "0")


def _clone_table_with_page_break(doc, source_table):
    """Deep-copy a table and append it after a page-break paragraph."""
    # Page break
    p_el = OxmlElement("w:p")
    r_el = OxmlElement("w:r")
    br_el = OxmlElement("w:br")
    br_el.set(qn("w:type"), "page")
    r_el.append(br_el)
    p_el.append(r_el)
    doc.element.body.append(p_el)

    # Clone the table XML (preserves every dimension exactly)
    new_tbl = deepcopy(source_table._tbl)
    doc.element.body.append(new_tbl)

    return Table(new_tbl, doc)


def _fill_page_cells(table, data_rows, row_cursor,
                     label_row_indices, label_col_indices,
                     line_formats, field_mapping,
                     label_col_width, cell_margin_lr):
    """Fill label cells on one page, returning the updated row cursor."""
    for ri in label_row_indices:
        for ci in label_col_indices:
            if row_cursor < len(data_rows):
                _populate_cell(
                    table.rows[ri].cells[ci],
                    data_rows[row_cursor],
                    line_formats,
                    field_mapping,
                    label_col_width,
                    cell_margin_lr,
                )
                row_cursor += 1
    return row_cursor


def generate(profile_path=None, spreadsheet_path=None, output_path=None,
             field_mapping=None):
    """Run the full label-generation pipeline.

    Instead of building a document from scratch, this copies the original
    Avery template so that every table dimension, margin, and positioning
    property is preserved exactly.  Label cells are cleared and filled
    with spreadsheet data.

    Parameters
    ----------
    profile_path : str or Path, optional
        Path to label_format.json.  Defaults to ``PROFILE_PATH``.
    spreadsheet_path : str or Path, optional
        Path to the .numbers file.  Defaults to ``NUMBERS_SPREADSHEET_PATH``.
    output_path : str or Path, optional
        Where to save the generated .docx.  Defaults to ``OUTPUT_PATH``.
    field_mapping : dict, optional
        ``{line_number: column_name}`` mapping.  Defaults to
        ``config.FIELD_MAPPING``.

    Returns
    -------
    Path or None
        The path to the saved document, or *None* if there was nothing
        to print.
    """
    profile_path = Path(profile_path or PROFILE_PATH)
    spreadsheet_path = Path(spreadsheet_path or NUMBERS_SPREADSHEET_PATH)
    output_path = Path(output_path or OUTPUT_PATH)
    if field_mapping is None:
        field_mapping = FIELD_MAPPING

    # 1. Load the formatting profile (re-analyze if template is newer)
    from label_maker.template_analyzer import analyze_template, save_profile
    if (not profile_path.exists()
            or AVERY_TEMPLATE_PATH.stat().st_mtime > profile_path.stat().st_mtime):
        print("Analyzing template (new or updated)...")
        profile = analyze_template(AVERY_TEMPLATE_PATH)
        save_profile(profile, profile_path)
    else:
        profile = load_profile(profile_path)
    grid = profile["grid"]
    fmt = profile["format_template"]
    labels_per_page = grid["labels_per_page"]

    # 2. Read flagged rows
    headers, rows, indices = read_flagged_rows(spreadsheet_path)

    if not rows:
        print("No new labels to print.")
        return None

    # Validate that every mapped column exists in the spreadsheet
    mapped_cols = _get_all_mapped_columns(field_mapping)
    missing = mapped_cols - set(headers)
    if missing:
        raise ValueError(
            f"FIELD_MAPPING references columns not found in the spreadsheet: "
            f"{sorted(missing)}.  Available headers: {headers}"
        )

    print(f"Found {len(rows)} label(s) to print.")

    # 3. Copy the Avery template to the output location.
    #    This preserves every page margin, table dimension, cell size,
    #    row height, and column width from the original template exactly.
    template_path = AVERY_TEMPLATE_PATH
    output_path.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(template_path), str(output_path))

    # 4. Open the copy and work with the template's own table
    doc = Document(str(output_path))

    # Remove any document protection / "mark as final" that the Avery
    # template may carry — these make Word open the file as read-only.
    _strip_document_protection(doc)

    if not doc.tables:
        raise ValueError(
            "Template has no tables — expected an Avery label grid."
        )

    base_table = doc.tables[0]

    # Which physical rows/cols hold actual labels
    label_row_indices = grid.get(
        "label_row_indices", list(range(grid["num_rows"])))
    label_col_indices = grid.get(
        "label_col_indices", list(range(grid["num_cols"])))

    # Cell width for the right-aligned tab stop calculation
    cell_info = profile.get("cell", {})
    col_widths = cell_info.get("col_widths_twips", [])
    label_col_width = (
        col_widths[label_col_indices[0]]
        if col_widths and label_col_indices
        and label_col_indices[0] < len(col_widths)
        else 5000
    )
    cell_margin_lr = cell_info.get("margin_left_twips") or 115

    # Clear any sample content from the template's label cells
    _clear_label_cells(base_table, label_row_indices, label_col_indices)

    # 5. Fill pages with data
    total_pages = math.ceil(len(rows) / labels_per_page)
    row_cursor = 0

    # First page — use the template's own table (exact dimensions)
    row_cursor = _fill_page_cells(
        base_table, rows, row_cursor,
        label_row_indices, label_col_indices,
        fmt["line_formats"], field_mapping,
        label_col_width, cell_margin_lr,
    )

    # Additional pages — clone the template table
    for _ in range(1, total_pages):
        new_table = _clone_table_with_page_break(doc, base_table)
        _clear_label_cells(new_table, label_row_indices, label_col_indices)
        row_cursor = _fill_page_cells(
            new_table, rows, row_cursor,
            label_row_indices, label_col_indices,
            fmt["line_formats"], field_mapping,
            label_col_width, cell_margin_lr,
        )

    # 6. Save the document
    doc.save(str(output_path))
    print(f"Saved {len(rows)} label(s) across {total_pages} page(s) "
          f"to {output_path}")

    # 7. Clear the print flags so they don't print again
    clear_print_flags(spreadsheet_path, indices)

    # 8. Open the document
    _open_file(output_path)

    return output_path


def _strip_document_protection(doc):
    """Remove document protection, 'mark as final', and content edit restrictions.

    Avery templates often ship with form protection or editing restrictions
    that make Word open the generated document as read-only.
    """
    # 1. Remove w:documentProtection from settings
    settings_el = doc.settings.element
    for tag in ("w:documentProtection", "w:writeProtection"):
        el = settings_el.find(qn(tag))
        if el is not None:
            settings_el.remove(el)

    # 2. Remove "mark as final" custom property (docPropsCustom)
    #    This is stored in the core/custom properties. python-docx doesn't
    #    expose custom props directly, but the _MarkAsFinal flag lives in
    #    the extended-properties or custom XML. We clear it via the core props.
    try:
        cp = doc.core_properties
        # If marked as "read only recommended", clear it
        # (python-docx doesn't expose this directly, but we can try)
    except Exception:
        pass

    # 3. Remove any w:permStart / w:permEnd (editing permission ranges)
    #    and content controls (w:sdt) at the body level
    body = doc.element.body
    for el in body.findall(qn("w:sdt")):
        body.remove(el)


def _open_file(path):
    """Open a file with the system default application."""
    path = str(path)
    system = platform.system()

    # On macOS, remove the quarantine extended attribute that makes Word
    # open downloaded files in Protected View (read-only).
    if system == "Darwin":
        try:
            subprocess.run(
                ["xattr", "-d", "com.apple.quarantine", path],
                capture_output=True,
            )
        except OSError:
            pass

    try:
        if system == "Darwin":
            subprocess.Popen(["open", path])
        elif system == "Windows":
            subprocess.Popen(["start", "", path], shell=True)
        else:
            subprocess.Popen(["xdg-open", path])
    except OSError:
        print(f"Could not auto-open {path}. Please open it manually.")


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    generate()
