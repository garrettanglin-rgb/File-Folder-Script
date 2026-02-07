"""Generate a print-ready Avery label document from spreadsheet data.

Reads the formatting profile (label_format.json), pulls flagged rows
from the Numbers spreadsheet, builds a new Word document that replicates
the exact Avery grid layout and formatting, and opens the result.
"""

import math
import platform
import subprocess
import sys
from pathlib import Path

from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, Twips, RGBColor

from label_maker.config import (
    AVERY_TEMPLATE_PATH,
    FIELD_MAPPING,
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
# Low-level XML helpers (python-docx doesn't expose every property)
# ---------------------------------------------------------------------------

def _set_row_height(row, height_twips):
    """Set an exact row height via the underlying XML."""
    trPr = row._tr.get_or_add_trPr()
    trHeight = trPr.find(qn("w:trHeight"))
    if trHeight is None:
        trHeight = OxmlElement("w:trHeight")
        trPr.append(trHeight)
    trHeight.set(qn("w:val"), str(height_twips))
    trHeight.set(qn("w:hRule"), "exact")


def _set_cell_margins(cell, top=0, bottom=0, left=0, right=0):
    """Set per-cell margins (in twips) via tcMar XML."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = tcPr.find(qn("w:tcMar"))
    if tcMar is None:
        tcMar = OxmlElement("w:tcMar")
        tcPr.append(tcMar)
    for side, val in [("top", top), ("bottom", bottom),
                      ("left", left), ("right", right)]:
        el = tcMar.find(qn(f"w:{side}"))
        if el is None:
            el = OxmlElement(f"w:{side}")
            tcMar.append(el)
        el.set(qn("w:w"), str(val))
        el.set(qn("w:type"), "dxa")


def _set_cell_width(cell, width_twips):
    """Set the preferred cell width in twips."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcW = tcPr.find(qn("w:tcW"))
    if tcW is None:
        tcW = OxmlElement("w:tcW")
        tcPr.append(tcW)
    tcW.set(qn("w:w"), str(width_twips))
    tcW.set(qn("w:type"), "dxa")


def _hide_table_borders(table):
    """Remove all borders from the table (invisible grid)."""
    tblPr = table._tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tblPr)
    borders = tblPr.find(qn("w:tblBorders"))
    if borders is None:
        borders = OxmlElement("w:tblBorders")
        tblPr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = borders.find(qn(f"w:{edge}"))
        if el is None:
            el = OxmlElement(f"w:{edge}")
            borders.append(el)
        el.set(qn("w:val"), "none")
        el.set(qn("w:sz"), "0")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "auto")


def _set_fixed_layout(table):
    """Force fixed table layout so column widths are respected."""
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")


def _center_table(table):
    """Centre the table horizontally on the page."""
    tblPr = table._tbl.tblPr
    jc = tblPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        tblPr.append(jc)
    jc.set(qn("w:val"), "center")


def _set_grid_cols(table, col_widths_twips):
    """Write explicit <w:tblGrid><w:gridCol> entries."""
    tblGrid = table._tbl.find(qn("w:tblGrid"))
    if tblGrid is None:
        tblGrid = OxmlElement("w:tblGrid")
        table._tbl.insert(1, tblGrid)
    # Clear existing gridCols
    for gc in list(tblGrid.findall(qn("w:gridCol"))):
        tblGrid.remove(gc)
    for w in col_widths_twips:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        tblGrid.append(gc)


# ---------------------------------------------------------------------------
# Core document-building logic
# ---------------------------------------------------------------------------

def _set_paragraph_spacing(paragraph, line_fmt):
    """Apply paragraph-level spacing properties from a line format dict."""
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


def _format_run(run, line_fmt):
    """Apply run-level font properties from a line format dict."""
    font = run.font
    if line_fmt.get("font_name"):
        font.name = line_fmt["font_name"]
    if line_fmt.get("font_size_pt"):
        font.size = Pt(line_fmt["font_size_pt"])
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
            fmt = line_formats[fmt_index]
        else:
            fmt = _DEFAULT_FMT

        para = cell.paragraphs[0] if li == 0 else cell.add_paragraph()

        if isinstance(columns, tuple):
            left_col, right_col = columns
            _apply_line_pair(
                para, fmt,
                data_row.get(left_col, ""),
                data_row.get(right_col, ""),
                cell_width_twips,
                cell_margin_lr_twips,
            )
        else:
            _apply_line_single(para, fmt, data_row.get(columns, ""))


def _build_page_table(doc, profile, section):
    """Create and return a single-page label table matching the profile.

    Replicates the exact physical table from the template, including any
    gutter columns and spacer rows, so the output matches the original
    Avery sheet layout precisely.
    """
    grid = profile["grid"]
    num_rows = grid["num_rows"]       # physical rows (incl. spacers)
    num_cols = grid["num_cols"]       # physical cols (incl. gutters)
    cell_info = profile.get("cell", {})
    page = profile["page"]

    # --- Per-column widths from template XML, or uniform fallback ---
    col_widths = cell_info.get("col_widths_twips", [])
    if not col_widths or len(col_widths) != num_cols:
        avail_w = int((page["page_width_in"] - page["margin_left_in"]
                       - page["margin_right_in"]) * 1440)
        col_widths = [avail_w // num_cols] * num_cols

    # --- Per-row heights from template XML, or uniform fallback ---
    all_row_heights = cell_info.get("all_row_heights_twips", [])
    if not all_row_heights or len(all_row_heights) != num_rows:
        avail_h = int((page["page_height_in"] - page["margin_top_in"]
                       - page["margin_bottom_in"]) * 1440)
        all_row_heights = [avail_h // num_rows] * num_rows

    # --- Create the physical table ---
    table = doc.add_table(rows=num_rows, cols=num_cols)
    _hide_table_borders(table)
    _set_fixed_layout(table)
    _center_table(table)
    _set_grid_cols(table, col_widths)

    # --- Cell margins ---
    margin_l = cell_info.get("margin_left_twips") or 115
    margin_r = cell_info.get("margin_right_twips") or 115
    margin_t = cell_info.get("margin_top_twips") or 0
    margin_b = cell_info.get("margin_bottom_twips") or 0
    cell_margins = {
        "top": margin_t, "bottom": margin_b,
        "left": margin_l, "right": margin_r,
    }

    label_col_set = set(grid.get("label_col_indices", list(range(num_cols))))
    label_row_set = set(grid.get("label_row_indices", list(range(num_rows))))

    for ri, row in enumerate(table.rows):
        h = all_row_heights[ri] if ri < len(all_row_heights) else None
        if h is not None:
            _set_row_height(row, h)
        for ci, cell_obj in enumerate(row.cells):
            w = col_widths[ci] if ci < len(col_widths) else col_widths[0]
            _set_cell_width(cell_obj, w)
            # Only apply label formatting to actual label cells
            if ri in label_row_set and ci in label_col_set:
                cell_obj.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                _set_cell_margins(cell_obj, **cell_margins)

    # The label column width (for tab-stop calculation)
    label_col_indices = grid.get("label_col_indices", list(range(num_cols)))
    label_col_width = (
        col_widths[label_col_indices[0]]
        if label_col_indices and label_col_indices[0] < len(col_widths)
        else col_widths[0]
    )

    return table, label_col_width, margin_l


def generate(profile_path=None, spreadsheet_path=None, output_path=None,
             field_mapping=None):
    """Run the full label-generation pipeline.

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

    # 1. Load the formatting profile
    profile = load_profile(profile_path)
    page = profile["page"]
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

    # 3. Build the Word document
    doc = Document()

    # Page setup matching the template
    section = doc.sections[0]
    section.page_width = Inches(page["page_width_in"])
    section.page_height = Inches(page["page_height_in"])
    section.top_margin = Inches(page["margin_top_in"])
    section.bottom_margin = Inches(page["margin_bottom_in"])
    section.left_margin = Inches(page["margin_left_in"])
    section.right_margin = Inches(page["margin_right_in"])

    # Remove the empty default paragraph
    if doc.paragraphs:
        p = doc.paragraphs[0]._element
        p.getparent().remove(p)

    # Which physical rows/cols hold actual labels
    label_row_indices = grid.get(
        "label_row_indices", list(range(grid["num_rows"])))
    label_col_indices = grid.get(
        "label_col_indices", list(range(grid["num_cols"])))

    # Determine how many pages we need
    total_pages = math.ceil(len(rows) / labels_per_page)

    row_cursor = 0
    for page_num in range(total_pages):
        if page_num > 0:
            # Add a section break for a new page
            new_section = doc.add_section()
            new_section.page_width = section.page_width
            new_section.page_height = section.page_height
            new_section.top_margin = section.top_margin
            new_section.bottom_margin = section.bottom_margin
            new_section.left_margin = section.left_margin
            new_section.right_margin = section.right_margin
            current_section = new_section
        else:
            current_section = section

        table, label_col_width, cell_margin_lr = _build_page_table(
            doc, profile, current_section)

        # Fill only the label cells (skip gutter cols & spacer rows)
        for ri in label_row_indices:
            for ci in label_col_indices:
                if row_cursor < len(rows):
                    _populate_cell(
                        table.rows[ri].cells[ci],
                        rows[row_cursor],
                        fmt["line_formats"],
                        field_mapping,
                        label_col_width,
                        cell_margin_lr,
                    )
                    row_cursor += 1
                # else: cell stays empty (blank label)

    # 4. Save the document
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    print(f"Saved {len(rows)} label(s) across {total_pages} page(s) "
          f"to {output_path}")

    # 5. Clear the print flags so they don't print again
    clear_print_flags(spreadsheet_path, indices)

    # 6. Open the document
    _open_file(output_path)

    return output_path


def _open_file(path):
    """Open a file with the system default application."""
    path = str(path)
    system = platform.system()
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
