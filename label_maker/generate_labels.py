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

def _apply_line_format(paragraph, line_fmt, text):
    """Apply a line_format dict to a paragraph, inserting *text*."""
    paragraph.alignment = _ALIGN_MAP.get(line_fmt["alignment"],
                                         WD_ALIGN_PARAGRAPH.LEFT)
    pf = paragraph.paragraph_format

    # Line spacing
    rule = _SPACING_RULE_MAP.get(line_fmt["line_spacing_rule"],
                                 WD_LINE_SPACING.SINGLE)
    pf.line_spacing_rule = rule
    if rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
        pf.line_spacing = Pt(line_fmt["line_spacing_value"])
    else:
        pf.line_spacing = line_fmt["line_spacing_value"]

    # Paragraph spacing
    sb = line_fmt.get("space_before_pt")
    sa = line_fmt.get("space_after_pt")
    pf.space_before = Pt(sb) if sb else Pt(0)
    pf.space_after = Pt(sa) if sa else Pt(0)

    # Run-level formatting
    run = paragraph.add_run(str(text) if text is not None else "")
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


def _populate_cell(cell, data_row, line_formats, field_mapping):
    """Fill a single table cell with label data using the format profile.

    *data_row* is a dict keyed by header name.  *field_mapping* is an
    ordered dict mapping 1-based line numbers to spreadsheet column
    names (from ``config.FIELD_MAPPING``).  *line_formats* comes from
    the profile's ``format_template.line_formats``.

    Each mapped line gets the formatting of the corresponding template
    line.  If the template has fewer line_formats than the mapping, the
    last available format is reused.
    """
    sorted_lines = sorted(field_mapping.items())  # [(1, "File Number"), ...]

    for li, (line_num, column_name) in enumerate(sorted_lines):
        text = data_row.get(column_name, "")
        # Pick the matching template line format; fall back to the last
        # one if the mapping has more lines than the template defined.
        fmt_index = min(li, len(line_formats) - 1)
        fmt = line_formats[fmt_index]

        if li == 0:
            para = cell.paragraphs[0]
        else:
            para = cell.add_paragraph()
        _apply_line_format(para, fmt, text)


def _build_page_table(doc, profile, section):
    """Create and return a single-page label table matching the profile."""
    grid = profile["grid"]
    num_rows = grid["num_rows"]
    num_cols = grid["num_cols"]

    table = doc.add_table(rows=num_rows, cols=num_cols)

    # Table-level properties
    _hide_table_borders(table)
    _set_fixed_layout(table)
    _center_table(table)

    # Column widths from the original template (twips / "dxa")
    col_width = 5126  # from template gridCol
    _set_grid_cols(table, [col_width] * num_cols)

    # Row heights + cell properties
    row_height = 1915  # from template trHeight
    cell_margins = {"top": 86, "bottom": 86, "left": 144, "right": 144}

    for ri, row in enumerate(table.rows):
        _set_row_height(row, row_height)
        for ci, cell in enumerate(row.cells):
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            _set_cell_margins(cell, **cell_margins)
            _set_cell_width(cell, col_width)

    return table


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
    mapped_cols = set(field_mapping.values())
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

        table = _build_page_table(doc, profile, current_section)

        # Fill cells with data
        for ri in range(grid["num_rows"]):
            for ci in range(grid["num_cols"]):
                if row_cursor < len(rows):
                    _populate_cell(
                        table.rows[ri].cells[ci],
                        rows[row_cursor],
                        fmt["line_formats"],
                        field_mapping,
                    )
                    row_cursor += 1
                # else: cell stays empty (blank label)

    # 4. Save the document
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(output_path))
    print(f"Saved {len(rows)} label(s) across {total_pages} page(s) to {output_path}")

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
