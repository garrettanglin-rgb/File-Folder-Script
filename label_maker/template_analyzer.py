"""Analyze an Avery label Word document and extract its formatting profile."""

import json
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Emu


# Map python-docx alignment enums to readable strings
_ALIGNMENT_MAP = {
    WD_ALIGN_PARAGRAPH.LEFT: "left",
    WD_ALIGN_PARAGRAPH.CENTER: "center",
    WD_ALIGN_PARAGRAPH.RIGHT: "right",
    WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
    None: "left",  # Word default when unset
}

# Map line-spacing rule enums to readable strings
_SPACING_RULE_MAP = {
    WD_LINE_SPACING.SINGLE: "single",
    WD_LINE_SPACING.ONE_POINT_FIVE: "1.5",
    WD_LINE_SPACING.DOUBLE: "double",
    WD_LINE_SPACING.EXACTLY: "exactly",
    WD_LINE_SPACING.AT_LEAST: "at_least",
    WD_LINE_SPACING.MULTIPLE: "multiple",
    None: "single",
}


def _emu_to_pt(emu_val):
    """Convert EMU (English Metric Units) to points. Returns None if input is None."""
    if emu_val is None:
        return None
    return round(emu_val / 12700, 1)


def _emu_to_inches(emu_val):
    """Convert EMU to inches. Returns None if input is None."""
    if emu_val is None:
        return None
    return round(emu_val / 914400, 4)


def _analyze_run(run):
    """Extract formatting from a single text run."""
    font = run.font
    return {
        "text": run.text,
        "font_name": font.name,
        "font_size_pt": _emu_to_pt(font.size),
        "bold": bool(font.bold) if font.bold is not None else False,
        "italic": bool(font.italic) if font.italic is not None else False,
        "underline": bool(font.underline) if font.underline is not None else False,
        "color_rgb": str(font.color.rgb) if font.color and font.color.rgb else None,
    }


def _analyze_paragraph(para):
    """Extract formatting from a single paragraph (one text line in a label)."""
    pf = para.paragraph_format
    alignment = _ALIGNMENT_MAP.get(para.alignment, "left")

    # Determine line spacing value and rule
    spacing_rule = _SPACING_RULE_MAP.get(pf.line_spacing_rule, "single")
    if pf.line_spacing is None:
        line_spacing_value = 1.0
    elif pf.line_spacing_rule in (WD_LINE_SPACING.EXACTLY, WD_LINE_SPACING.AT_LEAST):
        # Value is in EMU — convert to points
        line_spacing_value = _emu_to_pt(pf.line_spacing)
    else:
        # Value is a float multiplier (e.g. 1.0, 1.5, 2.0)
        line_spacing_value = float(pf.line_spacing)

    # Consolidate run-level formatting. A paragraph can have multiple runs
    # if formatting changes mid-line, but for label templates the typical
    # case is one run per paragraph. We capture every run individually so
    # nothing is lost.
    runs = [_analyze_run(r) for r in para.runs]

    # Derive dominant / first-run formatting for convenience
    first = runs[0] if runs else {}

    return {
        "text": para.text,
        "alignment": alignment,
        "line_spacing_rule": spacing_rule,
        "line_spacing_value": line_spacing_value,
        "space_before_pt": _emu_to_pt(pf.space_before),
        "space_after_pt": _emu_to_pt(pf.space_after),
        "font_name": first.get("font_name"),
        "font_size_pt": first.get("font_size_pt"),
        "bold": first.get("bold", False),
        "italic": first.get("italic", False),
        "underline": first.get("underline", False),
        "color_rgb": first.get("color_rgb"),
        "runs": runs,
    }


def _analyze_cell(cell):
    """Analyze all paragraphs (lines) inside a single label cell."""
    paragraphs = cell.paragraphs
    # Skip cells that are completely empty (no visible text)
    lines = [_analyze_paragraph(p) for p in paragraphs if p.text.strip()]
    return {
        "num_lines": len(lines),
        "lines": lines,
    }


def analyze_template(docx_path):
    """Open an Avery label .docx and return a complete formatting profile.

    Parameters
    ----------
    docx_path : str or Path
        Path to the Word document to analyze.

    Returns
    -------
    dict
        A dictionary describing the page layout, grid structure, and
        per-line formatting of every label in the template.
    """
    docx_path = Path(docx_path)
    doc = Document(str(docx_path))

    # --- Page / section info ---
    section = doc.sections[0]
    page_info = {
        "page_width_in": _emu_to_inches(section.page_width),
        "page_height_in": _emu_to_inches(section.page_height),
        "margin_top_in": _emu_to_inches(section.top_margin),
        "margin_bottom_in": _emu_to_inches(section.bottom_margin),
        "margin_left_in": _emu_to_inches(section.left_margin),
        "margin_right_in": _emu_to_inches(section.right_margin),
    }

    # --- Table / grid detection ---
    if not doc.tables:
        raise ValueError("No tables found in the document. Avery label templates use a table grid.")

    table = doc.tables[0]
    num_rows = len(table.rows)
    num_cols = len(table.columns)
    labels_per_page = num_rows * num_cols

    grid_info = {
        "num_rows": num_rows,
        "num_cols": num_cols,
        "labels_per_page": labels_per_page,
    }

    # --- Per-label analysis ---
    labels = []
    for ri, row in enumerate(table.rows):
        for ci, cell in enumerate(row.cells):
            label_data = _analyze_cell(cell)
            label_data["row"] = ri
            label_data["col"] = ci
            labels.append(label_data)

    # --- Build a canonical "format template" from the first label ---
    # This gives consumers a quick reference for the expected line formatting
    # without iterating every label.
    first_label = labels[0] if labels else {}
    line_formats = []
    for line in first_label.get("lines", []):
        line_formats.append({
            "font_name": line["font_name"],
            "font_size_pt": line["font_size_pt"],
            "bold": line["bold"],
            "italic": line["italic"],
            "underline": line["underline"],
            "alignment": line["alignment"],
            "line_spacing_rule": line["line_spacing_rule"],
            "line_spacing_value": line["line_spacing_value"],
            "space_before_pt": line["space_before_pt"],
            "space_after_pt": line["space_after_pt"],
            "color_rgb": line["color_rgb"],
        })

    format_template = {
        "num_lines_per_label": first_label.get("num_lines", 0),
        "line_formats": line_formats,
    }

    profile = {
        "source_file": str(docx_path.resolve()),
        "page": page_info,
        "grid": grid_info,
        "format_template": format_template,
        "labels": labels,
    }

    return profile


def save_profile(profile, output_path):
    """Write a formatting profile to a JSON file.

    Parameters
    ----------
    profile : dict
        The profile dictionary returned by ``analyze_template``.
    output_path : str or Path
        Destination path for the JSON file.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with open(output_path, "w") as f:
        json.dump(profile, f, indent=2)
    print(f"Formatting profile saved to {output_path}")


def load_profile(json_path):
    """Load a previously saved formatting profile from JSON.

    Parameters
    ----------
    json_path : str or Path
        Path to the JSON profile file.

    Returns
    -------
    dict
        The formatting profile dictionary.
    """
    with open(json_path) as f:
        return json.load(f)


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    from label_maker.config import AVERY_TEMPLATE_PATH

    docx_file = Path(AVERY_TEMPLATE_PATH)
    output_file = docx_file.parent / "label_format.json"

    print(f"Analyzing template: {docx_file}")
    profile = analyze_template(docx_file)

    print(f"\nPage: {profile['page']['page_width_in']}\" x {profile['page']['page_height_in']}\"")
    print(f"Grid: {profile['grid']['num_rows']} rows x {profile['grid']['num_cols']} cols "
          f"= {profile['grid']['labels_per_page']} labels/page")
    print(f"Lines per label: {profile['format_template']['num_lines_per_label']}")

    for i, lf in enumerate(profile["format_template"]["line_formats"]):
        style_parts = []
        if lf["bold"]:
            style_parts.append("Bold")
        if lf["italic"]:
            style_parts.append("Italic")
        if lf["underline"]:
            style_parts.append("Underline")
        style_str = "+".join(style_parts) if style_parts else "Regular"
        print(f"  Line {i + 1}: {lf['font_name']} {lf['font_size_pt']}pt, "
              f"{style_str}, {lf['alignment']}, spacing={lf['line_spacing_value']}")

    save_profile(profile, output_file)
