"""Read label data from an Apple Numbers spreadsheet.

Opens a .numbers file, dynamically detects column headers, and returns
only the rows flagged for printing (based on a "Print" or "New" column).
After labels are generated, ``clear_print_flags`` writes the flags out
so the same rows are not printed again.
"""

from pathlib import Path

from numbers_parser import Document


# Values in the flag column that mean "print this row"
_PRINT_FLAGS = {"y", "x", "yes"}

# Column names recognised as the print-flag column (case-insensitive)
_FLAG_COLUMN_NAMES = {"print", "new"}


def _find_flag_column(headers):
    """Return the index of the print-flag column, or raise if not found."""
    for i, name in enumerate(headers):
        if name is not None and name.strip().lower() in _FLAG_COLUMN_NAMES:
            return i
    raise ValueError(
        f"No print-flag column found. Expected a column named one of "
        f"{sorted(_FLAG_COLUMN_NAMES)} (case-insensitive). "
        f"Found headers: {headers}"
    )


def read_headers(spreadsheet_path, sheet_index=0, table_index=0):
    """Return the list of column headers from the first row.

    Only columns that actually contain a header value are returned
    (trailing empty columns from the default Numbers grid are excluded).

    Parameters
    ----------
    spreadsheet_path : str or Path
        Path to the .numbers file.
    sheet_index : int
        Sheet index (default 0).
    table_index : int
        Table index within the sheet (default 0).

    Returns
    -------
    list[str]
        Column header names.
    """
    doc = Document(str(spreadsheet_path))
    table = doc.sheets[sheet_index].tables[table_index]
    raw = [table.cell(0, c).value for c in range(table.num_cols)]
    # Trim trailing None columns
    while raw and raw[-1] is None:
        raw.pop()
    return raw


def read_flagged_rows(spreadsheet_path, sheet_index=0, table_index=0):
    """Read the spreadsheet and return only the rows flagged for printing.

    The function dynamically reads headers from row 0, locates the
    print-flag column ("Print" or "New"), and filters data rows whose
    flag value is in ``_PRINT_FLAGS``.

    Parameters
    ----------
    spreadsheet_path : str or Path
        Path to the .numbers file.
    sheet_index : int
        Sheet index (default 0).
    table_index : int
        Table index within the sheet (default 0).

    Returns
    -------
    tuple[list[str], list[dict], list[int]]
        A 3-tuple of:
        - **headers** – column names (excluding trailing empties).
        - **rows** – list of dicts keyed by header name for each
          flagged row.
        - **flagged_row_indices** – the zero-based *data* row indices
          (i.e. row 1 in the spreadsheet = index 0 here) so they can
          be passed to ``clear_print_flags`` later.
    """
    doc = Document(str(spreadsheet_path))
    table = doc.sheets[sheet_index].tables[table_index]

    # --- Discover headers (first row) ---
    all_headers = [table.cell(0, c).value for c in range(table.num_cols)]
    # Trim trailing empty columns
    while all_headers and all_headers[-1] is None:
        all_headers.pop()
    num_used_cols = len(all_headers)

    flag_col = _find_flag_column(all_headers)

    # --- Iterate data rows (skip header row 0) ---
    flagged_rows = []
    flagged_indices = []

    for ri in range(1, table.num_rows):
        flag_val = table.cell(ri, flag_col).value
        if flag_val is not None and str(flag_val).strip().lower() in _PRINT_FLAGS:
            row_dict = {}
            for ci in range(num_used_cols):
                row_dict[all_headers[ci]] = table.cell(ri, ci).value
            flagged_rows.append(row_dict)
            flagged_indices.append(ri)

    return all_headers, flagged_rows, flagged_indices


def clear_print_flags(spreadsheet_path, flagged_row_indices,
                      sheet_index=0, table_index=0):
    """Clear the print-flag column for the given rows and save the file.

    After labels have been generated successfully, call this to blank
    out the flag so those rows won't be picked up on the next run.

    Parameters
    ----------
    spreadsheet_path : str or Path
        Path to the .numbers file (will be overwritten).
    flagged_row_indices : list[int]
        Row indices as returned by ``read_flagged_rows``.
    sheet_index : int
        Sheet index (default 0).
    table_index : int
        Table index within the sheet (default 0).
    """
    spreadsheet_path = Path(spreadsheet_path)
    doc = Document(str(spreadsheet_path))
    table = doc.sheets[sheet_index].tables[table_index]

    # Re-detect the flag column (file may have been edited externally)
    all_headers = [table.cell(0, c).value for c in range(table.num_cols)]
    while all_headers and all_headers[-1] is None:
        all_headers.pop()
    flag_col = _find_flag_column(all_headers)

    for ri in flagged_row_indices:
        table.write(ri, flag_col, "")

    doc.save(str(spreadsheet_path))
    print(f"Cleared print flags for {len(flagged_row_indices)} row(s) in {spreadsheet_path.name}")


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    from label_maker.config import NUMBERS_SPREADSHEET_PATH

    path = Path(NUMBERS_SPREADSHEET_PATH)
    print(f"Reading spreadsheet: {path}\n")

    headers, rows, indices = read_flagged_rows(path)
    print(f"Headers: {headers}")
    print(f"Flagged rows: {len(rows)} of {len(rows) + len(indices)}")  # rough total
    print()

    for i, row in enumerate(rows):
        print(f"  [{indices[i]}] {row}")

    print(f"\nSpreadsheet row indices to clear: {indices}")
