"""
Tensile Test Mechanical Properties Analyzer

Scans a folder for *-*.is_tens_Exports directories, reads CSV files inside,
extracts:
  - Max tensile stress (拉伸应力) -> Operation X
  - Elongation at break point (拉伸应变) -> Operation V

Writes results to 力学性能.xlsx organized by sample type (column) and
sample number (row).
"""

import sys
import os
import re
import csv
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


def find_csv_header_index(headers, target):
    """Find the 1-based column index where `target` is a substring of the header."""
    for idx, h in enumerate(headers, start=1):
        if target in h:
            return idx
    return None


def _open_csv(filepath):
    """Try to open a CSV file, falling back through common Chinese encodings."""
    for enc in ("utf-8-sig", "gbk", "gb2312", "gb18030", "utf-8"):
        try:
            with open(filepath, "r", encoding=enc) as f:
                csv.reader(f).__next__()  # Test reading row 1
            return enc
        except (UnicodeDecodeError, StopIteration):
            continue
    return "utf-8-sig"  # Last resort


def read_csv_values(filepath, col_idx, start_row=3):
    """Read numeric values from a CSV column, starting at `start_row` (1-based)."""
    enc = _open_csv(filepath)
    values = []
    with open(filepath, "r", encoding=enc) as f:
        reader = csv.reader(f)
        for row_num, row in enumerate(reader, start=1):
            if row_num < start_row:
                continue
            if col_idx <= len(row):
                val = row[col_idx - 1].strip().strip('"')
                try:
                    values.append(float(val))
                except (ValueError, TypeError):
                    continue
    return values


def get_csv_headers(filepath):
    """Return the header row (row 1) of a CSV file."""
    enc = _open_csv(filepath)
    with open(filepath, "r", encoding=enc) as f:
        reader = csv.reader(f)
        return next(reader)


def operation_x(filepath):
    """Max tensile stress (拉伸应力) from data starting at row 3."""
    headers = get_csv_headers(filepath)
    col = find_csv_header_index(headers, "拉伸应力")
    if col is None:
        print(f"  WARNING: '拉伸应力' column not found")
        return None
    values = read_csv_values(filepath, col)
    if not values:
        print(f"  WARNING: No numeric data in '拉伸应力' column")
        return None
    return max(values)


def operation_v(filepath):
    """
    Detect elongation at break from the strain column:
    Find the first row i where strain[i+1] - strain[i] is negative AND
    represents a rapid decrease compared to the preceding trend.

    Uses a rolling window: if the current diff is negative and falls more
    than K standard deviations below the rolling mean, it's flagged as a
    rapid anomalous decrease (indicating fracture).
    """
    WINDOW = 20   # Number of preceding diffs to establish trend
    K = 4.0       # Sigma threshold for "rapid decrease"

    headers = get_csv_headers(filepath)
    col = find_csv_header_index(headers, "拉伸应变")
    if col is None:
        print(f"  WARNING: '拉伸应变' column not found")
        return None
    values = read_csv_values(filepath, col)
    if len(values) < WINDOW + 2:
        print(f"  WARNING: Not enough strain data points")
        return None

    # Compute consecutive differences
    diffs = [values[i + 1] - values[i] for i in range(len(values) - 1)]

    import statistics
    for i in range(WINDOW, len(diffs)):
        window = diffs[i - WINDOW:i]
        mean = statistics.mean(window)
        stdev = statistics.stdev(window) if len(window) >= 2 else 0.0001
        threshold = mean - K * stdev

        if diffs[i] < 0 and diffs[i] < threshold:
            print(f"  [V] Trigger at data row {i + 3}: "
                  f"diff={diffs[i]:.6f}, window_mean={mean:.6f}, "
                  f"threshold={threshold:.6f}")
            return values[i]  # strain at row i, before the drop

    # Fallback: if no negative anomalous diff, report the final strain value
    # (test ended without sharp strain reversal)
    print(f"  [V] No anomalous strain drop detected; using final strain value")
    return values[-1]


def parse_dirname(dirname):
    """
    Parse 'a-b' from directory names like '1-2.is_tens_Exports'
    or '1_2.is_tens_Exports'. Returns (a, b) or None.
    """
    # Strip the .is_tens_Exports suffix to isolate a-b
    name = dirname.replace(".is_tens_Exports", "").replace(".is_tens_Exports", "")
    match = re.match(r"^(\d+)[-_](\d+)$", name)
    if match:
        return int(match.group(1)), int(match.group(2))
    return None


def scan_dirs(folder):
    """
    Find all directories matching '*-*.is_tens_Exports' or
    '*_*.is_tens_Exports', each containing at least one CSV.
    Returns list of (dir_path, csv_path).
    """
    folder_path = Path(folder)
    if not folder_path.is_dir():
        print(f"ERROR: '{folder}' is not a valid directory.")
        sys.exit(1)

    results = []
    for entry in sorted(folder_path.iterdir()):
        if not entry.is_dir():
            continue
        if ".is_tens_Exports" not in entry.name:
            continue
        csv_files = sorted(entry.glob("*.csv"))
        if not csv_files:
            print(f"SKIP: '{entry.name}' — no CSV files inside")
            continue
        results.append((entry, csv_files[0]))
    return results


def main():
    if len(sys.argv) < 2:
        print("Usage: python analyze_tensile.py <folder_path> [output_path]")
        sys.exit(1)

    folder = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else "力学性能.xlsx"

    entries = scan_dirs(folder)
    if not entries:
        print(f"No .is_tens_Exports directories with CSV files found in '{folder}'.")
        sys.exit(0)

    print(f"Found {len(entries)} directories to process.\n")

    out_wb = Workbook()
    out_ws = out_wb.active
    out_ws.title = "力学性能"

    processed = 0
    skipped = 0

    for dir_path, csv_path in entries:
        parsed = parse_dirname(dir_path.name)
        if parsed is None:
            print(f"SKIP: '{dir_path.name}' — cannot parse a-b from name")
            skipped += 1
            continue

        a, b = parsed
        print(f"Processing: {dir_path.name}/  ->  {csv_path.name}  (type={a}, num={b})")

        # Operation X: max stress -> column a, row b
        result_x = operation_x(str(csv_path))
        if result_x is not None:
            col_letter = get_column_letter(a)
            cell_ref = f"{col_letter}{b}"
            out_ws[cell_ref] = round(result_x, 2)
            print(f"  [X] 抗拉强度 (max 拉伸应力): {result_x:.2f} -> {cell_ref}")

        # Operation V: elongation -> column a, row b+10
        result_v = operation_v(str(csv_path))
        if result_v is not None:
            col_letter = get_column_letter(a)
            cell_ref = f"{col_letter}{b + 10}"
            out_ws[cell_ref] = round(result_v, 2)
            print(f"  [V] 断裂伸长率 (strain[i]): {result_v:.2f} -> {cell_ref}")

        processed += 1
        print()

    out_wb.save(output_path)
    print(f"Done. {processed} processed, {skipped} skipped.")
    print(f"Output: {os.path.abspath(output_path)}")


if __name__ == "__main__":
    main()
