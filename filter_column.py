"""
filter_columns.py

Takes two Excel files:
  1. A "data" file (file1)
  2. A "masterfile" (file2)

Creates a new Excel file containing only the columns from file1 whose
column headers also appear in file2's masterfile (column names are matched
case-insensitively and with leading/trailing whitespace stripped).

Usage:
    python filter_columns.py --data file1.xlsx --master file2.xlsx --output result.xlsx

Optional flags:
    --data-sheet      Sheet name/index to read from file1 (default: first sheet)
    --master-sheet    Sheet name/index to read from file2 (default: first sheet)
    --case-sensitive  Match column names exactly (case-sensitive, no trimming)
"""

import argparse
import re
import sys
import pandas as pd


def normalize(col_name, case_sensitive=False):
    name = str(col_name)
    # Replace non-breaking spaces / tabs / newlines with a regular space
    name = name.replace("\xa0", " ").replace("\t", " ").replace("\n", " ")
    # Collapse multiple internal spaces into one, and strip leading/trailing
    name = re.sub(r"\s+", " ", name).strip()
    if not case_sensitive:
        name = name.lower()
    return name


def main():
    parser = argparse.ArgumentParser(description="Filter file1 columns to only those present in masterfile.")
    parser.add_argument("--data", required=True, help="Path to file1 (the data file to filter)")
    parser.add_argument("--master", required=True, help="Path to file2 (the masterfile)")
    parser.add_argument("--output", required=True, help="Path to write the resulting Excel file")
    parser.add_argument("--data-sheet", default=0, help="Sheet name or index in the data file (default: first sheet)")
    parser.add_argument("--master-sheet", default=0, help="Sheet name or index in the masterfile (default: first sheet)")
    parser.add_argument("--case-sensitive", action="store_true", help="Match column names exactly, case-sensitive")
    args = parser.parse_args()

    # Allow numeric sheet args to be passed as ints
    def parse_sheet(val):
        try:
            return int(val)
        except (ValueError, TypeError):
            return val

    data_sheet = parse_sheet(args.data_sheet)
    master_sheet = parse_sheet(args.master_sheet)

    try:
        df_data = pd.read_excel(args.data, sheet_name=data_sheet)
    except Exception as e:
        print(f"Error reading data file '{args.data}': {e}", file=sys.stderr)
        sys.exit(1)

    try:
        df_master = pd.read_excel(args.master, sheet_name=master_sheet)
    except Exception as e:
        print(f"Error reading masterfile '{args.master}': {e}", file=sys.stderr)
        sys.exit(1)

    master_cols_normalized = {
        normalize(c, args.case_sensitive) for c in df_master.columns
    }

    matched_columns = [
        c for c in df_data.columns
        if normalize(c, args.case_sensitive) in master_cols_normalized
    ]

    if not matched_columns:
        print("Warning: No matching columns found between the data file and the masterfile.", file=sys.stderr)

    result_df = df_data[matched_columns]

    try:
        result_df.to_excel(args.output, index=False)
    except Exception as e:
        print(f"Error writing output file '{args.output}': {e}", file=sys.stderr)
        sys.exit(1)

    print(f"Done. {len(matched_columns)} matching column(s) written to '{args.output}'.")
    print("Matched columns:", matched_columns)

    unmatched = [c for c in df_data.columns if c not in matched_columns]
    if unmatched:
        print("Columns in data file NOT found in masterfile (excluded):", unmatched)


if __name__ == "__main__":
    main()
