EXCEL COLUMN FILTER

A small Python utility that compares the column headers of an Excel file against a "masterfile" of approved column names, and 
produces a new Excel file containing only the matching columns (with their data intact).

Column name matching is case-insensitive and ignores leading/trailing/extra whitespace, so headers like "VENDOR NAME " 
and "vendor name" are treated as the same column.

WHY ? 

Useful when you regularly receive Excel files with inconsistent or extra columns, and want to standardize them down to a known, 
approved set of fields defined in a separate reference (master) file.

HOW IT WORKS ? 

1. Reads all columns from your data file (--data).
2. Reads all columns from your masterfile (--master).
3. Normalizes each column name (lowercases, strips whitespace, collapses internal spaces) before comparing.
4. Keeps only the columns from the data file whose normalized name also appears in the masterfile.
5. Writes those columns (with their original data) to a new Excel file (--output).


Columns present only in the masterfile, or only in the data file, are simply excluded from the output — no error is raised, 
but a summary is printed to the console.


INSTALLATION

1. bashgit clone https://github.com/<your-username>/<your-repo>.git
2. cd <your-repo>
3. pip install -r requirements.txt

USAGE

bashpython filter_columns.py --data file1.xlsx --master masterfile.xlsx --output result.xlsx

NOTE : Here file1.xlsx is the orignal excel file you need to make changes in.
            masterfile.xlsx is the masterfile from which you need to compare the columns of file 1.
            result.xlsx is the new file which will be created. 


EXAMPLE

bashpython filter_columns.py \
  --data file1_sample.xlsx \
  --master masterfile_sample.xlsx \
  --output result.xlsx

SAMPLE OUTPUT:

Done. 4 matching column(s) written to 'result.xlsx'.
Matched columns: ['VENDOR NAME ', ' Invoice Number', 'Amount (USD)', 'Sales  Region']
Columns in data file NOT found in masterfile (excluded): ['invoice_date', 'Internal Notes']

SAMPLE FILES

file1_sample.xlsx and masterfile_sample.xlsx are included in this repo to demonstrate the column-matching behavior, 
including mismatched casing and whitespace.

NOTES

Underscore vs. space (e.g. invoice_date vs. Invoice Date) is treated as a genuinely different name and will not match. 
Only casing and whitespace differences are normalized. If no columns match between the two files, the script still runs 
and produces an empty output file, with a warning printed to the console.
