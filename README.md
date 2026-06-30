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

1. git clone https://github.com/ayush7-hash/project1.git
   
Note on python vs python3 (and pip vs pip3): Depending on your OS, the command to run Python may be python or python3, and similarly the package installer may be pip or pip3. To check which ones to use, run:

bashpython --version
python3 --version
pip --version
pip3 --version

Use whichever python/python3 command returns a version starting with 3.x (e.g. Python 3.12.1), and the matching pip/pip3 command (whichever one doesn't return "command not found"). On Windows, python and pip usually work out of the box. On macOS/Linux, you'll often need python3 and pip3 instead, unless you've set up aliases (see below). All commands in this README use python/pip — substitute python3/pip3 if that's what your system requires.


Repo Access

HTTPS with a Personal Access Token (PAT)

Go to GitHub → profile picture → Settings → Developer settings → Personal access tokens → Tokens (classic).
Click Generate new token (classic), give it a name and expiration, and check the repo scope.Make sure to check the repo scope , or 
else your installation will not work.
Click Generate token and copy it immediately — GitHub only shows it once.

Run:

git clone https://github.com/ayush7-hash/project1.git


When prompted:


   Username for 'https://github.com': <your-github-username>
   Password for 'https://<your-github-username>@github.com': <paste-your-token-here>

Use your token, not your actual GitHub account password — passwords are rejected.

2. cd project1
3. pip install -r requirements.txt (use pip or pip3 as per your system)


USAGE

python filter_columns.py --data file1.xlsx --master masterfile.xlsx --output result.xlsx (use Python or Python3 as per your system)

NOTE : Here file1.xlsx is the orignal excel file you need to make changes in.
            masterfile.xlsx is the masterfile from which you need to compare the columns of file 1.
            result.xlsx is the new file which will be created. 


EXAMPLE

python filter_columns.py \
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
