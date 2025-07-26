# tabkit
Command line tools to read/write XLSX and CSV files, and more !

# Usage:

`xl2csv`
```text
usage: xl2csv.py [-h] [-nc] [-s [SHEET ...]] [-f [FILTER ...]] [files ...]

Convert one or more Excel (.xlsx) workbooks into unified CSV output.

positional arguments:
  files                 pattern(s) to match input filepaths. Note: considers only .xlsx files.

options:
  -h, --help            show this help message and exit
  -nc, --no-context     Dont Include `workbook_name` and `sheet_name` columns in every output row (default OFF).
  -s [SHEET ...], --sheet [SHEET ...]
                        pattern(s) to match sheet names inside workbook
  -f [FILTER ...], --filter [FILTER ...]
                        pattern(s) to match row values inside worksheets and filter
```

`csv2xl`
```text
usage: csv2xl.py [-h] [-s] [-o OUTPUT] [-w]

Convert piped CSV data into one or more Excel (.xlsx) workbooks and worksheets.

options:
  -h, --help            show this help message and exit
  -s, --split           Uses first 2 rows as workbook name and sheet name and splits data accordingly into workbooks and sheets
  -o OUTPUT, --output OUTPUT
                        Output dir (defaults to current directory)
  -w, --overwrite       pattern(s) to match row values inside worksheets and filter
```

`csv2tab`
```text
(some commands that write to stdout) | csv2tab
```
