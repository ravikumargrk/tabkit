# usage: csv2xl [options] <input-csv-file(s)>
# Convert CSV data into one or more Excel (.xlsx) workbooks and worksheets.
# If `workbook_name` and `sheet_name` columns are present and mapped, splits data into respective files and sheets.

# Options:
#   -h, --help            Show this help message and exit.
#   -s, --split           uses first 2 rows as workbook name and sheet name and splits accordingly
#                         Note: Raises error if number of fields < 2 for any rows or first field does not end with .xlsx
#                         Default value OFF (Writes all output to 1 workbook and 1 sheet)
#   -o, --output DIR      Specify output directory.
#                         Defaults to current directory.
#                         Note: default workbook name is Book{n}.xlsx integer (n) will be set so as to not overwrite existing workbooks in given path
#   -w, --overwrite       Overwrites if needed (Does not use integer naming)
#                
# Examples:
#   # Rebuild XLSX files from CSV with workbook_name and sheet_name fields as routing directives
#   csv2xl data_with_context.csv --map-context workbook_name,sheet_name --output ./output_dir/
#   # Convert plain CSV to a single XLSX file ignoring context columns
#   csv2xl flat_data.csv --literal --output report.xlsx

import argparse
parser = argparse.ArgumentParser(description='Convert piped CSV data into one or more Excel (.xlsx) workbooks and worksheets.')
parser.add_argument('-s', '--split', action='store_true', help='Uses first 2 rows as workbook name and sheet name and splits data accordingly into workbooks and sheets')
parser.add_argument('-o', '--output', type=str, default='.', help='Output dir (defaults to current directory)')
parser.add_argument('-w', '--overwrite', action='store_true', help='pattern(s) to match row values inside worksheets and filter', )

arg_dict = vars(parser.parse_args())

WB_DEFAULT = 'Book.xlsx'
ST_DEFAULT = 'Sheet'

import sys
import csv 

if sys.stdin.isatty():
    # parser.
    usage_str = parser.format_usage()
    parser.usage = '(some commands that print to stdout) |' + usage_str[6:]
    parser.print_usage()
    exit(0)
else:
    pass

# contains dict[str, set] = {workbook:{sheet1, sheet2}}
meta_data:dict[str, dict] = {}

def convert(field: str):
    """Convert field to appropriate type (int, float, or string)."""
    field = field.strip()
    if not field:
        return ""
    if field.isnumeric():
        return int(field)
    else:
        try:
            res = float(field)
            return res
        except ValueError:
            return field

def createWorkBook(file, sheet):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    meta_data[file] = {}
    meta_data[file]['object'] = wb
    meta_data[file]['sheets'] = {}
    meta_data[file]['sheets'][sheet] = {}
    meta_data[file]['sheets'][sheet]['object'] = ws
    meta_data[file]['sheets'][sheet]['write_row'] = 1
    
def createSheet(file, sheet):
    wb:Workbook = meta_data[file]['object']
    ws = wb.create_sheet(title=sheet)
    meta_data[file]['sheets'][sheet] = {}
    meta_data[file]['sheets'][sheet]['object'] = ws
    meta_data[file]['sheets'][sheet]['write_row'] = 1

def getSaveFilename(path:str):
    if arg_dict['split']:
        import os 
        if not os.path.exists(path):
            return path 
    
    if '.' in path:
        filename, ext = path[:path.rfind('.')], path[path.rfind('.'):]
        cut_start = len(filename)
        cut_end = -len(ext)
    else:
        filename, ext = path, ''
        cut_start = len(filename)
        cut_end = None

    # check all prexisting filenames
    from glob import glob
    # idx = len(glob(filename+'*'+ext))
    existing_numeric_suffixes = [int(existing_filename[cut_start: cut_end]) for existing_filename in glob(filename+'*'+ext) if existing_filename[cut_start: cut_end].isnumeric()]
    max_idx = max(existing_numeric_suffixes) if existing_numeric_suffixes else 0
    return filename + f'{max_idx+1}' + ext

def save_all_workbooks():
    """Save all created workbooks to files."""
    for filename, data in meta_data.items():
        wb: Workbook = data['object']
        new_filename = getSaveFilename(filename)
        if not new_filename == filename:
            if arg_dict['overwrite']:
                new_filename = filename
                print('Warning: overwriting', filename)
            else:
                print('Warning:', filename, 'already exists so saving workbook as', new_filename)
        wb.save(new_filename)

from openpyxl import Workbook
from openpyxl.worksheet.worksheet import Worksheet

def append(file, sheet, row:list[str]):
    new = False
    if file not in meta_data:
        createWorkBook(file, sheet)
        # apppend line
        new = True
    else:
        if not sheet in meta_data[file]['sheets']:
            createSheet(file, sheet)
            # append line
            new = True
    # Parse CSV line properly using csv module

    # by now we have file and sheet
    
    ws:Worksheet = meta_data[file]['sheets'][sheet]['object']
    current_row = meta_data[file]['sheets'][sheet]['write_row']
    for idx, field in enumerate(row):
        ws.cell(row=current_row, column=(idx+1), value=convert(field))
    meta_data[file]['sheets'][sheet]['write_row'] = current_row + 1

# init:
# if the workbook is not given
# use default workbook name
# if sheet name not given 
# use default sheet name

# if the workbook name is changed:
# if sheet name not given 
# use default sheet name

# note: no empty workbook names

try:
    # Read all input from stdin
    input_non_empty_flag = False
    
    for row in csv.reader(sys.stdin):
        
        # create workbook only when there is a line to write
        # note we may have currentSheet as None at this point
        # but currentWorkBook can never be None.

        if arg_dict['split']:
            if len(row) < 2:
                print('Error: split mode is used but input row contains less than 2 fields:\nInput row from stdin:', row , file=sys.stderr)
                exit(1)

            if not row[0].endswith('.xlsx'):
                print('Error: split mode is used but input row\'s first field does not contain workbook name ending with xlsx.\nInput row from stdin:', row , file=sys.stderr)
                exit(1)
            
            currentWorkBook, currentSheet, row = row[0], row[1], row[2:]
        else:
            currentWorkBook = WB_DEFAULT
            currentSheet = ST_DEFAULT

        # append row at right place
        append(currentWorkBook, currentSheet, row)

        input_non_empty_flag = True
    
    save_all_workbooks()

    # Input error handling
    if not input_non_empty_flag:
        print('<empty stream>', file=sys.stderr)
        sys.exit(1)
    
except KeyboardInterrupt:
    print('\nOperation cancelled by user', file=sys.stderr)
    sys.exit(1)

