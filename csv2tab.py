# usage: stdout | csv2tab

import argparse
parser = argparse.ArgumentParser(description='Display CSV data from stdin as a formatted table on the terminal.')
parser.usage = '(some commands that write to stdout) | csv2tab'

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

if sys.stdin.isatty():
    parser.print_usage()
    sys.exit(1)
else:
    pass

import shutil
import textwrap
from tabulate import tabulate

MIN_COL_WIDTH = 1

def wrap_cell(text:str, width):
    return "\n".join(
        textwrap.wrap(
            text, 
            width=int(width),
            replace_whitespace=False
        )
    ).strip() if text else ""

def get_width(cell:str):
    max_ln_len = 0
    for line in cell.splitlines():
        max_ln_len = max(len(line), max_ln_len)
    return max_ln_len

def get_col_widths(rows):
    n_cols = max([len(row) for row in rows])
    col_widths = [MIN_COL_WIDTH for i in range(n_cols)]
    for row in rows:
        for idx, cell in enumerate(row):
            col_widths[idx] = max(col_widths[idx], get_width(cell))
    return col_widths

def reduce_col_widths(col_widths, delta):
    red_col_widths = [c for c in col_widths]
    for i in range(delta):
        # decrease the max column by 1
        red_col_widths[red_col_widths.index(max(red_col_widths))] -= 1
    return red_col_widths

def preprocess_data(rows):
    # Get terminal size
    term_width = shutil.get_terminal_size().columns # check cross compatibility works on windows

    # data = data.strip()

    if not rows:
        print('<empty stream>')
        exit(0)

    rows = [[cell.replace('\\n', '\n') for cell in row] for row in rows]
    
    col_widths = get_col_widths(rows)

    n_cols = len(col_widths)
    available_width = term_width - (2*2 + (n_cols-1)*3) - 1
        # 2 borders takes 2 pos each
        # n columns have n-1 seperators taking 3 pos each
        # when table size == term size, it adds line breaks for each line printed. so, -1
    actual_width = sum(col_widths)

    prop_widths = reduce_col_widths(col_widths=col_widths, delta=(actual_width - available_width))

    # print(col_widths, sum(col_widths), available_width)
    # print(prop_widths, sum(prop_widths), term_width)

    if any(w < MIN_COL_WIDTH for w in prop_widths):
        print('terminal size too small, zoom out!')
        exit(0)

    # Wrap each cell
    wrapped_rows = [
        [wrap_cell(cell, prop_widths[idx]) for idx, cell in enumerate(row)]
        for row in rows
    ]

    # Print table
    # print(tabulate(wrapped_rows, tablefmt="grid"))
    return wrapped_rows

# Usage
if __name__ == "__main__":
    import csv 
    csv_reader = csv.reader(sys.stdin)
    rows = list(csv_reader)
    processed_rows = preprocess_data(rows)
    print(tabulate(processed_rows, tablefmt='grid'))
