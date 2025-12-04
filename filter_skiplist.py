
import argparse
parser = argparse.ArgumentParser(description='filters csv stream of SMV workbooks if skiplist keyword is defected in 8th field')
# parser.add_argument('format', help='format string ("line" variable stores current line.)')
parser.usage = '(csv stream of smv-analysis-workbooks) |' + parser.format_usage()[6:]

arg_dict = vars(parser.parse_args())

import sys
import io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

if sys.stdin.isatty():
    parser.print_usage()
    sys.exit(1)
else:
    pass

def get(lst, idx, default=''):
    return lst[idx] if -len(lst) <= idx < len(lst) else default

sheets=['appr_0110', 'decl_0110', 'appr_0130', 'decl_0130', 'appr_0410', 'decl_0410']

# Usage
import csv 
if __name__ == "__main__":
    
    writer = csv.writer(sys.stdout, delimiter=',', lineterminator='\n')
    for line in csv.reader(sys.stdin):
        try:
            if any(sh in get(line,0).strip().lower() for sh in sheets):
                if get(line, 8).strip().lower()=='skiplist':
                    # record = [eval('f"' + fs + '"') for fs in format_str]
                    writer.writerow(line)
        except KeyboardInterrupt:
            break 
        except BrokenPipeError:
            break
    
