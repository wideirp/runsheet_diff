import os
from pathlib import Path
import argparse
import itertools
import openpyxl as xl

parser = argparse.ArgumentParser()

parser.add_argument("--runsheet", help="path to runsheet")
parser.add_argument("--sheet", help="name of sheet; usually Runsheet")
parser.add_argument(
    "--col", help="column to get Volume & Page info; usually C")
parser.add_argument("--row_start", help="first row of instruments; usually 34")
parser.add_argument("--instruments", help="path to instruments folder")

args = parser.parse_args()

rs = xl.load_workbook(args.runsheet)
wb = rs[args.sheet]

rs_fns = [cell.value.strip()
          for cell in wb[args.col] if (cell.value and cell.row >= int(args.row_start) and cell.row <= 2000)]

inst_fns = [Path(fn).stem.strip()
            for fn in os.listdir(args.instruments) if fn[0] != '.']

inst_missing = [fn for fn in rs_fns if fn not in inst_fns]

rs_missing = [fn for fn in inst_fns if fn not in rs_fns]


# OUTPUT
# Missing from Runsheet    Missing from Instruments Folder
temp = '{0:25}  {1}'

print(temp.format(f"Runsheet Total: {
      len(rs_fns)}", f"Instrument Total: {len(inst_fns)}"))
print(temp.format("Missing from Runsheet", "Missing from Instruments Folder"))
print(temp.format("---------------------", "-------------------------------"))
for row in itertools.zip_longest(rs_missing, inst_missing, fillvalue=""):
    print(temp.format(row[0], row[1]))
