# Purpose: to generate a list of { name: 'A1', color: 'fff5ca' } from .xlsx file

import pandas as pd

INPUT_FILE = '/Users/zhaoxue/Documents/beads/fuse-bead-tool/Colors.xlsx'
OUTPUT_FILE = '/Users/zhaoxue/Documents/beads/fuse-bead-tool/list.txt'
COLOR_HEX_HEADER = 'Color (Hex)'
COLOR_ID_HEADER = 'Mard'

# load excel file
df = pd.read_excel(INPUT_FILE)

# select columns
data = df[[COLOR_HEX_HEADER, COLOR_ID_HEADER]]

# compose output
result_list = [f"{{ name: '{row[COLOR_ID_HEADER]}', color: '{row[COLOR_HEX_HEADER]}' }}" for _, row in data.iterrows()]

# print to file
with open(OUTPUT_FILE, 'w') as f:
    f.write(',\n'.join(result_list))