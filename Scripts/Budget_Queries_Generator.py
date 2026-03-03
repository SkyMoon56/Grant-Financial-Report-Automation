import pandas as pd
import os
import warnings
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows

warnings.filterwarning("ignore", category = UserWarning, module = 'openpyxl')

#Find the files
possible_paths = [
    os.path.join(os.path.expanduser("~"), "OneDrive", "Desktop", "SBSC"),
    os.path.join(os.path.expanduser("~"), "Desktop", "SBSC"),
    r"C:\Users\user\OneDrive\Desktop\SBSC"
]

SCRIPT_DIR = None
for p in possible_paths:
    if os.path.exists(p):
        SCRIPT_DIR = p
        break
if not SCRIPT_DIR:
    print("ERROR: Could not find the SBSC folder anywhere.")
    sys.exit()

 os.chdir(SCRIPT_DIR)
print(f"Confirmed Working Directory: {SCRIPT_DIR}")

all_files = os.listdir(SCRIPT_DIR)
budget_match = [f for f in all_files if "REFS_BUD" in f]
project_match = [f for f in all_files if "OU_SPNSR" in f]