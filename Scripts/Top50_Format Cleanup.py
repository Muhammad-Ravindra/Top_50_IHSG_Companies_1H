import pandas as pd
import os
import openpyxl
from pathlib import Path
from utils import month_num

Top50_Raw_Folders = Path("../Top 50 Market Cap/Raw")
Top50_Cleaned_Folders = Path("../Top 50 Market Cap/Cleaned")

def cleanup(file):
    curr_month = str(file).split(" ")[5]
    wb = openpyxl.load_workbook(os.path.join(Top50_Raw_Folders,file))
    ws = wb.active

    for mr in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(mr))

    ws["B2"].value = None
    ws["F4"].value = "Market Cap IDR (Billions)"
    ws["F5"].value = None
    ws["G4"].value = f"Market Cap {ws["G5"].value}"
    ws["G5"].value = None
    ws.delete_rows(5,1)
    ws.delete_rows(56,3)
    ws.delete_rows(1,3)
    ws.delete_cols(1,1)
    wb.save(os.path.join(Top50_Cleaned_Folders,f"50 Biggest Market Cap {month_num.get(curr_month)}.xlsx"))

for f in os.listdir(Top50_Raw_Folders):
    cleanup(f)