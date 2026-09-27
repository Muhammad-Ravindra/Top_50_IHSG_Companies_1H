import pandas as pd
import openpyxl
from pathlib import Path
import os

stock_list_path = Path("../Daftar Saham")
project_main_dir = Path(__file__).parent

stock_sector_list = os.listdir(stock_list_path)

current_rows = 2

aggregated_wb = openpyxl.Workbook()
ws = aggregated_wb.active

# Make headers
ws.cell(1,1).value = "Code"
ws.cell(1,2).value = "Sector"

for stock_sector in stock_sector_list:
    current_sector = str(stock_sector).split("-")[1].strip()
    sectoral_stock_df = pd.read_excel(stock_list_path / stock_sector,usecols=[1])
    sectoral_count = sectoral_stock_df.count().values[0]

    for i,rows in enumerate(ws.iter_rows(current_rows,current_rows+sectoral_count,1,1)):
        rows[0].value = sectoral_stock_df.iloc[i-1,0]

    for rows in ws.iter_rows(current_rows,current_rows+sectoral_count,2,2):
        rows[0].value = current_sector

    current_rows += sectoral_count

aggregated_wb.save("../aggregated_code_sector.xlsx")