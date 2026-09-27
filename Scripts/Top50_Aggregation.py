import pandas as pd
import openpyxl
import os
from pathlib import Path
import calendar
import datetime

marketcap_cleaned = Path("../Top 50 Market Cap/Cleaned")

with pd.ExcelWriter("../aggregated_top_50.xlsx",engine="openpyxl") as writer:
    aggregated_df = []

    for files in os.listdir(marketcap_cleaned):
        curr_month = int(files.split(" ")[4][1])
        df = pd.read_excel(marketcap_cleaned/files  )
        df["months"] = datetime.date(2026, curr_month, calendar.monthrange(2026,curr_month)[1])

        aggregated_df.append(df)

    aggregated_df = pd.concat(aggregated_df,ignore_index=True)
    aggregated_df.to_excel(writer)