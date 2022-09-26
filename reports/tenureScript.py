# -*- coding: utf-8 -*-
"""
Created on Mon Jun 20 11:42:42 2022

@author: jacob.sterling
"""

import pandas as pd
from pathlib import Path
import datetime
from utils import functions

import numpy as np

homePath = Path.home() / "advance.online/J Drive - Exec Reports/Margins Reports"

margins = pd.DataFrame([], columns=["Client Name","PAYNO", "Surname_x", "Forename","Solution","Sdc Option","Year","CHQDATE", "Email Address", "Count of", "WEEKS_PAID","NI_NO","FREQ"])

# joinersAxm = {
#     2020: pd.read_csv("joiners axm 2020.csv",encoding="latin"),
#     2021: pd.read_csv("joiners axm 2021.csv",encoding="latin"),
#     2022: pd.read_csv("joiners axm 2022.csv",encoding="latin"),
#     }


joinersIo = {
    2020: pd.read_csv("data/joiners 2020.csv",encoding="latin"),
    2021: pd.read_csv("data/joiners 2021.csv",encoding="latin"),
    2022: pd.read_csv("data/joiners 2022.csv",encoding="latin"),
    }

for year in range(2021, 2023):
    marginsReport = pd.read_excel(homePath / rf"Margins {year}-{year + 1}/Margins Report {year}-{year + 1}.xlsx", sheet_name = "Core Data")
    
    #joinersIo[year] = joinersIo[year][~joinersIo[year]["Pay No"].isin(joinersAxm[year]["Pay No"])]
    
    #joiners = pd.concat([joinersAxm[year], joinersIo[year]])
    
    joiners = joinersIo[year]
    
    joiners["Pay No"] = joiners["Pay No"].apply(lambda x: functions.PAYNO_Convert(x))
    
    marginsReport["PAYNO"] = marginsReport["PAYNO"].apply(lambda x: functions.PAYNO_Convert(x))
    
    marginsReport["Year"] = rf"{year}-{year + 1}"
    
    yearReport = marginsReport.dropna(subset = ["PAYNO"]).merge(joiners.dropna(subset = ["Pay No"]).drop_duplicates(subset=["Pay No"]), how = "left", left_on="PAYNO", right_on="Pay No").drop_duplicates(subset=["PAYNO"])[margins.columns]

    #margins = pd.concat([margins, yearReport[yearReport["Solution"] == "PAYE"]])
    margins = pd.concat([margins, yearReport[yearReport["Client Name"].str.contains("CLEARWATER PEOPLE SOLUTIONS LTD") | yearReport["Client Name"].str.contains("CLEARWATER PEOPLE'S SOLUTIONS LTD")]])
    
    #(yearReport["CHQDATE"] >= pd.Timestamp.now() - pd.Timedelta(days=365) )
    
margins["Full Name"] = margins["Forename"] + " " + margins["Surname_x"]

margins["Tenure"] = margins["WEEKS_PAID"].str.count(",") + 1

pivot = pd.pivot_table(margins.fillna(0), values = ["Tenure"], columns = ["Year", "F"], index=["Full Name", "Email Address", "PAYNO"], aggfunc={"Tenure": sum}, fill_value=np.nan, margins=True)

with pd.ExcelWriter("Tenure Clearwater.xlsx") as writer:
    pivot.to_excel(writer, sheet_name="Pivot")
    margins.to_excel(writer, sheet_name="Data", index= False)

writer.save()