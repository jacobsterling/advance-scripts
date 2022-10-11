import pandas as pd
from pathlib import Path

res = pd.DataFrame([], columns = ["CLI", "Start time"])

for file in Path().absolute().glob("*"):
    if file.name.__contains__("PSC"):
        calls = pd.read_excel(file, sheet_name="Call lists", usecols=["CLI", "Start time", "Title"])
        
        calls["Method"] = "outbound" if file.name.__contains__("outbounds") else "inbound"
        
        for i, row in calls.iterrows():
            if row["Method"] == "outbound":
                calls.at[i, "CLI"] = row["Title"].split(" ")[3]
        
        calls = calls.drop(columns = "Title")
        
        res = pd.concat([res, calls])

res.to_csv("call list.csv", index=False)