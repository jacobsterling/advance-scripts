# -*- coding: utf-8 -*-
"""
Created on Tue Aug  9 12:59:59 2022

@author: jacob.sterling
"""

class rebates:
    def __init__(self):
        from pathlib import Path
        from formats import taxYear
        from functions import tax_calcs
        from datetime import datetime
        from datetime import date
        self.pd = __import__('pandas')
        
        yearAbbr = taxYear().Year_format1("-")
        year = taxYear().Year("-")
        yearc = taxYear().Yearc()
        self.week = tax_calcs().tax_week_calc()
        
        monthNum = int(input(rf"Enter Month Number ({year}): "))
        
        homePath = Path.home() / "advance.online"
        rebatesPath = homePath / rf"J Drive - Finance/Rebates Reports/Rebates {yearAbbr}"
        marginsPath = homePath / rf"J Drive - Exec Reports/Margins Reports/Margins {year}"
        dataPath = marginsPath / rf"Data/Week {self.week}"
        
        rebateDetails = self.pd.read_csv("Rebate+Details.csv", skiprows=6, na_values=['-'])
        marginsReport = self.pd.read_excel(marginsPath / rf"Margins Report 22-23.xlsx", na_values=['-'] ,sheet_name= ['Core Data','PAYE Data','Accounts 2'])
        
        #{year}
        
        margins = marginsReport['Core Data']
        self.paye = marginsReport['PAYE Data']
        self.paye.columns = self.paye.iloc[4, :]
        self.paye = self.paye.iloc[5:, 0:].reset_index(drop=True)
        
        #accounts = marginsReport['Accounts 2'][ marginsReport['Accounts 2']["Account Type"] == "Agency" ]["Office Number","Account Name","Account No"]
        #accounts.loc[~accounts["Nominal Code PSF"].isna() , "Account No"] = accounts.loc[~accounts["Nominal Code PSF"].isna(), "Nominal Code PSF"]
        rebateDetails.loc[~rebateDetails["Nominal Code PSF"].isna() , "Account No"] = rebateDetails.loc[~rebateDetails["Nominal Code PSF"].isna(), "Nominal Code PSF"]
        
        self.monthMargins = margins[(margins["CHQDATE"].dt.month == monthNum) & (margins["CHQDATE"].dt.year == yearc)].merge(rebateDetails, validate="many_to_one", how = "outer")
        self.monthMargins["Client Name"] = self.monthMargins["Client Name"].str.upper()
        self.rebates = self.monthMargins[~self.monthMargins["Record Id"].isna()]
        self.unmergedRebates = rebateDetails[~rebateDetails["Account No"].isin(self.rebates["Account No"])]
        self.unmergedMargins = self.monthMargins[self.monthMargins["Record Id"].isna()]
    
    @staticmethod
    def valueCheck(v: str, margin: float):
        if v.__contains__("x"):
            return float(v.replace("x", ""))*margin
        else:
            return float(v)
    
    
    def run(self):
        for i, row in self.rebates.iterrows():
            mn = 22 #default min margin
            
            def strip_number(str):
                if not self.pd.isnull(str):
                    return float(str.replace("Â","").replace("£","").replace(" ", ""))
                else: 
                    return None
        
            match row["Solution"]:
                case "CIS": 
                    mn = strip_number(row["Retained Margin CIS"])
                case "PAYE":
                    match row["Type"]:
                        case "Fixed Expenses":
                            mn = strip_number(row["Retained Margin Umbrella with Expenses"])
                        case "Mileage Only":
                            mn = strip_number(row["Retained Margin Umbrella with Expenses"])
                        case "Not Under SDC":
                            mn = strip_number(row["Retained Margin Umbrella no Expenses"])  
                case "SE":
                    mn = strip_number(row["Retained Margin Non CIS (SE)"])
                            
            if not self.pd.isnull(row["Default Rebate"]):
                try:
                    value = strip_number(row["Default Rebate"])
                except ValueError:
                    raise Exception(f"Invalid Default Rebate format for offno: {row['Office Number']}")
            else:
                value = 0
            
            margin = float(row["Margins"])
            
            if not self.pd.isnull(row["Solution Conditions"]):
                for condition in row["Solution Conditions"].split(","):
                    try:
                        solution, v = condition.split(": ")
                    except ValueError:
                        raise Exception(f"Invalid Rebate Condition format: {condition}, for offno: {row['Office Number']}")
                
                if solution == row["Solution.1"]:
                    value = strip_number(v)
                    break
            
            if not self.pd.isnull(row["Margin Conditions"]):
                for condition in row["Margin Conditions"].split("\n"):
                    try:
                        x, mx, v = condition.split(",")
                    except ValueError:
                        raise Exception(f"Invalid Rebate Condition format: {condition}, for offno: {row['Office Number']}")
                    
                    if x != "x":
                        try:
                            mn = float(x.replace(" ", ""))    
                        except ValueError:
                            raise Exception(f"Invalid Rebate range format: {mn} {mx}, for offno: {row['Office Number']}")         
                    
                    try: 
                        mx = float(mx.replace(" ", ""))    
                    except ValueError:
                        mx = None
                        
                    if mx:
                        if mn <= margin <= mx:
                            value = self.valueCheck(v, margin)
                            break
                    elif mn <= margin:
                        value = self.valueCheck(v, margin)
                        break
                    
            if row["Client Name"] == "CORRIE" and self.pd.isnull(row["PAYNO"]):
                self.rebates.loc[i, "Rebate"] = float(self.paye.loc[row["Week Number"] - 1, "Rebate"])
            else:
                self.rebates.loc[i, "Rebate"] = value
                
        self.export()
        
    def export(self):
        import numpy as np
        
        pivot = self.pd.pivot_table(self.rebates, values=['Rebate', 'Margins', "Count of"], index=["Client Name","Account No"], aggfunc={'Count of': np.sum, 'Margins': np.sum,'Rebate': np.sum}, fill_value=0)
        
        with self.pd.ExcelWriter("Rebates.xlsx") as writer:
            pivot.to_excel(writer, sheet_name="Rebates")
            self.rebates.to_excel(writer, sheet_name="Rebate Data" ,index= False)
            self.unmergedRebates.to_excel(writer, sheet_name="Unmerged Rebates", index= False)
            self.unmergedMargins.to_excel(writer, sheet_name="Unmerged Margins", index= False)
            self.monthMargins.to_excel(writer, sheet_name="Core Data", index= False)
        writer.save()

rebates = rebates().run()
