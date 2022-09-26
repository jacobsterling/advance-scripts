# -*- coding: utf-8 -*-
"""
Created on Tue Aug  9 12:59:59 2022

@author: jacob.sterling
"""

#column for if profit below £10

                
class rebates:
    def __init__(self):
        from pathlib import Path
        from utils import formats
        from utils import functions
        from datetime import datetime
        from datetime import date
        
        self.pd = __import__('pandas')
        
        yearAbbr = formats.taxYear().Year_format1("-")
        year = formats.taxYear().Year("-")
        yearc = formats.taxYear().yearc
        self.week = functions.tax_calcs().tax_week_calc()
        
        monthNum = int(input(rf"Enter Month Number ({year}): "))
        
        homePath = Path.home() / "advance.online"
        rebatesPath = homePath / rf"J Drive - Finance/Rebates Reports/Rebates {yearAbbr}"
        marginsPath = homePath / rf"J Drive - Exec Reports/Margins Reports/Margins {year}"
        groupsPath = Path.home() / "OneDrive - advance.online/Documents/data/groups.xlsx"
        
        self.exceptions = self.pd.read_excel("rebate exceptions.xlsx")
        groups = self.pd.read_excel(groupsPath)
        
        marginsReport = self.pd.read_excel(marginsPath / rf"Margins Report 2022-2023.xlsx", na_values=['-'] ,sheet_name= ['Core Data','PAYE Data','Accounts 2'])
        
        #{year}
        
        self.margins = marginsReport['Core Data']
        self.paye = marginsReport['PAYE Data']
        self.paye.columns = self.paye.iloc[4, :]
        self.paye = self.paye.iloc[5:, 0:].reset_index(drop=True)
        
        accounts = marginsReport['Accounts 2'][ marginsReport['Accounts 2']["Account Type"] == "Agency" ][["Office Number","Account Name","Account No", "Nominal Code PSF", "Default Rebate", "Rebate Conditions"]]
        
        accounts.loc[~accounts["Nominal Code PSF"].isna() , "Account No"] = accounts.loc[~accounts["Nominal Code PSF"].isna(), "Nominal Code PSF"]
                
        rebateDetails = accounts[~accounts["Default Rebate"].isna() | ~accounts["Rebate Conditions"].isna()]
        
        self.margins = self.margins[(marginsReport['Core Data']["CHQDATE"].dt.month == monthNum) & (marginsReport['Core Data']["CHQDATE"].dt.year == yearc)]
        
        self.rebates = self.margins[(marginsReport['Core Data']["CHQDATE"].dt.month == monthNum) & (marginsReport['Core Data']["CHQDATE"].dt.year == yearc)]
        
        for i, row in groups.iterrows():
            if not self.pd.isnull(row['Office Number']):
                self.rebates.loc[self.rebates['Client Name'] == row['Client Name'], 'Office Number'] = float(row['Office Number'])
            
            self.margins.loc[self.margins['Client Name'] == row['Client Name'], 'Group Name'] = row['Name Change']
            self.rebates.loc[self.rebates['Client Name'] == row['Client Name'], 'Client Name'] = row['Name Change']
    
        self.rebates = self.rebates.merge(rebateDetails, validate="many_to_one", how = "outer")
        
        self.unmerged = self.rebates[self.rebates["Account Name"].isna()]
        self.rebates = self.rebates[~self.rebates["Account Name"].isna()]
        
        
    @staticmethod
    def valueCheck(v: str, margin: float):
        if v.__contains__("x"):
            return float(v.replace("x", ""))*margin
        else:
            return float(v)
    
    def run(self):
        import re
        
        condition_match = re.compile(r"(?:(?:(?:([><=]{1,2}) ([0-9]{1,2}(?:\.[0-9]{1,2})?)?)|((?:[a-zA-Z]{2,8} ?){1,3})) = ([0-9]{1,2}(?:\.[0-9]{1,2})?x?))")
        
        for i, row in self.rebates.iterrows():
            def format_error(format, msg):
                print(rf"OFFNO {row['Office Number']}: Invalid Rebate Condition Format: {format}, {msg}")
                return
            
            value = None
            margin = float(row["Margins"])
            
            try:
                min_margin = float(row["Fee Type"])
            except ValueError:
                # unmanaged fee type
                min_margin = 0
            
            if margin > 0 and row["PAYNO"] not in self.exceptions["PAYNO"]:
                if row["Client Name"] == "CORRIE" and self.pd.isnull(row["PAYNO"]):
                    value = float(self.paye.loc[row["Week Number"] - 1, "Rebate"])
                    
                elif not self.pd.isnull(self.rebates.at[i, "Rebate Conditions"]):
                    for condition in condition_match.finditer(row["Rebate Conditions"]):
                        group = condition.groups()
                        
                        if group[3]:
                            v = None
                            
                            if group[3].__contains__("x"):
                                try:
                                    multiplier = float(group[3].replace("x", ""))
                                    v = margin * multiplier
                                except ValueError:
                                    format_error(group, "invalid multiplier")
                            else:
                                try:
                                    v = float(group[3])
                                except ValueError:
                                    format_error(group, "value not a number")
                            
                            if not self.pd.isnull(v):
                                match group[0]:
                                    case "<":
                                        if group[1]:
                                            if margin < float(group[1]) and margin >= min_margin:
                                                value = v
                                        else:
                                            format_error(group, "no operator value")
                                    
                                    case "<=":
                                        if group[1]:
                                            if margin <= float(group[1]) and margin >= min_margin:
                                                value = v
                                        else:
                                            format_error(group, "no operator value")
                                            
                                    case ">":
                                        if group[1]:
                                            if margin > float(group[1]):
                                                value = v
                                        else:
                                            format_error(group, "no operator value")
                                    
                                    case ">=":
                                        if group[1]:
                                            if margin >= float(group[1]):
                                                value = v
                                        else:
                                            format_error(group, "no operator value")
                                    
                                    case "=":
                                        if group[1]:
                                            if margin == float(group[1]):
                                                value = v
                                        else:
                                            format_error(group, "no operator value")
                                    
                                    case "==":
                                        if group[1]:
                                            if margin == float(group[1]):
                                                value = v
                                        else:
                                            format_error(group, "no operator value")
                                            
                                    case None:
                                        if group[2]:
                                            if group[2] == row["Solution.1"]:
                                                value = v

                                        else:
                                            format_error(group, "no solution")
                                            
                                    case other:
                                        format_error(group, "unmanaged operator")
                            else:
                                print(v)
                                format_error(group, "no condition value")  
                        else:
                            format_error(group, "no rebate value")
                                
                if not self.pd.isnull(self.rebates.at[i, "Default Rebate"]) and not value and margin >= min_margin:
                    value = float(self.rebates.at[i, "Default Rebate"].replace("£", "").replace("Â","").replace(" ", ""))    
                
                self.rebates.at[i, "Rebate"] = value if value else 0
            else:
                if row["PAYNO"] in self.exceptions["PAYNO"] and row["Client Name"] in self.exceptions["Client Name"]:
                    print(row["PAYNO"])
                self.rebates.at[i, "Rebate"] = 0
        self.export()
        
    def export(self):
        import numpy as np
        
        pivot = self.pd.pivot_table(self.rebates, columns = ["CHQDATE"], values=['Rebate', 'Margins', "Count of"], index=["Client Name","Account No"], aggfunc={'Count of': np.sum, 'Margins': np.sum,'Rebate': np.sum}, fill_value=0)
        
        with self.pd.ExcelWriter("Rebates.xlsx") as writer:
            pivot.to_excel(writer, sheet_name="Rebates")
            self.rebates.to_excel(writer, sheet_name="Rebate Data" ,index= False)
            self.margins.to_excel(writer, sheet_name="Core Data", index= False)
            self.unmerged.to_excel(writer, sheet_name="Unmerged", index= False)
        writer.save()

rebates().run()
