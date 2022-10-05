# -*- coding: utf-8 -*-
"""
Created on Tue Aug  9 12:59:59 2022

@author: jacob.sterling
"""

#column for if profit below £10

                
class rebates:
    def __init__(self):
        from pathlib import Path
        from utils.formats import taxYear
        from utils import functions
        import datetime
        self.datetime = datetime.datetime
        
        self.pd = __import__('pandas')
        
        self.yearAbbr = taxYear().Year_format1("-")
        year = taxYear().Year("-")
        yearc = taxYear().yearc
        self.week = functions.tax_calcs().tax_week()
        
        CORE_ACCOUNTS = ['Adam Shaw','Dave Levenston','Gerry Hunnisett','Sam Amos']
        
        monthNum = int(input(rf"Enter Month Number ({year}): "))
        
        homePath = Path.home() / "advance.online"
        self.rebatesPath = homePath / rf"J Drive - Finance/Rebates Reports/Rebates {self.yearAbbr}"
        marginsPath = homePath / rf"J Drive - Exec Reports/Margins Reports/Margins {year}"
        groupsPath = Path.home() / "OneDrive - advance.online/Documents/data/groups.xlsx"
        
        self.exceptions = self.pd.read_excel("rebateExceptions.xlsx")
        groups = self.pd.read_excel(groupsPath)
        
        marginsReport = self.pd.read_excel(marginsPath / rf"Margins Report 2022-2023.xlsx", sheet_name= ['Core Data','PAYE Data','Accounts 2'], na_values="-")
        
        self.margins = marginsReport['Core Data']
        self.margins['Client Name'] =  self.margins['Client Name'].str.upper()
        self.paye = marginsReport['PAYE Data']
        self.paye.columns = self.paye.iloc[4, :]
        self.paye = self.paye.iloc[5:, 0:].reset_index(drop=True)
        
        accounts = marginsReport['Accounts 2'][ marginsReport['Accounts 2']["Account Type"] == "Agency" ]
        
        accounts.loc[~accounts["Nominal Code PSF"].isna() , "Account No"] = accounts.loc[~accounts["Nominal Code PSF"].isna(), "Nominal Code PSF"]
                
        rebateDetails = accounts[~accounts["Rebate Conditions"].isna()].drop_duplicates(subset= "Office Number")
        
        self.margins = self.margins[(marginsReport['Core Data']["CHQDATE"].dt.month == monthNum) & (marginsReport['Core Data']["CHQDATE"].dt.year == yearc)]
        self.margins['Group Name'] = self.margins['Client Name']
        
        self.chqdates = self.margins["CHQDATE"].unique()
        
        for i, row in groups.iterrows():
            if not self.pd.isnull(row['Office Number']):
                self.margins.loc[self.margins['Client Name'] == row['Client Name'], 'Office Number'] = float(row['Office Number'])
            
            self.margins.loc[self.margins['Client Name'] == row['Client Name'], 'Group Name'] = row['Name Change']
    
        self.margins = self.margins.merge(rebateDetails, validate="many_to_one", how = "outer")
        
        self.unmergedRebates = self.margins[self.margins["Client Name"].isna()]
        
        self.margins = self.margins.dropna(subset = "Client Name")
        
        self.margins.loc[~self.margins["Account Owner"].isin(CORE_ACCOUNTS), "Account Owner"] = "Unmanned"
        
    def run(self):
        import re
        
        condition_match = re.compile(r"(?:(?:(?:([><=]{1,2}) ?([0-9]{1,2}(?:\.[0-9]{1,2})?)?)|((?:[a-zA-Z]{2,8} ?){1,3})) ?= (x?\/?[0-9]{1,2}(?:\.[0-9]{1,2})?x?))")
        
        def format_error(format, msg):
                print(rf"OFFNO {row['Office Number']}: Invalid Rebate Condition Format: {format}, {msg}")
                return
        
        for i, row in self.margins.iterrows():
            
            value = None
            margin = float(row["Margins"])
            
            if row["PAYNO"] not in self.exceptions["PAYNO"]:
                if row["Client Name"] in ["CORRIE", "MASTER PEACE RECRUITMENT"] and self.pd.isnull(row["PAYNO"]):
                    value = float(self.paye.loc[row["Week Number"] - 1, "Rebate"]) if row["Client Name"] == "CORRIE" else 0 # or rebate * count of ?? masterpiece
                    
                elif not self.pd.isnull(self.margins.at[i, "Rebate Conditions"]):
                    for condition in condition_match.finditer(row["Rebate Conditions"]):
                        group = condition.groups()
                        
                        if group[3]:
                            v = None
                            
                            if group[3].__contains__("x"):
                                try:
                                    if group[3].__contains__("/"):
                                        divisor = float(group[3].replace("x", "").replace("/", ""))
                                        v = margin / divisor if margin > 0 else 0
                                    else:
                                        multiplier = float(group[3].replace("x", ""))
                                        v = margin * multiplier if margin > 0 else 0
                                    
                                except ValueError:
                                    format_error(group, "invalid multiplier")
                            else:
                                try:
                                    v = float(group[3])
                                except ValueError:
                                    format_error(group, "value not a number")
                            
                            if v:
                                match group[0]:
                                    case "<":
                                        if group[1]:
                                            if margin < float(group[1]) and margin < 0:
                                                value = v
                                        else:
                                            format_error(group, "no operator value")
                                    
                                    case "<=":
                                        if group[1]:
                                            if margin <= float(group[1]) and margin < 0:
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
                            format_error(group, "no rebate value")
                
                self.margins.at[i, "Rebate"] = value if value else 0
            else:
                if row["PAYNO"] in self.exceptions["PAYNO"] and row["Client Name"] in self.exceptions["Client Name"]:
                    print(row["PAYNO"])
                self.margins.at[i, "Rebate"] = 0
        self.export()
        
    def export(self):
        import numpy as np
        from utils.functions import tax_calcs
        
        max = self.pd.to_datetime(str(np.max(self.chqdates)))
        min = self.pd.to_datetime(str(np.min(self.chqdates)))
        period = tax_calcs().period(max)
        month = max.strftime("%B")
        yA, _ = self.yearAbbr.split("-")
        
        rebateDir = self.rebatesPath / rf"{month} {yA}"
        
        self.unmergedMargins = self.margins[self.margins["Account Name"].isna()]
        
        self.margins["Revenue"] = self.margins['Margins'] - self.margins['Rebate']
        
        self.margins["Average Margin"] = self.margins['Margins']
        
        self.margins.loc[self.margins["Average Margin"] <= 0, "Average Margin"] = np.nan
        
        rebates = self.pd.pivot_table(self.margins[~self.margins["Account Name"].isna()], columns = ["CHQDATE"], values=['Rebate', 'Margins'], index=['Group Name', "Client Name"], aggfunc={'Margins': np.sum,'Rebate': np.sum}, fill_value=0, margins = True)
        
        netMarginCore = self.pd.pivot_table(self.margins[self.margins["Account Owner"] != "Unmanned"], values=['Rebate', 'Margins', "Count of", "Revenue", "Average Margin"], index=["Account Owner", 'Group Name', "Client Name"], aggfunc={'Count of': np.sum, "Average Margin": np.mean, 'Margins': np.sum,'Rebate': np.sum,"Revenue": np.sum, }, fill_value=0, margins = True)
        
        netMarginOther = self.pd.pivot_table(self.margins[self.margins["Account Owner"] == "Unmanned"], values=['Rebate', 'Margins', "Count of", "Revenue", "Average Margin"], index=["Client Name"], aggfunc={'Count of': np.sum, "Average Margin": np.mean, 'Margins': np.sum,'Rebate': np.sum,"Revenue": np.sum, }, fill_value=0, margins = True)
        
        with self.pd.ExcelWriter(rebateDir / rf"{month} py Rebates {self.yearAbbr}.xlsx") as writer:
            rebates.to_excel(writer, sheet_name="Rebates")
            netMarginCore.to_excel(writer, sheet_name="Net Margins Core")
            netMarginOther.to_excel(writer, sheet_name="Net Margins Other")
            self.margins.to_excel(writer, sheet_name="Core Data", index= False)
            self.unmergedRebates.to_excel(writer, sheet_name="Unmerged Rebates", index= False)
            self.unmergedMargins.to_excel(writer, sheet_name="Unmerged Margins", index= False)
            wb = writer.book
            money_fmt = wb.add_format({'num_format': '£#,##0.#0'})
            ws = writer.sheets['Net Margins Core']
            ws.set_column('F:H', 12, money_fmt)
            ws.set_column('D:D', 12, money_fmt)
            ws = writer.sheets['Net Margins Other']
            ws.set_column('D:F', 12, money_fmt)
            ws.set_column('B:B', 12, money_fmt)
            ws = writer.sheets['Rebates']
            ws.set_column('C:N', 12, money_fmt)
        writer.save()

        self.margins["Group Sum"] = self.margins["Rebate"]*1.2
        upload = self.pd.pivot_table(self.margins[~self.margins["Account Name"].isna()], values=['Group Sum', 'Count of'], index=["Group Name", "Account No", "Account Name"], aggfunc={'Group Sum': np.sum, 'Count of': np.sum}, fill_value=0).reset_index()
        
        upload["Month"] = month
        
        upload["Rebate end week"] = max.strftime("%d/%m/%Y")
        
        upload['Group Sum'] = upload['Group Sum'].round(decimals=2)
        
        upload["Period"] = str(period) if period < 10 else "0" + str(period)
        
        psfUpload = upload.rename(columns={"Account No": "Account Code", "Group Name": "Merit Name", "Rebate end week":"Date to Accrue for"}).drop(columns=["Account Name", 'Count of'])

        psfUpload["Year"] = self.yearAbbr.replace("-", "/")
         
        psfUpload.to_csv(rebateDir / rf"{month} py Rebates {self.yearAbbr} - psf import.csv", index=False)
        
        crmUpload = upload.rename(columns={"Account Name": "CRM Name", 'Count of': "Total Margins", "Group Sum":"Total Amount"}).drop(columns=["Account No", "Group Name"])
        
        crmUpload["Rebate start week"] = min.strftime("%d/%m/%Y")
        
        crmUpload.to_csv(rebateDir / rf"{month} py Rebates {self.yearAbbr} - crm import.csv", index=False)
    
        
rebates().run()
