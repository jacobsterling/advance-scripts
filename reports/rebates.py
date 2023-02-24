# -*- coding: utf-8 -*-
"""
Created on Tue Aug  9 12:59:59 2022

@author: jacob.sterling
"""

# need to create a column for if profit below £10
   
class rebates:
    def __init__(self):
        import datetime
        from pathlib import Path

        from utils import functions, formats
    
        self.datetime = datetime.datetime

        self.pd = __import__('pandas')
        
        # defining constants such as tax year, week, core accounts and the user input
        self.yearAbbr = formats.taxYear().Year_format1("-")
        year = formats.taxYear().Year("-")
        self.week = functions.tax_calcs().tax_week()
        
        CORE_ACCOUNTS = ['Adam Shaw','Dave Levenston','Gerry Hunnisett','Sam Amos']
        
        userInput = input(rf"Enter Month Number + Year (xx {self.yearAbbr}): ")
        
        # user input validation
        monthNum, yearInput = userInput.split(" ")
        
        monthNum = int(monthNum)
        
        mn, mx = self.yearAbbr.split("-")
        
        yearc = int(rf"20{mx}") if mx == yearInput else int(rf"20{mn}")
        self.yA = int(mx) if mx == yearInput else int(mn)
        
        if monthNum < 1:
            year = formats.taxYear().Yearp("-")
            monthNum += 12
        
        # define the paths of the files
        homePath = Path.home() / "advance.online"
        self.rebatesPath = homePath / rf"J Drive - Operations/Finance/Agency Rebates/Rebates {self.yearAbbr}"
        marginsPath = homePath / rf"J Drive - Exec Reports/Margins Reports/Margins {year}"
        # used to identify groups
        groupsPath = Path.home() / "J Drive - Operations/python-utils/groups.xlsx"
        
        # workers to exclude rebates from
        self.exceptions = self.pd.read_excel("rebateExceptions.xlsx")
        groups = self.pd.read_excel(groupsPath)
        
        marginsReport = self.pd.read_excel(marginsPath / rf"Margins Report 2022-2023.xlsx", sheet_name= ['Core Data','PAYE Data','Accounts 2'], na_values="-")
        
        # reading margins report for clients, paye data and accounts
        self.margins = marginsReport['Core Data']
        self.margins['Client Name'] =  self.margins['Client Name'].str.upper()
        self.paye = marginsReport['PAYE Data']
        self.paye.columns = self.paye.iloc[4, :]
        self.paye = self.paye.iloc[5:, 0:].reset_index(drop=True)
        
        accounts = marginsReport['Accounts 2'][ marginsReport['Accounts 2']["Account Type"] == "Agency" ]
        
        # for psf upload
        accounts.loc[~accounts["Nominal Code PSF"].isna() , "Account No"] = accounts.loc[~accounts["Nominal Code PSF"].isna(), "Nominal Code PSF"]
        
        # rebate conditions from CRM
        rebateDetails = accounts[~accounts["Rebate Conditions"].isna()].drop_duplicates(subset= "Office Number")
        
        self.margins = self.margins[(marginsReport['Core Data']["CHQDATE"].dt.month == monthNum) & (marginsReport['Core Data']["CHQDATE"].dt.year == yearc)]
        self.margins['Group Name'] = self.margins['Client Name']
        
        # all chqdates from core data
        self.chqdates = self.margins["CHQDATE"].unique()
        
        # grouping the agencies by group
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
        
        #regex from rebate conditions
        condition_match = re.compile(r"(?:(?:(?:([><=]{1,2}) ?([0-9]{1,2}(?:\.[0-9]{1,2})?)?)|((?:[a-zA-Z]{2,8} ?){1,3})) ?= (x?\/?[0-9]{1,2}(?:\.[0-9]{1,2})?x?))")
        
        # custom error
        def format_error(format, msg):
                print(rf"OFFNO {row['Office Number']}: Invalid Rebate Condition Format: {format}, {msg}")
                return
        
        # iterating though margin data
        for i, row in self.margins.iterrows():
            
            # margin value and value of rebate
            margin, value = float(row["Margins"]), None
            
            # filtering exceptions
            if row["PAYNO"] not in self.exceptions["PAYNO"].values and row["Client Name"] not in self.exceptions["Client Name"]:
                
                # PAYE rebate calculation
                if row["Client Name"] in ["CORRIE", "MASTER PEACE RECRUITMENT"] and self.pd.isnull(row["PAYNO"]):
                    value = float(self.paye.loc[row["Week Number"] - 1, "Rebate"]) if row["Client Name"] == "CORRIE" else 0 # or rebate * count of ?? masterpiece
                
                # operates on only agencies with rebate conditions
                elif not self.pd.isnull(row["Rebate Conditions"]):
                    for condition in condition_match.finditer(row["Rebate Conditions"]):
                        group, v = condition.groups(), None
                        
                        # if the rebate value exists
                        if group[3]:
                            # if value contains x (x represents the margin value)
                            if group[3].__contains__("x"):
                                try:
                                    # if condition contains division operator margin value is divided by number given else multiplied
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
                                if group[0]: # optional oporator else use as solution (group 2)
                                    if group[1]:
                                        # match operator then calculate rebate accordingly
                                        match group[0]:
                                            case "<":
                                                if margin < float(group[1]) and margin < 0:
                                                    value = v
                                            
                                            case "<=":
                                                if margin <= float(group[1]) and margin < 0:
                                                    value = v
                                                    
                                            case ">":
                                                if margin > float(group[1]):
                                                    value = v
                                            
                                            case ">=":
                                                if margin >= float(group[1]):
                                                    value = v
                                            
                                            case "=" | "==":
                                                    if margin == float(group[1]):
                                                        value = v
                                                    
                                            case other:
                                                format_error(group, "unmanaged operator")
                                    else:
                                        format_error(group, "no operator value")
                                else:
                                    if group[2]:
                                        if group[2] == row["Solution.1"]:
                                            value = v

                                    else:
                                        format_error(group, "no solution")
                        else:
                            format_error(group, "no rebate value")
            self.margins.at[i, "Rebate"] = value if value else 0
        self.export()
        
    def export(self):
        # creates the rebate files
        import numpy as np
        from utils import functions
        
        max = str(np.max(self.chqdates))
        min = str(np.min(self.chqdates))
        
        period = functions.tax_calcs().period(max, "%Y-%m-%dT%H:%M:%S.%f000")
        month = self.pd.to_datetime(max).strftime("%B")
        
        rebateDir = self.rebatesPath / rf"{month} {self.yA}"
        
        self.unmergedMargins = self.margins[self.margins["Account Name"].isna()]
        
        self.margins["Revenue"] = self.margins['Margins'] - self.margins['Rebate']
        
        self.margins["Average Margin"] = self.margins['Margins']
        
        self.margins.loc[self.margins["Average Margin"] <= 0, "Average Margin"] = np.nan
        
        self.margins["CHQDATE"] = self.margins["CHQDATE"].apply(lambda x: x.strftime(format="%d/%m/%Y"))
        
        rebates = self.pd.pivot_table(self.margins[~self.margins["Account Name"].isna()], columns = ["CHQDATE"], values=['Rebate', 'Margins'], index=['Group Name', "Client Name"], aggfunc={'Margins': np.sum,'Rebate': np.sum}, fill_value=0, margins = True)
        
        rebates.loc[("RSS INFRASTRUCTURE LTD","RSS INFRASTRUCTURE LTD") , ('Rebate', 'All')] -= 1000
        
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
            money_fmt = wb.add_format({'num_format': r'£#,##0.#0'})
            ws = writer.sheets['Net Margins Core']
            ws.set_column('F:H', 12, money_fmt)
            ws.set_column('D:D', 12, money_fmt)
            ws = writer.sheets['Net Margins Other']
            ws.set_column('D:F', 12, money_fmt)
            ws.set_column('B:B', 12, money_fmt)
            ws = writer.sheets['Rebates']
            ws.set_column('C:N', 12, money_fmt)
        writer.save()

        upload = self.pd.pivot_table(self.margins[(~self.margins["Account Name"].isna()) & (self.margins["Rebate"] > 0)], values=['Rebate', 'Count of'], index=["Account Name", "Account No", "Group Name"], aggfunc={'Rebate': np.sum, 'Count of': np.sum}, fill_value=0).reset_index().rename(columns={"Rebate": 'Group Sum'})
        
        upload.loc[upload["Group Name"] == "RSS INFRASTRUCTURE LTD" , 'Group Sum'] -= 1000
        
        upload["Month"] = month
        
        upload["Rebate end week"] = self.pd.to_datetime(max).strftime("%d/%m/%Y")
        
        upload["Group Sum"] = upload["Group Sum"].map('£{:,.2f}'.format)
        
        PSF_UPLOAD_EXCEPTIONS = []
                                 
        psfUpload = upload[~upload["Account Name"].isin(PSF_UPLOAD_EXCEPTIONS)].rename(columns={"Account No": "Account Code", "Group Name": "Merit Name", "Rebate end week":"Date to Accrue for"}).drop(columns=['Count of', "Account Name"]).reindex(columns=["Merit Name","Group Sum","Account Code", "Month", "Date to Accrue for"])

        psfUpload["Period"] = "'" + str(period) if period > 9 else "'0" + str(period)
        
        psfUpload["Year"] = self.yearAbbr.replace("-", "/")
         
        psfUpload.to_csv(rebateDir / rf"{month} py Rebates {self.yearAbbr} - psf import.csv", index=False, encoding="latin")
        
        CRM_UPLOAD_EXCEPTIONS = ["Advanced Resource Managers", "Search Consultancy Manchester","NRL Glasgow","Scantec Personnel Ltd", "Manpower", "Rullion Build Glasgow", "White Label Recruitment"]
        
        crmUpload = upload[~upload["Account Name"].isin(CRM_UPLOAD_EXCEPTIONS)].rename(columns={"Account Name": "CRM Name", 'Count of': "Total Margins", "Group Sum": "Total Amount"}).drop(columns=["Account No", "Group Name"])
        
        crmUpload["Rebate start week"] = self.pd.to_datetime(min).strftime("%d/%m/%Y")
        
        crmUpload["Total Amount"] = crmUpload["Total Amount"].apply(lambda x: float(x.replace("£", '').replace(",", '')))
        
        crmUpload["Send Rebate Info"] = False
        
        crmUpload.to_csv(rebateDir / rf"{month} py Rebates {self.yearAbbr} - crm import.csv", index=False)
        
rebates().run()
