# -*- coding: utf-8 -*-
"""
Created on Thu May  5 15:26:06 2022

@author: jacob.sterling
"""

# class BankUpload:
#     def __init__(self):
import datetime

start = datetime.datetime.now()

print('_______________________________________________________________________')
print('')
print('Initializing Script...')
print(start)

from pathlib import Path
from utils.formats import day
from utils.formats import taxYear
from utils.functions import tax_calcs
from remitReaders.armPayrollScript import armPayroll

import pandas as pd
import numpy as np

dayToday = day().dayTodayFormat()

dayToday1 = day().dayTodayFormat1()

yearc = taxYear().yearc

Year = taxYear().Year(" - ")

yearFormat1 = taxYear().Year_format1("/")

Week = tax_calcs().tax_week() - 1

pertempsPath = Path.home() / rf"advance.online/J Drive - Operations/Remittances and invoices/Pertemps/Tax Year {Year}/Week {Week}"

homePath = Path.home() / r"advance.online/J Drive - Finance"

branchCodesPath = homePath / r"Bank and Cash/PSF Files"

outputPath = homePath / rf"Bank and Cash/PSF Files/{yearc}"

bankPath = homePath / rf"1 Private Finance/Bank and Cash/Bank statements/Advance Contracting RBS/Advance Contracting Solutions/{yearc}"

print(bankPath)

clientsPath = Path("M:\Work\PSF\Invoice")

print('_______________________________________________________________________')
print('')
print('Reading Data...')
print(datetime.datetime.now().time())

for file in branchCodesPath.glob("*"):
    if file.name.__contains__("Accounts") and file.suffix == ".csv":
        accounts = pd.read_csv(file).drop_duplicates(subset="Description")
    if file.name.__contains__('Branch Codes') and file.suffix == ".csv":
        branchCodes = pd.read_csv(file)
    if file.name.__contains__('Worker Ref') and file.suffix == ".csv":
        workerRef = pd.read_csv(file)
        
# fileC, fileCStat = None, 0
# for file in clientsPath.glob("*"):
#     if file.name.__contains__(f"IO{Week}_Clients"):
#         fileC = file
#         fileCStat = fileC.stat().st_mtime
# clients = pd.read_csv( fileC, header = None)

fileC, fileCStat = None, 0
for file in bankPath.glob("*"):
    print(file.name)
    if file.name.__contains__(dayToday):
        fileC = file
        break
    elif file.stat().st_mtime > fileCStat:
        fileC = file
        fileCStat = fileC.stat().st_mtime

print('_______________________________________________________________________')
print('')
print('Formatting Bank Statement...')
print(datetime.datetime.now().time())

bankStatement = pd.read_csv( fileC , encoding="latin").dropna(subset=['Credit'])
bankStatement = bankStatement[["Date","Narrative #1","Narrative #2","Credit"]].rename(columns={"Credit":"Value","Narrative #1":"Description","Narrative #2":"UF1"})
bankStatement = bankStatement.merge(accounts,how = "left")

if bankStatement["Description"].str.contains("ADVANCED RESOURCE").any() or bankStatement["Description"].str.contains("OPTAMOR LIMITED").any(): 
    arm = armPayroll(Week).readPDF()

for i, row in bankStatement.iterrows():
    date = pd.to_datetime(row["Date"], format="%d/%m/%Y")
    week = tax_calcs().tax_week(date)
    match row["Description"]:
        case "ADVANCED RESOURCE" | "OPTAMOR LIMITED":
            for j, ref in arm.iterrows():
                if str(row["UF1"]) == str(ref["Remittance Ref"]):
                    print('_______________________________________________________________________')
                    print('')
                    print(f'ARM Detected, Changing Remittance Ref {row["UF1"]} to {ref["Invoice Number"]}...')
                    UF1 = ref["Invoice Number"]
        
        # case "KEEN THINKING LIMI":
        #     workerRef
        #     pass
        
        case other:
            week = "0" + str(week) if week < 10 else str(week)
            
            if not pd.isnull(row["UF1"]):
                UF1 = str(row["UF1"]) + week
            else:
                UF1 = week

    bankStatement.at[i, "UF1"] = UF1

if input("Type 'y' to merge pertemps: ") == 'y':
    import pdfplumber
    import re
    print('_______________________________________________________________________')
    print('')
    print(f'Reading Pertemps PDF Week {Week}...')
    print(datetime.datetime.now().time())
    
    result = pd.DataFrame([],columns = ["UF1","Branch Code","Surname","Forenames","Date","Rate","Value"])
    newLinePattern = re.compile(r"^([0-9]+\/[0-9A-Z]+) ([0-9]+)\/([0-9]+) (.*) ([0-9]{2}\/[0-9]{2}\/[0-9]{4}) (.*) ([0-9]+\.?[0-9]?) (£[1-9],?[0-9]*\.[0-9]{2}) (£[1-9],?[0-9]*\.[0-9]{2}) (£[1-9],?[0-9]*\.[0-9]{2}) (£[1-9],?[0-9]*\.[0-9]{2}) [0-9\-]+ [0-9]+ (Advance Contracting Solutions Ltd)$")
    nextLinePattern = re.compile("^([0-9]{2}\/[0-9]{2}\/[0-9]{4}) (.*) ([0-9]\.?[0-9]?) (£[0-9]+\.[0-9]{2}) (£[0-9]+\.[0-9]{2}) (£[0-9]+\.[0-9]{2}) (£[0-9]+\.[0-9]{2})$")
    for file in pertempsPath.glob("*"):
        if file.suffix in [".pdf",".PDF"]:
            pdf = pdfplumber.open(file)
            for page in pdf.pages:
                text = page.extract_text()
                for line in text.split("\n"):
                   
                    if newLinePattern.match(line):
                        
                        groups = re.search(newLinePattern, line).groups()
                        
                        branchCode = groups[1]
                            
                        UF1 = groups[2]
                        names = groups[3].replace(" Pertemps Medical Ltd"," ").replace(" Ventures S/E Weekly"," ").split(" ")
                        surname = names[0]
                        fornames = ""
                        for name in names[1:]:
                            if name != " ":
                                fornames = fornames + name + " "
                        fornames =  fornames[:-1]
                        
                        date = groups[4]
                        hours = groups[6]
                        rate = float(groups[7].replace(',','').replace(' ','').replace('£',''))
                        value = float(groups[8].replace(',','').replace(' ','').replace('£',''))
                        
                        dfTemp = pd.DataFrame([[UF1, branchCode, surname, fornames, date, rate ,value]],columns = result.columns)
                        result = pd.concat([result,dfTemp]).reset_index(drop = True)
                    
                    elif nextLinePattern.match(line):
                        
                        groups = re.search(nextLinePattern, line).groups()
                        
                        date = groups[0]
                        hours = groups[2]
                        rate = float(groups[3].replace(',','').replace(' ','').replace('£',''))
                        value = float(groups[4].replace(',','').replace(' ','').replace('£',''))
                        
                        dfTemp = pd.DataFrame([[UF1, branchCode, surname, fornames, date, rate ,value]],columns = result.columns)
                        result = pd.concat([result,dfTemp]).reset_index(drop = True)
    
    result["Branch Code"] = result["Branch Code"].astype(str)
    branchCodes["Branch Code"] = branchCodes["Branch Code"].astype(str)
    pertemps = pd.merge(result,branchCodes, how="left",validate="many_to_one")
    
    pertemps = pertemps.groupby(['Surname','Forenames']).agg({'Date':'first','Value':np.sum,'UF1':'first','Account':'first'}).reset_index(drop=True)
    pertemps["Description"] = "NETWORK VENTURES"
    
    bankStatement = pd.concat([bankStatement,pertemps]).reset_index(drop=True)

print('_______________________________________________________________________')
print('')
print('Exporting Import and Error File...')
print(datetime.datetime.now().time())

bankStatement["Document Type"] = "JRNL"
bankStatement["Year"] = yearFormat1
bankStatement["Period"] = bankStatement["Date"].apply(lambda x: tax_calcs().period(x, "%d/%m/%Y"))
bankStatement["Nominal"] = 5310

errorFile = bankStatement[bankStatement["Account"].isnull()]

for i, row in errorFile.iterrows():
    if row["Description"] not in ["NETWORK VENTURES"]:
        account = input(f"Enter missing account code for description; {row['Description']} (press enter to skip): ")
        if account != "":
            bankStatement.at[i, "Account"] = account
            accounts = pd.concat([accounts,pd.DataFrame([[row["Description"],account]], columns = accounts.columns)])
                                 
errorFile = errorFile.reset_index(drop = True)
errorFile = errorFile.reset_index(drop = False).rename(columns={"index":"Row No"}).reindex(columns = ["Document Type","Row No","Year","Period","Date","Nominal","Account","Value","Description","UF1"])
errorFile.to_csv(outputPath / rf"Error Py {dayToday1}.csv",index = False)

bankStatement = bankStatement[~bankStatement["Account"].isnull()]
bankStatement["Value"] = bankStatement["Value"]*-1

bankStatementRev = bankStatement[["Document Type","Year","Period","Date","Value","Account","Description","UF1"]]
bankStatementRev["Nominal"] = 5200
bankStatementRev["Account"] = ""
bankStatementRev["Value"] = bankStatementRev["Value"]*-1

bankStatement = pd.concat([bankStatement, bankStatementRev]).reset_index(drop = True)
bankStatement = bankStatement.reset_index(drop = False).rename(columns={"index":"Row No"})

bankStatement["Row No"] = bankStatement["Row No"] + 1

bankStatement = bankStatement.reindex(columns=["Document Type","Row No","Year","Period","Date","Nominal","Account","Value","Description","UF1"])

bankStatement.to_csv(outputPath / rf"Bank Upload Py {dayToday1}.csv", index = False)

accounts.to_csv(branchCodesPath / r"Accounts.csv", index = False)
branchCodes.to_csv(branchCodesPath / r"Branch Codes.csv", index = False)

print('_______________________________________________________________________')
print('')
print('Done.')


