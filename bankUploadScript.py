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
import pandas as pd
from Formats import taxYear
from Formats import day
from Functions import tax_calcs
from armPayrollScript import armPayroll
import numpy as np

dayToday = day().dayTodayFormat()

dayToday1 = day().dayTodayFormat1()

yearc = taxYear().yearc

Year = taxYear().Year(" - ")

yearFormat1 = taxYear().Year_format1("/")

Week = tax_calcs().tax_week_calc()

pertempsPath = Path.home() / rf"advance.online/J Drive - Operations/Remittances and invoices/Pertemps/Tax Year {Year}/Week {Week}"

homePath = Path.home() / r"advance.online/J Drive - Finance"

branchCodesPath = homePath / r"Bank and Cash/PSF Files"

outputPath = homePath / rf"Bank and Cash/PSF Files/{yearc}"

bankPath = homePath / rf"1 Private Finance/Bank and Cash/Bank statements/Advance Contracting RBS/Advance Contracting Solutions/{yearc}"

clientsPath = Path("M:\Work\PSF\Invoice")

print('_______________________________________________________________________')
print('')
print('Reading Data...')
print(datetime.datetime.now().time())

for file in branchCodesPath.glob("*"):
    if file.name.__contains__("Bank Actual Upload Template") and file.suffix == ".xlsm":
        template = pd.read_excel(file, sheet_name = ["Accounts","Branch Codes"])
        accounts = template['Accounts']
        branchCodes = template['Branch Codes'].rename(columns={"PS Agency Code":"Account"})
        
fileC, fileCStat = None, 0
for file in clientsPath.glob("*"):
    if file.name.__contains__(f"IO{Week}_Clients"):#file.stat().st_mtime > fileCStat and 
        fileC = file
        fileCStat = fileC.stat().st_mtime
clients = pd.read_csv( fileC, header = None)

fileC = None
for file in bankPath.glob("*"):
    if file.name.__contains__(dayToday):
        fileC = file
        break
    elif file.stat().st_mtime > fileC.stat().st_mtime:
        fileC = file

bankStatement = pd.read_csv( fileC ).dropna(subset=['Credit'])
bankStatement = bankStatement[["Date","Narrative #1","Narrative #2","Credit"]].rename(columns={"Credit":"Value","Narrative #1":"Description","Narrative #2":"UF1"})
bankStatement = bankStatement.merge(accounts,how = "left")

for i, row in bankStatement.iterrows():
    date = pd.to_datetime(row["Date"])
    
    week = tax_calcs().tax_week_calc(pd.to_datetime(date))
    
    if row["Description"] in ["ADVANCED RESOURCE","OPTAMOR LIMITED"]:
        df = armPayroll(week).readPDF()
        for j, ref in df.iterrows():
            if row["UF1"] == ref["Remittance Ref"]:
                UF1 = ref["Invoice Number"]
    else:
        week = "0" + str(week) if week < 10 else str(week)
        
        if not pd.isnull(row["UF1"]):
            UF1 = str(row["UF1"]) + week
        else:
            UF1 = week

    bankStatement.at[i, "UF1"] = UF1

if input("Type 'y' to merge pertemps: ") == 'y':
    import pdfplumber
    import re
    
    newLinePattern = re.compile(r"((^[0-9A-Z][0-9]{2})\S[A][0-9]{5}).*")
    namePattern = re.compile(r"([A-Z]{1}[a-z]+)")
    datePattern = re.compile(r"([0-9]{2}/[0-9]{2}/[0-9]{4})")
    valuesPattern = re.compile(r"([£][0-9]+[.][0-9]{2})")
    
    result = pd.DataFrame([],columns = ["UF1","Branch Code","Surname","Forenames","Date","Rate","Value","Gross"])
    
    for file in pertempsPath.glob("*"):
        if file.suffix in [".pdf",".PDF"]:
            pdf = pdfplumber.open(file)
            for page in pdf.pages:
                text = page.extract_text()
                
                for line in text.split("\n"):
                    if newLinePattern.match(line):
    
                        pertempsId = re.search(newLinePattern,line).groups()
                        
                        UF1 = pertempsId[0]
                        
                        try:
                            branchCode = int(pertempsId[-1])
                        except ValueError:
                            branchCode = pertempsId[-1]
                    try:
                        date = re.search(datePattern,line).group(1)
                    except AttributeError:
                        continue
                    
                    nameGroup = re.findall(namePattern,line)
                    
                    if nameGroup:
                        surname = nameGroup[0]
                        
                        forenames = ""
                        for forename in nameGroup[0:]:
                            forenames += forename + " "
                        forenames = forenames[:-1]
                        forenames = forenames.replace(surname+" ","").replace("Weekly ","").replace("Ventures ","").replace(" Advance Contracting Solutions Ltd","")
                    
                    values = re.findall(valuesPattern,line)
                    rate = float(values[0].replace('£',''))
                    value = float(values[1].replace('£',''))
                    gross = float(values[2].replace('£',''))
                    
                    dfTemp = pd.DataFrame([[UF1, branchCode, surname, forenames, date, rate ,value ,gross]],columns = result.columns)
                    result = pd.concat([result,dfTemp]).reset_index(drop = True)
    
    result["Branch Code"] = result["Branch Code"].astype(str)
    branchCodes["Branch Code"] = branchCodes["Branch Code"].astype(str)
    pertemps = pd.merge(result,branchCodes, how="left",validate="many_to_one")
    
    pertemps = pertemps.groupby(['Surname','Forenames']).agg({'Date':'first','Value':np.sum,'UF1':'first','Account':'first'}).reset_index(drop=True)
    pertemps["Description"] = "NETWORK VENTURES"
    
    bankStatement = pd.concat([bankStatement,pertemps]).reset_index(drop=True)

bankStatement["Document Type"] = "JRNL"
bankStatement["Year"] = yearFormat1
bankStatement["Period"] = bankStatement["Date"].apply(lambda x: pd.to_datetime(x)).dt.strftime("%m")
bankStatement["Nominal"] = 5310

errorFile = bankStatement[bankStatement["Account"].isnull()].reset_index(drop = True)
errorFile = errorFile.reset_index(drop = False).rename(columns={"index":"Row No"}).reindex(["Document Type","Row No","Year","Period","Date","Nominal","Account","Value","Description","UF1"])
errorFile.to_csv(outputPath / rf"Bank Error Py {dayToday1}",index = False)

bankStatement = bankStatement[~bankStatement["Account"].isnull()]

bankStatementRev = bankStatement
bankStatementRev["Nominal"] = 5200
bankStatementRev["Value"] = bankStatementRev["Value"]*-1

bankStatement = pd.concat([bankStatement, bankStatementRev]).reset_index(drop = True)
bankStatement = bankStatement.reset_index(drop = False).rename(columns={"index":"Row No"})

bankStatement = bankStatement.reindex(["Document Type","Row No","Year","Period","Date","Nominal","Account","Value","Description","UF1"])

bankStatement.to_csv(outputPath / rf"Bank Upload Py {dayToday1}.csv", index = False)

