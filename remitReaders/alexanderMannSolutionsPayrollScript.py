# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
from pathlib import Path
from tabula import read_pdf
from Formats import day
from Formats import taxYear
import pdfplumber
import re


Week = input('Enter Week Number: ')
Day = day().dayFormat(input('Enter "m" for Monday or "f" for Friday: '))
Year = taxYear().Year(' - ')
    
file_path = Path(rf'C:\Users\jacob.sterling\advance.online\J Drive - Operations\Remittances and invoices\Alexander Mann\Tax Year {Year}\Week {Week}')

file_path = file_path / fr"{Day}" if Day == "Friday" else file_path

result = pd.DataFrame([],columns = ["Worker Name","Timesheet","Description","Rate","Total"])

newLinePattern = re.compile(r"(([A-Z][a-zA-Z]+.)+)([A-Z][a-zA-Z]+.) (\[([A-Z0-9]+)\])")
datePattern = re.compile(r"([0-9]{2}/[0-9]{2}/[0-9]{4}){1}")
timesheetPattern = re.compile(r"([A-Z]{2}[0-9]+)")
totalPattern = re.compile(r"(.*)( [1-9],?[0-9]+\.[0-9]{2}) [A-Z]$")
valuesPattern = re.compile(r".([1-9],?([0-9]+)?\.[0-9]{2})")
extraTotalPattern = re.compile(r"(.*)([0-9]+?\.[0-9]{2}) [A-Z]$")

for pdf in file_path.glob('*'):
    if pdf.suffix in [".PDF",".pdf"] and (pdf.name).__contains__('SELFBILL_'):
        
        
def readPDF(self, pdf)  
    print(f'Reading {pdf.name}......')
    pdf = pdfplumber.open(pdf)
    for page in pdf.pages:
        text = page.extract_text()
        for line in text.split("\n"):
            if newLinePattern.match(line):
                nameGroup = re.search(newLinePattern,line).groups()
                timesheet = nameGroup[-1]
                name = line.replace(timesheet,"").replace("[","").replace("]","")
                
            if datePattern.match(line):
                date = re.findall(datePattern,line)[0]
            
            if totalPattern.match(line) or extraTotalPattern.match(line):
                
                if totalPattern.match(line):
                    amountGroup = re.search(totalPattern,line).groups()
                    amount = float(amountGroup[1].replace(",","").replace(" ",""))
                    desc = amountGroup[0].replace(date + " ","")
                    
                    rates = re.findall(valuesPattern,desc)
                    rate = float(rates[-1][0].replace(",","").replace(" ",""))
                    
                    desc = desc.replace(rates[-1][0],"")
                elif extraTotalPattern.match(line):
                    amountGroup = re.search(extraTotalPattern,line).groups()
                    amount = float(amountGroup[1].replace(",","").replace(" ",""))
                    desc = amountGroup[0]
                    rate = 0
                    
                # "Daily Rate" if rate.__contains__("STD") else "Company Income"
                for prefix in ["MR","MRS","MISS","MS"]:
                    if prefix == name.split(" ")[0].upper():
                        name = name.replace(prefix + " ","")
                
                result = pd.concat([result, pd.DataFrame([[name[:-1] if name[-1] == "" else name,
                                                           date,
                                                           desc,
                                                           rate,
                                                           amount]],
                                                         columns=result.columns)]).reset_index(drop=True)

result.to_csv(file_path / rf"AXM Py Import {Day} Week {Week}.csv")
                        
