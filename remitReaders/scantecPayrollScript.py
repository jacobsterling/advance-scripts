# -*- coding: utf-8 -*-
"""
Created on Thu Feb 24 09:44:11 2022

@author: jacob.sterling
"""

import pandas as pd
from pathlib import Path
import pdfplumber
import re
from utils import formats

Year = formats.taxYear().Year(' - ')

homePath = Path.home() / r"advance.online"

filePath = homePath / rf"J Drive - Operations/Remittances and invoices/Scantec/Tax Year {Year}"

def runWeek(Week: int = None):
    Week = input("Enter week number: ") if not Week else Week

    result = pd.DataFrame([], columns = ["Worker Name","UF1","Description", "Hours", "Rate", "Amount", "PDF Name"])
    totals = pd.DataFrame([], columns = ["Day", "PDF Name", "PDF Total", "Calulated Total"])
    
    days = formats.day.abb_dic
    
    for d in days.keys():
        df, total = runDay(d, Week)
        result = pd.concat([result,df])
        total["Day"] = days[d]
        totals = pd.concat([totals, total])
    
    totals["Difference"] = totals["PDF Total"] - totals["Calulated Total"]
    
    return result, totals

def runDay(Day: str = None, Week: int = None):
    Week = input("Enter week number: ") if not Week else Week
    Day = input("Enter 'TU' for teusday or 'F' for friday") if not Day else Day
    Day = formats.day().dayFormat(Day)
    
    dayPath = filePath / rf"Week {Week}/{Day}"
    
    result = pd.DataFrame([], columns = ["Worker Name","UF1","Description", "Hours", "Rate", "Amount", "PDF Name"])
    totals = pd.DataFrame([], columns = ["PDF Name", "PDF Total","Calulated Total"])
    
    for file in dayPath.glob('*'):
        if file.is_file() and file.suffix in [".PDF",".pdf"]:
            df, total = scantecPDFReader(file)
            result = pd.concat([result,df])
            temp = pd.DataFrame([[file.name, total, df["Amount"].sum()]], columns = totals.columns)
            totals = pd.concat([totals, temp])
        
    return result, totals

def scantecPDFReader(file):
    newLinePattern = re.compile(r"^(.*) [0-9]?\/?\[([0-9\/]+)\] ?([0-9\/]+)? ([A-Z0-9]+) (.* )?([1-9][0-9]*\.[0-9]{2}) ([0-9],?[0-9]*\.[0-9]{2}) ([1-9],?[0-9]*\.[0-9]{2})[A-Z][1]$")
    totalPattern = re.compile(r"^[A-Z][0-9] ([1-9]+,?[0-9]*\.[0-9]{2}) ([0-9]+\.[0-9]{2}%) ([1-9]+,?[0-9]*\.[0-9]{2}) Total ([1-9]+,?[0-9]*\.[0-9]{2})")
    
    result = pd.DataFrame([], columns = ["Worker Name","UF1","Description", "Hours", "Rate", "Amount", "PDF Name"])
    total = 0

    pdf = pdfplumber.open(file)
    print(rf"Reading {file.name}....")
    for page in pdf.pages:
        text = page.extract_text()
        for line in text.split("\n"):
            if newLinePattern.match(line):
                groups = re.search(newLinePattern, line).groups()
                name = groups[0].replace("MR ","").replace("MRS ","").replace("MISS ","")
                UF1 = groups[1] if not groups[1].__contains__("/") else "NA"
                date = groups[2]
                timesheet = groups[3]
                description = groups[4]
                hours = float(groups[5].replace(",",""))
                rate = float(groups[6].replace(",",""))
                amount = float(groups[7].replace(",",""))
                result = pd.concat([result, pd.DataFrame([[name,
                                                           UF1,
                                                           description,
                                                           hours,
                                                           rate,
                                                           amount,
                                                           file.name
                    ]], columns = result.columns)]).reset_index(drop=True)
            elif totalPattern.match(line):
                 total = float(re.search(totalPattern, line).groups()[-1].replace(",","").replace(" ",""))
            else:
                print(line)
    return result, total
