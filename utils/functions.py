# -*- coding: utf-8 -*-
"""
Created on Wed Mar 30 12:47:04 2022

@author: jacob.sterling
"""
def age(birthdate):
    from datetime import date
    DOB = str(birthdate).split('/')
    today = date.today()
    age = today.year - int(DOB[2]) - ((today.month, today.day) < (int(DOB[1]), int(DOB[0])))
    return age

def PAYNO_Check(payno):
    try:
        int(payno)
        return True
    except ValueError:
        return False

def PAYNO_Convert(payno):
    import numpy as np
    try:
        return int(payno)
    except ValueError:
        return np.nan
    
def check_index(item, l):
    n = 0
    for k in l:
        if item in k:
            return n
        n += 1
    return None

def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)

class tax_calcs:
    def __init__(self):
        from utils.formats import taxYear
        import datetime
        self.datetime = datetime.datetime
        self.taxYearRange = taxYear().Year('-')
        self.pd = __import__('pandas')
        self.date = datetime.date
        self.timedelta = datetime.timedelta
        self.today = self.date.today()
    
    def chqdate(self, w: int):
        
        df = self.tax_week_map()
        
        rs, re = df.loc[df['Week'] == w, 'Range'].astype(str).values[0].split('/')
        
        for period in self.pd.period_range(start = rs, end = re, freq = 'D'):
            date = self.datetime.strptime(str(period), '%Y-%m-%d')
            if date.weekday() == 4:
                return date
    
    def from_iso(self, w : int):
        year1, year2 = self.taxYearRange.split('-')
        weekRange = self.pd.period_range(start = f'{year1}-01-01', end = f'{year2}-12-01', freq = 'W')
        df = self.pd.DataFrame(range(1,len(weekRange)+1,1),columns=['Week'], index = weekRange).reset_index().rename(columns={'index':'Range'})
        date = self.pd.to_datetime(df.loc[df['Week'] == w,'Range'].values[0])
        return self.tax_week(date)
    
    def tax_week(self, d = None):
        date = self.pd.to_datetime(d) if d else self.today
        
        df = self.tax_week_map()
        
        return df.loc[(df['RangeS'].dt.date <= date) & (df['RangeE'].dt.date >= date), 'Week'].astype(int).values[0]
    
    def tax_week_map(self):
        import math
        
        year1, year2 = self.taxYearRange.split('-')
        weekRange = self.pd.period_range(start = f'{year1}-04-06', end = f'{year2}-04-05', freq = 'W')
        df = self.pd.DataFrame(range(1,len(weekRange)+1,1),columns=['Week'], index = weekRange).reset_index().rename(columns={'index':'Range'})
        df[['RangeS','RangeE']] = df['Range'].astype(str).str.split('/',expand=True)
        df['RangeE'] = df['RangeE'].apply(lambda x: self.datetime.strptime(x, '%Y-%m-%d'))
        df['RangeS'] = df['RangeS'].apply(lambda x: self.datetime.strptime(x, '%Y-%m-%d'))
        
        df['Period'] = df['RangeE'].dt.strftime("%m").astype(int)
        df['Fiscal Period'] = df["Week"].apply(lambda x: math.ceil(x / 4) if math.ceil(x / 4) < 13 else math.ceil(x / 4) - 12)
        
        df.loc[ df['Week'].isin(range(1,14)), "Quarter" ] = "Q1"
        df.loc[ df['Week'].isin(range(14,27)), "Quarter" ] = "Q2"
        df.loc[ df['Week'].isin(range(27,40)), "Quarter" ] = "Q3"
        df.loc[ df['Week'].isin(range(40,54)), "Quarter"] = "Q4"

        return df
        
    def period(self, d = None, frt: str = None):
        date = self.pd.to_datetime(d, format = frt) if d else self.today
        
        df = self.tax_week_map()
        
        return df.loc[(df['RangeS'].dt.date <= date) & (df['RangeE'].dt.date >= date), 'Fiscal Period'].astype(int).values[0]
