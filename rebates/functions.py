# -*- coding: utf-8 -*-
"""
Created on Wed Mar 30 12:47:04 2022

@author: jacob.sterling
"""

def PAYNO_Check(payno):
    try:
        int(payno)
        return True
    except ValueError:
        return False

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
        from formats import taxYear 
        import datetime as datetime
        self.datetime = datetime
        self.taxYearRange = taxYear().Year('-')
        self.pd = __import__('pandas')
        self.date = datetime.date
        self.timedelta = datetime.timedelta
        self.today = self.date.today()
        
    def tax_week_calc(self, d = None):
        d = self.today if not d else d
        year1, year2 = self.taxYearRange.split('-')
        try:
            d = int(d)
            weekRange = self.pd.period_range(start = f'{year1}-01-01', end = f'{year2}-12-01', freq = 'W')
            df = self.pd.DataFrame(range(1,len(weekRange)+1,1),columns=['Week'], index = weekRange).reset_index().rename(columns={'index':'Range'})
            date = self.pd.to_datetime(df.loc[df['Week'] == d,'Range'].values[0])
        except TypeError:
            date = self.pd.to_datetime(d)
        weekRange = self.pd.period_range(start = f'{year1}-04-06', end = f'{year2}-04-05', freq = 'W')
        df = self.pd.DataFrame(range(1,len(weekRange)+1,1),columns=['Week'], index = weekRange).reset_index().rename(columns={'index':'Range'})
        df[['RangeS','RangeE']] = df['Range'].astype(str).str.split('/',expand=True)
        df['RangeE'] = df['RangeE'].apply(lambda x: self.datetime.datetime.strptime(x, '%Y-%m-%d'))
        df['RangeS'] = df['RangeS'].apply(lambda x: self.datetime.datetime.strptime(x, '%Y-%m-%d'))
        for i, row in df.iterrows():
            if row['RangeS'] <= date <= row['RangeE']:
                return int(row['Week'])
        
    def tax_month_calc(self, d = None):
        d = self.today if not d else d
        date = self.pd.to_datetime(d) - self.timedelta(5)
        YM = date.strftime("%Y-%m")
        
        year1, year2 = self.taxYearRange.split('-')
        periodRange = self.pd.period_range(start = f'{year1}-04-06', end = f'{year2}-04-05', freq = 'M')
        df = self.pd.DataFrame(range(1,len(periodRange)+1,1),columns=['Period'], index = periodRange).reset_index().rename(columns={'index':'Range'})
        
        return df.loc[df['Range'] == f'{YM}','Period'].astype(int).values[0]
    
    @staticmethod
    def Quarter_Calc():
        quarter_range = list()
        while quarter_range == list():
            quarter = 'Q' + str(input('Enter Quarter (1, 2, 3, 4): '))
            if quarter == 'Q1':
                quarterp = 'Q4'
                quarter_range = list(range(1,14))
            elif quarter == 'Q2':
                quarterp = 'Q1'
                quarter_range = list(range(14,27))
            elif quarter == 'Q3':
                quarterp = 'Q2'
                quarter_range = list(range(27,40))
            elif quarter == 'Q4':
                quarterp = 'Q3'
                quarter_range = list(range(40,53))
            else:
                print('Incorrect input.')
        return quarter_range, quarter, quarterp
    

