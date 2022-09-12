# -*- coding: utf-8 -*-
"""
Created on Wed Mar 30 10:53:28 2022

@author: jacob.sterling
"""

class dayFormat:
    def __init__(self, abbreviation):
        import datetime
        datetime.datetime.now()
        date = datetime.date
        self.today = date.today()
        
        if abbreviation.upper() == 'M':
            self.Day = "Monday"
        if abbreviation.upper() == 'TU':
            self.Day = "Tuesday"
        elif abbreviation.upper() == 'W':
            self.Day = "Wednesday"
        if abbreviation.upper() == 'TH':
            self.Day = "Thursday"
        elif abbreviation.upper() == 'F':
            self.Day = "Friday"
        elif abbreviation.upper() == 'SA':
            self.Day = "Satuday"
        elif abbreviation.upper() == 'SU':
            self.Day = "Sunday"
        elif abbreviation.upper() == 'T':
            self.Day = "Tuesday"
        else:
            self.Day = self.today.strftime("%A")
            
    def Day(self):
        return self.Day
    
class taxYear:
    def __init__(self, div):
        import datetime
        datetime.datetime.now()
        self.div = div
        date = datetime.date
        self.today = date.today()
        Year = self.today.year
        self.yearppp = Year - 3
        self.yearpp = Year - 2
        self.yearp = Year - 1
        self.yearc = Year
        self.yearcc = Year + 1
    
    def Yearpp(self):
        if self.today.isocalendar()[1] > 39:
            return (f'{self.yearpp}{self.div}{self.yearp}')
        else:
            return (f'{self.yearppp}{self.div}{self.yearpp}')
            
    def Yearp(self):
        if self.today.isocalendar()[1] > 39:
            return (f'{self.yearp}{self.div}{self.yearc}')
        else:
            return (f'{self.yearpp}{self.div}{self.yearp}')
            
    def Year(self):
        if self.today.isocalendar()[1] > 39:
            return (f'{self.yearc}{self.div}{self.yearcc}')
        else:
            return (f'{self.yearp}{self.div}{self.yearc}')