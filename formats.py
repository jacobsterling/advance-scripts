# -*- coding: utf-8 -*-
"""
Created on Wed Mar 30 10:53:28 2022

@author: jacob.sterling
"""




class day:
    def __init__(self):
        import datetime
        datetime.datetime.now()
        date = datetime.date
        self.today = date.today()
        
        
    def dayFormat(self, abbreviation):
        abb_dic = {'M':"Monday", "TU":"Tuesday","W":"Wednesday",'TH':"Thursday","F":"Friday","SA":"Satuday","SU":"Sunday"}
        
        if abbreviation.upper() in abb_dic.keys():
            Day = abb_dic[abbreviation.upper()]
        else:
            try:
                for key in abb_dic.keys():
                    if key.__contains__(abbreviation.upper()):
                        Day = abb_dic[key]
                        break
            except TypeError:
                Day = self.today.strftime("%A")
        return Day
        
    def dayToday(self):
        return self.today.strftime("%d/%m/%Y")
    
    def dayTodayFormat(self):
        return self.today.strftime("%d-%b-%Y")
    
    def dayTodayFormat1(self):
        return self.today.strftime("%d%m%Y")
    
    def dayPeriod(self):
        return self.today.strftime("%m")
    
class taxYear:
    def __init__(self, Week = None):
        import datetime
        datetime.datetime.now()
        date = datetime.date
        today = date.today()
        self.week = today.isocalendar()[1] if not Week else Week

        Year = today.year
        self.yearppp = Year - 3
        self.yearpp = Year - 2
        self.yearp = Year - 1
        self.yearc = Year
        self.yearcc = Year + 1
    
    def Yearpp(self, div):
        if self.week > 14:
            return (f'{self.yearpp}{div}{self.yearp}')
        else:
            return (f'{self.yearppp}{div}{self.yearpp}')
            
    def Yearp(self, div):
        if self.week  > 14:
            return (f'{self.yearp}{div}{self.yearc}')
        else:
            return (f'{self.yearpp}{div}{self.yearp}')
            
    def Year(self, div):
        if self.week > 14:
            return (f'{self.yearc}{div}{self.yearcc}')
        else:
            return (f'{self.yearp}{div}{self.yearc}')
    
    def Yearp_format1(self, div):
        if self.week > 14:
            return ('{yearp}{div}{yearc}').format(yearp = self.yearp-2000, yearc = self.yearc-2000, div = div)
        else:
            return ('{yearpp}{div}{yearp}').format(yearpp = self.yearpp-2000, yearp = self.yearp-2000, div = div)
        
    def Year_format1(self, div):
        if self.week > 14:
            return ('{yearc}{div}{yearcc}').format(yearc = self.yearc-2000, yearcc = self.yearcc-2000, div = div)
        else:
            return ('{yearp}{div}{yearc}').format(yearp = self.yearp-2000, yearc = self.yearc-2000, div = div)
    
    def Year_format2(self):
        return ('{yearc}').format(yearc = self.yearc-2000)
    
    def yearc(self):
        return self.yearc