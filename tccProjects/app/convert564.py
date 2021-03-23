#-----------------------------------------------------------------------------#

## IMPORTS ##

import xlrd
import calendar
import datetime
import csv

#-----------------------------------------------------------------------------#

## CLASS ##

class Convert564():

    def __init__(self, *args, **kwargs):

        self.roster = self.open_workbook()
        #print(self.roster)

        self.create_csv()

    def open_workbook(self):
        
        wb = xlrd.open_workbook('G:/Customer Reporting/564 & 470 -Pacific Radiology/Latest Roster 564.xlsx')
        sheet = wb.sheet_by_index(0)

        heading_row = self.get_heading_row(sheet)

        areas = [
            'WHANGANUI', 'TARANAKI', 'NELSON', 'OAMARU & DUNSTAN', 'ALL OTHERS'
            ]

        roster = {}
        for i in range(len(areas)):

            column = i + 2
            area = areas[i]

            area_roster = self.extract_sheet_data(sheet, heading_row, column)
            roster[area] = area_roster

        return roster

    def get_heading_row(self, sheet):

        for r in range(sheet.nrows):
            if sheet.cell_value(r,0) == 'Date - ':
                return r

    def extract_sheet_data(self, sheet, heading_row, column):

        #        
        roster = []

        #
        date = None
        for r in range(heading_row + 1, sheet.nrows+1):
            
            try:
                time = sheet.cell_value(r,1)
            except:
                break

            if time == '12am-7am':
                date = datetime.datetime(*xlrd.xldate_as_tuple(sheet.cell_value(r+2,0), 0))

            role, time = self.convertTime(date, time)

            staff = sheet.cell_value(r, column)

            roster.append([time, staff, role])
        
        return roster

    def convertTime(self, date, time):

        role = 'Radiologist'
        try:
            t = time.replace(' ', '').split('-')[1]
        except:
            t = time
        t = t.replace('.', '')

        if 'am' in t:
            t = t.split('am')[0]
            if int(t) < 12:
                t = datetime.timedelta(hours=int(t))
            elif '30' in t:
                hour = int(t.split('30')[0])
                t = datetime.timedelta(hours=hour, minutes=30)
        elif 'pm' in t:
            t = t.split('pm')[0]
            if int(t) < 12:
                t = datetime.timedelta(hours=int(t)+12)
            elif '30' in t:
                hour = int(t.split('30')[0]) + 12
                t = datetime.timedelta(hours=hour, minutes=30)
        else:
            t = datetime.timedelta(hours=23, minutes=59, seconds=59)
            
        print(t)
        dt = date + t
        dt_string = dt.strftime("%d/%m/%y %I:%M:00 %p")

        return role, dt_string

    def convertTime1(self, date, time):

        role = 'Radiologist'

        t = time.replace(' ', '').split('-')[1]
        t = t.replace('.', '')
        if t == '7am':
            t = datetime.timedelta(hours=7)
        elif t == '8am':
            t = datetime.timedelta(hours=8)
        elif t == '830am':
            t = datetime.timedelta(hours=8, minutes=30)
        elif t == '330pm':
            t = datetime.timedelta(hours=15, minutes=30)
        elif t == '5pm':
            t = datetime.timedelta(hours=17)
        elif t == '9pm':
            t = datetime.timedelta(hours=21)
        elif t == '10pm':
            t = datetime.timedelta(hours=22)
        elif t == 'midnight':
            t = datetime.timedelta(hours=23, minutes=59, seconds=59)
        else:
            if 'all' in t:
                role = time
                if '10pm' in t:
                    t = datetime.timedelta(hours=22)
                else:
                    t = datetime.timedelta(days=1,hours=7)
            else:
                t = datetime.timedelta(hours=23, minutes=59, seconds=59)

        dt = date + t
        dt_string = dt.strftime("%d/%m/%y %I:%M:00 %p")

        return role, dt_string

    def create_csv(self):

        # region service endtime name

        try:
            open('G:/Customer Reporting/564 & 470 -Pacific Radiology/Michael Import/564Roster.csv', 'w').close()
        except:
            pass

        csv_file = open('G:/Customer Reporting/564 & 470 -Pacific Radiology/Michael Import/564Roster.csv', 'w', newline='')
        writer = csv.writer(csv_file, delimiter=',')
        writer.writerow(['REGION', 'ROLE', 'DATETIME', 'NAME'])

        for region in self.roster:
            for row in self.roster[region]:
                writer.writerow([region, row[2], row[0], row[1]])
        csv_file.close()

#-----------------------------------------------------------------------------#

Convert564()

#-----------------------------------------------------------------------------#