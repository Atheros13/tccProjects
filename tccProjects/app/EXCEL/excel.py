#-----------------------------------------------------------------------------#

## IMPORTS ##

import xlrd
import calendar
import datetime
import csv

#-----------------------------------------------------------------------------#

## CLASS ##

class Excel():

    def __init__(self, filepath, filename, *args, **kwargs):

        self.filepath = filepath
        self.filename = filename

        self.wb = xlrd.open_workbook('%s%s' % (filepath, filename))
        self.sheet = self.wb.sheet_by_index(0)

