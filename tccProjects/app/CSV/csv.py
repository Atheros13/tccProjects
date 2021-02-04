## IMPORTS ##
import csv
import os

## CLASSES ##
class CSV():

    def __init__(self, filepath, filename, open_as, headings=[], *args, **kwargs):

        self.filepath = filepath
        self.filename = filename
        self.headings = headings
        self.open_as = open_as

        if open_as == 'a' and filename not in os.listdir(filepath) and headings != []:
            self.build_csv()
        elif open_as == 'w' and headings != []:
            self.build_csv()

    def build_csv(self):

        with open('%s%s' % (self.filepath, self.filename), 'w', newline="") as csv_file:
            writer = csv.writer(csv_file, delimiter=',')
            writer.writerow(self.headings)
            csv_file.close()

    def writerow(self, row):

        with open('%s%s' % (self.filepath, self.filename), self.open_as, newline="") as csv_file:
            writer = csv.writer(csv_file, delimiter=',')
            writer.writerow(row)

    def writerows(self, rows):

        with open('%s%s' % (self.filepath, self.filename), self.open_as, newline="") as csv_file:
            writer = csv.writer(csv_file, delimiter=',')
            writer.writerows(rows)