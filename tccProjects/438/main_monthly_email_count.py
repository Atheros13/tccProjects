from app.Outlook.outlook438 import Outlook438AllEmails as Outlook438

import calendar
import time
import webbrowser

class Main():

    def __init__(self, *args, **kwargs):

        self.outlook = Outlook438()
        self.open_notepad()

    def open_notepad(self):
        """Writes a .txt file of the email data and opens the file."""

        notepad = "Monthly Email Count - %s\n\n" % calendar.month_name[self.outlook.month]

        for folder in self.outlook.data:
            notepad += "%s - %s\n" % (folder, self.outlook.data[folder])

        file = open('G:\\Customer Reporting\\438-DVS\\MONTHlY Reports\\Automation_Monthly_Email_Count.txt', 'w')
        file.write(notepad)
        file.close()

        webbrowser.open('G:\\Customer Reporting\\438-DVS\\MONTHlY Reports\\Automation_Monthly_Email_Count.txt')



## ENGINE ##
Main()