"""

"""

## PYTHON IMPORTS ##
import sys
import webbrowser
import win32com.client

## APP IMPORTS ##
from app.Outlook.outlook461 import Outlook461 as Outlook

## CLASSES ##

class OpenJobEmail():

    def __init__(self, *args, **kwargs):

        self.outlook = Outlook()
        self.email = self.outlook.get_email()
        if self.email != False:
            self.process_email()

    def process_email(self):

        check = True
        for att in self.email.Attachments:
            if '.PDF' in att.FileName or ".pdf" in att.FileName:
                check = False
                att.SaveAsFile("G:\\Customer Reporting\\461-Independent Cellar Services\\Automation\\WorkOrder.PDF")
                self.email.FlagRequest = "Mark Complete"
                self.email.Move(self.outlook.frucor_jobs)
                webbrowser.open("G:\\Customer Reporting\\461-Independent Cellar Services\\Automation\\WorkOrder.PDF")
                sys.stdout.write("G:\\Customer Reporting\\461-Independent Cellar Services\\Automation\\WorkOrder.PDF")
    
        if check:            
            self.email.SaveAs("G:\\Customer Reporting\\461-Independent Cellar Services\\Automation\\WorkOrder.msg")
            self.email.FlagRequest = "Mark Complete"
            self.email.Move(self.outlook.frucor_jobs)
            webbrowser.open("G:\\Customer Reporting\\461-Independent Cellar Services\\Automation\\WorkOrder.msg")
            sys.stdout.write("G:\\Customer Reporting\\461-Independent Cellar Services\\Automation\\WorkOrder.msg")

## ENGINE ##
OpenJobEmail()
