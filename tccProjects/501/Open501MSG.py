"""

"""

## PYTHON IMPORTS ##
import sys
import webbrowser
import win32com.client

## APP IMPORTS ##
from app.Outlook.outlook501 import Outlook501 as Outlook

## CLASSES ##

class Open501MSG():

    def __init__(self, *args, **kwargs):

        self.outlook = Outlook()
        self.email = self.outlook.get_email()
        if self.email != False:
            self.process_email()

    def process_email(self):

        filename = "G:\\Customer Reporting\\501-Kordia\\Automation\\KordiaEmail.msg"

        try:
            self.email.SaveAs(filename)
        except:
            count = 1
            while True:
                try:
                    filename = "G:\\Customer Reporting\\501-Kordia\\Automation\\KordiaEmail%s.msg" % count
                    self.email.SaveAs(filename)
                    break
                except:
                    count += 1


        self.email.Subject = "ACTIONED - %s" % self.email.Subject
        self.email.FlagRequest = "Mark Complete"
        self.email.Move(self.outlook.inbox)

        webbrowser.open(filename)
        sys.stdout.write(filename)

## ENGINE ##
Open501MSG()