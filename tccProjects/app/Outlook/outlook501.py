## IMPORTS ##
import win32com.client

## CLASSES ##
class Outlook501():

    def __init__(self, *args, **kwargs):

        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.get_folders()

    def get_folders(self):
        """Creates references to specific folders. Only self.inbox will be a standard folder."""

        for r in self.outlook.Folders:
            if r.name == "info@thecallcentre.co.nz":

                for box in r.Folders:
                    if box.name == "Inbox":
                        self.inbox = box

                self.kordia = r.Folders.Item("TCC CUSTOMERS").Folders.Item('K').Folders.Item("Kordia").Folders.Item("**AUTOMATION")                

    def get_email(self):
        """Returns an email from the self.inbox folder."""

        for email in self.kordia.Items:
                return email
        return False

    def check_email(self, email):
        """Returns True if an email passes certain checks. Can be over ridden."""

        pass