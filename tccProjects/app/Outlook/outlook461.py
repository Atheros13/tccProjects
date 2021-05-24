## IMPORTS ##
import win32com.client

## CLASSES ##
class Outlook461():

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

                ics = r.Folders.Item("TCC CUSTOMERS").Folders.Item('I').Folders.Item("Independent Cellar Services (461 & 811)")                
                self.frucor_jobs = ics.Folders.Item("Frucor Jobs")
                self.frucor_job_summaries = ics.Folders.Item("Frucor Job Summaries")

    def get_email(self):
        """Returns an email from the self.inbox folder."""

        for email in self.inbox.Items:
            if self.check_email(email):
                return email
        return False

    def check_email(self, email):
        """Returns True if an email passes certain checks. Can be over ridden."""

        if email.Subject[:25] == "Frucor Chiller Service WO":
            return True

        email_address = None
        if email.SenderEmailType=='EX':
            email_address = email.Sender.GetExchangeUser().PrimarySmtpAddress
        else:
            email_address = email.SenderEmailAddress
        if email_address == "admin.ak@answerservices.co.nz":
            return True