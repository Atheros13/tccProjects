## IMPORTS ##
import win32com.client

## CLASSES ##
class Outlook():

    def __init__(self, *args, **kwargs):

        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.get_folders()

    def get_folders(self):
        """Creates references to specific folders. Only self.inbox will be a standard folder."""

        self.inbox = []

    def get_email(self):
        """Returns an email from the self.inbox folder."""

        for email in self.inbox.Items.reverse():
            if self.check_email(email):
                return email

    def check_email(self, email):
        """Returns True if an email passes certain checks. Can be over ridden."""

        return True