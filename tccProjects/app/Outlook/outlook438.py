## IMPORTS ##
import win32com.client

from app.Outlook.outlook import Outlook

## CLASSES ##
class Outlook438(Outlook):

    def __init__(self, button_action, *args, **kwargs):

        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        self.get_folders(button_action)

    def get_folders(self, button_action):

        for r in self.outlook.Folders:
            if r.name == "info@thecallcentre.co.nz":
  
                dvs = r.Folders.Item("TCC CUSTOMERS").Folders.Item('D').Folders.Item("DVS (438)")                
                self.sales = dvs.Folders.Item("* WEBSITE - Sales Leads (New & Upgrades)")
                self.technical = dvs.Folders.Item("* WEBSITE - Technical Issues")
                self.filters = dvs.Folders.Item("* WEBSITE - Filter/Service Enquiries")
                self.general = dvs.Folders.Item("* WEBSITE - General Enquiries")
                self.new_build = dvs.Folders.Item("* WEBSITE - New Build Enquiries")
                self.spare_parts = dvs.Folders.Item("* WEBSITE - Spare Parts Enquiries")
                self.follow_up = self.sales.Folders.Item("FOLLOW UP")

                if button_action == "Inbox":
                    for box in r.Folders:
                        if box.name == "Inbox" and button_action == "Inbox":
                            self.inbox = box
                elif button_action == "Follow":
                    self.inbox = self.follow_up

    def check_email(self, email):

        dvs_subjects = ['New submission from Free Consultation - Homepage', 'New submission from DVS Contact Us',
                        'New submission from Book A Consultation', 'New submission from Book Service']

        if email.Subject in dvs_subjects:
            return True

