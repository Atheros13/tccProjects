## IMPORTS ##
from app.Outlook.outlook import Outlook

## CLASSES ##
class Outlook438(Outlook):

    def get_folders(self):

        for r in self.outlook.Folders:
            if r.name == "info@thecallcentre.co.nz":
                
                for box in r.Folders:
                    if box.name == "Inbox":
                        self.inbox = box

                dvs = r.Folders.Item("TCC CUSTOMERS").Folders.Item('D').Folders.Item("DVS (438)")
                self.sales = dvs.Folders.Item("* WEBSITE - Sales Leads (New & Upgrades)")
                self.technical = dvs.Folders.Item("* WEBSITE - Technical Issues")
                self.filters = dvs.Folders.Item("* WEBSITE - Filter/Service Enquiries")
                self.general = dvs.Folders.Item("* WEBSITE - General Enquiries")
                self.new_build = dvs.Folders.Item("* WEBSITE - New Build Enquiries")
                self.spare_parts = dvs.Folders.Item("* WEBSITE - Spare Parts Enquiries")

    def check_email(self, email):

        dvs_subjects = ['New submission from Free Consultation - Homepage', 'New submission from DVS Contact Us',
                        'New submission from Book A Consultation', 'New submission from Book Service']

        if email.Subject in dvs_subjects:
            return True
