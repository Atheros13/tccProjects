## IMPORTS ##
import datetime
import win32com.client

from app.Outlook.outlook import Outlook

## CLASSES ##
class Outlook438WebsiteLeads(Outlook):

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
                self.follow_up = dvs.Folders.Item("* WEBSITE - FOLLOW UP")

                if button_action == "Inbox":
                    for box in r.Folders:
                        if box.name == "Inbox" and button_action == "Inbox":
                            self.inbox = box
                elif button_action == "Follow":
                    self.inbox = self.follow_up

    def check_email(self, email):

        dvs_subjects = ['New submission from Free Consultation - Homepage', 'New submission from DVS Contact Us',
                        'New submission from Book A Consultation', 'New submission from Book Service',
                        'New submission from DVS Upgrade Enquiry']

        if email.Subject in dvs_subjects:
            return True

class Outlook438AllEmails(Outlook):

    def __init__(self, *args, **kwargs):

        self.data = {}

        self.month = self.month()
        self.year = self.year()

        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

        self.process_folders()

        for folder in self.data:
            print(folder, self.data[folder])

    ## ?
    def month(self):

        if datetime.datetime.now().month == 1:
            return 12
        else:
            return datetime.datetime.now().month - 1

    def year(self):

        if datetime.datetime.now().month == 1:
            return datetime.datetime.now().year - 1
        else:
            return datetime.datetime.now().year
    
    ## PROPERTIES
    @property
    def specific_folders(self):

        names = """ * DAILY LASSOO SUMMARIES
                    * JOHN CHING
                    * WEBSITE - AA FILTER FORWARDS
                    * WEBSITE - Banner Leads
                    * WEBSITE - Competition Replies
                    * WEBSITE - Filter/Service Enquiries
                    * WEBSITE - General Enquiries
                    * WEBSITE - Jonathan Zhou
                    * WEBSITE - Lossnay Leads & Enquiries
                    * WEBSITE - Moved Into A House with DVS
                    * WEBSITE - New Build Enquiries
                    * WEBSITE - Remittance Advice
                    * WEBSITE - Rental Property Filter/Service Requests
                    * WEBSITE - Sales Leads (New & Upgrades)
                    * WEBSITE - Spare Parts Enquiries
                    * WEBSITE - Technical Issues
                    * WEBSITE - Tests
                    * WEBSITE: GET $100 off price quoted"""

        folder_names = []
        for n in names.split("\n"):
            folder_names.append(n.strip())

        return folder_names

    @property
    def general_folders(self):

        names = """Abe Wells
                    Absolute Cool
                    Accounts
                    Adams Electrical
                    Albert Numaga
                    All Safe Eletrical
                    Althea Thompson
                    Andi Lindsay
                    Andrew Maitai
                    Andy Colvin
                    Andy Markham
                    Anna-Marie Haggerty
                    Annika Werner
                    Aotea Cromwell
                    Aotea Electrical Wanaka
                    Aotea Oamaru
                    Bevin Young
                    Brendon Watson
                    Brent Wingham
                    Brett Lindsay
                    Brian Faircloth
                    Brian Lotter
                    Brooke Roberts / Kirk
                    Bruce Hagar
                    Carey Foster
                    Climate Systems/Gloria Burrows
                    CMC Electrical
                    COMPLAINTS
                    Complete Electrical
                    CONFIRMED EMAILS
                    Corin Stephens
                    Craig August CMC Electrical
                    Craig Campbell
                    Dave Marshall
                    Debbie Gut
                    Dee Karena
                    DG Sewell
                    Digger
                    Direct Electric LTD
                    DVS TEST Emails
                    E-ACCOUNTS LOGINS
                    ElectroNet/Nicola Anderson
                    Elite Sparkie
                    Frieda Vlaardingerbroek
                    Geri
                    Gordon Gower
                    Graham (Hoppy) Hopkins
                    Graham Elliott
                    Graham Thompson
                    Greg Brown
                    Hannah Parr
                    Helix Filter
                    Helrimu
                    Hirst Electrical
                    Installers
                    Intrepid Electrical
                    Jason Campbell
                    Jennie Whalley
                    Jeremy Hay
                    Johnstone Electrical
                    Jonathan Zhuo
                    JRS TEst
                    Keith Hamilton
                    Kelly Hewson
                    Kevin Middleton
                    Kimberley Gates
                    Lara Markham
                    Laser Electrical (Kaitaia)
                    Leishman Electrical
                    Lighting Electrical
                    Liz Sandes
                    Lotter Electrical
                    Low Electrical
                    Mainline
                    Mark Everett
                    Matt Richardson
                    Maurice Blackwell
                    Mike Van Velzen
                    Monique Richardson
                    Murray Corps
                    Oceane Chouet
                    Paul Castle
                    Paul Jackson
                    Paula Whitehouse
                    Peter Fraser
                    Peter Roberts
                    Peter Thomson
                    Phil Gilchrist
                    Phil Varley
                    Pippa Jackson
                    Posthaste
                    Proven Systems
                    RE Electrical
                    Replies to Error Email May
                    REPORTS
                    Rimu Electrical
                    Roderick Magnusson
                    Roger Pick
                    Ross Unkovich
                    Scott Marshall
                    Stefan HeatCo
                    Steve Bright
                    Stuart Lougher
                    Sue Roberts
                    Sullivan & Spillane
                    Time for a New FIlter?
                    Toni Wells
                    Tony Sandes
                    Total Group
                    Tracey Key
                    Vivienne Halliwell
                    Wayne Clifford
                    Wellington Filters"""

        folder_names = []
        for n in names.split("\n"):
            folder_names.append(n.strip())

        return folder_names

    ## PROCESS
    def process_folders(self):

        for r in self.outlook.Folders:
            if r.name == "info@thecallcentre.co.nz":

                for box in r.Folders:
                    if box.name == "Inbox":
                        self.inbox = box
  
                dvs = r.Folders.Item("TCC CUSTOMERS").Folders.Item('D').Folders.Item("DVS (438)")  

                for box in dvs.Folders:

                    emails = box.Items
                    emails.Sort("[ReceivedTime]", True)

                    for email in emails:

                        try:
                            email_date = datetime.datetime.strptime(str(email.ReceivedTime).split()[0], "%Y-%m-%d")
                        except:
                            continue

                        if email_date.month == datetime.datetime.now().month and email_date.year == datetime.datetime.now().year:
                            continue
                        elif email_date.month == self.month and email_date.year == self.year:
                            if box.name not in self.data:
                                self.data[box.name] = 1
                            else:
                                self.data[box.name] += 1
                        else:
                            break







## ENGINE ##
#Outlook438AllEmails()

