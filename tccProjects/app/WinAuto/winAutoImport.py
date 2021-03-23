## IMPORTS ##
from pywinauto.application import Application
from pywinauto.controls.win32_controls import ButtonWrapper

import time

## CLASS ##
class WinAutoImport():

    """ """

    def __init__(self, *args, mapping_filepath=None, clear=False, **kwargs):

        self.mapping_filepath = mapping_filepath

        if clear:
            self.remedy_supervisor()

        self.login()
        self.open_mapping()

    def remedy_supervisor(self):

        self.remedy = Application(backend="uia").start('C:\\Program Files\\BMC Software\\user\\aruser.exe')

        remedy = self.remedy.window(title_re="BMC Remedy User")
        remedy['User Name:Edit'].set_text("Supervisor 1")
        remedy['Password:Edit'].set_text("super")
        remedy.Button2.click()
        time.sleep(5)
        remedy.OKButton.click()

        #print(remedy.print_control_identifiers())

        rem = self.remedy.window(title_re="BMC Remedy User - MCC Metro Control NEW (Search)")
        #remedy["Client Code"].send_keys('190')

        print(rem.print_control_identifiers())

    def login(self):

        self.app = Application(backend="uia").start('C:\Program Files\BMC Software\ARSystem\dataimporttool\launcher.exe')

        login = self.app.window(title_re="BMC Remedy Data Import - Login")
        login.edit2.set_text("super")
        login.button3.click()

    def open_mapping(self):

        imp = self.app.window(title_re="BMC Remedy Data Import")
        imp["Open an existing mapping file."].click()
        time.sleep(5)
        imp["File name:Edit"].set_text(self.mapping_filepath)
        imp.Button15.click()
        imp["Start importing records from the import file"].click()
        time.sleep(5)
        imp["OKButton"].click()
        try:
            imp.type_keys("%{F4}")
        except:
            pass
        try:
            imp.send_keys("{VK_MENU}{F4}")
        except:
            pass

        