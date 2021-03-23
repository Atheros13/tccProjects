## IMPORTS ##
from pywinauto.application import Application
from pywinauto.controls.win32_controls import ButtonWrapper

import time

## CLASS ##
class WinAutoRemedy():

    """ """

    def __init__(self, *args, name=None, password=None, **kwargs):

        if name != None:
            self.name = name
            self.password = password
        else:
            self.name = "Supervisor 1"
            self.password = "super"

        self.login()

    def login(self):

        self.remedy = Application(backend="uia").start('C:\\Program Files\\BMC Software\\user\\aruser.exe')
        remedy = self.remedy.window(title_re="BMC Remedy User")
        remedy['User Name:Edit'].set_text(self.name)
        remedy['Password:Edit'].set_text(self.password)
        remedy.Button2.click()
        time.sleep(5)
        remedy.OKButton.click()
