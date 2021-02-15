from pywinauto.application import Application
from time import sleep
from pywinauto.controls.win32_controls import ButtonWrapper

app = Application(backend="uia").start('C:\Program Files\BMC Software\ARSystem\dataimporttool\launcher.exe')

login = app.window(title_re="BMC Remedy Data Import - Login")
login.edit2.set_text("super")
login.button3.click()

imp = app.window(title_re="BMC Remedy Data Import")
imp["Open an existing mapping file."].click()

imp["File name:Edit"].set_text("test")