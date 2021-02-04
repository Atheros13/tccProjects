## IMPORTS
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

## CLASS
class ChromeDriver():

    def __init__(self, *args, **kwargs):

        self.driver = webdriver.Chrome("G:\\Michael Atheros Work\\Automation\\chromedriver.exe")
