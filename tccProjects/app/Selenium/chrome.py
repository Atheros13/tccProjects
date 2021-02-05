## IMPORTS
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

## CLASS
class ChromeDriver():

    def __init__(self, *args, **kwargs):

        chromeOptions = webdriver.ChromeOptions() 
        chromeOptions.add_argument("--remote-debugging-port=9222") 

        self.driver = webdriver.Chrome("G:\\Automation\\Selenium\\chromedriver.exe", chrome_options=chromeOptions) #ChromeDriverManager().install())
