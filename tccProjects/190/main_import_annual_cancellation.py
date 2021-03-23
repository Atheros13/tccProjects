##


## IMPORTS ##
import os

from app.WinAuto.winAutoImport import WinAutoImport

## CLASS ##
class Main():

    def __init__(self, *args, **kwargs):

        # if there is a csv  
        if not os.path.exists("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\EMAIL\\csv_import.csv"):
            return

        # import data
        WinAutoImport(mapping_filepath="G:\Customer Reporting\\190 - NZCAR\Automation\CODE\IMPORT\email_cancellation.armx")

        # delete csv
        os.remove("G:\\Customer Reporting\\190 - NZCAR\\Automation\\CODE\\EMAIL\\csv_import.csv")

## ENGINE ##
Main()