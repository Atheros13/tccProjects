## IMPORTS ##
import win32com.client
import time

## CLASSES ##
class OutlookFolderSearch():

    def __init__(self, folder_name, *args, **kwargs):

        self.folder_name = folder_name

        self.outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        for folder in self.outlook.Folders:
            if folder.name == "info@thecallcentre.co.nz":    
                result = self.search_folders(folder)
                if result != None:
                    result += "-> %s" % folder.name
                    print(result)
                else:
                    print("FOLDER NOT FOUND - Check spelling of folder")
                time.sleep(360)
                    

    def search_folders(self, folder):
        """Creates references to specific folders. Only self.inbox will be a standard folder."""

        for f in folder.Folders:
            print(f.name)
            if self.folder_name in f.name:
                return "FOUND: %s " % f.name
            result = self.search_folders(f)
            if result != None:
                result += "-> %s" % f.name
                return result

## ENGINE ##
if len(sys.argv) > 1:

    folder_name = " ".join(sys.argv[1:])
    OutlookFolderSearch(folder_name)