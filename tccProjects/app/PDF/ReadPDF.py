## IMPORTS ##
from tika import parser

## CLASS ##
class ReadPDF():

    def __init__(self, filepath, filename, *args, **kwargs):

        pass


## TEST ##
parsed = str(parser.from_file('G:\\Michael Atheros - Work\\Frucor.PDF')["content"]).split()
print(parsed.split("\n"))