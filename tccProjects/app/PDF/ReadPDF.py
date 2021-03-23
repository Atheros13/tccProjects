## IMPORTS ##
from tika import parser

## CLASS ##
class ReadPDF():

    def __init__(self, filepath, filename, *args, **kwargs):

        self.content = str(parser.from_file("%s%s" % (filepath, filename))["content"])

    def build_content_lines():

        lines = []
        for line in self.content.split("\n"):
            if line not in ["", " "]:
                lines.append(line)
        return lines
