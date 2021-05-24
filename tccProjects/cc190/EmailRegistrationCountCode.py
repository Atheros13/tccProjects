## IMPORTS 
import sys

from cc190.EmailCount import EmailRegistrationCount


## CLASS ##

class Main():

    argv = sys.argv

    date_from = self.convert_date(argv[1])
    date_to = self.convert_date(argv[2])

    EmailRegistrationCount(date_from=date_from, date_to=date_to, no_spca=True)

    def convert_date(self, date):

        d = date.split("/")
        if len(d[0]) == 1:
            d[0] = '0%s' % d[0]
        return "%s-%s-%s" % (d[2], d[1], d[0])

## ENGINE ##
Main()