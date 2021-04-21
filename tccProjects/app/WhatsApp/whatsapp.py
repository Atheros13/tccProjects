## IMPORTS ##
import datetime
import pywhatkit


## CLASS ##
class WhatsAppAutomation():

    def __init__(self, *args, **kwargs):

        self.number_to = "+64226472984"
        self.message = 'Test one, two, three... anything but THAT.'

        self.send_message()

    def send_message(self, hour=None, minute=None):
        """ """

        # Sets the hour and minute values if they are not provided
        now = datetime.datetime.now()
        if hour == None:
            if now.minute >= 57:
                hour = now.hour + 1
                minute = 0
            else:
                hour = now.hour
                minute = now.minute + 1

        # 
        pywhatkit.sendwhatmsg(self.number_to, self.message, hour, minute)



## ENGINE ##
WhatsAppAutomation()
