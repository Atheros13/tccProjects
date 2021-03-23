## IMPORTS ##
import datetime


## APP IMPORTS ##
from app.PDF.readPDF import ReadPDF


## CLASS ##
class Main():

    def __init__(self, *args, language=None, timedate=None, earliest=None, latest=None, duration=None, **kwargs):

        #
        self.language = language
        self.timedate = self.convert_timedate(timedate)
        self.earliest = self.convert_timedate(earliest)
        self.latest = self.convert_timedate(latest)
        self.duration = self.convert_duration(duration)

        #
        self.languages = self.build_languages()
        self.language_groups = self.build_language_groups()
        self.interpreters = self.build_interpreters()
        self.bookings = self.build_bookings()

        # ENGINE
        if language != None:
            results = self.process()

        # TEST
        if False:
            self.test()

    ## CONVERT ##
    def convert_timedate(self, timedate):

        '''Returns a str of a date & time to a datetime object, or None if None provided.'''

        if timedate == None:
            return None

        # FORMAT EXAMPLE "16/03/2021 1:19:26 pm"
        timedate = timedate.replace("am", "AM")
        timedate = timedate.replace("pm", "PM")
        return datetime.datetime.strptime(timedate, '%d/%m/%Y %-I:%M:%S %p')

    def convert_duration(self, duration):

        '''Returns a duration as a timedelta or None if None provided.'''
        if duration == None:
            return None

        d = ""
        for c in duration:
            if c == ".":
                d += c
                continue
            try:
                int(c)
                d += c
            except:
                continue

        d = float(d)
        if d >= 10:
            return datetime.timedelta(minutes=d)
        else:
            return datetime.timedelta(minutes=d*60)

    ## BUILD ##
    def build_languages(self):

        return [
            "Afghani - Dari", "Afghani - Pushtu", "African - East Swahili", "African - Kirundi", "African Rwandan",
             "Tigrinya", "Arabic", "Arabic - Lebanese", "Assyrian", "Bengali", "Kiribati", "SriLanka - TAMIL",
             "Burmese", "Burmese - Karen", "Cambodian - Khmer", "Chinese - Cantonese", "Chinese - Hakka", "Chinese - Mandarin",
             "Chinese - Shanghainese", "Chinese - Taiwanese", "Chinese - Teochew", "Cook Island Maori", "Somali",
             "Filipino", "Filipino - Tagalog", "French", "German", "German Swiss", "Hungarian", "Fijian - HINDI", 
             "Indian - Gujarati", "Indian - HINDI", "Indian - PUNJABI", "Indian - Urdu", "Telugu - Indian", "Iranian - Persian-Farsi",
             "Italian", "Japanese", "Korean", "Kurdish", "Malay - Malay", "Niuean", "SIGN Language", "Persian", 
             "Portuguese", "Romanian", "Russian", "Samoan", "Serbian", "Serbo Croation", "Slavic - Bosnian", "Slavic Albanian",
             "Ukranian", "Spanish", "Lao", "Thai", "Tongan", "Vietnamese", 'ITS Admin or Meetings'
            ]

    def build_language_groups(self):

        return ['AFGHANI', 'AFRICAN', 'ARABIC / Assyrian', 'BENGALI - SE Asia', 'BURMESE', 'CAMBODIAN', 
                'CHINESE', 'COOK ISLAND MAO', 'ETHIOPIAN', 'FILIPINO', 'FRENCH', 'GERMAN', 'HUNGARIAN',
                'INDIAN', 'IRANIAN', 'ITALIAN', 'JAPANESE', 'KOREAN', 'KURDISH', 'MALAY', 'NIUEAN', 
                'NZSL Sign Language', 'ITS Work', 'PERSIAN', 'PORTUGUESE', 'ROMANIAN', 'RUSSIAN',
                'SAMOAN', 'SERBIAN', 'SERBO CROATION', 'SLAVIC (Slavonic La', 'SPANISH', 'THAI', 
                'TONGAN', 'TURKISH', 'TUVALUAN', 'VIETNAMESE'
            ]

    def build_interpreters(self):

        dict = {}

        rows = str(ReadPDF("G:\\Customer Reporting\\562-CMDHB\\", "Latest Interpreter List.pdf").content).split("\n")
        for i in range(len(rows)):
            row = rows[i].strip()
            if row == "":
                continue
            if row in self.languages and row not in dict:
                dict[row] = []
                ind = i+1
                while True:
 
                    try:
                        line = rows[ind]
                    except:
                        break

                    if line in self.languages or line in self.language_groups and ind != i + 1:
                        break

                    if "Printed" in line or "Firstname" in line or "Interpreter availability," in line:
                        ind += 1
                        continue

                    num_count = 0 
                    first_number_index = None
                    for l in range(len(line)):
                        try:
                            int(line[l])
                            num_count += 1
                            if first_number_index == None:
                                first_number_index = l
                        except:
                            continue
                    if num_count >= 9:
                        if "," not in line:
                            lastname = line[:line.index(' ')].upper()
                            firstname = line[line.index(' ')+1:first_number_index].strip().upper()
                        elif line.index(",") > line.index(" "):
                            lastname = line[:line.index(' ')].upper()
                            firstname = line[line.index(' ')+1:first_number_index].strip().upper()
                        else:
                            lastname = line[:line.index(',')].upper()
                            firstname = line[line.index(',')+1:first_number_index].strip().upper()
                                
                        phone = ""
                        space_count = 0
                        last_index = 0
                        for n in range(first_number_index, len(line)):
                            if line[n] == " " or line[n] == "-":
                                space_count += 1
                                if space_count == 3:
                                    last_index = n
                                    break
                                phone += line[n]
                                continue
                            try:
                                int(line[n])
                                phone += line[n]
                            except:
                                last_index = n
                                break
                        phone = phone.replace(" ", "")
                        phone = phone.replace("-", "")
                        if len(phone) > 12:
                            if "09" in phone:
                                phone = phone[:len(phone)-8]
                            else:
                                phone = phone[:len(phone)-6]

                        for n in range(last_index, len(line)):
                            if line[n] in [" ", "-"]:
                                continue
                            try:
                                int(line[n])
                            except:
                                notes = line[n:]
                                break
                        
                        dict[row].append([lastname, firstname, phone.strip(), notes])                        

                    ind += 1

        return dict

    def build_bookings(self):

        bookings = {}

        # set text variable for putting all readouts from the pdf's
        raw_text = ""

        # process all saved pdf's 
        for filename in ["Todays Job List.pdf", "Tomorrow or Weekend Job List.pdf", "Monday Job List.pdf"]:
            raw_text += ReadPDF("G:\\Customer Reporting\\562-CMDHB\\", filename).content
            raw_text += "\n"
        rows = []
        for row in raw_text.split("\n"):
            if row != "":
                rows.append(row)

        # 
        date = None
        for i in range(len(rows)):

            if len(rows[i]) > 10:
                if rows[i][2] == ":" and rows[i][7:10] == ".m.":
                    time = rows[i][:10]
                    time = time.replace("a.m.", "AM")
                    time = time.replace("p.m.", "PM")
                    dt = '%s %s' % (date, time)
                    #dt = datetime.datetime.strptime('%s %s' % (date, time), '%d %B %Y %I:%M %p')
                    language = rows[i+1]
                    interpreter = rows[i+2].upper()
                    if language in self.languages or language in self.language_groups:
                        if interpreter not in bookings:
                            bookings[interpreter] = [dt]  
                        else:
                            bookings[interpreter].append(dt)
            # DATE
            text = rows[i].split()
            if len(text) >= 3:
                try:
                    if int(text[2]) == int(datetime.datetime.now().year) or int(text[2])-1 == int(datetime.datetime.now().year):
                        day = text[0]
                        month = text[1]
                        year = text[2]

                        date = "%s %s %s" % (day, month, year)
                except:
                    pass

        return bookings

    ## PROCESS ##
    def process(self):

        # Get all interpreters for the language
        interpreters = []
        for data in self.interpreters[self.language]:

            interpreters.append(data)
   
        # Only Language
        if self.timedate == None:
            results = []
            for data in interpreters:
                results.append(self.process_interpreter_details(data))
            return results

        # No Time Flexibility
        if self.earliest == None:
            booked = []
            not_booked = []
            for data in interpreters:
                bookings = self.search_bookings(data)
                if bookings == None:
                    not_booked.append(self.process_interpreter_details(data))
                else:
                    clash = False
                    for booking in bookings:
                        clash = self.check_clash(booking)
                    if not clash:
                        booked.append(self.process_interpreter_details(data, booking_notes=bookings))
                return booked + not_booked

        # Time Flexibility


    def process_interpreter_details(self, data, booking_notes=None):

        lastname = data[0]
        firstname = data[1]
        number = data[2]
        notes = data[3]
        if booking_notes != None:
            notes += " - "

        return '%s %s %s - %s' % (lastname, firstname, number, notes)

    ## SEARCH ##
    def search_bookings(self, data):

        lastname = data[0]
        firstname = data[1]
        name1 = "%s %s" % (lastname, firstname)
        name2 = "%s %s" % (firstname, lastname)

        if name1 in self.bookings:
            return self.bookings[name1]
        elif name2 in self.bookings:
            return self.bookings[name2]
        elif lastname in self.bookings:
            return self.bookings[lastname]
        elif firstname in self.bookings:
            return self.bookings[firstname]
        else:
            return None

    ## CHECK #
    def check_clash(self, booking):

        booking_start = booking
        booking_end = booking + datetime.timedelta(hours=1)

        request_start = self.timedate
        request_end = self.timedate + self.duration

        if request_end <= booking_start or request_start >= booking_end:
            return False

        return True

    ## TEST ##
    def test(self):

        for data in self.interpreters["Indian - PUNJABI"]:

            lastname = data[0]
            firstname = data[1]
            name1 = "%s %s" % (lastname, firstname)
            name2 = "%s %s" % (firstname, lastname)

            number = data[2]
            notes = data[3]

            if name1 in self.bookings:
                print(firstname, lastname, number, notes, self.bookings[name1])
            elif name2 in self.bookings:
                print(firstname, lastname, number, notes, self.bookings[name2])
            elif lastname in self.bookings:
                print(firstname, lastname, number, notes, self.bookings[lastname])
            elif firstname in self.bookings:
                print(firstname, lastname, number, notes, self.bookings[firstname])
            else:
                print("NO BOOKINGS: %s %s %s %s" % (firstname, lastname, number, notes))
        print('\n')

    def test1(self):

        man = "Samoan"

        for data in self.interpreters[man]:

            lastname = data[0]
            firstname = data[1]
            name1 = "%s %s" % (lastname, firstname)
            name2 = "%s %s" % (firstname, lastname)

            number = data[2]
            notes = data[3]

            if name1 in self.bookings:
                print(name1, self.bookings[name1])
                continue
            elif name2 in self.bookings:
                print(name2, self.bookings[name2])
                continue
            elif lastname in self.bookings:
                print(lastname, self.bookings[lastname])
                continue
            elif firstname in self.bookings:
                print(firstname, self.bookings[firstname])
                continue

## ENGINE ##
Main()