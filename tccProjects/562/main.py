## IMPORTS ##
import csv
import datetime
import json
import sys
import webbrowser

## APP IMPORTS ##
from app.PDF.readPDF import ReadPDF

## CLASS ##
class Main():

    def __init__(self, *args, language=None, timedate=None, earliest=None, latest=None, duration=None, operator=None, **kwargs):

        #
        self.language = language
        self.timedate = self.convert_timedate(timedate)
        self.earliest = self.convert_timedate(earliest)
        self.latest = self.convert_timedate(latest)
        if self.earliest == self.latest:
            self.earliest = None
            self.latest = None
        self.duration = self.convert_duration(duration)
        self.operator = operator

        #
        self.languages = self.build_languages()
        self.language_groups = self.build_language_groups()

        # If this is called from the Supervisor button, process interpreters and bookings
        # (with booking date and time as a String), and convert to a Json file
        # which is saved on the G Drive
        if self.language == None:
            self.interpreters = self.build_interpreters()
            self.bookings = self.build_bookings()
            self.build_unavailable()
            self.convert_to()
        # If this is called from the 562 form, access the Json file, convert the date and time string
        # to a Datetime object and 
        else:
            self.interpreters = {}
            self.bookings = {}
            self.convert_from()

        # ENGINE
        if self.language != None:
            self.results = "LANGUAGE: %s\nREQUESTED FOR: %s for %s" % (self.language.upper(), timedate, duration)
            if earliest == None:
                earliest = timedate
            if latest == None:
                latest = timedate
            if earliest == timedate and latest == timedate:
                pass
            else:
                self.results += "\nEARLIEST START: %s\nLATEST START: %s" % (earliest, latest)
            self.results += "\n\n" 

            results = self.process()
            for box in results:
                for r in box:
                    self.results += "%s---\n" % r
                self.results += "\n"
            self.open_notepad()

        else:
            self.record_csv()
            print("SUPERVISOR AUTOMATION COMPLETE - CLOSE THIS SCREEN")

    ## CSV NOTE ##
    def record_csv(self):

        """This needs to be added in if people are still not running the CMDHB code, 
        so I can get a record of who does and doesn't run it."""

        pass

    ## CONVERT ##
    def convert_timedate(self, timedate):

        '''Returns a str of a date & time to a datetime object, or None if None provided.'''

        if timedate == None:
            return None

        # FORMAT EXAMPLE "16/03/2021 1:19:26 pm"
        timedate = timedate.replace("am", "AM")
        timedate = timedate.replace("pm", "PM")
        return datetime.datetime.strptime(timedate, '%d/%m/%Y %H:%M:%S %p')

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

        return ['AFGHANI', 'AFRICAN', 'ARABIC', 'ARABIC / Assyrian', 'BENGALI - SE Asia', 'BURMESE', 'CAMBODIAN', 
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

            # If the row contains a language, but that language is not already in the dict
            if row in self.languages and row not in dict:
                
                #print(row)

                dict[row] = []
                ind = i+1
                # after there is a valid language, go through line by line extracting all interpreters, 
                # phone numbers and notes
                while True:

                    # make sure this is not the last row
                    try:
                        line = rows[ind]
                    except:
                        break

                    if line in self.languages or line in self.language_groups and ind != i + 1:
                        break

                    if "Printed" in line or "Firstname" in line or "Interpreter availability," in line:
                        ind += 1
                        continue

                    # Go through
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
                    if num_count >= 9 and line[:5] != "Phone":
                        if "," not in line:
                            lastname = line[:line.index(' ')].upper()
                            firstname = line[line.index(' ')+1:first_number_index].strip().upper()
                        elif line.index(",") > line.index(" "):
                            lastname = line[:line.index(' ')].upper()
                            firstname = line[line.index(' ')+1:first_number_index].strip().upper()
                        else:
                            lastname = line[:line.index(',')].upper()
                            firstname = line[line.index(',')+1:first_number_index].strip().upper()

                        # Extract the Phone Number, but using the index (found above) for the first number.
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
                        # if the number is too long, then it is probably 2 numbers, a mobile and landline.
                        # The landline should be the one that starts with 09 after at least 8 other characters
                        # i.e. to prevent it pulling mobile numbers with 09 in it.
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
                                note = line[n:]
                                notes = self.build_notes(note, ind+1, rows)
                                break

                        # Decypher and Watis Exceptions
                        if firstname == "DECYPHER":
                            dict[row].append(["DECYPHER", "DECYPHER", "0274792419", "Decypher Interpreting: **TINT** Only. Email: info@decypher.co.nz, Phone: 0274792419 (afterhours)***TINT*** ONLY"])
                        elif firstname == "WATIS":
                            dict[row].append(["WATIS", "WATIS", "09 442 3211 extn 221", "Before using for Afterhours/Weekends/Public Holidays please phone ITS Manager Kim de Jong to get approval"])
                        else:             
                            dict[row].append([lastname, firstname, phone.strip(), notes])   

                    ind += 1

        #print(dict)
        return dict

    def build_unavailable(self):

        dict = {}

        rows = str(ReadPDF("G:\\Customer Reporting\\562-CMDHB\\", "ITS Availability List.pdf").content).split("\n")
        data = ""
        for i in range(len(rows)):
            row = rows[i].strip()

            if row == "":
                continue
            if "Name Dates Language" in row:
                continue
            if "ITS AVAILABILITY W/ENDING" in row:
                break

            data += "%s " % row.upper()

        data = data.split()

        available = []
        line = []
        for d in data:

            if d in self.language_groups:
                
                name = line[0]
                details = " ".join(line[1:])
                line = []

                for language in self.interpreters:

                    if d in language.upper():

                        for i in range(len(self.interpreters[language])):
                            row = self.interpreters[language][i]
                            firstname = row[0]
                            
                            if name == firstname or name in firstname:
                                self.interpreters[language][i][3] += "\nCHECK UNAVAILABILITY: %s" % details
                                break

            else:
                line.append(d)

    def build_notes(self, note, i, rows):

        """Extracts all text that could be a note for a specific 
        interpreter, then processes it so only the General and DO NOT USE WITH notes are included,
        then returns the note. """

        # Extract Text
        while True:

            # check if last line
            try:

                line = rows[i]
            except:  

                break

            # if crosses multiple pages, keep going
            if "Printed" in line or "Firstname" in line or "Interpreter availability," in line:
                i += 1
                continue

            # if it has moved to a new language, stop
            if line in self.languages or line in self.language_groups:
                break

            # Decypher exception
            if line[:5] == "Phone":
                break

            # check to see if it is a new interpreter line
            num_count = 0
            for l in range(len(line)):
                try:
                    int(line[l])
                    num_count += 1
                    if num_count == 9:
                        break
                except:
                    continue


            # OTHERWISE assume this is a note
            if note[-1] != " ":
                note += " %s" % line
            else:
                note += line

            i += 1

        ## Process Text

        # Remove qualification
        qualifications = ["AUT Liaison", "AIT Cert", "AIT Medical", "Akld Regional", "Unitec Cert", "UNITEC Cert",
                    "AIT Health", "AUT Health", "AIT Adv", "AUT Adv", "MIT Adv", "MIT Cert", "Court Interpreters",
                    "NAATI", "AIT Diploma", "AUT Diploma"
                ]
        for phrase in qualifications:
            if phrase in note:
                note = note[:note.index(phrase)]

        #Remove work hours and add DOES NOT WORK WITH
        work_hours = ["MMH Emp 2nd", "Group 2 less40", "Group 1+40 PM", "ITS Contractor",
                      "Part Time", "Full"]         
        for w in work_hours:
            if w in note:
                n = note.split(w)
                if ", ," in n[1]:
                    note = n[0]
                else:
                    note = n[0] + "\nDOES NOT WORK WITH: %s" % n[1]
                break

        # Weird Exceptions
        if "Lima DRYSDALE" in note:
            note = ""

        ## Return
        return note

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

            #print(i, rows[i])

            # if there are more than 10 characters on the line, continue
            if len(rows[i]) > 10:
                # if this is a row with a 24 hour time on it (therefore an individual booking)
                if rows[i][2] == ":" and rows[i][7:10] == ".m.":
                    # extracts the time,and converts 
                    time = rows[i][:10]
                    time = time.replace("a.m.", "AM")
                    time = time.replace("p.m.", "PM")

                    dt = '%s %s' % (date, time) # String
                    #dt = datetime.datetime.strptime('%s %s' % (date, time), '%d %B %Y %I:%M %p') # Datetime

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

    ## JSON ##
    def convert_to(self):

        with open('G:\\Customer Reporting\\562-CMDHB\\Automation\\interpreters.json', 'w') as outfile:
            json.dump(self.interpreters, outfile)

        with open("G:\\Customer Reporting\\562-CMDHB\\Automation\\bookings.json", 'w') as outfile:
            json.dump(self.bookings, outfile)

    def convert_from(self):

         with open('G:\\Customer Reporting\\562-CMDHB\\Automation\\bookings.json', 'r') as outfile:
            bookings = json.load(outfile)   

            for interpreter in bookings:
                self.bookings[interpreter] = []
                for dt_string in bookings[interpreter]:
                    dt_datetime = datetime.datetime.strptime(dt_string, '%d %B %Y %I:%M %p') # Datetime
                    self.bookings[interpreter].append(dt_datetime)

         with open('G:\\Customer Reporting\\562-CMDHB\\Automation\\interpreters.json', 'r') as outfile:
            self.interpreters = json.load(outfile)

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
            return [results]

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
                        if not clash:
                            clash = self.check_clash(booking)
                    if not clash:
                        booked.append(self.process_interpreter_details(data, booking_notes=bookings))

            return [booked, not_booked]

        # Flexible
        booked = []
        flex_booked = []
        not_booked = []
        for data in interpreters:
            bookings = self.search_bookings(data)
            if bookings == None:
                not_booked.append(self.process_interpreter_details(data))
            else:
                clash = False
                # check best time clash
                for booking in bookings:
                    if not clash:
                        clash = self.check_clash(booking)
                if not clash:
                    booked.append(self.process_interpreter_details(data, booking_notes=bookings))
                else:
                    # check flexible clash
                    clash = False
                    for b in range(len(bookings)):
                        booking = bookings[b]
                        try:
                            if b != 0:
                                previous_booking = bookings[b-1]
                            else:
                                previous_booking = None
                        except:
                            previous_booking = None
                        try:
                            next_booking = bookings[b+1]
                        except:
                            next_booking = None
                        if not clash:
                            clash = self.check_flexible_clash(booking, previous_booking, next_booking)
                    if not clash:
                        flex_booked.append(self.process_interpreter_details(data, booking_notes=bookings, flex=True))

        return [booked, flex_booked, not_booked]

    def process_interpreter_details(self, data, booking_notes=None, flex=False):

        lastname = data[0]
        firstname = data[1]
        number = data[2]

        message = "%s, %s - %s\n" % (lastname, firstname, number)

        # Booking
        if booking_notes != None:
            if not flex:
                message += "CURRENT BOOKINGS (Exact): "
            else:
                message += "CURRENT BOOKINGS (Flexibility): "

            for b in booking_notes:
                if b.date() == self.timedate.date():
                    message += "%s, " % b.time()

            message = message[:-2]
            message += "\n"

        # Notes
        message += "NOTES: %s\n" % data[3]

        return message

    ## SEARCH ##
    def search_bookings(self, data):

        lastname = data[0]
        firstname = data[1]
        name1 = "%s %s" % (lastname, firstname)
        name2 = "%s %s" % (firstname, lastname)

        bookings = []
        if name1 in self.bookings:
            for b in self.bookings[name1]:
                if b.date() == self.timedate.date():
                    bookings.append(b)
        elif name2 in self.bookings:
            for b in self.bookings[name2]:
                if b.date() == self.timedate.date():
                    bookings.append(b)
        elif lastname in self.bookings:
            for b in self.bookings[lastname]:
                if b.date() == self.timedate.date():
                    bookings.append(b)
        elif firstname in self.bookings:
            for b in self.bookings[firstname]:
                if b.date() == self.timedate.date():
                    bookings.append(b)
        if bookings != []:
            return bookings
        else:
            return None

    ## CHECK #
    def check_clash(self, booking):

        booking_start = booking
        booking_end = booking + datetime.timedelta(hours=1)

        request_start = self.timedate
        request_end = self.timedate + self.duration

        if request_end <= booking_start:
            return False
        elif request_start >= booking_end:
            return False
        return True

    def check_flexible_clash(self, booking, previous_booking, next_booking):

        """This seems to work - but I have left some of the mistaken code in case it turns
        out it was more correct than I thought """

        booking_start = booking
        booking_end = booking + datetime.timedelta(hours=1)

        if booking_start >= (self.earliest + self.duration):
            return False
        if False:
            if previous_booking == None:
                return False
            elif booking_start >= (previous_booking + datetime.timedelta(hours=1) + self.duration): # previous booking + one hour + the time for the requested booking
                return False
        if booking_end <= self.latest:
            return False
        if False:
            if next_booking == None:
                return False
            elif (booking_end + self.duration) >= next_booking: # if this would run into the next booking
                return False

        return True

    ## OPEN ##
    def open_notepad(self):
        """Writes a .txt file of the email data and opens the file."""

        file = open('G:\\Customer Reporting\\562-CMDHB\\Automation\\%s.txt' % self.operator, 'w')
        file.write(self.results)
        file.close()

        webbrowser.open('G:\\Customer Reporting\\562-CMDHB\\Automation\\%s.txt' % self.operator)

## ENGINE ##

# If the program is called from the form
if len(sys.argv) > 1:
    raw_argv = sys.argv[1].replace("@", " ")
    argv = raw_argv.split("###")
    language = argv[0]
    timedate= argv[1]
    earliest = argv[2]
    latest = argv[3]
    duration = argv[4]
    operator = argv[5]

    Main(language=language, timedate=timedate, earliest=earliest, latest=latest, duration=duration, operator=operator)
# If the program is called from the Supervisor .exe
else:
    #Main(language="Chinese - Mandarin", timedate="12/05/2021 08:30:00 am", earliest="12/05/2021 7:00:00 am", latest="12/05/2021 10:00:00 am", duration="1 hr", operator="Michael Atheros")
    Main()