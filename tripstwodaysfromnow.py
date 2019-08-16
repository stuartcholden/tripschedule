import csv, smtplib, ssl, xlrd, datetime, subprocess

from_address = "kandalore.trippers@gmail.com"
password = "shalominthehome"
to_address = "tripdirector@kandalore.com"


#input number you want to search
#number = input('Enter date to find\n')
now = datetime.datetime.now()
exceltup = (now.year, now.month, now.day)
today = int(xlrd.xldate.xldate_from_date_tuple((exceltup),0))
tomorrow = today + 1
dayleavingtext = "tomorrow"
intwodays = today + 2
inthreedays = today + 3


context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)
    with open("email.csv",encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Start Date'] == str(tomorrow):

                subject = "Subject: Trip " + row['TripID'] + ", " + dayleavingtext + "\n\n"


                if row['Tripper 1'].startswith('PJ'):
                    message1 = "Hi " + row['Staff1EmailName'] + " & " + row['Staff2EmailName'] + ",\n\nYou're leading the " + row['Route'] + " " + dayleavingtext + "."
                elif row['Tripper 2'] == "":
                    message1 = "Hi " + row['Tripper 1'] + ",\n\nYou're leading the " + row['Route'] + " " + dayleavingtext
                else:
                    message1 = "Hi " + row['Tripper 1'] + " & " + row['Tripper 2'] + ",\n\nYou're leading the " + row['Route'] + " " + dayleavingtext

                if not row['Staff1DisplayName'] == "":
                    staff1 = " with " + row['Staff1DisplayName']
                else:
                    staff1 = "."

                if not row['Staff2DisplayName'] == "":
                    staff2 = " & " + row['Staff2DisplayName']
                else:
                    staff2 = ""

                gear = " Here's what you'll need:\n\nBarrels: " + row['Barrels'] + "\nDish kits: " + row['Dish Kits'] + "\nAquatabs: " + row['Aquatabs'] + "\nBoats: " + row['Boats'] + "\nTents: " + row['Tents'] + "\nWhistles: " + row['Whistles'] + "\nRolls of toilet paper (pack the extra one seperately to avoid disaster): " + row['Toilet Paper']  + "\n"

                if not row['Paper Bags'] == "":
                    paperbags = "Paper bags: " + row['Paper Bags'] + "\n"
                else:
                    paperbags = ""

                if not row['Booking Reference'] == "":
                    bookingreference = "Booking Reference Number: " + row['Booking Reference'] + "\n"
                else:
                    bookingreference = ""

                if not row['Sites'] == "":
                    sites = "Sites: " + row['Sites'] + "\n"
                else:
                    sites = ""

                if not row['Drop Off Required'] == "":
                    dropoff = "Drop off: " + row['Drop Off Required'] + "\n"
                else:
                    dropoff = ""

                if not row['Pick Up Required'] == "":
                    pickup = "Pick up: " + row['Pick Up Required'] + "\n"
                else:
                    pickup = "" 

                if not row['SPOT'] == "":
                    spot = "SPOT: " + row['SPOT'] + "\n"
                else:
                    spot = ""

                if not row['Emerg Money'] == "":
                    money = "Emergency Money: " + row['Emerg Money'] + "\n"
                else:
                    money = ""

                campers = "\n\nHere are your campers, as of Trip Schedule " + row['tripscheduleversion'] + ":\n\n"

                if not row['Camper 1'] == "":
                    camper1 = row['Camper 1']+ "\n"
                else:
                    camper1 = ""
                if not row['Camper 2'] == "":
                    camper2 = row['Camper 2'] + "\n"
                else:
                    camper2 = ""
                if not row['Camper 3'] == "":
                    camper3 = row['Camper 3'] + "\n"
                else:
                    camper3 = ""
                if not row['Camper 4'] == "":
                    camper4 = row['Camper 4'] + "\n"
                else:
                    camper4 = ""
                if not row['Camper 5'] == "":
                    camper5 = row['Camper 5'] + "\n"
                else:
                    camper5 = ""
                if not row['Camper 6'] == "":
                    camper6 = row['Camper 6'] + "\n"
                else:
                    camper6 = ""
                if not row['Camper 7'] == "":
                    camper7 = row['Camper 7'] + "\n"
                else:
                    camper7 = ""
                if not row['Camper 8'] == "":
                    camper8 = row['Camper 8'] + "\n"
                else:
                    camper8 = ""
                if not row['Camper 9'] == "":
                    camper9 = row['Camper 9'] + "\n"
                else:
                    camper9 = ""
                if not row['Camper 10'] == "":
                    camper10 = row['Camper 10'] + "\n"
                else:
                    camper10 = ""
                if not row['Camper 11'] == "":
                    camper11 = row['Camper 11'] + "\n"
                else:
                    camper11 = ""
                if not row['Camper 12'] == "":
                    camper12 = row['Camper 12'] + "\n"
                else:
                    camper12 = ""
                if not row['Camper 13'] == "":
                    camper13 = row['Camper 13'] + "\n"
                else:
                    camper13 = ""
                if not row['Camper 14'] == "":
                    camper14 = row['Camper 14'] + "\n"
                else:
                    camper14 = ""
                if not row['Camper 15'] == "":
                    camper15 = row['Camper 15'] + "\n"
                else:
                    camper15 = ""
                if not row['Camper 16'] == "":
                    camper16 = row['Camper 16'] + "\n"
                else:
                    camper16 = ""
                if not row['Camper 17'] == "":
                    camper17 = row['Camper 17'] + "\n"
                else:
                    camper17 = ""
                if not row['Camper 18'] == "":
                    camper18 = row['Camper 18'] + "\n"
                else:
                    camper18 = ""
                if not row['Camper 19'] == "":
                    camper19 = row['Camper 19'] + "\n"
                else:
                    camper19 = ""
                if not row['Camper 20'] == "":
                    camper20 = row['Camper 20'] + "\n"
                else:
                    camper20 = ""
                if not row['Camper 21'] == "":
                    camper21 = row['Camper 21'] + "\n"
                else:
                    camper21 = ""
                if not row['Camper 22'] == "":
                    camper22 = row['Camper 22'] + "\n"
                else:
                    camper22 = ""
                if not row['Camper 23'] == "":
                    camper23 = row['Camper 23'] + "\n"
                else:
                    camper23 = ""
                if not row['Camper 24'] == "":
                    camper24 = row['Camper 24'] + "\n"
                else:
                    camper24 = ""
                if not row['Camper 25'] == "":
                    camper25 = row['Camper 25'] + "\n"
                else:
                    camper25 = ""
                if not row['Camper 26'] == "":
                    camper26 = row['Camper 26'] + "\n"
                else:
                    camper26 = ""
                if not row['Camper 27'] == "":
                    camper27 = row['Camper 27'] + "\n"
                else:
                    camper27 = ""
                if not row['Camper 28'] == "":
                    camper28 = row['Camper 28'] + "\n"
                else:
                    camper28 = ""
                if not row['Camper 29'] == "":
                    camper29 = row['Camper 29'] + "\n"
                else:
                    camper29 = ""
                if not row['Camper 30'] == "":
                    camper30 = row['Camper 30'] + "\n"
                else:
                    camper30 = ""

                if row['Praise be to Cthulhu, devourer of young love and purveyor of discounted leather jackets'] == "yes":
#                    signoff = subprocess.Popen("fortune fortune/cthulhu", stdout = subprocess.PIPE)
                    signoff = "\n\nPraise be to Cthulhu, devourer of young love and purveryor of fine discounted leather jackets,\n\nStu"
                else:
                    signoff = "\n\nSincerely,\n\nStu"

#                if row['Tripper 1'].startswith('PJ'):

                pjmessage = subject+message1+gear+paperbags+bookingreference+sites+dropoff+pickup+spot+money+campers+camper1+camper2+camper3+camper4+camper5+camper6+camper7+camper8+camper9+camper10+camper11+camper12+camper13+camper14+camper15+camper16+camper17+camper18+camper19+camper20+camper21+camper22+camper23+camper24+camper25+camper26+camper27+camper28+camper29+camper30+signoff

#                else:

                coremessage = subject+message1+staff1+staff2+gear+paperbags+bookingreference+sites+dropoff+pickup+spot+money+campers+camper1+camper2+camper3+camper4+camper5+camper6+camper7+camper8+camper9+camper10+camper11+camper12+camper13+camper14+camper15+camper16+camper17+camper18+camper19+camper20+camper21+camper22+camper23+camper24+camper25+camper26+camper27+camper28+camper29+camper30+signoff
                print(coremessage)

#                expmessage = 

#                x2message = 

#I think all the below lines can be modified to use a while loop and just send a message to any email cell that is filled out
                if row['Tripper 1'].startswith('PJ'):
                    print(pjmessage)
                    server.sendmail(from_address,row['email1']+", "+row['email2'],pjmessage)
                    print("Send to: " + row['email1'] + ", " + row['email2'])
                if row['email1'] == "":
                    server.sendmail(from_address,"tripdirector@kandalore.com","Subject: Trip " + row['TripID'] + " does not have an email for the lead tripper")
                    print("Subject: Trip " + row['TripID'] + " does not have an email for the lead tripper\n\nHopefully there's at least a tripper.")
                if row['Section'] == "Exp":
                    print(expmessage)
                    server.sendmail(from_address,row['email1']+", "+row['email2'],expmessage)
                if row['Section'] == "X2":
                    print(x2message)
                    server.sendmail(from_address,row['email1']+", "+row['email2'],expmessage)
                else:
                    print(coremessage)
                    server.sendmail(from_address,row['email1'],coremessage)

#                if not row['email2'] == "":
#                    server.sendmail(from_address,row['email2'],message)

def send_me_a_test_email():
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(from_address, password)
        with open("email.csv",encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Start Date'] == str(tomorrow):
                    server.sendmail(from_address,to_address,message)


def print_to_command_line():
    with open("email.csv",encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Start Date'] == str(tomorrow):
                print(message1)


def email_the_trippers():
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(from_address, password)
        with open("email.csv",encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Start Date'] == str(tomorrow):
                    server.sendmail(from_address,to_address,message)



#print_to_command_line()
#send_me_a_test_email()
#email_the_trippers()
