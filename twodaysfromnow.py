import csv, smtplib, ssl, xlrd, datetime, subprocess

import gmailapppassword
from_address = gmailapppassword.username
password = gmailapppassword.password

#input number you want to search
#number = input('Enter date to find\n')
now = datetime.datetime.now()
exceltup = (now.year, now.month, now.day)
today = int(xlrd.xldate.xldate_from_date_tuple((exceltup),0))
tomorrow = today + 1
dayleavingtext = "in two days"
intwodays = today + 2
inthreedays = today + 3


context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address,password)
    with open("email.csv",encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Start Date'] == str(intwodays):

                subject = "Subject: Trip " + row['TripID'] + ", " + dayleavingtext + "\n\n"


                if row['Tripper 1'].startswith('PJ'):
                    message1 = "Hi " + row['Staff1EmailName'] + " & " + row['Staff2EmailName'] + ",\n\nYou're leading the " + row['Route'] + " " + dayleavingtext + "."
                elif row['Tripper 2'] == "":
                    message1 = "Hi " + row['Tripper1HR'] + ",\n\nYou're leading the " + row['Route'] + " " + dayleavingtext
                else:
                    message1 = "Hi " + row['Tripper1HR'] + " & " + row['Tripper2HR'] + ",\n\nYou're leading the " + row['Route'] + " " + dayleavingtext

                if not row['Staff1EmailName'] == "":
                    staff1 = " with " + row['Staff1EmailName']
                else:
                    staff1 = "."

                if not row['Staff2EmailName'] == "":
                    staff2 = " & " + row['Staff2EmailName']
                else:
                    staff2 = ""

                if not row['PJ Helper'] == "":
                    pjhelper = ". " + row['PJ Helper'] + " will help you pack"
                else:
                    pjhelper = ""

                if row['Barrels'] == "0":
                    barrels = ""
                else:
                    barrels = "\nBarrels: " + row['Barrels']

                gear = ". Here's what you'll need:\n" + barrels + "\nDish kits: " + row['Dish Kits'] + "\nAquatabs: " + row['Aquatabs'] + "\nBoats: " + row['Boats'] + "\nTents: " + row['Tents'] + "\nWhistles: " + row['Whistles'] + "\nRolls of toilet paper (pack the extra one seperately to avoid disaster): " + row['Toilet Paper']  + "\n"

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

                if int(row['Days on Trip']) > 7:
                    extrasatbat = "Extra Sat Phone Battery: 1\n"
                elif int(row['Days on Trip']) > 12:
                    extrasatbat += "Extra Sat Phone Batteries: 2\n"
                else:
                    extrasatbat = ""

                if not row['Emerg Money'] == "":
                    money = "Emergency Money: " + row['Emerg Money'] + "\n"
                else:
                    money = ""

                if not row['Menu'] == "":
                    menu = "Menu: " + row['Menu'] + "\n"
                else:
                     menu = ""

                if not row['Trip Notes'] == "":
                    notes = "\nNotes: " + row['Trip Notes'] + "\n"
                else:
                    notes = ""

                campers = "\n\nHere are your campers, as of Trip Schedule " + row['tripscheduleversion'] + ":\n\nCabin: " + row['Cabin'] + "\n\n"

                for i in range(1, 30):
                    if row['Camper ' + str(i)] != "":
                        campers += row['Camper ' + str(i)] + "\n"
                    else:
                        campers += ""

                if row['Praise be to Cthulhu, devourer of young love and purveyor of discounted leather jackets'] == "yes":
                    signoff = "\n\nPraise be to Cthulhu, devourer of young love and purveryor of fine discounted leather jackets,\n\nStu"
                else:
                    signoff = "\n\nSincerely,\n\nStu"

                pjmessage = subject+message1+pjhelper+gear+paperbags+bookingreference+sites+dropoff+pickup+spot+notes+campers+signoff
                coremessage = subject+message1+staff1+staff2+gear+paperbags+bookingreference+sites+dropoff+pickup+spot+extrasatbat+money+menu+notes+campers+signoff
                expmessage = subject+message1+staff1+staff2+gear+paperbags+bookingreference+sites+dropoff+pickup+spot+extrasatbat+money+menu+notes+campers+signoff

                if row['email1'] == "":
                    server.sendmail(from_address,"tripdirector@kandalore.com","Subject: Trip " + row['TripID'] + " does not have an email for the lead tripper\n\nHopefully there's at least a tripper.")
                if row['Section'] == "Exp":
                    server.sendmail(from_address,row['email1']+", "+row['email2'],expmessage.encode('utf-8'))
                if row['Section'] == "X2":
                    server.sendmail(from_address,row['email1']+", "+row['email2'],expmessage.encode('utf-8'))
                if row['Section'] == "PG" or row['Section'] == "JG" or row['Section'] == "PB" or row['Section'] == "JB":
                    server.sendmail(from_address,[row['email1'], row['email2'], row['SectionHeadEmail'], row['DirectorofCampLifeEmail']],pjmessage.encode('utf-8'))
                else:
                    server.sendmail(from_address,row['email1'],coremessage.encode('utf-8'))
