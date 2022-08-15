import csv, smtplib, ssl, xlrd, openpyxl, datetime, subprocess, argparse, n2w

to_address = "tripdirector@kandalore.com"

import gmailapppassword
from_address = gmailapppassword.username
password = gmailapppassword.password

my_parser = argparse.ArgumentParser()
my_parser.add_argument('-t', '--test', action='store_true', help='Testing mode. All emails are sent to tripdirector@kandalore.com')

args = my_parser.parse_args()

print(vars(args))

if args.test is True:
    print('Wow it worked')

path = "/Users/stu/Documents/Tripping/Schedule.xlsm"

#input number you want to search
#number = input('Enter date to find\n')
now = datetime.datetime.now()
exceltup = (now.year, now.month, now.day)
today = int(xlrd.xldate.xldate_from_date_tuple((exceltup),0))
tomorrow = today + 1
dayleavingtext = "in two days"
intwodays = today + 2
inthreedays = today + 3

tripUID = input("Enter a trip UID:\n")

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)
    with open("email.csv",encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['TripID'] == tripUID:

                if not row['TripID'] == "":
                    daysuntildeparture = int(row['Start Date']) - today

                subject = "Subject: Trip " + row['TripID'] + ", " + " in " + str(n2w.convert(daysuntildeparture)) + " days" + "\n\n"

                pjhelpersubject = "Subject: Help " + row['Staff1EmailName'] + " & " + row['Staff2EmailName'] + " packing out PJ trip " + row['TripID']+"\n\n"

                if row['Tripper 1'].startswith('PJ'):
                    message1 = "Hi " + row['Staff1EmailName'] + " & " + row['Staff2EmailName'] + ",\n\nYou're leading the " + row['Route'] + " in " + str(n2w.convert(daysuntildeparture)) + " days."
                elif row['Tripper 2'] == "":
                    message1 = "Hi " + row['Tripper 1'] + ",\n\nYou're leading the " + row['Route'] + " in " + str(n2w.convert(daysuntildeparture)) + " days."
                else:
                    message1 = "Hi " + row['Tripper 1'] + " & " + row['Tripper 2'] + ",\n\nYou're leading the " + row['Route'] + " " + "in " + str(n2w.convert(daysuntildeparture)) + " days"

                if not row['Staff1DisplayName'] == "":
                    staff1 = " with " + row['Staff1DisplayName']
                else:
                    staff1 = "."

                if not row['Staff2DisplayName'] == "":
                    staff2 = " & " + row['Staff2DisplayName']
                else:
                    staff2 = ""

                if not row['PJ Helper'] == "":
                    pjhelper = " " + row['PJ Helper'] + " will help you pack.\n"
                else:
                    pjhelper = ""

                gear = "\n\nHere's what you'll need:\n\nBarrels: " + row['Barrels'] + "\nDish kits: " + row['Dish Kits'] + "\nAquatab kits: " + row['Packs of Aquatabs'] + "\nBoats: " + row['Boats'] + "\nTents: " + row['Tents'] + "\nWhistles: " + row['Whistles'] + "\nRolls of toilet paper (pack the extra one seperately to avoid disaster): " + row['Toilet Paper']  + "\n"

                if not row['Paper Bags'] == "":
                    menstrual = "Paper bags: " + row['Paper Bags'] + "\n"
                else:
                    menstrual = ""

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

                if str(row['Days on Trip']) == str(12):
                    extrasatbat = "Extra Sat Phone Battery: 1\n"
                if int(row['Days on Trip']) >= 20:
                    extrasatbat = "Extra Sat Phone Batteries: 2\n"
                else:
                    extrasatbat = "wut\n"

                if not row['Emerg Money'] == "None":
                    money = "Emergency Money: $" + row['Emerg Money'] + "\n"
                else:
                    money = ""

                if not row['Trip Notes'] == "":
                    tripnotes = "\nAdditional Notes: " + row['Trip Notes'] + "\n"
                else:
                    tripnotes = ""

                campers = "\n\nHere are your campers, as of Trip Schedule " + row['tripscheduleversion'] + ":\n\n"

                for i in range(1, 30):
                    if row['Camper ' + str(i)] != "":
                        campers += row['Camper ' + str(i)] + "\n"
                    else:
                        campers += ""



                if row['Praise be to Cthulhu, devourer of young love and purveyor of discounted leather jackets'] == "yes":
                    signoff = "\n\nPraise be to Cthulhu, devourer of young love and purveryor of fine discounted leather jackets,\n\nStu"
                else:
                    signoff = "\n\nSincerely,\n\nStu"

                pjmessage = subject+message1+pjhelper+gear+menstrual+bookingreference+sites+dropoff+pickup+spot+money+tripnotes+campers+signoff
                pjhelpermessage = pjhelpersubject + "Hi " + row['PJ Helper'] + ", could you help " + row['Staff1EmailName'] + " & " + row['Staff2EmailName'] + " pack out trip " + row['TripID'] + "on" 
                coremessage = subject+message1+staff1+staff2+gear+menstrual+bookingreference+sites+dropoff+pickup+spot+extrasatbat+money+tripnotes+campers+signoff
                errormessage = subject+"This is the catchall function. " + row['Tripper 1'] + " has not received this message."
                noleadtripper = subject + "Trip " + row['TripID'] + " does not have an email for the lead tripper."

                if row['Section'] == "PG" or row['Section'] == "JG" or row['Section'] == "PB" or row['Section'] == "JB" and row['Tripper 1'].startswith('PJ'):
                    server.sendmail(from_address,[row['email1'], row['email2'],row['SectionHeadEmail'],row['pjhelperemail'],row['DirectorofCampLifeEmail']],pjmessage.encode('utf-8'))

                if row['Section'] == "PG" or row['Section'] == "JG" or row['Section'] == "PB" or row['Section'] == "JB" and not row['Tripper 1'].startswith('PJ'):
                    server.sendmail(from_address,row['email1'],pjmessage.encode('utf-8'))

                if row['email1'] == "":
                    server.sendmail(from_address,row['tripdirectoremail'],noleadtripper.encode('utf-8'))

                if row['Section'] == "Exp":
                    server.sendmail(from_address,row['email1'],coremessage.encode('utf-8'))
                    server.sendmail(from_address,row['email2'],coremessage.encode('utf-8'))

                if row['Section'] == "X2":
                    server.sendmail(from_address,[row['email1'],row['email2']],coremessage.encode('utf-8')),

                if row['Trip Program'] == "Core":
                    server.sendmail(from_address,[row['email1'],row['email2']],coremessage.encode('utf-8'))


                if args.test is True:
                    print('Wow it worked')
