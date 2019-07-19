import csv, smtplib, ssl, xlrd, datetime

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


with open("email.csv",encoding="utf-8") as file:
    reader = csv.DictReader(file)
    for row in reader:
        if row['Start Date'] == str(tomorrow):

            message1 = "Subject: Trip " + row['TripID'] + ", " + dayleavingtext + "\n\n" + "Hi " + row['Tripper 1'] + ",\n\nYou're leading the " + row['Route'] + " " + dayleavingtext + " with " + row['Staff1DisplayName'] + " Here's what you'll need:\n\nDish kits: " + row['Dish Kits'] + "\nAquatabs: " + row['Aquatabs'] + "\nBoats: " + row['Boats'] + "\nTents: " + row['Tents'] + "\nWhistles: " + row['Whistles'] + "\nPaper bags: " + row['Paper Bags'] + "\nRolls of toilet paper (pack the extra one seperately to avoid disaster): " + row['Toilet Paper'] + "\nBarrels: " + row['Barrels'] + "\n"

            if not row['Sites'] == "":
                sites = "Sites: " + row['Sites']
            else:
                sites = ""

            if not row['SPOT'] == "":
                spot = "SPOT: " + row['SPOT']
            else:
                spot = ""

            campers = "\n\nYour campers are:\n\n"

            if not row['Camper 1'] == "":
                camper1 = row['Camper 1']+ "\n"
            if not row['Camper 2'] == "":
                camper2 = row['Camper 2'] + "\n"
            if not row['Camper 3'] == "":
                camper3 = row['Camper 3'] + "\n"
            if not row['Camper 4'] == "":
                camper4 = row['Camper 4'] + "\n"
            if not row['Camper 5'] == "":
                camper5 = row['Camper 5'] + "\n"
            if not row['Camper 6'] == "":
                camper6 = row['Camper 6'] + "\n"
            if not row['Camper 7'] == "":
                camper7 = row['Camper 7'] + "\n"
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

            if row['Praise be to Cthulhu, devourer of young love and purveyor of discounted leather jackets'] == "yes":
                signoff = "\n\nPraise be to Cthulhu,\n\nStu"
            else:
                signoff = "\n\nSincerely,\n\nStu"

            print(message1+sites+campers+camper1+camper2+camper3+camper4+camper5+camper6+camper7+camper8+camper9+camper10+camper11+camper12+signoff)



#message = message1 + message2

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
