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


def message1():
    with open("email.csv",encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Start Date'] == str(tomorrow):
                message1 = "Subject: Trip " + row['TripID'] + ", " + dayleavingtext + "\n\n" + "Hi " + row['Tripper 1'] + ",\n\nYou're leading the " + row['Route'] + " " + dayleavingtext + " with " + row['Staff1DisplayName'] + " Here's what you'll need:\n\nDish kits: " + row['Dish Kits'] + "\nAquatabs: " + row['Aquatabs'] + "\nBoats: " + row['Boats'] + "\nTents: " + row['Tents'] + "\nWhistles: " + row['Whistles'] + "\nPaper bags: " + row['Paper Bags'] + "\nRolls of toilet paper (pack the extra one seperately to avoid disaster): " + row['Toilet Paper'] + "\nBarrels: " + row['Barrels'] + "\nSites: " + row['Sites'] + "\n\n Your campers are:\n\n" + row['Camper 1'] + "\n" + row['Camper 2'] + "\n" + row['Camper 3'] + "\n" + row['Camper 4'] + "\n" + row['Camper 5'] + "\n" + row['Camper 6'] + "\n" + row['Camper 7'] + "\n" + row['Camper 8'] + "\n" + row['Camper 9'] + "\n" + row['Camper 10'] + "\n" + row['Camper 11'] + "\n" + row['Camper 12'] + "\n" + row['Camper 13'] + "\n" + row['Camper 14'] + "\n" + row['Camper 15'] + "\n" + row['Camper 16'] + "\n" + row['Camper 17'] + "\n" + row['Camper 18'] + "\n" + row['Camper 19'] + "\n" + row['Camper 20'] + "\n\nSincerely,\n\nStu" + "\n\n" 


def message2():
    with open("email.csv",encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Start Date'] == str(tomorrow):
	        if row['Camper 9'] == ""
		    message2 = "No camper nine"
		else:
		    message2 = row['Camper 9']

message1()
message2()

def send_me_a_test_email():
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(from_address, password)
        with open("email.csv",encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Start Date'] == str(tomorrow):
                    server.sendmail(from_address,to_address,message)



def email_the_trippers():
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.login(from_address, password)
        with open("email.csv",encoding="utf-8") as file:
            reader = csv.DictReader(file)
            for row in reader:
                if row['Start Date'] == str(tomorrow):
                    server.sendmail(from_address,to_address,message)



send_me_a_test_email()
#email_the_trippers()


#with open("email.csv",encoding="utf-8") as file:
#    reader = csv.DictReader(file)
#    for row in reader:
#        if row['Start Date'] == str(inthreedays):
#            testemail = "row['stuemail']"

#added comment
