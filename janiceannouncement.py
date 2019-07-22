import csv, smtplib, ssl, xlrd, datetime

from_address = "kandalore.trippers@gmail.com"
password = "shalominthehome"
to_address = "tripdirector@kandalore.com"


#input number you want to search
#number = input('Enter date to find\n')
now = datetime.datetime.now()
exceltup = (now.year, now.month, now.day)
today = int(xlrd.xldate.xldate_from_date_tuple((exceltup),0))
tomorrow = today + 10
dayleavingtext = "tomorrow"
intwodays = today + 11
inthreedays = today + 12


context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)
    with open("email.csv",encoding="utf-8") as file:
        reader = csv.DictReader(file)
        subject = "Subject: Trips leaving in the next few days\n\n"
        message1 = "Today:\n\n"
        message2 = "\n\nTomorrow:\n\n"
        message3 = "\n\nThe next day:\n\n"
        message4 = "\n\nThe day after that:\n\n"
        for row in reader:
            if row['Start Date'] == str(today):
                message1 += row['TripID'] + " - " + row['Tripper 1'] + " and the " + row['Section'] + " on the " + row['Route'] + ", \n"

            if row['Start Date'] == str(tomorrow):
                message2 += row['TripID'] + " - " + row['Tripper 1'] + " and the " + row['Section'] + " on the " + row['Route'] + ", \n"

            if row['Start Date'] == str(intwodays):
                message3 += row['TripID'] + " - " + row['Tripper 1'] + " and the " + row['Section'] + " on the " + row['Route'] + ", \n"

            if row['Start Date'] == str(inthreedays):
                message4 += row['TripID'] + " - " + row['Tripper 1'] + " and the " + row['Section'] + " on the " + row['Route'] + ", \n"

        server.sendmail(from_address,to_address,subject+message1+message2+message3+message4)
