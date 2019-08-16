import csv, smtplib, ssl, xlrd, datetime

from_address = "kandalore.trippers@gmail.com"
password = "shalominthehome"
to_address = "janice@kandalore.com"


#input number you want to search
#number = input('Enter date to find\n')
now = datetime.datetime.now()
exceltup = (now.year, now.month, now.day)
today = int(xlrd.xldate.xldate_from_date_tuple((exceltup),0))
tomorrow = today + 1

intwodays = today + 2
inthreedays = today + 3


context = ssl.create_default_context()
with smtplib.SMTP_SSL(host='smtp.gmail.com', port=465, context=context) as server:
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
                message1 += row['TripID'] + " - " + row['Tripper 1'] + " and the " + row['Section'] + "s on the " + row['Route'] + ", \n"

            if row['Start Date'] == str(tomorrow):
                message2 += row['TripID'] + " - " + row['Tripper 1'] + " and the " + row['Section'] + "s on the " + row['Route'] + ", \n"

            if row['Start Date'] == str(intwodays):
                message3 += row['TripID'] + " - " + row['Tripper 1'] + " and the " + row['Section'] + "s on the " + row['Route'] + ", \n"

            if row['Start Date'] == str(inthreedays):
                message4 += row['TripID'] + " - " + row['Tripper 1'] + " and the " + row['Section'] + "s on the " + row['Route'] + ", \n"

        server.sendmail(from_address,to_address,subject+message1+message2+message3+message4)
