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
            if row['Section'] == "PJG" and row['Start Date'] >= str(today):
                subject = "Subject: PJ Girls Trips that haven't left yet\n\n"
                greeting = "Hi " + row['SectionHeadName'] + ",\n\n"
                message1 = row['TripID'] + row['Tripper 1'] + row['Route']

                for i in range(30):
                    if row['Camper {i}'] != "":
                        campers = "\n" + row['Camper {i}']
                    else:
                        campers = ""

                message = subject + greeting + message1 + campers + "\n"
#        server.sendmail(from_address,to_address,message
                print(message)
