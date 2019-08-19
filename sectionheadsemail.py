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
        subject = "Subject: PJ Girls Trips that haven't left yet\n\n"
#        greeting = "Hi " + row['SectionHeadName'] + ",\n\n"
        greeting = "Hi Sophie,\n\n"
        print(subject + greeting)
        for row in reader:
            if row['Section'] == "PJG" and row['Start Date'] >= str(today):
                message1 = row['TripID'] + " " + row['Route'] + "\n\n"
                tripper = "Tripper: " + row['Tripper 1'] + "\n"
                staff = "Staff: " + row['Staff 1'] + "\n\n"
                dates = "Dates: " + row['Start Date'] + " to " + row['End Date'] + "\n\n"
                campers = "Campers:\n"

                for i in range(1, 30):
                    if row['Camper '+ str(i)] != "":
                        campers += row['Camper '+ str(i)] + "\n"
                    else:
                        campers += ""

                message = message1 + dates + tripper + staff + campers + "\n"
#        server.sendmail(from_address,to_address,message
                print(message)
