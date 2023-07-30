import csv, smtplib, ssl, xlrd, openpyxl, datetime, subprocess, argparse, n2w

import gmailapppassword
from_address = gmailapppassword.username
password = gmailapppassword.password


now = datetime.datetime.now()
exceltup = (now.year, now.month, now.day)
today = int(xlrd.xldate.xldate_from_date_tuple((exceltup),0))
tomorrow = today + 1
dayleavingtext = "in two days"
intwodays = today + 2
inthreedays = today + 3


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    for row in reader:
        if not row['Nights at Coop'] == "":
            to_address = row['LITDirectorEmails']
            subject = "Subject: Kandalore Trips at the Coop" + "\n\n"
            standardmessage = "Hi " + row['LITDirectorNames'] + ","
            signoff = "\n\nSincerely,\n\nStu"


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    trips = "\n\nHere are the trips with LITs that have changed since the last trip schedule:"
    LITneeded = ""
    for row in reader:
        if row['LITs on Trip'] == "LIT" and row['Changed Since Last Version'] == "Yes" and row['Session'].startswith('A') and not row['Trip Status'] == "Done":
            xlstartdate = xlrd.xldate_as_datetime(int(row['Start Date']), 0).strftime("%b. %d")
            xlenddate = xlrd.xldate_as_datetime(int(row['End Date']), 0).strftime("%b. %d") 
            if int(row['Total People on Trip']) % 2 == 0 and not row['Trip Program'] == "PJ":
                LITneeded = ""
            else:
                LITneeded = "\nLIT Needed"
            if row['Cabin'] == "0" or row['Cabin'] == 0 or row['Cabin'] == "":
                cabin = ""
            else:
                cabin = "\n" + row['Cabin'] + ":"
            trips += "\n\n" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"\n" + xlstartdate + " to " + xlenddate + "\n" + "Tripper: " + row['Tripper1HR'] + "\n" + "Staff: " + row['Staff List for Trip Program'] + LITneeded + "\n" + "\n" + "\nLIT(s):\n"
            for i in range(1, 4):
                if row['LIT ' + str(i)] != "" in row['LIT ' +str(i)]:
                    trips += row['LIT ' + str(i)] + "\n"
                else:
                    trips += ""


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    unchangedtrips = "\n\nAnd these are the trips wiht LITs that have not changed since the last trip schedule:"
    unchangedLITneeded = ""
    for row in reader:
        if row['LITs on Trip'] == "LIT" and row['Changed Since Last Version'] == "" and row['Session'].startswith('A') and not row['Trip Status'] == "Done":
            xlstartdate = xlrd.xldate_as_datetime(int(row['Start Date']), 0).strftime("%b. %d")
            xlenddate = xlrd.xldate_as_datetime(int(row['End Date']), 0).strftime("%b. %d") 
            if int(row['Total People on Trip']) % 2 == 0 and not row['Trip Program'] == "PJ":
                unchangedLITneeded = ""
            else:
                unchangedLITneeded = "\nLIT Needed"
            if row['Cabin'] == "0" or row['Cabin'] == 0 or row['Cabin'] == "":
                cabin = ""
            else:
                cabin = "\n" + row['Cabin'] + ":"
            unchangedtrips += "\n\n" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"\n" + xlstartdate + " to " + xlenddate + "\n" + "Tripper: " + row['Tripper1HR'] + "\n" + "Staff: " + row['Staff List for Trip Program'] + unchangedLITneeded + "\n" + "\nLIT(s):\n"
            for i in range(1, 4):
                if row['LIT ' + str(i)] != "" in row['LIT ' +str(i)]:
                    unchangedtrips += row['LIT ' + str(i)] + "\n"
                else:
                    unchangedtrips += ""


context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address,password)
    coremessage = subject+standardmessage+trips+unchangedtrips+signoff
    server.sendmail(from_address,to_address,coremessage.encode('utf-8'))
