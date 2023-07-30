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
        if row['International Section'] == "IL":
            to_address = row['InternationalLiaisonEmail']
            subject = "Subject: Trips with International Campers" + "\n\n"
            standardmessage = "Hi " + row['InternationalLiaisonName'] + ","
            signoff = "\n\nSincerely,\n\nStu"


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    trips = "\n\nHere are the trips in your section that have changed since the last trip schedule:"
    LITneeded = ""
    for row in reader:
        if row['International Section'] == "IL" and row['Changed Since Last Version'] == "Yes" and row['Session'].startswith('A') and not row['Trip Status'] == "Done":
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
            trips += "\n\n" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"\n" + xlstartdate + " to " + xlenddate + "\n" + "Tripper: " + row['Tripper 1'] + "\n" + "Staff: " + row['Staff List for Trip Program'] + LITneeded + "\n" + cabin  + "\n"
            for i in range(1, 30):
                if row['Camper ' + str(i)] != "" and "- I" in row['Camper ' +str(i)]:
                    trips += row['Camper ' + str(i)] + "\n"
                else:
                    trips += ""


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    unchangedtrips = "\n\nAnd these are the trips in your section that have not changed since the last trip schedule:"
    unchangedLITneeded = ""
    for row in reader:
        if row['International Section'] == "IL" and row['Changed Since Last Version'] == "" and row['Session'].startswith('A') and not row['Trip Status'] == "Done":
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
            unchangedtrips += "\n\n" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"\n" + xlstartdate + " to " + xlenddate + "\n" + "Tripper: " + row['Tripper 1'] + "\n" + "Staff: " + row['Staff List for Trip Program'] + unchangedLITneeded + "\n" + cabin  + "\n"
            for i in range(1, 30):
                if row['Camper ' + str(i)] != "" and "- I" in row['Camper ' +str(i)]:
                    unchangedtrips += row['Camper ' + str(i)] + "\n"
                else:
                    unchangedtrips += ""


context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address,password)
    coremessage = subject+standardmessage+trips+unchangedtrips+signoff
    server.sendmail(from_address,to_address,coremessage.encode('utf-8'))
