import csv, smtplib, ssl, xlrd, openpyxl, datetime, subprocess, argparse, n2w

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

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

section = input("Enter a Section:\n")
#section = ['PG','JG','IG','SG','PB','JB','IG','SG','Path','LIT']


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    for row in reader:
        if row['Section'] == section:
            recipients = row['SectionHeadEmail']
            subject = row['Section'] + " Section Trips"
            standardmessage = "Hi " + row['SectionHeadName'] + ","
            signoff = "Sincerely,<br><br>Stu"


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    changedtrips = "<br><br>Here are the trips in your section that have <i><span style=\"color:rgb(199,117,252);\">changed</span></i> since the last trip schedule:"
    LITneeded = ""
    for row in reader:
        if row['Section'] == section and row['Changed Since Last Version'] == "Yes" and row['Session'].startswith('B') and not row['Trip Status'] == "Done":
            xlstartdate = xlrd.xldate_as_datetime(int(row['Start Date']), 0).strftime("%b. %d")
            xlenddate = xlrd.xldate_as_datetime(int(row['End Date']), 0).strftime("%b. %d") 
            if int(row['Total People on Trip']) % 2 == 0 and not row['Trip Program'] == "PJ":
                LITneeded = ""
            else:
                LITneeded = "<br>LIT Needed"
            if row['PJ Helper'] == "":
                PJHelper = ""
            else:
                PJHelper = "<br><br>" + row['PJ Helper'] + " will help the staff get ready the day before."
            if row['Cabin'] == "0" or row['Cabin'] == 0 or row['Cabin'] == "":
                cabin = ""
            else:
                cabin = "<br>" + row['Cabin'] + ":"
            changedtrips += "<br><br>" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"<br>" + xlstartdate + " to " + xlenddate + "<br>" + "Tripper: " + row['Tripper1HR'] + "<br>" + "Staff: " + row['Staff List for Trip Program'] + LITneeded + PJHelper + "<br>" + "<i>" +  cabin + "</i>"  + "<br>"
            for i in range(1, 30):
                if row['Camper ' + str(i)] != "":
                    changedtrips += row['Camper ' + str(i)] + "<br>"
                else:
                    changedtrips += ""


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    unchangedtrips = "<br><br>And these are the trips in your section that have not changed since the last trip schedule:"
    unchangedLITneeded = ""
    for row in reader:
        if row['Section'] == section and row['Changed Since Last Version'] == "" and row['Session'].startswith('B') and not row['Trip Status'] == "Done":
            xlstartdate = xlrd.xldate_as_datetime(int(row['Start Date']), 0).strftime("%b. %d")
            xlenddate = xlrd.xldate_as_datetime(int(row['End Date']), 0).strftime("%b. %d") 
            if int(row['Total People on Trip']) % 2 == 0 and not row['Trip Program'] == "PJ":
                unchangedLITneeded = ""
            else:
                unchangedLITneeded = "<br>LIT Needed"
            if row['PJ Helper'] == "":
                PJHelper = ""
            else:
                PJHelper = "<br><br>" + row['PJ Helper'] + " will help the staff ready the day before."
            if row['Cabin'] == "0" or row['Cabin'] == 0 or row['Cabin'] == "":
                cabin = ""
            else:
                cabin = "<br>" + row['Cabin'] + ":"
            campers = ""
            for i in range(1, 30):
                if row['Camper ' + str(i)] != "":
                    campers += row['Camper ' + str(i)] + "<br>"
                else:
                    campers == ""
            unchangedtrips += "<br><br>" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"<br>" + xlstartdate + " to " + xlenddate + "<br>" + "Tripper: " + row['Tripper1HR'] + "<br>" + "Staff: " + row['Staff List for Trip Program'] + unchangedLITneeded + PJHelper + "<br>" + "<span style=\"color:rgb(199,117,252);\">" + cabin + "</span>" + "<br>" + "<span style=\"color:violet;\">" + campers + "</span>"

message = MIMEMultipart("alternative")
message["Subject"] = subject
message["From"] = from_address
message["To"] = recipients

html = """\
<html>
  <body>
    <p>{standardmessage}<br>
       {changedtrips}<br>
       {unchangedtrips}<br>
       {signoff}
    </p>
  </body>
</html>
""".format(**locals())

part1 = MIMEText(html, "html")

message.attach(part1)

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address,password)
#    coremessage = subject+standardmessage+trips+unchangedtrips+signoff
    server.sendmail("tripdirector@kandalore.com",str.split(recipients),message.as_string())
