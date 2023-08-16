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


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    for row in reader:
        if row['LITs on Trip'] == "LIT":
            to_address = row['LITDirectorEmails']
            standardmessage = "Hi " + row['LITDirectorNames'] + ","
            signoff = "<br><br>Sincerely,<br><br>Stu"


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    tripsneedinglits = "\n\nHere are the trips that need LITs:"
    for row in reader:
        if row['LIT Needed'] == "LIT Needed" and row['Session'].startswith('B') and not row['Trip Status'] == "Done" and not row['Trip Status'] == "Out":
            xlstartdate = xlrd.xldate_as_datetime(int(row['Start Date']), 0).strftime("%b. %d")
            xlenddate = xlrd.xldate_as_datetime(int(row['End Date']), 0).strftime("%b. %d") 
            if row['Cabin'] == "0" or row['Cabin'] == 0 or row['Cabin'] == "":
                cabin = ""
            else:
                cabin = "Cabin(s): " + row['Cabin']
            section = row['Section Full Name']
            tripsneedinglits += "<br><br>" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"<br>" + xlstartdate + " to " + xlenddate + "<br>" + "Tripper: " + row['Tripper1HR'] + "<br>" + "Staff: " + row['Staff List for Trip Program'] + "<br>Section: " + section + "<br>" + cabin + "<br>"

with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    changedtrips = "<br><br>Here are the trips with LITs that have <i><span style=\"color:rgb(199,117,252);\">changed</span></i> since the last trip schedule:"
    LITneeded = ""
    for row in reader:
        if row['LITs on Trip'] == "LIT" and row['Changed Since Last Version'] == "Yes" and row['Session'].startswith('B') and not row['Trip Status'] == "Done":
            xlstartdate = xlrd.xldate_as_datetime(int(row['Start Date']), 0).strftime("%b. %d")
            xlenddate = xlrd.xldate_as_datetime(int(row['End Date']), 0).strftime("%b. %d") 
            if row['Cabin'] == "0" or row['Cabin'] == 0 or row['Cabin'] == "":
                cabin = ""
            else:
                cabin = "Cabin(s): " + row['Cabin']
            section = row['Section Full Name']
            changedtrips += "<br><br>" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"<br>" + xlstartdate + " to " + xlenddate + "<br>" + "Tripper: " + row['Tripper1HR'] + "<br>" + "Staff: " + row['Staff List for Trip Program'] + "<br>" + "Section: " + section + cabin + "<br><br>LIT(s):<br>"
            for i in range(1, 4):
                if row['LIT ' + str(i)] != "" in row['LIT ' +str(i)]:
                    changedtrips += row['LIT ' + str(i)] + "<br>"
                else:
                    changedtrips += ""


with open("email.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)
    unchangedtrips = "<br><br>And these are the trips with LITs that have not changed since the last trip schedule:"
    unchangedLITneeded = ""
    for row in reader:
        if row['LITs on Trip'] == "LIT" and row['Changed Since Last Version'] == "" and row['Session'].startswith('B') and not row['Trip Status'] == "Done":
            xlstartdate = xlrd.xldate_as_datetime(int(row['Start Date']), 0).strftime("%b. %d")
            xlenddate = xlrd.xldate_as_datetime(int(row['End Date']), 0).strftime("%b. %d") 
            if row['Cabin'] == "0" or row['Cabin'] == 0 or row['Cabin'] == "":
                cabin = ""
            else:
                cabin = row['Cabin']
            section = row['Section Full Name']
            unchangedtrips += "<br><br>" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] +"<br>" + xlstartdate + " to " + xlenddate + "<br>" + "Tripper: " + row['Tripper1HR'] + "<br>" + "Staff: " + row['Staff List for Trip Program'] + "<br>Section: " + section + "<br>Cabin(s): " + cabin + "<br><br>LIT(s):<br>"
            for i in range(1, 4):
                if row['LIT ' + str(i)] != "" in row['LIT ' +str(i)]:
                    unchangedtrips += row['LIT ' + str(i)] + "<br>"
                else:
                    unchangedtrips += ""

message = MIMEMultipart("alternative")
message["Subject"] = "Trips involving LITs"
message["From"] = from_address
message["To"] = to_address

html = """\
<html>
  <body>
    <p>{standardmessage}<br><br>
       {tripsneedinglits}<br>
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
#    coremessage = subject+standardmessage+tripsneedinglits+trips+unchangedtrips+signoff
    server.sendmail(from_address,str.split(to_address),message.as_string())
