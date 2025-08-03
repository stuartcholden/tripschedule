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

driver = ["Gunnar","Stu","Connor O","Connor M","Haley","Harry","Derek","Raph","Zoe","Grace","Bill","Jamie"]
#do_driver = input("Enter a Driver:\n")
#section = ['PG','JG','IG','SG','PB','JB','IG','SG','Path','LIT']


input_file = open(r'drives.csv')
csv_reader = csv.DictReader(input_file)
driverset = set()
for row in csv_reader:
    if row['Driver'] != "" and row ['Driver'] != "0":
        driverset.add(row['Driver'])
print(driverset)


for x in driverset:
    with open("drives.csv",encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Driver'] == x:
                recipients = row['Driver Email']
                subject = "Your Upcoming Drives"
                standardmessage = "Hi " + row['Driver'] + ","
                signoff = "Sincerely,<br><br>Stu"


    with open("drives.csv",encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        drives = "<br>Here are your upcoming drives:"
        for row in reader:
            if row['Driver'] == x and row['Past'] != "Past":
                xldate = xlrd.xldate_as_datetime(int(row['Date of Drive']), 0).strftime("%b. %d")

                if row['Location Link'] != "":
                    dropofflocation = "<br>Drop Off Location: " + "<a href=" + row['Location Link'] + ">" + row['Location Name'] + "</a>"
                else:
                    dropofflocation = "<br>Drop Off Location: Consult with Stu or Connor"

                if row['Type'] == "Drop Off":
                    typedrive = "<span style=\"font-size:1.2em;color:rgb(20,159,236);\">" + row['Type'] + "</span>"
                elif row['Type'] == "Pick Up":
                    typedrive = "<span style=\"font-size:1.2em;color:rgb(251,0,7);\">" + row['Type'] + "</span>"
                else:
                    typedrive = ""

                if row['Pick Up Time'] != "":
                    pickuptime = "<br>Pick Up Time: " + row['Pick Up Time']
                else:
                    pickuptime = ""

                if row['Driving Time'] != "":
                    drivingtime = "<br>Driving time from Pine Crest: " + row['Driving Time']
                else:
                    drivingtime = "<br>Driving time from Pine Crest: ~u n k n o w n~"

                if row['Van'] != "":
                    van = "<br>Van: " + row['Van']
                else:
                    van = "Van not set yet"

                if row['Changes'] != "":
                    changes = "<br><br> <i><span style=\"color:red;\">Changes since last version</span></i>:<br>" + row['Changes']
                else:
                    changes = ""

                drives += "<br><br>" + typedrive + "<br>" +xldate + "<br>" + '{:0>3}'.format(row['TripID']) + " " + row['Route'] + "<br>Staff: " + row['Staff'] + pickuptime + dropofflocation + drivingtime + van + changes + "<br>"

    message = MIMEMultipart("alternative")
    message["Subject"] = subject
    message["From"] = from_address
    message["To"] = recipients

    html = """\
    <html>
      <body>
        <p>{standardmessage}<br>
           {drives}<br>
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
        server.sendmail("stu.holden@ymcagta.org",str.split(recipients),message.as_string())
