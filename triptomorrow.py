#!/usr/bin/env /opt/homebrew/bin/python3.9
import csv, smtplib, ssl, xlrd, datetime, subprocess

from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import gmailapppassword
from_address = gmailapppassword.username
password = gmailapppassword.password

#input number you want to search
#number = input('Enter date to find\n')
now = datetime.datetime.now()
exceltup = (now.year, now.month, now.day)
today = int(xlrd.xldate.xldate_from_date_tuple((exceltup),0))
tomorrow = today + 1
dayleavingtext = "tomorrow"
intwodays = today + 2
inthreedays = today + 3
infourdays = today + 4


context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address,password)
    with open("/users/stu/git/tripschedule/email.csv",encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Start Date'] == str(tomorrow):

                subject = "Trip " + row['TripID'] + ", " + dayleavingtext


                if row['Trip Program'] != 'Leadership' and row['Trip Active'] == "Yes":
                    message1 = "Hi " + row['PETrip Leaders'] + ",<br><br>You're leading the " + "<span style=\"color:darkviolet;\">" + row['Route'] + "</span> " + dayleavingtext + " with the " + row['Cabin'] + "."
                else:
                    message1 = ""

                if row['60l Barrels'] == "0" or row['60l Barrels'] == "None":
                    barrels = ""
                else:
                    barrels = "<br>60l Barrels: " + row['60l Barrels']

                if row['30l Barrels'] == "0" or row['30l Barrels'] == "None":
                    halfbarrels = ""
                else:
                    halfbarrels = "<br>30l Barrels: " + row['30l Barrels']

                gear = "<br><br>Here's what you'll need:<br>" + barrels +  halfbarrels + "<br>Green Scrubbies: " + row['Green Scrubbies'] + "<br>Steel Wool: " + row['Steel Wool'] + "<br>Aquatab packs of 50: " + row['Packs of Aquatabs'] + "<br>Boats: " + row['Boats'] + "<br>Tents: " + row['Tents'] + "<br>Whistles: " + row['Whistles'] + "<br>Rolls of toilet paper (pack the extra one seperately to avoid" + "<span style=\"color:crimson;\"> disaster</span>" + "): " + row['Toilet Paper Needed HR']  + "<br>"

                if not row['Paper Bags'] == "":
                    paperbags = "Paper bags: " + row['Paper Bags'] + "<br>"
                else:
                    paperbags = ""

                if not row['Booking Reference'] == "":
                    bookingreference = "Booking Reference Number: " + row['Booking Reference'] + "<br>"
                else:
                    bookingreference = ""

                if not row['Sites'] == "":
                    sites = "Sites: " + row['Sites'] + "<br>"
                else:
                    sites = ""

                if row['Drop Off Required'] != "":
                    dropofflocationlink = "<a href=" + row['drop off link'] + ">" + row['drop off name'] + "</a><br>"
                else:
                    dropofflocationlink = ""

                if not row['Drop Off Required'] == "" and row['Drop Off Van'] == "Scarlett Jovansson":
                    dovanwithcolour = "<br>Drop off: " + "<span style=\"color:crimson;\">" + row['Drop Off Van'] + "</span>"
                elif not row['Drop Off Required'] == "" and row['Drop Off Van'] == "Van der Waals Force":
                    dovanwithcolour = "<br>Drop off: " + "<span style=\"color:grey;\">" + row['Drop Off Van'] + "</span>"
                elif not row['Drop Off Required'] == "" and row['Drop Off Van'] == "Van Diesel":
                    dovanwithcolour = "<br>Drop off: " + "<span style=\"color:blue;\">" + row['Drop Off Van'] + "</span>"
                elif not row['Drop Off Required'] == "" and row['Drop Off Van'] == "Bluce Willis":
                    dovanwithcolour = "<br>Drop off: " + "<span style=\"color:blue;\">" + row['Drop Off Van'] + "</span>"
                else:
                    dovanwithcolour = ""

                if row['Drop Off Required'] == "":
                    dodriverandtime = ""
                elif row['Drop Off Required'] == "Canoe Docks" or row['Drop Off Required'] == "OP":
                    dodriverandtime = "<br>Departure: Paddling Out"
                else:
                    dodriverandtime = " feat. " +  row['Drop Off Driver'] + " at " + row['drop_off_time'] + "<br>" + row['DBDropOffDriveTime'] + " drive to"

                if row['Pick Up Required'] != "":
                    pickuplocationlink = "<a href=" + row['pick up link'] + ">" + row['pick up name'] + "</a><br>"
                else:
                    pickuplocationlink = ""

                if not row['Pick Up Required'] == "" and row['Pick Up Van'] == "Scarlett Jovansson":
                    puvanwithcolour = "<br>Pick Up: " + "<span style=\"color:crimson;\">" + row['Pick Up Van'] + "</span>"
                elif not row['Pick Up Required'] == "" and row['Pick Up Van'] == "Van der Waals Force":
                    puvanwithcolour = "<br>Pick Up: " + "<span style=\"color:grey;\">" + row['Pick Up Van'] + "</span>"
                elif not row['Pick Up Required'] == "" and row['Pick Up Van'] == "Van Diesel":
                    puvanwithcolour = "<br>Pick Up: " + "<span style=\"color:blue;\">" + row['Pick Up Van'] + "</span>"
                elif not row['Pick Up Required'] == "" and row['Pick Up Van'] == "Bluce Willis":
                    puvanwithcolour = "<br>Pick Up: " + "<span style=\"color:blue;\">" + row['Pick Up Van'] + "</span>"
                else:
                    puvanwithcolour = ""

                if row['Pick Up Required'] == "":
                    pudriverandtime = ""
                elif row['Pick Up Required'] == "Canoe Docks" or row['Pick Up Required'] == "OP":
                    pudriverandtime = "<br>Return: Paddling In"
                else:
                    pudriverandtime = " feat. " +  row['Pick Up Driver'] + " at " + row['pick_up_time'] + "<br>" + row['DBPickUpDriveTime'] + " drive to"

                if not row['SPOT or InReach'] == "":
                    spot = "SPOT or InReach: " + row['SPOT or InReach'] + "<br>"
                else:
                    spot = ""

                if not row['DBPhoneNumber'] == "":
                    phone = "Phone: " + row['DBPhoneNumber'] + "<br>"
                else:
                    phone = ""

                if not row['extra_sat_bat'] == "":
                    extrasatbat = "Extra Sat Phone Batteries: " + row['extra_sat_bat'] + "<br>"
                else:
                    extrasatbat = ""

                if not row['SPOT or InReach'] == "":
                    gpsdevice = "SPOT or InReach: " + row['SPOT or InReach'] + "<br>"
                else:
                    gpsdevice = ""

                if not row['Emerg Money'] == "":
                    money = "Emergency Money: $" + row['Emerg Money'] + "<br>"
                else:
                    money = ""

                if not row['Trip Notes'] == "":
                    notes = "<br>Notes: " + row['Trip Notes'] + "<br>"
                else:
                    notes = ""

#                for i in range(1, 30):
#                    if row['Camper ' + str(i)] != "":
#                        campers += row['Camper ' + str(i)] + "<br>"
#                    else:
#                        campers += ""

                if row['Praise be to Cthulhu, devourer of young love and purveyor of discounted leather jackets'] == "yes":
                    signoff = "<br><br>Praise be to Cthulhu, devourer of young love and purveryor of fine discounted leather jackets,<br><br>Stu"
                else:
                    signoff = "<br><br>In camp friendship,<br><br>Stu"

                
                message = MIMEMultipart("alternative")
                message["Subject"] = subject
                message["Reply-To"] = "stu.holden@ymcagta.org"
#                message["To"] = row['email1'] + ", " + row['email2']

                html = """\
                <html>
                  <body>
                       {message1}
                       {gear}
                       {paperbags}
                       {bookingreference}
                       {sites}
                       {phone}
                       {extrasatbat}
                       {gpsdevice}
                       {money}
                       {notes}
                       {dovanwithcolour}
                       {dodriverandtime}
                       {dropofflocationlink}
                       {puvanwithcolour}
                       {pudriverandtime}
                       {pickuplocationlink}
                       {signoff}
                    </p>
                  </body>
                </html>
                """.format(**locals())

                part1 = MIMEText(html, "html")

                message.attach(part1)
                
                if row['email1'] == "":
                    server.sendmail(from_address,row['tripdirectoremail'],"Subject: Trip " + row['TripID'] + " does not have an email for the lead tripper\n\nHopefully there's at least a tripper.")
                if row['Trip Active'] == "Yes" and row['Trip Program'] != "Leadership":
                    server.sendmail(from_address,[row['email1'],row['email2'],row['email3']],message.as_string())
