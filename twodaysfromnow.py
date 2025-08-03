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
tomorrow = today - 8
dayleavingtext = "in two days"
intwodays = today + 2
inthreedays = today + 3
infourdays = today + 4


context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address,password)
    with open("email.csv",encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Start Date'] == str(tomorrow):

                subject = "Trip " + row['TripID'] + ", " + dayleavingtext


                if row['Trip Program'] == 'Back Lakes':
                    message1 = "Hi " + row['Tripper1HR'] + " & " + row['Staff1EmailName'] + ",<br><br>You're leading the " + "<span style=\"color:darkviolet;\">" + row['Route'] + "</span> " + dayleavingtext + "."
                if row['Trip Program'] == 'Camper Trips':
                    message1 = "Hi " + row['PETrip Leaders'] + ",<br><br>You're leading the " + "<span style=\"color:darkviolet;\">" + row['Route'] + "</span> " + dayleavingtext + " with the " + row['Cabin'] + "."
                elif row['Trip Program'] == 'Leadership':
                    message1 = "Hi " + row['DBTrip Leaders'] + ",<br><br>You're leading the " + "<span style=\"color:darkviolet;\">" + row['Route'] + "</span> " + " " + dayleavingtext + "."
#                else:
#                    message1 = "Hi " + row['Tripper1HR'] + " & " + row['Tripper2HR'] + ",<br><br>You're leading the " + "<span style=\"color:darkviolet;\">" + row['Route'] + "</span> " + " " + dayleavingtext

                if not row['Staff1EmailName'] == "":
                    staff1 = " with " + row['Staff1EmailName'] + "."
                else:
                    staff1 = ""

                if not row['Staff2EmailName'] == "":
                    staff2 = " & " + row['Staff2EmailName'] + "."
                else:
                    staff2 = ""

#                if not row['PJ Helper'] == "":
#                    pjhelper = ". " + row['PJ Helper'] + " will help you pack"
#                else:
#                    pjhelper = ""

                if row['Barrels'] == "0":
                    barrels = ""
                else:
                    barrels = "<br>Barrels: " + row['Barrels']

                gear = "<br><br>Here's what you'll need:<br>" + barrels + "<br>Green Scrubbies: " + row['Green Scrubbies'] + "<br>Steel Wool: " + row['Steel Wool'] + "<br>Aquatab packs of 50: " + row['Packs of Aquatabs'] + "<br>Boats: " + row['Boats'] + "<br>Tents: " + row['Tents'] + "<br>Whistles: " + row['Whistles'] + "<br>Rolls of toilet paper (pack the extra one seperately to avoid" + "<span style=\"color:crimson;\"> disaster</span>" + "): " + row['Toilet Paper']  + "<br>"

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

                if not row['Drop Off Required'] == "":
                    dropoff = "Drop off: " + row['DBTrip Drop Off Driver'] + "<br>"
                else:
                    dropoff = ""

                if not row['Pick Up Required'] == "":
                    pickup = "Pick up: " + row['DBTrip Pick Up Driver'] + "<br>"
                else:
                    pickup = "" 

                if not row['SPOT or InReach'] == "":
                    spot = "SPOT or InReach: " + row['SPOT or InReach'] + "<br>"
                else:
                    spot = ""

                if not row['Phone'] == "":
                    phone = "Phone: " + row['Phone'] + " (" + row['phone_numbers'] + ")" + "<br>"
                else:
                    phone = ""

                if not row['extra_sat_bat'] == "":
                    extrasatbat = "Extra Sat Phone Batteries: " + row['extra_sat_bat'] + "<br>"
                else:
                    extrasatbat = ""

                if not row['Emerg Money'] == "":
                    money = "Emergency Money: " + row['Emerg Money'] + "<br>"
                else:
                    money = ""

                if not row['Menu'] == "":
                    menu = "Menu: " + row['Menu'] + "<br>"
                else:
                     menu = ""

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
                    signoff = "<br><br>Sincerely,<br><br>Stu"

                campertripsmessage = MIMEMultipart("alternative")
                campertripsmessage["Subject"] = subject
                campertripsmessage["From"] = from_address
                campertripsmessage["To"] = row['email1']

                html = """\
                <html>
                  <body>
                    <p>
                       {message1}
                       {gear}
                       {paperbags}
                       {phone}
                       {extrasatbat}
                       {spot}
                       {money}
                       <br><br>Logistics<br><br>
                       {bookingreference}
                       {sites}
                       {dropoff}
                       {pickup}
                       {notes}
                       {signoff}
                    </p>
                  </body>
                </html>
                """.format(**locals())

                part1 = MIMEText(html, "html")

                campertripsmessage.attach(part1)

                lsmessage = MIMEMultipart("alternative")
                lsmessage["Subject"] = subject
                lsmessage["From"] = from_address
                lsmessage["To"] = row['email1']

                html = """\
                <html>
                  <body>
                       {message1}
                       {gear}
                       {paperbags}
                       {bookingreference}
                       {sites}
                       {dropoff}
                       {pickup}
                       {menu}
                       {notes}
                       {signoff}
                    </p>
                  </body>
                </html>
                """.format(**locals())

                part1 = MIMEText(html, "html")

                lsmessage.attach(part1)
                
                pjmessage = MIMEMultipart("alternative")
                pjmessage["Subject"] = subject
                pjmessage["From"] = from_address
                pjmessage["To"] = row['email1']

                html = """\
                <html>
                  <body>
                       {gear}
                       {paperbags}
                       {bookingreference}
                       {sites}
                       {dropoff}
                       {pickup}
                       {spot}
                       {extrasatbat}
                       {money}
                       {menu}
                       {notes}
                       {signoff}
                    </p>
                  </body>
                </html>
                """.format(**locals())

                part1 = MIMEText(html, "html")

                pjmessage.attach(part1)

#                pjmessage = subject+message1+pjhelper+gear+paperbags+bookingreference+sites+dropoff+pickup+spot+notes+campers+signoff
#                coremessage = subject+message1+staff1+staff2+gear+paperbags+bookingreference+sites+dropoff+pickup+spot+extrasatbat+money+menu+notes+campers+signoff
#                expmessage = subject+message1+staff1+staff2+gear+paperbags+bookingreference+sites+dropoff+pickup+spot+extrasatbat+money+menu+notes+campers+signoff

                if row['email1'] == "":
                    server.sendmail(from_address,row['tripdirectoremail'],"Subject: Trip " + row['TripID'] + " does not have an email for the lead tripper\n\nHopefully there's at least a tripper.")
                if row['Trip Program'] == "Leadership":
                    server.sendmail(from_address,[row['email1'], row['email2']],lsmessage.as_string())
                if row['Trip Program'] == "Back Lakes" and row['Lead by a Tripper'] == "No":
                    server.sendmail(from_address,[row['email1'], row['email2'],],pjmessage.as_string())
                if row['Trip Program'] == "Back Lakes" and row['Lead by a Tripper'] == "Yes":
                    server.sendmail(from_address,row['email1'],pjmessage.as_string())
                if row['Lead by an X2'] == "Yes":
                    server.sendmail(from_address,[row['email1'], row['x2directoremail']],pjmessage.as_string())
                if row['Trip Program'] == "Camper Trips":
                    server.sendmail(from_address,row['email1'],campertripsmessage.as_string())
