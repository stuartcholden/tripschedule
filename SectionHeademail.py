import csv, smtplib, ssl, xlrd, openpyxl, datetime, subprocess, argparse, n2w

from_address = "kandalore.trippers@gmail.com"
password = "canhnxqvxgdllpxh"
to_address = "tripdirector@kandalore.com"

my_parser = argparse.ArgumentParser()
my_parser.add_argument('-t', '--test', action='store_true', help='Testing mode. All emails are sent to tripdirector@kandalore.com')

args = my_parser.parse_args()

print(vars(args))

if args.test is True:
    print('Wow it worked')


#input number you want to search
#number = input('Enter date to find\n')
now = datetime.datetime.now()
exceltup = (now.year, now.month, now.day)
today = int(xlrd.xldate.xldate_from_date_tuple((exceltup),0))
tomorrow = today + 1
dayleavingtext = "in two days"
intwodays = today + 2
inthreedays = today + 3

tripUID = input("Enter a trip UID:\n")

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)
    with open("email.csv",encoding="utf-8") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['TripID'] == tripUID:

                subject = "Subject:" + row['Section'] + " Section Trips" + "\n\n"

                if not row['TripID'] == "":
                    message1 = "Hi " + row['SectionHeadName'] + ",\n\nThese are the trips in your section:"

                campers = "\n\nHere are your campers, as of Trip Schedule " + row['tripscheduleversion'] + ":\n\n"

                for i in range(1, 30):
                    if row['Camper ' + str(i)] != "":
                        campers += row['Camper ' + str(i)] + "\n"
                    else:
                        campers += ""

                if row['Praise be to Cthulhu, devourer of young love and purveyor of discounted leather jackets'] == "yes":
                    signoff = "\n\nPraise be to Cthulhu, devourer of young love and purveryor of fine discounted leather jackets,\n\nStu"
                else:
                    signoff = "\n\nSincerely,\n\nStu"

                coremessage = subject+message1+campers+signoff
                errormessage = subject+"This is the catchall function. " + row['SectionHeadName'] + " has not received this message."
                noleadtripper = subject + "Trip " + row['TripID'] + " does not have an email for the lead tripper."

                if not row['TripID'] == "":
                    server.sendmail(from_address,[row['SectionHeadEmail']],coremessage)

                if row['email1'] == "":
                    server.sendmail(from_address,row['tripdirectoremail'],noleadtripper)

                if row['Section'] == "Exp":
                    server.sendmail(from_address,row['SectionHeadEmail'],coremessage)
                    server.sendmail(from_address,row['email2'],coremessage)

                if row['Section'] == "X2":
                    server.sendmail(from_address,[row['email1'], row['email2']],"NOT TRUE")


                else:
                    server.sendmail(from_address,row['tripdirectoremail'],errormessage)

                if args.test is True:
                    print('Wow it worked')
