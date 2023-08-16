import csv, smtplib, ssl, xlrd, openpyxl, datetime, subprocess, argparse, n2w

to_address = "tripdirector@kandalore.com"

import gmailapppassword
from_address = gmailapppassword.username
password = gmailapppassword.password

my_parser = argparse.ArgumentParser()
my_parser.add_argument('-t', '--test', action='store_true', help='Testing mode. All emails are sent to tripdirector@kandalore.com')

args = my_parser.parse_args()

print(vars(args))

if args.test is True:
    print('Wow it worked')

path = "/Users/stu/Documents/Tripping/Schedule.xlsm"

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

import emailconstructor
coremessage = emailconstructor.coremessage

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)
    with open("email.csv",encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['TripID'] == tripUID:

                if row['Section'] == "PG" or row['Section'] == "JG" or row['Section'] == "PB" or row['Section'] == "JB" and row['Tripper 1'].startswith('PJ'):
                    server.sendmail(from_address,[row['email1'], row['email2'],row['SectionHeadEmail'],row['pjhelperemail'],row['DirectorofCampLifeEmail']],pjmessage.encode('utf-8'))

                if row['Section'] == "PG" or row['Section'] == "JG" or row['Section'] == "PB" or row['Section'] == "JB" and not row['Tripper 1'].startswith('PJ'):
                    server.sendmail(from_address,row['email1'],pjmessage.encode('utf-8'))

                if row['email1'] == "":
                    server.sendmail(from_address,row['tripdirectoremail'],noleadtripper.encode('utf-8'))

                if row['Section'] == "Exp":
                    server.sendmail(from_address,row['email1'],coremessage.encode('utf-8'))
                    server.sendmail(from_address,row['email2'],coremessage.encode('utf-8'))

                if row['Section'] == "X2":
                    server.sendmail(from_address,[row['email1'],row['email2']],coremessage.encode('utf-8')),

                if row['Trip Program'] == "Core":
                    server.sendmail(from_address,[row['email1'],row['email2']],coremessage.encode('utf-8'))


                if args.test is True:
                    print('Wow it worked')
