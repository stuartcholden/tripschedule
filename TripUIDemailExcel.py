import csv, smtplib, ssl, xlrd, openpyxl, datetime, subprocess, argparse, n2w, warnings

import pandas as pd

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

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)
    with warnings.catch_warnings():
        warnings.filterwarnings("ignore", message="Data Validation extension is not supported")
        df = pd.read_excel('/Users/stu/Documents/Tripping/Schedule.xlsm', sheet_name='Trips')


        for row in df['0']:
            if int(row) == int(tripUID):
                subject = "Tripper is " + df.loc[df['0'] == int(tripUID)]['Tripper1HR'].values
                print(subject)
#    subject = df[df['TripID'] == int(tripUID)]
#    print(subject)
