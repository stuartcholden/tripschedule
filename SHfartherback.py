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

with open("email.csv",encoding="utf-8-sig") as file2:
    reader2 = csv.DictReader(file2)
    for row2 in reader2:
        if row2['Section'] == "LIT":
            core = ""
            core += row2['TripID']+"three"
            print(core)

print(core)

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)
    with open("email.csv",encoding="utf-8-sig") as file:
        reader = csv.DictReader(file)
        for row in reader:
            if row['Section'] == "LIT":
#                print(core)
                subject = "Subject:" + row['Section'] + " Section Trips" + "\n\n"
                message1 = "Hi " + row['SectionHeadName'] + ",\n\nThese are the trips in your section:" + row['TripID']

                signoff = "\n\nSincerely,\n\nStu"

                coremessage = subject+message1+signoff
                errormessage = subject+"This is the catchall function. " + row['SectionHeadName'] + " has not received this message."

        server.sendmail(from_address,"tripdirector@kandalore.com",coremessage.encode('utf-8'))


