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

#driver = ["Gunnar","Stu","Connor O","Connor M","Haley","Harry","Derek","Raph","Zoe","Grace","Bill","Jamie"]
#do_driver = input("Enter a Driver:\n")
#section = ['PG','JG','IG','SG','PB','JB','IG','SG','Path','LIT']



with open("drives.csv",encoding="utf-8-sig") as file:
    reader = csv.DictReader(file)


    for row in reader:
#        if row['Driver'] != "" and row['Driver'] != "0":
        uniqueNames = set()
        uniqueNames.add(row['Driver'])
#        driverset = set(driver)


#        driverlist = list(driverset)

#for item in driverset:

print(len(uniqueNames))
