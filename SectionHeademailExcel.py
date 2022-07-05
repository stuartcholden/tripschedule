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

#section = input("Enter a Section:\n")

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)

# To open the workbook
# workbook object is created
    wb_obj = openpyxl.load_workbook(path)
 
# Get workbook active sheet object
# from the active attribute
    sheet_obj = wb_obj.get_sheet_by_name('Trips')
 
# Cell objects also have a row, column,
# and coordinate attributes that provide
# location information for the cell.
 
# Note: The first row or
# column integer is 1, not 0.
 
# Cell object is created by using
# sheet object's cell() method.
    cell_obj = sheet_obj.cell(row = 17, column = 2)
    tripper = str(cell_obj)
 
# Print value of cell object
# using the value attribute
    print(cell_obj.value)

    coremessage = cell_obj.value
    server.sendmail(from_address,"tripdirector@kandalore.com",coremessage)
