import csv, xlrd, openpyxl, datetime, subprocess, argparse, n2w

path = "/Users/stu/Documents/Tripping/Schedule.xlsm"

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

        if not row['Start Date'] == "":
            daysuntildeparture = int(row['Start Date']) - today

        subject = "Subject: Trip " + row['TripID'] + ", " + " in " + str(n2w.convert(daysuntildeparture)) + " days" + "\n\n"

        pjhelpersubject = "Subject: Help " + row['Staff1EmailName'] + " & " + row['Staff2EmailName'] + " packing out PJ trip " + row['TripID']+"\n\n"

        if not row['PJ Kit'] == "":
            message1 = "Hi " + row['Tripper1HR'] + " & " + row['Staff1EmailName'] + " & " + row['Staff2EmailName'] + " & " + row['Staff3EmailName'] + " & " + row['Staff4EmailName'] + ",\n\nYou're leading the " + row['Route'] + " in " + str(n2w.convert(daysuntildeparture)) + " days"
        elif row['Tripper 2'] == "":
            message1 = "Hi " + row['Tripper1HR'] + ",\n\nYou're leading the " + row['Route'] + " in " + str(n2w.convert(daysuntildeparture)) + " days"
        else:
            message1 = "Hi " + row['Tripper1HR'] + " & " + row['Tripper2HR'] + ",\n\nYou're leading the " + row['Route'] + " " + "in " + str(n2w.convert(daysuntildeparture)) + " days"

        if not row['Staff1EmailName'] == "":
            staff1 = " with " + row['Staff1EmailName']
        else:
            staff1 = "."

        if not row['Staff2EmailName'] == "":
            staff2 = " with " + row['Staff2EmailName']
        else:
            staff2 = "."

        if not row['Staff3EmailName'] == "":
            staff3 = " with " + row['Staff3EmailName']
        else:
            staff3 = "."

        if not row['Staff4EmailName'] == "":
            staff4 = " with " + row['Staff4EmailName']
        else:
            staff4 = "."

        if not row['PJ Helper'] == "":
            pjhelper = " " + row['PJ Helper'] + " will help you pack. You'll be using row['PJ Kit'].\n"
        else:
            pjhelper = ""

        gear = "\n\nHere's what you'll need:\n\nBarrels: " + row['Barrels'] + "\nDish kits: " + row['Dish Kits'] + "\nAquatab kits: " + row['Packs of Aquatabs'] + "\nBoats: " + row['Boats'] + "\nTents: " + row['Tents'] + "\nWhistles: " + row['Whistles'] + "\nRolls of toilet paper (pack the extra one seperately to avoid disaster): " + row['Toilet Paper']  + "\n"

        if not row['Paper Bags'] == "":
            menstrual = "Paper bags: " + row['Paper Bags'] + "\n"
        else:
            menstrual = ""

        if not row['Booking Reference'] == "":
            bookingreference = "Booking Reference Number: " + row['Booking Reference'] + "\n"
        else:
            bookingreference = ""

        if not row['Sites'] == "":
            sites = "Sites: " + row['Sites'] + "\n"
        else:
            sites = ""

        if not row['Drop Off Required'] == "":
            dropoff = "Drop off: " + row['Drop Off Required'] + "\n"
        else:
            dropoff = ""

        if not row['Pick Up Required'] == "":
            pickup = "Pick up: " + row['Pick Up Required'] + "\n"
        else:
            pickup = "" 

        if not row['SPOT'] == "":
            spot = "SPOT: " + row['SPOT'] + "\n"
        else:
            spot = ""

        if not row['Days on Trip'] == "" and int(row['Days on Trip']) > 7:
            extrasatbat = "Extra Sat Phone Battery: 1\n"
        elif not row['Days on Trip'] == "" and int(row['Days on Trip']) > 12:
            extrasatbat += "Extra Sat Phone Batteries: 2\n"
        else:
            extrasatbat = ""

        if not row['Emerg Money'] == "None":
            money = "Emergency Money: $" + row['Emerg Money'] + "\n"
        else:
            money = ""

        if not row['Menu'] == "":
            menu = "Menu: " + row['Menu'] + "\n"
        else:
            menu = ""

        if not row['Trip Notes'] == "":
            tripnotes = "\nAdditional Notes: " + row['Trip Notes'] + "\n"
        else:
            tripnotes = ""

        campers = "\n\nHere are your campers, as of Trip Schedule " + row['tripscheduleversion'] + ":\n\nCabin: " + row['Cabin'] + "\n\n"

        for i in range(1, 30):
            if row['Camper ' + str(i)] != "":
                campers += row['Camper ' + str(i)] + "\n"
            else:
                campers += ""



        if row['Praise be to Cthulhu, devourer of young love and purveyor of discounted leather jackets'] == "yes":
            signoff = "\n\nPraise be to Cthulhu, devourer of young love and purveryor of fine discounted leather jackets,\n\nStu"
        else:
            signoff = "\n\nSincerely,\n\nStu"

        pjmessage = subject+message1+pjhelper+gear+menstrual+bookingreference+sites+dropoff+pickup+spot+money+tripnotes+campers+signoff
        pjhelpermessage = pjhelpersubject + "Hi " + row['PJ Helper'] + ", could you help " + row['Staff1EmailName'] + " & " + row['Staff2EmailName'] + " pack out trip " + row['TripID'] + "on" 
        coremessage = subject+message1+staff1+staff2+gear+menstrual+bookingreference+sites+dropoff+pickup+spot+extrasatbat+money+tripnotes+campers+signoff
        errormessage = subject+"This is the catchall function. " + row['Tripper 1'] + " has not received this message."
        noleadtripper = subject + "Trip " + row['TripID'] + " does not have an email for the lead tripper."

print(coremessage)
