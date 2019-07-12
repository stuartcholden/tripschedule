import csv, smtplib, ssl

message = """Subject: Your grade

Hi {Trip}, your grade is {section}"""
from_address = "kandalore.trippers@gmail.com"
password = "standardissuechotch"

context = ssl.create_default_context()
with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
    server.login(from_address, password)
    with open("email.csv") as file:
        reader = csv.reader(file)
        next(reader)  # Skip header row
        for Trip, email, Section in reader:
            server.sendmail(
                from_address,
                email,
                message.format(name=name,grade=grade),
            )
