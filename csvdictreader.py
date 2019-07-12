import csv
with open('email.csv', 'rb') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        print(row['Tripper 1'], row['Trip'])
