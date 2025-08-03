import openpyxl

book = openpyxl.load_workbook('/Users/Stu/Documents/Tripping/Current/Schedule.xlsm',data_only=True)

sheet = book["Driver Email"]
cells = sheet['O1': 'AB10']

for c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13, c14 in cells:
    print(f"{c1.value} {c2.value} {c3.value} {c4.value} {c9.value}")
