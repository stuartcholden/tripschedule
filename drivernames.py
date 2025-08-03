import openpyxl

# Open the spreadsheet
workbook = openpyxl.load_workbook("/Users/Stu/Documents/Tripping/Current/Schedule.xlsm")

# Get the first sheet
sheet = workbook["Driver Email"]

# Create a list to store the values
names = []

# Iterate over the columns in the sheet
for column in sheet.iter_cols():
    # Get the value of the first cell in the column 
    # (the cell with the column name)
    column_name = column[0].value
    # Check if the column is the "Name" column
    if column_name == "OrigDriver":
        # Iterate over the cells in the column
        for i, cell in enumerate(column):
            # Skip the first cell (the cell with the column name)
            if i == 0:
                continue
            # Add the value of the cell to the list
            names.append(cell.value)

# Print the list of names
print(names)
