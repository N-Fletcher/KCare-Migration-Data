import csv, os
import pandas as pd

#Input required. This value is entered with any data found in the file, and also is used to find the program mapping spreadsheet

agency = 'Live for Life'



myDirectory = os.getcwd()
odysseyDirectory = os.path.abspath(os.path.join(myDirectory, os.pardir))
mappingsDirectory = odysseyDirectory + '\Program Mapping Spreadsheets'

for file in os.listdir(mappingsDirectory):
    if file.__contains__(agency.split()[0]) & file.endswith('.xlsx'):
        input_file = mappingsDirectory + '\\' + file

if input_file == '':
    extendedDirectory = mappingsDirectory + '/Completed Migrations'

    for file in os.listdir(extendedDirectory):
        if file.__contains__(agency.split()[0]) & file.endswith('.xlsx'):
            input_file = mappingsDirectory + '/Completed Migrations/' + file



input = pd.read_excel(input_file, sheet_name='Results - Final Migration')
input.fillna('', inplace=True)




'''External Functions'''
# Function cleans up 'Took' column so that all values can be read as integers


def formatTime(value) -> float:

    if value.strip().__contains__("<1 minute"):
        return 0.5
    return value.split()[0]


# Function calculates updated, new and migrated values, accounting for commas
def calculate(updated, new) -> tuple[int, int, int]:

    # If statements account for values with commas (values in the thousands)
    if updated.__contains__(','):
        updated = int(row[5].strip().replace(',', ''))
    else:
        updated = int(row[5])

    if new.__contains__(','):
        new = int(row[6].strip().replace(',', ''))
    else:
        new = int(row[6])

    migrated = updated + new
    return updated, new, migrated


# Function appends a given row to the main data file
def appendRow(agency, date, system, category, document, total, new,
              updated, migrated, skipped, time, speed):

    # 'a' = append, rather than write which would wipe the file every time
    # by default, csv tool adds an extra line after every row, so 'newline = ""' disbales this
    with open("Migration Data.csv", "a", newline="") as all_data:
        writer = csv.writer(all_data)
        writer.writerow([agency, date, system, category, document, total, new,
                         updated, migrated, skipped, time, speed])


'''Main function'''
# Date should be a static value, so it's defined outside of the scope of loops
date = ''
# script tracks how many rows are appended for quality assurance
newRowsCount = 0

for _, row in input.iterrows():
    # There's only one date for an entire migration, so the code first checks
    # whether the date was already found. If not, it keeps checking each row
    # until it finds the date
    if date == '':
        temp = str(row[3])
        if temp.__contains__('/') | temp.__contains__('-'):
            date = temp.split()[0]

    # Constantly checks for header box that contains all of the migration info,
    # looks for 'FC' and 'GCM' keywords within that box
    if str(row[0]).lower().__contains__('(fc'):
        system = 'FC'
                             # script finds unintended 'fc' characters, '(' combats this
    elif str(row[0]).lower().__contains__('gcm'):
        system = 'GCM'

    # Checks to see whether row contains data first, then runs transfer code
    time_column = str(row[8])
    if (row[2] != '') & (time_column.__contains__('minute') | time_column.strip().isdigit()):
        # Column mappings specifically outlined to avoid headaches
        # later on should formatting change at all
        document = row[2]
        category = row[0]
        total = row[4]
        updated, new, migrated = calculate(str(row[5]), str(row[6]))
        skipped = row[7]
        time = formatTime(row[8])
        speed = float(migrated) / float(time)

        # writing function externalized to avoid quadruple nesting
        appendRow(agency, date, system, category, document, total, new,
                  updated, migrated, skipped, time, speed)

        newRowsCount += 1

print('Appended', newRowsCount, 'rows to data file.')
