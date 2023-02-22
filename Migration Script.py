import csv

''' Enter file and agency name here - this is the only manual input needed -
Any changes to main data file name or spreadsheet formatting require the code to be updated
'''
input_file = 'CSKF Final.csv'
agency = 'Creative Solutions for Kids & Families'


'''External Functions'''
#Function cleans up 'Took' column so that all values can be read as integers
def formatTime(value) -> float:

    if value.strip() == "<1 minute":
        return 0.5
    return value.replace("minutes", "").replace("minute", "").strip()


#Function calculates updated, new and migrated values, accounting for commas
def calculate(updated, new) -> tuple[int, int, int]:

    #If statements account for values with commas (values in the thousands)
    if updated.__contains__(','):
        updated = int(row[5].strip().replace(',', ''))
    else: updated = int(row[5])
            
    if new.__contains__(','):
        new = int(row[6].strip().replace(',', ''))
    else: new = int(row[6])

    migrated = updated + new
    return updated, new, migrated


#Function appends a given row to the main data file
def appendRow(agency, date, system, category, document, total, new, 
              updated, migrated, skipped, time, speed):
    
    #'a' = append, rather than write which would wipe the file every time
    #by default, csv tool adds an extra line after every row, so 'newline = ""' disbales this
    with open("Migration Data.csv", "a", newline= "") as all_data:
        writer = csv.writer(all_data)
        writer.writerow([agency, date, system, category, document, total, new, 
                         updated, migrated, skipped, time, speed])
        


'''Main function'''
with open(input_file, 'r') as input:
    reader = csv.reader(input)
    date = ''

    for row in reader:
        #There's only one date for an entire migration, so the code first checks
        #whether the date was already found. If not, it keeps checking each row 
        #until it finds the date
        if date == '':
            if row[3].__contains__('/') | row[3].__contains__('-'):
                date = row[3].split()[0]

        #Constantly checks for header box that contains all of the migration info,
        #looks for 'FC' and 'GCM' keywords within that box
        if row[0].lower().__contains__('(fc'): system = 'FC'
                                    #script finds unintended 'fc' characters, '(' combats this
        elif row[0].lower().__contains__('gcm'): system = 'GCM'

        #Checks to see whether row contains data first, then runs transfer code
        if (row[2] != '') & (row[8].__contains__('minute') | row[8].strip().isdigit()):
            #Column mappings specifically outlined to avoid headaches 
            #later on should formatting change at all
            document = row[2]
            category = row[0]
            total = row[4]
            updated, new, migrated = calculate(row[5], row[6])
            skipped = row[7]
            time = formatTime(row[8])
            speed = float(migrated) / float(time)

            #writing function externalized to avoid quadruple nesting
            appendRow(agency, date, system, category, document, total, new,
                      updated, migrated, skipped, time, speed)