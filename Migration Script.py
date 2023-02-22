import csv, pandas as pd

#Enter file and agency name here
input_file = 'CSKF Final.csv'
agency = 'CSKF'


#Function cleans up 'Took' column so that all values can be read as ints
def formatTime(value) -> int:

    if value.strip() == "<1 minute":
        return 0.5
    return value.replace("minutes", "").replace("minute", "").strip()


#Function appends a given row to the main data file
def appendRow(agency, date, system, category, document, total, new, 
              updated, migrated, skipped, time, speed):
    
    # with open("Migration Data.csv", "a") as all_data:
    #     writer = csv.writer(all_data)
    with open("tester.csv", "a") as all_data:
        writer = csv.writer(all_data)

        writer.writerow([agency, date, system, category, document, total, new, 
                         updated, migrated, skipped, time, speed])
        

#Function pulls the date of the migration out of the spreadsheet
#The date remains the same for the entire migration, so no need to record it for every row
def findDate(reader) -> str:
    for row in reader:
        if len(row) > 2:
            date = row[3]
            if date.__contains__('/'):
                return date.split()[0]
        


#Main
with open(input_file, 'r') as input:
    reader = csv.reader(input)
    date = findDate(reader)
    system = str

    for row in reader:
        print(row)
        document = row[2]

        #Constantly checks for header box that contains all of the migration info,
        #looks for 'FC' and 'GCM' keywords within that box
        if row[0].lower().__contains__('(fc/res)'): system = 'FC'
        elif row[0].lower().__contains__('gcm'): system = 'GCM'
        # print(isFC)

        #Checks to see whether row contains data
        if (document != '') & (row[5].isdigit()):
            category = row[0]
            total = row[4]

            if row[5].__contains__(','):
                updated = int(row[5].strip().replace(',', ''))
            else: updated = int(row[5])
            
            if row[6].__contains__(','):
                new = int(row[6].strip().replace(',', ''))
            else: new = int(row[6])
            
            migrated = updated + new
            skipped = row[7]
            time = formatTime(row[8])
            speed = float(migrated) / float(time)

            # print(agency, date, system, category, document, total, new,
            #            updated, migrated, skipped, time, speed)
            appendRow(agency, date, system, category, document, total, new,
                      updated, migrated, skipped, time, speed)