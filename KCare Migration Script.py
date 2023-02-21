import csv, pandas as pd

##Enter file and agency name here - SEPARATE FC AND GCM DATA BEFORE RUNNING SCRIPT
input_file = 'Elks final mig.csv'
agency = ''
fc_system = bool


def fix_times(file, reader):
    writer = csv.writer(file)

    for row in reader:
        time = row[8].lower()

        if time == "took" #| time.__contains__("total"):
            row[8] = ""
        if time.strip() == "<1 minute":
            row[8] = 1
        else:
            row[8] = row[8].replace("minutes", "").replace("minute", "").strip()

        print(row[8])


with open(input_file) as new:
    reader = csv.reader(new)
    fix_times(new, reader)

    # for row in reader:
    #     if (row[5] > 0 | row[6] > 0) & row[8] > 1:
    #         print(row)
    new.close()

# with open("Migration Data.csv", "a") as all:
#     writer = csv.writer(all)
#     writer.writerow(row)
#     all.close()

