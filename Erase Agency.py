import os, csv
import pandas as pd
from openpyxl import load_workbook

'''Enter agency to be wiped from main data file'''
agency = "Grace"

#Create a copy of data file so that data isn't lost while rewriting the main file
df = pd.read_csv('Migration Data.csv')
df.to_csv('Copy of ' + 'Migration Data.csv')

#The copy file will contain all of the data to be read(input), the main file 
#automatically gets wiped and written into(output) using the input from the copy file
input = open('Copy of Migration Data.csv', 'r')
output = open('Migration Data.csv', 'w', newline='')
writer = csv.writer(output)

#script tracks how many rows are discarded for quality assurance
count = 0
for row in csv.reader(input):

    #Condition checks whether each row's agency column contains the specified agency
    #If the row isn't for the specified agency, it goes back onto the file
    if not row[1].__contains__(agency):
        writer.writerow(row[1:])
    else: count+=1

input.close()
output.close()

print('Removed', count, 'rows from data file.')

#Script deletes copy file to avoid clutter within the folder and confusion with the user
os.remove('Copy of Migration Data.csv')