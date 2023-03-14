import os, sys
from Analysis_Pipeline import importData

# This is the only user input needed, it sets the name associated with the data that gets imported
agency = 'Elks Aidmore'


#Defines the input file name on the global scope.
# ** This should only be manually set if the desired mapping spreadsheet is in the same folder as the script
input_file = ''

#Defines all directories that contain program mapping spreadsheets to be referenced
myDirectory = os.getcwd()
odysseyDirectory = os.path.abspath(os.path.join(myDirectory, os.pardir))
mappingsDirectory = odysseyDirectory + '/Program Mapping Spreadsheets'
completedMappingsDirectory = mappingsDirectory + '/Completed Migrations'
oldMappingsDirectory = completedMappingsDirectory + '/Migrations pre-2023'

#Searches each directory one by one until the file for the agency is found.
for file in os.listdir(mappingsDirectory):
    if file.__contains__(agency.split()[0]) & file.endswith('.xlsx'):
        input_file = mappingsDirectory + '/' + file

if input_file == '':
    for file in os.listdir(completedMappingsDirectory):
        if file.__contains__(agency.split()[0]) & file.endswith('.xlsx'):
            input_file = completedMappingsDirectory + '/' + file

if input_file == '':
    for file in os.listdir(oldMappingsDirectory):
         if file.__contains__(agency.split()[0]) & file.endswith('.xlsx'):
            input_file = oldMappingsDirectory + '/' + file

# If the file isn't found by now, the import script can't run, so the program is ended
if input_file == '':
    sys.exit('No file found with given agency name')

#Runs script to import data into the analysis spreadsheet
importData(agency, input_file)