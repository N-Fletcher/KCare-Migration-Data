import pyodbc

con = pyodbc.connect('DRIVER={SQL Server};Server=devdbasg01.kaleidacare.com;Database=eRProjectOdyssey;Port=14.0.3294.2;User ID=DATACENTER\\nfletcher;Password=Seahawks325!')
cursor = con.cursor()

cursor.execute("""
               SELECT *
               FROM dbo.MigrationTaskLog
               WHERE LogLevel = 2
               """)

allEntries = cursor.fetchall()

for row in allEntries:
    id = row[1]
    timestamp = row[2]
    message = row[3]
    print(id, timestamp, message)