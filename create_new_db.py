from pymongo import MongoClient
import openpyxl

print "Opening DB"
client = MongoClient("mongodb://localhost:27017/")
db = client['test']
print "DB opened"

#Add FOR loop to find each file in the folder, extract data from file, import to DB and continue

print "Loading file"
filepath = 'C:\Users\WillemT\Documents\Work\Projects\Projects 2015\Data Analytics\Python_AMI_Data_Analysis\TrescimoAMIData\\'
filename = '1202124660978'
wb = openpyxl.load_workbook(filepath + filename + '.xlsx')
sheet = wb.get_sheet_by_name('Sheet1')
print "File loaded"
customer = str(sheet.rows[1][0].value)
#db[customer].remove()

print "Starting DB entry"
counti = 0
for i in range(1,11):
    date = str(sheet.rows[i][2].value)[:10]
    time = str(sheet.rows[i][2].value)[-8:-3]
    cursor = db[customer].find({date+'.'+time: {"$exists": True}}).limit(1)
    if cursor.count() == 0:
        db[customer].insert({date: {time: int(sheet.rows[i][3].value)}})
        counti += 1
print str(counti) + " entries added to DB"

#cursor = db[customer].find()
#for document in cursor:
#    print document
