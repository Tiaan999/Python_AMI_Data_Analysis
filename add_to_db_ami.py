import pymongo
import datetime as dt
from datetime import datetime
import os
import xlrd

main_start = datetime.now()
db = pymongo.MongoClient("mongodb://localhost:27017/")['test']
time_half_hour = ['00:00', '00:30', '01:00', '01:30', '02:00', '02:30', '03:00', '03:30',
                  '04:00', '04:30', '05:00', '05:30', '06:00', '06:30', '07:00', '07:30',
                  '08:00', '08:30', '09:00', '09:30', '10:00', '10:30', '11:00', '11:30',
                  '12:00', '12:30', '13:00', '13:30', '14:00', '14:30', '15:00', '15:30',
                  '16:00', '16:30', '17:00', '17:30', '18:00', '18:30', '19:00', '19:30',
                  '20:00', '20:30', '21:00', '21:30', '22:00', '22:30', '23:00', '23:30']

filepath = open('file_path.txt').read()
for file_f in os.listdir(filepath):
    if file_f.endswith(".xlsx"):
        wb = xlrd.open_workbook(filepath + file_f)
        if wb.sheet_names()[0] == 'Sheet1':
            start = datetime.now()
            sheet = wb.sheet_by_name('Sheet1')
            number_of_entries = sheet.nrows
            customer = str(sheet.row(1)[0])[7:-2]
            bulk = db[customer].initialize_ordered_bulk_op()
            for i in range(1, number_of_entries):
                date_time = datetime(1899, 12, 30) + dt.timedelta(days=sheet.row(i)[2].value)
                date = date_time.strftime('%Y/%m/%d')
                time = date_time.strftime('%H:%M')
                time_index = str(time_half_hour.index(time) + 1)
                value = int(sheet.row(i)[3].value)
                bulk.find({"date": date}).upsert().update({"$set": {'energy_kwh.' + time_index: value}})
            bulk.execute()
            end = datetime.now()
            print str(number_of_entries) + " entries were added to the DB in " + str((end - start).seconds) \
                + " seconds, from " + file_f + "."
            print db[customer].find().skip(db[customer].count() - 1)[0]
        else:
            start = datetime.now()
            sheet = wb.sheet_by_name('Meter Data')
            number_of_entries = sheet.nrows
            customer = str(sheet.row(0)[0])[-14:-1]
            bulk = db[customer].initialize_ordered_bulk_op()
            for i in range(10, number_of_entries):
                date_time = datetime.strptime(sheet.row(i)[0].value, '%Y/%m/%d %H:%M')
                date = date_time.strftime('%Y/%m/%d')
                time = date_time.strftime('%H:%M')
                time_index = str(time_half_hour.index(time) + 1)
                value = int(sheet.row(i)[1].value)
                bulk.find({"date": date}).upsert().update({"$set": {'energy_kwh.' + time_index: value}})
            bulk.execute()
            end = datetime.now()
            print str(number_of_entries) + " entries were added to the DB in " + str((end - start).seconds) \
                + " seconds, from x" + file_f + "."
            print db[customer].find().skip(db[customer].count() - 1)[0]

main_end = datetime.now()
print "Transfer from 22 files, was completed in " + str((main_end - main_start).seconds) + " seconds."
