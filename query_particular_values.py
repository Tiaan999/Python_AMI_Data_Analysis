from pymongo import MongoClient
from datetime import datetime

db = MongoClient("mongodb://localhost:27017/")['test']
collections = db.collection_names()
for coll in collections:
    first_date = datetime.strptime(db[coll].find().limit(1)[0]['date'], '%Y/%m/%d').date()
    last_date = datetime.strptime(db[coll].find().skip(db[coll].count() - 1)[0]['date'], '%Y/%m/%d').date()
    number_of_days_between_dates = (last_date - first_date).days + 1
    number_of_days_in_record = db[coll].count()
    percent_complete = float(number_of_days_in_record)/float(number_of_days_between_dates)*100
    print 'Customer: ' + coll + ', First: ' + str(first_date) + ', Last: ' + str(last_date) + ', Percentage complete: '\
          + str(("{0:.0f}".format(round(percent_complete)))) + '%'
