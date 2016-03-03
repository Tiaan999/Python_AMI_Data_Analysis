from pymongo import MongoClient

db = MongoClient("mongodb://localhost:27017/")['test']
collections = db.collection_names()
coll1 = collections[0]
collection1 = db[coll1].find({'date': {'$eq': '2011/06/29'}})
for doc in collection1:
    print doc
